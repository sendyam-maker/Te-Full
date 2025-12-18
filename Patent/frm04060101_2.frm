VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04060101_2 
   Appearance      =   0  '平面
   BorderStyle     =   1  '單線固定
   Caption         =   "專利公報資料維護"
   ClientHeight    =   6090
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
   ScaleHeight     =   6090
   ScaleWidth      =   8960
   Begin VB.TextBox textNA01 
      Height          =   270
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   7
      Top             =   2340
      Width           =   852
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3405
      Left            =   30
      TabIndex        =   64
      Top             =   2640
      Width           =   8895
      _ExtentX        =   15699
      _ExtentY        =   5997
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   "中文名稱"
      TabPicture(0)   =   "frm04060101_2.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label28"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label26"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label25"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label24"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label23"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Label22"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label21"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "Label20"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label19"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "Label18"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "Label17"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtcAppl(9)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtcAppl(8)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtcAppl(7)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtcAppl(6)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtcAppl(5)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtcAppl(4)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtcAppl(3)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtcAppl(2)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtcAppl(1)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtcAppl(0)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).ControlCount=   21
      TabCaption(1)   =   "英文名稱"
      TabPicture(1)   =   "frm04060101_2.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txteAppl(0)"
      Tab(1).Control(1)=   "txteAppl(1)"
      Tab(1).Control(2)=   "txteAppl(2)"
      Tab(1).Control(3)=   "txteAppl(3)"
      Tab(1).Control(4)=   "txteAppl(4)"
      Tab(1).Control(5)=   "txteAppl(5)"
      Tab(1).Control(6)=   "txteAppl(6)"
      Tab(1).Control(7)=   "txteAppl(7)"
      Tab(1).Control(8)=   "txteAppl(8)"
      Tab(1).Control(9)=   "txteAppl(9)"
      Tab(1).Control(10)=   "Label38"
      Tab(1).Control(11)=   "Label37"
      Tab(1).Control(12)=   "Label36"
      Tab(1).Control(13)=   "Label35"
      Tab(1).Control(14)=   "Label34"
      Tab(1).Control(15)=   "Label33"
      Tab(1).Control(16)=   "Label32"
      Tab(1).Control(17)=   "Label31"
      Tab(1).Control(18)=   "Label30"
      Tab(1).Control(19)=   "Label29"
      Tab(1).ControlCount=   20
      TabCaption(2)   =   "其他"
      TabPicture(2)   =   "frm04060101_2.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtTPB37"
      Tab(2).Control(1)=   "txtTPB36"
      Tab(2).Control(2)=   "txtTPB35"
      Tab(2).Control(3)=   "txtTPB34"
      Tab(2).Control(4)=   "Label43"
      Tab(2).Control(5)=   "Label42"
      Tab(2).Control(6)=   "Label41"
      Tab(2).Control(7)=   "Label40"
      Tab(2).Control(8)=   "Label39"
      Tab(2).ControlCount=   9
      Begin VB.TextBox txtTPB37 
         Height          =   600
         Left            =   -73530
         MaxLength       =   300
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   39
         Top             =   1650
         Width           =   7275
      End
      Begin VB.TextBox txtTPB36 
         Height          =   600
         Left            =   -73530
         MaxLength       =   600
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   38
         Top             =   1020
         Width           =   7275
      End
      Begin VB.TextBox txtTPB35 
         Height          =   264
         Left            =   -73530
         MaxLength       =   7
         TabIndex        =   37
         Top             =   720
         Width           =   1185
      End
      Begin VB.TextBox txtTPB34 
         Height          =   264
         Left            =   -73530
         MaxLength       =   7
         TabIndex        =   36
         Top             =   420
         Width           =   1185
      End
      Begin MSForms.TextBox txteAppl 
         Height          =   300
         Index           =   0
         Left            =   -73650
         TabIndex        =   26
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
         Index           =   1
         Left            =   -73650
         TabIndex        =   27
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
         Index           =   2
         Left            =   -73650
         TabIndex        =   28
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
         Index           =   3
         Left            =   -73650
         TabIndex        =   29
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
         Index           =   4
         Left            =   -73650
         TabIndex        =   30
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
         Index           =   5
         Left            =   -73650
         TabIndex        =   31
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
         Index           =   6
         Left            =   -73650
         TabIndex        =   32
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
         Index           =   7
         Left            =   -73650
         TabIndex        =   33
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
         Index           =   8
         Left            =   -73650
         TabIndex        =   34
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
         Index           =   9
         Left            =   -73650
         TabIndex        =   35
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
         Index           =   0
         Left            =   1380
         TabIndex        =   16
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
      Begin MSForms.TextBox txtcAppl 
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   17
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
         Index           =   2
         Left            =   1380
         TabIndex        =   18
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
         Index           =   3
         Left            =   1380
         TabIndex        =   19
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
         Index           =   4
         Left            =   1380
         TabIndex        =   20
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
         Index           =   5
         Left            =   1380
         TabIndex        =   21
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
         Index           =   6
         Left            =   1380
         TabIndex        =   22
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
         Index           =   7
         Left            =   1380
         TabIndex        =   23
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
         Index           =   8
         Left            =   1380
         TabIndex        =   24
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
         Index           =   9
         Left            =   1380
         TabIndex        =   25
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
         TabIndex        =   90
         Top             =   2310
         Width           =   2340
      End
      Begin VB.Label Label42 
         AutoSize        =   -1  'True
         Caption         =   "優先權國家 :"
         Height          =   180
         Left            =   -74790
         TabIndex        =   89
         Top             =   1680
         Width           =   990
      End
      Begin VB.Label Label41 
         AutoSize        =   -1  'True
         Caption         =   "優先權號 :"
         Height          =   180
         Left            =   -74790
         TabIndex        =   88
         Top             =   1050
         Width           =   810
      End
      Begin VB.Label Label40 
         AutoSize        =   -1  'True
         Caption         =   "最早優先權日 :"
         Height          =   180
         Left            =   -74790
         TabIndex        =   87
         Top             =   750
         Width           =   1170
      End
      Begin VB.Label Label39 
         AutoSize        =   -1  'True
         Caption         =   "申請日 :"
         Height          =   180
         Left            =   -74790
         TabIndex        =   86
         Top             =   480
         Width           =   630
      End
      Begin VB.Label Label38 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱1 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   85
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label Label37 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱2 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   84
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label36 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱3 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   83
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label35 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱4 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   82
         Top             =   1230
         Width           =   1080
      End
      Begin VB.Label Label34 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱5 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   81
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label Label33 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱6 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   80
         Top             =   1770
         Width           =   1080
      End
      Begin VB.Label Label32 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱7 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   79
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label Label31 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱8 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   78
         Top             =   2310
         Width           =   1080
      End
      Begin VB.Label Label30 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱9 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   77
         Top             =   2580
         Width           =   1080
      End
      Begin VB.Label Label29 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱10 :"
         Height          =   180
         Left            =   -74910
         TabIndex        =   76
         Top             =   2850
         Width           =   1170
      End
      Begin VB.Label Label17 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱1 :"
         Height          =   180
         Left            =   120
         TabIndex        =   75
         Top             =   420
         Width           =   1080
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱2 :"
         Height          =   180
         Left            =   120
         TabIndex        =   74
         Top             =   690
         Width           =   1080
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱3 :"
         Height          =   180
         Left            =   120
         TabIndex        =   73
         Top             =   960
         Width           =   1080
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱4 :"
         Height          =   180
         Left            =   120
         TabIndex        =   72
         Top             =   1230
         Width           =   1080
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱5 :"
         Height          =   180
         Left            =   120
         TabIndex        =   71
         Top             =   1500
         Width           =   1080
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱6 :"
         Height          =   180
         Left            =   120
         TabIndex        =   70
         Top             =   1770
         Width           =   1080
      End
      Begin VB.Label Label23 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱7 :"
         Height          =   180
         Left            =   120
         TabIndex        =   69
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label Label24 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱8 :"
         Height          =   180
         Left            =   120
         TabIndex        =   68
         Top             =   2310
         Width           =   1080
      End
      Begin VB.Label Label25 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱9 :"
         Height          =   180
         Left            =   120
         TabIndex        =   67
         Top             =   2580
         Width           =   1080
      End
      Begin VB.Label Label26 
         AutoSize        =   -1  'True
         Caption         =   "申請人名稱10 :"
         Height          =   180
         Left            =   120
         TabIndex        =   66
         Top             =   2850
         Width           =   1170
      End
      Begin VB.Label Label28 
         AutoSize        =   -1  'True
         Caption         =   "PS. 申請人名稱至41卷29期(20141011)開始匯入"
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
         Left            =   150
         TabIndex        =   65
         Top             =   3120
         Width           =   4230
      End
   End
   Begin VB.TextBox txtTPB13 
      Height          =   270
      Left            =   5700
      MaxLength       =   1
      TabIndex        =   15
      Top             =   2040
      Width           =   315
   End
   Begin VB.TextBox txtTPB12 
      Height          =   270
      Left            =   7710
      MaxLength       =   2
      TabIndex        =   14
      Top             =   1740
      Width           =   615
   End
   Begin VB.TextBox txtTPB11 
      Height          =   270
      Left            =   5700
      MaxLength       =   2
      TabIndex        =   13
      Top             =   1740
      Width           =   615
   End
   Begin VB.TextBox txtTPB10 
      Height          =   270
      Left            =   7080
      MaxLength       =   15
      TabIndex        =   12
      Top             =   1440
      Width           =   1725
   End
   Begin VB.TextBox txtTPB09 
      Height          =   270
      Left            =   5700
      MaxLength       =   1
      TabIndex        =   11
      Top             =   1440
      Width           =   315
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1530
      MaxLength       =   12
      TabIndex        =   0
      Top             =   540
      Visible         =   0   'False
      Width           =   2772
   End
   Begin VB.CommandButton Command1 
      Caption         =   "下一筆公告號(&N)"
      Height          =   405
      Index           =   1
      Left            =   1740
      TabIndex        =   41
      Top             =   60
      Width           =   1560
   End
   Begin VB.CommandButton Command1 
      Caption         =   "上一筆公告號(&P)"
      Height          =   405
      Index           =   0
      Left            =   120
      TabIndex        =   40
      Top             =   60
      Width           =   1560
   End
   Begin VB.TextBox text06_2 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   2430
      Locked          =   -1  'True
      TabIndex        =   45
      TabStop         =   0   'False
      Top             =   2040
      Width           =   1812
   End
   Begin VB.TextBox text09 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   5700
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Top             =   1140
      Width           =   2772
   End
   Begin VB.TextBox text07_1 
      Height          =   270
      Left            =   5700
      MaxLength       =   4
      TabIndex        =   9
      Top             =   510
      Width           =   852
   End
   Begin VB.TextBox text06_1 
      Height          =   270
      Left            =   1530
      MaxLength       =   3
      TabIndex        =   6
      Top             =   2040
      Width           =   852
   End
   Begin VB.TextBox text05 
      Height          =   270
      Left            =   2970
      MaxLength       =   2
      TabIndex        =   5
      Top             =   1740
      Width           =   852
   End
   Begin VB.TextBox text04 
      Height          =   270
      Left            =   1530
      MaxLength       =   2
      TabIndex        =   4
      Top             =   1740
      Width           =   852
   End
   Begin VB.TextBox text03 
      Height          =   270
      Left            =   1530
      MaxLength       =   7
      TabIndex        =   3
      Top             =   1440
      Width           =   2772
   End
   Begin VB.TextBox text02 
      Height          =   270
      Left            =   1530
      MaxLength       =   7
      TabIndex        =   2
      Top             =   1140
      Width           =   2772
   End
   Begin VB.TextBox text01 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1530
      Locked          =   -1  'True
      MaxLength       =   12
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   840
      Width           =   2772
   End
   Begin VB.CommandButton buttonOK 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   405
      Left            =   6840
      TabIndex        =   42
      Top             =   60
      Width           =   800
   End
   Begin VB.CommandButton buttonCancel 
      Cancel          =   -1  'True
      Caption         =   "回前畫面(&U)"
      CausesValidation=   0   'False
      Height          =   405
      Left            =   7680
      TabIndex        =   43
      Top             =   60
      Width           =   1200
   End
   Begin MSForms.ComboBox Combo1 
      Height          =   300
      Left            =   2430
      TabIndex        =   8
      Top             =   2310
      Width           =   1815
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3201;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox text07_2 
      Height          =   300
      Left            =   6600
      TabIndex        =   10
      Top             =   510
      Width           =   1815
      VariousPropertyBits=   671107099
      Size            =   "3201;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox text08 
      Height          =   255
      Left            =   5700
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   840
      Width           =   2775
      VariousPropertyBits=   671105055
      Size            =   "4895;450"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label45 
      Caption         =   "（臺灣要分到縣市統計用）"
      ForeColor       =   &H000000C0&
      Height          =   165
      Left            =   4320
      TabIndex        =   92
      Top             =   2400
      Width           =   2355
   End
   Begin VB.Label Label44 
      Caption         =   "地區名稱 :"
      Height          =   240
      Left            =   270
      TabIndex        =   91
      Top             =   2340
      Width           =   1215
   End
   Begin VB.Label Label27 
      AutoSize        =   -1  'True
      Caption         =   "案件屬性  :"
      Height          =   180
      Left            =   4800
      TabIndex        =   63
      Top             =   2070
      Width           =   855
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "產業別分類  :"
      Height          =   180
      Left            =   6630
      TabIndex        =   62
      Top             =   1770
      Width           =   1035
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "IPC分類  :"
      Height          =   180
      Left            =   4890
      TabIndex        =   61
      Top             =   1770
      Width           =   765
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "國際分類號 :"
      Height          =   180
      Left            =   6060
      TabIndex        =   60
      Top             =   1470
      Width           =   990
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "新型是否修正  :"
      Height          =   180
      Left            =   4440
      TabIndex        =   59
      Top             =   1470
      Width           =   1215
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "新申請案號 :"
      Height          =   180
      Left            =   270
      TabIndex        =   58
      Top             =   540
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
      Left            =   3360
      TabIndex        =   57
      Top             =   150
      Width           =   3405
   End
   Begin VB.Label Label10 
      Caption         =   "期"
      Height          =   255
      Left            =   3990
      TabIndex        =   56
      Top             =   1740
      Width           =   255
   End
   Begin VB.Label Label9 
      Caption         =   "卷"
      Height          =   255
      Left            =   2430
      TabIndex        =   55
      Top             =   1740
      Width           =   375
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "本所案號  :"
      Height          =   180
      Left            =   4440
      TabIndex        =   54
      Top             =   1140
      Width           =   855
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "事務所名稱 :"
      Height          =   180
      Left            =   4440
      TabIndex        =   53
      Top             =   840
      Width           =   990
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "代理人 :"
      Height          =   180
      Left            =   4440
      TabIndex        =   52
      Top             =   540
      Width           =   630
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "申請人國籍 :"
      Height          =   180
      Left            =   270
      TabIndex        =   51
      Top             =   2040
      Width           =   990
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "公報 :"
      Height          =   180
      Left            =   270
      TabIndex        =   50
      Top             =   1740
      Width           =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "公告日 :"
      Height          =   180
      Left            =   270
      TabIndex        =   49
      Top             =   1440
      Width           =   630
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "公告號 /證書號:"
      Height          =   180
      Left            =   270
      TabIndex        =   48
      Top             =   1140
      Width           =   1215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "申請案號 :"
      Height          =   180
      Left            =   270
      TabIndex        =   44
      Top             =   840
      Width           =   810
   End
End
Attribute VB_Name = "frm04060101_2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Morgan 2021/12/22 改成Form2.0 (text07_2,text08,txtcAppl,txteAppl,Combo1)
'Memo By Morgan 2012/12/11 智權人員欄已修改
'Memo by Morgan2010/12/28 申請案號欄已修改
'2010/12/3 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/11 日期欄已修改
Option Explicit
'Memo by Morgan 2008/11/18
'國內案公告不必控管是否有國外案未發文，若要管制則將於相同案之分案、發文、提申也加入國內案檢查

Dim m_EditMode As Integer
Dim m_DataKey As String
Dim m_CurrTPB02 As String
Dim m_CurrTPB03 As String
Dim m_CurrTPB04 As String
Dim m_CurrTPB05 As String
'Add By Cheng 2002/11/27
Dim m_blnUpdateCP22 As Boolean
'Add by Morgan 2004/7/14
Dim m_strNextFeeDate As String  '下次繳費日本所期限
Dim m_strNextDueDate As String  '下次繳費日法定期限
'edit by nickc 2007/02/02
'Dim pA(1 To T_PA) As String
Dim pa() As String

Dim m_str421CP09 As String '技術報告總收文號
Dim m_str421CP14 As String '技術報告承辦人
Dim m_str421CP48 As String '技術報告承辦期限
Dim m_str421EP06 As String '技術報告文件齊備日
'Add by Morgan 2004/7/29
Dim m_strPA11 As String    '申請案號
'Add by Morgan 2006/10/13
Dim m_strPA14 As String '預定公告日
Dim m_bol412 As Boolean '是否有發文延緩公告


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
            'Modify By Cheng 2002/11/06
'            OnWork
            If OnWork = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
            
            '若有未發文技術報告時發 Mail 通知承辦人
            If m_str421CP09 <> "" And m_str421CP14 <> "" Then
               Dim stPS As String
               stPS = "※注意，本案已公告已可承辦且承辦期限為 " & ChangeTStringToTDateString(Format(Val(m_str421CP48) - 19110000)) & "！"
               Call PUB_SendMail(strUserNum, m_str421CP14, m_str421CP09, "技術報告文件齊備通知", "", stPS)
               m_str421CP09 = "": m_str421CP14 = "": m_str421CP48 = "": m_str421EP06 = ""
            End If
            
            'Add by Morgan 2008/1/15
            If m_EditMode = "1" And Me.Text1.Visible = True And text09.Text <> "" Then
               MsgBox "原申請號為本所案件，下一程序之年費期限可能已修改但無法還原，請自行檢查並更正！"
            End If
            
            Select Case m_EditMode
               Case 0:
                  'Modify by Morgan 2004/8/2
                  'nNo = Val(text02.Text)
                  'nNo = nNo + 1
                  'm_CurrTPB02 = CStr(nNo)
                  nNo = Val(Mid(text02.Text, 2))
                  nNo = nNo + 1
                  m_CurrTPB02 = Left(text02.Text, 1) & CStr(nNo)
                  m_CurrTPB03 = text03
                  m_CurrTPB04 = text04
                  m_CurrTPB05 = text05
            End Select
            'Add By Cheng 2003/01/16
            '隱藏新申請案號欄位
            Me.Text1.Visible = False
            Me.Label12.Visible = False
            Me.Hide
            frm04060101_1.Show
            frm04060101_1.SetInputTPB01
         End If
      ' 刪除
      Case 3:
         strTit = "詢問"
         strMsg = "是否要刪除此筆資料?"
         If MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit) = vbYes Then
            'Modify By Cheng 2002/11/06
'            OnWork
            If OnWork = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub
            'Add By Cheng 2003/01/16
            '隱藏新申請案號欄位
            Me.Text1.Visible = False
            Me.Label12.Visible = False
            Me.Hide
            frm04060101_1.Show
            frm04060101_1.SetInputTPB01
         End If
      Case Else:
        'Add By Cheng 2002/11/21
        If m_EditMode = 1 Then
            Me.text02.SetFocus
        End If
        'Add By Cheng 2003/01/16
        '隱藏新申請案號欄位
        Me.Text1.Visible = False
        Me.Label12.Visible = False
        Me.Hide
        frm04060101_1.Show
        frm04060101_1.SetInputTPB01
   End Select
   frm04060101_1.UpdateRecord m_DataKey
EXITSUB:
End Sub
' 使用者按下取消的按鍵
Private Sub buttonCancel_Click()
   Dim strTit As String
   Dim strMsg As String
   Dim nResponse
   
   strTit = "詢問"
   strMsg = "你並未存檔, 確定離開嗎?"
    'Modify By Cheng 2002/11/22
    '取消顯示訊息
'   nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
'   If nResponse = vbYes Then
        'Add By Cheng 2002/11/21
        If m_EditMode = 1 Then
            Me.text02.SetFocus
        End If
        'Add By Cheng 2003/01/16
        '隱藏新申請案號欄位
        Me.Text1.Visible = False
        Me.Label12.Visible = False
      Me.Hide
      frm04060101_1.Show
      frm04060101_1.SetInputTPB01
'   End If
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

'Remove by Morgan 2011/5/12
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

Private Function IsTPB02Exist(ByVal strTPB02 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   IsTPB02Exist = False
   strSql = "SELECT * FROM TPBulletin " & _
            "WHERE TPB02 = '" & strTPB02 & "' AND " & _
                  "TPB01 <> '" & m_DataKey & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   If rsTmp.RecordCount > 0 Then
      IsTPB02Exist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 此模組在處理資料到資料庫的工作
'Modify By Cheng 2002/11/06
'Public Sub OnWork()
Public Function OnWork() As Boolean
   Dim strSql As String
   Dim strDate As String
   Dim strFreeAgentCode As String
   'Add By Cheng 2002/11/27
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim i As Integer, strcAppl(0 To 10) As String
   Dim streAppl(0 To 10) As String 'Add By Sindy 2018/11/12
   
'Add By Cheng 2002/11/06
On Error GoTo ErrorHandler
OnWork = True
cnnConnection.BeginTrans

   strDate = Empty
   If IsEmpty(text03) = False Then
      strDate = ChangeTStringToWString(text03)
   End If
   
   ' 代理人是鍵入名稱時
   'If IsEmpty(text07_1) = True And IsEmpty(text07_2) = False Then
   '   strFreeAgentCode = GetFreeAgentCode
   '   strSQL = "INSERT INTO TAgent (TA01, TA02, TA03) VALUES ('P','" & strFreeAgentCode & "','" & text07_2 & "')"
   '   cnnConnection.Execute strSQL
   '   text07_1 = strFreeAgentCode
   'End If
   
   'Add By Sindy 2017/2/21
   For i = 0 To 9
      strcAppl(i) = ""
      streAppl(i) = "" 'Add By Sindy 2018/11/12
   Next i
   For i = 0 To 9
      If Trim(txtcAppl(i)) <> "" Then
         strcAppl(i) = txtcAppl(i).Text
      End If
      'Add By Sindy 2018/11/12
      If Trim(txteAppl(i)) <> "" Then
         streAppl(i) = txteAppl(i).Text
      End If
      '2018/11/12 END
   Next i
   '2017/2/21 END
   Select Case m_EditMode
      ' 新增資料到國內專利公報檔
      Case 0:
'         If strDate <> Empty Then
            'Modify by Morgan 2004/7/30
            '加TPB09
            'Modify By Sindy 2013/10/29 +TPB10,TPB11,TPB12
            'Modify By Sindy 2019/9/4 +,Trim(Mid(Combo1.Text, 4))
            strSql = "Insert into TPBulletin " & _
                     "(TPB01,TPB02" & IIf(strDate <> Empty, ",TPB03", "") & ",TPB04,TPB05,TPB06,TPB07,TPB08,TPB09,TPB10,TPB11,TPB12,TPB13" & _
                     ",TPB14,TPB15,TPB16,TPB17,TPB18,TPB19,TPB20,TPB21,TPB22,TPB23" & _
                     ",TPB24,TPB25,TPB26,TPB27,TPB28,TPB29,TPB30,TPB31,TPB32,TPB33" & _
                     ",TPB34,TPB35,TPB36,TPB37,TPB38" & _
                     ") Values ('" & text01 & "','" & text02 & "'" & IIf(strDate <> Empty, "," & strDate, "") & ",'" & text04 & "','" & _
                     text05 & "','" & text06_1 & "','" & text07_1 & "','" & text08 & "'," & CNULL(txtTPB09.Text) & _
                     "," & CNULL(txtTPB10.Text) & "," & CNULL(Format(txtTPB11.Text, "00")) & "," & CNULL(txtTPB12.Text) & "," & CNULL(txtTPB13.Text) & _
                     "," & CNULL(ChgSQL(strcAppl(0))) & "," & CNULL(ChgSQL(strcAppl(1))) & "," & CNULL(ChgSQL(strcAppl(2))) & "," & CNULL(ChgSQL(strcAppl(3))) & "," & CNULL(ChgSQL(strcAppl(4))) & _
                     "," & CNULL(ChgSQL(strcAppl(5))) & "," & CNULL(ChgSQL(strcAppl(6))) & "," & CNULL(ChgSQL(strcAppl(7))) & "," & CNULL(ChgSQL(strcAppl(8))) & "," & CNULL(ChgSQL(strcAppl(9))) & _
                     "," & CNULL(ChgSQL(streAppl(0))) & "," & CNULL(ChgSQL(streAppl(1))) & "," & CNULL(ChgSQL(streAppl(2))) & "," & CNULL(ChgSQL(streAppl(3))) & "," & CNULL(ChgSQL(streAppl(4))) & _
                     "," & CNULL(ChgSQL(streAppl(5))) & "," & CNULL(ChgSQL(streAppl(6))) & "," & CNULL(ChgSQL(streAppl(7))) & "," & CNULL(ChgSQL(streAppl(8))) & "," & CNULL(ChgSQL(streAppl(9))) & _
                     "," & CNULL(DBDATE(txtTPB34), True) & "," & CNULL(DBDATE(txtTPB35), True) & "," & CNULL(txtTPB36) & "," & CNULL(txtTPB37) & "," & CNULL(Trim(Mid(Combo1.Text, 4))) & _
                     ")"
'         Else
'            'Modify by Morgan 2004/7/30
'            '加TPB09
'            'Modify By Sindy 2013/10/29 +TPB10,TPB11,TPB12
'            strSql = "Insert into TPBulletin " & _
'                     "(TPB01,TPB02,TPB04,TPB05,TPB06,TPB07,TPB08,TPB09,TPB10,TPB11,TPB12,TPB13" & _
'                     ",TPB14,TPB15,TPB16,TPB17,TPB18,TPB19,TPB20,TPB21,TPB22,TPB23" & _
'                     ",TPB24,TPB25,TPB26,TPB27,TPB28,TPB29,TPB30,TPB31,TPB32,TPB33" & _
'                     ",TPB34,TPB35,TPB36,TPB37" & _
'                     ") Values ('" & text01 & "','" & text02 & "','" & text04 & "','" & _
'                     text05 & "','" & text06_1 & "','" & text07_1 & "','" & text08 & "'," & CNULL(txtTPB09.Text) & _
'                     "," & CNULL(txtTPB10.Text) & "," & CNULL(txtTPB11.Text) & "," & CNULL(txtTPB12.Text) & "," & CNULL(txtTPB13.Text) & _
'                     "," & CNULL(strcAppl(0)) & "," & CNULL(strcAppl(1)) & "," & CNULL(strcAppl(2)) & "," & CNULL(strcAppl(3)) & "," & CNULL(strcAppl(4)) & _
'                     "," & CNULL(strcAppl(5)) & "," & CNULL(strcAppl(6)) & "," & CNULL(strcAppl(7)) & "," & CNULL(strcAppl(8)) & "," & CNULL(strcAppl(9)) & _
'                     "," & CNULL(streAppl(0)) & "," & CNULL(streAppl(1)) & "," & CNULL(streAppl(2)) & "," & CNULL(streAppl(3)) & "," & CNULL(streAppl(4)) & _
'                     "," & CNULL(streAppl(5)) & "," & CNULL(streAppl(6)) & "," & CNULL(streAppl(7)) & "," & CNULL(streAppl(8)) & "," & CNULL(streAppl(9)) & _
'                     "," & CNULL(DBDATE(txtTPB34), True) & "," & CNULL(DBDATE(txtTPB35), True) & "," & CNULL(txtTPB36) & "," & CNULL(txtTPB37) & _
'                     ")"
'         End If
         cnnConnection.Execute strSql
         
         'Modify by Morgan 2004/8/4
         '本所申請案才更新
         If text09.Text <> "" Then
            'Modify by Morgan 2004/8/2
            '93.8.1 以後公告號改為證書號
            If Val(text03) >= 930801 Then
               ' 更新專利基本檔的公告日,專利號數
               '2008/3/20 modify by sonia 同時更新公告號
               strSql = "UPDATE Patent SET PA14 = " & ChangeTStringToWString(text03) & _
                        ",PA22 = '" & text02 & "' ,PA15 = '" & text02 & "' " & _
                        " WHERE PA11 = '" & m_DataKey & "'"
            Else
            ' 更新專利基本檔的公告日及公告號
               strSql = "UPDATE Patent SET PA14 = " & ChangeTStringToWString(text03) & ", " & _
                        " PA15 = '" & text02 & "' " & _
                        " WHERE PA11 = '" & m_DataKey & "'"
            End If
            cnnConnection.Execute strSql
            
            'Add By Cheng 2002/11/27
            If m_blnUpdateCP22 Then
               StrSQLa = " Select * From Patent Where PA11='" & Me.text01.Text & "'"
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               While Not rsA.EOF
                   strSql = "Update CaseProgress Set CP22 ='N' Where " & ChgCaseprogress(rsA.Fields(0).Value & rsA.Fields(1).Value & rsA.Fields(2).Value & rsA.Fields(3).Value) & " And ((CP10 >= '101' AND CP10 <= '105' ) or cp10='125') And CP09 <'C' "
                   cnnConnection.Execute strSql
                   rsA.MoveNext
               Wend
            End If
         
            'Add by Morgan 2004/7/15
            '若用新法則更新下一程序年費期限
            If Val(text03) >= 930701 Then
               'Modify by Morgan 2005/2/17 繳10年以上會有錯
               'm_strNextDueDate = CompDate(0, Val(Right(pa(72), 1)), TransDate(text03, 2))
               strExc(0) = Right(pa(72), 2)
               If Left(strExc(0), 1) = "," Then strExc(0) = Right(strExc(0), 1)
               m_strNextDueDate = CompDate(0, Val(strExc(0)), TransDate(text03, 2))
               '2005/2/17 end
               m_strNextDueDate = CompDate(2, -1, m_strNextDueDate)
               'Added by Morgan 2014/10/28
               'Modified by Morgan 2014/11/20 外專改回舊規則
               If strSrvDate(1) >= 台灣案所限新規則啟用日 And pa(1) <> "FCP" Then
                  m_strNextFeeDate = PUB_GetOurDeadline(m_strNextDueDate)
               'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
               ElseIf strSrvDate(1) >= 外專台灣案所限新規則啟用日 And pa(1) = "FCP" Then
                  m_strNextFeeDate = PUB_GetFCPOurDeadline(m_strNextDueDate, 2)
               'end 2019/7/11
               'end 2014/10/28
                  m_strNextFeeDate = CompDate(2, -2, m_strNextDueDate)
               End If 'Added by Morgan 2014/10/28
               
               If pa(1) = "P" Then 'Add by Morgan 2008/1/15 P案才要抓工作天
                  m_strNextFeeDate = PUB_GetWorkDay1(m_strNextFeeDate, True)
               End If
               strSql = "update nextprogress set NP08=" & m_strNextFeeDate & ", np09=" & m_strNextDueDate & _
                  " where np07='605'  and np06 is null and np02='" & pa(1) & "'" & _
                  " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'"
                  
               cnnConnection.Execute strSql, intI
            End If
            '內專若有未發文技術報告時更新文件齊備日(=公告日)及承辦期限
            If pa(1) = "P" Then
               If PUB_ChkCPExist(pa, "421", 1, m_str421CP09, m_str421CP14) = True Then
                  m_str421EP06 = TransDate(text03, 2)
                  '更新文件齊備日
                  strSql = "Update EngineerProgress Set EP06=" & strSrvDate(1) & " Where EP02='" & m_str421CP09 & "' AND EP06 IS NULL"
                  cnnConnection.Execute strSql
                  
                  If PUB_IfSetCP48(m_str421CP09) Then 'Add by Morgan 2010/10/5
                  
                     'Modify by Morgan 2007/10/12 承辦期限改呼叫共用函數計算
                     'm_str421CP48 = PUB_GetEngDueDate(m_str421EP06, pa(1), "000", "421")
                     m_str421CP48 = Pub_GetHandleDay(pa(1), "000", "421", m_str421EP06, , m_str421CP09)
                     'end 2007/10/12
                     If Val(m_str421CP48) > 0 Then
                        '更新承辦期限
                        strSql = "Update CaseProgress Set CP48=" & m_str421CP48 & " Where CP09='" & m_str421CP09 & "' AND CP48 IS NULL"
                        cnnConnection.Execute strSql
                     End If
                  End If 'Add by Morgan 2010/10/5
                  
                  'Added by Morgan 2019/12/11 非FMP案更新齊備日承辦期限在 Trigger 設定
                  If Val(m_str421CP48) = 0 Then
                     strExc(0) = "select cp48 from caseprogress where cp09='" & m_str421CP09 & "'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        m_str421CP48 = "" & RsTemp(0)
                     End If
                  End If
                  'end 2019/12/11
                  
               End If
            End If
            'END 2004/7/15
         End If
      
      ' 更新專利公報檔的資料
      Case 1:
         'Modify By Cheng 20036/01/16
         '若非修改申請案號
         If Me.Text1.Visible = False Then
'             If strDate <> Empty Then
               'Modify by Morgan 2004/7/30
               '加TPB09
               'Modify By Sindy 2013/10/29 +TPB10,TPB11,TPB12
               'Modify By Sindy 2019/9/4 +,TPB38=" & CNULL(Trim(Mid(Combo1.Text, 4)))
               strSql = "Update TPBulletin " & _
                        "Set TPB01='" & text01 & "',TPB02='" & text02 & "'" & _
                        IIf(strDate <> Empty, ",TPB03=" & strDate, "") & ",TPB04='" & text04 & "'" & _
                        ",TPB05='" & text05 & "',TPB06='" & text06_1 & "',TPB07='" & text07_1 & "',TPB08='" & text08 & "',TPB09=" & CNULL(txtTPB09) & _
                        ",TPB10=" & CNULL(txtTPB10) & ",TPB11=" & CNULL(Format(txtTPB11.Text, "00")) & ",TPB12=" & CNULL(txtTPB12) & ",TPB13=" & CNULL(txtTPB13) & _
                        ",TPB14=" & CNULL(ChgSQL(strcAppl(0))) & ",TPB15=" & CNULL(ChgSQL(strcAppl(1))) & ",TPB16=" & CNULL(ChgSQL(strcAppl(2))) & ",TPB17=" & CNULL(ChgSQL(strcAppl(3))) & ",TPB18=" & CNULL(ChgSQL(strcAppl(4))) & _
                        ",TPB19=" & CNULL(ChgSQL(strcAppl(5))) & ",TPB20=" & CNULL(ChgSQL(strcAppl(6))) & ",TPB21=" & CNULL(ChgSQL(strcAppl(7))) & ",TPB22=" & CNULL(ChgSQL(strcAppl(8))) & ",TPB23=" & CNULL(ChgSQL(strcAppl(9))) & _
                        ",TPB24=" & CNULL(ChgSQL(streAppl(0))) & ",TPB25=" & CNULL(ChgSQL(streAppl(1))) & ",TPB26=" & CNULL(ChgSQL(streAppl(2))) & ",TPB27=" & CNULL(ChgSQL(streAppl(3))) & ",TPB28=" & CNULL(ChgSQL(streAppl(4))) & _
                        ",TPB29=" & CNULL(ChgSQL(streAppl(5))) & ",TPB30=" & CNULL(ChgSQL(streAppl(6))) & ",TPB31=" & CNULL(ChgSQL(streAppl(7))) & ",TPB32=" & CNULL(ChgSQL(streAppl(8))) & ",TPB33=" & CNULL(ChgSQL(streAppl(9))) & _
                        ",TPB34=" & CNULL(DBDATE(txtTPB34), True) & ",TPB35=" & CNULL(DBDATE(txtTPB35), True) & _
                        ",TPB36=" & CNULL(txtTPB36) & ",TPB37=" & CNULL(txtTPB37) & ",TPB38=" & CNULL(Trim(Mid(Combo1.Text, 4))) & _
                        " Where TPB01='" & text01 & "'"
'             Else
'               strSql = "Update TPBulletin " & _
'                        "Set TPB01='" & text01 & "'," & "TPB02='" & text02 & "'," & "TPB04='" & text04 & "'," & _
'                        "TPB05='" & text05 & "'," & "TPB06='" & text06_1 & "'," & "TPB07='" & text07_1 & "'," & "TPB08='" & text08 & "',TPB09=" & CNULL(txtTPB09) & _
'                        ",TPB10=" & CNULL(txtTPB10) & ",TPB11=" & CNULL(txtTPB11) & ",TPB12=" & CNULL(txtTPB12) & ",TPB13=" & CNULL(txtTPB13) & _
'                        ",TPB14=" & CNULL(strcAppl(0)) & ",TPB15=" & CNULL(strcAppl(1)) & ",TPB16=" & CNULL(strcAppl(2)) & ",TPB17=" & CNULL(strcAppl(3)) & ",TPB18=" & CNULL(strcAppl(4)) & _
'                        ",TPB19=" & CNULL(strcAppl(5)) & ",TPB20=" & CNULL(strcAppl(6)) & ",TPB21=" & CNULL(strcAppl(7)) & ",TPB22=" & CNULL(strcAppl(8)) & ",TPB23=" & CNULL(strcAppl(9)) & _
'                        ",TPB24=" & CNULL(streAppl(0)) & ",TPB25=" & CNULL(streAppl(1)) & ",TPB26=" & CNULL(streAppl(2)) & ",TPB27=" & CNULL(streAppl(3)) & ",TPB28=" & CNULL(streAppl(4)) & _
'                        ",TPB29=" & CNULL(streAppl(5)) & ",TPB30=" & CNULL(streAppl(6)) & ",TPB31=" & CNULL(streAppl(7)) & ",TPB32=" & CNULL(streAppl(8)) & ",TPB33=" & CNULL(streAppl(9)) & _
'                        ",TPB34=" & CNULL(DBDATE(txtTPB34), True) & ",TPB35=" & CNULL(DBDATE(txtTPB35), True) & _
'                        ",TPB36=" & CNULL(txtTPB36) & ",TPB37=" & CNULL(txtTPB37) & _
'                        " Where TPB01='" & text01 & "'"
'             End If
             cnnConnection.Execute strSql
             
            'Modify by Morgan 2004/8/4
            '本所申請案才更新
            If text09.Text <> "" Then
               'Modify by Morgan 2004/8/2
               '93.8.1 以後公告號改為證書號
               If Val(text03) >= 930801 Then
                  ' 更新專利基本檔的公告日,專利號數
                  '2008/3/20 modify by sonia 同時更新公告號
                  strSql = "UPDATE Patent SET PA14 = " & ChangeTStringToWString(text03) & _
                           ",PA22 = '" & text02 & "',PA15 = '" & text02 & "'  " & _
                           " WHERE PA11 = '" & m_DataKey & "'"
               Else
                  ' 更新專利基本檔的公告日及公告號
                  strSql = "UPDATE Patent SET PA14 = " & ChangeTStringToWString(text03) & ", " & _
                                             "PA15 = '" & text02 & "' " & _
                           " WHERE PA11 = '" & m_DataKey & "'"
               End If
               cnnConnection.Execute strSql
                
               'Add by Morgan 2008/1/15 修改也要更新
               '若用新法則更新下一程序年費期限
               If Val(text03) >= 930701 Then
                  'Modify by Morgan 2005/2/17 繳10年以上會有錯
                  'm_strNextDueDate = CompDate(0, Val(Right(pa(72), 1)), TransDate(text03, 2))
                  strExc(0) = Right(pa(72), 2)
                  If Left(strExc(0), 1) = "," Then strExc(0) = Right(strExc(0), 1)
                  m_strNextDueDate = CompDate(0, Val(strExc(0)), TransDate(text03, 2))
                  '2005/2/17 end
                  m_strNextDueDate = CompDate(2, -1, m_strNextDueDate)
                  'Added by Morgan 2014/10/28
                  If strSrvDate(1) >= 台灣案所限新規則啟用日 Then
                     m_strNextFeeDate = PUB_GetOurDeadline(m_strNextDueDate)
                  Else
                  'end 2014/10/28
                     m_strNextFeeDate = CompDate(2, -2, m_strNextDueDate)
                  End If 'Added by Morgan 2014/10/28
                  
                  If pa(1) = "P" Then 'Add by Morgan 2008/1/15 P案才要抓工作天
                     m_strNextFeeDate = PUB_GetWorkDay1(m_strNextFeeDate, True)
                  End If
                  
                  strSql = "update nextprogress set NP08=" & m_strNextFeeDate & ", np09=" & m_strNextDueDate & _
                     " where np07='605'  and np06 is null and np02='" & pa(1) & "'" & _
                     " and np03='" & pa(2) & "' and np04='" & pa(3) & "' and np05='" & pa(4) & "'"
                     
                  cnnConnection.Execute strSql, intI
               End If
            End If
             
         '修改申請案號
         Else
             '新增新申請案號資料
            'Modify by Morgan 2004/7/30
            '加TPB09並指定新增欄位
            'Modify By Sindy 2013/10/29 +TPB10,TPB11,TPB12
            'Modify By Sindy 2019/9/4 +,Trim(Mid(Combo1.Text, 4))
             strSql = "Insert Into TPBulletin" & _
                     "(TPB01, TPB02, TPB03, TPB04, TPB05, TPB06, TPB07, TPB08, TPB09,TPB10,TPB11,TPB12,TPB13" & _
                     ",TPB14,TPB15,TPB16,TPB17,TPB18,TPB19,TPB20,TPB21,TPB22,TPB23" & _
                     ",TPB24,TPB25,TPB26,TPB27,TPB28,TPB29,TPB30,TPB31,TPB32,TPB33" & _
                     ",TPB34,TPB35,TPB36,TPB37,TPB38" & _
                     ") Values ('" & Me.Text1.Text & "','" & Me.text02.Text & "'," & IIf(Me.text03.Text = "", "NULL", DBDATE(Me.text03.Text)) & ",'" & Me.text04.Text & "','" & _
                     Me.text05.Text & "','" & Me.text06_1.Text & "','" & Me.text07_1.Text & "','" & Me.text08.Text & "'," & CNULL(txtTPB09.Text) & _
                     "," & CNULL(txtTPB10.Text) & "," & CNULL(Format(txtTPB11.Text, "00")) & "," & CNULL(txtTPB12.Text) & "," & CNULL(txtTPB13.Text) & _
                     "," & CNULL(ChgSQL(strcAppl(0))) & "," & CNULL(ChgSQL(strcAppl(1))) & "," & CNULL(ChgSQL(strcAppl(2))) & "," & CNULL(ChgSQL(strcAppl(3))) & "," & CNULL(ChgSQL(strcAppl(4))) & _
                     "," & CNULL(ChgSQL(strcAppl(5))) & "," & CNULL(ChgSQL(strcAppl(6))) & "," & CNULL(ChgSQL(strcAppl(7))) & "," & CNULL(ChgSQL(strcAppl(8))) & "," & CNULL(ChgSQL(strcAppl(9))) & _
                     "," & CNULL(ChgSQL(streAppl(0))) & "," & CNULL(ChgSQL(streAppl(1))) & "," & CNULL(ChgSQL(streAppl(2))) & "," & CNULL(ChgSQL(streAppl(3))) & "," & CNULL(ChgSQL(streAppl(4))) & _
                     "," & CNULL(ChgSQL(streAppl(5))) & "," & CNULL(ChgSQL(streAppl(6))) & "," & CNULL(ChgSQL(streAppl(7))) & "," & CNULL(ChgSQL(streAppl(8))) & "," & CNULL(ChgSQL(streAppl(9))) & _
                     "," & CNULL(DBDATE(txtTPB34), True) & "," & CNULL(DBDATE(txtTPB35), True) & "," & CNULL(txtTPB36) & "," & CNULL(txtTPB37) & "," & CNULL(Trim(Mid(Combo1.Text, 4))) & _
                     ")"
             cnnConnection.Execute strSql
            
            'Modify by Morgan 2004/8/2
            '93.8.1 以後公告號改為證書號
            If Val(text03) >= 930801 Then
               ' 更新專利基本檔的公告日
               '2008/3/20 modify by sonia 同時更新公告號
               strSql = "UPDATE Patent SET PA14 = " & IIf(Me.text03.Text = "", "NULL", ChangeTStringToWString(text03)) & _
                      ",PA22 = '" & text02 & "',PA15 = '" & text02 & "' " & _
                      " WHERE PA11 = '" & Me.Text1.Text & "'"
            Else
               ' 更新專利基本檔的公告日及公告號
               strSql = "UPDATE Patent SET PA14 = " & IIf(Me.text03.Text = "", "NULL", ChangeTStringToWString(text03)) & ", " & _
                      " PA15 = '" & text02 & "' " & _
                      " WHERE PA11 = '" & Me.Text1.Text & "'"
            End If
            cnnConnection.Execute strSql
             
             '刪除舊申請案號
             strSql = "Delete From TPBulletin Where TPB01='" & text01.Text & "'"
             cnnConnection.Execute strSql
             
            'Modify by Morgan 2004/8/4
            '本所申請案才更新
            If text09.Text <> "" Then
               'Modify by Morgan 2004/8/3
               ' 更新專利基本檔的公告日及公告號
               'Modify by Morgan 2004/8/2
               '93.8.1 以後公告號改為證書號
               If Val(text03) >= 930801 Then
                  '2008/3/20 modify by sonia 同時更新公告號
                  strSql = "UPDATE Patent SET PA14 = NULL , PA22 = DECODE(PA21,NULL,NULL,PA22) , PA15 = DECODE(PA21,NULL,NULL,PA22) " & _
                      " WHERE PA11 = '" & Me.text01.Text & "'"
               Else
                  strSql = "UPDATE Patent SET PA14 = NULL , PA15 = NULL " & _
                      " WHERE PA11 = '" & Me.text01.Text & "'"
               End If
                cnnConnection.Execute strSql
            End If
         End If
      Case 3:
         strSql = "Delete From TPBulletin where TPB01 = '" & m_DataKey & "'"
         cnnConnection.Execute strSql
         'Add by Morgan 2007/12/12 刪除時基本資料也要清除
         If Val(text03) >= 930801 Then
            '2008/3/20 modify by sonia 同時更新公告號
            strSql = "UPDATE Patent SET PA14 = NULL , PA22 = DECODE(PA21,NULL,NULL,PA22), PA15 = DECODE(PA21,NULL,NULL,PA22) WHERE PA11 = '" & Me.text01.Text & "'"
         Else
            strSql = "UPDATE Patent SET PA14 = NULL , PA15 = NULL WHERE PA11 = '" & Me.text01.Text & "'"
         End If
         cnnConnection.Execute strSql
         'end 2007/12/12
   End Select
'Add By Cheng 2002/11/06
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
         'add by Morgan 2004/8/2
         Erase pa
         'Add by Morgan 2007/2/12
         ReDim pa(1 To TF_PA) As String
         
         If m_EditMode = 1 Or m_EditMode = 2 Or m_EditMode = 3 Then
            Set rsTmp = New ADODB.Recordset
            strSql = "Select * from TPBulletin where TPB01 = '" & m_DataKey & "'"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenDynamic
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               If IsNull(rsTmp.Fields("TPB02")) = False Then
                  text02 = rsTmp.Fields("TPB02")
               End If
               If IsNull(rsTmp.Fields("TPB03")) = False Then
                  text03 = ChangeWStringToTString(rsTmp.Fields("TPB03"))
               End If
               If IsNull(rsTmp.Fields("TPB04")) = False Then
                  text04 = rsTmp.Fields("TPB04")
               End If
               If IsNull(rsTmp.Fields("TPB05")) = False Then
                  text05 = rsTmp.Fields("TPB05")
               End If
               If IsNull(rsTmp.Fields("TPB06")) = False Then
                  text06_1 = rsTmp.Fields("TPB06")
               End If
               If IsNull(rsTmp.Fields("TPB07")) = False Then
                  text07_1 = rsTmp.Fields("TPB07")
               End If
               'Add by Morgan 2004/8/3
               txtTPB09.Text = "" & rsTmp.Fields("TPB09")
               'Add By Sindy 2013/10/29
               txtTPB10.Text = "" & rsTmp.Fields("TPB10")
               txtTPB11.Text = "" & rsTmp.Fields("TPB11")
               txtTPB12.Text = "" & rsTmp.Fields("TPB12")
               '2013/10/29 END
               
               'Add By Sindy 2017/2/21
               txtTPB13.Text = "" & rsTmp.Fields("TPB13")
               txtcAppl(0).Text = "" & rsTmp.Fields("TPB14")
               txtcAppl(1).Text = "" & rsTmp.Fields("TPB15")
               txtcAppl(2).Text = "" & rsTmp.Fields("TPB16")
               txtcAppl(3).Text = "" & rsTmp.Fields("TPB17")
               txtcAppl(4).Text = "" & rsTmp.Fields("TPB18")
               txtcAppl(5).Text = "" & rsTmp.Fields("TPB19")
               txtcAppl(6).Text = "" & rsTmp.Fields("TPB20")
               txtcAppl(7).Text = "" & rsTmp.Fields("TPB21")
               txtcAppl(8).Text = "" & rsTmp.Fields("TPB22")
               txtcAppl(9).Text = "" & rsTmp.Fields("TPB23")
               '2017/2/21 END
               
               'Add By Sindy 2018/11/12
               txteAppl(0).Text = "" & rsTmp.Fields("TPB24")
               txteAppl(1).Text = "" & rsTmp.Fields("TPB25")
               txteAppl(2).Text = "" & rsTmp.Fields("TPB26")
               txteAppl(3).Text = "" & rsTmp.Fields("TPB27")
               txteAppl(4).Text = "" & rsTmp.Fields("TPB28")
               txteAppl(5).Text = "" & rsTmp.Fields("TPB29")
               txteAppl(6).Text = "" & rsTmp.Fields("TPB30")
               txteAppl(7).Text = "" & rsTmp.Fields("TPB31")
               txteAppl(8).Text = "" & rsTmp.Fields("TPB32")
               txteAppl(9).Text = "" & rsTmp.Fields("TPB33")
               If IsNull(rsTmp.Fields("TPB34")) = False Then
                  txtTPB34 = ChangeWStringToTString(rsTmp.Fields("TPB34"))
               End If
               If IsNull(rsTmp.Fields("TPB35")) = False Then
                  txtTPB35 = ChangeWStringToTString(rsTmp.Fields("TPB35"))
               End If
               txtTPB36.Text = "" & rsTmp.Fields("TPB36")
               txtTPB37.Text = "" & rsTmp.Fields("TPB37")
               '2018/11/12 END
               
               'Add By Sindy 2019/9/4
               Combo1.Text = "" & rsTmp.Fields("TPB38")
               If IsEmptyText(Combo1.Text) = False Then
                  textNA01 = GetNationNo(Combo1.Text)
                  Combo1.Text = textNA01 & " " & Combo1.Text
               End If
               '2019/9/4 END
               
               'If UpdateCtrlData = True Then 'Modify By Sindy 2012/4/3 Mark
                  UpdateCtrlData = UpdateCtrlData(1)
               'End If
               'If UpdateCtrlData = True Then
                  UpdateCtrlData = UpdateCtrlData(2)
               'End If
               'If UpdateCtrlData = True Then
                  UpdateCtrlData = UpdateCtrlData(3)
               'End If
            Else
               UpdateCtrlData = False
            End If
            rsTmp.Close
         End If
         If m_EditMode = 0 Then
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
            'Modify by Morgan 2011/1/3 要控制抓專利否則會抓到商標的代理人
            strSql = "Select * from TAGENT where TA02 = '" & text07_1 & "' and TA01='P'"
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
                        "PA09 = '000' and pa23='1'"
         text09 = Empty
         m_strPA14 = ""
         m_bol412 = False
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenDynamic
         If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            text09 = rsTmp.Fields("PA01") & "-" & rsTmp.Fields("PA02") & "-" & rsTmp.Fields("PA03") & "-" & rsTmp.Fields("PA04")
            'Add by  Morgan 2004/7/14
            pa(1) = rsTmp.Fields("PA01")
            pa(2) = rsTmp.Fields("PA02")
            pa(3) = rsTmp.Fields("PA03")
            pa(4) = rsTmp.Fields("PA04")
            pa(14) = "" & rsTmp.Fields("PA14")
            pa(22) = "" & rsTmp.Fields("PA22")
            pa(72) = "" & rsTmp.Fields("PA72")
            pa(21) = "" & rsTmp.Fields("PA21") 'Add by Morgan 2008/1/15
            
            'Add By Cheng 2003/05/01
            '若為新增狀況
            If m_EditMode = 0 Then
                Me.text07_1.Text = "01"
                text07_1_Validate False
            End If
            
            'Add by Morgan 2006/10/13 只控制內專就好
            If "" & rsTmp.Fields("PA01") = "P" Then
               m_strPA14 = PUB_GetPrePA14(pa, m_bol412)
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
         'Add By Sindy 2013/10/29
         txtTPB10.Locked = True
         txtTPB11.Locked = True
         txtTPB12.Locked = True
         '2013/10/29 END
         'Add By Sindy 2017/2/21
         text07_2.Locked = True
         txtTPB13.Locked = True
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
         '2017/2/21 END
         'Add By Sindy 2018/11/12
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
         txtTPB34.Locked = True
         txtTPB35.Locked = True
         txtTPB36.Locked = True
         txtTPB37.Locked = True
         '2018/11/12 END
         Combo1.Locked = True: textNA01.Locked = True 'Add By Sindy 2019/9/4
      Case Else:
         text02.Locked = False
         text03.Locked = False
         text04.Locked = False
         text05.Locked = False
         text06_1.Locked = False
         text07_1.Locked = False
         'Add By Sindy 2013/10/29
         txtTPB10.Locked = False
         txtTPB11.Locked = False
         txtTPB12.Locked = False
         '2013/10/29 END
         'Add By Sindy 2017/2/21
         text07_2.Locked = False
         txtTPB13.Locked = False
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
         '2017/2/21 END
         'Add By Sindy 2018/11/12
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
         txtTPB34.Locked = False
         txtTPB35.Locked = False
         txtTPB36.Locked = False
         txtTPB37.Locked = False
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
    'Modify By Cheng 2002/11/06
'            OnWork
    If OnWork = False Then MsgBox "存檔失敗，請洽系統管理員 !", vbCritical: Exit Sub

      Select Case m_EditMode
         Case 0:
            nNo = Val(text02.Text)
            nNo = nNo + 1
            m_CurrTPB02 = CStr(nNo)
            m_CurrTPB03 = text03
            m_CurrTPB04 = text04
            m_CurrTPB05 = text05
      End Select
      frm04060101_1.UpdateRecord m_DataKey
   Else
      Exit Sub
   End If
 
   If Index = 0 Then
      '上一筆
      i = frm04060101_1.grdList.row - 1
      If i > 0 Then
         frm04060101_1.grdList.row = frm04060101_1.grdList.row - 1
         frm04060101_1.textQuery = frm04060101_1.grdList.TextMatrix(i, 1)
      Else
         MsgBox "已是第一筆了 !", vbInformation
      End If
      
   Else
      i = frm04060101_1.grdList.row + 1
      If i < frm04060101_1.grdList.Rows Then
         frm04060101_1.grdList.row = frm04060101_1.grdList.row + 1
         frm04060101_1.textQuery = frm04060101_1.grdList.TextMatrix(i, 1)
      Else
         MsgBox "已是最後一筆了 !", vbInformation
      End If
   End If
   Select Case m_EditMode
      Case "1"
         frm04060101_1.buttonMod_Click
        'Add By Cheng 2002/11/19
        If Me.text02.Enabled Then Me.text02.SetFocus
      Case "2"
         frm04060101_1.buttonQuery_Click
      Case "3"
         frm04060101_1.buttonDel_Click
   End Select
End Sub

Private Sub Form_Activate()
    'Add By Cheng 2003/01/16
    If Me.Text1.Visible Then
        Me.Text1.SetFocus
    End If
End Sub

Private Sub Form_Initialize()
'add by nickc 2007/02/02
ReDim pa(1 To TF_PA) As String
End Sub

Private Sub Form_Load()
Dim rsTmp As New ADODB.Recordset
   
   'Add By Sindy 2018/11/12
   SSTab1.Tab = 0
   
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
   frm04060101_1.Show
   'Set frm04060101_2 = Nothing 'Removed by Morgan 2021/12/22 form2.0會有問題，改在呼叫時清除記憶體變數
End Sub

Private Sub text02_Validate(Cancel As Boolean)
   Dim nLen As Integer
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(text02) = False Then
      'Modify by Morgan 2004/8/2
      ''93.8.1 以後公告號改為證書號
      'nLen = Len(text02)
      'text02 = String(6 - nLen, "0") & text02
      'Add by Morgan 2004/8/3
      '申請號若大於8碼則不檢查證書號是否重複
      'Modify by Morgan 2010/12/28 申請案號改碼數
      'If Not (Val(text03) >= 930801 And Len(m_strPA11) > 8) Then
      If Not (Val(text03) >= 930801 And Len(m_strPA11) > 9) Then
         
         If IsTPB02Exist(text02) = True Then
            Cancel = True
            strTit = "公告號/證書號"
            strMsg = "公告號/證書號已存在"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            text02_GotFocus
         End If
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
         strMsg = "請輸入正確的公告日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text03_GotFocus
         GoTo EXITSUB
      End If
        'Modify By Cheng 2003/01/21
        '公告日不能大於系統日
'      If DBDATE(text03) >= DBDATE(SystemDate()) Then
      If DBDATE(text03) > DBDATE(SystemDate()) Then
         Cancel = True
        'Modify By Cheng 2003/01/21
'         strMsg = "公告日必須小於系統日"
         strMsg = "公告日不能大於系統日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text03_GotFocus
      End If
   End If
EXITSUB:
End Sub

'910709 Sieg 412
Private Sub text04_LostFocus()
   If text04 <> "" Then
      If Not ChkText04 Then text04.SetFocus
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
   If Val(text04) + 62 <> Val(strTmp) Then
      MsgBox "公告日期與公報卷期不符，請重新輸入 !", vbCritical
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
   i = (j - 1) * 3
   j = Val(Right(text03, 2))
   If j >= 1 And j < 11 Then
      i = i + 1
   ElseIf j >= 11 And j < 21 Then
      i = i + 2
   ElseIf j >= 21 Then
      i = i + 3
   End If
   
   If Val(text05) <> i Then
      MsgBox "公告日期與公報卷期不符，請重新輸入 !", vbCritical
      ChkText05 = False
   End If
End Function

'910709 Sieg 412
Private Sub text05_LostFocus()
   If text05 <> "" Then
      If Not ChkText05 Then text05.SetFocus
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
        'Add By Sindy 2011/12/15
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
   
    'Add By Cheng 2003/01/14
    '欄位資料檢查有誤
    Cancel = False
    '若有輸入代理人代號
    If IsEmptyText(text07_1) = False Then
       If UpdateCtrlData(2) = False Then
            Cancel = True
            strMsg = "無此代理人資料"
            strTit = "錯誤"
            nResponse = MsgBox(strMsg, vbOKOnly, strTit)
            text07_1_GotFocus
            'Add By Cheng 2003/01/13
            '欄位資料檢查有誤
            Cancel = True
       End If
'       text07_2.Locked = True
'       text07_2.TabStop = False
    '若沒輸入代理人代號
    Else
'       text07_2.Locked = False
'       text07_2.TabStop = True
       text08 = Empty
       text07_2 = Empty
'       text07_2.SetFocus
    End If
End Sub

Public Sub UpdateData()
   '2005/5/19 ADD BY SONIA
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   '2005/5/19 END
   ' 先清除欄位內容
   Clear
   ' 更新 Caption
   Dim strCap As String
   Dim strTmp As String
   strCap = "專利公報資料維護"
   
   Command1(0).Visible = True
   Command1(1).Visible = True
   
   ' 更新第一個欄位
   text01 = m_DataKey
   
   Select Case m_EditMode
      Case 0:
         strTmp = " -- 新增"
         'Add by Morgan 2004/8/2
         '新型是否修正預設Y, 證書號預設 I M D
         m_strPA11 = m_DataKey
         If m_EditMode = 0 Then
            'Modify by Morgan 2010/12/28 申請案號改碼數
            'Select Case Mid(m_strPA11, 3, 1)
            Select Case Mid(m_strPA11, 4, 1)
               '發明
               Case "1"
                  If Left(m_CurrTPB02, 1) <> "I" Then m_CurrTPB02 = "I"
                  txtTPB09.Text = ""
               '新型
               Case "2"
                  If Left(m_CurrTPB02, 1) <> "M" Then m_CurrTPB02 = "M"
                  txtTPB09.Text = "N"
               '設計
               Case "3"
                  If Left(m_CurrTPB02, 1) <> "D" Then m_CurrTPB02 = "D"
                  txtTPB09.Text = ""
            End Select
            
         End If
   
         text02 = m_CurrTPB02
         text03 = m_CurrTPB03
         Command1(0).Visible = False
         Command1(1).Visible = False
         '2005/5/19 ADD BY SONIA
         'Modify by Morgan 2010/12/28 申請案號改碼數
         'If Len(m_DataKey) > 8 Then
         '   StrSQLa = "Select TPB02 From TPBulletin Where TPB01='" & ChgSQL(Mid(m_DataKey, 1, 8)) & "'"
         If Len(m_DataKey) > 9 Then
            StrSQLa = "Select TPB02 From TPBulletin Where TPB01='" & ChgSQL(Mid(m_DataKey, 1, 9)) & "'"
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                m_CurrTPB02 = rsA.Fields(0)
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
         End If
         '2005/5/19 END
      Case 1:
         'Add by Sindy 2017/9/21
         '新型是否修正預設Y, 證書號預設 I M D
         m_strPA11 = m_DataKey
         
         strTmp = " -- 修改"
         text02 = ""
         text03 = ""
      Case 2:
         strTmp = " -- 查詢"
      Case 3:
         strTmp = " -- 刪除"
   End Select
   Caption = strCap & strTmp
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
   Dim stPA22 As String '母案證書號
   Dim i As Integer, Cancel As Boolean
   
   CheckDataValid = False
   
   'Added by Morgan 2021/12/22 檢查畫面輸入欄位是否含有Unicode文字
   If PUB_ChkUniText(Me, , True, "TextBox") = False Then
       Exit Function
   End If
   'end 2021/12/22
   
   '若為修改申請案號
   If m_EditMode = 1 And Me.Text1.Visible = True Then
      '若未輸入新申請案號
      If Me.Text1.Text = "" Then
         strMsg = "請輸入新申請案號"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         Text1_GotFocus
         GoTo EXITSUB
      End If
      '若有輸入新的申請案號
      If Me.Text1.Text <> "" Then
           StrSQLa = "Select * From TPBulletin Where TPB01='" & ChgSQL(Me.Text1.Text) & "'"
           rsA.CursorLocation = adUseClient
           rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
           If rsA.RecordCount > 0 Then
               MsgBox "您輸入的新申請案號已存在, 請重新輸入!!!", vbExclamation + vbOKOnly
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
      strMsg = "請輸入公告號/證書號"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text02.SetFocus
      GoTo EXITSUB
      
   ElseIf Val(text03) >= 930801 Then
      If Len(text02.Text) <> 7 Then
         strMsg = "證書號必須為7碼"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text02.SetFocus
         GoTo EXITSUB
         
      '聯合案
      'Modify by Morgan 2010/12/28 申請案號改碼數
      'ElseIf Len(m_strPA11) > 8 Then
      ElseIf Len(m_strPA11) > 9 Then
         '若為本所案件需檢查與母案相同
         If text09.Text <> "" Then
            If IsPA22Ok(m_strPA11, text02.Text, stPA22) = False Then
               strTit = "證書號"
               strMsg = "證書號與母案證書號【" & stPA22 & "】不同，是否要繼續？"
               If MsgBox(strMsg, vbInformation + vbYesNo + vbDefaultButton2, strTit) = vbNo Then
                  text02.SetFocus
                  GoTo EXITSUB
               End If
            End If
         End If
         
      ElseIf IsTPB02Exist(text02) = True Then
         strTit = "證書號"
         strMsg = "證書號已存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text02.SetFocus
         GoTo EXITSUB
      End If
      
   Else
      nLen = Len(text02)
      text02 = String(6 - nLen, "0") & text02
      If IsTPB02Exist(text02) = True Then
         strTit = "公告號"
         strMsg = "公告號已存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text02.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   If IsEmptyText(text03) = True Then
      strMsg = "請輸入公告日"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text03.SetFocus
      GoTo EXITSUB
   Else
      If CheckIsTaiwanDate(text03, False) = False Then
         strMsg = "請輸入正確的公告日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text03.SetFocus
         GoTo EXITSUB
      End If
      '公告日不能大於系統日
      If DBDATE(text03) > DBDATE(SystemDate()) Then
         strMsg = "公告日不能大於系統日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text03.SetFocus
         GoTo EXITSUB
      End If
      If Val(pa(14)) > 0 Then
         If Val(text03) + 19110000 <> pa(14) Then
            MsgBox "公告日與第一次輸入【" & ChangeTStringToTDateString(Format(Val(pa(14)) - 19110000)) & "】不同，不可存檔！", vbCritical
            text03.SetFocus
            GoTo EXITSUB
         End If
      Else
         'Add by Morgan 2006/10/13 公告日與申請延緩公告的日期不同時提醒
         If m_bol412 = True Then
            If Val(text03) + 19110000 <> Val(m_strPA14) Then
               MsgBox "公告日與延緩公告日【" & ChangeTStringToTDateString(Format(Val(m_strPA14) - 19110000)) & "】不同！", vbCritical
            End If
         End If
         'end 2006/10/13
      End If
      '有發證日才檢查
      If Val(text03) >= 930801 And pa(22) <> "" And pa(21) <> "" Then
         If text02.Text <> pa(22) Then
            MsgBox "證書號與第一次輸入【" & pa(22) & "】不同，不可存檔！", vbCritical
            text02.SetFocus
            GoTo EXITSUB
         End If
      End If
   End If
   
   If IsEmptyText(text04) = True Then
      strMsg = "請輸入公告卷期"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text04.SetFocus
      GoTo EXITSUB
      
   ElseIf Not ChkText04 Then
      text04.SetFocus
      GoTo EXITSUB
   End If
   
   If IsEmptyText(text05) = True Then
      strMsg = "請輸入公告卷期"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text05.SetFocus
      GoTo EXITSUB
      
   ElseIf Not ChkText05 Then
      text05.SetFocus
      GoTo EXITSUB
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
   
   If IsEmptyText(text06_1) = True Then
      strMsg = "請輸入申請人國籍"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text06_1.SetFocus
      GoTo EXITSUB
      
   ElseIf Me.text06_1.Text = "000" Then
      strMsg = "申請人國籍不可輸入 000 !!!"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text06_1.SetFocus
      GoTo EXITSUB
      
   'Add By Sindy 2011/12/15
   ElseIf Me.text06_1.Text = "003" Then
      strMsg = "申請人國籍不可輸入 003 !!!"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text06_1.SetFocus
      GoTo EXITSUB
      
   ElseIf UpdateCtrlData(1) = False Then
      strMsg = "申請人國籍不正確"
      strTit = "資料檢核"
      nResponse = MsgBox(strMsg, vbOKOnly, strTit)
      text06_1.SetFocus
      GoTo EXITSUB
   End If
   'Modify by Morgan 2007/12/4
   'If PUB_ChkCPExist(pa, 自請撤回, 2) = True Then
   If Check413 = True Then
   'end 2007/12/4
      MsgBox "本案已申請自撤，應不予公告，請查明！"
      GoTo EXITSUB
   End If
   
   m_blnUpdateCP22 = False
   '若執行新增功能
   If m_EditMode = 0 Then
      strExc(0) = "SELECT PA16,PA20,PA14 FROM PATENT WHERE PA11='" & Me.text01.Text & "' AND PA09='000'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Dim stPA14 As String
         stPA14 = "" & RsTemp.Fields("PA14")
         If stPA14 <> "" Then
            If Val(stPA14) <> Val(text03) + 19110000 Then
               
            End If
         End If
         
         If (RsTemp.Fields("PA16") <> "1" Or IsNull(RsTemp.Fields("PA16"))) Or IsNull(RsTemp.Fields("PA20")) Then
            If MsgBox("此申請案為本所案件，但基本檔未輸入核准資料，您是否確定新增此筆資料 ?", vbCritical + vbOKCancel) = vbCancel Then
               GoTo EXITSUB
               
            '若有輸入代理人
            ElseIf Me.text07_1.Text <> "" Or Me.text07_2.Text <> "" Then
               '不更新是否出名欄(CP22)
               m_blnUpdateCP22 = False
               
            '若無輸入代理人
            Else
               '更新是否出名欄(CP22)
               m_blnUpdateCP22 = True
            End If
         End If
      End If
   End If
   
   ' 代理人
   If IsEmptyText(text07_1) = False Then
      If UpdateCtrlData(2) = False Then
         strMsg = "無此代理人資料"
         strTit = "錯誤"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         text07_1.SetFocus
         GoTo EXITSUB
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
         'Modif by Morgan 2011/5/12
         'strFreeAgentCode = GetFreeAgentCode
         strFreeAgentCode = PUB_GetFreeAgentCode("P")
         strMsg = "確定要新增代理人編號 <" & strFreeAgentCode & "> " & Chr(10) & Chr(13) & _
                        "　　　　　代理人名稱 <" & text07_2 & "> "
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton1, strTit)
         If nResponse = vbYes Then
            strSql = "INSERT INTO TAgent (TA01, TA02, TA03, TA04) VALUES ('P','" & strFreeAgentCode & "','" & Trim(text07_2) & "','" & Trim(text07_2) & "')"
            cnnConnection.Execute strSql
            text07_1 = strFreeAgentCode
            Me.text08.Text = "" & Me.text07_2.Text
            ' 儲存公告日
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
         GoTo EXITSUB
      End If
   End If
   
   '若為新型時是否修正攔一定要輸
   'Modify by Morgan 2010/12/28 申請案號改碼數
   'If Mid(m_strPA11, 3, 1) = "2" Then
   If Mid(m_strPA11, 4, 1) = "2" Then
      If txtTPB09.Text = "" Then
         MsgBox "新型案是否修正攔不可空白！", vbCritical
         txtTPB09.SetFocus
         GoTo EXITSUB
      End If
   End If
   
   'Add By Sindy 2013/10/29 IPC分類資料開始於101年01月
   If Val(DBDATE(text03)) >= 20120101 Then
      If IsEmptyText(txtTPB10) = True Then
         strMsg = "請輸入國際分類號"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPB10.SetFocus
         txtTPB10_GotFocus
         GoTo EXITSUB
      End If
      'Modify By Sindy 2025/10/21 mark,已經沒有在維護IPC分類,所以不鎖一定要輸入
'      If IsEmptyText(txtTPB11) = True Then
'         strMsg = "請輸入IPC分類"
'         strTit = "資料檢核"
'         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
'         txtTPB11.SetFocus
'         txtTPB11_GotFocus
'         GoTo EXITSUB
'      End If
      If IsEmptyText(txtTPB12) = True Then
         strMsg = "請輸入產業別分類 "
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPB12.SetFocus
         txtTPB12_GotFocus
         GoTo EXITSUB
      End If
   End If
   '2013/10/29 END
   
   'Add By Sindy 2017/2/21
   For i = 0 To 9
      Cancel = False
      Call txtcAppl_Validate(i, Cancel)
      If Cancel = True Then
         txtcAppl(i).SetFocus
         txtcAppl_GotFocus (i)
         GoTo EXITSUB
      End If
   Next i
   '2017/2/21 END
   
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
   txtTPB09.Text = ""
   'Add By Sindy 2013/10/29
   txtTPB10.Text = ""
   txtTPB11.Text = ""
   txtTPB12.Text = ""
   '2013/10/29 END
   'Add By Sindy 2017/2/21
   txtTPB13.Text = ""
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
   '2017/2/21 END
   'Add By Sindy 2018/11/12
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
   txtTPB34.Text = ""
   txtTPB35.Text = ""
   txtTPB36.Text = ""
   txtTPB37.Text = ""
   '2018/11/12 END
   Combo1.Text = "": textNA01.Text = "" 'Add By Sindy 2019/9/4
End Sub

Private Sub SetPrevState()
   text02 = m_CurrTPB02
   text03 = m_CurrTPB03
   text04 = m_CurrTPB04
   text05 = m_CurrTPB05
   'Modify by Morgan 2004/8/3
   'If text02 = Empty Then
   If Len(text02) <> 7 Then
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
        'Add By Cheng 2002/11/22
        On Error GoTo ErrorHandler
        cnnConnection.BeginTrans
         strTit = "代理人"
        'Modify By Cheng 2002/11/22
'         strMsg = "確定要新增代理人 <" & text07_2 & ">"
        'Add By Cheng 2002/11/22
        'Modif by Morgan 2011/5/12
         'strFreeAgentCode = GetFreeAgentCode
         strFreeAgentCode = PUB_GetFreeAgentCode("P")
         
         strMsg = "確定要新增代理人編號 <" & strFreeAgentCode & "> " & Chr(10) & Chr(13) & _
                        "　　　　　代理人名稱 <" & text07_2 & "> "
        'Modify By Cheng 2002/11/22
'         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton2, strTit)
         nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbDefaultButton1, strTit)
         If nResponse = vbYes Then
            'Modify By Cheng 2002/11/22
'            strFreeAgentCode = GetFreeAgentCode
            '91.5.23 modify by sonia
            strSql = "INSERT INTO TAgent (TA01, TA02, TA03, TA04) VALUES ('P','" & strFreeAgentCode & "','" & Trim(text07_2) & "','" & Trim(text07_2) & "')"
            'strSQL = "INSERT INTO TAgent (TA01, TA02, TA03) VALUES ('P','" & strFreeAgentCode & "','" & text07_2 & "')"
            '91.5.23 end
            cnnConnection.Execute strSql
            text07_1 = strFreeAgentCode
            'Add By Cheng 2002/11/11
            Me.text08.Text = "" & Me.text07_2.Text
            ' 儲存公告日
            If IsEmpty(text03) = False Then
               strSql = "UPDATE TAgent SET TA05 = " & DBDATE(text03) & " " & _
                        "WHERE TA01 = 'P' AND " & _
                              "TA02 = '" & strFreeAgentCode & "' "
               cnnConnection.Execute strSql
            End If
            'Add By Cheng 2002/11/22
            cnnConnection.CommitTrans
         'Add By Cheng 2003/02/24
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
    If Cancel = True Then text07_2_GotFocus
'Add By Cheng 2002/11/22
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
   'Add by Morgan 2004/8/2
   'edit by nickc 2007/07/11 切換輸入法改用API
   'text02.IMEMode = 2
   CloseIme
   
    'Modify By Cheng 2002/11/19
    If m_EditMode = 1 Then
        'Modify By Cheng 2002/11/21
        '最後一個字反白
        If Me.text02.Text <> "" Then
            Me.text02.SelStart = Len(Me.text02.Text) - 1
            Me.text02.SelLength = 1
        End If
    ElseIf m_EditMode = 0 Then
      text02.SelStart = Len(text02)
      text02.SelLength = 0
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
    'Add By Cheng 2003/01/16
    TextInverse Me.Text1
End Sub
'Add by Morgan 2010/12/28
Private Sub Text1_KeyPress(KeyAscii As Integer)
   If KeyAscii = Asc(" ") Or KeyAscii = Asc("-") Then
      KeyAscii = 0
      Beep
   End If
End Sub

'Add by Morgan 2004/7/29
'新型是否修正欄預設Y
Private Sub Text1_LostFocus()
   If Text1.Text <> "" Then
      m_strPA11 = Text1.Text
   Else
      m_strPA11 = text01.Text
   End If
   'Modify by Morgan 2010/12/28 申請案號改碼數
   'If Mid(m_strPA11, 3, 1) = "2" Then
   If Mid(m_strPA11, 4, 1) = "2" Then
      txtTPB09.Text = "N"
   Else
      txtTPB09.Text = ""
   End If
End Sub

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
   
   If m_EditMode = 1 Or m_EditMode = 2 Then
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
                  Combo1 = Empty
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
   End If
End Sub

'Add By Sindy 2017/2/21
Private Sub txtcAppl_GotFocus(Index As Integer)
   '切換輸入法改用API
   OpenIme
   TextInverse txtcAppl(Index)
End Sub

'Add By Sindy 2017/2/21
Private Sub txtcAppl_Validate(Index As Integer, Cancel As Boolean)
   '若不是修改狀態，將會出不去
   If m_EditMode <> 0 And m_EditMode <> 1 Then Exit Sub
   If txtcAppl(Index).Text = "" Then Exit Sub
    '中日文欄位尾碼加空白，最大可輸長度減一(因可能會有造字無法存入問題)
   If Not CheckLengthIsOK(txtcAppl(Index), txtcAppl(Index).MaxLength - 1) Then
      Cancel = True
   End If
End Sub

'Add By Sindy 2018/11/12
Private Sub txteAppl_GotFocus(Index As Integer)
   InverseTextBox txteAppl(Index)
End Sub

'Add by Morgan 2004/7/29
Private Sub txtTPB09_GotFocus()
   TextInverse txtTPB09
   'edit by nickc 2007/07/11 切換輸入法改用API
   'txtTPB09.IMEMode = 2
   CloseIme
End Sub
'Add by Morgan 2004/7/29
'只能輸入 Y N
Private Sub txtTPB09_KeyPress(KeyAscii As Integer)
   
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 Then
      '非新型不輸
      'Modify by Morgan 2010/12/28 申請案號改碼數
      'If Mid(m_strPA11, 3, 1) <> "2" Then
      If Mid(m_strPA11, 4, 1) <> "2" Then
         Beep
         KeyAscii = 0
      '只可輸Y,N
      ElseIf KeyAscii <> 78 And KeyAscii <> 89 Then
         Beep
         KeyAscii = 0
      End If
   End If
   
End Sub

Private Function IsPA22Ok(ByVal stPA11 As String, ByVal stPA22 As String, ByRef stMomPA22 As String) As Boolean

On Error GoTo ErrHnd

   'Modify by Morgan 2010/12/28 申請案號改碼數
   'strSql = "Select PA22 FROM PATENT where PA11='" & Left(stPA11, 8) & "' AND PA01='P' AND PA09='000' AND PA23='1' AND PA22 IS NOT NULL"
   strSql = "Select PA22 FROM PATENT where PA11='" & Left(stPA11, 9) & "' AND PA01='P' AND PA09='000' AND PA23='1' AND PA22 IS NOT NULL"
   
   CheckOC
   adoRecordset.CursorLocation = adUseClient
   adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If adoRecordset.RecordCount > 0 Then
      stMomPA22 = "" & adoRecordset.Fields("PA22")
      '若母案證書號為數字則只比較數字部分
      If IsNumeric(stMomPA22) Then stPA22 = Mid(stPA22, 2)
      If stPA22 = stMomPA22 Then
         IsPA22Ok = True
      End If
   Else
      stMomPA22 = ""
   End If
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbCritical
   CheckOC

End Function
'Add by Morgan 2007/12/4 檢查有發文申請程序的自請撤回
Private Function Check413() As Boolean
   strExc(0) = "select 1 from caseprogress a where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "' and cp10='413' and cp27>0 and cp57 is null" & _
      " and exists(select * from caseprogress b where b.cp09=a.cp43 and instr('101,102,103,104,105,107,301,302,303,304,305,306,307',b.cp10)>0)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Check413 = True
   End If
End Function

'Add By Sindy 2013/10/29
Private Sub txtTPB10_GotFocus()
   TextInverse Me.txtTPB10
End Sub
Private Sub txtTPB10_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub txtTPB11_GotFocus()
   TextInverse Me.txtTPB11
End Sub
Private Sub txtTPB12_GotFocus()
   TextInverse Me.txtTPB12
End Sub
Private Sub txtTPB13_GotFocus()
   TextInverse Me.txtTPB13
End Sub

'Add By Sindy 2018/11/12
Private Sub txtTPB34_GotFocus()
   InverseAll txtTPB34
End Sub
Private Sub txtTPB34_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(txtTPB34) = False Then
      If CheckIsTaiwanDate(txtTPB34, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的申請日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPB34_GotFocus
         GoTo EXITSUB
      End If
      If DBDATE(txtTPB34) > DBDATE(SystemDate()) Then
         Cancel = True
         strMsg = "申請日不能大於系統日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPB34_GotFocus
      End If
   End If
EXITSUB:
End Sub
Private Sub txtTPB35_GotFocus()
   InverseAll txtTPB35
End Sub
Private Sub txtTPB35_Validate(Cancel As Boolean)
   Dim strMsg As String
   Dim strTit As String
   Dim nResponse
   Cancel = False
   If IsEmptyText(txtTPB35) = False Then
      If CheckIsTaiwanDate(txtTPB35, False) = False Then
         Cancel = True
         strMsg = "請輸入正確的最早優先權日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPB35_GotFocus
         GoTo EXITSUB
      End If
      If DBDATE(txtTPB35) > DBDATE(SystemDate()) Then
         Cancel = True
         strMsg = "最早優先權日不能大於系統日"
         strTit = "資料檢核"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         txtTPB35_GotFocus
      End If
   End If
EXITSUB:
End Sub
'2018/11/12 END
