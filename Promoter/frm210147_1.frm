VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210147_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "簽核狀況查詢"
   ClientHeight    =   5784
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5784
   ScaleWidth      =   8952
   Tag             =   "加班資料"
   Begin VB.CommandButton cmdFile 
      Caption         =   "檢視回覆單"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4950
      TabIndex        =   0
      Top             =   30
      Width           =   1065
   End
   Begin VB.TextBox txtF0301 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   100
      Width           =   945
   End
   Begin VB.TextBox txtF0309 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   4860
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   400
      Width           =   1845
   End
   Begin VB.TextBox txtF0310 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   400
      Width           =   645
   End
   Begin VB.CommandButton cmdDel 
      Caption         =   "刪除(&D)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   7290
      TabIndex        =   2
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "修改(&M)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   6450
      TabIndex        =   1
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8130
      TabIndex        =   3
      Top             =   30
      Width           =   800
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm210147_1.frx":0000
      Height          =   1190
      Left            =   4716
      TabIndex        =   4
      Top             =   4400
      Width           =   4212
      _ExtentX        =   7430
      _ExtentY        =   2096
      _Version        =   393216
      FixedCols       =   0
      ScrollTrack     =   -1  'True
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
   Begin TabDlg.SSTab SSTab1 
      Height          =   3700
      Left            =   0
      TabIndex        =   14
      Top             =   665
      Width           =   8904
      _ExtentX        =   15706
      _ExtentY        =   6519
      _Version        =   393216
      Tab             =   1
      TabHeight       =   420
      TabCaption(0)   =   "結案單"
      TabPicture(0)   =   "frm210147_1.frx":0015
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame10"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "txtF0305"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "txtNP09"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "txtNP08"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "txtNP07"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "txtApplNo"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "txtCaseNo"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtAg"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "Label1(37)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtNation"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtCaseName"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "txtApplPerson"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "txtNP14"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "txtNP10"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "txtNP15"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtF0306"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "Label1(14)"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "Label1(13)"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "Label1(12)"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "Label1(11)"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label1(10)"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label1(9)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "Label1(8)"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "Label1(7)"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "Label1(6)"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "Label1(5)"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "Label1(4)"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "Label1(3)"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "Label1(2)"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "請款項目"
      TabPicture(1)   =   "frm210147_1.frx":0031
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(21)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(18)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(32)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(31)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(20)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "Label1(19)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblName(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblName(1)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblName(2)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblName(3)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(26)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(36)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "FrameCRC"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtA1K(1)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtA1K(0)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtA1K(3)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtA1K(4)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtA1K(5)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtA1K(6)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtA1K(2)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmdInfo"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "聯絡項目"
      TabPicture(2)   =   "frm210147_1.frx":004D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "Frame6"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "Frame7"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "txtFCPMemo"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "Label1(17)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).ControlCount=   5
      Begin VB.Frame Frame10 
         BorderStyle     =   0  '沒有框線
         Height          =   280
         Left            =   -69000
         TabIndex        =   108
         Top             =   360
         Width           =   1092
         Begin VB.CheckBox ChkClose 
            Caption         =   "閉卷"
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
            Height          =   255
            Left            =   120
            TabIndex        =   109
            Top             =   0
            Width           =   950
         End
      End
      Begin VB.CommandButton cmdInfo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "目前設定"
         Height          =   550
         Left            =   8250
         Style           =   1  '圖片外觀
         TabIndex        =   107
         Top             =   300
         Width           =   550
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   2
         Left            =   6984
         MaxLength       =   1
         TabIndex        =   88
         Top             =   372
         Width           =   255
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   6
         Left            =   6000
         MaxLength       =   9
         TabIndex        =   87
         Top             =   948
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   5
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   86
         Text            =   "txtA1K27"
         Top             =   948
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   4
         Left            =   6000
         MaxLength       =   9
         TabIndex        =   85
         Top             =   636
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   3
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   84
         Text            =   "txtA1K28"
         Top             =   636
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         BorderStyle     =   0  '沒有框線
         Height          =   270
         Index           =   0
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   83
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   1
         Left            =   3504
         MaxLength       =   1
         TabIndex        =   82
         Top             =   372
         Width           =   255
      End
      Begin VB.Frame Frame8 
         Height          =   1400
         Left            =   -69050
         TabIndex        =   76
         Top             =   250
         Width           =   2900
         Begin VB.CheckBox Chk9 
            Caption         =   "閉卷請款"
            Height          =   255
            Index           =   22
            Left            =   220
            TabIndex        =   80
            Top             =   700
            Width           =   2200
         End
         Begin VB.CheckBox Chk9 
            Caption         =   "請銷本所年費期限管制"
            Height          =   255
            Index           =   21
            Left            =   220
            TabIndex        =   79
            Top             =   400
            Width           =   2200
         End
         Begin VB.CheckBox Chk9 
            Caption         =   "不請款"
            Height          =   255
            Index           =   23
            Left            =   220
            TabIndex        =   78
            Top             =   1000
            Width           =   2200
         End
         Begin VB.CheckBox Chk7 
            Caption         =   "913 閉卷"
            Height          =   255
            Index           =   1
            Left            =   50
            TabIndex        =   77
            Top             =   120
            Width           =   1050
         End
         Begin MSForms.ComboBox CboState 
            Height          =   288
            Index           =   1
            Left            =   1150
            TabIndex        =   81
            Top             =   120
            Width           =   1596
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "2822;508"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3300
         Left            =   -74940
         TabIndex        =   69
         Top             =   250
         Width           =   2320
         Begin VB.TextBox txtCCD08 
            Height          =   270
            Left            =   1300
            MaxLength       =   8
            TabIndex        =   103
            Text            =   "txtCCD08"
            Top             =   960
            Width           =   850
         End
         Begin VB.CheckBox Chk6 
            Caption         =   "未付帳款"
            Height          =   255
            Index           =   3
            Left            =   50
            TabIndex        =   72
            Top             =   700
            Width           =   2200
         End
         Begin VB.CheckBox Chk6 
            Caption         =   "查本案前款均已付清"
            Height          =   255
            Index           =   2
            Left            =   50
            TabIndex        =   71
            Top             =   400
            Width           =   2200
         End
         Begin VB.CheckBox Chk6 
            Caption         =   "D/N run C 類工程師報告"
            Height          =   255
            Index           =   1
            Left            =   50
            TabIndex        =   70
            Top             =   100
            Width           =   2200
         End
         Begin VB.Label Label1 
            Caption         =   "該案未請款:"
            Height          =   252
            Index           =   38
            Left            =   60
            TabIndex        =   106
            Top             =   2520
            Width           =   1236
         End
         Begin VB.Label lblNotPay_CP 
            Caption         =   "該案未請款:"
            ForeColor       =   &H000000FF&
            Height          =   432
            Left            =   120
            TabIndex        =   105
            Top             =   2760
            Width           =   2100
         End
         Begin VB.Label lblNotPayCPN 
            Caption         =   "     管制催款日"
            Height          =   252
            Left            =   60
            TabIndex        =   104
            Top             =   960
            Width           =   1236
         End
         Begin MSForms.TextBox txtNotPay 
            Height          =   1300
            Left            =   48
            TabIndex        =   73
            Top             =   1230
            Width           =   2200
            VariousPropertyBits=   -1466941413
            ScrollBars      =   2
            Size            =   "3881;2293"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Frame Frame7 
         Height          =   3300
         Left            =   -72600
         TabIndex        =   57
         Top             =   250
         Width           =   3500
         Begin VB.CheckBox Chk8 
            Caption         =   "XXX"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   102
            Top             =   2760
            Visible         =   0   'False
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "需管制6個月補繳期"
            Height          =   255
            Index           =   19
            Left            =   96
            TabIndex        =   101
            Top             =   2520
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "期限屆,未獲指示"
            Height          =   255
            Index           =   11
            Left            =   96
            TabIndex        =   65
            Top             =   400
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "代理人指示不繳年費"
            Height          =   255
            Index           =   12
            Left            =   96
            TabIndex        =   64
            Top             =   650
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "A.未獲指示"
            Height          =   255
            Index           =   14
            Left            =   250
            TabIndex        =   63
            Top             =   1450
            Width           =   1200
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "透過其他管道代繳年費"
            Height          =   255
            Index           =   13
            Left            =   96
            TabIndex        =   62
            Top             =   900
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "B.已獲指示"
            Height          =   255
            Index           =   15
            Left            =   1620
            TabIndex        =   61
            Top             =   1450
            Width           =   1200
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "後續准駁簡單報告"
            Height          =   255
            Index           =   16
            Left            =   250
            TabIndex        =   60
            Top             =   2016
            Width           =   2000
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "駁不報告"
            Height          =   255
            Index           =   17
            Left            =   250
            TabIndex        =   59
            Top             =   2256
            Width           =   1200
         End
         Begin VB.CheckBox Chk7 
            Caption         =   "907 不續辦"
            Height          =   255
            Index           =   0
            Left            =   50
            TabIndex        =   58
            Top             =   120
            Width           =   1150
         End
         Begin VB.Label Label1 
            Caption         =   "後續准駁簡單報告："
            Height          =   252
            Index           =   15
            Left            =   96
            TabIndex        =   68
            Top             =   1750
            Width           =   1800
         End
         Begin VB.Label Label1 
            Caption         =   "不續辦請款："
            Height          =   252
            Index           =   16
            Left            =   96
            TabIndex        =   67
            Top             =   1200
            Width           =   1104
         End
         Begin MSForms.ComboBox CboState 
            Height          =   288
            Index           =   0
            Left            =   1250
            TabIndex        =   66
            Top             =   120
            Width           =   2180
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "3845;508"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Frame FrameCRC 
         BackColor       =   &H00C0FFFF&
         Caption         =   "請款項目"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000040C0&
         Height          =   2450
         Left            =   120
         TabIndex        =   42
         Top             =   1200
         Width           =   8712
         Begin VB.TextBox txtAmt 
            Alignment       =   1  '靠右對齊
            Height          =   270
            Index           =   0
            Left            =   1512
            TabIndex        =   45
            Top             =   200
            Width           =   1200
         End
         Begin VB.TextBox txtAmt 
            Alignment       =   1  '靠右對齊
            Height          =   270
            Index           =   1
            Left            =   3840
            MaxLength       =   6
            TabIndex        =   44
            Top             =   200
            Width           =   1200
         End
         Begin VB.TextBox txtAmt 
            Alignment       =   1  '靠右對齊
            Height          =   270
            Index           =   2
            Left            =   6120
            MaxLength       =   6
            TabIndex        =   43
            Top             =   200
            Width           =   1200
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridAMT 
            Height          =   1504
            Left            =   120
            TabIndex        =   56
            Top             =   564
            Width           =   8196
            _ExtentX        =   14457
            _ExtentY        =   2646
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            FormatString    =   "順序|代號|請款項目|金額|折扣|備註"
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   90
            X2              =   8300
            Y1              =   500
            Y2              =   500
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "總金額："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   22
            Left            =   720
            TabIndex        =   48
            Top             =   204
            Width           =   804
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "規費："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   24
            Left            =   3240
            TabIndex        =   47
            Top             =   204
            Width           =   600
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "點數："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   25
            Left            =   5520
            TabIndex        =   46
            Top             =   204
            Width           =   600
         End
      End
      Begin VB.TextBox txtF0305 
         BorderStyle     =   0  '沒有框線
         Height          =   285
         Left            =   -73950
         Locked          =   -1  'True
         TabIndex        =   21
         Text            =   "其他"
         Top             =   2240
         Width           =   7650
      End
      Begin VB.TextBox txtNP09 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -68460
         Locked          =   -1  'True
         TabIndex        =   20
         Text            =   "105/02/05"
         Top             =   1640
         Width           =   1065
      End
      Begin VB.TextBox txtNP08 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -71100
         Locked          =   -1  'True
         TabIndex        =   19
         Text            =   "105/02/05"
         Top             =   1640
         Width           =   1065
      End
      Begin VB.TextBox txtNP07 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -73950
         Locked          =   -1  'True
         TabIndex        =   18
         Text            =   "年費"
         Top             =   1640
         Width           =   1065
      End
      Begin VB.TextBox txtApplNo 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -71100
         Locked          =   -1  'True
         TabIndex        =   17
         Text            =   "201420105880.4"
         Top             =   372
         Width           =   1515
      End
      Begin VB.TextBox txtCaseNo 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -73950
         Locked          =   -1  'True
         TabIndex        =   16
         Text            =   "P-107255-0-00"
         Top             =   372
         Width           =   1275
      End
      Begin MSForms.TextBox txtAg 
         Height          =   288
         Left            =   -73944
         TabIndex        =   15
         Top             =   1320
         Width           =   7656
         VariousPropertyBits=   679495711
         Size            =   "13504;508"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   37
         Left            =   -74880
         TabIndex        =   100
         Top             =   1320
         Width           =   720
      End
      Begin MSForms.TextBox txtNation 
         Height          =   288
         Left            =   -71100
         TabIndex        =   41
         Top             =   1000
         Width           =   1608
         VariousPropertyBits=   679495711
         Size            =   "2831;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseName 
         Height          =   288
         Left            =   -73944
         TabIndex        =   40
         Top             =   670
         Width           =   7656
         VariousPropertyBits=   679495711
         Size            =   "13494;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtApplPerson 
         Height          =   288
         Left            =   -73944
         TabIndex        =   39
         Top             =   1000
         Width           =   1872
         VariousPropertyBits=   679495711
         Size            =   "3307;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtNP14 
         Height          =   288
         Left            =   -71280
         TabIndex        =   38
         Top             =   1930
         Width           =   1608
         VariousPropertyBits=   679495711
         Size            =   "2831;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtNP10 
         Height          =   288
         Left            =   -73944
         TabIndex        =   37
         Top             =   1930
         Width           =   1608
         VariousPropertyBits=   679495711
         Size            =   "2831;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtNP15 
         Height          =   516
         Left            =   -73944
         TabIndex        =   36
         Top             =   2550
         Width           =   7656
         VariousPropertyBits=   -1466939361
         ScrollBars      =   3
         Size            =   "13494;900"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtF0306 
         Height          =   480
         Left            =   -73944
         TabIndex        =   35
         Top             =   3100
         Width           =   7656
         VariousPropertyBits=   -1466939361
         ScrollBars      =   3
         Size            =   "13494;847"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "說明如下："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   14
         Left            =   -74880
         TabIndex        =   34
         Top             =   3100
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "結案記錄："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   13
         Left            =   -74880
         TabIndex        =   33
         Top             =   2240
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件備註："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   12
         Left            =   -74880
         TabIndex        =   32
         Top             =   2556
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "相關人："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   11
         Left            =   -72024
         TabIndex        =   31
         Top             =   1930
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   10
         Left            =   -74880
         TabIndex        =   30
         Top             =   1930
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   9
         Left            =   -69420
         TabIndex        =   29
         Top             =   1640
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   8
         Left            =   -72024
         TabIndex        =   28
         Top             =   1640
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "下一程序："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   7
         Left            =   -74880
         TabIndex        =   27
         Top             =   1640
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請國家："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   6
         Left            =   -72024
         TabIndex        =   26
         Top             =   1000
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   5
         Left            =   -74880
         TabIndex        =   25
         Top             =   1000
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   24
         Top             =   670
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請案號："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   -72024
         TabIndex        =   23
         Top             =   372
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   22
         Top             =   372
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "合併列印請款單：       (要印:Y)"
         Height          =   252
         Index           =   36
         Left            =   5520
         TabIndex        =   99
         Top             =   396
         Width           =   2556
      End
      Begin VB.Label Label1 
         Caption         =   "帳款已清：         (已清:Y)"
         Height          =   252
         Index           =   26
         Left            =   120
         TabIndex        =   98
         Top             =   396
         Width           =   2100
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(3)"
         Height          =   252
         Index           =   3
         Left            =   6960
         TabIndex        =   97
         Top             =   996
         Width           =   1500
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(2)"
         Height          =   252
         Index           =   2
         Left            =   2040
         TabIndex        =   96
         Top             =   996
         Width           =   2496
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(1)"
         Height          =   252
         Index           =   1
         Left            =   6960
         TabIndex        =   95
         Top             =   672
         Width           =   1500
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(0)"
         Height          =   252
         Index           =   0
         Left            =   2040
         TabIndex        =   94
         Top             =   648
         Width           =   2496
      End
      Begin VB.Label Label1 
         Caption         =   "固定列印對象："
         Height          =   252
         Index           =   19
         Left            =   4680
         TabIndex        =   93
         Top             =   996
         Width           =   1296
      End
      Begin VB.Label Label1 
         Caption         =   "請款對象："
         Height          =   252
         Index           =   20
         Left            =   120
         TabIndex        =   92
         Top             =   672
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "列印對象："
         Height          =   252
         Index           =   31
         Left            =   120
         TabIndex        =   91
         Top             =   996
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "固定請款對象："
         Height          =   252
         Index           =   32
         Left            =   4680
         TabIndex        =   90
         Top             =   672
         Width           =   1296
      End
      Begin VB.Label Label1 
         Caption         =   "列印申請人：       (要印:Y/改不印:N)"
         Height          =   252
         Index           =   18
         Left            =   2400
         TabIndex        =   89
         Top             =   396
         Width           =   3000
      End
      Begin MSForms.TextBox txtFCPMemo 
         Height          =   1450
         Left            =   -68920
         TabIndex        =   75
         Top             =   2050
         Width           =   2748
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "4847;2558"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "其他說明："
         Height          =   252
         Index           =   17
         Left            =   -69050
         TabIndex        =   74
         Top             =   1800
         Width           =   996
      End
      Begin VB.Label Label1 
         Caption         =   "其他說明："
         Height          =   252
         Index           =   21
         Left            =   5640
         TabIndex        =   54
         Top             =   1800
         Width           =   996
      End
      Begin VB.Label Label1 
         Caption         =   "固定列印對象："
         Height          =   252
         Index           =   23
         Left            =   -70320
         TabIndex        =   53
         Top             =   950
         Width           =   1296
      End
      Begin VB.Label Label1 
         Caption         =   "請款對象："
         Height          =   252
         Index           =   27
         Left            =   -74880
         TabIndex        =   52
         Top             =   636
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "列印對象："
         Height          =   252
         Index           =   28
         Left            =   -74880
         TabIndex        =   51
         Top             =   950
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "固定請款對象："
         Height          =   252
         Index           =   29
         Left            =   -70320
         TabIndex        =   50
         Top             =   636
         Width           =   1300
      End
      Begin VB.Label Label1 
         Caption         =   "列印申請人：       (要印:Y)"
         Height          =   252
         Index           =   30
         Left            =   -72000
         TabIndex        =   49
         Top             =   360
         Width           =   2496
      End
   End
   Begin VB.Label Label26 
      Caption         =   "Create ID:           Date         Time             Update ID:                Date                  Time"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   7.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Left            =   120
      TabIndex        =   55
      Top             =   5620
      Width           =   6672
   End
   Begin MSForms.TextBox txtF0310_2 
      Height          =   285
      Left            =   1650
      TabIndex        =   13
      Top             =   400
      Width           =   1605
      VariousPropertyBits=   679495711
      ScrollBars      =   3
      Size            =   "2831;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtF0407 
      Height          =   1190
      Left            =   936
      TabIndex        =   12
      Top             =   4400
      Width           =   3732
      VariousPropertyBits=   -1466939361
      ScrollBars      =   3
      Size            =   "6583;2099"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "簽核意見："
      Height          =   180
      Left            =   36
      TabIndex        =   11
      Top             =   4400
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "表單編號："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   36
      TabIndex        =   10
      Top             =   100
      Width           =   900
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "目前表單狀態："
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3555
      TabIndex        =   9
      Top             =   400
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "填表人員："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   8
      Top             =   400
      Width           =   900
   End
End
Attribute VB_Name = "frm210147_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/27 Form2.0已修改
'Create by Sindy 2015/1/9
Option Explicit
 
' 變數宣告區
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_NP01 As String
Dim m_NP22 As String

Dim dblPrevRow As Double
Dim m_PrevForm As Form '前一畫面
Dim m_PrevForm2 As Form 'Added by Morgan 2015/12/15
Dim m_F0302 As String
Dim m_F0303 As String
Dim m_F0304 As String
Dim m_F0308 As String
Dim m_F0316 As String '智權人員
Dim i As Integer
Dim m_strSaveFiles As String '附件
Dim m_strSaveFilesCP09 As String '附件總收文號
Dim m_AttachPath As String '附件暫存區
Public m_SignFlowEmp As String '簽核人員(因有可能人員休假職代代為操作)
Public m_stNP01 As String, m_stNP22 As String 'Add by Amy 2020/05/21 總收文號/下一程序序號
Public m_bolCallCloseMenu As Boolean 'Add By Sindy 2020/12/25 外部呼叫查詢:卷宗區
'Add by Amy 2025/04/10 FC結案單用
Public intFCState As Integer '0-非FC/1-商標/2-專利
Public m_NP07 As String '案件性質編號
Dim arrAMTCol() As String, arrAMTWidth() As String

Public Sub SetParent(ByRef fm As Form)
   If Not m_PrevForm Is Nothing Then Set m_PrevForm2 = m_PrevForm 'Added by Morgan 2015/12/15
   
   Set m_PrevForm = fm
   
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("簽核人員", "身份", "日期", "時間", "簽核結果", "B1104")
   arrGridHeadWidth = Array(1050, 600, 800, 800, 800, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Public Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strFlowDBType As String
Dim strTi06 As String 'Add by Amy 2020/05/21
'Add by Amy 2025/04/10 將Flow003中屬於結案單資料者拆至結案單主檔中
Dim strRCode As String, strRCodeN As String, strReason As String 'strRCode=原因代碼(原F0305)/原因代碼名稱/strReason=說明 or 備註(原F0306)
Dim strState As String '1-請款項目/2-指示程序項目
Dim strClose As String, strCCM07 As String, strNotPay As String, strMemo As String, strMsg As String, strCCM07N As String
Dim stA1K04 As String, stA1K27 As String, stA1K28 As String, stA1K29 As String
Dim strCCD08 As String, strNotPay_CP As String 'Add by Amy 2025/08/08
Dim stMergeBill As String, stCCM23 As String, stCCM24 As String, stCCM25 As String 'Add by Amy 2025/08/18

   Screen.MousePointer = vbHourglass
   Me.Enabled = False

   ClearField '清空欄位值
   'Modify by Amy 2020/05/21 +if
   '商標延展結案
   'Modify by Amy 2022/06/17 +卷宗區顯示T延展結案 txtF0301=T-開頭
   If txtF0301 = MsgText(601) Or Left(txtF0301, 2) = "T-" Then
       Me.Height = 4788 'Modify by Amy 2025/06/17 原:3165,改成頁籤後不夠高
       If Left(txtF0301, 2) = "T-" Then
            txtF0309 = "歸卷"
       Else
            txtF0309 = "處理中"
       End If
       strFlowDBType = "NP"
       m_F0303 = m_stNP01
       m_F0304 = m_stNP22
   '2022/06/17
   '有run結案流程
   Else
       strFlowDBType = GetFlowFormReadDBType(txtF0301, , m_CP01, m_CP02, m_CP03, m_CP04)
    
       '案件表單主檔
       'Modify by Amy 2025/04/10 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
      If strSrvDate(1) >= FCP結案單電子化啟用日 Then
         strSql = "select flow003.*,CloseCaseMain.*,decode(F0309," & ShowFlow表單狀態中文 & ") as F0309NM" & _
                " from flow003,CloseCaseMain" & _
                " where f0301='" & txtF0301 & "' And f0301=ccm01(+) "
      Else
       strSql = "select flow003.*,decode(F0309," & ShowFlow表單狀態中文 & ") as F0309NM" & _
                " from flow003" & _
                " where f0301='" & txtF0301 & "'"
       End If
       rsTmp.CursorLocation = adUseClient
       rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    
       If rsTmp.RecordCount > 0 Then
          txtF0301 = rsTmp.Fields("F0301")
          m_F0302 = rsTmp.Fields("F0302")
          'Modify by Amy 2025/04/10 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
         If strSrvDate(1) >= FCP結案單電子化啟用日 Then
            m_F0303 = rsTmp.Fields("ccm02")
            m_F0304 = "" & rsTmp.Fields("ccm03")
            strRCode = "" & rsTmp.Fields("ccm04")
            strReason = "" & rsTmp.Fields("ccm05")
            strClose = "" & rsTmp.Fields("ccm06") '是否閉卷
            strCCM07 = "" & rsTmp.Fields("ccm07") '狀態選項
            txtAmt(0) = "" & rsTmp.Fields("ccm08") '總金額
            txtAmt(1) = "" & rsTmp.Fields("ccm09")  '規費
            txtAmt(2) = "" & rsTmp.Fields("ccm10")  '點數
            'Add by Amy 2025/08/18 +請款項目資料
            stCCM25 = "" & rsTmp.Fields("ccm25") '帳款已清(填結案單時,帳款狀態)
            stMergeBill = "" & rsTmp.Fields("ccm19") '合併列印請款單
            stA1K04 = "" & rsTmp.Fields("ccm20") '列印申請人
            stA1K28 = "" & rsTmp.Fields("ccm21") '請款對象
            stA1K27 = "" & rsTmp.Fields("ccm22") '列印對象
            stCCM23 = "" & rsTmp.Fields("ccm23") '固定請款對象
            stCCM24 = "" & rsTmp.Fields("ccm24") '固定列印對象
            'end 2025/08/18
            If strCCM07 <> MsgText(601) Then
               strCCM07N = Pub_SetCloseCboState(Me.Name, , , " And AC02='" & strCCM07 & "'")
            End If
          Else
            m_F0303 = rsTmp.Fields("F0303")
            m_F0304 = "" & rsTmp.Fields("F0304")
            strRCode = rsTmp.Fields("F0305")
            strReason = "" & rsTmp.Fields("F0306")
          End If
          'end 2025/04/10
          
          'Modify By Sindy 2017/6/20
          m_F0316 = rsTmp.Fields("F0316") '智權人員
          txtF0310 = rsTmp.Fields("F0310"): txtF0310_2 = GetPrjSalesNM(rsTmp.Fields("F0310"))
          If m_F0316 <> txtF0310 Then
             txtF0310.ForeColor = &HFF0000
             txtF0310_2.ForeColor = &HFF0000
          Else
             txtF0310.ForeColor = &H80000008
             txtF0310_2.ForeColor = &H80000008
          End If
          '2017/6/20 END
          
          m_F0308 = rsTmp.Fields("F0308") '下一處理人員
          txtF0309 = rsTmp.Fields("F0309") & " " & rsTmp.Fields("F0309NM")
          If m_F0302 = Flow_結案單 Then
             'Mark by Amy 2025/06/02 備註過長時,Frame1鎖住,導致無法使用ScrollBar,故拿掉
'             Frame1.Top = 720
'             Frame1.Visible = True
             'end 2025/06/02
             'Modify by Amy 2025/04/10 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中,並改抓共用
'             strSql = "SELECT ROR02 FROM REASONOFRELIEF WHERE ROR01='" & rsTmp.Fields("F0305") & "'"
'             CheckOC
'             adoRecordset.CursorLocation = adUseClient
'             adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'             If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'                If IsNull(adoRecordset.Fields(0)) Then
'                   txtF0305 = rsTmp.Fields("F0305") + ""
'                Else
'                   txtF0305 = rsTmp.Fields("F0305") + "  " + adoRecordset.Fields(0)
'                End If
'             Else
'                txtF0305 = ""
'             End If
'             CheckOC
'             txtF0306 = "" & rsTmp.Fields("F0306")
            strRCodeN = strRCode
            Call Pub_SetCloseReason(intFCState, Me.Name, , strRCodeN)
            If strRCodeN = MsgText(601) Then
               txtF0305 = strRCode
            Else
               txtF0305 = strRCode & "  " & strRCodeN
            End If
            txtF0306 = strReason
            'end 2025/04/10
          End If
          Call UpdateCUID(rsTmp)
       Else
          Screen.MousePointer = vbDefault
          ShowNoData
          rsTmp.Close
          Set rsTmp = Nothing
          Exit Sub
       End If
       rsTmp.Close
   End If
   'end 2020/05/21

   '下一程序
   If strFlowDBType = "NP" Then
      'Modify by Amy 2020/05/21 +T102Inform
      strSql = "select *" & _
               " from nextprogress,casepropertymap,staff,T102Inform" & _
               " where np01='" & m_F0303 & "' and np22=" & m_F0304 & _
               " and np02=cpm01(+) and np07=cpm02(+) And np01=ti02(+) And np22=ti04(+)" & _
               " and np10=st01(+)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         m_CP01 = rsTmp.Fields("np02")
         m_CP02 = rsTmp.Fields("np03")
         m_CP03 = rsTmp.Fields("np04")
         m_CP04 = rsTmp.Fields("np05")
         txtNP07 = GetPrjState4(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04, rsTmp.Fields("np07"))
         txtNP07.Tag = rsTmp.Fields("np07") 'Add by Amy 2025/09/23
         txtNP08 = ChangeWStringToTDateString("" & rsTmp.Fields("np08"))
         txtNP09 = ChangeWStringToTDateString("" & rsTmp.Fields("np09"))
         txtNP10 = rsTmp.Fields("st02")
         txtNP14 = "" & rsTmp.Fields("np14")
         txtNP15 = "" & rsTmp.Fields("np15")
         'Add by Amy 2020/05/21 商標延展結案, 結案記錄抓ti05
         'Modofy by Amy 2022/06/17 +卷宗區顯示T延展結案 txtF0301=T-開頭
         If txtF0301 = MsgText(601) Or Left(txtF0301, 2) = "T-" Then
            txtF0305 = "" & rsTmp.Fields("ti05")
            txtF0310 = "" & rsTmp.Fields("ti03")
            txtF0310_2 = GetPrjSalesNM(txtF0310)
            strTi06 = "" & rsTmp.Fields("ti06")
         End If
      End If
      rsTmp.Close
   End If
   'Add by Amy 2025/04/10 +FC結案單明細
   If strSrvDate(1) >= FCP結案單電子化啟用日 Then
      '外商
      If intFCState = 1 Then
         If strClose = "Y" Then ChkClose.Value = vbChecked
         'Modify by Amy 2025/08/18 原只顯示,外商結案單增加欄位記錄,以利後續外商程序操作請款單
         'Call Pub_GetCloseA1KData(1, Me.Name, m_CP01, m_CP02, m_CP03, m_CP04, stA1K29, m_NP07, stA1K04, stA1K27, stA1K28)
         txtA1K(0) = stCCM25 '帳款已清
         txtA1K(1) = stA1K04 '列印申請人
         txtA1K(2) = stMergeBill '合併列印請款單
         txtA1K(3) = stA1K28: Call txtA1K_Validate(3, False) '請款對象
         txtA1K(4) = stCCM23: Call txtA1K_Validate(4, False) '固定請款對象
         txtA1K(5) = stA1K27: Call txtA1K_Validate(5, False) '列印對象
         txtA1K(6) = stCCM24: Call txtA1K_Validate(6, False) '固定列印對象
         Call Pub_CloseShowfrm210133_INV(0, Me.Name, txtF0301, Me.cmdInfo)
         'Modify by Amy 2025/10/02 SetA1KColor 改抓共用
         'Modify by Amy 2025/10/14 拿掉 If cmdInfo.Visible = True,避免設定未改,導致請款單資料錯誤
         Call Pub_CloseSetA1KDataColor(1, Me.Name, m_CP01, m_CP02, m_CP03, m_CP04, m_NP07, Me.txtA1K, 0, 6)
         'end 2025/08/18
      End If
      strState = Pub_GetField("CloseCaseDetail", "CCD01='" & txtF0301 & "'", "Distinct ccd02")
      '請款項目
      If strState = "1" Then
         If Pub_QueryFCCloseDetail(Val(strState), Me.Name, txtF0301, m_CP01, strMsg, GridAMT) = False Then
            If strMsg <> "無資料" Then
               strMsg = "讀取FC商標[結案明細]資料有誤,通知電腦中心！" & vbCrLf & strMsg
               GoTo EXITSUB
            End If
         Else
            Call Pub_SetCloseGridAmtWidth(Me.Name, GridAMT, arrAMTCol(), arrAMTWidth)
         End If
      '聯絡項目
      ElseIf strState = "2" Then
         'Add by Amy 2025/07/29 案件備註(原:備註)顯示灰色-Anny ex:P-129167 帶太久前資料,會混淆
         Label1(12).Enabled = False
         txtNP15.Enabled = False
         'end 2025/07/29
         'Add by Amy 2025/08/13
         CboState(0) = "": CboState(1) = "": Chk7(0).Value = vbUnchecked: Chk7(1).Value = vbUnchecked
         'end 2025/08/13
         '913 閉卷
         If strClose = "Y" Then
            CboState(1) = strCCM07N
            Chk7(1).Value = vbChecked
         '907 不續辦
         Else
            CboState(0) = strCCM07N
            Chk7(0).Value = vbChecked
         End If
         'Modify by Amy 2025/08/08 +strCCD08/strNotPay_CP ,未付帳款 及 管制催款日 都有資料,於程序操作完 解除期限/閉卷 後加行事曆
         If Pub_QueryFCCloseDetail(Val(strState), Me.Name, txtF0301, m_CP01, strMsg, , Chk6, Chk8, Chk9, strNotPay, strMemo, strCCD08, strNotPay_CP) = False Then
            strMsg = "讀取FC專利[結案明細]資料有誤,通知電腦中心！" & vbCrLf & strMsg
            GoTo EXITSUB
         End If
         If strNotPay <> MsgText(601) Then Chk6(3).Value = vbChecked
         txtNotPay = strNotPay
         'Add by Amy 2025/08/08
         txtCCD08 = strCCD08
         lblNotPay_CP.Caption = strNotPay_CP
         'end 2025/08/08
         txtFCPMemo = strMemo
      End If
   End If
   'end 2025/04/10

   '基本檔
   'Modify by Amy 2018/06/19 非P 案結案電子化,加入其他基本檔
'   strSql = "select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cu04,pa09,na03,pa91" & _
'            " from patent,customer,nation" & _
'            " where pa01='" & m_CP01 & "' and pa02='" & m_CP02 & "' and pa03='" & m_CP03 & "' and pa04='" & m_CP04 & "'" & _
'            " and substr(pa26,1,8)=cu01(+) and substr(pa26,9)=cu02(+) and pa09=na01(+) "
   'Modify by Amy 2025/04/10 +代理人
   strSql = ",FCAg,NVL(FA04,Decode(FA05,null,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) as AName "
   If intFCState > 0 Then strSql = ",FCAg,Decode(FA05,null,Nvl(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65) as AName "
   strSql = "select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cu04,pa09,na03,pa91" & strSql & _
               "From Customer,Nation,Fagent," & _
             "(Select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa11,pa09,pa91,pa26,PA75 as FCAg " & _
            "From patent where pa01='" & m_CP01 & "' and pa02='" & m_CP02 & "' and pa03='" & m_CP03 & "' and pa04='" & m_CP04 & "'" & _
    "Union Select tm01 as pa01,tm02 as pa02,tm03 as pa03,tm04 as pa04,tm05 as pa05,tm06 as pa06,tm07 as pa07,tm12 as pa11,tm10 as pa09,tm58 as pa91,tm23 as pa26,TM44 as FCAg " & _
            "From TradeMark where tm01='" & m_CP01 & "' and tm02='" & m_CP02 & "' and tm03='" & m_CP03 & "' and tm04='" & m_CP04 & "' " & _
    "Union Select lc01 as pa01,lc02 as pa02,lc03 as pa03,lc04 as pa04,lc05 as pa05,lc06 as pa06,lc07 as pa07,'' as pa11,lc15 as pa09,lc27 as pa91,lc11 as pa26,LC22 as FCAg " & _
            "From LawCase where lc01='" & m_CP01 & "' and lc02='" & m_CP02 & "' and lc03='" & m_CP03 & "' and lc04='" & m_CP04 & "' " & _
    "Union Select hc01 as pa01,hc02 as pa02,hc03 as pa03,hc04 as pa04,hc06 as pa05,'' as pa06,'' as pa07,'' as pa11,'' as pa09,hc12 as pa91,hc05 as pa26,'' as FCAg " & _
            "From HireCase where hc01='" & m_CP01 & "' and hc02='" & m_CP02 & "' and hc03='" & m_CP03 & "' and hc04='" & m_CP04 & "' " & _
    "Union Select sp01 as pa01,sp02 as pa02,sp03 as pa03,sp04 as pa04,sp05 as pa05,sp06 as pa06,sp07 as pa07,sp11 as pa11,sp09 as pa09,sp18 as pa91,sp08 as pa26,SP26 as FCAg " & _
            "From ServicePractice where sp01='" & m_CP01 & "' and sp02='" & m_CP02 & "' and sp03='" & m_CP03 & "' and sp04='" & m_CP04 & "' " & _
         ") Where substr(pa26,1,8)=cu01(+) and substr(pa26,9)=cu02(+) and pa09=na01(+) And SubStr(FCAg,1,8)=FA01(+) And Decode(SubStr(FCAg,9,1),'','0',SubStr(FCAg,9,1))=FA02(+)"
   'end 2025/04/10
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      txtCaseNo = rsTmp.Fields("pa01") & "-" & rsTmp.Fields("pa02") & "-" & rsTmp.Fields("pa03") & "-" & rsTmp.Fields("pa04")
      txtApplNo = "" & rsTmp.Fields("pa11")
      txtCaseName = rsTmp.Fields("pa05") & rsTmp.Fields("pa06") & rsTmp.Fields("pa07")
      txtApplPerson = rsTmp.Fields("cu04")
      txtNation = rsTmp.Fields("na03")
      txtAg = "" & rsTmp.Fields("AName") 'Add by Amy 2025/04/10 +代理人
      
      If txtNP10 = "" Then
         txtNP10 = GetPrjSalesNM(ShowCurrCP13(m_CP01, m_CP02, m_CP03, m_CP04, rsTmp.Fields("pa09"))) '智權人員
      End If
      txtNP15 = "" & rsTmp.Fields("pa91")
   End If
   rsTmp.Close
    
   'Modify by Amy 2020/05/21 +if 有run結案流程
   'Modify by Amy 2022/06/17 + txtF0301<>T-開頭
   Label23.Visible = False: txtF0407.Visible = False
   GRD1.Visible = False
   If txtF0301 <> MsgText(601) And Left(txtF0301, 2) <> "T-" Then
       Label23.Visible = True: txtF0407.Visible = True
       GRD1.Visible = True
       
       '案件表單流程備註檔
       SetFlow004TextBox txtF0407, txtF0301
       '案件表單簽核檔
       strSql = "SELECT ST02||nvl(F0208,'') 簽核人員,decode(F0202," & ShowFlow簽核人員身份 & ") 身份,sqldateT(F0205) 日期,sqltime6(F0206) 時間,decode(F0207," & ShowFlow簽核結果 & ") 簽核結果,F0204 FROM FLOW002,Staff WHERE F0201='" & txtF0301 & "' and F0204=ST01(+) order by decode(F0205,null,2,1) asc,F0205||sqltime6(F0206) asc,F0202,F0203 asc" 'order by F0205,F0202,F0203 asc
       rsTmp.CursorLocation = adUseClient
       rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
       If rsTmp.RecordCount > 0 Then
          Set GRD1.Recordset = rsTmp
       End If
    
       '若有資料游標停在第一筆
       GRD1.Visible = False
       GRD1.col = 0
       GRD1.row = 1
    '   If rsTmp.RecordCount > 0 Then
    '      For i = 0 To GRD1.Cols - 1
    '         GRD1.col = i
    '         GRD1.CellBackColor = &HFFC0C0
    '      Next i
    '   End If
       GRD1.Visible = True
   End If
   'end 2020/05/21
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True

   'Modify by Amy 2020/05/21 商標延展結案進入,不顯示修改鈕,Ti06=Y才有刪除鈕
   If m_bolCallCloseMenu = False Then 'Add By Sindy 2020/12/25 外部呼叫查詢:卷宗區
      cmdModify.Visible = False
      If txtF0301 = MsgText(601) Then
         cmdDel.Enabled = False
         If strTi06 = "Y" Then cmdDel.Enabled = True
      '下一處理人員=自己,才可修改,刪除
      ElseIf m_F0308 = m_SignFlowEmp And m_F0316 = m_SignFlowEmp Then
         cmdModify.Visible = True
         cmdModify.Enabled = True
         cmdDel.Enabled = True
         'cmdFile.Visible = False '回覆單
      '下一處理人員<>自己,只可查看簽核資料
      Else
         cmdModify.Visible = True
         cmdModify.Enabled = False
         cmdDel.Enabled = False
      End If
      'Modify By Sindy 2019/8/1 Move至此處
      '回覆單
      'Add by Amy 2025/04/10 資料夾未建立者先建,避免資料抓不到
      If m_AttachPath = "" Then
         If Pub_SetFilePathDelTmp("Close", 1, strExc(9), m_AttachPath) = False Then
            MsgBox "附件資料夾建立失敗" & vbCrLf & _
                           strExc(9) & vbCrLf & "請洽電腦中心!"
         End If
      End If
      'end 2025/04/10
      cmdFile.Visible = False
      'Modify By Sindy 2015/5/18
      'If PUB_ChkIsReplyFile(m_CP01, m_CP02, m_CP03, m_CP04, m_strSaveFiles, txtF0301, m_strSaveFilesCP09) = True Then
      'Modify by Amy 2025/04/10 +FC結案單
      cmdFile.Caption = "檢視回覆單"
      If strSrvDate(1) >= FCP結案單電子化啟用日 And intFCState > 0 Then cmdFile.Caption = "檢視電子檔"
      'If PUB_ChkIsReplyFile(m_CP01 & m_CP02 & m_CP03 & m_CP04, m_strSaveFiles, txtF0301, m_strSaveFilesCP09, txtF0301) = True Then
      '2015/5/18 END
      If PUB_ChkIsReplyFile(m_CP01 & m_CP02 & m_CP03 & m_CP04, m_strSaveFiles, txtF0301, m_strSaveFilesCP09, txtF0301, intFCState, m_AttachPath) = True Then
      'end 2025/04/10
         If m_strSaveFiles <> "" Then
            'Modify by Amy 2025/06/27 拿掉UCase(TypeName(m_PrevForm)) <> UCase("frm040119")由指示信判發程式進入結案單,仍要可使用卷宗區鈕
            '  由指示信判發進入且於指示信畫面上開啟指示信檔案,再進入此支會出現"檔案已開啟"的訊息-Morgan 說正常不應該如此操作,且訊息無誤,可先不處理
            If UCase(TypeName(m_PrevForm)) <> UCase("frm100101_L") Then
               cmdFile.Visible = True
            End If
   '            If PUB_GetAttachFile_CPP(m_CP01 & m_CP02 & m_CP03 & m_CP04, m_strSaveFiles, m_AttachPath) = False Then
   '               MsgBox "無法儲存欲開啟的檔案[ " & m_strSaveFiles & " ]！"
   '            End If
         End If
      End If
   End If
   
EXITSUB:
   'Add by Amy 2025/04/10 避免有錯,無法離開
   If strMsg <> MsgText(601) Then
      cmdModify.Enabled = False
      cmdDel.Enabled = False
      MsgBox strMsg
   End If
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   'end 2025/04/10
   Set rsTmp = Nothing
End Sub

Private Sub ClearField()
   txtF0309 = Empty

   'Frame1:結案單
   txtCaseNo = Empty
   txtApplNo = Empty
   txtCaseName = Empty
   txtApplPerson = Empty
   txtNation = Empty
   txtNP07 = Empty
   txtNP07.Tag = Empty 'Add by Amy 2025/09/23
   txtNP08 = Empty
   txtNP09 = Empty
   txtNP10 = Empty
   txtNP14 = Empty
   txtNP15 = Empty
   txtAg = Empty 'Add by Amy 2025/04/10

   txtF0407 = Empty
   GRD1.Clear
   SetGrd
End Sub

Private Sub cmdDel_Click()
'Dim rsTmp As New ADODB.Recordset
Dim strDel As String 'Add by Amy 2020/05/21
'Add by Amy 2025/04/10
Dim strF0316 As String, strTo As String, strSubject As String, strContent As String

On Error GoTo ErrHand

   If MsgBox("確定是否要刪除資料？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then Exit Sub

   Screen.MousePointer = vbHourglass

   cnnConnection.BeginTrans
   
   'Modify by Amy 2020/05/21
   '商標延展結案
   If txtF0301 = MsgText(601) Then
       strDel = "Delete T102Inform Where ti02='" & m_stNP01 & "' And ti04='" & m_stNP22 & "' "
       Pub_SeekTbLog strDel 'Add by Amy 2023/02/14 記錄Log,可能沒回覆單,而 2022/06/20有改log顯示,故加入以便查詢
       cnnConnection.Execute strDel
       
       '檔案改放 FTP,必須在DB資料刪除前執行
       PUB_DelFtpFile2 m_CP01 & m_CP02 & m_CP03 & m_CP04, " and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))>0 and CPP10='U'"
      
      '刪除卷宗區暫存的回覆單附件
      strDel = "DELETE FROM casepaperpdf WHERE CPP01='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "' and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))>0 and CPP10='U'"
      Pub_SeekTbLog strDel '記錄Log
      cnnConnection.Execute strDel
   Else
       'Modify by Amy 2025/04/10 +FC結案單
       'Modify by Amy 2025/08/28 +intFCState
       Call PUB_CloseFlowDataDel(txtF0301, m_CP01, m_CP02, m_CP03, m_CP04, , strF0316, intFCState)
       
    '   '流程主檔
    '   strSql = "DELETE FROM Flow003 WHERE F0301='" & txtF0301 & "'"
    '   Pub_SeekTbLog strSql '記錄Log
    '   cnnConnection.Execute strSql
    '
    '   '簽核檔
    '   strSql = "SELECT * FROM Flow002 WHERE F0201='" & txtF0301 & "' order by F0202,F0203 asc"
    '   rsTmp.CursorLocation = adUseClient
    '   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    '   If rsTmp.RecordCount > 0 Then
    '      With rsTmp
    '         .MoveFirst
    '         Do While Not .EOF
    '            strSql = "DELETE FROM Flow002 WHERE F0201='" & txtF0301 & "' and F0202='" & rsTmp.Fields("F0202") & "' and F0203=" & rsTmp.Fields("F0203")
    '            Pub_SeekTbLog strSql '記錄Log
    '            cnnConnection.Execute strSql
    '            .MoveNext
    '         Loop
    '      End With
    '   End If
    '   rsTmp.Close
    '
    '   '流程備註檔
    '   strSql = "SELECT * FROM Flow004 WHERE F0401='" & txtF0301 & "' order by F0402 asc "
    '   rsTmp.CursorLocation = adUseClient
    '   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    '   If rsTmp.RecordCount > 0 Then
    '      With rsTmp
    '         .MoveFirst
    '         Do While Not .EOF
    '            strSql = "DELETE FROM Flow004 WHERE F0401='" & txtF0301 & "' and F0402=" & rsTmp.Fields("F0402")
    '            Pub_SeekTbLog strSql '記錄Log
    '            cnnConnection.Execute strSql
    '            .MoveNext
    '         Loop
    '      End With
    '   End If
    '   rsTmp.Close
    '
    '   '更新下一程序
    '   strSql = "Update NextProgress Set NP24=null WHERE NP24='" & txtF0301 & "'"
    '   Pub_SeekTbLog strSql '記錄Log
    '   cnnConnection.Execute strSql
    '
    '   '刪除卷宗區暫存的回覆單附件
    '   strSql = "DELETE FROM casepaperpdf WHERE CPP01='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "' and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))>0 and CPP10='U'"
    '   Pub_SeekTbLog strSql '記錄Log
    '   cnnConnection.Execute strSql
    '
    '   Set rsTmp = Nothing
   End If
   'end 2020/05/21
   cnnConnection.CommitTrans
   
   'Add by Amy 2025/06/26 外專退回後刪除需通知承辦主管,若二級主為自己(ex:David),則不需發
   'Modify by Amy 2025/08/12 +外商
   If (strSrvDate(1) >= FCP結案單電子化啟用日 And intFCState = 2) _
     Or (strSrvDate(1) >= FCT結案單電子化啟用日 And intFCState = 1) Then
      strTo = PUB_GetFCPProSup(strF0316) 'FC承辦人員2級主管
      'Modify by Amy 2025/08/12 外商需通知3級主管
      strExc(9) = ""
      If intFCState = 1 Then
         strExc(9) = GetST52SelfList(strF0316, "st53")
         If strExc(9) <> "" Then strTo = strTo & ";" & strExc(9)
      End If
      If strTo <> "" And ((intFCState = 2 And strTo <> strUserNum) Or (intFCState = 1)) Then
      'end 2025/08/12
         strSubject = GetPrjSalesNM(strF0316) & m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & "電子結案單【已刪除】通知"
         strContent = GetPrjSalesNM(strF0316) & " 已將電子結案單編號 " & txtF0301 & _
                              " (" & m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 & ") 刪除"
   
         '含 特殊職代
         'Modify by Amy 2025/08/12 只發人事職代-Sindy
         '      ex:1140811 Anny休假,黃賢泰填的結案單FCP-066436-0-00 不該發莊瑄凡(A8013)及洪培堯(A5023)
         'PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , , , , True
         PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , , , , True, , , , , , , , , , 1
      End If
   End If
   'end 2025/04/10
   Screen.MousePointer = vbDefault
   
   Call cmdExit_Click '結束
   Exit Sub

ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox "刪除失敗！" & vbCrLf & Err.Description
End Sub

Public Sub cmdExit_Click()

   'Added by Morgan 2015/12/15
   If m_PrevForm Is Nothing Then
      Unload Me
   Else
   'end 2015/12/15
   
      'Modify by Amy 2025/04/10 +frm100101_L
      If m_PrevForm.Name <> "frm100101_L" And m_PrevForm.Name <> "frm110101_2" Then
         m_PrevForm.Hide
         m_PrevForm.QueryData
      End If
      m_PrevForm.Show
      
   'Modified by Morgan 2015/12/15
   'Unload Me
      If m_PrevForm2 Is Nothing Then
         Unload Me
      Else
         Set m_PrevForm = m_PrevForm2
         Set m_PrevForm2 = Nothing
         Me.Hide
      End If
   End If
   'end 2015/12/15
   
   
End Sub

Private Sub cmdFile_Click()
Dim ii As Integer, jj As Integer, arrData As Variant
'Dim hLocalFile As Long
Dim strMsg As String 'Add by Amy 2025/04/10

Screen.MousePointer = vbHourglass
'Modify by Amy 2025/09/23 +if bug-T延展結案(無Flow003資料),有回覆單資料會無法顯示
If (Left(txtCaseNo, 2) = "T-" Or Left(txtCaseNo, 3) = "TF-") And (txtNP07.Tag = "102" Or txtNP07.Tag = "109" Or txtNP07.Tag = "716") And m_stNP22 <> "" Then
   Call Pub_OpenReplayPDFOrMsg(intFCState, Me, m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04, m_stNP22, m_strSaveFiles, m_AttachPath, strMsg)
Else
   Call Pub_OpenReplayPDFOrMsg(intFCState, Me, m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04, txtF0301, m_strSaveFiles, m_AttachPath, strMsg)
End If
'end 2025/09/23
If strMsg <> "" Then
   MsgBox strMsg
End If
Screen.MousePointer = vbDefault
'end 2025/04/10

'Mark by Amy 2025/04/10 改至Pub_OpenReplayPDFOrMsg,以下不執行
'   If m_strSaveFiles <> "" Then
'      Screen.MousePointer = vbHourglass
''      If m_strSaveFilesCP09 <> "" Then
''         frm100101_L.m_strKey = m_strSaveFilesCP09
''      Else
'         frm100101_L.m_strKey = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04
''      End If
'      frm100101_L.SetParent Me
'      If frm100101_L.QueryData = True Then
'         'Modify By Sindy 2018/2/12
''         For ii = frm100101_L.GRD1.Rows - 1 To 1 Step -1
''            If InStr(frm100101_L.GRD1.TextMatrix(ii, 4), m_strSaveFiles) > 0 Then Exit For
''         Next
'         arrData = Split(m_strSaveFiles, ":")
'         For jj = 0 To UBound(arrData)
'            For ii = frm100101_L.GRD1.Rows - 1 To 1 Step -1
'               If InStr(frm100101_L.GRD1.TextMatrix(ii, 4), arrData(jj)) > 0 Then Exit For
'            Next ii
'            If ii > 0 Then
'               Call frm100101_L.FrmCallOpenFile(ii, IIf(UBound(arrData) = jj, True, False))
'               If UBound(arrData) = jj Then
'                  frm100101_L.Show
'                  Me.Hide
'               End If
'            Else
'               Unload frm100101_L
'               Screen.MousePointer = vbDefault
'               MsgBox "有回覆單電子檔:" & m_strSaveFiles & " (找不到電子檔)"
'               Exit Sub
'            End If
'         Next jj
'         '2018/2/12 END
''         If ii > 0 Then
''            Call frm100101_L.FrmCallOpenFile(ii)
''            frm100101_L.Show
''            Me.Hide
''         Else
''            Unload frm100101_L
''         End If
'      Else
'         Unload frm100101_L
'      End If
'      Screen.MousePointer = vbDefault
''      If Dir(m_strSaveFiles, vbDirectory) = "" Then
''         MsgBox "電子檔不存在！", vbExclamation
''         Exit Sub
''      End If
''      '開啟檔案
''      ShellExecute hLocalFile, "open", m_strSaveFiles, vbNullString, vbNullString, 1
''   Else
''      MsgBox "無電子檔！", vbExclamation
''      Exit Sub
'   End If
End Sub

'Add by Amy 2025/08/18
Private Sub cmdInfo_Click()
   Call Pub_CloseShowfrm210133_INV(1, Me.Name, m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04)
End Sub

Private Sub cmdModify_Click()
   Dim stMsg As String 'Add by Amy 2025/04/10
   
   'Modify by Amy 2025/04/10 +intFCState
   If intFCState = 0 Then
      '非FC 結案單
      frm210133.txt1(0) = m_CP01
      frm210133.txt1(1) = m_CP02
      frm210133.txt1(2) = m_CP03
      frm210133.txt1(3) = m_CP04
      frm210133.m_F0301 = txtF0301
      frm210133.SetParent Me
      If frm210133.doQuery = True Then
         frm210133.Show
      Else
         Unload frm210133
      End If
   Else
      frm210133_F.txt1(0) = m_CP01
      frm210133_F.txt1(1) = m_CP02
      frm210133_F.txt1(2) = m_CP03
      frm210133_F.txt1(3) = m_CP04
      frm210133_F.m_CCM01 = txtF0301
      frm210133_F.SetParent Me
      If frm210133_F.doQueryCloseCase(stMsg) = True Then
         frm210133_F.Show
      Else
         If stMsg <> MsgText(601) Then
            MsgBox stMsg
         End If
         Unload frm210133_F
      End If
   End If
   Me.Hide
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me

   Me.Height = 6300 'Modify by Amy 2025/04/10 原:6120
   
   Me.txtF0301.BackColor = &H8000000F
   Me.txtF0310.BackColor = &H8000000F
   Me.txtF0310_2.BackColor = &H8000000F
   Me.txtF0309.BackColor = &H8000000F
   'Modify by Amy 2025/04/10 改共用,避免沒建資料夾
   'm_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum
   If Pub_SetFilePathDelTmp("Close", 1, strExc(9), m_AttachPath) = False Then
      MsgBox "附件資料夾建立失敗" & vbCrLf & _
                     strExc(9) & vbCrLf & "請洽電腦中心!"
   End If
   'end 2025/04/10
   If m_SignFlowEmp = "" Then m_SignFlowEmp = strUserNum '簽核人員
   'Add by Amy 2025/04/10 +FC結案單
   Call ClearSSTab1And2
   Frame10.Visible = False
   SSTab1.TabVisible(1) = False
   SSTab1.TabVisible(2) = False
   Call SetLock(True)
   'Memo by Amy 2025/06/09 外商結案單延期上線, intFCState=0(未上線前,前畫面以時間控制)
   If strSrvDate(1) >= FCP結案單電子化啟用日 And intFCState > 0 Then
      If intFCState = 1 Then
         Frame10.Visible = True '為外商 閉卷勾選才顯示
         Frame10.Enabled = False
         Call Pub_SetCloseGridAmtWidth(Me.Name, GridAMT, arrAMTCol(), arrAMTWidth(), True)
      End If
      SSTab1.TabVisible(intFCState) = True
      '專利
      If intFCState = 2 Then
         Call Pub_GetFCPContactItem(Me.Name, Chk6, Chk8, Chk9) 'FCP-聯絡項目
      End If
   End If
   Me.SSTab1.Tab = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
   m_stNP01 = "": m_stNP22 = "" 'Add by Amy 2020/05/21
   'Add by Amy 2025/04/10
   intFCState = Empty
   m_NP07 = ""
   'end 2025/04/10
   Set m_PrevForm = Nothing
   Set frm210147_1 = Nothing
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
getGrdColRow GRD1, x, y, nCol, nRow
GRD1.col = nCol
GRD1.row = nRow
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   '上一筆資料列清除反白
   If dblPrevRow > 0 Then
      GRD1.col = 2
      GRD1.row = dblPrevRow
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = QBColor(15)
      Next i
   End If
   '目前資料列反白
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   dblPrevRow = GRD1.row
   For i = 0 To GRD1.Cols - 1
      GRD1.col = i
      GRD1.CellBackColor = &HFFC0C0
   Next i
End If
GRD1.Visible = True
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String
Dim strUName As String
Dim strUDate As String
Dim strUTime As String

   If IsNull(rsSrcTmp.Fields("F0310")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("F0310")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("F0310"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("F0311")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("F0311")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("F0311"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("F0312")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("F0312")) = False Then
         strTemp = rsSrcTmp.Fields("F0312")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("F0313")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("F0313")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("F0313"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("F0314")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("F0314")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("F0314"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("F0315")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("F0315")) = False Then
         strTemp = rsSrcTmp.Fields("F0315")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If

   ' 設定CUID中的文字
   Label26.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

'Add by Amy 2025/04/10
'清除欄位值
Private Sub ClearSSTab1And2()
   Dim obj As Object

   ChkClose.Value = vbUnchecked 'FCT閉卷
   '請款項目
   For Each obj In txtA1K
      obj.Text = ""
   Next
   For Each obj In txtAmt
      obj.Text = ""
   Next
   For Each obj In lblName
      obj.Caption = ""
   Next
   GridAMT.Clear
      
   '聯絡項目 頁籤
   Chk7(0).Value = vbUnchecked 'FCP不續辦
   Chk7(1).Value = vbUnchecked 'FCP閉卷
   For Each obj In Chk6
      obj.Value = vbUnchecked
   Next
   txtNotPay = "" '未付帳款
   txtFCPMemo = "" '其他說明
   
   '不續辦 勾選項目
   For Each obj In Chk8
      obj.Value = vbUnchecked
   Next
   '閉卷 勾選項目
   For Each obj In Chk9
      obj.Value = vbUnchecked
   Next
  
End Sub

Private Sub GridAMT_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim intCol As Integer
   
   intCol = GetColVal(arrAMTCol(), "備註")
   GridAMT.ToolTipText = ""
   If GridAMT.MouseRow <> 0 And GridAMT.MouseCol = intCol Then
      If GridAMT.TextMatrix(GridAMT.MouseRow, intCol) <> "" Then
         GridAMT.ToolTipText = GridAMT.TextMatrix(GridAMT.MouseRow, intCol)
      End If
   End If
End Sub

'Add by Amy 2025/04/10
Private Sub txtA1K_Validate(Index As Integer, Cancel As Boolean)
    Dim stName As String
   
   If Index >= 3 And Index <= 6 And txtA1K(Index) <> MsgText(601) Then
      If Left(txtA1K(Index), 1) = 代理人編號 Then
         Call ClsPDGetAgent(txtA1K(Index), stName)
      ElseIf Left(txtA1K(Index), 1) = 客戶編號 Then
         stName = GetCustomerName(txtA1K(Index))
      End If
      lblName(Index - 3).Caption = stName
   End If
End Sub

'Add by Amy 2025/04/10
Private Sub SetLock(IsLock As Boolean)
   Dim obj As Object
   
   'Mark by Amy 2025/04/10 FC備註過長時,Frame1鎖住,導致無法使用ScrollBar,故拿掉
   'Frame1.Enabled = Not IsLock
   
   '請款項目
   For Each obj In txtA1K
      obj.Locked = IsLock
   Next
   For Each obj In txtAmt
      obj.Locked = IsLock
   Next
   
   '聯絡項目 頁籤
   Frame6.Enabled = Not IsLock
   Frame7.Enabled = Not IsLock
   Frame8.Enabled = Not IsLock
   txtFCPMemo.Locked = IsLock
End Sub
'end 2025/04/10


