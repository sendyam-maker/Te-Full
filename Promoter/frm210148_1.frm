VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210148_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件簽核作業"
   ClientHeight    =   6084
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
   ScaleHeight     =   6084
   ScaleWidth      =   8952
   Tag             =   "加班資料"
   Begin VB.CommandButton cmdFile 
      Caption         =   "檢視回覆單"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   3330
      TabIndex        =   0
      Top             =   0
      Width           =   1065
   End
   Begin VB.CommandButton Command1 
      Caption         =   "案件進度(&C)"
      Height          =   360
      Index           =   1
      Left            =   3630
      Style           =   1  '圖片外觀
      TabIndex        =   16
      Top             =   0
      Width           =   1125
   End
   Begin VB.CommandButton Command1 
      Caption         =   "完整卷宗(&H)"
      Height          =   360
      Index           =   3
      Left            =   2460
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   0
      Width           =   1125
   End
   Begin VB.TextBox txtF0309 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   5175
      Locked          =   -1  'True
      TabIndex        =   13
      Top             =   390
      Width           =   1845
   End
   Begin VB.TextBox txtF0310 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   390
      Width           =   645
   End
   Begin VB.TextBox txtF0301 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   90
      Width           =   945
   End
   Begin VB.CommandButton cmdQueryNext 
      Caption         =   "查詢下一筆(&N)"
      Height          =   360
      Left            =   6630
      TabIndex        =   3
      Top             =   0
      Width           =   1365
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "退回(&B)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   5700
      TabIndex        =   2
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同意(&O)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4770
      TabIndex        =   1
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8040
      TabIndex        =   4
      Top             =   0
      Width           =   885
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm210148_1.frx":0000
      Height          =   1188
      Left            =   4716
      TabIndex        =   5
      Top             =   4700
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
      Height          =   3500
      Left            =   0
      TabIndex        =   19
      Top             =   672
      Width           =   8904
      _ExtentX        =   15706
      _ExtentY        =   6181
      _Version        =   393216
      Tab             =   1
      TabHeight       =   420
      TabCaption(0)   =   "結案單"
      TabPicture(0)   =   "frm210148_1.frx":0015
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame10"
      Tab(0).Control(1)=   "txtCaseNo"
      Tab(0).Control(2)=   "txtApplNo"
      Tab(0).Control(3)=   "txtNP07"
      Tab(0).Control(4)=   "txtNP08"
      Tab(0).Control(5)=   "txtNP09"
      Tab(0).Control(6)=   "txtF0305"
      Tab(0).Control(7)=   "Label1(37)"
      Tab(0).Control(8)=   "txtAg"
      Tab(0).Control(9)=   "txtNation"
      Tab(0).Control(10)=   "txtF0306"
      Tab(0).Control(11)=   "txtNP15"
      Tab(0).Control(12)=   "txtNP10"
      Tab(0).Control(13)=   "txtNP14"
      Tab(0).Control(14)=   "txtApplPerson"
      Tab(0).Control(15)=   "txtCaseName"
      Tab(0).Control(16)=   "Label1(2)"
      Tab(0).Control(17)=   "Label1(3)"
      Tab(0).Control(18)=   "Label1(4)"
      Tab(0).Control(19)=   "Label1(5)"
      Tab(0).Control(20)=   "Label1(6)"
      Tab(0).Control(21)=   "Label1(7)"
      Tab(0).Control(22)=   "Label1(8)"
      Tab(0).Control(23)=   "Label1(9)"
      Tab(0).Control(24)=   "Label1(10)"
      Tab(0).Control(25)=   "Label1(11)"
      Tab(0).Control(26)=   "Label1(12)"
      Tab(0).Control(27)=   "Label1(13)"
      Tab(0).Control(28)=   "Label1(14)"
      Tab(0).ControlCount=   29
      TabCaption(1)   =   "請款項目"
      TabPicture(1)   =   "frm210148_1.frx":0031
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(21)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(18)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(26)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblName(3)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "lblName(2)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblName(1)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblName(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Label1(19)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label1(20)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(31)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(32)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label1(33)"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "FrameCRC"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "txtA1K(2)"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "txtA1K(6)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "txtA1K(5)"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "txtA1K(4)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "txtA1K(3)"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "txtA1K(0)"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "txtA1K(1)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "cmdInfo"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "聯絡項目"
      TabPicture(2)   =   "frm210148_1.frx":004D
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Frame8"
      Tab(2).Control(1)=   "Frame7"
      Tab(2).Control(2)=   "Frame6"
      Tab(2).Control(3)=   "Label1(17)"
      Tab(2).Control(4)=   "txtFCPMemo"
      Tab(2).ControlCount=   5
      Begin VB.CommandButton cmdInfo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "目前設定"
         Height          =   550
         Left            =   8250
         Style           =   1  '圖片外觀
         TabIndex        =   112
         Top             =   300
         Width           =   550
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  '沒有框線
         Height          =   280
         Left            =   -69000
         TabIndex        =   113
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
            TabIndex        =   114
            Top             =   0
            Width           =   950
         End
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   1
         Left            =   3504
         MaxLength       =   1
         TabIndex        =   92
         Top             =   372
         Width           =   255
      End
      Begin VB.TextBox txtA1K 
         BorderStyle     =   0  '沒有框線
         Height          =   270
         Index           =   0
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   91
         Top             =   360
         Width           =   255
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   3
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   90
         Text            =   "txtA1K28"
         Top             =   636
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   4
         Left            =   6000
         MaxLength       =   9
         TabIndex        =   89
         Top             =   636
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         Height          =   270
         Index           =   5
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   88
         Text            =   "txtA1K27"
         Top             =   948
         Width           =   850
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
         Index           =   2
         Left            =   6984
         MaxLength       =   1
         TabIndex        =   86
         Top             =   372
         Width           =   255
      End
      Begin VB.Frame Frame8 
         Height          =   1350
         Left            =   -69050
         TabIndex        =   78
         Top             =   252
         Width           =   2900
         Begin VB.CheckBox Chk7 
            Caption         =   "913 閉卷"
            Height          =   255
            Index           =   1
            Left            =   50
            TabIndex        =   82
            Top             =   120
            Width           =   1050
         End
         Begin VB.CheckBox Chk9 
            Caption         =   "不請款"
            Height          =   255
            Index           =   23
            Left            =   220
            TabIndex        =   81
            Top             =   996
            Width           =   2200
         End
         Begin VB.CheckBox Chk9 
            Caption         =   "請銷本所年費期限管制"
            Height          =   255
            Index           =   21
            Left            =   220
            TabIndex        =   80
            Top             =   400
            Width           =   2200
         End
         Begin VB.CheckBox Chk9 
            Caption         =   "閉卷請款"
            Height          =   255
            Index           =   22
            Left            =   220
            TabIndex        =   79
            Top             =   700
            Width           =   2200
         End
         Begin MSForms.ComboBox CboState 
            Height          =   288
            Index           =   1
            Left            =   1150
            TabIndex        =   83
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
      Begin VB.Frame Frame7 
         Height          =   3150
         Left            =   -72600
         TabIndex        =   66
         Top             =   252
         Width           =   3500
         Begin VB.CheckBox Chk8 
            Caption         =   "XXX"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   107
            Top             =   2760
            Visible         =   0   'False
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "需管制6個月補繳期"
            Height          =   255
            Index           =   19
            Left            =   96
            TabIndex        =   106
            Top             =   2508
            Width           =   3100
         End
         Begin VB.CheckBox Chk7 
            Caption         =   "907 不續辦"
            Height          =   255
            Index           =   0
            Left            =   50
            TabIndex        =   76
            Top             =   120
            Width           =   1150
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "駁不報告"
            Height          =   255
            Index           =   17
            Left            =   250
            TabIndex        =   73
            Top             =   2196
            Width           =   1200
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "後續准駁簡單報告"
            Height          =   255
            Index           =   16
            Left            =   250
            TabIndex        =   72
            Top             =   1956
            Width           =   2000
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "B.已獲指示"
            Height          =   255
            Index           =   15
            Left            =   1620
            TabIndex        =   71
            Top             =   1428
            Width           =   1200
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "透過其他管道代繳年費"
            Height          =   255
            Index           =   13
            Left            =   96
            TabIndex        =   70
            Top             =   900
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "A.未獲指示"
            Height          =   255
            Index           =   14
            Left            =   250
            TabIndex        =   69
            Top             =   1428
            Width           =   1200
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "代理人指示不繳年費"
            Height          =   255
            Index           =   12
            Left            =   96
            TabIndex        =   68
            Top             =   650
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "期限屆,未獲指示"
            Height          =   255
            Index           =   11
            Left            =   96
            TabIndex        =   67
            Top             =   400
            Width           =   3100
         End
         Begin MSForms.ComboBox CboState 
            Height          =   288
            Index           =   0
            Left            =   1250
            TabIndex        =   77
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
         Begin VB.Label Label1 
            Caption         =   "不續辦請款："
            Height          =   252
            Index           =   16
            Left            =   96
            TabIndex        =   75
            Top             =   1200
            Width           =   1104
         End
         Begin VB.Label Label1 
            Caption         =   "後續准駁簡單報告："
            Height          =   252
            Index           =   15
            Left            =   96
            TabIndex        =   74
            Top             =   1750
            Width           =   1800
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3150
         Left            =   -74940
         TabIndex        =   61
         Top             =   250
         Width           =   2320
         Begin VB.TextBox txtCCD08 
            Height          =   270
            Left            =   1300
            MaxLength       =   8
            TabIndex        =   108
            Text            =   "txtCCD08"
            Top             =   960
            Width           =   850
         End
         Begin VB.CheckBox Chk6 
            Caption         =   "D/N run C 類工程師報告"
            Height          =   255
            Index           =   1
            Left            =   50
            TabIndex        =   64
            Top             =   100
            Width           =   2200
         End
         Begin VB.CheckBox Chk6 
            Caption         =   "查本案前款均已付清"
            Height          =   255
            Index           =   2
            Left            =   50
            TabIndex        =   63
            Top             =   400
            Width           =   2200
         End
         Begin VB.CheckBox Chk6 
            Caption         =   "未付帳款"
            Height          =   255
            Index           =   3
            Left            =   50
            TabIndex        =   62
            Top             =   700
            Width           =   2200
         End
         Begin VB.Label Label1 
            Caption         =   "該案未請款:"
            Height          =   252
            Index           =   38
            Left            =   60
            TabIndex        =   111
            Top             =   2520
            Width           =   1236
         End
         Begin VB.Label lblNotPay_CP 
            Caption         =   "該案未請款:"
            ForeColor       =   &H000000FF&
            Height          =   350
            Left            =   120
            TabIndex        =   110
            Top             =   2760
            Width           =   2100
         End
         Begin VB.Label lblNotPayCPN 
            Caption         =   "     管制催款日"
            Height          =   252
            Left            =   60
            TabIndex        =   109
            Top             =   960
            Width           =   1236
         End
         Begin MSForms.TextBox txtNotPay 
            Height          =   1296
            Left            =   48
            TabIndex        =   65
            Top             =   1230
            Width           =   2196
            VariousPropertyBits=   -1466941413
            ScrollBars      =   2
            Size            =   "3881;2293"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.TextBox txtCaseNo 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -73944
         Locked          =   -1  'True
         TabIndex        =   40
         Text            =   "P-107255-0-00"
         Top             =   372
         Width           =   1275
      End
      Begin VB.TextBox txtApplNo 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -71100
         Locked          =   -1  'True
         TabIndex        =   39
         Text            =   "201420105880.4"
         Top             =   372
         Width           =   1515
      End
      Begin VB.TextBox txtNP07 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -73944
         Locked          =   -1  'True
         TabIndex        =   38
         Text            =   "年費"
         Top             =   1550
         Width           =   1065
      End
      Begin VB.TextBox txtNP08 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -71100
         Locked          =   -1  'True
         TabIndex        =   37
         Text            =   "105/02/05"
         Top             =   1550
         Width           =   1065
      End
      Begin VB.TextBox txtNP09 
         BorderStyle     =   0  '沒有框線
         Height          =   255
         Left            =   -68460
         Locked          =   -1  'True
         TabIndex        =   36
         Text            =   "105/02/05"
         Top             =   1550
         Width           =   1065
      End
      Begin VB.TextBox txtF0305 
         BorderStyle     =   0  '沒有框線
         Height          =   315
         Left            =   -73944
         Locked          =   -1  'True
         TabIndex        =   35
         Text            =   "其他"
         Top             =   2640
         Width           =   7650
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
         Height          =   2200
         Left            =   120
         TabIndex        =   20
         Top             =   1200
         Width           =   8712
         Begin VB.TextBox txtAmt 
            Alignment       =   1  '靠右對齊
            Height          =   270
            Index           =   2
            Left            =   6120
            MaxLength       =   6
            TabIndex        =   23
            Top             =   200
            Width           =   1200
         End
         Begin VB.TextBox txtAmt 
            Alignment       =   1  '靠右對齊
            Height          =   270
            Index           =   1
            Left            =   3840
            MaxLength       =   6
            TabIndex        =   22
            Top             =   200
            Width           =   1200
         End
         Begin VB.TextBox txtAmt 
            Alignment       =   1  '靠右對齊
            Height          =   270
            Index           =   0
            Left            =   1512
            TabIndex        =   21
            Top             =   200
            Width           =   1200
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridAMT 
            Height          =   1504
            Left            =   80
            TabIndex        =   24
            Top             =   600
            Width           =   8244
            _ExtentX        =   14542
            _ExtentY        =   2646
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            FormatString    =   "順序|代號|請款項目|金額|折扣|備註"
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
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
            TabIndex        =   27
            Top             =   204
            Width           =   600
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
            TabIndex        =   26
            Top             =   204
            Width           =   600
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
            TabIndex        =   25
            Top             =   204
            Width           =   804
         End
         Begin VB.Line Line2 
            BorderWidth     =   2
            X1              =   90
            X2              =   8300
            Y1              =   500
            Y2              =   500
         End
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "代理人："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   37
         Left            =   -74880
         TabIndex        =   34
         Top             =   1240
         Width           =   720
      End
      Begin MSForms.TextBox txtAg 
         Height          =   288
         Left            =   -73944
         TabIndex        =   105
         Top             =   1240
         Width           =   7656
         VariousPropertyBits=   679495711
         Size            =   "13504;508"
         Value           =   "txtFg"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtNation 
         Height          =   288
         Left            =   -71100
         TabIndex        =   60
         Top             =   936
         Width           =   1608
         VariousPropertyBits=   679495711
         Size            =   "2831;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtF0306 
         Height          =   432
         Left            =   -73944
         TabIndex        =   59
         Top             =   2960
         Width           =   7656
         VariousPropertyBits=   -1466939361
         ScrollBars      =   3
         Size            =   "13494;758"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtNP15 
         Height          =   516
         Left            =   -73944
         TabIndex        =   58
         Top             =   2110
         Width           =   7656
         VariousPropertyBits=   -1466939361
         ScrollBars      =   3
         Size            =   "13504;900"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtNP10 
         Height          =   288
         Left            =   -73944
         TabIndex        =   57
         Top             =   1810
         Width           =   1608
         VariousPropertyBits=   679495711
         Size            =   "2831;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtNP14 
         Height          =   288
         Left            =   -71280
         TabIndex        =   56
         Top             =   1810
         Width           =   1608
         VariousPropertyBits=   679495711
         Size            =   "2831;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtApplPerson 
         Height          =   288
         Left            =   -73944
         TabIndex        =   55
         Top             =   936
         Width           =   1872
         VariousPropertyBits=   679495711
         Size            =   "3307;503"
         Value           =   "txtApplPerson"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCaseName 
         Height          =   288
         Left            =   -73944
         TabIndex        =   54
         Top             =   636
         Width           =   7656
         VariousPropertyBits=   679495711
         Size            =   "13494;503"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所案號："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   2
         Left            =   -74880
         TabIndex        =   53
         Top             =   372
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請案號："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   3
         Left            =   -72024
         TabIndex        =   52
         Top             =   372
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件名稱："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   4
         Left            =   -74880
         TabIndex        =   51
         Top             =   636
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請人："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   5
         Left            =   -74880
         TabIndex        =   50
         Top             =   936
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "申請國家："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   6
         Left            =   -72024
         TabIndex        =   49
         Top             =   936
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "下一程序："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   7
         Left            =   -74880
         TabIndex        =   48
         Top             =   1550
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "本所期限："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   8
         Left            =   -72024
         TabIndex        =   47
         Top             =   1550
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "法定期限："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   9
         Left            =   -69420
         TabIndex        =   46
         Top             =   1524
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "智權人員："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   10
         Left            =   -74880
         TabIndex        =   45
         Top             =   1810
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "相關人："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   11
         Left            =   -72024
         TabIndex        =   44
         Top             =   1788
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "案件備註："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   12
         Left            =   -74880
         TabIndex        =   43
         Top             =   2112
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "結案記錄："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   13
         Left            =   -74880
         TabIndex        =   42
         Top             =   2640
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "說明如下："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   14
         Left            =   -74880
         TabIndex        =   41
         Top             =   2960
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "列印申請人：       (要印:Y/改不印:N)"
         Height          =   252
         Index           =   33
         Left            =   2400
         TabIndex        =   103
         Top             =   396
         Width           =   3000
      End
      Begin VB.Label Label1 
         Caption         =   "固定請款對象："
         Height          =   252
         Index           =   32
         Left            =   4680
         TabIndex        =   102
         Top             =   672
         Width           =   1296
      End
      Begin VB.Label Label1 
         Caption         =   "列印對象："
         Height          =   252
         Index           =   31
         Left            =   120
         TabIndex        =   101
         Top             =   996
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "請款對象："
         Height          =   252
         Index           =   20
         Left            =   120
         TabIndex        =   100
         Top             =   672
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "固定列印對象："
         Height          =   252
         Index           =   19
         Left            =   4680
         TabIndex        =   99
         Top             =   996
         Width           =   1296
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(0)"
         Height          =   252
         Index           =   0
         Left            =   2040
         TabIndex        =   98
         Top             =   648
         Width           =   2496
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(1)"
         Height          =   252
         Index           =   1
         Left            =   6960
         TabIndex        =   97
         Top             =   672
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
         Caption         =   "lblName(3)"
         Height          =   252
         Index           =   3
         Left            =   6960
         TabIndex        =   95
         Top             =   996
         Width           =   1500
      End
      Begin VB.Label Label1 
         Caption         =   "帳款已清：         (已清:Y)"
         Height          =   252
         Index           =   26
         Left            =   120
         TabIndex        =   94
         Top             =   396
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "合併列印請款單：       (要印:Y)"
         Height          =   252
         Index           =   18
         Left            =   5520
         TabIndex        =   93
         Top             =   396
         Width           =   2556
      End
      Begin VB.Label Label1 
         Caption         =   "其他說明："
         Height          =   252
         Index           =   17
         Left            =   -69050
         TabIndex        =   85
         Top             =   1680
         Width           =   996
      End
      Begin MSForms.TextBox txtFCPMemo 
         Height          =   1500
         Left            =   -68920
         TabIndex        =   84
         Top             =   1920
         Width           =   2748
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "4847;2646"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "列印申請人：       (要印:Y)"
         Height          =   252
         Index           =   30
         Left            =   -72000
         TabIndex        =   33
         Top             =   360
         Width           =   2496
      End
      Begin VB.Label Label1 
         Caption         =   "固定請款對象："
         Height          =   252
         Index           =   29
         Left            =   -70320
         TabIndex        =   32
         Top             =   636
         Width           =   1300
      End
      Begin VB.Label Label1 
         Caption         =   "列印對象："
         Height          =   252
         Index           =   28
         Left            =   -74880
         TabIndex        =   31
         Top             =   950
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "請款對象："
         Height          =   252
         Index           =   27
         Left            =   -74880
         TabIndex        =   30
         Top             =   636
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "固定列印對象："
         Height          =   252
         Index           =   23
         Left            =   -70320
         TabIndex        =   29
         Top             =   950
         Width           =   1296
      End
      Begin VB.Label Label1 
         Caption         =   "其他說明："
         Height          =   252
         Index           =   21
         Left            =   5640
         TabIndex        =   28
         Top             =   1800
         Width           =   996
      End
   End
   Begin MSForms.TextBox txtNote 
      Height          =   468
      Left            =   936
      TabIndex        =   104
      Top             =   4220
      Width           =   7992
      VariousPropertyBits=   -1466939365
      MaxLength       =   200
      ScrollBars      =   3
      Size            =   "14097;829"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtF0407 
      Height          =   1188
      Left            =   936
      TabIndex        =   18
      Top             =   4700
      Width           =   3732
      VariousPropertyBits=   -1466939361
      ScrollBars      =   3
      Size            =   "6583;2099"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtF0310_2 
      Height          =   285
      Left            =   1650
      TabIndex        =   17
      Top             =   390
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
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "目前表單狀態："
      ForeColor       =   &H000000C0&
      Height          =   180
      Left            =   3870
      TabIndex        =   14
      Top             =   420
      Width           =   1260
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
      Left            =   60
      TabIndex        =   10
      Top             =   5900
      Width           =   6672
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "您的意見："
      Height          =   180
      Left            =   36
      TabIndex        =   9
      Top             =   4220
      Width           =   900
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "簽核意見："
      Height          =   180
      Left            =   36
      TabIndex        =   8
      Top             =   4700
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "表單編號："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   7
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "填表人員："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   6
      Top             =   420
      Width           =   900
   End
End
Attribute VB_Name = "frm210148_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/27 Form2.0已修改
'Create by Sindy 2015/1/12
Option Explicit
 
' 變數宣告區
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_CP13 As String

Dim dblPrevRow As Double
Dim m_PrevForm As Form '前一畫面
Dim m_F0302 As String
Public m_F0303 As String
Dim m_F0304 As String
Dim m_F0308 As String '下一處理人員
Dim m_F0309 As String '目前狀態
Dim m_F0316 As String '智權人員
Dim i As Integer
Dim strUpdDate As String, strUpdTime As String
Dim strContent As String, strSubject As String
Dim m_strSaveFiles As String '附件
Dim m_strSaveFilesCP09 As String '附件總收文號
Dim m_AttachPath As String '附件暫存區
Dim m_PA09 As String
Public m_SignFlowEmp As String '簽核人員(因有可能人員休假職代代為操作)
Public cmdState As Integer
Public cmdAction As Integer 'Add By Sindy 2019/3/25 卷宗區回傳值 1=同意 2=下一筆 9=結束
Public PreF0302 As String 'Modify by Amy 2022/09/16
'Add by Amy 2025/04/10 FC結案單用
Public intFCState As Integer '0-非FC/1-商標/2-專利
Public m_NP07 As String '案件性質編號
Dim strNation As String, arrAMTCol() As String, arrAMTWidth() As String

Public Sub SetParent(ByRef fm As Form)
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
Dim rsTmp As New ADODB.Recordset, strSql As String, strFlowDBType As String
Dim strNP06 As String 'Add By Sindy 2017/2/23
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
   Call ClearSSTab1And2(False) 'Add by Amy 2025/08/28 前畫面勾選多筆資料進入第2筆前筆資料會殘留
   strFlowDBType = GetFlowFormReadDBType(txtF0301, , m_CP01, m_CP02, m_CP03, m_CP04)
   
   'Add by Amy 2022/09/16
   If PreF0302 <> MsgText(601) Then strSql = "And f0302='" & PreF0302 & "' "
   
   '案件表單主檔
   'Modify by Amy 2022/09/16 +m_F0302
   'Modify by Amy 2025/04/10 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
   If strSrvDate(1) >= FCP結案單電子化啟用日 Then
      strSql = "select flow003.*,CloseCaseMain.*,decode(F0309," & ShowFlow表單狀態中文 & ") as F0309NM" & _
                  " from flow003,CloseCaseMain" & _
                  " where f0301='" & txtF0301 & "' And f0301=ccm01(+) " & strSql
   Else
      strSql = "select flow003.*,decode(F0309," & ShowFlow表單狀態中文 & ") as F0309NM" & _
               " from flow003" & _
               " where f0301='" & txtF0301 & "' " & strSql
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
         stMergeBill = "" & rsTmp.Fields("ccm19") '合併列印請款單
         stA1K04 = "" & rsTmp.Fields("ccm20") '列印申請人
         stA1K28 = "" & rsTmp.Fields("ccm21") '請款對象
         stA1K27 = "" & rsTmp.Fields("ccm22") '列印對象
         stCCM23 = "" & rsTmp.Fields("ccm23") '固定請款對象
         stCCM24 = "" & rsTmp.Fields("ccm24") '固定列印對象
         stCCM25 = "" & rsTmp.Fields("ccm25") '帳款已清(填結案單時,帳款狀態)
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
      m_F0309 = rsTmp.Fields("F0309") '目前狀態
      txtF0309 = rsTmp.Fields("F0309") & " " & rsTmp.Fields("F0309NM")
      If m_F0302 = Flow_結案單 Then
         'Mark by Amy 2025/04/10 FC備註過長時,Frame1鎖住,導致無法使用ScrollBar,故拿掉
'         Frame1.Top = 600
'         Frame1.Visible = True
         'end 2025/04/10
         'Modify by Amy 2025/04/10 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中,並改抓共用
'         strSql = "SELECT ROR02 FROM REASONOFRELIEF WHERE ROR01='" & rsTmp.Fields("F0305") & "'"
'         CheckOC
'         adoRecordset.CursorLocation = adUseClient
'         adoRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'         If adoRecordset.RecordCount <> 0 And adoRecordset.RecordCount > 0 Then
'            If IsNull(adoRecordset.Fields(0)) Then
'               txtF0305 = rsTmp.Fields("F0305") + ""
'            Else
'               txtF0305 = rsTmp.Fields("F0305") + "  " + adoRecordset.Fields(0)
'            End If
'         Else
'            txtF0305 = ""
'         End If
'         CheckOC
'         txtF0306 = "" & rsTmp.Fields("F0306")
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
   
   '下一程序
   If strFlowDBType = "NP" Then
      strSql = "select *" & _
               " from nextprogress,casepropertymap,staff" & _
               " where np01='" & m_F0303 & "' and np22=" & m_F0304 & _
               " and np02=cpm01(+) and np07=cpm02(+)" & _
               " and np10=st01(+)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         m_CP01 = rsTmp.Fields("np02")
         m_CP02 = rsTmp.Fields("np03")
         m_CP03 = rsTmp.Fields("np04")
         m_CP04 = rsTmp.Fields("np05")
         txtNP07 = GetPrjState4(m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04, rsTmp.Fields("np07"))
         txtNP08 = ChangeWStringToTDateString("" & rsTmp.Fields("np08"))
         txtNP09 = ChangeWStringToTDateString("" & rsTmp.Fields("np09"))
         txtNP10 = rsTmp.Fields("st02")
         txtNP14 = "" & rsTmp.Fields("np14")
         txtNP15 = "" & rsTmp.Fields("np15")
         strNP06 = "" & rsTmp.Fields("np06") 'Add By Sindy 2017/2/23
      End If
      rsTmp.Close
   End If
   'Add by Amy 2025/04/10 +FC結案單明細
   If strSrvDate(1) >= FCP結案單電子化啟用日 Then
      '外商
      If intFCState = 1 Then
         If strClose = "Y" Then ChkClose.Value = vbChecked
         'Modify by Amy 2025/08/18 原只顯示,外商結案單增加欄位記錄,以利後續外商程序操作請款單
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
         End If
      '聯絡項目
      ElseIf strState = "2" Then
         'Add by Amy 2025/07/29 案件備註(原:備註)顯示灰色-Anny ex:P-129167 帶太久前資料,會混淆
         Label1(12).Enabled = False
         txtNP15.Enabled = False
         'end 2025/07/29
         '913 閉卷
         If strClose = "Y" Then
            CboState(1) = strCCM07N
            Chk7(1).Value = vbChecked
         '907 不續辦
         Else
            CboState(0) = strCCM07N
            Chk7(0).Value = vbChecked
         End If
         'Modify by Amy 2025/08/08 strCCD08/strNotPay_CP ,未付帳款 及 管制催款日 都有資料,於程序操作完 解除期限/閉卷 後加行事曆
         If Pub_QueryFCCloseDetail(Val(strState), Me.Name, txtF0301, m_CP01, strMsg, , Chk6, Chk8, Chk9, strNotPay, strMemo, strCCD08, strNotPay_CP) = False Then
            If strMsg <> "無資料" Then
               strMsg = "讀取FC專利[結案明細]資料有誤,通知電腦中心！" & vbCrLf & strMsg
               GoTo EXITSUB
            End If
         End If
         If strNotPay <> MsgText(601) Then Chk6(3).Value = vbChecked
         txtNotPay = strNotPay
         'Add by Amy 2025/08/08
         txtCCD08 = strCCD08
         lblNotPay_CP.Caption = strNotPay_CP
         'end 2025/08/08
         txtFCPMemo = strMemo
      End If
      '外商
      If intFCState = 1 Then
         Call Pub_SetCloseGridAmtWidth(Me.Name, Me.GridAMT, arrAMTCol(), arrAMTWidth)
         GridAMT.col = 0
         GridAMT.row = 1
      End If
   End If
   'end 2025/04/10
   
   '基本檔
   'Modify by Amy 2018/06/19 非P 案結案電子化,加入其他基本檔
'   strSql = "select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cu04,pa09,na03,pa91" & _
'            " from patent,customer,nation" & _
'            " where pa01='" & m_CP01 & "' and pa02='" & m_CP02 & "' and pa03='" & m_CP03 & "' and pa04='" & m_CP04 & "'" & _
'            " and substr(pa26,1,8)=cu01(+) and substr(pa26,9)=cu02(+)" & _
'            " and pa09=na01(+)"
   'Modify by Amy 2025/04/10 +代理人/申請國家編號
   strSql = ",FCAg,NVL(FA04,Decode(FA05,null,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) as AName,na01 "
   If intFCState > 0 Then strSql = ",FCAg,Decode(FA05,null,Nvl(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65) as AName,na01 "
   strSql = "select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa11,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cu04,pa09,na03,pa91 " & strSql & _
                "From Customer,Nation,Fagent," & _
                "(Select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa11,pa09,pa91,pa26,PA75 as FCAg " & _
                "From patent where pa01='" & m_CP01 & "' and pa02='" & m_CP02 & "' and pa03='" & m_CP03 & "' and pa04='" & m_CP04 & "' " & _
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
   'end 2018/06/19
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      txtCaseNo = rsTmp.Fields("pa01") & "-" & rsTmp.Fields("pa02") & "-" & rsTmp.Fields("pa03") & "-" & rsTmp.Fields("pa04")
      txtApplNo = "" & rsTmp.Fields("pa11")
      txtCaseName = rsTmp.Fields("pa05") & rsTmp.Fields("pa06") & rsTmp.Fields("pa07")
      txtApplPerson = rsTmp.Fields("cu04")
      m_PA09 = rsTmp.Fields("PA09")
      txtNation = rsTmp.Fields("na03")
      'Add by Amy 2025/04/10 +代理人編號/申請國家編號
      txtAg = "" & rsTmp.Fields("AName")
      strNation = "" & rsTmp.Fields("na01")
      'end 2025/04/10
           
      m_CP13 = ShowCurrCP13(m_CP01, m_CP02, m_CP03, m_CP04, m_PA09)
      If txtNP10 = "" Then
         txtNP10 = GetPrjSalesNM(m_CP13) '智權人員
      End If
      txtNP15 = "" & rsTmp.Fields("pa91")
   End If
   rsTmp.Close
   
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
   
   '回覆單:
   'Add by Amy 2025/04/10 資料夾未建立者先建,避免資料抓不到
   If m_AttachPath = "" Then
      If Pub_SetFilePathDelTmp("Close", 1, strExc(9), m_AttachPath) = False Then
         MsgBox "附件資料夾建立失敗" & vbCrLf & _
                        strExc(9) & vbCrLf & "請洽電腦中心!"
      End If
   End If
   'end 2025/04/10
   'Modify By Sindy 2020/12/28 無回覆單,則顯示卷宗區
   cmdFile.Caption = "檢視回覆單"
'   cmdFile.Enabled = False
   'Modify By Sindy 2015/5/18
   'If PUB_ChkIsReplyFile(m_CP01, m_CP02, m_CP03, m_CP04, m_strSaveFiles, txtF0301, m_strSaveFilesCP09) = True Then
   'Modify by Amy 2025/04/10 +FC結案單
   'If PUB_ChkIsReplyFile(m_CP01 & m_CP02 & m_CP03 & m_CP04, m_strSaveFiles, txtF0301, m_strSaveFilesCP09, txtF0301) = True Then
   '2015/5/18 END
   If PUB_ChkIsReplyFile(m_CP01 & m_CP02 & m_CP03 & m_CP04, m_strSaveFiles, txtF0301, m_strSaveFilesCP09, txtF0301, intFCState, m_AttachPath) = True Then
   'end 2025/04/10
      If m_strSaveFiles <> "" Then
         If strSrvDate(1) >= FCP結案單電子化啟用日 And intFCState > 0 Then cmdFile.Caption = "檢視電子檔" 'Add by Amy 2025/04/10
'         cmdFile.Enabled = True
      Else
         cmdFile.Caption = "卷宗區" 'Modify By Sindy 2020/12/28
'         If PUB_GetAttachFile_CPP(m_CP01 & m_CP02 & m_CP03 & m_CP04, m_strSaveFiles, m_AttachPath) = False Then
'            MsgBox "無法儲存欲開啟的檔案[ " & m_strSaveFiles & " ]！"
'         End If
      End If
   Else
      cmdFile.Caption = "卷宗區" 'Modify By Sindy 2020/12/28
   End If
   
   'Add By Sindy 2017/2/23
   If strNP06 = "Y" Then '已收文
      '簽核主管
      If UCase(m_PrevForm.Name) = UCase("frm210148") Then
         cmdOK.Enabled = False
         MsgBox "此案已收文，請退回給智權人員！", vbExclamation
      '判發主管
      Else
         MsgBox "此案已收文！", vbExclamation
      End If
   End If
   '2017/2/23 END
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   
EXITSUB:
   'Add by Amy 2025/04/10 避免有錯,無法離開
   If strMsg <> MsgText(601) Then
      cmdOK.Enabled = False
      cmdBack.Enabled = False
      MsgBox strMsg
   End If
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   'end 2025/04/10
   Set rsTmp = Nothing
End Sub

Private Sub ClearField()
   'Frame1:結案單
   txtCaseNo = Empty
   txtApplNo = Empty
   txtCaseName = Empty
   txtApplPerson = Empty
   txtNation = Empty
   txtNP07 = Empty
   txtNP08 = Empty
   txtNP09 = Empty
   txtNP10 = Empty
   txtNP14 = Empty
   txtNP15 = Empty
   'Add by Amy 2025/04/10
   txtAg = Empty
   strNation = Empty
   'end 2025/04/10
   
   txtNote = Empty
   txtF0407 = Empty
   GRD1.Clear
   SetGrd
   cmdBack.Enabled = True 'Add By Sindy 2023/5/16
End Sub

Private Sub cmdFile_Click()
Dim ii As Integer, jj As Integer, arrData As Variant
'Dim hLocalFile As Long
Dim strMsg As String 'Add by Amy 2025/04/10

Screen.MousePointer = vbHourglass
Call Pub_OpenReplayPDFOrMsg(intFCState, Me, m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04, txtF0301, m_strSaveFiles, m_AttachPath, strMsg)
If strMsg <> "" Then
   MsgBox strMsg
End If
Screen.MousePointer = vbDefault
'end 2025/04/10
   
'Mark by Amy 2025/04/10 改至Pub_OpenReplayPDFOrMsg,以下不執行
'   'Modify By Sindy 2020/12/28 無回覆單,則顯示卷宗區
'   'If m_strSaveFiles <> "" Then
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
'               If InStr(frm100101_L.GRD1.TextMatrix(ii, 4), arrData(jj)) > 0 Then
'                  Exit For
'               End If
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
'               MsgBox "有回覆單電子檔:" & m_strSaveFiles
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
'
'   'End If
End Sub

'Add by Amy 2025/08/18
Private Sub cmdInfo_Click()
   Call Pub_CloseShowfrm210133_INV(1, Me.Name, m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04)
End Sub

Private Sub cmdok_Click()
Dim Cancel As Boolean

On Error GoTo ErrHand
   
   'Add By Sindy 2024/1/25
   Call txtNote_Validate(Cancel)
   If Cancel = True Then
      Exit Sub
   End If
   '2024/1/25 END
   
   Screen.MousePointer = vbHourglass
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   'Add by Sindy 2021/5/28 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Sub
   End If
   '2021/5/28 END
   
   cnnConnection.BeginTrans
   
   If Left(txtF0309, 2) = Flow_已完成 Or Left(txtF0309, 2) = Flow_判發重送 Then
      m_F0309 = Flow_歸檔
      '3.判發人員:
      '簽核檔
      strSql = "update FLOW002 set " & _
               "F0205='" & strUpdDate & "'" & _
               ",F0206='" & strUpdTime & "'" & _
               ",F0207='1'" & IIf(m_SignFlowEmp <> strUserNum, ",F0208='(代)',F0204='" & strUserNum & "'", "") & _
               " where F0201='" & txtF0301 & "' and F0202='3' and F0207 is null"
      cnnConnection.Execute strSql
      '表單主檔
      'Modify by Amy 2022/09/16 +m_F0302
      strSql = "update FLOW003 set " & _
               "F0309=" & CNULL(m_F0309) & _
               " where F0301='" & txtF0301 & "' And F0302='" & m_F0302 & "' "
      cnnConnection.Execute strSql
      '流程備註檔
      strSql = GetInsertFLOW004Sql(Trim(txtF0301), strUserNum, strUpdDate, strUpdTime, m_F0309, ChgSQL(Trim(txtNote.Text)))
      cnnConnection.Execute strSql
   Else
      '1.簽核主管:
      '簽核檔
      'Modify by Amy 2018/08/27 + and F0202='1' 因簽核主管=補看人員,補看筆那會一併被簽核 ex:T-212788
      strSql = "update FLOW002 set " & _
               "F0205='" & strUpdDate & "'" & _
               ",F0206='" & strUpdTime & "'" & _
               ",F0207='1'" & IIf(m_SignFlowEmp <> strUserNum, ",F0208='(代)',F0204='" & strUserNum & "'", "") & _
               " where F0201='" & txtF0301 & "' and F0202='1' and F0204='" & m_SignFlowEmp & "' and F0207 is null "
      cnnConnection.Execute strSql
      '流程備註檔
      If Trim(txtNote.Text) <> "" Then
         strSql = GetInsertFLOW004Sql(Trim(txtF0301), strUserNum, strUpdDate, strUpdTime, m_F0309, ChgSQL(Trim(txtNote.Text)))
         cnnConnection.Execute strSql
      End If
      '讀取下一處理人員
      If GetNextProPerson_Flow(Trim(txtF0301), Trim(m_F0316), m_F0308, m_F0309) = False Then GoTo ErrHand
   End If
   
   cnnConnection.CommitTrans
'   cmdok.Enabled = False
   
   '發E-Mail通知下一處理主管(多審核主管用)
   If m_F0309 = Flow_主管審核中 Then
      strContent = GetEMailContent_Flow(txtF0301, strSubject)
'      MsgBox "收件者：" & m_F0308 & GetPrjSalesNM(m_F0308) & vbCrLf & vbCrLf & _
'             "主　旨：" & strSubject & vbCrLf & vbCrLf & _
'             "內　容：" & strContent, vbInformation
      'Modify by Amy 2025/08/12 +if F字頭部門,只發人事職代-Sindy
      '     ex:1140811 Anny休假,黃賢泰填的結案單FCP-066436-0-00 不該發莊瑄凡(A8013)及洪培堯(A5023)
      If Left(PUB_GetST03(strUserNum), 1) = "F" Then
         PUB_SendMail strUserNum, m_F0308, "", strSubject, strContent, , , , , , , , , , , , , True, , , , , , , , , , 1
      Else
         'Modify By Sindy 2016/10/12 + 含特殊職代
         PUB_SendMail strUserNum, m_F0308, "", strSubject, strContent, , , , , , , , , , , , , True
      End If
   'Modify by Amy 2025/04/10 +FC承辦主管同意後,發信通知程序人員
   'Add by Amy 2018/08/27 發信通知承辦人員
   'ElseIf m_F0309 = Flow_處理中 And (m_CP01 = "CFT" Or m_CP01 = "CFC" Or m_CP01 = "S") Then
   ElseIf m_F0309 = Flow_處理中 _
     And ((m_CP01 = "CFT" Or m_CP01 = "CFC" Or (m_CP01 = "S" And strNation <> "000")) _
               Or (strSrvDate(1) >= FCP結案單電子化啟用日 And intFCState = 2) _
               Or (strSrvDate(1) >= FCT結案單電子化啟用日 And intFCState = 1)) Then
      'Modify By Sindy 2025/6/4
      strContent = GetEMailContent_Flow(txtF0301, strSubject)
      PUB_SendMail strUserNum, m_F0308, "", strSubject, strContent, , , , , , , , , , , , , True
'      strSubject = m_CP01 & "-" & m_CP02 & m_CP03 & m_CP04 & "結案單通知"
'      If m_CP01 = "CFT" Or m_CP01 = "CFC" Or m_CP01 = "S" Then
'         PUB_SendMail strUserNum, m_F0308, "", strSubject, 結案單外商CF操作路徑, , , , , , , , , , , , , True
'      ElseIf intFCState = 1 Then
'         PUB_SendMail strUserNum, m_F0308, "", strSubject, 結案單外商FC操作路徑, , , , , , , , , , , , , True
'      ElseIf intFCState = 2 Then
'         PUB_SendMail strUserNum, m_F0308, "", strSubject, 結案單外專FC操作路徑, , , , , , , , , , , , , True
'      End If
      '2025/6/4 END
   'end 2025/04/10
   End If
   
   Screen.MousePointer = vbDefault
   
   Me.txtF0301 = ""
   m_PrevForm.Show
   m_PrevForm.PubShowNextData
   If Me.txtF0301 = "" Then
      Unload Me
   End If
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox "簽核失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdBack_Click()
Dim SignPerson As String, intMaxF0203 As Integer
Dim strF0202_3 As String 'Add by Amy 2018/08/07
Dim rsTmp As New ADODB.Recordset
Dim Cancel As Boolean
Dim IsFMPY53374 As Boolean 'Add by Amy 2025/04/10
   
On Error GoTo ErrHand

   'Add By Sindy 2024/1/25
   Call txtNote_Validate(Cancel)
   If Cancel = True Then
      Exit Sub
   End If
   '2024/1/25 END
   '程序人員
   'Modify by Amy 2018/08/07 從下面搬上來先檢查,並增加補看人員
   'Modify by Amy 2020/03/06 +本所案號
   'Modify by Amy 2025/04/10 +if -FC結案單
   If intFCState = 0 Then
      SignPerson = GetSignOffEmp("NP", m_CP01, m_CP02, m_PA09, m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04) '程序人員
   Else
      IsFMPY53374 = PUB_FMPtoCheck(1, 2, PUB_GetST05(Pub_GetField("Flow003", "F0301='" & txtF0301 & "'", "F0316")), m_CP01, m_CP02, m_CP03, m_CP04)
      SignPerson = GetSignOffEmp("NP", m_CP01, m_CP02, m_PA09, m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04, m_NP07, m_F0316, IsFMPY53374)
   End If
   If SignPerson = MsgText(601) Then
      MsgBox "無設定程序人員，請通知電腦中心！"
      Exit Sub
   End If
   '補看人員
   'Modify by Amy 2021/06/28 +本所案號 for CFT/CFC/S 判斷補看人員
   'Modify by Amy 2025/04/10 +FC結案單,傳入案件性質及案件程序人員(for 案件程序人員的二級主管)
   'Modify by Amy 2025/06/28 +傳 案件性質/程序人員/操作人員/申請國家 for FC結案單
   '  T or FCT爭議案-Pub_GetSpecMan("內商爭議案程序主管") / FCT 非爭議案-案件程序人員的二級主管;CFT 及 S 非台灣 照舊結案單
   '  P[非]寰華FMP-Pub_GetSpecMan("專利處特定編號") / 'FCP/FG/P寰華 案件程序人員的二級主管;CFP 照舊結案單
   strF0202_3 = GetF0202_3(m_CP01, m_CP02, m_CP03, m_CP04, m_NP07, Left(SignPerson, 5), strUserNum, strNation)
   If strF0202_3 = MsgText(601) Then
      MsgBox "無設定補看人員，請通知電腦中心！"
      Exit Sub
   End If
   'end 2018/08/07
   
   'Add By Sindy 2023/5/16
   If Left(txtF0309, 2) = Flow_已完成 Or Left(txtF0309, 2) = Flow_判發重送 Then
      'Modify By Sindy 2023/5/16 and cp09=np24(+) => and cp09=np01(+)
      strSql = "select cp09,np24 from caseprogress,nextprogress where cp140='" & txtF0301 & "'" & _
               " and cp01=np02(+) and cp02=np03(+) and cp03=np04(+) and cp04=np05(+) and cp09=np01(+)"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If "" & rsTmp.Fields("np24") <> "" Then
            cmdBack.Enabled = False
            MsgBox "下一期限已收文，不可退回！", vbInformation
            Exit Sub
         End If
      End If
      rsTmp.Close
   End If
   '2023/5/16 END
   
   'Add by Sindy 2021/5/28 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Sub
   End If
   '2021/5/28 END
   
   If Trim(txtNote.Text) = "" Then
      MsgBox "您的意見不可以空白！", vbExclamation
      txtNote.SetFocus
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   cnnConnection.BeginTrans
   
   If Left(txtF0309, 2) = Flow_已完成 Or Left(txtF0309, 2) = Flow_判發重送 Then
      'Modify by Amy 2018/08/07 抓取程式人員搬至上面
      m_F0309 = Flow_判發退回 '判發人員退回程序:不需發E-Mail
      m_F0308 = Left(SignPerson, 5) '下一處理人員:程序
      '3.判發人員:
      '簽核檔
      strSql = "update FLOW002 set " & _
               "F0205='" & strUpdDate & "'" & _
               ",F0206='" & strUpdTime & "'" & _
               ",F0207='2'" & IIf(m_SignFlowEmp <> strUserNum, ",F0208='(代)',F0204='" & strUserNum & "'", "") & _
               " where F0201='" & txtF0301 & "' and F0202='3' and F0207 is null"
      cnnConnection.Execute strSql
      
      '判發人員退回程序時再新增2筆待簽核的記錄
      '程序人員
      intMaxF0203 = 0
      If SignPerson <> "" Then
         strSql = "SELECT max(F0203) FROM FLOW002 WHERE F0201='" & txtF0301 & "' and F0202='2'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            intMaxF0203 = RsTemp.Fields(0)
         End If
         intMaxF0203 = intMaxF0203 + 1
         strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(txtF0301) & ",'2'," & intMaxF0203 & "," & CNULL(Left(SignPerson, 5)) & ")"
         cnnConnection.Execute strSql
      End If
      '補看人員
      intMaxF0203 = 0
      'Modify by Amy 2018/08/07
      'If Flow_補看人員 <> "" Then
      If strF0202_3 <> MsgText(601) Then
         strSql = "SELECT max(F0203) FROM FLOW002 WHERE F0201='" & txtF0301 & "' and F0202='3'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            intMaxF0203 = RsTemp.Fields(0)
         End If
         intMaxF0203 = intMaxF0203 + 1
         strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(txtF0301) & ",'3'," & intMaxF0203 & "," & CNULL(strF0202_3) & ")"
         cnnConnection.Execute strSql
      End If
      
   Else
      m_F0309 = Flow_退回 '程序或簽核主管退回當事人:發E-Mail
      m_F0308 = m_F0316 '下一處理人員:當事人
      '1.簽核主管:
      '簽核檔
      strSql = "update FLOW002 set " & _
               "F0205='" & strUpdDate & "'" & _
               ",F0206='" & strUpdTime & "'" & _
               ",F0207='2'" & IIf(m_SignFlowEmp <> strUserNum, ",F0208='(代)',F0204='" & strUserNum & "'", "") & _
               " where F0201='" & txtF0301 & "' and F0204='" & m_SignFlowEmp & "' and F0207 is null "
      cnnConnection.Execute strSql
   End If
   
   '表單主檔
   'Modify by Amy 2022/09/16 +m_F0302
   strSql = "update FLOW003 set " & _
            "F0307='" & strUserNum & "'" & _
            ",F0308='" & m_F0308 & "'" & _
            ",F0309='" & m_F0309 & "'" & _
            " where F0301='" & txtF0301 & "' And F0302='" & m_F0302 & "' "
   cnnConnection.Execute strSql
   
   '流程備註檔
   If Trim(txtNote.Text) <> "" Then
      strSql = GetInsertFLOW004Sql(Trim(txtF0301), strUserNum, strUpdDate, strUpdTime, m_F0309, ChgSQL(Trim(txtNote.Text)))
      cnnConnection.Execute strSql
   End If
   
   cnnConnection.CommitTrans
'   cmdBack.Enabled = False
   
   '發E-Mail通知當事人
   'Modify by Amy 2019/11/27 +外商案件判發退回發信通知承辦
   'Modify by Amy 2025/04/10 +strNation 外商CF S案只負責非台灣 及 FC結案單要發信通知
   If m_F0309 = Flow_退回 _
     Or (m_F0309 = Flow_判發退回 And ((m_CP01 = "CFT" Or m_CP01 = "CFC" Or (m_CP01 = "S" And strNation <> "000")) _
                                                                        Or (strSrvDate(1) >= FCP結案單電子化啟用日 And intFCState = 2) _
                                                                        Or (strSrvDate(1) >= FCT結案單電子化啟用日 And intFCState = 1))) Then
      strContent = GetEMailContent_Flow(txtF0301, strSubject)
      If Trim(txtNote.Text) <> "" Then
         strSubject = strSubject & "；退回原因：" & Trim(txtNote.Text)
      End If
'      MsgBox "收件者：" & m_F0308 & GetPrjSalesNM(m_F0308) & vbCrLf & vbCrLf & _
'             "主　旨：" & strSubject & vbCrLf & vbCrLf & _
'             "內　容：" & strContent, vbInformation
      'Modify by Amy 2025/04/10 +if 主管簽核或補看 退回時,填單人員非智權人員時,退回的收受者掛原智權人員,副本給填單人員-Sindy
      '                             因FCP退回給原智權人員,非操作人員
      If m_F0309 = Flow_退回 And txtF0310 <> m_F0316 Then
         PUB_SendMail strUserNum, m_F0316, "", strSubject, strContent, , , , , , txtF0310
      Else
         PUB_SendMail strUserNum, m_F0308, "", strSubject, strContent
      End If
      'end 2025/04/10
   End If
   
   Screen.MousePointer = vbDefault
   
   Me.txtF0301 = ""
   m_PrevForm.Show
   m_PrevForm.PubShowNextData
   If Me.txtF0301 = "" Then
      Unload Me
   End If
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   'Resume
   MsgBox "退回失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdQueryNext_Click()
   Me.txtF0301 = ""
   m_PrevForm.Show
   m_PrevForm.PubShowNextData
   If Me.txtF0301 = "" Then
      Unload Me
   End If
End Sub

Private Sub cmdExit_Click()
   m_PrevForm.Hide
   m_PrevForm.QueryData
   m_PrevForm.Show
   Unload Me
End Sub

Private Sub Command1_Click(Index As Integer)
   Select Case Index
   Case 1, 3
      cmdState = Index
      PubShowNextData
      Exit Sub
   End Select
End Sub

Public Sub PubShowNextData()
Dim ii As Integer, jj As Integer
Dim arrData As Variant
Dim strMsg As String 'Add by Amy 2025/04/10
   
   Select Case cmdState
   Case 1 '進度
      Me.Enabled = False
'      For i = 1 To MSHFlexGrid1.Rows - 1
'         MSHFlexGrid1.col = 0
'         MSHFlexGrid1.row = i
'         If Trim(MSHFlexGrid1.Text) = "V" Then
'            MSHFlexGrid1.col = 0
'            MSHFlexGrid1.Text = ""
'            For j = 0 To MSHFlexGrid1.Cols - 1
'               MSHFlexGrid1.col = j
'               MSHFlexGrid1.CellBackColor = QBColor(15)
'            Next j
'            MSHFlexGrid1.col = 3
'            If Not IsNull(MSHFlexGrid1.Text) Then
               If fnSaveParentForm(Me) = False Then
                  Me.Enabled = True
                  Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               frm100101_2.Show
               'frm100101_2.Tag = Pub_RplStr(MSHFlexGrid1.TextMatrix(i, 3)) '本所案號
               frm100101_2.Tag = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 '本所案號
               frm100101_2.cmdOK(5).Visible = False '下一筆按鈕隱藏
               frm100101_2.StrMenu
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
'            End If
'         End If
'      Next i
'      Me.Enabled = True
   
   Case 3 '完整卷宗
      'Add By Sindy 2019/3/25
      '卷宗區點結束
      If cmdAction = 9 Then cmdAction = 0: Exit Sub
      '卷宗區點同意
      If cmdAction = 1 Then
         Call cmdok_Click
         cmdAction = 0 'Added by Morgan 2019/4/26
         Exit Sub
      End If
      '2019/3/25 END
      
      Me.Enabled = False
'      For i = 1 To MSHFlexGrid1.Rows - 1
'         MSHFlexGrid1.col = 0
'         MSHFlexGrid1.row = i
'         If Trim(MSHFlexGrid1.Text) = "V" Then
'            MSHFlexGrid1.col = 0
'            MSHFlexGrid1.Text = ""
'            For j = 0 To MSHFlexGrid1.Cols - 1
'               MSHFlexGrid1.col = j
'               MSHFlexGrid1.CellBackColor = QBColor(15)
'            Next j
'            MSHFlexGrid1.col = 3
'            If Not IsNull(MSHFlexGrid1.Text) Then
               If fnSaveParentForm(Me) = False Then
                  Me.Enabled = True
                  Exit Sub
               End If
               Screen.MousePointer = vbHourglass
               'Add by Amy 2025/04/10 避免有未改到,改成共用
               Call Pub_OpenReplayPDFOrMsg(intFCState, Me, m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04, txtF0301, m_strSaveFiles, m_AttachPath, strMsg _
                                 , True)
               If strMsg <> "" Then
                  MsgBox "有電子檔如下:" & vbCrLf & Replace(strMsg, ";", vbCrLf) & vbCrLf & " (找不到電子檔)"
               End If
               'end 2025/04/10
               
               'Mark by Amy 2025/04/10 改至Pub_OpenReplayPDFOrMsg,以下不執行
'               'frm100101_L.m_strKey = Pub_RplStr(MSHFlexGrid1.TextMatrix(i, 3)) '本所案號
'               frm100101_L.m_strKey = m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04 '本所案號
'               frm100101_L.SetParent Me
'               If frm100101_L.QueryData = True Then
'                  'Add By Sindy 2019/3/25
'                  cmdAction = 9 '預設結束
'                  frm100101_L.cmdok(1).Caption = "同意"
'                  frm100101_L.cmdok(1).Visible = True
'                  '2019/3/25 END
'                  If m_strSaveFiles <> "" Then
'                     'Modify By Sindy 2018/2/12
''                     For ii = frm100101_L.GRD1.Rows - 1 To 1 Step -1
''                        If InStr(frm100101_L.GRD1.TextMatrix(ii, 4), m_strSaveFiles) > 0 Then Exit For
''                     Next
'                     arrData = Split(m_strSaveFiles, "&")
'                     For jj = 0 To UBound(arrData)
'                        For ii = frm100101_L.GRD1.Rows - 1 To 1 Step -1
'                           If InStr(frm100101_L.GRD1.TextMatrix(ii, 4), arrData(jj)) > 0 Then Exit For
'                        Next ii
'                        If ii > 0 Then
'                           Call frm100101_L.FrmCallOpenFile(ii, IIf(UBound(arrData) = jj, True, False))
'                           If UBound(arrData) = jj Then
'                              frm100101_L.Show
'                              Me.Hide
'                           End If
'                        Else
'                           Unload frm100101_L
'                           Screen.MousePointer = vbDefault
'                           Me.Enabled = True
'                           MsgBox "有回覆單電子檔:" & m_strSaveFiles
'                           Exit Sub
'                        End If
'                     Next jj
'                     '2018/2/12 END
''                     If ii > 0 Then
''                        Call frm100101_L.FrmCallOpenFile(ii)
''                     End If
'                  End If
''                  frm100101_L.Show
''                  Me.Hide
'               Else
'                  Unload frm100101_L
'               End If
               'Mark end 2025/04/10
               Screen.MousePointer = vbDefault
               Me.Enabled = True
               Exit Sub
'            End If
'         End If
'      Next i
'      Me.Enabled = True
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   Me.txtF0301.BackColor = &H8000000F
   Me.txtF0310.BackColor = &H8000000F
   Me.txtF0310_2.BackColor = &H8000000F
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
         Call Pub_SetCloseGridAmtWidth(Me.Name, GridAMT, arrAMTCol(), arrAMTWidth, True)
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
   PreF0302 = "" 'Add by Amy 2022/09/16
   'Add by Amy 2025/04/10
   intFCState = Empty
   m_NP07 = ""
   'end 2025/04/10
   Set m_PrevForm = Nothing
   Set frm210148_1 = Nothing
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

'Add By Sindy 2021/5/28
Private Sub txtNote_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtNote
End Sub

'Add By Sindy 2024/1/25
Private Sub txtNote_GotFocus()
   TextInverse txtNote
End Sub
Private Sub txtNote_Validate(Cancel As Boolean)
   '檢查長度
   If Not CheckLengthIsOK(txtNote, txtNote.MaxLength) Then
      txtNote.SetFocus
      Cancel = True
   End If
End Sub
'2024/1/25 END

'Add by Amy 2025/04/10 清除欄位值
'Modify by Amy 2025/08/28 +bolCls:清項目名
Private Sub ClearSSTab1And2(Optional ByVal bolCls As Boolean = True)
   Dim obj As Object
   
   '外商
   SSTab1.TabVisible(1) = False
   Frame10.Visible = False '閉卷勾選
   If intFCState = 1 Then
      Frame10.Visible = True '閉卷勾選才顯示
      Frame10.Enabled = False
      SSTab1.TabVisible(1) = True
      'Modify by Amy 2025/10/02 SetA1KColor 改抓共用
      'Call SetA1KColor(0) 'Add by Amy 2025/08/29 前一筆有設顏色,第2筆沒顏色要清
      Call Pub_CloseSetA1KDataColor(1, Me.Name, "", "", "", "", "", Me.txtA1K, 0, 6)
   End If
   '都要先清空,前一筆為新結案單時,第2筆為舊結案單資料會殘留
   ChkClose.Value = vbUnchecked 'FCT閉卷
   
   '請款項目
   If SSTab1.TabVisible(1) = True Then
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
      GridAMT.Rows = 2
   End If
   
   '聯絡項目 頁籤
   If SSTab1.TabVisible(2) = True Then
      'FCP不續辦
      CboState(0) = ""
      Chk7(0).Value = vbUnchecked
      'FCP閉卷
      CboState(1) = ""
      Chk7(1).Value = vbUnchecked
      For Each obj In Chk6
         obj.Value = vbUnchecked
         If bolCls = True Then obj.Caption = ""
      Next
      txtNotPay = "" '未付帳款
      lblNotPay_CP.Caption = "" '收文未請款=該案未請款
      txtFCPMemo = "" '其他說明
      
      '不續辦 勾選項目
      For Each obj In Chk8
         obj.Value = vbUnchecked
         If bolCls = True Then obj.Caption = ""
      Next
      
      '閉卷 勾選項目
      For Each obj In Chk9
         obj.Value = vbUnchecked
         If bolCls = True Then obj.Caption = ""
      Next
   End If
End Sub

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


