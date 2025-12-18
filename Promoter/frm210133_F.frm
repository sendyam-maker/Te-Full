VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm210133_F 
   BorderStyle     =   1  '單線固定
   Caption         =   "案件結案單"
   ClientHeight    =   6780
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   9024
   ControlBox      =   0   'False
   BeginProperty Font 
      Name            =   "新細明體-ExtB"
      Size            =   9
      Charset         =   136
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6780
   ScaleWidth      =   9024
   Begin TabDlg.SSTab SSTab1 
      Height          =   3648
      Left            =   100
      TabIndex        =   17
      Top             =   3000
      Width           =   8800
      _ExtentX        =   15515
      _ExtentY        =   6435
      _Version        =   393216
      Tabs            =   4
      Tab             =   1
      TabsPerRow      =   4
      TabHeight       =   420
      TabCaption(0)   =   "結案原因"
      TabPicture(0)   =   "frm210133_F.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame1"
      Tab(0).Control(1)=   "Frame10"
      Tab(0).ControlCount=   2
      TabCaption(1)   =   "請款項目"
      TabPicture(1)   =   "frm210133_F.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label1(17)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1(18)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label1(19)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "Label1(20)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label1(21)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblName(0)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "lblName(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblName(2)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "lblName(3)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "Label1(26)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Label1(13)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Line2"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "txtA1K(1)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "FrameCRC"
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
      Tab(1).Control(21)=   "FrameAmt"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "Grid70X"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).ControlCount=   23
      TabCaption(2)   =   "聯絡項目"
      TabPicture(2)   =   "frm210133_F.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(14)"
      Tab(2).Control(1)=   "txtNotPay"
      Tab(2).Control(2)=   "txtFCPMemo"
      Tab(2).Control(3)=   "lblNotPayCPN"
      Tab(2).Control(4)=   "Label1(27)"
      Tab(2).Control(5)=   "lblNotPay_CP"
      Tab(2).Control(6)=   "Frame9"
      Tab(2).Control(7)=   "Frame6"
      Tab(2).Control(8)=   "Chk6(1)"
      Tab(2).Control(9)=   "Chk6(2)"
      Tab(2).Control(10)=   "Chk6(3)"
      Tab(2).Control(11)=   "Frame7"
      Tab(2).Control(12)=   "Frame8"
      Tab(2).Control(13)=   "txtCCD08"
      Tab(2).ControlCount=   14
      TabCaption(3)   =   "結案原因"
      TabPicture(3)   =   "frm210133_F.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Frame5"
      Tab(3).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid70X 
         Height          =   996
         Left            =   5520
         TabIndex        =   138
         Top             =   2520
         Visible         =   0   'False
         Width           =   2796
         _ExtentX        =   4932
         _ExtentY        =   1757
         _Version        =   393216
         Rows            =   1
         Cols            =   6
         FixedRows       =   0
         FixedCols       =   0
         BackColorFixed  =   12632256
         FormatString    =   "順序|代號|請款項目|金額|折扣|備註"
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.Frame FrameAmt 
         BorderStyle     =   0  '沒有框線
         Enabled         =   0   'False
         Height          =   400
         Left            =   120
         TabIndex        =   131
         Top             =   1140
         Width           =   8412
         Begin VB.TextBox txtAmt 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   2
            Left            =   5880
            MaxLength       =   6
            TabIndex        =   134
            Top             =   120
            Width           =   1200
         End
         Begin VB.TextBox txtAmt 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   1
            Left            =   3480
            MaxLength       =   6
            TabIndex        =   133
            Top             =   120
            Width           =   1200
         End
         Begin VB.TextBox txtAmt 
            Alignment       =   1  '靠右對齊
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   270
            Index           =   0
            Left            =   1032
            TabIndex        =   132
            Top             =   120
            Width           =   1200
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
            Left            =   5280
            TabIndex        =   137
            Top             =   120
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
            Left            =   2880
            TabIndex        =   136
            Top             =   120
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
            Left            =   240
            TabIndex        =   135
            Top             =   120
            Width           =   804
         End
      End
      Begin VB.TextBox txtCCD08 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Left            =   -73680
         MaxLength       =   8
         TabIndex        =   33
         Text            =   "txtCCD08"
         Top             =   1250
         Width           =   850
      End
      Begin VB.CommandButton cmdInfo 
         BackColor       =   &H00FFFFC0&
         Caption         =   "目前設定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   550
         Left            =   8040
         Style           =   1  '圖片外觀
         TabIndex        =   124
         Top             =   360
         Width           =   550
      End
      Begin VB.TextBox txtA1K 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   2
         Left            =   6984
         MaxLength       =   1
         TabIndex        =   19
         Top             =   336
         Width           =   255
      End
      Begin VB.Frame Frame10 
         BorderStyle     =   0  '沒有框線
         Height          =   380
         Left            =   -74820
         TabIndex        =   129
         Top             =   220
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
            TabIndex        =   130
            Top             =   120
            Width           =   950
         End
      End
      Begin VB.Frame Frame8 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   900
         Left            =   -69120
         TabIndex        =   108
         Top             =   700
         Width           =   2400
         Begin VB.CheckBox Chk9 
            Caption         =   "閉卷請款"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   96
            TabIndex        =   52
            Top             =   300
            Width           =   2200
         End
         Begin VB.CheckBox Chk9 
            Caption         =   "請銷本所年費期限管制"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   96
            TabIndex        =   51
            Top             =   0
            Width           =   2200
         End
         Begin VB.CheckBox Chk9 
            Caption         =   "不請款"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   96
            TabIndex        =   53
            Top             =   600
            Width           =   2200
         End
      End
      Begin VB.Frame Frame7 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2650
         Left            =   -72636
         TabIndex        =   105
         Top             =   700
         Width           =   3250
         Begin VB.CheckBox Chk8 
            Caption         =   "XX"
            Height          =   255
            Index           =   18
            Left            =   120
            TabIndex        =   126
            Top             =   2280
            Visible         =   0   'False
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "需管制6個月補繳期"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   96
            TabIndex        =   44
            Top             =   2016
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "期限屆,未獲指示"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   96
            TabIndex        =   37
            Top             =   0
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "代理人指示不繳年費"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   96
            TabIndex        =   38
            Top             =   250
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "A.未獲指示"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   252
            TabIndex        =   40
            Top             =   1000
            Width           =   1200
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "透過其他管道代繳年費"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   96
            TabIndex        =   39
            Top             =   500
            Width           =   3100
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "B.已獲指示"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   1620
            TabIndex        =   41
            Top             =   1000
            Width           =   1200
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "後續准駁簡單報告"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   252
            TabIndex        =   42
            Top             =   1500
            Width           =   2000
         End
         Begin VB.CheckBox Chk8 
            Caption         =   "駁不報告"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   252
            TabIndex        =   43
            Top             =   1716
            Width           =   1200
         End
         Begin VB.Label Label1 
            Caption         =   "後續准駁簡單報告："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   12
            Left            =   96
            TabIndex        =   107
            Top             =   1296
            Width           =   1800
         End
         Begin VB.Label Label1 
            Caption         =   "不續辦請款："
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   10
            Left            =   96
            TabIndex        =   106
            Top             =   804
            Width           =   1104
         End
      End
      Begin VB.TextBox txtA1K 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   6
         Left            =   6000
         MaxLength       =   9
         TabIndex        =   23
         Top             =   910
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   5
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   22
         Text            =   "A1K27"
         Top             =   910
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   4
         Left            =   6000
         MaxLength       =   9
         TabIndex        =   21
         Top             =   600
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   3
         Left            =   1080
         MaxLength       =   9
         TabIndex        =   20
         Text            =   "A1K28"
         Top             =   600
         Width           =   850
      End
      Begin VB.TextBox txtA1K 
         BorderStyle     =   0  '沒有框線
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   1080
         MaxLength       =   1
         TabIndex        =   104
         Top             =   320
         Width           =   255
      End
      Begin VB.Frame FrameCRC 
         BackColor       =   &H00C0FFFF&
         Caption         =   "請款項目編輯區"
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
         Height          =   2000
         Left            =   120
         TabIndex        =   96
         Top             =   1580
         Width           =   8412
         Begin VB.CommandButton cmdClsPay 
            BackColor       =   &H0080FF80&
            Caption         =   "清除"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7716
            Style           =   1  '圖片外觀
            TabIndex        =   30
            Top             =   120
            Width           =   525
         End
         Begin VB.CommandButton cmdDel 
            BackColor       =   &H0080FF80&
            Caption         =   "刪除"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   7080
            Style           =   1  '圖片外觀
            TabIndex        =   29
            Top             =   120
            Width           =   525
         End
         Begin VB.CommandButton cmdAdd 
            BackColor       =   &H0080FF80&
            Caption         =   "加入"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   300
            Left            =   6414
            Style           =   1  '圖片外觀
            TabIndex        =   28
            Top             =   120
            Width           =   525
         End
         Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridAMT 
            Height          =   1152
            Left            =   60
            TabIndex        =   97
            Top             =   840
            Width           =   8196
            _ExtentX        =   14457
            _ExtentY        =   2032
            _Version        =   393216
            Cols            =   6
            FixedCols       =   0
            BackColorFixed  =   12632256
            FormatString    =   "順序|代號|請款項目|金額|折扣|備註"
            _NumberOfBands  =   1
            _Band(0).Cols   =   6
         End
         Begin MSForms.TextBox TxtPay 
            Height          =   300
            Index           =   3
            Left            =   5136
            TabIndex        =   27
            Top             =   440
            Width           =   3120
            VariousPropertyBits=   -1466941413
            ScrollBars      =   2
            Size            =   "5503;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "序號"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   228
            Index           =   9
            Left            =   60
            TabIndex        =   117
            Top             =   220
            Width           =   396
         End
         Begin VB.Label lblPayItemN 
            Appearance      =   0  '平面
            BackColor       =   &H80000005&
            Caption         =   "lblPayItemN"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   288
            Left            =   1236
            TabIndex        =   116
            Top             =   448
            Width           =   1476
         End
         Begin MSForms.TextBox TxtPay 
            Height          =   300
            Index           =   2
            Left            =   4248
            TabIndex        =   26
            Top             =   440
            Width           =   864
            VariousPropertyBits=   671107099
            Size            =   "1531;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
            ParagraphAlign  =   2
         End
         Begin VB.Label Label1 
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
            Height          =   228
            Index           =   16
            Left            =   552
            TabIndex        =   102
            Top             =   220
            Width           =   1008
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "金額"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   228
            Index           =   99
            Left            =   2784
            TabIndex        =   101
            Top             =   220
            Width           =   768
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "折扣 (%)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   228
            Index           =   100
            Left            =   4236
            TabIndex        =   100
            Top             =   220
            Width           =   804
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "備註"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   228
            Index           =   15
            Left            =   5184
            TabIndex        =   99
            Top             =   220
            Width           =   552
         End
         Begin VB.Label LblCntItem 
            Alignment       =   1  '靠右對齊
            Appearance      =   0  '平面
            BackColor       =   &H80000005&
            BackStyle       =   0  '透明
            Caption         =   "1"
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
            Height          =   288
            Left            =   60
            TabIndex        =   98
            Top             =   448
            Width           =   456
         End
         Begin MSForms.TextBox TxtPay 
            Height          =   300
            Index           =   1
            Left            =   2748
            TabIndex        =   25
            Top             =   440
            Width           =   1500
            VariousPropertyBits=   671107099
            Size            =   "2646;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
            ParagraphAlign  =   2
         End
         Begin MSForms.TextBox TxtPay 
            Height          =   300
            Index           =   0
            Left            =   552
            TabIndex        =   24
            Top             =   440
            Width           =   660
            VariousPropertyBits=   671107099
            Size            =   "1154;529"
            Value           =   "123456"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.TextBox txtA1K 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   3504
         MaxLength       =   1
         TabIndex        =   18
         Top             =   336
         Width           =   255
      End
      Begin VB.CheckBox Chk6 
         Caption         =   "未付帳款"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   3
         Left            =   -74940
         TabIndex        =   45
         Top             =   960
         Width           =   2200
      End
      Begin VB.CheckBox Chk6 
         Caption         =   "查本案前款均已付清"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   -74940
         TabIndex        =   32
         Top             =   660
         Width           =   2200
      End
      Begin VB.CheckBox Chk6 
         Caption         =   "D/N run C 類工程師報告"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74940
         TabIndex        =   31
         Top             =   360
         Width           =   2200
      End
      Begin VB.Frame Frame5 
         Caption         =   "FCP"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3055
         Left            =   -74880
         TabIndex        =   85
         Top             =   360
         Width           =   8350
         Begin VB.OptionButton Option2 
            Caption         =   "(代號11)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   96
            TabIndex        =   69
            Top             =   2700
            Width           =   3000
         End
         Begin VB.OptionButton Option2 
            Caption         =   "(代號99)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   99
            Left            =   3700
            TabIndex        =   77
            Top             =   1170
            Width           =   1400
         End
         Begin VB.OptionButton Option2 
            Caption         =   "(代號09)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   9
            Left            =   100
            TabIndex        =   67
            Top             =   1908
            Width           =   3000
         End
         Begin VB.OptionButton Option2 
            Caption         =   "(代號10)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   100
            TabIndex        =   68
            Top             =   2340
            Width           =   3000
         End
         Begin VB.OptionButton Option2 
            Caption         =   "(代號14)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   14
            Left            =   3700
            TabIndex        =   73
            Top             =   550
            Width           =   4500
         End
         Begin VB.OptionButton Option2 
            Caption         =   "(代號15)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   3700
            TabIndex        =   75
            Top             =   860
            Width           =   4500
         End
         Begin VB.OptionButton Option2 
            Caption         =   "(代號23)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   100
            TabIndex        =   62
            Top             =   240
            Width           =   3500
         End
         Begin VB.OptionButton Option2 
            Caption         =   "(代號06)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   6
            Left            =   100
            TabIndex        =   66
            Top             =   1530
            Width           =   3000
         End
         Begin VB.OptionButton Option2 
            Caption         =   "(代號05)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   100
            TabIndex        =   65
            Top             =   1170
            Width           =   3000
         End
         Begin VB.OptionButton Option2 
            Caption         =   "(代號04)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   100
            TabIndex        =   64
            Top             =   860
            Width           =   2500
         End
         Begin VB.OptionButton Option2 
            Caption         =   "(代號02)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   100
            TabIndex        =   63
            Top             =   550
            Width           =   2325
         End
         Begin VB.OptionButton Option2 
            Caption         =   "(代號13)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   13
            Left            =   3700
            TabIndex        =   71
            Top             =   240
            Width           =   4500
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "請於下方敘明理由"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   7
            Left            =   5170
            TabIndex        =   112
            Top             =   1250
            Width           =   1440
         End
         Begin MSForms.TextBox txtReason 
            Height          =   1500
            Index           =   1
            Left            =   3756
            TabIndex        =   79
            Top             =   1488
            Width           =   4500
            VariousPropertyBits=   -1466941413
            MaxLength       =   500
            ScrollBars      =   2
            Size            =   "7937;2646"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Frame Frame6 
         Height          =   3300
         Left            =   -72700
         TabIndex        =   122
         Top             =   250
         Width           =   3500
         Begin VB.CheckBox Chk7 
            Caption         =   "907 不續辦"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   50
            TabIndex        =   35
            Top             =   120
            Width           =   1150
         End
         Begin MSForms.ComboBox CboState 
            Height          =   288
            Index           =   0
            Left            =   1250
            TabIndex        =   36
            Top             =   130
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
      Begin VB.Frame Frame9 
         Height          =   1520
         Left            =   -69130
         TabIndex        =   123
         Top             =   250
         Width           =   2850
         Begin VB.CheckBox Chk7 
            Caption         =   "913 閉卷"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   50
            TabIndex        =   49
            Top             =   120
            Width           =   1050
         End
         Begin MSForms.ComboBox CboState 
            Height          =   288
            Index           =   1
            Left            =   1150
            TabIndex        =   50
            Top             =   130
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
      Begin VB.Frame Frame1 
         Caption         =   "FCT"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3055
         Left            =   -74880
         TabIndex        =   84
         Top             =   360
         Width           =   8350
         Begin VB.OptionButton Option1 
            Caption         =   "已移轉，非本所辦理 "
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   96
            TabIndex        =   12
            Top             =   2700
            Width           =   3000
         End
         Begin VB.OptionButton Option1 
            Caption         =   "客戶指示不延展 (代號16)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   100
            TabIndex        =   6
            Top             =   550
            Width           =   2325
         End
         Begin VB.OptionButton Option1 
            Caption         =   "已轉由他所續辦 (代號02)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   100
            TabIndex        =   7
            Top             =   860
            Width           =   2500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "客戶已另案重提 (代號04)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   4
            Left            =   100
            TabIndex        =   8
            Top             =   1200
            Width           =   3000
         End
         Begin VB.OptionButton Option1 
            Caption         =   "客戶已遷移，無法連絡 (代號05)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   5
            Left            =   100
            TabIndex        =   9
            Top             =   1530
            Width           =   3000
         End
         Begin VB.OptionButton Option1 
            Caption         =   "客戶指示放棄本案 (代號15)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   100
            TabIndex        =   5
            Top             =   250
            Width           =   2500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "他所延展未變更代理人(管制下次延展期限) (代號19)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   3700
            TabIndex        =   14
            Top             =   550
            Width           =   4500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "客戶指示僅不續行此程序，待審查結果。"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   3700
            TabIndex        =   13
            Top             =   250
            Width           =   4500
         End
         Begin VB.OptionButton Option1 
            Caption         =   "由年費公司辦理延展"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   100
            TabIndex        =   11
            Top             =   2340
            Width           =   3000
         End
         Begin VB.OptionButton Option1 
            Caption         =   "客戶自行處理 (代號10)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   10
            Left            =   100
            TabIndex        =   10
            Top             =   1908
            Width           =   3000
         End
         Begin VB.OptionButton Option1 
            Caption         =   "其他 (代號99)"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   99
            Left            =   3700
            TabIndex        =   15
            Top             =   860
            Width           =   1400
         End
         Begin MSForms.TextBox txtReason 
            Height          =   1800
            Index           =   0
            Left            =   3750
            TabIndex        =   16
            Top             =   1200
            Width           =   4500
            VariousPropertyBits=   -1466941413
            ScrollBars      =   2
            Size            =   "7937;3175"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "請於下方敘明理由"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   180
            Index           =   5
            Left            =   5160
            TabIndex        =   111
            Top             =   936
            Width           =   1440
         End
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   240
         X2              =   8450
         Y1              =   1550
         Y2              =   1550
      End
      Begin VB.Label lblNotPay_CP 
         Caption         =   "該案未請款:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   430
         Left            =   -74880
         TabIndex        =   128
         Top             =   3120
         Width           =   2100
      End
      Begin VB.Label Label1 
         Caption         =   "該案未請款:"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   27
         Left            =   -74940
         TabIndex        =   127
         Top             =   2880
         Width           =   1236
      End
      Begin VB.Label lblNotPayCPN 
         Caption         =   "     管制催款日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   -74940
         TabIndex        =   125
         Top             =   1250
         Width           =   1230
      End
      Begin VB.Label Label1 
         Caption         =   "合併列印請款單：       (要印:Y)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   13
         Left            =   5520
         TabIndex        =   118
         Top             =   360
         Width           =   2550
      End
      Begin MSForms.TextBox txtFCPMemo 
         Height          =   1404
         Left            =   -69050
         TabIndex        =   60
         Top             =   2160
         Width           =   2780
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "4904;2476"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtNotPay 
         Height          =   1300
         Left            =   -74940
         TabIndex        =   34
         Top             =   1530
         Width           =   2196
         VariousPropertyBits=   -1466941413
         ScrollBars      =   2
         Size            =   "3873;2293"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "其他說明："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   14
         Left            =   -69130
         TabIndex        =   95
         Top             =   1920
         Width           =   996
      End
      Begin VB.Label Label1 
         Caption         =   "帳款已清：         (已清:Y)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   26
         Left            =   120
         TabIndex        =   103
         Top             =   360
         Width           =   2100
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(3)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   6960
         TabIndex        =   94
         Top             =   950
         Width           =   1500
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(2)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   2040
         TabIndex        =   93
         Top             =   950
         Width           =   2496
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(1)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   6960
         TabIndex        =   92
         Top             =   636
         Width           =   1500
      End
      Begin VB.Label lblName 
         Caption         =   "lblName(0)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   0
         Left            =   2040
         TabIndex        =   91
         Top             =   612
         Width           =   2496
      End
      Begin VB.Label Label1 
         Caption         =   "固定列印對象："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   21
         Left            =   4680
         TabIndex        =   90
         Top             =   950
         Width           =   1296
      End
      Begin VB.Label Label1 
         Caption         =   "請款對象："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   20
         Left            =   120
         TabIndex        =   89
         Top             =   636
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "列印對象："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   19
         Left            =   120
         TabIndex        =   88
         Top             =   950
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "固定請款對象："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   18
         Left            =   4680
         TabIndex        =   87
         Top             =   636
         Width           =   1300
      End
      Begin VB.Label Label1 
         Caption         =   "列印申請人：       (要印:Y / 改不印:N)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   17
         Left            =   2400
         TabIndex        =   86
         Top             =   360
         Width           =   3000
      End
   End
   Begin VB.CommandButton cmdFile 
      BackColor       =   &H00C0C0FF&
      Caption         =   "檢視/新增 電子檔"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   6350
      Style           =   1  '圖片外觀
      TabIndex        =   119
      Top             =   30
      Width           =   1600
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   7980
      TabIndex        =   120
      Top             =   30
      Width           =   975
   End
   Begin VB.Frame Frame4 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame4"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5010
      TabIndex        =   74
      Top             =   30
      Width           =   1815
      Begin VB.CommandButton cmdSend 
         Caption         =   "送出(&E)"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   0
         TabIndex        =   114
         Top             =   0
         Width           =   975
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   5010
      TabIndex        =   59
      Top             =   30
      Width           =   1485
      Begin VB.TextBox txtPCnt 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   0
         Left            =   750
         MaxLength       =   1
         TabIndex        =   61
         Text            =   "2"
         Top             =   30
         Width           =   270
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "列印(&P)　　份"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   47
         Top             =   0
         Width           =   1395
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "查詢(&S)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   3990
      TabIndex        =   4
      Top             =   30
      Width           =   975
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   3
      Left            =   2970
      MaxLength       =   2
      TabIndex        =   3
      Top             =   405
      Width           =   375
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   2
      Left            =   2640
      MaxLength       =   1
      TabIndex        =   2
      Top             =   405
      Width           =   270
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   1
      Left            =   1740
      MaxLength       =   6
      TabIndex        =   1
      Top             =   405
      Width           =   825
   End
   Begin VB.TextBox txt1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   270
      Index           =   0
      Left            =   1140
      MaxLength       =   3
      TabIndex        =   0
      Top             =   405
      Width           =   525
   End
   Begin VB.Frame Frame3 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame3"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   5010
      TabIndex        =   70
      Top             =   30
      Width           =   1875
      Begin VB.TextBox txtPCnt 
         Alignment       =   1  '靠右對齊
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   270
         Index           =   1
         Left            =   1230
         MaxLength       =   1
         TabIndex        =   72
         Text            =   "2"
         Top             =   30
         Width           =   270
      End
      Begin VB.CommandButton cmdok 
         Caption         =   "無期限閉卷(&P)　　份"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   0
         Style           =   1  '圖片外觀
         TabIndex        =   46
         Top             =   0
         Width           =   1815
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHGrid1 
      Height          =   1390
      Left            =   150
      TabIndex        =   83
      Top             =   1500
      Width           =   8692
      _ExtentX        =   15325
      _ExtentY        =   2455
      _Version        =   393216
      BackColor       =   -2147483624
      Cols            =   1
      FixedCols       =   0
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
      _Band(0).Cols   =   1
   End
   Begin MSForms.Label lblCU01Nm 
      Height          =   252
      Left            =   1160
      TabIndex        =   121
      Top             =   1260
      Width           =   5000
      Size            =   "8819;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "國　　籍："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   8
      Left            =   6216
      TabIndex        =   115
      Top             =   1260
      Width           =   900
   End
   Begin VB.Label lblCuNation 
      Caption         =   "lblCuNation"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7160
      TabIndex        =   113
      Top             =   1260
      Width           =   1704
   End
   Begin VB.Label lblApplyNation 
      Caption         =   "lblApplyNm"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7160
      TabIndex        =   110
      Top             =   740
      Width           =   1700
   End
   Begin MSForms.Label lblFCAgNm 
      Height          =   252
      Left            =   1160
      TabIndex        =   109
      Top             =   1000
      Width           =   5000
      Caption         =   "Y2776600"
      Size            =   "7117;450"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label lblFCAgNation 
      Caption         =   "lblFCAgNation"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7160
      TabIndex        =   82
      Top             =   996
      Width           =   1700
   End
   Begin VB.Label Label1 
      Caption         =   "國　　籍："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   3
      Left            =   6216
      TabIndex        =   81
      Top             =   996
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "代  理  人："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   216
      TabIndex        =   80
      Top             =   1000
      Width           =   996
   End
   Begin MSForms.Label lblCaseNm 
      Height          =   252
      Left            =   1160
      TabIndex        =   78
      Top             =   740
      Width           =   5000
      VariousPropertyBits=   27
      Size            =   "8819;444"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "label2"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   180
      Left            =   210
      TabIndex        =   76
      Top             =   0
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "案件名稱："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   11
      Left            =   210
      TabIndex        =   58
      Top             =   740
      Width           =   900
   End
   Begin VB.Label Label1 
      Caption         =   "申  請  人："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   216
      TabIndex        =   57
      Top             =   1260
      Width           =   996
   End
   Begin VB.Label Label1 
      Caption         =   "申請案號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   4
      Left            =   6216
      TabIndex        =   56
      Top             =   456
      Width           =   900
   End
   Begin VB.Label lblApplyCaseNo 
      Caption         =   "lblApplyCaseNo"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   7160
      TabIndex        =   55
      Top             =   456
      Width           =   1700
   End
   Begin VB.Label Label1 
      Caption         =   "申請國家："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   6
      Left            =   6216
      TabIndex        =   54
      Top             =   740
      Width           =   900
   End
   Begin VB.Line Line1 
      BorderWidth     =   2
      X1              =   1470
      X2              =   3030
      Y1              =   540
      Y2              =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "本所案號："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   180
      Index           =   0
      Left            =   210
      TabIndex        =   48
      Top             =   450
      Width           =   900
   End
End
Attribute VB_Name = "frm210133_F"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2025/04/10
Option Explicit

Public m_strSaveFiles As String '新增附件
Public m_CCM01 As String, m_NP01 As String, m_NP22 As String '結案單編號/總收文號/下一程序序號
Public m_strIR01 As String, m_strIR02 As String, m_strIR03 As String, m_strIR04 As String 'FCP系統收件區用

Dim m_AttachPath As String '附件暫存區路徑
Dim IsSetGirdColOK As Boolean, arrCol() As String, arrWidth() As String, arrAMTCol() As String, arrAMTWidth() As String
Dim m_PrevForm As Form '前一畫面
Dim i As Integer, j As Integer
Dim m_SetFlowEmp1 As String '設定簽核人員1
Dim strNP02 As String, strNP03 As String, strNP04 As String, strNP05 As String, m_NP07 As String
Dim strPreCase1 As String, strPreCase2 As String, strPreCase3 As String, strPreCase4 As String '前一筆案號
Dim m_CuNo1 As String, m_CU10 As String, m_CP13 As String '申請人1/申請人1國籍編號/客戶檔智權人員
Dim m_CuNo2 As String, m_CuNo3 As String, m_CuNo4 As String, m_CuNo5 As String '申請人2-5
Dim m_FCAgNo As String, m_FA10 As String, F_CCM01 As String  '代理人編號/代理人國籍編號/第1筆結案單號
Dim b_CuNo1 As String, b_CuNo2 As String, b_CuNo3 As String, b_CuNo4 As String, b_CuNo5 As String, b_FCAgNo As String '前一筆申請人1-5/前一筆代理人編號
Dim m_ApplyNA01 As String, m_IsClose As String, m_PA08 As String, m_PA10 As String '申請國家編號/是否已閉卷/專利種類/申請日
Dim m_CCM02 As String, m_CCM03 As String, m_CCM04 As String, m_CCM17 As String '總收文號或本所案號/下一程序序號/結案理由/外來郵件key
Dim SignPerson As String, m_F0202_2 As String, m_F0202_3 As String, m_F0308 As String, m_F0309 As String '承辦簽核主管/程序人員/補看人員/上一處理人員/下一處理人員/目前表單狀態
Dim stOptPerson As String, m_F0316 As String '可操作人員/智權人員
Dim m_row As Integer, nCol As Integer, nRow As Integer, intState As Integer '下一程序 列/請款項目 欄/請款項目 列 /外商 or 外專
Dim bolFMPY53374 As Boolean, bolFMP As Boolean '是寰華FMP案(外專程序結)/是FMP案(內專程序結)
Dim bolIROK As Boolean, bolGoNext As Boolean '系統收件區進入且第一筆案號結案ok /回前畫面時Run下一筆
Dim stPAndCFPMemo As String 'P非寰華及CFP給內專程序備註事項
Dim bolUpdNP24 As Boolean, stCCM02_New As String, stCCM03_New As String 'Add by Amy 2025/08/25
Dim stSalesAgent As String, stSalesST52 As String, stSalesST53 As String, stSalesST5455 As String  'Modify by Amy 承辦人員職代/主管ST52/主管53/主管5455(從Pub_FCInuptCloseLimit搬過來)
Dim stGridAmtCol As String 'Add by Amy 2025/11/07 設定金額

Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub

Private Function GetF0316(strNP10 As String) As String
   Dim stMan As String 'Add by Amy 2025/10/14
   
   'Add by Amy 2025/10/14
   If stSalesST52 <> "" Then stMan = stMan & ";" & stSalesST52
   If stSalesST53 <> "" Then stMan = stMan & ";" & stSalesST53
   If stSalesST5455 <> "" Then stMan = stMan & ";" & stSalesST5455
   If stMan <> "" Then stMan = Mid(stMan, 2)
   'end 2025/10/14
   
   '可能由st52 代為操作
   'Modify by Amy 2025/10/14 若操作者為目前智權人員之主管,以操作者為主
   'If strNP10 <> strUserNum And InStr(stOptPerson, strUserNum) > 0 Then
   If strNP10 <> strUserNum And InStr(stMan, strUserNum) > 0 Then
      GetF0316 = strUserNum
   'Add by Amy 2025/10/14 操作者為其目前智權人員之職代,以目前智權人員為主
   ' ex:FCP-061468 目前智權人員為B0004(李道昀) 請假,其職代 黃賢泰代為操作,簽核主管應為李道昀的主管Lisa(A6035)
   'Modify by Amy 2025/10/15 +mCP13 = strUserNum,無期限結案會抓不到人員 ex:1141015 結FCP-048229
   ElseIf InStr(stSalesAgent, strUserNum) > 0 Or m_CP13 = strUserNum Then
      GetF0316 = m_CP13
   'end 2025/10/14
   Else
      GetF0316 = strNP10
   End If
   '若智權人員已離職,則以Login人員代替
   If GetF0316 <> "" And (ChkStaffST04(GetF0316, False) = True Or Left(GetF0316, 1) <= "6") Then
      GetF0316 = strUserNum
   End If
   'Memo by Amy 2025/08/27 智權為P2006(商標智權人員),於[國內結案單]新增與退回之簽核人員不一致,Sindy與秀玲討論後都抓P2006之簽核人員
   '  ex:CFT-022430 (蒲璇已操作,無FC代理人) 無期限之結案單,Amy 測式[退回]再送出,發現新增時抓蒲璇之簽核人員,退回抓P2006之簽核人員
   '       將上述案號加FC代理人,進此支會出現無權限操作,若真有此狀況發生,再視情況調整,故此支先不改
End Function

Private Sub CboState_Change(Index As Integer)
   Dim stMsg As String
   
   If CboState(Index) <> MsgText(601) And Chk7(Index).Value = vbUnchecked Then
      '勾選彈過,再選下拉選單,只要彈一次
      If UCase(Me.ActiveControl.Name) = UCase("CboState") And Index = 1 Then
         '勾選 閉卷,顯示關連案
         Screen.MousePointer = vbHourglass
         If ChkCaseRelation(0, Me.Name, strNP02, strNP03, strNP04, strNP05, stMsg) = True Then
            MsgBox stMsg, vbInformation
         End If
         Screen.MousePointer = vbDefault
      End If
      Chk7(Index).Value = vbChecked
   End If
  
End Sub

Private Sub Chk6_Click(Index As Integer)
   If txt1(0) <> strNP02 Or txt1(1) <> strNP03 Then Exit Sub
   If UCase(Me.ActiveControl.Name) <> UCase("CHK6") Then Exit Sub
   
   '勾 D/N run C 類工程師報告
   If UCase(Me.ActiveControl.Name) = UCase("CHK6") And Me.ActiveControl.Index = 1 And Chk6(1).Value = vbChecked Then
      '由內專結,帶給內專人員確認帳單字樣
      If ((strNP02 = "P" And bolFMPY53374 = False) Or strNP02 = "CFP") Then
         If txtReason(1) = "" Then
            txtReason(1) = stPAndCFPMemo
         ElseIf InStr(txtReason(1), "本案大陸代理人是否有最終帳單") = 0 Then
            txtReason(1) = stPAndCFPMemo & vbCrLf & txtReason(1)
         End If
      End If
      '判斷無C類來函掛工程師,彈提醒 可繼續操作
      If ChkCClassCP14IsEngr(strNP02, strNP03, strNP04, strNP05) = False Then
         MsgBox "此案無C類來函掛工程師" & vbCrLf & _
                           "勾選「D/N run C 類工程師報告」" & _
                           "不會出outlook草稿!", vbInformation
      End If
   End If
End Sub

Private Sub Chk7_Click(Index As Integer)
   Dim stMsg As String
   
   If txt1(0) <> strNP02 Or txt1(1) <> strNP03 Then Exit Sub
   If UCase(Me.ActiveControl.Name) <> UCase("CHK7") Then Exit Sub
   
   Frame7.Enabled = True
   Frame8.Enabled = True
   CboState(0).Enabled = True
   CboState(1).Enabled = True
   Select Case Index
      Case 0 '不續辦
         If Chk7(Index).Value = vbChecked Then
            CboState(1) = ""
            Call ClearData("9")
            Chk7(1).Value = vbUnchecked
            CboState(1).Enabled = False
            Frame8.Enabled = False
         End If
      Case 1 '閉卷
         If Chk7(Index).Value = vbChecked Then
            CboState(0) = ""
            Call ClearData("8")
            If UCase(Me.ActiveControl.Name) = UCase("Chk7") And Index = 1 Then
               '勾選 閉卷,顯示關連案
               Screen.MousePointer = vbHourglass
               If ChkCaseRelation(0, Me.Name, strNP02, strNP03, strNP04, strNP05, stMsg) = True Then
                  MsgBox stMsg, vbInformation
               End If
               Screen.MousePointer = vbDefault
            End If
             Chk7(0).Value = vbUnchecked
             CboState(0).Enabled = False
             Frame7.Enabled = False
         End If
   End Select
End Sub

Private Sub Chk8_Click(Index As Integer)
   If txt1(0) <> strNP02 Or txt1(1) <> strNP03 Then Exit Sub
   If UCase(Me.ActiveControl.Name) <> UCase("CHK8") Then Exit Sub
   If Index < 14 Or Index > 17 Then Exit Sub
   
   Select Case Index
      Case 14, 15 '不續辦請款勾選 A.未獲指示 B.已獲指示
         If Chk8(Index).Value = vbChecked Then
            If Index = 14 Then
               Chk8(15).Value = vbUnchecked
            Else
               Chk8(14).Value = vbUnchecked
            End If
         End If
         '不續辦請款 勾選 A.未獲指示 B.已獲指示 且 由內專結
         If (UCase(Me.ActiveControl.Name) = UCase("CHK8") And (Me.ActiveControl.Index = 14 Or Me.ActiveControl.Index = 15)) _
           And Chk8(Index).Value = vbChecked And ((strNP02 = "P" And bolFMPY53374 = False) Or strNP02 = "CFP") Then
            If txtReason(1) = "" Then
               txtReason(1) = stPAndCFPMemo
            ElseIf InStr(txtReason(1), "本案大陸代理人是否有最終帳單") = 0 Then
               txtReason(1) = stPAndCFPMemo & vbCrLf & txtReason(1)
            End If
         End If
      Case 16, 17 '後續准駁簡單報告 A.後續准不報告 B.駁不報告
         If Chk8(Index).Value = vbChecked Then
            If Index = 16 Then
               Chk8(17).Value = vbUnchecked
            Else
               Chk8(16).Value = vbUnchecked
            End If
         End If
   End Select
End Sub

Private Sub Chk9_Click(Index As Integer)
   If UCase(Me.ActiveControl.Name) <> UCase("CHK9") Then Exit Sub
   '勾 閉卷請款 且 由內專結
   If UCase(Me.ActiveControl.Name) = UCase("CHK9") And Me.ActiveControl.Index = 22 _
     And Chk9(22).Value = vbChecked And ((strNP02 = "P" And bolFMPY53374 = False) Or strNP02 = "CFP") Then
      If txtReason(1) = "" Then
         txtReason(1) = stPAndCFPMemo
      ElseIf InStr(txtReason(1), "本案大陸代理人是否有最終帳單") = 0 Then
         txtReason(1) = stPAndCFPMemo & vbCrLf & txtReason(1)
      End If
   End If
End Sub

Private Sub cmdAdd_Click()
   If ChkSSTab1("A") = False Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   cmdAdd.Enabled = False
   
   If Val(LblCntItem.Caption) = 0 Then
      If GridAMT.TextMatrix(1, 1) = MsgText(601) Then
         nRow = GridAMT.Rows - 1
      Else
         nRow = GridAMT.Rows
         GridAMT.AddItem ""
      End If
   Else
      nRow = Val(LblCntItem.Caption)
   End If
   LblCntItem.Caption = nRow
   GridAMT.TextMatrix(nRow, 0) = nRow
   GridAMT.TextMatrix(nRow, 1) = TxtPay(0)
   GridAMT.TextMatrix(nRow, 2) = lblPayItemN
   GridAMT.TextMatrix(nRow, 3) = Format(TxtPay(1), "#,##0") '金額
   GridAMT.TextMatrix(nRow, 4) = TxtPay(2) '折扣
   GridAMT.TextMatrix(nRow, 5) = TxtPay(3)
   'Call RefreshGridAMTSeq '704/704 及 請款項目2碼 改順序
   Call SumGridAMT '計算 總金額/規費/點數
   Call SetGridColor(0, GridAMT, arrAMTCol) '先還原顏色
   Call SetGridColor(1, GridAMT, arrAMTCol, nRow)
   Call cmdClsPay_Click
   cmdAdd.Enabled = True
   TxtPay(0).SetFocus
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdClsPay_Click()
   Call ClearData("2.1")
   TxtPay(0).SetFocus
End Sub

Private Sub cmdDel_Click()
   Dim j As Integer
   
   If ChkSSTab1("D") = False Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   cmdDel.Enabled = False
   If Val(LblCntItem.Caption) > 0 Then
      If GridAMT.Rows - 1 = 1 Then
         GridAMT.Clear
         Call Pub_SetCloseGridAmtWidth(Me.Name, GridAMT, arrAMTCol(), arrAMTWidth)
      Else
         GridAMT.RemoveItem Val(LblCntItem.Caption)
      End If
      '重新整理順序
      For j = 1 To GridAMT.Rows - 1
         If Trim(GridAMT.TextMatrix(j, GetColVal(arrAMTCol, "請款項目"))) <> MsgText(601) Then
            GridAMT.TextMatrix(j, GetColVal(arrAMTCol, "順序")) = j
         End If
      Next j
   End If
   Call SumGridAMT '計算 總金額/規費/點數
   Call cmdClsPay_Click '清除
   cmdDel.Enabled = True
   TxtPay(0).SetFocus
   Screen.MousePointer = vbDefault
End Sub

Private Sub cmdFile_Click()
   
   If ChkForm("F") = False Then
      Exit Sub
   End If
   If intState = 0 Then Exit Sub
   
   Call frm090801_8.SetParent(Me)
   frm090801_8.intFCState = intState
   frm090801_8.m_strSaveFiles = Me.m_strSaveFiles
   frm090801_8.lblCaseNo = strNP02 & "-" & strNP03 & "-" & strNP04 & "-" & strNP05
   frm090801_8.Show vbModal
End Sub

Private Sub cmdInfo_Click()
   Call Pub_CloseShowfrm210133_INV(1, Me.Name, strNP02 & "-" & strNP03 & "-" & strNP04 & "-" & strNP05)
End Sub

Private Sub cmdok_Click(Index As Integer)
   Dim strMsg As String, IsOpen090201_2_1 As Boolean
   
   Select Case Index
      Case 2 '查詢
         Screen.MousePointer = vbHourglass
         MSHGrid1.MousePointer = flexHourglass
         If doQuery(strMsg) = False Then
            MSHGrid1.MousePointer = flexDefault
            Screen.MousePointer = vbDefault
            Exit Sub
         End If
         MSHGrid1.MousePointer = flexDefault
         Screen.MousePointer = vbDefault
         
      Case 0
         '檢查條件
         
      Case 1
         '檢查條件
        
      Case 3 '結束
         Unload Me
   Case Else
   End Select
End Sub

'清除欄位值
'stChoose:0-全部(全清)/1-查詢/2-請款項目/2.1 GridAMT 相關/3-聯絡項目
'                  4-只清部份變數
'                  8-Chk8不勾選/9-Chk9不勾選
Private Sub ClearData(stChoose As String)
   Dim obj As Object
   
   If stChoose = "0" Or stChoose = "1" Then
      ChkClose.Value = vbUnchecked 'FCT閉卷
      'FCT結案原因點選項目
      For Each obj In Option1
         obj.Value = False
         If stChoose = "0" Then obj.Caption = ""
      Next
      txtReason(0) = ""
      'FCP結案原因點選項目
      For Each obj In Option2
         obj.Value = False
         If stChoose = "0" Then obj.Caption = ""
      Next
      txtReason(1) = ""
   End If
    
   If stChoose = "0" Or stChoose = "1" Or stChoose = "4" Then
      '系統收件區進入再輸下一筆時,要保留附件
      If TypeName(m_PrevForm) = "Nothing" Then
         m_strSaveFiles = "" '附件
      End If
      '代理人
      m_FCAgNo = "" 'FC代理人編號
      lblFCAgNm = ""  'FC代理人名稱
      m_FA10 = ""  'FC代理人國籍編號
      lblFCAgNation = "" 'FC代理人國籍名稱
      '申請人
      m_CuNo1 = "" '申請人1編號
      lblCU01Nm.Caption = "" '申請人1名稱
      m_CU10 = "" '申請人1國籍編號
      lblCuNation = "" '申請人1國籍名稱
      m_CuNo2 = ""
      m_CuNo3 = ""
      m_CuNo4 = ""
      m_CuNo5 = ""
         
      m_CCM02 = "" '總收文號 or 本所案號
      m_CCM03 = "" '下一程序號
      m_CCM04 = "" '結案理由代號
      m_NP07 = "" '案件性質
      SignPerson = "" '承辦簽核主管
      m_F0202_2 = "" '程序人員
      m_F0202_3 = "" '補看人員
      m_F0308 = "" '下一處理人員
      m_F0309 = "" '目前表單狀態
      m_F0316 = "" '智權人員
      
      lblApplyCaseNo.Caption = "" '申請案號
      lblCaseNm.Caption = "" '案件名稱
      m_ApplyNA01 = "" '申請國家 編號
      lblApplyNation.Caption = "" '申請國家名稱
      m_IsClose = "" '是否已閉卷
      m_PA08 = "" '專利種類
      m_PA10 = ""  '申請日
      
      m_CP13 = "" '智權人員
      If stChoose = "1" Then
         GridAMT.Clear
         Call Pub_SetCloseGridAmtWidth(Me.Name, GridAMT, arrAMTCol(), arrAMTWidth)
      End If
   End If
   
'*** 請款項目 頁籤 ***
   If stChoose = "0" Or stChoose = "1" Or stChoose = "2" Then
      For Each obj In txtA1K
         obj.Text = ""
         obj.BackColor = &H80000005 'Add by Amy 2025/10/02
      Next
      For Each obj In txtAmt
         obj.Text = ""
      Next
      For Each obj In lblName
         obj.Caption = ""
      Next
   End If
   
   'GridAMT 相關
   If stChoose = "0" Or stChoose = "1" Or stChoose = "2.1" Then
      LblCntItem.Caption = "" '序號
      For Each obj In TxtPay
         obj.Text = ""
         obj.Tag = "" '不確定要不要用
      Next
      lblPayItemN = "" '請款項目名稱
   End If
   
'*** 聯絡項目 頁籤 ***
   If stChoose = "0" Or stChoose = "1" Or stChoose = "3" Then
      Chk7(0).Value = vbUnchecked 'FCP不續辦
      Chk7(1).Value = vbUnchecked 'FCP閉卷
      If stChoose = "1" Then
         CboState(0) = ""
         CboState(1) = ""
      End If
      For Each obj In Chk6
         obj.Value = vbUnchecked
         If stChoose = "0" Then obj.Caption = ""
      Next
      lblNotPay_CP.Caption = "" 'Add by Amy 2025/08/08 發文未請款
      txtCCD08 = "" 'Add by Amy 2025/08/08 管制催款日
      txtNotPay = "" '未付帳款
      txtFCPMemo = "" '其他說明
   End If
   
   If stChoose = "0" Or stChoose = "1" Or stChoose = "8" Then
      '不續辦 勾選項目
      For Each obj In Chk8
         obj.Value = vbUnchecked
         If stChoose = "0" Then obj.Caption = ""
      Next
   End If
   If stChoose = "0" Or stChoose = "1" Or stChoose = "9" Then
      '閉卷 勾選項目
      For Each obj In Chk9
         obj.Value = vbUnchecked
         If stChoose = "0" Then obj.Caption = ""
      Next
   End If
  
End Sub

Public Function doQuery(ByRef stMsg As String, Optional ByVal stFormN As String, Optional ByVal stCaseNo1 As String, Optional ByVal stCaseNo2 As String, Optional ByVal stCaseNo3 As String, Optional ByVal stCaseNo4 As String _
  , Optional ByVal pCRecNo As String, Optional ByVal stFileN As String) As Boolean
   Dim stAttach As String, lngRec As Long, bContinue As Boolean, nxtFrm As Form
   Dim stBackData As String, stNotPay As String, stNotPay_CP As String  '判斷前一筆代理人與申請人訊息 /未付帳款 /已發文未請款(1140807 加)
   Dim o_A1K04 As String, o_A1K27 As String, o_A1K28 As String, o_A1K29 As String, o_TM56 As String, o_TM69 As String '目前系統設定
   
On Error GoTo ErrHnd

   doQuery = False
   
   If stFormN = "" And TypeName(m_PrevForm) <> "Nothing" Then
      '系統收件區第2筆,查詢不會有stFormN
      stFormN = m_PrevForm.Name
   End If
   If (UCase(stFormN) = "FRM06010612" Or UCase(stFormN) = "FRM06010616") And bolIROK = False Then
      bolGoNext = False '回前畫面時是否Run下一筆
      m_CCM02 = pCRecNo
      txt1(0) = stCaseNo1: txt1(1) = stCaseNo2: txt1(2) = stCaseNo3: txt1(3) = stCaseNo4
      txt1(0).Enabled = False: txt1(1).Enabled = False: txt1(2).Enabled = False: txt1(3).Enabled = False
      cmdOK(2).Visible = False
      Me.Caption = Me.Caption & "（信件編號:" & m_strIR01 & "-" & m_strIR03 & "）"
   End If
   
   Screen.MousePointer = vbHourglass
   MSHGrid1.MousePointer = flexHourglass
   GridAMT.Clear
   Call Pub_SetCloseGridAmtWidth(Me.Name, GridAMT, arrAMTCol(), arrAMTWidth(), Not IsSetGirdColOK)
   MSHGrid1.Clear
   Call SetGridWidth
   
   '[非]系統收件區進入再輸下一筆,才全清資料
   If bolGoNext = False Then
      Call ClearData("1")
   End If

   bolFMPY53374 = False: bolFMP = False
   m_CCM01 = ""
   If ChkForm("Q") = False Then
      MSHGrid1.MousePointer = flexDefault
      Screen.MousePointer = vbDefault
      If UCase(stFormN) = "FRM06010612" Or UCase(stFormN) = "FRM06010616" Then
         '第1筆才直接回前畫面
         If bolGoNext = False Then Unload Me
      ElseIf UCase(stFormN) = "FRM210133_2" Then
         Unload Me
      Else
         txt1(0).SetFocus
      End If
      Exit Function
   End If
   If (UCase(stFormN) = "FRM06010612" Or UCase(stFormN) = "FRM06010616") And bolGoNext = False Then
      If Pub_SetFilePathDelTmp("Close", 1, strExc(9), m_AttachPath) = False Then
         stMsg = "附件資料夾建立失敗" & vbCrLf & _
                        strExc(9) & vbCrLf & "請洽電腦中心!"
         Exit Function
      End If
      '下載信件檔
      stAttach = m_AttachPath & "\" & Mid(stFileN, InStrRev(stFileN, "\") + 1)
      If PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, "", , stAttach, True, stFileN, , Me.Name) = False Then
         stAttach = ""
         Exit Function
      End If
      
      m_strSaveFiles = stAttach
      m_CCM17 = m_strIR01 & m_strIR02 & "." & m_strIR03
   End If
   
   '系統收件區進入再輸下一筆,比對代理人及申請人資料彈訊息
   'Modify by Amy 2025/08/28 +Pub_StrUserSt03 只有外專有
   If bolGoNext = True And Left(Pub_StrUserSt03, 2) = "F2" Then
      strExc(9) = b_CuNo1 & ";" & b_CuNo2 & ";" & b_CuNo3 & ";" & b_CuNo4 & ";" & b_CuNo5
      If ChkAgentAndApplySame(0, 1, Me.Name, UCase(txt1(0)), txt1(1), txt1(2), txt1(3), strPreCase1, strPreCase2, strPreCase3, _
                                                            strPreCase4, stBackData, b_FCAgNo, strExc(9)) = False Then
         If IsNumeric(stBackData) = True Then
            If Val(stBackData) = vbNo Then
               Exit Function
            End If
         ElseIf InStr(stBackData, "無此案號資料") > 0 Or InStr(stBackData, "請洽電腦中心") > 0 Then
            MsgBox stBackData, vbCritical, "操作錯誤！"
            Exit Function
         End If
      End If
   End If
   
   strNP02 = UCase(txt1(0))
   strNP03 = txt1(1)
   strNP04 = txt1(2)
   strNP05 = txt1(3)
   intState = 0
   'Modify by Amy 2025/08/18 +CFC
   If strNP02 = "FCT" Or strNP02 = "T" Or strNP02 = "S" Or strNP02 = "CFT" Or strNP02 = "CFC" Then
      intState = 1
   ElseIf strNP02 = "FCP" Or strNP02 = "P" Or strNP02 = "FG" Or strNP02 = "CFP" Then
      intState = 2
   End If
  
   '下一程序資料 (先抓到相關資料才可判斷是否已有結案單)
   bContinue = doQueryNP(strNP02, strNP03, strNP04, strNP05, stMsg)
   If bContinue = False Then
   Else
      '[無]期限閉卷 or 系統收件區進入 判斷是否已有結案單
      If Frame3.Visible = True Or m_CCM17 <> "" Then
         strExc(7) = "": strExc(8) = ""
         If m_CCM17 <> "" And MSHGrid1.Rows = 2 And MSHGrid1.TextMatrix(1, GetColVal(arrCol, "V")) = "V" Then
            strExc(7) = MSHGrid1.TextMatrix(1, GetColVal(arrCol, "NP01"))
            strExc(8) = MSHGrid1.TextMatrix(1, GetColVal(arrCol, "NP22"))
         End If
         If ChkFlowFormExists(Flow_結案單, strExc(7), strExc(8), strNP02, strNP03, strNP04, strNP05, , , , m_CCM17) = True Then
            bContinue = False
            stMsg = "此結案單已存在,不可重覆作業！"
         End If
      End If
   End If
   If bContinue = False Then
      MSHGrid1.MousePointer = flexDefault
      Screen.MousePointer = vbDefault
      If UCase(stFormN) = "FRM06010612" Or UCase(stFormN) = "FRM06010616" Then
         '第1筆才直接回前畫面
         If bolGoNext = False Then Unload Me
      ElseIf UCase(stFormN) = "FRM210133_2" Then
         Unload Me
      Else
         MsgBox stMsg, vbInformation
         txt1(0).SetFocus
      End If
      Exit Function
   End If
   
   '可操作結案單權限 (確定案子有資料,再判斷權限)
   '外商 由案號輸入進入,於上方已先檢查權限
   If UCase(stFormN) <> "FRM210133_2" Then
      If Pub_FCInuptCloseLimit(m_CP13, strNP02, strNP03, strNP04, strNP05, bolFMPY53374, bolFMP) = False Then
         MsgBox "無權限操作此案號", vbCritical, "操作錯誤！"
         Exit Function
      End If
   End If
   
   If stMsg <> "" Then
      'doQueryNP 回傳若為[無期限]需彈訊息
      MsgBox stMsg, vbInformation
   End If
   
   stMsg = ""
   MSHGrid1.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
   
   '外專 無前畫面 Or 由系統收件區進入者,才需抓系統資料,否則以 doQueryCloseCase 抓的資料為主
   If stFormN = "" Or UCase(stFormN) = UCase("frm06010616") Then
      txtNotPay = ""
      'Add by Amy 2025/08/08
      lblNotPay_CP.Caption = ""
      txtCCD08 = ""
      'end 2025/08/08
      If strNP02 = "FCP" Or strNP02 = "P" Or strNP02 = "FG" Or strNP02 = "CFP" Then
         '未付帳款 欄預帶 未付款+發文未請款的案件性質 資料
         'Modify by Amy 2025/08/08 未付帳款 及 管制催款日 都有資料,於程序操作完 解除期限/閉卷 後加行事曆(顯示 未付帳款),故將欄位拆開
         '未付帳款
         Call Pub_GetCloseA1KData(2, Me.Name, strNP02, strNP03, strNP04, strNP05, stNotPay)
         If stNotPay <> "" Then
            txtNotPay = "追蹤欠款：" & vbCrLf & stNotPay
            Call txtNotPay_Validate(False)
         End If
         '發文未請款的案件性質
         stNotPay_CP = GetNotPayCP10(2, Me.Name, strNP02, strNP03, strNP04, strNP05)
'         If txtNotPay <> "" And stNotPay_CP <> "" Then
'            txtNotPay = txtNotPay & vbCrLf & vbCrLf & "該案尚有已發文未請款：" & vbCrLf & stNotPay_CP
'         ElseIf stNotPay_CP <> "" Then
'            txtNotPay = "該案尚有已發文未請款：" & vbCrLf & stNotPay_CP
'         End If
         If Trim(txtNotPay) <> "" Then Chk6(3).Value = vbChecked
         lblNotPay_CP.Caption = stNotPay_CP
         'end 2025/08/08
      End If
   '外商 由案號輸入 or 系統收件區進入者,才需抓系統資料,否則以 doQueryCloseCase 抓的資料為主
   ElseIf UCase(stFormN) = UCase("frm210133_2") Or UCase(stFormN) = UCase("frm06010612") Then
      'Modify by Amy 2025/10/02 經理測式後覺得應該預帶 請款資料,並將其函數改成共用,避免有未改到
      'Call SetA1KData(0) 'Mark by Amy 2025/08/18 Sindy:先不預帶 請款資料
      txtA1K(0).BackColor = &H80000005
      Call Pub_CloseSetA1KDataColor(9, Me.Name, strNP02, strNP03, strNP04, strNP05, m_NP07, Me.txtA1K, 0, 6)
      Call txtA1K_Validate(3, False) '請款對象
      Call txtA1K_Validate(4, False) '固定請款對象
      Call txtA1K_Validate(5, False) '列印對象
      Call txtA1K_Validate(6, False) '固定列印對象
      'end 2025/10/02
      Call SetLock(1, False)
      
      '案件未付帳款 彈訊息
      Set nxtFrm = Forms(0).GetForm("frm210133_INV")
      nxtFrm.intQuery = 1
      Call nxtFrm.doQuery(strNP02, strNP03, strNP04, strNP05, lngRec)
      If lngRec = 0 Then
         'Add by Amy 2025/08/18 抓操作結案單當下之帳款狀態,不需像外專把資料存下,只需彈訊息和上Y-秀玲
         txtA1K(0) = "Y"
         txtA1K(0).BackColor = &HC0C0FF
         MsgBox "帳款已結清", vbInformation
      Else
         nxtFrm.Show vbModal
      End If
      nxtFrm.intQuery = Empty
      Set nxtFrm = Nothing
   End If
 
   If (UCase(stFormN) = "FRM06010612" Or UCase(stFormN) = "FRM06010616") Then
      bolGoNext = True
   End If
   doQuery = True
   Exit Function
   
ErrHnd:
   If Err.Number <> 0 Then
      stMsg = Err.Description
   End If
End Function

'抓結案單資料
Public Function doQueryCloseCase(ByRef stMsg As String) As Boolean
   Dim rsA As New ADODB.Recordset, intA As Integer
   Dim stA_Fix  As String, sta As String, stCCM07 As String, stCCM07N As String, stNotPay As String, stMemo As String
   Dim stA1K04 As String, stA1K27 As String, stA1K28 As String, stA1K29 As String
   Dim stNotPay_CP As String, stCCD08 As String 'Add by Amy 2025/08/08
   Dim stMergeBill As String, stCCM23 As String, stCCM24 As String, stCCM25 As String 'Add by Amy 2025/08/18
   
On Error GoTo ErrHnd
   doQueryCloseCase = False
   cmdOK(2).Visible = False
  
   '有結案單號
   'Modify by Amy 2025/08/01 下一程序同一總收文號(NP01)會有兩筆,但案號不同,大陸一案兩請會有此狀況,故國外結案單也加串NP,避免抓錯
   '     ex:國內結案單 CB4027634(IDS) CFP-034735,在主檔此文號是P案(P-134557) ,下一程序CB4027634 會有兩筆但案號不同
   stA_Fix = "Select CloseCaseMain.*,Flow003.*,cp09,cp01,cp02,cp03,cp04 From CloseCaseMain,Flow003,CaseProgress" & _
                        " Where CCM01='" & m_CCM01 & "' And CCM01=F0301(+) And CCM03 is null "
                        
   sta = stA_Fix & " And length(CCM02)=9 And CCM02=cp09(+)"
   sta = sta & " Union " & Replace(Replace(UCase(stA_Fix), "SELECT ", "SELECT Distinct "), "FLOW003.*,CP09", "FLOW003.*,'' as CP09") & " And length(CCM02)<>9 " & _
                              " And CP01(+)=SubStr(CCM02, 1, length(CCM02) - 9) And CP02(+)=SubStr(CCM02, length(CCM02)- 8, 6)" & _
                              " And CP03(+)=SubStr(CCM02, length(CCM02)- 2,1) And CP04(+)=SubStr(CCM02, length(CCM02)- 1,length(CCM02)) " & _
                        " Union Select CloseCaseMain.*,Flow003.*,np01,np02,np03,np04,np05 From CloseCaseMain,Flow003,NextProgress" & _
                              " Where CCM01='" & m_CCM01 & "' And CCM02=np01(+) And CCM03=np22(+) And CCM01=F0301(+) And CCM03 is not null "
   'end 2025/08/01
   intA = 1
   Set rsA = ClsLawReadRstMsg(intA, sta)
   If intA = 1 Then
      strNP02 = rsA.Fields("cp01"): txt1(0) = strNP02
      strNP03 = rsA.Fields("cp02"): txt1(1) = strNP03
      strNP04 = rsA.Fields("cp03"): txt1(2) = strNP04
      strNP05 = rsA.Fields("cp04"): txt1(3) = strNP05
      m_CCM02 = "" & rsA.Fields("CCM02") '總收文號 or 本所案號
      m_CCM03 = "" & rsA.Fields("CCM02") '下一程序號
      stCCM07 = "" & rsA.Fields("CCM07") 'FCP狀態
      If stCCM07 <> MsgText(601) Then
         stCCM07N = Pub_SetCloseCboState(TypeName(m_PrevForm), , , " And AC02='" & stCCM07 & "'")
      End If
      m_CCM17 = "" & rsA.Fields("CCM17") '外來郵件編號 ii01+ii02+.ii03
      intState = 1
      If strNP02 = "FCP" Or strNP02 = "P" Or strNP02 = "FG" Or strNP02 = "CFP" Then
         intState = 2
      Else
         txtAmt(0) = "" & rsA.Fields("CCM08") '總金額
         txtAmt(1) = "" & rsA.Fields("CCM09") '規費
         txtAmt(2) = "" & rsA.Fields("CCM10") '點數
         'Add by Amy 2025/08/18 +請款項目資料
         stMergeBill = "" & rsA.Fields("CCM19") '合併列印請款單
         stA1K04 = "" & rsA.Fields("CCM20") '列印申請人
         stA1K28 = "" & rsA.Fields("CCM21") '請款對象
         stA1K27 = "" & rsA.Fields("CCM22") '列印對象
         stCCM23 = "" & rsA.Fields("ccm23") '固定請款對象
         stCCM24 = "" & rsA.Fields("ccm24") '固定列印對象
         stCCM25 = "" & rsA.Fields("ccm25") '帳款已清(填結案單時,帳款狀態)
         'end 2025/08/18
      End If
      '結案理由代號/結案說明
      If "" & rsA.Fields("CCM04") <> MsgText(601) Then
         '商標
         If InStr(strNP02, "T") > 0 Then
            Option1(Val(rsA.Fields("CCM04"))).Value = True
            txtReason(0) = "" & rsA.Fields("CCM05")
         Else
            Option2(Val(rsA.Fields("CCM04"))).Value = True
            txtReason(1) = "" & rsA.Fields("CCM05")
         End If
      End If
      '是否閉卷
      If "" & rsA.Fields("CCM06") = "Y" Then
         If strNP02 = "FCP" Or strNP02 = "P" Or strNP02 = "FG" Or strNP02 = "CFP" Then
            Chk7(1).Value = vbChecked
            CboState(1) = stCCM07N
         Else
            ChkClose.Value = vbChecked
         End If
      ElseIf strNP02 = "FCP" Or strNP02 = "P" Or strNP02 = "FG" Or strNP02 = "CFP" Then
         Chk7(0).Value = vbChecked
         CboState(0) = stCCM07N
      End If
      
      If intState = 1 Then
         If Pub_QueryFCCloseDetail(intState, Me.Name, m_CCM01, strNP02, stMsg, Me.GridAMT) = False Then
            stMsg = "讀取FC商標結案資料有誤,通知電腦中心！" & vbCrLf & stMsg
            GoTo ErrHnd
         Else
            'Modify by Amy 2025/11/07 +stGridAmtCol及設定 金額/折扣靠右 ex:被退回不會設定
            Call Pub_SetCloseGridAmtWidth(Me.Name, GridAMT, arrAMTCol(), arrAMTWidth, , stGridAmtCol)
            Call SetGridColor(0, GridAMT, arrCol) 'Add by Amy 2025/11/07 設定 金額/折扣靠右
            'end 2025/11/07
         End If
         Call SetLock(1, False)
         Call PUB_GetNotPayA1K(" And a1k13='" & strNP02 & "' And  a1k14='" & strNP03 & "' And  a1k15='" & strNP04 & "' And a1k16='" & strNP05 & "'", , stA1K29)
         txtA1K(0) = stCCM25 '帳款已清
         txtA1K(1) = stA1K04 '列印申請人
         txtA1K(2) = stMergeBill '合併列印請款單
         txtA1K(3) = stA1K28: Call txtA1K_Validate(3, False) '請款對象
         txtA1K(4) = stCCM23: Call txtA1K_Validate(4, False) '固定請款對象
         txtA1K(5) = stA1K27: Call txtA1K_Validate(5, False) '列印對象
         txtA1K(6) = stCCM24: Call txtA1K_Validate(6, False) '固定列印對象
         'Modify by Amy 2025/10/02 改共用
         'Call SetA1KData(2)
         Call Pub_CloseSetA1KDataColor(1, Me.Name, strNP02, strNP03, strNP04, strNP05, m_NP07, Me.txtA1K, 0, 6)
      Else
         'Modify by Amy 2025/08/08 +stCCD08/stNotPay_CP ,未付帳款 及 管制催款日 都有資料,於程序操作完 解除期限/閉卷 後加行事曆
         If Pub_QueryFCCloseDetail(intState, Me.Name, m_CCM01, strNP02, stMsg, , Me.Chk6, Me.Chk8, Me.Chk9, stNotPay, stMemo, stCCD08, stNotPay_CP) = False Then
            stMsg = "讀取FC專利結案資料有誤,通知電腦中心！" & vbCrLf & stMsg
            GoTo ErrHnd
         End If
         'Add by Amy 2025/08/08
         lblNotPay_CP.Caption = stNotPay_CP
         txtCCD08 = stCCD08
         'end 2025/08/08
         txtNotPay = stNotPay
         txtFCPMemo = stMemo
      End If
      m_F0316 = "" & rsA.Fields("F0316") '智權人員
      
      '個人簽核(被退回)
      If UCase(TypeName(m_PrevForm)) = UCase("frm210147_1") Then
         If ChkForm("Q") = False Then
            Exit Function
         End If
         '下一程序資料
         Call doQueryNP(strNP02, strNP03, strNP04, strNP05, stMsg)
      End If
      
'*** 下載檔案 (此處有改, 需確認frm210133 是否也要改) ***
      If m_AttachPath = "" Then
         If Pub_SetFilePathDelTmp("Close", 1, strExc(9), m_AttachPath) = False Then
            MsgBox "附件資料夾建立失敗" & vbCrLf & _
                           strExc(9) & vbCrLf & "請洽電腦中心!", vbInformation
         End If
      End If
      If PUB_ChkIsReplyFile(strNP02 & strNP03 & strNP04 & strNP05, m_strSaveFiles, , , m_CCM01, intState, m_AttachPath) = True Then
         If m_strSaveFiles <> "" Then
            Call LoadFile(strNP02 & strNP03 & strNP04 & strNP05, m_strSaveFiles, m_AttachPath, stMsg)
            If stMsg <> "" Then
               MsgBox stMsg, vbCritical
           End If
         End If
      End If
'*** End 下載檔案 (此處有改, 需確認frm210133 是否也要改) ***
      doQueryCloseCase = True
   End If
   
ErrHnd:
   If Err.Number <> 0 Then stMsg = Err.Description
   Set rsA = Nothing
End Function

Private Function doQueryNP(m_NP02 As String, m_NP03 As String, m_NP04 As String, m_NP05 As String, ByRef stMsg As String) As Boolean
   Dim rsA As New ADODB.Recordset, strA As String
   Dim bolNotRun As Boolean
   
On Error GoTo ErrHnd

   doQueryNP = False
   MSHGrid1.Rows = 2
   MSHGrid1.Clear
   Frame3.Visible = False 'Add by Amy 2025/09/23 無期限閉卷,避免輸下一筆未還原
   
   If m_CCM01 <> MsgText(601) Then
      strA = strA & " And (NP24='" & m_CCM01 & "' Or length(NP24)=9) " 'length(NP24)=9:曾經收文過
   Else
      strA = strA & " And (NP24 is null Or length(NP24)=9) "
   End If
   
   strA = "Select ' ' AS V,Decode(SubStr(cp09,1,1),'C',Decode(cp05,'','',SubStr(cp05,1,4)-1911||'/'||SubStr(cp05,5,2)||'/'||SubStr(cp05,7,2)),'') as 來函收文日," & _
            "Decode(SubStr(cp09,1,1),'C',Decode('" & m_ApplyNA01 & "','000',C2.cpm03,C2.cpm04),'') as 來函性質," & _
            "Decode(SubStr(cp09,1,1),'C',cp09,'') as 來函總收文號,np07||' '||Decode('" & m_ApplyNA01 & "','000',C1.cpm03,C1.cpm04) as 下一程序," & _
            "Decode(np08,'','',SubStr(np08,1,4)-1911||'/'||SubStr(np08,5,2)||'/'||SubStr(np08,7,2)) as 本所期限," & _
            "Decode(np09,'','',SubStr(np09,1,4)-1911||'/'||SubStr(np09,5,2)||'/'||SubStr(np09,7,2)) as 法定期限," & _
            "st02 As 智權人員, np14 As 相關人, np15 As 備註,np07,np10,rownum as sort,NP01,NP22" & _
            " From NextProgress,CaseProgress,Staff,CasePropertyMap C1,CasePropertyMap C2" & _
            " Where NP02='" & m_NP02 & "' And NP03='" & m_NP03 & "' And NP04='" & m_NP04 & "' And NP05='" & m_NP05 & "'" & _
            " And np01=cp09(+) And np10=st01(+)" & strA & _
            " And np02=C1.cpm01(+) And np07=C1.cpm02(+) And cp01=C2.cpm01(+) And cp10=C2.cpm02(+)" & _
            " And np06 is null " & strNpSqlOfNoSalesDuty & _
            " Order by CP05 Desc, NP01 Desc, NP08 Desc"
      
      rsA.CursorLocation = adUseClient
      rsA.Open strA, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         doQueryNP = True
         Set MSHGrid1.Recordset = rsA
         'Call SetGrid1Data'Mark by Amy 2025/08/18 商標註冊費為全期, 故先不使用
         Call SetGridColor(0, MSHGrid1, arrCol)
         Call SetGridOneData(MSHGrid1)
      Else
         If Trim(lblCaseNm) = "" Then
            stMsg = "無案件資料！"
            bolNotRun = True
         Else
            doQueryNP = True
            Frame3.Visible = True '無期限閉卷
            stMsg = "無期限資料！"
            'Modify by Amy 2025/08/18 +CFC
            If m_NP02 = "FCT" Or m_NP02 = "T" Or m_NP02 = "S" Or m_NP02 = "CFT" Or m_NP02 = "CFC" Then
               ChkClose.Value = vbChecked
               Frame10.Enabled = False '閉卷,鎖住不可勾
            End If
         End If
      End If
      SetGridWidth
ErrHnd:
   If Err.Number <> 0 Then stMsg = Err.Description: bolNotRun = True
   If bolNotRun = True Then Exit Function
End Function

Private Sub SetData(ByRef GrdTp As MSHFlexGrid, arrTpCol() As String, intRow As Integer, Optional ByVal i As Integer)
   If UCase(GrdTp.Name) = "MSHGRID1" Then
      GrdTp.TextMatrix(intRow, 5) = Trim(GrdTp.TextMatrix(i, GetColVal(arrTpCol, "本所期限")))
      GrdTp.TextMatrix(intRow, 6) = Trim(GrdTp.TextMatrix(i, GetColVal(arrTpCol, "法定期限")))
      GrdTp.TextMatrix(intRow, 7) = Trim(GrdTp.TextMatrix(i, GetColVal(arrTpCol, "智權人員")))
      GrdTp.TextMatrix(intRow, 14) = Trim(GrdTp.TextMatrix(i, GetColVal(arrTpCol, "NP22")))
   Else
      LblCntItem.Caption = intRow
      TxtPay(0) = GrdTp.TextMatrix(intRow, GetColVal(arrTpCol, "代號"))
      lblPayItemN.Caption = GrdTp.TextMatrix(intRow, GetColVal(arrTpCol, "請款項目"))
      TxtPay(1) = Format(GrdTp.TextMatrix(intRow, GetColVal(arrTpCol, "金額")), "###0")
      TxtPay(2) = GrdTp.TextMatrix(intRow, GetColVal(arrTpCol, "折扣"))
      TxtPay(3) = GrdTp.TextMatrix(intRow, GetColVal(arrTpCol, "備註"))
   End If
End Sub

Private Sub SumGridAMT()
   Dim ii As Integer, intCol1 As Integer, intCol2 As Integer, intCol3 As Integer
   Dim stTP As String, stDisCount As String, stVal As String, stTotal As String, stSum As String
   
   '不論 新增 or 刪除 都要重算
   If GridAMT.Rows > 1 Then
      intCol1 = GetColVal(arrAMTCol(), "代號")
      intCol2 = GetColVal(arrAMTCol(), "金額")
      intCol3 = GetColVal(arrAMTCol(), "折扣")
      For ii = 1 To GridAMT.Rows - 1
         stTP = GridAMT.TextMatrix(ii, intCol1)
         stVal = Format(GridAMT.TextMatrix(ii, intCol2), "###0")
         stDisCount = GridAMT.TextMatrix(ii, intCol3)
         '折扣
         If stDisCount = "" Then stDisCount = "1"
         stDisCount = Val(stDisCount) / 100
         '98/98屬於規費
         If Right(stTP, 2) = "98" Or Right(stTP, 2) = "99" Then
            stSum = Val(stSum) + Val(stVal)
         End If
         stTotal = Val(stTotal) + (Val(stVal) * Val(stDisCount))
      Next ii
      txtAmt(0) = Format(stTotal, "#,##0") '總金額
      txtAmt(1) = Format(stSum, "#,##0") '規費
      txtAmt(2) = Format((CDbl(Val(stTotal)) - CDbl(Val(stSum))) / 1000, "####0.000")
   End If
End Sub

Private Sub cmdSend_Click()
   Dim Rs As New ADODB.Recordset, intI As Integer, ii As Integer, intChoose As Integer
   Dim strSubject As String, strContent As String
   Dim stCmd As String, stDate As String, stTime As String, stMsg As String, stCls As String, stAllCCD03 As String
   Dim bolModify  As Boolean, IsOK As Boolean
   Dim arrFile As Variant, stTemp As Variant
   
 On Error GoTo ErrHand
 
   If ChkForm("S") = False Then
      Exit Sub
   End If
   
   bolModify = False
   bolUpdNP24 = False: stCCM02_New = "": stCCM03_New = "" 'Add by Amy 2025/08/25
   If m_CCM01 <> MsgText(601) Then bolModify = True
      
   '無期限閉卷
   If Frame3.Visible = True Or MSHGrid1.TextMatrix(m_row, GetColVal(arrCol, "NP01")) = MsgText(601) Then
      'Modify by Amy 2025/08/25 +if 已送出後由有期限不續辦,改為閉卷
      If bolModify = True And Left(Pub_StrUserSt03, 2) = "F2" And Pub_GetField("NextProgress", "NP24='" & m_CCM01 & "'", "NP24") = m_CCM01 Then
         bolUpdNP24 = True
         stCCM02_New = strNP02 & strNP03 & strNP04 & strNP05
         stCCM03_New = ""
      Else
      'end 2025/08/25
         m_CCM02 = strNP02 & strNP03 & strNP04 & strNP05
         m_CCM03 = ""
      End If
   '有期限
   Else
      m_CCM02 = MSHGrid1.TextMatrix(m_row, GetColVal(arrCol, "NP01"))
      m_CCM03 = MSHGrid1.TextMatrix(m_row, GetColVal(arrCol, "NP22"))
   End If
   
   
   Screen.MousePointer = vbHourglass
   cmdSend.Enabled = False
   cnnConnection.BeginTrans
   
   stDate = strSrvDate(1)
   stTime = Right("000000" & ServerTime, 6)
   m_F0316 = GetF0316(m_F0316) '智權人員
   m_CCM04 = GetReason '結案原因代碼
   
   '新增
   If bolModify = False Then
      '表單編號自動給號
      m_CCM01 = AutoNo_FLOW("CLS", 5)
      '檢查是否還有自動給號資料不一致的問題
      stCmd = "select AU03 from autonumber where AU01='CLS'"
      intI = 1
      Set Rs = ClsLawReadRstMsg(intI, stCmd)
      If intI = 1 Then
         If Val(Rs.Fields("AU03")) <> Val(Right(m_CCM01, Len(m_CCM01) - 3)) Then
            stMsg = "系統自動給號(" & m_CCM01 & ")更新有誤，請洽電腦中心！"
            m_CCM01 = ""
            GoTo ErrHand
         End If
      End If
      '新增案件表單主檔
      'F0316=m_F0316 => F0316=m_CP13 因內專程序會幫雅娟、P1004代填, F0316存智權人員ID
      stCmd = "Insert into Flow003(F0301,F0302,F0307,F0310,F0311,F0312,F0316)" & _
                    " Values('" & m_CCM01 & "','" & Flow_結案單 & "','" & strUserNum & "'" & _
                     ",'" & strUserNum & "'," & stDate & "," & stTime & ",'" & m_CP13 & "')"
      cnnConnection.Execute stCmd, intI
      '結案單主檔
      IsOK = SaveCloseMain(1, stDate, stTime, stMsg)
   '修改
   Else
      '結案單主檔
      IsOK = SaveCloseMain(2, stDate, stTime, stMsg)
   End If
   If IsOK = False Then
      GoTo ErrHand
   End If
   
'*** 結案單明細 ***
   IsOK = False
   '請款項目
   If SSTab1.TabVisible(1) = True Then
      IsOK = Pub_SaveCloseDetail_FCT(Me.m_CCM01, Me.Name, Me.GridAMT, arrAMTCol(), stDate, stTime, stMsg)
   '聯絡項目
   ElseIf SSTab1.TabVisible(2) = True Then
      'Modify by Amy 2025/08/08 未付帳款 及 管制催款日 都有資料,於程序操作完 解除期限/閉卷 後加行事曆(顯示 未付帳款),故將欄位拆開
      If Chk7(0).Value = vbChecked Then
         '不續辦
         IsOK = Pub_SaveCloseDetail_FCP(Me.m_CCM01, Me.Name, "N", Me.Chk6, Me.txtNotPay, Me.Chk8, Me.txtFCPMemo, stDate, stTime, stMsg, stAllCCD03, Me.txtCCD08, Me.lblNotPay_CP.Caption)
      Else
         '閉卷
         IsOK = Pub_SaveCloseDetail_FCP(Me.m_CCM01, Me.Name, "Y", Me.Chk6, Me.txtNotPay, Me.Chk9, Me.txtFCPMemo, stDate, stTime, stMsg, stAllCCD03, Me.txtCCD08, Me.lblNotPay_CP.Caption)
      End If
      'end 2025/08/08
   End If
   If IsOK = False Then
      GoTo ErrHand
   End If
'*** End 結案單明細 ***
      
   '新增表單簽核檔
   stCmd = "Delete From FLOW002 where F0201=" & CNULL(m_CCM01)
   cnnConnection.Execute stCmd
   '簽核人員
   If SignPerson <> "" Then
      stTemp = Split(SignPerson, ",")
      For ii = 0 To UBound(stTemp)
         If stTemp(ii) <> MsgText(601) Then
            If ii = 0 And m_SetFlowEmp1 <> "" Then stTemp(ii) = m_SetFlowEmp1 '簽核人員1若有調整,已調整的為主
            stCmd = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(m_CCM01) & ",'1'," & (ii + 1) & "," & CNULL(CStr(stTemp(ii))) & ")"
            cnnConnection.Execute stCmd
         End If
      Next ii
   End If
   '*****若是代他人填單,簽核檔中若自己也是簽核人員之一時,一併確認掉
   stCmd = "update FLOW002 set " & _
                "F0205='" & stDate & "'" & _
                ",F0206='" & stTime & "'" & _
                ",F0207='1'" & _
                " where F0201='" & m_CCM01 & "' and F0204='" & strUserNum & "' and F0207 is null"
   cnnConnection.Execute stCmd
   '*****END
   '程序人員
   stCmd = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(m_CCM01) & ",'2',1," & CNULL(Left(m_F0202_2, 5)) & ")"
   cnnConnection.Execute stCmd
       
   '補看人員
   stCmd = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(m_CCM01) & ",'3',1," & CNULL(m_F0202_3) & ")"
   cnnConnection.Execute stCmd

   '更新下一程序
   If Val(m_CCM03) > 0 Then
      stCmd = "Update NextProgress Set NP24='" & m_CCM01 & "' WHERE NP01='" & m_CCM02 & "' and NP22=" & m_CCM03
      cnnConnection.Execute stCmd
   'Add by Amy 2025/08/25 +原有期限改無期限閉卷,NP24 要還原
   Else
      If bolUpdNP24 = True Then
         stCmd = "Update NextProgress Set NP24=Null WHERE NP24='" & m_CCM01 & "'"
         cnnConnection.Execute stCmd
      End If
   'end 2025/08/25
   End If

   '檔案處理
   If bolModify = True Then
      '修改時,記錄重送訊息
      stCmd = GetInsertFLOW004Sql(Trim(m_CCM01), strUserNum, stDate, stTime, Flow_重送, "")
      cnnConnection.Execute stCmd
      intChoose = 0
   Else
      intChoose = 1
   End If
   '儲存回覆單：存結案單號收進系統
   Call CloseSaveFile(intChoose, m_CCM01, strNP02, strNP03, strNP04, strNP05, False, intState, m_AttachPath, stMsg, m_strSaveFiles)
   If stMsg <> "" Then GoTo ErrHand
   
   '讀取下一處理人員
   If GetNextProPerson_Flow(m_CCM01, m_F0316, m_F0308, m_F0309) = False Then GoTo ErrHand
   
   cnnConnection.CommitTrans
   
   If UCase(TypeName(m_PrevForm)) = UCase("frm06010612") _
     Or UCase(TypeName(m_PrevForm)) = UCase("frm06010616") Then
      If bolIROK = False Then
         F_CCM01 = m_CCM01
         bolIROK = True
         '記錄第一筆案號
         strPreCase1 = strNP02: strPreCase2 = strNP03: strPreCase3 = strNP04: strPreCase4 = strNP05
         '記錄第一筆案號的代理人/申請人資訊
         b_FCAgNo = m_FCAgNo
         b_CuNo1 = m_CuNo1
         b_CuNo2 = m_CuNo2
         b_CuNo3 = m_CuNo3
         b_CuNo4 = m_CuNo4
         b_CuNo5 = m_CuNo5
      End If
   Else
      '[非] 外商/外專 系統收件區進入者,才刪檔匯入來源的檔案
      Call PUB_DelPCOrgFile(m_strSaveFiles)
   End If
   
   '發E-Mail通知下一處理主管(多審核主管用)
   If m_F0309 = Flow_主管審核中 Then
      If bolModify = True Then
         strContent = GetEMailContent_Flow(m_CCM01, strSubject, Flow_重送)
      Else
         strContent = GetEMailContent_Flow(m_CCM01, strSubject)
      End If
      '含 特殊職代
      'Modify by Amy 2025/08/12 只發人事職代-Sindy
      '     ex:1140811 Anny休假,黃賢泰填的結案單FCP-066436-0-00 不該發莊瑄凡(A8013)及洪培堯(A5023)
      'PUB_SendMail strUserNum, m_F0308, "", strSubject, strContent, , , , , , , , , , , , , True
      PUB_SendMail strUserNum, m_F0308, "", strSubject, strContent, , , , , , , , , , , , , True, , , , , , , , , , 1
   'Add By Sindy 2025/6/4 發信通知程序人員
   ElseIf m_F0309 = Flow_處理中 Then
      strContent = GetEMailContent_Flow(m_CCM01, strSubject)
      PUB_SendMail strUserNum, m_F0308, "", strSubject, strContent, , , , , , , , , , , , , True
   '2025/6/4 END
   End If
   
'*** P[非]寰華及CFP(內專程序結)要請款「出Outlook草稿」於此產生-邱子瑜 ***
   If intState = 2 Then
      '判斷是否要請款,stAllCCD03傳入目前所有之勾選,回傳-"0":寄 工程師+承辦/"1"-寄 承辦
      If Pub_ChkCloseInvoce(Me.Name, m_CCM01, strNP02, strNP03, strNP04, strNP05, IIf(bolFMPY53374 = False, "N", "Y"), stAllCCD03) = True Then
         If stAllCCD03 = "" Then
            MsgBox "開啟Outlook失敗,請洽電腦中心!", vbCritical, "操作錯誤！"
         Else
            strExc(9) = "907" '不續辦
            If Chk7(1).Value = vbChecked Then strExc(9) = "913" '閉卷
            If Pub_CloseOutLook(Me.Name, m_CCM01, strNP02, strNP03, strNP04, strNP05, m_ApplyNA01, strExc(9), stAllCCD03, stMsg) = False Then
               If stMsg <> "" Then
                  MsgBox stMsg, vbInformation '若無C類來函掛工程師,不需出草稿彈訊息
               Else
                  MsgBox "開啟Outlook失敗,請洽電腦中心!", vbCritical, "操作錯誤！"
               End If
            End If
         End If
      End If
   End If
'*** End P[非]寰華及CFP(內專程序結)要請款「出Outlook草稿」***
   
   cmdSend.Enabled = True
   Screen.MousePointer = vbDefault
   stCls = "1"
   If bolModify = True Then
        If TypeName(m_PrevForm) <> "Nothing" Then
            If UCase(TypeName(m_PrevForm)) = UCase("frm210147_1") Then
                m_PrevForm.cmdExit_Click
                Set m_PrevForm = Nothing
            End If
        End If
        Unload Me
   Else
'*** 新增 ***
      m_CCM02 = txt1(0) & "-" & txt1(1) & "-" & txt1(2) & "-" & txt1(3)
      '清除欄位值
      m_CCM01 = ""
      txt1(0) = "": txt1(1) = "": txt1(2) = "": txt1(3) = ""
      MSHGrid1.Clear
      MSHGrid1.Rows = 2
      SetGridWidth
      
      If UCase(TypeName(m_PrevForm)) <> "FRM06010612" Then GridAMT.Clear
      GridAMT.Rows = 2
      Call Pub_SetCloseGridAmtWidth(Me.Name, GridAMT, arrAMTCol(), arrAMTWidth)
      
      cmdOK(2).Default = True '查詢
      If TypeName(m_PrevForm) = "Nothing" Then
         '由案件結案單進入,按送出後Focus要停在系統類別欄
         txt1(0).SetFocus
      Else
         '外商/外專 系統收件區
         If UCase(TypeName(m_PrevForm)) = UCase("frm06010612") _
           Or UCase(TypeName(m_PrevForm)) = UCase("frm06010616") Then
            intI = MsgBox("此信件是否繼續進行多案結案？" & vbCrLf & _
                           "是：輸入下一筆" & vbCrLf & "否：此信件沖銷，回系統收件區", vbYesNo + vbDefaultButton2 + vbQuestion)
            stCls = "4"
            If intI = vbYes Then
               cmdOK(2).Visible = True
               cmdOK(2).Enabled = True
               txt1(0).Enabled = True: txt1(1).Enabled = True: txt1(2).Enabled = True: txt1(3).Enabled = True
               txt1(0).Locked = False: txt1(1).Locked = False: txt1(2).Locked = False: txt1(3).Locked = False
               Call ClearData(stCls)
            Else
               Unload Me
            End If
            Exit Sub
         End If
      End If
      Call ClearData(stCls)
'*** End 新增 ***
   End If
   'Add by Amy 2025/08/18
   If UCase(TypeName(m_PrevForm)) = UCase("frm210133_2") Then Unload Me
   Exit Sub
   
ErrHand:
   If bolModify = False Then m_CCM01 = ""
   cmdSend.Enabled = True
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   'Resume
   If Err.Number <> 0 Then
      stMsg = Err.Description
   End If
   If stMsg <> MsgText(601) Then
      MsgBox " 送出失敗,請洽電腦中心！" & vbCrLf & _
                         stMsg, vbInformation, "系統錯誤"
   End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
   '備註欄按enter鍵維持換行
   If KeyAscii = 13 Then
      If UCase(Me.ActiveControl.Name) = UCase("txtPay(3)") Or UCase(Me.ActiveControl.Name) = UCase("txtReason(0)") _
         Or UCase(Me.ActiveControl.Name) = UCase("txtNotPay") Or UCase(Me.ActiveControl.Name) = UCase("txtFCPMemo") Or UCase(Me.ActiveControl.Name) = UCase("txtReason(1)") Then
        Exit Sub
      End If
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   If Pub_SetFilePathDelTmp("Close", 1, strExc(9), m_AttachPath) = False Then
      MsgBox "附件資料夾建立失敗" & vbCrLf & _
                     strExc(9) & vbCrLf & "請洽電腦中心!", vbInformation, "系統錯誤"
   End If
   '提示文字
   Label2.Caption = "1.本人、帶人主管才能操作！"
   Frame2.Visible = False '列印'目前未使用
   Frame3.Visible = False '無期限閉卷
   Frame4.Visible = True
   
   If bolGoNext = False Then Call ClearData("0")
   If IsSetGirdColOK = False Then
      GridAMT.Clear
      Call Pub_SetCloseGridAmtWidth(Me.Name, Me.GridAMT, arrAMTCol(), arrAMTWidth(), True)
      MSHGrid1.Clear
      Call SetGridWidth
   End If
   Call SetLock(0, True)
     
   SSTab1.Tab = 0
   '商標
   If Left(Pub_StrUserSt03, 2) = "F1" Then
      SSTab1.TabVisible(2) = False
      SSTab1.TabVisible(3) = False
   '專利
   ElseIf Left(Pub_StrUserSt03, 2) = "F2" Then
      stPAndCFPMemo = "請於__月__日前確認並通知承辦，本案大陸代理人是否有最終帳單。（※ 若無需確認，請承辦自行移除文字。）"
      SSTab1.TabVisible(0) = False
      SSTab1.TabVisible(1) = False
      SSTab1.Tab = 2
   End If
   '商標
   If Left(Pub_StrUserSt03, 2) = "F1" Or Pub_StrUserSt03 = "M51" Then
      Call Pub_SetCloseReason(1, Me.Name, Me.Option1) '設定 Option1(index)=ROR01
      cmdSend.Default = False
   End If
   '專利
   'txtCCD08.Enabled = False 'Add by Amy 2025/08/08
   If Left(Pub_StrUserSt03, 2) = "F2" Or Pub_StrUserSt03 = "M51" Then
      If Left(Pub_StrUserSt03, 2) = "F2" Then Frame5.Caption = ""
      Call Pub_SetCloseReason(2, Me.Name, Me.Option2) '設定 Option2(index)=ROR01
      'FCP-907/913 狀態
      Call Pub_SetCloseCboState(Me.Name, Me.CboState(0), Me.CboState(1))
      'FCP-聯絡項目
      '都有的項目=Chk6,AC02=0開頭/不續辦 項目=Chk8,AC02=1開頭/閉卷 項目=Chk9,AC02=2開頭
      Call Pub_GetFCPContactItem(Me.Name, Me.Chk6, Me.Chk8, Me.Chk9)
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   If PUB_CheckFormExist("frm090201_2_1") = True Then
      Unload frm090201_2_1
   End If
   If PUB_CheckFormExist("frm210133_INV") = True Then
      Unload Forms(0).GetForm("frm210133_INV")
   End If
   If TypeName(m_PrevForm) <> "Nothing" Then
      If (UCase(TypeName(m_PrevForm)) = UCase("frm06010612") Or UCase(TypeName(m_PrevForm)) = UCase("frm06010616")) _
        And bolIROK = True Then
        '由系統收件區進入且未新更沖銷資料者(避免第2筆未結,按結束離開,不會沖銷,故寫於此)
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, "不續辦或閉卷", , , "結案單號=" & F_CCM01
         m_PrevForm.GoNext
      ElseIf UCase(TypeName(m_PrevForm)) <> UCase("frm06010612") And UCase(TypeName(m_PrevForm)) <> UCase("frm06010616") _
        And UCase(TypeName(m_PrevForm)) <> UCase("frm210147_1") And UCase(TypeName(m_PrevForm)) <> UCase("frm210133_2") Then
         m_PrevForm.PubShowNextData
      End If
      m_PrevForm.Show
      Set m_PrevForm = Nothing
   End If
   Set frm210133_F = Nothing
   m_strIR01 = "": m_strIR02 = "": m_strIR03 = "": m_strIR04 = ""
End Sub

Private Sub GridAMT_Click()
   nRow = GridAMT.MouseRow
   If nRow = 0 Then Exit Sub

   GridAMT.Visible = False
   Call SetData(GridAMT, arrAMTCol, nRow)
   Call SetGridColor(0, GridAMT, arrCol)
   Call SetGridColor(1, GridAMT, arrCol, nRow)
   GridAMT.Visible = True
End Sub

Private Sub MSHGrid1_Click()
   Dim intR As Integer, intCol As Integer
   Dim bolSetV As Boolean
   
   MSHGrid1.Visible = False
   MSHGrid1.row = MSHGrid1.MouseRow
   intCol = GetColVal(arrCol, "V")
   intR = MSHGrid1.row
   If Trim(MSHGrid1.TextMatrix(intR, intCol)) = "" Then
      bolSetV = True
   End If
   
   If intR >= 1 Then
      m_row = 0
      Call SetGridColor(0, MSHGrid1, arrCol)
      If bolSetV = True Then
         m_row = intR
         MSHGrid1.TextMatrix(m_row, intCol) = "V"
         Call SetGridColor(1, MSHGrid1, arrCol, m_row)
      End If
   End If

   MSHGrid1.Visible = True
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
   If SSTab1.Tab = 1 Then
      TxtPay(0).SetFocus
   End If
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub SetGridWidth()
   Dim stField As String, stWidth As String
   
   MSHGrid1.Visible = False
   If IsSetGirdColOK = False Then
      stField = "V,來函收文日,來函性質,來函總收文號,下一程序,本所期限,法定期限,智權人員,相關人,備註" & _
                        ",NP07,NP10,Sort,NP01,NP22"
      stWidth = "200,1000,1000,1000,1500,800,800,800,1000,3000" & _
                         ",0,0,0,0,0"
      arrCol = Split(stField, ",")
      arrWidth = Split(stWidth, ",")
      IsSetGirdColOK = True
   End If
   MSHGrid1.Cols = UBound(arrCol) + 1
   MSHGrid1.row = 0
   For i = 0 To MSHGrid1.Cols - 1
      MSHGrid1.col = i
      MSHGrid1.Text = arrCol(i)
      MSHGrid1.ColWidth(i) = Val(arrWidth(i))
      MSHGrid1.CellAlignment = flexAlignLeftCenter
   Next i
   MSHGrid1.Visible = True
End Sub

Private Function GetReason() As String
   Dim opt As Object
   
   GetReason = ""
   '商標
   'Modify by Amy 2025/08/18 +CFC
   If strNP02 = "FCT" Or strNP02 = "T" Or strNP02 = "S" Or strNP02 = "CFT" Or strNP02 = "CFC" Then
      For Each opt In Option1
         If opt.Value = True Then
            GetReason = Format(opt.Index, "00")
            Exit For
         End If
      Next
   ElseIf strNP02 = "FCP" Or strNP02 = "P" Or strNP02 = "FG" Or strNP02 = "CFP" Then
      For Each opt In Option2
         If opt.Value = True Then
            GetReason = Format(opt.Index, "00")
            Exit For
         End If
      Next
   End If
End Function

Private Function ChkFlowPerson(ByRef stMsg As String) As Boolean
   Dim stSignP2 As String
   
   ChkFlowPerson = False
   '承辦簽核人員(David 少輸,st52為自己)
   SignPerson = Left(GetSignOffEmp("CM1", strNP02, strNP03, m_ApplyNA01, , , m_F0316), 5)
   'Modify by Amy 2025/08/18 +if 承辦簽核人員 為操作者,不需簽核(不寫Flow002);外商+CFC
   If strUserNum = SignPerson Then
      SignPerson = ""
   End If
   'end 2025/08/18
   If InStr(strNP02, "T") > 0 Or strNP02 = "S" Or strNP02 = "CFC" Then
      '外商需加簽ST53
      stSignP2 = Left(GetSignOffEmp("CM2", strNP02, strNP03, m_ApplyNA01, , , m_F0316), 5)
      'Modify by Amy 2025/11/07 加iif 承辦簽核人員 為操作者,不需簽核時不需加","
      If stSignP2 <> "" Then SignPerson = IIf(SignPerson = "", SignPerson, SignPerson & ",") & stSignP2
   End If
   If SignPerson = "" Then
      stMsg = "無設定簽核人員，不可使用電子表單流程！"
      Exit Function
   End If
   
   '程序人員 (Memo CF案承辦為國外部者,仍會需要顯示其他頁籤輸入資料,故改用此支)
   'T -內商程序結 / FCT -外商程序結案(FCT爭議 原:內商結,114/7/2 秀玲詢問後,全部回到外商結);CFT 及 S 非台灣 照舊結案單
   'P[非]寰華FMP-內專程序結 / 'FCP/FG/P寰華 外專程序結;CFP 照舊結案單
   m_F0202_2 = GetSignOffEmp("NP", strNP02, strNP03, m_ApplyNA01, strNP02 & "-" & strNP03 & "-" & strNP04 & "-" & strNP05, m_NP07, m_F0316, bolFMPY53374)
   If m_F0202_2 = "" Then
      stMsg = "無設定程序人員，請通知電腦中心！"
      Exit Function
   End If
   '補看人員
   'T or FCT爭議案-Pub_GetSpecMan("內商爭議案程序主管") / FCT 非爭議案-案件程序人員的二級主管;CFT 及 S 非台灣 照舊結案單
   'P[非]寰華FMP-Pub_GetSpecMan("專利處特定編號") / 'FCP/FG/P寰華 案件程序人員的二級主管;CFP 照舊結案單
   m_F0202_3 = GetF0202_3(strNP02, strNP03, strNP04, strNP05, m_NP07, Left(m_F0202_2, 5), strUserNum, m_ApplyNA01)
   If m_F0202_3 = "" Then
      stMsg = "無設定補看人員，請通知電腦中心！"
      Exit Function
   End If
   ChkFlowPerson = True
End Function
 
Private Function ChkForm(stType As String) As Boolean
   Dim obj As Object, intIdx As Integer, intSSTabIdx As Integer
   Dim bolOK As Boolean, Is99NoMemo As Boolean, bCancel As Boolean, strMsg As String
   Dim o_A1K04 As String, o_A1K27 As String, o_A1K28 As String, o_A1K29 As String, o_TM56 As String, o_TM69 As String '目前系統設定
   
   ChkForm = False
   
   If Trim(txt1(0)) = "" Or Trim(txt1(1)) = "" Then
      MsgBox "本所案號不可以空白！", vbCritical, "操作錯誤！"
      If txt1(1) = "" Then txt1(1).SetFocus
      If txt1(0) = "" Then txt1(0).SetFocus
      Exit Function
   End If
   
   Select Case stType
      Case "F" '夾檔
         If stType = "F" Then
            If txt1(0) <> strNP02 Or txt1(1) <> strNP03 Then
               MsgBox "請先查詢要執行的資料！", vbCritical, "操作錯誤！"
               txt1_GotFocus 0
               Exit Function
            End If
         End If
      Case "Q" '查詢
         'Modify by Amy 2025/08/18 +CFC
         If Left(Pub_StrUserSt03, 2) = "F1" And Not (txt1(0) = "FCT" Or txt1(0) = "T" Or txt1(0) = "S" Or txt1(0) = "CFT" Or txt1(0) = "CFC") _
           Or Left(Pub_StrUserSt03, 2) = "F2" And Not (txt1(0) = "FCP" Or txt1(0) = "P" Or txt1(0) = "FG" Or txt1(0) = "CFP") Then
           MsgBox "系統別輸入錯誤！", vbCritical, "操作錯誤！"
           If txt1(0) = "" Then txt1(0).SetFocus
           Exit Function
         End If
         '檢查是否已有結案單
         If txt1(2) = MsgText(601) Then txt1(2) = "0"
         If txt1(3) = MsgText(601) Then txt1(3) = "00"
         'Memo by Amy 1140617 外商人員只要有FC代理人,都可能要輸請款項目,故 CFT/CFC及 S案台灣案有FC代理人,需從此支操作
         '     由外商系統操作結案單者,需先輸入案號,判斷
         '     [有]FC代理人進此支,可操作FCT案(含爭議案)、S台灣案、T(FMT)案、CFT/CFC有FC代理人之案件
         '     [無]FC代理人則進舊的( frm210133)
         If ChkCloseData(txt1(0), txt1(1), txt1(2), txt1(3), strMsg) = False Then
            MsgBox strMsg, vbCritical, "操作錯誤！"
            Exit Function
         End If
         
         '外商 由案號輸入進入,先檢查權限
         If UCase(TypeName(m_PrevForm)) = "FRM210133_2" Then
            If Pub_FCInuptCloseLimit(m_CP13, strNP02, strNP03, strNP04, strNP05) = False Then
               MsgBox "無權限操作此案號", vbCritical, "操作錯誤！"
               Exit Function
            End If
         End If
         
      Case "S" '送出
 '****** 送出 ******
         If txt1(0) <> strNP02 Or txt1(1) <> strNP03 Then
            MsgBox "請先查詢要執行的資料！", vbCritical, "操作錯誤！"
            txt1_GotFocus 0
            Exit Function
         End If
         '可操作結案單權限 (確定案子有資料,再判斷權限)
         If Pub_FCInuptCloseLimit(m_CP13, strNP02, strNP03, strNP04, strNP05, bolFMPY53374, bolFMP) = False Then
            MsgBox "無權限操作此案號", vbCritical, "操作錯誤！"
            Exit Function
         End If
         If m_IsClose = "Y" Then
            MsgBox "此案已閉卷，不可操作此作業！", vbExclamation
            Exit Function
         End If
         If Pub_FileIsOpen(Me.Name, m_strSaveFiles, strExc(9)) = True Then
            MsgBox "檔案正在使用中,需關閉之檔案如下:" & vbCrLf & _
                           Replace(strExc(9), ";", vbCrLf), vbCritical, "操作錯誤！"
            Exit Function
         End If
         
         'Add by Amy 2025/08/25
         '外專 判斷
         If Left(Pub_StrUserSt03, 2) = "F2" Then
            '無期限 不可勾選 不續辦 (Amy測退回由閉卷勾成不續辦,避免下方(ex:)問題加此判斷)
            '     ex:1140822 Anny 操作 FCP-064074 不續辦(有期限),共10件 程序從待處理區明細按「解除期限」鈕,不應該直接進閉卷畫面
            If Frame3.Visible = True And Chk7(0).Value = vbChecked Then
               If UCase(TypeName(m_PrevForm)) = UCase("frm210147_1") Then
                  MsgBox "原 閉卷 改為不續辦 不會帶出下一程序" & vbCrLf & _
                                 "請刪除此結案單重新操作", vbCritical, "操作錯誤！"
               Else
                  MsgBox "無期限不可勾選不續辦", vbCritical, "操作錯誤！"
               End If
               Exit Function
            End If
         End If
         
         '無期限閉卷
         If Frame3.Visible = True Then
            '檢查是否有資料重覆
            If ChkFlowFormExists(Flow_結案單, "", "", strNP02, strNP03, strNP04, strNP05, , , m_CCM01) = True Then
               MsgBox "此結案單已存在，不可重覆作業！", vbCritical, "操作錯誤！"
               txt1(0).SetFocus
               Exit Function
            End If
         '外專 有期限 勾閉卷,不需檢查是否勾選期限資料=無期限閉卷
         ElseIf Left(Pub_StrUserSt03, 2) = "F2" And Chk7(1).Value = vbChecked And m_row = 0 Then
            Frame3.Visible = True
         Else
            '有期限閉卷一定要勾選期限資料
            If m_row = 0 Then
               MsgBox "請勾選欲解除期限的資料！", vbCritical, "操作錯誤！"
               Exit Function
            End If
            
            '檢查是否有資料重覆
            strExc(8) = MSHGrid1.TextMatrix(m_row, GetColVal(arrCol, "NP01"))
            strExc(9) = MSHGrid1.TextMatrix(m_row, GetColVal(arrCol, "NP22"))
            If ChkFlowFormExists(Flow_結案單, strExc(8), strExc(9), strNP02, strNP03, strNP04, strNP05, , , m_CCM01) = True Then
               MsgBox "此結案單已存在，不可重覆作業！", vbCritical, "操作錯誤！"
               txt1(0).SetFocus
               Exit Function
            End If
            m_NP07 = MSHGrid1.TextMatrix(m_row, GetColVal(arrCol, "NP07"))
            m_F0316 = MSHGrid1.TextMatrix(m_row, GetColVal(arrCol, "NP10"))
         End If
         
         m_F0316 = GetF0316(m_F0316)
         '檢查簽核人員設定
         If ChkFlowPerson(strMsg) = False Then
            MsgBox strMsg, vbCritical, "操作錯誤！"
            Exit Function
         End If
         
         '商標
         'Modify by Amy 2025/08/18 +CFC
         If txt1(0) = "FCT" Or txt1(0) = "T" Or txt1(0) = "S" Or txt1(0) = "CFT" Or txt1(0) = "CFC" Then
            intSSTabIdx = 0
            intIdx = 0
            'FCT結案原因點選項目
            For Each obj In Option1
               If obj.Value = True Then
                  bolOK = True
                  If obj.Index = 99 And txtReason(0) = MsgText(601) Then
                     Is99NoMemo = True
                  End If
                  Exit For
               End If
            Next
         ElseIf txt1(0) = "FCP" Or txt1(0) = "P" Or txt1(0) = "FG" Or txt1(0) = "CFP" Then
            intSSTabIdx = 3
            intIdx = 1
            'FCP結案原因點選項目
            For Each obj In Option2
               If obj.Value = True Then
                  bolOK = True
                  If obj.Index = 99 And txtReason(1) = MsgText(601) Then
                     Is99NoMemo = True
                  End If
               End If
            Next
         End If
         If bolOK = False Then
            SSTab1.Tab = intSSTabIdx
            MsgBox "請勾選結案理由！", vbCritical, "操作錯誤！"
            Exit Function
         End If
         '勾選其他未輸理由
         If Is99NoMemo = True Then
            SSTab1.Tab = intSSTabIdx
            MsgBox "結案理由為「其他」" & vbCrLf & _
                             "理由不可為空！", vbCritical, "操作錯誤！"
            'Modify by Amy 2025/08/18 +CFC
            If txt1(0) = "FCT" Or txt1(0) = "T" Or txt1(0) = "S" Or txt1(0) = "CFT" Or txt1(0) = "CFC" Then
               txtReason(0).SetFocus
            ElseIf txt1(0) = "FCP" Or txt1(0) = "P" Or txt1(0) = "FG" Or txt1(0) = "CFP" Then
               txtReason(1).SetFocus
            End If
            Exit Function
         End If
         bCancel = False
         Call txtReason_Validate(intIdx, bCancel)
         If bCancel = True Then
            SSTab1.Tab = intSSTabIdx
            Exit Function
         End If
         
''*** 請款項目 相關 ***
         If SSTab1.TabVisible(1) = True Then
            intSSTabIdx = 1
            Call Pub_GetCloseA1KData(1, Me.Name, strNP02, strNP03, strNP04, strNP05, o_A1K29, m_NP07, o_A1K04, o_A1K27, o_A1K28, o_TM56, o_TM69)
            For Each obj In txtA1K
               bCancel = False: strMsg = "": strExc(8) = "": strExc(9) = ""
               If obj.Index <> 0 And obj.Index <> 2 Then
                  intIdx = obj.Index
                  Select Case obj.Index
                     Case 1
                        strExc(9) = o_A1K04
                        strExc(8) = "列印申請人"
                     Case 3
                        strExc(9) = o_A1K27
                        strExc(8) = "請款對象"
                     Case 4
                        strExc(9) = o_TM56
                        strExc(8) = "固定請款對象"
                     Case 5
                        strExc(9) = o_A1K28
                        strExc(8) = "列印對象"
                     Case 6
                        strExc(9) = o_TM69
                        strExc(8) = "固定列印對象"
                  End Select
                  'Modify by Amy 2025/10/02 原:有值且與目前資料相同不需輸,經理說直接帶至畫面中,故改為空時設回原值
'                  If obj.Index = 1 Then
'                     '畫面為Y,目前設定不印=Y or 畫面為N,目前設定不印=空白
'                     If (obj.Text = "Y" And strExc(9) = "Y") Or (obj.Text = "N" And strExc(9) = "") Then
'                        strMsg = strExc(8)
'                     End If
'                  ElseIf obj.Text <> "" And obj.Text = strExc(9) Then
'                      strMsg = strExc(8)
'                  End If
'                  If strMsg <> "" Then
'                     strMsg = "「" & strMsg & "」與目前資料相同" & vbCrLf & "不需輸入！"
'                     Exit For
'                  End If
                   If obj.Text = "" And strExc(9) <> "" Then obj.Text = strExc(9)
                   'end 2025/10/02
                  Call txtA1K_Validate(obj.Index, bCancel)
                  If bCancel = True Then Exit For
               '合併列印請款單
               ElseIf obj.Index = 2 Then
                  Call txtA1K_Validate(obj.Index, False)
               End If
            Next
            If bCancel = True Or strMsg <> "" Then
               If strMsg <> "" Then MsgBox strMsg, vbCritical, "操作錯誤！"
               SSTab1.Tab = intSSTabIdx
               txtA1K(intIdx).SetFocus
               Exit Function
            End If
            If Trim(TxtPay(0)) <> "" Then
               MsgBox "有請款項目未加入,請確認", vbCritical, "操作錯誤！"
               SSTab1.Tab = intSSTabIdx
               TxtPay(0).SetFocus
               Exit Function
            End If
            'Add by Amy 2025/11/07 檢查 沒勾「閉卷」且請款項目前3碼為704  Or 有勾「閉卷」且請款項目前3碼為703 不可存
            If ChkGridAmt(strMsg) = True Then
                MsgBox strMsg, vbCritical, "操作錯誤！"
                SSTab1.Tab = intSSTabIdx
                Exit Function
            End If
            'end 2025/11/07
         End If
''*** End 請款項目 相關 ***
'*** 聯絡項目 相關 ***
         If SSTab1.TabVisible(2) = True Then
            intSSTabIdx = 2
            bolOK = False: bCancel = False: strMsg = ""
            '未付帳款
            Call txtNotPay_Validate(bCancel)
            If bCancel = True Then
               SSTab1.Tab = intSSTabIdx
               Exit Function
            End If
            'Add by Amy 2025/08/08 管制催款日
            If txtCCD08 <> "" Then
               Call txtCCD08_Validate(bCancel)
               If bCancel = True Then
                  SSTab1.Tab = intSSTabIdx
                  Exit Function
               End If
            End If
            If Chk7(0).Value = vbUnchecked And Chk7(1).Value = vbUnchecked Then
               SSTab1.Tab = intSSTabIdx
               MsgBox "「907不續辦」及「913閉卷」" & vbCrLf & _
                                "請擇一勾選！", vbCritical, "操作錯誤！"
               Exit Function
            End If
            '907不續辦
            If Chk7(0).Value = vbChecked Then
               If CboState(0) = MsgText(601) Then
                  SSTab1.Tab = intSSTabIdx
                  MsgBox "「907不續辦」狀態不可為空！", vbCritical, "操作錯誤！"
                  CboState(0).SetFocus
                  Exit Function
               'Add by Amy 2025/08/11 「907不續辦」狀態選99-其他,其他說明必填-Anny
               ElseIf Left(CboState(0), Val(InStr(CboState(0), ".")) - 1) = "99" And txtFCPMemo = "" Then
                  MsgBox "「907不續辦」狀態為「99.其他」" & vbCrLf & _
                                  "其他說明不可為空！", vbCritical, "操作錯誤！"
                  txtFCPMemo.SetFocus
                  Exit Function
               'end 2025/08/11
               End If
               strMsg = "907不續辦"
               For Each obj In Chk8
                  If obj.Value = vbChecked Then
                     bolOK = True
                  End If
               Next
            End If
            '913閉卷
            If Chk7(1).Value = vbChecked Then
               If CboState(1) = MsgText(601) Then
                  SSTab1.Tab = intSSTabIdx
                  MsgBox "「913閉卷」狀態不可為空！", vbCritical, "操作錯誤！"
                  CboState(1).SetFocus
                  Exit Function
               'Add by Amy 2025/08/11 「913閉卷」狀態選99-其他,其他說明必填-Anny
               ElseIf Left(CboState(1), Val(InStr(CboState(1), ".")) - 1) = "99" And txtFCPMemo = "" Then
                  MsgBox "「913閉卷」狀態為「99.其他」" & vbCrLf & _
                                  "其他說明不可為空！", vbCritical, "操作錯誤！"
                  txtFCPMemo.SetFocus
                  Exit Function
               'end 2025/08/11
               End If
               strMsg = "913閉卷"
               For Each obj In Chk9
                  If obj.Value = vbChecked Then
                     bolOK = True
                  End If
               Next
            End If
            If bolOK = False Then
               SSTab1.Tab = intSSTabIdx
               MsgBox "請勾選「" & strMsg & "」之聯絡項目！", vbCritical, "操作錯誤！"
               Exit Function
            End If
         End If
'*** End 聯絡項目 相關 ***

         '檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
         If PUB_ChkUniText(Me, , True, "TextBox") = False Then
             Exit Function
         End If
         
         If ChkDataYN = False Then
            Exit Function
         End If
 '****** End 送出 ******
      Case Else
   End Select
   
   ChkForm = True
End Function

Private Function ChkSSTab1(stState As String) As Boolean
   Dim strMsg As String, bCancel As Boolean
   
   ChkSSTab1 = False
   If Trim(strNP02) = "" Or Trim(strNP02) = "" Then
      MsgBox "本所案號不可以空白！", vbCritical, "操作錯誤！"
      If txt1(1) = "" Then txt1(1).SetFocus
      If txt1(0) = "" Then txt1(0).SetFocus
      Exit Function
   End If
   
   Select Case stState
        Case "A" '加入鈕
         If TxtPay(0) = MsgText(601) Then
            MsgBox "請款項目不可為空！", vbCritical, "操作錯誤！"
            SSTab1.Tab = 1
            TxtPay(0).SetFocus
            Exit Function
         End If
         
         Call TxtPay_Validate(0, False)
         If lblPayItemN.Caption = MsgText(601) Then
            SSTab1.Tab = 1
            TxtPay(0).SetFocus
            Exit Function
         End If
         If TxtPay(1) = MsgText(601) Then
            MsgBox "金額不可為空！", vbCritical, "操作錯誤！"
            SSTab1.Tab = 1
            TxtPay(1).SetFocus
            Exit Function
         End If
         Call TxtPay_Validate(1, bCancel)
         If bCancel = True Then
            SSTab1.Tab = 1
            Exit Function
         End If
         '折扣
         Call TxtPay_Validate(2, bCancel)
         If bCancel = True Then
            SSTab1.Tab = 1
            Exit Function
         End If
         '備註
         Call TxtPay_Validate(3, bCancel)
         If bCancel = True Then
            SSTab1.Tab = 1
            Exit Function
         End If
      Case "D" '刪除鈕
         If LblCntItem.Caption = MsgText(601) Then strMsg = "請選擇欲刪除的資料！"
         If TxtPay(0) = MsgText(601) Then strMsg = "無資料可刪除！"
         If strMsg <> MsgText(601) Then
            MsgBox strMsg, vbCritical, "操作錯誤！"
            SSTab1.Tab = 1
            Exit Function
         End If
         If Val(LblCntItem.Caption) > 0 Then
            strMsg = "確定要刪除第" & LblCntItem.Caption & "筆資料？"
            If MsgBox(strMsg, vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Function
            End If
         End If
   End Select
   ChkSSTab1 = True
End Function

'檢查是否已有結案單
Private Function ChkCloseData(m_NP02 As String, m_NP03 As String, m_NP04 As String, m_NP05 As String, ByRef stMsg As String) As Boolean
   Dim RsQ As New ADODB.Recordset, intQ As Integer
   Dim stQ As String
   
   ChkCloseData = False
   
   '讀取相關資料及是否閉卷
   stQ = "Select TM12,CaseN,A.NA03 as ApplyNation,Nation,PKind,TM11,sClose,CU10" & _
               ",sApply,NVL(CU04,Decode(CU05,null,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CName,CU10,C.NA03 as CNA03" & _
               ",FCAg,Decode(FA05,null,Nvl(FA04,FA06),FA05||' '||FA63||' '||FA64||' '||FA65) as AName,FA10,F.NA03 as FNA03 " & _
               ",sApply2,sApply3,sApply4,sApply5 " & _
               "From(" & _
                         "Select TM12,TM05||TM06||TM07 as CaseN,TM23 as sApply,TM10 as Nation,' ' as PKind,TM11,TM29 as sClose,TM44 as FCAg " & _
                         ",tm78 as sApply2,tm79 as sApply3,tm80 as sApply4,tm81 as sApply5 From Trademark " & _
                         "Where TM01='" & m_NP02 & "' And TM02='" & m_NP03 & "' And TM03='" & m_NP04 & "' And TM04='" & m_NP05 & "' " & _
            "Union Select PA11,PA05||PA06||PA07,PA26,PA09 as Nation,PA08 as PKind,PA10,PA57 as sClose,PA75 as FCAg " & _
                           ",pa27 as sApply2,pa28 as sApply3,pa29 as sApply4,pa30 as sApply5 From Patent " & _
                         "Where PA01='" & m_NP02 & "' And PA02='" & m_NP03 & "' And PA03='" & m_NP04 & "' And PA04='" & m_NP05 & "' " & _
            "Union Select SP11,SP05||SP06||SP07,SP08,SP09 as Nation,' ' as PKind,0,SP15 as sClose,SP26 as FCAg " & _
                           ",sp58 as sApply2,sp59 as sApply3,sp65 as sApply4,sp66 as sApply5 From Servicepractice " & _
                         "Where SP01='" & m_NP02 & "' And SP02='" & m_NP03 & "' And SP03='" & m_NP04 & "' And SP04='" & m_NP05 & "' " & _
            "),Nation A,Customer,Nation C,Nation F,FAgent " & _
            "Where SubStr(sApply,1,8)=CU01(+) And Decode(SubStr(sApply,9,1),'','0',SubStr(sApply,9,1))=CU02(+) And CU10=C.NA01(+) " & _
            "And SubStr(FCAg,1,8)=FA01(+) And Decode(SubStr(FCAg,9,1),'','0',SubStr(FCAg,9,1))=FA02(+) And FA10=F.NA01(+) " & _
            "And Nation=A.NA01(+) "
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stQ)
   If intQ = 1 Then
      lblApplyCaseNo = "" & Trim(RsQ.Fields("TM12")) '申請案號
      lblCaseNm.Caption = "" & RsQ.Fields("CaseN") '案件名稱
      m_ApplyNA01 = "" & RsQ.Fields("Nation") '申請國家 編號
      lblApplyNation = m_ApplyNA01 & " " & RsQ.Fields("ApplyNation") '申請國家編號 名稱
      m_PA08 = "" & RsQ.Fields("PKind") '專利種類
      m_PA10 = "" & RsQ.Fields("TM11") '申請日
      m_IsClose = "" & RsQ.Fields("sClose") '是否閉卷
      
      'FC代理人
      m_FCAgNo = "" & Trim(RsQ.Fields("FCAg")) '編號
      lblFCAgNm.Caption = m_FCAgNo & " " & Trim(RsQ.Fields("AName"))   '編號 名稱
      m_FA10 = "" & Trim(RsQ.Fields("FA10")) '國籍編號
      lblFCAgNation = m_FA10 & " " & Trim(RsQ.Fields("FNA03")) '國籍編號 名稱
      
      '申請人
      m_CuNo1 = "" & RsQ.Fields("sApply")
      lblCU01Nm.Caption = m_CuNo1 & " " & RsQ.Fields("CName") '申請人1編號 名稱
      m_CU10 = Mid("" & Trim(RsQ.Fields("CU10")), 1, 3) '申請人1國籍編號
      lblCuNation = m_CU10 & " " & Trim(RsQ.Fields("CNA03")) '國籍編號 名稱
      m_CP13 = ShowCurrCP13(m_NP02, m_NP03, m_NP04, m_NP05, m_ApplyNA01) '承辦人員
      m_CuNo2 = "" & RsQ.Fields("sApply2")
      m_CuNo3 = "" & RsQ.Fields("sApply3")
      m_CuNo4 = "" & RsQ.Fields("sApply4")
      m_CuNo5 = "" & RsQ.Fields("sApply5")
      
      If m_IsClose = "Y" Then
         stMsg = "此案已閉卷，不可操作此作業！"
         Set RsQ = Nothing
         If txt1(1).Enabled = True Then txt1(1).SetFocus
         Exit Function
      End If
   End If
      
   ChkCloseData = True
   Set RsQ = Nothing
End Function

'Memo by Amy 2025/08/18 延用國內結案單寫法(Frm210133),以前分一期/二期目前都為全期-詢問Sindy後可先不用
Private Sub SetGrid1Data()
   Dim ii As Integer, intRow As Integer
   
   MSHGrid1.Visible = False
   intRow = MSHGrid1.Rows - 1
   For ii = 1 To MSHGrid1.Rows - 1
      If strNP02 = "FCT" Or strNP02 = "T" Then
         If m_ApplyNA01 < "010" And Trim(MSHGrid1.TextMatrix(ii, GetColVal(arrCol, "NP07"))) = "715" Then
            intRow = intRow + 1
            MSHGrid1.AddItem ("")
            MSHGrid1.TextMatrix(intRow, GetColVal(arrCol, "下一程序")) = "717 " & GetCaseTypeName(strNP02, "717", 0)
            MSHGrid1.TextMatrix(intRow, GetColVal(arrCol, "NP07")) = "717"
            MSHGrid1.TextMatrix(intRow, GetColVal(arrCol, "Sort")) = CStr(MSHGrid1.TextMatrix(ii, GetColVal(arrCol, "Sort"))) & "-1"
            Call SetData(MSHGrid1, arrCol, intRow, ii)
         End If
      End If
   Next ii
   MSHGrid1.Visible = True
End Sub

'intChoose:0-還原顏色=[不選取] / 1-設定[選取]顏色
Private Sub SetGridColor(intChoose As Integer, ByRef GrdTp As MSHFlexGrid, arrTpCol() As String, Optional ByVal intR As Integer = 0)
   Dim i As Integer, j As Integer, intStart As Integer, intEnd As Integer, intVCol As Integer
   Dim stTmp As String
   
   If intChoose = 1 And intR <= 0 Then Exit Sub
   
   GrdTp.Visible = False
   '還原顏色=[不選取]
   If intChoose = 0 Then
      intStart = 0: intEnd = GrdTp.Cols - 1
      If UCase(GrdTp.Name) = "MSHGRID1" Then intVCol = GetColVal(arrCol, "V")
      For j = 1 To GrdTp.Rows - 1
         GrdTp.row = j
         For i = intStart To intEnd
            GrdTp.col = i
            If UCase(GrdTp.Name) = "MSHGRID1" Then
               If MSHGrid1.Text = "V" Then
                  MSHGrid1.Text = " "
               End If
            'Add by Amy 2025/11/07 +GridAmt 設定
            ElseIf UCase(GrdTp.Name) = "GRIDAMT" Then
               '金額/折扣 靠右
               stTmp = Replace(GrdTp.Text, ",", "")
               If InStr(";" & stGridAmtCol, ";" & i) > 0 Then
                  GrdTp.Text = Format(stTmp, "#,##0")
                  GrdTp.CellAlignment = flexAlignRightCenter
               End If
            'end 2025/11/07
            End If
            '未選取(依狀態設定顏色)
            GrdTp.CellBackColor = &H80000018 '設回
         Next i
      Next j
   '選取(勾選)
   ElseIf intChoose = 1 Then
      GrdTp.row = intR
      For i = 0 To GrdTp.Cols - 1
         GrdTp.col = i
         GrdTp.CellBackColor = &HFFC0C0 '整列底 藍色
      Next i
   End If
   
   GrdTp.Visible = True
End Sub

'查詢只有一筆資料Grid顏色設定
Private Sub SetGridOneData(ByRef GrdTp As MSHFlexGrid)
   GrdTp.Visible = False
   With GrdTp
      If .Rows = 2 Then
         .row = 1
         .col = 1
         If .Text <> "V" Then
            If UCase(GrdTp.Name) = "MSHGRID1" Then m_row = 1
            .row = 1
            .col = 0
            .Text = "V"
            '變過色仍需要再跑,因為只有一筆選取時要變藍色
            Call SetGridColor(1, GrdTp, arrCol, 1)
         End If
      End If
   End With
   GrdTp.Visible = True
End Sub

Private Function SaveCloseMain(intEdit As Integer, stDate As String, stTime As String, ByRef stMsg As String) As Boolean
   Dim stCmd As String, stTmp As String, intC As Integer
   Dim stCCM05 As String, stCCM06 As String, stCCM07 As String, stCCM08 As String, stCCM09 As String, stCCM10 As String
   Dim stCCM19 As String, stCCM20 As String, stCCM21 As String, stCCM22 As String
   Dim stCCM23 As String, stCCM24 As String, stCCM25 As String 'Add by Amy 2025/08/18
   
 On Error GoTo ErrHand
   SaveCloseMain = False: stMsg = ""
   
   stCCM06 = ""
   '商標
   'Modify by Amy 2025/08/18 +CFC
   If strNP02 = "FCT" Or strNP02 = "T" Or strNP02 = "S" Or strNP02 = "CFT" Or strNP02 = "CFC" Then
      stCCM05 = txtReason(0)
      '閉卷
      If ChkClose.Value = vbChecked Then
         stCCM06 = "Y"
      End If
   '專利
   ElseIf strNP02 = "FCP" Or strNP02 = "P" Or strNP02 = "FG" Or strNP02 = "CFP" Then
      stCCM05 = txtReason(1)
      '閉卷
      If Chk7(1).Value = vbChecked Then
         stCCM06 = "Y"
         stCCM07 = Left(CboState(1), Val(InStr(CboState(1), ".")) - 1)
      '不續辦
      Else
         stCCM07 = Left(CboState(0), Val(InStr(CboState(0), ".")) - 1)
      End If
   End If
   stCCM08 = Val(Format(txtAmt(0), "###0")) '總金額
   stCCM09 = Val(Format(txtAmt(1), "###0")) '規費
   stCCM10 = Val(txtAmt(2)) '點數
   stCCM19 = txtA1K(2) '合併列印請款單
   stCCM20 = txtA1K(1) '是否列印申請人
   stCCM21 = txtA1K(3) '請款對象
   stCCM22 = txtA1K(5) '列印對象
   'Add by Amy 2025/08/18
   stCCM23 = txtA1K(4) '固定請款對象
   stCCM24 = txtA1K(6) '固定列印對象
   stCCM25 = txtA1K(0) '帳款已結清
   
   '新增
   If intEdit = 1 Then
      'Modify by Amy  ccm03(下一程序號)可能為空,加CNULL(ChgSQL())
      'Modify by Amy 2025/08/18 +stCCM23/24/25
      stCmd = "Insert into CloseCaseMain(CCM01,CCM02,CCM03,CCM04,CCM05,CCM06,CCM07,CCM08,CCM09,CCM10" & _
                                                                        ",CCM11,CCM12,CCM13,CCM17,CCM18,CCM19,CCM20,CCM21,CCM22,CCM23,CCM24,CCM25) " & _
                        "Values ('" & m_CCM01 & "','" & m_CCM02 & "'," & CNULL(ChgSQL(m_CCM03)) & ",'" & m_CCM04 & "'," & CNULL(ChgSQL(stCCM05)) & "" & _
                        ",'" & stCCM06 & "','" & stCCM07 & "'," & CNULL(stCCM08, True) & "," & CNULL(stCCM09, True) & "," & CNULL(stCCM10, True) & _
                        ",'" & strUserNum & "'," & stDate & "," & stTime & "," & CNULL(ChgSQL(m_CCM17)) & ",'F'," & CNULL(ChgSQL(stCCM19)) & "," & CNULL(ChgSQL(stCCM20)) & _
                        "," & CNULL(ChgSQL(stCCM21)) & "," & CNULL(ChgSQL(stCCM22)) & "," & CNULL(ChgSQL(stCCM23)) & "," & CNULL(ChgSQL(stCCM24)) & "," & CNULL(ChgSQL(stCCM25)) & ")"
   '修改
   Else
      stTmp = stTmp & ",CCM04='" & m_CCM04 & "'" '結案理由代號
      stTmp = stTmp & ",CCM05=" & CNULL(ChgSQL(stCCM05)) '結案說明
      stTmp = stTmp & ",CCM06=" & CNULL(ChgSQL(stCCM06)) '是否閉卷
      
      If strNP02 = "FCP" Or strNP02 = "P" Or strNP02 = "FG" Or strNP02 = "CFP" Then
         'Add by Amy 2025/08/25 已送出後由有期限不續辦,改為閉卷
         If bolUpdNP24 = True Then
            stTmp = stTmp & ",CCM02=" & CNULL(ChgSQL(stCCM02_New)) '總收文號/本所案號
            stTmp = stTmp & ",CCM03=" & CNULL(ChgSQL(stCCM03_New)) '下一程序號
         End If
         'end 2025/08/25
         stTmp = stTmp & ",CCM07='" & stCCM07 & "'" '狀態選項
      Else
         stTmp = stTmp & ",CCM08=" & CNULL(stCCM08) '總金額
         stTmp = stTmp & ",CCM09=" & CNULL(stCCM09) '規費
         stTmp = stTmp & ",CCM10=" & CNULL(stCCM10) '點數
         stTmp = stTmp & ",CCM19=" & CNULL(ChgSQL(stCCM19)) '是否列印申請人
         stTmp = stTmp & ",CCM20=" & CNULL(ChgSQL(stCCM20)) '是否列印請款單
         stTmp = stTmp & ",CCM21=" & CNULL(ChgSQL(stCCM21)) '請款對象
         stTmp = stTmp & ",CCM22=" & CNULL(ChgSQL(stCCM22)) '列印對象
         'Add by Amy 2025/08/18 +stCCM23/24/25
         stTmp = stTmp & ",CCM23=" & CNULL(ChgSQL(stCCM23)) '固定請款對象
         stTmp = stTmp & ",CCM24=" & CNULL(ChgSQL(stCCM24)) '固定列印對象
         stTmp = stTmp & ",CCM25=" & CNULL(ChgSQL(stCCM25)) '帳款已結清
      End If
      If stTmp <> MsgText(601) Then
         stTmp = stTmp & ",CCM14='" & strUserNum & "'"
         stTmp = stTmp & ",CCM15=" & stDate
         stTmp = stTmp & ",CCM16=" & stTime
         'Modify by Amy  ccm03(下一程序號)可能為空
         If Val(m_CCM03) > 0 Then
            stCmd = "And CCM03='" & m_CCM03 & "'"
         End If
         stCmd = "Update CloseCaseMain" & _
                                 " Set " & Mid(stTmp, 2) & _
                                 " Where CCM01='" & m_CCM01 & "' And CCM02='" & m_CCM02 & "' " & stCmd
      End If
      
   End If
   cnnConnection.Execute stCmd, intC
   SaveCloseMain = True

ErrHand:
   If Err.Number <> 0 Then
      stMsg = Err.Description
   End If
End Function

Private Sub txtA1K_GotFocus(Index As Integer)
   TextInverse txtA1K(Index)
End Sub

Private Sub txtA1K_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   '列印申請人
   If Index = 1 Then
      If KeyAscii <> 8 And KeyAscii <> Asc("Y") And KeyAscii <> Asc("N") Then
         KeyAscii = 0
         Beep
      End If
   '帳款已清/合併列印請款單
   ElseIf (Index = 0 Or Index = 2) Then
      If KeyAscii <> 8 And KeyAscii <> Asc("Y") Then
         KeyAscii = 0
         Beep
      End If
   End If
End Sub

Private Sub txtA1K_Validate(Index As Integer, Cancel As Boolean)
   Dim stName As String, stData As String
   
   If Index >= 3 And Index <= 6 Then
      If txtA1K(Index) <> MsgText(601) Then
         If Left(txtA1K(Index), 1) = 代理人編號 Then
            Call ClsPDGetAgent(txtA1K(Index), stName)
         ElseIf Left(txtA1K(Index), 1) = 客戶編號 Then
            stName = GetCustomerName(txtA1K(Index))
         End If
         txtA1K(Index) = GetNewFagent(txtA1K(Index)) '未滿8碼補0
      End If
      lblName(Index - 3).Caption = stName
      If txtA1K(Index) <> MsgText(601) And lblName(Index - 3).Caption = "" Then
         Cancel = True
         Exit Sub
      End If
   End If
   If UCase(Me.ActiveControl.Name) = UCase("txtA1K") Then
      'Modify by Amy 2025/10/02 改共用
      'Call SetA1KData(1, Index)
      Call Pub_CloseSetA1KDataColor(2, Me.Name, strNP02, strNP03, strNP04, strNP05, m_NP07, Me.txtA1K, Index, Index)
   End If
End Sub

'Add by Amy 2025/08/08 管制催款日
Private Sub txtCCD08_Validate(Cancel As Boolean)
   If Trim(txtCCD08) = MsgText(601) Then Exit Sub
   
   Cancel = False
   If CheckIsTaiwanDate(txtCCD08) = False Then
      Cancel = True
   End If
   If ChkWorkDay(Val(txtCCD08) + 19110000) = False Then
      Cancel = True
      MsgBox "請輸入工作日！", vbCritical, "操作錯誤！"
   End If
End Sub

Private Sub txtNotPay_GotFocus()
   TextInverse txtNotPay
End Sub

Private Sub txtNotPay_Validate(Cancel As Boolean)
   txtCCD08.Enabled = False 'Add by Amy 2025/08/08
   If Trim(txtNotPay) = MsgText(601) Then Exit Sub
   
   Cancel = False
   If CheckLengthIsOK(txtNotPay, 500, , "未付帳款") = False Then
      Cancel = True
      Call txtNotPay_GotFocus
   End If
   
   'Add by Amy 2025/08/08 外專程序結案的 未付帳款 欄位中有值,則可輸「管制催款日」
   If (strNP02 = "P" And bolFMPY53374 = True) Or strNP02 <> "P" Then
      txtCCD08.Enabled = True
   End If
End Sub

Private Sub TxtPay_GotFocus(Index As Integer)
   TextInverse TxtPay(Index)
End Sub

'請款項目
Private Sub TxtPay_Validate(Index As Integer, Cancel As Boolean)
   If TxtPay(Index) = MsgText(601) Then Exit Sub
   Select Case Index
      Case 0 '請款項目
         lblPayItemN.Caption = ""
         If txt1(0) <> MsgText(601) And TxtPay(Index) <> MsgText(601) Then
            lblPayItemN = A1j03Query(txt1(0), TxtPay(Index))
            If lblPayItemN.Caption = MsgText(601) Then
               Cancel = True
               MsgBox "無此請款項目請確認！", vbCritical, "操作錯誤！"
               TxtPay(Index).SetFocus
            End If
            'Add by Amy 2025/10/23 測式時輸錯導致後面請款單輸入抓不到資料,可過不檔-秀玲
            If ChkClose.Value = vbChecked And Left(TxtPay(Index), 3) = "703" Then
               'Modify by Amy 2025/11/07 薛經理說要檔住
'               If MsgBox("勾選「閉卷」確定要請 [703-" & GetCaseTypeName(strNP02, "703", 0) & "]相關之請款項目？" & vbCrLf & _
'                           "是:繼續操作" & vbCrLf & _
'                           "否:回前畫面,再確認", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'                  Cancel = True
'                  TxtPay(Index).SetFocus
'               End If
               Cancel = True
               MsgBox "請款項目與結案原因不同" & vbCrLf & _
                              "(有勾選「閉卷」,但請款項目輸 703-" & GetCaseTypeName(strNP02, "703", 0) & ")" & vbCrLf & _
                              "請修正", vbCritical, "操作錯誤！"
               TxtPay(Index).SetFocus
               'end 2025/11/07
            End If
            If ChkClose.Value = vbUnchecked And Left(TxtPay(Index), 3) = "704" Then
               'Modify by Amy 2025/11/07 薛經理說要檔住
               MsgBox "請款項目與結案原因不同" & vbCrLf & _
                              "(沒勾選「閉卷」,但請款項目輸 704-" & GetCaseTypeName(strNP02, "704", 0) & ")" & vbCrLf & _
                              "請修正", vbCritical, "操作錯誤！"
               Cancel = True
               TxtPay(Index).SetFocus
               'end 2025/11/07
            End If
            'end 2025/10/23
         End If
      Case 1 '金額
         If IsNumeric(TxtPay(Index)) = False Then
            Cancel = True
            Call TxtPay_GotFocus(Index)
            Exit Sub
         End If
      Case 2 '折扣
         If IsNumeric(TxtPay(Index)) = False Then
            Cancel = True
            Call TxtPay_GotFocus(Index)
            Exit Sub
         End If
         If Val(TxtPay(Index)) > 100 Then
            Cancel = True
            MsgBox "折扣不可大於100%！", vbExclamation
            Call TxtPay_GotFocus(Index)
            Exit Sub
         End If
      Case 3 '備註
         If CheckLengthIsOK(TxtPay(Index), 500, , "備註") = False Then
            Cancel = True
            Exit Sub
         End If
   End Select
   
End Sub

'其他-說明
Private Sub txtReason_GotFocus(Index As Integer)
   TextInverse txtReason(Index)
End Sub

Private Sub txtReason_Validate(Index As Integer, Cancel As Boolean)
   If Trim(txtReason(Index)) = MsgText(601) Then Exit Sub
   
   Cancel = False
   If CheckLengthIsOK(txtReason(Index), 500, , "理由") = False Then
      Cancel = True
      Call txtReason_GotFocus(Index)
   End If
End Sub

'傳入承辦人員編及案號,判斷FC結案單操作權限
Private Function Pub_FCInuptCloseLimit(ByVal stSales As String, ByVal m_NP02 As String, ByVal m_NP03 As String, ByVal m_NP04 As String, ByVal m_NP05 As String, _
  Optional ByRef IsFMPY53374 As Boolean, Optional ByRef IsFMP As Boolean) As Boolean
   Dim stTP As String
   
   Pub_FCInuptCloseLimit = False
   stTP = stSales
   stSalesAgent = "": stSalesST52 = "": stSalesST53 = "": stSalesST5455 = "" 'Add by Amy 2025/10/14
   'Modify by Amy 2025/09/18 A8013 無法操作 Anny (休假)案子 ex:FCP-052707
   'stSalesAgent = GetCaseDutyAgent(stSales, "", False, , , "1") '承辦人員職代
   stSalesAgent = GetCaseDutyAgent(stSales, "", False) '承辦人員職代
   If stSalesAgent <> MsgText(601) Then stTP = stTP & "," & stSalesAgent
   stSalesST52 = Left(GetSignOffEmp("CM1", m_NP02, m_NP03, m_ApplyNA01, , , stSales), 5) 'FC承辦人員2級主管
   '專利
   If m_NP02 = "P" Then
      IsFMPY53374 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, m_NP02, m_NP03, m_NP04, m_NP05) 'FMP寰華案件
      IsFMP = PUB_ChkIsFMP(m_NP02, m_NP03, m_NP04, m_NP05)  '是否為FMP案件
   End If
   If stSalesST52 <> MsgText(601) Then stTP = stTP & "," & stSalesST52
   'Modify by Amy 2025/09/30 測式外商結案單發現FCT-038823系統收件區可操作,但直接由結案單卻無權限,故調整為各層級主管都可操作(外商規則一樣)-Sindy
   '  ex:FCT-038823 目前智權為沈佳穎,st52為沈佳穎,st53=80030
   'If InStr(m_NP02, "T") > 0 Then
      stSalesST53 = Left(GetSignOffEmp("CM2", m_NP02, m_NP03, m_ApplyNA01, , , stSales), 5) 'FC承辦人員3級主管
   'End If
   If stSalesST53 <> MsgText(601) Then stTP = stTP & "," & stSalesST53
   stSalesST5455 = GetSignOffEmp("CM3", m_NP02, m_NP03, m_ApplyNA01, , , stSales)
   If stSalesST5455 <> MsgText(601) Then stTP = stTP & "," & stSalesST5455
   'end 2025/09/30
   If InStr(stTP, strUserNum) = 0 Then Exit Function
    
   stOptPerson = stTP
   Pub_FCInuptCloseLimit = True
End Function
'intChoose:0-全部 /1-只設定請款單 資訊 /2-只設定總金額、規費、點數
Private Sub SetLock(intChoose As Integer, IsLock As Boolean)
   Dim obj As Object
  
   '請款項目
   If intChoose = 0 Or intChoose = 1 Then
      '請款單 資訊
      For Each obj In txtA1K
         If IsLock = False Then
            '只有 帳款已清 不可改
            If obj.Index <> 0 Then
               obj.Locked = False
            End If
         ElseIf IsLock = True Then
            obj.Locked = True
         End If
      Next
   End If
   If intChoose = 0 Or intChoose = 2 Then
      '總金額、規費、點數
      For Each obj In txtAmt
         obj.Locked = IsLock
      Next
   End If
End Sub

'顯示關連案
Private Sub Show090201_2_1()
   If strNP02 <> "FCP" And strNP02 <> "P" And strNP02 <> "FG" And strNP02 <> "CFP" Then Exit Sub
   Me.Hide
   '有聯合案,衍生案,分割案,一案兩請,香港案,澳門案 彈訊息
   frm090201_2_1.SetParent Me
   frm090201_2_1.cmdOK(1).Visible = False
   frm090201_2_1.cmdOK(4).Visible = False
   frm090201_2_1.StrMenu (strNP02 & "-" & strNP03 & "-" & strNP04 & "-" & strNP05)
   frm090201_2_1.Show
End Sub

Private Sub LoadFile(ByVal stKey As String, strPathFile As String, strLoadPath As String, ByRef stMsg As String)
   Dim ii As Integer, stTmp As String, stRepPath As String, stMsg_Open As String, stMsg_EFile As String, arrTmp
   
   stMsg = ""
   stRepPath = strLoadPath
   If Right(stRepPath, 1) <> "\" Then stRepPath = stRepPath & "\"
   arrTmp = Split(strPathFile, "&")
   For ii = LBound(arrTmp) To UBound(arrTmp)
      '傳入之strLoadPath 最後不可為\
      If PUB_GetAttachFile_CPP(stKey, "" & Replace(arrTmp(ii), stRepPath, ""), strLoadPath, , stTmp) = False Then
         If InStr(stTmp, "檔案已開啟") > 0 Then
            stMsg_Open = stMsg_Open & ";" & stTmp
         ElseIf stTmp <> "" Then
            stMsg_EFile = stMsg_EFile & ";" & stTmp
         End If
      End If
   Next ii
   If stMsg_Open <> "" Or stMsg_EFile <> "" Then
      If stMsg_Open <> "" Then
         stMsg = Replace(Mid(stMsg_Open, 2), ";", vbCrLf)
      End If
      If stMsg_EFile <> "" Then
         If stMsg <> "" Then stMsg = stMsg & vbCrLf
         stMsg = "附件下載失敗:" & vbCrLf & Replace(Mid(stMsg_EFile, 2), ";", vbCrLf) & vbCrLf & _
                        "請洽電腦中心"
      End If
   End If
End Sub

'多筆結案單存檔(FC結案單可上傳.Pdf 及 .Msg)
'intChoose:0-刪檔後再新增(結案單修改) / 1-只新增 /2-只刪檔
'intState:0-智權部結案單 / 1-外商 / 2-外專
Private Sub CloseSaveFile(intChoose As Integer, stCCM01 As String, stCaseNo1 As String, stCaseNo2 As String, stCaseNo3 As String, stCaseNo4 As String _
                , bolLog As Boolean, intState As Integer, stLoadPath As String, ByRef stMsg As String, Optional ByVal m_stSaveFiles As String)
   Dim ii As Integer, bolChk As Boolean, stFileTyp As String, stCmd As String, stWhr As String, stCPP12 As String, stTmp As String, stCCM17 As String
   Dim arrFile
   
   'Memo by Amy frm210133_F 一定會有結案單號 與 frm210133 可能為紙本(結案號為空),才需判斷副檔名
   
   stMsg = ""
   '刪除檔案
   If intChoose = 0 Or intChoose = 2 Then
      bolChk = PUB_ChkIsReplyFile(stCaseNo1 & stCaseNo2 & stCaseNo3 & stCaseNo4, , , , stCCM01, intState, stLoadPath)
      If bolChk = True Then
         stWhr = " And CPP11='" & stCCM01 & "'"
         '檔案改放 FTP,必須在DB資料刪除前執行
         PUB_DelFtpFile2 stCaseNo1 & stCaseNo2 & stCaseNo3 & stCaseNo4, stWhr
         stCmd = "DELETE From CasepaperPdf Where CPP01='" & stCaseNo1 & stCaseNo2 & stCaseNo3 & stCaseNo4 & "' " & stWhr
         If bolLog = True Then Pub_SeekTbLog stCmd '記錄Log
         cnnConnection.Execute stCmd
      End If
   End If
   
   '上傳檔案
   If intChoose = 0 Or intChoose = 1 Then
      arrFile = Split(m_stSaveFiles, "&")
      For ii = LBound(arrFile) To UBound(arrFile)
         stTmp = arrFile(ii)
         If InStrRev(stTmp, " (") > 0 Then
            '排除 C:\Program Files (x86) 狀況
            If UCase(Mid(stTmp, InStrRev(stTmp, " (") + 1, Len("(X86)"))) <> "(X86)" Then
               stTmp = Left(stTmp, InStrRev(stTmp, " (") - 1)
            End If
         End If
         stFileTyp = EMP_回覆單: stCPP12 = ""
         If Right(UCase(stTmp), 4) = ".MSG" Then
            stFileTyp = "RX"
            stCPP12 = "F"
         End If
         '由系統收件區帶入者才需加 stCCM17,由此支再加入者不需帶 (因檔名命名不同)
         stCCM17 = ""
         If InStr(arrFile(ii), m_CCM17) > 0 Then stCCM17 = m_CCM17
         If PUB_UpdReplyFile("" & arrFile(ii), "", stCaseNo1, stCaseNo2, stCaseNo3, stCaseNo4, , stCCM01, stFileTyp, , Me.Name, stCPP12, stCCM17) = False Then
            'PUB_UpdReplyFile 上傳失敗會彈訊息
            stMsg = ";上傳失敗"
            Exit For
         End If
      Next ii
   End If
   If stMsg <> "" Then stMsg = Mid(stMsg, 2)
End Sub

'彈是否詢問
Private Function ChkDataYN() As Boolean
   ChkDataYN = False
   
   If m_strSaveFiles = "" Then
      If MsgBox("郵件電子檔未匯入，要匯入郵件電子檔？" & vbCrLf & _
                           "是:回前畫面匯入電子檔" & vbCrLf & _
                           "否:不需匯入電子檔,繼續送出", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
         Exit Function
      End If
   End If
   'P與CFP 內專結
   If (strNP02 = "P" And bolFMPY53374 = False) Or strNP02 = "CFP" Then
      '結案原因 有「本案大陸代理人是否有最終帳單」字樣
      If InStr(txtReason(1), "本案大陸代理人是否有最終帳單") > 0 Then
         '[未勾選] D/N run C 類工程師報告/閉卷請款/不續辦項目中之「A.未獲指示」及「B.已獲指示」,彈詢問-潘子微
         If Chk6(1).Value = vbUnchecked And Chk9(22).Value = vbUnchecked _
           And Chk8(14).Value = vbUnchecked And Chk8(15).Value = vbUnchecked Then
            If MsgBox("結案原因有給「內專程序」與代理人確認最終帳單之字樣" & vbCrLf & _
                               "但未勾選「D/N run C 類工程師報告」、「閉卷請款」" & vbCrLf & _
                               "                 「 未獲指示」、「已獲指示」是否需回前畫面刪除字樣？" & vbCrLf & _
                               "是:回前畫面刪除字樣" & vbCrLf & _
                               "否:繼續送出", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
               Exit Function
            End If
         '[有勾選] 上述項目,但未改日期
         ElseIf InStr(txtReason(1), "__月__日") > 0 _
           And (Chk6(1).Value = vbChecked Or Chk9(22).Value = vbChecked _
           Or Chk8(14).Value = vbChecked Or Chk8(15).Value = vbChecked) Then
            If MsgBox("結案原因有給「內專程序」與代理人確認最終帳單之字樣" & vbCrLf & _
                               "但日期為「__月__日」是否需回前畫面修改？" & vbCrLf & _
                               "是:回前畫面修改" & vbCrLf & _
                               "否:繼續送出", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbYes Then
               Exit Function
            End If
         End If
      End If
   End If
   
   ChkDataYN = True
End Function

'判斷是否有C類來函掛工程師
Private Function ChkCClassCP14IsEngr(ByVal stCase1 As String, ByVal stCase2 As String, ByVal stCase3 As String, ByVal stCase4 As String) As Boolean
   Dim RsQ As New ADODB.Recordset, intQ As Integer, strQ As String
   
   ChkCClassCP14IsEngr = False
   strQ = GetCClassCP14Sql(stCase1, stCase2, stCase3, stCase4)
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      ChkCClassCP14IsEngr = True
   End If
   Set RsQ = Nothing
End Function

'傳入案號抓進度已發文未請款之案件性質
'回傳資料 intChoose:0-案件性質編號+名稱 / 1-案件性質編號 / 2-案件性質名稱
Private Function GetNotPayCP10(intChoose As Integer, ByVal stFormN As String, ByVal stCNo1 As String, ByVal stCNo2 As String, ByVal stCNo3 As String, ByVal stCNo4 As String) As String
   Dim RsQ As New ADODB.Recordset, intQ As Integer, strQ As String, strTmp(2) As String
   
   GetNotPayCP10 = ""
   strQ = "Select PA01,PA02,PA03,PA04,PA09,CP09,CP10,CP27,CP60,Nvl(Decode(pa09,'000',CPM03,CPM04),CP10) as CP10N " & _
               "From CaseProgress,CasePropertyMap,( " & _
                     "Select PA01,PA02,PA03,PA04,PA09 From Patent " & _
                     "Where PA01='" & stCNo1 & "' And PA02='" & stCNo2 & "' And PA03='" & stCNo3 & "' And PA04='" & stCNo4 & "' " & _
               "Union Select SP01,SP02,SP03,SP04,SP09 From ServicePractice " & _
                     "Where Sp01='" & stCNo1 & "' And Sp02='" & stCNo2 & "' And SP03='" & stCNo3 & "' And SP04='" & stCNo4 & "' " & _
               ") Where PA01=CP01(+) And PA02=CP02(+) And PA03=CP03(+) And PA04=CP04(+) " & _
               "And CP01=CPM01(+) And CP10=CPM02(+) " & _
               "And cp158>0 And cp159=0 And cp16>0 And cp20||cp60 IS NULL "
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      RsQ.MoveFirst
      Do While RsQ.EOF = False
         strTmp(1) = RsQ.Fields("CP10")
         strTmp(2) = RsQ.Fields("CP10N")
         '案件性質 編號
         If intChoose = 1 Then
            strTmp(0) = strTmp(0) & "," & strTmp(1)
         '案件性質 名稱
         ElseIf intChoose = 2 Then
            strTmp(0) = strTmp(0) & "," & strTmp(2)
         '案件性質 編號+名稱
         Else
            strTmp(0) = strTmp(0) & "," & strTmp(1) & "-" & strTmp(2)
         End If
         RsQ.MoveNext
      Loop
      If strTmp(0) <> "" Then GetNotPayCP10 = Mid(strTmp(0), 2)
   End If
   Set RsQ = Nothing
End Function

'Mark by Amy 2025/10/02 不使用,改寫至共用
'intChoose:0-抓全部 資料/1-比對不同設定顏色(單筆)/2-讀取CCM資料設定顏色
Private Sub SetA1KData(intChoose As Integer, Optional ByVal idx As Integer)
'   Dim o_A1K04 As String, o_A1K27 As String, o_A1K28 As String, o_A1K29 As String, o_TM56 As String, o_TM69 As String '目前系統設定
'   Dim ii As Integer, intS As Integer, intE As Integer, stData As String
'
'   Call Pub_GetCloseA1KData(1, Me.Name, strNP02, strNP03, strNP04, strNP05, o_A1K29, m_NP07, o_A1K04, o_A1K27, o_A1K28, o_TM56, o_TM69)
'   '抓全部 資料
'   If intChoose = 0 Then
'      txtA1K(0) = o_A1K29 '帳款已清
'      txtA1K(1) = o_A1K04 '列印申請人
'      txtA1K(2) = "" '合併列印請款單
'      txtA1K(3) = o_A1K28: Call txtA1K_Validate(3, False) '請款對象
'      txtA1K(4) = o_TM56: Call txtA1K_Validate(4, False) '固定請款對象
'      txtA1K(5) = o_A1K27: Call txtA1K_Validate(5, False) '列印對象
'      txtA1K(6) = o_TM69: Call txtA1K_Validate(6, False) '固定列印對象
'   '比對不同設定顏色
'   Else
'      intS = 1: intE = 6
'      If intChoose = 1 Then
'         intS = idx: intE = idx
'      End If
'      For ii = intS To intE
'         stData = ""
'         Select Case ii
'            Case 1 '列印申請人
'               stData = o_A1K04
'            Case 2 '合併列印請款單
'               stData = ""
'            Case 3 '請款對象
'               stData = o_A1K28
'            Case 4 '固定請款對象
'               stData = o_TM56
'            Case 5 '列印對象
'               stData = o_A1K27
'            Case 6 '固定列印對象
'               stData = o_TM69
'         End Select
'         txtA1K(ii).BackColor = &H80000005
'         '資料與目前不同,顯示顏色
'         If txtA1K(ii) <> stData And txtA1K(ii) <> "" Then
'             txtA1K(ii).BackColor = &HC0C0FF
'         End If
'      Next ii
'   End If
End Sub

'Add by Amy 2025/11/07
Private Function ChkGridAmt(stMsg) As Boolean
   Dim ii As Integer, intCol As Integer, stTmp As String
   
   ChkGridAmt = False: stMsg = ""
   intCol = GetColVal(arrAMTCol(), "代號")
   For ii = 1 To GridAMT.Rows - 1
      stTmp = GridAMT.TextMatrix(ii, intCol)
      If stTmp <> MsgText(601) Then
         'Memo 此處有修改要確認 TxtPay_Validate 是否也要改
         '有勾閉卷 且 請款項目輸703開頭
         If ChkClose.Value = vbChecked And Left(stTmp, 3) = "703" Then
            ChkGridAmt = True
            stMsg = "請款項目與結案原因不同" & vbCrLf & _
                              "(有勾選「閉卷」,但請款項目輸 703-" & GetCaseTypeName(strNP02, "703", 0) & ")" & vbCrLf & _
                              "請修正"
         '沒勾閉卷 且 請款項目輸704開頭
         ElseIf ChkClose.Value = vbUnchecked And Left(stTmp, 3) = "704" Then
            ChkGridAmt = True
            stMsg = "請款項目與結案原因不同" & vbCrLf & _
                              "(沒勾選「閉卷」,但請款項目輸 704-" & GetCaseTypeName(strNP02, "704", 0) & ")" & vbCrLf & _
                              "請修正"
         End If
      End If
   Next ii
End Function

Private Sub RefreshGridAMTSeq()
   Dim i As Integer, j As Integer, intAMTRow As Integer, intCol1 As Integer, intCol2 As Integer, stData As String, arrDel
   Dim bol703704 As Boolean, int70XRow As Integer, intSeqno As Integer, stStep As String, stData2 As String '有703704/序號/請款項目代號
   
   intCol1 = GetColVal(arrAMTCol, "代號")
   intCol2 = GetColVal(arrAMTCol, "順序")
   intAMTRow = 1: int70XRow = 1
   '重新整理順序
   For j = 1 To GridAMT.Rows - 1
      intSeqno = 0
      stData = GridAMT.TextMatrix(j, intCol1)
      If Trim(stData) <> MsgText(601) Then
         '703/704 開頭 Or 長度2
         If Left(stData, 3) = "703" Or Left(stData, 3) = "704" Or Len(stData) = 2 Then
            If stData = "703" Or stData = "704" Then
               bol703704 = True
               intSeqno = 1
            Else
               If Grid70X.Rows - 1 = 0 And bol703704 = False Then
                  intSeqno = 2
               '703704 開頭 Or 目前 請款代號 長度2
               Else
                  intSeqno = Grid70X.Rows + 1 '序號
               End If
            End If
            '固定只有2列,GridAMT.RemoveItem 2會錯
            If GridAMT.Rows >= 3 And intSeqno <> 0 Then
               Grid70X.AddItem ""
               Grid70X.TextMatrix(int70XRow, 0) = intSeqno
               For i = 1 To Grid70X.Cols - 1
                  '1-請款代號/2-請款項目名/3-金額/4-折扣/5-備註
                  Grid70X.TextMatrix(int70XRow, i) = GridAMT.TextMatrix(j, i)
               Next i
               int70XRow = int70XRow + 1
               GridAMT.TextMatrix(j, intCol2) = "X"
               stStep = "2"
            End If
         Else
            GridAMT.TextMatrix(j, intCol2) = intAMTRow
            intAMTRow = intAMTRow + 1
         End If
      End If
   Next j
   
   If GridAMT.Rows >= 3 Then
      '資料只有2列
      If GridAMT.Rows = 3 Then
         stData = GridAMT.TextMatrix(1, 1)
         stData2 = GridAMT.TextMatrix(2, 1)
         '第一筆 703 /704 開頭 且 第二筆不是703/704 巾=開頭且大於3碼 Or 2碼,交換
         If ((Left(stData, 3) = "703" Or Left(stData, 3) = "704") And Left(stData2, 3) <> "703" And Left(stData2, 3) <> "704" And Len(stData2) >= 3) _
           Or Len(stData) = 2 Then
            stStep = "2"
            GridAMT.RemoveItem 1
         End If
      Else
         If Grid70X.Rows - 1 > 0 Then
            stStep = "1"
            Grid70X.Sort = 1
            '刪除 順序=X
            j = GridAMT.Rows - 1
            Do While j <> 1
               If GridAMT.TextMatrix(j, intCol2) = "X" Then
                  GridAMT.RemoveItem j
                  j = j - 1
               End If
               j = j - 1
            Loop
         End If
      End If
      '重新整理順序
      If stStep <> "" Then
         '*** 再放回 ***
         If Val(stStep) >= 1 Then
            For j = 1 To Grid70X.Rows - 1
               GridAMT.AddItem ""
               intAMTRow = GridAMT.Rows - 1
               GridAMT.TextMatrix(intAMTRow, intCol2) = intAMTRow - 1 '順序
               For i = 1 To GridAMT.Cols - 1
                  '1-請款代號/2-請款項目名/3-金額/4-折扣/5-備註
                  GridAMT.TextMatrix(intAMTRow, i) = Grid70X.TextMatrix(j, i)
               Next i
               intAMTRow = intAMTRow + 1
            Next j
         End If
         '*** 重新整理順序 ***
         If Val(stStep) >= 2 Then
            For j = 1 To GridAMT.Rows - 1
               If Trim(GridAMT.TextMatrix(j, intCol1)) <> MsgText(601) Then
                  GridAMT.TextMatrix(j, intCol2) = j '順序
               End If
            Next j
         End If
      End If
   End If
   Grid70X.Clear
   Grid70X.Rows = 1
End Sub
