VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm030210 
   BorderStyle     =   1  '單線固定
   Caption         =   "接洽單產生器"
   ClientHeight    =   5720
   ClientLeft      =   140
   ClientTop       =   2420
   ClientWidth     =   8990
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5720
   ScaleWidth      =   8990
   Begin VB.Frame FrameTM136 
      Height          =   280
      Left            =   6120
      TabIndex        =   67
      Top             =   1170
      Width           =   2590
      Begin VB.TextBox TextTM 
         Height          =   270
         Index           =   136
         Left            =   1050
         MaxLength       =   1
         TabIndex        =   11
         Top             =   0
         Width           =   345
      End
      Begin VB.Label Label73 
         AutoSize        =   -1  'True
         Caption         =   "註冊證形式：       1:電子 2:紙本"
         Height          =   180
         Left            =   0
         TabIndex        =   68
         Top             =   40
         Width           =   2520
         WordWrap        =   -1  'True
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   4180
      Left            =   30
      TabIndex        =   41
      Top             =   1500
      Width           =   8920
      _ExtentX        =   15734
      _ExtentY        =   7373
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabHeight       =   441
      TabCaption(0)   =   "案件性質 / 請款內容"
      TabPicture(0)   =   "frm030210.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Frame3"
      Tab(0).Control(1)=   "Grid1(0)"
      Tab(0).Control(2)=   "Frame2"
      Tab(0).Control(3)=   "Grid1(1)"
      Tab(0).Control(4)=   "Label6"
      Tab(0).ControlCount=   5
      TabCaption(1)   =   "請款對象 / 申請人"
      TabPicture(1)   =   "frm030210.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Label14(3)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lblAppl(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label14(2)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lblAppl(0)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Label14(1)"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "lblTM(56)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "Label14(0)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "lblTM(69)"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "Label8"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "lblTM(78)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "lblTM(23)"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "Label10"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "lblTM(79)"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "Label11"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "lblTM(80)"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "Label12"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "lblTM(81)"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "Label13"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "Frame1"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "TextTM(69)"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "TextTM(56)"
      Tab(1).Control(20).Enabled=   0   'False
      Tab(1).Control(21)=   "TextAppl(0)"
      Tab(1).Control(21).Enabled=   0   'False
      Tab(1).Control(22)=   "TextAppl(1)"
      Tab(1).Control(22).Enabled=   0   'False
      Tab(1).Control(23)=   "TextTM(23)"
      Tab(1).Control(23).Enabled=   0   'False
      Tab(1).Control(24)=   "TextTM(78)"
      Tab(1).Control(24).Enabled=   0   'False
      Tab(1).Control(25)=   "TextTM(79)"
      Tab(1).Control(25).Enabled=   0   'False
      Tab(1).Control(26)=   "TextTM(80)"
      Tab(1).Control(26).Enabled=   0   'False
      Tab(1).Control(27)=   "TextTM(81)"
      Tab(1).Control(27).Enabled=   0   'False
      Tab(1).ControlCount=   28
      Begin VB.Frame Frame3 
         Height          =   220
         Left            =   -69330
         TabIndex        =   79
         Top             =   1140
         Width           =   2230
         Begin VB.Label Label16 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "項目目前筆數："
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
            Left            =   120
            TabIndex        =   81
            Top             =   0
            Width           =   1330
         End
         Begin VB.Label LblCntItem 
            BackColor       =   &H00C0FFFF&
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
            Height          =   180
            Left            =   1470
            TabIndex        =   80
            Top             =   0
            Width           =   630
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Height          =   1040
         Index           =   0
         Left            =   -74340
         TabIndex        =   74
         Top             =   3060
         Width           =   7550
         _ExtentX        =   13317
         _ExtentY        =   1834
         _Version        =   393216
         Cols            =   7
         FixedCols       =   0
         BackColorFixed  =   -2147483638
         BackColorBkg    =   12648384
         HighLight       =   0
         FormatString    =   "順序|代碼|案件性質|總金額|規費|點數|電子送件"
         _NumberOfBands  =   1
         _Band(0).Cols   =   7
      End
      Begin VB.TextBox TextTM 
         Height          =   300
         Index           =   81
         Left            =   960
         MaxLength       =   9
         TabIndex        =   33
         Text            =   "999999999"
         Top             =   3120
         Width           =   950
      End
      Begin VB.TextBox TextTM 
         Height          =   300
         Index           =   80
         Left            =   960
         MaxLength       =   9
         TabIndex        =   32
         Text            =   "999999999"
         Top             =   2820
         Width           =   950
      End
      Begin VB.TextBox TextTM 
         Height          =   300
         Index           =   79
         Left            =   960
         MaxLength       =   9
         TabIndex        =   31
         Text            =   "999999999"
         Top             =   2520
         Width           =   950
      End
      Begin VB.TextBox TextTM 
         Height          =   300
         Index           =   78
         Left            =   960
         MaxLength       =   9
         TabIndex        =   30
         Text            =   "999999999"
         Top             =   2220
         Width           =   950
      End
      Begin VB.TextBox TextTM 
         Height          =   300
         Index           =   23
         Left            =   960
         MaxLength       =   9
         TabIndex        =   29
         Text            =   "999999999"
         Top             =   1920
         Width           =   950
      End
      Begin VB.TextBox TextAppl 
         Height          =   300
         Index           =   1
         Left            =   1290
         MaxLength       =   9
         TabIndex        =   28
         Text            =   "999999999"
         Top             =   1530
         Width           =   950
      End
      Begin VB.TextBox TextAppl 
         Height          =   300
         Index           =   0
         Left            =   1290
         MaxLength       =   9
         TabIndex        =   27
         Text            =   "999999999"
         Top             =   1230
         Width           =   950
      End
      Begin VB.TextBox TextTM 
         Height          =   300
         Index           =   56
         Left            =   1290
         MaxLength       =   9
         TabIndex        =   26
         Text            =   "999999999"
         Top             =   930
         Width           =   950
      End
      Begin VB.TextBox TextTM 
         Height          =   300
         Index           =   69
         Left            =   1290
         MaxLength       =   9
         TabIndex        =   25
         Text            =   "999999999"
         Top             =   630
         Width           =   950
      End
      Begin VB.Frame Frame1 
         Height          =   250
         Left            =   120
         TabIndex        =   46
         Top             =   330
         Width           =   8650
         Begin VB.CheckBox Chk1 
            Caption         =   "不請款"
            Height          =   210
            Index           =   0
            Left            =   270
            TabIndex        =   22
            Top             =   0
            Width           =   1240
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "合併列印請款單"
            Height          =   210
            Index           =   1
            Left            =   2040
            TabIndex        =   23
            Top             =   0
            Width           =   1600
         End
         Begin VB.CheckBox Chk1 
            Caption         =   "列印申請人"
            Height          =   210
            Index           =   2
            Left            =   4230
            TabIndex        =   24
            Top             =   0
            Width           =   1600
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   1070
         Left            =   -74910
         TabIndex        =   42
         Top             =   270
         Width           =   8720
         Begin VB.CommandButton cmdOK 
            BackColor       =   &H0080C0FF&
            Caption         =   "修改"
            Height          =   320
            Index           =   2
            Left            =   6870
            Style           =   1  '圖片外觀
            TabIndex        =   18
            Top             =   150
            Width           =   525
         End
         Begin VB.CommandButton cmdOK 
            BackColor       =   &H0080C0FF&
            Caption         =   "清除"
            Height          =   320
            Index           =   4
            Left            =   8010
            Style           =   1  '圖片外觀
            TabIndex        =   20
            Top             =   150
            Width           =   525
         End
         Begin VB.CommandButton cmdOK 
            BackColor       =   &H0080C0FF&
            Caption         =   "刪除"
            Height          =   320
            Index           =   3
            Left            =   7440
            Style           =   1  '圖片外觀
            TabIndex        =   19
            Top             =   150
            Width           =   525
         End
         Begin VB.CommandButton cmdOK 
            BackColor       =   &H0080C0FF&
            Caption         =   "加入"
            Height          =   320
            Index           =   1
            Left            =   6300
            Style           =   1  '圖片外觀
            TabIndex        =   17
            Top             =   150
            Width           =   525
         End
         Begin VB.CheckBox chkWebApp 
            BackColor       =   &H00C0FFFF&
            Caption         =   "電子送件"
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
            Left            =   3630
            TabIndex        =   13
            Top             =   210
            Width           =   1170
         End
         Begin VB.CommandButton cmdOK 
            BackColor       =   &H00FFFFC0&
            Caption         =   "刪除案件性質"
            Height          =   320
            Index           =   0
            Left            =   4890
            Style           =   1  '圖片外觀
            TabIndex        =   21
            Top             =   150
            Width           =   1190
         End
         Begin VB.Label LblItem 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
            Caption         =   "LblItem"
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9
               Charset         =   136
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   1410
            TabIndex        =   73
            Top             =   840
            Width           =   620
         End
         Begin VB.Label Label4 
            BackColor       =   &H00C0FFFF&
            Caption         =   "僅可輸入: + - * / ( ) 數字"
            ForeColor       =   &H000000FF&
            Height          =   190
            Left            =   3240
            TabIndex        =   72
            Top             =   840
            Width           =   2080
         End
         Begin VB.Line Line1 
            BorderWidth     =   2
            X1              =   90
            X2              =   8520
            Y1              =   510
            Y2              =   510
         End
         Begin MSForms.ComboBox Combo1 
            Height          =   290
            Left            =   1350
            TabIndex        =   12
            Top             =   180
            Width           =   2150
            VariousPropertyBits=   679495707
            DisplayStyle    =   3
            Size            =   "3792;512"
            MatchEntry      =   1
            ShowDropButtonWhen=   2
            FontName        =   "新細明體"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "案件性質："
            Height          =   230
            Index           =   1
            Left            =   390
            TabIndex        =   71
            Top             =   240
            Width           =   920
         End
         Begin VB.Label LblCnt 
            AutoSize        =   -1  'True
            BackColor       =   &H00C0FFFF&
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
            Height          =   180
            Left            =   150
            TabIndex        =   70
            Top             =   240
            Width           =   90
         End
         Begin MSForms.TextBox Textdisc 
            Height          =   300
            Left            =   6840
            TabIndex        =   16
            Top             =   540
            Width           =   380
            VariousPropertyBits=   671105051
            MaxLength       =   2
            Size            =   "670;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "折扣：         %"
            Height          =   230
            Index           =   4
            Left            =   6300
            TabIndex        =   45
            Top             =   630
            Width           =   1190
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "金額/算式："
            Height          =   230
            Index           =   99
            Left            =   2190
            TabIndex        =   44
            Top             =   630
            Width           =   980
         End
         Begin VB.Label Label1 
            BackColor       =   &H00C0FFFF&
            Caption         =   "請款項目代號："
            Height          =   230
            Index           =   10
            Left            =   90
            TabIndex        =   43
            Top             =   630
            Width           =   1310
         End
         Begin MSForms.TextBox TextItem 
            Height          =   300
            Left            =   1410
            TabIndex        =   14
            Top             =   540
            Width           =   590
            VariousPropertyBits=   671105051
            MaxLength       =   5
            Size            =   "1041;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
         Begin MSForms.TextBox TextAmt 
            Height          =   300
            Left            =   3210
            TabIndex        =   15
            Top             =   540
            Width           =   2960
            VariousPropertyBits=   671105051
            Size            =   "5221;529"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   180
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid Grid1 
         Height          =   1700
         Index           =   1
         Left            =   -74910
         TabIndex        =   69
         Top             =   1350
         Width           =   8510
         _ExtentX        =   15011
         _ExtentY        =   2999
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         BackColorFixed  =   -2147483638
         BackColorBkg    =   12648384
         HighLight       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
         FormatString    =   "案件性質|代號|請款項目|金額/算式|計算後金額|折扣"
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin VB.Label Label6 
         Caption         =   "總計："
         Height          =   260
         Left            =   -74910
         TabIndex        =   75
         Top             =   3090
         Width           =   590
      End
      Begin VB.Label Label13 
         Caption         =   "申請人5:"
         Height          =   250
         Left            =   90
         TabIndex        =   64
         Top             =   3150
         Width           =   800
      End
      Begin MSForms.Label lblTM 
         Height          =   260
         Index           =   81
         Left            =   1920
         TabIndex        =   63
         Top             =   3150
         Width           =   6890
         Size            =   "12153;459"
         BorderStyle     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label12 
         Caption         =   "申請人4:"
         Height          =   250
         Left            =   90
         TabIndex        =   62
         Top             =   2850
         Width           =   800
      End
      Begin MSForms.Label lblTM 
         Height          =   260
         Index           =   80
         Left            =   1920
         TabIndex        =   61
         Top             =   2850
         Width           =   6890
         Size            =   "12153;459"
         BorderStyle     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label11 
         Caption         =   "申請人3:"
         Height          =   250
         Left            =   90
         TabIndex        =   60
         Top             =   2550
         Width           =   800
      End
      Begin MSForms.Label lblTM 
         Height          =   260
         Index           =   79
         Left            =   1920
         TabIndex        =   59
         Top             =   2550
         Width           =   6890
         Size            =   "12153;459"
         BorderStyle     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label10 
         Caption         =   "申請人2:"
         Height          =   250
         Left            =   90
         TabIndex        =   58
         Top             =   2250
         Width           =   800
      End
      Begin MSForms.Label lblTM 
         Height          =   260
         Index           =   23
         Left            =   1920
         TabIndex        =   57
         Top             =   1950
         Width           =   6890
         Size            =   "12153;459"
         BorderStyle     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.Label lblTM 
         Height          =   260
         Index           =   78
         Left            =   1920
         TabIndex        =   56
         Top             =   2250
         Width           =   6890
         Size            =   "12153;459"
         BorderStyle     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label8 
         Caption         =   "申請人1:"
         Height          =   250
         Left            =   90
         TabIndex        =   55
         Top             =   1950
         Width           =   800
      End
      Begin MSForms.Label lblTM 
         Height          =   260
         Index           =   69
         Left            =   2250
         TabIndex        =   54
         Top             =   660
         Width           =   6560
         Size            =   "11571;459"
         BorderStyle     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         Caption         =   "固定列印對象:"
         Height          =   250
         Index           =   0
         Left            =   90
         TabIndex        =   53
         Top             =   660
         Width           =   1190
      End
      Begin MSForms.Label lblTM 
         Height          =   260
         Index           =   56
         Left            =   2250
         TabIndex        =   52
         Top             =   960
         Width           =   6560
         Size            =   "11571;459"
         BorderStyle     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         Caption         =   "固定請款對象:"
         Height          =   250
         Index           =   1
         Left            =   90
         TabIndex        =   51
         Top             =   960
         Width           =   1190
      End
      Begin MSForms.Label lblAppl 
         Height          =   260
         Index           =   0
         Left            =   2250
         TabIndex        =   50
         Top             =   1260
         Width           =   6560
         Size            =   "11571;459"
         BorderStyle     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         Caption         =   "請款對象:"
         Height          =   250
         Index           =   2
         Left            =   90
         TabIndex        =   49
         Top             =   1260
         Width           =   1190
      End
      Begin MSForms.Label lblAppl 
         Height          =   260
         Index           =   1
         Left            =   2250
         TabIndex        =   48
         Top             =   1560
         Width           =   6560
         Size            =   "11571;459"
         BorderStyle     =   1
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label14 
         Caption         =   "列印對象:"
         Height          =   250
         Index           =   3
         Left            =   90
         TabIndex        =   47
         Top             =   1560
         Width           =   1190
      End
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "尋找(&F)"
      Height          =   375
      Left            =   3080
      TabIndex        =   4
      Top             =   110
      Width           =   800
   End
   Begin VB.TextBox TextKey 
      Height          =   300
      Index           =   1
      Left            =   960
      MaxLength       =   3
      TabIndex        =   0
      TabStop         =   0   'False
      Text            =   "FCT"
      Top             =   160
      Width           =   550
   End
   Begin VB.TextBox TextKey 
      Height          =   300
      Index           =   2
      Left            =   1512
      MaxLength       =   6
      TabIndex        =   1
      Top             =   160
      Width           =   860
   End
   Begin VB.TextBox TextKey 
      Height          =   300
      Index           =   3
      Left            =   2380
      MaxLength       =   1
      TabIndex        =   2
      Top             =   160
      Width           =   260
   End
   Begin VB.TextBox TextKey 
      Height          =   300
      Index           =   4
      Left            =   2650
      MaxLength       =   2
      TabIndex        =   3
      Top             =   160
      Width           =   380
   End
   Begin VB.CommandButton cmdWord 
      Caption         =   "Word 編輯(&W)"
      Enabled         =   0   'False
      Height          =   400
      Left            =   6690
      TabIndex        =   5
      Top             =   40
      Width           =   1190
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7980
      TabIndex        =   6
      Top             =   40
      Width           =   800
   End
   Begin VB.Label Label15 
      Caption         =   "智權人員:"
      Height          =   250
      Left            =   4080
      TabIndex        =   78
      Top             =   210
      Width           =   800
   End
   Begin MSForms.Label lblEmp 
      Height          =   260
      Left            =   4920
      TabIndex        =   77
      Top             =   210
      Width           =   590
      Size            =   "1041;459"
      FontName        =   "新細明體"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblFM2 
      Height          =   260
      Index           =   0
      Left            =   5550
      TabIndex        =   76
      Top             =   210
      Width           =   950
      Size            =   "1676;459"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblTM 
      Height          =   260
      Index           =   15
      Left            =   4920
      TabIndex        =   66
      Top             =   510
      Width           =   2300
      Size            =   "4057;459"
      BorderStyle     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label lblTM 
      Height          =   260
      Index           =   12
      Left            =   960
      TabIndex        =   65
      Top             =   510
      Width           =   2300
      Size            =   "4057;459"
      BorderStyle     =   1
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextCP 
      Height          =   300
      Index           =   6
      Left            =   960
      TabIndex        =   8
      Top             =   1170
      Width           =   950
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1676;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextCP 
      Height          =   300
      Index           =   48
      Left            =   4950
      TabIndex        =   10
      Top             =   1170
      Width           =   950
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1676;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextCP 
      Height          =   300
      Index           =   7
      Left            =   2980
      TabIndex        =   9
      Top             =   1170
      Width           =   950
      VariousPropertyBits=   671105051
      MaxLength       =   7
      Size            =   "1676;529"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label9 
      Caption         =   "承辦期限 :"
      Height          =   250
      Left            =   4080
      TabIndex        =   40
      Top             =   1190
      Width           =   860
   End
   Begin VB.Label Label2 
      Caption         =   "本所期限 :"
      Height          =   260
      Left            =   90
      TabIndex        =   39
      Top             =   1190
      Width           =   860
   End
   Begin VB.Label Label25 
      Caption         =   "法定期限 :"
      Height          =   250
      Left            =   2110
      TabIndex        =   38
      Top             =   1190
      Width           =   860
   End
   Begin MSForms.ComboBox CboTmName 
      Height          =   330
      Left            =   960
      TabIndex        =   7
      Top             =   800
      Width           =   7910
      VariousPropertyBits=   679495711
      DisplayStyle    =   3
      Size            =   "13952;582"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "本所案號:"
      Height          =   250
      Index           =   0
      Left            =   90
      TabIndex        =   37
      Top             =   180
      Width           =   800
   End
   Begin VB.Label Label3 
      Caption         =   "申請案號:"
      Height          =   250
      Left            =   90
      TabIndex        =   36
      Top             =   510
      Width           =   800
   End
   Begin VB.Label Label5 
      Caption         =   "審定號數:"
      Height          =   250
      Left            =   4080
      TabIndex        =   35
      Top             =   510
      Width           =   800
   End
   Begin VB.Label Label7 
      Caption         =   "商標名稱:"
      Height          =   260
      Left            =   90
      TabIndex        =   34
      Top             =   840
      Width           =   800
   End
End
Attribute VB_Name = "frm030210"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Sindy 2024/10/11
Option Explicit

Dim sp() As String, tm() As String
Dim m_strSys As String
Dim m_strArrCP10 As String, m_strArrItem As String
Dim nCol As Long, nRow As Long
Dim arrTmp As Variant


Private Sub Chk1_Click(Index As Integer)
Dim i As Integer
   
   '擇1
   If Chk1(Index).Value = 1 Then
      For i = 0 To 2
         If i <> Index Then
            Chk1(i).Value = 0
         End If
      Next i
   End If
   If Chk1(0).Value = 1 Then '不請款
      Me.cmdOK(3).Enabled = False
   Else
      Me.cmdOK(3).Enabled = True
   End If
End Sub

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdok_Click(Index As Integer)
Dim arrTmp As Variant
Dim strCP10 As String, strCP10Name As String, strItem As String
Dim intJ As Integer, ii As Integer
Dim intDRow As Integer
Dim strDelVal As String
Dim dblAmt As Double
   
On Error GoTo ErrHnd
   
   '點選的案件性質
   If Trim(Combo1.Text) <> "" Then
      arrTmp = Split(Me.Combo1.Text, " ")
      strCP10 = Trim(arrTmp(0))
      strCP10Name = Trim(arrTmp(1))
   Else
      strCP10 = ""
   End If
   If Index = 0 Or Index = 2 Or Index = 3 Then
      If Grid1(0).Rows <= 1 Then
         Exit Sub
      End If
   End If
   If (Index = 0 Or Index = 1) And strCP10 = "" Then
      MsgBox "案件性質不可空白!!", vbInformation
      Combo1.SetFocus
      Exit Sub
   End If
   
   If Chk1(0).Value = 0 Then '非不請款才檢查
      If Index = 1 Or Index = 2 Then
         If Trim(TextItem) = "" Then
            MsgBox "請款項目代號不可空白!!", vbInformation
            TextItem.SetFocus
            Exit Sub
         ElseIf Trim(Me.LblItem) = "" Then
            MsgBox "請款項目代號錯誤，請重新輸入!!", vbInformation
            TextItem.SetFocus
            Exit Sub
         End If
         If Trim(TextAmt) = "" Then
            MsgBox "金額/算式不可空白!!", vbInformation
            TextAmt.SetFocus
            Exit Sub
         End If
      End If
   End If
   
   '目前已加入的案件性質
   m_strArrCP10 = ""
   For intJ = 1 To Grid1(0).Rows - 1
      If Grid1(0).TextMatrix(intJ, 1) <> "" Then
         m_strArrCP10 = m_strArrCP10 & "," & Grid1(0).TextMatrix(intJ, 1)
      End If
   Next intJ
   If m_strArrCP10 <> "" Then m_strArrCP10 = Mid(m_strArrCP10, 2)
   '目前已加入的請款項目
   m_strArrItem = ""
   For intJ = 1 To Grid1(1).Rows - 1
      If Grid1(1).TextMatrix(intJ, 1) <> "" Then
         m_strArrItem = m_strArrItem & "," & Trim(Grid1(1).TextMatrix(intJ, 0)) & Grid1(1).TextMatrix(intJ, 1)
      End If
   Next intJ
   If m_strArrItem <> "" Then m_strArrItem = Mid(m_strArrItem, 2)
   
   If Chk1(0).Value = 0 Then '非不請款才檢查
      If Index = 1 Or Index = 2 Then
         If InStr(m_strArrItem, strCP10 & TextItem) > 0 And Index = 1 Then
            MsgBox "請款項目代號重覆了!!", vbInformation
            TextItem.SetFocus
            Exit Sub
         Else
            'Modify By Sindy 2024/11/18
            If Trim(LblItem.Caption) = "" Then
               MsgBox "請款項目與案件性質不符，請重新輸入!!", vbInformation
               TextItem.SetFocus
               Exit Sub
            Else
               If Len(Trim(TextItem)) > 4 Then
                  If Left(Trim(TextItem), 3) <> strCP10 Then
                     MsgBox "請款項目與案件性質不符，請重新輸入!!", vbInformation
                     TextItem.SetFocus
                     Exit Sub
                  End If
               End If
            End If
         End If
      End If
   End If
   
   Select Case Index
      '*********************************************
      Case 0 '刪除案件性質
      '*********************************************
         If m_strArrCP10 = "" Then
            cmdOK(0).Enabled = False
            Exit Sub
         Else
            If InStr(m_strArrCP10, strCP10) = 0 Then
               MsgBox "無此案件性質可刪除，請重新點選!!", vbInformation
               Combo1.SetFocus
               Exit Sub
            End If
         End If
         '移除此案件性質及相關的請款項目
         If (Grid1(0).Rows - 1) = 1 Then
            Grid1(0).Clear
            InitGrid 7, Grid1(0)
            Call Grid1Head(0)
         Else
            For intJ = Grid1(0).Rows - 1 To 1 Step -1
               If Grid1(0).TextMatrix(intJ, 1) = strCP10 Then
                  Grid1(0).RemoveItem intJ
               End If
            Next intJ
         End If
         If (Grid1(1).Rows - 1) = 1 Then
            Grid1(1).Clear
            InitGrid 6, Grid1(1)
            Call Grid1Head(1)
         Else
            For intJ = Grid1(1).Rows - 1 To 1 Step -1
               If Grid1(1).TextMatrix(intJ, 1) = strCP10 Then
                  Grid1(1).RemoveItem intJ
               End If
            Next intJ
         End If
         
      '*********************************************
      Case 1 '加入(案件性質及請款項目一併加入)
      '*********************************************
         '新增案件性質
         If InStr(m_strArrCP10, strCP10) = 0 Then
            '順序/案件性質代碼/案件性質/電子送件
            LblCnt.Caption = Grid1(0).Rows
            Grid1(0).AddItem ""
            Grid1(0).TextMatrix(LblCnt, 0) = LblCnt
            Grid1(0).TextMatrix(LblCnt, 1) = strCP10
            Grid1(0).TextMatrix(LblCnt, 2) = strCP10Name
            If m_strArrCP10 <> "" Then
               m_strArrCP10 = m_strArrCP10 & "," & strCP10
            Else
               m_strArrCP10 = strCP10
            End If
            Grid1(0).TextMatrix(LblCnt, 6) = IIf(chkWebApp.Value = 1, "Y", "")
         End If
         
         '案件性質代碼/代號/請款項目/金額(算式)/計算後金額/折扣
         If Chk1(0).Value = 0 Then '非不請款
            If CountAmt(dblAmt) = False Then
               TextAmt.SetFocus
               Exit Sub
            End If
            intDRow = Grid1(1).Rows
            Grid1(1).AddItem ""
            Grid1(1).TextMatrix(intDRow, 0) = strCP10
            Grid1(1).TextMatrix(intDRow, 1) = TextItem
            Grid1(1).TextMatrix(intDRow, 2) = LblItem
            Grid1(1).TextMatrix(intDRow, 3) = TextAmt
            Grid1(1).TextMatrix(intDRow, 4) = dblAmt
            Grid1(1).TextMatrix(intDRow, 5) = Textdisc
   '         m_strArrItem = IIf(m_strArrItem = "", "", m_strArrItem & ",") & Trim(Grid1(1).TextMatrix(intDRow, 0)) & Grid1(1).TextMatrix(intDRow, 1)
            Call GetGridData(Grid1(1), 1, intDRow)
         End If
         
      '*********************************************
      Case 2 '(該筆)修改
      '*********************************************
         If Chk1(0).Value = 0 Then '非不請款
            If Val(LblCntItem.Caption) = 0 Then
               MsgBox "無點選欲修改的請款項目!!", vbInformation
               Exit Sub
            ElseIf Grid1(0).TextMatrix(LblCnt, 1) <> Grid1(1).TextMatrix(LblCntItem.Caption, 0) Then
               MsgBox "案件性質與請款項目不一致，不可修改!!", vbInformation
               Exit Sub
            End If
            If CountAmt(dblAmt) = False Then
               TextAmt.SetFocus
               Exit Sub
            End If
            Grid1(1).TextMatrix(LblCntItem.Caption, 3) = TextAmt
            Grid1(1).TextMatrix(LblCntItem.Caption, 4) = dblAmt
            Grid1(1).TextMatrix(LblCntItem.Caption, 5) = Textdisc
            Call GetGridData(Grid1(1), 1, LblCntItem.Caption)
         End If
         Grid1(0).TextMatrix(LblCnt, 6) = IIf(chkWebApp.Value = 1, "Y", "")
         
      '*********************************************
      Case 3 '刪除
      '*********************************************
         If Trim(TextItem) = "" Then
            MsgBox "請款項目代號空白，請點選欲刪除的請款項目!!"
            Exit Sub
         End If
         If InStr(m_strArrItem, ",") = 0 And Left(m_strArrItem, Len(strCP10)) = strCP10 Then
            If MsgBox("最後一筆確定要移除嗎？會連同僅剩的案件性質一併刪除？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
               Exit Sub
            Else
               Call cmdok_Click(0)
               Exit Sub
            End If
         End If
         If MsgBox("確定要移除第" & LblCntItem.Caption & "筆請款項目資料嗎？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
            Exit Sub
         End If
         If Val(LblCntItem.Caption) = 0 Then
            MsgBox "無點選欲刪除的請款項目!!", vbInformation
            Exit Sub
         Else
            '移除該筆請款項目
            strCP10 = Trim(Grid1(1).TextMatrix(LblCntItem.Caption, 0)) '案件性質
            strDelVal = Trim(Grid1(1).TextMatrix(LblCntItem.Caption, 0)) & TextItem '請款項目
            If InStr(m_strArrItem, "," & strDelVal) > 0 Then
               m_strArrItem = Replace(m_strArrItem, "," & strDelVal, "")
            ElseIf InStr(m_strArrItem, strDelVal) > 0 Then
               m_strArrItem = Replace(m_strArrItem, strDelVal, "")
            End If
            Grid1(1).RemoveItem Me.LblCnt
         End If
         '若此案件性質已無請款項目也一併移除
         If InStr(m_strArrItem, "," & strCP10) = 0 Then
            If Left(m_strArrItem, Len(strCP10)) <> strCP10 Then
               For intJ = Grid1(0).Rows - 1 To 1 Step -1
                  If Grid1(0).TextMatrix(intJ, 1) = strCP10 Then
                     Grid1(0).RemoveItem intJ
                  End If
               Next intJ
            End If
         End If
         
      '*********************************************
      Case 4 '清除(請款項目相關欄位值)
      '*********************************************
         Call ClearItem
         Exit Sub
   End Select
'****************************************************************
'重新整理
'****************************************************************
   If Chk1(0).Value = 1 Then '不請款刪除全部請款項目
      Grid1(1).Clear
      InitGrid 6, Grid1(1)
      Call Grid1Head(1)
   End If
   '案件性質的順序
   If Index = 0 Or Index = 3 Then
      For intJ = 1 To Grid1(0).Rows - 1
         Grid1(0).TextMatrix(intJ, 0) = intJ
      Next intJ
   End If
   '目前已加入的案件性質
   m_strArrCP10 = ""
   For intJ = 1 To Grid1(0).Rows - 1
      If Grid1(0).TextMatrix(intJ, 1) <> "" Then
         m_strArrCP10 = m_strArrCP10 & "," & Grid1(0).TextMatrix(intJ, 1)
      End If
   Next intJ
   If m_strArrCP10 <> "" Then m_strArrCP10 = Mid(m_strArrCP10, 2)
   '目前已加入的請款項目
   m_strArrItem = ""
   For intJ = 1 To Grid1(1).Rows - 1
      If Grid1(1).TextMatrix(intJ, 1) <> "" Then
         m_strArrItem = m_strArrItem & "," & Trim(Grid1(1).TextMatrix(intJ, 0)) & Grid1(1).TextMatrix(intJ, 1)
      End If
   Next intJ
   If m_strArrItem <> "" Then m_strArrItem = Mid(m_strArrItem, 2)
'****************************************************************

   '**************************
   '計算
   '**************************
   If (Grid1(0).Rows - 1) = 0 Then
      LblCnt = ""
   Else
      For intJ = 1 To Grid1(0).Rows - 1
         strCP10 = Grid1(0).TextMatrix(intJ, 1) '案件性質
         Grid1(0).TextMatrix(intJ, 3) = ""
         Grid1(0).TextMatrix(intJ, 4) = ""
         Grid1(0).TextMatrix(intJ, 5) = ""
         For ii = 1 To Grid1(1).Rows - 1
            If strCP10 = Grid1(1).TextMatrix(ii, 0) Then
               '總金額
               Grid1(0).TextMatrix(intJ, 3) = Val(Grid1(0).TextMatrix(intJ, 3)) + Val(Grid1(1).TextMatrix(ii, 4))
               '規費
               strItem = Val(Grid1(1).TextMatrix(ii, 1)) '請款項目
               'modify by sonia 2024/12/23 請款項目末2碼為98者為代收代付，也要列入規費
               If Len(Trim(strItem)) = 5 And (Right(Trim(strItem), 2) = 99 Or Right(Trim(strItem), 2) = 98) Then
                  Grid1(0).TextMatrix(intJ, 4) = Val(Grid1(0).TextMatrix(intJ, 4)) + Val(Grid1(1).TextMatrix(ii, 4))
               End If
            End If
         Next ii
         If Grid1(1).Rows - 1 > 0 Then
            '點數
            Grid1(0).TextMatrix(intJ, 5) = Format((Val(Grid1(0).TextMatrix(intJ, 3)) - Val(Grid1(0).TextMatrix(intJ, 4))) / 1000, "0.000")
         End If
      Next intJ
   End If
   Grid1(0).row = Val(LblCnt)
   '**************************
   
   If InStr("2", m_strSys) > 0 Then '商標
      If InStr(m_strArrCP10, "717") > 0 Then
         FrameTM136.Visible = True
      Else
         FrameTM136.Visible = False
      End If
      If InStr(m_strArrCP10, "102") > 0 Or InStr(m_strArrCP10, "301") > 0 Or InStr(m_strArrCP10, "501") > 0 Then
         TextTM(23).Enabled = True
         TextTM(78).Enabled = True
         TextTM(79).Enabled = True
         TextTM(80).Enabled = True
         TextTM(81).Enabled = True
      Else
         TextTM(23).Enabled = False
         TextTM(78).Enabled = False
         TextTM(79).Enabled = False
         TextTM(80).Enabled = False
         TextTM(81).Enabled = False
         '有異動過改回原資料
         If tm(23) <> TextTM(23) Or tm(78) <> TextTM(78) Or tm(79) <> TextTM(79) Or _
            tm(80) <> TextTM(80) Or tm(81) <> TextTM(81) Then
            Call GetTMXYName(True) '取得名稱
         End If
      End If
   End If
   
   '下一程序抓本所期限,法定期限
   '應剔除程序管制之案件性質,若有一筆以上帶期限先到且未過期者
   If m_strArrCP10 <> "" Then
      strExc(0) = "select distinct np09,np08,np23 from nextprogress" & _
                  " where np02='" & tm(1) & "' and np03='" & tm(2) & "'" & _
                  " and np04='" & tm(3) & "' and np05='" & tm(4) & "'" & _
                  " and np06 is null and (np09>=" & strSrvDate(1) & " or np08>=" & strSrvDate(1) & ")" & strNpSqlOfNoSalesDuty & _
                  " and instr('" & m_strArrCP10 & "',np07)>0" & _
                  " order by np09 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Val("" & RsTemp.Fields("np08")) > 0 Then
            TextCP(6) = TransDate(RsTemp.Fields("np08"), 1)
         End If
         If Val("" & RsTemp.Fields("np09")) > 0 Then
            TextCP(7) = TransDate(RsTemp.Fields("np09"), 1)
         End If
      End If
   End If
   
   '結束的相關控制:
   Grid1(0).Visible = False '要執行Grid的Visible(False 到 True)資料才會更新
   Grid1(0).Visible = True
   Grid1(1).Visible = False
   Grid1(1).Visible = True
   '清除欄位值
   Select Case Index
      Case 0 '刪除案件性質
         Call ClearItem
         
      Case 1 '加入
         cmdOK(0).Enabled = True
         Call ClearItem(False)
         
      Case 2 '修改
         Call ClearItem(False)
         
      Case 3 '刪除
         Call ClearItem
         If Grid1(0).Rows = 1 Then
            cmdOK(0).Enabled = False
         End If
   End Select
   If Me.Grid1(0).Rows > 1 Then
      Me.CmdWord.Enabled = True
   Else
      Me.CmdWord.Enabled = False
   End If
   
   Exit Sub
   
ErrHnd:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

Private Function CountAmt(dblAmt As Double) As Boolean
   dblAmt = 0
   CountAmt = False
   '計算後金額:
   strExc(0) = "select " & TextAmt.Text & " from dual "
   intI = 1
   strExc(10) = ""
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0), , True, strExc(10))
   If intI = 1 Then
      CountAmt = True
      dblAmt = RsTemp.Fields(0)
      If Val(Textdisc) > 0 Then
         dblAmt = dblAmt / 100 * Textdisc
      End If
   Else
      If strExc(10) <> "" Then
         If InStr(strExc(10), "遺漏表示式") > 0 Then
            MsgBox "算式( " & TextAmt.Text & " )有誤！請檢查", vbInformation
         Else
            MsgBox strExc(10), vbInformation
         End If
      Else
         MsgBox strExc(0) & vbCrLf & "算式有誤！", vbInformation
      End If
   End If
End Function

Private Sub ClearItem(Optional bolClsCPM As Boolean = True)
   If bolClsCPM = True Then
      LblCnt.Caption = Empty
      Combo1.Text = ""
      chkWebApp.Value = 0
   End If
   TextItem.Text = ""
   TextAmt.Text = ""
   Textdisc.Text = ""
   LblItem.Caption = Empty
   LblCntItem.Caption = Empty
End Sub

Private Sub cmdQuery_Click()
Dim strNation As String
   
   Me.CmdWord.Enabled = False
   Me.cmdOK(0).Enabled = False
   Me.cmdOK(1).Enabled = False
   Me.cmdOK(2).Enabled = False
   Me.cmdOK(3).Enabled = False
   Me.cmdOK(4).Enabled = False
   
   If Trim(TextKey(1)) = "" Then
      MsgBox "請輸入本所案號!!", vbCritical
      TextKey(1).SetFocus
      Call TextKey_GotFocus(1)
      Exit Sub
   End If
   If Trim(TextKey(2)) = "" Then
      MsgBox "請輸入本所案號!!", vbCritical
      TextKey(2).SetFocus
      Call TextKey_GotFocus(2)
      Exit Sub
   End If
   If TextKey(3) = "" Then TextKey(3) = "0"
   If TextKey(4) = "" Then TextKey(4) = "00"
   tm(1) = TextKey(1)
   tm(2) = TextKey(2)
   tm(3) = TextKey(3)
   tm(4) = TextKey(4)
   
   If Grid1(0).Rows > 1 Then
      If MsgBox("是否清除請款項目資料嗎？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
         Call ClearForm(False, False)
      Else
         Call ClearForm(False)
      End If
   Else
      Call ClearForm(False)
   End If
   m_strSys = CheckSys(tm(1))
   FrameTM136.Visible = True
   '基本檔
   If InStr("2", m_strSys) > 0 Then '商標
      If ClsPDReadTrademarkDatabase(tm(), 國外_CF) = False Then
         Exit Sub
      Else
         strNation = tm(10)
         If tm(7) <> "" Then
            CboTmName.AddItem "外：" & tm(7), 0
         End If
         If tm(6) <> "" Then
            CboTmName.AddItem "英：" & tm(6), 0
         End If
         CboTmName.AddItem "中：" & tm(5), 0
         CboTmName.ListIndex = 0
         Me.LblTM(12) = tm(12)
         Me.LblTM(15) = tm(15)
         Me.TextTM(136) = tm(136)
      End If
   ElseIf InStr("5,6", m_strSys) > 0 Then '服務
      sp(1) = tm(1)
      sp(2) = tm(2)
      sp(3) = tm(3)
      sp(4) = tm(4)
      If ClsPDReadServicePracticeDatabase(sp(), 國外_CF) = False Then
         Exit Sub
      Else
         strNation = sp(9)
         If sp(7) <> "" Then
            CboTmName.AddItem "外：" & sp(7), 0
         End If
         If sp(6) <> "" Then
            CboTmName.AddItem "英：" & sp(6), 0
         End If
         CboTmName.AddItem "中：" & sp(5), 0
         CboTmName.ListIndex = 0
         Me.LblTM(12) = sp(11)
         Me.LblTM(15) = sp(14)
         FrameTM136.Visible = False
      End If
   End If
   '目前智權人員
   If tm(1) = "FCT" Then
      lblEmp.Caption = PUB_GetFCTSalesNo(tm(1), tm(2), tm(3), tm(4))
   ElseIf tm(1) = "S" Then
      If strNation = "000" Then
         lblEmp.Caption = PUB_GetFCTSalesNo(tm(1), tm(2), tm(3), tm(4))
      Else
         lblEmp.Caption = PUB_GetAKindSalesNo(tm(1), tm(2), tm(3), tm(4))
      End If
   Else
      lblEmp.Caption = PUB_GetAKindSalesNo(tm(1), tm(2), tm(3), tm(4))
   End If
   lblFM2(0).Caption = GetPrjSalesNM(lblEmp.Caption)
   Call GetTMXYName '取得名稱
   '案件性質
   Call frm090801_New_SetComboCase(1, "", Combo1, Me.TextKey(1).Text, "000 台灣", , , , , "101")
   Me.cmdOK(1).Enabled = True
   Me.cmdOK(2).Enabled = True
   Me.cmdOK(3).Enabled = True
   Me.cmdOK(4).Enabled = True
End Sub

'取得名稱
Private Sub GetTMXYName(Optional onlyReadX As Boolean = False)
Dim strTemp As String

   If InStr("2", m_strSys) > 0 Then '商標
      If onlyReadX = False Then
         Me.TextTM(69) = tm(69)
         Me.TextTM(56) = tm(56)
      End If
      Me.TextTM(23) = tm(23)
      Me.TextTM(78) = tm(78)
      Me.TextTM(79) = tm(79)
      Me.TextTM(80) = tm(80)
      Me.TextTM(81) = tm(81)
   ElseIf InStr("5,6", m_strSys) > 0 Then '服務
      If onlyReadX = False Then
         Me.TextTM(69) = sp(67)
         Me.TextTM(56) = sp(37)
      End If
      Me.TextTM(23) = sp(8)
      Me.TextTM(78) = sp(58)
      Me.TextTM(79) = sp(59)
      Me.TextTM(80) = sp(65)
      Me.TextTM(81) = sp(66)
   End If
   If onlyReadX = False Then
      If TextTM(69) <> "" Then
         'Modify By Sindy 2025/11/13
         'If GetAgentAndState(TextTM(69), strTemp) = True Then Me.lblTM(69).Caption = strTemp
         Me.LblTM(69).Caption = ""
         If GetAgentAndState(TextTM(69), strTemp, , False) = True Then
            Me.LblTM(69).Caption = strTemp
         Else
            If ClsPDGetCustomer(TextTM(69), strTemp) = True Then
               Me.LblTM(69).Caption = strTemp
            Else
               'TextTM(69) = ""
               Me.SSTab1.Tab = 1
               If Me.TextTM(69).Enabled = True Then TextTM(69).SetFocus
            End If
         End If
         '2025/11/13 END
      End If
      If TextTM(56) <> "" Then
         'Modify By Sindy 2025/11/13
         'If GetAgentAndState(TextTM(56), strTemp) = True Then Me.lblTM(56).Caption = strTemp
         Me.LblTM(56).Caption = ""
         If GetAgentAndState(TextTM(56), strTemp, , False) = True Then
            Me.LblTM(56).Caption = strTemp
         Else
            If ClsPDGetCustomer(TextTM(56), strTemp) = True Then
               Me.LblTM(56).Caption = strTemp
            Else
               'TextTM(56) = ""
               Me.SSTab1.Tab = 1
               If Me.TextTM(56).Enabled = True Then TextTM(56).SetFocus
            End If
         End If
         '2025/11/13 END
      End If
   End If
   If TextTM(23) <> "" Then
      If ClsPDGetCustomer(TextTM(23), strTemp) = True Then Me.LblTM(23).Caption = strTemp
   End If
   If TextTM(78) <> "" Then
      If ClsPDGetCustomer(TextTM(78), strTemp) = True Then Me.LblTM(78).Caption = strTemp
   End If
   If TextTM(79) <> "" Then
      If ClsPDGetCustomer(TextTM(79), strTemp) = True Then Me.LblTM(79).Caption = strTemp
   End If
   If TextTM(80) <> "" Then
      If ClsPDGetCustomer(TextTM(80), strTemp) = True Then Me.LblTM(80).Caption = strTemp
   End If
   If TextTM(81) <> "" Then
      If ClsPDGetCustomer(TextTM(81), strTemp) = True Then Me.LblTM(81).Caption = strTemp
   End If
End Sub

Private Sub ClearForm(Optional bolClrPKey As Boolean = True, Optional bolClrItem As Boolean = True)
Dim objTextKey As Object
Dim objText As Object
Dim objLbl As Object
Dim objChk As Object
   
   '****************************************
   If bolClrPKey = True Then
      '保留原輸入的系統類別
      TextKey(1).Tag = TextKey(1).Text
      For Each objTextKey In Me.TextKey
         objTextKey.Text = Empty
      Next
      TextKey(1).Text = TextKey(1).Tag
   End If
   '****************************************
   '屬案件資料
   '****************************************
   For Each objText In Me.TextTM
      objText.Text = Empty
      objText.Tag = Empty
      If objText.Index = 23 Or objText.Index = 78 Or objText.Index = 79 Or objText.Index = 80 Or objText.Index = 81 Then
         objText.Enabled = False
      End If
   Next
   For Each objLbl In Me.LblTM
      objLbl.Caption = Empty
   Next
   TextCP(6).Text = Empty
   TextCP(7).Text = Empty
   TextCP(48).Text = Empty
   lblEmp.Caption = Empty
   lblFM2(0).Caption = Empty
   '****************************************
   
   '請款項目/其他資料
   If bolClrItem = True Then
      For Each objChk In Me.Chk1
         objChk.Value = vbUnchecked
      Next
      CboTmName.Clear
      Combo1.Clear
      LblCnt.Caption = Empty
      chkWebApp.Value = vbUnchecked
      TextItem.Text = Empty
      TextAmt.Text = Empty
      Textdisc.Text = Empty
      For Each objText In Me.TextAppl
         objText.Text = Empty
         objText.Tag = Empty
      Next
      For Each objLbl In Me.lblAppl
         objLbl.Caption = Empty
      Next
      LblItem.Caption = Empty
      m_strArrCP10 = ""
      InitGrid 7, Grid1(0)
      Call Grid1Head(0)
      InitGrid 6, Grid1(1)
      Call Grid1Head(1)
   End If
   
   LblCntItem.Caption = ""
   Me.SSTab1.Tab = 0
End Sub

Private Sub cmdWord_Click()
Dim m_FileName As String
Dim i As Integer, jj As Integer, kk As Integer, intJ As Integer
Dim strName As String, strText As String
Dim intIndex As Integer
Dim strCPM As String
   
On Error GoTo ErrHand
   
   CmdWord.Enabled = False
   '取得樣本檔
   m_FileName = "$" & TextKey(1) & "-" & TextKey(2) & "-" & TextKey(3) & "-" & TextKey(4) & ".Order." & ServerTime & ".doc"
   Call PUB_GetSampleFile(m_FileName, "T22-000001-0-00")
   
   If Dir(App.path & "\" & m_FileName) <> "" Then
      Screen.MousePointer = vbHourglass
      '判斷word是否已開啟
      If g_WordAp Is Nothing Then
RestarWord:
         Set g_WordAp = New Word.Application
         g_WordAp.Visible = True 'False
      End If
'         If Dir(PUB_Getdesktop & "\" & m_TempFileName) <> "" Then
'            Kill PUB_Getdesktop & "\" & m_TempFileName
'         End If
      g_WordAp.Documents.Open App.path & "\" & m_FileName
'         g_WordAp.ActiveDocument.SaveAs PUB_Getdesktop & "\" & m_TempFileName
'         g_WordAp.ActiveDocument.Close
'         g_WordAp.Documents.Open PUB_Getdesktop & "\" & m_TempFileName
      With g_WordAp
         .Selection.WholeStory
         .Selection.Copy
         
         For i = 1 To 18
            strName = ""
            strText = ""
            If i = 1 Then
               strName = "本所案號"
               strText = TextKey(1) & "-" & TextKey(2) & IIf(TextKey(3) & TextKey(4) <> "000", "-" & TextKey(3) & "-" & TextKey(4), "")
            ElseIf i = 2 Then
               strName = "智權人員"
               strText = lblEmp & lblFM2(0)
            ElseIf i = 3 Then
               strName = "商標名稱"
               strText = Mid(CboTmName.Text, 3)
            ElseIf i = 4 Then
               strName = "案件性質"
               For jj = 1 To Grid1(0).Rows - 1
                  If strText <> "" Then
                     strText = strText & vbCrLf
                     strCPM = strCPM & ","
                  End If
                  strExc(10) = "(" & jj & ")" & Grid1(0).TextMatrix(jj, 1) & " " & Grid1(0).TextMatrix(jj, 2)
                  strCPM = strCPM & Grid1(0).TextMatrix(jj, 1)
                  strText = strText & convForm(strExc(10), 26) & IIf(Grid1(0).TextMatrix(jj, 6) = "Y", "電子送件：Y", "")
               Next jj
            ElseIf i = 5 Then
               strName = "法定期限"
               strText = ChangeTStringToTDateString(TextCP(7))
            ElseIf i = 6 Then
               strName = "本所期限"
               strText = ChangeTStringToTDateString(TextCP(6))
            ElseIf i = 7 Then
               strName = "承辦期限"
               strText = ChangeTStringToTDateString(TextCP(48))
            ElseIf i = 8 Then
               strName = "證書形式"
               If FrameTM136.Visible = True Then
                  strText = IIf(TextTM(136) = "1", "證書形式：電子", IIf(TextTM(136) = "2", "證書形式：紙本", ""))
               End If
            ElseIf i = 9 Then
               strName = "申請人"
               For jj = 1 To 5
                  If jj = 1 Then intIndex = 23
                  If jj = 2 Then intIndex = 78
                  If jj = 3 Then intIndex = 79
                  If jj = 4 Then intIndex = 80
                  If jj = 5 Then intIndex = 81
                  If TextTM(intIndex) <> "" Then
                     For kk = 1 To 2
                        If kk = 1 Then
                           strText = "申請人" & jj & "："
                           If jj = 1 Then
                              Call wordFindCol(g_WordAp, strName, strText)
                           Else
                              .Selection.MoveRight Unit:=wdCell '可以新增一列
                              .Selection.TypeText strText
                           End If
                        Else
                           strText = TextTM(intIndex) & " " & LblTM(intIndex)
                           .Selection.MoveRight 'Unit:=wdCharacter, Count:=1
                           .Selection.TypeText strText
                        End If
                     Next kk
                  End If
               Next jj
               strName = "" '*****
            ElseIf i = 10 Then
               strName = "款1"
               strText = IIf(Chk1(0).Value = 1, "Ｖ", "")
            ElseIf i = 11 Then
               strName = "款2"
               strText = IIf(Chk1(1).Value = 1, "Ｖ", "")
            ElseIf i = 12 Then
               strName = "款3"
               strText = IIf(Chk1(2).Value = 1, "Ｖ", "")
            ElseIf i = 13 Then
               strName = "固定列印對象"
               strText = IIf(TextTM(69) <> "", TextTM(69), "")
            ElseIf i = 14 Then
               strName = "固定請款對象"
               strText = IIf(TextTM(56) <> "", TextTM(56), "")
            ElseIf i = 15 Then
               strName = "請款對象"
               strText = IIf(TextAppl(0) <> "", TextAppl(0), "")
            ElseIf i = 16 Then
               strName = "列印對象"
               strText = IIf(TextAppl(1) <> "", TextAppl(1), "")
            ElseIf i = 17 Then
               strName = "總金額"
               For jj = 1 To Grid1(0).Rows - 1
                  If Grid1(0).TextMatrix(jj, 2) <> "" Then
                     For kk = 1 To 6
                        If kk = 1 Then
                           strText = "(" & jj & ")總金額："
                           If jj = 1 Then
                              Call wordFindCol(g_WordAp, strName, strText)
                           Else
                              .Selection.MoveRight Unit:=wdCell
                              .Selection.MoveRight Unit:=wdCell '可以新增一列
                              .Selection.TypeText strText
                           End If
                        Else
                           If kk = 2 Then
                              strText = Format(Grid1(0).TextMatrix(jj, 3), "##,##0")
                           ElseIf kk = 3 Then
                              strText = "規費："
                           ElseIf kk = 4 Then
                              strText = Format(Grid1(0).TextMatrix(jj, 4), "##,##0")
                           ElseIf kk = 5 Then
                              strText = "點數："
                           Else
                              strText = Format(Grid1(0).TextMatrix(jj, 5), "##,##0.000")
                           End If
                           .Selection.MoveRight
                           .Selection.TypeText strText
                        End If
                     Next kk
                  End If
               Next jj
               strName = "" '*****
            ElseIf i = 18 Then
               strName = "代號"
               arrTmp = Split(strCPM, ",")
               For intJ = 0 To UBound(arrTmp) '案件性質
                  For jj = 1 To Grid1(1).Rows - 1 '請款項目
                     If Grid1(1).TextMatrix(jj, 0) = arrTmp(intJ) Then
                        For kk = 1 To 5
                           strText = Grid1(1).TextMatrix(jj, kk)
                           If kk = 1 Then
                              If jj = 1 And intJ = 0 Then
                                 Call wordFindCol(g_WordAp, strName, strText)
                              Else
                                 .Selection.MoveRight Unit:=wdCell '可以新增一列
                                 .Selection.TypeText strText
                              End If
                           Else
                              .Selection.MoveRight
                              .Selection.TypeText strText
                           End If
                        Next kk
                     End If
                  Next jj
               Next intJ
               strName = "" '*****
            End If
            'Find並且置換
            Call wordFindCol(g_WordAp, strName, strText)
ReadNext:
         Next i
      End With
      Screen.MousePointer = vbDefault
'         g_WordAp.ActiveDocument.Save
'         g_WordAp.ActiveDocument.Close
'         MsgBox "檔案已存放在：" & PUB_Getdesktop & "\" & m_TempFileName
      MsgBox "資料已產生完畢!!!"
   Else
      MsgBox "無接洽單的樣本!!!"
   End If
   
   CmdWord.Enabled = True
   Set g_WordAp = Nothing
   Exit Sub
   
ErrHand:
   CmdWord.Enabled = True
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   End If
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Sub

Private Sub wordFindCol(g_WordAp As Object, strName As String, strText As String)
   'Find並且置換
   If Trim(strName) <> "" Then
      g_WordAp.Selection.Find.ClearFormatting
      g_WordAp.Selection.Find.Text = "|#" & strName & "#|"
      g_WordAp.Selection.Find.Replacement.Text = ""
      g_WordAp.Selection.Find.Forward = True
      g_WordAp.Selection.Find.Wrap = wdFindContinue
      g_WordAp.Selection.Find.Format = False
      g_WordAp.Selection.Find.MatchCase = False
      g_WordAp.Selection.Find.MatchWholeWord = False
      g_WordAp.Selection.Find.MatchWildcards = False
      g_WordAp.Selection.Find.MatchSoundsLike = False
      g_WordAp.Selection.Find.MatchAllWordForms = False
      g_WordAp.Selection.Find.MatchByte = True
      g_WordAp.Selection.Find.Execute
      g_WordAp.Selection.Delete
'               If bolFontBorders = True Then
'                  g_WordAp.Selection.Font.Borders(1).LineStyle = wdLineStyleSingle '字元要加框線
'               End If
      g_WordAp.Selection.TypeText strText
      'g_WordAp.Selection.Font.Underline = wdUnderlineSingle '加底線
   End If
End Sub

Private Sub Combo1_Validate(Cancel As Boolean)
Dim intJ As Integer
Dim bolFind As Boolean
Dim strCP10 As String
   
   If Combo1.Text = "" Then Combo1.Tag = "": Exit Sub
   If Combo1.Tag <> Combo1.Text Then
      arrTmp = Split(Me.Combo1.Text, " ")
      strCP10 = Trim(arrTmp(0))
      For intJ = 0 To Combo1.ListCount - 1
         If InStr(Combo1.List(intJ), strCP10 & " ") > 0 Then
            bolFind = True
            Exit For
         End If
      Next intJ
      If bolFind = False Then
         MsgBox "案件性質錯誤，請重新輸入 !", vbCritical
         Combo1.SetFocus
         Cancel = True
         Combo1.Text = ""
      End If
      Combo1.Tag = Combo1.Text
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   ReDim Preserve tm(1 To TF_TM) As String
   ReDim Preserve sp(1 To tf_SP) As String
   ClearForm
End Sub

Private Sub Grid1Head(Index As Integer)
Dim intJ As Integer
   
   If Index = 0 Then
      FixGrid Grid1(Index)
      With Grid1(Index)
         .Visible = False
         .row = 0
         .col = 0: .ColWidth(0) = 500: .Text = "順序"
         .CellAlignment = flexAlignCenterCenter
         .col = 1: .ColWidth(1) = 0: .Text = "案件性質代碼"
         .CellAlignment = flexAlignCenterCenter
         .col = 2: .ColWidth(2) = 2000: .Text = "案件性質"
         .CellAlignment = flexAlignCenterCenter
         .col = 3: .ColWidth(3) = 1200: .Text = "總金額"
         .CellAlignment = flexAlignCenterCenter
         .col = 4: .ColWidth(4) = 1200: .Text = "規費"
         .CellAlignment = flexAlignCenterCenter
         .col = 5: .ColWidth(5) = 1200: .Text = "點數"
         .CellAlignment = flexAlignCenterCenter
         .col = 6: .ColWidth(6) = 1000: .Text = "電子送件"
         .CellAlignment = flexAlignCenterCenter
   '      If .Cols > 5 Then
   '          For intJ = 6 To .Cols
   '              .col = intJ
   '              .ColWidth(intJ) = 0
   '          Next intJ
   '      End If
         .Visible = True
         If .Rows > 1 Then .row = 1
      End With
   Else
      FixGrid Grid1(Index)
      With Grid1(Index)
         .Visible = False
         .row = 0
         .col = 0: .ColWidth(0) = 800: .Text = "案件性質"
         .CellAlignment = flexAlignCenterCenter
         .col = 1: .ColWidth(1) = 800: .Text = "代號"
         .CellAlignment = flexAlignCenterCenter
         .col = 2: .ColWidth(2) = 1500: .Text = "請款項目"
         .CellAlignment = flexAlignCenterCenter
         .col = 3: .ColWidth(3) = 2800: .Text = "金額/算式"
         .CellAlignment = flexAlignCenterCenter
         .col = 4: .ColWidth(4) = 1200: .Text = "計算後金額"
         .CellAlignment = flexAlignCenterCenter
         .col = 5: .ColWidth(5) = 1000: .Text = "折扣"
         .CellAlignment = flexAlignCenterCenter
   '      If .Cols > 5 Then
   '          For intJ = 6 To .Cols
   '              .col = intJ
   '              .ColWidth(intJ) = 0
   '          Next intJ
   '      End If
         .Visible = True
         If .Rows > 1 Then .row = 1
      End With
   End If
End Sub

Private Sub Grid1_Click(Index As Integer)
If Chk1(0).Value = 0 And Index = 0 Then Exit Sub
If nRow <= 0 Then Exit Sub

Grid1(Index).Visible = False
Grid1(Index).col = nCol 'Grid1(Index).MouseCol
Grid1(Index).row = nRow 'Grid1(Index).MouseRow
LblCntItem.Caption = Grid1(Index).row
If Grid1(Index).row <> 0 Then
   Call GetGridData(Grid1(Index), Index, Grid1(Index).row)
End If
Grid1(Index).Visible = True
End Sub

Private Sub Grid1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
'Dim nCol As Long, nRow As Long
   
   getGrdColRow Grid1(Index), x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   Grid1(Index).col = nCol
   Grid1(Index).row = nRow
'   If Me.Grid1(Index).row < 1 And Me.Grid1(Index).Text <> "V" Then
'      If Me.Grid1(Index).Text = "點數" Then
'         If m_blnColOrderAsc = True Then
'            Me.Grid1(Index).Sort = 3  '數值昇冪
'            m_blnColOrderAsc = False
'         Else
'            Me.Grid1(Index).Sort = 4 '數值降冪
'            m_blnColOrderAsc = True
'         End If
'      Else
'         If m_blnColOrderAsc = True Then
'            Me.Grid1(Index).Sort = 5 '字串昇冪
'            m_blnColOrderAsc = False
'         Else
'            Me.Grid1(Index).Sort = 6 '字串降冪
'            m_blnColOrderAsc = True
'         End If
'      End If
'   End If
End Sub

'取得資料視窗中的資料列欄位值
Private Sub GetGridData(oGrid As MSHFlexGrid, Index As Integer, intRow As Integer)
Dim j As Integer, i As Integer
Dim dblPrevRow As Double
   
   'oGrid.Visible = False
   '檢查目前那一列為反白列
   dblPrevRow = 0
   For j = 1 To oGrid.Rows - 1
      oGrid.col = 2
      oGrid.row = j
      If oGrid.CellBackColor <> QBColor(15) Then
         dblPrevRow = j
         Exit For
      End If
   Next j
   '上一筆資料列清除反白
   If dblPrevRow <> intRow Then
      If dblPrevRow > 0 And dblPrevRow <= (oGrid.Rows - 1) Then
         oGrid.row = dblPrevRow
         For i = 0 To oGrid.Cols - 1
            oGrid.col = i
            If oGrid.CellBackColor <> QBColor(15) Then
               oGrid.CellBackColor = QBColor(15) '白
            End If
         Next i
      End If
   End If
   If intRow > 0 Then
      LblCntItem.Caption = intRow
      '目前資料列反白
      oGrid.row = intRow
      dblPrevRow = oGrid.row
      'If oGrid.CellBackColor = QBColor(15) Then
         For i = 0 To oGrid.Cols - 1
            oGrid.col = i
            'If oGrid.CellBackColor = QBColor(15) Then
               oGrid.CellBackColor = &HFFC0C0 '反白
            'End If
         Next i
      'End If

      '顯示資料於畫面上
      If oGrid.TextMatrix(intRow, 1) <> "" Then
         If Index = 0 Then
            LblCnt.Caption = oGrid.TextMatrix(intRow, 0) '順序
            Combo1.Text = Trim(oGrid.TextMatrix(intRow, 1)) & " " & Trim(oGrid.TextMatrix(intRow, 2)) '案件性質
            Combo1.Tag = Combo1.Text
            If Trim(oGrid.TextMatrix(intRow, 6)) = "Y" Then '電子送件
               chkWebApp.Value = 1
            End If
            TextItem = ""
            TextAmt = ""
            Textdisc = ""
            LblItem = ""
         Else
            TextItem = Trim(oGrid.TextMatrix(intRow, 1))
            TextAmt = Trim(oGrid.TextMatrix(intRow, 3))
            Textdisc = Trim(oGrid.TextMatrix(intRow, 5))
            LblItem = Trim(oGrid.TextMatrix(intRow, 2))
            For j = 1 To Grid1(0).Rows - 1
               If Trim(oGrid.TextMatrix(intRow, 0)) = Trim(Grid1(0).TextMatrix(j, 1)) Then
                  LblCnt.Caption = Trim(Grid1(0).TextMatrix(j, 0)) '順序
                  Combo1.Text = Trim(Grid1(0).TextMatrix(j, 1)) & " " & Trim(Grid1(0).TextMatrix(j, 2)) '案件性質
                  Combo1.Tag = Combo1.Text
                  If Trim(Grid1(0).TextMatrix(j, 6)) = "Y" Then '電子送件
                     chkWebApp.Value = 1
                  End If
                  Exit For
               End If
            Next j
         End If
      End If
   End If
   'oGrid.Visible = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm030210 = Nothing
End Sub

Private Sub TextItem_GotFocus()
   TextInverse TextItem
End Sub

Private Sub TextItem_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TextItem_Validate(Cancel As Boolean)
   LblItem.Caption = ""
   If TextItem = MsgText(601) Then
      Exit Sub
   End If
   
   LblItem = A1j03Query(TextKey(1), TextItem)
End Sub

Private Sub TextAmt_GotFocus()
   TextInverse TextAmt
End Sub

Private Sub TextAmt_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
   If Not (KeyAscii >= 48 And KeyAscii <= 57) And KeyAscii <> 8 And KeyAscii <> Asc("*") And KeyAscii <> Asc("/") _
      And KeyAscii <> Asc("+") And KeyAscii <> Asc("-") _
      And KeyAscii <> Asc("(") And KeyAscii <> Asc(")") Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub Textdisc_GotFocus()
   TextInverse Textdisc
End Sub

Private Sub Textdisc_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub Textdisc_LostFocus()
   If Textdisc = "" Then Exit Sub
   If Textdisc > 100 Or Textdisc < 0 Then
      Textdisc = ""
   End If
End Sub

Private Sub TextTM_GotFocus(Index As Integer)
  TextInverse TextTM(Index)
End Sub

Private Sub TextTM_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 23, 56, 69, 78, 79, 80, 81
         KeyAscii = UpperCase(KeyAscii)
      Case 6, 7, 48, 136
         KeyAscii = Pub_NumAscii(KeyAscii)
   End Select
End Sub

Private Sub TextTM_LostFocus(Index As Integer)
Dim strTemp As String
   
   If TextTM(Index) = "" Then Exit Sub
   Select Case Index
      Case 23, 78, 79, 80, 81
         Me.LblTM(Index).Caption = ""
         If ClsPDGetCustomer(TextTM(Index), strTemp) = True Then Me.LblTM(Index).Caption = strTemp
      Case 56, 79
         Me.LblTM(Index).Caption = ""
         If GetAgentAndState(TextTM(Index), strTemp, , False) = True Then
            Me.LblTM(Index).Caption = strTemp
         Else
            If ClsPDGetCustomer(TextTM(Index), strTemp) = True Then
               Me.LblTM(Index).Caption = strTemp
            Else
               TextTM(Index) = ""
            End If
         End If
   End Select
End Sub

Private Sub TextAppl_GotFocus(Index As Integer)
  TextInverse TextAppl(Index)
End Sub

Private Sub TextAppl_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub TextAppl_LostFocus(Index As Integer)
Dim strTemp As String
   
   If TextAppl(Index) = "" Then Exit Sub
   Me.lblAppl(Index).Caption = ""
   If GetAgentAndState(TextAppl(Index), strTemp, , False) = True Then
      Me.lblAppl(Index).Caption = strTemp
   Else
      If ClsPDGetCustomer(TextAppl(Index), strTemp) = True Then
         Me.lblAppl(Index).Caption = strTemp
      Else
         TextAppl(Index) = ""
      End If
   End If
End Sub

Private Sub TextKey_GotFocus(Index As Integer)
  TextInverse TextKey(Index)
End Sub

Private Sub TextKey_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 1
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub TextKey_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 1
         If TextKey(Index) <> "FCT" And TextKey(Index) <> "S" Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
            TextInverse TextKey(Index)
            Cancel = True
         End If
   End Select
End Sub
