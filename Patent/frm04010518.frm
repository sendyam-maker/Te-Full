VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm04010518 
   BorderStyle     =   1  '單線固定
   Caption         =   "專利處收件夾信件處理"
   ClientHeight    =   7640
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   9340
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7640
   ScaleWidth      =   9340
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   270
      Index           =   9
      Left            =   8460
      TabIndex        =   78
      Top             =   2250
      Width           =   195
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   270
      Index           =   8
      Left            =   7530
      TabIndex        =   75
      Top             =   2250
      Width           =   195
   End
   Begin VB.CommandButton cmdDelRow 
      BackColor       =   &H00C0E0FF&
      Caption         =   "刪除"
      Height          =   270
      Left            =   7950
      Style           =   1  '圖片外觀
      TabIndex        =   71
      Top             =   2730
      Width           =   675
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   270
      Index           =   7
      Left            =   6630
      TabIndex        =   62
      Top             =   2240
      Width           =   195
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   270
      Index           =   6
      Left            =   5700
      TabIndex        =   61
      Top             =   2240
      Width           =   195
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   270
      Index           =   5
      Left            =   4800
      TabIndex        =   60
      Top             =   2240
      Width           =   195
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   270
      Index           =   4
      Left            =   3870
      TabIndex        =   59
      Top             =   2240
      Width           =   195
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   270
      Index           =   3
      Left            =   2970
      TabIndex        =   58
      Top             =   2240
      Width           =   195
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   270
      Index           =   2
      Left            =   2040
      TabIndex        =   57
      Top             =   2240
      Width           =   195
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   270
      Index           =   1
      Left            =   1110
      TabIndex        =   56
      Top             =   2240
      Width           =   195
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   270
      Index           =   0
      Left            =   180
      TabIndex        =   55
      Top             =   2240
      Width           =   195
   End
   Begin VB.CommandButton cmdHandRecv 
      Caption         =   "人工啟動接收"
      Height          =   270
      Left            =   7650
      Style           =   1  '圖片外觀
      TabIndex        =   51
      Top             =   1980
      Width           =   1245
   End
   Begin VB.CommandButton cmdRecOutlookQ 
      Caption         =   "郵件接收狀況"
      Height          =   270
      Left            =   6360
      TabIndex        =   50
      Top             =   1980
      Width           =   1245
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "畫面更新(&Q)"
      Height          =   330
      Left            =   4020
      TabIndex        =   1
      Top             =   0
      Width           =   1275
   End
   Begin VB.CommandButton cmdUpdRow 
      BackColor       =   &H00C0E0FF&
      Caption         =   "更正"
      Height          =   270
      Left            =   5400
      Style           =   1  '圖片外觀
      TabIndex        =   35
      Top             =   1980
      Width           =   675
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "記錄查詢"
      Height          =   330
      Left            =   7200
      TabIndex        =   4
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   255
      Left            =   5160
      TabIndex        =   8
      Top             =   7800
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtPathPatent 
      Height          =   270
      Left            =   2040
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "C:\Patent"
      Top             =   7680
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "信件轉入及分類"
      Height          =   330
      Left            =   4620
      TabIndex        =   0
      Top             =   8310
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "整批轉寄"
      Height          =   330
      Left            =   6270
      TabIndex        =   3
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "重新分類"
      Height          =   330
      Left            =   5340
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8115
      TabIndex        =   5
      Top             =   0
      Width           =   800
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame Frame1 
      Caption         =   "資料修改區"
      ForeColor       =   &H00000080&
      Height          =   1670
      Left            =   90
      TabIndex        =   9
      Top             =   360
      Width           =   9200
      Begin VB.TextBox txtPI23 
         Height          =   420
         Left            =   540
         MaxLength       =   200
         TabIndex        =   74
         Top             =   1200
         Width           =   8600
      End
      Begin VB.ComboBox cboPI06 
         Height          =   260
         Left            =   5970
         TabIndex        =   14
         Text            =   "cboPI06"
         Top             =   360
         Width           =   1545
      End
      Begin VB.TextBox txtPI21 
         Height          =   270
         Left            =   7110
         MaxLength       =   2
         TabIndex        =   22
         Top             =   930
         Width           =   375
      End
      Begin VB.TextBox txtPI19 
         Height          =   270
         Left            =   6020
         MaxLength       =   6
         TabIndex        =   20
         Top             =   930
         Width           =   855
      End
      Begin VB.TextBox txtPI18 
         Height          =   270
         Left            =   5520
         MaxLength       =   3
         TabIndex        =   19
         Top             =   930
         Width           =   495
      End
      Begin VB.TextBox txtPI20 
         Height          =   270
         Left            =   6870
         MaxLength       =   1
         TabIndex        =   21
         Top             =   930
         Width           =   255
      End
      Begin VB.ComboBox cboPI05 
         Height          =   260
         ItemData        =   "frm04010518.frx":0000
         Left            =   5970
         List            =   "frm04010518.frx":0002
         Style           =   2  '單純下拉式
         TabIndex        =   17
         Top             =   660
         Width           =   1545
      End
      Begin VB.Label Label10 
         Caption         =   "備註:"
         Height          =   260
         Left            =   60
         TabIndex        =   73
         Top             =   1200
         Width           =   440
      End
      Begin MSForms.ComboBox cboPI06x 
         Height          =   300
         Left            =   5970
         TabIndex        =   13
         Top             =   60
         Width           =   1550
         VariousPropertyBits=   746604571
         DisplayStyle    =   3
         Size            =   "2725;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cboPI06x"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox List1 
         Height          =   1040
         Left            =   7500
         TabIndex        =   15
         Top             =   180
         Width           =   1640
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "2884;1378"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtPI17 
         Height          =   590
         Left            =   510
         TabIndex        =   70
         Top             =   210
         Width           =   4760
         VariousPropertyBits=   -1399830505
         BackColor       =   -2147483633
         BorderStyle     =   1
         ScrollBars      =   3
         Size            =   "8396;1041"
         Value           =   "txtPI17"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         Caption         =   "本所案號:"
         Height          =   230
         Left            =   4710
         TabIndex        =   46
         Top             =   960
         Width           =   800
      End
      Begin VB.Label Label1 
         Caption         =   "(點二下可移除資料)"
         ForeColor       =   &H000000C0&
         Height          =   230
         Index           =   1
         Left            =   7560
         TabIndex        =   33
         Top             =   30
         Width           =   1580
      End
      Begin VB.Label LblPI12 
         Appearance      =   0  '平面
         BorderStyle     =   1  '單線固定
         Caption         =   "Label7"
         ForeColor       =   &H80000008&
         Height          =   230
         Left            =   1260
         TabIndex        =   18
         Top             =   840
         Width           =   3350
      End
      Begin VB.Label Label6 
         Caption         =   "收信日期時間:"
         Height          =   170
         Left            =   60
         TabIndex        =   16
         Top             =   870
         Width           =   1160
      End
      Begin VB.Label Label5 
         Caption         =   "收受者:"
         Height          =   260
         Left            =   5340
         TabIndex        =   12
         Top             =   390
         Width           =   620
      End
      Begin VB.Label Label4 
         Caption         =   "分　類:"
         Height          =   260
         Left            =   5340
         TabIndex        =   11
         Top             =   690
         Width           =   620
      End
      Begin VB.Label Label3 
         Caption         =   "主旨:"
         Height          =   260
         Left            =   60
         TabIndex        =   10
         Top             =   240
         Width           =   440
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3710
      Left            =   30
      TabIndex        =   37
      Top             =   2460
      Width           =   9270
      _ExtentX        =   16351
      _ExtentY        =   6544
      _Version        =   393216
      Tabs            =   10
      Tab             =   1
      TabsPerRow      =   10
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "P程序1"
      TabPicture(0)   =   "frm04010518.frx":0004
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "GRD1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "P程序2"
      TabPicture(1)   =   "frm04010518.frx":0020
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "GRD1(1)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "P程序3"
      TabPicture(2)   =   "frm04010518.frx":003C
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GRD1(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "P程序4"
      TabPicture(3)   =   "frm04010518.frx":0058
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GRD1(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "CFP程序1"
      TabPicture(4)   =   "frm04010518.frx":0074
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "GRD1(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "CFP程序2"
      TabPicture(5)   =   "frm04010518.frx":0090
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "GRD1(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "CFP程序3"
      TabPicture(6)   =   "frm04010518.frx":00AC
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "GRD1(6)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "CFP程序4"
      TabPicture(7)   =   "frm04010518.frx":00C8
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "GRD1(7)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "其他"
      TabPicture(8)   =   "frm04010518.frx":00E4
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "GRD1(8)"
      Tab(8).ControlCount=   1
      TabCaption(9)   =   "垃圾信箱"
      TabPicture(9)   =   "frm04010518.frx":0100
      Tab(9).ControlEnabled=   0   'False
      Tab(9).Control(0)=   "GRD1(9)"
      Tab(9).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm04010518.frx":011C
         Height          =   3350
         Index           =   0
         Left            =   -74970
         TabIndex        =   38
         Top             =   330
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   5909
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm04010518.frx":0131
         Height          =   3350
         Index           =   1
         Left            =   30
         TabIndex        =   39
         Top             =   330
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   5909
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm04010518.frx":0146
         Height          =   3350
         Index           =   2
         Left            =   -74970
         TabIndex        =   40
         Top             =   330
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   5909
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm04010518.frx":015B
         Height          =   3350
         Index           =   3
         Left            =   -74970
         TabIndex        =   41
         Top             =   330
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   5909
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm04010518.frx":0170
         Height          =   3350
         Index           =   4
         Left            =   -74970
         TabIndex        =   42
         Top             =   330
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   5909
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm04010518.frx":0185
         Height          =   3350
         Index           =   5
         Left            =   -74970
         TabIndex        =   43
         Top             =   330
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   5909
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm04010518.frx":019A
         Height          =   3350
         Index           =   6
         Left            =   -74970
         TabIndex        =   44
         Top             =   330
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   5909
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm04010518.frx":01AF
         Height          =   3350
         Index           =   7
         Left            =   -74970
         TabIndex        =   45
         Top             =   330
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   5909
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm04010518.frx":01C4
         Height          =   3350
         Index           =   8
         Left            =   -74970
         TabIndex        =   48
         Top             =   330
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   5909
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm04010518.frx":01D9
         Height          =   3350
         Index           =   9
         Left            =   -74970
         TabIndex        =   77
         Top             =   330
         Width           =   9150
         _ExtentX        =   16140
         _ExtentY        =   5909
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   10
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   6
      End
   End
   Begin VB.TextBox txtShowTransMsg 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H00C0FFC0&
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   24
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   705
      Left            =   990
      Locked          =   -1  'True
      TabIndex        =   49
      Text            =   "信件分類作業中，請稍候..."
      Top             =   0
      Visible         =   0   'False
      Width           =   7160
   End
   Begin SHDocVwCtl.WebBrowser WebBrowserP 
      CausesValidation=   0   'False
      Height          =   1485
      Left            =   2850
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   8190
      Width           =   1605
      ExtentX         =   2831
      ExtentY         =   2619
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   0
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
   Begin VB.ComboBox Combo5 
      Height          =   260
      ItemData        =   "frm04010518.frx":01EE
      Left            =   300
      List            =   "frm04010518.frx":01F0
      Style           =   2  '單純下拉式
      TabIndex        =   67
      Top             =   6330
      Visible         =   0   'False
      Width           =   1485
   End
   Begin VB.TextBox TxtIPDept 
      Height          =   285
      Left            =   1020
      TabIndex        =   23
      Top             =   8130
      Visible         =   0   'False
      Width           =   8505
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   140
      Index           =   8
      Left            =   7770
      TabIndex        =   76
      Top             =   2290
      Width           =   350
   End
   Begin VB.Label Label9 
      Caption         =   $"frm04010518.frx":01F2
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
      Height          =   750
      Left            =   60
      TabIndex        =   72
      Top             =   6840
      Width           =   9210
   End
   Begin MSForms.TextBox txtPI11 
      Height          =   300
      Left            =   1530
      TabIndex        =   69
      Top             =   0
      Visible         =   0   'False
      Width           =   1365
      VariousPropertyBits=   746604575
      Size            =   "2408;529"
      Value           =   "txtPI11"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextII17 
      Height          =   300
      Left            =   5880
      TabIndex        =   68
      Top             =   7710
      Visible         =   0   'False
      Width           =   2385
      VariousPropertyBits=   746604575
      Size            =   "4207;529"
      Value           =   "TextII17"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label1 
      Caption         =   "備註：雙擊”主旨”開啟信件"
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   0
      Left            =   2100
      TabIndex        =   65
      Top             =   2040
      Width           =   2480
   End
   Begin VB.Label LblCC 
      Caption         =   "分信至其他部門信箱將加發副本：(Patent);(TM);(IPDept)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   180
      Left            =   60
      TabIndex        =   66
      Top             =   6630
      Width           =   9210
   End
   Begin VB.Label Label8 
      Alignment       =   1  '靠右對齊
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
      Caption         =   "其他信箱的收受者："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   180
      Left            =   60
      TabIndex        =   64
      Top             =   6180
      Visible         =   0   'False
      Width           =   1760
   End
   Begin VB.Label LblReceiver 
      BackColor       =   &H0080FFFF&
      BorderStyle     =   1  '單線固定
      Caption         =   "LblReceiver"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00004000&
      Height          =   420
      Left            =   1830
      TabIndex        =   63
      Top             =   6180
      Visible         =   0   'False
      Width           =   7440
   End
   Begin MSForms.TextBox TextBoxP 
      Height          =   345
      Left            =   30
      TabIndex        =   53
      Top             =   7830
      Width           =   2025
      VariousPropertyBits=   -1400879077
      ScrollBars      =   2
      Size            =   "3572;609"
      Value           =   "Find簡體字"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox3 
      Height          =   345
      Left            =   2100
      TabIndex        =   52
      Top             =   7830
      Width           =   2025
      VariousPropertyBits=   -1400879077
      ScrollBars      =   2
      Size            =   "3572;609"
      Value           =   "Find簡體字"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H0080FF80&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   140
      Index           =   9
      Left            =   8700
      TabIndex        =   47
      Top             =   2280
      Visible         =   0   'False
      Width           =   350
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   140
      Index           =   7
      Left            =   6870
      TabIndex        =   25
      Top             =   2280
      Width           =   350
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   140
      Index           =   0
      Left            =   420
      TabIndex        =   32
      Top             =   2280
      Width           =   350
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   140
      Index           =   1
      Left            =   1350
      TabIndex        =   31
      Top             =   2280
      Width           =   350
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   140
      Index           =   2
      Left            =   2280
      TabIndex        =   30
      Top             =   2280
      Width           =   350
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   140
      Index           =   3
      Left            =   3210
      TabIndex        =   29
      Top             =   2280
      Width           =   350
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   140
      Index           =   4
      Left            =   4080
      TabIndex        =   28
      Top             =   2280
      Width           =   350
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   140
      Index           =   5
      Left            =   5010
      TabIndex        =   27
      Top             =   2280
      Width           =   350
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   140
      Index           =   6
      Left            =   5940
      TabIndex        =   26
      Top             =   2280
      Width           =   350
   End
   Begin VB.Label LblTotCnt 
      AutoSize        =   -1  'True
      Caption         =   "總筆數(1~10頁籤):"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   260
      TabIndex        =   36
      Top             =   2040
      Width           =   1400
   End
   Begin VB.Label TodayTotCnt 
      AutoSize        =   -1  'True
      Caption         =   "今日總筆數："
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   600
      TabIndex        =   34
      Top             =   80
      Width           =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "信件資料夾："
      Height          =   195
      Left            =   870
      TabIndex        =   6
      Top             =   7710
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.Label LblCntIPDept 
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   5580
      TabIndex        =   24
      Top             =   7890
      Visible         =   0   'False
      Width           =   4125
   End
End
Attribute VB_Name = "frm04010518"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/4/14 Form2.0已修改
'Create By Sindy 2016/8/25
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim dblPrevRow(0 To 9) As Double '記錄目前點選那一筆
Dim m_AttachPath As String
Dim nCol As Long, nRow As Long
Dim pa() As String
Dim sp() As String
Dim m_OldKey As String
Dim bolCboPI06_KeyPress As Boolean 'Add By Sindy 2021/4/14
'Add By Sindy 2025/1/9
Dim intStart As Integer
Dim quyIndex As Integer
Dim int_Pcnt As Integer
'2025/1/9 END


'分類
Private Sub cboPI05_Click()
Dim strUser As String
Dim varTemp As Variant, ii As Integer 'Add By Sindy 2020/3/19
   
   If cboPI05.ListIndex >= 0 Then
      List1.Clear: List1.Tag = ""
      Select Case Trim(Left(cboPI05.Text, 2))
         Case "1" 'P程序1
            strUser = Pub_GetSpecMan("專利處轉信非台灣程序1")
         Case "2" 'P程序2
            strUser = Pub_GetSpecMan("專利處轉信非台灣程序2")
         'Modify By Sindy 2018/6/21
'         Case "3" '亞洲
'            strUser = Pub_GetSpecMan("專利處轉信亞洲程序")
'         Case "4" '歐洲
'            strUser = Pub_GetSpecMan("專利處轉信歐洲程序")
'         Case "5" '美洋非(單)
'            strUser = Pub_GetSpecMan("專利處轉信美洋非洲單號程序")
'         Case "6" '美洋非(雙)
'            strUser = Pub_GetSpecMan("專利處轉信美洋非洲雙號程序")
         Case "3" '美日(單)
            strUser = Pub_GetSpecMan("專利處轉信美日單號程序")
         Case "4" '美日(雙)
            strUser = Pub_GetSpecMan("專利處轉信美日雙號程序")
         Case "5" '美日外(單)
            strUser = Pub_GetSpecMan("專利處轉信美日以外單號程序")
         Case "6" '美日外(雙)
            strUser = Pub_GetSpecMan("專利處轉信美日以外雙號程序")
         '2018/6/21 END
         'Add By Sindy 2020/3/18
         Case 7, 8
         Case Else
            'Modify By Sindy 2022/3/21
            If InStr(cboPI05.Text, "人員空缺") = 0 Then
            '2022/3/21 END
               varTemp = Split(cboPI05.Text, " ")
               If UBound(varTemp) > 0 Then
                  For ii = 0 To Combo5.ListCount - 1
                     If InStr(Combo5.List(ii), varTemp(0)) > 0 Then
                        varTemp = Split(Combo5.List(ii), " ")
                        strUser = varTemp(1)
                        Exit For
                     End If
                  Next ii
               End If
            End If
         '2020/3/18 END
      End Select
      If strUser <> "" Then
         cboPI06.Text = ""
         List1.Tag = strUser & " " & GetPrjSalesNM(strUser)
         List1.AddItem strUser & " " & GetPrjSalesNM(strUser)
         bolCboPI06_KeyPress = False 'Add By Sindy 2021/4/14
      End If
   End If
End Sub

'收受者
Private Sub cboPI06_Click()
'   If bolCboPI06_KeyPress = True Then Exit Sub 'Add By Sindy 2021/4/14
   If cboPI06.ListIndex >= 0 Then
      If InStr(List1.Tag, cboPI06.List(cboPI06.ListIndex)) = 0 Then
         If Trim(List1.Tag) = "" Then List1.Clear
         List1.AddItem cboPI06.List(cboPI06.ListIndex)
         bolCboPI06_KeyPress = False 'Add By Sindy 2021/4/14
         List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboPI06.List(cboPI06.ListIndex)
         cboPI06.Text = ""
      End If
   End If
   If cboPI06.Enabled = False Then
      cboPI06.Text = ""
   End If
End Sub

Private Sub cboPI06_Validate(Cancel As Boolean)
   If cboPI06.Text <> "" Then
      Call cboPI06_LostFocus
'      '檢查人員是否存在或離職
'      If ChkStaffST04(Left(cboPI06, 5)) = True Then
'         cboPI06.SetFocus
'         Call cboPI06_GotFocus
'         Exit Sub
'      End If
      'If Len(Trim(cboPI06.Text)) = 5 Then
         'cboPI06.Text = Left(cboPI06.Text, 5) & " " & GetStaffName(Left(cboPI06.Text, 5), True)
         If List1.ListCount = 0 Then List1.Clear: List1.Tag = ""
         If InStr(List1.Tag, cboPI06.Text) = 0 Then
            List1.AddItem cboPI06.Text
            bolCboPI06_KeyPress = False 'Add By Sindy 2021/4/14
            List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboPI06.Text
         End If
         cboPI06.Text = ""
         'cboPI06.SetFocus 'Add By Sindy 2018/4/24
      'End If
   End If
End Sub
Private Sub cboPI06_GotFocus()
   cboPI06.SelStart = 0
   cboPI06.SelLength = Len(cboPI06.Text)
End Sub
Private Sub cboPI06_KeyPress(KeyAscPI As Integer) 'MSForms.ReturnInteger
   bolCboPI06_KeyPress = True 'Add By Sindy 2021/4/14
   KeyAscPI = UpperCase(KeyAscPI)
End Sub
Private Sub cboPI06_LostFocus()
Dim strText As String
Dim bolFind As Boolean, ii As Integer
   
   If cboPI06.Text <> "" Then
      '依員工姓名抓取員工編號
      strText = GetPrjSalesNM_2(cboPI06.Text)
      If strText <> "" Then
         cboPI06.Text = strText & " " & cboPI06.Text
      Else
         '依員工編號抓取員工姓名
         strText = GetPrjSalesNM(Left(cboPI06.Text, 5))
         If strText <> "" Then
            'Add By Sindy 2021/4/14
            '檢查人員是否離職
            If ChkStaffST04(Left(cboPI06.Text, 5)) = True Then
               cboPI06.SetFocus
               Call cboPI06_GotFocus
               cboPI06.Text = ""
               Exit Sub
            Else
            '2021/4/14 END
               cboPI06.Text = Left(cboPI06.Text, 5) & " " & strText
            End If
         Else
            'Add By Sindy 2021/4/14 檢查是否有在List清單裡, 沒有則不可加入
            bolFind = False
            For ii = 0 To cboPI06.ListCount - 1
               'If cboPI06.Text = cboPI06.List(ii) Then
               If InStr(cboPI06.List(ii), cboPI06.Text) > 0 Then
                  cboPI06.Text = cboPI06.List(ii)
                  bolFind = True: Exit For
               End If
            Next ii
            If bolFind = False Then
               cboPI06.Text = ""
            End If
            '2021/4/14 END
         End If
      End If
   End If
End Sub

'清除反白,並且檢查是否有更新過資料
Private Sub CancelRowColor(Index As Integer, intRow As Integer)
Dim j As Integer

   '清除反白
   GRD1(Index).TextMatrix(intRow, 0) = ""
'   If Not (GRD1(Index).TextMatrix(intRow, 10) = GRD1(Index).TextMatrix(intRow, 7) And _
'           GRD1(Index).TextMatrix(intRow, 11) = GRD1(Index).TextMatrix(intRow, 6) And _
'           GRD1(Index).TextMatrix(intRow, 19) = GRD1(Index).TextMatrix(intRow, 15) And _
'           GRD1(Index).TextMatrix(intRow, 20) = GRD1(Index).TextMatrix(intRow, 16) And _
'           GRD1(Index).TextMatrix(intRow, 21) = GRD1(Index).TextMatrix(intRow, 17) And _
'           GRD1(Index).TextMatrix(intRow, 22) = GRD1(Index).TextMatrix(intRow, 18)) And _
'      (GRD1(Index).TextMatrix(intRow, 10) = "" And GRD1(Index).TextMatrix(intRow, 10) = "") Then
   If GRD1(Index).TextMatrix(intRow, 10) <> "" Or _
      GRD1(Index).TextMatrix(intRow, 11) <> "" Or _
      GRD1(Index).TextMatrix(intRow, 19) <> "" Or _
      GRD1(Index).TextMatrix(intRow, 20) <> "" Or _
      GRD1(Index).TextMatrix(intRow, 21) <> "" Or _
      GRD1(Index).TextMatrix(intRow, 22) <> "" Then
      GRD1(Index).TextMatrix(intRow, 0) = "!"
      cmdSave.Enabled = True: cmdSave.BackColor = &HC0FFC0 '*****
   End If
   GRD1(Index).col = 0
   GRD1(Index).row = intRow
   For j = 0 To GRD1(Index).Cols - 1
      GRD1(Index).col = j
      GRD1(Index).CellBackColor = QBColor(15)
   Next j
   Me.Tag = Replace(Me.Tag, "," & intRow, "") '清除筆數
End Sub

'Add By Sindy 2018/1/5
Private Sub ChkTab_Click(Index As Integer)
   If Val(LblRow(Index)) > 0 Then
      ChkTab(Index).Visible = True
   Else
      ChkTab(Index).Visible = False
   End If
   If InStr(SSTab1.TabCaption(Index), "垃圾信箱") > 0 Then
      ChkTab(Index).Visible = False 'Add By Sindy 2022/1/6 不用轉寄
   End If
End Sub

''Add By Sindy 2018/1/4
'Private Sub Check1_Click()
'Dim mChk As CheckBox
'
'   For Each mChk In ChkTab
'      mChk.Value = Check1.Value
'   Next
'End Sub

'刪除鍵
Private Sub cmdDelRow_Click()
Dim bolHavdDel As Boolean
Dim i As Integer, j As Integer
   
   bolHavdDel = False
   '先檢查是否有資料要刪除
   If GRD1(SSTab1.Tab).Rows - 1 < 1 Then Exit Sub
   If GRD1(SSTab1.Tab).Rows - 1 >= 1 And GRD1(SSTab1.Tab).TextMatrix(1, 13) = "" Then Exit Sub
   For i = 1 To GRD1(SSTab1.Tab).Rows - 1
      If GRD1(SSTab1.Tab).TextMatrix(i, 0) = "V" Then
         bolHavdDel = True
         Exit For
      End If
   Next i
   If bolHavdDel = False Then
      MsgBox "請至少勾選一筆要刪除的資料！"
      Exit Sub
   Else
      If MsgBox("確定要刪除信件嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
         For i = 1 To GRD1(SSTab1.Tab).Rows - 1
            If GRD1(SSTab1.Tab).TextMatrix(i, 0) = "V" Then
               Call CancelRowColor(SSTab1.Tab, i) '清除反白,並且檢查是否有更新過資料
            End If
         Next i
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   Screen.MousePointer = vbHourglass
   For i = 1 To GRD1(SSTab1.Tab).Rows - 1
      If GRD1(SSTab1.Tab).TextMatrix(i, 0) = "V" Then
         strExc(0) = "update PatentInput set " & _
                     " PI07='Y',PI08=" & strSrvDate(1) & ",PI09=" & Right("000000" & ServerTime, 6) & ",PI10='" & strUserNum & "',PI16=" & strSrvDate(1) & _
                     " where PI01=" & GRD1(SSTab1.Tab).TextMatrix(i, 8) & _
                       " and PI02=" & GRD1(SSTab1.Tab).TextMatrix(i, 9) & _
                       " and PI03='" & ChgSQL(GRD1(SSTab1.Tab).TextMatrix(i, 13)) & "'"
         cnnConnection.Execute strExc(0)
         '******************************************************
         '清除反白
         GRD1(SSTab1.Tab).TextMatrix(i, 0) = ""
         GRD1(SSTab1.Tab).col = 0
         GRD1(SSTab1.Tab).row = i
         For j = 0 To GRD1(SSTab1.Tab).Cols - 1
            GRD1(SSTab1.Tab).col = j
            GRD1(SSTab1.Tab).CellBackColor = QBColor(15)
         Next j
         Me.Tag = Replace(Me.Tag, "," & i, "") '清除筆數
         '******************************************************
         LblTotCnt.Caption = "總筆數(" & intStart + 1 & "~" & quyIndex + 1 & "頁籤): " & _
                             Val(Replace(LblTotCnt.Caption, "總筆數(" & intStart + 1 & "~" & quyIndex + 1 & "頁籤):", "")) - 1
         strExc(0) = LblRow(SSTab1.Tab).Caption
         strExc(0) = Val(strExc(0)) - 1
         'Modify By Sindy 2025/1/10
         LblRow(SSTab1.Tab).Caption = strExc(0): ChkTab_Click (SSTab1.Tab)
'         If SSTab1.Tab = 0 Then LblRow(0).Caption = strExc(0): ChkTab_Click (0) 'P程序1
'         If SSTab1.Tab = 1 Then LblRow(1).Caption = strExc(0): ChkTab_Click (1) 'P程序2
'         If SSTab1.Tab = 2 Then LblRow(2).Caption = strExc(0): ChkTab_Click (2) 'CFP程序1 美日(單) --亞洲
'         If SSTab1.Tab = 3 Then LblRow(3).Caption = strExc(0): ChkTab_Click (3) 'CFP程序2 美日(雙) --歐洲
'         If SSTab1.Tab = 4 Then LblRow(4).Caption = strExc(0): ChkTab_Click (4) 'CFP程序3 美日外(單) --美洋非(單)
'         If SSTab1.Tab = 5 Then LblRow(5).Caption = strExc(0): ChkTab_Click (5) 'CFP程序4 美日外(雙) --美洋非(雙)
'         If SSTab1.Tab = 6 Then LblRow(6).Caption = strExc(0): ChkTab_Click (6) '其他
'         If SSTab1.Tab = 7 Then LblRow(7).Caption = strExc(0): ChkTab_Click (7) '垃圾信箱
         '2025/1/10 END
         GRD1(SSTab1.Tab).RowHeight(i) = 0
      End If
   Next i
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox " 刪除失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdExit_Click()
   '先更新資料
   If cmdSave.Enabled = True Then
      If MsgBox("資料有異動，是否要存檔？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbYes Then
         Call CmdSave_Click
      End If
   End If
   
   Unload Me
End Sub

'Add By Sindy 2017/11/15
Private Sub cmdHandRecv_Click()
   strExc(0) = "select mrl01 from mailreceivelog" & _
               " where mrl01='" & Left(Patent收件匣, 2) & "'" & _
               " and mrl09='A'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      cmdHandRecv.BackColor = &HC0FFC0
      MsgBox "正在等待信件接件！", vbInformation
'      Timer1.Interval = 100
   Else
      If MsgBox("確定要新增「接收信件」的排程嗎？" & vbCrLf & vbCrLf & _
                "(啟動排程至少需等1分鐘後...)", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbYes Then
         strExc(0) = "insert into mailreceivelog(mrl01,mrl02,mrl03,mrl05,mrl09)" & _
                     "values('" & Left(Patent收件匣, 2) & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",'" & strUserNum & "','A')"
         cnnConnection.Execute strExc(0)
         cmdHandRecv.BackColor = &HC0FFC0
'         Timer1.Interval = 100
      Else
         cmdHandRecv.BackColor = &H8000000F
'         Timer1.Interval = 0
      End If
   End If
End Sub

Private Sub cmdHistory_Click()
Dim nFrm As Form
   
   '檢查表單是否已開啟，若是，則關閉
   For Each nFrm In Forms
      If StrComp(nFrm.Name, "frm06010613", vbTextCompare) = 0 Then
         Unload frm06010613
      End If
   Next
   
   Call frm06010613.SetParent(Me)
   frm06010613.m_WorkType = 0 '信箱主檔 Add By Sindy 2017/12/12
'   frm06010613.Combo1.Enabled = False
'   'frm06010613.cboIR13 = strUserNum & " " & GetPrjSalesNM(strUserNum) '轉寄人員
'   frm06010613.cboIR13.Tag = frm06010613.cboIR13.Text
'   frm06010613.txtDate(2) = strSrvDate(2) '轉寄起始日期
'   frm06010613.txtDate(3) = strSrvDate(2) '轉寄截止日期
'   frm06010613.Caption = "信件記錄查詢 - 信件主檔"
   frm06010613.Show
   Me.Hide
End Sub

Private Sub cmdQuery_Click()
   QueryData
End Sub

Private Sub cmdRecOutlookQ_Click()
   frm06010615.m_QueryType = "P"
   frm06010615.Hide
   frm06010615.cmdQuery_Click
   frm06010615.Show vbModal
End Sub

'整批轉寄
Private Sub cmdSendMail_Click()
Dim intTab As Long
Dim strFileName As String
Dim tmpArr As Variant
Dim strTo As String
Dim strSubject As String
Dim strContext As String
Dim strEMailTo As String '串要發通知信的人員
Dim strUpdTime As String
Dim strRecordRow As String '記錄處理到那一筆資料 Modify By Sindy 2018/1/2
Dim bolReadFile As Boolean 'Add By Sindy 2018/1/3
Dim i As Integer, j As Integer
Dim strPI11 As String, strPI12 As String, strPI13 As String, strPI17 As String
Dim strTi03 As String, strTi03_2 As String
Dim bolSaveEFile As Boolean, stFtpPath As String
Dim strToCC As String 'Add By Sindy 2019/7/17
   
   Screen.MousePointer = vbHourglass
   '先更新資料
   If cmdSave.Enabled = True Then
      Call CmdSave_Click
   Else
      Call QueryData
   End If
   
On Error GoTo ErrHand
   
   cmdSendMail.Enabled = False
   strEMailTo = ""
   For intTab = intStart To quyIndex 'P程序1~垃圾信箱
      If ChkTab(intTab).Value = 1 And ChkTab(intTab).Visible = True Then 'Add By Sindy 2018/1/4
      For i = 1 To GRD1(intTab).Rows - 1
         If Trim(GRD1(intTab).TextMatrix(i, 13)) <> "" And _
            Trim(GRD1(intTab).TextMatrix(i, 7)) <> "" Then '有檔名流水號有收受者時,就更新資料
            strRecordRow = SSTab1.TabCaption(intTab) & ":" & GRD1(intTab).TextMatrix(i, 8) & "-" & _
                           GRD1(intTab).TextMatrix(i, 9) & "-" & GRD1(intTab).TextMatrix(i, 13) & _
                           "-" & GRD1(intTab).TextMatrix(i, 7) 'Modify By Sindy 2018/1/2
            'Add By Sindy 2017/12/22 收受者有非程序時,先下載檔案
            tmpArr = Split(GRD1(intTab).TextMatrix(i, 7), ";")
            bolReadFile = False 'Add By Sindy 2018/1/3
            For j = 0 To UBound(tmpArr)
               If tmpArr(j) <> "" Then
                  'Add By Sindy 2018/4/10 讀不到ST03值的收受者資料可能有問題,不處理
                  If InStr(tmpArr(j), "@") = 0 Then
                     If PUB_GetST03(CStr(tmpArr(j))) = "" Then
                        'Re:P-116847案不續辦-年費
                        '人員輸入AA3014,輸錯;讀不到ST03值的收受者資料,此判斷是需要的
                        'PUB_SendMail strUserNum, "97038", "", GRD1(intTab).TextMatrix(i, 8) & "-" & GRD1(intTab).TextMatrix(i, 9) & "-" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)), GRD1(intTab).TextMatrix(i, 1) & vbCrLf & GRD1(intTab).TextMatrix(i, 7) & "讀不到ST03值的收受者資料可能有問題,不處理"
                        GoTo ReadNext
                     End If
                  End If
                  '2018/4/10 END
                  If PUB_GetST03(CStr(tmpArr(j))) <> "P12" Then '非專利程序
                     '讀取檔案
                     strRecordRow = strRecordRow & "[讀檔]"
                     'If Dir(m_AttachPath & "\" & strFileName, vbDirectory) = "" Then
                     If bolReadFile = False Then 'Add By Sindy 2018/1/3
                        strFileName = Mid(GRD1(intTab).TextMatrix(i, 14), InStrRev(GRD1(intTab).TextMatrix(i, 14), "/") + 1)
                        If GetAttachFile(GRD1(intTab).TextMatrix(i, 8), GRD1(intTab).TextMatrix(i, 9), GRD1(intTab).TextMatrix(i, 13), strFileName, m_AttachPath & "\" & strFileName) = True Then
                           bolReadFile = True 'Add By Sindy 2018/1/3
                           '信包信裡面沒有附件檔,等待下載信件
                           Do While Dir(strFileName, vbDirectory) = ""
                              DoEvents
                           Loop
                        'Add By Sindy 2018/1/3
                        Else
                           Call QueryData
                           cmdSendMail.Enabled = True
                           Exit Sub
                        '2018/1/3 END
                        End If
                     End If
                  End If
               End If
            Next j
            '2017/12/22 END
            
'            'Add By Sindy 2018/4/24 再檢查一次電子檔是否存在
'            If bolReadFile = True Then
'               If Dir(strFileName) <> "" Then
'                  MsgBox "電子檔不存在(" & strFileName & ")!", vbExclamation
'                  GoTo ReadNext
'               End If
'            End If
            
            cnnConnection.BeginTrans
            strRecordRow = strRecordRow & "[存檔(1)]"
            strUpdTime = Right("000000" & ServerTime, 6)
            GRD1(intTab).TextMatrix(i, 7) = PUB_IR04DataMakeUp(GRD1(intTab).TextMatrix(i, 7))
            tmpArr = Split(GRD1(intTab).TextMatrix(i, 7), ";")
            strTo = "": strToCC = "" 'Add Sindy 2022/2/7
            For j = 0 To UBound(tmpArr)
               If tmpArr(j) <> "" Then
                  If InStr(tmpArr(j), "@") > 0 Then tmpArr(j) = Mid(tmpArr(j), 1, InStr(tmpArr(j), "@") - 1)
                  strRecordRow = strRecordRow & "[存檔(2)]"
                  '新增郵件轉信讀取記錄
                  strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR15)" & _
                              " values(" & GRD1(intTab).TextMatrix(i, 8) & _
                                       "," & GRD1(intTab).TextMatrix(i, 9) & _
                                       ",'" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)) & "'" & _
                                       ",'" & tmpArr(j) & "'," & strSrvDate(1) & "," & _
                                       strUpdTime & ",'" & strUserNum & "','Y')"
                  cnnConnection.Execute strExc(0)
                  
                  'Add By Sindy 2019/6/20 轉入商標處信件系統
                  If UCase(tmpArr(j)) = "TM" And strSrvDate(1) >= TM分信系統啟用日 Then
                     '讀取專利處收件夾資料
                     strExc(0) = "select pi11,pi12,pi13,pi17 from PatentInput" & _
                                 " where pi01=" & GRD1(intTab).TextMatrix(i, 8) & _
                                 " and pi02=" & GRD1(intTab).TextMatrix(i, 9) & _
                                 " and pi03='" & GRD1(intTab).TextMatrix(i, 13) & "'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strPI11 = "" & RsTemp.Fields("pi11")
                        strPI12 = "" & RsTemp.Fields("pi12")
                        strPI13 = "" & RsTemp.Fields("pi13")
                        strPI17 = "" & RsTemp.Fields("pi17")
                     End If
                     
                     '商標處收件夾資料,最大流水號
                     'Modify By Sindy 2019/12/2 自動給號,才能 Keep PKey
                     strTi03 = AutoNoByDate("T", 4)
                     '2019/12/2 END
                     strTi03_2 = strSrvDate(1) & strUpdTime & "." & strTi03 & ".msg"
                     '存實體檔案到TMInput
                     bolSaveEFile = PUB_PutFtpFile(strFileName, strSrvDate(1), strTi03_2, stFtpPath, UCase("TMInput"))
                     If bolSaveEFile = True Then
                        '存資料到商標處收件夾資料
                        strSql = "insert into TMInput(Ti01,Ti02,Ti03,Ti04,Ti05,Ti11,Ti12,Ti13,Ti14,Ti17,Ti15)" & _
                                 " values(" & strSrvDate(1) & "," & strUpdTime & _
                                 ",'" & strTi03 & "','" & strUserNum & "',null" & _
                                 "," & CNULL(ChgSQL(strPI11)) & "," & strPI12 & "," & CNULL(strPI13) & _
                                 ",'" & ChgSQL(stFtpPath) & "','" & ChgSQL(strPI17) & "','Patent')"
                        cnnConnection.Execute strSql
                        
                        'Add By Sindy 2019/7/17 增加副本給商標處主管
                        If OL_SendNotifyMailCC("Patent", "TM", strFileName, strPI17, strSrvDate(1), strUpdTime, strTi03, OL_TmMailCC, strSrvDate(1), strUpdTime) = False Then
                           GoTo ErrHand
                        End If
                        
                        '更新專利處收件夾資料:轉寄人員,商標處新流水號
                        strExc(0) = "update PatentInput set " & _
                                    " PI10='" & strUserNum & "',PI22='" & strTi03 & "'" & _
                                    " where PI01=" & GRD1(intTab).TextMatrix(i, 8) & _
                                      " and PI02=" & GRD1(intTab).TextMatrix(i, 9) & _
                                      " and PI03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)) & "'" & _
                                      " and PI08=0"
                        cnnConnection.Execute strExc(0)
                        
                        '該收受者上刪除日期時間人員
                        strExc(0) = "update InputRecord set " & _
                                    " ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
                                    " where ir01=" & GRD1(intTab).TextMatrix(i, 8) & _
                                      " and ir02=" & GRD1(intTab).TextMatrix(i, 9) & _
                                      " and ir03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)) & "'" & _
                                      " and ir04='" & tmpArr(j) & "'"
                        cnnConnection.Execute strExc(0)

'                        GRD1(intTab).TextMatrix(i, 7) = Replace(GRD1(intTab).TextMatrix(i, 7), "tm", "") '收受者拿掉tm
'                        GRD1(intTab).TextMatrix(i, 7) = Replace(GRD1(intTab).TextMatrix(i, 7), ";;", ";")
'                        If GRD1(intTab).TextMatrix(i, 7) = ";" Then GRD1(intTab).TextMatrix(i, 7) = ""
'                        If GRD1(intTab).TextMatrix(i, 7) <> "" Then
'                           If Left(GRD1(intTab).TextMatrix(i, 7), 1) = ";" Then GRD1(intTab).TextMatrix(i, 7) = Mid(GRD1(intTab).TextMatrix(i, 7), 2)
'                           If Right(GRD1(intTab).TextMatrix(i, 7), 1) = ";" Then GRD1(intTab).TextMatrix(i, 7) = Mid(GRD1(intTab).TextMatrix(i, 7), 1, Len(GRD1(intTab).TextMatrix(i, 7)) - 1)
'                        End If
                     End If
                  
                  'Modify By Sindy 2017/12/22
                  '非程序人員時，轉寄Outlook並且該收受者上刪除日期時間人員
                  ElseIf PUB_GetST03(CStr(tmpArr(j))) <> "P12" Then '非專利程序
                  
                     '寄發信件
                     strRecordRow = strRecordRow & "[寄信]"
                     
                     'Add By Sindy 2019/7/17 增加副本給國外部信件處理人
                     If UCase(tmpArr(j)) = "IPDEPT" Then
                        If strToCC <> "" Then strToCC = strToCC & ";"
                        strToCC = Pub_GetSpecMan("國外部信件處理人")
                     ElseIf UCase(tmpArr(j)) = "TM" Then
                        If strToCC <> "" Then strToCC = strToCC & ";"
                        strToCC = OL_TmMailCC
                     End If
                     '2019/7/17 END
                     
                     'Add By Sindy 2022/2/7
                     If strTo <> "" Then strTo = strTo & ";"
                     strTo = strTo & Trim(tmpArr(j))
                     '2022/2/7 END
                     
'                     'Modify By Sindy 2018/10/17
'                     '還是先維持用個人名義轉寄信件
'                     PUB_SendMail strUserNum, tmpArr(j), "", GRD1(intTab).TextMatrix(i, 1), vbCrLf & "信件內容參附件！", , strFileName, , , , strToCC, , , , , False
'                     'PUB_SendMail strUserNum, tmpArr(j), "", GRD1(intTab).TextMatrix(i, 1), vbCrLf & "信件內容參附件！", , strFileName, , , , , "patent@taie.com.tw", , , , False
'                     '2018/10/17 END
'                     If bolMailSendOk = False Then GoTo ErrHand
'                     '該收受者上刪除日期時間人員
'                     strExc(0) = "update InputRecord set " & _
'                                 " ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
'                                 " where ir01=" & GRD1(intTab).TextMatrix(i, 8) & _
'                                   " and ir02=" & GRD1(intTab).TextMatrix(i, 9) & _
'                                   " and ir03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)) & "'" & _
'                                   " and ir04='" & tmpArr(j) & "'"
'                     cnnConnection.Execute strExc(0)
                  End If
                  '2017/12/22 END
               End If
            Next j
            'Add By Sindy 2022/2/7
            '秀玲寄-
            '洪副理 您好：原信是同時寄到IPDEPT及PATENT信箱，PATENT的部分是由林慧汶轉寄您及Monica、May三人，
            '程式原是分3 封信寄發，已請SINDY調整程式，改為一封信同時寄發三人，
            '這樣您就不會再有重覆收到正副本不同的2封信，也可以看出同時發給May。
            If strTo <> "" Then
               '還是先維持用個人名義轉寄信件
               PUB_SendMail strUserNum, strTo, "", GRD1(intTab).TextMatrix(i, 1), vbCrLf & "信件內容參附件！", , strFileName, , , , strToCC, , , , , False
               If bolMailSendOk = False Then GoTo ErrHand
               '該收受者上刪除日期時間人員
               strExc(0) = "update InputRecord set " & _
                           " ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
                           " where ir01=" & GRD1(intTab).TextMatrix(i, 8) & _
                             " and ir02=" & GRD1(intTab).TextMatrix(i, 9) & _
                             " and ir03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)) & "'" & _
                             " and ir04 in('" & Replace(strTo, ";", "','") & "')"
               cnnConnection.Execute strExc(0)
            End If
            '2022/2/7 END
            
            strExc(0) = "update PatentInput set " & _
                        " PI08=" & strSrvDate(1) & ",PI09=" & strUpdTime & ",PI10='" & strUserNum & "'" & _
                        " where PI01=" & GRD1(intTab).TextMatrix(i, 8) & _
                          " and PI02=" & GRD1(intTab).TextMatrix(i, 9) & _
                          " and PI03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)) & "'"
            cnnConnection.Execute strExc(0)
            
            '檢查信件收受者全部已處理或已刪除,主檔就可以掛上msg檔刪除日期,等待AutoBatchDay一個月後刪除實體檔
            Call SavePatentInput(GRD1(intTab).TextMatrix(i, 8), GRD1(intTab).TextMatrix(i, 9), GRD1(intTab).TextMatrix(i, 13))
            
            cnnConnection.CommitTrans
            
'            '刪除PC端檔案
'            'Call fs.DeleteFile(m_AttachPath & "\" & strFileName)
'            Kill strFileName
            
            '串要發通知信的人員
            tmpArr = Split(GRD1(intTab).TextMatrix(i, 7), ";")
            For j = 0 To UBound(tmpArr)
               If tmpArr(j) <> "" Then
                  'Modify By Sindy 2017/12/22
                  If PUB_GetST03(CStr(tmpArr(j))) = "P12" Then '專利程序
                  '2017/12/22 END
                     If InStr(strEMailTo, tmpArr(j)) = 0 Then
                        strEMailTo = strEMailTo & ";" & tmpArr(j)
                     End If
                  End If
               End If
            Next j
            
         End If
ReadNext:
      Next i
      End If
   Next intTab
   
   Call PUB_SendNotifyMail(strEMailTo, True) '寄發通知信
   Call PUB_SendMailCache(, , False) 'Add By Sindy 2019/7/17
   DoEvents
   cmdSendMail.Enabled = True
   
   Call QueryData
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   
   Call PUB_SendNotifyMail(strEMailTo, True) '寄發通知信
   
   cmdSendMail.Enabled = True
   MsgBox " 整批轉寄失敗！" & vbCrLf & "●筆數:" & strRecordRow & vbCrLf & Err.Description
   Call QueryData '重新查詢
End Sub

'Add By Sindy 2018/1/2
Private Sub SavePatentInput(strPI01 As String, strPI02 As String, strPI03 As String)
   '若信件收受者全部已處理或已刪除,主檔就可以掛上msg檔刪除日期,等待AutoBatchDay一個月後刪除實體檔
   strExc(0) = "select ir01 from InputRecord" & _
               " where ir01=" & strPI01 & _
                 " and ir02=" & strPI02 & _
                 " and ir03='" & strPI03 & "'" & _
                 " and ir08=0" 'and ir05=0 and ir08=0 : 若信件收受者全部已讀取或已刪除
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      '更新"無"Msg檔刪除日期
      strExc(0) = "update PatentInput set" & _
                  " pi16=" & strSrvDate(1) & _
                  " where pi01=" & strPI01 & _
                    " and pi02=" & strPI02 & _
                    " and pi03='" & strPI03 & "'" & _
                    " and pi16=0"
      cnnConnection.Execute strExc(0)
   End If
End Sub

Private Function GetAttachFile(ByVal strPkey1 As String, ByVal strPkey2 As String, _
                               ByVal strPkey3 As String, ByRef pFileName As String, _
                               Optional pSavePath As String) As Boolean
Dim stAttPath As String
   
On Error GoTo ErrHnd

   If pSavePath = "" Then
      If Dir(m_AttachPath, vbDirectory) = "" Then
         MkDir m_AttachPath
      End If
      stAttPath = m_AttachPath & "\" & pFileName
   Else
      '改傳完整的檔案路徑:路徑+檔名
      If InStr(pSavePath, m_AttachPath) > 0 Then
         If Dir(m_AttachPath, vbDirectory) = "" Then
            MkDir m_AttachPath
         End If
      End If
      stAttPath = pSavePath
   End If
   
   GetAttachFile = PUB_GetAttachFile_IImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
   
   Exit Function
   
ErrHnd:
   If Err.NUMBER = 70 Then
      MsgBox ChgSQL(pFileName) & "檔案已開啟！", vbCritical
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function

'重新分類
Private Sub CmdSave_Click()
Dim intTab As Long
Dim i As Integer
   
On Error GoTo ErrHand
   
   Screen.MousePointer = vbHourglass
   For intTab = intStart To quyIndex
      For i = 1 To GRD1(intTab).Rows - 1
         If GRD1(intTab).TextMatrix(i, 13) <> "" And GRD1(intTab).RowHeight(i) > 0 Then '有資料
            If GRD1(intTab).TextMatrix(i, 0) = "!" Then
               'Modify By Sindy 2017/12/21
'               strExc(0) = "update PatentInput set" & _
'                           " PI05='" & grd1(intTab).TextMatrix(i, 11) & "',PI06='" & grd1(intTab).TextMatrix(i, 10) & "'" & _
'                           ",PI18='" & grd1(intTab).TextMatrix(i, 19) & "',PI19='" & grd1(intTab).TextMatrix(i, 20) & "'" & _
'                           ",PI20='" & grd1(intTab).TextMatrix(i, 21) & "',PI21='" & grd1(intTab).TextMatrix(i, 22) & "'" & _
'                           " where PI01=" & grd1(intTab).TextMatrix(i, 8) & _
'                             " and PI02=" & grd1(intTab).TextMatrix(i, 9) & _
'                             " and PI03='" & ChgSQL(grd1(intTab).TextMatrix(i, 13)) & "'"
'               cnnConnection.Execute strExc(0)
               'Modify By Sindy 2022/7/14 + 儲存PI10:分類人員
               strExc(0) = "update PatentInput set" & _
                           " PI05='" & GRD1(intTab).TextMatrix(i, 11) & "',PI06='" & GRD1(intTab).TextMatrix(i, 10) & "',PI10='" & strUserNum & "'"
               If GRD1(intTab).TextMatrix(i, 19) <> "" Then
                  strExc(0) = strExc(0) & _
                           ",PI18='" & GRD1(intTab).TextMatrix(i, 19) & "',PI19='" & GRD1(intTab).TextMatrix(i, 20) & "'" & _
                           ",PI20='" & GRD1(intTab).TextMatrix(i, 21) & "',PI21='" & GRD1(intTab).TextMatrix(i, 22) & "'"
               End If
               'Add By Sindy 2024/4/19
               '備註
               strExc(0) = strExc(0) & ",PI23='" & GRD1(intTab).TextMatrix(i, 24) & "'"
               '2024/4/19 END
               strExc(0) = strExc(0) & _
                           " where PI01=" & GRD1(intTab).TextMatrix(i, 8) & _
                             " and PI02=" & GRD1(intTab).TextMatrix(i, 9) & _
                             " and PI03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)) & "'"
               cnnConnection.Execute strExc(0)
               '2017/12/21 END
            End If
         End If
      Next i
   Next intTab
   Call QueryData
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Screen.MousePointer = vbDefault
   MsgBox " 重新分類失敗！" & vbCrLf & Err.Description
End Sub

'信件轉入及分類
Private Sub cmdTrans_Click()
Dim oFileSys As New FileSystemObject
Dim oFolder As Folder
Dim strErrText As String
   
   '先更新資料
   If cmdSave.Enabled = True Then
      Call CmdSave_Click
   Else
      Call QueryData
   End If
   
   If txtPathPatent = "" Then
      MsgBox "信件資料夾不可空白！"
      Exit Sub
   End If
   If Dir(txtPathPatent, vbDirectory) = "" Then
      MkDir txtPathPatent
   End If
   Set oFolder = oFileSys.GetFolder(txtPathPatent.Text)
   If oFolder.files.Count = 0 Then
      MsgBox "此目錄尚無信件！"
      Set oFolder = Nothing
      Exit Sub
   End If
   
On Error GoTo ErrHand
   
   cmdTrans.Enabled = False
   TxtIPDept.Visible = True
   LblCntIPDept.Visible = True
   If PUB_PatentTransMail(Me, , strErrText, , "N") = False Then
      GoTo ErrHand
   End If
   PUB_SaveLastDate Me.Name, strUserNum & "PATH", txtPathPatent.Text
   MsgBox "信件轉入完成！" & IIf(oFolder.files.Count > 0, vbCrLf & vbCrLf & "(尚有未轉入的信件，詳情請至資料夾查看)", "")
   
   GetTodayTotCnt '重新計算今日總筆數
   
   cmdTrans.Enabled = True
   TxtIPDept.Visible = False
   LblCntIPDept.Visible = False
   Call QueryData
   Exit Sub
   
ErrHand:
   MsgBox strErrText, vbExclamation
   cmdTrans.Enabled = True
   Call QueryData
End Sub

Private Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim intTab As Integer
Dim dblTotCnt As Double
Dim ii As Integer
Dim strPI05 As String
Dim strConSql As String 'Add By Sindy 2022/7/15
Dim tmpArr As Variant
   
   cmdSendMail.Enabled = True 'Add By Sindy 2018/1/3
   cmdUpdRow.Enabled = False
   cmdSave.Enabled = False: cmdSave.BackColor = &H8000000F
   m_blnColOrderAsc = True
   QueryData = False
   
   Screen.MousePointer = vbHourglass
   'Modify By Sindy 2020/6/16 從Form Load Move過來
   'Add By Sindy 2018/1/5
   strSql = "SELECT count(*) FROM patentinput WHERE pi05 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If RsTemp.Fields(0) > 0 Then
         txtShowTransMsg.Top = 1830
         txtShowTransMsg.ZOrder '移至頂層
         txtShowTransMsg.Visible = True
         DoEvents '*****
         Call PUB_IPDeptChangePatent(Me)
         txtShowTransMsg.Visible = False
         DoEvents '*****
      End If
   End If
   '2018/1/5 END
   '2020/6/16 END
   
   dblTotCnt = 0
   For intTab = intStart To quyIndex
      GRD1(intTab).Clear
      Call SetGrd(intTab)
      'Modify By Sindy 2020/3/19 (intTab + 1) => strPI05
      'Modify By Sindy 2025/1/9
      'strPI05 = intTab + 1
      If intTab >= 0 And intTab <= cboPI05.ListCount - 1 Then
         strPI05 = Trim(Left(cboPI05.List(intTab), 3))
      Else
         strPI05 = ""
      End If
      If strPI05 = "" Then
         SSTab1.TabVisible(intTab) = False
         LblRow(intTab).Visible = False
         LblRow(intTab).Caption = ""
         Call ChkTab_Click(intTab)
      Else
      '2025/1/9 END
         'Add By Sindy 2022/7/15 8.垃圾信箱
         '雅娟的垃圾信箱: 排除是雅娟自己分到垃圾信箱的
         '玫音的垃圾信箱: 只能看到是由雅娟分過來的
         strConSql = ""
         'If intTab = 7 Then '垃圾信箱
         LblRow(intTab).BackColor = &HC0C0FF
         If InStr(cboPI05.List(intTab), "垃圾信箱") > 0 Then
            ChkTab(intTab).Visible = False
            LblRow(intTab).BackColor = &H80FF80
            If strUserNum = "79075" Then '雅娟
               strConSql = " and (pi10<>'" & strUserNum & "' or pi10 is null)"
            ElseIf strUserNum = "99043" Then '玫音
               strConSql = " and pi10='79075'"
            End If
         End If
         '2022/7/15 END
         
         If intTab > Combo5.ListCount - 1 Then
            tmpArr = Split(Replace(cboPI05.List(intTab), "  ", " "), " ")
            SSTab1.TabCaption(intTab) = Trim(tmpArr(1))
         End If
         
         'Modify By Sindy 2022/2/9 + ,getmailbox(pi01,pi03)|| : 分類前面加信箱來源和收件者信箱
         'Modify By Sindy 2024/4/19 + ,PI23 as 備註
         strSql = "select '' V,PI17 主旨,getmailbox(pi01,pi03)||decode(PI05," & Show專利處信件分類 & ",PI05) 分類,decode(pi18,null,'',pi18||'-'||pi19||'-'||pi20||'-'||pi21) 本所案號" & _
                  ",decode(nvl(st02,''),'','',st02) 收受者,sqldatet(PI12)||' '||sqltime6(PI13) 收信日期時間,PI05,PI06,PI01,PI02,'' newPI06,'' newPI05,PI15 系統記錄,PI03,PI14 FTP路徑檔名" & _
                  ",PI18,PI19,PI20,PI21,'' newPI18,'' newPI19,'' newPI20,'' newPI21,PI11,PI23 as 備註" & _
                  " From PatentInput,staff" & _
                  " where PI08=0 and PI05='" & strPI05 & "'" & _
                  " and PI06=st01(+)" & strConSql & _
                  " order by PI12 desc,PI13 desc"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            Set GRD1(intTab).Recordset = rsTmp
            'Modify By Sindy 2025/1/9
            LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
            dblTotCnt = dblTotCnt + rsTmp.RecordCount
   '         If intTab = 0 Then SSTab1.TabCaption(intTab) = "P程序1": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   '         If intTab = 1 Then SSTab1.TabCaption(intTab) = "P程序2": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   '         'Modify By Sindy 2018/6/21
   ''         If intTab = 2 Then SSTab1.TabCaption(intTab) = "亞洲": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   ''         If intTab = 3 Then SSTab1.TabCaption(intTab) = "歐洲": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   ''         If intTab = 4 Then SSTab1.TabCaption(intTab) = "美洋非(單)": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   ''         If intTab = 5 Then SSTab1.TabCaption(intTab) = "美洋非(雙)": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   ''         If intTab = 2 Then SSTab1.TabCaption(intTab) = "美日(單)": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   ''         If intTab = 3 Then SSTab1.TabCaption(intTab) = "美日(雙)": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   ''         If intTab = 4 Then SSTab1.TabCaption(intTab) = "美日外(單)": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   ''         If intTab = 5 Then SSTab1.TabCaption(intTab) = "美日外(雙)": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   '         '2018/6/21 END
   '         'Modify By Sindy 2020/3/19
   '         If intTab = 2 Then LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   '         If intTab = 3 Then LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   '         If intTab = 4 Then LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   '         If intTab = 5 Then LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   '         '2020/3/19 END
   '         If intTab = 6 Then SSTab1.TabCaption(intTab) = "其他": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   '         If intTab = 7 Then SSTab1.TabCaption(intTab) = "垃圾信箱": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   '         'If intTab = 8 Then SSTab1.TabCaption(intTab) = "國外部匯入": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
   '         If intTab <> 8 Then dblTotCnt = dblTotCnt + rsTmp.RecordCount
            '2025/1/9 END
            QueryData = True
            '解析收受者
            For ii = 1 To GRD1(intTab).Rows - 1
               GRD1(intTab).TextMatrix(ii, 4) = PUB_ReadUserData(GRD1(intTab).TextMatrix(ii, 7))
            Next ii
         Else
            'Modify By Sindy 2025/1/9
            LblRow(intTab).Caption = "": Call ChkTab_Click(intTab)
   '         If intTab = 0 Then SSTab1.TabCaption(intTab) = "P程序1": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   '         If intTab = 1 Then SSTab1.TabCaption(intTab) = "P程序2": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   '         'Modify By Sindy 2018/6/21
   ''         If intTab = 2 Then SSTab1.TabCaption(intTab) = "亞洲": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   ''         If intTab = 3 Then SSTab1.TabCaption(intTab) = "歐洲": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   ''         If intTab = 4 Then SSTab1.TabCaption(intTab) = "美洋非(單)": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   ''         If intTab = 5 Then SSTab1.TabCaption(intTab) = "美洋非(雙)": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   ''         If intTab = 2 Then SSTab1.TabCaption(intTab) = "美日(單)": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   ''         If intTab = 3 Then SSTab1.TabCaption(intTab) = "美日(雙)": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   ''         If intTab = 4 Then SSTab1.TabCaption(intTab) = "美日外(單)": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   ''         If intTab = 5 Then SSTab1.TabCaption(intTab) = "美日外(雙)": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   '         '2018/6/21 END
   '         'Modify By Sindy 2020/3/19
   '         If intTab = 2 Then LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   '         If intTab = 3 Then LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   '         If intTab = 4 Then LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   '         If intTab = 5 Then LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   '         '2020/3/19 END
   '         If intTab = 6 Then SSTab1.TabCaption(intTab) = "其他": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   '         If intTab = 7 Then SSTab1.TabCaption(intTab) = "垃圾信箱": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
   '         'If intTab = 8 Then SSTab1.TabCaption(intTab) = "國外部匯入": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab8).Visible = False
            '2025/1/9 END
         End If
         rsTmp.Close
      
         GRD1(intTab).Visible = False
         GRD1(intTab).col = 0
         GRD1(intTab).row = 1
         GRD1(intTab).Visible = True
      End If
   Next intTab
   'If intStart <> quyIndex Then
      LblTotCnt.Caption = "總筆數(" & intStart + 1 & "~" & quyIndex + 1 & "頁籤): " & dblTotCnt
   'End If
   
   'Add By Sindy 2017/11/15
   strExc(0) = "select mrl01 from mailreceivelog" & _
               " where mrl01='" & Left(Patent收件匣, 2) & "'" & _
               " and mrl09='A'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      cmdHandRecv.BackColor = &HC0FFC0
   Else
      cmdHandRecv.BackColor = &H8000000F
   End If
   '2017/11/15 END
   
   ClearDetail
   'Modify By Sindy 2025/1/16
   'SSTab1.Tab = 6 '預設在其他
   For ii = 0 To SSTab1.Tabs - 1
      If SSTab1.TabCaption(ii) = "其他" Then
         SSTab1.Tab = ii
         Exit For
      End If
   Next ii
   '2025/1/16 END
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'清除單筆明細資料
Private Sub ClearDetail()
Dim i As Integer
   
   txtPI17.Text = ""
   txtPI11.Text = ""
   LblPI12.Caption = ""
   cboPI05.Enabled = False
   cboPI05.ListIndex = -1
   'cboPI05.Text = ""
   cboPI06.Enabled = False
   cboPI06.ListIndex = -1
   cboPI06.Text = ""
   txtPI18.Text = ""
   txtPI19.Text = ""
   txtPI20.Text = ""
   txtPI21.Text = ""
   List1.Clear
   List1.Tag = ""
   Frame1.Tag = "" '記錄目前那個tab
   Me.Tag = "" '記錄grd點選那幾筆資料列
   For i = intStart To quyIndex
      dblPrevRow(i) = 0
   Next i
   txtPI23.Text = "" 'Add By Sindy 2024/4/19
End Sub

'更正單筆視窗資料
Private Sub cmdUpdRow_Click()
Dim tmpArr As Variant
Dim strUser As String
Dim strText As String
Dim intUpdRow As Integer
Dim jj As Integer
Dim bolCancl As Boolean
Dim SelManyRow As Boolean
Dim i As Integer
Dim strTempEmp As String 'Add By Sindy 2019/7/5
   
   'Add By Sindy 2022/3/21
   If InStr(cboPI05.Text, "人員空缺") > 0 Then
      MsgBox "人員空缺，請選擇有效的分類！", vbExclamation
      Exit Sub
   End If
   '2022/3/21 END
   
   'Add By Sindy 2017/12/21
   SelManyRow = False '是否選取多筆
   jj = 0
   If Val(Frame1.Tag) >= 0 And Me.Tag <> "" Then
      For i = 1 To GRD1(SSTab1.Tab).Rows - 1
         If GRD1(SSTab1.Tab).TextMatrix(i, 0) = "V" Then
            jj = jj + 1
            If jj > 1 Then
               SelManyRow = True '選取多筆
            End If
         End If
      Next i
   End If
   '2017/12/21 END
   
   'Frame1.Tag : 第幾個頁籤
   'Me.Tag : GRD筆數
   If Val(Frame1.Tag) >= 0 And Me.Tag <> "" Then
      If GRD1(Frame1.Tag).TextMatrix(dblPrevRow(Frame1.Tag), 13) <> "" Then '有資料
         If cboPI05.Text = "" Then
            MsgBox "分類不可空白！", vbExclamation
            cboPI05.SetFocus
            Exit Sub
         End If
         If Trim(Left(cboPI05, 2)) <> 7 And Trim(Left(cboPI05, 2)) <> 8 Then '其他和垃圾信箱不須控制即時輸入收受者
            If cmdUpdRow.Enabled = True And List1.Tag = "" Then
               MsgBox "收受者不可空白！", vbExclamation
               cboPI06.SetFocus
               Exit Sub
            End If
'         Else
'            If List1.Tag = "" And (txtPI18 = "" Or txtPI19 = "") Then
'               MsgBox "收受者 或 本所案號不可空白！", vbExclamation
'               txtPI18.SetFocus
'               Exit Sub
'            End If
         End If
         
'         bolCancl = False
'         Call txtPI21_Validate(bolCancl)
'         If bolCancl = True Then
'            Exit Sub
'         End If
         
         If SSTab1.Tab <> Frame1.Tag Then
            MsgBox "記錄欲更新的頁籤有誤！SSTab1.Tab=" & SSTab1.Tab & " Frame1.Tag=" & Frame1.Tag, vbExclamation
            Exit Sub
         End If
         
         Screen.MousePointer = vbHourglass
         For i = 1 To GRD1(SSTab1.Tab).Rows - 1
            If GRD1(SSTab1.Tab).TextMatrix(i, 0) = "V" Then
               intUpdRow = i
               If dblPrevRow(Frame1.Tag) <> intUpdRow Then
                  dblPrevRow(Frame1.Tag) = intUpdRow
               End If
               
               'Add By Sindy 2022/2/22 針對分類有*號時，要增加提醒訊息。
               If InStr(GRD1(SSTab1.Tab).TextMatrix(i, 2), "*") > 0 And Trim(Left(cboPI05.Text, 2)) = "8" Then '*號進垃圾信箱
                  If MsgBox(GRD1(SSTab1.Tab).TextMatrix(i, 1) & vbCrLf & vbCrLf & _
                         "提醒：信件狀態有一邊是直接刪除的，會在信箱代號旁加上*號，例如(P,F*)，若人員處理信件時發現非該單位信件，請轉寄回該信箱。" & vbCrLf & vbCrLf & _
                         "確定要丟垃圾信箱嗎？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
                     Screen.MousePointer = vbDefault
                     Exit Sub
                  End If
               End If
               '2022/2/22 END
               
               '分類
               'If Left(cboPI05.Text, 1) <> GRD1(Frame1.Tag).TextMatrix(intUpdRow, 6) Then
                  'Modify By Sindy 2022/2/9 分類前面加信箱來源和收件者信箱
                  If InStrRev(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2), ")") > 0 Then
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2) = Mid(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2), 1, InStrRev(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2), ")")) & Trim(Mid(cboPI05.Text, 3))
                  Else
                  '2022/2/9 END
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2) = Trim(Mid(cboPI05.Text, 3))
                  End If
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 11) = Trim(Left(cboPI05.Text, 2))
               'End If
               
               '收受者
               List1.Tag = ""
               For jj = 0 To List1.ListCount - 1
                  'Add By Sindy 2019/7/5 檢查收受者是否已存在記錄檔裡
                  If InStr(List1.List(jj), "@") > 0 Then
                     strTempEmp = Trim(Left(List1.List(jj), InStr(List1.List(jj), "@") - 1))
                  Else
                     strTempEmp = Trim(Left(List1.List(jj), 6))
                     'Add By Sindy 2021/3/8 Account,Jerry_lin
                     strSql = "SELECT st01,st02 FROM staff " & _
                              " WHERE st01='" & strTempEmp & "'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 0 Then
                        strTempEmp = Trim(List1.List(jj))
                     End If
                     '2021/3/8 END
                  End If
                  strSql = "SELECT ir03 FROM inputrecord " & _
                           " WHERE ir01=" & GRD1(Frame1.Tag).TextMatrix(intUpdRow, 8) & _
                           " and ir02=" & GRD1(Frame1.Tag).TextMatrix(intUpdRow, 9) & _
                           " and ir03=" & CNULL(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 13)) & _
                           " and ir04=" & CNULL(strTempEmp)
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     MsgBox strTempEmp & " 已存在收受者記錄檔裡，不可重覆！", vbExclamation
                     Screen.MousePointer = vbDefault
                     Exit Sub
                  End If
                  '2019/7/5 END
                  
                  'Add By Sindy 2018/10/17
                  If InStr(List1.List(jj), "@") > 0 Then
                     List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & Trim(List1.List(jj))
                  Else
                  '2018/10/17 END
                     List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & strTempEmp 'Trim(Left(List1.List(jj), 6))
                  End If
               Next jj
               If List1.Tag = "" Then
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 4) = ""
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 10) = ""
               Else
                  If List1.Tag <> GRD1(Frame1.Tag).TextMatrix(intUpdRow, 7) Then
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 4) = PUB_ReadUserData(List1.Tag)
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 10) = List1.Tag
                     Call SetList1(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 10))
                  End If
               End If
               
               'Add By Sindy 2024/4/19
               '備註
               GRD1(Frame1.Tag).TextMatrix(intUpdRow, 24) = txtPI23.Text
               '2024/4/19 END
               
               If SelManyRow = False Then '單筆時
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 19) = txtPI18
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 20) = txtPI19
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 21) = txtPI20
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 22) = txtPI21
                  If txtPI18 = "" And txtPI19 = "" And txtPI20 = "" And txtPI21 = "" Then
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 3) = ""
                  Else
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 3) = txtPI18 & "-" & txtPI19 & "-" & txtPI20 & "-" & txtPI21
                  End If
               End If
               Call CancelRowColor(Frame1.Tag, intUpdRow) '清除反白,並且檢查是否有更新過資料
            End If
         Next i
         cmdUpdRow.Enabled = False
         Screen.MousePointer = vbDefault
      End If
   End If
End Sub

Private Sub SetList1(strText As String)
Dim tmpArr As Variant
Dim strTempName As String
Dim j As Integer
   
   '收受者
   tmpArr = Split(strText, ";")
   List1.Clear
   List1.Tag = ""
   For j = 0 To UBound(tmpArr)
      If tmpArr(j) <> "" Then
         strTempName = ""
         If Len(tmpArr(j)) = 5 And InStr(tmpArr(j), "@") = 0 Then '員工編號
            strTempName = GetPrjSalesNM(CStr(tmpArr(j)))
         End If
         If strTempName <> "" Then
            List1.AddItem tmpArr(j) & " " & strTempName
         Else
            List1.AddItem tmpArr(j)
         End If
         bolCboPI06_KeyPress = False 'Add By Sindy 2021/4/14
      End If
   Next j
End Sub

Private Sub Command2_Click()
Dim stFileName As String
   
On Error GoTo ErrHnd
   
   stFileName = "*.msg"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "msg檔案 (*.msg)|*.msg"
      .InitDir = IIf(txtPathPatent <> "", txtPathPatent, PUB_Getdesktop)
      .MaxFileSize = 5000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         txtPathPatent.Text = Mid(.FileName, 1, InStrRev(.FileName, "\") - 1)
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.NUMBER <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
Dim ii As Integer, jj As Integer
Dim varTemp As Variant
Dim strPtab As String, strCFPtab As String
Dim tmpArr As Variant
   
   MoveFormToCenter Me
   
   ReDim pa(1 To TF_PA) As String
   ReDim sp(1 To tf_SP) As String
   
   If PUB_GetLastDate(Me.Name, strUserNum & "PATH") <> "" Then
      txtPathPatent = PUB_GetLastDate(Me.Name, strUserNum & "PATH")
   End If
   
   If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
      cboPI06x.Visible = False
   End If
   
   'SSTab1的Tab數量:index起迄
   intStart = 0
   quyIndex = 9
   '組合下拉選單
   '分類
   cboPI05.Clear
   Combo5.Clear 'Add By Sindy 2025/1/10
   'Add By Sindy 2025/1/9
   If strSrvDate(1) >= P業務區劃分啟用日 Then
      Call PUB_AddItemCFPHandler(cboPI05, Combo5, , "P")
      For ii = 0 To Combo5.ListCount - 1
         If ii > 3 Then
            Exit For '只有4個頁籤
         End If
         strPtab = strPtab & "," & ii
         varTemp = Split(Combo5.List(ii), " ")
         If UBound(varTemp) > 0 Then
            SSTab1.TabCaption(ii) = varTemp(0) & " "
            For jj = 0 To cboPI05.ListCount - 1
               If InStr(cboPI05.List(jj), varTemp(0)) > 0 Then
                  varTemp = Split(cboPI05.List(jj), " ")
                  SSTab1.TabCaption(ii) = SSTab1.TabCaption(ii) & varTemp(1)
                  Exit For
               End If
            Next jj
         End If
      Next ii
      '固定4個頁籤,其他補"人員空缺"
      For ii = ii To 3
         SSTab1.TabCaption(ii) = "人員空缺"
         cboPI05.AddItem "   (人員空缺)"
      Next ii
      int_Pcnt = 4 'Add By Sindy 2025/1/9
   Else
   '2025/1/9 END
      cboPI05.AddItem "1  P程序1"
      cboPI05.AddItem "2  P程序2"
      'Add By Sindy 2025/1/9
      Combo5.AddItem "1  P程序1"
      Combo5.AddItem "2  P程序2"
      SSTab1.TabCaption(0) = "P程序1"
      SSTab1.TabCaption(1) = "P程序2"
      strPtab = strPtab & ",0,1"
      int_Pcnt = 2
      '2025/1/9 END
   End If
   'Modify By Sindy 2018/6/21
'   cboPI05.AddItem "3  亞洲"
'   cboPI05.AddItem "4  歐洲"
'   cboPI05.AddItem "5  美洋非(單)"
'   cboPI05.AddItem "6  美洋非(雙)"
   'Modify by Sindy 2020/3/18
   '109/4/1以後改業務區劃分
   If strSrvDate(1) >= CFP業務區劃分啟用日 Then
      Call PUB_AddItemCFPHandler(cboPI05, Combo5)
      For ii = int_Pcnt To Combo5.ListCount - 1
         If ii > 3 + int_Pcnt Then
            Exit For '只有4個頁籤
         End If
         strCFPtab = strCFPtab & "," & ii
         varTemp = Split(Combo5.List(ii), " ")
         If UBound(varTemp) > 0 Then
'            'Add By Sindy 2025/1/9
'            If strSrvDate(1) >= P業務區劃分啟用日 Then
               SSTab1.TabCaption(ii) = varTemp(0) & " "
               For jj = 0 To cboPI05.ListCount - 1
                  If InStr(cboPI05.List(jj), varTemp(0)) > 0 Then
                     varTemp = Split(cboPI05.List(jj), " ")
                     SSTab1.TabCaption(ii) = SSTab1.TabCaption(ii) & varTemp(1)
                     Exit For
                  End If
               Next jj
'            Else
'            '2025/1/9 END
'               SSTab1.TabCaption(ii + 2) = varTemp(0) & " "
'               For jj = 0 To cboPI05.ListCount - 1
'                  If InStr(cboPI05.List(jj), varTemp(0)) > 0 Then
'                     varTemp = Split(cboPI05.List(jj), " ")
'                     SSTab1.TabCaption(ii + 2) = SSTab1.TabCaption(ii + 2) & varTemp(1)
'                     Exit For
'                  End If
'               Next jj
'            End If
         End If
      Next ii
      'Add By Sindy 2022/3/21 CFP固定4個頁籤,其他補"人員空缺"
      For ii = ii To 3 + int_Pcnt
         SSTab1.TabCaption(2 + ii) = "人員空缺"
         cboPI05.AddItem "   (人員空缺)" ', intI 'Added by Morgan 2022/3/18 CFP減1人,先補空白,否則後面的index會不正確
      Next ii
      '2022/3/21 END
   Else
   '2020/3/18 END
      cboPI05.AddItem "3  美日(單)": SSTab1.TabCaption(2) = "美日(單)"
      cboPI05.AddItem "4  美日(雙)": SSTab1.TabCaption(3) = "美日(雙)"
      cboPI05.AddItem "5  美日外(單)": SSTab1.TabCaption(4) = "美日外(單)"
      cboPI05.AddItem "6  美日外(雙)": SSTab1.TabCaption(5) = "美日外(雙)"
      '2018/6/21 END
   End If
   
'   'Added by Morgan 2022/3/18 CFP減1人,先補空白,否則後面的index會不正確
'   If cboPI05.ListCount < 6 Then
'      For intI = cboPI05.ListCount To 5
'         cboPI05.AddItem "", intI
'      Next
'   End If
'   'end 2022/3/18
   
   cboPI05.AddItem "7  其他"
   cboPI05.AddItem "8  垃圾信箱"
   '收受者
   cboPI06.Clear
   cboPI06.AddItem ""
   'Add By Sindy 2018/4/27
   '固定加 David
   cboPI06.AddItem "77015 " & GetPrjSalesNM("77015")
   'Add By Sindy 2018/10/17
   '固定加 ipdept@taie.com.tw
   cboPI06.AddItem "ipdept@taie.com.tw"
   cboPI06.AddItem "TM@taie.com.tw" 'Add By Sindy 2019/2/19
   
   strSql = "SELECT a0902,st01,st02 FROM staff,acc090" & _
            " WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9'" & _
            " and st03>='P10' and st03<='P19' order by st03,st01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            cboPI06.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   'Add By Sindy 2017/2/14
   '鎖定資料:顯示提示訊息
   If Pub_StrUserSt03 <> "M51" Then 'Add By Sindy 2020/9/17
      If PUB_GetLock(Me.Name, m_OldKey, Me.Caption) = False Then
   '      cmdTrans.Enabled = False
   '      cmdSave.Enabled = False
   '      cmdSendMail.Enabled = False
   '      cmdHistory.Enabled = False
   '      cmdDelRow.Enabled = False
   '      cmdUpdRow.Enabled = False
   '      SSTab1.Enabled = False
   '   Else
   '      Me.Enabled = True
      End If
      '2017/2/14 end
   End If
   
   'Add By Sindy 2017/11/23
   If Dir(App.path & "\executePatent.txt") <> "" Then
      WebBrowserP.Navigate App.path & "\executePatent.txt"
      DoEvents
      TextBoxP = Replace(Replace(WebBrowserP.Document.Body.innerhtml, "<PRE>", ""), "</PRE>", "")
   Else
      TextBoxP = ""
   End If
   '2017/11/23 END
   
   If Pub_StrUserSt03 = "M51" Then
      cmdTrans.Visible = True
   End If
   
   'Modify By Sindy 2025/1/10
   If strPtab <> "" Then strPtab = Mid(strPtab, 2) & ",8"
   If strCFPtab <> "" Then strCFPtab = Mid(strCFPtab, 2) & ",8"
   'Add By Sindy 2018/5/9
   If Pub_strUserST05 = "73" Or Pub_strUserST05 = "75" Then 'P
      tmpArr = Split(strPtab, ",")
      For ii = 0 To UBound(tmpArr)
         ChkTab(tmpArr(ii)).Value = 1
      Next ii
'      ChkTab(0).Value = 1
'      ChkTab(1).Value = 1
'      ChkTab(6).Value = 1
   ElseIf Pub_strUserST05 = "83" Or Pub_strUserST05 = "85" Then 'CFP
      tmpArr = Split(strCFPtab, ",")
      For ii = 0 To UBound(tmpArr)
         ChkTab(tmpArr(ii)).Value = 1
      Next ii
'      ChkTab(2).Value = 1
'      ChkTab(3).Value = 1
'      ChkTab(4).Value = 1
'      ChkTab(5).Value = 1
'      ChkTab(6).Value = 1
   Else
      ChkTab(Combo5.ListCount).Value = 1
'      ChkTab(6).Value = 1
   End If
   '2018/5/9 END
   '2025/1/10 END
   
   QueryData
   GetTodayTotCnt '今日總筆數
   
   'Add By Sindy 2019/7/17
   'modify by sonia 2019/8/20 郭雅娟要求應薛經理改文字
   'LblCC.Caption = "其他信箱會加發副本給主管：Patent(" & PUB_ReadUserData(OL_PatMailCC) & ");TM(" & PUB_ReadUserData(OL_TmMailCC) & ");IPDept(" & PUB_ReadUserData(Pub_GetSpecMan("國外部信件處理人")) & ")"
   lblCC.Caption = "分信至其他部門信箱將加發副本：" & PUB_ReadUserData(OL_PatMailCC) & "(Patent);" & PUB_ReadUserData(OL_TmMailCC) & "(TM);" & PUB_ReadUserData(Pub_GetSpecMan("國外部信件處理人")) & "(IPDept)"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2017/2/14
   '清除鎖定資料
   strSql = "Delete from LockRec where LR01='" & Me.Name & "' and LR02='" & strUserNum & "'"
   adoTaie.Execute strSql
'   If PUB_GetLock("", m_OldKey) = False Then
'      Cancel = 1
'      Exit Sub
'   End If
   '2017/2/14 end
   DestroyToolTip '清除物件
   Set frm04010518 = Nothing
End Sub

Private Sub SetGrd(Index As Integer)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2024/4/19 +, "備註"
   '                        0    1       2       3           4         5               6       7       8       9       10         11         12          13      14             15      16      17      18      19         20         21         22         23      24
   arrGridHeadText = Array("V", "主旨", "分類", "本所案號", "收受者", "收信日期時間", "PI05", "PI06", "PI01", "PI02", "newPI06", "newPI05", "系統記錄", "PI03", "FTP路徑檔名", "PI18", "PI19", "PI20", "PI21", "newPI18", "newPI19", "newPI20", "newPI21", "PI11", "備註")
   arrGridHeadWidth = Array(200, 3500, 950, 1200, 1000, 1500, 0, 0, 0, 0, 0, 0, 900, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2000)
   'arrGridHeadWidth = Array(200, 3500, 950, 1200, 1000, 1500, 800, 800, 800, 800, 800, 800, 900, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800, 800)
   GRD1(Index).Visible = False
   GRD1(Index).Cols = UBound(arrGridHeadText) + 1
   GRD1(Index).Rows = 2
   For iRow = 0 To GRD1(Index).Cols - 1
      GRD1(Index).row = 0
      GRD1(Index).col = iRow
      GRD1(Index).Text = arrGridHeadText(iRow)
      GRD1(Index).ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1(Index).CellAlignment = flexAlignCenterCenter
   Next iRow
   GRD1(Index).Visible = True
End Sub

'開啟附件
Private Sub GRD1_DblClick(Index As Integer)
Dim strFileName As String
   
   GRD1(Index).row = GRD1(Index).MouseRow
   GRD1(Index).col = GRD1(Index).MouseCol
   nRow = GRD1(Index).row
   nCol = GRD1(Index).col
   If nCol = 1 And nRow > 0 Then
      dblPrevRow(Index) = nRow
      If GRD1(Index).TextMatrix(dblPrevRow(Index), 14) <> "" Then
         '讀取檔案
         Screen.MousePointer = vbHourglass
         strFileName = Mid(GRD1(Index).TextMatrix(dblPrevRow(Index), 14), InStrRev(GRD1(Index).TextMatrix(dblPrevRow(Index), 14), "\") + 1)
         strFileName = Mid(strFileName, InStrRev(strFileName, "/") + 1)
         Call PUB_ChkFileTypeOpenExE(strFileName) 'Add By Sindy 2017/9/13
         If GetAttachFile(GRD1(Index).TextMatrix(dblPrevRow(Index), 8), GRD1(Index).TextMatrix(dblPrevRow(Index), 9), GRD1(Index).TextMatrix(dblPrevRow(Index), 13), strFileName, m_AttachPath & "\" & strFileName) = True Then
            ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
'         Else
'            MsgBox "無此郵件！", vbInformation
         End If
         Screen.MousePointer = vbDefault
      End If
   End If
End Sub

'今日總筆數
Private Function GetTodayTotCnt()
   strSql = "SELECT count(*) FROM PatentInput WHERE PI01=" & strSrvDate(1)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      TodayTotCnt = "今日總筆數：" & "" & RsTemp.Fields(0)
   End If
End Function

Private Sub GRD1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   'grd1(Index).ToolTipText = ""
   If GRD1(Index).MouseRow <> 0 Then
      If GRD1(Index).MouseCol = 1 Or GRD1(Index).MouseCol = 4 Or GRD1(Index).MouseCol = 12 Then
         If iRow <> GRD1(Index).MouseRow Or iCol <> GRD1(Index).MouseCol Then
            'If GRD1(Index).TextMatrix(GRD1(Index).MouseRow, GRD1(Index).MouseCol) <> "" Then
               'grd1(Index).ToolTipText = grd1(Index).TextMatrix(grd1(Index).MouseRow, grd1(Index).MouseCol)
               CreateToolTip GetHWndForToolTip(GRD1(Index)), GRD1(Index).TextMatrix(GRD1(Index).MouseRow, GRD1(Index).MouseCol)
               iRow = GRD1(Index).MouseRow
               iCol = GRD1(Index).MouseCol
            'End If
         End If
      ElseIf GRD1(Index).MouseCol = 5 Then '收信日期時間
         CreateToolTip GetHWndForToolTip(GRD1(Index)), GRD1(Index).TextMatrix(GRD1(Index).MouseRow, 24)
         iRow = GRD1(Index).MouseRow
         iCol = GRD1(Index).MouseCol
      End If
   End If
End Sub

Private Sub Grd1_Click(Index As Integer)
Dim tmpArr As Variant, strTempName As String
Dim strKeep As String
Dim rsTmp As New ADODB.Recordset
Dim i As Integer, j As Integer
Dim ii As Integer

On Error GoTo ErrHand

GRD1(Index).Visible = False
GRD1(Index).row = GRD1(Index).MouseRow
GRD1(Index).col = GRD1(Index).MouseCol
nRow = GRD1(Index).row
nCol = GRD1(Index).col
If nRow = 0 Then
   If Me.GRD1(Index).row < 1 And Me.GRD1(Index).Text <> "V" Then
      If Me.GRD1(Index).Text = "無" Then
         If m_blnColOrderAsc = True Then
            Me.GRD1(Index).Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1(Index).Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1(Index).Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1(Index).Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
Else
   GRD1(Index).row = nRow 'GRD1(Index).MouseRow
   GRD1(Index).col = 0
   If GRD1(Index).TextMatrix(GRD1(Index).row, 13) <> "" Then
      '清除反白
      If GRD1(Index).CellBackColor = &HFFC0C0 Then
         'If nCol <> 1 Then
            Call CancelRowColor(Index, GRD1(Index).row) '清除反白,並且檢查是否有更新過資料
            If dblPrevRow(Index) = GRD1(Index).row Then
               dblPrevRow(Index) = 0
               '重新預設目前筆數
               If Me.Tag <> "" Then
                  tmpArr = Split(Me.Tag, ",")
                  dblPrevRow(Index) = tmpArr(UBound(tmpArr))
               Else
                  Call ClearDetail
               End If
            End If
         'End If
      Else
         '將點選資料列反白
         strKeep = GRD1(Index).TextMatrix(GRD1(Index).row, 0)
         GRD1(Index).TextMatrix(GRD1(Index).row, 0) = "V"
         GRD1(Index).col = 0
         GRD1(Index).row = nRow
         For i = 0 To GRD1(Index).Cols - 1
            GRD1(Index).col = i
            GRD1(Index).CellBackColor = &HFFC0C0
         Next i
         dblPrevRow(Index) = GRD1(Index).row '記錄目前筆數
         Me.Tag = Me.Tag & "," & dblPrevRow(Index)
      End If
      
      '**********************************************************************
      '顯示明細資料
      '**********************************************************************
      If Val(dblPrevRow(Index)) > 0 Then
         LblReceiver.Caption = PUB_GetMailInputData(GRD1(Index).TextMatrix(dblPrevRow(Index), 8), ChgSQL(GRD1(Index).TextMatrix(dblPrevRow(Index), 13)))
         If LblReceiver.Caption <> "" Then
            Label8.Visible = True
            LblReceiver.Visible = True
         Else
            Label8.Visible = False
            LblReceiver.Visible = False
         End If
         
         '主旨
         txtPI17.Text = GRD1(Index).TextMatrix(dblPrevRow(Index), 1)
         txtPI17.SetFocus 'Add by Sindy 2021/4/13 Form2.0才會顯示出捲抽
         '收信日期時間
         LblPI12.Caption = GRD1(Index).TextMatrix(dblPrevRow(Index), 5)
         '寄件者
         txtPI11.Text = GRD1(Index).TextMatrix(dblPrevRow(Index), 23)
         'Add By Sindy 2024/4/19
         '備註
         txtPI23.Text = GRD1(Index).TextMatrix(dblPrevRow(Index), 24)
         '2024/4/19 END
         
         '分類
         'Modify By Sindy 2025/1/10
         cboPI05.ListIndex = -1
         For ii = 0 To cboPI05.ListCount - 1
            tmpArr = Split(Replace(cboPI05.List(ii), "  ", " "), " ")
            If InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), tmpArr(1)) > 0 Or _
               InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), tmpArr(0)) > 0 Then
               cboPI05.ListIndex = ii
               Exit For
            End If
         Next ii
'         'Modify By Sindy 2024/10/7
'         'If GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "P程序1" Then
'         If InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "P程序1") > 0 Then
'            cboPI05.ListIndex = 0
'         'ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "P程序2" Then
'         ElseIf InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "P程序2") > 0 Then
'            cboPI05.ListIndex = 1
'         'Modify By Sindy 2018/6/21
''         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "亞洲" Then
''            cboPI05.ListIndex = 2
''         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "歐洲" Then
''            cboPI05.ListIndex = 3
''         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "美洋非(單)" Then
''            cboPI05.ListIndex = 4
''         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "美洋非(雙)" Then
''            cboPI05.ListIndex = 5
'         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "美日(單)" Then
'            cboPI05.ListIndex = 2
'         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "美日(雙)" Then
'            cboPI05.ListIndex = 3
'         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "美日外(單)" Then
'            cboPI05.ListIndex = 4
'         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "美日外(雙)" Then
'            cboPI05.ListIndex = 5
'         '2018/6/21 END
'         'ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "其他" Then
'         ElseIf InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "其他") > 0 Then
'            cboPI05.ListIndex = 6
'         'ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "垃圾信箱" Then
'         ElseIf InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "垃圾信箱") > 0 Then
'            cboPI05.ListIndex = 7
'         Else
'            cboPI05.ListIndex = -1
'            'Modify By Sindy 2020/3/19
'            For ii = 2 To 5
'               'If m_strPi05_CFP(ii) = GRD1(Index).TextMatrix(dblPrevRow(Index), 2) Then
'               If InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), m_strPi05_CFP(ii)) > 0 Then
'                  cboPI05.ListIndex = ii
'                  Exit For
'               End If
'            Next ii
'            '2020/3/19 END
'         End If
'         '2024/10/7 END
         '2025/1/10 END
         
         '本所案號
         txtPI18 = ""
         txtPI19 = ""
         txtPI20 = ""
         txtPI21 = ""
         If GRD1(Index).TextMatrix(dblPrevRow(Index), 19) <> "" Or strKeep = "!" Then
            txtPI18 = GRD1(Index).TextMatrix(dblPrevRow(Index), 19)
            txtPI19 = GRD1(Index).TextMatrix(dblPrevRow(Index), 20)
            txtPI20 = GRD1(Index).TextMatrix(dblPrevRow(Index), 21)
            txtPI21 = GRD1(Index).TextMatrix(dblPrevRow(Index), 22)
         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 15) <> "" Then
            txtPI18 = GRD1(Index).TextMatrix(dblPrevRow(Index), 15)
            txtPI19 = GRD1(Index).TextMatrix(dblPrevRow(Index), 16)
            txtPI20 = GRD1(Index).TextMatrix(dblPrevRow(Index), 17)
            txtPI21 = GRD1(Index).TextMatrix(dblPrevRow(Index), 18)
         End If
         
         'Modify By Sindy 2018/1/3 Move此處
         '收受者
         cboPI06.ListIndex = -1
         If GRD1(Index).TextMatrix(dblPrevRow(Index), 10) <> "" Or strKeep = "!" Then
            tmpArr = Split(GRD1(Index).TextMatrix(dblPrevRow(Index), 10), ";")
         Else
            tmpArr = Split(GRD1(Index).TextMatrix(dblPrevRow(Index), 7), ";")
         End If
         List1.Clear
         For j = 0 To UBound(tmpArr)
            If tmpArr(j) <> "" Then
               strTempName = ""
               If Len(tmpArr(j)) = 5 And InStr(tmpArr(j), "@") = 0 Then '員工編號
                  strTempName = GetPrjSalesNM(CStr(tmpArr(j)))
               End If
               If strTempName <> "" Then
                  List1.AddItem tmpArr(j) & " " & strTempName
                  cboPI06.Text = tmpArr(j) & " " & strTempName
               Else
                  List1.AddItem tmpArr(j)
                  cboPI06.Text = tmpArr(j)
               End If
               bolCboPI06_KeyPress = False 'Add By Sindy 2021/4/14
               If InStr(List1.Tag, cboPI06.Text) = 0 Then List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboPI06.Text
               cboPI06.Text = ""
            End If
         Next j
         
         '設定
         cmdUpdRow.Enabled = False
         'List1.Tag = ""
         Frame1.Tag = Index '記錄那一個GRD1
         cboPI05.Enabled = True
         cboPI06.Enabled = True
         List1.Enabled = True
         txtPI18.Enabled = True
         txtPI19.Enabled = True
         txtPI20.Enabled = True
         txtPI21.Enabled = True
'         If Index = 8 Then '國外部匯入
'            cboPI05.Enabled = False
'            cboPI06.Enabled = False
'            List1.Enabled = False
'            txtPI18.Enabled = False
'            txtPI19.Enabled = False
'            txtPI20.Enabled = False
'            txtPI21.Enabled = False
'         Else
            If UBound(Split(Me.Tag, ",")) > 0 Then
               cmdUpdRow.Enabled = True
            End If
'         End If
      End If
      '**********************************************************************
   End If
End If
GRD1(Index).Visible = True

Set rsTmp = Nothing
Exit Sub

ErrHand:
   Set rsTmp = Nothing
   MsgBox Err.Description
End Sub

'點二下可刪除List1資料列
Private Sub List1_DblClick(Cancel As MSForms.ReturnBoolean)
Dim strText As String

   If List1.ListIndex >= 0 Then
      strText = List1.List(List1.ListIndex)
      List1.RemoveItem List1.ListIndex
      If List1.Tag <> "" Then
         List1.Tag = Replace(List1.Tag, strText, "")
         List1.Tag = Replace(List1.Tag, ";;", ";")
         If Left(List1.Tag, 1) = ";" Then List1.Tag = Mid(List1.Tag, 2)
         If Right(List1.Tag, 1) = ";" Then List1.Tag = Mid(List1.Tag, 1, Len(List1.Tag) - 1)
      End If
      'MsgBox List1.Tag
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim tmpArr As Variant
Dim i As Integer
   
   If SSTab1.Enabled = False Then Exit Sub 'Add By Sindy 2017/2/14
   'Modify By Sindy 2022/1/6 人員不分出去的信件, 分類至垃圾信箱(不用輸入收受者); 統一由雅娟做2次檢查才能刪除;
   'cmdDelRow.Enabled = True
   'Modify By Sindy 2022/7/15
   '開放玫音也可以看到刪除鍵, 但垃圾信箱看到的信件是由雅娟分過去垃圾信箱的
   '雅娟的部分, 就排除是雅娟自己分過去的
   'Modify By Sindy 2025/1/10
   'If SSTab1.Tab = 7 And (strUserNum = "79075" Or strUserNum = "99043") Then
   If InStr(SSTab1.TabCaption(SSTab1.Tab), "垃圾信箱") > 0 And (strUserNum = "79075" Or strUserNum = "99043") Then
      cmdDelRow.Enabled = True
      cmdDelRow.Visible = True
   Else
      cmdDelRow.Enabled = False
      cmdDelRow.Visible = False
   End If
   '2022/1/6 END
   cmdUpdRow.Enabled = True
   If SSTab1.Tag <> "" And PreviousTab <> SSTab1.Tab Then
'      If PreviousTab = 8 Then
'         tmpArr = Split(Me.Tag, ",")
'         For i = 1 To UBound(tmpArr)
'            If Val(tmpArr(i)) > 0 Then
'               Call CancelRowColor(PreviousTab, Val(tmpArr(i))) '清除反白,並且檢查是否有更新過資料
'            End If
'         Next i
'         Call ClearDetail '清除單筆明細資料
'      Else
         If MsgBox("您已勾選資料尚未處理，" & vbCrLf & vbCrLf & _
                   "確定要放棄處理嗎？", vbInformation + vbYesNo + vbDefaultButton2, "警示詢問") = vbYes Then
            tmpArr = Split(Me.Tag, ",")
            For i = 1 To UBound(tmpArr)
               If Val(tmpArr(i)) > 0 Then
                  Call CancelRowColor(PreviousTab, Val(tmpArr(i))) '清除反白,並且檢查是否有更新過資料
               End If
            Next i
            Call ClearDetail '清除單筆明細資料
         Else
            SSTab1.Tag = ""
            SSTab1.Tab = PreviousTab
            Exit Sub '***
         End If
'      End If
   End If
   
'   If SSTab1.Tab = 8 Then
''      If SSTab1.Tag <> "" Then
''         TmpArr = Split(Me.Tag, ",")
''         For i = 1 To UBound(TmpArr)
''            If Val(TmpArr(i)) > 0 Then
''               Call CancelRowColor(PreviousTab, Val(TmpArr(i))) '清除反白,並且檢查是否有更新過資料
''            End If
''         Next i
''         Call ClearDetail '清除單筆明細資料
''      End If
'      cmdDelRow.Enabled = False
'      cmdUpdRow.Enabled = False
'      If grd1(SSTab1.Tab).Rows - 1 > 0 Then
'         If grd1(SSTab1.Tab).TextMatrix(1, 13) <> "" And grd1(SSTab1.Tab).RowHeight(1) > 0 Then '有資料
'            If MsgBox("要做IPDept資料分類動作嗎？", vbInformation + vbYesNo + vbDefaultButton2, "警示詢問") = vbYes Then
'               txtShowTransMsg.Top = 1830
'               txtShowTransMsg.ZOrder '移至頂層
'               txtShowTransMsg.Visible = True
'               DoEvents '*****
'               Screen.MousePointer = vbHourglass
'               Call TransIPDeptData
'               txtShowTransMsg.Visible = False
'               Screen.MousePointer = vbDefault
'            End If
'         End If
'      End If
'   End If
End Sub

Private Sub SSTab1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   SSTab1.Tag = Me.Tag
End Sub

Private Sub txtPI18_GotFocus()
   TextInverse txtPI18
   CloseIme
End Sub

Private Sub txtPI18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPI18_Validate(Cancel As Boolean)
   If txtPI18 <> "" Then
      txtPI18 = UCase(txtPI18)
      If ChkSysName(txtPI18) = True Then
         If txtPI18 <> "P" And txtPI18 <> "PS" And _
            txtPI18 <> "CFP" And txtPI18 <> "CPS" Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtPI18
End Sub

Private Sub txtPI19_GotFocus()
   TextInverse txtPI19
End Sub

Private Sub txtPI20_GotFocus()
   TextInverse txtPI20
End Sub

Private Sub txtPI20_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtPI20_LostFocus()
   If txtPI18 <> "" And txtPI19 <> "" And txtPI20 = "" Then txtPI20 = "0"
End Sub

Private Sub txtPI21_GotFocus()
   TextInverse txtPI21
End Sub

Private Sub txtPI21_LostFocus()
   If txtPI18 <> "" And txtPI19 <> "" And txtPI21 = "" Then txtPI21 = "00"
End Sub

Private Sub txtPI21_Validate(Cancel As Boolean)
Dim strPI05 As String
Dim strPI06 As String
Dim ii As Integer, varTemp As Variant 'Add By Sindy 2020/3/19
   
   If txtPI18 <> "" And txtPI19 <> "" Then
      If txtPI20 = "" Then txtPI20 = "0"
      If txtPI21 = "" Then txtPI21 = "00"
      If PUB_PatentByChkCP14(txtPI17, txtPI11, txtPI18, txtPI19, txtPI20, txtPI21, strPI05, strPI06, True) = True Then
         If strPI05 <> Trim(Left(cboPI05, 2)) Then
            'Modify By Sindy 2020/3/19
            'cboPI05.ListIndex = Val(strPI05) - 1
            For ii = 0 To cboPI05.ListCount - 1
               varTemp = Split(cboPI05.List(ii), " ")
               If UBound(varTemp) > 0 Then
                  If varTemp(0) = strPI05 Then
                     cboPI05.ListIndex = ii
                     Exit For
                  End If
               End If
            Next ii
            '2020/3/19 END
         End If
      Else
         cboPI05.ListIndex = 6 '其他
         cboPI06.Text = ""
         List1.Clear
         txtPI18 = "": txtPI19 = "": txtPI20 = "": txtPI21 = ""
         txtPI18.SetFocus
      End If
   End If
End Sub

Private Sub cboPI06x_GotFocus()
   cboPI06x.SelStart = 0
   cboPI06x.SelLength = Len(cboPI06x.Text)
End Sub
Private Sub cboPI06x_KeyPress(KeyAscPI As MSForms.ReturnInteger)
   bolCboPI06_KeyPress = True 'Add By Sindy 2021/4/14
   KeyAscPI = UpperCase(KeyAscPI)
End Sub

'Add By Sindy 2024/4/19
Private Sub txtPI23_GotFocus()
   OpenIme
   TextInverse txtPI23
End Sub
Private Sub txtPI23_Validate(Cancel As Boolean)
   If txtPI23.Text <> "" Then
      If Not CheckLengthIsOK(txtPI23, txtPI23.MaxLength, False) Then
         Cancel = True
         MsgBox "備註內容太長", vbOKOnly, "檢核資料"
         txtPI23_GotFocus
      End If
   End If
End Sub
'2024/4/19 END
