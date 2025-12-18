VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm06010611 
   BorderStyle     =   1  '單線固定
   Caption         =   "國外部收件夾信件處理"
   ClientHeight    =   7420
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   9000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7420
   ScaleWidth      =   9000
   Begin VB.CommandButton Command1 
      Caption         =   "轉寄"
      Height          =   345
      Left            =   2040
      TabIndex        =   52
      Top             =   0
      Visible         =   0   'False
      Width           =   795
   End
   Begin VB.CommandButton cmdHandRecv 
      Caption         =   "人工啟動接收"
      Height          =   330
      Left            =   7710
      Style           =   1  '圖片外觀
      TabIndex        =   45
      Top             =   1590
      Width           =   1245
   End
   Begin VB.CommandButton cmdRecOutlookQ 
      Caption         =   "郵件接收狀況"
      Height          =   330
      Left            =   6420
      TabIndex        =   44
      Top             =   1590
      Width           =   1245
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢(&Q)"
      Height          =   330
      Left            =   3000
      TabIndex        =   0
      Top             =   0
      Width           =   800
   End
   Begin VB.CommandButton cmdUpdRow 
      Caption         =   "更正"
      Height          =   285
      Left            =   5400
      TabIndex        =   40
      Top             =   1635
      Width           =   675
   End
   Begin VB.CommandButton cmdDelRow 
      Caption         =   "刪除"
      Height          =   285
      Left            =   4650
      TabIndex        =   39
      Top             =   1635
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
   Begin VB.TextBox TxtIPDept 
      Height          =   285
      Left            =   90
      TabIndex        =   27
      Top             =   570
      Visible         =   0   'False
      Width           =   8355
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   255
      Left            =   4350
      TabIndex        =   10
      Top             =   330
      Width           =   345
   End
   Begin VB.TextBox txtPathIPDept 
      Height          =   270
      Left            =   1230
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "C:\IPDept"
      Top             =   300
      Width           =   3105
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3915
      Left            =   60
      TabIndex        =   6
      Top             =   2160
      Width           =   8835
      _ExtentX        =   15593
      _ExtentY        =   6914
      _Version        =   393216
      Tabs            =   9
      Tab             =   5
      TabsPerRow      =   9
      TabHeight       =   520
      TabCaption(0)   =   "個案"
      TabPicture(0)   =   "frm06010611.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "GRD1(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "外商"
      TabPicture(1)   =   "frm06010611.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRD1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "外專"
      TabPicture(2)   =   "frm06010611.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GRD1(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "專利處"
      TabPicture(3)   =   "frm06010611.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GRD1(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "外法"
      TabPicture(4)   =   "frm06010611.frx":0070
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "GRD1(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "新知"
      TabPicture(5)   =   "frm06010611.frx":008C
      Tab(5).ControlEnabled=   -1  'True
      Tab(5).Control(0)=   "GRD1(5)"
      Tab(5).Control(0).Enabled=   0   'False
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "財務"
      TabPicture(6)   =   "frm06010611.frx":00A8
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "GRD1(6)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "其他"
      TabPicture(7)   =   "frm06010611.frx":00C4
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "GRD1(7)"
      Tab(7).ControlCount=   1
      TabCaption(8)   =   "開拓"
      TabPicture(8)   =   "frm06010611.frx":00E0
      Tab(8).ControlEnabled=   0   'False
      Tab(8).Control(0)=   "GRD1(8)"
      Tab(8).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm06010611.frx":00FC
         Height          =   3495
         Index           =   0
         Left            =   -74940
         TabIndex        =   7
         Top             =   360
         Width           =   8685
         _ExtentX        =   15311
         _ExtentY        =   6174
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|收受者|收信日期時間"
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
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm06010611.frx":0111
         Height          =   3500
         Index           =   1
         Left            =   -74940
         TabIndex        =   11
         Top             =   360
         Width           =   8685
         _ExtentX        =   15311
         _ExtentY        =   6174
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|收受者|收信日期時間"
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
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm06010611.frx":0126
         Height          =   3500
         Index           =   2
         Left            =   -74940
         TabIndex        =   12
         Top             =   360
         Width           =   8685
         _ExtentX        =   15311
         _ExtentY        =   6174
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|收受者|收信日期時間"
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
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm06010611.frx":013B
         Height          =   3500
         Index           =   3
         Left            =   -74940
         TabIndex        =   13
         Top             =   360
         Width           =   8685
         _ExtentX        =   15311
         _ExtentY        =   6174
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|收受者|收信日期時間"
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
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm06010611.frx":0150
         Height          =   3500
         Index           =   4
         Left            =   -74940
         TabIndex        =   14
         Top             =   360
         Width           =   8685
         _ExtentX        =   15311
         _ExtentY        =   6174
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|收受者|收信日期時間"
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
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm06010611.frx":0165
         Height          =   3500
         Index           =   5
         Left            =   60
         TabIndex        =   15
         Top             =   360
         Width           =   8685
         _ExtentX        =   15311
         _ExtentY        =   6174
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|收受者|收信日期時間"
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
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm06010611.frx":017A
         Height          =   3500
         Index           =   6
         Left            =   -74940
         TabIndex        =   16
         Top             =   360
         Width           =   8685
         _ExtentX        =   15311
         _ExtentY        =   6174
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|收受者|收信日期時間"
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
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm06010611.frx":018F
         Height          =   3500
         Index           =   7
         Left            =   -74940
         TabIndex        =   17
         Top             =   360
         Width           =   8685
         _ExtentX        =   15311
         _ExtentY        =   6174
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|收受者|收信日期時間"
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
         _Band(0).Cols   =   5
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm06010611.frx":01A4
         Height          =   3495
         Index           =   8
         Left            =   -74940
         TabIndex        =   43
         Top             =   360
         Width           =   8685
         _ExtentX        =   15311
         _ExtentY        =   6174
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|收受者|收信日期時間"
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
         _Band(0).Cols   =   5
      End
   End
   Begin VB.CommandButton cmdTrans 
      Caption         =   "信件轉入及分類"
      Height          =   330
      Left            =   3840
      TabIndex        =   1
      Top             =   0
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
   Begin VB.Frame Frame1 
      Caption         =   "資料修改區"
      ForeColor       =   &H00000080&
      Height          =   1065
      Left            =   60
      TabIndex        =   18
      Top             =   600
      Width           =   8835
      Begin VB.ComboBox cboII05 
         Height          =   300
         ItemData        =   "frm06010611.frx":01B9
         Left            =   5640
         List            =   "frm06010611.frx":01D5
         Style           =   2  '單純下拉式
         TabIndex        =   25
         Top             =   660
         Width           =   1500
      End
      Begin MSForms.ComboBox cboII06 
         Height          =   285
         Left            =   5640
         TabIndex        =   20
         Top             =   330
         Width           =   1500
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2646;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cboII06"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ListBox List1 
         Height          =   765
         Left            =   7140
         TabIndex        =   21
         Top             =   240
         Width           =   1635
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "2884;1270"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   165
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtII17 
         Height          =   465
         Left            =   570
         TabIndex        =   49
         Top             =   210
         Width           =   4335
         VariousPropertyBits=   -1399830505
         BackColor       =   -2147483633
         ScrollBars      =   3
         Size            =   "7646;820"
         Value           =   "txtPI11"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label1 
         Caption         =   "(點二下可移除資料)"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   1
         Left            =   7230
         TabIndex        =   37
         Top             =   30
         Width           =   1575
      End
      Begin VB.Label LblII12 
         Caption         =   "Label7"
         Height          =   225
         Left            =   1320
         TabIndex        =   26
         Top             =   780
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "收信日期時間:"
         Height          =   165
         Left            =   120
         TabIndex        =   24
         Top             =   780
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "收受者:"
         Height          =   255
         Left            =   5010
         TabIndex        =   23
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "分　類:"
         Height          =   255
         Left            =   5010
         TabIndex        =   22
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "主旨:"
         Height          =   255
         Left            =   120
         TabIndex        =   19
         Top             =   240
         Width           =   435
      End
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   -210
      Top             =   1830
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label9 
      Caption         =   $"frm06010611.frx":0213
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
      Left            =   90
      TabIndex        =   51
      Top             =   6630
      Width           =   8880
   End
   Begin VB.Label Label7 
      Caption         =   "主旨有 URGENT 字樣，前頭加入●符號提醒"
      ForeColor       =   &H000000FF&
      Height          =   225
      Left            =   90
      TabIndex        =   48
      Top             =   6090
      Width           =   3675
   End
   Begin VB.Label LblCC 
      AutoSize        =   -1  'True
      Caption         =   $"frm06010611.frx":031B
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
      Height          =   270
      Left            =   90
      TabIndex        =   47
      Top             =   6330
      Width           =   8790
   End
   Begin VB.Label Label1 
      Caption         =   "備註：雙擊”主旨”開啟信件"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   1920
      TabIndex        =   46
      Top             =   1770
      Width           =   2535
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   8
      Left            =   8190
      TabIndex        =   42
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   7
      Left            =   7170
      TabIndex        =   29
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   0
      Left            =   390
      TabIndex        =   36
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   1
      Left            =   1350
      TabIndex        =   35
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   2
      Left            =   2340
      TabIndex        =   34
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   3
      Left            =   3300
      TabIndex        =   33
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   4
      Left            =   4260
      TabIndex        =   32
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   5
      Left            =   5250
      TabIndex        =   31
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H00C0C0FF&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   6
      Left            =   6240
      TabIndex        =   30
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label LblTotCnt 
      Caption         =   "總筆數:"
      ForeColor       =   &H00C00000&
      Height          =   225
      Left            =   255
      TabIndex        =   41
      Top             =   1770
      Width           =   1335
   End
   Begin VB.Label TodayTotCnt 
      AutoSize        =   -1  'True
      Caption         =   "今日總筆數："
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   960
      TabIndex        =   38
      Top             =   60
      Width           =   1080
   End
   Begin VB.Label Label2 
      Caption         =   "信件資料夾："
      Height          =   195
      Left            =   60
      TabIndex        =   8
      Top             =   330
      Width           =   1125
   End
   Begin VB.Label LblCntIPDept 
      ForeColor       =   &H00C00000&
      Height          =   165
      Left            =   4740
      TabIndex        =   28
      Top             =   390
      Visible         =   0   'False
      Width           =   4125
   End
   Begin MSForms.TextBox TextII17 
      Height          =   300
      Left            =   0
      TabIndex        =   50
      Top             =   0
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
End
Attribute VB_Name = "frm06010611"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/4/14 Form2.0已修改
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Dim dblPrevRow(0 To 8) As Double '記錄目前點選那一筆
Private Declare Function FtpRenameFile Lib "wininet.dll" Alias "FtpRenameFileA" (ByVal hFtpSession As Long, ByVal lpszExisting As String, ByVal lpszNew As String) As Boolean
Dim m_AttachPath As String
Dim m_InputGrid As MSHFlexGrid
Dim m_InputCol As Integer, m_InputRow As Integer
Dim nCol As Long, nRow As Long
Public m_QuyII01 As String, m_QuyII02 As String, m_QuyII03 As String
Dim bolCboII06_KeyPress As Boolean 'Add By Sindy 2021/4/14
Dim m_OldKey As String
Dim m_strFileName As String


'分類
Private Sub cboII05_Click()
'   If Left(cboII05.Text, 1) = "Z" Then
'      cboII06.Tag = Left(cboII05.Text, 1)
'      List1.Clear
'      List1.Tag = ""
'      cboII06.ListIndex = 0
   'ElseIf Index = 0 And Left(cboII05.Text, 1) <> cboII05.Tag And cboII05.ListIndex >= 0 Then
   '取消Left(cboII05.Text, 1) <> cboII05.Tag判斷，使其可以再恢後原グ分類
   If cboII05.ListIndex >= 0 Then
      'Modify By Sindy 2016/5/16 + And cboII05.Enabled = True
      If cmdUpdRow.Enabled = False And cboII05.Enabled = True Then
         Call cmdUpdRow_Click
      Else
         If cboII06.Tag <> Left(cboII05.Text, 1) Then
            'Add By Sindy 2024/8/12 有可能已選好收受者,所以不要清除List1
            If Left(cboII05.Text, 1) <> "Z" Then
            '2024/8/12 END
               List1.Clear
               List1.Tag = ""
            End If
         End If
         cboII06.Tag = Left(cboII05.Text, 1)
         Select Case Left(cboII05.Text, 1)
            Case "2" '外商
               cboII06.ListIndex = 1
            Case "3" '外專
               cboII06.ListIndex = 2
            Case "4" '專利處
               cboII06.ListIndex = 3
            Case "5" '外法
               cboII06.ListIndex = 4
            Case "6" '新知
               cboII06.ListIndex = 5
            Case "7" '財務
               cboII06.ListIndex = 6
            'Add By Sindy 2016/6/15
            Case "8" '開拓
               cboII06.ListIndex = 8 '7=taieacc Add By Sindy 2024/10/16
            '2016/6/15 END
            Case Else
            'Case "Z" '其他
               cboII06.ListIndex = 0
         End Select
      End If
   End If
End Sub

'收受者
Private Sub cboII06_Click()
   If bolCboII06_KeyPress = True Then Exit Sub 'Add By Sindy 2021/4/14
   If cboII06.ListIndex >= 0 Then
      'Add By Sindy 2016/4/19 點選收受者新知時,分類則為新知
      If cboII06 = "新知" And List1.Tag = "" Then
         cboII06.Tag = "6"
         cboII05.ListIndex = 5
      'Add By Sindy 2016/6/15
      ElseIf cboII06 = "開拓" And List1.Tag = "" Then
         cboII06.Tag = "8"
         cboII05.ListIndex = 7
      End If
      '2016/4/19 END
      If InStr(List1.Tag, cboII06.List(cboII06.ListIndex)) = 0 Then
         If List1.Tag = "" Then List1.Clear
         List1.AddItem cboII06.List(cboII06.ListIndex)
         bolCboII06_KeyPress = False 'Add By Sindy 2021/4/14
         List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboII06.List(cboII06.ListIndex)
      End If
   End If
   'Add By Sindy 2016/5/16
   If cboII06.Enabled = False Then
      cboII06.Text = ""
   End If
   '2016/5/16 END
End Sub

Private Sub cboII06_Validate(Cancel As Boolean)
   If cboII06.Text <> "" Then
      Call cboII06_LostFocus
'      '檢查人員是否存在或離職
'      If ChkStaffST04(Left(cboII06, 5)) = True Then
'         cboII06.SetFocus
'         Call cboII06_GotFocus
'         Exit Sub
'      End If
      'If Len(Trim(cboII06.Text)) = 5 Then
         'cboII06.Text = Left(cboII06.Text, 5) & " " & GetStaffName(Left(cboII06.Text, 5), True)
         'Add By Sindy 2016/4/18
         If InStr(List1.Tag, cboII06.Text) = 0 Then
            If List1.Tag = "" Then List1.Clear
            List1.AddItem cboII06.Text
            bolCboII06_KeyPress = False 'Add By Sindy 2021/4/14
            List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboII06.Text
         End If
         '2016/4/18 END
         cboII06.Text = ""
      'End If
   End If
End Sub

'Add By Sindy 2016/5/18
Private Sub cboII06_GotFocus()
   cboII06.SelStart = 0
   cboII06.SelLength = Len(cboII06.Text)
End Sub
Private Sub cboII06_KeyPress(KeyAscii As MSForms.ReturnInteger)
   bolCboII06_KeyPress = True 'Add By Sindy 2021/4/14
   KeyAscii = UpperCase(KeyAscii)
End Sub
Private Sub cboII06_LostFocus()
Dim strText As String
Dim bolFind As Boolean, ii As Integer
   
   If cboII06.Text <> "" Then
      '依員工姓名抓取員工編號
      strText = GetPrjSalesNM_2(cboII06.Text)
      If strText <> "" Then
         cboII06.Text = strText & " " & cboII06.Text
      Else
         '依員工編號抓取員工姓名
         strText = GetPrjSalesNM(Left(cboII06.Text, 5))
         If strText <> "" Then
            'Add By Sindy 2021/4/14
            '檢查人員是否離職
            If ChkStaffST04(Left(cboII06.Text, 5)) = True Then
               cboII06.SetFocus
               Call cboII06_GotFocus
               cboII06.Text = ""
               Exit Sub
            Else
            '2021/4/14 END
               cboII06.Text = Left(cboII06.Text, 5) & " " & strText
            End If
         Else
            'Add By Sindy 2021/4/14 檢查是否有在List清單裡, 沒有則不可加入
            bolFind = False
            For ii = 0 To cboII06.ListCount - 1
               'If cboII06.Text = cboII06.List(ii) Then
               If InStr(cboII06.List(ii), cboII06.Text) > 0 Then
                  cboII06.Text = cboII06.List(ii)
                  bolFind = True: Exit For
               End If
            Next ii
            If bolFind = False Then
               cboII06.Text = ""
            End If
            '2021/4/14 END
         End If
      End If
   End If
End Sub
'2016/5/18 END

'Add By Sindy 2016/5/13
'清除反白,並且檢查是否有更新過資料要還原
Private Sub CancelRowColor(Index As Integer, intRow As Integer)
   '清除反白
   GRD1(Index).TextMatrix(intRow, 0) = ""
   If GRD1(Index).TextMatrix(intRow, 9) <> "" Or _
      GRD1(Index).TextMatrix(intRow, 10) <> "" Then
      GRD1(Index).TextMatrix(intRow, 0) = "!"
   End If
   GRD1(Index).col = 0
   GRD1(Index).row = intRow
   For j = 0 To GRD1(Index).Cols - 1
      GRD1(Index).col = j
      GRD1(Index).CellBackColor = QBColor(15)
   Next j
   Me.Tag = Replace(Me.Tag, "," & intRow, "") '清除筆數
End Sub

'Add By Sindy 2016/5/13
'刪除鍵
Private Sub cmdDelRow_Click()
Dim bolHavdDel As Boolean
   
   bolHavdDel = False
   '先檢查是否有資料要刪除
   If GRD1(SSTab1.Tab).Rows - 1 < 1 Then Exit Sub
   If GRD1(SSTab1.Tab).Rows - 1 >= 1 And GRD1(SSTab1.Tab).TextMatrix(1, 12) = "" Then Exit Sub
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
               Call CancelRowColor(SSTab1.Tab, i) '清除反白,並且檢查是否有更新過資料要還原
            End If
         Next i
         Exit Sub
      End If
   End If
   
On Error GoTo ErrHand
   
   Screen.MousePointer = vbHourglass
   For i = 1 To GRD1(SSTab1.Tab).Rows - 1
      If GRD1(SSTab1.Tab).TextMatrix(i, 0) = "V" Then
         'Add By Sindy 2022/2/22 針對分類有*號時，要增加提醒訊息。
         If InStr(GRD1(SSTab1.Tab).TextMatrix(i, 2), "*") > 0 Then
            If MsgBox(GRD1(SSTab1.Tab).TextMatrix(i, 1) & vbCrLf & vbCrLf & _
                   "提醒：信件狀態有一邊是直接刪除的，會在信箱代號旁加上*號，例如(P*,F)，若人員處理信件時發現非該單位信件，請轉寄回該信箱。" & vbCrLf & vbCrLf & _
                   "確定要刪除信件嗎？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
               'Call CancelRowColor(SSTab1.Tab, i) '清除反白,並且檢查是否有更新過資料要還原
               Screen.MousePointer = vbDefault
               Exit Sub
            End If
         End If
         '2022/2/22 END
         
         strExc(0) = "update IPDeptInput set " & _
                     " ii07='Y',ii08=" & strSrvDate(1) & ",ii09=to_char(sysdate, 'HH24MISS'),ii10='" & strUserNum & "',ii16=" & strSrvDate(1) & _
                     " where ii01=" & GRD1(SSTab1.Tab).TextMatrix(i, 7) & _
                       " and ii02=" & GRD1(SSTab1.Tab).TextMatrix(i, 8) & _
                       " and ii03='" & ChgSQL(GRD1(SSTab1.Tab).TextMatrix(i, 12)) & "'"
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
         LblTotCnt.Caption = "總筆數: " & Val(Replace(LblTotCnt.Caption, "總筆數:", "")) - 1
         strExc(0) = LblRow(SSTab1.Tab).Caption
         strExc(0) = Val(strExc(0)) - 1
         If SSTab1.Tab = 0 Then SSTab1.Caption = "個案": LblRow(0).Caption = strExc(0)
         If SSTab1.Tab = 1 Then SSTab1.Caption = "外商": LblRow(1).Caption = strExc(0)
         If SSTab1.Tab = 2 Then SSTab1.Caption = "外專": LblRow(2).Caption = strExc(0)
         If SSTab1.Tab = 3 Then SSTab1.Caption = "專利處": LblRow(3).Caption = strExc(0)
         If SSTab1.Tab = 4 Then SSTab1.Caption = "外法": LblRow(4).Caption = strExc(0)
         If SSTab1.Tab = 5 Then SSTab1.Caption = "新知": LblRow(5).Caption = strExc(0)
         If SSTab1.Tab = 6 Then SSTab1.Caption = "財務": LblRow(6).Caption = strExc(0)
         If SSTab1.Tab = 7 Then SSTab1.Caption = "其他": LblRow(7).Caption = strExc(0)
         If SSTab1.Tab = 8 Then SSTab1.Caption = "開拓": LblRow(8).Caption = strExc(0) 'Add By Sindy 2016/6/15
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
               " where mrl01='" & Left(IPDept收件匣, 2) & "'" & _
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
                     "values('" & Left(IPDept收件匣, 2) & "'," & strSrvDate(1) & ",to_char(sysdate, 'HH24MISS'),'" & strUserNum & "','A')"
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
   
   Call frm06010613.SetParent(Me) 'Modify By Sindy 2016/9/30
   frm06010613.m_WorkType = "0" '信箱主檔
   'frm06010613.QueryData
   frm06010613.Show
   Me.Hide
End Sub

''Add By Sindy 2017/3/23
'Private Sub cmdManualRecv_Click()
'Dim bolCancel As Boolean '中斷
'
'   If cmdManualRecv.Caption = "啟動接收郵件" Then
'      bolCancel = False
'      cmdManualRecv.Caption = "中斷接收郵件"
'      Call frmTaOutLook.Hide
'   Else
'      bolCancel = True '中斷
'      cmdManualRecv.Caption = "啟動接收郵件"
'   End If
'   If frmTaOutLook.userControlFCPin(bolCancel) = True Then
'      GetTodayTotCnt
'      Call QueryData
'   End If
'End Sub

Private Sub cmdQuery_Click()
   QueryData
End Sub

Private Sub cmdRecOutlookQ_Click()
   frm06010615.m_QueryType = "F"
   frm06010615.Hide
   frm06010615.cmdQuery_Click
   frm06010615.Show vbModal
End Sub

'整批轉寄
Private Sub cmdSendMail_Click()
Dim intTab As Long
Dim strFileName As String
Dim tmpArr As Variant
Dim strTo As String, strToSys As String
Dim strSubject As String
Dim strContext As String
Dim strEMailTo As String '串要發通知信的人員
Dim strUpdTime As String 'Add By Sindy 2016/5/13
'Add By Sindy 2016/9/13
Dim strII11 As String, strII12 As String
Dim strII13 As String, strII17 As String
Dim strPI03 As String, strPI03_2 As String
Dim stFtpPath As String, bolSaveEFile As Boolean
'2016/9/13 END
Dim strTableName As String, strUpdWhere As String, strII20 As String 'Add By Sindy 2017/11/21
Dim i As Integer 'Add By Sindy 2018/10/26
Dim strTi03 As String, strTi03_2 As String
Dim strToCC As String
Dim bolSendMailErr  As Boolean
Dim strF1xEmp As String, strF2xEmp As String
Dim varTmp As Variant, jj As Integer
   
   Screen.MousePointer = vbHourglass
   '先更新資料
   If cmdSave.Enabled = True Then
      Call CmdSave_Click
   Else
      Call QueryData
   End If
   
On Error GoTo ErrHand
   
   cmdSendMail.Enabled = False 'Add By Sindy 2016/5/10
   strEMailTo = ""
   'Modify By Sindy 2016/11/15 副所長決定信件只要幫David轉發即可
   '　　其他還是發到個人的Outlook信箱,人員才不用2邊看信
   '    且有時工程師在一個月內處理不完信件,其信件就被系統刪掉等問題
   For intTab = 0 To 8
      If (GRD1(intTab).Rows - 1) > 0 Then
         For i = 1 To GRD1(intTab).Rows - 1
            If Trim(GRD1(intTab).TextMatrix(i, 12)) <> "" And _
               Trim(GRD1(intTab).TextMatrix(i, 6)) <> "" Then '有檔名有收受者時,則處理資料(轉寄)
               
               '產生實體檔案
               'Modify By Sindy 2016/10/4
               'strFileName = GRD1(intTab).TextMatrix(i, 12)
               strFileName = Mid(GRD1(intTab).TextMatrix(i, 13), InStrRev(GRD1(intTab).TextMatrix(i, 13), "/") + 1)
               '2016/10/4 END
               If GetAttachFile(GRD1(intTab).TextMatrix(i, 7), GRD1(intTab).TextMatrix(i, 8), GRD1(intTab).TextMatrix(i, 12), strFileName, "", m_AttachPath & "\" & strFileName) = False Then
                  GoTo ReadNext
               End If
               'Add By Sindy 2017/1/10 信包信裡面沒有附件檔,等待下載信件
               Do While Dir(strFileName, vbDirectory) = ""
                  DoEvents
               Loop
               '2017/1/10 END
               
               strUpdTime = Right("000000" & ServerTime, 6) 'Add By Sindy 2016/5/13
               
               cnnConnection.BeginTrans
               
               strPI03 = "" 'Add By Sindy 2022/2/14
               strTi03 = "" 'Add By Sindy 2022/2/14
               strPI03_2 = ""
               strTi03_2 = ""
               'Add By Sindy 2016/9/13 檢查收受者若是有專利處(patent)信件也複製一份至專利處收件夾資料
               '                       若純寄patent則國外部信件要上刪除日期註記
               '                       若有其他單位人員則不需要上註記
               'Modify By Sindy 2019/6/20
               GRD1(intTab).TextMatrix(i, 6) = PUB_IR04DataMakeUp(GRD1(intTab).TextMatrix(i, 6))
               If InStr(UCase(GRD1(intTab).TextMatrix(i, 6)), "PATENT") > 0 Or _
                  (strSrvDate(1) >= TM分信系統啟用日 And _
                  InStr(UCase(GRD1(intTab).TextMatrix(i, 6)), "TM") > 0) Then
                  '取得國外部收件夾資料
                  strExc(0) = "select ii11,ii12,ii13,ii17 from IPDeptInput" & _
                              " where ii01=" & GRD1(intTab).TextMatrix(i, 7) & _
                              " and ii02=" & GRD1(intTab).TextMatrix(i, 8) & _
                              " and ii03='" & GRD1(intTab).TextMatrix(i, 12) & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strII11 = "" & RsTemp.Fields("ii11")
                     strII12 = "" & RsTemp.Fields("ii12")
                     strII13 = "" & RsTemp.Fields("ii13")
                     strII17 = "" & RsTemp.Fields("ii17")
                  End If
                  
                  '新增專利處收件夾
                  If InStr(UCase(GRD1(intTab).TextMatrix(i, 6)), "PATENT") > 0 Then
'                           strExc(0) = "select count(*) from PatentInput" & _
'                                       " where PI01=" & strSrvDate(1)
'                           intI = 1
'                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                           If intI = 1 Then
'                              intPI03 = Val(RsTemp.Fields(0)) + 1
'                           Else
'                              intPI03 = 1
'                           End If
'                           strPI03 = "P" & Format(intPI03, "0000")
                     'Modify By Sindy 2019/12/2 自動給號,才能 Keep PKey
                     strPI03 = AutoNoByDate("P", 4)
                     '2019/12/2 END
                     strPI03_2 = strSrvDate(1) & strUpdTime & "." & strPI03 & ".msg"
                     bolSaveEFile = PUB_PutFtpFile(strFileName, strSrvDate(1), strPI03_2, stFtpPath, UCase("PatentInput"))
                     If bolSaveEFile = True Then
                        '存資料到專利處收件夾資料
                        strSql = "insert into PatentInput(PI01,PI02,PI03,PI04,PI05,PI11,PI12,PI13,PI14,PI17,PI15)" & _
                                 " values(" & strSrvDate(1) & "," & strUpdTime & _
                                 ",'" & strPI03 & "','" & strUserNum & "',null" & _
                                 "," & CNULL(ChgSQL(strII11)) & "," & strII12 & "," & CNULL(strII13) & _
                                 ",'" & ChgSQL(stFtpPath) & "','" & ChgSQL(strII17) & "','IPDept')"
                        cnnConnection.Execute strSql
                        
                        'Add By Sindy 2022/8/11
                        '新增郵件轉信讀取記錄(直接沖銷)
                        strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR15,IR08,IR09,IR10)" & _
                                    " values(" & GRD1(intTab).TextMatrix(i, 7) & _
                                             "," & GRD1(intTab).TextMatrix(i, 8) & _
                                             ",'" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'" & _
                                             ",'patent'," & strSrvDate(1) & "," & _
                                             strUpdTime & ",'" & strUserNum & "','Y'," & _
                                             strSrvDate(1) & "," & strUpdTime & ",'" & strUserNum & "')"
                        cnnConnection.Execute strExc(0)
                        
                        'Add By Sindy 2019/7/5 增加副本給專利處主管
                        If OL_SendNotifyMailCC("IPDept", "Patent", strFileName, strII17, strSrvDate(1), strUpdTime, strPI03, OL_PatMailCC, strSrvDate(1), strUpdTime) = False Then
                           GoTo ErrHand
                        End If
                     End If
                     
                     '收受者拿掉patent
                     GRD1(intTab).TextMatrix(i, 6) = Replace(UCase(GRD1(intTab).TextMatrix(i, 6)), UCase("patent"), "")
                     GRD1(intTab).TextMatrix(i, 6) = Replace(GRD1(intTab).TextMatrix(i, 6), ";;", ";")
                     If GRD1(intTab).TextMatrix(i, 6) = ";" Then GRD1(intTab).TextMatrix(i, 6) = ""
                     If GRD1(intTab).TextMatrix(i, 6) <> "" Then
                        If Left(GRD1(intTab).TextMatrix(i, 6), 1) = ";" Then GRD1(intTab).TextMatrix(i, 6) = Mid(GRD1(intTab).TextMatrix(i, 6), 2)
                        If Right(GRD1(intTab).TextMatrix(i, 6), 1) = ";" Then GRD1(intTab).TextMatrix(i, 6) = Mid(GRD1(intTab).TextMatrix(i, 6), 1, Len(GRD1(intTab).TextMatrix(i, 6)) - 1)
                     End If
                  End If
                  
                  '新增商標處收件夾
                  If (strSrvDate(1) >= TM分信系統啟用日 And _
                     InStr(UCase(GRD1(intTab).TextMatrix(i, 6)), "TM") > 0) Then
                     
                     'Modify By Sindy 2019/12/2 自動給號,才能 Keep PKey
                     strTi03 = AutoNoByDate("T", 4)
                     '2019/12/2 END
                     strTi03_2 = strSrvDate(1) & strUpdTime & "." & strTi03 & ".msg"
                     bolSaveEFile = PUB_PutFtpFile(strFileName, strSrvDate(1), strTi03_2, stFtpPath, UCase("TMInput"))
                     If bolSaveEFile = True Then
                        '存資料到商標處收件夾資料
                        strSql = "insert into TMInput(TI01,TI02,TI03,TI04,TI05,TI11,TI12,TI13,TI14,TI17,TI15)" & _
                                 " values(" & strSrvDate(1) & "," & strUpdTime & _
                                 ",'" & strTi03 & "','" & strUserNum & "',null" & _
                                 "," & CNULL(ChgSQL(strII11)) & "," & strII12 & "," & CNULL(strII13) & _
                                 ",'" & ChgSQL(stFtpPath) & "','" & ChgSQL(strII17) & "','IPDept')"
                        cnnConnection.Execute strSql
                        
                        'Add By Sindy 2022/8/11
                        '新增郵件轉信讀取記錄(直接沖銷)
                        strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR15,IR08,IR09,IR10)" & _
                                    " values(" & GRD1(intTab).TextMatrix(i, 7) & _
                                             "," & GRD1(intTab).TextMatrix(i, 8) & _
                                             ",'" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'" & _
                                             ",'TM'," & strSrvDate(1) & "," & _
                                             strUpdTime & ",'" & strUserNum & "','Y'," & _
                                             strSrvDate(1) & "," & strUpdTime & ",'" & strUserNum & "')"
                        cnnConnection.Execute strExc(0)

                        'Add By Sindy 2019/7/5 增加副本給商標處主管
                        If OL_SendNotifyMailCC("IPDept", "TM", strFileName, strII17, strSrvDate(1), strUpdTime, strTi03, OL_TmMailCC, strSrvDate(1), strUpdTime) = False Then
                           GoTo ErrHand
                        End If
                     End If
                     
                     '收受者拿掉TM
                     GRD1(intTab).TextMatrix(i, 6) = Replace(UCase(GRD1(intTab).TextMatrix(i, 6)), UCase("tm"), "")
                     GRD1(intTab).TextMatrix(i, 6) = Replace(GRD1(intTab).TextMatrix(i, 6), ";;", ";")
                     If GRD1(intTab).TextMatrix(i, 6) = ";" Then GRD1(intTab).TextMatrix(i, 6) = ""
                     If GRD1(intTab).TextMatrix(i, 6) <> "" Then
                        If Left(GRD1(intTab).TextMatrix(i, 6), 1) = ";" Then GRD1(intTab).TextMatrix(i, 6) = Mid(GRD1(intTab).TextMatrix(i, 6), 2)
                        If Right(GRD1(intTab).TextMatrix(i, 6), 1) = ";" Then GRD1(intTab).TextMatrix(i, 6) = Mid(GRD1(intTab).TextMatrix(i, 6), 1, Len(GRD1(intTab).TextMatrix(i, 6)) - 1)
                     End If
                  End If
                  
                  '記錄流水號
                  strExc(0) = "update IPDeptInput set " & _
                              " ii10='" & strUserNum & "',ii15='" & strPI03 & IIf(strPI03 <> "" And strTi03 <> "", ",", "") & strTi03 & "'" & _
                              " where ii01=" & GRD1(intTab).TextMatrix(i, 7) & _
                                " and ii02=" & GRD1(intTab).TextMatrix(i, 8) & _
                                " and ii03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'" & _
                                " and ii08=0"
                  cnnConnection.Execute strExc(0)
                  '收受者只分給 PATENT 或 TM時, 要上轉寄日期時間人員,刪除實體檔日期
                  If GRD1(intTab).TextMatrix(i, 6) = "" Then
                     strExc(0) = "update IPDeptInput set " & _
                                 " ii08=" & strSrvDate(1) & ",ii09=" & strUpdTime & ",ii10='" & strUserNum & "',ii16=" & strSrvDate(1) & _
                                 " where ii01=" & GRD1(intTab).TextMatrix(i, 7) & _
                                   " and ii02=" & GRD1(intTab).TextMatrix(i, 8) & _
                                   " and ii03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'" & _
                                   " and ii08=0"
                     cnnConnection.Execute strExc(0)
                  End If
               End If
               
               strTo = GRD1(intTab).TextMatrix(i, 6) '收受者
               strToSys = "": strExc(10) = "": bolSendMailErr = False
               If strTo <> "" Then
                  'Modify By Sindy 2022/5/25 外專部門(F22,F23)信件收錄進系統收件區,其他維持Outlook轉寄
                  'Modify By Sindy 2022/8/10 David說新知,財務,開拓也不要列入系統收件區
                  If Val(strSrvDate(1)) >= 外專信件沖銷啟用日 And intTab <> 5 And intTab <> 6 And intTab <> 8 Then
                     tmpArr = Split(strTo, ";")
                     For j = 0 To UBound(tmpArr)
                        If tmpArr(j) <> "" Then
                           'Add By Sindy 2023/7/14 檢查要新增的收受者是否已有紀錄
                           strExc(0) = "select * from inputrecord" & _
                                       " where IR01=" & GRD1(intTab).TextMatrix(i, 7) & _
                                       " and IR03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'" & _
                                       " and IR04='" & tmpArr(j) & "'"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 0 Then
                           '2023/7/14 END
                              If PUB_GetST03(CStr(tmpArr(j))) <> "" Then
                                 '外專部門(F22,F23)信件收錄進系統收件區
                                 'Modify By Sindy 2023/5/24 + (Val(strSrvDate(1)) >= 外商CF信件沖銷啟用日 And PUB_IPDept_IsCFMail(strChkEmp) = True)
                                 If Trim(PUB_GetST03(CStr(tmpArr(j)))) = "F22" Or _
                                    Trim(PUB_GetST03(CStr(tmpArr(j)))) = "F23" Or _
                                    (Val(strSrvDate(1)) >= 外商CF信件沖銷啟用日 And PUB_IPDept_IsCFMail(CStr(tmpArr(j))) = True) Then
                                    
                                    '收錄進系統收件區的收受者
                                    If strToSys <> "" Then strToSys = strToSys & ";"
                                    strToSys = strToSys & Trim(tmpArr(j))
                                    '記錄要發通知信的收受者
                                    If InStr(strEMailTo, Trim(tmpArr(j))) = 0 Then
                                       If strEMailTo <> "" Then strEMailTo = strEMailTo & ";"
                                       strEMailTo = strEMailTo & Trim(tmpArr(j))
                                    End If
                                    '新增郵件轉信讀取記錄
                                    strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR15)" & _
                                                " values(" & GRD1(intTab).TextMatrix(i, 7) & _
                                                         "," & GRD1(intTab).TextMatrix(i, 8) & _
                                                         ",'" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'" & _
                                                         ",'" & tmpArr(j) & "'," & strSrvDate(1) & "," & _
                                                         strUpdTime & ",'" & strUserNum & "','Y')"
                                    cnnConnection.Execute strExc(0)
                                 '非外專部門
                                 Else
                                    '維持Outlook轉寄的收受者
                                    If strExc(10) <> "" Then strExc(10) = strExc(10) & ";"
                                    strExc(10) = strExc(10) & Trim(tmpArr(j))
                                    
                                    'Add By Sindy 2022/8/11
                                    '新增郵件轉信讀取記錄(直接沖銷)
                                    strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR15,IR08,IR09,IR10)" & _
                                                " values(" & GRD1(intTab).TextMatrix(i, 7) & _
                                                         "," & GRD1(intTab).TextMatrix(i, 8) & _
                                                         ",'" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'" & _
                                                         ",'" & Trim(tmpArr(j)) & "'," & strSrvDate(1) & "," & _
                                                         strUpdTime & ",'" & strUserNum & "','Y'," & _
                                                         strSrvDate(1) & "," & strUpdTime & ",'" & strUserNum & "')"
                                    cnnConnection.Execute strExc(0)
                                 End If
                              '無部門別
                              Else
                                 '維持Outlook轉寄的收受者
                                 If strExc(10) <> "" Then strExc(10) = strExc(10) & ";"
                                 strExc(10) = strExc(10) & Trim(tmpArr(j))
                                 
                                 'Add By Sindy 2022/8/11
                                 '新增郵件轉信讀取記錄(直接沖銷)
                                 strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR15,IR08,IR09,IR10)" & _
                                             " values(" & GRD1(intTab).TextMatrix(i, 7) & _
                                                      "," & GRD1(intTab).TextMatrix(i, 8) & _
                                                      ",'" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'" & _
                                                      ",'" & Trim(tmpArr(j)) & "'," & strSrvDate(1) & "," & _
                                                      strUpdTime & ",'" & strUserNum & "','Y'," & _
                                                      strSrvDate(1) & "," & strUpdTime & ",'" & strUserNum & "')"
                                 cnnConnection.Execute strExc(0)
                              End If
                           End If
                        End If
                     Next j
                     If strToSys <> "" Then strTo = strExc(10) '*****
                  End If
                  '2022/5/25 END
               End If
               
               '用Outlookl轉寄
               If strTo <> "" Then
                  strSubject = GRD1(intTab).TextMatrix(i, 1) 'Left(strFileName, Len(strFileName) - 4) Modify By Sindy 2016/4/14 改放主旨
                  
                  'Add By Sindy 2022/8/11
                  strExc(10) = ""
                  If strToSys <> "" Or InStr(UCase(GRD1(intTab).TextMatrix(i, 3)), UCase("patent")) > 0 Or _
                        InStr(UCase(GRD1(intTab).TextMatrix(i, 3)), UCase("TM")) > 0 Then
                     If strToSys <> "" Then strExc(10) = PUB_ReadUserData(strToSys)
                     If InStr(UCase(GRD1(intTab).TextMatrix(i, 3)), UCase("patent")) > 0 Then
                        If strExc(10) <> "" Then strExc(10) = strExc(10) & ","
                        strExc(10) = strExc(10) & "patent"
                     End If
                     If InStr(UCase(GRD1(intTab).TextMatrix(i, 3)), UCase("TM")) > 0 Then
                        If strExc(10) <> "" Then strExc(10) = strExc(10) & ","
                        strExc(10) = strExc(10) & "TM"
                     End If
                  End If
                  If strExc(10) <> "" Then strExc(10) = "同時信件已分給下列人員：" & strExc(10) & vbCrLf & vbCrLf
                  '2022/8/11 END
                  
                  'modify by sonia 2018/11/6 改內容,原為"同仁收到由INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 請及時轉寄或通報國外部顏副理(77015@taie.com.tw)，以便國外部迅速轉發適當收信同仁, 謝謝!"
                  strContext = strExc(10) & "信件內容參附件" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                               "同仁收到由 INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 或該封e-mail非貴單位之案件, 請務必及時再轉寄 ipdept@taie.com.tw 以便國外部迅速轉發適當收信同仁, 謝謝!"
                  
                  '轉寄信件
                  'PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , strFileName
                  'Modify By Sindy 2016/4/14 + 寄件者為ipdept
                  'Modify By Sindy 2017/6/2
                  '*注意* 若要寄patent信箱,要傳入完整email:patent@taie.com.tw
                  '********************
                  'Add By Sindy 2017/11/21 SendMail記錄
                  PStr_SendMailKey1 = GRD1(intTab).TextMatrix(i, 7)
                  PStr_SendMailKey2 = GRD1(intTab).TextMatrix(i, 8)
                  PStr_SendMailKey3 = ChgSQL(GRD1(intTab).TextMatrix(i, 12))
                  strTableName = "IPDEPTINPUT"
                  strUpdWhere = "II01=" & PStr_SendMailKey1 & " and II02=" & PStr_SendMailKey2 & " and II03='" & PStr_SendMailKey3 & "'"
                  strExc(0) = "update " & strTableName & _
                              " set II20='S',II21=" & strSrvDate(1) & ",II22=to_char(sysdate, 'HH24MISS')" & _
                              " where " & strUpdWhere
                  cnnConnection.Execute strExc(0)
                  '2017/11/21 END
                  '********************
                  If intTab = 5 Then '新知不轉職代
                     PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , strFileName, , , , , IIf(strSrvDate(1) >= "20170706", "inbound@taie.com.tw", "ipdept@taie.com.tw"), , , True, False, , , , strUpdWhere, strTableName, , , , , , , , "1"
                  Else
                  '2017/6/2 END
                     PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , strFileName, , , , strToCC, IIf(strSrvDate(1) >= "20170706", "inbound@taie.com.tw", "ipdept@taie.com.tw"), , , , False, , , , strUpdWhere, strTableName, , , , , , , , "1"
                  End If
                  '2016/4/14 END
                  If bolMailSendOk = True Then
                     'Modify By Sindy 2018/10/29 Mark
'                     'Add By Sindy 2017/11/21
'                     If PStr_SendMailKey1 = GRD1(intTab).TextMatrix(i, 7) And _
'                        PStr_SendMailKey2 = GRD1(intTab).TextMatrix(i, 8) And _
'                        PStr_SendMailKey3 = ChgSQL(GRD1(intTab).TextMatrix(i, 12)) Then
'                        strII20 = "Y"
'                     Else
'                        '不同
'                        If PStr_SendMailKey1 & PStr_SendMailKey2 & PStr_SendMailKey3 = "" Then
'                           strII20 = "N"
'                        Else
'                           strII20 = PStr_SendMailKey1 & PStr_SendMailKey2 & PStr_SendMailKey3
'                        End If
'                     End If
'                     '2017/11/21 END
                     'Modify By Sindy 2018/10/29
                     strExc(0) = "select ii20 from " & strTableName & _
                                 " where " & strUpdWhere
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        If "" & RsTemp.Fields(0) = "T" Then
                           strExc(0) = "update " & strTableName & _
                                       " set II20='Y'" & _
                                       " where " & strUpdWhere
                           cnnConnection.Execute strExc(0)
                           'Modify By Sindy 2018/10/31 移至此處更新,有SendMail成功T才更新轉寄日期
'                           strExc(0) = "update IPDeptInput set " & _
'                                       " ii08=" & strSrvDate(1) & ",ii09=to_char(sysdate, 'HH24MISS'),ii10='" & strUserNum & "',ii16=" & strSrvDate(1) & _
'                                       " where ii01=" & GRD1(intTab).TextMatrix(i, 7) & _
'                                         " and ii02=" & GRD1(intTab).TextMatrix(i, 8) & _
'                                         " and ii03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'" & _
'                                         " and ii08=0"
'                           cnnConnection.Execute strExc(0)
                        Else
                           strExc(0) = "update " & strTableName & _
                                       " set II20='" & PStr_SendMailKey1 & PStr_SendMailKey2 & PStr_SendMailKey3 & "'" & _
                                       " where " & strUpdWhere
                           cnnConnection.Execute strExc(0)
                           If UCase(pub_DbTerminalName) = 正式資料庫電腦名稱 Then bolSendMailErr = True
                        End If
                     End If
                     '2018/10/29 END
                  End If
                  
                  '********************
                  'Add By Sindy 2017/11/21 SendMail記錄
                  PStr_SendMailKey1 = ""
                  PStr_SendMailKey2 = ""
                  PStr_SendMailKey3 = ""
                  strTableName = ""
                  strUpdWhere = ""
                  '2017/11/21 END
                  '********************
               End If
               
               '寄信失敗,RollbackTrans
               If bolSendMailErr = True Then
                  cnnConnection.RollbackTrans
               Else
               
                  If (strToSys <> "" Or strTo <> "") And bolSendMailErr = False Then
                     '上轉寄日期時間人員
                     strExc(0) = "update IPDeptInput set " & _
                                 " ii08=" & strSrvDate(1) & ",ii09=" & strUpdTime & ",ii10='" & strUserNum & "'" & _
                                 " where ii01=" & GRD1(intTab).TextMatrix(i, 7) & _
                                   " and ii02=" & GRD1(intTab).TextMatrix(i, 8) & _
                                   " and ii03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'" & _
                                   " and ii08=0"
                     cnnConnection.Execute strExc(0)
                  End If
                  
                  'Add By Sindy 2022/5/27 無收錄進系統收件區的信件, 才能直接上刪除實體檔日期
                  If strToSys = "" And strTo <> "" And bolSendMailErr = False Then
                  '2022/5/27 END
                     '上刪除實體檔日期
                     strSql = "update IPDeptInput set" & _
                              " ii16=" & strSrvDate(1) & _
                              " where ii01=" & GRD1(intTab).TextMatrix(i, 7) & _
                                " and ii02=" & GRD1(intTab).TextMatrix(i, 8) & _
                                " and ii03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'" & _
                                " and ii08>0" '*****
                     cnnConnection.Execute strSql
                  End If
                  cnnConnection.CommitTrans
               End If
               
               '刪除PC端檔案
               'Call fs.DeleteFile(m_AttachPath & "\" & strFileName)
               Kill strFileName
            End If
ReadNext:
         Next i
      End If
   Next intTab
   'Modify By Sindy 2023/5/24
   'Call PUB_SendNotifyMail(strEMailTo, True) '寄發通知信
   '區分部門
   strF1xEmp = "": strF2xEmp = ""
   varTmp = Split(strEMailTo, ";")
   For jj = 0 To UBound(varTmp)
      If Left(PUB_GetST03(CStr(varTmp(jj))), 2) = "F1" Then '外商
         strF1xEmp = strF1xEmp & ";" & varTmp(jj)
      Else
         strF2xEmp = strF2xEmp & ";" & varTmp(jj)
      End If
   Next jj
   If strF1xEmp <> "" Then
      strF1xEmp = Mid(strF1xEmp, 2)
      Call PUB_SendNotifyMail(strF1xEmp, True) '寄發通知信
   End If
   If strF2xEmp <> "" Then
      strF2xEmp = Mid(strF2xEmp, 2)
      Call PUB_SendNotifyMail(strF2xEmp, True) '寄發通知信
   End If
   '2023/5/24 END
   Call PUB_SendMailCache(, , False) 'Add By Sindy 2019/7/17
'   DoEvents 'Add By Sindy 2016/5/10
'
'   '除了專利處外,其他同新知及開拓方式處理
''   '最後處理新知mail
''   For j = 0 To 1
''      If j = 0 Then intTab = 5 '新知
''      If j = 1 Then intTab = 8 '開拓
''   Set objOutLook = CreateObject("Outlook.Application") 'Add By Sindy 2017/3/10
'   For intTab = 0 To 8
'      'If intTab <> 3 Then '踢除專利處
'         For i = 1 To GRD1(intTab).Rows - 1
'            If GRD1(intTab).TextMatrix(i, 12) <> "" And _
'               GRD1(intTab).TextMatrix(i, 6) <> "" Then '有檔名有收受者時,則處理資料(轉寄)
'               '讀取檔案
'               'Modify By Sindy 2016/10/4
'               'strFileName = GRD1(intTab).TextMatrix(i, 12)
'               strFileName = Mid(GRD1(intTab).TextMatrix(i, 13), InStrRev(GRD1(intTab).TextMatrix(i, 13), "/") + 1)
'               '2016/10/4
'               strTo = GRD1(intTab).TextMatrix(i, 6)
'
'               '用Outlookl轉寄
'               If strTo <> "" Then
'                  strSubject = GRD1(intTab).TextMatrix(i, 1) 'Left(strFileName, Len(strFileName) - 4) Modify By Sindy 2016/4/14 改放主旨
'                  'modify by sonia 2018/11/6 改內容,原為"同仁收到由INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 請及時轉寄或通報國外部顏副理(77015@taie.com.tw)，以便國外部迅速轉發適當收信同仁, 謝謝!"
'                  strContext = "信件內容參附件" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
'                               "同仁收到由 INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 或該封e-mail非貴單位之案件, 請務必及時再轉寄 ipdept@taie.com.tw 以便國外部迅速轉發適當收信同仁, 謝謝!"
'                  cnnConnection.BeginTrans
'                  If GetAttachFile(GRD1(intTab).TextMatrix(i, 7), GRD1(intTab).TextMatrix(i, 8), GRD1(intTab).TextMatrix(i, 12), strFileName, "", m_AttachPath & "\" & strFileName) = True Then
'                     'Add By Sindy 2017/1/10 信包信裡面沒有附件檔,等待下載信件
'                     Do While Dir(strFileName, vbDirectory) = ""
'                        DoEvents
'                     Loop
'                     '2017/1/10 END
'                     'Modify By Sindy 2016/4/12 因為新知檔案很大,開啟信件寄送會很花時間,改為信包信方式
'                     '新知:開啟新郵件
'         '            Call OpenNeweMail(strTo, strSubject, strContext, strFileName)
'
'   '                  '*** 轉寄 *** 無法異動寄件者還是有問題
'   '                  Set objMail = objOutLook.CreateItemFromTemplate(strFileName) '原始信
'   '                  objMail.Forward
'   '                  objMail.Display
'   '                  '移除原信的收件人及副本;密件副本不會留在msg中
'   '                  For ii = objMail.Recipients.Count To 1 Step -1
'   '                     objMail.Recipients.Remove ii
'   '                  Next ii
'   '                  objMail.SenderName = "ipdept" '?????不行,還是唯讀
'   '                  '副本.cc
'   '                  '收件者.To
'   '                  objMail.To = "97038" 'strTo
'   '                  '密件副本.BCC
'   '                  objMail.Recipients.add "97038" '97038@taie.com.tw
'   '                  objMail.Send
'   '                  DoEvents
'   '                  '*** END
'
'                     '轉寄信件
'                     'PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , strFileName
'                     'Modify By Sindy 2016/4/14 + 寄件者為ipdept
'                     'Modify By Sindy 2017/6/2
'                     '*注意* 若要寄patent信箱,要傳入完整email:patent@taie.com.tw
'                     '********************
'                     'Add By Sindy 2017/11/21 SendMail記錄
'                     PStr_SendMailKey1 = GRD1(intTab).TextMatrix(i, 7)
'                     PStr_SendMailKey2 = GRD1(intTab).TextMatrix(i, 8)
'                     PStr_SendMailKey3 = ChgSQL(GRD1(intTab).TextMatrix(i, 12))
'                     strTableName = "IPDEPTINPUT"
'                     strUpdWhere = "II01=" & PStr_SendMailKey1 & " and II02=" & PStr_SendMailKey2 & " and II03='" & PStr_SendMailKey3 & "'"
'                     strExc(0) = "update " & strTableName & _
'                                 " set II20='S',II21=" & strSrvDate(1) & ",II22=to_char(sysdate, 'HH24MISS')" & _
'                                 " where " & strUpdWhere
'                     cnnConnection.Execute strExc(0)
'                     '2017/11/21 END
'                     '********************
'                     If intTab = 5 Then '新知不轉職代
'                        PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , strFileName, , , , , IIf(strSrvDate(1) >= "20170706", "inbound@taie.com.tw", "ipdept@taie.com.tw"), , , True, False, , , , strUpdWhere, strTableName
'                     Else
'                     '2017/6/2 END
'                        'Add By Sindy 2019/7/5 增加副本給商標處主管
'                        strToCC = ""
'                        If InStr(UCase(GRD1(intTab).TextMatrix(i, 6)), "TM") > 0 Then
'                           strToCC = OL_TmMailCC
'                        End If
'                        '2019/7/5 END
'                        PUB_SendMail strUserNum, strTo, "", strSubject, strContext, , strFileName, , , , strToCC, IIf(strSrvDate(1) >= "20170706", "inbound@taie.com.tw", "ipdept@taie.com.tw"), , , , False, , , , strUpdWhere, strTableName
'                     End If
'                     '2016/4/14 END
'                     If bolMailSendOk = True Then
'                        'Modify By Sindy 2018/10/29 Mark
'   '                     'Add By Sindy 2017/11/21
'   '                     If PStr_SendMailKey1 = GRD1(intTab).TextMatrix(i, 7) And _
'   '                        PStr_SendMailKey2 = GRD1(intTab).TextMatrix(i, 8) And _
'   '                        PStr_SendMailKey3 = ChgSQL(GRD1(intTab).TextMatrix(i, 12)) Then
'   '                        strII20 = "Y"
'   '                     Else
'   '                        '不同
'   '                        If PStr_SendMailKey1 & PStr_SendMailKey2 & PStr_SendMailKey3 = "" Then
'   '                           strII20 = "N"
'   '                        Else
'   '                           strII20 = PStr_SendMailKey1 & PStr_SendMailKey2 & PStr_SendMailKey3
'   '                        End If
'   '                     End If
'   '                     '2017/11/21 END
'                        'Modify By Sindy 2018/10/29
'                        strExc(0) = "select ii20 from " & strTableName & _
'                                    " where " & strUpdWhere
'                        intI = 1
'                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                        If intI = 1 Then
'                           If "" & RsTemp.Fields(0) = "T" Then
'                              strExc(0) = "update " & strTableName & _
'                                          " set II20='Y'" & _
'                                          " where " & strUpdWhere
'                              cnnConnection.Execute strExc(0)
'                              'Modify By Sindy 2018/10/31 移至此處更新,有SendMail成功T才更新轉寄日期
'                              strExc(0) = "update IPDeptInput set " & _
'                                          " ii08=" & strSrvDate(1) & ",ii09=to_char(sysdate, 'HH24MISS'),ii10='" & strUserNum & "',ii16=" & strSrvDate(1) & _
'                                          " where ii01=" & GRD1(intTab).TextMatrix(i, 7) & _
'                                            " and ii02=" & GRD1(intTab).TextMatrix(i, 8) & _
'                                            " and ii03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'" & _
'                                            " and ii08=0"
'                              cnnConnection.Execute strExc(0)
'                           Else
'                              strExc(0) = "update " & strTableName & _
'                                          " set II20='" & PStr_SendMailKey1 & PStr_SendMailKey2 & PStr_SendMailKey3 & "'" & _
'                                          " where " & strUpdWhere
'                              cnnConnection.Execute strExc(0)
'                           End If
'                        End If
'                        '2018/10/29 END
'
'                        'Modify By Sindy 2018/10/31 Mark
'   '                     strExc(0) = "update IPDeptInput set " & _
'   '                                 " ii08=" & strSrvDate(1) & ",ii09=to_char(sysdate, 'HH24MISS'),ii10='" & strUserNum & "',ii16=" & strSrvDate(1) & _
'   '                                 " where ii01=" & Grd1(intTab).TextMatrix(i, 7) & _
'   '                                   " and ii02=" & Grd1(intTab).TextMatrix(i, 8) & _
'   '                                   " and ii03='" & ChgSQL(Grd1(intTab).TextMatrix(i, 12)) & "'" & _
'   '                                   " and ii08=0"
'   '                     cnnConnection.Execute strExc(0)
'                        'Modify By Sindy 2018/10/29 Mark
'   '                     strExc(0) = "update IPDeptInput set " & _
'   '                                 " ii20='" & strII20 & "'" & _
'   '                                 " where ii01=" & GRD1(intTab).TextMatrix(i, 7) & _
'   '                                   " and ii02=" & GRD1(intTab).TextMatrix(i, 8) & _
'   '                                   " and ii03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'"
'   '                     cnnConnection.Execute strExc(0)
'
'         '            '直接刪除File Server Msg檔案
'         '            If PUB_DelFtpFile2(GRD1(intTab).TextMatrix(i, 7), " and ii02=" & GRD1(intTab).TextMatrix(i, 8) & " and upper(ii03)='" & ChgSQL(UCase(GRD1(intTab).TextMatrix(i, 12))) & "'", "ipdeptinput") = True Then
'         '               strExc(0) = "update IPDeptInput set " & _
'         '                           " ii14=null" & _
'         '                           " where ii01=" & GRD1(intTab).TextMatrix(i, 7) & _
'         '                             " and ii02=" & GRD1(intTab).TextMatrix(i, 8) & _
'         '                             " and ii03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'"
'         '               cnnConnection.Execute strExc(0)
'         '            End If
'                     End If
'                     '********************
'                     'Add By Sindy 2017/11/21 SendMail記錄
'                     PStr_SendMailKey1 = ""
'                     PStr_SendMailKey2 = ""
'                     PStr_SendMailKey3 = ""
'                     strTableName = ""
'                     strUpdWhere = ""
'                     '2017/11/21 END
'                     '********************
'                  End If
'               End If
'
'               cnnConnection.CommitTrans
'      '         If i <> GRD1(intTab).Rows - 1 Then
'      '            'MsgBox "是否繼續下一筆？", vbInformation, "請確認"
'      '            If MsgBox("是否繼續下一筆？", vbExclamation + vbYesNo + vbDefaultButton1, "重要訊息！") = vbNo Then
'      '               Exit For
'      '            End If
'      '         End If
'            End If
'         Next i
'      'End If
'   'Next j
'   Next intTab
   
   cmdSendMail.Enabled = True 'Add By Sindy 2016/5/10
   
   Call QueryData
   
   'Modify By Sindy 2018/10/26 信件有遺失,轉寄資訊正常,但確實寄信備份網頁系統找不到信件
'select ii08,ii09,ii20,ii21,ii22,ii17 from ipdeptinput where ii01='20181025' and ii03 in('F0292','F0304','F0293','F0262');
'/*
'      II08       II09 II20                       II21       II22 II17
'---------- ---------- -------------------- ---------- ---------- --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'  20181025     141308 Y                      20181025     141310 未傳遞的主旨: Mail Delivery Failure
'  20181026     143250 Y                      20181026     143256 Mail Delivery Failure
'  20181026     143249 Y                      20181026     143255 IMPORTANT NOTICE
'  20181026     143249 Y                      20181026     143254 Out of Office Notice
'*/
   strExc(0) = "select count(*) from ipdeptinput where ii20<>'Y' and ii20 is not null" & _
               " and ii01>=20181001" & _
               " order by ii12 asc,ii13 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp.Fields(0) > 0 Then
         PUB_SendMail strUserNum, "97038", "", "檢查信件是否有遺失(" & RsTemp.Fields(0) & "筆)", strExc(0), , , , , , , , , , , False
      End If
   End If
   '2018/10/26 END
   
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   '********************
   'Add By Sindy 2017/11/21 SendMail記錄
   PStr_SendMailKey1 = ""
   PStr_SendMailKey2 = ""
   PStr_SendMailKey3 = ""
   strTableName = ""
   strUpdWhere = ""
   '2017/11/21 END
   '********************
   cnnConnection.RollbackTrans
   
   'Add By Sindy 2017/11/21
   If Err.Number = 70 Then
      MsgBox ChgSQL(m_AttachPath & "\" & strFileName) & "檔案已開啟！", vbCritical
   Else
   '2017/11/21 END
      MsgBox " 整批轉寄失敗！" & vbCrLf & Err.Description, vbCritical
   End If
   
   'Modify By Sindy 2023/5/24
   'Call PUB_SendNotifyMail(strEMailTo, True) '寄發通知信
   '區分部門
   strF1xEmp = "": strF2xEmp = ""
   varTmp = Split(strEMailTo, ";")
   For jj = 0 To UBound(varTmp)
      If Left(PUB_GetST03(CStr(varTmp(jj))), 2) = "F1" Then '外商
         strF1xEmp = strF1xEmp & ";" & varTmp(jj)
      Else
         strF2xEmp = strF2xEmp & ";" & varTmp(jj)
      End If
   Next jj
   If strF1xEmp <> "" Then
      strF1xEmp = Mid(strF1xEmp, 2)
      Call PUB_SendNotifyMail(strF1xEmp, True) '寄發通知信
   End If
   If strF2xEmp <> "" Then
      strF2xEmp = Mid(strF2xEmp, 2)
      Call PUB_SendNotifyMail(strF2xEmp, True) '寄發通知信
   End If
   '2023/5/24 END
   Call QueryData '重新查詢
   cmdSendMail.Enabled = True 'Add By Sindy 2016/9/2
   Screen.MousePointer = vbDefault
End Sub

'呼叫新郵件
Private Sub OpenNeweMail(strTo As String, strSubject As String, _
                         strContext As String, strAttach As String)
Dim objOutLook As Object
Dim objMail As Object
Dim ArrStr As Variant
Dim jj As Integer
   
   '呼叫新郵件：
   Set objOutLook = CreateObject("Outlook.Application")
   'Set objMail = objOutLook.CreateItem(0) '新郵件
   Set objMail = objOutLook.CreateItemFromTemplate(strAttach) '原始信
   
   objMail.PrintOut '列印郵件及附件,附件本身在電腦中按滑鼠右鍵是可以列印的
   '附件
   For jj = objMail.Attachments.Count To 1 Step -1 '個數
      objMail.Attachments.Item(jj).SaveAsFile "c:\" & objMail.Attachments.Item(jj).DisplayName '另存檔案
   Next jj
   
   '移除原信的收件人及副本;密件副本不會留在msg中
   For jj = objMail.Recipients.Count To 1 Step -1
      objMail.Recipients.Remove jj
   Next jj
   
   '副本.cc
   '收件者.To
   objMail.To = strTo
   '密件副本.BCC
   
'   '主旨.Subject
'   objMail.Subject = strSubject
'
'   '加附件
'   If strAttach <> "" Then
'      ArrStr = Split(strAttach, ";")
'      For jj = 0 To UBound(ArrStr)
'         objMail.Attachments.Add ArrStr(jj)
'      Next jj
'   End If
'
'   '內文.Body
'   objMail.Body = strContext
   
   objMail.Display
'   objMail.Send
   
   Set objMail = Nothing
   Set objOutLook = Nothing
End Sub

'重新分類
Private Sub CmdSave_Click()
Dim intTab As Long
   
On Error GoTo ErrHand
   
   Screen.MousePointer = vbHourglass
   For intTab = 0 To 8 '7
      For i = 1 To GRD1(intTab).Rows - 1
         If GRD1(intTab).TextMatrix(i, 12) <> "" And GRD1(intTab).RowHeight(i) > 0 Then '有資料
'            If GRD1(intTab).TextMatrix(i, 9) <> "" Or _
'               GRD1(intTab).TextMatrix(i, 10) <> "" Then
            If GRD1(intTab).TextMatrix(i, 0) = "!" Then
               strExc(0) = "update IPDeptInput set " & _
                           " ii05='" & GRD1(intTab).TextMatrix(i, 10) & "',ii06='" & GRD1(intTab).TextMatrix(i, 9) & "'" & _
                           " where ii01=" & GRD1(intTab).TextMatrix(i, 7) & _
                             " and ii02=" & GRD1(intTab).TextMatrix(i, 8) & _
                             " and ii03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 12)) & "'"
               cnnConnection.Execute strExc(0)
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
'Dim objOutLook As Object
'Dim objMail As Object
'Dim strII03 As String, strII03_2 As String, strII11 As String, strII12 As String, strII13 As String
'Dim strUpdTime As String
'Dim stFtpPath As String
'Dim strII05 As String, strII06 As String, strII17 As String
'Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
''Dim strFileName As String
Dim oFileSys As New FileSystemObject
Dim oFolder As Folder
'Dim oFile As File
'Dim strCP09 As String
'Dim fs, f
'Dim bolSaveEFile As Boolean
'Dim lngRonCnt As Long
'Dim bolConnect As Boolean
'Dim intII03 As Integer 'Add By Sindy 2016/10/4
Dim strErrText As String
   
   '先更新資料
   If cmdSave.Enabled = True Then
      Call CmdSave_Click
   Else
      Call QueryData
   End If
   
   If txtPathIPDept = "" Then
      MsgBox "信件資料夾不可空白！"
      Exit Sub
   End If
   If Dir(txtPathIPDept, vbDirectory) = "" Then
      MkDir txtPathIPDept
   End If
   Set oFolder = oFileSys.GetFolder(txtPathIPDept.Text)
   If oFolder.files.Count = 0 Then
      MsgBox "此目錄尚無信件！"
      Set oFolder = Nothing
      Exit Sub
   End If
   
On Error GoTo ErrHand
   
'   cmdTrans.Enabled = False
'   TxtIPDept.Visible = True
'   LblCntIPDept.Visible = True
'   Set objOutLook = CreateObject("Outlook.Application")
'   Set fs = CreateObject("Scripting.FileSystemObject")
'   lngRonCnt = 0
'   For Each oFile In oFolder.files
'      lngRonCnt = lngRonCnt + 1
'      LblCntIPDept.Caption = "已處理件數 / 剩餘件數：" & lngRonCnt & " / " & oFolder.files.Count
'      DoEvents
''      TxtIPDept = oFile.Name
'      If UCase(Right(Trim(oFile.Name), 4)) = UCase(".msg") Then
''         If TxtIPDept <> oFile.Name Then
''            'Add By Sindy 2016/10/4 因為”|”ex. 主旨:Owner of TW-Patent No. 141540 (our ref: DR015-1-TW)  |  TW-Patent No. 97777 (our ref: DR014-1-TW)
''            If InStr(TxtIPDept, "|") > 0 Then
''               oFile.Name = Replace(TxtIPDept, "|", "_")
''               TxtIPDept = oFile.Name
''            End If
''            If InStr(TxtIPDept, "?") > 0 Then
''            '2016/10/4 END
''               oFile.Name = Replace(TxtIPDept, "?", "_")
''               TxtIPDept = oFile.Name
''            End If
''            If InStr(TxtIPDept, "o") > 0 Then 'o(德文上面有2點)
''            '2016/10/4 END
''               oFile.Name = Replace(TxtIPDept, "o", "_")
''               TxtIPDept = oFile.Name
''            End If
''            If InStr(TxtIPDept, "a") > 0 Then
''            '2016/10/4 END
''               oFile.Name = Replace(TxtIPDept, "a", "_") 'a(德文上面有2點)
''               TxtIPDept = oFile.Name
''            End If
''         End If
'         Call PUB_ExLetterTransTxt(oFile, TxtIPDept)
'
'         Set objMail = objOutLook.CreateItemFromTemplate(txtPathIPDept.Text & "\" & oFile.Name)
'         Screen.MousePointer = vbHourglass
'
'         'strII03 = Trim(oFile.Name)
'         strII17 = objMail.Subject
'         Text2 = strII17 'Add By Sindy 2016/4/21 Re: ML/kc 中?特許出願201510920053.X　貴所整理番?31565－CN　弊所整理番?：P-112987
'         DoEvents
''         If strII17 <> objMail.Subject Then
''            MsgBox "主旨抓的有誤，請洽電腦中心！"
''            GoTo ErrHand
''         End If
'         'If InStr(strII03, "未傳遞的主旨") = 0 And InStr(strII03, "延遲的傳遞") = 0 And Left(strII03, 3) <> "已讀取" Then
'         If objMail.Class = 46 Then '46.olReport
'            strII11 = ""
'            strII12 = "0"
'            strII13 = ""
'         '43.olMail
'         Else
'            If objMail.SenderEmailType = "EX" Then
'               strII11 = objMail.SenderName
'            Else
'               If objMail.SenderName = objMail.SenderEmailAddress Then
'                  strII11 = objMail.SenderEmailAddress
'               Else
'                  strII11 = objMail.SenderName & "[" & objMail.SenderEmailAddress & "]"
'               End If
'            End If
'            strII12 = Format(objMail.SentOn, "YYYYMMDD")
'            strII13 = Format(objMail.SentOn, "HHMMSS") 'Receivedtime
'         End If
'         'Modify By Sindy 2016/4/21 strII17-->Text2
'         'strII05 = ToSortOut(strII17, strII11, strII06, strCP01, strCP02, strCP03, strCP04)
'         strII05 = ToSortOut(Text2, strII11, strII06, strCP01, strCP02, strCP03, strCP04)
'
'         strUpdTime = Right("000000" & ServerTime, 6)
'         strCP09 = ""
'         '個案
'         If strII05 = "1" Then
'            If strCP01 <> "" And strCP02 <> "" Then
'               '該案號最大收文日最小Create日期時間的總收文號
'               strExc(0) = "select cp09 from caseprogress" & _
'                           " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
'                           " and cp05=(select max(cp05) from caseprogress" & _
'                           " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "')" & _
'                           " order by cp66 asc,cp67 asc"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  strCP09 = RsTemp.Fields("CP09")
'               End If
'            End If
'            '有總收文號:欲收錄到卷宗區所以檔名長度有限制
'            If strCP09 <> "" Then
''               If LenB(strII03) > 74 Then
''                  strII03 = LeftB(strII03, 66) & ".rx.msg" '必須取偶數,不可奇數
''               End If
''               '郵件副檔名要取為.rx.msg
''               If InStr(UCase(strII03), UCase(".rx.msg")) = 0 Then strII03 = Left(strII03, Len(strII03) - 4) & ".rx.msg"
''               'modify by sonia 2016/4/8 1.應檢查IPDeptInput,否則寫入IPDeptInput會違反唯一的限制條件,2.與DAVID討論暫不放卷宗區,以免系統直接放入沒用的信件
''               ''檢查資料庫中是否已有今天相同的檔名存在,若有,檔名再加時間
''               'strExc(0) = "select cpp02 from casepaperpdf" & _
''               '            " where cpp01=" & CNULL(strCP09) & _
''               '            " and upper(cpp02)=upper('" & ChgSQL(strII03) & "')"
''               'intI = 1
''               'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''               'If intI = 1 Then
''               '   strII03 = Trim(Left(strII03, Len(strII03) - 7)) & "," & strUpdTime & ".rx.msg"
''               '   '加了時間還是有可能重覆,再加當日當時筆數(流水號)
''               '   strExc(0) = "select cpp02 from casepaperpdf" & _
''               '               " where cpp01=" & CNULL(strCP09) & _
''               '               " and upper(cpp02)=upper('" & ChgSQL(strII03) & "')"
''               '   intI = 1
''               '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''               '   If intI = 1 Then
''               '      strExc(0) = "select count(*) from IPDeptInput" & _
''               '                  " where ii01=" & strSrvDate(1)
''               '      intI = 1
''               '      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''               '      If intI = 1 Then
''               '         strExc(0) = Val(RsTemp.Fields(0)) + 1
''               '      End If
''               '      strII03 = Trim(Left(strII03, Len(strII03) - 7)) & "," & strExc(0) & ".rx.msg"
''               '   End If
''               'End If
''               '檢查資料庫中是否已有今天相同的檔名存在,若有,檔名再加時間
''               strExc(0) = "select ii03 from IPDeptInput" & _
''                           " where ii01=" & strSrvDate(1) & _
''                           " and upper(ii03)=upper('" & ChgSQL(strII03) & "')"
''               intI = 1
''               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''               If intI = 1 Then
''                  strII03 = Trim(Left(strII03, Len(strII03) - 4)) & "," & strUpdTime & ".msg"
''                  '加了時間還是有可能重覆,再加當日當時筆數(流水號)
''                  strExc(0) = "select ii03 from IPDeptInput" & _
''                              " where ii01=" & strSrvDate(1) & _
''                              " and upper(ii03)=upper('" & ChgSQL(strII03) & "')"
''                  intI = 1
''                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''                  If intI = 1 Then
''                     strExc(0) = "select count(*) from IPDeptInput" & _
''                                 " where ii01=" & strSrvDate(1)
''                     intI = 1
''                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''                     If intI = 1 Then
''                        strExc(0) = Val(RsTemp.Fields(0)) + 1
''                     End If
''                     strII03 = Trim(Left(strII03, Len(strII03) - 4)) & "," & strExc(0) & ".msg"
''                  End If
''               End If
''               'end 2016/4/8
'            Else
'               strII05 = "Z" '其他
'            End If
'         End If
'
''         '非個案
''         If strII05 <> "1" Then
''            If LenB(strII03) > 74 Then
''               strII03 = LeftB(strII03, 70) & ".msg"
''            End If
''            '檢查資料庫中是否已有今天相同的檔名存在,若有,檔名再加時間
''            strExc(0) = "select ii03 from IPDeptInput" & _
''                        " where ii01=" & strSrvDate(1) & _
''                        " and upper(ii03)=upper('" & ChgSQL(strII03) & "')"
''            intI = 1
''            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''            If intI = 1 Then
''               strII03 = Trim(Left(strII03, Len(strII03) - 4)) & "," & strUpdTime & ".msg"
''               '加了時間還是有可能重覆,再加當日當時筆數(流水號)
''               strExc(0) = "select ii03 from IPDeptInput" & _
''                           " where ii01=" & strSrvDate(1) & _
''                           " and upper(ii03)=upper('" & ChgSQL(strII03) & "')"
''               intI = 1
''               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''               If intI = 1 Then
''                  strExc(0) = "select count(*) from IPDeptInput" & _
''                              " where ii01=" & strSrvDate(1)
''                  intI = 1
''                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''                  If intI = 1 Then
''                     strExc(0) = Val(RsTemp.Fields(0)) + 1
''                  End If
''                  strII03 = Trim(Left(strII03, Len(strII03) - 4)) & "," & strExc(0) & ".msg"
''               End If
''            End If
''         End If
'
'         cnnConnection.BeginTrans
'         bolConnect = True
'         '存實體檔案到File Server
'         '檢查若為個案必須儲存到卷宗區
'         'CANCEL BY SONIA 2016/4/8 與DAVID討論暫不放卷宗區,以免系統直接放入沒用的信件
'         'If strCP09 <> "" Then
'         '   Set f = fs.GetFile(txtPathIPDept.Text & "\" & oFile.Name)
'          '  bolSaveEFile = SaveAttFile_PDF(strCP09, txtPathIPDept.Text & "\" & oFile.Name, strII03, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), True, "F", "Y", , , stFtpPath)
'         'Else
'         'END 2016/4/8
'         '國外部信件區
'            strExc(0) = "select count(*) from IPDeptInput" & _
'                        " where ii01=" & strSrvDate(1)
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            'Modify By Sindy 2016/10/4
'            If intI = 1 Then
'               intII03 = Val(RsTemp.Fields(0)) + 1
'            Else
'               intII03 = 1
'            End If
'            strII03 = "F" & Format(intII03, "0000")
'            '2016/10/4 END
'            strII03_2 = strSrvDate(1) & strUpdTime & "." & strII03 & ".msg"
'            bolSaveEFile = PUB_PutFtpFile(txtPathIPDept.Text & "\" & oFile.Name, strSrvDate(1), strII03_2, stFtpPath, "IPDEPTINPUT")
'         'End If  'CANCEL BY SONIA 2016/4/8
'         If bolSaveEFile = True Then
'            '存資料到DB
'            'MODIFY BY SONIA 2016/4/11 不放卷宗區所以也不存總收文號,否則不同日之信件會開到同一個MSG檔,例:
'            'strSql = "insert into IPDeptInput(ii01,ii02,ii03,ii04,ii05,ii06,ii11,ii12,ii13,ii14,ii17,ii18)" & _
'                     " values(" & strSrvDate(1) & "," & strUpdTime & _
'                     ",'" & ChgSQL(strII03) & "','" & strUserNum & "'" & _
'                     ",'" & strII05 & "','" & strII06 & "'" & _
'                     "," & CNULL(ChgSQL(strII11)) & "," & CNULL(strII12) & "," & CNULL(strII13) & _
'                     ",'" & ChgSQL(stFtpPath) & "','" & ChgSQL(strII17) & "','" & strII05 & "')"
'            'Modify By Sindy 2016/4/27 寄件者長度太長,截取長度100 ex.MAILER-DAEMON@heramailgw12.hera.idc.justsystem.co.jp[MAILER-DAEMON@heramailgw12.hera.idc.justsystem.co.jp]
'            If Len(strII11) > 100 Then
'               strII11 = Mid(strII11, 1, 100)
'            End If
'            '2016/4/27 END
'            strSql = "insert into IPDeptInput(ii01,ii02,ii03,ii04,ii05,ii06,ii11,ii12,ii13,ii14,ii17,ii18)" & _
'                     " values(" & strSrvDate(1) & "," & strUpdTime & _
'                     ",'" & ChgSQL(strII03) & "','" & strUserNum & "'" & _
'                     ",'" & strII05 & "','" & strII06 & "'" & _
'                     "," & CNULL(ChgSQL(strII11)) & "," & strII12 & "," & CNULL(strII13) & _
'                     ",'" & ChgSQL(stFtpPath) & "','" & ChgSQL(strII17) & "','" & strII05 & "')"
'            cnnConnection.Execute strSql
'            '刪除PC端檔案
'            'Kill 刪不掉 "C:\IPdept\【轉知】(1) 經濟部智慧財產局來函，自105年4月1日起提出發明專利加速審查、專利審查高速公路與支援利用專利審查高速公路之專利申請案尚未公開者，不必再申請提早公開；(2) 經濟部智慧財產局來函，公告修正「發明專利加速審查申請書及其申請須知」、「發明專利PPH審查申請書及其申請須知」與「發明專利TW-SUPA審查申請書」.msg"
'            'Kill txtPathIPDept.Text & "\" & oFile.Name
'            Call fs.DeleteFile(txtPathIPDept.Text & "\" & oFile.Name)
'         End If
'         cnnConnection.CommitTrans
'         bolConnect = False
'      End If
'   Next
'   LblCntIPDept.Caption = "已處理件數 / 剩餘件數：" & lngRonCnt & " / " & oFolder.files.Count '最後再讀一次資料夾的檔案數
   
   cmdTrans.Enabled = False
   TxtIPDept.Visible = True
   LblCntIPDept.Visible = True
   'Modify By Sindy 2017/7/11
   'If PUB_IPDeptTransMail(Me, , strErrText) = False Then
   'strProFileName="N" : 手動拖拉.Msg檔至資料夾再匯入
   If PUB_IPDeptTransMail_New(Me, , strErrText, , "N") = False Then
   '2017/7/11 END
      GoTo ErrHand
   End If
   PUB_SaveLastDate Me.Name, strUserNum & "PATH", txtPathIPDept.Text
   MsgBox "信件轉入完成！" & IIf(oFolder.files.Count > 0, vbCrLf & vbCrLf & "(尚有未轉入的信件，詳情請至資料夾查看)", "")
   
   GetTodayTotCnt  'add by sonia 2016/4/1 重新計算今日總筆數
   
   cmdTrans.Enabled = True
   TxtIPDept.Visible = False
   LblCntIPDept.Visible = False
   Call QueryData
'   Screen.MousePointer = vbDefault
'   Set f = Nothing
'   Set fs = Nothing
'   Set oFolder = Nothing
'   Set oFile = Nothing
'   Set oFileSys = Nothing
'   Set objMail = Nothing
'   Set objOutLook = Nothing
   Exit Sub
   
ErrHand:
'   Screen.MousePointer = vbDefault
'   If bolConnect = True Then cnnConnection.RollbackTrans
'   If Err.Number <> 0 Then MsgBox " 信件轉入失敗！" & vbCrLf & Err.Description
   MsgBox strErrText, vbExclamation
   cmdTrans.Enabled = True
   Call QueryData
'   Set f = Nothing
'   Set fs = Nothing
'   Set oFolder = Nothing
'   Set oFile = Nothing
'   Set oFileSys = Nothing
'   Set objMail = Nothing
'   Set objOutLook = Nothing
End Sub

Private Function QueryData(Optional ByVal quyIndex As Integer = 99) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim intTab As Integer
Dim intStar As Integer
Dim dblTotCnt As Double
   
   cmdUpdRow.Enabled = False
   cmdSave.Enabled = False: cmdSave.BackColor = &H8000000F
   m_blnColOrderAsc = True
   QueryData = False
   
   Screen.MousePointer = vbHourglass
   If quyIndex = 99 Then '查全部
      intStar = 0
      quyIndex = 8 '7
   Else
      intStar = quyIndex
   End If
   For intTab = intStar To quyIndex
      GRD1(intTab).Clear
      Call SetGrd(intTab)
      'Modify By Sindy 2016/5/13 '' 刪除 ==> '' V
      'Modify By Sindy 2022/2/11 + ,getmailbox(ii01,ii03)|| : 分類前面加信箱來源和收件者信箱
      strSql = "select '' V,ii17 主旨,getmailbox(ii01,ii03)||decode(ii05," & Show國外部信件分類 & ") 分類" & _
               ",decode(nvl(st02,''),'','',st02) 收受者,sqldatet(ii12)||' '||sqltime6(ii13) 收信日期時間,ii05,ii06,ii01,ii02,'' newII06,'' newII05,ii15,ii03 檔名,ii14 FTP路徑檔名,ii18 系統記錄" & _
               " From ipdeptinput,staff" & _
               " where ii08=0 and ii05='" & IIf(intTab = 7, "Z", IIf(intTab > 7, intTab, intTab + 1)) & "'" & _
               " and ii06=st01(+)" & _
               " order by ii12 asc,ii13 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         Set GRD1(intTab).Recordset = rsTmp
         If intTab = 0 Then SSTab1.TabCaption(intTab) = "個案": LblRow(0).Visible = True: LblRow(0).Caption = rsTmp.RecordCount
         If intTab = 1 Then SSTab1.TabCaption(intTab) = "外商": LblRow(1).Visible = True: LblRow(1).Caption = rsTmp.RecordCount
         If intTab = 2 Then SSTab1.TabCaption(intTab) = "外專": LblRow(2).Visible = True: LblRow(2).Caption = rsTmp.RecordCount
         If intTab = 3 Then SSTab1.TabCaption(intTab) = "專利處": LblRow(3).Visible = True: LblRow(3).Caption = rsTmp.RecordCount
         If intTab = 4 Then SSTab1.TabCaption(intTab) = "外法": LblRow(4).Visible = True: LblRow(4).Caption = rsTmp.RecordCount
         If intTab = 5 Then SSTab1.TabCaption(intTab) = "新知": LblRow(5).Visible = True: LblRow(5).Caption = rsTmp.RecordCount
         If intTab = 6 Then SSTab1.TabCaption(intTab) = "財務": LblRow(6).Visible = True: LblRow(6).Caption = rsTmp.RecordCount
         If intTab = 7 Then SSTab1.TabCaption(intTab) = "其他": LblRow(7).Visible = True: LblRow(7).Caption = rsTmp.RecordCount
         If intTab = 8 Then SSTab1.TabCaption(intTab) = "開拓": LblRow(8).Visible = True: LblRow(8).Caption = rsTmp.RecordCount 'Add By Sindy 2016/6/15
         dblTotCnt = dblTotCnt + rsTmp.RecordCount
         QueryData = True
         '解析收受者
         For i = 1 To GRD1(intTab).Rows - 1
            GRD1(intTab).TextMatrix(i, 3) = PUB_ReadUserData(GRD1(intTab).TextMatrix(i, 6))
            'Add By Sindy 2019/11/14 主旨裡有 URGENT 字樣者,主旨前頭加入●符號
            If InStr(UCase(GRD1(intTab).TextMatrix(i, 1)), "URGENT") > 0 Then
               GRD1(intTab).TextMatrix(i, 1) = "●" & GRD1(intTab).TextMatrix(i, 1)
            End If
            '2019/11/14 END
         Next i
      Else
         If intTab = 0 Then SSTab1.TabCaption(intTab) = "個案": LblRow(0).Visible = False
         If intTab = 1 Then SSTab1.TabCaption(intTab) = "外商": LblRow(1).Visible = False
         If intTab = 2 Then SSTab1.TabCaption(intTab) = "外專": LblRow(2).Visible = False
         If intTab = 3 Then SSTab1.TabCaption(intTab) = "專利處": LblRow(3).Visible = False
         If intTab = 4 Then SSTab1.TabCaption(intTab) = "外法": LblRow(4).Visible = False
         If intTab = 5 Then SSTab1.TabCaption(intTab) = "新知": LblRow(5).Visible = False
         If intTab = 6 Then SSTab1.TabCaption(intTab) = "財務": LblRow(6).Visible = False
         If intTab = 7 Then SSTab1.TabCaption(intTab) = "其他": LblRow(7).Visible = False
         If intTab = 8 Then SSTab1.TabCaption(intTab) = "開拓": LblRow(8).Visible = False 'Add By Sindy 2016/6/15
      End If
      rsTmp.Close
      
      GRD1(intTab).Visible = False
      GRD1(intTab).col = 0
      GRD1(intTab).row = 1
      GRD1(intTab).Visible = True
   Next intTab
   If intStar <> quyIndex Then
      LblTotCnt.Caption = "總筆數: " & dblTotCnt
   End If
   
   'Add By Sindy 2017/11/15
   strExc(0) = "select mrl01 from mailreceivelog" & _
               " where mrl01='" & Left(IPDept收件匣, 2) & "'" & _
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
   SSTab1.Tab = 7 '預設在其他
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'清除單筆明細資料
Private Sub ClearDetail()
   txtII17.Text = ""
   LblII12.Caption = ""
   cboII05.Enabled = False
   cboII05.ListIndex = -1
   cboII06.Text = ""
   cboII06.Enabled = False
   cboII06.ListIndex = -1
   List1.Clear
   Frame1.Tag = "" '記錄目前那個tab
'   cboII05.Tag = ""
   List1.Tag = ""
   Me.Tag = "" '記錄grd點選那幾筆資料列
   For i = 0 To 8 '7
      dblPrevRow(i) = 0
   Next i
End Sub

'更正單筆視窗資料
Private Sub cmdUpdRow_Click()
Dim tmpArr As Variant
Dim tmpArr2 As Variant 'Add By Sindy 2024/10/15
Dim strUser As String
Dim strText As String
Dim intUpdRow As Integer 'Add By Sindy 2016/5/16
   
   'Modify By Sindy 2016/5/16 + And Me.Tag <> ""
   'If Frame1.Tag >= 0 Then
   If Val(Frame1.Tag) >= 0 And Me.Tag <> "" Then
      If GRD1(Frame1.Tag).TextMatrix(dblPrevRow(Frame1.Tag), 12) <> "" Then
         If cboII05.Text = "" Then
            MsgBox "分類不可空白！", vbExclamation
            cboII05.SetFocus
            Exit Sub
         ElseIf SSTab1.Tab <> 0 Then '非個案頁籤
            If cboII05.ListIndex = 0 Then
               MsgBox "分類不可選個案！", vbExclamation
               cboII05.SetFocus
               Exit Sub
            End If
         End If
         '+ if："其他"調整後可以再改回"其他"
         'If Not (Left(cboII05.Text, 1) = "Z" And Left(cboII05.Text, 1) = cboII05.Tag) Then
         'If Not (Left(cboII05.Text, 1) = "Z" And Frame1.Tag = "7") Then
         'END
         'Modify By Sindy 2019/4/22
         If Left(cboII05.Text, 1) <> "Z" Then
         '2019/4/22 END
            If cmdUpdRow.Enabled = True And List1.Tag = "" Then
               MsgBox "收受者不可空白！", vbExclamation
               If cboII06.Enabled = True Then
                  cboII06.SetFocus
               End If
               Exit Sub
            End If
         End If
         
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
               
               '分類
               'If Left(cboII05.Text, 1) <> cboII05.Tag Then
               If Left(cboII05.Text, 1) <> GRD1(Frame1.Tag).TextMatrix(intUpdRow, 5) Then
                  cmdSave.Enabled = True: cmdSave.BackColor = &HC0FFC0
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 0) = "!"
                  
                  'Modify By Sindy 2022/2/11 分類前面加信箱來源和收件者信箱
                  If InStrRev(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2), ")") > 0 Then
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2) = Mid(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2), 1, InStrRev(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2), ")")) & Trim(Mid(cboII05.Text, 2))
                  Else
                  '2022/2/9 END
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2) = Trim(Mid(cboII05.Text, 2))
                  End If
                  
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 10) = Left(cboII05.Text, 1)
                  If List1.Enabled = False Then
                     Select Case Left(cboII05.Text, 1)
                        Case "2" '外商
                           strUser = Pub_GetSpecMan("國外部轉信外商群組")
                        Case "3" '外專
                           strUser = Pub_GetSpecMan("國外部轉信外專群組")
                        Case "4" '專利處
                           strUser = "patent" '"patent@taie.com.tw"
                        Case "5" '外法
                           'modify by sonia 2016/4/1 改國外部轉信外法群組
                           'strUser = Pub_GetSpecMan("國外部轉信外法英文組群組") & ";99021"
                           strUser = Pub_GetSpecMan("國外部轉信外法群組") & ";" & Pub_GetSpecMan("國外部轉信外專承辦日文組長") '99021
                        Case "6" '新知:閻副所長,資訊分享區
                           strUser = Pub_GetSpecMan("國外部轉信新知群組")
                        Case "7" '財務
                           strUser = "account" '"account@taie.com.tw"
                        'Add By Sindy 2016/6/15
                        Case "8" '開拓:閻副所長
                           strUser = Pub_GetSpecMan("國外部轉信開拓群組")
                        '2016/6/15 END
                     End Select
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 3) = PUB_ReadUserData(strUser)
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 9) = strUser
                     Call SetList1(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 9))
                  End If
               '恢復原欄位值
               'ElseIf GRD1(Frame1.Tag).TextMatrix(intUpdRow, 0) = "!" Then
               ElseIf GRD1(Frame1.Tag).TextMatrix(intUpdRow, 10) <> "" Then
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 0) = ""
                  
                  'Modify By Sindy 2022/2/11 分類前面加信箱來源和收件者信箱
                  If InStrRev(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2), ")") > 0 Then
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2) = Mid(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2), 1, InStrRev(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2), ")")) & Trim(Mid(cboII05.Text, 2))
                  Else
                  '2022/2/9 END
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2) = Trim(Mid(cboII05.Text, 2))
                  End If
                  
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 3) = PUB_ReadUserData(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 6))
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 9) = ""
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 10) = ""
                  Call SetList1(GRD1(Frame1.Tag).TextMatrix(intUpdRow, 6))
               End If
               
               '收受者
               If List1.Tag = "" And List1.Enabled = True Then
                  If GRD1(Frame1.Tag).TextMatrix(intUpdRow, 6) <> "" Then
                     cmdSave.Enabled = True: cmdSave.BackColor = &HC0FFC0
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 0) = "!"
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 3) = ""
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 9) = ""
                     If GRD1(Frame1.Tag).TextMatrix(intUpdRow, 10) = "" Then
                        GRD1(Frame1.Tag).TextMatrix(intUpdRow, 10) = Left(cboII05.Text, 1)
                     End If
                  End If
               ElseIf List1.Tag <> "" Then
                  tmpArr = Split(List1.Tag, ";")
                  strUser = ""
                  For j = 0 To UBound(tmpArr)
                     If tmpArr(j) <> "" Then
                        'Add By Sindy 2023/5/24
                        strExc(0) = "select * from inputrecord" & _
                                    " where IR01=" & GRD1(Frame1.Tag).TextMatrix(intUpdRow, 7) & _
                                    " and IR03='" & GRD1(Frame1.Tag).TextMatrix(intUpdRow, 12) & "'" & _
                                    " and IR04='" & tmpArr(j) & "'"
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           Screen.MousePointer = vbDefault
                           MsgBox tmpArr(j) & " 此收受者已有收過此信件！", vbExclamation
                           cboII06.SetFocus
                           Exit Sub
                        End If
                        '2023/5/24 END
                        
                        Select Case UCase(tmpArr(j))
                           Case "外商群組"
                              strUser = strUser & IIf(strUser = "", "", ";") & Pub_GetSpecMan("國外部轉信外商群組")
                           Case "外專群組"
                              strUser = strUser & IIf(strUser = "", "", ";") & Pub_GetSpecMan("國外部轉信外專群組")
                           Case UCase("patent")
                              strUser = strUser & IIf(strUser = "", "", ";") & "patent" '"patent@taie.com.tw"
                           Case "外法群組"
                              'modify by sonia 2016/4/1 改國外部轉信外法群組
                              'strUser = strUser & IIf(strUser = "", "", ";") & Pub_GetSpecMan("國外部轉信外法英文組群組") & ";99021"
                              strUser = strUser & IIf(strUser = "", "", ";") & Pub_GetSpecMan("國外部轉信外法群組") & ";" & Pub_GetSpecMan("國外部轉信外專承辦日文組長") '99021
                           Case UCase("account")
                              strUser = strUser & IIf(strUser = "", "", ";") & "account" '"account@taie.com.tw"
                           'Modify By Sindy 2019/4/22
                           Case UCase("10F傳真機") '10F傳真機
                              strUser = strUser & IIf(strUser = "", "", ";") & "25011666" '"25011666@taie.com.tw"
                           Case "新知"
                              strUser = strUser & IIf(strUser = "", "", ";") & Pub_GetSpecMan("國外部轉信新知群組")
                           Case "代理人通知"
                              'Modify By Sindy 2021/4/6 INBOUND代理人通知增加 99033.Elvan & A4024.Widen
                              'Modify By Sindy 2025/3/7 99033;A4024 改用 國外互惠新案收文通知名單
                              '                         目的是要抓國外業拓人員,因此設定即為"國外業拓"
                              strUser = strUser & IIf(strUser = "", "", ";") & _
                                 Pub_GetSpecMan("國外部轉信外商群組") & ";" & _
                                 Pub_GetSpecMan("國外部轉信外專群組") & ";patent;" & Pub_GetSpecMan("國外互惠新案收文通知名單")
                           'Add By Sindy 2016/6/15
                           Case "開拓"
                              strUser = strUser & IIf(strUser = "", "", ";") & Pub_GetSpecMan("國外部轉信開拓群組")
                           '2016/6/15 END
                           'Add By Sindy 2020/4/29
                           Case UCase("國內信件")
                              strUser = strUser & IIf(strUser = "", "", ";") & Pub_GetSpecMan("國內信件管理人員")
                           '2020/4/29 END
                           Case Else '員工編號
                              'Add By Sindy 2024/10/15
                              If InStr(Trim(tmpArr(j)), " ") > 0 Then
                                 tmpArr2 = Split(Trim(tmpArr(j)), " ")
                                 strUser = strUser & IIf(strUser = "", "", ";") & tmpArr2(0)
                              Else
                              '2024/10/15 END
                                 strUser = strUser & IIf(strUser = "", "", ";") & Left(Trim(UCase(tmpArr(j))), 5)
                              End If
                        End Select
                     End If
                  Next j
                  '過濾是否有收受者重覆的資料
                  If strUser <> "" And InStr(strUser, ";") > 0 Then
                     strText = strUser
                     tmpArr = Split(strText, ";")
                     strUser = ""
                     For j = 0 To UBound(tmpArr)
                        If tmpArr(j) <> "" Then
                           If InStr(strUser, tmpArr(j)) = 0 Then
                              'Add By Sindy 2023/5/24
                              strExc(0) = "select * from inputrecord" & _
                                          " where IR01=" & GRD1(Frame1.Tag).TextMatrix(intUpdRow, 7) & _
                                          " and IR03='" & GRD1(Frame1.Tag).TextMatrix(intUpdRow, 12) & "'" & _
                                          " and IR04='" & tmpArr(j) & "'"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 1 Then
                                 Screen.MousePointer = vbDefault
                                 MsgBox tmpArr(j) & " 此收受者已有收過此信件！", vbExclamation
                                 cboII06.SetFocus
                                 Exit Sub
                              End If
                              '2023/5/24 END
                              strUser = strUser & IIf(strUser = "", "", ";") & tmpArr(j)
                           End If
                        End If
                     Next j
                  End If
                  cmdSave.Enabled = True: cmdSave.BackColor = &HC0FFC0
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 0) = "!"
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 3) = PUB_ReadUserData(strUser)
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 9) = strUser
                  If GRD1(Frame1.Tag).TextMatrix(intUpdRow, 10) = "" Then
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 10) = Left(cboII05.Text, 1)
                  End If
               End If
               
               Call CancelRowColor(Frame1.Tag, intUpdRow) '清除反白,並且檢查是否有更新過資料要還原
            End If
         Next i
         Me.List1.Clear 'Add By Sindy 2023/5/24
         cmdUpdRow.Enabled = False 'Add By Sindy 2016/5/17
         Screen.MousePointer = vbDefault
      End If
   End If
End Sub

Private Sub SetList1(strText As String)
Dim tmpArr As Variant
Dim strTempName As String
   
   '收受者
   tmpArr = Split(strText, ";")
   List1.Clear
   'Modify By Sindy 2019/4/22
   'List1.Tag = "" 'Add By Sindy 2016/5/16
   
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
         bolCboII06_KeyPress = False 'Add By Sindy 2021/4/14
      End If
   Next j
End Sub

Private Sub Command1_Click()
Dim objOutLook As Object
Dim objMail As Object
Dim myForward As Object
Dim jj As Integer
   
   If m_strFileName <> "" Then
      Set objOutLook = CreateObject("Outlook.Application")
      Set objMail = objOutLook.CreateItemFromTemplate(m_strFileName) 'oForm.txtPathIPDept.Text & "\" & oFile.Name
      
      '*** 轉寄 *** 會用inbound名義寄出
'目前問題是內文加文字怕會有亂碼問題
'        寄件者程式無法自動帶
      Set myForward = objMail.Forward '轉寄
      'Set myForward = objMail.Reply '回覆
      '移除原信的收件人及副本;密件副本不會留在msg中
      For jj = myForward.Recipients.Count To 1 Step -1
         myForward.Recipients.Remove jj
      Next jj
      myForward.Recipients.add "97038"
      '副本
      myForward.cc = ""
      '主旨增加,當個案且有案號時,顯示歸入那一個案號
      myForward.Subject = "RE: " & myForward.Subject & "【TEST Saved】"
      'myForward.senderemailaddress = "ipdept@taie.com.tw"
      'myForward.sentonbehalfofname = "ipdept"
      myForward.Display
      'myForward.Send
      DoEvents
      
      Set myForward = Nothing
      Set objMail = Nothing
      Set objOutLook = Nothing
      '*** END
   End If
End Sub

Private Sub Command2_Click()
'   Dim Shl As Object, Fd As Object
'   Set Shl = CreateObject("Shell.Application")
'   Set Fd = Shl.BrowseForFolder(hwnd, "請選取資料夾", 0, "C:\")
'   If Not Fd Is Nothing Then
'      txtPathIPDept.Text = Fd.Items.Item.path
'   End If
Dim stFileName As String
   
On Error GoTo ErrHnd
   
   stFileName = "*.msg"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "msg檔案 (*.msg)|*.msg"
      .InitDir = IIf(txtPathIPDept <> "", txtPathIPDept, PUB_Getdesktop)
      .MaxFileSize = 5000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         txtPathIPDept.Text = Mid(.FileName, 1, InStrRev(.FileName, "\") - 1)
      End If
   End With
   Exit Sub
ErrHnd:
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   If PUB_GetLastDate(Me.Name, strUserNum & "PATH") <> "" Then
      txtPathIPDept = PUB_GetLastDate(Me.Name, strUserNum & "PATH")
   End If
   QueryData
   
   '組合下拉選單
   '分類
   cboII05.Clear
   cboII05.AddItem "1 個案"
   cboII05.AddItem "2 外商"
   cboII05.AddItem "3 外專"
   cboII05.AddItem "4 專利處"
   cboII05.AddItem "5 外法"
   cboII05.AddItem "6 新知"
   cboII05.AddItem "7 財務"
   cboII05.AddItem "8 開拓" 'Add By Sindy 2016/6/15
   cboII05.AddItem "Z 其他"
   '收受者
   cboII06.Clear
   cboII06.AddItem ""
   cboII06.AddItem "外商群組"
   cboII06.AddItem "外專群組"
   cboII06.AddItem "patent"
   cboII06.AddItem "外法群組"
   cboII06.AddItem "新知"
   cboII06.AddItem "account 國外財務信箱"
   cboII06.AddItem "taieacc 國內財務信箱" 'Add By Sindy 2024/10/15
   cboII06.AddItem "開拓" 'Add By Sindy 2016/6/15
   cboII06.AddItem "代理人通知"
   cboII06.AddItem "TM" 'Add By Sindy 2019/2/19
   cboII06.AddItem "10F傳真機" '"25011666" 'Add By Sindy 2019/4/22 10F傳真機
   cboII06.AddItem "國內信件" 'Add By Sindy 2020/4/29 國內信件
   'Modify By Sindy 2016/4/18 特殊群組外再加外專承辦組即可,方便David分類轉寄使用,若有其他人輸員工編號按Tab鍵即可
'   strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' order by st03,st01 asc"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      With RsTemp
'         RsTemp.MoveFirst
'         Do While RsTemp.EOF = False
'            cboII06.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
'            RsTemp.MoveNext
'         Loop
'      End With
'   End If
   strSql = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' and st03='F23' order by st03,st01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With RsTemp
         RsTemp.MoveFirst
         Do While RsTemp.EOF = False
            cboII06.AddItem Trim(RsTemp.Fields("st01")) & " " & Trim(RsTemp.Fields("st02"))
            RsTemp.MoveNext
         Loop
      End With
   End If
   '2016/4/18 END
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   'Add By Sindy 2021/7/28
   '鎖定資料:顯示提示訊息
   If Pub_StrUserSt03 <> "M51" Then
      If PUB_GetLock(Me.Name, m_OldKey, Me.Caption) = False Then
      End If
   End If
   '2021/7/28 END
   
   GetTodayTotCnt  'add by sonia 2016/4/1 加入今日總筆數
'   Timer1.Interval = 100
   
   'Add By Sindy 2019/7/17
   'modify by sonia 2019/8/20 郭雅娟要求應薛經理改文字
   'LblCC.Caption = "其他信箱會加發副本給主管：Patent(" & PUB_ReadUserData(OL_PatMailCC) & ");TM(" & PUB_ReadUserData(OL_TmMailCC) & ");IPDept(" & PUB_ReadUserData(Pub_GetSpecMan("國外部信件處理人")) & ")"
   LblCC.Caption = "分信至其他部門信箱將加發副本：" & PUB_ReadUserData(OL_PatMailCC) & "(Patent);" & PUB_ReadUserData(OL_TmMailCC) & "(TM);" & PUB_ReadUserData(Pub_GetSpecMan("國外部信件處理人")) & "(IPDept)"
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Add By Sindy 2021/7/28
   '清除鎖定資料
   strSql = "Delete from LockRec where LR01='" & Me.Name & "' and LR02='" & strUserNum & "'"
   adoTaie.Execute strSql
'   If PUB_GetLock("", m_OldKey) = False Then
'      Cancel = 1
'      Exit Sub
'   End If
   '2021/7/28 END
   
   DestroyToolTip '清除物件
   Set frm06010611 = Nothing
End Sub

Private Sub SetGrd(Index As Integer)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   'Modify By Sindy 2016/10/4 + FTP路徑檔名
   '                        0    1       2       3         4               5       6       7       8       9          10         11      12      13             14
   arrGridHeadText = Array("V", "主旨", "分類", "收受者", "收信日期時間", "II05", "II06", "II01", "II02", "newII06", "newII05", "II15", "檔名", "FTP路徑檔名", "系統記錄")
   arrGridHeadWidth = Array(200, 3500, 950, 1900, 1500, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2000)
   'Modify By Sindy 2019/4/22
   'arrGridHeadWidth = Array(200, 1200, 950, 900, 900, 800, 800, 800, 800, 800, 800, 800, 800, 800)
   '2019/4/22 END
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
   GRD1(Index).row = GRD1(Index).MouseRow
   GRD1(Index).col = GRD1(Index).MouseCol
   nRow = GRD1(Index).row
   nCol = GRD1(Index).col
   'If GRD1(Index).MouseCol = 1 Then
   If GRD1(Index).col = 1 Then
      If GRD1(Index).TextMatrix(dblPrevRow(Index), 12) <> "" Then
         '讀取檔案
         Screen.MousePointer = vbHourglass
         'Modify By Sindy 2016/10/4
         'm_strFileName = GRD1(Index).TextMatrix(dblPrevRow(Index), 12)
         m_strFileName = Mid(GRD1(Index).TextMatrix(dblPrevRow(Index), 13), InStrRev(GRD1(Index).TextMatrix(dblPrevRow(Index), 13), "/") + 1)
         '2016/10/4
         Call PUB_ChkFileTypeOpenExE(m_strFileName) 'Add By Sindy 2017/9/13
         If GetAttachFile(GRD1(Index).TextMatrix(dblPrevRow(Index), 7), GRD1(Index).TextMatrix(dblPrevRow(Index), 8), GRD1(Index).TextMatrix(dblPrevRow(Index), 12), m_strFileName, "", m_AttachPath & "\" & m_strFileName) = True Then
            ShellExecute 0, "open", m_strFileName, vbNullString, vbNullString, 1
'         Else
'            MsgBox "無此郵件！", vbInformation
         End If
         Screen.MousePointer = vbDefault
      End If
   End If
End Sub

'add by sonia 2016/4/1 加入今日總筆數
Private Function GetTodayTotCnt()
   strSql = "SELECT count(*) FROM IPDeptInput WHERE ii01=" & strSrvDate(1)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      TodayTotCnt = "今日總筆數：" & "" & RsTemp.Fields(0)
   End If
End Function
'end 2016/4/1

Private Function GetAttachFile(ByVal strPkey1 As String, ByVal strPkey2 As String, ByVal strPkey3 As String, _
                               ByRef pFileName As String, ByVal strCP09 As String, _
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
   
   If Dir(stAttPath) <> "" Then Kill stAttPath 'Add By Sindy 2017/11/21
'   If strCP09 <> "" Then '個案
'      GetAttachFile = PUB_GetAttachFile_CPP(strCP09, pFileName, stAttPath, True)
'      'ADD BY SONIA 2016/4/8 因之前放入個案,故個案讀不到加入下面語法
'      If GetAttachFile = False Then
'         GetAttachFile = PUB_GetAttachFile_IImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
'      End If
'      'END 2016/4/8
'   Else
      GetAttachFile = PUB_GetAttachFile_IImsg(strPkey1, strPkey2, strPkey3, pFileName, stAttPath, True)
'   End If
   
   Exit Function
   
ErrHnd:
   GetAttachFile = False
   If Err.Number = 70 Then
      MsgBox ChgSQL(pFileName) & "檔案已開啟！", vbCritical
   Else
      MsgBox Err.Description, vbCritical
   End If
End Function

Private Sub GRD1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   'grd1(Index).ToolTipText = ""
   If GRD1(Index).MouseRow <> 0 And _
      (GRD1(Index).MouseCol = 1 Or GRD1(Index).MouseCol = 3) Then
      If iRow <> GRD1(Index).MouseRow Or iCol <> GRD1(Index).MouseCol Then
         If GRD1(Index).TextMatrix(GRD1(Index).MouseRow, GRD1(Index).MouseCol) <> "" Then
            'grd1(Index).ToolTipText = grd1(Index).TextMatrix(grd1(Index).MouseRow, grd1(Index).MouseCol)
            CreateToolTip GetHWndForToolTip(GRD1(Index)), GRD1(Index).TextMatrix(GRD1(Index).MouseRow, GRD1(Index).MouseCol)
            iRow = GRD1(Index).MouseRow
            iCol = GRD1(Index).MouseCol
         End If
      End If
   End If
End Sub

'Modify By Sindy 2016/5/13
Private Sub Grd1_Click(Index As Integer)
Dim tmpArr As Variant, strTempName As String

On Error GoTo ErrHand

cmdUpdRow.Enabled = False
cboII05.Enabled = False
cboII06.Enabled = False
List1.Enabled = False
Frame1.Tag = ""
'cboII05.Tag = ""
List1.Tag = ""

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
'If GRD1(Index).MouseRow <> 0 Then
   GRD1(Index).row = nRow 'GRD1(Index).MouseRow
   GRD1(Index).col = 0
   If GRD1(Index).TextMatrix(GRD1(Index).row, 12) <> "" Then
      '清除反白
      'If GRD1(Index).TextMatrix(GRD1(Index).row, 0) = "V" Then
      If GRD1(Index).CellBackColor = &HFFC0C0 Then
'         GRD1(index).TextMatrix(GRD1(index).row, 0) = ""
'         GRD1(index).col = 0
'         For i = 0 To GRD1(index).Cols - 1
'            GRD1(index).col = i
'            GRD1(index).CellBackColor = QBColor(15)
'         Next i
         Call CancelRowColor(Index, GRD1(Index).row) '清除反白,並且檢查是否有更新過資料要還原
         If dblPrevRow(Index) = GRD1(Index).row Then
            dblPrevRow(Index) = 0
            '重新預設目前筆數
            If Me.Tag <> "" Then
               tmpArr = Split(Me.Tag, ",")
               dblPrevRow(Index) = tmpArr(UBound(tmpArr))
            'Add By Sindy 2016/9/14
            Else
               Call ClearDetail
            '2016/9/14 END
            End If
         End If
'         If dblPrevCnt(Index) > 0 Then
'            dblPrevCnt(Index) = Val(dblPrevCnt(Index)) - 1
'         End If
         'MsgBox "dblPrevRow(Index)=" & dblPrevRow(Index) & vbCrLf & "dblPrevCnt(Index)=" & dblPrevCnt(Index) & vbCrLf & "Me.Tag=" & Me.Tag
      Else
         '將點選資料列反白
         GRD1(Index).TextMatrix(GRD1(Index).row, 0) = "V"
         GRD1(Index).col = 0
         GRD1(Index).row = nRow
         For i = 0 To GRD1(Index).Cols - 1
            GRD1(Index).col = i
            GRD1(Index).CellBackColor = &HFFC0C0
         Next i
'         If Val(dblPrevCnt(Index)) = 0 Then
'            Me.Tag = "" '還原值
'         End If
         dblPrevRow(Index) = GRD1(Index).row '記錄目前筆數
         Me.Tag = Me.Tag & "," & dblPrevRow(Index)
'         dblPrevCnt(Index) = Val(dblPrevCnt(Index)) + 1
         'MsgBox "dblPrevRow(Index)=" & dblPrevRow(Index) & vbCrLf & "dblPrevCnt(Index)=" & dblPrevCnt(Index) & vbCrLf & "Me.Tag=" & Me.Tag
      End If
      
      '**********************************************************************
      '顯示明細資料
      '**********************************************************************
      If Val(dblPrevRow(Index)) > 0 Then
         '主旨
         txtII17.Text = GRD1(Index).TextMatrix(dblPrevRow(Index), 1)
         txtII17.SetFocus 'Add by Sindy 2021/4/13 Form2.0才會顯示出捲抽
         '收受者
         cboII06.ListIndex = 0
         If GRD1(Index).TextMatrix(dblPrevRow(Index), 9) <> "" Then
            tmpArr = Split(GRD1(Index).TextMatrix(dblPrevRow(Index), 9), ";")
         Else
            tmpArr = Split(GRD1(Index).TextMatrix(dblPrevRow(Index), 6), ";")
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
               Else
                  List1.AddItem tmpArr(j)
               End If
               bolCboII06_KeyPress = False 'Add By Sindy 2021/4/14
            End If
         Next j
         '分類
'         cboII05.Tag = GRD1(Index).TextMatrix(dblPrevRow(Index), 5)
         If InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "個案") > 0 Then
            cboII05.ListIndex = 0
         ElseIf InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "外商") > 0 Then
            cboII05.ListIndex = 1
         ElseIf InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "外專") > 0 Then
            cboII05.ListIndex = 2
         ElseIf InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "專利處") > 0 Then
            cboII05.ListIndex = 3
         ElseIf InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "外法") > 0 Then
            cboII05.ListIndex = 4
         ElseIf InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "新知") > 0 Then
            cboII05.ListIndex = 5
         ElseIf InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "財務") > 0 Then
            cboII05.ListIndex = 6
         ElseIf InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "其他") > 0 Then
            cboII05.ListIndex = 8
         'Add By Sindy 2016/6/15
         ElseIf InStr(GRD1(Index).TextMatrix(dblPrevRow(Index), 2), "開拓") > 0 Then
            cboII05.ListIndex = 7
         '2016/6/15 END
         End If
         '收信日期時間
         LblII12.Caption = GRD1(Index).TextMatrix(dblPrevRow(Index), 4)
         '設定
         Frame1.Tag = Index '記錄那一個GRD1
         cboII05.Enabled = False
         cboII06.Enabled = False
         List1.Enabled = False
         If Index = 0 Then '個案
         'ElseIf Index = 7 Or UBound(Split(Me.Tag, ",")) > 1 Then '其他
         ElseIf Index = 7 Then  '其他
            cmdUpdRow.Enabled = True
            cboII05.Enabled = True
            cboII06.Enabled = True
            List1.Enabled = True
         Else
            'Add By Sindy 2016/5/17
            If UBound(Split(Me.Tag, ",")) > 0 Then
               cmdUpdRow.Enabled = True
            End If
            '2016/5/17 END
            cboII05.Enabled = True
         End If
      End If
      '**********************************************************************
   End If
End If
GRD1(Index).Visible = True

Exit Sub

ErrHand:
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

'Add By Sindy 2016/5/13
Private Sub SSTab1_Click(PreviousTab As Integer)
Dim tmpArr As Variant
   
   If SSTab1.Tag <> "" And PreviousTab <> SSTab1.Tab Then
      If MsgBox("您已勾選資料尚未處理，" & vbCrLf & vbCrLf & _
                "確定要放棄處理嗎？", vbInformation + vbYesNo + vbDefaultButton2, "警示詢問") = vbYes Then
         tmpArr = Split(Me.Tag, ",")
         For i = 1 To UBound(tmpArr)
            If Val(tmpArr(i)) > 0 Then
               Call CancelRowColor(PreviousTab, Val(tmpArr(i))) '清除反白,並且檢查是否有更新過資料要還原
            End If
         Next i
         Call ClearDetail '清除單筆明細資料
         Exit Sub
      Else
         SSTab1.Tag = ""
         SSTab1.Tab = PreviousTab
         Exit Sub
      End If
   End If
End Sub
Private Sub SSTab1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   SSTab1.Tag = Me.Tag
End Sub

''Add By Sindy 2017/11/15
'Private Sub Timer1_Timer()
'   strExc(0) = "select mrl01 from mailreceivelog" & _
'               " where mrl01='" & Left(IPDept收件匣, 2) & "'" & _
'               " and mrl09='A'"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      cmdHandRecv.BackColor = &HC0FFC0
'      MsgBox "正在等待信件接件！", vbInformation
'      Timer1.Interval = 100
'   Else
'      cmdHandRecv.BackColor = &H8000000F
'      Timer1.Interval = 0
'   End If
'End Sub
