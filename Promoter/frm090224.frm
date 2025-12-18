VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm090224 
   BorderStyle     =   1  '單線固定
   Caption         =   "商標處收件夾信件處理"
   ClientHeight    =   6732
   ClientLeft      =   4080
   ClientTop       =   2160
   ClientWidth     =   8952
   ControlBox      =   0   'False
   Icon            =   "frm090224.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6732
   ScaleWidth      =   8952
   Begin VB.CommandButton cmdQuery 
      Caption         =   "畫面更新(&Q)"
      Height          =   330
      Left            =   4020
      TabIndex        =   39
      Top             =   0
      Width           =   1275
   End
   Begin VB.CommandButton cmdHistory 
      Caption         =   "記錄查詢"
      Height          =   330
      Left            =   7200
      TabIndex        =   2
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdSendMail 
      Caption         =   "整批轉寄"
      Height          =   330
      Left            =   6270
      TabIndex        =   1
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "重新分類"
      Enabled         =   0   'False
      Height          =   330
      Left            =   5340
      Style           =   1  '圖片外觀
      TabIndex        =   0
      Top             =   0
      Width           =   885
   End
   Begin VB.CommandButton cmdExit 
      Cancel          =   -1  'True
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   330
      Left            =   8115
      TabIndex        =   3
      Top             =   0
      Width           =   800
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   240
      Index           =   4
      Left            =   6000
      TabIndex        =   49
      Top             =   1935
      Value           =   1  '核取
      Width           =   195
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   240
      Index           =   3
      Left            =   4530
      TabIndex        =   48
      Top             =   1935
      Value           =   1  '核取
      Width           =   195
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   240
      Index           =   2
      Left            =   3060
      TabIndex        =   47
      Top             =   1935
      Value           =   1  '核取
      Width           =   195
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   240
      Index           =   1
      Left            =   1650
      TabIndex        =   46
      Top             =   1935
      Value           =   1  '核取
      Width           =   195
   End
   Begin VB.CheckBox ChkTab 
      Caption         =   "Check1"
      Height          =   240
      Index           =   0
      Left            =   150
      TabIndex        =   45
      Top             =   1935
      Value           =   1  '核取
      Width           =   195
   End
   Begin VB.CommandButton cmdRecOutlookQ 
      Caption         =   "郵件接收狀況"
      Height          =   270
      Left            =   7680
      TabIndex        =   40
      Top             =   1680
      Width           =   1245
   End
   Begin VB.CommandButton cmdUpdRow 
      BackColor       =   &H00C0E0FF&
      Caption         =   "更正"
      Height          =   270
      Left            =   5400
      Style           =   1  '圖片外觀
      TabIndex        =   28
      Top             =   1650
      Width           =   675
   End
   Begin VB.CommandButton cmdDelRow 
      BackColor       =   &H00C0E0FF&
      Caption         =   "刪除"
      Height          =   270
      Left            =   4680
      Style           =   1  '圖片外觀
      TabIndex        =   27
      Top             =   1650
      Width           =   675
   End
   Begin VB.CommandButton Command2 
      Caption         =   "<="
      Height          =   255
      Left            =   8100
      TabIndex        =   6
      Top             =   180
      Visible         =   0   'False
      Width           =   345
   End
   Begin VB.TextBox txtPathPatent 
      Height          =   270
      Left            =   4980
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "C:\TM"
      Top             =   150
      Visible         =   0   'False
      Width           =   3105
   End
   Begin VB.Frame Frame1 
      Caption         =   "資料修改區"
      ForeColor       =   &H00000080&
      Height          =   1275
      Left            =   90
      TabIndex        =   7
      Top             =   390
      Width           =   8835
      Begin VB.TextBox txtTi21 
         Height          =   270
         Left            =   6750
         MaxLength       =   2
         TabIndex        =   19
         Top             =   960
         Width           =   375
      End
      Begin VB.TextBox txtTi19 
         Height          =   270
         Left            =   5655
         MaxLength       =   6
         TabIndex        =   17
         Top             =   960
         Width           =   855
      End
      Begin VB.TextBox txtTi18 
         Height          =   270
         Left            =   5160
         MaxLength       =   3
         TabIndex        =   16
         Top             =   960
         Width           =   495
      End
      Begin VB.TextBox txtTi20 
         Height          =   270
         Left            =   6510
         MaxLength       =   1
         TabIndex        =   18
         Top             =   960
         Width           =   255
      End
      Begin VB.ComboBox cboTi05 
         Height          =   300
         ItemData        =   "frm090224.frx":000C
         Left            =   5610
         List            =   "frm090224.frx":000E
         Style           =   2  '單純下拉式
         TabIndex        =   13
         Top             =   660
         Width           =   1530
      End
      Begin MSForms.ListBox List1 
         Height          =   975
         Left            =   7170
         TabIndex        =   12
         Top             =   240
         Width           =   1635
         ScrollBars      =   3
         DisplayStyle    =   2
         Size            =   "2884;1561"
         MatchEntry      =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   165
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.ComboBox cboTi06 
         Height          =   285
         Left            =   5610
         TabIndex        =   11
         Top             =   360
         Width           =   1515
         VariousPropertyBits=   679495707
         DisplayStyle    =   3
         Size            =   "2672;503"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         Value           =   "cboTi06"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtTI17 
         Height          =   705
         Left            =   570
         TabIndex        =   56
         Top             =   210
         Width           =   4365
         VariousPropertyBits=   -1399830505
         BackColor       =   -2147483633
         ScrollBars      =   3
         Size            =   "7699;1244"
         Value           =   "txtTI17"
         SpecialEffect   =   0
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label7 
         Caption         =   "本所案號:"
         Height          =   225
         Left            =   4350
         TabIndex        =   36
         Top             =   990
         Width           =   795
      End
      Begin VB.Label Label1 
         Caption         =   "(點二下可移除資料)"
         ForeColor       =   &H000000C0&
         Height          =   225
         Index           =   1
         Left            =   7230
         TabIndex        =   25
         Top             =   60
         Width           =   1575
      End
      Begin VB.Label LblTi12 
         Caption         =   "Label7"
         Height          =   225
         Left            =   1260
         TabIndex        =   15
         Top             =   990
         Width           =   3015
      End
      Begin VB.Label Label6 
         Caption         =   "收信日期時間:"
         Height          =   165
         Left            =   90
         TabIndex        =   14
         Top             =   990
         Width           =   1155
      End
      Begin VB.Label Label5 
         Caption         =   "收受者:"
         Height          =   255
         Left            =   4980
         TabIndex        =   10
         Top             =   390
         Width           =   615
      End
      Begin VB.Label Label4 
         Caption         =   "分　類:"
         Height          =   255
         Left            =   4980
         TabIndex        =   9
         Top             =   690
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "主旨:"
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   240
         Width           =   435
      End
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   3915
      Left            =   60
      TabIndex        =   30
      Top             =   2160
      Width           =   8835
      _ExtentX        =   15579
      _ExtentY        =   6900
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   8.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "MCTF"
      TabPicture(0)   =   "frm090224.frx":0010
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "GRD1(0)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "大陸案"
      TabPicture(1)   =   "frm090224.frx":002C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "GRD1(1)"
      Tab(1).ControlCount=   1
      TabCaption(2)   =   "個人"
      TabPicture(2)   =   "frm090224.frx":0048
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "GRD1(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "非大陸案"
      TabPicture(3)   =   "frm090224.frx":0064
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "GRD1(3)"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "其他"
      TabPicture(4)   =   "frm090224.frx":0080
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "GRD1(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "其他信箱"
      TabPicture(5)   =   "frm090224.frx":009C
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "GRD1(5)"
      Tab(5).ControlCount=   1
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
         Bindings        =   "frm090224.frx":00B8
         Height          =   3495
         Index           =   0
         Left            =   60
         TabIndex        =   31
         Top             =   360
         Width           =   8685
         _ExtentX        =   15325
         _ExtentY        =   6160
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.6
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
         Bindings        =   "frm090224.frx":00CD
         Height          =   3495
         Index           =   1
         Left            =   -74940
         TabIndex        =   32
         Top             =   360
         Width           =   8685
         _ExtentX        =   15325
         _ExtentY        =   6160
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.6
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
         Bindings        =   "frm090224.frx":00E2
         Height          =   3500
         Index           =   2
         Left            =   -74940
         TabIndex        =   33
         Top             =   360
         Width           =   8685
         _ExtentX        =   15325
         _ExtentY        =   6160
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.6
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
         Bindings        =   "frm090224.frx":00F7
         Height          =   3500
         Index           =   3
         Left            =   -74940
         TabIndex        =   34
         Top             =   360
         Width           =   8685
         _ExtentX        =   15325
         _ExtentY        =   6160
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.6
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
         Bindings        =   "frm090224.frx":010C
         Height          =   3495
         Index           =   4
         Left            =   -74940
         TabIndex        =   35
         Top             =   360
         Width           =   8685
         _ExtentX        =   15325
         _ExtentY        =   6160
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.6
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
         Bindings        =   "frm090224.frx":0121
         Height          =   3495
         Index           =   5
         Left            =   -74940
         TabIndex        =   52
         Top             =   360
         Width           =   8685
         _ExtentX        =   15325
         _ExtentY        =   6160
         _Version        =   393216
         Cols            =   6
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         AllowUserResizing=   3
         FormatString    =   "V|主旨|分類|本所案號|收受者|收信日期時間"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "新細明體-ExtB"
            Size            =   9.6
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
      TabIndex        =   38
      Text            =   "信件分類作業中，請稍候..."
      Top             =   1250
      Visible         =   0   'False
      Width           =   7005
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD2 
      Bindings        =   "frm090224.frx":0136
      Height          =   1875
      Left            =   120
      TabIndex        =   43
      Top             =   7140
      Width           =   2805
      _ExtentX        =   4953
      _ExtentY        =   3302
      _Version        =   393216
      Rows            =   0
      FixedRows       =   0
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin SHDocVwCtl.WebBrowser WebBrowserT 
      CausesValidation=   0   'False
      Height          =   1725
      Left            =   2970
      TabIndex        =   44
      TabStop         =   0   'False
      Top             =   7140
      Width           =   1605
      ExtentX         =   2831
      ExtentY         =   3043
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
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   8460
      Top             =   90
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSForms.TextBox txtTi11 
      Height          =   300
      Left            =   2460
      TabIndex        =   57
      Top             =   0
      Visible         =   0   'False
      Width           =   1365
      VariousPropertyBits=   746604575
      Size            =   "2408;529"
      Value           =   "txtTi11"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      Caption         =   "信件資料夾："
      Height          =   195
      Left            =   3900
      TabIndex        =   4
      Top             =   210
      Visible         =   0   'False
      Width           =   1065
   End
   Begin VB.Label LblCC 
      AutoSize        =   -1  'True
      Caption         =   $"frm090224.frx":014B
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
      Height          =   360
      Left            =   90
      TabIndex        =   54
      Top             =   6360
      Width           =   8820
   End
   Begin VB.Label Label1 
      Caption         =   "備註：雙擊”主旨”開啟信件"
      ForeColor       =   &H000000C0&
      Height          =   195
      Index           =   0
      Left            =   3600
      TabIndex        =   53
      Top             =   6120
      Width           =   2535
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
      Left            =   255
      TabIndex        =   51
      Top             =   6120
      Visible         =   0   'False
      Width           =   1755
   End
   Begin VB.Label LblReceiver 
      AutoSize        =   -1  'True
      BackColor       =   &H0080FFFF&
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
      Height          =   180
      Left            =   2040
      TabIndex        =   50
      Top             =   6120
      Visible         =   0   'False
      Width           =   1020
   End
   Begin MSForms.TextBox TextBoxT 
      Height          =   345
      Left            =   120
      TabIndex        =   42
      Top             =   6750
      Width           =   2025
      VariousPropertyBits=   -1400879077
      ScrollBars      =   2
      Size            =   "3572;609"
      Value           =   "Find簡體字"
      FontName        =   "Arial Unicode MS"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox TextBox3 
      Height          =   345
      Left            =   2190
      TabIndex        =   41
      Top             =   6750
      Width           =   2025
      VariousPropertyBits=   -1400879077
      ScrollBars      =   2
      Size            =   "3572;609"
      Value           =   "Find簡體字"
      FontName        =   "Arial Unicode MS"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblRow 
      Alignment       =   2  '置中對齊
      BackColor       =   &H0080FF80&
      Caption         =   "999"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Index           =   5
      Left            =   7470
      TabIndex        =   37
      Top             =   1980
      Visible         =   0   'False
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
      TabIndex        =   24
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
      Left            =   1890
      TabIndex        =   23
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
      Left            =   3300
      TabIndex        =   22
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
      Left            =   4770
      TabIndex        =   21
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
      Left            =   6210
      TabIndex        =   20
      Top             =   1980
      Width           =   345
   End
   Begin VB.Label LblTotCnt 
      AutoSize        =   -1  'True
      Caption         =   "總筆數(1~8頁籤):"
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   255
      TabIndex        =   29
      Top             =   1710
      Width           =   1335
   End
   Begin VB.Label TodayTotCnt 
      AutoSize        =   -1  'True
      Caption         =   "今日總筆數："
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   600
      TabIndex        =   26
      Top             =   80
      Width           =   1080
   End
   Begin MSForms.TextBox TextII17 
      Height          =   300
      Left            =   0
      TabIndex        =   55
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
Attribute VB_Name = "frm090224"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/4/14 Form2.0已修改
'Create By Sindy 2019/4/18
Option Explicit

'(執行各項功能的權限)
Dim m_bInsert As Boolean
Dim m_bUpdate As Boolean
Dim m_bDelete As Boolean
Dim m_bQuery As Boolean
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim dblPrevRow(0 To 5) As Double '記錄目前點選那一筆
Dim m_AttachPath As String
Dim nCol As Long, nRow As Long
Dim m_OldKey As String
Dim m_strUserList As String
Dim bolCboTi06_KeyPress As Boolean 'Add By Sindy 2021/4/14


'分類
Private Sub cboTi05_Click()
Dim strUser As String
   
   If cboTi05.ListIndex >= 0 Then
      List1.Clear: List1.Tag = ""
'      Select Case Left(cboTi05.Text, 1)
'         Case "1" 'P程序1
'            strUser = Pub_GetSpecMan("專利處轉信非台灣程序1")
'         Case "2" 'P程序2
'            strUser = Pub_GetSpecMan("專利處轉信非台灣程序2")
'         'Modify By Sindy 2018/6/21
''         Case "3" '亞洲
''            strUser = Pub_GetSpecMan("專利處轉信亞洲程序")
''         Case "4" '歐洲
''            strUser = Pub_GetSpecMan("專利處轉信歐洲程序")
''         Case "5" '美洋非(單)
''            strUser = Pub_GetSpecMan("專利處轉信美洋非洲單號程序")
''         Case "6" '美洋非(雙)
''            strUser = Pub_GetSpecMan("專利處轉信美洋非洲雙號程序")
'         Case "3" '美日(單)
'            strUser = Pub_GetSpecMan("專利處轉信美日單號程序")
'         Case "4" '美日(雙)
'            strUser = Pub_GetSpecMan("專利處轉信美日雙號程序")
'         Case "5" '美日外(單)
'            strUser = Pub_GetSpecMan("專利處轉信美日以外單號程序")
'         Case "6" '美日外(雙)
'            strUser = Pub_GetSpecMan("專利處轉信美日以外雙號程序")
'         '2018/6/21 END
'      End Select
'      If strUser <> "" Then
'         cboTi06.Text = ""
'         List1.Tag = strUser & " " & GetPrjSalesNM(strUser)
'         List1.AddItem strUser & " " & GetPrjSalesNM(strUser)
'      End If
   End If
End Sub

'Add By Sindy 2020/8/25
Private Sub ChkSetEmp()
Dim strEmp As String
Dim tmpArr As Variant
Dim j As Integer
   
   'Modify By Sindy 2021/4/13 + And cboTi06.Tag <> cboTi06.Text
   If cboTi06.Text <> "" And cboTi06.Tag <> cboTi06.Text Then
      'Modify By Sindy 2020/9/15
      'strEmp = Trim(Left(cboTi06.Text, 6))
      If InStr(cboTi06.Text, " ") > 0 Then
         strEmp = Mid(Trim(cboTi06.Text), 1, InStr(cboTi06.Text, " ") - 1)
      ElseIf InStr(cboTi06.Text, "@") > 0 Then
         If InStr(UCase(cboTi06.Text), "@TAIE.COM.TW") > 0 Then
            strEmp = Mid(Trim(cboTi06.Text), 1, InStr(cboTi06.Text, "@") - 1)
         Else
            strEmp = Trim(cboTi06.Text)
         End If
      Else
         'strEmp = Trim(Left(cboTi06.Text, 6))
         strEmp = Trim(cboTi06.Text)
      End If
      '2020/9/15 END
      cboTi06.Tag = cboTi06.Text 'Add By Sindy 2021/4/13
      
      Call PUB_TM_ToSortOut_sub(strEmp, True)
      tmpArr = Split(strEmp, ";")
      For j = 0 To UBound(tmpArr)
         cboTi06.Text = Trim(tmpArr(j))
         'Call cboTi06_LostFocus
         Call GetCboTi06_StaffNm
         If List1.ListCount = 0 Then List1.Clear: List1.Tag = ""
         If InStr(List1.Tag, cboTi06.Text) = 0 Then
            List1.AddItem cboTi06.Text
            List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboTi06.Text
         End If
      Next
      bolCboTi06_KeyPress = False 'Add By Sindy 2021/4/14
      If cboTi06.Enabled = True Then
         cboTi06.Text = ""
      End If
   End If
End Sub

'收受者
Private Sub cboTi06_Click()
   If bolCboTi06_KeyPress = True Then Exit Sub 'Add By Sindy 2021/4/14
   If cboTi06.ListIndex >= 0 Then
      Call ChkSetEmp
'      If InStr(List1.Tag, cboTi06.List(cboTi06.ListIndex)) = 0 Then
'         If Trim(List1.Tag) = "" Then List1.Clear
'         List1.AddItem cboTi06.List(cboTi06.ListIndex)
'         List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboTi06.List(cboTi06.ListIndex)
'         cboTi06.Text = ""
'      End If
   End If
'   If cboTi06.Enabled = False Then
'      cboTi06.Text = ""
'   End If
End Sub

Private Sub cboTi06_Validate(Cancel As Boolean)
Dim strEmp As String
Dim tmpArr As Variant
Dim j As Integer
   
   Call ChkSetEmp
'   If cboTi06.Text <> "" Then
'      Call PUB_TM_ToSortOut_sub(strEmp, True)
'      tmpArr = Split(strEmp, ";")
'      For j = 0 To UBound(tmpArr)
'         cboTi06.Text = Trim(tmpArr(j))
'         Call cboTi06_LostFocus
'         If List1.ListCount = 0 Then List1.Clear: List1.Tag = ""
'         If InStr(List1.Tag, cboTi06.Text) = 0 Then
'            List1.AddItem cboTi06.Text
'            List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboTi06.Text
'         End If
'      Next
'      cboTi06.Text = ""
'''      '檢查人員是否存在或離職
'''      If ChkStaffST04(Left(cboTi06, 5)) = True Then
'''         cboTi06.SetFocus
'''         Call cboTi06_GotFocus
'''         Exit Sub
'''      End If
''      'If Len(Trim(cboTi06.Text)) = 5 Then
''         'cboTi06.Text = Left(cboTi06.Text, 5) & " " & GetStaffName(Left(cboTi06.Text, 5), True)
''         If List1.ListCount = 0 Then List1.Clear: List1.Tag = ""
''         If InStr(List1.Tag, cboTi06.Text) = 0 Then
''            List1.AddItem cboTi06.Text
''            List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboTi06.Text
''         End If
''         cboTi06.Text = ""
''         'cboTi06.SetFocus 'Add By Sindy 2018/4/24
''      'End If
'   End If
End Sub

Private Sub cboTi06_GotFocus()
   cboTi06.SelStart = 0
   cboTi06.SelLength = Len(cboTi06.Text)
End Sub
Private Sub cboTi06_KeyPress(KeyAscTi As MSForms.ReturnInteger)
   bolCboTi06_KeyPress = True 'Add By Sindy 2021/4/14
   KeyAscTi = UpperCase(KeyAscTi)
End Sub
Private Sub cboTi06_LostFocus()
   Call GetCboTi06_StaffNm
End Sub

Private Sub GetCboTi06_StaffNm()
Dim strText As String
Dim bolFind As Boolean, ii As Integer
   
   If cboTi06.Text <> "" Then
      '依員工姓名抓取員工編號
      strText = GetPrjSalesNM_2(cboTi06.Text)
      If strText <> "" Then
         cboTi06.Text = strText & " " & cboTi06.Text
      Else
         '依員工編號抓取員工姓名
         strText = GetPrjSalesNM(Left(cboTi06.Text, 5))
         If strText <> "" Then
            'Add By Sindy 2021/4/14
            '檢查人員是否離職
            If ChkStaffST04(Left(cboTi06.Text, 5)) = True Then
               cboTi06.SetFocus
               Call cboTi06_GotFocus
               cboTi06.Text = ""
               Exit Sub
            Else
            '2021/4/14 END
               cboTi06.Text = Left(cboTi06.Text, 5) & " " & strText
            End If
         Else
            'Add By Sindy 2021/4/14 檢查是否有在List清單裡, 沒有則不可加入
            bolFind = False
            For ii = 0 To cboTi06.ListCount - 1
               'If cboTi06.Text = cboTi06.List(ii) Then
               If InStr(cboTi06.List(ii), cboTi06.Text) > 0 Then
                  cboTi06.Text = cboTi06.List(ii)
                  bolFind = True: Exit For
               End If
            Next ii
            If bolFind = False Then
               cboTi06.Text = ""
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

Private Sub ChkTab_Click(Index As Integer)
   If Val(LblRow(Index)) > 0 Then
      ChkTab(Index).Visible = True
   Else
      ChkTab(Index).Visible = False
   End If
End Sub

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
         strExc(0) = "update TMInput set " & _
                     " Ti07='Y',Ti08=" & strSrvDate(1) & ",Ti09=" & Right("000000" & ServerTime, 6) & ",Ti10='" & strUserNum & "',Ti16=" & strSrvDate(1) & _
                     " where Ti01=" & GRD1(SSTab1.Tab).TextMatrix(i, 8) & _
                       " and Ti02=" & GRD1(SSTab1.Tab).TextMatrix(i, 9) & _
                       " and Ti03='" & ChgSQL(GRD1(SSTab1.Tab).TextMatrix(i, 13)) & "'"
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
         LblTotCnt.Caption = "總筆數(1~5頁籤): " & Val(Replace(LblTotCnt.Caption, "總筆數(1~5頁籤):", "")) - 1
         strExc(0) = LblRow(SSTab1.Tab).Caption
         strExc(0) = Val(strExc(0)) - 1
         If SSTab1.Tab = 0 Then LblRow(0).Caption = strExc(0): ChkTab_Click (0) '台灣案
         If SSTab1.Tab = 1 Then LblRow(1).Caption = strExc(0): ChkTab_Click (1) '非台灣案
         If SSTab1.Tab = 2 Then LblRow(2).Caption = strExc(0): ChkTab_Click (2) 'MCTF
         If SSTab1.Tab = 3 Then LblRow(3).Caption = strExc(0): ChkTab_Click (3) '個人
         If SSTab1.Tab = 4 Then LblRow(4).Caption = strExc(0): ChkTab_Click (4) '其他
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
   frm06010615.m_QueryType = "T"
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
Dim strRecordRow As String '記錄處理到那一筆資料
Dim bolReadFile As Boolean
Dim i As Integer, j As Integer
Dim strTi11 As String, strTi12 As String, strTi13 As String, strTi17 As String
Dim intPI03 As String, strPI03 As String, strPI03_2 As String
Dim stFtpPath As String, bolSaveEFile As Boolean
Dim strToCC As String 'Add By Sindy 2019/7/17
Dim strSender As String 'Add By Sindy 2020/8/24
   
   If m_bInsert = False Then cmdSendMail.Enabled = False: Exit Sub
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
   For intTab = 0 To 4 'MCTF~其他
      If ChkTab(intTab).Value = 1 And ChkTab(intTab).Visible = True Then
      For i = 1 To GRD1(intTab).Rows - 1
         If GRD1(intTab).TextMatrix(i, 13) <> "" And _
            GRD1(intTab).TextMatrix(i, 7) <> "" Then '有檔名流水號有收受者時,就更新資料
            strRecordRow = SSTab1.TabCaption(intTab) & ":" & GRD1(intTab).TextMatrix(i, 8) & "-" & GRD1(intTab).TextMatrix(i, 9) & "-" & GRD1(intTab).TextMatrix(i, 13) 'Modify By Sindy 2018/1/2
            'Add By Sindy 2020/8/24
            strSender = GRD1(intTab).TextMatrix(i, 7)
            Call PUB_TM_ToSortOut_sub(strSender)
            GRD1(intTab).TextMatrix(i, 7) = strSender
            '2020/8/24 END
            '逐一檢查收受者
            tmpArr = Split(GRD1(intTab).TextMatrix(i, 7), ";")
            bolReadFile = False
            For j = 0 To UBound(tmpArr)
               If tmpArr(j) <> "" Then
                  'Add By Sindy 2020/11/4
                  If InStr(OL_TmMail需排除的收受者, tmpArr(j)) > 0 Then
                     GoTo ReadNext
                  End If
                  '2020/11/4 END
                  
                  'Add By Sindy 2018/4/10 讀不到ST03值的收受者資料可能有問題,不處理
'                  UCase(tmpArr(j)) <> "PATENT" And _
'                  UCase(tmpArr(j)) <> "IPDEPT" And
                  If PUB_GetST03(CStr(tmpArr(j))) = "" And _
                     Len(tmpArr(j)) = 5 And _
                     InStr(tmpArr(j), "@") = 0 Then
                     If PUB_GetST03(CStr(tmpArr(j))) = "" Then
                        GoTo ReadNext
                     End If
                  End If
                  '2018/4/10 END
                  '收受者 非商標處人員 或 @ 或 Patent 時,先下載檔案
'                  If Left(PUB_GetST03(CStr(tmpArr(j))), 2) <> "P2" Or _
'                     UCase(tmpArr(j)) = "PATENT" Or _
'                     InStr(tmpArr(j), "@") > 0 Then
                  If Left(PUB_GetST03(CStr(tmpArr(j))), 2) <> "P2" Then
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
                  strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR15)" & _
                              " values(" & GRD1(intTab).TextMatrix(i, 8) & _
                                       "," & GRD1(intTab).TextMatrix(i, 9) & _
                                       ",'" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)) & "'" & _
                                       ",'" & tmpArr(j) & "'," & strSrvDate(1) & "," & _
                                       strUpdTime & ",'" & strUserNum & "','Y')"
                  cnnConnection.Execute strExc(0)
                  
                  If UCase(tmpArr(j)) = "PATENT" Then
                     strRecordRow = strRecordRow & "[匯入專利處收件夾]"
                     '讀取商標處收件夾資料
                     strExc(0) = "select Ti11,Ti12,Ti13,Ti17 from TMInput" & _
                                 " where Ti01=" & GRD1(intTab).TextMatrix(i, 8) & _
                                 " and Ti02=" & GRD1(intTab).TextMatrix(i, 9) & _
                                 " and Ti03='" & GRD1(intTab).TextMatrix(i, 13) & "'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strTi11 = "" & RsTemp.Fields("Ti11")
                        strTi12 = "" & RsTemp.Fields("Ti12")
                        strTi13 = "" & RsTemp.Fields("Ti13")
                        strTi17 = "" & RsTemp.Fields("Ti17")
                     End If
                     
                     '讀取專利處收件夾資料
'                     strExc(0) = "select count(*) from PatentInput" & _
'                                 " where PI01=" & strSrvDate(1)
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        intPI03 = Val(RsTemp.Fields(0)) + 1
'                     Else
'                        intPI03 = 1
'                     End If
'                     strPI03 = "P" & Format(intPI03, "0000")
                     'Modify By Sindy 2019/12/2 自動給號,才能 Keep PKey
                     strPI03 = AutoNoByDate("P", 4)
                     '2019/12/2 END
                     strPI03_2 = strSrvDate(1) & strUpdTime & "." & strPI03 & ".msg"
                     '存實體檔案到PatentInput
                     bolSaveEFile = PUB_PutFtpFile(strFileName, strSrvDate(1), strPI03_2, stFtpPath, UCase("PatentInput"))
                     If bolSaveEFile = True Then
                        '新增資料到專利處收件夾資料
                        strSql = "insert into PatentInput(PI01,PI02,PI03,PI04,PI05,PI11,PI12,PI13,PI14,PI17,PI15)" & _
                                 " values(" & strSrvDate(1) & "," & strUpdTime & _
                                 ",'" & strPI03 & "','" & strUserNum & "',null" & _
                                 "," & CNULL(ChgSQL(strTi11)) & "," & strTi12 & "," & CNULL(strTi13) & _
                                 ",'" & ChgSQL(stFtpPath) & "','" & ChgSQL(strTi17) & "','TM')"
                        cnnConnection.Execute strSql
                        'Add By Sindy 2019/7/5 增加副本給專利處主管
                        If OL_SendNotifyMailCC("TM", "Patent", strFileName, strTi17, strSrvDate(1), strUpdTime, strPI03, OL_PatMailCC, strSrvDate(1), strUpdTime) = False Then
                           GoTo ErrHand
                        End If
                        
                        '更新商標處收件夾資料
                        strExc(0) = "update TMInput set " & _
                                    " Ti10='" & strUserNum & "',Ti22='" & strPI03 & "'" & _
                                    " where Ti01=" & GRD1(intTab).TextMatrix(i, 8) & _
                                      " and Ti02=" & GRD1(intTab).TextMatrix(i, 9) & _
                                      " and Ti03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)) & "'" & _
                                      " and Ti08=0"
                        cnnConnection.Execute strExc(0)
                        '該收受者上刪除日期時間人員
                        strExc(0) = "update InputRecord set " & _
                                    " ir08=" & strSrvDate(1) & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
                                    " where ir01=" & GRD1(intTab).TextMatrix(i, 8) & _
                                      " and ir02=" & GRD1(intTab).TextMatrix(i, 9) & _
                                      " and ir03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)) & "'" & _
                                      " and ir04='" & tmpArr(j) & "'"
                        cnnConnection.Execute strExc(0)
                        
'                        GRD1(intTab).TextMatrix(i, 7) = Replace(GRD1(intTab).TextMatrix(i, 7), "patent", "") '收受者拿掉patent
'                        GRD1(intTab).TextMatrix(i, 7) = Replace(GRD1(intTab).TextMatrix(i, 7), ";;", ";")
'                        If GRD1(intTab).TextMatrix(i, 7) = ";" Then GRD1(intTab).TextMatrix(i, 7) = ""
'                        If GRD1(intTab).TextMatrix(i, 7) <> "" Then
'                           If Left(GRD1(intTab).TextMatrix(i, 7), 1) = ";" Then GRD1(intTab).TextMatrix(i, 7) = Mid(GRD1(intTab).TextMatrix(i, 7), 2)
'                           If Right(GRD1(intTab).TextMatrix(i, 7), 1) = ";" Then GRD1(intTab).TextMatrix(i, 7) = Mid(GRD1(intTab).TextMatrix(i, 7), 1, Len(GRD1(intTab).TextMatrix(i, 7)) - 1)
'                        End If
                     End If
                     
                  '轉寄Outlook並且該收受者上刪除日期時間人員
                  'ElseIf Left(PUB_GetST03(CStr(tmpArr(j))), 2) <> "P2" Or InStr(tmpArr(j), "@") > 0 Then
                  ElseIf Left(PUB_GetST03(CStr(tmpArr(j))), 2) <> "P2" Then
                     strRecordRow = strRecordRow & "[寄信]"
                     
                     'Add By Sindy 2019/7/17 增加副本給國外部信件處理人
                     If UCase(tmpArr(j)) = "IPDEPT" Then
                        If strToCC <> "" Then strToCC = strToCC & ";"
                        strToCC = Pub_GetSpecMan("國外部信件處理人")
                     End If
                     '2019/7/17 END
                     
                     'Add By Sindy 2022/2/7
                     If strTo <> "" Then strTo = strTo & ";"
                     strTo = strTo & Trim(tmpArr(j))
                     '2022/2/7 END
                                          
'                     '還是先維持用個人名義轉寄信件
'                     PUB_SendMail strUserNum, tmpArr(j), "", GRD1(intTab).TextMatrix(i, 1), vbCrLf & "信件內容參附件！", , strFileName, , , , strToCC, , , , , False
'                     'PUB_SendMail strUserNum, tmpArr(j), "", GRD1(intTab).TextMatrix(i, 1), vbCrLf & "信件內容參附件！", , strFileName, , , , , "patent@taie.com.tw", , , , False
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
            
            strExc(0) = "update TMInput set " & _
                        " Ti08=" & strSrvDate(1) & ",Ti09=" & strUpdTime & ",Ti10='" & strUserNum & "'" & _
                        " where Ti01=" & GRD1(intTab).TextMatrix(i, 8) & _
                          " and Ti02=" & GRD1(intTab).TextMatrix(i, 9) & _
                          " and Ti03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)) & "'"
            cnnConnection.Execute strExc(0)
            Call SaveTMInput(GRD1(intTab).TextMatrix(i, 8), GRD1(intTab).TextMatrix(i, 9), GRD1(intTab).TextMatrix(i, 13))
            cnnConnection.CommitTrans
            
'            '刪除PC端檔案
'            'Call fs.DeleteFile(m_AttachPath & "\" & strFileName)
'            Kill strFileName
            
            '串要發通知信的人員
            tmpArr = Split(GRD1(intTab).TextMatrix(i, 7), ";")
            For j = 0 To UBound(tmpArr)
               If tmpArr(j) <> "" Then
                  If Left(PUB_GetST03(CStr(tmpArr(j))), 2) = "P2" Then '商標處
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
   Call PUB_SendMailCache 'Add By Sindy 2019/7/17
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

Private Sub SaveTMInput(strTi01 As String, strTi02 As String, strTi03 As String)
   '若信件收受者全部已處理或已刪除,主檔就可以掛上msg檔刪除日期,等待AutoBatchDay一個月後刪除實體檔
   strExc(0) = "select ir01 from InputRecord" & _
               " where ir01=" & strTi01 & _
                 " and ir02=" & strTi02 & _
                 " and ir03='" & strTi03 & "'" & _
                 " and ir08=0" 'and ir05=0 and ir08=0 : 若信件收受者全部已讀取或已刪除
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 0 Then
      '更新"無"Msg檔刪除日期
      strExc(0) = "update TMInput set" & _
                  " Ti16=" & strSrvDate(1) & _
                  " where Ti01=" & strTi01 & _
                    " and Ti02=" & strTi02 & _
                    " and Ti03='" & strTi03 & "'" & _
                    " and Ti16=0"
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
   If Err.Number = 70 Then
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
   
   If m_bInsert = False Then cmdSave.Enabled = False: Exit Sub
   Screen.MousePointer = vbHourglass
   For intTab = 0 To 4
      For i = 1 To GRD1(intTab).Rows - 1
         If GRD1(intTab).TextMatrix(i, 13) <> "" And GRD1(intTab).RowHeight(i) > 0 Then '有資料
            If GRD1(intTab).TextMatrix(i, 0) = "!" Then
               strExc(0) = "update TMInput set" & _
                           " Ti05='" & GRD1(intTab).TextMatrix(i, 11) & "',Ti06='" & GRD1(intTab).TextMatrix(i, 10) & "'"
               If GRD1(intTab).TextMatrix(i, 19) <> "" Then
                  strExc(0) = strExc(0) & _
                           ",Ti18='" & GRD1(intTab).TextMatrix(i, 19) & "',Ti19='" & GRD1(intTab).TextMatrix(i, 20) & "'" & _
                           ",Ti20='" & GRD1(intTab).TextMatrix(i, 21) & "',Ti21='" & GRD1(intTab).TextMatrix(i, 22) & "'"
               End If
               strExc(0) = strExc(0) & _
                           " where Ti01=" & GRD1(intTab).TextMatrix(i, 8) & _
                             " and Ti02=" & GRD1(intTab).TextMatrix(i, 9) & _
                             " and Ti03='" & ChgSQL(GRD1(intTab).TextMatrix(i, 13)) & "'"
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

Private Function QueryData(Optional ByVal quyIndex As Integer = 99) As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim intTab As Integer
Dim intStar As Integer
Dim dblTotCnt As Double
Dim i As Integer
   
   cmdSendMail.Enabled = True
   cmdUpdRow.Enabled = False
   cmdSave.Enabled = False: cmdSave.BackColor = &H8000000F
   m_blnColOrderAsc = True
   QueryData = False
   
   Screen.MousePointer = vbHourglass
   '代表匯入未分信
   strSql = "SELECT count(*) FROM TMInput WHERE Ti05 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If RsTemp.Fields(0) > 0 Then
         txtShowTransMsg.Top = 1830
         txtShowTransMsg.ZOrder '移至頂層
         txtShowTransMsg.Visible = True
         DoEvents '*****
         Call PUB_IPDeptChangeTM(Me)
         txtShowTransMsg.Visible = False
         DoEvents '*****
      End If
   End If
   
   If quyIndex = 99 Then '查全部
      intStar = 0
      quyIndex = 5
   Else
      intStar = quyIndex
   End If
   For intTab = intStar To quyIndex
      GRD1(intTab).Clear
      Call SetGrd(intTab)
      strSql = "select '' V,Ti17 主旨,decode(Ti05," & Show商標處信件分類 & ") 分類,decode(Ti18,null,'',Ti18||'-'||Ti19||'-'||Ti20||'-'||Ti21) 本所案號" & _
               ",decode(nvl(st02,''),'','',st02) 收受者,sqldatet(Ti12)||' '||sqltime6(Ti13) 收信日期時間,Ti05,Ti06,Ti01,Ti02,'' newTi06,'' newTi05,Ti15 系統記錄,Ti03,Ti14 FTP路徑檔名" & _
               ",Ti18,Ti19,Ti20,Ti21,'' newTi18,'' newTi19,'' newTi20,'' newTi21,Ti11" & _
               " From TMInput,staff" & _
               " where Ti08=0 and Ti05='" & (intTab + 1) & "'" & _
               " and Ti06=st01(+)" & _
               " order by Ti12 desc,Ti13 desc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         Set GRD1(intTab).Recordset = rsTmp
         If intTab = 0 Then SSTab1.TabCaption(intTab) = "MCTF": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
         If intTab = 1 Then SSTab1.TabCaption(intTab) = "大陸案": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
         If intTab = 2 Then SSTab1.TabCaption(intTab) = "個人": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
         If intTab = 3 Then SSTab1.TabCaption(intTab) = "非大陸案": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
         If intTab = 4 Then SSTab1.TabCaption(intTab) = "其他": LblRow(intTab).Visible = True: LblRow(intTab).Caption = rsTmp.RecordCount: Call ChkTab_Click(intTab)
         If intTab <> 5 Then dblTotCnt = dblTotCnt + rsTmp.RecordCount
         QueryData = True
         '解析收受者
         For i = 1 To GRD1(intTab).Rows - 1
            GRD1(intTab).TextMatrix(i, 4) = PUB_ReadUserData(GRD1(intTab).TextMatrix(i, 7))
         Next i
      Else
         If intTab = 0 Then SSTab1.TabCaption(intTab) = "MCTF": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
         If intTab = 1 Then SSTab1.TabCaption(intTab) = "大陸案": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
         If intTab = 2 Then SSTab1.TabCaption(intTab) = "個人": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
         If intTab = 3 Then SSTab1.TabCaption(intTab) = "非大陸案": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
         If intTab = 4 Then SSTab1.TabCaption(intTab) = "其他": LblRow(intTab).Caption = "": Call ChkTab_Click(intTab) ': LblRow(intTab).Visible = False
      End If
      rsTmp.Close
      
      GRD1(intTab).Visible = False
      GRD1(intTab).col = 0
      GRD1(intTab).row = 1
      GRD1(intTab).Visible = True
   Next intTab
   If intStar <> quyIndex Then
      LblTotCnt.Caption = "總筆數(1~5頁籤): " & dblTotCnt
   End If
   
   ClearDetail
   SSTab1.Tab = 4 '預設在其他
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Function

'清除單筆明細資料
Private Sub ClearDetail()
Dim i As Integer
   
   txtTI17.Text = ""
   txtTi11.Text = ""
   LblTi12.Caption = ""
   cboTi05.Enabled = False
   cboTi05.ListIndex = -1
   cboTi06.Enabled = False
   cboTi06.ListIndex = -1
   cboTi06.Text = ""
   txtTi18.Text = ""
   txtTi19.Text = ""
   txtTi20.Text = ""
   txtTi21.Text = ""
   cboTi06.Tag = "" 'Add By Sindy 2021/4/13
   List1.Clear
   List1.Tag = ""
   Frame1.Tag = "" '記錄目前那個tab
   Me.Tag = "" '記錄grd點選那幾筆資料列
   For i = 0 To 5
      dblPrevRow(i) = 0
   Next i
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
   
   'Frame1.Tag : 第幾個頁籤
   'Me.Tag : GRD筆數
   If Val(Frame1.Tag) >= 0 And Me.Tag <> "" Then
      If GRD1(Frame1.Tag).TextMatrix(dblPrevRow(Frame1.Tag), 13) <> "" Then '有資料
         If cboTi05.Text = "" Then
            MsgBox "分類不可空白！", vbExclamation
            cboTi05.SetFocus
            Exit Sub
         End If
         If Left(cboTi05, 1) <> 5 Then '其他不須控制即時輸入收受者
            If cmdUpdRow.Enabled = True And List1.Tag = "" Then
               MsgBox "收受者不可空白！", vbExclamation
               cboTi06.SetFocus
               Exit Sub
            End If
'         Else
'            If List1.Tag = "" And (txtTi18 = "" Or txtTi19 = "") Then
'               MsgBox "收受者 或 本所案號不可空白！", vbExclamation
'               txtTi18.SetFocus
'               Exit Sub
'            End If
         End If
         
'         bolCancl = False
'         Call txtTi21_Validate(bolCancl)
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
               '分類
               'If Left(cboTi05.Text, 1) <> GRD1(Frame1.Tag).TextMatrix(intUpdRow, 6) Then
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 2) = Trim(Mid(cboTi05.Text, 2))
                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 11) = Left(cboTi05.Text, 1)
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
               '本所案號
               'Modify By Sindy 2017/12/21
'               If SelManyRow = True Then '選取多筆
'                  grd1(Frame1.Tag).TextMatrix(intUpdRow, 19) = grd1(Frame1.Tag).TextMatrix(intUpdRow, 15)
'                  grd1(Frame1.Tag).TextMatrix(intUpdRow, 20) = grd1(Frame1.Tag).TextMatrix(intUpdRow, 16)
'                  grd1(Frame1.Tag).TextMatrix(intUpdRow, 21) = grd1(Frame1.Tag).TextMatrix(intUpdRow, 17)
'                  grd1(Frame1.Tag).TextMatrix(intUpdRow, 22) = grd1(Frame1.Tag).TextMatrix(intUpdRow, 18)
'               Else
               '2017/12/21 END
               If SelManyRow = False Then '單筆時
'                  If txtTi18 <> grd1(Frame1.Tag).TextMatrix(intUpdRow, 15) Or _
'                     txtTi19 <> grd1(Frame1.Tag).TextMatrix(intUpdRow, 16) Or _
'                     txtTi20 <> grd1(Frame1.Tag).TextMatrix(intUpdRow, 17) Or _
'                     txtTi21 <> grd1(Frame1.Tag).TextMatrix(intUpdRow, 18) Then
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 19) = txtTi18
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 20) = txtTi19
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 21) = txtTi20
                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 22) = txtTi21
                     If txtTi18 = "" And txtTi19 = "" And txtTi20 = "" And txtTi21 = "" Then
                        GRD1(Frame1.Tag).TextMatrix(intUpdRow, 3) = ""
                     Else
                        GRD1(Frame1.Tag).TextMatrix(intUpdRow, 3) = txtTi18 & "-" & txtTi19 & "-" & txtTi20 & "-" & txtTi21
                     End If
'                  ElseIf grd1(Frame1.Tag).TextMatrix(intUpdRow, 19) <> grd1(Frame1.Tag).TextMatrix(intUpdRow, 15) Or _
'                     grd1(Frame1.Tag).TextMatrix(intUpdRow, 20) <> grd1(Frame1.Tag).TextMatrix(intUpdRow, 16) Or _
'                     grd1(Frame1.Tag).TextMatrix(intUpdRow, 21) <> grd1(Frame1.Tag).TextMatrix(intUpdRow, 17) Or _
'                     grd1(Frame1.Tag).TextMatrix(intUpdRow, 22) <> grd1(Frame1.Tag).TextMatrix(intUpdRow, 18) Then
'                     grd1(Frame1.Tag).TextMatrix(intUpdRow, 19) = grd1(Frame1.Tag).TextMatrix(intUpdRow, 15)
'                     grd1(Frame1.Tag).TextMatrix(intUpdRow, 20) = grd1(Frame1.Tag).TextMatrix(intUpdRow, 16)
'                     grd1(Frame1.Tag).TextMatrix(intUpdRow, 21) = grd1(Frame1.Tag).TextMatrix(intUpdRow, 17)
'                     grd1(Frame1.Tag).TextMatrix(intUpdRow, 22) = grd1(Frame1.Tag).TextMatrix(intUpdRow, 18)
'                  End If
               End If
'               If grd1(Frame1.Tag).TextMatrix(intUpdRow, 19) = "" And _
'                  grd1(Frame1.Tag).TextMatrix(intUpdRow, 20) = "" And _
'                  grd1(Frame1.Tag).TextMatrix(intUpdRow, 21) = "" And _
'                  grd1(Frame1.Tag).TextMatrix(intUpdRow, 22) = "" Then
'                  grd1(Frame1.Tag).TextMatrix(intUpdRow, 3) = ""
'               Else
'                  grd1(Frame1.Tag).TextMatrix(intUpdRow, 3) = grd1(Frame1.Tag).TextMatrix(intUpdRow, 19) & "-" & grd1(Frame1.Tag).TextMatrix(intUpdRow, 20) & "-" & grd1(Frame1.Tag).TextMatrix(intUpdRow, 21) & "-" & grd1(Frame1.Tag).TextMatrix(intUpdRow, 22)
'               End If
               
'               '收受者
'               If List1.Tag = "" And List1.Enabled = True Then
'                  If GRD1(Frame1.Tag).TextMatrix(intUpdRow, 7) <> "" Then
'                     cmdSave.Enabled = True: cmdSave.BackColor = &HC0FFC0
'                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 0) = "!"
'                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 4) = ""
'                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 10) = ""
'                     If GRD1(Frame1.Tag).TextMatrix(intUpdRow, 11) = "" Then
'                        GRD1(Frame1.Tag).TextMatrix(intUpdRow, 11) = Left(cboTi05.Text, 1)
'                     End If
'                  End If
'               ElseIf List1.Tag <> "" Then
'                  TmpArr = Split(List1.Tag, ";")
'                  strUser = ""
'                  For j = 0 To UBound(TmpArr)
'                     If TmpArr(j) <> "" Then
'                        strUser = strUser & IIf(strUser = "", "", ";") & Left(Trim(UCase(TmpArr(j))), 5)
'                     End If
'                  Next j
'                  '過濾是否有收受者重覆的資料
'                  If strUser <> "" And InStr(strUser, ";") > 0 Then
'                     strText = strUser
'                     TmpArr = Split(strText, ";")
'                     strUser = ""
'                     For j = 0 To UBound(TmpArr)
'                        If TmpArr(j) <> "" Then
'                           If InStr(strUser, TmpArr(j)) = 0 Then
'                              strUser = strUser & IIf(strUser = "", "", ";") & TmpArr(j)
'                           End If
'                        End If
'                     Next j
'                  End If
'                  cmdSave.Enabled = True: cmdSave.BackColor = &HC0FFC0
'                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 0) = "!"
'                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 4) = PUB_ReadUserData(strUser)
'                  GRD1(Frame1.Tag).TextMatrix(intUpdRow, 10) = strUser
'                  If GRD1(Frame1.Tag).TextMatrix(intUpdRow, 11) = "" Then
'                     GRD1(Frame1.Tag).TextMatrix(intUpdRow, 11) = Left(cboTi05.Text, 1)
'                  End If
'               End If
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
         bolCboTi06_KeyPress = False 'Add By Sindy 2021/4/14
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
   If Err.Number <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   m_bInsert = IsUserHasRightOfFunction(Me.Name, strAdd, False)
   m_bUpdate = IsUserHasRightOfFunction(Me.Name, strEdit, False)
   m_bDelete = IsUserHasRightOfFunction(Me.Name, strDel, False)
   m_bQuery = IsUserHasRightOfFunction(Me.Name, strFind, False)
   
   ReDim pa(1 To TF_PA) As String
   ReDim sp(1 To tf_SP) As String
   
   If PUB_GetLastDate(Me.Name, strUserNum & "PATH") <> "" Then
      txtPathPatent = PUB_GetLastDate(Me.Name, strUserNum & "PATH")
   End If
   
   '組合下拉選單
   '分類
   cboTi05.Clear
   cboTi05.AddItem "1 MCTF"
   cboTi05.AddItem "2 大陸案"
   cboTi05.AddItem "3 個人"
   cboTi05.AddItem "4 非大陸案"
   cboTi05.AddItem "5 其他"
   '收受者
   cboTi06.Clear
   cboTi06.AddItem "": m_strUserList = ""
   '商標處人員
   m_strUserList = PUB_AddComboTMailEmp(Left(Pub_GetSpecMan("商標處信件處理人"), 5), cboTi06, True)
   cboTi06.AddItem "A2008 " & GetPrjSalesNM("A2008") 'Add By Sindy 2020/6/5 + 婉莘,請款時要寄的
   'Add By Sindy 2021/10/25 因林經理即將退休,往後巨京管理的相關信件,需分信給江協理麻煩在收受者名單裡增加江協理 (98020), 以利分信
   cboTi06.AddItem "98020 " & GetPrjSalesNM("98020")
   cboTi06.AddItem "ipdept"
   cboTi06.AddItem "patent"
   cboTi06.AddItem "Jerry_lin" 'Add By Sindy 2020/9/15
   cboTi06.AddItem "account" 'Add By Sindy 2020/9/15
   cboTi06.AddItem "bd" 'Add By Sindy 2022/5/10
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
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
   End If
   
   If Dir(App.path & "\executeTM.txt") <> "" Then
      WebBrowserT.Navigate App.path & "\executeTM.txt"
      DoEvents
      TextBoxT = Replace(Replace(WebBrowserT.Document.Body.innerhtml, "<PRE>", ""), "</PRE>", "")
   Else
      TextBoxT = ""
   End If
   
   SSTab1.TabVisible(5) = False
   
   QueryData
   GetTodayTotCnt '今日總筆數
   
   'Add By Sindy 2019/7/17
   'modify by sonia 2019/8/20 郭雅娟要求應薛經理改文字
   'LblCC.Caption = "其他信箱會加發副本給主管：Patent(" & PUB_ReadUserData(OL_PatMailCC) & ");TM(" & PUB_ReadUserData(OL_TmMailCC) & ");IPDept(" & PUB_ReadUserData(Pub_GetSpecMan("國外部信件處理人")) & ")"
   LblCC.Caption = "分信至其他部門信箱將加發副本：" & PUB_ReadUserData(OL_PatMailCC) & "(Patent);" & PUB_ReadUserData(OL_TmMailCC) & "(TM);" & PUB_ReadUserData(Pub_GetSpecMan("國外部信件處理人")) & "(IPDept)"
   
   'Modify By Sindy 2020/7/7 Mark
'   If m_bInsert = False Then
'      cmdSave.Enabled = False
'      cmdSendMail.Enabled = False
'   Else
'      cmdSave.Enabled = True
'      cmdSendMail.Enabled = True
'   End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   DestroyToolTip '清除物件
End Sub

Private Sub Form_Unload(Cancel As Integer)
   '清除鎖定資料
   strSql = "Delete from LockRec where LR01='" & Me.Name & "' and LR02='" & strUserNum & "'"
   adoTaie.Execute strSql
   
   DestroyToolTip '清除物件
   Set frm090224 = Nothing
End Sub

Private Sub SetGrd(Index As Integer)
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   '                        0    1       2       3           4         5               6       7       8       9       10         11         12          13      14             15      16      17      18      19         20         21         22         23
   arrGridHeadText = Array("V", "主旨", "分類", "本所案號", "收受者", "收信日期時間", "Ti05", "Ti06", "Ti01", "Ti02", "newTi06", "newTi05", "系統記錄", "Ti03", "FTP路徑檔名", "Ti18", "Ti19", "Ti20", "Ti21", "newTi18", "newTi19", "newTi20", "newTi21", "Ti11")
   arrGridHeadWidth = Array(200, 3500, 950, 1200, 2000, 1500, 0, 0, 0, 0, 0, 0, 900, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0)
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
   If nCol = 1 And nRow > 0 Then '雙擊”主旨”開啟信件
      dblPrevRow(Index) = nRow
      If GRD1(Index).TextMatrix(dblPrevRow(Index), 14) <> "" Then
         '讀取檔案
         Screen.MousePointer = vbHourglass
         strFileName = Mid(GRD1(Index).TextMatrix(dblPrevRow(Index), 14), InStrRev(GRD1(Index).TextMatrix(dblPrevRow(Index), 14), "\") + 1)
         strFileName = Mid(strFileName, InStrRev(strFileName, "/") + 1)
         Call PUB_ChkFileTypeOpenExE(strFileName)
         If GetAttachFile(GRD1(Index).TextMatrix(dblPrevRow(Index), 8), GRD1(Index).TextMatrix(dblPrevRow(Index), 9), GRD1(Index).TextMatrix(dblPrevRow(Index), 13), strFileName, m_AttachPath & "\" & strFileName) = True Then
            ShellExecute 0, "open", strFileName, vbNullString, vbNullString, 1
         End If
         Screen.MousePointer = vbDefault
      End If
   End If
End Sub

'今日總筆數
Private Function GetTodayTotCnt()
   strSql = "SELECT count(*) FROM TMInput WHERE Ti01=" & strSrvDate(1)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      TodayTotCnt = "今日總筆數：" & "" & RsTemp.Fields(0)
   End If
End Function

Private Sub GRD1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Static iRow As Integer, iCol As Integer
   
   'grd1(Index).ToolTipText = ""
   If GRD1(Index).MouseRow <> 0 And _
      (GRD1(Index).MouseCol = 1 Or GRD1(Index).MouseCol = 4 Or GRD1(Index).MouseCol = 12) Then
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

Private Sub Grd1_Click(Index As Integer)
Dim tmpArr As Variant, strTempName As String
Dim strKeep As String
Dim rsTmp As New ADODB.Recordset
Dim i As Integer, j As Integer

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
         txtTI17.Text = GRD1(Index).TextMatrix(dblPrevRow(Index), 1)
         txtTI17.SetFocus 'Add by Sindy 2021/4/13 Form2.0才會顯示出捲抽
         '收信日期時間
         LblTi12.Caption = GRD1(Index).TextMatrix(dblPrevRow(Index), 5)
         '寄件者
         txtTi11.Text = GRD1(Index).TextMatrix(dblPrevRow(Index), 23)
         
         '分類
         If GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "MCTF" Then
            cboTi05.ListIndex = 0
         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "大陸案" Then
            cboTi05.ListIndex = 1
         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "個人" Then
            cboTi05.ListIndex = 2
         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "非大陸案" Then
            cboTi05.ListIndex = 3
         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 2) = "其他" Then
            cboTi05.ListIndex = 4
         Else
            cboTi05.ListIndex = -1
         End If
         '本所案號
         txtTi18 = ""
         txtTi19 = ""
         txtTi20 = ""
         txtTi21 = ""
         If GRD1(Index).TextMatrix(dblPrevRow(Index), 19) <> "" Or strKeep = "!" Then
            txtTi18 = GRD1(Index).TextMatrix(dblPrevRow(Index), 19)
            txtTi19 = GRD1(Index).TextMatrix(dblPrevRow(Index), 20)
            txtTi20 = GRD1(Index).TextMatrix(dblPrevRow(Index), 21)
            txtTi21 = GRD1(Index).TextMatrix(dblPrevRow(Index), 22)
         ElseIf GRD1(Index).TextMatrix(dblPrevRow(Index), 15) <> "" Then
            txtTi18 = GRD1(Index).TextMatrix(dblPrevRow(Index), 15)
            txtTi19 = GRD1(Index).TextMatrix(dblPrevRow(Index), 16)
            txtTi20 = GRD1(Index).TextMatrix(dblPrevRow(Index), 17)
            txtTi21 = GRD1(Index).TextMatrix(dblPrevRow(Index), 18)
         End If
         
         '收受者
         cboTi06.Tag = "" 'Add By Sindy 2021/4/13
         List1.Tag = "" 'Add By Sindy 2021/4/13
         cboTi06.ListIndex = -1
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
                  If InStr(List1.Tag, tmpArr(j) & " " & strTempName) = 0 Then List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & tmpArr(j) & " " & strTempName
                  'cboTi06.Text = tmpArr(j) & " " & strTempName
               Else
                  List1.AddItem tmpArr(j)
                  If InStr(List1.Tag, tmpArr(j) & " " & strTempName) = 0 Then List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & tmpArr(j) & " " & strTempName
                  'cboTi06.Text = tmpArr(j)
               End If
               bolCboTi06_KeyPress = False 'Add By Sindy 2021/4/14
               'If InStr(List1.Tag, cboTi06.Text) = 0 Then List1.Tag = List1.Tag & IIf(List1.Tag = "", "", ";") & cboTi06.Text
               'cboTi06.Text = ""
            End If
         Next j
         
         '設定
         cmdUpdRow.Enabled = False
         Frame1.Tag = Index '記錄那一個GRD1
         cboTi05.Enabled = True
         cboTi06.Enabled = True
         List1.Enabled = True
         txtTi18.Enabled = True
         txtTi19.Enabled = True
         txtTi20.Enabled = True
         txtTi21.Enabled = True
         If Index = 5 Then '其他信箱匯入
            cboTi05.Enabled = False
            cboTi06.Enabled = False
            List1.Enabled = False
            txtTi18.Enabled = False
            txtTi19.Enabled = False
            txtTi20.Enabled = False
            txtTi21.Enabled = False
         Else
            If UBound(Split(Me.Tag, ",")) > 0 Then
               cmdUpdRow.Enabled = True
            End If
         End If
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
   End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
Dim tmpArr As Variant
Dim i As Integer
   
   If SSTab1.Enabled = False Then Exit Sub
   cmdDelRow.Enabled = True
   cmdUpdRow.Enabled = True
   If SSTab1.Tag <> "" And PreviousTab <> SSTab1.Tab Then
      If PreviousTab = 8 Then
         tmpArr = Split(Me.Tag, ",")
         For i = 1 To UBound(tmpArr)
            If Val(tmpArr(i)) > 0 Then
               Call CancelRowColor(PreviousTab, Val(tmpArr(i))) '清除反白,並且檢查是否有更新過資料
            End If
         Next i
         Call ClearDetail '清除單筆明細資料
      Else
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
      End If
   End If
End Sub

Private Sub SSTab1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   SSTab1.Tag = Me.Tag
End Sub

Private Sub txtTi18_GotFocus()
   TextInverse txtTi18
   CloseIme
End Sub

Private Sub txtTi18_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtTi18_Validate(Cancel As Boolean)
   If txtTi18 <> "" Then
      txtTi18 = UCase(txtTi18)
      If ChkSysName(txtTi18) = True Then
         If Left(txtTi18, 1) <> "T" And txtTi18 <> "FCT" Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbCritical
            Cancel = True
         End If
      Else
         Cancel = True
      End If
   End If
   If Cancel Then TextInverse txtTi18
End Sub

Private Sub txtTi19_GotFocus()
   TextInverse txtTi19
End Sub

Private Sub txtTi20_GotFocus()
   TextInverse txtTi20
End Sub

Private Sub txtTi20_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtTi20_LostFocus()
   If txtTi18 <> "" And txtTi19 <> "" And txtTi20 = "" Then txtTi20 = "0"
End Sub

Private Sub txtTi21_GotFocus()
   TextInverse txtTi21
End Sub

Private Sub txtTi21_LostFocus()
   If txtTi18 <> "" And txtTi19 <> "" And txtTi21 = "" Then txtTi21 = "00"
End Sub

Private Sub txtTi21_Validate(Cancel As Boolean)
Dim strTi05 As String
Dim strTi06 As String
   
   If txtTi18 <> "" And txtTi19 <> "" Then
      If txtTi20 = "" Then txtTi20 = "0"
      If txtTi21 = "" Then txtTi21 = "00"
      
      strExc(0) = "select tm12,tm15,tm10,'T' as sys_type from trademark" & _
                  " where tm01='" & txtTi18 & "'" & _
                    " and tm02='" & txtTi19 & "'" & _
                    " and tm03='" & txtTi20 & "'" & _
                    " and tm04='" & txtTi21 & "'" & _
                  " union select sp11,sp14,sp09,'S' as sys_type from servicepractice" & _
                  " where sp01='" & txtTi18 & "'" & _
                    " and sp02='" & txtTi19 & "'" & _
                    " and sp03='" & txtTi20 & "'" & _
                    " and sp04='" & txtTi21 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Left(txtTi18, 1) <> "T" And txtTi18 <> "FCT" Then
            MsgBox "系統類別錯誤，請重新輸入 !", vbExclamation, "警告！"
            Me.txtTi18.SetFocus
            Cancel = True
            Exit Sub
         Else
            If RsTemp.Fields("tm10") = "020" Then
               '2.大陸案
               strTi05 = "2"
            Else
               '4.非大陸案
               strTi05 = "4"
            End If
            If strTi05 <> Trim(Left(cboTi05, 2)) Then cboTi05.ListIndex = Val(strTi05) - 1
            Exit Sub
         End If
      Else
         MsgBox "本所案號輸入錯誤！", vbExclamation, "警告！"
         Me.txtTi18.SetFocus
         Cancel = True
         Exit Sub
      End If
      If Cancel = True Then
         cboTi05.ListIndex = 4 '其他
         cboTi06.Text = ""
         List1.Clear
         txtTi18 = "": txtTi19 = "": txtTi20 = "": txtTi21 = ""
         txtTi18.SetFocus
      End If
   End If
End Sub
