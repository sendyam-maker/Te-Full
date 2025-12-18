VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm880019 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "新郵件"
   ClientHeight    =   6156
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   8952
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6156
   ScaleWidth      =   8952
   StartUpPosition =   3  '系統預設值
   Begin TabDlg.SSTab SSTab1 
      Height          =   6135
      Left            =   60
      TabIndex        =   12
      TabStop         =   0   'False
      Top             =   30
      Width           =   8865
      _ExtentX        =   15642
      _ExtentY        =   10816
      _Version        =   393216
      TabHeight       =   520
      TabCaption(0)   =   " "
      TabPicture(0)   =   "frm880019.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Label2(0)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label3"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Label4"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Label5"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblSendMailDt"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "Label6"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "txtSubject"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "txtCopy"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "txtReceiver"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "txtBCC"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lblSender"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "FramePrint"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "SSTab2"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "cmdReceiver(1)"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "cmdReceiver(0)"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "cmdExit"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "cmdSend"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "cmdAttach"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtAttachment"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "chkPrint"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdReceiver(2)"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "CommonDialog1"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "chkImportant"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdNoSend"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "chkReceipt"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).ControlCount=   26
      TabCaption(1)   =   " "
      TabPicture(1)   =   "frm880019.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lstMailBox"
      Tab(1).Control(1)=   "Command2(1)"
      Tab(1).Control(2)=   "Command2(0)"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "Tab 2"
      TabPicture(2)   =   "frm880019.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Command3(1)"
      Tab(2).Control(1)=   "Command3(0)"
      Tab(2).Control(2)=   "MSHFlexGrid1"
      Tab(2).ControlCount=   3
      Begin VB.CheckBox chkReceipt 
         Caption         =   "讀取回條"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   144
         TabIndex        =   36
         Top             =   3000
         Visible         =   0   'False
         Width           =   1104
      End
      Begin VB.CommandButton cmdNoSend 
         Caption         =   "不寄"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   7650
         TabIndex        =   33
         Top             =   5700
         Visible         =   0   'False
         Width           =   1095
      End
      Begin VB.CheckBox chkImportant 
         Caption         =   "高重要性"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   135
         TabIndex        =   32
         Top             =   3240
         Visible         =   0   'False
         Width           =   1104
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1764
         Top             =   3492
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.CommandButton cmdReceiver 
         Caption         =   "密件副本..."
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   2
         Left            =   120
         TabIndex        =   4
         Top             =   1530
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   -73680
         TabIndex        =   27
         Top             =   5310
         Width           =   1095
      End
      Begin VB.CommandButton Command3 
         Caption         =   "確定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   -74805
         TabIndex        =   26
         Top             =   5310
         Width           =   1095
      End
      Begin VB.CheckBox chkPrint 
         Caption         =   "列印"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000C0&
         Height          =   180
         Left            =   135
         TabIndex        =   11
         Top             =   3468
         Visible         =   0   'False
         Width           =   810
      End
      Begin VB.TextBox txtAttachment 
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1065
         Left            =   1260
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  '垂直捲軸
         TabIndex        =   8
         Top             =   2580
         Width           =   7530
      End
      Begin VB.CommandButton cmdAttach 
         Caption         =   "附件..."
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Left            =   120
         TabIndex        =   7
         Top             =   2580
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "確定"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   -74805
         TabIndex        =   16
         Top             =   5310
         Width           =   1095
      End
      Begin VB.CommandButton Command2 
         Caption         =   "取消"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   -73680
         TabIndex        =   15
         Top             =   5310
         Width           =   1095
      End
      Begin VB.ListBox lstMailBox 
         Height          =   3924
         ItemData        =   "frm880019.frx":0054
         Left            =   -74910
         List            =   "frm880019.frx":0056
         Style           =   1  '項目包含核取方塊
         TabIndex        =   14
         Top             =   450
         Width           =   8700
      End
      Begin VB.CommandButton cmdSend 
         Caption         =   "寄送"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   225
         TabIndex        =   9
         Top             =   5700
         Width           =   1095
      End
      Begin VB.CommandButton cmdExit 
         Caption         =   "取消"
         CausesValidation=   0   'False
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   1350
         TabIndex        =   10
         Top             =   5700
         Width           =   1095
      End
      Begin VB.CommandButton cmdReceiver 
         Caption         =   "收件者..."
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   0
         Left            =   120
         TabIndex        =   0
         Top             =   450
         Width           =   1095
      End
      Begin VB.CommandButton cmdReceiver 
         Caption         =   "副本..."
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         Index           =   1
         Left            =   120
         TabIndex        =   2
         Top             =   990
         Width           =   1095
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
         Height          =   4845
         Left            =   -74910
         TabIndex        =   25
         Top             =   390
         Width           =   8670
         _ExtentX        =   15304
         _ExtentY        =   8551
         _Version        =   393216
         Cols            =   4
         HighLight       =   0
         AllowUserResizing=   1
         FormatString    =   "V|檔案名稱|副檔名說明|最後修改時間"
         _NumberOfBands  =   1
         _Band(0).Cols   =   4
      End
      Begin TabDlg.SSTab SSTab2 
         Height          =   1965
         Left            =   60
         TabIndex        =   29
         Top             =   3690
         Width           =   8745
         _ExtentX        =   15431
         _ExtentY        =   3471
         _Version        =   393216
         Tabs            =   2
         TabsPerRow      =   2
         TabHeight       =   520
         TabCaption(0)   =   "內文"
         TabPicture(0)   =   "frm880019.frx":0058
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "txtContent"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "HTML格式"
         TabPicture(1)   =   "frm880019.frx":0074
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "WebBrowser1"
         Tab(1).ControlCount=   1
         Begin SHDocVwCtl.WebBrowser WebBrowser1 
            Height          =   1485
            Left            =   -75000
            TabIndex        =   31
            Top             =   300
            Width           =   8700
            ExtentX         =   15346
            ExtentY         =   2619
            ViewMode        =   0
            Offline         =   0
            Silent          =   0
            RegisterAsBrowser=   0
            RegisterAsDropTarget=   1
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
            Location        =   ""
         End
         Begin MSForms.TextBox txtContent 
            Height          =   1635
            Left            =   0
            TabIndex        =   30
            Top             =   270
            Width           =   8700
            VariousPropertyBits=   -1466941413
            ScrollBars      =   2
            Size            =   "15346;2884"
            FontName        =   "新細明體-ExtB"
            FontHeight      =   195
            FontCharSet     =   136
            FontPitchAndFamily=   34
         End
      End
      Begin VB.Frame FramePrint 
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   8.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   480
         Left            =   2520
         TabIndex        =   22
         Top             =   5580
         Visible         =   0   'False
         Width           =   5055
         Begin VB.CommandButton cmdPrint 
            Caption         =   "列印"
            CausesValidation=   0   'False
            BeginProperty Font 
               Name            =   "新細明體"
               Size            =   9.6
               Charset         =   136
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   3870
            TabIndex        =   34
            Top             =   120
            Width           =   1095
         End
         Begin VB.ComboBox Combo1 
            Height          =   300
            Left            =   675
            Style           =   2  '單純下拉式
            TabIndex        =   23
            Top             =   120
            Width           =   3150
         End
         Begin VB.Label Label2 
            Caption         =   "印表機"
            Height          =   225
            Index           =   1
            Left            =   90
            TabIndex        =   24
            Top             =   180
            Width           =   585
         End
      End
      Begin MSForms.Label lblSender 
         Height          =   252
         Left            =   1260
         TabIndex        =   35
         Top             =   180
         Width           =   1608
         VariousPropertyBits=   27
         Caption         =   "lblSender"
         Size            =   "2831;450"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtBCC 
         Height          =   525
         Left            =   1260
         TabIndex        =   5
         Top             =   1500
         Width           =   7530
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "13282;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtReceiver 
         Height          =   525
         Left            =   1260
         TabIndex        =   1
         Top             =   420
         Width           =   7530
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "13282;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtCopy 
         Height          =   525
         Left            =   1260
         TabIndex        =   3
         Top             =   960
         Width           =   7530
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "13282;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin MSForms.TextBox txtSubject 
         Height          =   525
         Left            =   1260
         TabIndex        =   6
         Top             =   2040
         Width           =   7500
         VariousPropertyBits=   -1467989989
         ScrollBars      =   2
         Size            =   "13229;926"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label6 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "密件副本："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   195
         TabIndex        =   28
         Top             =   1530
         Width           =   1005
      End
      Begin VB.Label lblSendMailDt 
         AutoSize        =   -1  'True
         Caption         =   "寄件日期："
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
         Left            =   5100
         TabIndex        =   21
         Top             =   150
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label Label5 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "附件："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   20
         Top             =   2640
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "收件者："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   405
         TabIndex        =   19
         Top             =   510
         Width           =   795
      End
      Begin VB.Label Label3 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "主旨："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   18
         Top             =   2100
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "寄件者："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   192
         Index           =   0
         Left            =   408
         TabIndex        =   17
         Top             =   192
         Width           =   792
      End
      Begin VB.Label Label1 
         Alignment       =   2  '置中對齊
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "副本："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   585
         TabIndex        =   13
         Top             =   1050
         Width           =   615
      End
   End
   Begin VB.Menu mnuPopUp 
      Caption         =   "附件"
      Visible         =   0   'False
      Begin VB.Menu mnuFn 
         Caption         =   "開啟"
         Index           =   0
      End
      Begin VB.Menu mnuFn 
         Caption         =   "刪除"
         Index           =   1
      End
      Begin VB.Menu mnuFn 
         Caption         =   "取消"
         Index           =   2
      End
      Begin VB.Menu mnuFn 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   3
      End
      Begin VB.Menu mnuFn 
         Caption         =   "下載"
         Index           =   4
      End
      Begin VB.Menu mnuFn 
         Caption         =   "下載全部"
         Index           =   5
      End
      Begin VB.Menu mnuFn 
         Caption         =   ""
         Enabled         =   0   'False
         Index           =   6
      End
      Begin VB.Menu mnuFn 
         Caption         =   "壓縮加密"
         Index           =   7
      End
   End
End
Attribute VB_Name = "frm880019"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2022/02/18 改成Form2.0 ; txtReceiver、txtCopy、txtBCC、txtSubject、txtContent、lblSender(2023/3/7更換)
'Created by Morgan 2014/12/18
Option Explicit
      
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long
Dim m_stFiles As String, m_InitDir As String
Dim m_selText As String, m_selStart As Integer
Public m_bolDone As Boolean
Public m_bolSaveMail As Boolean 'Add By Sindy 2015/9/10
Public m_CP09 As String, m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String, m_CP10 As String 'Add By Sindy 2015/9/10
Dim m_PrevForm As Form '前一畫面 Add By Sindy 2015/9/10
Public m_SMB02 As String, m_SMB03 As String 'Add By Sindy 2015/9/10
Public m_SMB11 As String 'Add By Sindy 2018/8/31
Dim m_DefaultPrinter As String
Dim strPrinter As String
Public m_bolAutoMail As Boolean '是否自動寄送 Added by Morgan 2016/3/28
Public m_AddMailCache As String 'Added by Lydia 2020/08/17 傳入Y，回傳MailCache語法；不自動發email (FCP年費發文承辦Email)
Public m_bolAttFromCpp As Boolean '附件是否選自卷宗區 Added by Morgan 2016/5/20
Public m_bolPLetter As Boolean '是否為專利指示信 Added by Morgan 2018/9/14
Public m_bolTLetter As Boolean '是否為商標指示信 Added by Sindy 2020/7/21
Dim m_AttachPath As String 'Added by Morgan 2016/5/23
Public m_LP01 As String 'Add by Amy 2020/01/02
Public m_RedText As String 'Added by Morgan 2022/2/17
Public m_isCFFagent As Boolean 'Add By Sindy 2022/3/14 收件人是否為CF代理人
Public m_CustCaseNo As String 'Added by Morgan 2023/7/19 客戶案件案號
Public m_CU13 As String 'Added by Morgan 2024/5/15
Dim m_DefSendler As String 'Added by Lydia 2024/07/05 預設寄件人


Private Sub AddFile(pFileName As String)
   'Modified by Morgan 2018/9/17 修正只符合前面部分檔名時會漏掉問題
   'If InStr(m_stFiles, pFileName) = 0 Then
   If InStr(m_stFiles, pFileName & ";") = 0 Then
   'end 2018/9/17
      m_stFiles = m_stFiles & pFileName & ";"
      txtAttachment = txtAttachment & GetFileDesc(pFileName) & ";" & vbCrLf
   End If
End Sub

Private Function GetFileDesc(pFilePath As String) As String
   Dim fs, f
   
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFile(pFilePath)
   GetFileDesc = f.Name & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)"
End Function

Private Sub cmdAttach_Click()
   If m_bolAttFromCpp Then
      SetCppAttach
   Else
      SetLocalAttach
   End If
End Sub

'卷宗區及原始檔
Private Sub SetCppAttach()
   Dim stConCpp As String, stVTB1 As String, stVTB2 As String
   
   SetGrid True
   
   '要排除的附件 altr, worksheet, info, cust.pdf, reply.pdf
   'Modify By Sindy 2017/6/14 cp01=efc01(+) ==> instr(cp01||',ALL',efc01(+))>0
   'Modified by Morgan 2018/9/6 +剔除 altr,worksheet
   'modify by sonia 2018/12/22 '.INFO.'改為'.INFO',否則CFP-028921之INFO1及INFO2不會剔除
   'stConCpp = " and INSTR(lower(CPP02),'.altr.')=0 and INSTR(lower(CPP02),'.worksheet.')=0 and INSTR(lower(CPP02),'.info.')=0 and instr(lower(cpp02),'.cust.pdf')=0  and instr(lower(cpp02),'.reply.pdf')=0"
   'Modified by Morgan 2019/1/22 +剔除 ack
   'Added by Morgan 2021/3/16 +剔除 order,cus
   stConCpp = " and INSTR(lower(CPP02),'.order.pdf')=0 and INSTR(lower(CPP02),'.cus.pdf')=0 and INSTR(lower(CPP02),'.altr.')=0 and INSTR(lower(CPP02),'.worksheet.')=0 and INSTR(lower(CPP02),'.ack.')=0 and INSTR(lower(CPP02),'.info')=0 and instr(lower(cpp02),'.cust.pdf')=0  and instr(lower(cpp02),'.reply.pdf')=0"
   'Modified by Morgan 2018/9/4 +原始檔
   'Modified by Morgan 2018/9/6 CFP案要能選所有未發文或同日發文的檔案
   'Modified by Morgan 2018/9/7 副檔說明改用db自訂函數GETFILEDESC(原程式會抓到多個說明導致檔案重複列出)
   If m_CP01 = "CFP" Then
      stVTB1 = "select b.*,c.* from caseprogress a,caseprogress b,CasepaperPDF c" & _
         " where a.cp09='" & m_CP09 & "' and b.cp01(+)=a.cp01" & _
         " and b.cp02(+)=a.cp02 and b.cp03(+)=a.cp03 and b.cp04(+)=a.cp04" & _
         " and (b.cp27=a.cp27 or (b.cp158=0 and b.cp159=0))" & _
         " and cpp01(+)=b.cp09 and (cpp10 is null or cpp10='X' or cpp10='Y')" & _
         " and lower(substr(cpp02,-4))='.pdf'" & stConCpp
      'Added by Morgan 2019/1/22 +EU也可點選子案附件
      'Modified by Morgan 2021/3/16 不必限制國家,集體都要能選--玫音 Ex:CFP-032272
      stVTB1 = stVTB1 & " union select b.*,c.* from caseprogress a,patent,caseprogress b,CasepaperPDF c" & _
         " where a.cp09='" & m_CP09 & "' and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04" & _
         " and b.cp01(+)=a.cp01 and b.cp02(+)=a.cp02 and b.cp03(+)<>a.cp03 and b.cp04(+)=a.cp04" & _
         " and (b.cp27=a.cp27 or (b.cp158=0 and b.cp159=0))" & _
         " and cpp01(+)=b.cp09 and (cpp10 is null or cpp10='X' or cpp10='Y')" & _
         " and lower(substr(cpp02,-4))='.pdf'" & stConCpp
      
      stVTB2 = "select b.*,c.* from caseprogress a,caseprogress b,Casepaperfile c" & _
         " where a.cp09='" & m_CP09 & "' and b.cp01(+)=a.cp01" & _
         " and b.cp02(+)=a.cp02 and b.cp03(+)=a.cp03 and b.cp04(+)=a.cp04" & _
         " and (b.cp27=a.cp27 or (b.cp158=0 and b.cp159=0))" & _
         " and cpf01(+)=b.cp09" & _
         " and cpf01 is not null and (cpf10 is null or cpf10='X')"
      'Added by Morgan 2019/1/22 +EU也可點選子案附件
      'Modified by Morgan 2021/3/16 不必限制國家,集體都要能選--玫音 Ex:CFP-032272
      stVTB2 = stVTB2 & " union select b.*,c.* from caseprogress a,patent,caseprogress b,Casepaperfile c" & _
         " where a.cp09='" & m_CP09 & "' and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04" & _
         " and b.cp01(+)=a.cp01 and b.cp02(+)=a.cp02 and b.cp03(+)<>a.cp03 and b.cp04(+)=a.cp04" & _
         " and (b.cp27=a.cp27 or (b.cp158=0 and b.cp159=0))" & _
         " and cpf01(+)=b.cp09" & _
         " and cpf01 is not null and (cpf10 is null or cpf10='X')"
   Else
      'Modified by Morgan 2023/7/19 +開放附件 ".att." 也可選--茹曣(經理確認)
      stVTB1 = "select * from caseprogress,CasepaperPDF" & _
         " where cp09='" & m_CP09 & "' and cpp01(+)=cp09" & _
         " and (cpp10 is null or cpp10='X' or cpp10='Y')" & _
         " and (lower(substr(cpp02,-4))='.pdf' or instr(lower(cpp02),'.att.')>0) " & stConCpp
      
      stVTB2 = "select * from caseprogress,Casepaperfile" & _
         " where cp09='" & m_CP09 & "' and cpf01(+)=cp09" & _
         " and cpf01 is not null and (cpf10 is null or cpf10='X')"
   End If
   'Modified by Morgan 2025/3/28 +CPP19
   strExc(0) = "select distinct '' V,decode(cpp02,null,'',cpp02||' ('||Round(cpp03 / 1024, 2)||' KB)') as 檔案名稱" & _
      ",GETFILEDESC(cpp02,CP01,CP10) as 副檔名說明" & _
      ",sqldatet(cpp08)||' '||sqltime(cpp09)||decode(' ('||cpp05||cpp12||')',' ()','',' ('||cpp05||cpp12||')') as 最後修改時間" & _
      ",GETFILESORT(cpp02,CP01,CP10) as sort,cpp02,'1' src,cpp14,cpp19" & _
      " from (" & stVTB1 & ") X"
      
   If (m_bolPLetter And (m_CP01 = "CFP" Or m_CP01 = "CPS")) Then
      strExc(0) = strExc(0) & " union select distinct '' V,decode(cpf02,null,'',cpf02||' ('||Round(cpf03 / 1024, 2)||' KB)') as 檔案名稱" & _
      ",GETFILEDESC(cpf02,CP01,CP10) as 副檔名說明" & _
      ",sqldatet(cpf08)||' '||sqltime(cpf09)||decode(' ('||cpf05||cpf12||')',' ()','',' ('||cpf05||cpf12||')') as 最後修改時間" & _
      ",GETFILESORT(cpf02,CP01,CP10) as sort,cpf02,'2' src,cpf13,''" & _
      " from (" & stVTB2 & ") X"
   End If
   
   strExc(0) = strExc(0) & " order by src, sort"
   'end 2018/9/4

   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      SSTab1.TabVisible(2) = True
      'Modified by Morgan 2018/9/4 +原始檔
      If (m_bolPLetter And (m_CP01 = "CFP" Or m_CP01 = "CPS")) Then
         SSTab1.TabCaption(2) = "卷宗區PDF檔及原始檔"
      Else
         SSTab1.TabCaption(2) = "卷宗區PDF檔"
      End If
      'end 2018/9/4
      SSTab1.Tab = 2
      With MSHFlexGrid1
      .Visible = False
      Set .Recordset = RsTemp
      SetGrid
      .Visible = True
      End With
      SSTab1.TabVisible(0) = False
   Else
      MsgBox "卷宗區沒有附件！", vbExclamation
   End If
End Sub

Private Sub SetGrid(Optional pReset As Boolean = False)
   Dim iCol As Integer
   Dim arrGridHeadWidth
   Dim iUbound As Integer

   arrGridHeadWidth = Array(240, 3500, 1000, 2300)
   iUbound = UBound(arrGridHeadWidth)
   
   With MSHFlexGrid1
   If pReset = True Then
      .Clear
      .Rows = 2
   End If
   .FixedCols = 0
   .FormatString = "V|檔案名稱|副檔名說明|最後修改時間"
   For iCol = 0 To .Cols - 1
      If iCol <= iUbound Then
         .ColWidth(iCol) = arrGridHeadWidth(iCol)
         .ColAlignment(iCol) = flexAlignLeftCenter
      Else
         .ColWidth(iCol) = 0
      End If
   Next
   End With
End Sub

Private Sub SetLocalAttach()
   Dim stFileName As String
   Dim sFile
   Dim ii As Integer

On Error GoTo ErrHnd
   
   stFileName = "*.*"
   With CommonDialog1
      .CancelError = True
      .FileName = stFileName
      .Filter = "All Files (*.*)|*.*"
      .InitDir = m_InitDir
      .MaxFileSize = 3000
      .Flags = cdlOFNHideReadOnly Or cdlOFNPathMustExist Or cdlOFNAllowMultiselect Or cdlOFNExplorer Or cdlOFNNoDereferenceLinks
      .ShowOpen
      If .FileName <> "" Then
         If InStr(.FileName, ChrW$(0)) > 0 Then
            sFile = Split(.FileName, ChrW$(0))
            m_InitDir = sFile(0)
            For ii = 1 To UBound(sFile)
               If InStr(sFile(ii), "\") > 0 Then
                  stFileName = sFile(ii)
               Else
                  stFileName = sFile(0) & "\" & sFile(ii)
               End If
               AddFile stFileName
            Next
         Else
            strExc(1) = Left(.FileName, InStrRev(.FileName, "\"))
            If m_InitDir <> strExc(1) Then
               m_InitDir = strExc(1)
               SaveSetting "TAIE", strUserNum, UCase(Me.Name) & "Dir", m_InitDir
            End If
            AddFile .FileName
         End If
      End If
   End With
   Exit Sub
   
ErrHnd:
   If Err.NUMBER <> 32755 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub cmdExit_Click()
   'Modified by Morgan 2015/6/17
   'Unload Me
   Me.Hide
   'end 2015/6/17
End Sub

Private Function UpdateAppForm() As Boolean
   
On Error GoTo ErrHnd

   strSql = "update appform set af11=19221111,af14='" & strUserNum & "' where af01='" & m_CP09 & "' and af11=0"
   cnnConnection.Execute strSql, intI
   UpdateAppForm = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
End Function

Private Sub cmdNoSend_Click()
   'Added by Morgan 2021/3/30
   Dim stLP49 As String '確收備註
   
   If cmdNoSend.Caption = "確收" Then
      '待處理區
      If UCase(TypeName(m_PrevForm)) = UCase("frm210149") Then
         strSql = "update casepaperpdf set cpp12=cpp12 where cpp01='" & m_CP09 & "' and instr(lower(cpp02),'.cack.')>0"
         cnnConnection.Execute strSql, intI
         If intI = 0 Then
            If MsgBox("本收文號卷宗區沒有確收信(.CACK.)，是否確收？", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
               stLP49 = InputBox("確收備註:")
               If stLP49 = "" Then MsgBox "確收備註不可空白！", vbCritical: Exit Sub
            Else
               Exit Sub
            End If
         End If
      End If
      
      strSql = "update letterprogress set lp46='" & strUserNum & "',lp47=to_char(sysdate,'yyyymmdd'),lp48=to_char(sysdate,'hh24miss'),lp49='" & ChgSQL(stLP49) & "' where lp01='" & m_CP09 & "' and lp47=0"
      cnnConnection.Execute strSql, intI
      If intI = 1 Then
         m_bolDone = True
         Me.Hide
      End If

   'end 2021/3/30
   ElseIf MsgBox("是否確認不寄送？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
      If UpdateAppForm() = True Then
         m_bolDone = True
         Me.Hide
      End If
   End If
End Sub

'Add By Sindy 2015/10/15 寄件備份列印
Private Sub cmdPrint_Click()
Dim iLine As Integer

   Set Printer = Printers(Combo1.ListIndex)
   Printer.EndDoc
   Printer.Orientation = 1
   
   Printer.Font.Name = "新細明體"
   iLine = 3
   Printer.Font.Size = 14
   strExc(0) = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2)) & " " & Format(Right("000000" & ServerTime, 6), "##:##:##")
   Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth(strExc(0)) - 1500
   Printer.CurrentY = iLine * 300
   Printer.Print strExc(0)
   
   Printer.Font.Size = 16
   Printer.FontBold = True
   iLine = iLine + 2
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print strUserName
   iLine = iLine + 1
   Printer.Line (500, iLine * 300 + 50)-(10500, iLine * 300 + 50), , B
   Printer.FontBold = False
   
   Printer.Font.Size = 11
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "寄件者：" & lblSender
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print lblSendMailDt
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "收件者：" & txtReceiver
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "副本：" & txtCopy
   'Add By Sindy 2018/5/14
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "密件副本：" & txtBCC
   '2018/5/14 END
   iLine = iLine + 1
   Printer.CurrentX = 500
   Printer.CurrentY = iLine * 300
   Printer.Print "主旨：" & txtSubject
   iLine = iLine + 1
'   Printer.CurrentX = 500
'   Printer.CurrentY = iLine * 300
'   Printer.Print "附件：" & txtAttachment
   PUB_PrintFontIntoBox "附件：" & Replace(txtAttachment, vbCrLf, ""), 500, iLine * 300, 10500, (iLine + 5) * 300, False, False, , 11
   
   iLine = iLine + 7
'   Printer.CurrentX = 500
'   Printer.CurrentY = iLine * 300
'   Printer.Print txtContent
   PUB_PrintFontIntoBox txtContent, 500, iLine * 300, 10500, 51 * 300, False, False, , 11
   
   Printer.EndDoc
   
   'Modified by Morgan 2016/3/30
   '自動寄發的不要彈訊息(整批發文才不會中斷)
   'Modified by Lydia 2020/08/17
   'If Not m_bolAutoMail Then
   If Not (m_bolAutoMail Or m_AddMailCache <> "") Then
      ShowPrintOk
   End If
End Sub

Private Sub cmdReceiver_Click(Index As Integer)
   SSTab1.TabVisible(1) = True
   If Index = 0 Then
      SSTab1.TabCaption(1) = "收件者"
      SetMailBox txtReceiver.Text
   'Modify By Sindy 2018/5/14
   ElseIf Index = 1 Then
      SSTab1.TabCaption(1) = "副本"
      SetMailBox txtCopy.Text
   Else
      SSTab1.TabCaption(1) = "密件副本"
      SetMailBox txtBCC.Text
   '2018/5/14 END
   End If
   SSTab1.Tab = 1
   SSTab1.TabVisible(0) = False
End Sub

Private Sub SetMailBox(pMailList As String)
   Dim ii As Integer, jj As Integer
   Dim ArrMail() As String
   
   If lstMailBox.ListCount > 0 Then
      ii = -1
      For jj = 0 To lstMailBox.ListCount - 1
         lstMailBox.Selected(jj) = False
         If Left(lstMailBox.List(jj), 6) = strUserNum & " " Then
            ii = jj
         End If
      Next
      If SSTab1.TabCaption(1) = "收件者" And ii <> -1 Then
         lstMailBox.RemoveItem ii
      End If
   End If
   
   ArrMail = Split(pMailList, ";")
   For ii = LBound(ArrMail) To UBound(ArrMail)
      ArrMail(ii) = Trim(ArrMail(ii))
      If ArrMail(ii) <> "" Then
         For jj = 0 To lstMailBox.ListCount - 1
            If lstMailBox.List(jj) = ArrMail(ii) Then
               lstMailBox.Selected(jj) = True
               Exit For
            End If
         Next
         If jj = lstMailBox.ListCount Then
            lstMailBox.AddItem ArrMail(ii), jj
            lstMailBox.Selected(jj) = True
         End If
      End If
   Next
End Sub

Public Sub cmdSend_Click()
Dim strUpdDate As String, strUpdTime As String
Dim strSender() As String, strReceiver As String, strCopy As String, strAtt As String, strBCC As String
Dim strTemp As String, strCDate As String, strCTime As String
Dim bolHadShowMsg As Boolean 'Add By Sindy 2018/12/6
Dim iSignId As Integer 'Added by Morgan 2019/11/19 簽名檔代碼:0=無, 1=patent-中文, 2=patent-英文, 3=ipdept-英文
Dim stContent As String 'Added by Morgan 2019/11/20
'Dim strUpd As String 'Add by Amy 2020/01/02
Dim strUpdSQL As String 'Add By Sindy 2023/3/9
   
    'Added by Lydia 2022/02/18 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
    'Removed by Morgan 2023/5/11 此表單可不必檢查(內文可能會有?號)
    'If PUB_ChkUniText(Me, , True, "TextBox") = False Then
    '    Exit Sub
    'End If
    'end 2023/5/11

   'Added by Lydia 2020/08/17 回傳MailCache語法
   If cmdSend.Caption = "確定" And m_AddMailCache = "Y" Then
        Call ProcMailCache
        Me.Hide
        Exit Sub
   End If
   
   'Add By Sindy 2018/8/31
   If cmdSend.Caption = "轉寄" Then
      cmdAttach.Visible = True
      cmdReceiver(0).Visible = True
      cmdReceiver(1).Visible = True
      cmdReceiver(2).Visible = True
      txtReceiver.Locked = False
      txtCopy.Locked = False
      txtSubject = "FW: " & txtSubject
      txtSubject.Locked = False
'      txtAttachment.Locked = False
'      txtAttachment = "" '附件無實體檔,須重新夾帶
      txtContent = vbCrLf & vbCrLf & "-------------------- Original message --------------------" & vbCrLf & _
                   "From: " & lblSender & vbCrLf & _
                   "Sent: " & Replace(lblSendMailDt.Caption, "寄件日期：", "") & vbCrLf & _
                   "To: " & txtReceiver & vbCrLf & _
                   "Subject: " & txtSubject.Tag & vbCrLf & vbCrLf & txtContent
      'Add By Sindy 2021/5/3 吳碧梧(智權.經理.Wu):想呈現每次給客戶會稿的明確資訊內容
      If frm880019.txtContent.Tag <> "" Then
         txtContent = txtContent.Tag & txtContent.Text
      End If
      '2021/5/3 END
      txtContent.Locked = False
      cmdSend.Caption = "寄送"
      Exit Sub
   End If
   '2018/8/31 END
   'Add By Sindy 2018/9/13 附件要回寫到歷程附件檔
   Dim arrFile As Variant
   Dim ii As Integer, jj As Integer
   Dim fs, f
   Dim bolFind As Boolean 'Add By Sindy 2018/10/5
   'Modify By Sindy 2020/7/21 + And m_bolTLetter = False
   If UCase(TypeName(m_PrevForm)) = UCase("frm090202_2") And m_bolTLetter = False Then
      If m_stFiles <> "" Then
         m_PrevForm.lstAtt(0).Clear '存E-Mail中的附件
         arrFile = Split(m_stFiles, ";")
         Set fs = CreateObject("Scripting.FileSystemObject")
         For ii = LBound(arrFile) To UBound(arrFile) - 1
            'Add By Sindy 2018/12/6 檔案是否正在使用中
            If PUB_ChkFileOpening(CStr(arrFile(ii)), bolHadShowMsg) = True Then
               If bolHadShowMsg = False Then
                  MsgBox arrFile(ii) & vbCrLf & "檔案正在使用中，請關閉才可執行送出！", vbExclamation
               End If
               Exit Sub
            End If
            '2018/12/6 END
'            'Add By Sindy 2018/10/5 若在EMail視窗加的電子檔也要放入歷程附件區,以利存檔
'            bolFind = False
'            If arrFile(ii) <> "" Then
'               For jj = 0 To m_PrevForm.lstAtt(0).ListCount - 1
'                  If InStr(m_PrevForm.lstAtt(0).List(jj), arrFile(ii)) > 0 Then
'                     bolFind = True
'                     Exit For
'                  End If
'               Next jj
'               If bolFind = False Then
'            '2018/10/5 END
                  Set f = fs.GetFile(arrFile(ii))
                  m_PrevForm.lstAtt(0).AddItem arrFile(ii) & " (" & Round(f.Size / 1024, 2) & " KB)" & " #" & Format(f.DateLastModified, "YYYYMMDDHHMMSS") & "#"
'               End If
'            End If
         Next ii
'         'Add By Sindy 2018/10/16 移除歷程附件區在EMail附件區不存在的電子檔
'         For jj = m_PrevForm.lstAtt(0).ListCount - 1 To 0 Step -1
'            strTemp = Mid(m_PrevForm.lstAtt(0).List(jj), 1, InStrRev(m_PrevForm.lstAtt(0).List(jj), "(") - 1)
'            If InStr(UCase(m_stFiles), UCase(Trim(strTemp))) = 0 Then
'               m_PrevForm.lstAtt(0).RemoveItem jj
'            End If
'         Next jj
'         '2018/10/16 END
      Else
         m_PrevForm.lstAtt(0).Clear '全部清除
      End If
   '2018/9/13 END
   Else
      'Modify By Sindy 2019/1/21 有附件就要檢查是否檔案正在使用中，懷疑使用中會導至寄出時附件遺失
      arrFile = Split(m_stFiles, ";")
      Set fs = CreateObject("Scripting.FileSystemObject")
      For ii = LBound(arrFile) To UBound(arrFile) - 1
         '檔案是否正在使用中
         If PUB_ChkFileOpening(CStr(arrFile(ii)), bolHadShowMsg) = True Then
            If bolHadShowMsg = False Then
               MsgBox arrFile(ii) & vbCrLf & "檔案正在使用中，請關閉才可執行送出！", vbExclamation
            End If
            Exit Sub
         End If
      Next ii
      '2019/1/21 END
   End If
   
   'Add By Sindy 2019/2/23
   'Modified by Lydia 2019/06/19 判斷附件按鈕才提醒(因為工程師認翻譯呈報主管Email不用附件)
   'If Trim(m_stFiles) = "" Then
   If Trim(m_stFiles) = "" And cmdAttach.Visible = True Then
      If MsgBox("無附件，是否確認寄送？", vbYesNo + vbDefaultButton2 + vbQuestion, "郵寄") = vbNo Then
         Exit Sub
      End If
   End If
   '2019/2/23 END
   
   'Add by Amy 2019/01/02 有傳lp01資料需更新 E化寄送人員/日期/時間
   'Removed by Morgan 2021/4/1 移到寄信成功後並改記錄與寄件備份相同的時間以便連結
   'If m_LP01 <> MsgText(601) Then
   '     strUpd = "Update LetterProgress Set lp38='" & strUserNum & "',lp39=to_char(sysdate,'yyyymmdd'),lp40=to_char(sysdate,'hh24miss') " & _
   '                    "Where lp01='" & m_LP01 & "'"
   '     cnnConnection.Execute strUpd
   'End If
   'end 2021/4/1
   
   strReceiver = GetMailList(txtReceiver.Text)
   If txtReceiver = "" Then MsgBox "請輸入收件者！", vbExclamation: Exit Sub
   strCopy = GetMailList(txtCopy.Text)
   strBCC = GetMailList(txtBCC.Text) 'Add By Sindy 2018/5/14
   strAtt = Replace(m_stFiles, ";", "*")
   If ChkAttSize(strAtt) = False Then Exit Sub 'Added by Morgan 2022/5/3 檢查附件大小是否超過
   
   'Added by Morgan 2023/9/19
   '檢查畫面的附件檔案數量是否與傳送的相同
   If txtAttachment <> "" Then
      If Len(txtAttachment) - Len(Replace(txtAttachment, ";", "")) <> Len(strAtt) - Len(Replace(strAtt, "*", "")) Then
         MsgBox "附件檔案數量檢查失敗，請重新操作！", vbExclamation: Exit Sub
      End If
   End If
   'end 2023/9/19
   
   'Modified by Morgan 2023/5/18 改鎖表單以避免點到關閉或其他按鈕
   'cmdSend.Enabled = False 'Add By Sindy 2022/2/22 防止按2次
   Me.Enabled = False
   'end 2023/5/18
   
   'Modified by Morgan 2015/3/27
   'PUB_SendMail strUserNum, strReceiver, "", txtSubject.Text, txtContent.Text, , strAtt, , True, , strCopy, , , , True
   strSender = Split(Trim(lblSender))
'      txtContent.Text = txtContent.Text & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
'                  "<BR>" & _
'                  " <P><B><I><FONT face=新細明體 size=1>Tai E International Patent &amp; Law" & _
'                  " Office</FONT></I></B><I></I><FONT face=新細明體 size=1> / </FONT><FONT face=Arial" & _
'                  " size=1>kc</FONT><BR><I><FONT face=新細明體 size=1>9F, 112, Section 2, Chang-An East" & _
'                  " Road Taipei, Taiwan, R.O.C. ;P.O. Box 46-478, Taipei, Taiwan," & _
'                  " R.O.C.</FONT></I><FONT face=新細明體 size=1><BR></FONT><I><FONT face=新細明體" & _
'                  " size=1>Tel: 886-2-25061023, 886-2-25081531 Fax: 886-2-25068147, 886-2-25076571," & _
'                  " 886-2-25090804,</FONT></I><BR><I><FONT face=新細明體 size=1>886-2-25064319<U> URL:" & _
'                  " &lt;<A href=http://www.taie.com.tw/" & _
'                  " target=_blank>http://www.taie.com.tw</A>&gt;</U>" & _
'                  " E-mail:patent@taie.com.tw</FONT><FONT face=Courier New" & _
'                  " size=1>&nbsp;</FONT></I>&nbsp;<FONT face=新細明體 size=1><BR>************* Email" & _
'                  " Confidentiality Notice ********************<BR>This e-mail transmission is" & _
'                  " intended only for the use of the individual<BR>or entity to which it is" & _
'                  " addressed, and may contain information that is<BR>privileged, confidential and" & _
'                  " exempt from disclosure under applicable<BR>law.</FONT><FONT face=Courier New" & _
'                  " size=1></FONT>&nbsp;<FONT face=新細明體 size=1> If the reader is not the intended" & _
'                  " recipient, you are hereby<BR>notified that any dissemination, distribution or" & _
'                  " copying of this<BR>communication is strictly prohibited.</FONT><FONT" & _
'                  " face=Courier New size=1></FONT>&nbsp;<FONT face=新細明體 size=1> If you have" & _
'                  " received this<BR>transmission in error, please notify us immediately, and return" & _
'                  " the<BR>original message to us at the above address.</FONT><FONT" & _
'                  " face=Courier New size=1></FONT>&nbsp;<FONT face=新細明體 size=1> We greatly" & _
'                  " appreciate your<BR>cooperation. </FONT></P><BR>"
   
   'Modified by Morgan 2019/11/19 +簽名檔
   'PUB_SendMail strSender(0), strReceiver, "", Trim(txtSubject.Text), Trim(txtContent.Text), , strAtt, True, True, , strCopy, strSender(0), , , True, , strBCC, , , , , , IIf(chkImportant.Value = vbChecked, True, False)
   stContent = Trim(txtContent.Text)
   iSignId = 0
   'Modify By Sindy 2020/1/13 商標處程序
   'If Left(Pub_StrUserSt03, 2) = "P2" Then
   'Modify By Sindy 2020/2/19 商標處程序也會是智權人員身分
   '                          ,會操作待會稿區的客戶會稿E-Mail,要排除圖片簽名檔,因為在內文裡面已有簽名檔
   'Modify By Sindy 2020/10/14 改判斷 Left(Pub_StrUserSt03, 2) = "P2"
'   If Pub_StrUserSt03 = "P22" And _
'      InStr(stContent, "本信件僅授權於指定之收信人取閱之用") = 0 Then
   If Left(Pub_StrUserSt03, 2) = "P2" And _
      InStr(stContent, "本信件僅授權於指定之收信人取閱之用") = 0 Then
      iSignId = 4
   '2020/1/13 END
   ElseIf m_bolPLetter Then
      iSignId = PUB_GetSignID(m_CP01)
   End If
   
   'Added by Morgan 2022/2/17 特殊內容設紅色字
   If m_RedText <> "" Then
      stContent = PUB_Text2Html(stContent)
      m_RedText = PUB_Text2Html(m_RedText)
      stContent = Replace(stContent, m_RedText, "<span style='color:red'>" & m_RedText & "</span>")
   End If
   'end 2022/2/17
   
   'Modify By Sindy 2023/3/9 發生信有寄出去,但記錄且是空白的 ex:T-131764(通知期限-延展)
   '將SQL語法寄信前先組好,寄信成功後再寫進DB
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   strUpdSQL = "insert into smailbackup(smb01,smb02,smb03,smb04,smb05,smb06,smb07,smb08,smb09,smb10,smb11)" & _
               " values('" & m_CP09 & "'," & strUpdDate & "," & strUpdTime & _
                    ",'" & ChgSQL(Trim(lblSender.Caption)) & "'" & _
                    ",'" & ChgSQL(Trim(txtReceiver)) & "'" & _
                    ",'" & ChgSQL(Trim(txtCopy)) & "'" & _
                    ",'" & ChgSQL(Trim(txtSubject)) & "'" & _
                    ",'" & ChgSQL(Trim(txtAttachment)) & "'" & _
                    ",'" & ChgSQL(Trim(txtContent)) & "'" & _
                    ",'" & ChgSQL(Trim(txtBCC)) & "'," & CNULL(m_SMB11, True) & ")"
   '2023/3/9 END
   'Modified by Morgan 2023/9/15 +chkReceipt
   PUB_SendMail strSender(0), strReceiver, "", Trim(txtSubject.Text), stContent, , strAtt, True, True, , strCopy, strSender(0), , , True, , strBCC, , , , , , IIf(chkImportant.Value = vbChecked, True, False), , iSignId, IIf(chkReceipt.Value = vbChecked, True, False)
   'end 2019/11/19
   
   'end 2015/3/27
   If bolMailSendOk = True Then
      'Modified by Morgan 2015/6/17
      'Unload Me
      m_bolDone = True
      'Add By Sindy 2015/9/10
      If m_bolSaveMail = True And m_CP09 <> "" Then
'         strUpdDate = strSrvDate(1)
'         strUpdTime = Right("000000" & ServerTime, 6)
         '存EMail寄件備份
         'Modify By Sindy 2018/8/31 + smb11
         'Modified by Morgan 2018/9/14 修正主旨,內文有單引號會當問題(+ChgSQL)
         'Modified by Morgan 2018/10/30 修正收件者有單引號會當問題(文字全+ChgSQL)
         'Modify By Sindy 2023/3/9 改在寄信前先組語法
'         strSql = "insert into smailbackup(smb01,smb02,smb03,smb04,smb05,smb06,smb07,smb08,smb09,smb10,smb11)" & _
'                  " values('" & m_CP09 & "'," & strUpdDate & "," & strUpdTime & _
'                          ",'" & ChgSQL(Trim(lblSender.Caption)) & "'" & _
'                          ",'" & ChgSQL(Trim(txtReceiver)) & "'" & _
'                          ",'" & ChgSQL(Trim(txtCopy)) & "'" & _
'                          ",'" & ChgSQL(Trim(txtSubject)) & "'" & _
'                          ",'" & ChgSQL(Trim(txtAttachment)) & "'" & _
'                          ",'" & ChgSQL(Trim(txtContent)) & "'" & _
'                          ",'" & ChgSQL(Trim(txtBCC)) & "'," & CNULL(m_SMB11, True) & ")"
'         cnnConnection.Execute strSql
         cnnConnection.Execute strUpdSQL
'         If txtSubject = "" And txtContent = "" Then
''            發生txtSubject = 空白 And txtContent = 空白
''            txtSubject.Visible = False
''            txtSubject.Enabled = True
''            txtContent.Visible = False
''            txtContent.Enabled = True
'            PUB_SendMail strUserNum, "97038", "", "[frm880019]發生信有寄出去,但記錄且是空白的", strUpdSQL & vbCrLf & vbCrLf & _
'                  "發生txtSubject=空白 And txtContent=空白" & vbCrLf & vbCrLf & _
'                  "txtSubject.Visible = " & txtSubject.Visible & vbCrLf & _
'                  "txtSubject.Enabled = " & txtSubject.Enabled & vbCrLf & _
'                  "txtContent.Visible = " & txtContent.Visible & vbCrLf & _
'                  "txtContent.Enabled = " & txtContent.Enabled & vbCrLf
'         End If
         '2023/3/9 END
         
         'Added by Morgan 2019/11/22 從PUB_SendOrderLetterP移來(因發生有寄信但仍顯示於待處理區問題,可能是FMP案列印EMail後訊息沒回且不正常結束造成)
         If m_bolPLetter = True Then
            strSql = "update appform set af11=" & strUpdDate & ",af12=" & strUpdTime & ",af14='" & strUserNum & "' where af01='" & m_CP09 & "'"
            cnnConnection.Execute strSql, intI
         End If
         'end 2019/11/22
      
         'Added by Morgan 2021/4/1 從上面移來並改記錄與寄件備份相同的時間以便連結
         If m_LP01 <> MsgText(601) Then
              strSql = "Update LetterProgress Set lp38='" & strUserNum & "',lp39=" & strUpdDate & ",lp40=" & strUpdTime & " Where lp01='" & m_LP01 & "'"
              cnnConnection.Execute strSql, intI
         End If
   
         m_SMB02 = strUpdDate 'Added by Morgan 2015/11/24
         m_SMB03 = strUpdTime 'Added by Morgan 2015/11/24
         
         '產生本所案號+案件性質+Email.menu
         '注意 : 資料來源不存”S”,是因為當重新歸檔時不可清除此筆記錄
         'Modify By Sindy 2020/2/19 電子檔名,本所案號使用函數 PUB_CaseNo2FileName
         strSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,cpp08,cpp09,cpp10)" & _
                  " values('" & m_CP09 & "','" & PUB_CaseNo2FileName(m_CP01, m_CP02, m_CP03, m_CP04) & _
                          "." & m_CP10 & "." & strUpdDate & strUpdTime & "." & EMP_Email & ".menu',0," & _
                          strUpdDate & "," & strUpdTime & ",'Y')"
         cnnConnection.Execute strSql
         
         'Add By Sindy 2015/10/15
         If chkPrint.Value = 1 And chkPrint.Visible = True Then
            lblSendMailDt.Visible = True
            strTemp = TAIWANDATE(strUpdDate)
            strCDate = Format(strTemp, "###/##/##")
            strTemp = strUpdTime
            strCTime = Format(strTemp, "##:##:##")
            lblSendMailDt.Caption = "寄件日期：" & strCDate & " " & strCTime
            Call cmdPrint_Click
         End If
         '2015/10/15 END
      End If
      '2015/9/10 END
      
      Me.Hide
      'end 2015/6/17
   End If
   
   'Modified by Morgan 2023/5/18 改鎖表單以避免點到關閉或其他按鈕
   'cmdSend.Enabled = True 'Add By Sindy 2022/2/22
   Me.Enabled = True
   'end 2023/5/18
End Sub

Private Sub Command2_Click(Index As Integer)
   If Index = 0 Then
      If SSTab1.TabCaption(1) = "收件者" Then
         txtReceiver.Text = ""
      'Modify By Sindy 2018/5/14
      ElseIf SSTab1.TabCaption(1) = "副本" Then
         txtCopy.Text = ""
      Else
         txtBCC.Text = ""
      '2018/5/14 END
      End If
      For intI = 0 To Me.lstMailBox.ListCount - 1
         If lstMailBox.Selected(intI) = True Then
            If SSTab1.TabCaption(1) = "收件者" Then
               txtReceiver.Text = txtReceiver.Text & lstMailBox.List(intI) & "; "
            'Modify By Sindy 2018/5/14
            ElseIf SSTab1.TabCaption(1) = "副本" Then
               txtCopy.Text = txtCopy.Text & lstMailBox.List(intI) & "; "
            Else
               txtBCC.Text = txtBCC.Text & lstMailBox.List(intI) & "; "
            '2018/5/14 END
            End If
         End If
      Next
   End If
   SSTab1.TabVisible(0) = True
   SSTab1.Tab = 0
   Me.SSTab1.TabVisible(1) = False
   If SSTab1.TabCaption(1) = "收件者" Then
      txtReceiver.SetFocus
   'Add By Sindy 2018/5/14
   ElseIf SSTab1.TabCaption(1) = "副本" Then
      txtCopy.SetFocus
   Else
      txtBCC.SetFocus
   '2018/5/14 END
   End If
End Sub

Private Sub Command3_Click(Index As Integer)
   Dim stFileName As String
   Dim stFtpPath As String 'Added by Morgan 2018/9/4
   Dim stTableName As String 'Added by Morgan 2018/9/11
   Dim stCPP19 As String 'Added by Morgan 2025/3/28
   
   If Index = 0 Then
      For intI = 1 To MSHFlexGrid1.Rows - 1
         If MSHFlexGrid1.TextMatrix(intI, 0) = "V" Then
            'Modified by Morgan 2018/9/4 +原始檔
            'Modified by Morgan 2018/9/11 因可能會選到該案其他收文號的檔案改寫法
            stTableName = ""
            '卷宗區
            If MSHFlexGrid1.TextMatrix(intI, 6) = "1" Then
               stTableName = "CASEPAPERPDF"
            '原始檔
            ElseIf MSHFlexGrid1.TextMatrix(intI, 6) = "2" Then
               stTableName = "CASEPAPERFILE"
            End If
            If stTableName <> "" Then
               'Modified by Morgan 2023/7/19
               If m_CustCaseNo <> "" Then
                  stFileName = m_AttachPath & "\" & PUB_FilterEFileSymbol(m_CustCaseNo) & "." & MSHFlexGrid1.TextMatrix(intI, 5)
               Else
               'end 2023/7/19
                  stFileName = m_AttachPath & "\" & MSHFlexGrid1.TextMatrix(intI, 5)
               End If
               stFtpPath = MSHFlexGrid1.TextMatrix(intI, 7)
               stCPP19 = MSHFlexGrid1.TextMatrix(intI, 8) 'Added by Morgan 2025/3/28
               If PUB_GetFtpFile(stFtpPath, stFileName, stTableName, True, , stCPP19 <> "") Then
                  AddFile stFileName
               End If
            End If
            'end 2018/9/11
         End If
      Next
   End If
   SSTab1.TabVisible(0) = True
   SSTab1.Tab = 0
   SSTab1.TabVisible(2) = False
End Sub

'Added by Morgan 2016/3/28
Private Sub Form_Activate()
Dim strText As String
Static bActivated As Boolean

   'Add By Sindy 2020/10/14
   'Modified by Morgan 2021/3/4 改判斷有寄送按鈕時詢問(因E化客戶指定信箱不可修改)
   'If txtReceiver.Locked = False Then
   If cmdSend.Visible = True Then
      If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or Pub_StrUserSt03 = "M51" Then
         strText = InputBox("要更改寄件者嗎?")
         If strText <> "" Then
            lblSender = strText & " (" & GetPrjSalesNM(strText) & ")"
         End If
         Screen.MousePointer = vbDefault 'Added by Lydia 2023/08/18
      End If
   End If
   '2020/10/14 END
   
   '自動寄送
   If m_bolAutoMail Then
      cmdSend_Click
   'Added by Morgan 2018/9/6
   'Modified by Morgan 2018/9/12
   ElseIf m_bolPLetter Or InStr(txtContent.Text, "&nbsp;") > 0 Then
      SSTab2.TabVisible(1) = True
      'SSTab2.Tab = 1 'Removed by Morgan 2019/11/20 落款改有圖載入比較久,改先不切換頁籤
      If m_bolPLetter Then
         If m_CP01 = "CFP" Or m_CP01 = "CPS" Then
            cmdNoSend.Visible = True
            txtContent.FontName = "Times New Roman"
         End If
      End If
      
      'Added by Morgan 2019/11/22
      '檢視模式(原始內文會有 TAG, 改用HTML顯示)
      If cmdSend.Visible = False Then
         SSTab2.Tab = 1
         SSTab2.TabCaption(1) = SSTab2.TabCaption(0)
         SSTab2.TabVisible(0) = False
      End If
      'end 2019/11/22
   'Added by Lydia 2020/08/17 傳入Y，回傳MailCache語法
   ElseIf m_AddMailCache = "Y" Then
       cmdSend.Caption = "確定"
       cmdAttach.Enabled = False
       cmdReceiver(2).Enabled = False
   'end 2020/08/17
   End If
   
   'Added by Morgan 2022/3/29
   If Not bActivated Then
      If txtContent.Enabled = True And Me.Visible = True Then
         txtContent.SetFocus
         txtContent.SelStart = 0
      End If
      bActivated = True
   End If
   'end 2022/3/29
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SSTab1.TabVisible(1) = False
   SSTab1.TabVisible(2) = False
   
   'Added by Morgan 2015/3/27
   '專利處人員固定以patent@taie.com.tw為寄件者 - 郭雅娟
   'Modified by Morgan 2024/5/15
   'If Left(Pub_StrUserSt15, 2) = "P1" Then
   If Left(Pub_StrUserSt15, 2) = "P1" Or m_CU13 = "P1004" Then
   'end 20224/5/15
      lblSender = "patent@taie.com.tw"
   'Add by Sindy 2024/10/9 財務扣繳
   ElseIf Pub_StrUserSt15 = "M31" Then
      lblSender = "taieacc@taie.com.tw"
   'Added by Morgan 2024/8/5 創新智權專利
   ElseIf m_CU13 = "30015" Then
      lblSender = "inno.ip@taie.com.tw"
   'Add By Sindy 2021/7/6 W20.顧問服務組
   ElseIf Left(Pub_StrUserSt15, 2) = "W2" Then
      lblSender = "ACS01@taie.com.tw"
   '2021/7/6 END
   'Modify By Sindy 2018/5/14
   '商標處人員固定以tm@taie.com.tw為寄件者 - 商標公用信箱
   'Modify By Sindy 2018/5/25 桂英會有智權人員身份,會用自己的信箱發信給客戶
   'ElseIf Left(Pub_StrUserSt15, 2) = "P2" Then
   ElseIf InStr(lblSender, "@taie.com.tw") = 0 Then
'      lblSender = "tm@taie.com.tw"
'   '2018/5/14 END
'   Else
      lblSender = strUserNum & " (" & strUserName & ")"
   End If
   'end 2015/3/27
   
   'Added by Lydia 2024/07/05 預設寄件人
   If m_DefSendler <> "" Then
      lblSender = Replace(PUB_ReadUserData(m_DefSendler, True, "1"), ",", ";")
   End If
   'end 2024/07/05
   
   PUB_SetPrinter Me.Name, Combo1, m_DefaultPrinter '抓系統中目前預設的印表機
   strPrinter = PUB_GetOsDefaultPrinter '抓控制台目前預設的印表機
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   
   'Added by Morgan 2018/9/5
   WebBrowser1.Navigate "about:blank"
   SSTab2.Tab = 0
   SSTab2.TabVisible(1) = False
   'end 2018/9/5
End Sub

Private Sub Form_Unload(Cancel As Integer)
   
   'Add By Sindy 2015/9/11
   If TypeName(m_PrevForm) <> "Nothing" Then
      'm_PrevForm.Show
      Set m_PrevForm = Nothing
   End If
   '2015/9/11 END
   
   Set frm880019 = Nothing
End Sub

Public Sub SetAttach(pFiles As String)
   Dim arrFile() As String
   Dim ii As Integer
   
   m_stFiles = ""
   arrFile = Split(pFiles, ";")
   'm_InitDir = PUB_Getdesktop
   'If arrFile(0) <> "" Then If InStrRev(arrFile(0), "\") > 0 Then m_InitDir = Left(arrFile(0), InStrRev(arrFile(0), "\") - 1)
   m_InitDir = GetSetting("TAIE", strUserNum, UCase(Me.Name) & "Dir", "")
   If m_InitDir = "" Then m_InitDir = GetMyDocPath
   txtAttachment.Text = ""
   For ii = LBound(arrFile) To UBound(arrFile)
      If arrFile(ii) <> "" Then
         m_stFiles = m_stFiles & arrFile(ii) & ";"
         'txtAttachment = txtAttachment & Mid(arrFile(ii), InStrRev(arrFile(ii), "\") + 1) & "; "
         txtAttachment = txtAttachment & GetFileDesc(arrFile(ii)) & ";" & vbCrLf
      End If
   Next
End Sub
'Added by Morgan 2021/3/18
'設定寶齡富錦 Y55435 案件信箱
Public Sub SetPBFEmail(pCustNo As String)
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
      
   txtBCC = strUserNum & " (" & strUserName & "); "
      
   stSQL = "select pcc08||' ('||nvl(pcc05,nvl(pcc03,pcc04))||')',pcc02,pcc08,pcc05 from potcustcont where pcc01='" & Left(pCustNo, 8) & "' and instr(pcc08,'@')>0 order by pcc02"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      Do While Not .EOF
         lstMailBox.AddItem .Fields(0)
         If .Fields("pcc02") <= "02" Then
            lstMailBox.Selected(lstMailBox.ListCount - 1) = True
            txtReceiver = txtReceiver & .Fields(0) & "; "
            
         'Modified by Morgan 2021/8/23 接洽人會異動,副本改抓"美代"或"美申"開頭
         'ElseIf .Fields("pcc02") <= "12" Then
         ElseIf Left(.Fields("pcc05"), 2) = "美代" Or Left(.Fields("pcc05"), 2) = "美申" Then
            txtCopy = txtCopy & .Fields(0) & "; "
         End If
         .MoveNext
      Loop
      End With
   End If
   
   Set rsQuery = Nothing
End Sub

'Added by Morgan 2018/10/30
'設定E化客戶相關信箱
Public Sub SetECustEmail(pCustNo As String)
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
      
   'txtReceiver.Locked = True
   'cmdReceiver(0).Enabled = False
   'Modified by Morgan 2021/3/16 改放密件副本--文雄
   'txtCopy = strUserNum & " (" & strUserName & "); "
   'Modified by Morgan 2022/3/22
   'txtBCC = strUserNum & " (" & strUserName & "); "
   txtBCC = ChkSpecMailReciver(strUserNum, m_CP09, False, False)
   Call CheckMail(txtBCC)
   'end 2022/3/22
   'end 2021/3/16
   
   stSQL = "select cu176,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cu04,cu185,cu186,cu187,cu188" & _
      " from customer where cu01='" & Left(pCustNo, 8) & "' and cu02='" & Mid(pCustNo, 9) & "'"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      If InStr(m_CP01, "T") > 0 And Not IsNull(.Fields("cu187")) Then
         txtReceiver = .Fields("cu187")
         txtCopy = "" & .Fields("cu188")
      Else
         txtReceiver = .Fields("cu176")
         txtCopy = "" & .Fields("cu185")
      End If
      txtReceiver = txtReceiver & " (指定信箱:" & .Fields("cu04") & ")" & "; "
      
      'Added by Morgan 2025/2/27
      If InStr("," & .Fields("cu186") & ",", "," & 勾選讀取回條 & ",") > 0 Then
         chkReceipt.Value = 1
      End If
      'end 2025/2/27
      End With
      m_bolAttFromCpp = True 'Added by Morgan 2021/11/11 給客戶的檔案應該都要先上傳卷宗區
   End If
   
   Set rsQuery = Nothing
End Sub

'Modified by Morgan 2015/11/24 +pNoCopy:不預設副本收受人
'Modify By Sindy 2018/5/11 + Optional strTxtCopy As String = "" : 指定副本收件人
'                          , Optional strTxtBCC As String = "" : 指定密件副本收件人
Public Sub SetEmail(ByVal pCustNo As String, ByVal pContactNo As String, Optional pFCAgent As String, Optional pFCContactNo As String, Optional pNoCopy As Boolean = False, _
   Optional strTxtCopy As String = "", Optional strTxtBCC As String = "")
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stMainMail As String
   Dim varTemp As Variant, ii As Integer, strTemp As String 'Add By Sindy 2018/5/11
   Dim strCP13 As String
   
   'Modified by Morgan 2014/12/30
   '預設寄代表信箱若有預設接洽人有設且信箱不同於代表信箱時也要寄
   '有FC代理人時不抓客戶EMail
   txtReceiver = ""
   lstMailBox.Clear
   'Add by Sindy 2019/8/7 pFCAgent 此變數值,有人傳入E-Mail,排除此狀況
   stSQL = ""
   If pFCAgent <> "" And InStr(pFCAgent, "@") = 0 Then
   '2019/8/7 END
      pFCAgent = ChangeCustomerL(pFCAgent) 'Add by Sindy 2019/8/6
      '代理人
      'Modify By Sindy 2018/8/30 +客戶/代理人姓名
      stSQL = "select fa16 cu20,fa80 cu116,fa81 cu117,fa82 cu118,'' CU127,NVL(FA04,DECODE(FA05,NULL,FA06,FA05||' '||FA63||' '||FA64||' '||FA65)) as cu04 from fagent where fa01='" & Left(pFCAgent, 8) & "' and fa02='0'"
      pCustNo = pFCAgent
      pContactNo = pFCContactNo
   Else
      'Add by Sindy 2019/8/7 pCustNo 此變數值,有人傳入E-Mail,排除此狀況
      'Modified by Lydia 2022/10/12 直接傳入代表信箱
      'If InStr(pCustNo, "@") = 0 Then
      If pFCAgent <> "" And InStr(pFCAgent, "@") > 0 Then
           txtReceiver = pFCAgent
      ElseIf pCustNo <> "" And InStr(pCustNo, "@") = 0 Then
      '2019/8/7 END
         pCustNo = ChangeCustomerL(pCustNo) 'Add by Sindy 2019/8/6
         '客戶
         'Modify By Sindy 2018/8/30 +客戶/代理人姓名
         stSQL = "select cu20,cu116,cu117,cu118,CU127,NVL(CU04,DECODE(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as cu04 from customer where cu01='" & Left(pCustNo, 8) & "' and cu02='0'"
      End If
   End If
   If stSQL <> "" Then
      intR = 1
      Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         With rsQuery
         'Modify By Sindy 2018/8/30 +客戶/代理人姓名
         If InStr("" & .Fields("cu20"), "@") > 0 Then
            stMainMail = .Fields("cu20")
            lstMailBox.AddItem .Fields("cu20") & " (代表信箱:" & .Fields("cu04") & ")"
            lstMailBox.Selected(lstMailBox.ListCount - 1) = True
            txtReceiver = txtReceiver & .Fields("cu20") & " (代表信箱:" & .Fields("cu04") & ")" & "; "
         End If
         If InStr("" & .Fields("cu116"), "@") > 0 Then
            lstMailBox.AddItem .Fields("cu116") & " (其他信箱1:" & .Fields("cu04") & ")"
         End If
         If InStr("" & .Fields("cu117"), "@") > 0 Then
            lstMailBox.AddItem .Fields("cu117") & " (其他信箱2:" & .Fields("cu04") & ")"
         End If
         If InStr("" & .Fields("cu118"), "@") > 0 Then
            lstMailBox.AddItem .Fields("cu118") & " (其他信箱3:" & .Fields("cu04") & ")"
         End If
         If pContactNo = "" Then
            pContactNo = "" & .Fields("CU127")
         End If
         End With
      End If
   End If
   
   '接洽人
   stSQL = "select pcc08||' ('||nvl(pcc05,nvl(pcc03,pcc04))||')',pcc02,pcc08 from potcustcont where pcc01='" & Left(pCustNo, 8) & "' and instr(pcc08,'@')>0"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With rsQuery
      Do While Not .EOF
         lstMailBox.AddItem .Fields(0)
         If pContactNo = .Fields("pcc02") Then
            If LCase(stMainMail) <> LCase(.Fields("pcc08")) Then
               lstMailBox.Selected(lstMailBox.ListCount - 1) = True
               txtReceiver = txtReceiver & .Fields(0) & "; "
            End If
         End If
         .MoveNext
      Loop
      End With
   End If
   
   'Modify By Sindy 2018/5/11
   '指定密件副本收件人
   If strTxtBCC <> "" Then
      varTemp = Split(strTxtBCC, ";")
      For ii = LBound(varTemp) To UBound(varTemp)
         If varTemp(ii) <> "" Then
            'Add By Sindy 2022/1/21
            If GetPrjSalesNM(CStr(varTemp(ii))) <> "" Then
            '2022/1/21 END
               strTemp = varTemp(ii) & " (" & GetPrjSalesNM(CStr(varTemp(ii))) & ")"
            Else
               strTemp = varTemp(ii)
            End If
            lstMailBox.AddItem strTemp
            txtBCC = txtBCC & strTemp & "; "
         End If
      Next
   End If
   
   'Add Sindy 2022/1/25 案件智權人員為顧服組,非顧服組操作,副本收件人尚未掛W2001顧服組時,副本要加掛W2001顧服組
   'Modify By Sindy 2022/1/26 秀玲:不論是任何人操作寄信及客戶會稿，只要是顧服組的客戶，副本一律加ACS01@taie.com.tw ==> 拿掉 And Left(Pub_StrUserSt15, 2) <> "W2"
   strCP13 = PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)
   'Modify By Sindy 2022/3/14 排除收件人為CF代理人
   If strCP13 = "W2001" And InStr(strTxtCopy, "W2001") = 0 And m_isCFFagent = False Then
      'Modified by Lydia 2024/05/14
      'strTxtCopy = "W2001"
      strTxtCopy = Pub_GetSpecMan("ACS郵件通知主管")
   End If
   '2022/1/25 END
   '有指定副本收件人
   If strTxtCopy <> "" Then
      'Modify By Sindy 2022/1/21 W2001客戶之案件，不管任何人操作，在客戶會稿的歷程，自動預設ACS01@taie.com.tw為副本收受者。
      'Modify By Sindy 2022/3/14 排除收件人為CF代理人
      If InStr(UCase(strTxtCopy), UCase("W2001")) > 0 And m_isCFFagent = False Then
         'Modified by Lydia 2024/05/14
         'strTxtCopy = Replace(strTxtCopy, UCase("W2001"), "ACS01@taie.com.tw")
         strTxtCopy = Replace(strTxtCopy, UCase("W2001"), Pub_GetSpecMan("ACS郵件通知主管"))
      End If
      '2022/1/21 END
      varTemp = Split(strTxtCopy, ";")
      For ii = LBound(varTemp) To UBound(varTemp)
         If varTemp(ii) <> "" Then
            'Add By Sindy 2022/1/21
            If GetPrjSalesNM(CStr(varTemp(ii))) <> "" Then
            '2022/1/21 END
               strTemp = varTemp(ii) & " (" & GetPrjSalesNM(CStr(varTemp(ii))) & ")"
            Else
               strTemp = varTemp(ii)
            End If
            lstMailBox.AddItem strTemp
            txtCopy = txtCopy & strTemp & "; "
         End If
      Next
   Else
   '2018/5/11 END
      If pNoCopy = False Then
         txtCopy = strUserNum & " (" & strUserName & ")"
         lstMailBox.AddItem txtCopy
         txtCopy = txtCopy & "; "
      End If
   End If
   Set rsQuery = Nothing
End Sub

Private Sub mnuFn_Click(Index As Integer)
   '開啟
   If Index = 0 Then
      If m_selText <> "" Then fnOpenAttFile
   '刪除
   ElseIf Index = 1 Then
      If m_selText <> "" Then fnClearAttFile
   'Added by Morgan 2024/8/13
   '下載
   ElseIf Index = 4 Then
      If m_selText <> "" Then fnDownAttFile
   '下載全部
   ElseIf Index = 5 Then
      If m_stFiles <> "" Then fnDownAttFile True
   '壓縮加密
   ElseIf Index = 7 Then
     If m_stFiles <> "" Then fnZipAttFile
   'end 2024/8/13
   End If
End Sub

Private Sub fnClearAttFile()
   Dim stFiles As String, intIdx As Integer, ii As Integer
   Dim files() As String
   
   If txtAttachment.SelStart > 0 Then
      stFiles = Left(txtAttachment.Text, txtAttachment.SelStart)
      files = Split(stFiles, ";")
      intIdx = UBound(files)
   End If
   
   files = Split(m_stFiles, ";")
   If txtAttachment.SelStart = 0 Then
      intIdx = LBound(files)
   End If
   
   m_stFiles = ""
   txtAttachment = ""
   For ii = LBound(files) To UBound(files)
      If ii <> intIdx And files(ii) <> "" Then
         m_stFiles = m_stFiles & files(ii) & ";"
         txtAttachment = txtAttachment & GetFileDesc(files(ii)) & ";" & vbCrLf
      End If
   Next
   'txtAttachment.SelText = ""
   
   'If Mid(txtAttachment, 1, 2) = vbCrLf Then
   '   txtAttachment = Mid(txtAttachment, 3)
   'End If
End Sub

'Added by Morgan 2014/12/25
Private Sub fnOpenAttFile()
   Dim hLocalFile As Long
   Dim stFiles As String, intIdx As Integer
   Dim files() As String
   
   If cmdAttach.Visible = False Then Exit Sub 'Add By Sindy 2018/9/6
   
   If m_selText <> "" Then
      If m_selStart > 0 Then
         stFiles = Left(txtAttachment.Text, m_selStart)
         files = Split(stFiles, ";")
         intIdx = UBound(files)
      End If
      
      files = Split(m_stFiles, ";")
      If m_selStart = 0 Then
         intIdx = LBound(files)
      End If
      
      If files(intIdx) <> "" Then
         ShellExecute hLocalFile, "open", files(intIdx), vbNullString, vbNullString, 1
      End If
      txtAttachment.SelStart = m_selStart
      txtAttachment.SelLength = Len(m_selText)
   End If
End Sub
'Added by Morgan 2024/8/13
'下載附件
Private Sub fnDownAttFile(Optional pAll As Boolean = False)
   Dim hLocalFile As Long
   Dim stFiles As String, intIdx As Integer
   Dim files() As String
   Dim stFolderPath As String, stFullName As String
   Dim bolDone As Boolean
   
   If cmdAttach.Visible = False Then Exit Sub
   
   '讀取前次設定路徑
   stFolderPath = GetSetting("TAIE", "P", UCase(Me.Name) & "Dir", "")
   If stFolderPath <> "" Then
      If PUB_ChkDir(stFolderPath) = False Then
         stFolderPath = PUB_Getdesktop
      End If
   Else
      stFolderPath = PUB_Getdesktop
   End If
   stFolderPath = PUB_GetFolder(Me.hWnd, stFolderPath, "請選取下載檔案要存放的資料夾:")
   If Trim(stFolderPath) <> "" Then 'they did not hit cancel
      SaveSetting "TAIE", "P", UCase(Me.Name) & "Dir", stFolderPath
   Else
      Exit Sub
   End If
   If Right(Trim(stFolderPath), 1) <> "\" Then
      stFolderPath = Trim(stFolderPath) & "\"
   End If

   If pAll Then
      files = Split(m_stFiles, ";")
      For intIdx = LBound(files) To UBound(files)
         If files(intIdx) <> "" Then
            stFullName = stFolderPath & Mid(files(intIdx), InStrRev(files(intIdx), "\") + 1)
            If Dir(stFullName) <> "" Then
               If MsgBox("檔案【" & stFullName & "】已存在是否要覆蓋??", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                  stFullName = ""
               End If
            End If
            If stFullName <> "" Then
               FileCopy files(intIdx), stFullName
               If Dir(stFullName) <> "" Then
                  bolDone = True
               End If
            End If
         End If
      Next
      
   Else
   
      If m_selStart > 0 Then
         stFiles = Left(txtAttachment.Text, m_selStart)
         files = Split(stFiles, ";")
         intIdx = UBound(files)
      End If
      
      files = Split(m_stFiles, ";")
      If m_selStart = 0 Then
         intIdx = LBound(files)
      End If
      
      If files(intIdx) <> "" Then
         If stFolderPath <> "" Then
            stFullName = stFolderPath & Mid(files(intIdx), InStrRev(files(intIdx), "\") + 1)
            If Dir(stFullName) <> "" Then
               If MsgBox("檔案【" & stFullName & "】已存在是否要覆蓋??", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
                  stFullName = ""
               End If
            End If
            If stFullName <> "" Then
               FileCopy files(intIdx), stFullName
               If Dir(stFullName) <> "" Then
                  bolDone = True
               End If
            End If
         End If
      End If
   End If
   
   If bolDone Then
      MsgBox "下載完成!!", vbInformation
   End If
   
End Sub

'Added by Morgan 2024/8/13
'壓縮加密
Private Sub fnZipAttFile()
   Dim hLocalFile As Long
   Dim stFiles As String, intIdx As Integer
   Dim files() As String
   Dim stFolderPath As String, stFullName As String
   Dim bolDone As Boolean
   Dim stTMPFilePath As String, stTMPZip As String, stZipFileName As String, stZipPwd As String
   
   If cmdAttach.Visible = False Then Exit Sub
   
   stTMPZip = m_AttachPath & "\$TEMP.ZIP"
   If Dir(stTMPZip) <> "" Then
      Kill stTMPZip
   End If
      
   stTMPFilePath = m_AttachPath & "\$ZIPTEMP"
   If Dir(stTMPFilePath, vbDirectory) = "" Then
      MkDir stTMPFilePath
   Else
      If Dir(stTMPFilePath & "\*.*") <> "" Then
         Kill stTMPFilePath & "\*.*"
      End If
   End If
   
   
   '讀取前次設定路徑
   stFolderPath = GetSetting("TAIE", "P", UCase(Me.Name) & "Dir2", "")
   If stFolderPath <> "" Then
      If PUB_ChkDir(stFolderPath) = False Then
         stFolderPath = PUB_Getdesktop
      End If
   Else
      stFolderPath = PUB_Getdesktop
   End If
   stFolderPath = PUB_GetFolder(Me.hWnd, stFolderPath, "請選取壓縮檔要存放的資料夾:")
   If Trim(stFolderPath) <> "" Then 'they did not hit cancel
      SaveSetting "TAIE", "P", UCase(Me.Name) & "Dir2", stFolderPath
   Else
      Exit Sub
   End If
   If Right(Trim(stFolderPath), 1) <> "\" Then
      stFolderPath = Trim(stFolderPath) & "\"
   End If
   'Modified by Morgan 2024/8/28 +客戶案件案號內的符號轉全形
   stZipFileName = IIf(m_CustCaseNo <> "", PUB_FilterEFileSymbol(m_CustCaseNo) & ".", "") & PUB_CaseNo2FileName(m_CP01, m_CP02, m_CP03, m_CP04) & "." & m_CP10 & ".ZIP"
   stZipPwd = InputBox("請輸入壓縮檔的密碼！" & vbCrLf & vbCrLf & vbCrLf & "若不須密碼可直接按確定或取消。", "壓縮檔:【" & stZipFileName & "】")
   
   stZipFileName = stFolderPath & stZipFileName
   
   If Dir(stZipFileName) <> "" Then
      If MsgBox("檔案【" & stZipFileName & "】已存在是否要覆蓋??", vbYesNo + vbDefaultButton2 + vbExclamation) = vbNo Then
         stFullName = ""
      End If
   End If
   
   files = Split(m_stFiles, ";")
   For intIdx = LBound(files) To UBound(files)
      If files(intIdx) <> "" Then
         stFullName = stTMPFilePath & Mid(files(intIdx), InStrRev(files(intIdx), "\"))
         FileCopy files(intIdx), stFullName
         SetAttr stFullName, vbNormal
      End If
   Next
   
   'Modified by Morgan 2024/8/14 中文檔名無法加密，要先壓到暫存再更名
   If PUB_ZipFile(stTMPFilePath & "\*.*", stTMPZip, stZipPwd) = True Then
      FileCopy stTMPZip, stZipFileName
      m_stFiles = ""
      txtAttachment = ""
      AddFile stZipFileName
      MsgBox "壓縮完成!!", vbInformation
   Else
      MsgBox " 壓縮失敗!!", vbCritical
   End If
   Exit Sub
   
End Sub

'Modified by Lydia 2022/02/21
'Private Sub SelectText(pTextBox As TextBox)
Private Sub SelectText(pTextBox As Control)
   Dim iStart As Integer, iEnd As Integer
   If pTextBox.Text <> "" Then
      intI = pTextBox.SelStart
      If intI > 0 Then
         iStart = InStrRev(pTextBox.Text, ";", intI)
         iEnd = InStr(intI, pTextBox.Text, ";")
         If iEnd > iStart Then
            pTextBox.SelStart = iStart
            pTextBox.SelLength = iEnd - iStart
         End If
      End If
   End If
End Sub

Private Sub MSHFlexGrid1_Click()
   If MSHFlexGrid1.MouseRow > 0 Then
      If MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, 0) = "" Then
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, 0) = "V"
      Else
         MSHFlexGrid1.TextMatrix(MSHFlexGrid1.MouseRow, 0) = ""
      End If
   End If
End Sub

Private Sub MSHFlexGrid1_DblClick()
   Dim stFileName As String, hLocalFile As Long
   Dim stFtpPath As String 'Added by Morgan 2018/9/4
   
   If MSHFlexGrid1.MouseRow > 0 Then
      'Modified by Morgan 2018/9/4 +原始檔
      intI = MSHFlexGrid1.MouseRow
      '卷宗區
      If MSHFlexGrid1.TextMatrix(intI, 6) = "1" Then
         stFileName = MSHFlexGrid1.TextMatrix(intI, 5)
         If PUB_GetAttachFile_CPP(m_CP09, stFileName, m_AttachPath, False) Then
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
         End If
      '原始檔
      ElseIf MSHFlexGrid1.TextMatrix(intI, 6) = "2" Then
         stFileName = m_AttachPath & "\" & MSHFlexGrid1.TextMatrix(intI, 5)
         stFtpPath = MSHFlexGrid1.TextMatrix(intI, 7)
         If PUB_GetFtpFile(stFtpPath, stFileName, "CASEPAPERFILE", True) Then
            ShellExecute hLocalFile, "open", stFileName, vbNullString, vbNullString, 1
         End If
      End If
      'end 2018/9/4
   End If
End Sub

Private Sub txtAttachment_DblClick()
   If m_SMB02 <> "" Then Exit Sub 'Add By Sindy 2015/9/11
   
   If Not m_PrevForm Is Nothing Then 'Added by Morgan 2015/12/10
      If UCase(m_PrevForm.Name) = UCase("frm090202_4_1") Then Exit Sub 'Add By Sindy 2015/9/21
   End If
   
   ' Avoid the 'disabled' gray text by locking updates
   LockWindowUpdate txtAttachment.hWnd

   ' A disabled TextBox will not display a context menu
   txtAttachment.Enabled = False

   ' Give the previous line time to complete
   DoEvents

   ' Enable the control again
   txtAttachment.Enabled = True

   ' Unlock updates
   LockWindowUpdate 0&

   txtAttachment.SetFocus
   txtAttachment.SelLength = 0
   fnOpenAttFile
End Sub

Private Sub txtAttachment_KeyDown(KeyCode As Integer, Shift As Integer)
   If KeyCode = 46 Then
      If txtAttachment.SelLength > 0 Then
         mnuFn_Click 1
      End If
   End If
End Sub

Private Sub txtAttachment_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   If cmdSend.Visible = False Then Exit Sub 'Added by Morgan 2024/8/13
   
   If Button = vbRightButton Then
      ' Avoid the 'disabled' gray text by locking updates
         LockWindowUpdate txtAttachment.hWnd
   
         ' A disabled TextBox will not display a context menu
         txtAttachment.Enabled = False
   
         ' Give the previous line time to complete
         DoEvents
   
         If txtAttachment.SelLength > 0 Then
            mnuFn(0).Enabled = True
            mnuFn(1).Enabled = True
            'Added by Morgan 2024/8/13
            mnuFn(3).Visible = True
            mnuFn(4).Visible = True
            mnuFn(5).Visible = True
            mnuFn(6).Visible = True
            mnuFn(7).Visible = True
            'end 2024/8/13
         Else
            mnuFn(0).Enabled = False
            mnuFn(1).Enabled = False
            'Added by Morgan 2024/8/13
            mnuFn(3).Visible = False
            mnuFn(4).Visible = False
            If m_stFiles <> "" Then
               mnuFn(5).Visible = True
               mnuFn(6).Visible = True
               mnuFn(7).Visible = True
            Else
               mnuFn(5).Visible = False
               mnuFn(6).Visible = False
               mnuFn(7).Visible = False
            End If
            'end 2024/8/13
         End If
         
         ' Display our own context menu
         PopupMenu mnuPopUp
   
         ' Enable the control again
         txtAttachment.Enabled = True
         txtAttachment.SetFocus
         ' Unlock updates
         LockWindowUpdate 0&
         
   ElseIf Button = vbLeftButton Then
   
      SelectText txtAttachment
      m_selText = txtAttachment.SelText
      m_selStart = txtAttachment.SelStart
   End If
End Sub

Private Sub txtCopy_Click()
   SelectText txtCopy
End Sub

Private Sub txtContent_Change()
   PUB_RefreshText txtContent
End Sub

Private Sub txtCopy_Validate(Cancel As Boolean)
   If CheckMail(txtCopy) = False Then
      Cancel = True
   'Removed by Morgan 2015/3/27 開放可自行刪除
   'Else
   '   If InStr(txtCopy, strUserNum & " ") = 0 And InStr(LCase(txtCopy), LCase(strUserNum & "@taie.com.tw ")) = 0 Then
   '      txtCopy = strUserNum & " (" & strUserName & "); " & txtCopy
   '   End If
   End If
End Sub

'Add By Sindy 2018/5/14
Private Sub txtBCC_Click()
   SelectText txtBCC
End Sub
Private Sub txtBCC_Validate(Cancel As Boolean)
   If CheckMail(txtBCC) = False Then
      Cancel = True
   End If
End Sub
'2018/5/14 END

Private Sub txtReceiver_Click()
   SelectText txtReceiver
End Sub

Private Sub txtReceiver_Validate(Cancel As Boolean)
   If CheckMail(txtReceiver) = False Then
      Cancel = True
   End If
End Sub
'Modified by Lydia 2022/02/21
'Private Function CheckMail(pMailBox As TextBox) As Boolean
Private Function CheckMail(pMailBox As Control) As Boolean
   Dim ArrMail() As String
   Dim arrMailBox() As String
   Dim ii As Integer, strNew As String
   
   ArrMail = Split(pMailBox.Text, ";")
   For ii = LBound(ArrMail) To UBound(ArrMail)
      ArrMail(ii) = Trim(ArrMail(ii))
      If ArrMail(ii) <> "" Then
         arrMailBox = Split(ArrMail(ii))
         If InStr(arrMailBox(0), "@") = 0 Then
            'Added by Lydia 2020/09/11 排除backup
            If UCase(arrMailBox(0)) = "BACKUP" Then
            Else
               strExc(0) = "select st01||' ('||st02||')',st04 from staff where st01='" & UCase(arrMailBox(0)) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp("st04") = "2" Then
                     MsgBox RsTemp(0) & " 已離職請修正！", vbExclamation
                     CheckMail = False
                     Exit Function
                  Else
                     ArrMail(ii) = RsTemp(0)
                  End If
               Else
                  MsgBox "[ " & arrMailBox(0) & " ] 格式錯誤！", vbExclamation
                  CheckMail = False
                  Exit Function
               End If
            End If
         End If
         strNew = strNew & ArrMail(ii) & "; "
      End If
   Next
   CheckMail = True
   pMailBox.Text = strNew
End Function

Private Function GetMailList(pText As String) As String
   Dim ArrMail() As String
   Dim arrMailBox() As String
   Dim ii As Integer, strNew As String
   
   ArrMail = Split(pText, ";")
   For ii = LBound(ArrMail) To UBound(ArrMail)
      ArrMail(ii) = Trim(ArrMail(ii))
      If ArrMail(ii) <> "" Then
         arrMailBox = Split(ArrMail(ii))
         strNew = strNew & arrMailBox(0) & ";"
      End If
   Next
   GetMailList = strNew
End Function

'Add By Sindy 2015/9/11
'Modified by Lydia 2024/07/05 + pSendler
Public Sub SetParent(ByRef fm As Form, Optional ByVal pSendler As String = "")
   Set m_PrevForm = fm
   m_DefSendler = pSendler
End Sub

'Add By Sindy 2015/9/11
Public Function QueryData() As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strTemp As String, strCDate As String, strCTime As String
Dim strConSql As String
   
   QueryData = True
   
   Screen.MousePointer = vbHourglass
   Me.Enabled = False
   'Add By Sindy 2018/8/31
   If m_SMB11 = "" Then
      strConSql = " And smb02=" & m_SMB02 & _
                  " And smb03=" & m_SMB03
   Else
      strConSql = " And smb11=" & m_SMB11
   End If
   '2018/8/31 END
   '寄件備份
   strSql = "Select * From smailbackup" & _
            " Where smb01='" & m_CP09 & "'" & strConSql & _
            " order by smb02 desc,smb03 desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      'Modify By Sindy 2021/7/6
      If InStr("" & rsTmp.Fields("smb04"), "W2001") = 0 Then
      '2021/7/6 END
         lblSender = "" & rsTmp.Fields("smb04")
      End If
      txtReceiver = "" & rsTmp.Fields("smb05")
      txtCopy = "" & rsTmp.Fields("smb06")
      txtBCC = "" & rsTmp.Fields("smb10") 'Add By Sindy 2018/5/14
      'Modify By Sindy 2018/10/30
      'txtSubject = "" & rsTmp.Fields("smb07")
      'txtAttachment = "" & rsTmp.Fields("smb08")
      If cmdSend.Caption = "轉寄" Then
         txtSubject.Tag = "" & rsTmp.Fields("smb07")
      Else
         txtSubject = "" & rsTmp.Fields("smb07")
         txtAttachment = "" & rsTmp.Fields("smb08")
      End If
      '2018/10/30 END
      txtContent = "" & rsTmp.Fields("smb09")
      lblSendMailDt.Visible = True
      strTemp = TAIWANDATE(rsTmp.Fields("smb02"))
      strCDate = Format(strTemp, "###/##/##")
      strTemp = rsTmp.Fields("smb03")
      strCTime = Format(strTemp, "##:##:##")
      lblSendMailDt.Caption = "寄件日期：" & strCDate & " " & strCTime
      If "" & rsTmp.Fields("smb11") <> "" Then
         Me.Caption = Me.Caption & " (歷程序號:" & rsTmp.Fields("smb11") & ")"
      End If
      rsTmp.Close
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      QueryData = False
      rsTmp.Close
      Set rsTmp = Nothing
      Unload Me
      Exit Function
   End If
   
   Screen.MousePointer = vbDefault
   Me.Enabled = True
   
EXITSUB:
   Set rsTmp = Nothing
End Function

Private Sub SSTab2_Click(PreviousTab As Integer)
   If PreviousTab = 0 And SSTab2.Tab = 1 Then
      ShowHTML
   End If
End Sub

'Added by Morgan 2018/9/5
Private Sub ShowHTML()
   Dim stContent As String, stHTML As String
   Dim strSignFile As String, iSignId As Integer, strSignHead As String, strSignBody As String, strSignAtt As String, strMhtFile As String
   'Added by Morgan 2022/3/29
   Dim bUTF8 As Boolean
   bUTF8 = True
   'end 2022/3/29
   
   WebBrowser1.Navigate "about:blank"
   DoEvents
   
   stContent = txtContent.Text
   
   'Added by Morgan 2019/11/20 專利落款改有圖,預覽用MIME(mht)格式
   iSignId = PUB_GetSignID(m_CP01)
   If m_bolPLetter And iSignId > 0 Then
      'Modified by Morgan 2022/3/24 +UTF8
      stContent = PrepText(stContent, bUTF8)
      
      stHTML = "MIME-Version: 1.0" & vbCrLf & _
         "Content-Type: multipart/related;" & vbCrLf & _
         "            type=""text/html"";" & vbCrLf & _
         "            boundary=""" & cBoundaryB & """" & vbCrLf & vbCrLf
         
      'Modified by Morgan 2022/3/24 +UTF8
      stHTML = stHTML & PUB_GetContentMIME(stContent, iSignId, , bUTF8)

'      stHTML = "MIME-Version: 1.0" & vbCrLf & _
'         "Content-Type: multipart/related;" & vbCrLf & _
'         "            type=""text/html"";" & vbCrLf & _
'         "            boundary=""" & cBoundaryB & """" & vbCrLf & vbCrLf & _
'         cDASH2 & cBoundaryB & vbCrLf & _
'         "Content-Type: text/html; charset=""Big5""" & vbCrLf & _
'         "Content-Transfer-Encoding: quoted-printable" & vbCrLf & vbCrLf & _
'         "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbCrLf & _
'         "<HTML><HEAD>" & vbCrLf & strSignHead & vbCrLf & _
'         "<STYLE>BODY {MARGIN-TOP: 0px; FONT-SIZE: 10pt; MARGIN-LEFT: 10px }</STYLE>" & vbCrLf
'
'      '簽名檔格式設定
'      strSignFile = App.path & "\$SignHead.txt"
'      If PUB_ReadDB2File(strSignFile, 60) = True Then
'         strSignHead = PUB_ReadTextFile(strSignFile)
'         strSignHead = PUB_FixFirstDot(strSignHead)
'         stHTML = stHTML & strSignHead & vbCrLf
'      End If
'
'      stHTML = stHTML & "</HEAD>" & vbCrLf
'      stHTML = stHTML & "<BODY><DIV>" & vbCrLf & stContent & vbCrLf & "</DIV>" & vbCrLf
'
'      '簽名檔內文
'      strSignFile = App.path & "\$SignBody.txt"
'      If PUB_ReadDB2File(strSignFile, 60 + iSignId) = True Then
'         strSignBody = PUB_ReadTextFile(strSignFile)
'         strSignBody = PUB_FixFirstDot(strSignBody)
'         stHTML = stHTML & strSignBody & vbCrLf
'      End If
'
'      stHTML = stHTML & "</BODY></HTML>" & vbCrLf & vbCrLf
'
'      '簽名檔圖
'      strSignFile = App.path & "\$SignAttPic1.txt"
'      If PUB_ReadDB2File(strSignFile, 57) = True Then
'         strSignAtt = PUB_ReadTextFile(strSignFile)
'         stHTML = stHTML & cDASH2 & cBoundaryB & vbCrLf
'         stHTML = stHTML & strSignAtt & vbCrLf & vbCrLf
'      End If
'
'      strSignFile = App.path & "\$SignAttPic2.txt"
'      If PUB_ReadDB2File(strSignFile, 58) = True Then
'         strSignAtt = PUB_ReadTextFile(strSignFile)
'         stHTML = stHTML & cDASH2 & cBoundaryB & vbCrLf
'         stHTML = stHTML & strSignAtt & vbCrLf & vbCrLf
'      End If
'
'      stHTML = stHTML & cDASH2 & cBoundaryB & cDASH2 & vbCrLf & vbCrLf
      
      
      strMhtFile = App.path & "\$Letter.mht"
      If SaveTxt(stHTML, strMhtFile) = True Then
         WebBrowser1.Navigate strMhtFile
      End If
   Else
   'end 2019/11/20
   
      stContent = PUB_Text2Html(stContent)
      stHTML = "<!DOCTYPE HTML PUBLIC ""-//W3C//DTD HTML 4.0 Transitional//EN"">" & vbCrLf & _
               "<HTML><HEAD>" & vbCrLf & _
               "<STYLE>BODY {MARGIN-TOP: 0px; FONT-SIZE: 10pt; MARGIN-LEFT: 10px }</STYLE>" & vbCrLf & _
               "</HEAD>" & vbCrLf & _
               "<BODY><DIV>" & vbCrLf & stContent & vbCrLf & _
               "</DIV></BODY></HTML>"
      WebBrowser1.Document.Write (stHTML)
      
   End If 'Added by Morgan 2019/11/20
   'WebBrowser1.Refresh
End Sub

'Added by Morgan 2019/11/20
Private Function SaveTxt(pText As String, pFilePath As String) As Boolean
   Dim fN As Integer, arrTxt() As String, ii As Integer
   
On Error GoTo ErrHnd

   fN = FreeFile
   Open pFilePath For Output As fN
   arrTxt = Split(pText, vbCrLf)
   For ii = LBound(arrTxt) To UBound(arrTxt)
      Print #fN, arrTxt(ii)
   Next
   Close #fN

   SaveTxt = True
   
ErrHnd:
   If Err.NUMBER <> 0 Then MsgBox Err.Description, vbExclamation
   If fN <> 0 Then Close #fN
End Function
'Added by Morgan 2019/11/20
'本文加字型的 Tag
Private Function AddFontTag(pText As String, pSId As Integer) As String
   '中文
   If pSId = 1 Then
      '標楷體 14pt
      'AddFontTag = "<span style=""font-size:14.0pt;font-family:標楷體"">" & vbCrLf & pText & vbCrLf & "&nbsp;<span>"
      'AddFontTag = "<span style=3D""font-size:14.0pt;font-family:=BC=D0=B7=A2=C5=E9"">" & vbCrLf & pText & vbCrLf & "&nbsp;</span>"
      AddFontTag = "<span style=""font-size:14.0pt;font-family:標楷體"">" & vbCrLf & pText & vbCrLf & "&nbsp;</span>"
   '英文
   Else
      'Times New Roman 12pt
      AddFontTag = "<span style=""font-size:12pt;font-family:&quot;Times New Roman&quot;,&quot;serif&quot;"">" & vbCrLf & pText & vbCrLf & "&nbsp;</span>"
   End If
End Function

'Added by Lydia 2020/08/17 回傳MailCache語法
Private Sub ProcMailCache()
   m_AddMailCache = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                        "values('" & strUserNum & "','" & GetMailList(txtReceiver.Text) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                        ",'" & ChgSQL(txtSubject) & "','" & ChgSQL(txtContent) & "'," & CNULL(GetMailList(txtCopy.Text)) & ")"
End Sub

'Added by Morgan 2022/5/3
'Modified by Morgan 2023/12/19 配合IDS附件較大改為30M--郭
'檢查附件大小是否超過 20MB
Private Function ChkAttSize(pAtt As String) As Boolean
   Dim files() As String
   Dim fs
   Dim lngAttSize As Long, ii As Integer
   
   If pAtt <> "" Then
      files = Split(pAtt, cAST)
      lngAttSize = 0
      Set fs = CreateObject("Scripting.FileSystemObject")
      For ii = LBound(files) To UBound(files)
         If files(ii) <> "" Then
            lngAttSize = lngAttSize + fs.GetFile(files(ii)).Size
         End If
      Next
      
      If lngAttSize > 30& * 1024 * 1024 Then
         MsgBox "附件大小(" & Format(lngAttSize / 1024&, "###,###") & " KB)超過 30MB，寄信取消！", vbCritical, "寄信錯誤"
         Exit Function
      End If
   End If
   ChkAttSize = True
End Function
