VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form Frmacc44j0_1 
   AutoRedraw      =   -1  'True
   Caption         =   "說明"
   ClientHeight    =   4110
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6585
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   4110
   ScaleWidth      =   6585
   Begin TabDlg.SSTab SSTab1 
      Height          =   3900
      Left            =   50
      TabIndex        =   0
      Top             =   100
      Width           =   6400
      _ExtentX        =   11298
      _ExtentY        =   6879
      _Version        =   393216
      Tabs            =   5
      Tab             =   4
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "104年傳票"
      TabPicture(0)   =   "Frmacc44j0_1.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Image1"
      Tab(0).Control(1)=   "Label1(8)"
      Tab(0).Control(2)=   "Label1(7)"
      Tab(0).Control(3)=   "Label1(6)"
      Tab(0).Control(4)=   "Label1(5)"
      Tab(0).Control(5)=   "Label1(4)"
      Tab(0).Control(6)=   "Label1(3)"
      Tab(0).Control(7)=   "Label1(2)"
      Tab(0).Control(8)=   "Label1(1)"
      Tab(0).Control(9)=   "Label1(0)"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "105年傳票"
      TabPicture(1)   =   "Frmacc44j0_1.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label1(12)"
      Tab(1).Control(1)=   "Label1(11)"
      Tab(1).Control(2)=   "Label1(10)"
      Tab(1).Control(3)=   "Label1(9)"
      Tab(1).Control(4)=   "Label2(0)"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "106年傳票"
      TabPicture(2)   =   "Frmacc44j0_1.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label1(17)"
      Tab(2).Control(1)=   "Label2(2)"
      Tab(2).Control(2)=   "Label1(16)"
      Tab(2).Control(3)=   "Label2(1)"
      Tab(2).Control(4)=   "Label1(15)"
      Tab(2).Control(5)=   "Label1(14)"
      Tab(2).ControlCount=   6
      TabCaption(3)   =   "111年傳票"
      TabPicture(3)   =   "Frmacc44j0_1.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "Label1(18)"
      Tab(3).Control(1)=   "Label1(13)"
      Tab(3).ControlCount=   2
      TabCaption(4)   =   "112年傳票"
      TabPicture(4)   =   "Frmacc44j0_1.frx":0070
      Tab(4).ControlEnabled=   -1  'True
      Tab(4).Control(0)=   "Label1(19)"
      Tab(4).Control(0).Enabled=   0   'False
      Tab(4).Control(1)=   "Label1(20)"
      Tab(4).Control(1).Enabled=   0   'False
      Tab(4).Control(2)=   "Label1(21)"
      Tab(4).Control(2).Enabled=   0   'False
      Tab(4).Control(3)=   "Label1(22)"
      Tab(4).Control(3).Enabled=   0   'False
      Tab(4).ControlCount=   4
      Begin VB.Label Label1 
         Caption         =   "　 （J公司D112030094及95 轉至 1公司D112032089及79）               導致若只抓 11203月2492 貸方合計不等於 11203月收款"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   600
         Index           =   22
         Left            =   120
         TabIndex        =   26
         Top             =   1680
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "　 ACS-000085及000091,110年收款時,只做 J公司2492,未做收款      於11203月從 J公司調至 1公司"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   21
         Left            =   90
         TabIndex        =   25
         Top             =   1200
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "● M0101 會計科目 2492 只抓1公司原因："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   20
         Left            =   60
         TabIndex        =   24
         Top             =   900
         Width           =   4380
      End
      Begin VB.Label Label1 
         Caption         =   "● M0101 資料於112年01月起正式啟用"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   19
         Left            =   60
         TabIndex        =   23
         Top             =   450
         Width           =   4380
      End
      Begin VB.Label Label1 
         Caption         =   "　 D111041155~D111041157為自行輸入之實績保留由顧服組轉      至其他個人。"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Index           =   18
         Left            =   -74940
         TabIndex        =   22
         Top             =   810
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "● 傳票資料導致報表全年與各月不合："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   13
         Left            =   -74940
         TabIndex        =   21
         Top             =   450
         Width           =   4380
      End
      Begin VB.Label Label1 
         Caption         =   "　 2.高國碩及陳頌恩10月由中一區調中二區"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   250
         Index           =   17
         Left            =   -74880
         TabIndex        =   20
         Top             =   2160
         Width           =   6000
      End
      Begin VB.Label Label2 
         Caption         =   "(會與系統產生之報表  全所當月實績及期末實績不合，不影響整年       度報表)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.75
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   495
         Index           =   2
         Left            =   -74520
         TabIndex        =   19
         Top             =   1560
         Width           =   5800
      End
      Begin VB.Label Label1 
         Caption         =   " 　  D106033096及 D106040039 瑞婷於3月份將實績轉為點數保        留，自行修改存留之報表"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Index           =   16
         Left            =   -74880
         TabIndex        =   18
         Top             =   1080
         Width           =   6000
      End
      Begin VB.Label Label2 
         Caption         =   "● 調部門："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   1
         Left            =   -74880
         TabIndex        =   17
         Top             =   360
         Width           =   4380
      End
      Begin VB.Label Label1 
         Caption         =   "　 1.李承翰4月由專利工程師調北四區"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   15
         Left            =   -74880
         TabIndex        =   16
         Top             =   720
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "　  由於高國碩及陳頌恩調區導致整年度報表與各月不合"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   14
         Left            =   -74760
         TabIndex        =   15
         Top             =   2490
         Width           =   6000
      End
      Begin VB.Image Image1 
         Height          =   135
         Left            =   -74760
         Top             =   120
         Visible         =   0   'False
         Width           =   135
      End
      Begin VB.Label Label1 
         Caption         =   $"Frmacc44j0_1.frx":008C
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   900
         Index           =   12
         Left            =   -74880
         TabIndex        =   14
         Top             =   2040
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "● 傳票資料導致報表全年與各月不合："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   11
         Left            =   -74880
         TabIndex        =   13
         Top             =   1320
         Width           =   4380
      End
      Begin VB.Label Label1 
         Caption         =   "　 1.D105040343~44 有更正2月及3月份保留及實績"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   10
         Left            =   -74880
         TabIndex        =   12
         Top             =   1680
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "　 1.D105013094 項次002 智權人員改為20011"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   9
         Left            =   -74880
         TabIndex        =   11
         Top             =   720
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "　 1.D104010569~70 蘇特助 銷退扣應收入,傳票做扣保留"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   8
         Left            =   -74880
         TabIndex        =   10
         Top             =   2520
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "　 2.D104011343~46 蘇特助 銷退扣應收入,傳票做扣保留"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   7
         Left            =   -74880
         TabIndex        =   9
         Top             =   2880
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "  因會計科目做4191,導致1月期末是錯的(1月月報表會有問題)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   300
         Index           =   6
         Left            =   -74880
         TabIndex        =   8
         Top             =   3240
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "● 傳票資料導致報表全年與各月不合："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   5
         Left            =   -74880
         TabIndex        =   7
         Top             =   2160
         Width           =   4380
      End
      Begin VB.Label Label1 
         Caption         =   "　 4.D104103363 項次004~006 智權人員由S29->20031"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   4
         Left            =   -74880
         TabIndex        =   6
         Top             =   1800
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "　 3.D104084940 項次002 智權人員由S29->20031"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   3
         Left            =   -74880
         TabIndex        =   5
         Top             =   1440
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "　 2.D104074706 項次004~006 智權人員由S29->20031"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   2
         Left            =   -74880
         TabIndex        =   4
         Top             =   1080
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "　 1.D104070122 項次002 智權人員由S29->20031 (J公司)"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         Index           =   1
         Left            =   -74880
         TabIndex        =   1
         Top             =   720
         Width           =   6000
      End
      Begin VB.Label Label1 
         Caption         =   "● 106/03/23~24修改資料(秀玲)："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   -74880
         TabIndex        =   3
         Top             =   360
         Width           =   4380
      End
      Begin VB.Label Label2 
         Caption         =   "● 106/03/23~24修改資料(秀玲)："
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.25
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   300
         Index           =   0
         Left            =   -74880
         TabIndex        =   2
         Top             =   360
         Width           =   4380
      End
   End
End
Attribute VB_Name = "Frmacc44j0_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Amy 2022/06/16 Form2.0已修改 (無需修改)
Option Explicit

Private Sub Form_Load()
    Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   PUB_InitForm Me, Me.Width, Me.Height
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strFormName = MsgText(601)
    Set Frmacc44j0_1 = Nothing
End Sub


