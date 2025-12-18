VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm140112 
   BorderStyle     =   1  '單線固定
   Caption         =   "預約作業"
   ClientHeight    =   6660
   ClientLeft      =   156
   ClientTop       =   156
   ClientWidth     =   10524
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6660
   ScaleWidth      =   10524
   Begin TabDlg.SSTab SSTab1 
      Height          =   5988
      Left            =   24
      TabIndex        =   17
      Top             =   624
      Width           =   10440
      _ExtentX        =   18415
      _ExtentY        =   10562
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   529
      TabMaxWidth     =   5292
      ForeColor       =   16711680
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "會議室查詢"
      TabPicture(0)   =   "frm140112.frx":0000
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Check1"
      Tab(0).Control(1)=   "cmdMove(0)"
      Tab(0).Control(2)=   "cmdMove(3)"
      Tab(0).Control(3)=   "Command1"
      Tab(0).Control(4)=   "cmdMove(1)"
      Tab(0).Control(5)=   "cmdMove(2)"
      Tab(0).Control(6)=   "grdDataList"
      Tab(0).Control(7)=   "Label2"
      Tab(0).Control(8)=   "Combo1"
      Tab(0).ControlCount=   9
      TabCaption(1)   =   "日期查詢"
      TabPicture(1)   =   "frm140112.frx":001C
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "Combo2"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Label1"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "Label3"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "MGrid2"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "Timer3"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdMoveD(2)"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "cmdMoveD(1)"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "Command2"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "cmdMoveD(3)"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "cmdMoveD(0)"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "Check2"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).ControlCount=   11
      Begin VB.CheckBox Check2 
         Caption         =   "含星期六、日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   132
         TabIndex        =   32
         Top             =   480
         Width           =   1545
      End
      Begin VB.CommandButton cmdMoveD 
         Caption         =   "▲▲"
         BeginProperty Font 
            Name            =   "@新細明體"
            Size            =   6.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   1812
         TabIndex        =   31
         ToolTipText     =   "上週"
         Top             =   468
         Width           =   450
      End
      Begin VB.CommandButton cmdMoveD 
         Caption         =   "▼▼"
         BeginProperty Font 
            Name            =   "@新細明體"
            Size            =   6.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   6432
         TabIndex        =   30
         ToolTipText     =   "下週"
         Top             =   468
         Width           =   450
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H0080FFFF&
         Caption         =   "今日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   120
         Style           =   1  '圖片外觀
         TabIndex        =   29
         ToolTipText     =   "點我回到今日"
         Top             =   792
         Width           =   1000
      End
      Begin VB.CommandButton cmdMoveD 
         Caption         =   "▲"
         BeginProperty Font 
            Name            =   "@新細明體"
            Size            =   6.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   2292
         TabIndex        =   28
         ToolTipText     =   "前日"
         Top             =   468
         Width           =   450
      End
      Begin VB.CommandButton cmdMoveD 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "@新細明體"
            Size            =   6.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   5928
         TabIndex        =   27
         ToolTipText     =   "次日"
         Top             =   468
         Width           =   450
      End
      Begin VB.Timer Timer3 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   3264
         Top             =   72
      End
      Begin VB.CheckBox Check1 
         Caption         =   "含星期六、日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   -74868
         TabIndex        =   23
         Top             =   492
         Width           =   1545
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "▲▲"
         BeginProperty Font 
            Name            =   "@新細明體"
            Size            =   6.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   0
         Left            =   -73188
         TabIndex        =   22
         ToolTipText     =   "上個月"
         Top             =   468
         Width           =   450
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "▼▼"
         BeginProperty Font 
            Name            =   "@新細明體"
            Size            =   6.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   3
         Left            =   -66720
         TabIndex        =   21
         ToolTipText     =   "下個月"
         Top             =   468
         Width           =   450
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H0080FFFF&
         Caption         =   "本週"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9.6
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   500
         Left            =   -74868
         Style           =   1  '圖片外觀
         TabIndex        =   20
         ToolTipText     =   "點我回到今日"
         Top             =   804
         Width           =   1000
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "▲"
         BeginProperty Font 
            Name            =   "@新細明體"
            Size            =   6.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   1
         Left            =   -72708
         TabIndex        =   19
         ToolTipText     =   "上週"
         Top             =   468
         Width           =   450
      End
      Begin VB.CommandButton cmdMove 
         Caption         =   "▼"
         BeginProperty Font 
            Name            =   "@新細明體"
            Size            =   6.6
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   285
         Index           =   2
         Left            =   -67212
         TabIndex        =   18
         ToolTipText     =   "下週"
         Top             =   468
         Width           =   450
      End
      Begin VB.Timer Timer1 
         Interval        =   1000
         Left            =   48
         Top             =   -72
      End
      Begin VB.Timer Timer2 
         Enabled         =   0   'False
         Interval        =   50
         Left            =   432
         Top             =   -24
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdDataList 
         Height          =   5030
         Left            =   -74916
         TabIndex        =   24
         Top             =   780
         Width           =   8688
         _ExtentX        =   15325
         _ExtentY        =   8869
         _Version        =   393216
         BackColor       =   -2147483624
         Rows            =   49
         Cols            =   7
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid MGrid2 
         Height          =   5028
         Left            =   84
         TabIndex        =   33
         Top             =   780
         Width           =   10224
         _ExtentX        =   18034
         _ExtentY        =   8869
         _Version        =   393216
         BackColor       =   -2147483624
         Rows            =   49
         Cols            =   7
         AllowBigSelection=   0   'False
         ScrollTrack     =   -1  'True
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
         _Band(0).Cols   =   7
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "預約日期可以自行輸入民國年月日"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   204
         Left            =   7032
         TabIndex        =   36
         Top             =   480
         Width           =   3240
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "預約"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   2832
         TabIndex        =   35
         Top             =   504
         Width           =   456
      End
      Begin MSForms.ComboBox Combo2 
         Height          =   348
         Left            =   3336
         TabIndex        =   34
         Top             =   432
         Width           =   2568
         VariousPropertyBits=   545343515
         DisplayStyle    =   3
         Size            =   "4530;614"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體"
         FontHeight      =   240
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "預約"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   11.4
            Charset         =   136
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   228
         Left            =   -71616
         TabIndex        =   26
         Top             =   480
         Width           =   456
      End
      Begin MSForms.ComboBox Combo1 
         Height          =   348
         Left            =   -71088
         TabIndex        =   25
         Top             =   432
         Width           =   3840
         VariousPropertyBits=   545343515
         DisplayStyle    =   7
         Size            =   "6773;529"
         MatchEntry      =   1
         ShowDropButtonWhen=   2
         FontName        =   "新細明體-ExtB"
         FontHeight      =   195
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
   End
   Begin VB.Frame frmColorSample 
      BorderStyle     =   0  '沒有框線
      Caption         =   "Frame1"
      Height          =   525
      Left            =   60
      TabIndex        =   5
      Top             =   0
      Width           =   4000
      Begin VB.Label lblColor 
         Appearance      =   0  '平面
         BorderStyle     =   1  '單線固定
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   6
         Left            =   2580
         TabIndex        =   16
         Top             =   60
         Width           =   195
      End
      Begin VB.Label lblColorDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "教育訓練預約"
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
         Index           =   3
         Left            =   2820
         TabIndex        =   15
         Top             =   60
         Width           =   1170
      End
      Begin VB.Label lblColorDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "非上班時段"
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
         Index           =   4
         Left            =   1545
         TabIndex        =   14
         Top             =   60
         Width           =   975
      End
      Begin VB.Label lblColor 
         Appearance      =   0  '平面
         BorderStyle     =   1  '單線固定
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   4
         Left            =   1275
         TabIndex        =   13
         Top             =   60
         Width           =   195
      End
      Begin VB.Label lblColor 
         Appearance      =   0  '平面
         BorderStyle     =   1  '單線固定
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   3
         Left            =   1275
         TabIndex        =   12
         Top             =   300
         Width           =   195
      End
      Begin VB.Label lblColorDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "已被預約"
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
         Index           =   2
         Left            =   1740
         TabIndex        =   11
         Top             =   300
         Width           =   780
      End
      Begin VB.Label lblColor 
         Appearance      =   0  '平面
         BorderStyle     =   1  '單線固定
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   2
         Left            =   1500
         TabIndex        =   10
         Top             =   300
         Width           =   195
      End
      Begin VB.Label lblColorDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "目前點選"
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
         Index           =   1
         Left            =   255
         TabIndex        =   9
         Top             =   60
         Width           =   780
      End
      Begin VB.Label lblColor 
         Appearance      =   0  '平面
         BorderStyle     =   1  '單線固定
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   1
         Left            =   45
         TabIndex        =   8
         Top             =   60
         Width           =   195
      End
      Begin VB.Label lblColorDesc 
         AutoSize        =   -1  'True
         BackStyle       =   0  '透明
         Caption         =   "週期性預約"
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
         Index           =   0
         Left            =   255
         TabIndex        =   7
         Top             =   300
         Width           =   975
      End
      Begin VB.Label lblColor 
         Appearance      =   0  '平面
         BorderStyle     =   1  '單線固定
         ForeColor       =   &H80000008&
         Height          =   195
         Index           =   0
         Left            =   45
         TabIndex        =   6
         Top             =   300
         Width           =   195
      End
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "新增"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   0
      Left            =   4140
      TabIndex        =   4
      Top             =   90
      Width           =   910
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "修改"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   1
      Left            =   5070
      TabIndex        =   3
      Top             =   90
      Width           =   910
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "刪除"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   2
      Left            =   6000
      TabIndex        =   2
      Top             =   90
      Width           =   910
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "檢視"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   3
      Left            =   6930
      TabIndex        =   1
      Top             =   90
      Width           =   910
   End
   Begin VB.CommandButton cmdFunc 
      Caption         =   "結束"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Index           =   4
      Left            =   7920
      TabIndex        =   0
      Top             =   90
      Width           =   960
   End
End
Attribute VB_Name = "frm140112"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo by Lydia 2025/01/09 增加以日期查詢：目前只有顯示會議室，其他預約有寫相關程式但未顯示
'Memo by Morgan 2021/4/22 改成Form2.0 (grdDataList)
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Morgan 2011/5/16
Option Explicit

Dim m_selRow As Long, m_selCol As Long, m_bolFree As Boolean
Dim m_selColor As Long '點選顏色
Dim m_ocpColor As Long '已預約顏色
Dim m_ocpColor2 As Long '已預約顏色2
Dim m_ocpColor3 As Long '週期性預約顏色
Dim m_ocpColor4 As Long '教育訓練預約顏色 'Add by Amy 2019/12/09

Dim m_offColor As Long '例假日顏色
Dim m_dftColor As Long '預設顏色
Dim m_dftColor2 As Long '預設顏色2
Dim m_StartDate As String
Dim m_arrWorkDay(4) As Boolean '是否工作日
Dim m_SelDate As String '選取日期
Dim m_bolRetrievalSys As Boolean '檢索系統維護權限 Added by Morgan 2015/5/4
Dim m_bolSalesAndCar As Boolean '是否智權人員借車 Added by Morgan 2015/8/14
Dim stUsers As String '公務車可登記人員 Added by Morgan 2019/7/12
'Add by Amy 2019/11/12
Public m_SN01 As String, m_Title As String, m_SN12 As String '教育訓練編號/教育訓練標題/教育訓練建立者
'Add by Amy 2019/12/09
Dim bolShowBlock As Boolean  '教育訓練進入有預約,show 閃爍
Public stOldRR1 As String, stOldRR2 As String, stOldRR3 As String, stOldRR4 As String '教育訓練會議室/日期/時間起/迄
Dim intGrdR As Integer, intGrdC As Integer '教育訓練
'Add by Amy 2020/01/14 前畫面查詢
Public bolReadOnly As Boolean
Dim strNowRR2 As String '新增後日期
Dim strCombo1ItemData As String 'Added by Morgan 2021/4/22
'Added by Lydia 2025/01/09
Dim m_MeetCnt As Integer '會議室數量
Dim m_MeetNo() As String '會議室編號MR01
Dim m_MeetType() As String '1=會議室/2=其他(公務車...etc)
Dim m_StdDate2 As String, m_PassStd As String '日期查詢的條件日期
Dim m_Sel2Date As String '選取日期
Dim m_sel2Row As Long, m_sel2Col As Long, m_bol2Free As Boolean
Dim m_idX As Integer '查詢條件:1=會議室,2=日期
Dim bolWorkDay2 As Boolean '是否為工作天


Private Sub Check1_Click()
   If Check1.Value = vbChecked Then
      cmdMove(1).ToolTipText = "前5天"
      cmdMove(2).ToolTipText = "後5天"
   Else
      cmdMove(1).ToolTipText = "上週"
      cmdMove(2).ToolTipText = "下週"
   End If
   RefreshGridData
End Sub

Private Sub cmdFunc_Click(Index As Integer)
   'Added by Lydia 2025/01/09
   If m_idX = 2 Then
      Select Case Index
      Case 0
         ShowNextForm2 0
      Case 1
         ShowNextForm2 1
      Case 2
         ShowNextForm2 2
      Case 3
         ShowNextForm2 3
      Case 4
         Unload Me
      End Select
   Else
   'end 2025/01/09
      Select Case Index
      Case 0
         ShowNextForm 0
      Case 1
         ShowNextForm 1
      Case 2
         ShowNextForm 2
      Case 3
         ShowNextForm 3
      Case 4
         Unload Me
      End Select
   End If
End Sub

Public Sub cmdMove_Click(Index As Integer) 'Modify by Amy 2020/01/14 原:Private
   RefreshGridData Index + 1
End Sub

Private Sub Combo1_Click()
   If Combo1.ListIndex <> Val(Combo1.Tag) Then
      RefreshGridData
      Combo1.Tag = Combo1.ListIndex
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   SetDataListWidth
   SetRoomChoice
   'Add by Amy 2019/12/09 帶入教育訓練資料
   If m_SN01 <> MsgText(601) Then
        '會議室
        Combo1.ListIndex = Val(stOldRR1) - 1
        Combo1.Tag = Combo1.ListIndex
        '日期
        If stOldRR2 <> MsgText(601) Then
            m_StartDate = stOldRR2
        End If
   End If
   'end 2019/12/09
   RefreshGridData
   
   'modify by sonia 2019/6/10 專利郭經理要求開放所有人可預約IPTECT
'   m_bolRetrievalSys = False 'Added by Morgan 2015/5/4
'   'Added by Morgan 2013/9/12
'   'modify by sonia 2015/4/30 開放閻副所長81040及電腦中心人員
'   'modify by sonia 2015/5/1  林總指示再開放外專工程師
'   'If Left(Pub_StrUserSt03, 2) = "P1" Then
'   'Modified by Morgan 2016/4/21 開放智權人員也能預約--文雄
'   If Left(Pub_StrUserSt03, 2) = "P1" Or Pub_StrUserSt03 = "M51" Or Left(Pub_StrUserSt03, 1) = "S" Or Pub_StrUserSt03 = "F21" Or strUserNum = "81040" Or strUserNum = "69009" Then
'      'Modified by Morgan 2015/5/4 改所有人都能讀,有權限的才能維護
'      'Me.Caption = "會議室/檢索系統預約作業"
'      m_bolRetrievalSys = True
'      'end 2015/5/4
'   End If
'   'end 2013/9/12
   m_bolRetrievalSys = True
   'end 2019/6/10
   Command1.Caption = Command1.Caption & vbCrLf & Format(strSrvDate(2), "###/##/##")
   'Added by Lydia 2025/01/09
   If m_SN01 = "" Then
      Command2.Caption = Command2.Caption & vbCrLf & Format(strSrvDate(2), "###/##/##")
      RefreshGridData2
      Combo2.Tag = Combo2.Text
   Else
      Me.SSTab1.TabVisible(1) = False  '先隱藏
   End If
   Me.SSTab1.Tab = 0
   'end 2025/01/09
End Sub

Private Sub SetRoomChoice()
   'Modified by Morgan 2019/7/12 stUsers 改全域變數
   'Dim stCon As String, stUsers As String
   Dim stCon As String
      
   'Modified by Morgan 2015/8/13 +總務或智權部可借用公務車
   'Modified by Morgan 2018/4/10 特殊人員改抓設定
   'If Pub_StrUserSt03 = "M51" Or strUserNum = "A4023" Or strUserNum = "94007" Or Pub_StrUserSt03 = "M11" Or Pub_StrUserSt03 = "M10" Or Left(Pub_StrUserSt15, 1) = "S" Then
   stUsers = Pub_GetSpecMan("公務車可登記人員")
   'Modified by Morgan 2019/7/12 部門也抓設定(+客戶服務組 W10--文雄)
   'If Pub_StrUserSt03 = "M51" Or InStr(stUsers, strUserNum) > 0 Or Pub_StrUserSt03 = "M11" Or Pub_StrUserSt03 = "M10" Or Left(Pub_StrUserSt15, 1) = "S" Then
   If Pub_StrUserSt03 = "M51" Or InStr(stUsers, strUserNum) > 0 Or InStr(stUsers, Pub_StrUserSt15) > 0 Or Left(Pub_StrUserSt15, 1) = "S" Then
   'end 2019/7/12
   
   'Modified by Morgan 2019/12/16 是否可預約改判斷 mr06='Y'
   '   stCon = " or mr01=9"
   Else
      stCon = " and mr01<>9"
   'end 2019/12/16
   
   End If
   
   'Modify by Morgan 2011/5/16 目前只開放 5f 及 中型
   'Modified by Morgan 2013/9/12 +專利處加檢索系統
   'modify by sonia 2015/4/30 開放閻副所長81040及電腦中心人員
   'modify by sonia 2015/5/1  林總指示再開放外專工程師
   'If Left(Pub_StrUserSt03, 2) = "P1" Then
   'Modified by Morgan 2015/5/4 改所有人都能讀,有權限的才能維護
   'If Left(Pub_StrUserSt03, 2) = "P1" Or Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "F21" Or strUserNum = "81040" Then
   'Modified by Morgan 2017/10/20 +mr01>=10
      'Modified by Morgan 2019/3/15 +排序mr05
      'Modified by Morgan 2019/12/16 是否可預約改判斷 mr06='Y'
      'strExc(0) = "select * from meetingroom where mr01<3 or mr01=8 Or mr01 >= 10 " & stCon & " order by nvl(mr05,mr01) desc"
      strExc(0) = "select * from meetingroom where mr06='Y' " & stCon & " order by nvl(mr05,mr01) desc"
      'end 2019/12/16
   'Else
   '   strExc(0) = "select * from meetingroom where mr01<3 order by mr01 desc"
   'End If
   'end 2015/5/4
   'end 2013/9/12
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      Do While Not .EOF
         Combo1.AddItem .Fields("mr02") & " " & .Fields("mr03"), 0
         'Modified by Morgan 2021/4/22
         'Combo1.ITEMDATA(0) = .Fields("mr01")
         strCombo1ItemData = .Fields("mr01") & "," & strCombo1ItemData
         'end 2021/4/22
         .MoveNext
      Loop
      End With
      Combo1.ListIndex = 0
      Combo1.Tag = Combo1.ListIndex
   End If
   
   Call SetCombo2 'Added by Lydia 2025/01/09
End Sub

Private Sub Command1_Click()
   Screen.MousePointer = vbHourglass
   m_StartDate = strSrvDate(1)
   RefreshGridData
   Screen.MousePointer = vbDefault
End Sub

Private Function GetDateStr(pDate As Date) As String
   Select Case Weekday(pDate)
      Case 1
         GetDateStr = Format(pDate, "M/D") & "(日)"
      Case 2
         GetDateStr = Format(pDate, "M/D") & "(一)"
      Case 3
         GetDateStr = Format(pDate, "M/D") & "(二)"
      Case 4
         GetDateStr = Format(pDate, "M/D") & "(三)"
      Case 5
         GetDateStr = Format(pDate, "M/D") & "(四)"
      Case 6
         GetDateStr = Format(pDate, "M/D") & "(五)"
      Case 7
         GetDateStr = Format(pDate, "M/D") & "(六)"
   End Select
End Function

Private Sub SetDataListWidth()
   Dim ii As Integer
   '點選
   m_selColor = RGB(&H35, &H35, &HFF)
   lblColor(1).BackColor = m_selColor
   '已預約1
   m_ocpColor = RGB(&H90, &HEE, &H90)
   lblColor(3).BackColor = m_ocpColor
   '已預約2
   m_ocpColor2 = RGB(&H0, &HFF, &HFF)
   lblColor(2).BackColor = m_ocpColor2
   '教育訓練預約 'Add by Amy 2019/12/09
   m_ocpColor4 = RGB(&HCC, &H77, &H22)
   lblColor(6).BackColor = m_ocpColor4
   '週期性預約
   m_ocpColor3 = RGB(&HDD, &HA0, &HDD)
   'm_ocpColor3 = RGB(&H90, &HEE, &H90)
   lblColor(0).BackColor = m_ocpColor3
   '非工作時段
   m_offColor = RGB(&HA9, &HA9, &HA9)
   lblColor(4).BackColor = m_offColor
   '底色
   m_dftColor = RGB(&HFF, &HFA, &HCD)
   '底色2
   m_dftColor2 = &HFFFFFF
   
   With grdDataList
      .Visible = False
      .GridColorBand(0) = RGB(&H69, &H69, &H69)
      .Clear
      .Rows = 49: .Cols = 7
      .FixedRows = 1
      .row = 0
      .RowHeight(.row) = 500
      
      .col = 0: .ColWidth(.col) = 1020
      .col = 1: .ColWidth(.col) = 400
      
      '設定日期欄位
      For ii = 0 To 4
         .col = .col + 1: .ColWidth(.col) = 1380: .CellFontBold = True: .CellFontSize = 15
      Next
      '設定時間欄位
      For ii = 0 To 23
         .row = .row + 1
         .RowHeight(.row) = 200
         .col = 0: .Text = Format(ii, "00"): .CellFontBold = True: .CellFontSize = 16
         
         .col = 1: .Text = "00": .CellFontBold = False: .CellFontSize = 7
         .CellAlignment = flexAlignLeftTop: .CellBackColor = m_dftColor
         .row = .row + 1
         .RowHeight(.row) = 200
         .col = 0: .Text = Format(ii, "00"): .CellFontBold = True
         .col = 1: .Text = "": .CellFontBold = False: .CellFontSize = 9
         .CellBackColor = m_dftColor2
      Next
      
      .MergeCol(0) = True
      .MergeCol(1) = True
      .MergeCells = flexMergeRestrictColumns
      .ColAlignmentFixed = flexAlignCenterCenter
      .TopRow = 17
      .Visible = True
   End With
   
   ResetGrid True
   
   'Added by Lydia 2025/01/09 會議室數量
   'If m_SN01 = "" Then
      strExc(0) = "select count(*) cnt from meetingroom where mr06='Y' and mr04='1'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_MeetCnt = Val("" & RsTemp.Fields("cnt"))
         ReDim m_MeetNo(0 To m_MeetCnt + 1) As String
         ReDim m_MeetType(0 To m_MeetCnt + 1) As String
      End If
      With MGrid2
         .Visible = False
         .GridColorBand(0) = RGB(&H69, &H69, &H69)
         .Clear
         .Rows = 49: .Cols = m_MeetCnt + 2
         .FixedRows = 1
         .row = 0
         .RowHeight(.row) = 500
         
         .col = 0: .ColWidth(.col) = 1020
         .col = 1: .ColWidth(.col) = 400
         
         '設定會議室欄位
         strExc(0) = "select * from meetingroom where mr06='Y' and mr04='1' order by mr05 "
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               .col = RsTemp.Fields("mr05") + 1
               .ColWidth(.col) = 1380: .CellFontBold = True: .CellFontSize = 10
               .Text = RsTemp.Fields("mr02") & vbCrLf & RsTemp.Fields("mr03")
               .WordWrap = True
               m_MeetNo(RsTemp.Fields("mr05") + 1) = "" & RsTemp.Fields("mr01")
               m_MeetType(RsTemp.Fields("mr05") + 1) = "" & RsTemp.Fields("mr04")
               RsTemp.MoveNext
            Loop
         End If
         '設定時間欄位
         For ii = 0 To 23
            .row = .row + 1
            .RowHeight(.row) = 200
            .col = 0: .Text = Format(ii, "00"): .CellFontBold = True: .CellFontSize = 16
            
            .col = 1: .Text = "00": .CellFontBold = False: .CellFontSize = 7
            .CellAlignment = flexAlignLeftTop: .CellBackColor = m_dftColor
            .row = .row + 1
            .RowHeight(.row) = 200
            .col = 0: .Text = Format(ii, "00"): .CellFontBold = True
            .col = 1: .Text = "": .CellFontBold = False: .CellFontSize = 9
            .CellBackColor = m_dftColor2
         Next
         
         .MergeCol(0) = True
         .MergeCol(1) = True
         .MergeCells = flexMergeRestrictColumns
         .ColAlignmentFixed = flexAlignCenterCenter
         .TopRow = 17
         .Visible = True
      End With
      ResetGrid2 True
      MGrid2.Visible = True
   'End If
   'end 2025/01/09
End Sub

Public Sub RefreshGridData(Optional pChoice As Integer)
   Dim strColName As String, ii As Integer, dtTmp As Date
   Dim strEndDate As String, iRow As Integer, iCol As Integer
   Dim iPos1 As Integer, iPos2 As Integer, iWday As Integer
   Dim iHalfHours As Integer, iMaxBytes As Single, iBytes As Single, strText As String
   Dim iWeekDay As Integer, stSign As String
   
   strNowRR2 = "" 'Add by Amy 2020/01/14
   If m_StartDate = "" Then
      m_StartDate = strSrvDate(1)
   End If
   If Check1.Value = 0 Then
      iWeekDay = Weekday(Format(m_StartDate, "####/##/##"))
      If iWeekDay <> 2 Then
         m_StartDate = CompDate(2, 2 - iWeekDay, m_StartDate)
      End If
   End If
   
   Select Case pChoice
      Case 1 '上個月
         m_StartDate = CompDate(1, -1, m_StartDate)
         If Check1.Value = 0 Then
            iWeekDay = Weekday(Format(m_StartDate, "####/##/##"))
            If iWeekDay <> 2 Then
               m_StartDate = CompDate(2, 2 - iWeekDay, m_StartDate)
            End If
         End If
         
      Case 2 '上週或前5天
         '上週
         If Check1.Value = 0 Then
            m_StartDate = CompDate(2, -7, m_StartDate)
         '前5天
         Else
            m_StartDate = CompDate(2, -5, m_StartDate)
         End If
      
      Case 3 '下週或後5天
         '下週
         If Check1.Value = 0 Then
            m_StartDate = CompDate(2, 7, m_StartDate)
         '後5天
         Else
            m_StartDate = CompDate(2, 5, m_StartDate)
         End If
         
      Case 4 '下個月
         m_StartDate = CompDate(1, 1, m_StartDate)
         iWeekDay = Weekday(Format(m_StartDate, "####/##/##"))
         If iWeekDay <> 2 Then
            m_StartDate = CompDate(2, 2 - iWeekDay, m_StartDate)
         End If
   End Select
   
   '清除資料
   ResetGrid , True
   
   With grdDataList
   .Visible = False
   .TextMatrix(0, 1) = ((m_StartDate \ 10000) - 1911) & "年" '年
   .WordWrap = True
   For ii = 0 To 4
      .col = 2 + ii
      .MergeCol(.col) = False
      
      strEndDate = CompDate(2, ii, m_StartDate)
      m_arrWorkDay(ii) = ChkWorkDay(strEndDate)
      
      dtTmp = Format(strEndDate, "####/##/##")
      strColName = GetDateStr(dtTmp)
      '.TextMatrix(0, iCol) = strColName
      .row = 0
      .Text = strColName
      If strEndDate = strSrvDate(1) Then
         .CellBackColor = Command1.BackColor
      Else
         .CellBackColor = .BackColorFixed
      End If
      iCol = .col
      If Combo1.ListIndex >= 0 Then
         'Modify by Amy 2019/12/09 +rr20
         strExc(0) = "select rr02,rr03,rr04,rr05,rr07,rr08,rr16,rr20 from RoomReservation" & _
            " where rr01=" & PUB_GetItemData(strCombo1ItemData, Combo1.ListIndex) & " and rr02=" & strEndDate & _
            " and rr05='N' and rr18=0" & _
            " union select rd04,rr03,rr04,rr05,rr07,rr08,rr16,rr20 from RoomResDetail,RoomReservation" & _
            " where rd01=" & PUB_GetItemData(strCombo1ItemData, Combo1.ListIndex) & " and rd04=" & strEndDate & _
            " and rd05 is null and rr18=0 and rr01(+)=rd01 and rr02(+)=rd02 and rr03(+)=rd03" & _
            " order by rr03"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            Do While Not RsTemp.EOF
               strText = "" & RsTemp.Fields("rr08")
               'Moidfy by Amy 2019/12/09 +教育訓練預約顯示※
               If Not IsNull(RsTemp("rr20")) = True Then
                  If m_SN01 <> MsgText(601) And m_SN01 = RsTemp("rr20") Then
                    '目前預約
                    'Add by Amy 2020/01/14
                    bolShowBlock = True
                    strNowRR2 = "" & RsTemp("rr02")
                    'end 2020/01/14
                    strText = "★" & strText
                    intGrdC = iCol
                  Else
                    strText = "※" & strText
                  End If
               ElseIf RsTemp("rr05") <> "N" Then
                  strText = "◎" & strText
               Else
                  If stSign = "◆" Then
                     stSign = "◇"
                  Else
                     stSign = "◆"
                  End If
                  strText = stSign & strText
               End If
               iBytes = GetTextLength(strText)
               iHalfHours = (2 * (RsTemp.Fields("rr04") \ 100) + ((RsTemp.Fields("rr04") Mod 100) / 30)) - (2 * (RsTemp.Fields("rr03") \ 100) + ((RsTemp.Fields("rr03") Mod 100) / 30))
               '除非只有一格(半小時),否則最後一行固定帶預約人
               If iHalfHours > 1 Then
                  If IsNull(RsTemp.Fields("rr07")) Then
                     iMaxBytes = 14 * iHalfHours
                     If iBytes > iMaxBytes Then
                        strText = PUB_StrToStr(strText, iMaxBytes - 3) & "..."
                     End If
                  'Modified by Morgan 2015/8/13 借車加次數
                  ElseIf RsTemp.Fields("rr16") > 0 Then
                     iMaxBytes = 14 * (iHalfHours - 1)
                     If iBytes > iMaxBytes Then
                        strText = PUB_StrToStr(strText, iMaxBytes - 3) & "..."
                     End If
                     strText = strText & vbCrLf & " .. " & GetStaffName(RsTemp.Fields("rr07"), True) & "(第" & RsTemp.Fields("rr16") & "次)"
                  'end 2015/8/13
                  Else
                     iMaxBytes = 14 * (iHalfHours - 1)
                     'Modified by Morgan 2017/12/21 配合顯示視訊訊息改以內容為主，長度超過則不顯示預約人
                     If iBytes > iMaxBytes Then
                        strText = PUB_StrToStr(strText, iMaxBytes - 3 + 14) & "..."
                     Else
                        strText = strText & vbCrLf & " ........... " & GetStaffName(RsTemp.Fields("rr07"), True)
                     End If
                     'end 2017/12/21
                  End If
               ElseIf iBytes > 14 Then
                  strText = PUB_StrToStr(strText, 14)
               End If
                  
               iPos1 = 1 + 2 * RsTemp.Fields("rr03") \ 100 + IIf(RsTemp.Fields("rr03") Mod 100 >= 30, 1, 0)
               iPos2 = 1 + 2 * RsTemp.Fields("rr04") \ 100 + IIf(RsTemp.Fields("rr04") Mod 100 >= 30, 0, -1)
               For iRow = iPos1 To iPos2
                  .row = iRow
                  .col = iCol
                  .CellFontSize = 9
                  .Text = strText
                  .CellAlignment = 1
               Next
               RsTemp.MoveNext
            Loop
         End If
      End If
      .MergeCol(.col) = True
   Next
   
   '設定顏色
   ResetGrid True
   m_selRow = 0
   m_selCol = 0
   m_SelDate = ""
   DisableFunc
   .Visible = True
   End With
   
   'Added by Morgan 2015/8/14
   'Modified by Morgan 2016/10/21 +總務(M11,M10)也比照智權人員規則
   'Modified by Morgan 2019/7/12 部門也抓設定(+客戶服務組 W10--文雄)
   'If Combo1.ITEMDATA(Combo1.ListIndex) = 9 And (Left(Pub_StrUserSt15, 1) = "S" Or Pub_StrUserSt15 = "M11" Or Pub_StrUserSt15 = "M10") Then
   If PUB_GetItemData(strCombo1ItemData, Combo1.ListIndex) = 9 And (Left(Pub_StrUserSt15, 1) = "S" Or InStr(stUsers, Pub_StrUserSt15) > 0) Then
   'end 2019/7/12
      m_bolSalesAndCar = True
   Else
      m_bolSalesAndCar = False
   End If
   'end 2015/8/14
   'Add by Amy 2019/12/09 避免捲動 造成抓的Row是錯的重抓Row
   Call GetFilkerRow
End Sub

'檢查是否為上班時間
Private Function GetDftColor() As Long
   With grdDataList
   If m_arrWorkDay(.col - 2) = False Or Val(.TextMatrix(.row, 0)) < 8 Or Val(.TextMatrix(.row, 0)) > 17 Then
      GetDftColor = m_offColor
   ElseIf .TextMatrix(.row, 1) = "00" Then
      GetDftColor = m_dftColor
   Else
      GetDftColor = m_dftColor2
   End If
   End With
End Function

Private Sub ResetGrid(Optional bolColor As Boolean, Optional bolData As Boolean)
   Dim ii As Integer, jj As Integer, strDate As String, strLstText As String
   Dim bolVisivle As Boolean
  
   Timer1.Enabled = False
   With grdDataList
      bolVisivle = .Visible
      .Visible = False
      For jj = 2 To .Cols - 1
         .col = jj
         strLstText = ""
         For ii = 1 To .Rows - 1
            .row = ii
            If bolData Then
               .Text = ""
            End If
            .CellForeColor = vbBlack
            If bolColor Then
               If .Text <> "" Then
                  If strLstText <> .Text Then
                     strLstText = .Text
                  End If
                  
                  '週期性
                  If Left(strLstText, 1) = "◎" Then
                     .CellBackColor = m_ocpColor3
                  Else
                     'Modify by Amy 2019/12/09 +教育訓練預約顯示※
                     If Left(strLstText, 1) = "※" Or Left(strLstText, 1) = "★" Then
                        .CellBackColor = m_ocpColor4
                        If strNowRR2 <> "" Then 'Added by Lydia 2025/01/09
                           'Modify by Amy 2020/01/14 + 判斷日期,避免按上下週也閃
                           If Left(strLstText, 1) = "★" And bolShowBlock = True _
                             And Mid(.TextMatrix(0, jj), 1, Val(InStr(.TextMatrix(0, jj), "(")) - 1) = Val(Mid(strNowRR2, 5, 2)) & "/" & Val(Mid(strNowRR2, 7, 2)) Then
                               Timer1.Enabled = True
                           End If
                        End If
                     ElseIf Left(strLstText, 1) = "◆" Then
                        .CellBackColor = m_ocpColor
                     Else
                        .CellBackColor = m_ocpColor2
                     End If
                  End If
               Else
                  .CellBackColor = GetDftColor
               End If
            End If
         Next
      Next
      .Visible = bolVisivle
   End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strR(1 To 4) As String, bolHasRR20 As Boolean 'Add by Amy 2019/12/09
    'Add by Amy 2019/11/12 帶回資料
    If m_SN01 <> MsgText(601) Then
        'Modify by Amy 2019/12/09 判斷與前畫面不同彈訊息(進此畫面未操作預約,會將舊資料帶回)
        bolHasRR20 = GetRR20Info(strR())
        strR(3) = IIf(Val(strR(3)) = 0, "", String(4 - Len(strR(3)), "0") & strR(3))
        strR(4) = IIf(Val(strR(4)) = 0, "", String(4 - Len(strR(4)), "0") & strR(4))
        If stOldRR1 & stOldRR2 & stOldRR3 & stOldRR4 <> strR(1) & strR(2) & strR(3) & strR(4) Then
            '教育訓練進入刪除資料就結束
            If strR(1) & strR(2) & strR(3) & strR(4) = MsgText(601) Then
                MsgBox "無會議室預約資料請確認！"
            Else
                MsgBox "會議室預約時段有修改,將覆蓋前畫面資料"
            End If
        End If
            
        If bolHasRR20 = True Then
            With frm140113
                .cboRoom = .GetMeetingRoom(strR(1), True)
                .cboRoom.Tag = strR(1)
                .MaskEdBox1.Text = ChangeWStringToWDateString(strR(2))
                .MaskEdBox1.Mask = ADFormat
                .lblWeek = GetWeekDay(CDate(.MaskEdBox1))
                .cboTime(0) = Format(strR(3), "00:00")
                .cboTime(0).Tag = .cboTime(0)
                .cboTime(1) = Format(strR(4), "00:00")
                .cboTime(1).Tag = .cboTime(1)
            End With
        End If
        m_SN01 = ""
        m_Title = ""
        m_SN12 = ""
        Me.Tag = ""
        bolShowBlock = False
        bolReadOnly = False 'Add by Amy 2020/01/14
        tmpBol = fnCancelNowFormAndShowParentForm(Me)
    Else
        Set frm140112 = Nothing
    End If
    'end 2019/11/12
End Sub

Private Sub grdDataList_DblClick()
   If m_selRow > 0 And m_selCol > 0 Then
      With grdDataList
      .row = m_selRow
      .col = m_selCol
      '點選日期
      m_SelDate = Format((Val(.TextMatrix(0, 1)) + 1911) & "/" & Mid(.TextMatrix(0, .col), 1, Len(grdDataList.TextMatrix(0, .col)) - 3), "YYYYMMDD")
      'Added by Morgan 2012/11/20 跨年問題
      If Val(.TextMatrix(0, .col)) <> Val(.TextMatrix(0, 2)) And Val(.TextMatrix(0, 2)) = 12 Then
         m_SelDate = m_SelDate + 10000
      End If
      'end 2012/11/20
      
      If m_SelDate >= strSrvDate(1) Then
         If .Text = "" Then
            If cmdFunc(0).Enabled = True Then
               m_idX = 1 'Added by Lydia 2025/01/09
               cmdFunc(0).Value = True
            End If
         Else
            If CheckRight = True Then
               ShowNextForm 1
            Else
               ShowNextForm 3
            End If
         End If
      Else
         If .Text <> "" Then
            ShowNextForm 3
         End If
      End If
      End With
   End If
End Sub

Private Sub grdDataList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim iRow As Integer
   Dim bolChange As Boolean
   Dim strDate As String
   Dim strDownText As String
  
   If x > grdDataList.ColWidth(0) + grdDataList.ColWidth(1) And y > grdDataList.RowHeight(0) Then
      With grdDataList
      .Visible = False
      If Button = 1 Or Button = 2 Then m_idX = 1 'Added by Lydia 2025/01/09
      '右鍵
      If Button = 2 Then
         SetCellActive x, y
      End If
      
      If Button = 1 Or (Button = 2 And .CellBackColor <> m_selColor) Then
      
         m_selRow = .row
         m_selCol = .col
         strDownText = .Text
            
         ResetGrid True '所有顏色還原
         
         .row = m_selRow
         .col = m_selCol
         
         '已預約
         If .Text <> "" Then
            iRow = m_selRow
            Do While iRow > 1
               iRow = iRow - 1
               If .TextMatrix(iRow, m_selCol) <> strDownText Then
                  iRow = iRow + 1
                  Exit Do
               End If
            Loop
            .row = iRow '開始
            .CellBackColor = m_selColor
            .CellForeColor = vbWhite
            
            Do While iRow < .Rows - 1
               iRow = iRow + 1
               If .TextMatrix(iRow, m_selCol) <> strDownText Then
                  iRow = iRow - 1 '結束
                  Exit Do
               End If
               .row = iRow
               .CellBackColor = m_selColor
               .CellForeColor = vbWhite
            Loop

            m_selRow = .row
            '讓 MouseMove 不動作
            If m_bolSalesAndCar Then
               m_bolFree = True
            Else
               m_bolFree = False
            End If
         '未預約
         Else
            .CellBackColor = m_selColor
            m_bolFree = True
         End If
         'SetCommand
      End If
      .Visible = True
      
      '右鍵
      If Button = 2 Then
         SetCommand 'Added by Morgan 2015/8/14
         If .CellBackColor <> GetDftColor Then
            mdiMain.mnuPopItem(0).Enabled = cmdFunc(0).Enabled
            mdiMain.mnuPopItem(1).Enabled = cmdFunc(1).Enabled
            mdiMain.mnuPopItem(2).Enabled = cmdFunc(2).Enabled
            mdiMain.mnuPopItem(3).Enabled = cmdFunc(3).Enabled
            PopupMenu mdiMain.mnuPop
         End If
      End If
      End With
   End If
End Sub

Private Sub DisableFunc()
   cmdFunc(0).Enabled = False
   cmdFunc(1).Enabled = False
   cmdFunc(2).Enabled = False
   cmdFunc(3).Enabled = False
End Sub
Private Sub SetCommand()
   Dim ii As Integer, stFromTime As String, stToTime As String, bolBooked As Boolean, bolAddable As Boolean
   
   With grdDataList
   DisableFunc
   
   'Added by Morgan 2015/5/4
   'Modified by Morgan 2015/5/4 改所有人都能讀,有權限的才能維護
   If m_bolRetrievalSys = False And PUB_GetItemData(strCombo1ItemData, Combo1.ListIndex) = 8 Then
      Exit Sub
   End If
   'end 2015/5/4
   
   If m_selCol > 1 And m_selRow > 0 Then
      .col = m_selCol
      .row = m_selRow
      m_SelDate = Format((Val(.TextMatrix(0, 1)) + 1911) & "/" & Mid(.TextMatrix(0, .col), 1, Len(grdDataList.TextMatrix(0, .col)) - 3), "YYYYMMDD")
      'Added by Morgan 2012/11/20 跨年問題
      If Val(.TextMatrix(0, .col)) <> Val(.TextMatrix(0, 2)) And Val(.TextMatrix(0, 2)) = 12 Then
         m_SelDate = m_SelDate + 10000
      End If
      'end 2012/11/20
      
      '假日不開放借車,維持以紙本方式申請
      If PUB_GetItemData(strCombo1ItemData, Combo1.ListIndex) = 9 Then
         If ChkWorkDay(m_SelDate) = False Then Exit Sub
      End If
      
      bolBooked = False
      bolAddable = False
      For ii = 1 To .Rows - 1
         .row = ii
         If .CellBackColor = m_selColor Then
            If .Text = "" Then
               bolAddable = True
            Else
               bolBooked = True
            End If
            If stFromTime = "" Then
               stFromTime = .TextMatrix(.row, 0) & IIf(.TextMatrix(.row, 1) = "00", "00", "30")
            End If
            stToTime = IIf(.TextMatrix(.row, 1) = "00", .TextMatrix(.row, 0) & "30", Format(Val(.TextMatrix(.row, 0)) + 1, "00") & "00")
         End If
      Next
      .row = m_selRow
         
      '智權人員借車
      If m_bolSalesAndCar Then
         If m_SelDate > strSrvDate(1) Or (m_SelDate >= strSrvDate(1) And Val(stToTime & "00") > ServerTime) Then
            If bolAddable Then
               If bolBooked Then
                  If CheckAddable(m_SelDate, stFromTime, stToTime) = True Then
                     cmdFunc(0).Enabled = True
                  End If
               Else
                  cmdFunc(0).Enabled = True
               End If
               
            ElseIf bolBooked Then
            
               If CheckRight = True Then
                  cmdFunc(1).Enabled = True
                  cmdFunc(2).Enabled = True
                  
               '非當日第1次優先借車檢查
               ElseIf m_SelDate > strSrvDate(1) Then
                  If CheckAddable(m_SelDate, stFromTime, stToTime) = True Then
                     cmdFunc(0).Enabled = True
                  End If
               End If
               cmdFunc(3).Enabled = True
            End If
            
         ElseIf bolBooked Then
            cmdFunc(3).Enabled = True
            
         End If
         
      Else
      'end 2015/8/13
         If .Text = "" Then
            If m_SelDate >= strSrvDate(1) Then
               'Add by Amy 2020/02/06 過去時間不可新增(因可能為教育訓練的預約會議室修改 ex:新增教育訓練->預約好會議室->改今天已過去之時間為開始時間)
               If (m_SelDate = strSrvDate(1) And Format(stFromTime & "00", "000000") >= Format(ServerTime, "000000")) Or m_SelDate > strSrvDate(1) Then
                    cmdFunc(0).Enabled = True
                End If
               'end 2020/02/06
            End If
         Else
            If bolReadOnly = False Then 'Add by Amy 2020/01/14 +if 教育訓練查詢進入,只可看
                If Pub_StrUserSt03 = "M51" Or m_SelDate > strSrvDate(1) Or (m_SelDate >= strSrvDate(1) And Val(stToTime & "00") > ServerTime) Then
                   If CheckRight = True Then
                      cmdFunc(1).Enabled = True
                      cmdFunc(2).Enabled = True
                   End If
                End If
            End If
            cmdFunc(3).Enabled = True
         End If
      End If
   End If
   End With
End Sub

'將指定座標所在的儲存格設定為作用中
Private Sub SetCellActive(Px As Single, Py As Single)
   Dim iRow As Integer, iCol As Integer, bVisible As Boolean
   With grdDataList
   bVisible = .Visible
   .Visible = False
   For iRow = .TopRow To .Rows - 1
      For iCol = 2 To .Cols - 1
         .row = iRow
         .col = iCol
         If Px >= .CellLeft And Px <= .CellLeft + .CellWidth And Py >= .CellTop And Py <= .CellTop + .CellHeight Then
            GoTo flgDown
         End If
      Next
   Next
   .row = m_selRow
   .col = m_selCol
   
flgDown:
   .Visible = bVisible
   End With
End Sub

Private Sub grdDataList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim bolChange As Boolean, iLstRow As Integer, lLstColor As Long
   
On Error Resume Next

   If Button = 1 Then
      If x > grdDataList.ColWidth(0) + grdDataList.ColWidth(1) And y > grdDataList.RowHeight(0) And m_bolFree Then
         With grdDataList
         iLstRow = .row
         lLstColor = .CellBackColor
         If y < .CellTop Then
            If m_selRow - 1 > 1 Then
               m_selRow = m_selRow - 1
               bolChange = True
            End If
         End If
         If y > .CellTop + .CellHeight Then
            If m_selRow + 1 < .Rows Then
               m_selRow = m_selRow + 1
               bolChange = True
            End If
         End If
         
         If bolChange = True Then
            .row = m_selRow
            If lLstColor = m_selColor Then
               '移回頭要還原前格顏色
               If .CellBackColor = m_selColor Then
                  .row = iLstRow
                  .CellBackColor = GetDftColor
                  .CellForeColor = vbBlack
                  .row = m_selRow
               'Added by Morgan 2015/8/13
               ElseIf m_bolSalesAndCar Then
                  .CellBackColor = m_selColor
                  .CellForeColor = vbWhite
               'end 2015/8/13
               ElseIf .CellBackColor = GetDftColor Then
                  .CellBackColor = m_selColor
                  .CellForeColor = vbWhite
               End If
            End If
         End If
         End With
      End If
   End If
End Sub
'index:0=新增;1=修改;2=刪除;3=內容
Public Sub ShowNextForm(Index As Integer)
   Dim ii As Integer
   Dim stStart As String, stEnd As String
   Dim rsTmp As ADODB.Recordset
   Dim stStartDate As String, stEndDate As String
     
   With grdDataList
   '新增
   If Index = 0 Then
       m_SelDate = ""
      If m_selRow > 0 And m_selCol > 1 Then
         .row = m_selRow
         .col = m_selCol
         m_SelDate = Format((Val(.TextMatrix(0, 1)) + 1911) & "/" & Mid(.TextMatrix(0, .col), 1, Len(grdDataList.TextMatrix(0, .col)) - 3), "YYYYMMDD")
         'Added by Morgan 2012/11/20 跨年問題
         If Val(.TextMatrix(0, .col)) <> Val(.TextMatrix(0, 2)) And Val(.TextMatrix(0, 2)) = 12 Then
            m_SelDate = m_SelDate + 10000
         End If
         'end 2012/11/20

         .Visible = False
         For ii = 1 To .Rows - 1
            .row = ii
            .col = m_selCol
            If .CellBackColor = m_selColor Then
               If stStart = "" Then
                  stStart = .TextMatrix(.row, 0) & ":" & IIf(.TextMatrix(.row, 1) = "00", "00", "30")
               End If
               stEnd = IIf(.TextMatrix(.row, 1) = "00", .TextMatrix(.row, 0) & ":" & "30", Format(Val(.TextMatrix(.row, 0)) + 1, "00") & ":" & "00")
            End If
         Next
         .row = m_selRow
         .Visible = True
      End If
      
      'Added by Morgan 2015/8/20
      If m_bolSalesAndCar Then
         If FirstTimeCheck(m_SelDate, Replace(stEnd, ":", "")) = False Then
            'Modified by Morgan 2023/11/10
            'MsgBox "本時段之後已有第一次預約，無法新增！", vbExclamation
            MsgBox "當週第1次預約以前時段不可預約！", vbExclamation
            GoTo EXITSUB
         End If
      End If
      'end 2015/8/20
      
      'Added by Morgan 2017/12/18
      '5F或中型會議室新增預約時提醒視訊備註 --Robert
      If PUB_GetItemData(strCombo1ItemData, Combo1.ListIndex) = 1 Or PUB_GetItemData(strCombo1ItemData, Combo1.ListIndex) = 2 Then
         'Modified by Morgan 2020/11/11
         'MsgBox "如使用視訊，請加註「視訊」及「所別」！" & vbCrLf & "Ex:(視訊.全所),(視訊.北.中.南.高)..." & vbCrLf & vbCrLf & "(※:其他樓層，若同時使用視訊，請協調！)", vbExclamation, "使用視訊提醒"
         MsgBox "若為教育訓練，請逕至「教育訓練登入作業」登錄及預約！" & vbCrLf & vbCrLf & "如使用視訊，請加註「視訊」及「所別」！" & vbCrLf & "Ex:(視訊.全所),(視訊.北.中.南.高)...", vbExclamation, "使用提醒"
         'end 2020/11/11
      End If
      'end 2017/12/18
          
      With frm140112_1
      .m_State = "A" '新增
      .m_Users = stUsers 'Added by Morgan 2019/7/12
      .strRR20 = m_SN01 'Add by Amy 2019/11/12 教育訓練編號
      
      For ii = 0 To Combo1.ListCount - 1
         .cboRoom.AddItem Combo1.List(ii), ii
         'Modified by Morgan 2021/4/22
         '.cboRoom.ItemData(ii) = Combo1.ITEMDATA(ii)
         .cboRoomItemData = .cboRoomItemData & PUB_GetItemData(strCombo1ItemData, ii) & ","
         'end 2021/4/22
      Next
      .cboRoom.ListIndex = Combo1.ListIndex
      
      'Added by Morgan 2015/8/13
      If Pub_StrUserSt03 <> "M51" Then
         .Height = 4050
      End If
      'end 2015/8/13
      If m_SelDate = "" Then
         stStartDate = strSrvDate(2)
         If Format(Now, "N") < 30 Then
            .cboTime(0) = Format(Now, "HH:00")
            .cboTime(1) = Format(Now, "HH:30")
         Else
            .cboTime(0) = Format(Now, "HH:30")
            .cboTime(1) = Format(Val(Format(Now, "H")) + 1, "00") & ":00"
         End If
      Else
         If m_SelDate >= strSrvDate(1) Then
            stStartDate = TransDate(m_SelDate, 1)
         End If
         .cboTime(0) = stStart
         .cboTime(1) = stEnd
      End If
      .MaskEdBox1.Mask = ""
      .MaskEdBox1 = CFDate(stStartDate)
      .MaskEdBox1.Mask = DFormat
      .MaskEdBox2.Mask = DFormat
      .txtUser = strUserNum
      'Add by Amy 2019/11/12 預帶教育訓練標題
      If m_Title <> MsgText(601) Then
        If InStr(.txtContent, m_Title) = 0 Then
            .txtContent = m_Title & "-" & .txtContent  '教育訓練標題
        End If
      End If
      If m_SN01 <> MsgText(601) Then
        .cboRoom.Tag = .cboRoom.ListIndex
        .cboTime(0).Tag = .cboTime(0)
        .cboTime(1).Tag = .cboTime(1)
        .MaskEdBox1.Tag = .MaskEdBox1
        .MaskEdBox2.Tag = .MaskEdBox2
      End If
      'end 2019/11/12
      .SetEnable
      .Show vbModal
      End With
      'Add by Amy 2020/02/06 非畫面上日期需更新頁面
      If m_SN01 <> MsgText(601) Then RefreshGrid
   
   '不是新增
   Else
      strExc(1) = PUB_GetItemData(strCombo1ItemData, Combo1.ListIndex)
      strExc(2) = DBDATE(m_SelDate)
      strExc(3) = .TextMatrix(.row, 0) & IIf(.TextMatrix(.row, 1) = "00", "00", "30")
      '週期性
      'Modify by Amy 2019/01/24 +rr20
      If Left(.Text, 1) = "◎" Then
         strExc(0) = " select rr01,rr02,rr03,rr04,rr05,rr06,RR07,rr08,rr09,rr20" & _
            ",s1.st02 C1,sqldatet(rr11) C2,rr12 C3" & _
            ",s2.st02 C4,sqldatet(rr14) C5,rr15 C6,rr16" & _
            " from RoomResDetail,RoomReservation,staff s1,staff s2 where rd01=" & strExc(1) & " and rd04=" & strExc(2) & _
            " and rd05 is null and RR18=0 and rr01(+)=rd01 and rr02(+)=rd02 and rr03(+)=rd03" & _
            " and rr03<=" & strExc(3) & " and rr04>" & strExc(3) & _
            " and s1.st01(+)=rr10 and s2.st01(+)=rr13"
         
      Else
         strExc(0) = "select rr01,rr02,rr03,rr04,rr05,rr06,RR07,rr08,rr09,rr20" & _
            ",s1.st02 C1,sqldatet(rr11) C2,rr12 C3" & _
            ",s2.st02 C4,sqldatet(rr14) C5,rr15 C6,rr16" & _
            " from RoomReservation,staff s1,staff s2 where rr01=" & strExc(1) & " and rr02=" & strExc(2) & _
            " and rr05='N' and rr03<=" & strExc(3) & " and rr04>" & strExc(3) & " and RR18=0" & _
            " and s1.st01(+)=rr10 and s2.st01(+)=rr13"
      End If
      'end 2019/01/24
         
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With frm140112_1
            Select Case Index
            Case 1 '修改
               .m_State = "E"
            Case 2 '刪除
               .m_State = "D"
            Case 3 '內容
               .m_State = "S"
            End Select
            .m_Users = stUsers 'Added by Morgan 2019/7/12
            .strRR20 = m_SN01 'Add by Amy 2019/11/12 教育訓練編號
           
            For ii = 0 To Combo1.ListCount - 1
               .cboRoom.AddItem Combo1.List(ii), ii
               'Modified by Morgan 2021/4/22
               '.cboRoom.ItemData(ii) = Combo1.ITEMDATA(ii)
               .cboRoomItemData = .cboRoomItemData & PUB_GetItemData(strCombo1ItemData, ii) & ","
            Next
            .cboRoom.ListIndex = Combo1.ListIndex
            'Modified by Morgan 2021/4/22
            '.cboRoom.Tag = .cboRoom.ItemData(.cboRoom.ListIndex)
            .cboRoom.Tag = PUB_GetItemData(.cboRoomItemData, .cboRoom.ListIndex)

            'Added by Morgan 2015/8/13
            If Pub_StrUserSt03 <> "M51" Then
               .Height = 4050
            End If
            'end 2015/8/13
      
            stStartDate = TransDate(rsTmp("RR02"), 1)
            .MaskEdBox1.Mask = ""
            .MaskEdBox1 = CFDate(stStartDate)
            .MaskEdBox1.Mask = DFormat
            .MaskEdBox1.Tag = .MaskEdBox1

            .cboTime(0) = Format(rsTmp("RR03"), "00:00")
            .cboTime(0).Tag = .cboTime(0)
            
            .cboTime(1) = Format(rsTmp("RR04"), "00:00")
            .cboTime(1).Tag = .cboTime(1)
            
            .txtUser = "" & rsTmp("RR07")
            .txtUser.Tag = .txtUser
            .SetlblOldUser
            
            .txtContent = "" & rsTmp("RR08")
            'Add by Amy 2019/11/12 教育訓練有標題帶教育訓練標題
            If InStr("" & rsTmp("RR08"), m_Title) = 0 Then
                .txtContent = .txtContent & m_Title
            End If
            .txtContent.Tag = .txtContent
            .Text1 = "" & rsTmp("RR20")
            
            'Added by Morgan 2015/8/14
            If rsTmp("rr16") > 0 Then
               .lblTimes = "第" & rsTmp("rr16") & "次"
            End If
            'end 2015/8/14
            
            If rsTmp("RR09") = "N" Then
               .Check3.Value = 0
            Else
               .Check3.Value = 1
            End If
            .Check3.Tag = .Check3.Value
            
            .lblCreateData = "Create : " & rsTmp("C1") & " " & _
              " " & rsTmp("C2") & " " & _
              " " & Format(rsTmp("C3"), "00:00:00") & String(10, " ")
            If Not IsNull(rsTmp("C4")) Then
               .lblCreateData = .lblCreateData & _
                 "Update : " & rsTmp("C4") & " " & _
                 " " & rsTmp("C5") & " " & _
                 " " & Format("" & rsTmp("C6"), "00:00:00")
            End If
              
            'Modified by Morgan 2019/7/3
            'If rsTmp("RR05") = "1" Then
            '   .Check1.Value = 1
            If "" & rsTmp("RR05") <> "N" Then
               If rsTmp("RR05") = "1" Then
                  .Check1.Value = 1
               ElseIf rsTmp("RR05") = "2" Then
                  .Check2.Value = 1
               End If
            'end 2019/7/3
               stEndDate = TransDate(rsTmp("RR06"), 1)
               .MaskEdBox2.Mask = ""
               .MaskEdBox2 = CFDate(stEndDate)
               .MaskEdBox2.Mask = DFormat
               .ReadDetail m_SelDate
            End If
            
            .SetEnable
            .Show vbModal
         End With
         'Add by Amy 2020/02/06 非畫面上日期需更新頁面
         If m_SN01 <> MsgText(601) Then RefreshGrid
      Else
         MsgBox "無法讀取該筆資料！"
      End If
   End If
EXITSUB:
   End With
   Set rsTmp = Nothing
End Sub

'Modified by Lydia 2025/01/09 +傳入會議室pMeetRoom
Private Function CheckRight(Optional ByVal pMeetRoom As String) As Boolean
   'Add by Amy 2019/11/12
   Dim strQ As String
   Dim strDeptNo As String '前畫面部門
   
   If ChkStaffID(strUserNum, True) = True Then Exit Function
   If Pub_StrUserSt03 = "M51" Then CheckRight = True: Exit Function
   
   'Added by Lydia 2025/01/09 日期查詢
   If pMeetRoom <> "" Then
      strExc(1) = pMeetRoom
      strExc(2) = DBDATE(m_Sel2Date)
      strExc(3) = MGrid2.TextMatrix(MGrid2.row, 0) & IIf(MGrid2.TextMatrix(MGrid2.row, 1) = "00", "00", "30")
   Else
   'end 2025/01/09
      strExc(1) = PUB_GetItemData(strCombo1ItemData, Combo1.ListIndex)
      strExc(2) = DBDATE(m_SelDate)
      strExc(3) = grdDataList.TextMatrix(grdDataList.row, 0) & IIf(grdDataList.TextMatrix(grdDataList.row, 1) = "00", "00", "30")
   End If
   'Modify by Amy 2019/01/24 +rr20
   strQ = "select rr01,rr02,rr03,rr04,rr05,rr06,RR07,rr08,rr09,rr10,rr20 from RoomReservation" & _
      " where rr01=" & strExc(1) & " and rr02=" & strExc(2) & _
      " and rr05='N' and rr18=0 and (rr02>" & strSrvDate(1) & " or rr04>to_char(sysdate,'hh24mi')) and rr03<=" & strExc(3) & " and rr04>" & strExc(3) & _
      " union select rr01,rr02,rr03,rr04,rr05,rr06,RR07,rr08,rr09,rr10,rr20 from RoomResDetail,RoomReservation" & _
      " where rd01=" & strExc(1) & " and rd04=" & strExc(2) & _
      " and rd05 is null and rr18=0 and rr01(+)=rd01 and rr02(+)=rd02 and rr03(+)=rd03" & _
      " and rr03<=" & strExc(3) & " and rr04>" & strExc(3)
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strQ)
   If intI = 1 Then
       'Modify by Amy 2019/11/12 +教育訓練判斷
       '教育訓練進入,但由畫面自選時間
       If m_SN01 <> MsgText(601) Then
            strDeptNo = Me.Tag
            If m_SN01 = "" & RsTemp("RR20") Then
                '建立者
                If "" & RsTemp("RR10") = strUserNum Then
                    CheckRight = True
                'F21外專工程師同組,其他外專人員同部門
                ElseIf (strDeptNo = "F21" And PUB_GetStaffST16(strUserNum) = PUB_GetStaffST16(m_SN12)) _
                  Or (Left(strDeptNo, 2) = "F2" And strDeptNo <> "F21" And strDeptNo = GetST15(m_SN12)) Then
                    CheckRight = True
                ElseIf Left(strDeptNo, 2) <> "F2" And (Left(strDeptNo, 2) = Left(GetST15(m_SN12), 2) Or strDeptNo = "M51") Then
                    CheckRight = True
                End If
            End If
       '預約作業進入 , 建立人才可維護, 有教育訓練需於教育訓練需修改
       ElseIf RsTemp("RR10") = strUserNum And IsNull(RsTemp("RR20")) Then
            CheckRight = True
       End If
    
   End If
End Function
'Added by Morgan 2015/8/20
'當週第1次預約以前的時段不可再預約
Public Function FirstTimeCheck(pDate As String, pEndTime As String, Optional pUserNo As String) As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stEndDate As String
   
   If pUserNo = "" Then pUserNo = strUserNum
   
   intR = Weekday(ChangeWStringToWDateString(pDate))
   If intR = 7 Then
      stEndDate = pDate
   Else
      stEndDate = CompDate(2, 7 - intR, pDate)
   End If
   
   '當日取消也要算
   stSQL = "select 1 from RoomReservation" & _
      " where rr07='" & pUserNo & "' and rr01=9 and ((rr02>" & pDate & " and rr02<=" & stEndDate & ") or (rr02=" & pDate & " and rr03>=" & pEndTime & ")) and rr18=0 and rr16=1"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      FirstTimeCheck = False
   Else
      FirstTimeCheck = True
   End If
   Set rsQuery = Nothing
End Function

'Added by Morgan 2015/8/13
'本次借車次
Public Function GetTimes(pDate As String, Optional pUserNo As String) As Integer
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stStartDate As String, stEndDate As String
   
   If pUserNo = "" Then pUserNo = strUserNum
   
   intR = Weekday(ChangeWStringToWDateString(pDate))
   If intR = 1 Then
      stStartDate = pDate
   Else
      stStartDate = CompDate(2, -1 * (intR - 1), pDate)
   End If
   If intR = 7 Then
      stEndDate = pDate
   Else
      stEndDate = CompDate(2, 7 - intR, pDate)
   End If
   
   '當日取消也要算
   stSQL = "select nvl(count(*),0)+1 from RoomReservation" & _
      " where rr07='" & pUserNo & "' and rr01=9 and rr02>=" & stStartDate & " and rr02<=" & stEndDate & " and (rr18=0 or rr02=rr18)"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      GetTimes = rsQuery(0)
   End If
   Set rsQuery = Nothing
End Function

'檢查是否可新增借車預約
Public Function CheckAddable(pDate As String, pFromTime As String, pToTime As String, Optional pUserNo As String) As Boolean
   Dim stSQL As String, intR As Integer, stCon As String
   Dim rsQuery As ADODB.Recordset
   Dim stST15 As String
   
   If pUserNo = "" Then pUserNo = strUserNum
   
   '智權部同仁非當日借車時第1次借車優於非第1次
   If pDate > strSrvDate(1) Then
      stST15 = PUB_GetStaffST15(pUserNo, 1)
      'Modified by Morgan 2016/10/21 +總務(M11,M10)也比照智權人員規則
      'Modified by Morgan 2019/7/12 部門也抓設定(+客戶服務組 W10--文雄)
      'If Left(stST15, 1) = "S" Or stST15 = "M11" Or stST15 = "M10" Then
      If Left(stST15, 1) = "S" Or InStr(stUsers, stST15) > 0 Then
      'end 2019/7/12
         If GetTimes(pDate, pUserNo) = 1 Then
            stCon = " and (rr16 is null or rr16=1)"
         End If
      End If
   End If
   
   stSQL = "select * from RoomReservation" & _
      " where rr01=9 and rr02=" & pDate & " and rr03>=" & pFromTime & " and rr04<=" & pToTime & " and rr18=0" & stCon
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      CheckAddable = False
   Else
      CheckAddable = True
   End If
   Set rsQuery = Nothing
End Function

Private Sub grdDataList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   'Add by Amy 2019/12/09 避免timer觸發後無法連續選Grid
   If bolShowBlock = True Then
        Call GetFilkerRow
        Call CloseTimer
        bolShowBlock = False
   End If
   'end 2019/12/09
   SetCommand
End Sub

Private Function GetRR20Info(ByRef strR() As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim stQ As String, intQ As Integer
    
    GetRR20Info = False
    stQ = "Select * From RoomReservation Where RR20=" & Val(m_SN01)
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
       strR(1) = "" & RsQ.Fields("RR01") '會議室編號
       strR(2) = "" & RsQ.Fields("RR02") '日期
       strR(3) = "" & RsQ.Fields("RR03") '開始時間
       strR(4) = "" & RsQ.Fields("RR04") '結束時間
       GetRR20Info = True
    End If
    Set RsQ = Nothing
End Function


'Add by Amy 2019/12/09
Private Sub Timer1_Timer()
    Dim ii  As Integer
    
    If bolShowBlock = False Then Exit Sub
    
    grdDataList.row = intGrdR
    grdDataList.col = intGrdC
    If grdDataList.CellBackColor = m_ocpColor4 Then
        grdDataList.CellBackColor = RGB(&HFF, 0, 0)
    Else
        grdDataList.CellBackColor = m_ocpColor4
    End If
End Sub

Private Sub CloseTimer()
    Timer1.Enabled = False
    grdDataList.row = intGrdR
    grdDataList.col = intGrdC
    grdDataList.CellBackColor = m_ocpColor4
End Sub

Private Sub GetFilkerRow()
    Dim ii  As Integer
    
    For ii = 1 To grdDataList.row
        If Left(grdDataList.TextMatrix(ii, intGrdC), 1) = "★" Then
            intGrdR = ii
            Exit For
        End If
    Next ii
End Sub

'Add by Amy 2020/02/06 更新頁面
Private Sub RefreshGrid()
     Dim strRR(1 To 4) As String, strWeek As String
     
     If GetRR20Info(strRR()) = True Then
        m_StartDate = strRR(2)
        strWeek = GetWeekDay(ChangeWStringToWDateString(m_StartDate))
        If Right(strWeek, 1) = "六" Or Right(strWeek, 1) = "日" Then
            Check1.Value = 1
        Else
            Check1.Value = 0
        End If
        RefreshGridData
    End If
End Sub

Private Sub Timer2_Timer()
   Timer2.Enabled = False
   'Added by Lydia 2025/01/09
   If m_idX = 2 Then  '日期查詢
      ShowNextForm2 Val(Timer2.Tag)
   Else
   'end 2025/01/09
      ShowNextForm Val(Timer2.Tag)
   End If
End Sub

'Added by Lydia 2025/01/09
Private Sub Combo2_LostFocus()
Dim strTmp As String

   strTmp = GetCombo2Date("2", Combo2.Text, True)
   If strTmp = "" Then
      MsgBox "預設日期為今天！"
      Combo2.ListIndex = 0
   Else
      Combo2.Text = strTmp
   End If
   
   If Trim(Combo2.Text) <> Combo2.Tag Then
      If ChkDayCbo = False Then
         Call SetCombo2(Combo2.Text)
      Else
         RefreshGridData2
         Combo2.Tag = Trim(Combo2.Text)
      End If
   End If
End Sub

'Added by Lydia 2025/01/09
Private Function ChkDayCbo() As Boolean
Dim strTmp As String, bolFound As Boolean, intP As Integer
   
   bolFound = False
   strTmp = Trim(Combo2.Text)
   If Combo2.ListCount > 0 And strTmp <> "" Then
      For intP = 0 To Combo2.ListCount - 1
         If strTmp = Trim(Combo2.List(intP)) Then
            bolFound = True
            Exit For
         End If
      Next intP
   End If
   ChkDayCbo = bolFound
End Function

'Added by Lydia 2025/01/09
Private Sub Combo2_Click()
   If Trim(Combo2.Text) <> Combo2.Tag Then
      RefreshGridData2
      Combo2.Tag = Trim(Combo2.Text)
   End If
End Sub

'Added by Lydia 2025/01/09
Private Sub ResetGrid2(Optional bolColor As Boolean, Optional bolData As Boolean)
   Dim ii As Integer, jj As Integer, strDate As String, strLstText As String
   Dim bolVisivle As Boolean
  
   Timer3.Enabled = False
   With MGrid2
      bolVisivle = .Visible
      .Visible = False
      For jj = 2 To .Cols - 1
         .col = jj
         strLstText = ""
         For ii = 1 To .Rows - 1
            .row = ii
            If bolData Then
               .Text = ""
            End If
            .CellForeColor = vbBlack
            If bolColor Then
               If .Text <> "" Then
                  If strLstText <> .Text Then
                     strLstText = .Text
                  End If
                  '週期性
                  If Left(strLstText, 1) = "◎" Then
                     .CellBackColor = m_ocpColor3
                  Else
                     'Modify by Amy 2019/12/09 +教育訓練預約顯示※
                     If Left(strLstText, 1) = "※" Or Left(strLstText, 1) = "★" Then
                        .CellBackColor = m_ocpColor4
                        If strNowRR2 <> "" Then
                           'Modify by Amy 2020/01/14 + 判斷日期,避免按上下週也閃
                           If Left(strLstText, 1) = "★" And bolShowBlock = True _
                             And Mid(.TextMatrix(0, jj), 1, Val(InStr(.TextMatrix(0, jj), "(")) - 1) = Val(Mid(strNowRR2, 5, 2)) & "/" & Val(Mid(strNowRR2, 7, 2)) Then
                               Timer3.Enabled = True
                           End If
                        End If
                     ElseIf Left(strLstText, 1) = "◆" Then
                        .CellBackColor = m_ocpColor
                     Else
                        .CellBackColor = m_ocpColor2
                     End If
                  End If
               Else
                  .CellBackColor = GetDftColor2
               End If
            End If
         Next
      Next
      .Visible = bolVisivle
   End With
End Sub

'Added by Lydia 2025/01/09
'檢查是否為上班時間
Private Function GetDftColor2() As Long
   With MGrid2
   If bolWorkDay2 = False Or Val(.TextMatrix(.row, 0)) < 8 Or Val(.TextMatrix(.row, 0)) > 18 Then   '非工作時段
      GetDftColor2 = m_offColor
   ElseIf .TextMatrix(.row, 1) = "00" Then
      GetDftColor2 = m_dftColor
   Else
      GetDftColor2 = m_dftColor2
   End If
   End With
End Function

'Added by Lydia 2025/01/09
Public Sub RefreshGridData2(Optional ByVal pChoice As Integer, Optional ByVal pBolSet As Boolean)
   Dim intA As Integer
   Dim iRow As Integer, iCol As Integer
   Dim iPos1 As Integer, iPos2 As Integer, iWday As Integer
   Dim iHalfHours As Integer, iMaxBytes As Single, iBytes As Single, strText As String
   Dim iWeekDay As Integer, stSign As String
   
   strNowRR2 = ""

   bolWorkDay2 = False
   m_StdDate2 = GetCombo2Date("1", Combo2.Text, False)
   If m_StdDate2 = "" Then
      m_StdDate2 = strSrvDate(1)
   End If
   m_StdDate2 = DBDATE(m_StdDate2)
   
   Select Case pChoice
      Case 1 '上週
         m_StdDate2 = CompDate(2, -7, m_StdDate2)
      Case 2 '前一天
         m_StdDate2 = CompDate(2, -1, m_StdDate2)
      Case 3 '次日
         m_StdDate2 = CompDate(2, 1, m_StdDate2)
      Case 4 '下週
         m_StdDate2 = CompDate(2, 7, m_StdDate2)
   End Select
   If Check2.Value = 0 Then
      iWeekDay = Weekday(Format(m_StdDate2, "####/##/##"))
      If iWeekDay < 2 Or iWeekDay = 7 Then
         m_StdDate2 = CompWorkDay(1, m_StdDate2, IIf(pChoice <= 2, 1, 0))
         MsgBox "跳過星期六、日！", vbInformation
      End If
   End If
   bolWorkDay2 = ChkWorkDay(m_StdDate2)
   
   '清除資料
   ResetGrid2 , True
   
   If pBolSet = True Then
      strText = GetCombo2Date("2", TransDate(m_StdDate2, 1))
      Combo2.Tag = strText  '避免重覆觸發
      Combo2.Text = strText
      If ChkDayCbo = False Then
         Call SetCombo2(strText, IIf(pChoice > 0, True, False))
      End If
   End If
   
   With MGrid2
   .Visible = False
   .TextMatrix(0, 1) = ((m_StdDate2 \ 10000) - 1911) & "年" '年
   .WordWrap = True
   For intA = 1 To m_MeetCnt
      .col = intA + 1
      .MergeCol(.col) = False

      iCol = .col
      strExc(0) = " select rr02,rr03,rr04,rr05,rr07,rr08,rr16,rr20 from roomreservation where rr02=" & m_StdDate2 & " and rr05='N' and rr18=0 and rr01=" & intA & _
                  " Union select rd04,rr03,rr04,rr05,rr07,rr08,rr16,rr20 From roomresdetail, roomreservation Where RD04=" & m_StdDate2 & " And RD05 Is Null and rr18=0 and rr01(+)=rd01 and rr02(+)=rd02 and rr03(+)=rd03 and rd01=" & intA & _
                  " order by rr03"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Do While Not RsTemp.EOF
            strText = "" & RsTemp.Fields("rr08")
            'Moidfy by Amy 2019/12/09 +教育訓練預約顯示※
            If Not IsNull(RsTemp("rr20")) = True Then
               If m_SN01 <> MsgText(601) And m_SN01 = RsTemp("rr20") Then
                 '目前預約
                 'Add by Amy 2020/01/14
                 bolShowBlock = True
                 strNowRR2 = "" & RsTemp("rr02")
                 'end 2020/01/14
                 strText = "★" & strText
                 intGrdC = iCol
               Else
                 strText = "※" & strText
               End If
            ElseIf RsTemp("rr05") <> "N" Then
               strText = "◎" & strText
            Else
               If stSign = "◆" Then
                  stSign = "◇"
               Else
                  stSign = "◆"
               End If
               strText = stSign & strText
            End If
            iBytes = GetTextLength(strText)
            iHalfHours = (2 * (RsTemp.Fields("rr04") \ 100) + ((RsTemp.Fields("rr04") Mod 100) / 30)) - (2 * (RsTemp.Fields("rr03") \ 100) + ((RsTemp.Fields("rr03") Mod 100) / 30))
            '除非只有一格(半小時),否則最後一行固定帶預約人
            If iHalfHours > 1 Then
               If IsNull(RsTemp.Fields("rr07")) Then
                  iMaxBytes = 14 * iHalfHours
                  If iBytes > iMaxBytes Then
                     strText = PUB_StrToStr(strText, iMaxBytes - 3) & "..."
                  End If
               'Modified by Morgan 2015/8/13 借車加次數
               ElseIf RsTemp.Fields("rr16") > 0 Then
                  iMaxBytes = 14 * (iHalfHours - 1)
                  If iBytes > iMaxBytes Then
                     strText = PUB_StrToStr(strText, iMaxBytes - 3) & "..."
                  End If
                  strText = strText & vbCrLf & " .. " & GetStaffName(RsTemp.Fields("rr07"), True) & "(第" & RsTemp.Fields("rr16") & "次)"
               'end 2015/8/13
               Else
                  iMaxBytes = 14 * (iHalfHours - 1)
                  'Modified by Morgan 2017/12/21 配合顯示視訊訊息改以內容為主，長度超過則不顯示預約人
                  If iBytes > iMaxBytes Then
                     strText = PUB_StrToStr(strText, iMaxBytes - 3 + 14) & "..."
                  Else
                     strText = strText & vbCrLf & " ........... " & GetStaffName(RsTemp.Fields("rr07"), True)
                  End If
                  'end 2017/12/21
               End If
            ElseIf iBytes > 14 Then
               strText = PUB_StrToStr(strText, 14)
            End If
               
            iPos1 = 1 + 2 * RsTemp.Fields("rr03") \ 100 + IIf(RsTemp.Fields("rr03") Mod 100 >= 30, 1, 0)
            iPos2 = 1 + 2 * RsTemp.Fields("rr04") \ 100 + IIf(RsTemp.Fields("rr04") Mod 100 >= 30, 0, -1)
            For iRow = iPos1 To iPos2
               .row = iRow
               .col = iCol
               .CellFontSize = 9
               .Text = strText
               .CellAlignment = 1
            Next
            RsTemp.MoveNext
         Loop
      End If
      .MergeCol(.col) = True
   Next intA
   
   '設定顏色
   ResetGrid2 True
   m_sel2Row = 0
   m_sel2Col = 0
   m_Sel2Date = ""
   DisableFunc
   .Visible = True
   End With

   If PUB_GetItemData(strCombo1ItemData, Combo1.ListIndex) = 9 And (Left(Pub_StrUserSt15, 1) = "S" Or InStr(stUsers, Pub_StrUserSt15) > 0) Then
      m_bolSalesAndCar = True
   Else
      m_bolSalesAndCar = False
   End If
   
   'Add by Amy 2019/12/09 避免捲動 造成抓的Row是錯的重抓Row
   Call GetFilkerRow2
End Sub

'Added by Lydia 2025/01/09
Private Sub Check2_Click()

   Call SetCombo2(Combo2.Text)
   
End Sub

'Added by Lydia 2025/01/09
Private Sub GetFilkerRow2()
    Dim ii  As Integer
    
    For ii = 1 To MGrid2.row
        If Left(MGrid2.TextMatrix(ii, intGrdC), 1) = "★" Then
            intGrdR = ii
            Exit For
        End If
    Next ii
End Sub

'Added by Lydia 2025/01/09
Private Sub RefreshGrid2()
     Dim strRR(1 To 4) As String, strWeek As String
     
     If GetRR20Info(strRR()) = True Then
        m_StdDate2 = strRR(2)
        strWeek = GetWeekDay(ChangeWStringToWDateString(m_StdDate2))
        If Right(strWeek, 1) = "六" Or Right(strWeek, 1) = "日" Then
            Check2.Value = 1
        Else
            Check2.Value = 0
        End If
        RefreshGridData2
    End If
End Sub

'Added by Lydia 2025/01/09
'index:0=新增;1=修改;2=刪除;3=內容
Private Sub ShowNextForm2(Index As Integer)
   Dim ii As Integer
   Dim stStart As String, stEnd As String
   Dim rsTmp As ADODB.Recordset
   Dim stStartDate As String, stEndDate As String
     
   With MGrid2
   '新增
   If Index = 0 Then
       m_Sel2Date = ""
      If m_sel2Row > 0 And m_sel2Col > 1 Then
         .row = m_sel2Row
         .col = m_sel2Col
         m_Sel2Date = m_StdDate2
         '跨年問題
         If Val(.TextMatrix(0, .col)) <> Val(.TextMatrix(0, 2)) And Val(.TextMatrix(0, 2)) = 12 Then
            m_Sel2Date = m_Sel2Date + 10000
         End If
         
         .Visible = False
         For ii = 1 To .Rows - 1
            .row = ii
            .col = m_sel2Col
            If .CellBackColor = m_selColor Then
               If stStart = "" Then
                  stStart = .TextMatrix(.row, 0) & ":" & IIf(.TextMatrix(.row, 1) = "00", "00", "30")
               End If
               stEnd = IIf(.TextMatrix(.row, 1) = "00", .TextMatrix(.row, 0) & ":" & "30", Format(Val(.TextMatrix(.row, 0)) + 1, "00") & ":" & "00")
            End If
         Next
         .row = m_sel2Row
         .Visible = True
      End If
      
      'Added by Morgan 2015/8/20
      If m_bolSalesAndCar And m_MeetNo(m_sel2Col) = "9" Then
         If FirstTimeCheck(m_Sel2Date, Replace(stEnd, ":", "")) = False Then
            'Modified by Morgan 2023/11/10
            'MsgBox "本時段之後已有第一次預約，無法新增！", vbExclamation
            MsgBox "當週第1次預約以前時段不可預約！", vbExclamation
            GoTo EXITSUB
         End If
      End If
      'end 2015/8/20
      
      'Added by Morgan 2017/12/18
      '5F或中型會議室新增預約時提醒視訊備註 --Robert
      If m_MeetType(m_sel2Col) = "1" Then
         'Modified by Morgan 2020/11/11
         'MsgBox "如使用視訊，請加註「視訊」及「所別」！" & vbCrLf & "Ex:(視訊.全所),(視訊.北.中.南.高)..." & vbCrLf & vbCrLf & "(※:其他樓層，若同時使用視訊，請協調！)", vbExclamation, "使用視訊提醒"
         MsgBox "若為教育訓練，請逕至「教育訓練登入作業」登錄及預約！" & vbCrLf & vbCrLf & "如使用視訊，請加註「視訊」及「所別」！" & vbCrLf & "Ex:(視訊.全所),(視訊.北.中.南.高)...", vbExclamation, "使用提醒"
         'end 2020/11/11
      End If
      'end 2017/12/18
          
      With frm140112_1
      .m_State = "A" '新增
      .m_Users = stUsers 'Added by Morgan 2019/7/12
      .strRR20 = m_SN01 'Add by Amy 2019/11/12 教育訓練編號
      
      For ii = 2 To UBound(m_MeetNo)
         .cboRoom.AddItem Replace(MGrid2.TextMatrix(0, m_sel2Col), vbCrLf, " "), ii - 2
         .cboRoomItemData = .cboRoomItemData & m_MeetNo(ii) & ","
      Next
      .cboRoom.ListIndex = m_sel2Col - 2
      
      'Added by Morgan 2015/8/13
      If Pub_StrUserSt03 <> "M51" Then
         .Height = 4050
      End If
      'end 2015/8/13
      If m_Sel2Date = "" Then
         stStartDate = strSrvDate(2)
         If Format(Now, "N") < 30 Then
            .cboTime(0) = Format(Now, "HH:00")
            .cboTime(1) = Format(Now, "HH:30")
         Else
            .cboTime(0) = Format(Now, "HH:30")
            .cboTime(1) = Format(Val(Format(Now, "H")) + 1, "00") & ":00"
         End If
      Else
         If m_Sel2Date >= strSrvDate(1) Then
            stStartDate = TransDate(m_Sel2Date, 1)
         End If
         .cboTime(0) = stStart
         .cboTime(1) = stEnd
      End If
      .MaskEdBox1.Mask = ""
      .MaskEdBox1 = CFDate(stStartDate)
      .MaskEdBox1.Mask = DFormat
      .MaskEdBox2.Mask = DFormat
      .txtUser = strUserNum
      'Add by Amy 2019/11/12 預帶教育訓練標題
      If m_Title <> MsgText(601) Then
        If InStr(.txtContent, m_Title) = 0 Then
            .txtContent = m_Title & "-" & .txtContent  '教育訓練標題
        End If
      End If
      If m_SN01 <> MsgText(601) Then
        .cboRoom.Tag = .cboRoom.ListIndex
        .cboTime(0).Tag = .cboTime(0)
        .cboTime(1).Tag = .cboTime(1)
        .MaskEdBox1.Tag = .MaskEdBox1
        .MaskEdBox2.Tag = .MaskEdBox2
      End If
      'end 2019/11/12
      .SetEnable
      .Show vbModal
      End With
      'Add by Amy 2020/02/06 非畫面上日期需更新頁面
      If m_SN01 <> MsgText(601) Then RefreshGrid2
   
   '不是新增
   Else
      strExc(1) = m_MeetNo(m_sel2Col)
      strExc(2) = DBDATE(m_Sel2Date)
      strExc(3) = .TextMatrix(.row, 0) & IIf(.TextMatrix(.row, 1) = "00", "00", "30")
      '週期性
      'Modify by Amy 2019/01/24 +rr20
      If Left(.Text, 1) = "◎" Then
         strExc(0) = " select rr01,rr02,rr03,rr04,rr05,rr06,RR07,rr08,rr09,rr20" & _
            ",s1.st02 C1,sqldatet(rr11) C2,rr12 C3" & _
            ",s2.st02 C4,sqldatet(rr14) C5,rr15 C6,rr16" & _
            " from RoomResDetail,RoomReservation,staff s1,staff s2 where rd01=" & strExc(1) & " and rd04=" & strExc(2) & _
            " and rd05 is null and RR18=0 and rr01(+)=rd01 and rr02(+)=rd02 and rr03(+)=rd03" & _
            " and rr03<=" & strExc(3) & " and rr04>" & strExc(3) & _
            " and s1.st01(+)=rr10 and s2.st01(+)=rr13"
         
      Else
         strExc(0) = "select rr01,rr02,rr03,rr04,rr05,rr06,RR07,rr08,rr09,rr20" & _
            ",s1.st02 C1,sqldatet(rr11) C2,rr12 C3" & _
            ",s2.st02 C4,sqldatet(rr14) C5,rr15 C6,rr16" & _
            " from RoomReservation,staff s1,staff s2 where rr01=" & strExc(1) & " and rr02=" & strExc(2) & _
            " and rr05='N' and rr03<=" & strExc(3) & " and rr04>" & strExc(3) & " and RR18=0" & _
            " and s1.st01(+)=rr10 and s2.st01(+)=rr13"
      End If
      'end 2019/01/24
         
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         With frm140112_1
            Select Case Index
            Case 1 '修改
               .m_State = "E"
            Case 2 '刪除
               .m_State = "D"
            Case 3 '內容
               .m_State = "S"
            End Select
            .m_Users = stUsers
            .strRR20 = m_SN01 '教育訓練編號
           
            For ii = 2 To UBound(m_MeetNo)
               .cboRoom.AddItem Replace(MGrid2.TextMatrix(0, m_sel2Col), vbCrLf, " "), ii - 2
               .cboRoomItemData = .cboRoomItemData & m_MeetNo(ii) & ","
            Next
            .cboRoom.ListIndex = m_sel2Col - 2
            .cboRoom.Tag = m_MeetNo(m_sel2Col)
            
            If Pub_StrUserSt03 <> "M51" Then
               .Height = 4050
            End If

            stStartDate = TransDate(rsTmp("RR02"), 1)
            .MaskEdBox1.Mask = ""
            .MaskEdBox1 = CFDate(stStartDate)
            .MaskEdBox1.Mask = DFormat
            .MaskEdBox1.Tag = .MaskEdBox1

            .cboTime(0) = Format(rsTmp("RR03"), "00:00")
            .cboTime(0).Tag = .cboTime(0)
            
            .cboTime(1) = Format(rsTmp("RR04"), "00:00")
            .cboTime(1).Tag = .cboTime(1)
            
            .txtUser = "" & rsTmp("RR07")
            .txtUser.Tag = .txtUser
            .SetlblOldUser
            
            .txtContent = "" & rsTmp("RR08")
            '教育訓練有標題帶教育訓練標題
            If InStr("" & rsTmp("RR08"), m_Title) = 0 Then
                .txtContent = .txtContent & m_Title
            End If
            .txtContent.Tag = .txtContent
            .Text1 = "" & rsTmp("RR20")
            
            If rsTmp("rr16") > 0 Then
               .lblTimes = "第" & rsTmp("rr16") & "次"
            End If
            
            If rsTmp("RR09") = "N" Then
               .Check3.Value = 0
            Else
               .Check3.Value = 1
            End If
            .Check3.Tag = .Check3.Value
            
            .lblCreateData = "Create : " & rsTmp("C1") & " " & _
              " " & rsTmp("C2") & " " & _
              " " & Format(rsTmp("C3"), "00:00:00") & String(10, " ")
            If Not IsNull(rsTmp("C4")) Then
               .lblCreateData = .lblCreateData & _
                 "Update : " & rsTmp("C4") & " " & _
                 " " & rsTmp("C5") & " " & _
                 " " & Format("" & rsTmp("C6"), "00:00:00")
            End If
              
            If "" & rsTmp("RR05") <> "N" Then
               If rsTmp("RR05") = "1" Then
                  .Check1.Value = 1
               ElseIf rsTmp("RR05") = "2" Then
                  .Check2.Value = 1
               End If
               stEndDate = TransDate(rsTmp("RR06"), 1)
               .MaskEdBox2.Mask = ""
               .MaskEdBox2 = CFDate(stEndDate)
               .MaskEdBox2.Mask = DFormat
               .ReadDetail m_Sel2Date
            End If
            
            .SetEnable
            .Show vbModal
         End With
         '非畫面上日期需更新頁面
         If m_SN01 <> MsgText(601) Then RefreshGrid
      Else
         MsgBox "無法讀取該筆資料！"
      End If
   End If
EXITSUB:
   End With
   Set rsTmp = Nothing
End Sub

'Added by Lydia 2025/01/09
Private Sub MGrid2_DblClick()
   If m_sel2Row > 0 And m_sel2Col > 0 Then
      With MGrid2
      .row = m_sel2Row
      .col = m_sel2Col
      '點選日期
      m_Sel2Date = m_StdDate2
      If Val(.TextMatrix(0, .col)) <> Val(.TextMatrix(0, 2)) And Val(.TextMatrix(0, 2)) = 12 Then
         m_Sel2Date = m_Sel2Date + 10000
      End If
      
      If m_Sel2Date >= strSrvDate(1) Then
         If .Text = "" Then
            If cmdFunc(0).Enabled = True Then
               m_idX = 2
               cmdFunc(0).Value = True
            End If
         Else
            If CheckRight(m_MeetNo(m_sel2Col)) = True Then
               ShowNextForm2 1
            Else
               ShowNextForm2 3
            End If
         End If
      Else
         If .Text <> "" Then
            ShowNextForm2 3
         End If
      End If
      End With
   End If
End Sub

'Added by Lydia 2025/01/09
Private Sub MGrid2_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim iRow As Integer
   Dim bolChange As Boolean
   Dim strDate As String
   Dim strDownText As String
  
   If x > MGrid2.ColWidth(0) + MGrid2.ColWidth(1) And y > MGrid2.RowHeight(0) Then
      With MGrid2
      .Visible = False
      If Button = 1 Or Button = 2 Then m_idX = 2
      '右鍵
      If Button = 2 Then
         SetCellActive2 x, y
      End If
      
      If Button = 1 Or (Button = 2 And .CellBackColor <> m_selColor) Then
      
         m_sel2Row = .row
         m_sel2Col = .col
         strDownText = .Text
            
         ResetGrid2 True '所有顏色還原
         
         .row = m_sel2Row
         .col = m_sel2Col
         
         '已預約
         If .Text <> "" Then
            iRow = m_sel2Row
            Do While iRow > 1
               iRow = iRow - 1
               If .TextMatrix(iRow, m_sel2Col) <> strDownText Then
                  iRow = iRow + 1
                  Exit Do
               End If
            Loop
            .row = iRow '開始
            .CellBackColor = m_selColor
            .CellForeColor = vbWhite
            
            Do While iRow < .Rows - 1
               iRow = iRow + 1
               If .TextMatrix(iRow, m_sel2Col) <> strDownText Then
                  iRow = iRow - 1 '結束
                  Exit Do
               End If
               .row = iRow
               .CellBackColor = m_selColor
               .CellForeColor = vbWhite
            Loop

            m_sel2Row = .row
            '讓 MouseMove 不動作
            If m_bolSalesAndCar And m_MeetNo(m_sel2Col) = "9" Then
               m_bol2Free = True
            Else
               m_bol2Free = False
            End If
         '未預約
         Else
            .CellBackColor = m_selColor
            m_bol2Free = True
         End If
      End If
      .Visible = True
      
      '右鍵
      If Button = 2 Then
         SetCommand2 'Added by Morgan 2015/8/14
         If .CellBackColor <> GetDftColor2 Then
            mdiMain.mnuPopItem(0).Enabled = cmdFunc(0).Enabled
            mdiMain.mnuPopItem(1).Enabled = cmdFunc(1).Enabled
            mdiMain.mnuPopItem(2).Enabled = cmdFunc(2).Enabled
            mdiMain.mnuPopItem(3).Enabled = cmdFunc(3).Enabled
            PopupMenu mdiMain.mnuPop
         End If
      End If
      End With
   End If
End Sub

'Added by Lydia 2025/01/09
'將指定座標所在的儲存格設定為作用中
Private Sub SetCellActive2(Px As Single, Py As Single)
   Dim iRow As Integer, iCol As Integer, bVisible As Boolean
   With MGrid2
   bVisible = .Visible
   .Visible = False
   For iRow = .TopRow To .Rows - 1
      For iCol = 2 To .Cols - 1
         .row = iRow
         .col = iCol
         If Px >= .CellLeft And Px <= .CellLeft + .CellWidth And Py >= .CellTop And Py <= .CellTop + .CellHeight Then
            GoTo flgDown
         End If
      Next
   Next
   .row = m_sel2Row
   .col = m_sel2Col
   
flgDown:
   .Visible = bVisible
   End With
End Sub

'Added by Lydia 2025/01/09
Private Sub MGrid2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   Dim bolChange As Boolean, iLstRow As Integer, lLstColor As Long
   
On Error Resume Next

   If Button = 1 Then
      If x > MGrid2.ColWidth(0) + MGrid2.ColWidth(1) And y > MGrid2.RowHeight(0) And m_bol2Free Then
         With MGrid2
         iLstRow = .row
         lLstColor = .CellBackColor
         If y < .CellTop Then
            If m_sel2Row - 1 > 1 Then
               m_sel2Row = m_sel2Row - 1
               bolChange = True
            End If
         End If
         If y > .CellTop + .CellHeight Then
            If m_sel2Row + 1 < .Rows Then
               m_sel2Row = m_sel2Row + 1
               bolChange = True
            End If
         End If
         
         If bolChange = True Then
            .row = m_sel2Row
            If lLstColor = m_selColor Then
               '移回頭要還原前格顏色
               If .CellBackColor = m_selColor Then
                  .row = iLstRow
                  .CellBackColor = GetDftColor2
                  .CellForeColor = vbBlack
                  .row = m_sel2Row
               ElseIf .CellBackColor = GetDftColor2 Then
                  .CellBackColor = m_selColor
                  .CellForeColor = vbWhite
               End If
            End If
         End If
         End With
      End If
   End If
End Sub

'Added by Lydia 2025/01/09
Private Sub MGrid2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   'Add by Amy 2019/12/09 避免timer觸發後無法連續選Grid
   If bolShowBlock = True Then
        Call GetFilkerRow2
        Call CloseTimer3
        bolShowBlock = False
   End If
   'end 2019/12/09
   SetCommand2
End Sub

'Added by Lydia 2025/01/09
Private Sub Timer3_Timer()
    Dim ii  As Integer
    
    If bolShowBlock = False Then Exit Sub
    
    MGrid2.row = intGrdR
    MGrid2.col = intGrdC
    If MGrid2.CellBackColor = m_ocpColor4 Then
        MGrid2.CellBackColor = RGB(&HFF, 0, 0)
    Else
        MGrid2.CellBackColor = m_ocpColor4
    End If
End Sub

'Added by Lydia 2025/01/09
Private Sub CloseTimer3()
    Timer3.Enabled = False
    MGrid2.row = intGrdR
    MGrid2.col = intGrdC
    MGrid2.CellBackColor = m_ocpColor4
End Sub

'Added by Lydia 2025/01/09
Private Sub SetCommand2()
   Dim ii As Integer, stFromTime As String, stToTime As String, bolBooked As Boolean, bolAddable As Boolean
   
   With MGrid2
   DisableFunc

   'Modified by Morgan 2015/5/4 改所有人都能讀,有權限的才能維護
   If m_bolRetrievalSys = False And m_MeetNo(m_sel2Col) = "8" Then
      Exit Sub
   End If
   'end 2015/5/4
   
   If m_sel2Col > 1 And m_sel2Row > 0 Then
      .col = m_sel2Col
      .row = m_sel2Row
      m_Sel2Date = m_StdDate2
      '跨年問題
      If Val(.TextMatrix(0, .col)) <> Val(.TextMatrix(0, 2)) And Val(.TextMatrix(0, 2)) = 12 Then
         m_Sel2Date = m_Sel2Date + 10000
      End If
      
      '假日不開放借車,維持以紙本方式申請
      If m_MeetNo(m_sel2Col) = "9" Then
         If ChkWorkDay(m_Sel2Date) = False Then Exit Sub
      End If
      
      bolBooked = False
      bolAddable = False
      For ii = 1 To .Rows - 1
         .row = ii
         If .CellBackColor = m_selColor Then
            If .Text = "" Then
               bolAddable = True
            Else
               bolBooked = True
            End If
            If stFromTime = "" Then
               stFromTime = .TextMatrix(.row, 0) & IIf(.TextMatrix(.row, 1) = "00", "00", "30")
            End If
            stToTime = IIf(.TextMatrix(.row, 1) = "00", .TextMatrix(.row, 0) & "30", Format(Val(.TextMatrix(.row, 0)) + 1, "00") & "00")
         End If
      Next
      .row = m_sel2Row
         
      '智權人員借車
      If m_bolSalesAndCar And m_MeetNo(m_sel2Col) = "9" Then
         If m_Sel2Date > strSrvDate(1) Or (m_Sel2Date >= strSrvDate(1) And Val(stToTime & "00") > ServerTime) Then
            If bolAddable Then
               If bolBooked Then
                  If CheckAddable(m_Sel2Date, stFromTime, stToTime) = True Then
                     cmdFunc(0).Enabled = True
                  End If
               Else
                  cmdFunc(0).Enabled = True
               End If
               
            ElseIf bolBooked Then
            
               If CheckRight(m_MeetNo(m_sel2Col)) = True Then
                  cmdFunc(1).Enabled = True
                  cmdFunc(2).Enabled = True
                  
               '非當日第1次優先借車檢查
               ElseIf m_Sel2Date > strSrvDate(1) Then
                  If CheckAddable(m_Sel2Date, stFromTime, stToTime) = True Then
                     cmdFunc(0).Enabled = True
                  End If
               End If
               cmdFunc(3).Enabled = True
            End If
            
         ElseIf bolBooked Then
            cmdFunc(3).Enabled = True
            
         End If
         
      Else
      'end 2015/8/13
         If .Text = "" Then
            If m_Sel2Date >= strSrvDate(1) Then
               'Add by Amy 2020/02/06 過去時間不可新增(因可能為教育訓練的預約會議室修改 ex:新增教育訓練->預約好會議室->改今天已過去之時間為開始時間)
               If (m_Sel2Date = strSrvDate(1) And Format(stFromTime & "00", "000000") >= Format(ServerTime, "000000")) Or m_Sel2Date > strSrvDate(1) Then
                    cmdFunc(0).Enabled = True
                End If
               'end 2020/02/06
            End If
         Else
            If bolReadOnly = False Then 'Add by Amy 2020/01/14 +if 教育訓練查詢進入,只可看
                If Pub_StrUserSt03 = "M51" Or m_Sel2Date > strSrvDate(1) Or (m_Sel2Date >= strSrvDate(1) And Val(stToTime & "00") > ServerTime) Then
                   If CheckRight(m_MeetNo(m_sel2Col)) = True Then
                      cmdFunc(1).Enabled = True
                      cmdFunc(2).Enabled = True
                   End If
                End If
            End If
            cmdFunc(3).Enabled = True
         End If
      End If
   End If
   End With
End Sub

'Added by Lydia 2025/01/09
Private Function GetCombo2Date(ByVal pKind As String, ByVal pText As String, Optional ByVal bolChk As Boolean = False) As String
Dim strMid As String
   
   GetCombo2Date = ""
   If Trim(pText) <> "" Then
      strMid = pText
      If InStr(strMid, "(") > 0 Then
         strMid = Trim(Mid(strMid, 1, InStr(pText, "(") - 1))
      End If
      strMid = Replace(Trim(Left(strMid, 9)), "/", "")
      If bolChk = True Then
         If ChkDate(strMid) = False Then
            Exit Function
         End If
      End If
      GetCombo2Date = strMid '民國年月日
   End If
   If pKind = "2" And GetCombo2Date <> "" Then  '民國年/月/日(星期X)
      GetCombo2Date = ChangeTStringToTDateString(strMid) & " (" & Right(GetWeekDay(ChangeTStringToWDateString(strMid)), 1) & ")"
   End If
End Function

'Added by Lydia 2025/01/09
Private Sub cmdMoveD_Click(Index As Integer)
   RefreshGridData2 Index + 1, True
   
End Sub

'Added by Lydia 2025/01/09
Private Sub Command2_Click()
   Screen.MousePointer = vbHourglass
   Combo2.Text = ChangeTStringToTDateString(strSrvDate(2))
   RefreshGridData2 , True
   Screen.MousePointer = vbDefault
End Sub

'Added by Lydia 2025/01/09
Private Sub SetCombo2(Optional ByVal pDate1 As String, Optional ByVal pBolDay As Boolean)
Dim strTmp As String, pDate2 As String
   If pDate1 = "" Then
      pDate1 = strSrvDate(2)
   Else
      pDate1 = GetCombo2Date("1", pDate1)
   End If
   If pDate1 <> "" Then
      strTmp = Combo2.Text
      pDate1 = DBDATE(pDate1)
      Combo2.Clear
      For intI = 0 To IIf(Check2.Value = 1, 6, 4)
         If Check2.Value = vbChecked Or pBolDay = True Then
            If intI = 0 Then
               pDate2 = pDate1
            Else
               pDate2 = CompDate(2, intI, pDate1)
            End If
         Else
            pDate2 = CompWorkDay(intI + 1, pDate1)
         End If
         pDate2 = TransDate(pDate2, 1)
         Combo2.AddItem GetCombo2Date("2", pDate2, False)
      Next intI
      Combo2.ListIndex = 0
      If Combo2.Text <> strTmp Then
         Call Combo2_Click
      End If
   End If
End Sub
