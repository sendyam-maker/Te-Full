VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm170239 
   BorderStyle     =   1  '單線固定
   Caption         =   "年終獎金明細"
   ClientHeight    =   5660
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8930
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5660
   ScaleWidth      =   8930
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  '沒有框線
      ClipControls    =   0   'False
      Height          =   4215
      Left            =   120
      TabIndex        =   10
      Top             =   960
      Width           =   8535
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "工作天不滿一年者列印計算公式"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   26
         Left            =   240
         TabIndex        =   66
         Top             =   2700
         Visible         =   0   'False
         Width           =   2520
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "稅率%"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   18
         Left            =   5295
         TabIndex        =   65
         Top             =   3150
         Width           =   495
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   19
         Left            =   4680
         TabIndex        =   64
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "紅　　利："
         Height          =   180
         Index           =   24
         Left            =   3720
         TabIndex        =   63
         Top             =   1140
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "代扣補充保費："
         Height          =   180
         Index           =   17
         Left            =   6075
         TabIndex        =   62
         Top             =   3390
         Width           =   1260
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "代扣補充保費＝[(年終獎金＋紅利－缺勤扣款)－(４＊投保金額)]＊當年費率%"
         ForeColor       =   &H000000FF&
         Height          =   360
         Index           =   13
         Left            =   195
         TabIndex        =   61
         Top             =   3390
         Width           =   5895
      End
      Begin VB.Label lblDsp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "台一國際專利商標開發(股)公司"
         Height          =   180
         Index           =   0
         Left            =   915
         TabIndex        =   60
         Top             =   240
         Width           =   2460
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "公司別："
         Height          =   180
         Index           =   10
         Left            =   180
         TabIndex        =   59
         Top             =   240
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "PS : 代扣稅額＝ (年終獎金＋特殊功績獎金＋紅利－缺勤扣款)＊"
         ForeColor       =   &H000000FF&
         Height          =   180
         Index           =   7
         Left            =   240
         TabIndex        =   58
         Top             =   3150
         Width           =   4980
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "應發金額："
         Height          =   180
         Index           =   4
         Left            =   6435
         TabIndex        =   57
         Top             =   1140
         Width           =   900
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   17
         Left            =   7440
         TabIndex        =   56
         Top             =   3885
         Width           =   810
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   14
         Left            =   7440
         TabIndex        =   55
         Top             =   2700
         Width           =   810
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   10
         Left            =   7440
         TabIndex        =   54
         Top             =   1140
         Width           =   810
      End
      Begin VB.Line Line4 
         BorderWidth     =   2
         X1              =   120
         X2              =   8400
         Y1              =   780
         Y2              =   780
      End
      Begin VB.Line Line3 
         BorderWidth     =   2
         X1              =   120
         X2              =   8400
         Y1              =   1440
         Y2              =   1440
      End
      Begin VB.Line Line2 
         BorderWidth     =   2
         X1              =   120
         X2              =   8400
         Y1              =   2580
         Y2              =   2580
      End
      Begin VB.Line Line1 
         BorderWidth     =   2
         X1              =   120
         X2              =   8400
         Y1              =   2955
         Y2              =   2955
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "核發獎金基數："
         Height          =   180
         Index           =   32
         Left            =   3690
         TabIndex        =   53
         Top             =   480
         Width           =   1200
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "100 %"
         Height          =   180
         Index           =   5
         Left            =   5025
         TabIndex        =   52
         Top             =   480
         Width           =   450
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "考績："
         Height          =   180
         Index           =   31
         Left            =   1860
         TabIndex        =   51
         Top             =   480
         Width           =   540
      End
      Begin VB.Label lblDsp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "不得參加"
         Height          =   180
         Index           =   4
         Left            =   2445
         TabIndex        =   50
         Top             =   480
         Width           =   720
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "年度工作總天數："
         Height          =   180
         Index           =   30
         Left            =   5910
         TabIndex        =   49
         Top             =   480
         Width           =   1440
      End
      Begin VB.Label lblDsp 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "365"
         Height          =   180
         Index           =   6
         Left            =   7980
         TabIndex        =   48
         Top             =   480
         Width           =   270
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "基準月數："
         Height          =   180
         Index           =   12
         Left            =   180
         TabIndex        =   47
         Top             =   480
         Width           =   900
      End
      Begin VB.Label lblDsp 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "4.0"
         Height          =   180
         Index           =   3
         Left            =   1125
         TabIndex        =   46
         Top             =   480
         Width           =   225
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "平均基準月薪："
         Height          =   180
         Index           =   1
         Left            =   6090
         TabIndex        =   45
         Top             =   240
         Width           =   1200
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "公傷假時數："
         Height          =   180
         Index           =   15
         Left            =   3525
         TabIndex        =   44
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "實領金額："
         Height          =   180
         Index           =   11
         Left            =   6435
         TabIndex        =   43
         Top             =   3885
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "流產假時數："
         Height          =   180
         Index           =   9
         Left            =   3525
         TabIndex        =   42
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "代扣稅額："
         Height          =   180
         Index           =   8
         Left            =   6435
         TabIndex        =   41
         Top             =   3150
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "扣  除        病假時數："
         Height          =   180
         Index           =   6
         Left            =   180
         TabIndex        =   40
         Top             =   1560
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "借支金額："
         Height          =   180
         Index           =   5
         Left            =   6435
         TabIndex        =   39
         Top             =   3630
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "部門："
         Height          =   180
         Index           =   14
         Left            =   3690
         TabIndex        =   38
         Top             =   240
         Width           =   540
      End
      Begin VB.Label lblDsp 
         BackColor       =   &H00FFFFFF&
         Caption         =   "專利處英文顧問"
         Height          =   180
         Index           =   1
         Left            =   4245
         TabIndex        =   37
         Top             =   240
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "年     終       獎      金："
         Height          =   180
         Index           =   16
         Left            =   180
         TabIndex        =   36
         Top             =   900
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "事假時數："
         Height          =   180
         Index           =   18
         Left            =   975
         TabIndex        =   35
         Top             =   1800
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "特  殊  功  績  獎  金："
         Height          =   180
         Index           =   19
         Left            =   180
         TabIndex        =   34
         Top             =   1140
         Width           =   1710
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "缺勤扣款："
         Height          =   180
         Index           =   20
         Left            =   3705
         TabIndex        =   33
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "未休特別假時數："
         Height          =   180
         Index           =   21
         Left            =   5900
         TabIndex        =   32
         Top             =   900
         Width           =   1440
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "曠職時數："
         Height          =   180
         Index           =   22
         Left            =   975
         TabIndex        =   31
         Top             =   2040
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "應領金額："
         Height          =   180
         Index           =   23
         Left            =   6435
         TabIndex        =   30
         Top             =   2700
         Width           =   900
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "產假時數："
         Height          =   180
         Index           =   25
         Left            =   3705
         TabIndex        =   29
         Top             =   1560
         Width           =   900
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   2
         Left            =   7440
         TabIndex        =   28
         Top             =   240
         Width           =   810
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   7
         Left            =   1980
         TabIndex        =   27
         Top             =   900
         Width           =   810
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   8
         Left            =   1980
         TabIndex        =   26
         Top             =   1140
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "未休特別假代金："
         Height          =   180
         Index           =   2
         Left            =   3180
         TabIndex        =   25
         Top             =   900
         Width           =   1440
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   9
         Left            =   4680
         TabIndex        =   24
         Top             =   900
         Width           =   810
      End
      Begin VB.Label lbl 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "999  日  99  時"
         Height          =   180
         Index           =   0
         Left            =   7320
         TabIndex        =   23
         Top             =   900
         Width           =   1080
      End
      Begin VB.Label lbl 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "999  日  99  時"
         Height          =   180
         Index           =   1
         Left            =   1980
         TabIndex        =   22
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label lbl 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "999  日  99  時"
         Height          =   180
         Index           =   2
         Left            =   1980
         TabIndex        =   21
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label lbl 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "999  日  99  時"
         Height          =   180
         Index           =   3
         Left            =   1980
         TabIndex        =   20
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lbl 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "999  日  99  時"
         Height          =   180
         Index           =   4
         Left            =   4680
         TabIndex        =   19
         Top             =   1560
         Width           =   1080
      End
      Begin VB.Label lbl 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "999  日  99  時"
         Height          =   180
         Index           =   5
         Left            =   4680
         TabIndex        =   18
         Top             =   1800
         Width           =   1080
      End
      Begin VB.Label lbl 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "999  日  99  時"
         Height          =   180
         Index           =   6
         Left            =   4680
         TabIndex        =   17
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   11
         Left            =   4680
         TabIndex        =   16
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   12
         Left            =   7440
         TabIndex        =   15
         Top             =   3630
         Width           =   810
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   13
         Left            =   7440
         TabIndex        =   14
         Top             =   2280
         Width           =   810
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "扣除金額："
         Height          =   180
         Index           =   3
         Left            =   6435
         TabIndex        =   13
         Top             =   2280
         Width           =   900
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   15
         Left            =   7440
         TabIndex        =   12
         Top             =   3150
         Width           =   810
      End
      Begin VB.Label lblDsp 
         Alignment       =   1  '靠右對齊
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         Caption         =   "99,999,999"
         Height          =   180
         Index           =   16
         Left            =   7440
         TabIndex        =   11
         Top             =   3390
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "查詢(&O)"
      Default         =   -1  'True
      Height          =   400
      Left            =   6840
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   420
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   400
      Left            =   7740
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   420
      Width           =   800
   End
   Begin VB.TextBox txtYear 
      Height          =   285
      Left            =   840
      MaxLength       =   3
      TabIndex        =   2
      Top             =   510
      Width           =   600
   End
   Begin VB.HScrollBar HScroll1 
      Height          =   255
      Left            =   1785
      Max             =   200
      Min             =   150
      TabIndex        =   1
      Top             =   5250
      Value           =   200
      Width           =   4785
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   4455
      Top             =   0
   End
   Begin MSForms.ComboBox cboUser 
      Height          =   300
      Left            =   840
      TabIndex        =   0
      Top             =   120
      Width           =   2400
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "4233;529"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "年度："
      Height          =   180
      Index           =   0
      Left            =   300
      TabIndex        =   9
      Top             =   555
      Width           =   540
   End
   Begin VB.Label lblTest 
      Alignment       =   2  '置中對齊
      Appearance      =   0  '平面
      BackColor       =   &H80000005&
      BorderStyle     =   1  '單線固定
      Caption         =   "這是濃淡設定預覽"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   6735
      TabIndex        =   8
      Top             =   5280
      Width           =   1905
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "薪資資料濃淡設定"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   5280
      Width           =   1560
   End
   Begin VB.Label lblTimeOut 
      Appearance      =   0  '平面
      AutoSize        =   -1  'True
      BackColor       =   &H0000FFFF&
      Caption         =   "若您未繼續移動滑鼠,將會於 59 秒後登出"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   5040
      TabIndex        =   6
      Top             =   90
      Visible         =   0   'False
      Width           =   3435
   End
   Begin VB.Label lblStaffNo 
      AutoSize        =   -1  'True
      Caption         =   "員工："
      Height          =   180
      Left            =   300
      TabIndex        =   5
      Top             =   180
      Width           =   540
   End
End
Attribute VB_Name = "frm170239"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/7/14 Form2.0已修改
'create by sonia 2016/1/11
Option Explicit
Dim m_iCol As Integer, m_iRow As Integer
Dim m_StaffNoCon As String

Private Sub cboUser_KeyPress(KeyAscii As MSForms.ReturnInteger)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub cboUser_Validate(Cancel As Boolean)
   Dim ii As Integer
   For ii = 0 To cboUser.ListCount - 1
      If InStr(cboUser.List(ii), cboUser) > 0 Then
         cboUser.ListIndex = ii
      End If
   Next
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

Private Sub cmdSearch_Click()
      
   If txtYear = "" Then
      MsgBox "請輸入年度！", vbExclamation
      txtYear.SetFocus
      Exit Sub
   End If
   
   '尚未啟用
   If Val(Pub_MaxYBYear) = 0 Then
      MsgBox "尚無任何年終獎金資料可查詢！", vbExclamation
      txtYear.SetFocus
      Exit Sub
   End If
   
   '不可早於啟用年月
   If Val(txtYear) < Val(Left(Pub_StartYM, 4) - 1912) Then
      MsgBox "不可早於啟用年度 " & Val(Left(Pub_StartYM, 4) - 1912) & " 年！", vbExclamation
      txtYear.SetFocus
      Exit Sub
   End If
   
   '不可大於年終入帳最大年度
   If Val(txtYear) > Val(Pub_MaxYBYear) - 1912 Then
      MsgBox "不可大於年終獎金已入帳年度 " & Val((Pub_MaxYBYear) - 1912) & " 年！", vbExclamation
      txtYear.SetFocus
      Exit Sub
   End If
   
   cboUser_Validate False
   If cboUser.ListIndex < 0 Then
      MsgBox "請選擇員工！", vbExclamation
      cboUser.SetFocus
      Exit Sub
   End If
   
   QueryData

End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   'modify by sonia 2025/2/5 年終獎金資料應以"年終獎金作業可操作人員"來檢查權限
   'PUB_AddSalaryUser cboUser, False
   PUB_AddBonusUser cboUser, False
   
   If Val(Pub_MaxYBYear) <> 0 Then
      txtYear = Val(Pub_MaxYBYear) - 1912
   Else
      txtYear = Val(Left(Pub_StartYM, 4) - 1912)
   End If
   
   InitialField
   PUB_SetForeColorScroll HScroll1
   PUB_EnableSalaryTimer
   
End Sub

'add by 2025/2/5 可查年終獎金作業可操作人員名單加入下拉選單(參考PUB_AddSalaryUser)
Public Sub PUB_AddBonusUser(pCombo As Object, Optional bolFnoDisplay As Boolean = True)
Dim ii As Integer, arrTmp() As String
   
   arrTmp = Split(Pub_StaffBonusList, ";")
   pCombo.Clear
   For ii = LBound(arrTmp) To UBound(arrTmp)
      If arrTmp(ii) <> "" Then
         If bolFnoDisplay Or Left(Right(arrTmp(ii), 5), 1) <> "F" Then
            pCombo.AddItem arrTmp(ii)
         End If
      End If
   Next
   If pCombo.ListCount > 0 Then
      pCombo.ListIndex = 0
   End If
End Sub
'end 2025/2/5

' 初始化欄位陣列
Private Sub InitialField()
Dim i As Integer

      For i = 0 To 19
         lblDsp(i) = ""
         If i <= 6 Then
            Lbl(i) = ""
         End If
      Next i
      
      'add by sonia 2018/1/11 特殊功績獎金及紅利二欄有值才出現
      Label1(19).Visible = False
      Label1(24).Visible = False
      Label1(7) = "PS : 代扣稅額＝ (年終獎金－缺勤扣款)＊"
      lblDsp(18).Left = 3900
      lblDsp(18) = "稅率%"
      Label1(13) = "代扣補充保費＝[(年終獎金－缺勤扣款)－(４＊投保金額)]＊"
      Label1(26).Visible = False    '工作天不滿一年者才顯示計算公式
      'end 2018/1/11
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170239 = Nothing
End Sub

Private Sub HScroll1_Change()
   HScroll1_Scroll
End Sub

Private Sub HScroll1_Scroll()
   SetDataColor
   PUB_SaveForeColor HScroll1
End Sub

Private Sub SetDataColor()
Dim i As Integer
   
   lblTest.ForeColor = PUB_GetColor(HScroll1.Value)

   For i = 0 To 19
      lblDsp(i).ForeColor = lblTest.ForeColor
      If i <= 6 Then
         Lbl(i).ForeColor = lblTest.ForeColor
      End If
   Next i
    
End Sub

Private Sub Timer1_Timer()
   PUB_ShowSalaryCountDown lblTimeOut
End Sub

Private Sub txtYear_GotFocus()
   TextInverse txtYear
   CloseIme
End Sub

Private Sub txtYear_KeyPress(KeyAscii As Integer)
   If KeyAscii <> 8 And Not IsNumeric(Chr(KeyAscii)) Then
      KeyAscii = 0
      Beep
   End If
End Sub

Private Sub QueryData()
Dim strDay As String, strHour As String
Dim arrTmp() As String
Dim i As Integer
Dim strTemp(1 To 6) As String
Dim m_taxrate As String       '非固定之薪資所得扣繳稅率
Dim strNHIRate As Double      '補充保費費率  2016/2/24 add by sonia
Dim m_YearDay As Long         '年度總天數    2018/1/12 add by sonia

   InitialField
   arrTmp = Split(cboUser, " ")
   cboUser.Tag = arrTmp(1)
   
   'Modified by Morgan 2023/12/20 年終隔年才發，要加1年判斷是否帶新部門
   'strExc(0) = "SELECT YearBonus.* FROM YearBonus,acc080,acc090" & _
      " WHERE yb01 = '" & Val(txtYear + 1911) & "' and yb02= '" & cboUser.Tag & "'"
   strExc(0) = "SELECT a.*,a0802,decode(sign(yb01-" & Left(新部門啟用日, 4) & "+1),-1,a0902,a0922) DepName FROM YearBonus a,acc080,acc090,acc090new" & _
      " WHERE yb01 = '" & Val(txtYear + 1911) & "' and yb02= '" & cboUser.Tag & "'" & _
      " and a0801(+)=yb24 and a0901(+)=yb03 and a0921(+)=yb03"
   'end 2023/12/20
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      With RsTemp
      '公司名稱
      'Modified by Morgan 2023/12/20
      'lblDsp(0) = CompNameQuery(.Fields("yb24"))
      lblDsp(0) = .Fields("a0802")
      'end 2023/12/20
      '部門名稱
      'Modified by Morgan 2023/12/20
      'lblDsp(1) = GetDepartmentName(.Fields("yb03"))
      lblDsp(1) = "" & .Fields("DepName")
      'end 2023/12/20
      '年終獎金基準月數
      lblDsp(3) = GetYearBonusMonth(.Fields("yb01") - 1911, .Fields("yb02"))
      '年度工作總天數
      lblDsp(6) = GetYearWorkDay(.Fields("yb01") - 1911, .Fields("yb02"))
      '考績及核發獎金基數
      If GetYearMerit(.Fields("yb01") - 1911, .Fields("yb02"), strExc(1), strExc(2)) = True Then
         lblDsp(4) = strExc(1)
         lblDsp(5) = strExc(2)
      End If
      
      lblDsp(2) = Val("" & .Fields("yb04"))   '平均基準月薪
      lblDsp(7) = Val("" & .Fields("yb05"))   '年終獎金
      lblDsp(8) = Val("" & .Fields("yb06"))   '特殊功績獎金
      lblDsp(9) = Val("" & .Fields("yb08"))   '未休假代金
      strTemp(1) = Val("" & .Fields("yb09"))  '扣年終病假時數
      strTemp(2) = Val("" & .Fields("yb10"))  '扣年終事假時數
      strTemp(3) = Val("" & .Fields("yb11"))  '扣年終曠職時數
      strTemp(4) = Val("" & .Fields("yb12"))  '扣年終產假時數
      strTemp(5) = Val("" & .Fields("yb13"))  '扣年終流產假時數
      strTemp(6) = Val("" & .Fields("yb14"))  '扣年終公傷假時數
      lblDsp(11) = Val("" & .Fields("yb15"))  '缺勤扣款
      lblDsp(12) = Val("" & .Fields("yb16"))  '借支扣款
      lblDsp(15) = Val("" & .Fields("yb17"))  '代扣稅額
      lblDsp(16) = Val("" & .Fields("yb25"))  '補充保費
      lblDsp(19) = Val("" & .Fields("yb26"))  '紅利  add by sonia 2018/1/11
   
      Call Pub_GetSpecWorkHour(cboUser.Tag, Val(txtYear) + 19111231)   'add by sonia 2018/2/1
      
      '未休假時數
      If Val("" & .Fields("yb07")) = 0 Then
         Lbl(0) = "  0 日  0 時"
      Else
         'modify by sonia 2018/2/1 每日8小時改用上班特殊時數PUB_intWkHour
         'strDay = (.Fields("yb07") * 10) \ (8 * 10)
         'strHour = Round(.Fields("yb07") - (strDay * 8), 1)
         strDay = (.Fields("yb07") * 10) \ (PUB_intWkHour * 10)
         strHour = Round(.Fields("yb07") - (strDay * PUB_intWkHour), 1)
         'end 2018/2/1
         Lbl(0) = strDay + " 日 " + strHour + " 時"
      End If
      '6項缺勤扣款時數
      For i = 1 To 6
         If Val(strTemp(i)) = 0 Then
            Lbl(i) = "  0 日  0 時"
         Else
            'modify by sonia 2018/2/1 每日8小時改用上班特殊時數PUB_intWkHour
            'strDay = (strTemp(i) * 10) \ (8 * 10)
            'strHour = Round(strTemp(i) - (strDay * 8), 1)
            strDay = (strTemp(i) * 10) \ (PUB_intWkHour * 10)
            strHour = Round(strTemp(i) - (strDay * PUB_intWkHour), 1)
            'end 2018/2/1
            Lbl(i) = strDay + " 日 " + strHour + " 時"
         End If
      Next i
    
      '計算應發金額,應領金額,實領金額
      lblDsp(10) = "": lblDsp(13) = "": lblDsp(14) = "": lblDsp(17) = ""
      'modify by sonia 2018/1/11 +紅利yb26
      lblDsp(10) = Val(lblDsp(7)) + Val(lblDsp(8)) + Val(lblDsp(9)) + Val(lblDsp(19))    '應發金額
      'modify by sonia 2018/1/30 婧瑄說扣除金額不含借支,應領金額不扣,實領金額才扣
      'lblDsp(13) = Val(lblDsp(11)) + Val(lblDsp(12))                                    '扣除金額
      lblDsp(13) = Val(lblDsp(11))                                                       '扣除金額
      lblDsp(14) = Val(lblDsp(10)) - Val(lblDsp(13))                                     '應領金額
      'modify by sonia 2018/1/30 婧瑄說扣除金額不含借支,應領金額不扣,實領金額才扣
      'lblDsp(17) = Val(lblDsp(14)) - Val(lblDsp(15)) - Val(lblDsp(16))                  '實領金額
      lblDsp(17) = Val(lblDsp(14)) - Val(lblDsp(15)) - Val(lblDsp(16)) - Val(lblDsp(12)) '實領金額
      For i = 2 To 19
         Select Case i
            Case 2, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 19
               lblDsp(i) = Format(lblDsp(i), "#,###,##0")
            Case Else
         End Select
      Next i
   
      '代扣稅額稅率
      m_taxrate = 0
      strExc(0) = "select oc04 from OtherSalaryCode where oc01='01'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         m_taxrate = "" & RsTemp.Fields(0)
      End If
      lblDsp(18) = m_taxrate & "%"
   
      'modify by sonia 2018/1/11 特殊功績獎金及紅利二欄有值才出現
      'Label1(13) = "        代扣補充保費＝[(年終獎金＋特殊功績獎金－缺勤扣款)　　　　　　　　　－(４＊投保金額)]＊"
      Label1(19).Visible = False: lblDsp(8).Visible = False
      Label1(24).Visible = False: lblDsp(19).Visible = False
      If Val(lblDsp(8)) + Val(lblDsp(19)) = 0 Then
         Label1(7) = "PS : 代扣稅額＝ (年終獎金－缺勤扣款)＊"
         lblDsp(18).Left = 3565
         Label1(13) = "代扣補充保費＝[(年終獎金－缺勤扣款)－(４＊投保金額)]＊"
      ElseIf Val(lblDsp(8)) > 0 And Val(lblDsp(19)) = 0 Then
         Label1(19).Visible = True: lblDsp(8).Visible = True
         Label1(7) = "PS : 代扣稅額＝ (年終獎金＋特殊功績獎金－缺勤扣款)＊"
         lblDsp(18).Left = 4885
         Label1(13) = "代扣補充保費＝[(年終獎金＋特殊功績獎金－缺勤扣款)　　　　　　　　　　　　　　－(４＊投保金額)]＊"
      ElseIf Val(lblDsp(8)) = 0 And Val(lblDsp(19)) > 0 Then
         Label1(24).Visible = True: lblDsp(19).Visible = True
         Label1(7) = "PS : 代扣稅額＝ (年終獎金＋紅利－缺勤扣款)＊"
         lblDsp(18).Left = 4225
         Label1(13) = "代扣補充保費＝[(年終獎金＋紅利－缺勤扣款)－(４＊投保金額)]＊"
      Else
         Label1(19).Visible = True: lblDsp(8).Visible = True
         Label1(24).Visible = True: lblDsp(19).Visible = True
         Label1(7) = "PS : 代扣稅額＝ (年終獎金＋特殊功績獎金＋紅利－缺勤扣款)＊"
         lblDsp(18).Left = 5400
         Label1(13) = "代扣補充保費＝[(年終獎金＋特殊功績獎金＋紅利－缺勤扣款)　　　　　　　　　　　－(４＊投保金額)]＊"
      End If
      'end 2018/1/11
      
      'add by sonia 2016/2/24 抓補充保費費率
      strNHIRate = PUB_GetNhiRate(Val("" & .Fields("yb19")))
      Label1(13) = Label1(13) & strNHIRate & "%"
      'end 2016/2/24
      
      'add by sonia 2018/1/12 工作天不滿一年者列印計算公式
      Label1(26).Visible = False: Label1(24) = ""
      '取得計算年度之總天數
      If PUB_GetMonthDays((Val(txtYear) + 1911), 2) = 28 Then
         m_YearDay = 365
      Else
         m_YearDay = 366
      End If
      If lblDsp(6) <> m_YearDay Then
         Label1(26) = "年終獎金＝" & Format(lblDsp(2), "##,##0") & " * " & lblDsp(3) & " * " & lblDsp(5) & " * " & lblDsp(6) & " / " & m_YearDay
         Label1(26).Visible = True
      Else
         Label1(24) = "紅　　利："
      End If
      'end 2018/1/12
      
      End With
   Else
      MsgBox "查無符合條件資料！", vbExclamation
      txtYear.SetFocus
   End If
End Sub

' 取得年度工作總天數
Public Function GetYearWorkDay(ByVal strYear As String, ByVal StrStaff As String) As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   GetYearWorkDay = Empty
   strSql = "SELECT sum(sm27) FROM SalaryMonth WHERE sm01='" & StrStaff & "' and sm02>= '" & Val(strYear) + 1911 & "01' and sm02<='" & Val(strYear) + 1911 & "12' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields(0)) = False Then
         GetYearWorkDay = rsTmp.Fields(0)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   PUB_ResetSalaryTimer Me
End Sub


