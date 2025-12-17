VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180301_03 
   BorderStyle     =   1  '單線固定
   Caption         =   "出缺勤查詢－明細"
   ClientHeight    =   5730
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5730
   ScaleWidth      =   8950
   Tag             =   "加班資料"
   Begin VB.TextBox txtB1008_14 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      Height          =   200
      Left            =   4860
      Locked          =   -1  'True
      TabIndex        =   64
      TabStop         =   0   'False
      Text            =   "@可補休：剩餘 3.5 天"
      Top             =   930
      Width           =   3990
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   645
      Left            =   2820
      TabIndex        =   51
      Top             =   2070
      Visible         =   0   'False
      Width           =   2295
      Begin VB.ComboBox cboETime 
         Height          =   300
         ItemData        =   "frm180301_03.frx":0000
         Left            =   1290
         List            =   "frm180301_03.frx":0002
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   53
         Top             =   330
         Width           =   1005
      End
      Begin VB.ComboBox cboSTime 
         Height          =   300
         ItemData        =   "frm180301_03.frx":0004
         Left            =   1290
         List            =   "frm180301_03.frx":0006
         Locked          =   -1  'True
         Style           =   2  '單純下拉式
         TabIndex        =   52
         Top             =   0
         Width           =   1005
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "迄日下班時段："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   6
         Left            =   0
         TabIndex        =   55
         Top             =   360
         Width           =   1260
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "起日上班時段："
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   5
         Left            =   0
         TabIndex        =   54
         Top             =   30
         Width           =   1260
      End
   End
   Begin VB.TextBox txtB1018 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   4860
      Locked          =   -1  'True
      TabIndex        =   50
      Top             =   420
      Width           =   1845
   End
   Begin VB.TextBox txtB1008_2 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      Height          =   195
      Left            =   4860
      Locked          =   -1  'True
      TabIndex        =   49
      TabStop         =   0   'False
      Text            =   "@特別假：7天  已休3天"
      Top             =   690
      Width           =   3990
   End
   Begin VB.TextBox txtB1001 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   4
      Top             =   120
      Width           =   945
   End
   Begin VB.TextBox txtB1007_1 
      Height          =   300
      Left            =   3300
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   12
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1007_2 
      Height          =   300
      Left            =   4140
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   13
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1006 
      Height          =   300
      Left            =   3810
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   11
      Top             =   1290
      Width           =   945
   End
   Begin VB.Frame Frame03 
      BorderStyle     =   0  '沒有框線
      Height          =   885
      Left            =   5670
      TabIndex        =   26
      Top             =   1140
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox txtB1014 
         Height          =   315
         Left            =   540
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   16
         Top             =   120
         Width           =   225
      End
      Begin MSForms.TextBox txtB1015 
         Height          =   285
         Left            =   540
         TabIndex        =   59
         Top             =   480
         Width           =   2535
         VariousPropertyBits=   679495711
         ScrollBars      =   3
         Size            =   "4471;503"
         FontName        =   "新細明體-ExtB"
         FontHeight      =   180
         FontCharSet     =   136
         FontPitchAndFamily=   34
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "差程："
         Height          =   180
         Left            =   30
         TabIndex        =   29
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "(1:長程 2:短程 3:大陸 4:國外)"
         Height          =   180
         Left            =   780
         TabIndex        =   28
         Top             =   180
         Width           =   2235
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "地點："
         Height          =   180
         Left            =   30
         TabIndex        =   27
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.ComboBox CboB1002 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   300
      ItemData        =   "frm180301_03.frx":0008
      Left            =   960
      List            =   "frm180301_03.frx":000A
      Locked          =   -1  'True
      TabIndex        =   6
      Top             =   690
      Width           =   1695
   End
   Begin VB.TextBox txtB1003 
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   5
      Top             =   420
      Width           =   645
   End
   Begin VB.ComboBox CboB1008 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   300
      ItemData        =   "frm180301_03.frx":000C
      Left            =   3300
      List            =   "frm180301_03.frx":000E
      Locked          =   -1  'True
      TabIndex        =   7
      Top             =   690
      Width           =   1515
   End
   Begin VB.TextBox txtB1005_2 
      Height          =   300
      Left            =   1770
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1004 
      Height          =   300
      Left            =   1410
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   8
      Top             =   1290
      Width           =   945
   End
   Begin VB.TextBox txtB1005_1 
      Height          =   300
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   9
      Top             =   1620
      Width           =   585
   End
   Begin VB.CommandButton cmdQueryNext 
      Caption         =   "查詢下一筆(&N)"
      Height          =   360
      Left            =   6630
      TabIndex        =   0
      Top             =   30
      Width           =   1365
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8040
      TabIndex        =   1
      Top             =   30
      Width           =   800
   End
   Begin VB.Frame Frame01 
      BorderStyle     =   0  '沒有框線
      Height          =   495
      Left            =   960
      TabIndex        =   33
      Top             =   2040
      Width           =   1836
      Begin VB.TextBox txtB1010 
         Height          =   315
         Left            =   990
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   3
         Top             =   30
         Width           =   525
      End
      Begin VB.TextBox txtB1009 
         Height          =   315
         Left            =   240
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   2
         Top             =   30
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "共            日             時"
         Height          =   180
         Left            =   60
         TabIndex        =   34
         Top             =   90
         Width           =   1665
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm180301_03.frx":0010
      Height          =   2175
      Left            =   4710
      TabIndex        =   17
      Top             =   3540
      Width           =   4215
      _ExtentX        =   7426
      _ExtentY        =   3828
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
   Begin VB.TextBox txtB1020 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   6420
      Locked          =   -1  'True
      TabIndex        =   42
      TabStop         =   0   'False
      Tag             =   "v"
      Top             =   3420
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.TextBox txtB1019 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   1500
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   43
      Tag             =   "v"
      Top             =   3420
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.TextBox txtB1019_2 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   2280
      Locked          =   -1  'True
      TabIndex        =   44
      TabStop         =   0   'False
      Tag             =   "v"
      Top             =   3420
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtB1021 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   255
      Left            =   7500
      Locked          =   -1  'True
      TabIndex        =   47
      TabStop         =   0   'False
      Tag             =   "v"
      Top             =   3420
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.Frame Frame02 
      BorderStyle     =   0  '沒有框線
      Height          =   705
      Left            =   30
      TabIndex        =   30
      Top             =   2250
      Visible         =   0   'False
      Width           =   1875
      Begin VB.TextBox txtB101213 
         Height          =   315
         Left            =   750
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   15
         Top             =   360
         Width           =   705
      End
      Begin VB.TextBox txtB1030 
         Height          =   315
         Left            =   750
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   14
         Top             =   30
         Width           =   705
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "假日-共                     時"
         Height          =   180
         Left            =   60
         TabIndex        =   32
         Top             =   390
         Width           =   1725
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "實際時數"
         Height          =   180
         Left            =   60
         TabIndex        =   31
         Top             =   90
         Width           =   720
      End
   End
   Begin MSForms.Label Label27 
      Height          =   195
      Left            =   5520
      TabIndex        =   63
      Top             =   2460
      Width           =   3315
      VariousPropertyBits=   27
      Caption         =   "Update ID:           Date         Time             "
      Size            =   "5847;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label26 
      Height          =   195
      Left            =   5520
      TabIndex        =   62
      Top             =   2160
      Width           =   3315
      VariousPropertyBits=   27
      Caption         =   "Create ID:           Date         Time             "
      Size            =   "5847;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtB1011 
      Height          =   675
      Left            =   960
      TabIndex        =   61
      Top             =   2820
      Width           =   7965
      VariousPropertyBits=   -1466939361
      ScrollBars      =   3
      Size            =   "14049;1191"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtB1207 
      Height          =   2175
      Left            =   960
      TabIndex        =   60
      Top             =   3540
      Width           =   3705
      VariousPropertyBits=   -1466939361
      ScrollBars      =   3
      Size            =   "6535;3836"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtB1003_2 
      Height          =   285
      Left            =   1650
      TabIndex        =   58
      Top             =   420
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
   Begin VB.Label LblEndW 
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   4800
      TabIndex        =   57
      Top             =   1350
      Width           =   825
   End
   Begin VB.Label LblStarW 
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   2400
      TabIndex        =   56
      Top             =   1350
      Width           =   825
   End
   Begin VB.Label Label24 
      AutoSize        =   -1  'True
      Caption         =   "目前表單狀態："
      Height          =   180
      Left            =   3555
      TabIndex        =   48
      Top             =   420
      Width           =   1260
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "簽核意見："
      Height          =   255
      Left            =   30
      TabIndex        =   41
      Top             =   3540
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "表單編號："
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   1
      Left            =   30
      TabIndex        =   40
      Top             =   120
      Width           =   900
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   3150
      X2              =   4950
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "～"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2850
      TabIndex        =   39
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "時"
      Height          =   180
      Left            =   3930
      TabIndex        =   38
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      Height          =   180
      Index           =   2
      Left            =   3390
      TabIndex        =   37
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "            迄"
      Height          =   180
      Left            =   3570
      TabIndex        =   36
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "分"
      Height          =   180
      Left            =   4770
      TabIndex        =   35
      Top             =   1710
      Width           =   180
   End
   Begin VB.Line Line3 
      BorderWidth     =   3
      X1              =   930
      X2              =   4980
      Y1              =   1980
      Y2              =   1980
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "表單類別："
      Height          =   180
      Left            =   30
      TabIndex        =   25
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      ForeColor       =   &H00000000&
      Height          =   255
      Index           =   0
      Left            =   30
      TabIndex        =   24
      Top             =   420
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "假別："
      Height          =   180
      Left            =   2730
      TabIndex        =   23
      Top             =   750
      Width           =   540
   End
   Begin VB.Line Line1 
      BorderWidth     =   3
      X1              =   960
      X2              =   2760
      Y1              =   1230
      Y2              =   1230
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "分"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   2400
      TabIndex        =   22
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "時間：                 起                    "
      Height          =   180
      Left            =   390
      TabIndex        =   21
      Top             =   1020
      Width           =   2385
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   17
      Left            =   990
      TabIndex        =   20
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "時"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1560
      TabIndex        =   19
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "事由："
      Height          =   255
      Left            =   390
      TabIndex        =   18
      Top             =   2820
      Width           =   540
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "人事室簽收日期／時間："
      Height          =   255
      Index           =   4
      Left            =   4410
      TabIndex        =   46
      Tag             =   "v"
      Top             =   3420
      Visible         =   0   'False
      Width           =   1980
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "人事室簽收人員："
      Height          =   255
      Index           =   3
      Left            =   30
      TabIndex        =   45
      Tag             =   "v"
      Top             =   3420
      Visible         =   0   'False
      Width           =   1440
   End
End
Attribute VB_Name = "frm180301_03"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Sindy 2011/8/8
Option Explicit

Dim i As Integer, j As Integer
Public m_SA02 As Double, m_SA03 As Double
Dim strUpdDate As String, strUpdTime As String
Dim strContent As String, strSubject As String
Dim dblPrevRow As Double
Dim m_PrevForm As Form '前一畫面 'Add By Sindy 2013/6/21


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
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   
   '出缺勤電子簽核主檔
   'Modify By Sindy 2016/12/27 +,B1030
   strSql = "Select B1001,B1002,B1003,B1004,substr(ltrim(to_char('0000'||to_char(B1005),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),3,2) B1005,B1006,substr(ltrim(to_char('0000'||to_char(B1007),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),3,2) B1007,B1008||' '||AC03 B1008,B1009,B1010,B1011,B1012,B1013,B1014,B1015,B1016,B1017," & B1018CName & " B1018,B1019,B1020,B1021,B1022,B1023,B1024,B1025,B1026,B1027,B1028,B1029,B1030 " & _
            "From ABS010, allcode " & _
            "Where ac01(+)='04' and B1008=ac02(+) " & _
            "and B1001='" & Me.txtB1001 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("B1001")) Then txtB1001 = rsTmp.Fields("B1001")
      If Not IsNull(rsTmp.Fields("B1002")) Then CboB1002 = GetB1002Value(rsTmp.Fields("B1002"))
      If Not IsNull(rsTmp.Fields("B1003")) Then txtB1003 = rsTmp.Fields("B1003"): txtB1003_2 = GetPrjSalesNM(rsTmp.Fields("B1003"))
      If Not IsNull(rsTmp.Fields("B1004")) Then txtB1004 = ChangeWStringToTString(rsTmp.Fields("B1004"))
      If Not IsNull(rsTmp.Fields("B1005")) Then txtB1005_1 = Left(rsTmp.Fields("B1005"), 2): txtB1005_2 = Right(rsTmp.Fields("B1005"), 2)
      If Not IsNull(rsTmp.Fields("B1006")) Then txtB1006 = ChangeWStringToTString(rsTmp.Fields("B1006"))
      If Not IsNull(rsTmp.Fields("B1007")) Then txtB1007_1 = Left(rsTmp.Fields("B1007"), 2): txtB1007_2 = Right(rsTmp.Fields("B1007"), 2)
      If Not IsNull(rsTmp.Fields("B1008")) Then CboB1008 = rsTmp.Fields("B1008")
      If Not IsNull(rsTmp.Fields("B1009")) Then txtB1009 = rsTmp.Fields("B1009")
      If Not IsNull(rsTmp.Fields("B1010")) Then txtB1010 = rsTmp.Fields("B1010")
      If Not IsNull(rsTmp.Fields("B1011")) Then txtB1011 = rsTmp.Fields("B1011")
'      If Not IsNull(rsTmp.Fields("B1012")) Then txtB1012 = rsTmp.Fields("B1012")
'      If Not IsNull(rsTmp.Fields("B1013")) Then txtB1013 = rsTmp.Fields("B1013")

      'Add By Sindy 2021/8/11
      SetB102829Combo cboSTime, 1, txtB1004, txtB1003
      SetB102829Combo cboETime, 2, txtB1004, txtB1003
      '2021/8/11 END
      
      'Add By Sindy 2016/12/26
      If Not IsNull(rsTmp.Fields("B1012")) Then
         Label16.Caption = "平日-共                     時"
         txtB101213.Text = rsTmp.Fields("B1012")
      ElseIf Not IsNull(rsTmp.Fields("B1013")) Then
         Label16.Caption = "假日-共                     時"
         txtB101213.Text = rsTmp.Fields("B1013")
      End If
      If Not IsNull(rsTmp.Fields("B1030")) Then
         txtB1030 = rsTmp.Fields("B1030")
      Else
         txtB1030 = txtB101213
      End If
      '2016/12/26 END
      If Not IsNull(rsTmp.Fields("B1014")) Then txtB1014 = rsTmp.Fields("B1014")
      If Not IsNull(rsTmp.Fields("B1015")) Then txtB1015 = rsTmp.Fields("B1015")
      If Not IsNull(rsTmp.Fields("B1018")) Then txtB1018 = rsTmp.Fields("B1018")
      'Add By Sindy 2016/11/24
      If Not IsNull(rsTmp.Fields("B1028")) And Val("" & rsTmp.Fields("B1028")) > 0 Then
         For i = 0 To cboSTime.ListCount - 1
            If cboSTime.List(i) = Format(rsTmp.Fields("B1028"), "00:00") Then
               cboSTime.ListIndex = i
               Exit For
            End If
         Next i
         Frame1.Visible = True
      Else
         Frame1.Visible = False
      End If
      If Not IsNull(rsTmp.Fields("B1029")) And Val("" & rsTmp.Fields("B1029")) > 0 Then
         For i = 0 To cboETime.ListCount - 1
            If cboETime.List(i) = Format(rsTmp.Fields("B1029"), "00:00") Then
               cboETime.ListIndex = i
               Exit For
            End If
         Next i
         Frame1.Visible = True
      Else
         Frame1.Visible = False
      End If
      '2016/11/24 END
      
      If Not IsNull(rsTmp.Fields("B1019")) Then '已收單
         txtB1019 = rsTmp.Fields("B1019")
         txtB1019_2 = GetPrjSalesNM(rsTmp.Fields("B1019"))
         txtB1008_2.Visible = False
         txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
      Else '未收單
         txtB1008_2.Visible = True
         txtB1008_2 = GetCurrSpecRestDay(Trim(txtB1003), , Left(txtB1004, 3))
         'Add By Sindy 2024/12/10
         txtB1008_14.Visible = True
         txtB1008_14 = GetCurrFor14RestDay(Trim(txtB1003), , txtB1004)
         '2024/12/10 END
      End If
      If Not IsNull(rsTmp.Fields("B1020")) Then txtB1020 = ChangeWStringToTDateString(rsTmp.Fields("B1020"))
      If Not IsNull(rsTmp.Fields("B1021")) Then txtB1021 = Format(Right("0" & Trim(rsTmp.Fields("B1021")), 6), "##:##:##")
      
      Call CboB1002_Click
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Unload Me
      frm180301_01.Show
      Exit Sub
   End If
   
   Call QueryOther
   
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
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
   If Left(Trim(CboB1002), 2) <> "01" Then
      txtB1008_2.Visible = False
      txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub QueryOther()
   If Trim(txtB1001) <> "" Then
      '出缺勤電子簽核主檔
      strSql = "Select * From ABS010 Where B1001='" & Me.txtB1001 & "' "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            Label26.Visible = True
            Label27.Visible = True
            Call UpdateCUID(RsTemp)
         End If
      End If
      '表單流程備註檔
      SetABS012TextBox txtB1207, txtB1001
      '表單簽核檔
      strSql = "SELECT ST02||nvl(B1108,'') 簽核人員," & B1102CName & " 身份,sqldateT(B1105) 日期,sqltime6(B1106) 時間," & B1107CName & " 簽核結果,B1104 FROM ABS011,Staff WHERE B1101='" & txtB1001 & "' and B1104=ST01(+) order by B1102,B1103 asc "
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.RecordCount > 0 Then
            Set GRD1.Recordset = RsTemp
         End If
      End If
      
      'Add By Sindy 2020/8/14 顯示星期幾
      If Val(txtB1004) > 0 Then
         LblStarW.Caption = "(" & GetWeekDay(CDate(Format(DBDATE(txtB1004), "####/##/##"))) & ")"
      End If
      If Val(txtB1006) > 0 Then
         LblEndW.Caption = "(" & GetWeekDay(CDate(Format(DBDATE(txtB1006), "####/##/##"))) & ")"
      End If
      '2020/8/14 END
   End If
End Sub

Public Sub QueryData_2()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   
   '員工請假資料
   'Modify By Sindy 2024/10/18 +,SA19
   strSql = "Select SA01,SA02,substr(ltrim(to_char('0000'||to_char(SA03),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(SA03),'0000')),3,2) SA03,SA04,substr(ltrim(to_char('0000'||to_char(SA05),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(SA05),'0000')),3,2) SA05,SA06||' '||AC03 SA06,SA07,SA08,SA09,B1011," & B1018CName & " B1018,B1019,B1020,B1021,SA10,SA11,SA12,SA13,SA14,SA15,SA16,SA17,SA19 " & _
            "From Staff_Absence,allcode,ABS010 " & _
            "Where ac01(+)='04' and SA06=ac02(+) and SA09=B1001(+) " & _
            "and SA01='" & Me.txtB1003 & "' and SA02='" & m_SA02 & "' and SA03='" & m_SA03 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      CboB1002 = GetB1002Value("01")
      If Not IsNull(rsTmp.Fields("SA01")) Then txtB1003 = rsTmp.Fields("SA01"): txtB1003_2 = GetPrjSalesNM(rsTmp.Fields("SA01"))
      If Not IsNull(rsTmp.Fields("SA02")) Then txtB1004 = ChangeWStringToTString(rsTmp.Fields("SA02"))
      If Not IsNull(rsTmp.Fields("SA03")) Then txtB1005_1 = Left(rsTmp.Fields("SA03"), 2): txtB1005_2 = Right(rsTmp.Fields("SA03"), 2)
      If Not IsNull(rsTmp.Fields("SA04")) Then txtB1006 = ChangeWStringToTString(rsTmp.Fields("SA04"))
      If Not IsNull(rsTmp.Fields("SA05")) Then txtB1007_1 = Left(rsTmp.Fields("SA05"), 2): txtB1007_2 = Right(rsTmp.Fields("SA05"), 2)
      If Not IsNull(rsTmp.Fields("SA06")) Then CboB1008 = rsTmp.Fields("SA06")
      If Not IsNull(rsTmp.Fields("SA07")) Then txtB1009 = rsTmp.Fields("SA07")
      If Not IsNull(rsTmp.Fields("SA08")) Then txtB1010 = rsTmp.Fields("SA08")
      If Not IsNull(rsTmp.Fields("SA09")) Then txtB1001 = rsTmp.Fields("SA09")
      If Not IsNull(rsTmp.Fields("SA19")) Then txtB1011 = rsTmp.Fields("SA19") 'Modify By Sindy 2024/10/18 紙本請假事由
      
      'Add By Sindy 2021/8/11
      SetB102829Combo cboSTime, 1, txtB1004, txtB1003
      SetB102829Combo cboETime, 2, txtB1004, txtB1003
      '2021/8/11 END
      
      'Add By Sindy 2016/11/24
      If Not IsNull(rsTmp.Fields("SA16")) And Val("" & rsTmp.Fields("SA16")) > 0 Then
         For i = 0 To cboSTime.ListCount - 1
            If cboSTime.List(i) = Format(rsTmp.Fields("SA16"), "00:00") Then
               cboSTime.ListIndex = i
               Exit For
            End If
         Next i
         Frame1.Visible = True
      Else
         Frame1.Visible = False
      End If
      If Not IsNull(rsTmp.Fields("SA17")) And Val("" & rsTmp.Fields("SA17")) > 0 Then
         For i = 0 To cboETime.ListCount - 1
            If cboETime.List(i) = Format(rsTmp.Fields("SA17"), "00:00") Then
               cboETime.ListIndex = i
               Exit For
            End If
         Next i
         Frame1.Visible = True
      Else
         Frame1.Visible = False
      End If
      '2016/11/24 END
      
      If txtB1001 <> "" Then
         txtB1008_2.Visible = False
         txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
      End If
      '出缺勤簽核主檔資料
      If Not IsNull(rsTmp.Fields("B1011")) Then txtB1011 = rsTmp.Fields("B1011")
      If Not IsNull(rsTmp.Fields("B1018")) Then txtB1018 = rsTmp.Fields("B1018")
      If Not IsNull(rsTmp.Fields("B1019")) Then '已收單
         txtB1019 = rsTmp.Fields("B1019")
         txtB1019_2 = GetPrjSalesNM(rsTmp.Fields("B1019"))
         txtB1008_2.Visible = False
         txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
      Else '未收單
         txtB1008_2.Visible = True
         txtB1008_2 = GetCurrSpecRestDay(Trim(txtB1003), , Left(txtB1004, 3))
         'Add By Sindy 2024/12/10
         txtB1008_14.Visible = True
         txtB1008_14 = GetCurrFor14RestDay(Trim(txtB1003), , txtB1004)
         '2024/12/10 END
      End If
      If Not IsNull(rsTmp.Fields("B1020")) Then txtB1020 = ChangeWStringToTDateString(rsTmp.Fields("B1020"))
      If Not IsNull(rsTmp.Fields("B1021")) Then txtB1021 = Format(Right("0" & Trim(rsTmp.Fields("B1021")), 6), "##:##:##")
      
      Call CboB1002_Click
      
      If Trim(txtB1001) = "" Then
         Label26.Visible = True
         Label27.Visible = True
         Call UpdateCUID2(rsTmp)
      End If
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Unload Me
      frm180301_01.Show
      Exit Sub
   End If
   rsTmp.Close
      
   Call QueryOther
   
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
   
   Screen.MousePointer = vbDefault
   
   If Left(Trim(CboB1002), 2) <> "01" Then
      txtB1008_2.Visible = False
      txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Public Sub QueryData_3()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   
   '員工加班資料
   'Modify By Sindy 2012/6/18 +and So03=" & m_SA03 & "
   'Modify By Sindy 2016/12/27 +,So15
   strSql = "Select So01,So02,substr(ltrim(to_char('0000'||to_char(So03),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(So03),'0000')),3,2) So03,substr(ltrim(to_char('0000'||to_char(So04),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(So04),'0000')),3,2) So04,So05,So06,So13,B1011," & B1018CName & " B1018,B1019,B1020,B1021,so07,so08,so09,so10,so11,so12,So15 " & _
            "From Staff_Overtime,ABS010 " & _
            "Where So01='" & Me.txtB1003 & "' and So02='" & m_SA02 & "' and So03=" & m_SA03 & " and So13=B1001(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      CboB1002 = GetB1002Value("02")
      If Not IsNull(rsTmp.Fields("So01")) Then txtB1003 = rsTmp.Fields("So01"): txtB1003_2 = GetPrjSalesNM(rsTmp.Fields("So01"))
      If Not IsNull(rsTmp.Fields("So02")) Then txtB1004 = ChangeWStringToTString(rsTmp.Fields("So02"))
      If Not IsNull(rsTmp.Fields("So03")) Then txtB1005_1 = Left(rsTmp.Fields("So03"), 2): txtB1005_2 = Right(rsTmp.Fields("So03"), 2)
      If Not IsNull(rsTmp.Fields("So04")) Then txtB1007_1 = Left(rsTmp.Fields("So04"), 2): txtB1007_2 = Right(rsTmp.Fields("So04"), 2)
'      If Not IsNull(rsTmp.Fields("So05")) Then txtB1012 = rsTmp.Fields("So05")
'      If Not IsNull(rsTmp.Fields("So06")) Then txtB1013 = rsTmp.Fields("So06")
      'Add By Sindy 2016/12/26
      If Not IsNull(rsTmp.Fields("So05")) Then
         Label16.Caption = "平日-共                     時"
         txtB101213.Text = rsTmp.Fields("So05")
      ElseIf Not IsNull(rsTmp.Fields("So06")) Then
         Label16.Caption = "假日-共                     時"
         txtB101213.Text = rsTmp.Fields("So06")
      End If
      If Not IsNull(rsTmp.Fields("So15")) Then
         txtB1030 = rsTmp.Fields("So15")
      Else
         txtB1030 = txtB101213
      End If
      '2016/12/26 END
      If Not IsNull(rsTmp.Fields("So13")) Then txtB1001 = rsTmp.Fields("So13")
      
      If txtB1001 <> "" Then
         txtB1008_2.Visible = False
         txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
      End If
      '出缺勤簽核主檔資料
      If Not IsNull(rsTmp.Fields("B1011")) Then txtB1011 = rsTmp.Fields("B1011")
      If Not IsNull(rsTmp.Fields("B1018")) Then txtB1018 = rsTmp.Fields("B1018")
      If Not IsNull(rsTmp.Fields("B1019")) Then '已收單
         txtB1019 = rsTmp.Fields("B1019")
         txtB1019_2 = GetPrjSalesNM(rsTmp.Fields("B1019"))
         txtB1008_2.Visible = False
         txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
      Else '未收單
         txtB1008_2.Visible = True
         txtB1008_2 = GetCurrSpecRestDay(Trim(txtB1003), , Left(txtB1004, 3))
         'Add By Sindy 2024/12/10
         txtB1008_14.Visible = True
         txtB1008_14 = GetCurrFor14RestDay(Trim(txtB1003), , txtB1004)
         '2024/12/10 END
      End If
      If Not IsNull(rsTmp.Fields("B1020")) Then txtB1020 = ChangeWStringToTDateString(rsTmp.Fields("B1020"))
      If Not IsNull(rsTmp.Fields("B1021")) Then txtB1021 = Format(Right("0" & Trim(rsTmp.Fields("B1021")), 6), "##:##:##")
      
      Call CboB1002_Click
      
      If Trim(txtB1001) = "" Then
         Label26.Visible = True
         Label27.Visible = True
         Call UpdateCUID3(rsTmp)
      End If
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Unload Me
      frm180301_01.Show
      Exit Sub
   End If
   rsTmp.Close
   
   Call QueryOther
   
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
   
   Screen.MousePointer = vbDefault
   
   If Left(Trim(CboB1002), 2) <> "01" Then
      txtB1008_2.Visible = False
      txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Public Sub QueryData_4()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   Screen.MousePointer = vbHourglass
   
   '員工出差資料
   strSql = "Select SB01,SB02,substr(ltrim(to_char('0000'||to_char(SB03),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(SB03),'0000')),3,2) SB03,SB04,substr(ltrim(to_char('0000'||to_char(SB05),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(SB05),'0000')),3,2) SB05,SB06,SB07,SB08,SB09,SB10,B1011," & B1018CName & " B1018,B1019,B1020,B1021,SB11,SB12,SB13,SB14,SB15,SB16,SB17,SB18 " & _
            "From Staff_Busi_Trip,ABS010 " & _
            "Where SB01='" & Me.txtB1003 & "' and SB02='" & m_SA02 & "' and SB03='" & m_SA03 & "' and SB10=B1001(+) "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      CboB1002 = GetB1002Value("03")
      If Not IsNull(rsTmp.Fields("SB01")) Then txtB1003 = rsTmp.Fields("SB01"): txtB1003_2 = GetPrjSalesNM(rsTmp.Fields("SB01"))
      If Not IsNull(rsTmp.Fields("SB02")) Then txtB1004 = ChangeWStringToTString(rsTmp.Fields("SB02"))
      If Not IsNull(rsTmp.Fields("SB03")) Then txtB1005_1 = Left(rsTmp.Fields("SB03"), 2): txtB1005_2 = Right(rsTmp.Fields("SB03"), 2)
      If Not IsNull(rsTmp.Fields("SB04")) Then txtB1006 = ChangeWStringToTString(rsTmp.Fields("SB04"))
      If Not IsNull(rsTmp.Fields("SB05")) Then txtB1007_1 = Left(rsTmp.Fields("SB05"), 2): txtB1007_2 = Right(rsTmp.Fields("SB05"), 2)
      If Not IsNull(rsTmp.Fields("SB06")) Then txtB1009 = rsTmp.Fields("SB06")
      If Not IsNull(rsTmp.Fields("SB07")) Then txtB1010 = rsTmp.Fields("SB07")
      If Not IsNull(rsTmp.Fields("SB08")) Then txtB1014 = rsTmp.Fields("SB08")
      If Not IsNull(rsTmp.Fields("SB09")) Then txtB1015 = rsTmp.Fields("SB09")
      If Not IsNull(rsTmp.Fields("SB10")) Then txtB1001 = rsTmp.Fields("SB10")
      
      'Add By Sindy 2021/8/11
      SetB102829Combo cboSTime, 1, txtB1004, txtB1003
      SetB102829Combo cboETime, 2, txtB1004, txtB1003
      '2021/8/11 END
      
      'Add By Sindy 2016/11/24
      If Not IsNull(rsTmp.Fields("SB17")) And Val("" & rsTmp.Fields("SB17")) > 0 Then
         For i = 0 To cboSTime.ListCount - 1
            If cboSTime.List(i) = Format(rsTmp.Fields("SB17"), "00:00") Then
               cboSTime.ListIndex = i
               Exit For
            End If
         Next i
         Frame1.Visible = True
      Else
         Frame1.Visible = False
      End If
      If Not IsNull(rsTmp.Fields("SB18")) And Val("" & rsTmp.Fields("SB18")) > 0 Then
         For i = 0 To cboETime.ListCount - 1
            If cboETime.List(i) = Format(rsTmp.Fields("SB18"), "00:00") Then
               cboETime.ListIndex = i
               Exit For
            End If
         Next i
         Frame1.Visible = True
      Else
         Frame1.Visible = False
      End If
      '2016/11/24 END
      
      If txtB1001 <> "" Then
         txtB1008_2.Visible = False
         txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
      End If
      '出缺勤簽核主檔資料
      If Not IsNull(rsTmp.Fields("B1011")) Then txtB1011 = rsTmp.Fields("B1011")
      If Not IsNull(rsTmp.Fields("B1018")) Then txtB1018 = rsTmp.Fields("B1018")
      If Not IsNull(rsTmp.Fields("B1019")) Then '已收單
         txtB1019 = rsTmp.Fields("B1019")
         txtB1019_2 = GetPrjSalesNM(rsTmp.Fields("B1019"))
         txtB1008_2.Visible = False
         txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
      Else '未收單
         txtB1008_2.Visible = True
         txtB1008_2 = GetCurrSpecRestDay(Trim(txtB1003), , Left(txtB1004, 3))
         'Add By Sindy 2024/12/10
         txtB1008_14.Visible = True
         txtB1008_14 = GetCurrFor14RestDay(Trim(txtB1003), , txtB1004)
         '2024/12/10 END
      End If
      If Not IsNull(rsTmp.Fields("B1020")) Then txtB1020 = ChangeWStringToTDateString(rsTmp.Fields("B1020"))
      If Not IsNull(rsTmp.Fields("B1021")) Then txtB1021 = Format(Right("0" & Trim(rsTmp.Fields("B1021")), 6), "##:##:##")
      
      Call CboB1002_Click
      
      If Trim(txtB1001) = "" Then
         Label26.Visible = True
         Label27.Visible = True
         Call UpdateCUID4(rsTmp)
      End If
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Unload Me
      frm180301_01.Show
      Exit Sub
   End If
   rsTmp.Close
   
   Call QueryOther
   
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
   
   Screen.MousePointer = vbDefault
   
   If Left(Trim(CboB1002), 2) <> "01" Then
      txtB1008_2.Visible = False
      txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

Private Sub ClearField()
   txtB1001 = Empty
   CboB1002 = Empty
   txtB1003 = Empty
   txtB1003_2 = Empty
   txtB1003 = strUserNum
   txtB1003_2 = strUserName
   txtB1004 = Empty
   txtB1005_1 = Empty
   txtB1005_2 = Empty
   txtB1006 = Empty
   txtB1007_1 = Empty
   txtB1007_2 = Empty
   CboB1008 = Empty
   txtB1008_2 = Empty
   txtB1008_14 = Empty 'Add By Sindy 2024/12/10
   txtB1009 = Empty
   txtB1010 = Empty
   txtB1011 = Empty
   txtB1030 = Empty
   txtB101213 = Empty
   txtB1014 = Empty
   txtB1015 = Empty
   txtB1207 = Empty
   txtB1018 = Empty
   txtB1019 = Empty
   txtB1019_2 = Empty
   txtB1020 = Empty
   txtB1021 = Empty
   GRD1.Clear
   SetGrd
   Label26.Visible = False
   Label27.Visible = False
   LblStarW.Caption = "" 'Add By Sindy 2020/9/7
   LblEndW.Caption = "" 'Add By Sindy 2020/9/7
End Sub

Private Sub Form_Load()
'   If frm180301.cmdok(1).Tag = "" Then
      MoveFormToCenter Me
'   End If
   
   txtB1001.BackColor = &H8000000F
   txtB1003.BackColor = &H8000000F
   txtB1003_2.BackColor = &H8000000F
   txtB1008_2.BackColor = &H8000000F
   txtB1008_14.BackColor = &H8000000F 'Add By Sindy 2024/12/10
   txtB1030.BackColor = &H8000000F
   txtB101213.BackColor = &H8000000F
   
   txtB1004.BackColor = &H8000000F
   txtB1005_1.BackColor = &H8000000F
   txtB1005_2.BackColor = &H8000000F
   txtB1006.BackColor = &H8000000F
   txtB1007_1.BackColor = &H8000000F
   txtB1007_2.BackColor = &H8000000F
   txtB1009.BackColor = &H8000000F
   txtB1010.BackColor = &H8000000F
   txtB1014.BackColor = &H8000000F
   txtB1015.BackColor = &H8000000F
   txtB1011.BackColor = &H8000000F
   cboSTime.BackColor = &H8000000F
   cboETime.BackColor = &H8000000F
   
   '清空欄位值
   ClearField
   
   'Add By Sindy 2013/7/3
   If UCase(m_PrevForm.Name) = UCase("frm180301_01") Then
      cmdExit.Visible = True
   Else
      cmdExit.Visible = False
   End If
   '2013/7/3 END
End Sub

Private Sub Form_Unload(Cancel As Integer)
   'Set m_PrevForm = Nothing 'Add By Sindy 2013/6/21
   Set frm180301_03 = Nothing
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

Private Sub cmdExit_Click()
   Unload Me
   Unload frm180301_01
   'Add By Sindy 2022/10/28
   If TypeName(m_PrevForm) <> "frm160005" And TypeName(m_PrevForm) <> "frm160004" And TypeName(m_PrevForm) <> "frm160003" Then
   '2022/10/28 END
      Unload frm180301
   End If
End Sub

Private Sub cmdQueryNext_Click()
   'Modify By Sindy 2013/6/21
'   frm180301_01.Show
'   frm180301_01.PubShowNextData
   m_PrevForm.Show
   Unload Me
   'Add By Sindy 2022/10/28
   '下班逾30分鐘原因確認(frm160018_1) Add By Sindy 2025/10/30
   If TypeName(m_PrevForm) <> "frm160005" And _
      TypeName(m_PrevForm) <> "frm160004" And _
      TypeName(m_PrevForm) <> "frm160003" And _
      TypeName(m_PrevForm) <> "frm160018_1" Then
   '2022/10/28 END
      m_PrevForm.PubShowNextData
   End If
   '2013/6/21 END
End Sub

Private Sub CboB1002_Click()
   If Left(CboB1002.Text, 2) = 表單類別_請假 Then
      txtB1008_2.Visible = True
      txtB1008_14.Visible = True 'Add By Sindy 2024/12/10
      Label10.Visible = True
      CboB1008.Visible = True
      Label1(2).Visible = True
      txtB1006.Visible = True
      Frame01.Visible = True
      Frame02.Visible = False
      Frame03.Visible = False
   ElseIf Left(CboB1002.Text, 2) = 表單類別_加班 Then
      txtB1008_2.Visible = False
      txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
      Label10.Visible = False
      CboB1008.Visible = False
      Label1(2).Visible = False
      txtB1006.Visible = False
      Frame01.Visible = False
      Frame02.Visible = True
      Frame03.Visible = False
      Frame02.Top = 2040 'Add By Sindy 2016/11/24
      Frame02.Left = 900 '1140
   ElseIf Left(CboB1002.Text, 2) = 表單類別_出差 Then
      txtB1008_2.Visible = False
      txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
      Label10.Visible = False
      CboB1008.Visible = False
      Label1(2).Visible = True
      txtB1006.Visible = True
      Frame01.Visible = True
      Frame02.Visible = False
      Frame03.Visible = True
   End If
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
   
   If IsNull(rsSrcTmp.Fields("B1022")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1022")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("B1022"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1023")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1023")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("B1023"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1024")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1024")) = False Then
         strTemp = rsSrcTmp.Fields("B1024")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1025")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1025")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("B1025"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1026")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1026")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("B1026"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("B1027")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1027")) = False Then
         strTemp = rsSrcTmp.Fields("B1027")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label26.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ")
   Label27.Caption = "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID2(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("SA10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SA10")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("SA10"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SA11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SA11")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SA11"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SA12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SA12")) = False Then
         strTemp = rsSrcTmp.Fields("SA12")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SA13")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SA13")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("SA13"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SA14")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SA14")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SA14"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SA15")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SA15")) = False Then
         strTemp = rsSrcTmp.Fields("SA15")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label26.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ")
   Label27.Caption = "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID3(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("so07")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("so07")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("so07"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("so08")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("so08")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("so08"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("so09")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("so09")) = False Then
         strTemp = rsSrcTmp.Fields("so09")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("so10")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("so10")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("so10"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("so11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("so11")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("so11"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("so12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("so12")) = False Then
         strTemp = rsSrcTmp.Fields("so12")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label26.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ")
   Label27.Caption = "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID4(ByRef rsSrcTmp As ADODB.Recordset)
   Dim strTemp As String
   Dim strCName As String
   Dim strCDate As String
   Dim strCTime As String
   Dim strUName As String
   Dim strUDate As String
   Dim strUTime As String
   
   If IsNull(rsSrcTmp.Fields("SB11")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SB11")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("SB11"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SB12")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SB12")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SB12"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SB13")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SB13")) = False Then
         strTemp = rsSrcTmp.Fields("SB13")
         strCTime = Format(strTemp, "##:##:##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SB14")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SB14")) = False Then
         strUName = GetStaffName(rsSrcTmp.Fields("SB14"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SB15")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SB15")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("SB15"))
         strUDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("SB16")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("SB16")) = False Then
         strTemp = rsSrcTmp.Fields("SB16")
         strUTime = Format(strTemp, "##:##:##")
      End If
   End If
   
   ' 設定CUID中的文字
   Label26.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ")
   Label27.Caption = "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub

'Add By Sindy 2013/6/21
Public Sub SetParent(ByRef fm As Form)
   Set m_PrevForm = fm
End Sub
