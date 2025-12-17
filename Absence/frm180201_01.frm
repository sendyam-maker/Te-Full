VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180201_01 
   BorderStyle     =   1  '單線固定
   Caption         =   "簽核作業"
   ClientHeight    =   5750
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8950
   Tag             =   "加班資料"
   Begin VB.TextBox txtB1008_14 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      Height          =   195
      Left            =   4860
      Locked          =   -1  'True
      TabIndex        =   59
      TabStop         =   0   'False
      Text            =   "@可補休：剩餘 3.5 天"
      Top             =   960
      Width           =   3990
   End
   Begin VB.ComboBox cboETime 
      Height          =   300
      ItemData        =   "frm180201_01.frx":0000
      Left            =   1800
      List            =   "frm180201_01.frx":0002
      Locked          =   -1  'True
      Style           =   2  '單純下拉式
      TabIndex        =   47
      Top             =   2400
      Visible         =   0   'False
      Width           =   1005
   End
   Begin VB.ComboBox cboSTime 
      Height          =   300
      ItemData        =   "frm180201_01.frx":0004
      Left            =   1800
      List            =   "frm180201_01.frx":0006
      Locked          =   -1  'True
      Style           =   2  '單純下拉式
      TabIndex        =   46
      Top             =   2070
      Visible         =   0   'False
      Width           =   1005
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
      TabIndex        =   13
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1007_2 
      Height          =   300
      Left            =   4140
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   14
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1006 
      Height          =   300
      Left            =   3780
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   12
      Top             =   1290
      Width           =   945
   End
   Begin VB.TextBox txtB1008_2 
      Appearance      =   0  '平面
      BorderStyle     =   0  '沒有框線
      Height          =   195
      Left            =   4860
      Locked          =   -1  'True
      TabIndex        =   8
      TabStop         =   0   'False
      Text            =   "@特別假：7天  已休3天"
      Top             =   750
      Width           =   3990
   End
   Begin VB.Frame Frame02 
      BorderStyle     =   0  '沒有框線
      Height          =   645
      Left            =   5730
      TabIndex        =   33
      Top             =   2040
      Visible         =   0   'False
      Width           =   1875
      Begin VB.TextBox txtB101213 
         Enabled         =   0   'False
         Height          =   315
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   18
         Top             =   330
         Width           =   705
      End
      Begin VB.TextBox txtB1030 
         Enabled         =   0   'False
         Height          =   315
         Left            =   720
         Locked          =   -1  'True
         MaxLength       =   4
         TabIndex        =   17
         Top             =   0
         Width           =   705
      End
      Begin VB.Label Label16 
         AutoSize        =   -1  'True
         Caption         =   "假日-共                     時"
         Height          =   180
         Left            =   30
         TabIndex        =   35
         Top             =   360
         Width           =   1725
      End
      Begin VB.Label Label18 
         AutoSize        =   -1  'True
         Caption         =   "實際時數"
         Height          =   180
         Left            =   30
         TabIndex        =   34
         Top             =   60
         Width           =   720
      End
   End
   Begin VB.Frame Frame03 
      BorderStyle     =   0  '沒有框線
      Height          =   885
      Left            =   5730
      TabIndex        =   29
      Top             =   1140
      Visible         =   0   'False
      Width           =   3135
      Begin VB.TextBox txtB1014 
         Height          =   315
         Left            =   540
         Locked          =   -1  'True
         MaxLength       =   1
         TabIndex        =   19
         Top             =   120
         Width           =   225
      End
      Begin MSForms.TextBox txtB1015 
         Height          =   285
         Left            =   540
         TabIndex        =   53
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
         TabIndex        =   32
         Top             =   180
         Width           =   540
      End
      Begin VB.Label Label20 
         AutoSize        =   -1  'True
         Caption         =   "(1:長程 2:短程 3:大陸 4:國外)"
         Height          =   180
         Left            =   780
         TabIndex        =   31
         Top             =   180
         Width           =   2235
      End
      Begin VB.Label Label19 
         AutoSize        =   -1  'True
         Caption         =   "地點："
         Height          =   180
         Left            =   30
         TabIndex        =   30
         Top             =   480
         Width           =   540
      End
   End
   Begin VB.ComboBox CboB1002 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      Height          =   300
      ItemData        =   "frm180201_01.frx":0008
      Left            =   960
      List            =   "frm180201_01.frx":000A
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
      ItemData        =   "frm180201_01.frx":000C
      Left            =   3300
      List            =   "frm180201_01.frx":000E
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
      TabIndex        =   11
      Top             =   1620
      Width           =   585
   End
   Begin VB.TextBox txtB1004 
      Height          =   300
      Left            =   1380
      Locked          =   -1  'True
      MaxLength       =   7
      TabIndex        =   9
      Top             =   1290
      Width           =   945
   End
   Begin VB.TextBox txtB1005_1 
      Height          =   300
      Left            =   960
      Locked          =   -1  'True
      MaxLength       =   2
      TabIndex        =   10
      Top             =   1620
      Width           =   585
   End
   Begin VB.CommandButton cmdQueryNext 
      Caption         =   "查詢下一筆(&N)"
      Height          =   360
      Left            =   6660
      TabIndex        =   2
      Top             =   30
      Width           =   1365
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "退回當事人(&B)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   5310
      TabIndex        =   1
      Top             =   30
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同意(&O)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   4500
      TabIndex        =   0
      Top             =   30
      Width           =   800
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束(&X)"
      CausesValidation=   0   'False
      Height          =   360
      Left            =   8040
      TabIndex        =   3
      Top             =   30
      Width           =   800
   End
   Begin VB.Frame Frame01 
      BorderStyle     =   0  '沒有框線
      Height          =   495
      Left            =   3150
      TabIndex        =   36
      Top             =   2070
      Width           =   1965
      Begin VB.TextBox txtB1010 
         Height          =   315
         Left            =   1110
         Locked          =   -1  'True
         MaxLength       =   5
         TabIndex        =   16
         Top             =   30
         Width           =   525
      End
      Begin VB.TextBox txtB1009 
         Height          =   315
         Left            =   270
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   15
         Top             =   30
         Width           =   525
      End
      Begin VB.Label Label9 
         AutoSize        =   -1  'True
         Caption         =   "共              日               時"
         Height          =   180
         Left            =   60
         TabIndex        =   37
         Top             =   90
         Width           =   1845
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Bindings        =   "frm180201_01.frx":0010
      Height          =   1635
      Left            =   4710
      TabIndex        =   20
      Top             =   3810
      Width           =   4215
      _ExtentX        =   7444
      _ExtentY        =   2893
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
   Begin MSForms.Label Label26 
      Height          =   195
      Left            =   240
      TabIndex        =   58
      Top             =   5520
      Width           =   7905
      VariousPropertyBits=   27
      Caption         =   "CREATE :                                                    UPDATE : "
      Size            =   "13944;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtB1003_2 
      Height          =   285
      Left            =   1650
      TabIndex        =   57
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
   Begin MSForms.TextBox txtNote 
      Height          =   555
      Left            =   960
      TabIndex        =   56
      Top             =   3240
      Width           =   7845
      VariousPropertyBits=   -1466939365
      MaxLength       =   200
      ScrollBars      =   3
      Size            =   "13838;979"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtB1011 
      Height          =   585
      Left            =   960
      TabIndex        =   55
      Top             =   2670
      Width           =   7845
      VariousPropertyBits=   -1466939361
      ScrollBars      =   3
      Size            =   "13838;1032"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtB1207 
      Height          =   1635
      Left            =   990
      TabIndex        =   54
      Top             =   3810
      Width           =   3675
      VariousPropertyBits=   -1466939361
      ScrollBars      =   3
      Size            =   "6482;2884"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label LblEndW 
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   4770
      TabIndex        =   52
      Top             =   1350
      Width           =   825
   End
   Begin VB.Label LblStarW 
      ForeColor       =   &H00808000&
      Height          =   195
      Left            =   2370
      TabIndex        =   51
      Top             =   1350
      Width           =   825
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "迄日下班時段："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   5
      Left            =   510
      TabIndex        =   49
      Top             =   2460
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "起日上班時段："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   3
      Left            =   510
      TabIndex        =   48
      Top             =   2130
      Visible         =   0   'False
      Width           =   1260
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "您的意見："
      Height          =   180
      Left            =   30
      TabIndex        =   45
      Top             =   3270
      Width           =   900
   End
   Begin VB.Label Label23 
      AutoSize        =   -1  'True
      Caption         =   "簽核意見："
      Height          =   180
      Left            =   30
      TabIndex        =   44
      Top             =   3810
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "表單編號："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   1
      Left            =   30
      TabIndex        =   43
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
      TabIndex        =   42
      Top             =   1650
      Width           =   255
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "時"
      Height          =   180
      Left            =   3930
      TabIndex        =   41
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "日期"
      Height          =   180
      Index           =   2
      Left            =   3390
      TabIndex        =   40
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label Label22 
      AutoSize        =   -1  'True
      Caption         =   "            迄"
      Height          =   180
      Left            =   3570
      TabIndex        =   39
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "分"
      Height          =   180
      Left            =   4770
      TabIndex        =   38
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
      TabIndex        =   28
      Top             =   750
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      ForeColor       =   &H00000000&
      Height          =   180
      Index           =   0
      Left            =   30
      TabIndex        =   27
      Top             =   420
      Width           =   900
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "假別："
      Height          =   180
      Left            =   2730
      TabIndex        =   26
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
      TabIndex        =   25
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "時間：                 起                    "
      Height          =   180
      Left            =   390
      TabIndex        =   24
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
      TabIndex        =   23
      Top             =   1350
      Width           =   360
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "時"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   1560
      TabIndex        =   22
      Top             =   1710
      Width           =   180
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "事由："
      Height          =   180
      Left            =   390
      TabIndex        =   21
      Top             =   2760
      Width           =   540
   End
   Begin VB.Label Label12 
      ForeColor       =   &H000000C0&
      Height          =   525
      Left            =   3300
      TabIndex        =   50
      Top             =   420
      Width           =   5625
   End
End
Attribute VB_Name = "frm180201_01"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Create by Sindy 2011/8/8
Option Explicit

Dim m_B1003 As String
Dim m_B1017 As String
Dim m_B1018 As String
Dim i As Integer, j As Integer
Dim strB1102 As String, strB1103 As String
Dim strUpdDate As String, strUpdTime As String
Dim strContent As String, strSubject As String
Dim dblPrevRow As Double


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
   strSql = "Select B1001,B1002,B1003,B1004,substr(ltrim(to_char('0000'||to_char(B1005),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1005),'0000')),3,2) B1005,B1006,substr(ltrim(to_char('0000'||to_char(B1007),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1007),'0000')),3,2) B1007,B1008||' '||AC03 B1008,B1009,B1010,B1011,B1012,B1013,B1014,B1015,B1016,B1017," & B1018CName & " B1018,B1019,B1020,B1021,B1022,B1023,B1024,B1025,B1026,B1027,substr(ltrim(to_char('0000'||to_char(B1028),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1028),'0000')),3,2) B1028,substr(ltrim(to_char('0000'||to_char(B1029),'0000')),1,2)||':'||substr(ltrim(to_char('0000'||to_char(B1029),'0000')),3,2) B1029,B1030 " & _
            "From ABS010, allcode " & _
            "Where ac01(+)='04' and B1008=ac02(+) " & _
            "and B1001='" & Me.txtB1001 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   m_B1003 = "": m_B1017 = "": m_B1018 = ""
   If rsTmp.RecordCount > 0 Then
      If Not IsNull(rsTmp.Fields("B1001")) Then txtB1001 = rsTmp.Fields("B1001")
      If Not IsNull(rsTmp.Fields("B1002")) Then CboB1002 = GetB1002Value(rsTmp.Fields("B1002"))
      If Not IsNull(rsTmp.Fields("B1003")) Then txtB1003 = rsTmp.Fields("B1003"): m_B1003 = rsTmp.Fields("B1003"): txtB1003_2 = GetPrjSalesNM(rsTmp.Fields("B1003"))
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
      
      'Add By Sindy 2021/8/13
      SetB102829Combo cboSTime, 1, txtB1004, txtB1003
      SetB102829Combo cboETime, 2, txtB1004, txtB1003
      '2021/8/13 END
      
      'Add By Sindy 2020/8/14 顯示星期幾
      If Val(txtB1004) > 0 Then
         LblStarW.Caption = "(" & GetWeekDay(CDate(Format(DBDATE(txtB1004), "####/##/##"))) & ")"
      End If
      If Val(txtB1006) > 0 Then
         LblEndW.Caption = "(" & GetWeekDay(CDate(Format(DBDATE(txtB1006), "####/##/##"))) & ")"
      End If
      '2020/8/14 END
      
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
      'If Not IsNull(rsTmp.Fields("B1016")) Then m_B1016 = rsTmp.Fields("B1016")
      If Not IsNull(rsTmp.Fields("B1017")) Then m_B1017 = rsTmp.Fields("B1017")
      If Not IsNull(rsTmp.Fields("B1018")) Then Call GetB1018CodeOrCName(m_B1018, rsTmp.Fields("B1018"))
      
      If Not IsNull(rsTmp.Fields("B1028")) And rsTmp.Fields("B1028") <> "00:00" Then
         For i = 0 To cboSTime.ListCount - 1
            If cboSTime.List(i) = Format(Format(rsTmp.Fields("B1028"), "hhmm"), "00:00") Then
               cboSTime.ListIndex = i
               Exit For
            End If
         Next i
         Label1(3).Visible = True
         cboSTime.Visible = True
      Else
         Label1(3).Visible = False
         cboSTime.Visible = False
      End If
      If Not IsNull(rsTmp.Fields("B1029")) And rsTmp.Fields("B1029") <> "00:00" Then
         For i = 0 To cboETime.ListCount - 1
            If cboETime.List(i) = Format(Format(rsTmp.Fields("B1029"), "hhmm"), "00:00") Then
               cboETime.ListIndex = i
               Exit For
            End If
         Next i
         Label1(5).Visible = True
         cboETime.Visible = True
      Else
         Label1(5).Visible = False
         cboETime.Visible = False
      End If
      
      If IsNull(rsTmp.Fields("B1019")) Then
         txtB1008_2.Visible = True
         'Modify By Sindy 2014/12/3
         'txtB1008_2 = GetCurrSpecRestDay(Trim(txtB1003))
         txtB1008_2 = GetCurrSpecRestDay(Trim(txtB1003), , Left(txtB1004, 3))
         '2014/12/3 END
         'Add By Sindy 2024/12/10
         txtB1008_14.Visible = True
         txtB1008_14 = GetCurrFor14RestDay(Trim(txtB1003), , txtB1004)
         '2024/12/10 END
      Else
         txtB1008_2.Visible = False
         txtB1008_14.Visible = False 'Add By Sindy 2024/12/10
      End If
      
      Call CboB1002_Click
      
      Call UpdateCUID(rsTmp)
   Else
      Screen.MousePointer = vbDefault
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   rsTmp.Close
   
   If Trim(txtB1001) <> "" Then
      '表單流程備註檔
      SetABS012TextBox txtB1207, txtB1001
      '表單簽核檔
      strSql = "SELECT ST02||nvl(B1108,'') 簽核人員," & B1102CName & " 身份,sqldateT(B1105) 日期,sqltime6(B1106) 時間," & B1107CName & " 簽核結果,B1104 FROM ABS011,Staff WHERE B1101='" & txtB1001 & "' and B1104=ST01(+) order by B1102,B1103 asc "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         Set GRD1.Recordset = rsTmp
      End If
   End If
   
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
   'Add By Sindy 2012/2/13
   '檢查表單職代是否符合人事職代裡的設定資料
   If ChkIsDutyAgent(Trim(txtB1001), Trim(txtB1003)) = False Then
      MsgBox "提醒：此表單職代並非系統中所設定的職代！", vbExclamation
   End If
   '2012/2/13 End
   
   'Add By Sindy 2015/12/25 增加檢查同仁加班合計是否有超過46小時
   'Modify By Sindy 2016/12/26
   'Label12 = PUB_PerFormRemindMsg(Left(CboB1002, 2), "1", txtB1003, txtB1004, txtB1012, txtB1013, False)
   Label12 = PUB_PerFormRemindMsg(Left(CboB1002, 2), "1", txtB1003, txtB1004, txtB101213, False) 'Modify By Sindy 2021/7/23 + txtB101213
   
   'Add By Sindy 2020/5/28
   'Modify By Sindy 2020/10/26
   'Call PUB_ChkSerialRest_ToSir(txtB1001)
   Call PUB_ChkSerialRest_ToSir(txtB1001, , , Left(CboB1002, 2), _
         txtB1003, DBDATE(txtB1004), DBDATE(txtB1006), Val(txtB1009), Val(txtB1010))
   
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
   GRD1.Clear
   SetGrd
   LblStarW.Caption = Empty 'Add By Sindy 2020/8/14
   LblEndW.Caption = Empty 'Add By Sindy 2020/8/14
End Sub

Private Sub cmdBack_Click()
Dim strTo As String
   
On Error GoTo ErrHand
   
   If Trim(txtNote.Text) = "" Then
      MsgBox "您的意見不可以空白！", vbExclamation
      txtNote.SetFocus
      Exit Sub
   End If
   
   'Add by Sindy 2021/5/28 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Sub
   End If
   '2021/5/28 END
   
   Screen.MousePointer = vbHourglass
   
   m_B1018 = 退回
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   cmdBack.Enabled = False 'Add By Sindy 2013/12/6
   cnnConnection.BeginTrans
   
   '簽核檔
   strSql = "SELECT B1102,B1103 FROM ABS011 " & _
            "WHERE B1101='" & txtB1001 & "' and B1104='" & strUserNum & "' and B1107 is null order by B1102,B1103 "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strB1102 = RsTemp.Fields("B1102")
      strB1103 = RsTemp.Fields("B1103")
   End If
   strSql = "update ABS011 set " & _
            "B1105='" & strUpdDate & "'" & _
            ",B1106='" & strUpdTime & "'" & _
            ",B1107='2'" & _
            " where B1101='" & txtB1001 & "' and B1102='" & strB1102 & "' and B1103=" & strB1103 & " and B1104='" & strUserNum & "' "
   cnnConnection.Execute strSql
   '流程備註檔
   If Trim(txtNote.Text) <> "" Then
      strSql = GetInsertABS012Sql(Trim(txtB1001), strUserNum, strUpdDate, strUpdTime, m_B1018, ChgSQL(Trim(txtNote.Text)))
      cnnConnection.Execute strSql
   End If
   '主檔
   strSql = "update ABS010 set " & _
            "B1016='" & strUserNum & "'" & _
            ",B1017='" & txtB1003 & "'" & _
            ",B1018='" & m_B1018 & "'" & _
            " where B1001='" & txtB1001 & "' "
   cnnConnection.Execute strSql
   
   cnnConnection.CommitTrans
   
   '發E-Mail通知當事人及已簽核的審核主管
   strTo = GetBossB1107_2_1(txtB1001)
   If strTo <> "" Then
      strContent = GetEMailContent(txtB1001, strSubject, 退回通知主管, "被(" & strUserName & ")退回")
      If Trim(txtNote.Text) <> "" Then
         strSubject = strSubject & "；退回原因：" & Trim(txtNote.Text)
      End If
      'PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True
      '正本：當事人　副本：已簽核的審核主管
      PUB_SendMail strUserNum, txtB1003, "", strSubject, strContent, , , , , , strTo, , , , True
   Else
      strContent = GetEMailContent(txtB1001, strSubject)
      If Trim(txtNote.Text) <> "" Then
         strSubject = strSubject & "；退回原因：" & Trim(txtNote.Text)
      End If
      PUB_SendMail strUserNum, txtB1003, "", strSubject, strContent, , , , , , , , , , True
   End If
   
   Screen.MousePointer = vbDefault
   
   'tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Unload Me
   frm180201.Show
   frm180201.PubShowNextData
   Exit Sub
   
ErrHand:
   cmdBack.Enabled = True 'Add By Sindy 2013/12/6
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox "退回失敗！" & vbCrLf & Err.Description
End Sub

Private Sub cmdExit_Click()
   Unload Me
   frm180201.QueryData
   frm180201.Show
End Sub

Private Sub cmdok_Click()
'Add By Sindy 2012/1/4
Dim strB1001 As String, strB1002 As String, strB1003 As String
Dim strB1004 As String, strB1005 As String
Dim strB1006 As String, strB1007 As String
Dim strB1008 As String, strB1009 As String, strB1010 As String
Dim strB101213 As String
Dim strB1014 As String, strB1015 As String
Dim strB1028 As String, strB1029 As String
'2012/1/4 End
Dim strB1030 As String 'Add By Sindy 2016/12/26
   
On Error GoTo ErrHand
   
   'Add by Sindy 2021/5/28 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Sub
   End If
   '2021/5/28 END
   
   Screen.MousePointer = vbHourglass
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   cmdOK.Enabled = False 'Add By Sindy 2013/12/6
   cnnConnection.BeginTrans
   
   '流程備註檔
   If Trim(txtNote.Text) <> "" Then
      strSql = GetInsertABS012Sql(Trim(txtB1001), strUserNum, strUpdDate, strUpdTime, m_B1018, ChgSQL(Trim(txtNote.Text)))
      cnnConnection.Execute strSql
   End If
   
   Do While m_B1017 = strUserNum '以防下一處理人員再讀取到此時的簽核主管
      '簽核檔
      strSql = "update ABS011 set " & _
               "B1105='" & strUpdDate & "'" & _
               ",B1106='" & strUpdTime & "'" & _
               ",B1107='1'" & _
               " where B1101='" & txtB1001 & "' and B1104='" & strUserNum & "' and B1107 is null "
      cnnConnection.Execute strSql
      '讀取下一處理人員
      If GetNextProPerson(Trim(txtB1001), Trim(txtB1003), m_B1017, strUserNum) = False Then GoTo ErrHand
   Loop
   
   If m_B1017 = "M21" Then 'Modify By Sindy 2011/10/17 人事處不簽收,最高審核主管簽核完畢,系統自動簽收進人事系統
'      If AutoM21Receive = False Then
'         Screen.MousePointer = vbDefault
'         Exit Sub
'      End If
      'Modify By Sindy 2012/1/4 寫成共用函數
      strB1001 = Trim(txtB1001)
      strB1002 = Trim(Left(CboB1002, 2))
      strB1003 = Trim(txtB1003)
      strB1004 = DBDATE(txtB1004)
      strB1005 = txtB1005_1 & Format("00" & txtB1005_2, "00")
      strB1006 = DBDATE(txtB1006)
      strB1007 = txtB1007_1 & Format("00" & txtB1007_2, "00")
      strB1008 = Trim(Left(CboB1008, 2))
      strB1009 = txtB1009
      strB1010 = txtB1010
      strB1030 = txtB1030
      strB101213 = txtB101213
      strB1014 = txtB1014
      strB1015 = txtB1015
      strB1028 = IIf(cboSTime.Visible = False, "", Format(cboSTime, "hhmm"))
      strB1029 = IIf(cboETime.Visible = False, "", Format(cboETime, "hhmm"))
      If PUB_AutoM21Receive(strUserNum, strUpdDate, strUpdTime, strB1001, strB1002, strB1003, strB1004, strB1005, strB1006, strB1007, strB1008, strB1009, strB1010, strB1030, strB101213, strB1014, strB1015, strB1028, strB1029, strSubject, strContent) = False Then
         'Screen.MousePointer = vbDefault
         'Exit Sub
         GoTo ErrHand
      End If
      '2012/1/4 End
   Else
      '發E-Mail通知下一處理人員
      cnnConnection.CommitTrans
      strContent = GetEMailContent(txtB1001, strSubject)
      PUB_SendMail strUserNum, m_B1017, "", strSubject, strContent, , , , , , , , , , True
   End If
   
   Screen.MousePointer = vbDefault
   
   'tmpBol = fnCancelNowFormAndShowParentForm(Me)
   Unload Me
   frm180201.Show
   frm180201.PubShowNextData
   Exit Sub
   
ErrHand:
   cmdOK.Enabled = True 'Add By Sindy 2013/12/6
   Screen.MousePointer = vbDefault
   cnnConnection.RollbackTrans
   MsgBox "簽核失敗！" & vbCrLf & Err.Description
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   txtB1001.BackColor = &H8000000F
   txtB1003.BackColor = &H8000000F
   txtB1003_2.BackColor = &H8000000F
   'txtB1008_2.BackColor = &H8000000F
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
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180201_01 = Nothing
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

Private Sub cmdQueryNext_Click()
   Unload Me
   frm180201.Show
   frm180201.PubShowNextData
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
      Frame02.Left = 3150 '900
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

Private Sub txtNote_GotFocus()
   InverseTextBox txtNote
   OpenIme
End Sub

Private Sub txtNote_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
   If Button = 2 Then Forms(0).PopupMenu2 txtNote
End Sub

Private Sub txtNote_Validate(Cancel As Boolean)
If txtNote <> "" Then
   If CheckLengthIsOK(txtNote, txtNote.MaxLength) = False Then
      Call txtNote_GotFocus
      Cancel = True
      Exit Sub
   End If
End If
CloseIme
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
   'm_B1023 = ""
   If IsNull(rsSrcTmp.Fields("B1023")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("B1023")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("B1023"))
         strCDate = Format(strTemp, "###/##/##")
         'm_B1023 = rsSrcTmp.Fields("B1023")
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
              " " & strCTime & String(10, " ") & _
              "UPDATE : " & strUName & " " & _
              " " & strUDate & " " & _
              " " & strUTime
End Sub
