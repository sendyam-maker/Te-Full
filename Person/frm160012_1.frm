VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160012_1 
   BorderStyle     =   1  '單線固定
   Caption         =   "異常確認作業"
   ClientHeight    =   5750
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5750
   ScaleWidth      =   8950
   Begin VB.TextBox Text1 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   585
      Left            =   6360
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   41
      Text            =   "frm160012_1.frx":0000
      Top             =   3360
      Width           =   1725
   End
   Begin VB.TextBox textB1411 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   5910
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   35
      Top             =   3360
      Width           =   405
   End
   Begin VB.TextBox textB1404 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   270
      Left            =   1290
      Locked          =   -1  'True
      TabIndex        =   3
      Top             =   1140
      Width           =   1035
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "打卡明細"
      Height          =   375
      Left            =   7350
      TabIndex        =   11
      Top             =   600
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  '沒有框線
      Height          =   435
      Left            =   6330
      TabIndex        =   29
      Top             =   2700
      Width           =   2475
      Begin VB.TextBox textSA06 
         Height          =   285
         Left            =   720
         MaxLength       =   3
         TabIndex        =   7
         Top             =   120
         Width           =   675
      End
      Begin VB.Label Label22 
         AutoSize        =   -1  'True
         Caption         =   "分"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   9
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   180
         Left            =   1470
         TabIndex        =   31
         Top             =   140
         Width           =   190
      End
      Begin VB.Label Label21 
         AutoSize        =   -1  'True
         Caption         =   "曠職："
         Height          =   180
         Left            =   150
         TabIndex        =   30
         Top             =   135
         Width           =   540
      End
   End
   Begin VB.TextBox textB1402 
      Height          =   270
      Left            =   1290
      MaxLength       =   7
      TabIndex        =   1
      Top             =   510
      Width           =   1035
   End
   Begin VB.ComboBox cboB1403 
      Height          =   300
      ItemData        =   "frm160012_1.frx":003A
      Left            =   1290
      List            =   "frm160012_1.frx":0044
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   810
      Width           =   1125
   End
   Begin VB.ComboBox cboSTime 
      Height          =   300
      ItemData        =   "frm160012_1.frx":0058
      Left            =   4770
      List            =   "frm160012_1.frx":0068
      Style           =   2  '單純下拉式
      TabIndex        =   4
      Top             =   1140
      Visible         =   0   'False
      Width           =   1125
   End
   Begin VB.TextBox textB1409 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   285
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   1
      TabIndex        =   23
      Top             =   3000
      Width           =   405
   End
   Begin VB.TextBox textB1408 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Enabled         =   0   'False
      Height          =   270
      Left            =   1290
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   13
      Top             =   2700
      Width           =   1035
   End
   Begin VB.TextBox textB1407 
      Height          =   885
      Left            =   1290
      MaxLength       =   200
      TabIndex        =   6
      Top             =   1770
      Width           =   6195
   End
   Begin VB.TextBox textB1405 
      Height          =   285
      Left            =   1290
      MaxLength       =   1
      TabIndex        =   5
      Top             =   1440
      Width           =   405
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "人事確認"
      Default         =   -1  'True
      Height          =   375
      Left            =   3810
      TabIndex        =   8
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdBack 
      Caption         =   "回前畫面"
      Height          =   375
      Left            =   6800
      TabIndex        =   9
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   375
      Left            =   7950
      TabIndex        =   10
      Top             =   120
      Width           =   795
   End
   Begin VB.TextBox textB1401 
      Height          =   270
      Left            =   1290
      MaxLength       =   6
      TabIndex        =   0
      Top             =   210
      Width           =   1035
   End
   Begin VB.CommandButton cmdABS 
      Caption         =   "查詢當日請假資料"
      Height          =   375
      Left            =   4965
      TabIndex        =   12
      Top             =   120
      Width           =   1785
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   765
      Left            =   6960
      TabIndex        =   14
      Top             =   4590
      Visible         =   0   'False
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   1341
      _Version        =   393216
      Cols            =   14
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
      _Band(0).Cols   =   14
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdList 
      Height          =   2025
      Left            =   1650
      TabIndex        =   33
      Top             =   3330
      Width           =   3210
      _ExtentX        =   5662
      _ExtentY        =   3581
      _Version        =   393216
      Cols            =   3
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "刷卡日期|刷卡時間|人事補登"
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
      _Band(0).Cols   =   3
   End
   Begin MSForms.Label LblB1408 
      Height          =   225
      Left            =   2400
      TabIndex        =   44
      Top             =   2730
      Width           =   1365
      BackColor       =   12632256
      VariousPropertyBits=   27
      Size            =   "2408;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label23 
      Height          =   195
      Left            =   60
      TabIndex        =   43
      Top             =   5490
      Width           =   4125
      VariousPropertyBits=   27
      Caption         =   "CREATE : "
      Size            =   "7276;344"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label Label12 
      Height          =   225
      Left            =   2400
      TabIndex        =   42
      Top             =   240
      Width           =   1365
      BackColor       =   12632256
      VariousPropertyBits=   27
      Size            =   "2408;397"
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "(1.請假 2.遲到/早退 3.忘打卡 4.洽公請主管批示 5.指紋異常 6.因公未打卡 7.颱風假 8.其他)"
      Height          =   240
      Left            =   1740
      TabIndex        =   20
      Top             =   1500
      Width           =   6960
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "處理結果："
      Height          =   180
      Left            =   4980
      TabIndex        =   40
      Top             =   3360
      Width           =   900
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "確認日期："
      Height          =   180
      Left            =   5010
      TabIndex        =   39
      Top             =   3990
      Width           =   900
   End
   Begin VB.Label LblB1412 
      Caption         =   "LblB1412"
      Height          =   255
      Left            =   5940
      TabIndex        =   38
      Top             =   3990
      Width           =   1305
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "確認時間："
      Height          =   180
      Left            =   5010
      TabIndex        =   37
      Top             =   4290
      Width           =   900
   End
   Begin VB.Label LblB1413 
      Caption         =   "LblB1413"
      Height          =   255
      Left            =   5940
      TabIndex        =   36
      Top             =   4290
      Width           =   1305
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "當日打卡明細："
      Height          =   180
      Left            =   360
      TabIndex        =   34
      Top             =   3390
      Width           =   1260
   End
   Begin VB.Label Label13 
      Caption         =   "備註：有修改個人確認或原因時，會發E-Mail通知當事人。"
      ForeColor       =   &H000000C0&
      Height          =   255
      Left            =   4230
      TabIndex        =   32
      Top             =   5460
      Width           =   4665
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "批示日期："
      Height          =   180
      Left            =   3870
      TabIndex        =   28
      Top             =   2730
      Width           =   900
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "上班時段："
      Height          =   180
      Left            =   3840
      TabIndex        =   27
      Top             =   1200
      Visible         =   0   'False
      Width           =   900
   End
   Begin VB.Label LblB1410 
      Caption         =   "LblB1410"
      Height          =   255
      Left            =   4800
      TabIndex        =   26
      Top             =   2700
      Width           =   1305
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "（Y.同意 N.不同意）"
      Height          =   180
      Left            =   1740
      TabIndex        =   25
      Top             =   3060
      Width           =   1635
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "主管批示："
      Height          =   180
      Left            =   360
      TabIndex        =   24
      Top             =   3060
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "主管代號："
      Height          =   180
      Index           =   0
      Left            =   360
      TabIndex        =   22
      Top             =   2730
      Width           =   900
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "未打卡原因："
      Height          =   180
      Left            =   180
      TabIndex        =   21
      Top             =   1830
      Width           =   1080
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "個人確認："
      Height          =   180
      Left            =   360
      TabIndex        =   19
      Top             =   1500
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "打卡時間："
      Height          =   180
      Left            =   360
      TabIndex        =   18
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "時　　段："
      Height          =   180
      Left            =   360
      TabIndex        =   17
      Top             =   870
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "日　　期："
      Height          =   180
      Left            =   360
      TabIndex        =   16
      Top             =   540
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   1
      Left            =   330
      TabIndex        =   15
      Top             =   240
      Width           =   930
   End
End
Attribute VB_Name = "frm160012_1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/17 Form2.0已修改
'Create By Sindy 2013/6/21
Option Explicit

' 變數宣告區
Dim m_EditMode As Integer
Public m_WorkType As Integer '作業功能 : 0.異常確認 1.新增正常時間異常打卡
Public m_B1401 As String '員工代號
Public m_B1402 As String '日期
Public m_B1403 As String '打卡類別
Public bolClose As Boolean
Dim m_B1405 As String, m_B1407 As String


Private Sub cmdBack_Click()
   Unload Me
   frm160012.Show
   If m_WorkType = 0 Then
      frm160012.cmdok_Click
   End If
End Sub

'打卡明細
Private Sub cmdDetail_Click()
   If textB1401 = "" Or textB1402 = "" Then Exit Sub
   Call frm180303_1.SetParent(Me)
   frm180303_1.m_B1401 = textB1401
   frm180303_1.m_B1402 = DBDATE(textB1402)
   If frm180303_1.QueryData = True Then
      frm180303_1.Show vbModal '強制回應表單
   Else
      Unload frm180303_1
   End If
End Sub

' 當日打卡明細初始化列表
Private Sub InitialGridList()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("刷卡日期", "刷卡時間", "人事補登")
   arrGridHeadWidth = Array(1000, 1000, 800)
   grdList.Visible = False
   grdList.Cols = UBound(arrGridHeadText) + 1
   grdList.Rows = 2
   For iRow = 0 To grdList.Cols - 1
      grdList.row = 0
      grdList.col = iRow
      grdList.Text = arrGridHeadText(iRow)
      grdList.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grdList.CellAlignment = flexAlignCenterCenter
   Next
   grdList.Visible = True
End Sub

' 當日打卡明細查詢
Public Function PollRecordQueryData() As Boolean
   Dim stSQL As String
   
   PollRecordQueryData = False
   
   Call Pub_GetSpecWorkHour(textB1401, textB1402) '特殊人員的工作時數 Add By Sindy 2025/9/4
   
   Screen.MousePointer = vbHourglass
   Me.grdList.MousePointer = flexHourglass
   InitialGridList
   
   stSQL = "select sqldatet(pr01) as 刷卡日期,sqltime6(pr02) as 刷卡時間,decode(pr08,999,'Y','') as 人事補登"
   stSQL = stSQL & " from staff,staffcarddata,pollrecord where scd01(+)=st01 and pr03(+)=scd02 and pr01>0" & _
                    " and st01='" & textB1401 & "' and pr01=" & DBDATE(textB1402)
   stSQL = stSQL & " order by pr02 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, stSQL)
   If intI = 1 Then
      Set grdList.Recordset = RsTemp
      grdList.row = 1
      PollRecordQueryData = True
'   Else
'      ShowNoData
'      Me.grdList.MousePointer = flexDefault
'      Screen.MousePointer = vbDefault
'      Exit Function
   End If
   Me.grdList.MousePointer = flexDefault
   Screen.MousePointer = vbDefault
End Function

Private Sub cmdExit_Click()
   Unload Me
   'Unload frm160012
   If frm160012.m_QueryType = 3 Then
      frm160012.cmdQuery3_Click
   ElseIf frm160012.m_QueryType = 2 Then
      frm160012.cmdQuery2_Click
   Else
      frm160012.cmdQuery_Click
   End If
   frm160012.Show
End Sub

'請假、出差、加班資料
Private Sub cmdABS_Click()
Dim rsTmp As New ADODB.Recordset
   
   grd2.Clear
   SetGrd2
   If PUB_QueryData_ABS(textB1401, textB1402, rsTmp) = True Then
      Set grd2.Recordset = rsTmp
      Call PubShowNextData
      Exit Sub
   End If
End Sub

Private Sub SetGrd2()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("V", "員工代號", "表單編號", "TableID", "SA02", "SA03")
   arrGridHeadWidth = Array(800, 800, 800, 800, 800, 800)
   'grd2.Visible = False
   grd2.Cols = UBound(arrGridHeadText) + 1
   grd2.Rows = 2
   For iRow = 0 To grd2.Cols - 1
      grd2.row = 0
      grd2.col = iRow
      grd2.Text = arrGridHeadText(iRow)
      grd2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      grd2.CellAlignment = flexAlignCenterCenter
   Next
   'grd2.Visible = True
End Sub

'查詢出缺勤明細資料
Public Sub PubShowNextData()
Dim i As Integer
   
   Me.Enabled = False
   For i = 1 To grd2.Rows - 1
      grd2.col = 0
      grd2.row = i
      If Trim(grd2.Text) = "V" Then
         grd2.Text = ""
         grd2.col = 2 '表單編號
         Screen.MousePointer = vbHourglass
         Me.Hide
         Call frm180301_03.SetParent(Me)
         If grd2.TextMatrix(i, 3) = "1" Then '出缺勤
            frm180301_03.txtB1001 = Pub_RplStr(grd2.Text)
            frm180301_03.QueryData
         Else
            frm180301_03.txtB1003 = Pub_RplStr(grd2.TextMatrix(i, 1))
            frm180301_03.m_SA02 = Pub_RplStr(grd2.TextMatrix(i, 4))
            frm180301_03.m_SA03 = Pub_RplStr(grd2.TextMatrix(i, 5))
            If grd2.TextMatrix(i, 3) = "2" Then '請假
               frm180301_03.QueryData_2
            ElseIf grd2.TextMatrix(i, 3) = "3" Then '加班
               frm180301_03.QueryData_3
            ElseIf grd2.TextMatrix(i, 3) = "4" Then '出差
               frm180301_03.QueryData_4
            End If
         End If
         frm180301_03.Show
         Screen.MousePointer = vbDefault
         Me.Enabled = True
         Exit Sub
      End If
   Next i
   Me.Enabled = True
End Sub

'查詢資料
Private Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
On Error GoTo ErrHand
   
   InitialGridList
   
   Screen.MousePointer = vbHourglass
   
   '打卡異常資料
   strSql = "select *" & _
            " from abs014" & _
            " where B1401='" & m_B1401 & "'" & _
            " and B1402=" & DBDATE(m_B1402) & _
            " and B1403='" & m_B1403 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      textB1401 = "" & rsTmp.Fields("B1401")
      Label12.Caption = GetPrjSalesNM(textB1401)
      textB1402 = ChangeWStringToTString("" & rsTmp.Fields("B1402"))
      If "" & rsTmp.Fields("B1403") = "A" Then
         cboB1403.ListIndex = 0
      Else
         cboB1403.ListIndex = 1
      End If
      If "" & rsTmp.Fields("B1404") <> "" Then
         textB1404.Text = Format(Right("000000" & rsTmp.Fields("B1404"), 6), "##:##:##")
      End If
      '已輸入個人確認時,不可修改
      If Trim("" & rsTmp.Fields("B1405")) <> "" Then
         m_B1405 = Trim("" & rsTmp.Fields("B1405"))
         textB1405.Text = "" & rsTmp.Fields("B1405")
         If textB1405.Text = "4" And Trim(textB1408.Text) <> "" Then
            textB1405.Enabled = False
            textB1407.Enabled = False
         End If
'      '可輸入
'      Else
'         textB1405.Enabled = True
'         textB1407.Enabled = True
'         '預設值 : 如果上班打卡時間是大於等於9:30時,則一定要請假
'         If textB1404 <> "" Then
'            If "" & rsTmp.Fields("B1403") = "A" And Val(Replace(Left(textB1404, Len(textB1404) - 2), ":", "")) >= 930 _
'               And (m_B1401 <> "99029" And m_B1401 <> "96006") Then
'               textB1405.Text = "1"
''               textB1405.Enabled = False
'            End If
'         End If
      End If
      m_B1407 = Trim("" & rsTmp.Fields("B1407"))
      textB1407.Text = "" & rsTmp.Fields("B1407")
'      If IsNull(rsTmp.Fields("B1406")) Or "" & rsTmp.Fields("B1406") = "" Then
'         cboSTime.ListIndex = 0
'      Else
'         If "" & rsTmp.Fields("B1406") = "800" Then
'            cboSTime.ListIndex = 1
'         ElseIf "" & rsTmp.Fields("B1406") = "830" Then
'            cboSTime.ListIndex = 2
'         ElseIf "" & rsTmp.Fields("B1406") = "900" Then
'            cboSTime.ListIndex = 3
'         End If
'      End If
      '主管
      textB1408.Text = "" & rsTmp.Fields("B1408")
      LblB1408.Caption = GetPrjSalesNM(textB1408)
      LblB1410.Caption = ChangeWStringToTString("" & rsTmp.Fields("B1410"))
      textB1409.Text = "" & rsTmp.Fields("B1409")
'      textB1411.Text = "" & rsTmp.Fields("B1411")
'      LblB1412.Caption = ChangeWStringToTString("" & rsTmp.Fields("B1412"))
'      LblB1413.Caption = "" & rsTmp.Fields("B1413")
      Call UpdateCUID(rsTmp)
      If textB1405.Enabled = True Then
         textB1405.TabIndex = 0
      Else
         textB1407.TabIndex = 0
      End If
      '主管批示不同意才需要輸入曠職(時)
      If textB1405 = "4" And textB1409 = "N" Then
         Frame1.Visible = True
         textSA06.SetFocus
      End If
      
      textB1411.Text = "" & rsTmp.Fields("B1411")
      LblB1412.Caption = ChangeWStringToTString("" & rsTmp.Fields("B1412"))
      If "" & rsTmp.Fields("B1413") <> "" Then
         LblB1413.Caption = Format("" & rsTmp.Fields("B1413"), "00:00:00")
      End If
      
      '當日打卡明細
      Call PollRecordQueryData
   Else
      ShowNoData
      rsTmp.Close
      Set rsTmp = Nothing
      Screen.MousePointer = vbDefault
      Exit Sub
   End If
   
   rsTmp.Close
   Set rsTmp = Nothing
   Screen.MousePointer = vbDefault
   Exit Sub
   
ErrHand:
   Set rsTmp = Nothing
   Screen.MousePointer = vbDefault
End Sub

' 更新 Create 及 Update 的人
Private Sub UpdateCUID(ByRef rsSrcTmp As ADODB.Recordset)
Dim strTemp As String
Dim strCName As String
Dim strCDate As String
Dim strCTime As String

   If IsNull(rsSrcTmp.Fields("b1414")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("b1414")) = False Then
         strCName = GetStaffName(rsSrcTmp.Fields("b1414"), True)
      End If
   End If
   If IsNull(rsSrcTmp.Fields("b1415")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("b1415")) = False Then
         strTemp = TAIWANDATE(rsSrcTmp.Fields("b1415"))
         strCDate = Format(strTemp, "###/##/##")
      End If
   End If
   If IsNull(rsSrcTmp.Fields("b1416")) = False Then
      If IsEmptyText(rsSrcTmp.Fields("b1416")) = False Then
         strTemp = Right("000000" & rsSrcTmp.Fields("b1416"), 6)
         strCTime = Format(strTemp, "00:00:00")
      End If
   End If
   
   '設定CUID中的文字
   Label23.Caption = "CREATE : " & strCName & " " & _
              " " & strCDate & " " & _
              " " & strCTime & String(10, " ")
End Sub

'確定
Private Sub cmdok_Click()
Dim strUpdDate As String
Dim strUpdTime As String
Dim strSubject As String, strContent As String, strTo As String
   
   If CheckDataValid = False Then Exit Sub
   If TxtValidate = False Then Exit Sub
   
On Error GoTo ErrHand
   
   strUpdDate = Format(Now, "YYYYMMDD")
   strUpdTime = Format(time, "HHMMSS")
   
   cnnConnection.BeginTrans
   '新增異常資料
   If m_EditMode = 1 Then
      '" & CNULL(Format(cboSTime.Text, "hhmm")) & "
      If textB1404 <> "" Then
         textB1404 = Replace(textB1404, ":", "")
      End If
      strSql = "insert into ABS014(b1401,b1402,b1403,b1404,b1405,b1406,b1407" & IIf(textB1405 <> "", ",b1411,b1412,b1413", "") & ")" & _
               " VALUES(" & CNULL(textB1401) & "," & CNULL(DBDATE(textB1402)) & "," & CNULL(Left(Trim(cboB1403), 1)) & "," & _
               CNULL(textB1404) & "," & CNULL(textB1405) & ",null," & CNULL(textB1407)
      If textB1405 <> "" Then
         strSql = strSql & ",'B'," & CNULL(strUpdDate) & "," & CNULL(strUpdTime)
      End If
      strSql = strSql & ")"
      Pub_SeekTbLog strSql
      cnnConnection.Execute strSql
   '確認
   Else
      '當事人尚未輸入個人確認欄位值,代表是人事處先確認
      If m_B1405 = "" Then
         'B.人事處先確認
         '",B1406=" & CNULL(Format(cboSTime.Text, "hhmm"))
         strSql = "update ABS014 set B1405='" & textB1405 & "'" & _
                                   ",B1407='" & textB1407 & "'" & _
                                   ",B1411='B'" & _
                                   ",B1412=" & strUpdDate & _
                                   ",B1413=" & strUpdTime & _
                  " where B1401='" & textB1401 & "' and B1402=" & DBDATE(textB1402) & " and B1403='" & Left(Trim(cboB1403), 1) & "'"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      Else
         'C.人事處已確認
         strSql = "update ABS014 set B1405='" & textB1405 & "'" & _
                                   ",B1407='" & textB1407 & "'" & _
                                   ",B1411='C'" & _
                                   ",B1412=" & strUpdDate & _
                                   ",B1413=" & strUpdTime & _
                  " where B1401='" & textB1401 & "' and B1402=" & DBDATE(textB1402) & " and B1403='" & Left(Trim(cboB1403), 1) & "'"
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      End If
   End If
   Select Case textB1405
      Case "2" '遲到
         '檢查出缺勤資料是否已存在,若不存在,則新增
         If IsStaffAssExist(textB1401, DBDATE(textB1402)) = False Then
            strSql = "insert into Staff_Assist(SA01,SA02,SA04) VALUES(" & CNULL(textB1401) & "," & CNULL(DBDATE(textB1402)) & ",1)"
         Else
            strSql = "update Staff_Assist set SA04=nvl(SA04,0)+1" & _
                     " where SA01='" & textB1401 & "' and SA02=" & DBDATE(textB1402)
         End If
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      Case "3" '忘打卡
         '檢查出缺勤資料是否已存在,若不存在,則新增
         If IsStaffAssExist(textB1401, DBDATE(textB1402)) = False Then
            strSql = "insert into Staff_Assist(SA01,SA02,SA03) VALUES(" & CNULL(textB1401) & "," & CNULL(DBDATE(textB1402)) & ",1)"
         Else
            strSql = "update Staff_Assist set SA03=nvl(SA03,0)+1" & _
                     " where SA01='" & textB1401 & "' and SA02=" & DBDATE(textB1402)
         End If
         Pub_SeekTbLog strSql
         cnnConnection.Execute strSql
      Case "4" '未打卡請主管批示
         If textB1409 = "N" And Val(textSA06) > 0 Then '記曠職
            '檢查出缺勤資料是否已存在,若不存在,則新增
            If IsStaffAssExist(textB1401, DBDATE(textB1402)) = False Then
               strSql = "insert into Staff_Assist(SA01,SA02,SA06) VALUES(" & CNULL(textB1401) & "," & CNULL(DBDATE(textB1402)) & "," & Val(textSA06) & ")"
            Else
               strSql = "update Staff_Assist set SA06=nvl(SA06,0)+" & Val(textSA06) & _
                        " where SA01='" & textB1401 & "' and SA02=" & DBDATE(textB1402)
            End If
            Pub_SeekTbLog strSql
            cnnConnection.Execute strSql
         End If
   End Select
   
   cnnConnection.CommitTrans
   
   '發EMail通知當事者
   strSubject = ""
   strContent = ""
   If m_EditMode = 1 Then '新增異常資料
      strSubject = GetPrjSalesNM(textB1401) & " " & ChangeWStringToTDateString(DBDATE(textB1402)) & " " & IIf(Left(Trim(cboB1403), 1) = "A", "上", "下") & "班打卡異常，人事處新增通知！"
   Else
      '有修改資料
      If (m_B1405 <> "" And m_B1405 <> Trim(textB1405)) Or _
         (m_B1407 <> "" And m_B1407 <> Trim(textB1407)) Then
         strSubject = GetPrjSalesNM(textB1401) & " " & ChangeWStringToTDateString(DBDATE(textB1402)) & " " & IIf(Left(Trim(cboB1403), 1) = "A", "上", "下") & "班打卡異常，人事處修改通知！"
      End If
   End If
   If strSubject <> "" Then
      strContent = "員工姓名：" & Label12 & vbCrLf
      strContent = strContent & "日　　期：" & ChangeTStringToTDateString(textB1402) & vbCrLf
      strContent = strContent & "時　　段：" & IIf(Left(cboB1403, 1) = "A", "上班", "下班") & vbCrLf
      strContent = strContent & "打卡時間：" & Format(textB1404, "##:##:##") & vbCrLf
      strContent = strContent & "個人確認：" & textB1405 & "  (1.請假 2.遲到/早退 3.忘打卡(扣10元) 4.洽公請主管批示 5.指紋異常 6.因公未打卡 7.颱風假)" & vbCrLf
      strContent = strContent & "未打卡原因：" & textB1407 & vbCrLf
      
      strTo = PUB_GetST59(textB1401)
      If IsNull(strTo) Or strTo = "" Then
         strTo = textB1401
      End If
      'Modify By Sindy 2017/10/18 劉經理說不寄信
'      PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True
   End If
   
   If m_WorkType = 0 Then
      Call cmdBack_Click
   Else
      Call cmdExit_Click
   End If
   Exit Sub
   
ErrHand:
    cnnConnection.RollbackTrans
    MsgBox "資料存檔失敗！" & vbCrLf & Err.Description
End Sub

'檢查 Staff_Assist 資料是否存在
Private Function IsStaffAssExist(strST01 As String, strDate As String) As Boolean
Dim Rs As ADODB.Recordset
   
   IsStaffAssExist = False
   strExc(0) = "select * from Staff_Assist where sa01='" & strST01 & "' and sa02=" & DBDATE(strDate)
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      IsStaffAssExist = True
   End If
   Rs.Close
   Set Rs = Nothing
End Function

'檢查 ABS014 資料是否存在
Private Function IsABS014Exist(strST01 As String, strDate As String, strKind As String) As Boolean
Dim Rs As ADODB.Recordset
   
   IsABS014Exist = False
   strExc(0) = "select * from ABS014 where b1401='" & strST01 & "' and b1402=" & DBDATE(strDate) & " and b1403='" & strKind & "'"
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      IsABS014Exist = True
   End If
   Rs.Close
   Set Rs = Nothing
End Function

Private Sub Form_Load()
   MoveFormToCenter Me
   m_EditMode = 0
   If m_WorkType = 0 Then
      Me.Caption = "異常確認"
      cmdok.Caption = "人事確認"
      Label13.Caption = "備註：有修改個人確認或原因時，會發E-Mail通知當事人。"
      cmdBack.Enabled = True
   Else
      Me.Caption = "新增正常時間異常打卡"
      cmdok.Caption = "新增"
      m_EditMode = 1 '1.新增 2.修改 3.刪除 4.查詢
      Label13.Caption = "備註：會發E-Mail通知當事人。"
      cmdBack.Enabled = False
   End If
   Call ClearData
   If m_EditMode <> 1 Then
      textB1401.Enabled = False
      textB1402.Enabled = False
      cboB1403.Enabled = False
      Call QueryData
'      textB1404.Enabled = False
   Else '新增
      textB1401.Enabled = True
      textB1402.Enabled = True
      cboB1403.Enabled = True
'      textB1404.Enabled = False
   End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160012_1 = Nothing
End Sub

'清除欄位值
Private Sub ClearData()
   textB1401.Text = "": Label12.Caption = ""
   textB1402.Text = ""
   cboB1403.ListIndex = 0
   textB1404.Text = ""
   textB1405.Text = "": m_B1405 = ""
'   cboSTime.ListIndex = 0
   textB1407.Text = "": m_B1407 = ""
   textB1408.Text = "": LblB1408.Caption = ""
   textB1409.Text = ""
   LblB1410.Caption = ""
'   textB1411.Text = ""
'   LblB1412.Caption = ""
'   LblB1413.Caption = ""
   textSA06 = ""
   Label23 = Empty
   grd2.Clear
   Frame1.Visible = False
   textB1411.Text = ""
   LblB1412.Caption = ""
   LblB1413.Caption = ""
End Sub

'取得打卡時間
Private Sub GetPr02(strDate As String, strST01 As String)
Dim Rs As ADODB.Recordset
   
   textB1404 = ""
   strExc(0) = "select scd01,pr01,nvl(min(pr02),'') as min_pr02,nvl(max(pr02),'') as max_pr02 from pollrecord,staffcarddata where pr03=scd02(+) and pr01=" & DBDATE(strDate) & " and scd01='" & strST01 & "' group by scd01,pr01"
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      Rs.MoveFirst
      If Left(Trim(cboB1403), 1) = "A" Then
         textB1404 = "" & Rs.Fields("min_pr02")
      Else
         textB1404 = "" & Rs.Fields("max_pr02")
         
      End If
   End If
   If textB1404 <> "" Then
      textB1404 = Format(Right("000000" & textB1404, 6), "##:##:##")
   End If
   Rs.Close
   Set Rs = Nothing
End Sub

Private Function CheckDataValid() As Boolean
Dim strTit As String
Dim strMsg As String
Dim nResponse
   
   CheckDataValid = False
   
   '新增時
   If m_EditMode = 1 Then
      If textB1401 = "" Then
         MsgBox "員工代號不可空白 !!!"
         textB1401.SetFocus
         Exit Function
      End If
      If textB1402 = "" Then
         MsgBox "日期不可空白 !!!"
         textB1402.SetFocus
         Exit Function
      End If
      '檢查資料是否已存在
      If IsABS014Exist(textB1401, DBDATE(textB1402), Left(Trim(cboB1403), 1)) = True Then
         strTit = "新增資料"
         strMsg = "該筆記錄已存在"
         nResponse = MsgBox(strMsg, vbOKOnly, strTit)
         textB1402.SetFocus
         Exit Function
      End If
   End If
   
   '非新增或是新增一併確認時
   If m_EditMode <> 1 Or textB1405 <> "" Then
      If textB1405 = "" Then
         MsgBox "個人確認不可空白 !!!"
         If textB1405.Enabled = True Then textB1405.SetFocus
         Exit Function
      End If
'      If Left(Trim(cboB1403), 1) = "A" Then
'         If cboSTime = "" Then
'            MsgBox "上班時段不可空白 !!!"
'            cboSTime.SetFocus
'            Exit Function
'         End If
'      End If
      If textB1405 = "4" And textB1407 = "" Then
         MsgBox "未打卡原因不可空白 !!!"
         If textB1407.Enabled = True Then textB1407.SetFocus
         Exit Function
      'Add By Sindy 2013/8/30
      ElseIf textB1405 = "8" And textB1407 = "" Then
         MsgBox "請輸入未打卡原因 !!!"
         If textB1407.Enabled = True Then textB1407.SetFocus
         Exit Function
      '2013/8/30 END
      End If
   End If
   
   If textB1405 = "1" Then '填請假,是否有假單存在
      If CheckIsPersonRestSector(textB1401, DBDATE(textB1402), "00:00", DBDATE(textB1402), "24:00", "") = False Then
         MsgBox "無假單 !!!"
         If textB1405.Enabled = True Then textB1405.SetFocus
         Exit Function
      End If
   End If
   If textB1405 = "4" Then
'      If textB1408 = "" Then
'         MsgBox "請輸入批示主管的員工代號 !!!"
'         textB1408.SetFocus
'         Exit Function
'      ElseIf textB1408 = textB1401 Then
'         MsgBox "批示主管輸入錯誤 !!!"
'         textB1408.SetFocus
'         Exit Function
'      End If
      If textB1409 = "N" Then '主管不同意時,必須輸入曠職(時)
         If textSA06 = "" Then
            'Modify By Sindy 2025/8/28
            'MsgBox "曠職(時)不可空白 !!!"
            MsgBox "曠職(分)不可空白 !!!"
            '2025/8/28 END
            textSA06.SetFocus
            Exit Function
         End If
      ElseIf textB1409 = "" Then
         MsgBox "主管尚未批示 !!!"
         Exit Function
      End If
   End If
   
   '提醒,上班打卡時間是否大於等於9:30
   If textB1404 <> "" Then
      If textB1405.Text <> "1" And _
         Left(cboB1403, 1) = "A" And Val(Replace(Left(textB1404, Len(textB1404) - 2), ":", "")) >= 930 Then
         If MsgBox("此人上班打卡時間已超過9:30是否應該請假？", vbYesNo + vbDefaultButton2) = vbYes Then
            Exit Function
         End If
      End If
   End If
   
   'Add By Sindy 2013/8/7
   If textB1411 <> "" Then
      If MsgBox("此筆異常，已在 " & ChangeWStringToTDateString(DBDATE(LblB1412.Caption)) & " " & LblB1413.Caption & " 確認了！" & vbCrLf & vbCrLf & _
                "確定還要重新確認嗎？", vbYesNo + vbDefaultButton2) = vbNo Then
         Exit Function
      End If
      If m_B1405 = textB1405 Then
         MsgBox "欲重新確認，請重新輸入『個人確認』欄 !!!"
         textB1405.SetFocus
         Exit Function
      End If
   End If
   
   CheckDataValid = True
End Function

Private Function TxtValidate() As Boolean
Dim Cancel As Boolean

   TxtValidate = False
   
   If Me.textB1401.Enabled = True Then
      Cancel = False
      textB1401_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textB1402.Enabled = True Then
      Cancel = False
      textB1402_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If m_EditMode = 1 Then
      cboB1403_Validate Cancel
   End If
   If Me.textB1405.Enabled = True Then
      Cancel = False
      textB1405_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
   If Me.textSA06.Enabled = True Then
      Cancel = False
      textSA06_Validate Cancel
      If Cancel = True Then
         Exit Function
      End If
   End If
'   If Me.textB1408.Enabled = True Then
'      Cancel = False
'      textB1408_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
'   If Me.cboSTime.Enabled = True Then
'      Cancel = False
'      cboSTime_Validate Cancel
'      If Cancel = True Then
'         Exit Function
'      End If
'   End If
   
   'Add by Sindy 2021/9/1 檢查畫面的 TextBox, ComboBox 是否含有Unicode文字
   If PUB_ChkUniText(Me) = False Then
      Exit Function
   End If
   '2021/9/1 END
   
   TxtValidate = True
End Function

Private Sub textB1401_GotFocus()
   InverseTextBox textB1401
End Sub

Private Sub textB1401_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub textB1401_Validate(Cancel As Boolean)
Dim Rs As New ADODB.Recordset
   
   If textB1401.Text = "" Then Label12 = ""
   
   If textB1401 <> "" Then
      ' 檢查員工編號規則
      If ChkStaffST04(textB1401) Then
         Call textB1401_GotFocus
         Cancel = True
         Exit Sub
      End If
      Label12 = GetStaffName(textB1401, True)
      If Label12 = "" Then
         MsgBox "員工編號錯誤！查無此員工！", vbInformation
         Call textB1401_GotFocus
         Cancel = True
         Exit Sub
      End If
      If textB1402 <> "" Then
         Call PollRecordQueryData
      End If
   End If
End Sub

Private Sub textB1402_GotFocus()
   InverseTextBox textB1402
End Sub

Private Sub textB1402_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textB1402_Validate(Cancel As Boolean)
   If textB1402 = "" Then Exit Sub
   If textB1402 <> "" Then
      If ChkDate(textB1402) = False Then
         Call textB1402_GotFocus
         Cancel = True
         Exit Sub
      End If
      If textB1401 <> "" Then
         Call GetPr02(textB1402, textB1401)
         Call PollRecordQueryData
      End If
   End If
End Sub

Private Sub cboB1403_Validate(Cancel As Boolean)
   If textB1401 <> "" And textB1402 <> "" Then
      Call GetPr02(textB1402, textB1401)
   End If
End Sub

'Private Sub textB1404_GotFocus()
'   InverseTextBox textB1404
'End Sub
'
'Private Sub textB1404_KeyPress(KeyAscii As Integer)
'   KeyAscii = Pub_NumAscii(KeyAscii)
'End Sub
'
'Private Sub textB1404_Validate(Cancel As Boolean)
'   If textB1404 = "" Then Exit Sub
'   If textB1404 <> "" Then
'      If Len(textB1404) > 6 Then
'         MsgBox "時間輸入錯誤 !!!"
'         Call textB1404_GotFocus
'         Cancel = True
'         Exit Sub
'      End If
'   End If
'End Sub

Private Sub textB1405_GotFocus()
   InverseTextBox textB1405
   CloseIme
End Sub

Private Sub textB1405_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub textB1405_Validate(Cancel As Boolean)
   If textB1405 = "" Then Exit Sub
   If textB1405 <> "" Then
      Select Case textB1405
         'Modify By Sindy 2013/8/30 +8
         Case 1, 2, 3, 4, 5, 6, 7, 8
'            If textB1405 = "1" And textB1407 = "" Then
'               textB1407 = "已請假"
'            ElseIf textB1405 = "2" And textB1407 = "" Then
'               textB1407 = "遲到 或 早退"
'            ElseIf textB1405 = "3" And textB1407 = "" Then
'               textB1407 = "忘打卡"
'            ElseIf textB1405 = "5" And textB1407 = "" Then
'               textB1407 = "指紋異常"
'            ElseIf textB1405 = "6" And textB1407 = "" Then
'               textB1407 = "因公未打卡"
'            Else
               If m_B1405 <> "4" And textB1405 = "4" Then
                  MsgBox "人事處不可輸入4 !!!"
                  Call textB1405_GotFocus
                  Cancel = True
                  Exit Sub
               End If
'            End If
         Case Else
            'Modify By Sindy 2013/8/30 +8
            MsgBox "個人確認只可輸入1~8 !!!"
            Call textB1405_GotFocus
            Cancel = True
            Exit Sub
      End Select
   End If
   CloseIme
End Sub

Private Sub textB1407_GotFocus()
   InverseTextBox textB1407
   OpenIme
End Sub

Private Sub textSA06_GotFocus()
   InverseTextBox textSA06
   CloseIme
End Sub

Private Sub textSA06_KeyPress(KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii, True)
End Sub

Private Sub textSA06_Validate(Cancel As Boolean)
   If textSA06 <> "" Then
       If CheckLengthIsOK(textSA06, textSA06.MaxLength) = False Then
           Call textSA06_GotFocus
           Cancel = True
           Exit Sub
       End If
       'Modify By Sindy 2025/8/28
'       If textSA06.Text >= 8 Then
'           Call textSA06_GotFocus
'           MsgBox "曠職(時)不可超過8小時!!!", vbExclamation + vbOKOnly
'           Cancel = True
'           Exit Sub
'       End If
       strExc(10) = PUB_intWkHour * 60
       If textSA06.Text >= strExc(10) Then
           Call textSA06_GotFocus
           MsgBox "曠職(分)不可超過 " & strExc(10) & " 分鐘!!!", vbExclamation + vbOKOnly
           Cancel = True
           Exit Sub
       End If
       '2025/8/28 END
   End If
   CloseIme
End Sub

'Private Sub textB1408_GotFocus()
'   InverseTextBox textB1408
'End Sub
'
'Private Sub textB1408_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'End Sub
'
'Private Sub textB1408_Validate(Cancel As Boolean)
'Dim rs As New ADODB.Recordset
'
'   If textB1408.Text = "" Then LblB1408 = ""
'
'   If textB1408 <> "" Then
'      ' 檢查員工編號規則
'      If ChkStaffST04(textB1408) Then
'         Call textB1408_GotFocus
'         Cancel = True
'         Exit Sub
'      End If
'      LblB1408 = GetStaffName(textB1408, True)
'      If Label12 = "" Then
'         MsgBox "主管的員工編號錯誤！查無此員工！", vbInformation
'         Call textB1408_GotFocus
'         Cancel = True
'         Exit Sub
'      End If
'   End If
'End Sub

'Private Sub cboSTime_GotFocus()
''   InverseTextBox cboSTime
'End Sub
'
'Private Sub cboSTime_KeyPress(KeyAscii As Integer)
'   KeyAscii = Pub_NumAscii(KeyAscii)
'End Sub
'
'Private Sub cboSTime_Validate(Cancel As Boolean)
'Dim bolChkOK As Boolean
'Dim i As Integer
'
'   If cboSTime <> "" Then
'      bolChkOK = False
'      For i = 0 To cboSTime.ListCount - 1
'        If cboSTime.Text = cboSTime.List(i) Then
'           bolChkOK = True
'           Exit For
'        End If
'      Next i
'      If bolChkOK = False Then
'         Call cboSTime_GotFocus
'         MsgBox "輸入錯誤 !!!", vbExclamation + vbOKOnly
'         Cancel = True
'         Exit Sub
'      End If
'   End If
'End Sub
