VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180105 
   BorderStyle     =   1  '單線固定
   Caption         =   "打卡異常個人處理"
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   4935
      Left            =   60
      TabIndex        =   6
      Top             =   780
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   8714
      _Version        =   393216
      Cols            =   12
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|日期|部門|姓名|時段|打卡時間|有無請假|個人確認|未打卡原因|主管|主管批示|批示日期"
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
      _Band(0).Cols   =   12
   End
   Begin VB.CommandButton cmdABS 
      Caption         =   "查詢當日請假資料"
      Height          =   345
      Left            =   3037
      TabIndex        =   3
      Top             =   60
      Width           =   2145
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "打卡明細"
      Height          =   345
      Left            =   1980
      TabIndex        =   2
      Top             =   60
      Width           =   975
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "畫面更新(&Q)"
      Default         =   -1  'True
      Height          =   345
      Left            =   6711
      TabIndex        =   0
      Top             =   60
      Width           =   1185
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   345
      Left            =   7980
      TabIndex        =   1
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox textB1401 
      Appearance      =   0  '平面
      BackColor       =   &H8000000F&
      BorderStyle     =   0  '沒有框線
      Height          =   285
      Left            =   1110
      Locked          =   -1  'True
      MaxLength       =   6
      TabIndex        =   5
      Top             =   540
      Width           =   585
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確認處理方式"
      Height          =   345
      Left            =   5264
      TabIndex        =   4
      Top             =   60
      Width           =   1365
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   705
      Left            =   4230
      TabIndex        =   8
      Top             =   120
      Visible         =   0   'False
      Width           =   840
      _ExtentX        =   1482
      _ExtentY        =   1252
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin MSForms.TextBox txtST02 
      Height          =   285
      Left            =   1740
      TabIndex        =   10
      Top             =   540
      Width           =   1605
      VariousPropertyBits=   679495711
      BackColor       =   -2147483633
      ScrollBars      =   3
      Size            =   "2831;503"
      SpecialEffect   =   0
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "(註：可雙擊點選，即可進入確認處理方式)"
      ForeColor       =   &H000000C0&
      Height          =   165
      Left            =   5190
      TabIndex        =   9
      Top             =   570
      Width           =   3585
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   285
      Left            =   180
      TabIndex        =   7
      Top             =   540
      Width           =   900
   End
End
Attribute VB_Name = "frm180105"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/5/28 Form2.0已修改
'Create By Sindy 2013/6/21
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Public bolClose As Boolean


Private Sub cmdABS_Click()
Dim rsTmp As New ADODB.Recordset
   
   GRD2.Clear
   SetGrd2
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If GRD1.TextMatrix(i, 0) = "V" Then
         GRD1.Text = ""
         For j = 0 To GRD1.Cols - 1
            GRD1.col = j
            GRD1.CellBackColor = QBColor(15)
         Next j
         If PUB_QueryData_ABS(GRD1.TextMatrix(i, 13), GRD1.TextMatrix(i, 1), rsTmp) = True Then
            Set GRD2.Recordset = rsTmp
            Call PubShowNextData
            Exit Sub
         End If
      End If
   Next i
End Sub

Private Sub SetGrd2()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer

   arrGridHeadText = Array("V", "員工代號", "表單編號", "TableID", "SA02", "SA03")
   arrGridHeadWidth = Array(800, 800, 800, 800, 800, 800)
   'grd2.Visible = False
   GRD2.Cols = UBound(arrGridHeadText) + 1
   GRD2.Rows = 2
   For iRow = 0 To GRD2.Cols - 1
      GRD2.row = 0
      GRD2.col = iRow
      GRD2.Text = arrGridHeadText(iRow)
      GRD2.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD2.CellAlignment = flexAlignCenterCenter
   Next
   'grd2.Visible = True
End Sub

'查詢出缺勤明細資料
Public Sub PubShowNextData()
Dim i As Integer
Dim bolSelV As Boolean
   
   bolSelV = False
   Me.Enabled = False
   For i = 1 To GRD2.Rows - 1
      GRD2.col = 0
      GRD2.row = i
      If Trim(GRD2.Text) = "V" Then
         bolSelV = True
         GRD2.Text = ""
         GRD2.col = 2 '表單編號
         Screen.MousePointer = vbHourglass
         Me.Hide
         Call frm180301_03.SetParent(Me)
         If GRD2.TextMatrix(i, 3) = "1" Then '出缺勤
            frm180301_03.txtB1001 = Pub_RplStr(GRD2.Text)
            frm180301_03.QueryData
         Else
            frm180301_03.txtB1003 = Pub_RplStr(GRD2.TextMatrix(i, 1))
            frm180301_03.m_SA02 = Pub_RplStr(GRD2.TextMatrix(i, 4))
            frm180301_03.m_SA03 = Pub_RplStr(GRD2.TextMatrix(i, 5))
            If GRD2.TextMatrix(i, 3) = "2" Then '請假
               frm180301_03.QueryData_2
            ElseIf GRD2.TextMatrix(i, 3) = "3" Then '加班
               frm180301_03.QueryData_3
            ElseIf GRD2.TextMatrix(i, 3) = "4" Then '出差
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
   If bolSelV = False Then
      Call cmdABS_Click
   End If
End Sub

'打卡明細
Private Sub cmdDetail_Click()
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If GRD1.TextMatrix(i, 0) = "V" Then
         GRD1.Text = ""
         For j = 0 To GRD1.Cols - 1
            GRD1.col = j
            GRD1.CellBackColor = QBColor(15)
         Next j
         bolClose = False
         Call frm180303_1.SetParent(Me)
         frm180303_1.m_B1401 = GRD1.TextMatrix(i, 13)
         frm180303_1.m_B1402 = GRD1.TextMatrix(i, 1)
         If frm180303_1.QueryData = True Then
            frm180303_1.Show vbModal '強制回應表單
         Else
            Unload frm180303_1
         End If
         If bolClose = True Then
            Exit Sub
         End If
      End If
   Next i
End Sub

Private Sub cmdExit_Click()
   Unload Me
End Sub

'查詢資料
Private Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim subSQL1 As String, subSQL2 As String
Dim strStarDate As String, strEndDate As String
   
   GRD1.Clear
   SetGrd
   
   strExc(0) = "select b1402 from abs014 where b1401='" & strUserNum & "' and b1411 is null order by b1402 asc"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
   strStarDate = strSrvDate(1)
   strEndDate = strSrvDate(1)
   If intI = 1 Then
      rsTmp.MoveFirst
      strStarDate = "" & rsTmp.Fields(0)
      rsTmp.MoveLast
      strEndDate = "" & rsTmp.Fields(0)
   End If
   rsTmp.Close
   
   m_blnColOrderAsc = True
   Screen.MousePointer = vbHourglass
   
   '有無請假
   subSQL1 = "(SELECT SA01 as Userid,SA02 as date1,SA04 as date2,'Y' as IsY FROM staff_Absence,staff s1 WHERE sa01='" & strUserNum & "' and SA01=s1.st01(+)" & _
             " and ((SA02>=" & DBDATE(strStarDate) & " and SA02<=" & DBDATE(strEndDate) & ") or (SA04>=" & DBDATE(strStarDate) & " and SA04<=" & DBDATE(strEndDate) & ") or (" & DBDATE(strStarDate) & " between SA02 and SA04) or (" & DBDATE(strEndDate) & " between SA02 and SA04))" & _
             " union SELECT SB01 as Userid,SB02 as date1,SB04 as date2,'Y' as IsY FROM staff_busi_trip,staff s1 WHERE sb01='" & strUserNum & "' and SB01=s1.st01(+)" & _
             " and ((SB02>=" & DBDATE(strStarDate) & " and SB02<=" & DBDATE(strEndDate) & ") or (SB04>=" & DBDATE(strStarDate) & " and SB04<=" & DBDATE(strEndDate) & ") or (" & DBDATE(strStarDate) & " between SB02 and SB04) or (" & DBDATE(strEndDate) & " between SB02 and SB04))" & _
             " union SELECT B1003 as Userid,B1004 as date1,B1006 as date2,'Y' as IsY FROM ABS010,staff s1 WHERE B1003='" & strUserNum & "' and B1002 in('01','03') and B1018 not in('" & 退回 & "','" & 註銷 & "','" & 已核准 & "') and B1003=s1.st01(+)" & _
             " and ((B1004>=" & DBDATE(strStarDate) & " and B1004<=" & DBDATE(strEndDate) & ") or (B1006>=" & DBDATE(strStarDate) & " and B1006<=" & DBDATE(strEndDate) & ") or (" & DBDATE(strStarDate) & " between B1004 and B1006) or (" & DBDATE(strEndDate) & " between B1004 and B1006))" & _
             ") V1"
   '每日的最小和最大打卡時間
   subSQL2 = "(select scd01,pr01,nvl(min(pr02),0) as min_pr02,nvl(max(pr02),0) as max_pr02 from pollrecord,staffcarddata where pr03=scd02(+) and scd01='" & strUserNum & "' and pr01>=" & DBDATE(strStarDate) & " and pr01<=" & DBDATE(strEndDate) & " group by scd01,pr01) V2"
   '尚無處理結果的打卡異常資料
   'Modify By Sindy 2023/12/20
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "select '' as V,sqldatet(b1402) as 日期,A0922 as 部門,s1.ST02 as 姓名,decode(b1403,'A','上班','P','下班',b1403) as 時段," & _
               "sqltime6(b1404) as 打卡時間,V1.IsY as 有無請假,ac03 as 個人確認," & _
               "b1407 as 未打卡原因,s2.st02 as 主管,decode(b1405||b1408,'4','退回',decode(b1409,'Y','同意','N','不同意',b1409)) as 主管批示,sqldatet(b1410) as 批示日期," & _
               "s1.ST93 as ST03,b1401,b1406,V2.min_pr02 as 第一筆打卡時間,V2.max_pr02 as 最後一筆打卡時間" & _
               " from abs014,staff s1,staff s2,ACC090NEW,allcode," & subSQL1 & "," & subSQL2 & _
               " where b1401='" & strUserNum & "' and b1401=s1.st01(+)" & _
               " and b1408=s2.st01(+)" & _
               " and s1.ST93=A0921(+)" & _
               " and ac01(+)='10' and b1405=ac02(+) and B1411 is null" & _
               " and B1401=V1.Userid(+) and B1402 between V1.date1(+) and V1.date2(+)" & _
               " and B1401=V2.scd01(+) and B1402=V2.pr01(+)" & _
               " order by b1402 desc,b1401 asc,b1403 asc"
   Else
   '2023/12/20 END
      strSql = "select '' as V,sqldatet(b1402) as 日期,A0902 as 部門,s1.ST02 as 姓名,decode(b1403,'A','上班','P','下班',b1403) as 時段," & _
               "sqltime6(b1404) as 打卡時間,V1.IsY as 有無請假,ac03 as 個人確認," & _
               "b1407 as 未打卡原因,s2.st02 as 主管,decode(b1405||b1408,'4','退回',decode(b1409,'Y','同意','N','不同意',b1409)) as 主管批示,sqldatet(b1410) as 批示日期," & _
               "s1.ST03 as ST03,b1401,b1406,V2.min_pr02 as 第一筆打卡時間,V2.max_pr02 as 最後一筆打卡時間" & _
               " from abs014,staff s1,staff s2,ACC090,allcode," & subSQL1 & "," & subSQL2 & _
               " where b1401='" & strUserNum & "' and b1401=s1.st01(+)" & _
               " and b1408=s2.st01(+)" & _
               " and s1.ST03=A0901(+)" & _
               " and ac01(+)='10' and b1405=ac02(+) and B1411 is null" & _
               " and B1401=V1.Userid(+) and B1402 between V1.date1(+) and V1.date2(+)" & _
               " and B1401=V2.scd01(+) and B1402=V2.pr01(+)" & _
               " order by b1402 desc,b1401 asc,b1403 asc"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
   Else
      ShowNoData
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
      Exit Sub
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
   If rsTmp.RecordCount > 0 Then
      GRD1.TextMatrix(1, 0) = "V"
      For i = 0 To GRD1.Cols - 1
         GRD1.col = i
         GRD1.CellBackColor = &HFFC0C0
      Next i
   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'確認處理方式
Public Sub cmdok_Click()
Dim strKind As String
Dim bolSelisV As Boolean
   
   bolSelisV = False
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If GRD1.TextMatrix(i, 0) = "V" Then
         bolSelisV = True
         GRD1.Text = ""
         For j = 0 To GRD1.Cols - 1
            GRD1.col = j
            GRD1.CellBackColor = QBColor(15)
         Next j
         If GRD1.TextMatrix(i, 4) = "上班" Then
            strKind = "A"
         Else
            strKind = "P"
         End If
         frm180105_1.m_B1401 = Trim(GRD1.TextMatrix(i, 13))
         frm180105_1.m_B1402 = DBDATE(GRD1.TextMatrix(i, 1))
         frm180105_1.m_B1403 = strKind
         frm180105_1.Show
         Me.Hide
         Exit For
      End If
   Next i
   If bolSelisV = False Then
      Call cmdQuery_Click
   End If
End Sub

'畫面更新
Private Sub cmdQuery_Click()
   Call QueryData
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   textB1401 = strUserNum
   txtST02.Text = strUserName
   Call cmdQuery_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180105 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("V", "日期", "部門", "姓名", "時段", _
                           "打卡時間", "有無請假", "個人確認", "未打卡原因", "主管", _
                           "主管批示", "批示日期", "ST03", _
                           "b1401", "b1406", "第一筆打卡時間", "最後一筆打卡時間")
   arrGridHeadWidth = Array(200, 800, 800, 700, 500, _
                            800, 700, 800, 1000, 700, _
                            700, 800, 0, _
                            0, 0, 0, 0)
   GRD1.Visible = False
   GRD1.Cols = UBound(arrGridHeadText) + 1
   GRD1.Rows = 2
   For iRow = 0 To GRD1.Cols - 1
      GRD1.row = 0
      GRD1.col = iRow
      GRD1.Text = arrGridHeadText(iRow)
      GRD1.ColWidth(iRow) = arrGridHeadWidth(iRow)
      GRD1.CellAlignment = flexAlignCenterCenter
   Next
   GRD1.Visible = True
End Sub

Private Sub GRD1_DblClick()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   '先清空全部已選取的資料列
   For j = 1 To GRD1.Rows - 1
      If GRD1.TextMatrix(j, 1) <> "" Then
         If GRD1.TextMatrix(j, 0) = "V" Then
            GRD1.col = 0
            GRD1.row = j
            GRD1.Text = ""
            For i = 0 To GRD1.Cols - 1
               GRD1.col = i
               GRD1.CellBackColor = QBColor(15)
            Next i
         End If
      End If
   Next j
   '該筆資料列變成已選取
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   If GRD1.TextMatrix(GRD1.MouseRow, 1) <> "" Then
'      If GRD1.Text = "V" Then
'         GRD1.Text = ""
'         For i = 0 To GRD1.Cols - 1
'            GRD1.col = i
'            GRD1.CellBackColor = QBColor(15)
'         Next i
'      Else
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
         Call cmdok_Click
'      End If
   End If
End If
GRD1.Visible = True
End Sub

Private Sub grd1_SelChange()
GRD1.Visible = False
If GRD1.MouseRow <> 0 Then
   GRD1.col = 0
   GRD1.row = GRD1.MouseRow
   If GRD1.TextMatrix(GRD1.MouseRow, 1) <> "" Then
      If GRD1.Text = "V" Then
         GRD1.Text = ""
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = QBColor(15)
         Next i
      Else
         GRD1.Text = "V"
         For i = 0 To GRD1.Cols - 1
            GRD1.col = i
            GRD1.CellBackColor = &HFFC0C0
         Next i
      End If
   End If
End If
GRD1.Visible = True
End Sub

Private Sub grd1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim nCol As Long, nRow As Long
   getGrdColRow GRD1, x, y, nCol, nRow
   If nCol < 0 Or nRow < 0 Then Exit Sub
   If nCol = 2 Then nCol = 12 '部門別置換為使用部門別代碼做排序
   GRD1.col = nCol
   GRD1.row = nRow
   If Me.GRD1.row < 1 And Me.GRD1.Text <> "V" Then
      If Me.GRD1.Text = "部門別" Then
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 3  '數值昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 4 '數值降冪
            m_blnColOrderAsc = True
         End If
      Else
         If m_blnColOrderAsc = True Then
            Me.GRD1.Sort = 5 '字串昇冪
            m_blnColOrderAsc = False
         Else
            Me.GRD1.Sort = 6 '字串降冪
            m_blnColOrderAsc = True
         End If
      End If
   End If
End Sub
