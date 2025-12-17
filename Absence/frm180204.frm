VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm180204 
   BorderStyle     =   1  '單線固定
   Caption         =   "打卡異常主管處理"
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
   Begin VB.CommandButton cmdOK 
      Caption         =   "退回當事者"
      Height          =   345
      Index           =   2
      Left            =   5235
      TabIndex        =   5
      Top             =   60
      Width           =   1305
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "同意"
      Height          =   345
      Index           =   0
      Left            =   4410
      TabIndex        =   4
      Top             =   60
      Width           =   795
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   345
      Left            =   7980
      TabIndex        =   7
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "畫面更新(&Q)"
      Default         =   -1  'True
      Height          =   345
      Left            =   6590
      TabIndex        =   6
      Top             =   60
      Width           =   1335
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "不同意"
      Height          =   345
      Index           =   1
      Left            =   3585
      TabIndex        =   3
      Top             =   60
      Width           =   765
   End
   Begin VB.CommandButton cmdDetail 
      Caption         =   "打卡明細"
      Height          =   345
      Left            =   690
      TabIndex        =   1
      Top             =   60
      Width           =   945
   End
   Begin VB.CommandButton cmdABS 
      Caption         =   "查詢當日請假資料"
      Height          =   345
      Left            =   1695
      TabIndex        =   2
      Top             =   60
      Width           =   1845
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3945
      Left            =   60
      TabIndex        =   0
      Top             =   900
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   6967
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|日期|部門|姓名|時段|打卡時間|有無請假|未打卡原因"
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
      _Band(0).Cols   =   8
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grd2 
      Height          =   1155
      Left            =   7020
      TabIndex        =   8
      Top             =   480
      Visible         =   0   'False
      Width           =   1830
      _ExtentX        =   3228
      _ExtentY        =   2028
      _Version        =   393216
      Cols            =   14
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      _NumberOfBands  =   1
      _Band(0).Cols   =   14
   End
   Begin MSForms.ComboBox cboB1408 
      Height          =   285
      Left            =   2160
      TabIndex        =   14
      Top             =   570
      Width           =   2010
      VariousPropertyBits=   679495707
      DisplayStyle    =   3
      Size            =   "3545;503"
      MatchEntry      =   1
      ShowDropButtonWhen=   2
      FontName        =   "新細明體-ExtB"
      FontHeight      =   180
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin VB.Label Label4 
      Caption         =   "改送其他簽核主管批示："
      Height          =   225
      Left            =   120
      TabIndex        =   13
      Top             =   570
      Width           =   1995
   End
   Begin VB.Label Label5 
      Caption         =   "　　　3.改送其他簽核主管批示，輸入主管按下改送即可；會發Email通知該主管。"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   60
      TabIndex        =   12
      Top             =   5400
      Width           =   7035
   End
   Begin VB.Label Label3 
      Caption         =   "　　　2.退回：欲請當事者另送其他主管進行簽核，且發退回E-Mail通知當事者。"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   60
      TabIndex        =   11
      Top             =   5160
      Width           =   7035
   End
   Begin VB.Label Label2 
      Caption         =   "備註：1.批示為不同意時， 會同時發E-Mail通知當事者。"
      ForeColor       =   &H000000C0&
      Height          =   195
      Left            =   60
      TabIndex        =   10
      Top             =   4920
      Width           =   7035
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      ForeColor       =   &H000000C0&
      Height          =   225
      Left            =   60
      TabIndex        =   9
      Top             =   420
      Width           =   2475
   End
End
Attribute VB_Name = "frm180204"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2023/12/22 修改抓新部門程式
'Memo By Sindy 2021/5/28 Form2.0已修改
'Create By Sindy 2013/6/21
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Dim m_NotReason As String
Public bolClose As Boolean


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
Dim bolSelV As Boolean
   
   bolSelV = False
   Me.Enabled = False
   For i = 1 To grd2.Rows - 1
      grd2.col = 0
      grd2.row = i
      If Trim(grd2.Text) = "V" Then
         bolSelV = True
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
   If bolSelV = False Then
      Call cmdABS_Click
   End If
End Sub

'Add By Sindy 2014/2/17
Private Sub cboB1408_Click()
   If Trim(cboB1408.Text) <> "" Then
      If Left(cboB1408.Text, 5) = strUserNum Then
         MsgBox "請輸入欲改送的簽核主管，不可以輸入自己！", , "沒有資料"
         cboB1408.SetFocus
         Exit Sub
      End If
      For i = 1 To GRD1.Rows - 1
         GRD1.col = 0
         GRD1.row = i
         If GRD1.TextMatrix(i, 0) = "V" Then
            If GRD1.TextMatrix(i, 13) = strUserNum Then
               MsgBox "請輸入欲改送的簽核主管，不可以輸入當事者（第" & i & "筆）！", , "沒有資料"
               cboB1408.SetFocus
               Exit Sub
            End If
         End If
      Next i
      cmdOK(2).Caption = "改送簽核主管"
   Else
      cmdOK(2).Caption = "退回當事者"
   End If
End Sub
Private Sub cboB1408_LostFocus()
   Call cboB1408_Click
End Sub
Private Sub cboB1408_Validate(Cancel As Boolean)
   If cboB1408.Text <> "" And Len(Trim(cboB1408.Text)) = 5 Then
      For i = 1 To cboB1408.ListCount - 1
         If Trim(cboB1408.Text) = Left(Trim(cboB1408.List(i)), 5) Then
            cboB1408.ListIndex = i
         End If
      Next i
      If Len(Trim(cboB1408.Text)) = 5 Then
         cboB1408.Text = Trim(cboB1408.Text) & " " & GetPrjSalesNM(Trim(cboB1408.Text))
      End If
   End If
End Sub

Private Sub cmdABS_Click()
Dim rsTmp As New ADODB.Recordset
   
   grd2.Clear
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
'         If QueryData_ABS(GRD1.TextMatrix(i, 13), GRD1.TextMatrix(i, 1)) = True Then
'            Exit Sub
'         End If
         If PUB_QueryData_ABS(GRD1.TextMatrix(i, 13), GRD1.TextMatrix(i, 1), rsTmp) = True Then
            Set grd2.Recordset = rsTmp
            Call PubShowNextData
            Exit Sub
         End If
      End If
   Next i
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
Private Sub QueryData(bolChange As Boolean)
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim subSQL1 As String, subSQL2 As String
Dim strStarDate As String, strEndDate As String, strUserId As String
Dim strCon As String
Dim strB0101 As String 'Add By Sindy 2014/2/17
   
   cboB1408.Clear 'Add By Sindy 2014/2/17
   GRD1.Clear
   SetGrd
   strCon = ""
'   If Trim(cmdQuery.Caption) = "不同意的資料" Then
'      strCon = " and b1409='N' and b1411 is null"
'   ElseIf Trim(cmdQuery.Caption) = "未批示的資料" Then
      strCon = " and b1409 is null and b1411 is null"
'   End If
   
   '請假日期區間
   strExc(0) = "select b1402 from abs014 where b1408='" & strUserNum & "'" & strCon & " order by b1402 asc"
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
   '欲查詢員工代號
   strExc(0) = "select distinct b1401 from abs014 where b1408='" & strUserNum & "'" & strCon
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
   strUserId = ""
   If intI = 1 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         strUserId = strUserId & "'" & Trim("" & rsTmp.Fields(0)) & "',"
         rsTmp.MoveNext
      Loop
      If strUserId <> "" Then
         strUserId = Left(strUserId, Len(strUserId) - 1)
      End If
   End If
   If strUserId = "" Then strUserId = "''"
   rsTmp.Close
   
   m_blnColOrderAsc = True
   Screen.MousePointer = vbHourglass
   
   '有無請假
   subSQL1 = "(SELECT SA01 as Userid,SA02 as date1,SA04 as date2,'Y' as IsY FROM staff_Absence,staff s1 WHERE sa01 in(" & strUserId & ") and SA01=s1.st01(+)" & _
             " and ((SA02>=" & DBDATE(strStarDate) & " and SA02<=" & DBDATE(strEndDate) & ") or (SA04>=" & DBDATE(strStarDate) & " and SA04<=" & DBDATE(strEndDate) & ") or (" & DBDATE(strStarDate) & " between SA02 and SA04) or (" & DBDATE(strEndDate) & " between SA02 and SA04))" & _
             " union SELECT SB01 as Userid,SB02 as date1,SB04 as date2,'Y' as IsY FROM staff_busi_trip,staff s1 WHERE sb01 in(" & strUserId & ") and SB01=s1.st01(+)" & _
             " and ((SB02>=" & DBDATE(strStarDate) & " and SB02<=" & DBDATE(strEndDate) & ") or (SB04>=" & DBDATE(strStarDate) & " and SB04<=" & DBDATE(strEndDate) & ") or (" & DBDATE(strStarDate) & " between SB02 and SB04) or (" & DBDATE(strEndDate) & " between SB02 and SB04))" & _
             " union SELECT B1003 as Userid,B1004 as date1,B1006 as date2,'Y' as IsY FROM ABS010,staff s1 WHERE B1003 in(" & strUserId & ") and B1002 in('01','03') and B1018 not in('" & 退回 & "','" & 註銷 & "','" & 已核准 & "') and B1003=s1.st01(+)" & _
             " and ((B1004>=" & DBDATE(strStarDate) & " and B1004<=" & DBDATE(strEndDate) & ") or (B1006>=" & DBDATE(strStarDate) & " and B1006<=" & DBDATE(strEndDate) & ") or (" & DBDATE(strStarDate) & " between B1004 and B1006) or (" & DBDATE(strEndDate) & " between B1004 and B1006))" & _
             ") V1"
   '每日的最小和最大打卡時間
   subSQL2 = "(select scd01,pr01,nvl(min(pr02),0) as min_pr02,nvl(max(pr02),0) as max_pr02 from pollrecord,staffcarddata where pr03=scd02(+) and scd01 in(" & strUserId & ") and pr01>=" & DBDATE(strStarDate) & " and pr01<=" & DBDATE(strEndDate) & " group by scd01,pr01) V2"
   '尚無處理結果的打卡異常資料
   strSql = "select '' as V,sqldatet(b1402) as 日期," & IIf(strSrvDate(1) >= 新部門啟用日, "A0922", "A0902") & " as 部門,s1.ST02 as 姓名,decode(b1403,'A','上班','P','下班',b1403) as 時段," & _
            "sqltime6(b1404) as 打卡時間,V1.IsY as 有無請假,ac03 as 個人確認," & _
            "b1407 as 未打卡原因,s2.st02 as 主管,decode(b1409,'Y','同意','N','不同意',b1409) as 主管批示,sqldatet(b1410) as 批示日期," & _
            IIf(strSrvDate(1) >= 新部門啟用日, "s1.ST93", "s1.ST03") & " as ST03,b1401,b1406,V2.min_pr02 as 第一筆打卡時間,V2.max_pr02 as 最後一筆打卡時間" & _
            " from abs014,staff s1,staff s2," & IIf(strSrvDate(1) >= 新部門啟用日, "ACC090NEW", "ACC090") & ",allcode," & subSQL1 & "," & subSQL2 & _
            " where b1408='" & strUserNum & "' and b1401=s1.st01(+)" & _
            " and b1408=s2.st01(+)" & _
            IIf(strSrvDate(1) >= 新部門啟用日, " and s1.ST93=A0921(+)", " and s1.ST03=A0901(+)") & _
            " and ac01(+)='10' and b1405=ac02(+)" & strCon & _
            " and B1401=V1.Userid(+) and B1402 between V1.date1(+) and V1.date2(+)" & _
            " and B1401=V2.scd01(+) and B1402=V2.pr01(+)" & _
            " order by b1402 desc,b1401 asc,b1403 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
'      Label1.Caption = "下列為" & Trim(cmdQuery.Caption) & ":"
   Else
      ShowNoData
'      Label1.Caption = "目前沒有" & Trim(cmdQuery.Caption) & "!"
'      MsgBox "目前沒有" & Trim(cmdQuery.Caption) & "!!", , "沒有資料"
      Screen.MousePointer = vbDefault
      rsTmp.Close
      Set rsTmp = Nothing
'      Call GetcmdQueryCaption
'      If Trim(cmdQuery.Caption) = "未批示的資料" Then
'         Call QueryData(True)
'      End If
      Exit Sub
   End If
   
   '若有資料游標停在第一筆
   GRD1.Visible = False
   GRD1.col = 0
   GRD1.row = 1
'   If rsTmp.RecordCount > 0 Then
'      GRD1.TextMatrix(1, 0) = "V"
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   GRD1.Visible = True
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
   'Add By Sindy 2014/2/17 改送其他簽核主管批示,人員下拉選單
   '待簽核人員
   strB0101 = ""
   For i = 1 To GRD1.Rows - 1
      strB0101 = ",'" & GRD1.TextMatrix(i, 13) & "'"
   Next i
   If strB0101 <> "" Then strB0101 = Mid(strB0101, 2)
   '簽核主管
   strExc(0) = "select b0108 from abs001 where b0101 in(" & strB0101 & ") and b0108 is not null" & _
               " union select b0109 from abs001 where b0101 in(" & strB0101 & ") and b0109 is not null" & _
               " union select b0110 from abs001 where b0101 in(" & strB0101 & ") and b0110 is not null" & _
               " union select b0111 from abs001 where b0101 in(" & strB0101 & ") and b0111 is not null" & _
               " order by 1 asc"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsTmp.MoveFirst
      cboB1408.AddItem ""
      Do While Not rsTmp.EOF
         If rsTmp.Fields(0) <> strUserNum Then
            cboB1408.AddItem rsTmp.Fields(0) & " " & GetPrjSalesNM(rsTmp.Fields(0))
         End If
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   '2014/2/17 END
   
'   If bolChange = True Then
'      Call GetcmdQueryCaption
'   End If
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'Private Sub GetcmdQueryCaption()
'   If Trim(cmdQuery.Caption) = "不同意的資料" Then
'      cmdQuery.Caption = "未批示的資料"
'   ElseIf Trim(cmdQuery.Caption) = "未批示的資料" Then
'      cmdQuery.Caption = "不同意的資料"
'   End If
'End Sub

'同意 或 不同意
Public Sub cmdok_Click(Index As Integer)
Dim bolSelisV As Boolean
   
   m_NotReason = ""
   bolSelisV = False
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If GRD1.TextMatrix(i, 0) = "V" Then
         If bolSelisV = False Then
            If Index = 1 Then
               If MsgBox("確定執行不同意嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                  Exit Sub
               End If
               m_NotReason = InputBox("請輸入不同意的理由：")
               If m_NotReason = "" Then Exit Sub
            ElseIf Index = 2 Then
               If cmdOK(2).Caption = "退回當事者" Then
                  If MsgBox("確定退回當事者嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                     Exit Sub
                  End If
               Else
                  If cboB1408.Text = "" Then
                     MsgBox "請輸入其他簽核主管！"
                     cboB1408.SetFocus
                     Exit Sub
                  Else
                     If MsgBox("確定改送" & Mid(cboB1408.Text, 7) & "嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
                        Exit Sub
                     End If
                  End If
               End If
            End If
         End If
         
         bolSelisV = True
         GRD1.Text = ""
         For j = 0 To GRD1.Cols - 1
            GRD1.col = j
            GRD1.CellBackColor = QBColor(15)
         Next j
         '更新資料
         If UpdateData(GRD1.TextMatrix(i, 13), GRD1.TextMatrix(i, 1), _
              IIf(Trim(GRD1.TextMatrix(i, 4)) = "上班", "A", "P"), _
              IIf(Index = 2, "", IIf(Index = 0, "Y", "N")), GRD1.TextMatrix(i, 8)) = False Then
            Exit For
         End If
      End If
   Next i
   If bolSelisV = True Then
'      Call GetcmdQueryCaption
      Call QueryData(True)
   ElseIf GRD1.TextMatrix(1, 1) <> "" Then
      MsgBox "未勾選資料！"
   End If
End Sub

'更新資料
Private Function UpdateData(strB1401 As String, strB1402 As String, strB1403 As String, strB1409 As String, strB1407 As String) As Boolean
Dim strSubject As String, strContent As String
Dim strST59 As String, strTo As String
   
On Error GoTo ErrHand
   
   UpdateData = True
   cnnConnection.BeginTrans
   If strB1409 = "" And Trim(cboB1408.Text) = "" Then '退回當事者
      strSql = "UPDATE ABS014 SET B1408=null" & _
               " WHERE B1401='" & strB1401 & "' and B1402=" & DBDATE(strB1402) & " and B1403='" & strB1403 & "'"
      
   ElseIf strB1409 = "" And Trim(cboB1408.Text) <> "" Then '改送其他簽核主管
      strSql = "UPDATE ABS014 SET B1408='" & Left(Trim(cboB1408.Text), 5) & "'" & _
               " WHERE B1401='" & strB1401 & "' and B1402=" & DBDATE(strB1402) & " and B1403='" & strB1403 & "'"
   '同意/不同意
   Else
      strSql = "UPDATE ABS014 SET B1409='" & strB1409 & "'" & _
                                ",B1410=" & strSrvDate(1)
      If strB1409 = "N" Then
         strSql = strSql & ",B1407=B1407||';不同意的理由:" & m_NotReason & "'"
      End If
      strSql = strSql & " WHERE B1401='" & strB1401 & "' and B1402=" & DBDATE(strB1402) & " and B1403='" & strB1403 & "'"
   End If
   Pub_SeekTbLog strSql
   cnnConnection.Execute strSql
   cnnConnection.CommitTrans
   
   If strB1409 = "N" Or strB1409 = "" Then
      If strB1409 = "N" Then '不同意時發Mail通知當事者
         strSubject = GetPrjSalesNM(strB1401) & " " & ChangeWStringToTDateString(DBDATE(strB1402)) & " " & IIf(Left(Trim(strB1403), 1) = "A", "上", "下") & "班未打卡，主管不同意！"
         strContent = strSubject & vbCrLf & vbCrLf
         strContent = strContent & "未打卡原因：" & strB1407 & vbCrLf & vbCrLf
         strContent = strContent & "請　注　意：您的主管不同意您的未打卡批示。" & vbCrLf & vbCrLf
         strContent = strContent & "不同意理由：" & m_NotReason & vbCrLf
         
      ElseIf strB1409 = "" Then
         If Trim(cboB1408.Text) = "" Then '退回時發Mail通知當事者
            strSubject = GetPrjSalesNM(strB1401) & " " & ChangeWStringToTDateString(DBDATE(strB1402)) & " " & IIf(Left(Trim(strB1403), 1) = "A", "上", "下") & "班未打卡，主管退回！"
            strContent = strSubject & vbCrLf & vbCrLf
            strContent = strContent & "未打卡原因：" & strB1407 & vbCrLf & vbCrLf
            strContent = strContent & "請　注　意：" & GetPrjSalesNM(strUserNum) & "將您的未打卡批示退回，請另送其他主管進行簽核。" & vbCrLf
         
         Else '改送其他簽核主管時發Mail通知主管批示
            strSubject = GetPrjSalesNM(strB1401) & " " & ChangeWStringToTDateString(DBDATE(strB1402)) & " " & IIf(Left(Trim(strB1403), 1) = "A", "上", "下") & "班未打卡，請主管簽核！"
            strContent = strSubject & vbCrLf & vbCrLf
            strContent = strContent & "未打卡原因：" & strB1407 & vbCrLf
            strContent = strContent & "處　　　理：請至案件管理系統（一般作業->出缺勤作業->簽核->打卡異常主管處理）中，進行處理。" & vbCrLf
            PUB_SendMail strUserNum, Left(Trim(cboB1408.Text), 5), "", strSubject, strContent, , , , , , , , , , True
            Exit Function
         End If
      End If
      
      strST59 = PUB_GetST59(strB1401)
      If Not IsNull(strST59) And strST59 <> "" Then
         strTo = strST59
      Else
         strTo = strB1401
      End If
      
      PUB_SendMail strUserNum, strTo, "", strSubject, strContent, , , , , , , , , , True
   End If
   Exit Function
   
ErrHand:
    UpdateData = False
    cnnConnection.RollbackTrans
    MsgBox GetPrjSalesNM(strB1401) & " " & ChangeWStringToTDateString(DBDATE(strB1402)) & " " & IIf(Left(Trim(strB1403), 1) = "A", "上", "下") & "班未打卡，資料存檔失敗！" & vbCrLf & Err.Description
End Function

'畫面更新
Private Sub cmdQuery_Click()
   Call QueryData(True)
End Sub

Private Sub Form_Load()
Dim strData As String
Dim strTemp As Variant
   
   MoveFormToCenter Me
   Label1.Caption = ""
   
   'Add By Sindy 2014/2/18
   '檢查是否有A.表單待簽核
   strData = ChkIsAbsenceMustPro
   strTemp = Split(strData, ",")
   For i = 0 To UBound(strTemp)
      If strTemp(i) = "A" Then
         MsgBox "您尚有表單待簽核的資料！" & vbCrLf & vbCrLf & _
                "請進入「簽核作業」，進行處理。"
      End If
   Next i
   '2014/2/18 END
   
   Call cmdQuery_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm180204 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("V", "日期", "部門", "姓名", "時段", _
                           "打卡時間", "有無請假", "個人確認", "未打卡原因", "主管", _
                           "主管批示", "批示日期", "ST03", _
                           "b1401", "b1406", "第一筆打卡時間", "最後一筆打卡時間")
   arrGridHeadWidth = Array(200, 800, 900, 700, 600, _
                            800, 800, 0, 3700, 0, _
                            0, 0, 0, _
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
'   MsgBox GRD1.MouseRow & " 雙擊"
End Sub

'Add By Sindy 2018/3/22
Private Sub GRD1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   GRD1.ToolTipText = ""
   If GRD1.MouseRow <> 0 And GRD1.MouseCol > 0 Then
      GRD1.ToolTipText = GRD1.TextMatrix(GRD1.MouseRow, GRD1.MouseCol)
   End If
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
