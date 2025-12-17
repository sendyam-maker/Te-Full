VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160012 
   BorderStyle     =   1  '單線固定
   Caption         =   "打卡異常處理作業"
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
   Begin VB.CommandButton cmdQuery3 
      Caption         =   "查詢系統確認且為請假處理的資料"
      Height          =   345
      Left            =   4290
      TabIndex        =   11
      Top             =   870
      Width           =   3045
   End
   Begin VB.CommandButton Command1 
      Caption         =   "立即接收打卡資料"
      Height          =   345
      Left            =   7110
      TabIndex        =   23
      Top             =   5400
      Visible         =   0   'False
      Width           =   1785
   End
   Begin VB.CommandButton cmdAddCard 
      Caption         =   "補輸打卡資料"
      Height          =   345
      Left            =   5850
      TabIndex        =   13
      Top             =   480
      Width           =   1485
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "取消全選"
      Height          =   345
      Left            =   7410
      TabIndex        =   15
      Top             =   870
      Width           =   1485
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1170
      MaxLength       =   3
      TabIndex        =   3
      Top             =   660
      Width           =   495
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   1800
      MaxLength       =   3
      TabIndex        =   4
      Top             =   660
      Width           =   495
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "結束"
      Height          =   345
      Left            =   7980
      TabIndex        =   10
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   1170
      MaxLength       =   6
      TabIndex        =   5
      Top             =   960
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   2220
      MaxLength       =   6
      TabIndex        =   6
      Top             =   960
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1170
      MaxLength       =   7
      TabIndex        =   0
      Top             =   30
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2220
      MaxLength       =   7
      TabIndex        =   1
      Top             =   30
      Width           =   915
   End
   Begin VB.CommandButton cmdQuery 
      Caption         =   "查詢"
      Default         =   -1  'True
      Height          =   345
      Left            =   4080
      TabIndex        =   7
      Top             =   90
      Width           =   915
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "frm160012.frx":0000
      Left            =   1170
      List            =   "frm160012.frx":000D
      Style           =   2  '單純下拉式
      TabIndex        =   2
      Top             =   330
      Width           =   1155
   End
   Begin VB.CommandButton cmdEmail 
      Caption         =   "E-Mail異常通知"
      Height          =   345
      Left            =   7410
      TabIndex        =   14
      Top             =   480
      Width           =   1485
   End
   Begin VB.CommandButton cmdQuery2 
      Caption         =   "查詢個人未確認資料"
      Height          =   345
      Left            =   5070
      TabIndex        =   8
      Top             =   90
      Width           =   1845
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "新增異常資料"
      Height          =   345
      Left            =   4290
      TabIndex        =   12
      Top             =   480
      Width           =   1485
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "確認作業"
      Height          =   345
      Left            =   6990
      TabIndex        =   9
      Top             =   90
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   4125
      Left            =   60
      TabIndex        =   16
      Top             =   1260
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   7267
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
   Begin VB.Label Label4 
      Caption         =   "(註：可雙擊點選，即可進入異常確認作業)"
      ForeColor       =   &H000000C0&
      Height          =   165
      Left            =   2880
      TabIndex        =   22
      Top             =   5460
      Width           =   3585
   End
   Begin VB.Label Label1 
      Caption         =   "共 0 筆"
      ForeColor       =   &H000000C0&
      Height          =   210
      Left            =   150
      TabIndex        =   21
      Top             =   5460
      Width           =   1395
   End
   Begin VB.Line Line3 
      X1              =   1590
      X2              =   2070
      Y1              =   780
      Y2              =   780
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "部  門  別："
      Height          =   180
      Left            =   210
      TabIndex        =   20
      Top             =   690
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2040
      X2              =   2250
      Y1              =   150
      Y2              =   150
   End
   Begin VB.Line Line1 
      X1              =   2040
      X2              =   2250
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Left            =   210
      TabIndex        =   19
      Top             =   990
      Width           =   900
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "日　　期："
      Height          =   180
      Left            =   210
      TabIndex        =   18
      Top             =   60
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "時　　段："
      Height          =   180
      Left            =   210
      TabIndex        =   17
      Top             =   390
      Width           =   900
   End
End
Attribute VB_Name = "frm160012"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2021/6/16 Form2.0已修改
'Create By Sindy 2013/6/21
Option Explicit

Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Public m_QueryType As Integer '記錄目前是那一種查詢狀況 1.查詢 2.未確認資料查詢 3.查詢系統確認且為請假處理的資料


'新增打卡資料
Private Sub cmdAddCard_Click()
   frm160012_2.Show
   Me.Hide
End Sub

'E-Mail異常通知
Private Sub cmdEmail_Click()
Dim strKind As String
Dim min_pr02 As String, max_pr02 As String
   
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If GRD1.TextMatrix(i, 0) = "V" Then
         GRD1.Text = ""
         For j = 0 To GRD1.Cols - 1
            GRD1.col = j
            GRD1.CellBackColor = QBColor(15)
         Next j
         min_pr02 = Format(GRD1.TextMatrix(i, 15), "000000")
         max_pr02 = Format(GRD1.TextMatrix(i, 16), "000000")
         If GRD1.TextMatrix(i, 4) = "上班" Then
            strKind = "A"
         Else
            strKind = "P"
            '檢查上午是否有異常資料(是否已有補入上班時段)
            strSql = "select * from abs014 where b1401='" & GRD1.TextMatrix(i, 13) & "'" & _
                     " and b1402=" & DBDATE(GRD1.TextMatrix(i, 1)) & _
                     " and b1403='A'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If RsTemp.RecordCount > 0 Then
                  GRD1.TextMatrix(i, 15) = "" & RsTemp.Fields("b1404")
                  min_pr02 = "" & RsTemp.Fields("b1404")
                  If Val(Format("" & RsTemp.Fields("b1406"), "0000")) > 0 Then
                     GRD1.TextMatrix(i, 15) = Val(Format("" & RsTemp.Fields("b1406"), "0000") & "00")
                  End If
               End If
            End If
         End If
         Call StaffCardErrSendMail(GRD1.TextMatrix(i, 13), GRD1.TextMatrix(i, 1), strKind, GRD1.TextMatrix(i, 5), Val(Left(Format(GRD1.TextMatrix(i, 15), "000000"), 4)), Val(Left(Format(GRD1.TextMatrix(i, 16), "000000"), 4)), min_pr02, max_pr02, strUserNum)
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
Dim strCon As String, strConAsc As String
Dim subSQL1 As String, subSQL2 As String
Dim intHideRow As Integer 'Add By Sindy 2013/11/19
   
   Label1.Caption = "共 0 筆"
   GRD1.Clear
   SetGrd
   strCon = "": strConAsc = ""
   
   If Val(txt1(0)) = 0 Or Val(txt1(1)) = 0 Then
      MsgBox "請輸入起迄日期！", vbExclamation, "操作錯誤！"
      If Val(txt1(0)) = 0 Then txt1(0).SetFocus
      If Val(txt1(1)) = 0 Then txt1(1).SetFocus
      Exit Sub
   End If
   
   '只查詢個人未確認資料
   If m_QueryType = 2 Then
      strCon = strCon & " and B1411 is null and B1405 is null"
   '查詢系統確認且為請假處理的資料
   ElseIf m_QueryType = 3 Then
      strCon = strCon & " and B1411='A' and B1405='1'"
   Else
      strCon = strCon & " and B1411 is null"
   End If
   
   '時段
   If Combo1.Text <> "" Then
      strCon = strCon & " and b1403='" & Left(Trim(Combo1.Text), 1) & "'"
   End If
   '部門別
   If txt1(2) <> "" And txt1(3) <> "" Then
      'Modify By Sindy 2023/12/22
      If strSrvDate(1) >= 新部門啟用日 Then
         strCon = strCon & " and s1.ST93>='" & txt1(2) & "' and s1.ST93<='" & txt1(3) & "'"
         strConAsc = strConAsc & " and s1.ST93>='" & txt1(2) & "' and s1.ST93<='" & txt1(3) & "'"
      Else
      '2023/12/22 END
         strCon = strCon & " and s1.ST03>='" & txt1(2) & "' and s1.ST03<='" & txt1(3) & "'"
         strConAsc = strConAsc & " and s1.ST03>='" & txt1(2) & "' and s1.ST03<='" & txt1(3) & "'"
      End If
   End If
   '員工代號
   If txt1(4) <> "" And txt1(5) <> "" Then
      strCon = strCon & " and s1.ST01>='" & txt1(4) & "' and s1.ST01<='" & txt1(5) & "'"
      strConAsc = strConAsc & " and s1.ST01>='" & txt1(4) & "' and s1.ST01<='" & txt1(5) & "'"
   End If
   
   m_blnColOrderAsc = True
   Screen.MousePointer = vbHourglass
   
   '有無請假
   subSQL1 = "(SELECT SA01 as Userid,SA02 as date1,SA04 as date2,'Y' as IsY,sqldatet(SA02) as 起始日期,substr(sqltime(SA03||'00'),1,5) as 起始時間,sqldatet(SA04) as 迄止日期,substr(sqltime(SA05||'00'),1,5) as 迄止時間 FROM staff_Absence,staff s1 WHERE SA01=s1.st01(+)" & _
             " and ((SA02>=" & DBDATE(txt1(0)) & " and SA02<=" & DBDATE(txt1(1)) & ") or (SA04>=" & DBDATE(txt1(0)) & " and SA04<=" & DBDATE(txt1(1)) & ") or (" & DBDATE(txt1(0)) & " between SA02 and SA04) or (" & DBDATE(txt1(1)) & " between SA02 and SA04))" & strConAsc & _
             " union SELECT SB01 as Userid,SB02 as date1,SB04 as date2,'Y' as IsY,sqldatet(SB02) as 起始日期,substr(sqltime(SB03||'00'),1,5) as 起始時間,sqldatet(SB04) as 迄止日期,substr(sqltime(SB05||'00'),1,5) as 迄止時間 FROM staff_busi_trip,staff s1 WHERE SB01=s1.st01(+)" & _
             " and ((SB02>=" & DBDATE(txt1(0)) & " and SB02<=" & DBDATE(txt1(1)) & ") or (SB04>=" & DBDATE(txt1(0)) & " and SB04<=" & DBDATE(txt1(1)) & ") or (" & DBDATE(txt1(0)) & " between SB02 and SB04) or (" & DBDATE(txt1(1)) & " between SB02 and SB04))" & strConAsc & _
             " union SELECT B1003 as Userid,B1004 as date1,B1006 as date2,'Y' as IsY,sqldatet(B1004) as 起始日期,substr(sqltime(B1005||'00'),1,5) as 起始時間,sqldatet(B1006) as 迄止日期,substr(sqltime(B1007||'00'),1,5) as 迄止時間 FROM ABS010,staff s1 WHERE B1002 in('01','03') and B1018 not in('" & 退回 & "','" & 註銷 & "','" & 已核准 & "') and B1003=s1.st01(+)" & _
             " and ((B1004>=" & DBDATE(txt1(0)) & " and B1004<=" & DBDATE(txt1(1)) & ") or (B1006>=" & DBDATE(txt1(0)) & " and B1006<=" & DBDATE(txt1(1)) & ") or (" & DBDATE(txt1(0)) & " between B1004 and B1006) or (" & DBDATE(txt1(1)) & " between B1004 and B1006))" & strConAsc & _
             ") V1"
   '每日的最小和最大打卡時間
   subSQL2 = "(select scd01,pr01,nvl(min(pr02),0) as min_pr02,nvl(max(pr02),0) as max_pr02 from pollrecord,staffcarddata where pr03=scd02(+) and  pr01>=" & DBDATE(txt1(0)) & " and pr01<=" & DBDATE(txt1(1)) & " group by scd01,pr01) V2"
   '打卡異常資料
   'Modify By Sindy 2023/12/22
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "select '' as V,sqldatet(b1402) as 日期,A0922 as 部門,s1.ST02 as 姓名,decode(b1403,'A','上班','P','下班',b1403) as 時段," & _
               "sqltime6(b1404) as 打卡時間,V1.IsY as 有無請假,ac03 as 個人確認," & _
               "b1407 as 未打卡原因,s2.st02 as 主管,decode(b1409,'Y','同意','N','不同意',b1409) as 主管批示,sqldatet(b1410) as 批示日期," & _
               "s1.ST93 as ST03,b1401,b1406,V2.min_pr02 as 第一筆打卡時間,V2.max_pr02 as 最後一筆打卡時間," & _
               "V1.起始日期 as 起始日期,V1.起始時間 as 起始時間,V1.迄止日期 as 迄止日期,V1.迄止時間 as 迄止時間" & _
               " from abs014,staff s1,staff s2,ACC090NEW,allcode," & subSQL1 & "," & subSQL2 & _
               " where b1401=s1.st01(+)" & _
               " and b1408=s2.st01(+)" & _
               " and s1.ST93=A0921(+)" & _
               " and ac01(+)='10' and b1405=ac02(+)" & _
               " and b1402>=" & DBDATE(txt1(0)) & " and b1402<=" & DBDATE(txt1(1)) & strCon & _
               " and B1401=V1.Userid(+) and B1402 between V1.date1(+) and V1.date2(+)" & _
               " and B1401=V2.scd01(+) and B1402=V2.pr01(+)" & _
               " order by b1402 desc,b1401 asc,b1403 asc"
   Else
   '2023/12/22 END
      strSql = "select '' as V,sqldatet(b1402) as 日期,A0902 as 部門,s1.ST02 as 姓名,decode(b1403,'A','上班','P','下班',b1403) as 時段," & _
               "sqltime6(b1404) as 打卡時間,V1.IsY as 有無請假,ac03 as 個人確認," & _
               "b1407 as 未打卡原因,s2.st02 as 主管,decode(b1409,'Y','同意','N','不同意',b1409) as 主管批示,sqldatet(b1410) as 批示日期," & _
               "s1.ST03 as ST03,b1401,b1406,V2.min_pr02 as 第一筆打卡時間,V2.max_pr02 as 最後一筆打卡時間," & _
               "V1.起始日期 as 起始日期,V1.起始時間 as 起始時間,V1.迄止日期 as 迄止日期,V1.迄止時間 as 迄止時間" & _
               " from abs014,staff s1,staff s2,ACC090,allcode," & subSQL1 & "," & subSQL2 & _
               " where b1401=s1.st01(+)" & _
               " and b1408=s2.st01(+)" & _
               " and s1.ST03=A0901(+)" & _
               " and ac01(+)='10' and b1405=ac02(+)" & _
               " and b1402>=" & DBDATE(txt1(0)) & " and b1402<=" & DBDATE(txt1(1)) & strCon & _
               " and B1401=V1.Userid(+) and B1402 between V1.date1(+) and V1.date2(+)" & _
               " and B1401=V2.scd01(+) and B1402=V2.pr01(+)" & _
               " order by b1402 desc,b1401 asc,b1403 asc"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      'Add By Sindy 2013/11/19 ex:102/11/18-71011-P
      '不是3.查詢系統確認且為請假處理的資料時,過濾重覆資料
      intHideRow = 0
      If m_QueryType <> 3 Then
         For i = GRD1.Rows - 1 To 1 Step -1
            If i > 1 Then
               If Trim(GRD1.TextMatrix(i, 1)) = Trim(GRD1.TextMatrix(i - 1, 1)) And _
                  Trim(GRD1.TextMatrix(i, 4)) = Trim(GRD1.TextMatrix(i - 1, 4)) And _
                  Trim(GRD1.TextMatrix(i, 13)) = Trim(GRD1.TextMatrix(i - 1, 13)) Then
                  GRD1.RowHeight(i) = 0
                  intHideRow = intHideRow + 1
               End If
            End If
         Next i
      End If
      '2013/11/19 END
      Label1.Caption = "共 " & rsTmp.RecordCount - intHideRow & " 筆"
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
'   If rsTmp.RecordCount > 0 Then
'      For i = 0 To GRD1.Cols - 1
'         GRD1.col = i
'         GRD1.CellBackColor = &HFFC0C0
'      Next i
'   End If
   GRD1.Visible = True
   
   rsTmp.Close
   Screen.MousePointer = vbDefault
   
EXITSUB:
   Set rsTmp = Nothing
End Sub

'異常確認
Public Sub cmdok_Click()
Dim strKind As String
Dim bolSelect As Boolean
   
   bolSelect = False
   For i = 1 To GRD1.Rows - 1
      GRD1.col = 0
      GRD1.row = i
      If GRD1.TextMatrix(i, 0) = "V" Then
         bolSelect = True
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
         frm160012_1.m_WorkType = 0
         frm160012_1.m_B1401 = Trim(GRD1.TextMatrix(i, 13))
         frm160012_1.m_B1402 = DBDATE(GRD1.TextMatrix(i, 1))
         frm160012_1.m_B1403 = strKind
         frm160012_1.Show
         Me.Hide
         Exit For
      End If
   Next i
   If bolSelect = False Then
      If m_QueryType = 3 Then
         Call cmdQuery3_Click
      ElseIf m_QueryType = 2 Then
         Call cmdQuery2_Click
      Else
         Call cmdQuery_Click
      End If
   End If
End Sub

'新增正常時間異常打卡
Private Sub cmdAdd_Click()
   frm160012_1.m_WorkType = 1
   frm160012_1.Show
   Me.Hide
End Sub

'查詢異常未處理
Public Sub cmdQuery_Click()
   m_QueryType = 1
   cmdEmail.Enabled = True
   Call QueryData
End Sub

'未個人確認的
Public Sub cmdQuery2_Click()
   m_QueryType = 2
   cmdEmail.Enabled = True
   Call QueryData
End Sub

'查詢系統確認且為請假處理的資料
Public Sub cmdQuery3_Click()
   m_QueryType = 3
   cmdEmail.Enabled = False
   Call QueryData
End Sub

Private Sub ChangSelect()
   If cmdSelect.Caption = "全選" Then
      cmdSelect.Caption = "取消全選"
   ElseIf cmdSelect.Caption = "取消全選" Then
      cmdSelect.Caption = "全選"
   End If
End Sub

Private Sub cmdSelect_Click()
Dim k As Integer
   
   If cmdSelect.Caption = "全選" Then
      GRD1.Visible = False
      For k = 1 To GRD1.Rows - 1
         GRD1.col = 0
         GRD1.row = k
         If Trim(GRD1.TextMatrix(k, 1)) <> "" Then
            If Trim(GRD1.Text) = "" Then
               GRD1.Text = "V"
               For i = 0 To GRD1.Cols - 1
                  GRD1.col = i
                  GRD1.CellBackColor = &HFFC0C0
               Next i
            End If
         End If
      Next k
      GRD1.Visible = True
   ElseIf cmdSelect.Caption = "取消全選" Then
      GRD1.Visible = False
      For k = 1 To GRD1.Rows - 1
         GRD1.col = 0
         GRD1.row = k
         If Trim(GRD1.TextMatrix(k, 1)) <> "" Then
            If Trim(GRD1.Text) = "V" Then
               GRD1.Text = ""
               For i = 0 To GRD1.Cols - 1
                  GRD1.col = i
                  GRD1.CellBackColor = QBColor(15)
               Next i
            End If
         End If
      Next k
      GRD1.Visible = True
   End If
   Call ChangSelect
End Sub

Private Sub Command1_Click()
   PollingData True
End Sub

Private Sub PollingData(Optional bNotAuto As Boolean, Optional bUpdateTime As Boolean)
   Dim iRecs As Integer
   Dim arrIpList
   Dim ii As Integer
   Dim strShowMsg As String
   
   strShowMsg = ""
   HTAips = GetHtaIP()
   If HTAips <> "" Then
      arrIpList = Split(HTAips, ";")
      For ii = LBound(arrIpList) To UBound(arrIpList)
         HTAip = arrIpList(ii)
         If HTAip <> "" Then
         
'            If bUpdateTime = True Then
'               If HTAWriteTime(True) = True Then
'                  lstHistory.AddItem Now & "--> (" & HTAip & ") 指紋機時間已同步！", 0
'               Else
'                  lstHistory.AddItem Now & "--> (" & HTAip & ") 指紋機時間同步失敗！", 0
'               End If
'            End If
'
'            lstHistory.AddItem Now & " --> (" & HTAip & ") 開始接收刷卡紀錄" & IIf(bNotAuto, "(手動)", ""), 0
            If HTAPolling(iRecs, True) = True Then
               If iRecs = 0 Then
                  'MsgBox "(" & HTAip & ") 沒有新刷卡紀錄可接收！"
                  strShowMsg = strShowMsg & "(" & HTAip & ") 沒有新刷卡紀錄可接收！" & vbCrLf
               Else
                  'MsgBox "(" & HTAip & ") 接收完成共" & iRecs & "筆！"
                  strShowMsg = strShowMsg & "(" & HTAip & ") 接收完成共" & iRecs & "筆！" & vbCrLf
               End If
            Else
               'MsgBox "(" & HTAip & ") 刷卡紀錄接收失敗！"
               strShowMsg = strShowMsg & "(" & HTAip & ") 刷卡紀錄接收失敗！" & vbCrLf
            End If
         End If
      Next
      MsgBox strShowMsg
   Else
      MsgBox "考勤機IP未設定！"
   End If
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   txt1(0) = Left(ChangeWStringToTString(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##")))), 5) & "01"
   txt1(1) = strSrvDate(2)
   Call ChangSelect
   Call cmdQuery_Click
End Sub

Private Sub Form_Unload(Cancel As Integer)
'   Call PUB_GetLock("", UCase(Me.Name)) 'Add By Sindy 2019/2/21
   Set frm160012 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("V", "日期", "部門", "姓名", "時段", _
                           "打卡時間", "有無請假", "個人確認", "未打卡原因", "主管", _
                           "主管批示", "批示日期", "ST03", _
                           "b1401", "b1406", "第一筆打卡時間", "最後一筆打卡時間", _
                           "起始日期", "起始時間", "迄止日期", "迄止時間")
   '查詢系統確認且為請假處理的資料
   If m_QueryType = 3 Then
      arrGridHeadWidth = Array(200, 800, 800, 700, 500, _
                               800, 0, 0, 1000, 0, _
                               0, 0, 0, _
                               0, 0, 0, 0, _
                               900, 900, 900, 900)
   Else
      arrGridHeadWidth = Array(200, 800, 800, 700, 500, _
                               800, 700, 800, 1000, 700, _
                               800, 800, 0, _
                               0, 0, 0, 0, _
                               0, 0, 0, 0)
   End If
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

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 2, 3, 4, 5
         KeyAscii = UpperCase(KeyAscii)
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If CheckIsTaiwanDate(txt1(Index), False) = False And Trim(txt1(Index)) <> "" Then
            Call txt1_GotFocus(Index)
            Cancel = True
            MsgBox "請輸入民國日期不含/！", vbInformation, "輸入日期錯誤"
            Exit Sub
         End If
         
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
            If Val(txt1(Index)) > Val(txt1(Index + 1)) Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 1 Then
            If RunNick2(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 2, 3
         If Index = 2 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 3 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 4, 5
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 4 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 5 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
   End Select
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
