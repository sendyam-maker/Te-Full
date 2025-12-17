VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frm160018 
   BorderStyle     =   1  '單線固定
   Caption         =   "下班逾30分鐘原因確認"
   ClientHeight    =   5950
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   8950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5950
   ScaleWidth      =   8950
   Begin VB.CommandButton cmdOK 
      Caption         =   "全選(&A)"
      Height          =   345
      Index           =   3
      Left            =   5610
      TabIndex        =   10
      Top             =   90
      Width           =   800
   End
   Begin VB.CommandButton CmdExcel 
      Caption         =   "Excel清單"
      Height          =   345
      Left            =   6480
      TabIndex        =   12
      Top             =   870
      Width           =   1070
   End
   Begin VB.CheckBox Check1 
      Caption         =   "尚未輸入逾時原因"
      Height          =   250
      Left            =   3540
      TabIndex        =   8
      Top             =   960
      Width           =   2200
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   7
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   7
      Top             =   930
      Width           =   470
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   6
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   6
      Top             =   930
      Width           =   470
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "明細資料"
      Height          =   345
      Index           =   1
      Left            =   6480
      TabIndex        =   11
      Top             =   90
      Width           =   1070
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1440
      MaxLength       =   3
      TabIndex        =   2
      Top             =   330
      Width           =   470
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   3
      Top             =   330
      Width           =   470
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "結束"
      Height          =   345
      Index           =   2
      Left            =   7620
      TabIndex        =   13
      Top             =   90
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   1440
      MaxLength       =   6
      TabIndex        =   4
      Top             =   630
      Width           =   740
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   5
      Left            =   2310
      MaxLength       =   6
      TabIndex        =   5
      Top             =   630
      Width           =   740
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1440
      MaxLength       =   7
      TabIndex        =   0
      Top             =   30
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2490
      MaxLength       =   7
      TabIndex        =   1
      Top             =   30
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "查詢"
      Default         =   -1  'True
      Height          =   345
      Index           =   0
      Left            =   4620
      TabIndex        =   9
      Top             =   90
      Width           =   915
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GRD1 
      Height          =   3920
      Left            =   60
      TabIndex        =   14
      Top             =   1260
      Width           =   8850
      _ExtentX        =   15610
      _ExtentY        =   6914
      _Version        =   393216
      Cols            =   9
      FixedCols       =   0
      ScrollTrack     =   -1  'True
      AllowUserResizing=   3
      FormatString    =   "V|員工編號|姓名|日期|打卡起迄時間|上班時段|有填加班單|逾時原因|備註"
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
      _Band(0).Cols   =   9
   End
   Begin VB.ComboBox cboDept 
      Height          =   260
      Index           =   1
      Left            =   7110
      TabIndex        =   22
      Text            =   "cboDept"
      Top             =   5130
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.ComboBox cboDept 
      Height          =   260
      Index           =   0
      Left            =   6210
      TabIndex        =   23
      Text            =   "cboDept"
      Top             =   5130
      Visible         =   0   'False
      Width           =   1965
   End
   Begin VB.Label Label5 
      Caption         =   "備註：您有權限查詢的部門別為"
      ForeColor       =   &H00000080&
      Height          =   390
      Left            =   270
      TabIndex        =   21
      Top             =   5490
      Width           =   8640
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "所　　別："
      Height          =   180
      Left            =   480
      TabIndex        =   20
      Top             =   960
      Width           =   900
   End
   Begin VB.Line Line4 
      X1              =   1860
      X2              =   2340
      Y1              =   1050
      Y2              =   1050
   End
   Begin VB.Label Label4 
      Caption         =   "(註：可雙擊點選，即可進入明細資料)"
      ForeColor       =   &H00FF0000&
      Height          =   200
      Left            =   2070
      TabIndex        =   19
      Top             =   5220
      Width           =   3590
   End
   Begin VB.Label Label1 
      Caption         =   "共 0 筆"
      ForeColor       =   &H00FF0000&
      Height          =   210
      Left            =   270
      TabIndex        =   18
      Top             =   5220
      Width           =   1400
   End
   Begin VB.Line Line3 
      X1              =   1860
      X2              =   2340
      Y1              =   450
      Y2              =   450
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "部  門  別："
      Height          =   180
      Left            =   480
      TabIndex        =   17
      Top             =   360
      Width           =   900
   End
   Begin VB.Line Line2 
      X1              =   2310
      X2              =   2520
      Y1              =   150
      Y2              =   150
   End
   Begin VB.Line Line1 
      X1              =   2130
      X2              =   2340
      Y1              =   750
      Y2              =   750
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Left            =   480
      TabIndex        =   16
      Top             =   660
      Width           =   900
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "日　　期："
      Height          =   180
      Left            =   480
      TabIndex        =   15
      Top             =   60
      Width           =   900
   End
End
Attribute VB_Name = "frm160018"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Sindy 2025/10/14
Option Explicit

Public m_IsAbsBossST03 As String
Public m_strEmp As String '所屬簽核的人員
Dim m_blnColOrderAsc As Boolean '欄位資料由小到大排序
Dim i As Integer, j As Integer
Public cmdState As Integer '紀錄作用按鍵
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim m_intColumn As Integer


'查詢資料
Public Sub QueryData()
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCon As String, strConAsc As String
Dim subSQL1 As String, subSQL2 As String
   
   Label1.Caption = "共 0 筆"
   GRD1.Clear
   SetGrd
   strCon = "": strConAsc = ""
   
   If Check1.Value = 0 Then
      If Val(txt1(0)) = 0 Or Val(txt1(1)) = 0 Then
         MsgBox "請輸入起迄日期！", vbExclamation, "操作錯誤！"
         If Val(txt1(0)) = 0 Then txt1(0).SetFocus
         If Val(txt1(1)) = 0 Then txt1(1).SetFocus
         Exit Sub
      End If
'   Else
'      '尚未輸入逾時原因
'      strCon = strCon & " and B1504 is null"
'      '抓資料的起迄日期
'      strExc(0) = "SELECT min(b1502),max(b1502) FROM abs015 WHERE b1504 is null"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         If txt1(0) > TransDate(RsTemp.Fields(0), 1) Then '預設值
'            txt1(0) = TransDate(RsTemp.Fields(0), 1)
'         End If
'         If txt1(1) < TransDate(RsTemp.Fields(1), 1) Then '預設值
'            txt1(1) = TransDate(RsTemp.Fields(1), 1)
'         End If
'      End If
   End If
   
   '部門別
   If txt1(2) <> "" And txt1(3) <> "" Then
      strCon = strCon & " and ST93>='" & txt1(2) & "' and ST93<='" & txt1(3) & "'"
      strConAsc = strConAsc & " and s1.ST93>='" & txt1(2) & "' and s1.ST93<='" & txt1(3) & "'"
   End If
   '員工代號
   If txt1(4) <> "" And txt1(5) <> "" Then
      strCon = strCon & " and B1501>='" & txt1(4) & "' and B1501<='" & txt1(5) & "'"
      strConAsc = strConAsc & " and s1.ST01>='" & txt1(4) & "' and s1.ST01<='" & txt1(5) & "'"
   End If
   '所別
   If txt1(6) <> "" And txt1(7) <> "" Then
      strCon = strCon & " and ST06>='" & txt1(6) & "' and ST06<='" & txt1(7) & "'"
      strConAsc = strConAsc & " and s1.ST06>='" & txt1(6) & "' and s1.ST06<='" & txt1(6) & "'"
   End If
   
   '所屬簽核的人員
   If m_IsAbsBossST03 <> "" Then
      If m_strEmp <> "" Then
         strCon = strCon & " and ST01 in(" & m_strEmp & ")"
         strConAsc = strConAsc & " and s1.ST01 in(" & m_strEmp & ")"
      End If
   End If
   '部門別權限限制
   If m_IsAbsBossST03 <> "" Then
      strCon = strCon & " and ST93 in (" & m_IsAbsBossST03 & ")"
      strConAsc = strConAsc & " and s1.ST93 in (" & m_IsAbsBossST03 & ")"
   End If
   
   If Check1.Value = 1 Then
      '尚未輸入逾時原因
      strCon = strCon & " and B1504 is null"
      '抓資料的起迄日期
      strExc(0) = "SELECT min(b1502),max(b1502) FROM abs015,staff WHERE b1504 is null and b1501=st01(+)" & strCon
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If IsNull(RsTemp.Fields(0)) = False Then
            If txt1(0) > TransDate(RsTemp.Fields(0), 1) Then
               txt1(0) = TransDate(RsTemp.Fields(0), 1)
            End If
         End If
         If IsNull(RsTemp.Fields(1)) = False Then
            If txt1(1) < TransDate(RsTemp.Fields(1), 1) Then
               txt1(1) = TransDate(RsTemp.Fields(1), 1)
            End If
         End If
      End If
   End If
   '日期
   strCon = strCon & " and B1502>=" & DBDATE(txt1(0)) & " and B1502<=" & DBDATE(txt1(1))
   
   m_blnColOrderAsc = True
   Screen.MousePointer = vbHourglass
   
   '有無加班
   subSQL1 = "(SELECT SO01 as Userid,SO02 as date1,'Y' as IsY" & _
             " FROM staff_overtime,staff s1 WHERE SO01=s1.st01(+)" & _
             " and SO02>=" & DBDATE(txt1(0)) & " and SO02<=" & DBDATE(txt1(1)) & strConAsc & _
             " union SELECT B1003 as Userid,B1004 as date1,'Y' as IsY FROM ABS010,staff s1 WHERE B1002 in('02')" & _
             " and B1018 not in('" & 註銷 & "','" & 已核准 & "') and B1003=s1.st01(+)" & _
             " and B1004>=" & DBDATE(txt1(0)) & " and B1004<=" & DBDATE(txt1(1)) & strConAsc & _
             ") V1"
   '每日的最小和最大打卡時間
   subSQL2 = "(select scd01,pr01,substr('000000'||nvl(min(pr02),0),-6) as min_pr02,substr('000000'||nvl(max(pr02),0),-6) as max_pr02" & _
             " from pollrecord,staffcarddata where pr03=scd02(+)" & _
             " and pr01>=" & DBDATE(txt1(0)) & " and pr01<=" & DBDATE(txt1(1)) & " group by scd01,pr01) V2"
   strSql = "select '' as V,st01 as 員工編號,ST02 as 姓名,sqldatet(b1502) as 日期," & _
            "substr(V2.min_pr02,1,2)||':'||substr(V2.min_pr02,3,2)||':'||substr(V2.min_pr02,-2)||'~'||substr(V2.max_pr02,1,2)||':'||substr(V2.max_pr02,3,2)||':'||substr(V2.max_pr02,-2) as 打卡起迄時間,b1503 as 上班時段,V1.IsY as 有填加班單," & _
            "decode(b1504,null,b1504,b1504||'.'||ac03) as 逾時原因,b1505 備註" & _
            " from abs015,staff,ACC090NEW,allcode," & subSQL1 & "," & subSQL2 & _
            " where b1501=st01(+)" & _
            " and ST93=A0921(+)" & _
            " and ac01(+)='18' and b1504=ac02(+)" & strCon & _
            " and B1501=V1.Userid(+) and B1502=V1.date1(+)" & _
            " and B1501=V2.scd01(+) and B1502=V2.pr01(+)" & _
            " order by b1502 asc,b1501 asc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Set GRD1.Recordset = rsTmp
      Label1.Caption = "共 " & rsTmp.RecordCount & " 筆"
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

Private Sub CmdExcel_Click()
   PrintExcel
End Sub

Sub PrintExcel()
Dim ii As Integer
Dim strTempFile As String
   
   Set xlsAnnuity = New Excel.Application
   xlsAnnuity.Visible = True
   xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
   xlsAnnuity.Workbooks.add
   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   wksAnnuity.Cells.NumberFormatLocal = "@" '文字
   xlsAnnuity.ActiveWindow.Zoom = 100 '畫面比例100%太大了,調整為75%
   '把Excel的警告訊息關掉
   xlsAnnuity.DisplayAlerts = False
   
   wksAnnuity.PageSetup.PaperSize = 9 'A4
   wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
   'wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
   wksAnnuity.PageSetup.LeftMargin = 0 '邊界
   wksAnnuity.PageSetup.RightMargin = 0
   wksAnnuity.PageSetup.TopMargin = xlsAnnuity.InchesToPoints(0.4)
   wksAnnuity.PageSetup.BottomMargin = xlsAnnuity.InchesToPoints(0.5)
   wksAnnuity.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
   
'   xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
'   xlsAnnuity.Workbooks.add
'   Set wksAnnuity = xlsAnnuity.Worksheets(1)
   wksAnnuity.Activate
   
   '設定各欄位長度
   wksAnnuity.Columns("A:A").ColumnWidth = 9
   wksAnnuity.Columns("B:B").ColumnWidth = 9
   wksAnnuity.Columns("C:C").ColumnWidth = 9
   wksAnnuity.Columns("D:D").ColumnWidth = 15
   wksAnnuity.Columns("E:E").ColumnWidth = 12
   wksAnnuity.Columns("F:F").ColumnWidth = 12
   wksAnnuity.Columns("G:G").ColumnWidth = 15
   wksAnnuity.Columns("H:H").ColumnWidth = 15
   
   '標題
   m_intColumn = 1
   xlsAnnuity.Range("D" & m_intColumn).Value = "下班逾30分鐘原因確認清單明細"
   xlsAnnuity.Range("A" & m_intColumn & ":" & "H" & m_intColumn).Select
   With xlsAnnuity.Selection
      .HorizontalAlignment = xlCenter '置中
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = True
   End With
   With xlsAnnuity.Selection.Font
      .Bold = True '粗體
      .Name = "新細明體"
      .Size = 16
   End With
   m_intColumn = m_intColumn + 1
   xlsAnnuity.Range("A" & m_intColumn).Value = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   m_intColumn = m_intColumn + 2
   xlsAnnuity.Range("A" & m_intColumn).Value = "員工編號"
   xlsAnnuity.Range("B" & m_intColumn).Value = "姓名"
   'xlsAnnuity.Range("B" & m_intColumn).HorizontalAlignment = xlCenter
   xlsAnnuity.Range("C" & m_intColumn).Value = "日期"
   'xlsAnnuity.Range("C" & m_intColumn).HorizontalAlignment = xlCenter
   xlsAnnuity.Range("D" & m_intColumn).Value = "打卡起迄時間"
   xlsAnnuity.Range("E" & m_intColumn).Value = "上班時段"
   xlsAnnuity.Range("F" & m_intColumn).Value = "填寫加班單"
   xlsAnnuity.Range("G" & m_intColumn).Value = "逾時原因"
   xlsAnnuity.Range("H" & m_intColumn).Value = "備註"
   xlsAnnuity.Range("A" & m_intColumn & ":" & "M" & m_intColumn).Select
'   With xlsAnnuity.Selection
'       .HorizontalAlignment = xlCenter
'       .VerticalAlignment = xlCenter
'       .WrapText = False
'       .Orientation = 0
'       .AddIndent = False
'       .IndentLevel = 0
'       .ShrinkToFit = False
'       .ReadingOrder = xlContext
'       .MergeCells = False
'   End With
'   xlsAnnuity.Selection.Borders(xlDiagonalDown).LineStyle = xlNone
'   xlsAnnuity.Selection.Borders(xlDiagonalUp).LineStyle = xlNone
'   xlsAnnuity.Selection.Borders(xlEdgeLeft).LineStyle = xlNone
   With xlsAnnuity.Selection.Borders(xlEdgeTop)
       .LineStyle = xlContinuous
       .ColorIndex = xlAutomatic
       .tintandshade = 0
       .Weight = xlThin
   End With
   With xlsAnnuity.Selection.Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .ColorIndex = xlAutomatic
       .tintandshade = 0
       .Weight = xlThin
   End With
   
   '列印明細
   With GRD1
   For ii = 1 To .Rows - 1
      m_intColumn = m_intColumn + 1
      xlsAnnuity.Range("A" & m_intColumn).Value = .TextMatrix(ii, 1)
      xlsAnnuity.Range("B" & m_intColumn).Value = .TextMatrix(ii, 2)
      xlsAnnuity.Range("C" & m_intColumn).Value = .TextMatrix(ii, 3)
      xlsAnnuity.Range("D" & m_intColumn).Value = .TextMatrix(ii, 4)
      xlsAnnuity.Range("E" & m_intColumn).Value = .TextMatrix(ii, 5)
      xlsAnnuity.Range("F" & m_intColumn).Value = .TextMatrix(ii, 6)
      xlsAnnuity.Range("G" & m_intColumn).Value = .TextMatrix(ii, 7)
      xlsAnnuity.Range("H" & m_intColumn).Value = .TextMatrix(ii, 8)
   Next ii
   m_intColumn = m_intColumn + 2
   xlsAnnuity.Range("A" & m_intColumn).Value = "共 " & .Rows - 1 & " 筆"
   End With
   xlsAnnuity.ActiveSheet.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
'   '列印標題
'   xlsAnnuity.ActiveSheet.PageSetup.PrintTitleRows = "$1:$4"
'
'   strTempFile = App.path & "\$$demo.pdf"
'   If Val(xlsAnnuity.Version) < 12 Then
'      xlsAnnuity.Workbooks(1).SaveAs FileName:=strTempFile, FileFormat:=-4143
'   Else
'      xlsAnnuity.Workbooks(1).SaveAs FileName:=strTempFile, FileFormat:=56
'   End If
'   xlsAnnuity.Workbooks(1).PrintOut
   
'   'xlTypePDF   0  PDF：可攜式文件格式檔案 (.pdf)
'   'Quality:=xlQualityStandard 0  標準品質
'   '各參數解釋:
'   'Type 必要  XlFixedFormatTyp  匯出目標的檔案格式類型。
'   'FileName   選用  Variant  要儲存之檔案的檔案名稱。 可以包含完整路徑，否則 Microsoft Excel 會將檔案儲存在目前的資料夾中。
'   'Quality 選用  Variant  選用 XlFixedFormatQuality。 這會指定已發佈檔案的品質。
'   'IncludeDocProperties 選用  Variant  若要包含檔案屬性，則為 True 。否則 為 False。
'   'IgnorePrintAreas  選用  Variant  True 是表示忽略所有發佈時設定的列印範圍;否則 為 False。
'   'From 選用  Variant  要發佈的起始頁碼。 如果省略此引數，將從頭開始列印。
'   'To   選用  Variant  要發佈的最後一頁頁碼。 如果省略此引數，將發佈至最後一頁。
'   'OpenAfterPublish 選用  Variant  True 是表示在發佈後在檢視器中顯示檔案;否則 為 False。
'   'FixedFormatExtClassPtr 選用  Variant  FixedFormatExt 類別的指標。
'   xlsAnnuity.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strTempFile, Quality:=0, _
'   IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=True
'   'ShellExecute 0, "open", strTempFile, vbNullString, vbNullString, 1
   
'   xlsAnnuity.Workbooks.Close 'SaveChanges:=False
'   xlsAnnuity.Quit
   Set xlsAnnuity = Nothing
End Sub

Public Sub cmdok_Click(Index As Integer)
   cmdState = Index
   
   PubShowNextData
End Sub

Public Sub PubShowNextData()
Dim strKind As String
Dim bolSelect As Boolean

   Select Case cmdState
      Case 0 '查詢
         Call QueryData
      Case 1 '明細資料
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
               frm160018_1.m_B1501 = GRD1.TextMatrix(i, 1) '員工代號
               frm160018_1.m_B1502 = DBDATE(GRD1.TextMatrix(i, 3)) '日期
               frm160018_1.Show
               Me.Hide
               Exit For
            End If
         Next i
         If bolSelect = False Then QueryData '重新查詢
      Case 2 '離開
         Unload Me
      Case 3 '全選
         GRD1.Visible = False
         If GRD1.Rows > 1 Then
            If GRD1.TextMatrix(1, 1) <> "" Then
               For j = 1 To GRD1.Rows - 1
                  GRD1.col = 0
                  GRD1.row = j
                  GRD1.Text = "V"
                  For i = 0 To GRD1.Cols - 1
                     GRD1.col = i
                     GRD1.CellBackColor = &HFFC0C0
                  Next i
               Next j
            End If
         End If
         GRD1.Visible = True
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   
   Call PUB_SetQFormCol_ABS(m_IsAbsBossST03, m_strEmp, Me.Name, txt1(0), txt1(1), txt1(4), txt1(5), _
      cboDept(0), cboDept(1), txt1(2), txt1(3), txt1(6), txt1(7), Me.Label5)
   If GetStaffDepartment(strUserNum) = "M51" Or _
      GetStaffDepartment(strUserNum) = "M21" Then
      CmdExcel.Visible = True
   Else
      CmdExcel.Visible = False
   End If
   Check1.Value = 1
   '人事一進入此作業預設抓全部尚未輸入逾時原因
   If GetStaffDepartment(strUserNum) = "M21" Then
      txt1(2).Text = ""
      txt1(3).Text = ""
      txt1(4).Text = ""
      txt1(5).Text = ""
      txt1(6).Text = ""
      txt1(7).Text = ""
   End If
   Call cmdok_Click(0) '查詢
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm160018 = Nothing
End Sub

Private Sub SetGrd()
   Dim arrGridHeadText, arrGridHeadWidth
   Dim iRow As Integer
   
   arrGridHeadText = Array("V", "員工編號", "姓名", "日期", "打卡起迄時間", _
                           "上班時段", "有填加班單", "逾時原因", "備註")
   arrGridHeadWidth = Array(200, 600, 800, 800, 1400, _
                            1400, 600, 1000, 1500)
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
      Case 0, 1, 6, 7
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
      Case 6, 7
         If Index = 6 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 7 Then
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
         Call cmdok_Click(1)
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
