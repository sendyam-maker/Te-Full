VERSION 5.00
Begin VB.Form frm160201 
   BorderStyle     =   1  '單線固定
   Caption         =   "出缺勤統計通知單"
   ClientHeight    =   3260
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   5440
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3260
   ScaleWidth      =   5440
   Begin VB.CommandButton cmdok 
      Caption         =   "電子化線上確認(&E)"
      Height          =   435
      Index           =   2
      Left            =   1700
      TabIndex        =   14
      Top             =   90
      Width           =   1785
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   495
      Left            =   4650
      ScaleHeight     =   460
      ScaleWidth      =   650
      TabIndex        =   13
      Top             =   1290
      Visible         =   0   'False
      Width           =   690
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   150
      TabIndex        =   10
      Top             =   2610
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   11
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   12
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   2010
      MaxLength       =   5
      TabIndex        =   1
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   4
      Left            =   2790
      MaxLength       =   6
      TabIndex        =   5
      Top             =   1590
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2010
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1590
      Width           =   705
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   2640
      MaxLength       =   3
      TabIndex        =   3
      Top             =   1230
      Width           =   555
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2010
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1230
      Width           =   555
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4455
      TabIndex        =   7
      Top             =   90
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3510
      TabIndex        =   6
      Top             =   90
      Width           =   915
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "　　部門代號，因為在寄發Mail時會有其影響"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   3
      Left            =   150
      TabIndex        =   16
      Top             =   2280
      Width           =   3555
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "註：若處理全公司的出缺勤通知單時，不要輸入員工代號及"
      ForeColor       =   &H000000C0&
      Height          =   180
      Index           =   2
      Left            =   150
      TabIndex        =   15
      Top             =   2070
      Width           =   4680
   End
   Begin VB.Line Line2 
      X1              =   2730
      X2              =   2970
      Y1              =   1740
      Y2              =   1740
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "列印年月："
      Height          =   180
      Left            =   1080
      TabIndex        =   9
      Top             =   870
      Width           =   900
   End
   Begin VB.Line Line1 
      X1              =   2580
      X2              =   2970
      Y1              =   1380
      Y2              =   1380
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工代號："
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   8
      Top             =   1620
      Width           =   900
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門代號："
      Height          =   180
      Left            =   1080
      TabIndex        =   0
      Top             =   1260
      Width           =   900
   End
End
Attribute VB_Name = "frm160201"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/21 Form2.0已修改
'Memo By Sindy 2012/12/3 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
'Create by SINDY 2008/12/23
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_str As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
'Dim strTemp(1 To 40) As String
'Dim strTemp(1 To 52) As String
Dim strTemp(1 To 54) As String 'Modify By Sindy 2025/11/13
Dim PaperX As Double
'Dim paperY As Double
Dim iLine As Integer
'Dim LongPrintCurCnt As Long
Public intChoose As Integer 'Add By Sindy 2011/9/2 0.人事系統 1.出缺勤系統
Dim m_Device, m_iPages As Integer  'Add By Sindy 2011/9/2
Dim douExtRate As Double '字型位置縮放比
'Public gShowMsg As Boolean '顯示警示訊息
'Add By Sindy 2022/1/21
Dim strPrinter As String
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim m_intColumn As Integer
Dim strTempFile As String
'2022/1/21 END


Private Sub GetSql()
   m_StrSQL = ""
   If txt1(1) <> "" Then
       'Modify By Sindy 2023/12/22 部門調整改抓ST93
       m_StrSQL = m_StrSQL & " and st93>='" & txt1(1) & "' "
   End If
   If txt1(2) <> "" Then
       'Modify By Sindy 2023/12/22 部門調整改抓ST93
       m_StrSQL = m_StrSQL & " and st93<='" & txt1(2) & "' "
   End If
   If txt1(3) <> "" Then
       m_StrSQL = m_StrSQL & " and st01>='" & txt1(3) & "' "
   End If
   If txt1(4) <> "" Then
       m_StrSQL = m_StrSQL & " and st01<='" & txt1(4) & "' "
   End If
End Sub

Public Sub cmdok_Click(Index As Integer)
Dim strTit As String
Dim strMsg As String
Dim nResponse
Dim i As Integer, intRow As Integer
Dim strMeSql As String
Dim intCurrRow As Integer

Select Case Index
Case 0, 2
        If intChoose <> 1 Then '出缺勤系統
            If txt1(0) = "" Then
                MsgBox "列印年月不可以空白！", vbInformation, "操作錯誤！"
                txt1(0).SetFocus
                Exit Sub
            End If
            If Len(txt1(0)) <= 3 Then
                MsgBox "列印年月輸入錯誤！", vbInformation, "操作錯誤！"
                txt1(0).SetFocus
                Exit Sub
            End If
        End If
        
'         'Add By Sindy 2013/7/11
'         '檢查該列印年月是否尚有未處理的打卡異常資料,若有,則不可執行此作業
'         strMeSql = "SELECT count(*) FROM ABS014 WHERE substr(B1402,1,6)='" & Left(DBDATE(txt1(0) & "01"), 6) & "' and B1411 is null"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strMeSql)
'         If intI = 1 Then
'            If RsTemp.Fields(0) > 0 Then
'               MsgBox "該月份尚有人事未確認的打卡異常資料，不可執行此作業！", vbInformation, "操作錯誤！"
'               Exit Sub
'            End If
'         End If
'         '2013/7/11 END
         
        GetSql
        If Index = 0 Then '列印
            If intChoose = 1 Then '出缺勤系統
'               txt1(1) = "": txt1(2) = ""
'               For i = 1 To 2
'                  If i = 1 Then
'                     '第一筆先讀取自己的
'                     strMeSql = "SELECT * FROM ABS013 WHERE B1301='04' and B1302='" & strUserNum & "' and B1303='" & strUserNum & "'"
'                  ElseIf i = 2 Then
'                     '簽核他人的
'                     strMeSql = "SELECT * FROM ABS013 WHERE B1301='04' and B1302<>'" & strUserNum & "' and B1303='" & strUserNum & "' order by B1302 asc"
'                  End If
                  m_StrSQL = ""
                  If txt1(0) <> "" Then
                      m_StrSQL = m_StrSQL & " and B1314='" & Val(txt1(0)) + 191100 & "' "
                  End If
                  If txt1(3) <> "" Then
                      m_StrSQL = m_StrSQL & " and B1302>='" & txt1(3) & "' "
                  End If
                  If txt1(4) <> "" Then
                      m_StrSQL = m_StrSQL & " and B1302<='" & txt1(4) & "' "
                  End If
                  strMeSql = "SELECT * FROM ABS013 WHERE B1301='04' and B1303='" & strUserNum & "' " & m_StrSQL
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strMeSql)
                  If intI = 1 Then
                     With RsTemp
                        intRow = intRow + .RecordCount
                        .MoveFirst
                        intCurrRow = 0
                        Do While Not .EOF
                           intCurrRow = intCurrRow + 1
                           txt1(0) = Val(RsTemp.Fields("B1314")) - 191100
                           txt1(3) = RsTemp.Fields("B1302"): txt1(4) = RsTemp.Fields("B1302")
                           frm180203.m_B1301 = "" & RsTemp.Fields("B1301")
                           frm180203.m_B1302 = "" & RsTemp.Fields("B1302")
                           frm180203.m_B1303 = "" & RsTemp.Fields("B1303")
'                           If i = 2 And intCurrRow = 1 Then
'                              gShowMsg = True '簽核他人出缺勤資料時,第一筆資料要顯示警示訊息
'                           Else
'                              gShowMsg = False
'                           End If
                           GetSql
'                           Call StrMenu1(False)
                           Call StrMenu_Excel(False)
                           If bolfrm180203ExitForm = True Then GoTo GoToEnd
                           .MoveNext
                           'Add By Sindy 2011/11/4 檢查是否已確認完畢資料
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strMeSql)
                           If intI <> 1 Then Exit Do
                           '2011/11/4 End
                        Loop
                     End With
                  End If
'               Next i
GoToEnd:
               If intRow = 0 Then nResponse = MsgBox("無資料!", vbExclamation + vbOKOnly, "每月出缺勤統計確認")
            Else
               Screen.MousePointer = vbHourglass
'               Call StrMenu1(False)
               Call StrMenu_Excel(False)
            End If
            
        ElseIf Index = 2 Then '電子化線上確認
            '檢查資料有無存在
            strSql = "SELECT count(*) FROM ABS013,Staff WHERE B1301='04' and B1302=ST01(+) " & m_StrSQL
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If RsTemp.Fields(0) > 0 Then
                  '提示訊息
                  strTit = "詢問"
                  strMsg = "有" & RsTemp.Fields(0) & "筆資料已存在,確定是否要重新產生資料?"
                  nResponse = MsgBox(strMsg, vbYesNo + vbCritical + vbQuestion, strTit)
                  If nResponse = vbNo Then Screen.MousePointer = vbDefault: Exit Sub
               End If
            End If
            Screen.MousePointer = vbHourglass
'            Call StrMenu1(True)
            Call StrMenu_Excel(True)
        End If
Case 1
        Unload Me
End Select
Screen.MousePointer = vbDefault
End Sub

Sub StartupExcel()
   
   '啟動Excel
   If cmdOK(0).Tag = "" Then
      cmdOK(0).Tag = "Excel" '要啟動Excel
      
      '預設A4紙張/橫式/比例 80%/水平置中/邊界左右都改0
      Set xlsAnnuity = New Excel.Application
      'xlsAnnuity.Visible = True
      xlsAnnuity.SheetsInNewWorkbook = 1 '預設工作表數量
      xlsAnnuity.Workbooks.add
      Set wksAnnuity = xlsAnnuity.Worksheets(1)
      xlsAnnuity.ActiveWindow.Zoom = 75 '畫面比例100%太大了,調整為75%
      '把Excel的警告訊息關掉
      xlsAnnuity.DisplayAlerts = False
      
      wksAnnuity.PageSetup.PaperSize = 9 'A4
      'wksAnnuity.PageSetup.Orientation = xlLandscape '橫印
      wksAnnuity.PageSetup.Orientation = wdOrientLandscape '直印
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
      wksAnnuity.Columns("A:A").ColumnWidth = 14
      wksAnnuity.Columns("B:B").ColumnWidth = 8
      wksAnnuity.Columns("C:C").ColumnWidth = 10
      wksAnnuity.Columns("D:D").ColumnWidth = 10
      wksAnnuity.Columns("E:E").ColumnWidth = 8
      wksAnnuity.Columns("F:F").ColumnWidth = 10
      wksAnnuity.Columns("G:G").ColumnWidth = 10
   
   '關閉Excel
   Else
      'xlsAnnuity.Selection.Cells.Select
      xlsAnnuity.Range("A1:" & "B" & m_intColumn).Select
      xlsAnnuity.Selection.RowHeight = 14.5 '列高
      strTempFile = App.path & "\$$" & strUserNum
      
   '      If Val(xlsAnnuity.Version) < 12 Then
   '         xlsAnnuity.Workbooks(1).SaveAs FileName:=strTempFile & ".xls", FileFormat:=-4143
   '      Else
   '         xlsAnnuity.Workbooks(1).SaveAs FileName:=strTempFile & ".xls", FileFormat:=56
   '      End If
      If intChoose <> 1 Then 'Add By Sindy 2011/9/2 intChoose:0.人事系統 1.出缺勤系統
         xlsAnnuity.Workbooks(1).PrintOut
      Else
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
         xlsAnnuity.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strTempFile & ".pdf", Quality:=0, _
         IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False
         'ShellExecute 0, "open", strTempFile, vbNullString, vbNullString, 1
      End If
      
      xlsAnnuity.Workbooks.Close 'SaveChanges:=False
      xlsAnnuity.Quit
      Set xlsAnnuity = Nothing
   End If
End Sub

'Add By Sindy 2022/1/21
Sub PrintTitle_Excel()
   '標題
   m_intColumn = m_intColumn + 1
   xlsAnnuity.Range("A" & m_intColumn).Value = "出缺勤統計通知單"
   xlsAnnuity.Range("A" & m_intColumn & ":" & "G" & m_intColumn).Select
   With xlsAnnuity.Selection
      .HorizontalAlignment = xlCenter '置中
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = True '合併儲存格
   End With
'   With xlsAnnuity.Selection.Font
'      .Bold = True '粗體
'      .Name = "新細明體"
'      .Size = 16
'   End With
   m_intColumn = m_intColumn + 1
   If Len(Trim(txt1(0))) = 4 Then
      xlsAnnuity.Range("A" & m_intColumn).Value = Left(Trim(txt1(0)), 2) & "  年  " & Right(Trim(txt1(0)), 2) & "  月"
   Else
      xlsAnnuity.Range("A" & m_intColumn).Value = Left(Trim(txt1(0)), 3) & "  年  " & Right(Trim(txt1(0)), 2) & "  月"
   End If
   xlsAnnuity.Range("A" & m_intColumn & ":" & "G" & m_intColumn).Select
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
   m_intColumn = m_intColumn + 1
   xlsAnnuity.Range("A" & m_intColumn).Value = "員工姓名：" & strTemp(1) & "　" & strTemp(2) & "　　部門：" & strTemp(3)
   xlsAnnuity.Range("A" & m_intColumn & ":" & "E" & m_intColumn).Select
   xlsAnnuity.Selection.MergeCells = True '合併儲存格
   xlsAnnuity.Range("F" & m_intColumn).Value = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   xlsAnnuity.Range("F" & m_intColumn & ":" & "G" & m_intColumn).Select
   xlsAnnuity.Selection.HorizontalAlignment = xlRight '靠右
   xlsAnnuity.Selection.MergeCells = True '合併儲存格
   
   m_intColumn = m_intColumn + 1
   xlsAnnuity.Range("A" & m_intColumn).Value = "項　　目"
   xlsAnnuity.Range("A" & m_intColumn).HorizontalAlignment = xlCenter
   xlsAnnuity.Range("A" & m_intColumn & ":" & "A" & m_intColumn).Select
   '下框線
   With xlsAnnuity.Selection.Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .ColorIndex = xlAutomatic
       .tintandshade = 0
       .Weight = xlThin
   End With
   xlsAnnuity.Range("C" & m_intColumn).Value = "當 月 統 計"
   xlsAnnuity.Range("C" & m_intColumn & ":" & "D" & m_intColumn).Select
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
   '下框線
   With xlsAnnuity.Selection.Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .ColorIndex = xlAutomatic
       .tintandshade = 0
       .Weight = xlThin
   End With
   xlsAnnuity.Range("F" & m_intColumn).Value = "年 度 累 計"
   xlsAnnuity.Range("F" & m_intColumn & ":" & "G" & m_intColumn).Select
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
   '下框線
   With xlsAnnuity.Selection.Borders(xlEdgeBottom)
       .LineStyle = xlContinuous
       .ColorIndex = xlAutomatic
       .tintandshade = 0
       .Weight = xlThin
   End With
'   '上框線
'   With xlsAnnuity.Selection.Borders(xlEdgeTop)
'       .LineStyle = xlContinuous
'       .ColorIndex = xlAutomatic
'       .tintandshade = 0
'       .Weight = xlThin
'   End With
'   xlsAnnuity.Range("C" & m_intColumn & ":G" & m_intColumn).Select
'   xlsAnnuity.Selection.HorizontalAlignment = xlRight '靠右
End Sub

Sub PrintDetail_Excel(StrST01 As String)
Dim m_j As Integer
Dim dblDay As Double, dblHour As Double
Dim dblRestDay As Double 'Add By Sindy 2017/11/3
Dim strText As String

For m_j = 1 To 23 '22
   If m_j = 1 Then
      strText = "忘 打 卡"
   ElseIf m_j = 2 Then strText = "遲      到"
   ElseIf m_j = 3 Then strText = "曠      職"
   ElseIf m_j = 4 Then strText = "事      假"
   ElseIf m_j = 5 Then strText = "家庭照顧假"
   ElseIf m_j = 6 Then strText = "防疫照顧假"
   ElseIf m_j = 7 Then strText = "病      假"
   ElseIf m_j = 8 Then strText = "生 理 假"
   ElseIf m_j = 9 Then
      GoTo goStep
      strText = "健 檢 假"
   ElseIf m_j = 10 Then strText = "公      假"
   ElseIf m_j = 11 Then strText = "特 別 假"
   ElseIf m_j = 12 Then strText = "出      差"
   ElseIf m_j = 13 Then strText = "加      班"
   ElseIf m_j = 14 Then strText = "婚      假"
   ElseIf m_j = 15 Then strText = "產 檢 假"
   ElseIf m_j = 16 Then strText = "產      假"
   ElseIf m_j = 17 Then strText = "流 產 假"
   ElseIf m_j = 18 Then strText = "陪 產 假"
   ElseIf m_j = 19 Then strText = "喪      假"
   ElseIf m_j = 20 Then strText = "公 傷 假"
   ElseIf m_j = 21 Then strText = "補      休"
   'Modify By Sindy 2025/11/17
   ElseIf m_j = 22 Then strText = "天災不給薪"
   '2025/11/17 END
   Else
      strText = "其      他"
   End If
   m_intColumn = m_intColumn + 1
   xlsAnnuity.Range("A" & m_intColumn).Value = strText
   
   '當月統計
   If m_j = 1 Or m_j = 2 Then '1.忘打卡 2.遲到
      xlsAnnuity.Range("D" & m_intColumn).Value = strTemp(m_j + 4) & " 次"
   'Modify By Sindy 2025/11/17 3.曠職
   ElseIf m_j = 3 Then
      xlsAnnuity.Range("D" & m_intColumn).Value = strTemp(m_j + 4) & " 分"
   '2025/11/17 END
   ElseIf m_j = 13 Then '加班
      xlsAnnuity.Range("C" & m_intColumn).Value = "0 日"
      xlsAnnuity.Range("D" & m_intColumn).Value = strTemp(m_j + 4) & " 時"
   Else
     ' ”\”整除運算子不可使用於有小數位之數值, 因此先將數值*10做運算
     'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
     'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
     'Modify By Sindy 2012/7/9 上班時數為特殊者
'         If strST01 = "99029" Then
'            dblDay = (strTemp(m_j + 4) * 10) \ (5 * 10)
'            dblHour = Round(strTemp(m_j + 4) - (dblDay * 5), 1)
      dblDay = (strTemp(m_j + 4) * 10) \ (PUB_intWkHour * 10)
      dblHour = Round(strTemp(m_j + 4) - (dblDay * PUB_intWkHour), 1)
      xlsAnnuity.Range("C" & m_intColumn).Value = dblDay & " 日"
      xlsAnnuity.Range("D" & m_intColumn).Value = dblHour & " 時"
   End If
   
   If PUB_bSpecY <> True Then 'Add By Sindy 2012/7/11 上班時數特殊者且為過渡期時不列印年度累計
      '年度累計
      If m_j = 1 Or m_j = 2 Then '1.忘打卡 2.遲到
         xlsAnnuity.Range("G" & m_intColumn).Value = strTemp(m_j + 29) & " 次"
      'Modify By Sindy 2025/11/17 3.曠職
      ElseIf m_j = 3 Then
         xlsAnnuity.Range("G" & m_intColumn).Value = strTemp(m_j + 29) & " 分"
      '2025/11/17 END
      ElseIf m_j = 13 Then '加班
         xlsAnnuity.Range("F" & m_intColumn).Value = "0 日"
         xlsAnnuity.Range("G" & m_intColumn).Value = strTemp(m_j + 29) & " 時"
      Else
        ' ”\”整除運算子不可使用於有小數位之數值, 因此先將數值*10做運算
        'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
        'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
        'Modify By Sindy 2012/7/9 上班時數為特殊者
'         If strST01 = "99029" Then
'            dblDay = (strTemp(m_j + 26) * 10) \ (5 * 10)
'            dblHour = Round(strTemp(m_j + 26) - (dblDay * 5), 1)
         dblDay = (strTemp(m_j + 29) * 10) \ (PUB_intWkHour * 10)
         dblHour = Round(strTemp(m_j + 29) - (dblDay * PUB_intWkHour), 1)
         xlsAnnuity.Range("F" & m_intColumn).Value = dblDay & " 日"
         xlsAnnuity.Range("G" & m_intColumn).Value = dblHour & " 時"
      End If
   End If
   
goStep:
Next m_j

m_intColumn = m_intColumn + 1
'Modify By Sindy 2017/11/3
If Val(txt1(0) & "01") + 19110000 < 20180101 Then
   xlsAnnuity.Range("A" & m_intColumn).Value = "本年度特別假天數　" & strTemp(4) & "　天"
Else
   'Modify By Sindy 2018/2/8
   dblRestDay = Fix(Val(strTemp(4)))
   If dblRestDay = 0 Then
      xlsAnnuity.Range("A" & m_intColumn).Value = "本年度特別假天數　" & dblRestDay & "　天"
   Else
   '2018/2/8 END
      xlsAnnuity.Range("A" & m_intColumn).Value = "本年度特別假天數　" & dblRestDay & "　天" & _
         IIf((strTemp(4) * PUB_intWkHour) - (dblRestDay * PUB_intWkHour), "　" & (strTemp(4) * PUB_intWkHour) - (dblRestDay * PUB_intWkHour) & "　小時", "")
   End If
End If
'2017/11/3 END

m_intColumn = m_intColumn + 1
'Add By Sindy 2011/9/2
If intChoose = 1 Then '出缺勤系統
'   iLine = iLine + 3
'   m_Device.CurrentX = 1500 * douExtRate
'   m_Device.CurrentY = (iLine * 230) * douExtRate
   xlsAnnuity.Range("A" & m_intColumn).Value = "備註：如資料有誤，請電洽人事處。"
'   m_iPages = m_iPages + 1
'   If m_iPages > 1 Then
'      SetPic m_iPages - 1
'   End If
   '2010/9/16 End
Else
   xlsAnnuity.Range("A" & m_intColumn).Value = "上級批示　　　　　人事主管　　　　　部門主管　　　　　閱後蓋章　　　　　"
   xlsAnnuity.Range("A" & m_intColumn & ":" & "G" & m_intColumn).Select
   xlsAnnuity.Selection.MergeCells = True '合併儲存格
End If
End Sub

'Add By Sindy 2022/1/21
Sub StrMenu_Excel(bolEPrint As Boolean)
Dim strSql As String
Dim strSDate As String, strEDate As String, strSYDate As String, strEYDate As String
'Dim dblHour(18) As Double, dblCnt(18) As Double
'Dim dblHour(22) As Double, dblCnt(22) As Double...移到basPerson共用變數區
Dim strST14 As String 'Add By Sindy 2011/9/2
Dim strSendEmailTo As String 'Add By Sindy 2011/9/2
Dim intECnt As Integer, intPCnt As Integer 'Add By Sindy 2011/9/15

strSDate = ChangeTStringToWString(txt1(0) * 100 + 1)
strEDate = ChangeTStringToWString(txt1(0) * 100 + 31)
strSYDate = Left(Trim(strSDate), 4) & "0101"
strEYDate = strEDate
cmdOK(0).Tag = "" 'Excel:要啟動Excel
m_intColumn = 0

'63001.董事長 67004.副董事長 L01.律師 不印出缺勤統計表
'Modify By Sindy 2011/11/7 增加判斷到職日及復職日
'm_str = "SELECT ST01,ST02,a1.a0901||' '||a1.a0902,ST40,ST14 " & _
'               "FROM Staff,Acc090 a1,SalaryData " & _
'               "WHERE (ST04='1' or (ST04='2' and ST51>=" & Val(txt1(0) & "01") + 19110000 & ")) and ST01=SD01 and (sd02 not in('P','F') or sd02 is null) " & _
'               "and ST03=a1.a0901(+) " & _
'               "and ST01 not in('63001','67004') and ST03<>'L01' " & m_StrSQL & _
'               "and " & Val(txt1(0) & "31") + 19110000 & ">=st13 " & _
'               "and (" & Val(txt1(0) & "31") + 19110000 & ">=(select max(sc02) from Staff_Change where sc01=ST01 and sc03='02') or 0=(select count(*) from Staff_Change where sc01=ST01 and sc03='02')) " & _
'               "ORDER BY ST03,ST01 ASC "
'Modify By Sindy 2014/10/9 +ABS001
'Modify By Sindy 2016/10/11 L01.法務處律師只有3個人不打卡,其他都要(and ST03<>'L01'取消,改寫法)
'Modify By Sindy 2021/6/4 + 排除 98099.江郁仁
'Modify By Sindy 2023/2/4 and substr(st01,4,1)<>'9'
'Modify By Sindy 2023/12/22 部門調整改抓ST93
m_str = "SELECT ST01,ST02,decode(st93,null,'(舊)'||a1.A0901||' '||a1.A0902,a0921||' '||a0922),YV04,ST14,B0108 " & _
               "FROM Staff,Acc090NEW,Acc090 a1,SalaryData,YearVacation,ABS001 " & _
               "WHERE (ST04='1' or (ST04='2' and ST51>=" & Val(txt1(0) & "01") + 19110000 & ")) and ST01=SD01 and (sd02 not in('P','F') or sd02 is null) " & _
               "and ST03=a1.a0901(+) and ST93=a0921(+) " & _
               "and ST01 not in('63001','67004','98099','" & Replace(Pub_GetSpecMan("不用打卡的律師"), ";", "','") & "') " & m_StrSQL & _
               "and " & Val(txt1(0) & "31") + 19110000 & ">=st13 " & _
               "and (" & Val(txt1(0) & "31") + 19110000 & ">=(select max(sc02) from Staff_Change where sc01=ST01 and sc03='02') or 0=(select count(*) from Staff_Change where sc01=ST01 and sc03='02')) " & _
               "and YV01(+)=" & Left(strSDate, 4) & " and YV02(+)=ST01 and ST01=B0101(+) " & _
               "and substr(st01,4,1)<>'9' " & _
               "ORDER BY ST93,ST01 ASC "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
   '設定使用者所選擇的印表機成預設印表機
   PUB_SetOsDefaultPrinter Combo1
   
   If bolEPrint = True Then '電子化線上確認
      On Error GoTo ErrHand
      cnnConnection.BeginTrans
      '刪除資料
      strSql = "DELETE FROM ABS013 WHERE B1301='04' and B1302 in(SELECT B1302 FROM ABS013,Staff WHERE B1301='04' and B1302=ST01(+) " & m_StrSQL & ")"
      cnnConnection.Execute strSql
   End If
   
   With m_rs
      m_rs.MoveFirst
      Do While Not m_rs.EOF
         'Add By Sindy 2011/9/2
         strST14 = ""
         If Not IsNull(m_rs.Fields("ST14")) Then strST14 = m_rs.Fields("ST14")
         '執行E化並且有公司信箱者,產生待確認資料
         'Modify By Sindy 2014/10/9 增加判斷若無審核主管1也是列印紙本不走電子化
         'If bolEPrint = True And strST14 <> "99997" Then
         'Modify By Sindy 2018/9/3 劉經理指示107/8月不出79017林慧汶的出缺勤資料
         'Modify By Sindy 2023/2/4 閻所長(81040)沒有簽核主管(因所長已是最大,沒有主管),但還是要走電子簽核
         If bolEPrint = True And strST14 <> "99997" And ("" & m_rs.Fields("B0108") <> "" Or m_rs.Fields("ST01") = "81040") _
            And Not (txt1(3) = "" And m_rs.Fields("ST01") = "79017" And Val(txt1(0)) + 191100 = 201808) Then
         '2014/10/9 END
            strSql = "insert into ABS013(B1301,B1302,B1303,B1314) " & _
                     "values('04'," & CNULL(CheckStr(m_rs.Fields("ST01"))) & "," & _
                     CNULL(CheckStr(m_rs.Fields("ST01"))) & "," & Val(txt1(0)) + 191100 & ")"
            cnnConnection.Execute strSql
            
            intECnt = intECnt + 1
            '記錄要發通知確認的E-Mail人員
            strSendEmailTo = strSendEmailTo & CheckStr(m_rs.Fields("ST01")) & ";"
            GoTo GoToNext
         End If
         '2011/9/2 End
         
         If cmdOK(0).Tag = "" Then Call StartupExcel '要啟動Excel
         If (intPCnt Mod 2 = 0) And intPCnt > 0 Then
            '換頁
            wksAnnuity.Range("A" & m_intColumn + 1).Select
            wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
         End If
         
         'Add By Sindy 2022/1/24
         If intChoose = 1 Then '出缺勤系統
            Load frmpic002
            frmpic002.Label1.Caption = "電子檔產生中...請稍候..."
            frmpic002.Show
            frmpic002.ZOrder 0
         End If
         '2022/1/24 END
         
         intPCnt = intPCnt + 1
         For m_i = 1 To 54 '52 '42
             strTemp(m_i) = ""
         Next m_i
         strTemp(1) = CheckStr(m_rs.Fields("ST01"))
         strTemp(2) = CheckStr(m_rs.Fields("ST02"))
         strTemp(3) = CheckStr(m_rs.Fields(2))
         'strTemp(4) = CheckStr(m_rs.Fields("ST40"))
         strTemp(4) = CheckStr(m_rs.Fields("YV04"))
         strSql = " and ST01='" & Trim(strTemp(1)) & "' "
         
         '取得各假別時數-當月統計
         If PUB_GetAbsenceHour(strSql, strSDate, strEDate, dblHour(), dblCnt()) = True Then
            '10.產假=產假+17.扣年終產假
            '11.流產假=流產假+18.扣年終流產假
            dblHour(10) = dblHour(10) + dblHour(17)
            dblHour(11) = dblHour(11) + dblHour(18)
            '填入變數
            strTemp(5) = dblHour(1)
            strTemp(6) = dblHour(2)
            strTemp(7) = dblHour(3)
            strTemp(8) = dblHour(5)   '05.事假
            strTemp(9) = dblHour(22)  'Add By Sindy 2014/12/9 22.家庭照顧假
            strTemp(10) = dblHour(24)  'Add By Sindy 2020/2/4 24.防疫照顧假
            strTemp(11) = dblHour(6)  '06.病假
            strTemp(12) = dblHour(20) 'Add By Sindy 2014/12/9 20.生理假
            strTemp(13) = dblHour(23) 'Add By Sindy 2015/1/5 23.健檢假
            strTemp(14) = dblHour(7)  '公假
            strTemp(15) = dblHour(8)
            strTemp(16) = dblHour(4)  '04.出差
            strTemp(17) = dblHour(16) '16.加班
            strTemp(18) = dblHour(9)  '09.婚假
            strTemp(19) = dblHour(21) 'Add By Sindy 2014/12/9 21.產檢假
            strTemp(20) = dblHour(10) '10.產假
            strTemp(21) = dblHour(11) '11.流產假
            strTemp(22) = dblHour(19) 'Add By Sindy 2012/1/4 19.陪產假
            strTemp(23) = dblHour(12) '喪假
            strTemp(24) = dblHour(13) '公傷假
            strTemp(25) = dblHour(14) '補休
            strTemp(26) = dblHour(25) 'Add By Sindy 2025/11/14 25.天災不給薪
            strTemp(27) = dblHour(15) '其他
            strTemp(28) = dblHour(17) '17.扣年終產假
            strTemp(29) = dblHour(18) '18.扣年終流產假
         End If
         '取得各假別時數-年度統計
         If PUB_GetAbsenceHour(strSql, strSYDate, strEYDate, dblHour(), dblCnt()) = True Then
            '10.產假=產假+17.扣年終產假
            '11.流產假=流產假+18.扣年終流產假
            dblHour(10) = dblHour(10) + dblHour(17)
            dblHour(11) = dblHour(11) + dblHour(18)
            '填入變數
            strTemp(30) = dblHour(1)
            strTemp(31) = dblHour(2)
            strTemp(32) = dblHour(3)
            strTemp(33) = dblHour(5)  '05.事假
            strTemp(34) = dblHour(22) 'Add By Sindy 2014/12/9 22.家庭照顧假
            strTemp(35) = dblHour(24) 'Add By Sindy 2020/2/4 24.防疫照顧假
            strTemp(36) = dblHour(6)  '06.病假
            strTemp(37) = dblHour(20) 'Add By Sindy 2014/12/9 20.生理假
            strTemp(38) = dblHour(23) 'Add By Sindy 2015/1/5 23.健檢假
            strTemp(39) = dblHour(7)  '公假
            strTemp(40) = dblHour(8)
            strTemp(41) = dblHour(4)  '04.出差
            strTemp(42) = dblHour(16) '16.加班
            strTemp(43) = dblHour(9)  '09.婚假
            strTemp(44) = dblHour(21) 'Add By Sindy 2014/12/9 21.產檢假
            strTemp(45) = dblHour(10) '10.產假
            strTemp(46) = dblHour(11) '11.流產假
            strTemp(47) = dblHour(19) 'Add By Sindy 2012/1/4 19.陪產假
            strTemp(48) = dblHour(12) '喪假
            strTemp(49) = dblHour(13) '公傷假
            strTemp(50) = dblHour(14) '補休
            strTemp(51) = dblHour(25) 'Add By Sindy 2025/11/14 25.天災不給薪
            strTemp(52) = dblHour(15) '其他
            strTemp(53) = dblHour(17) '17.扣年終產假
            strTemp(54) = dblHour(18) '18.扣年終流產假
         End If
         
         Call Pub_GetSpecWorkHour(strTemp(1), strSDate) 'Add By Sindy 2012/7/9 上班時數為特殊者
         If intPCnt Mod 2 = 0 Then
            m_intColumn = m_intColumn + 2
         End If
         PrintTitle_Excel '列印表頭
         Call PrintDetail_Excel(strTemp(1))  '列印表中、表尾
GoToNext:
         m_rs.MoveNext
      Loop
   End With
   
   If cmdOK(0).Tag = "Excel" Then Call StartupExcel '要關閉Excel
   
   PUB_SetOsDefaultPrinter strPrinter '復原系統預設印表機
Else
   Screen.MousePointer = vbDefault
   MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
   Exit Sub
End If

'Add By Sindy 2011/9/2
If intChoose = 1 Then '出缺勤系統
'   SetPic m_iPages
'   frm180203.m_ImageW = m_Device.Width
'   frm180203.m_ImageH = m_Device.Height
'   frm180203.m_iPages = m_iPages
   frm180203.Caption = "每月出缺勤統計確認"
   If Dir(strTempFile & ".pdf") <> "" Then
      frm180203.WebBrowser1.Navigate strTempFile & ".pdf"
   Else
      frm180203.WebBrowser1.Navigate "about:blank"
   End If
   Unload frmpic002
   'Me.Hide
'   If gShowMsg = True Then '簽核他人出缺勤資料時,第一筆資料要顯示警示訊息
'      MsgBox "簽核其他人每月出缺勤資料！", vbExclamation
'   End If
   frm180203.Show vbModal '強制回應表單
   Unload Me
Else
   If bolEPrint = True Then
      cnnConnection.CommitTrans
      '發通知確認的E-Mail
      If strSendEmailTo <> "" Then
         If Right(strSendEmailTo, 1) = ";" Then strSendEmailTo = Left(strSendEmailTo, Len(strSendEmailTo) - 1)
'         'Modify By Sindy 2013/3/5
'         'PUB_SendMail strUserNum, strSendEmailTo, "", "每月出缺勤統計待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行每月出缺勤統計確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
'         PUB_SendMail strUserNum, "taie_alluser@taie.com.tw", "", "每月出缺勤統計待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行每月出缺勤統計確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
'         '2013/3/5 End
         'Modify By Sindy 2014/3/5 若有輸入查詢條件時,依當事者各別發Mail
         If m_StrSQL <> "" Then
            PUB_SendMail strUserNum, strSendEmailTo, "", "每月出缺勤統計待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行每月出缺勤統計確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
         Else
            PUB_SendMail strUserNum, "taie_alluser@taie.com.tw", "", "每月出缺勤統計待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行每月出缺勤統計確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
         End If
         '2014/3/5 END
      End If
   End If
   Screen.MousePointer = vbDefault
   If intECnt = 0 Then
      MsgBox "列印紙本 " & intPCnt & " 筆完成!!", , "列印成功"
   Else
      MsgBox "列印紙本 " & intPCnt & " 筆及產生電子檔 " & intECnt & " 筆完成!!", , "列印成功"
   End If
End If

Exit Sub

ErrHand:
   If Err.Number <> 0 Then
      If bolEPrint = True Then '電子化線上確認
         cnnConnection.RollbackTrans
         MsgBox " 更新失敗！" & vbCrLf & Err.Description
      End If
   End If
'2011/9/2 End
End Sub

'Sub StrMenu1(bolEPrint As Boolean)
'Dim strSql As String
'Dim strSDate As String, strEDate As String, strSYDate As String, strEYDate As String
''Dim dblHour(18) As Double, dblCnt(18) As Double
''Dim dblHour(22) As Double, dblCnt(22) As Double...移到basPerson共用變數區
'Dim strST14 As String 'Add By Sindy 2011/9/2
'Dim strSendEmailTo As String 'Add By Sindy 2011/9/2
'Dim intECnt As Integer, intPCnt As Integer 'Add By Sindy 2011/9/15
'
'strSDate = ChangeTStringToWString(txt1(0) * 100 + 1)
'strEDate = ChangeTStringToWString(txt1(0) * 100 + 31)
'strSYDate = Left(Trim(strSDate), 4) & "0101"
'strEYDate = strEDate
'
'Set Printer = Printers(Combo1.ListIndex)
'
''XP自定紙張需手動設定並將印表機預設為該紙張
''9 x
''If pub_OS = "1" Then
''   m_Device.Height = 2880
''   m_Device.Width = 13000
''Else
''   m_Device.PaperSize = PUB_GetPaperSize(3)
''End If
''m_Device.Font = "@新細明體"
''m_Device.FontSize = 12
'
''Add By Sindy 2011/9/2
'm_iPages = 1
'If intChoose = 1 Then '出缺勤系統
'   Set m_Device = Picture1
'   m_Device.AutoRedraw = True
'   m_Device.Width = 9048 '11899
'   m_Device.Height = 5700 '6120 '7000 '16838
'   m_Device.AutoSize = True
'   douExtRate = m_Device.Height / 8142 '16836
'   DelPic
'Else
'   Set m_Device = Printer
'   m_Device.EndDoc
'   m_Device.Orientation = 1 '1.直印 2.橫印
'   'Modify By Sindy 2014/9/4 中一刀改A4列印
'   'm_Device.PaperSize = PUB_GetPaperSize(3)  '中一刀
'   douExtRate = 1
'   'm_Device.PaperSize = 39 '中一刀
'   m_Device.PaperSize = 9  'PDF
'End If
'
''63001.董事長 67004.副董事長 L01.律師 不印出缺勤統計表
''Modify By Sindy 2011/11/7 增加判斷到職日及復職日
''m_str = "SELECT ST01,ST02,a1.a0901||' '||a1.a0902,ST40,ST14 " & _
''               "FROM Staff,Acc090 a1,SalaryData " & _
''               "WHERE (ST04='1' or (ST04='2' and ST51>=" & Val(txt1(0) & "01") + 19110000 & ")) and ST01=SD01 and (sd02 not in('P','F') or sd02 is null) " & _
''               "and ST03=a1.a0901(+) " & _
''               "and ST01 not in('63001','67004') and ST03<>'L01' " & m_StrSQL & _
''               "and " & Val(txt1(0) & "31") + 19110000 & ">=st13 " & _
''               "and (" & Val(txt1(0) & "31") + 19110000 & ">=(select max(sc02) from Staff_Change where sc01=ST01 and sc03='02') or 0=(select count(*) from Staff_Change where sc01=ST01 and sc03='02')) " & _
''               "ORDER BY ST03,ST01 ASC "
''Modify By Sindy 2014/10/9 +ABS001
''Modify By Sindy 2016/10/11 L01.法務處律師只有3個人不打卡,其他都要(and ST03<>'L01'取消,改寫法)
''Modify By Sindy 2021/6/4 + 排除 98099.江郁仁
''Modify By Sindy 2023/2/4 and substr(st01,4,1)<>'9'
'm_str = "SELECT ST01,ST02,a1.a0921||' '||a1.a0922,YV04,ST14,B0108 " & _
'               "FROM Staff,Acc090NEW a1,SalaryData,YearVacation,ABS001 " & _
'               "WHERE (ST04='1' or (ST04='2' and ST51>=" & Val(txt1(0) & "01") + 19110000 & ")) and ST01=SD01 and (sd02 not in('P','F') or sd02 is null) " & _
'               "and ST93=a1.a0921(+) " & _
'               "and ST01 not in('63001','67004','98099','" & Replace(Pub_GetSpecMan("不用打卡的律師"), ";", "','") & "') " & m_StrSQL & _
'               "and " & Val(txt1(0) & "31") + 19110000 & ">=st13 " & _
'               "and (" & Val(txt1(0) & "31") + 19110000 & ">=(select max(sc02) from Staff_Change where sc01=ST01 and sc03='02') or 0=(select count(*) from Staff_Change where sc01=ST01 and sc03='02')) " & _
'               "and YV01(+)=" & Left(strSDate, 4) & " and YV02(+)=ST01 and ST01=B0101(+) " & _
'               "and substr(st01,4,1)<>'9' " & _
'               "ORDER BY ST93,ST01 ASC "
'If m_rs.State = 1 Then m_rs.Close
'm_rs.CursorLocation = adUseClient
'm_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
''LongPrintCurCnt = 0
'If Not m_rs.EOF And Not m_rs.BOF Then
'   If bolEPrint = True Then '電子化線上確認
'      On Error GoTo ErrHand
'      cnnConnection.BeginTrans
'      '刪除資料
'      strSql = "DELETE FROM ABS013 WHERE B1301='04' and B1302 in(SELECT B1302 FROM ABS013,Staff WHERE B1301='04' and B1302=ST01(+) " & m_StrSQL & ")"
'      cnnConnection.Execute strSql
'   End If
'
'    With m_rs
'        m_rs.MoveFirst
'        'PrintTitle '列印表頭
'        Do While Not m_rs.EOF
''            LongPrintCurCnt = LongPrintCurCnt + 1
'
'            'Add By Sindy 2011/9/2
'            strST14 = ""
'            If Not IsNull(m_rs.Fields("ST14")) Then strST14 = m_rs.Fields("ST14")
'            '執行E化並且有公司信箱者,產生待確認資料
'            'Modify By Sindy 2014/10/9 增加判斷若無審核主管1也是列印紙本不走電子化
'            'If bolEPrint = True And strST14 <> "99997" Then
'            'Modify By Sindy 2018/9/3 劉經理指示107/8月不出79017林慧汶的出缺勤資料
'            If bolEPrint = True And strST14 <> "99997" And "" & m_rs.Fields("B0108") <> "" _
'               And Not (txt1(3) = "" And m_rs.Fields("ST01") = "79017" And Val(txt1(0)) + 191100 = 201808) Then
'            '2014/10/9 END
'               strSql = "insert into ABS013(B1301,B1302,B1303,B1314) " & _
'                        "values('04'," & CNULL(CheckStr(m_rs.Fields("ST01"))) & "," & _
'                        CNULL(CheckStr(m_rs.Fields("ST01"))) & "," & Val(txt1(0)) + 191100 & ")"
'               cnnConnection.Execute strSql
'
'               intECnt = intECnt + 1
'               '記錄要發通知確認的E-Mail人員
'               strSendEmailTo = strSendEmailTo & CheckStr(m_rs.Fields("ST01")) & ";"
'               GoTo GoToNext
'            End If
'            '2011/9/2 End
'
'            intPCnt = intPCnt + 1
'            For m_i = 1 To 52 '42
'                strTemp(m_i) = ""
'            Next m_i
'            strTemp(1) = CheckStr(m_rs.Fields("ST01"))
'            strTemp(2) = CheckStr(m_rs.Fields("ST02"))
'            strTemp(3) = CheckStr(m_rs.Fields(2))
'            'strTemp(4) = CheckStr(m_rs.Fields("ST40"))
'            strTemp(4) = CheckStr(m_rs.Fields("YV04"))
'            strSql = " and ST01='" & Trim(strTemp(1)) & "' "
'
'            '取得各假別時數-當月統計
'            If PUB_GetAbsenceHour(strSql, strSDate, strEDate, dblHour(), dblCnt()) = True Then
'               '10.產假=產假+17.扣年終產假
'               '11.流產假=流產假+18.扣年終流產假
'               dblHour(10) = dblHour(10) + dblHour(17)
'               dblHour(11) = dblHour(11) + dblHour(18)
'               '填入變數
'               strTemp(5) = dblHour(1)
'               strTemp(6) = dblHour(2)
'               strTemp(7) = dblHour(3)
'               strTemp(8) = dblHour(5)   '05.事假
'               strTemp(9) = dblHour(22)  'Add By Sindy 2014/12/9 22.家庭照顧假
'               strTemp(10) = dblHour(24)  'Add By Sindy 2020/2/4 24.防疫照顧假
'               strTemp(11) = dblHour(6)  '06.病假
'               strTemp(12) = dblHour(20) 'Add By Sindy 2014/12/9 20.生理假
'               strTemp(13) = dblHour(23) 'Add By Sindy 2015/1/5 23.健檢假
'               strTemp(14) = dblHour(7)  '公假
'               strTemp(15) = dblHour(8)
'               strTemp(16) = dblHour(4)  '04.出差
'               strTemp(17) = dblHour(16) '16.加班
'               strTemp(18) = dblHour(9)  '09.婚假
'               strTemp(19) = dblHour(21) 'Add By Sindy 2014/12/9 21.產檢假
'               strTemp(20) = dblHour(10) '10.產假
'               strTemp(21) = dblHour(11) '11.流產假
'               strTemp(22) = dblHour(19) 'Add By Sindy 2012/1/4 19.陪產假
'               strTemp(23) = dblHour(12)
'               strTemp(24) = dblHour(13)
'               strTemp(25) = dblHour(14)
'               strTemp(26) = dblHour(15)
'               strTemp(27) = dblHour(17) '17.扣年終產假
'               strTemp(28) = dblHour(18) '18.扣年終流產假
'            End If
'            '取得各假別時數-年度統計
'            If PUB_GetAbsenceHour(strSql, strSYDate, strEYDate, dblHour(), dblCnt()) = True Then
'               '10.產假=產假+17.扣年終產假
'               '11.流產假=流產假+18.扣年終流產假
'               dblHour(10) = dblHour(10) + dblHour(17)
'               dblHour(11) = dblHour(11) + dblHour(18)
'               '填入變數
'               strTemp(29) = dblHour(1)
'               strTemp(30) = dblHour(2)
'               strTemp(31) = dblHour(3)
'               strTemp(32) = dblHour(5)  '05.事假
'               strTemp(33) = dblHour(22) 'Add By Sindy 2014/12/9 22.家庭照顧假
'               strTemp(34) = dblHour(24) 'Add By Sindy 2020/2/4 24.防疫照顧假
'               strTemp(35) = dblHour(6)  '06.病假
'               strTemp(36) = dblHour(20) 'Add By Sindy 2014/12/9 20.生理假
'               strTemp(37) = dblHour(23) 'Add By Sindy 2015/1/5 23.健檢假
'               strTemp(38) = dblHour(7)  '公假
'               strTemp(39) = dblHour(8)
'               strTemp(40) = dblHour(4)  '04.出差
'               strTemp(41) = dblHour(16) '16.加班
'               strTemp(42) = dblHour(9)  '09.婚假
'               strTemp(43) = dblHour(21) 'Add By Sindy 2014/12/9 21.產檢假
'               strTemp(44) = dblHour(10) '10.產假
'               strTemp(45) = dblHour(11) '11.流產假
'               strTemp(46) = dblHour(19) 'Add By Sindy 2012/1/4 19.陪產假
'               strTemp(47) = dblHour(12)
'               strTemp(48) = dblHour(13)
'               strTemp(49) = dblHour(14)
'               strTemp(50) = dblHour(15)
'               strTemp(51) = dblHour(17) '17.扣年終產假
'               strTemp(52) = dblHour(18) '18.扣年終流產假
'            End If
'
'            Call Pub_GetSpecWorkHour(strTemp(1), strSDate) 'Add By Sindy 2012/7/9 上班時數為特殊者
'            'Modify By Sindy 2014/9/5
'            If intPCnt Mod 2 = 0 Then
'               'Printer.NewPage
'               iLine = iLine + 4
'            Else
'               iLine = 1 '新頁重頭列印
'            End If
'            '2014/9/5 END
'            PrintTitle '列印表頭
'            Call PrintDetail(strTemp(1))  '列印表中、表尾
'            'Modify By Sindy 2014/9/5
'            If intPCnt Mod 2 = 0 Then
'               Printer.NewPage
'            End If
'            '2014/9/5 END
''            '每二筆才換新頁
''            If LongPrintCurCnt Mod 2 = 0 Then
''               m_Device.NewPage
''            End If
'
''            If iLine >= 30 Then
''                If .AbsolutePosition <> .RecordCount Then
''                    m_Device.NewPage
''                    'PrintTitle '列印表頭
''                End If
''            End If
'GoToNext:
'            m_rs.MoveNext
'        Loop
'    End With
'Else
'   MsgBox "無符合條件的資料!!!", vbExclamation + vbOKOnly
'   Exit Sub
'End If
''Add By Sindy 2011/9/2
'If intChoose = 1 Then '出缺勤系統
'   SetPic m_iPages
'   frm180203.m_ImageW = m_Device.Width
'   frm180203.m_ImageH = m_Device.Height
'   frm180203.m_iPages = m_iPages
'   frm180203.Caption = "每月出缺勤統計確認"
'   Me.Hide
''   If gShowMsg = True Then '簽核他人出缺勤資料時,第一筆資料要顯示警示訊息
''      MsgBox "簽核其他人每月出缺勤資料！", vbExclamation
''   End If
'   frm180203.Show vbModal '強制回應表單
'   Unload Me
'Else
'   If bolEPrint = True Then
'      cnnConnection.CommitTrans
'      '發通知確認的E-Mail
'      If strSendEmailTo <> "" Then
'         If Right(strSendEmailTo, 1) = ";" Then strSendEmailTo = Left(strSendEmailTo, Len(strSendEmailTo) - 1)
''         'Modify By Sindy 2013/3/5
''         'PUB_SendMail strUserNum, strSendEmailTo, "", "每月出缺勤統計待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行每月出缺勤統計確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
''         PUB_SendMail strUserNum, "taie_alluser@taie.com.tw", "", "每月出缺勤統計待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行每月出缺勤統計確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
''         '2013/3/5 End
'         'Modify By Sindy 2014/3/5 若有輸入查詢條件時,依當事者各別發Mail
'         If m_StrSQL <> "" Then
'            PUB_SendMail strUserNum, strSendEmailTo, "", "每月出缺勤統計待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行每月出缺勤統計確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
'         Else
'            PUB_SendMail strUserNum, "taie_alluser@taie.com.tw", "", "每月出缺勤統計待確認通知", "各位同仁：" & vbCrLf & "　　請至案件管理系統的一般作業\出缺勤作業\簽核項目中，進行每月出缺勤統計確認。" & vbCrLf & vbCrLf & vbCrLf & "　　　　　　　　　　　　　　　人事處", , , , , , , , , , True
'         End If
'         '2014/3/5 END
'      End If
'   End If
'   m_Device.EndDoc
'   'ShowPrintOk
'   If intECnt = 0 Then
'      MsgBox "列印紙本 " & intPCnt & " 筆完成!!", , "列印成功"
'   Else
'      MsgBox "列印紙本 " & intPCnt & " 筆及產生電子檔 " & intECnt & " 筆完成!!", , "列印成功"
'   End If
'End If
'
'Exit Sub
'
'ErrHand:
'   If bolEPrint = True Then '電子化線上確認
'      cnnConnection.RollbackTrans
'      MsgBox " 更新失敗！" & vbCrLf & Err.Description
'   End If
''2011/9/2 End
'End Sub

Sub PrintTitle()
GetPleft

PaperX = 12000
'paperY = 7500

'''列印行數
''If LongPrintCurCnt Mod 2 <> 0 Then
'   iLine = 1 '新頁重頭列印
''Else
''   iLine = iLine + 7 '接續列印
''End If

m_Device.Font.Size = 12 * douExtRate
m_Device.Font.Underline = False
m_Device.FontBold = False

'm_Device.CurrentX = m_Device.ScaleWidth / 2 - (m_Device.TextWidth("出缺勤統計通知單") / 2)
m_Device.CurrentX = (PaperX / 2 - (m_Device.TextWidth("出缺勤統計通知單") / 2)) * douExtRate
m_Device.CurrentY = (iLine * 230) * douExtRate
m_Device.Print "出缺勤統計通知單"
iLine = iLine + 2
If Len(Trim(txt1(0))) = 4 Then
   'm_Device.CurrentX = m_Device.ScaleWidth / 2 - (m_Device.TextWidth(Left(Trim(txt1(0)), 2) & "  年  " & Right(Trim(txt1(0)), 2) & "  月") / 2)
   m_Device.CurrentX = (PaperX / 2 - (m_Device.TextWidth(Left(Trim(txt1(0)), 2) & "  年  " & Right(Trim(txt1(0)), 2) & "  月") / 2)) * douExtRate
   m_Device.CurrentY = (iLine * 230) * douExtRate
   m_Device.Print Left(Trim(txt1(0)), 2) & "  年  " & Right(Trim(txt1(0)), 2) & "  月"
Else
   'm_Device.CurrentX = m_Device.ScaleWidth / 2 - (m_Device.TextWidth(Left(Trim(txt1(0)), 3) & "  年  " & Right(Trim(txt1(0)), 2) & "  月") / 2)
   m_Device.CurrentX = (PaperX / 2 - (m_Device.TextWidth(Left(Trim(txt1(0)), 3) & "  年  " & Right(Trim(txt1(0)), 2) & "  月") / 2)) * douExtRate
   m_Device.CurrentY = (iLine * 230) * douExtRate
   m_Device.Print Left(Trim(txt1(0)), 3) & "  年  " & Right(Trim(txt1(0)), 2) & "  月"
End If
'm_Device.CurrentX = m_Device.ScaleWidth - m_Device.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
m_Device.CurrentX = (PaperX - m_Device.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000) * douExtRate
m_Device.CurrentY = (iLine * 230) * douExtRate
m_Device.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))

iLine = iLine + 1
'm_Device.CurrentX = m_Device.ScaleWidth - m_Device.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 1000
'm_Device.CurrentY = iLine * 230
'm_Device.Print "頁　　次：" & m_Device.Page

iLine = iLine + 1
m_Device.CurrentX = 1500 * douExtRate
m_Device.CurrentY = (iLine * 230) * douExtRate
m_Device.Print "員工姓名：" & strTemp(1) & "　" & strTemp(2)
m_Device.CurrentX = 6500 * douExtRate
m_Device.CurrentY = (iLine * 230) * douExtRate
m_Device.Print "部門：" & strTemp(3)

iLine = iLine + 2
m_Device.CurrentX = PLeft(1) * douExtRate
m_Device.CurrentY = (iLine * 230) * douExtRate
m_Device.Print "項　　目"
m_Device.CurrentX = PLeft(2) * douExtRate
m_Device.CurrentY = (iLine * 230) * douExtRate
m_Device.Print "  當 月 統 計"
If PUB_bSpecY <> True Then 'Add By Sindy 2012/7/11 上班時數特殊者且為過渡期時不列印年度累計
   m_Device.CurrentX = PLeft(3) * douExtRate
   m_Device.CurrentY = (iLine * 230) * douExtRate
   m_Device.Print "  年 度 累 計"
End If

iLine = iLine + 1
m_Device.CurrentX = PLeft(4) * douExtRate
m_Device.CurrentY = (iLine * 230) * douExtRate
m_Device.Print String(20, "-")
m_Device.CurrentX = PLeft(5) * douExtRate
m_Device.CurrentY = (iLine * 230) * douExtRate
m_Device.Print String(30, "-")
If PUB_bSpecY <> True Then 'Add By Sindy 2012/7/11 上班時數特殊者且為過渡期時不列印年度累計
   m_Device.CurrentX = PLeft(6) * douExtRate
   m_Device.CurrentY = (iLine * 230) * douExtRate
   m_Device.Print String(30, "-")
End If

iLine = iLine + 1
End Sub

Sub GetPleft()
'Modify By Sindy 2014/9/4
''明細抬頭
'PLeft(1) = 2500
'PLeft(2) = 5500
'PLeft(3) = 8500
''明細(-----)
'PLeft(4) = 2200
'PLeft(5) = 5100
'PLeft(6) = 8100
''明細內文
'PLeft(7) = 6000
'PLeft(8) = 7000
'PLeft(9) = 9000
'PLeft(10) = 10000
'明細抬頭
PLeft(1) = 2000
PLeft(2) = 5000
PLeft(3) = 8000
'明細(-----)
PLeft(4) = 1700
PLeft(5) = 4600
PLeft(6) = 7600
'明細內文
PLeft(7) = 5500
PLeft(8) = 6500
PLeft(9) = 8500
PLeft(10) = 9500
'2014/9/4 END
End Sub

Sub PrintDetail(StrST01 As String)
Dim m_j As Integer
Dim dblDay As Double, dblHour As Double
Dim dblRestDay As Double 'Add By Sindy 2017/11/3
   
'For m_j = 1 To 16
For m_j = 1 To 22 '21 '20
    m_Device.CurrentX = PLeft(1) * douExtRate
    m_Device.CurrentY = (iLine * 230) * douExtRate
    If m_j = 1 Then
         m_Device.Print "忘 打 卡"
    ElseIf m_j = 2 Then
         m_Device.Print "遲      到"
    ElseIf m_j = 3 Then
         m_Device.Print "曠      職"
    ElseIf m_j = 4 Then
         m_Device.Print "事      假"
    'Add By Sindy 2014/12/9
    ElseIf m_j = 5 Then
         m_Device.Print "家庭照顧假"
    'Add By Sindy 2020/2/4
    ElseIf m_j = 6 Then
         m_Device.Print "防疫照顧假"
    ElseIf m_j = 7 Then
         m_Device.Print "病      假"
    'Add By Sindy 2014/12/9
    ElseIf m_j = 8 Then
         m_Device.Print "生 理 假"
    'Add By Sindy 2015/1/5
    ElseIf m_j = 9 Then
         GoTo goStep
         m_Device.Print "健 檢 假"
    ElseIf m_j = 10 Then
         m_Device.Print "公      假"
    ElseIf m_j = 11 Then
         m_Device.Print "特 別 假"
    ElseIf m_j = 12 Then
         m_Device.Print "出      差"
    ElseIf m_j = 13 Then
         m_Device.Print "加      班"
    ElseIf m_j = 14 Then
         m_Device.Print "婚      假"
    'Add By Sindy 2014/12/9
    ElseIf m_j = 15 Then
         m_Device.Print "產 檢 假"
    ElseIf m_j = 16 Then
         m_Device.Print "產      假"
    ElseIf m_j = 17 Then
         m_Device.Print "流 產 假"
    'Add By Sindy 2012/1/4
    ElseIf m_j = 18 Then
         m_Device.Print "陪 產 假"
    ElseIf m_j = 19 Then
         m_Device.Print "喪      假"
    ElseIf m_j = 20 Then
         m_Device.Print "公 傷 假"
    ElseIf m_j = 21 Then
         m_Device.Print "補      休"
    Else
         m_Device.Print "其      他"
    End If
    
    '當月統計
    If m_j = 1 Or m_j = 2 Then '1.忘打卡 2.遲到
         m_Device.CurrentX = (PLeft(8) - m_Device.TextWidth(strTemp(m_j + 4) & "次")) * douExtRate
         m_Device.CurrentY = (iLine * 230) * douExtRate
         m_Device.Print strTemp(m_j + 4) & "次"
    ElseIf m_j = 13 Then '加班
         m_Device.CurrentX = (PLeft(7) - m_Device.TextWidth("0 日")) * douExtRate
         m_Device.CurrentY = (iLine * 230) * douExtRate
         m_Device.Print "0 日"
         m_Device.CurrentX = (PLeft(8) - m_Device.TextWidth(strTemp(m_j + 4) & " 時")) * douExtRate
         m_Device.CurrentY = (iLine * 230) * douExtRate
         m_Device.Print strTemp(m_j + 4) & " 時"
    Else
         ' ”\”整除運算子不可使用於有小數位之數值, 因此先將數值*10做運算
         'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
         'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
         'Modify By Sindy 2012/7/9 上班時數為特殊者
'         If strST01 = "99029" Then
'            dblDay = (strTemp(m_j + 4) * 10) \ (5 * 10)
'            dblHour = Round(strTemp(m_j + 4) - (dblDay * 5), 1)
         dblDay = (strTemp(m_j + 4) * 10) \ (PUB_intWkHour * 10)
         dblHour = Round(strTemp(m_j + 4) - (dblDay * PUB_intWkHour), 1)
         m_Device.CurrentX = (PLeft(7) - m_Device.TextWidth(dblDay & " 日")) * douExtRate
         m_Device.CurrentY = (iLine * 230) * douExtRate
         m_Device.Print dblDay & " 日"
         m_Device.CurrentX = (PLeft(8) - m_Device.TextWidth(dblHour & " 時")) * douExtRate
         m_Device.CurrentY = (iLine * 230) * douExtRate
         m_Device.Print dblHour & " 時"
    End If
    
    If PUB_bSpecY <> True Then 'Add By Sindy 2012/7/11 上班時數特殊者且為過渡期時不列印年度累計
       '年度累計
       If m_j = 1 Or m_j = 2 Then '1.忘打卡 2.遲到
            m_Device.CurrentX = (PLeft(10) - m_Device.TextWidth(strTemp(m_j + 28) & "次")) * douExtRate
            m_Device.CurrentY = (iLine * 230) * douExtRate
            m_Device.Print strTemp(m_j + 28) & "次"
       ElseIf m_j = 13 Then '加班
            m_Device.CurrentX = (PLeft(9) - m_Device.TextWidth("0 日")) * douExtRate
            m_Device.CurrentY = (iLine * 230) * douExtRate
            m_Device.Print "0 日"
            m_Device.CurrentX = (PLeft(10) - m_Device.TextWidth(strTemp(m_j + 28) & " 時")) * douExtRate
            m_Device.CurrentY = (iLine * 230) * douExtRate
            m_Device.Print strTemp(m_j + 28) & " 時"
       Else
            ' ”\”整除運算子不可使用於有小數位之數值, 因此先將數值*10做運算
            'Modify By Sindy 2010/7/14 99029伊恩一天只上4個小時
            'Modify By Sindy 2011/3/8 99029伊恩一天只上5個小時
            'Modify By Sindy 2012/7/9 上班時數為特殊者
   '         If strST01 = "99029" Then
   '            dblDay = (strTemp(m_j + 26) * 10) \ (5 * 10)
   '            dblHour = Round(strTemp(m_j + 26) - (dblDay * 5), 1)
            dblDay = (strTemp(m_j + 28) * 10) \ (PUB_intWkHour * 10)
            dblHour = Round(strTemp(m_j + 28) - (dblDay * PUB_intWkHour), 1)
            m_Device.CurrentX = (PLeft(9) - m_Device.TextWidth(dblDay & " 日")) * douExtRate
            m_Device.CurrentY = (iLine * 230) * douExtRate
            m_Device.Print dblDay & " 日"
            m_Device.CurrentX = (PLeft(10) - m_Device.TextWidth(dblHour & " 時")) * douExtRate
            m_Device.CurrentY = (iLine * 230) * douExtRate
            m_Device.Print dblHour & " 時"
       End If
    End If
    
    iLine = iLine + 1
goStep:
Next m_j
m_Device.CurrentX = PLeft(1) * douExtRate
m_Device.CurrentY = (iLine * 230) * douExtRate
'Modify By Sindy 2017/11/3
If Val(txt1(0) & "01") + 19110000 < 20180101 Then
   m_Device.Print "本年度特別假天數　" & strTemp(4) & "　天"
Else
   'Modify By Sindy 2018/2/8
   dblRestDay = Fix(Val(strTemp(4)))
   If dblRestDay = 0 Then
      m_Device.Print "本年度特別假天數　" & dblRestDay & "　天"
   Else
   '2018/2/8 END
      m_Device.Print "本年度特別假天數　" & dblRestDay & "　天" & _
         IIf((strTemp(4) * PUB_intWkHour) - (dblRestDay * PUB_intWkHour), "　" & (strTemp(4) * PUB_intWkHour) - (dblRestDay * PUB_intWkHour) & "　小時", "")
   End If
End If
'2017/11/3 END
'Add By Sindy 2011/9/2
If intChoose = 1 Then '出缺勤系統
   iLine = iLine + 3
   m_Device.CurrentX = 1500 * douExtRate
   m_Device.CurrentY = (iLine * 230) * douExtRate
   m_Device.Print "備註：如資料有誤，請電洽人事處。"
   
   m_iPages = m_iPages + 1
   If m_iPages > 1 Then
      SetPic m_iPages - 1
   End If
   '2010/9/16 End
Else
   iLine = iLine + 2
   m_Device.CurrentX = 1500 * douExtRate
   m_Device.CurrentY = (iLine * 230) * douExtRate
   m_Device.Print "上級批示"
   m_Device.CurrentX = 4000 * douExtRate
   m_Device.CurrentY = (iLine * 230) * douExtRate
   m_Device.Print "人事主管"
   m_Device.CurrentX = 6500 * douExtRate
   m_Device.CurrentY = (iLine * 230) * douExtRate
   m_Device.Print "部門主管"
   m_Device.CurrentX = 9000 * douExtRate
   m_Device.CurrentY = (iLine * 230) * douExtRate
   m_Device.Print "閱後蓋章"
   
   iLine = iLine + 1
End If
End Sub

Private Sub Form_Load()
Dim SeekPrint As Integer, SeekPrintL As Integer
Dim strSql As String, i As Integer, j As Integer
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
'   strSystemKind = GetSystemKindByNick
'   strSql = Printer.DeviceName
'   SeekPrintL = Printer.Orientation
'   For i = 0 To Printers.Count - 1
'      Set Printer = Printers(i)
'      Combo1.AddItem Printer.DeviceName, j
'      j = j + 1
'      If Printer.DeviceName = strSql Then
'         SeekPrint = i
'      End If
'   Next i
'
'   Set Printer = Printers(SeekPrint)
'   Combo1.Text = Combo1.List(SeekPrint)
   
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2022/1/21
End Sub

Private Sub Form_Unload(Cancel As Integer)
'Dim strText As String
'
'   '一進入系統,檢查是否有須要開啟此作業
'   If pub_CallNextABSForm = True Then
'      strText = ChkIsAbsenceMustPro
'      If InStr(1, strText, "D") > 0 Then
'         'Set frm160102 = Nothing
'         frm160102.intChoose = 1
'         frm160102.Hide
'         Call frm160102.cmdOK_Click(0)
'      End If
'      If InStr(1, strText, "D") = Len(strText) Then
'         pub_CallNextABSForm = False
'      End If
'   End If
   
   Set frm160201 = Nothing
'   If pub_CallNextABSForm = False Then
'      Call Forms(0).SysStartCallForm 'Add By Sindy 2011/10/7
'   End If
End Sub

'Add By Sindy 2011/9/2
Private Sub DelPic()
   Dim strPicFileName As String
   strPicFileName = App.path & "\$tmp_*.tmp"
   If Dir(strPicFileName) <> "" Then
      Kill strPicFileName
   End If
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
End Sub

'Add By Sindy 2011/9/2
Private Sub SetPic(idx As Integer)
Dim strPicFileName As String
   
   strPicFileName = App.path & "\$tmp_" & idx & ".tmp"
   SavePicture Picture1.Image, strPicFileName
   '要用覆蓋的否則會錯誤--VB Bug
   m_Device.Line (0, 0)-(m_Device.Width, m_Device.Height), QBColor(15), BF
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 1, 2, 3, 4
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "01") = False Then
                Call txt1_GotFocus(Index)
                Cancel = True
                Exit Sub
            End If
         End If
      Case 1, 2
         If Index = 1 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 2 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case 3, 4
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 3 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 4 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
      Case Else
   End Select
End Sub
