VERSION 5.00
Begin VB.Form frm160109 
   BorderStyle     =   1  '單線固定
   Caption         =   "晉升、真除列印"
   ClientHeight    =   3050
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3050
   ScaleWidth      =   4950
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2640
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1320
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2760
      MaxLength       =   7
      TabIndex        =   1
      Top             =   960
      Width           =   735
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1950
      MaxLength       =   7
      TabIndex        =   0
      Top             =   960
      Width           =   735
   End
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   30
      TabIndex        =   7
      Top             =   2400
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   4
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   8
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   2
      Left            =   1950
      MaxLength       =   6
      TabIndex        =   2
      Top             =   1320
      Width           =   615
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3915
      TabIndex        =   6
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2970
      TabIndex        =   5
      Top             =   60
      Width           =   915
   End
   Begin VB.Line Line1 
      X1              =   2550
      X2              =   2910
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line2 
      X1              =   2580
      X2              =   2940
      Y1              =   1080
      Y2              =   1080
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "異動日期："
      Height          =   180
      Left            =   1020
      TabIndex        =   10
      Top             =   1020
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   1020
      TabIndex        =   9
      Top             =   1380
      Width           =   900
   End
End
Attribute VB_Name = "frm160109"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/25 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
'Create by Sindy 2009/01/13
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_str2 As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 10) As Integer
Dim strTemp(1 To 10) As String
Dim iPgae As Integer, iLine As Integer
Dim strType As String
'Add By Sindy 2022/1/25
Dim strPrinter As String
Dim xlsAnnuity As New Excel.Application
Dim wksAnnuity As New Worksheet
Dim m_intColumn As Integer
'2022/1/25 END


Private Sub cmdok_Click(Index As Integer)
Select Case Index
Case 0
        If txt1(0) = "" And txt1(1) = "" And txt1(2) = "" And txt1(3) = "" Then
            MsgBox "請至少輸入一項列印條件！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        If (txt1(0) = "" And txt1(1) <> "") Then
            MsgBox "起始日期不可空白！", vbInformation, "操作錯誤！"
            txt1(0).SetFocus
            Exit Sub
        End If
        If (txt1(0) <> "" And txt1(1) = "") Then
            MsgBox "迄止日期不可空白！", vbInformation, "操作錯誤！"
            txt1(1).SetFocus
            Exit Sub
        End If
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            m_StrSQL = m_StrSQL & " and sc02>='" & DBDATE(txt1(0)) & "' "
        End If
        If txt1(1) <> "" Then
            m_StrSQL = m_StrSQL & " and sc02<='" & DBDATE(txt1(1)) & "' "
        End If
        If txt1(2) <> "" Then
            m_StrSQL = m_StrSQL & " and sc01>='" & txt1(2) & "' "
        End If
        If txt1(3) <> "" Then
            m_StrSQL = m_StrSQL & " and sc01<='" & txt1(3) & "' "
        End If
        'StrMenu1
        StrMenu_Excel
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
Case Else
End Select
End Sub

'Add By Sindy 2022/1/25
Sub StrMenu_Excel()

cmdok(0).Tag = "" 'Excel:要啟動Excel
m_intColumn = 0
'Modify By Sindy 2023/12/28 部門調整改抓ST93
m_str = "SELECT SC01,SC02,nvl(A0922,'(舊)'||a1.A0902) a0902,ST02,a3.ac03,a2.ac03,decode(SC03,'05','晉升','06','真除',SC03),SC04 " & _
             "FROM staff_change,staff,acc090 a1,acc090NEW,allcode a2,allcode a3 " & _
             "WHERE SC01=ST01(+) " & _
             "and SC04=a1.a0901(+) and SC04=a0921(+) and SC02<20240101 " & _
             "and '01'=a2.ac01(+) " & _
             "and SC05=a2.ac02(+) " & _
             "and '02'=a3.ac01(+) " & _
             "and SC06=a3.ac02(+) " & _
             "and SC03 in ('05','06') " & m_StrSQL
m_str = m_str & " union " & _
             "SELECT SC01,SC02,a0922 a0902,ST02,a3.ac03,a2.ac03,decode(SC03,'05','晉升','06','真除',SC03),SC04 " & _
             "FROM staff_change,staff,acc090NEW,allcode a2,allcode a3 " & _
             "WHERE SC01=ST01(+) " & _
             "and SC04=a0921(+) and SC02>=20240101 " & _
             "and '01'=a2.ac01(+) " & _
             "and SC05=a2.ac02(+) " & _
             "and '02'=a3.ac01(+) " & _
             "and SC06=a3.ac02(+) " & _
             "and SC03 in ('05','06') " & m_StrSQL
m_str = m_str & " order by SC04,SC01,SC02 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
   '設定使用者所選擇的印表機成預設印表機
   PUB_SetOsDefaultPrinter Combo1
   
   With m_rs
      .MoveFirst
      If cmdok(0).Tag = "" Then Call StartupExcel '要啟動Excel
      PrintTitle_Excel
      strType = ""
      
      Do While Not .EOF
         For m_i = 1 To 10
            strTemp(m_i) = ""
         Next m_i
         
         strTemp(1) = CheckStr(m_rs.Fields(0)) '員工編號
         strTemp(2) = CheckStr(m_rs.Fields(1)) '異動日期
         strTemp(3) = CheckStr(m_rs.Fields(2)) '部門
         strTemp(4) = CheckStr(m_rs.Fields(3)) '姓名
         
         ' 查詢此筆資料的前一筆異動資料
         m_str2 = "SELECT a3.ac03,a2.ac03 " & _
                        "FROM staff_change,allcode a2,allcode a3 " & _
                        "WHERE SC01='" & strTemp(1) & "' " & _
                        "and SC02 = " & _
                        "(SELECT max(SC02) FROM staff_change " & _
                        "WHERE SC01='" & strTemp(1) & "' and SC02<'" & strTemp(2) & "') " & _
                        "and '01'=a2.ac01(+) and SC05=a2.ac02(+) " & _
                        "and '02'=a3.ac01(+) and SC06=a3.ac02(+) "
         If m_rs2.State = 1 Then m_rs2.Close
         m_rs2.CursorLocation = adUseClient
         m_rs2.Open m_str2, cnnConnection, adOpenStatic, adLockReadOnly
         If Not m_rs2.EOF And Not m_rs2.BOF Then
            strTemp(5) = CheckStr(m_rs2.Fields(0)) '原職位
            strTemp(6) = CheckStr(m_rs2.Fields(1)) '原職稱
         End If
         
         strTemp(7) = CheckStr(m_rs.Fields(4)) '現職
         strTemp(8) = CheckStr(m_rs.Fields(5)) '現職稱
         strTemp(9) = CheckStr(m_rs.Fields(6)) '原因
         
'           If iLine > 34 Or iLine = 1 Then
'              'If .AbsolutePosition <> .RecordCount Then
'                  If strType <> "" Then Printer.NewPage
'                  iLine = 1
'                  PrintTitle '列印表頭
'              'End If
'           End If
         
         PrintDetail_Excel
         
         strType = strTemp(1)
         .MoveNext
      Loop
   End With
   
   If cmdok(0).Tag = "Excel" Then Call StartupExcel '要關閉Excel
   PUB_SetOsDefaultPrinter strPrinter '復原系統預設印表機
Else
   ShowNoData
   Exit Sub
End If

ShowPrintOk
End Sub

Sub StartupExcel()
   
   '啟動Excel
   If cmdok(0).Tag = "" Then
      cmdok(0).Tag = "Excel" '要啟動Excel
      
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
      wksAnnuity.Columns("A:A").ColumnWidth = 15
      wksAnnuity.Columns("B:B").ColumnWidth = 15
      wksAnnuity.Columns("C:C").ColumnWidth = 25
      wksAnnuity.Columns("D:D").ColumnWidth = 15
      wksAnnuity.Columns("E:E").ColumnWidth = 25
      wksAnnuity.Columns("F:F").ColumnWidth = 15
      wksAnnuity.Columns("G:G").ColumnWidth = 15
   
   '關閉Excel
   Else
      xlsAnnuity.ActiveSheet.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
      'Modify By Sindy 2022/1/19 列印標題
      xlsAnnuity.ActiveSheet.PageSetup.PrintTitleRows = "$1:$4"
      
      'xlsAnnuity.Selection.Cells.Select
'      xlsAnnuity.Range("A1:" & "G" & m_intColumn).Select
'      xlsAnnuity.Selection.RowHeight = 14.5 '列高
      
      xlsAnnuity.Workbooks(1).PrintOut
      
      xlsAnnuity.Workbooks.Close 'SaveChanges:=False
      xlsAnnuity.Quit
      Set xlsAnnuity = Nothing
   End If
End Sub

'Add By Sindy 2022/1/25
Sub PrintTitle_Excel()
   
   '標題
   m_intColumn = m_intColumn + 1
   xlsAnnuity.Range("A" & m_intColumn).Value = ChangeTStringToTDateString(txt1(0)) & " 起晉升、真除同仁名單"
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
   xlsAnnuity.Range("A" & m_intColumn).Value = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   xlsAnnuity.Range("A" & m_intColumn & ":" & "G" & m_intColumn).Select
   With xlsAnnuity.Selection
      .HorizontalAlignment = xlRight '靠右
      .VerticalAlignment = xlCenter
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .IndentLevel = 0
      .ShrinkToFit = False
      .ReadingOrder = xlContext
      .MergeCells = True
   End With
   m_intColumn = m_intColumn + 2
   xlsAnnuity.Range("A" & m_intColumn).Value = "部　門"
   xlsAnnuity.Range("B" & m_intColumn).Value = "姓　名"
   xlsAnnuity.Range("C" & m_intColumn).Value = "原　職　位"
   xlsAnnuity.Range("D" & m_intColumn).Value = "原　職　稱"
   xlsAnnuity.Range("E" & m_intColumn).Value = "現　職"
   xlsAnnuity.Range("F" & m_intColumn).Value = "現　職　稱"
   xlsAnnuity.Range("G" & m_intColumn).Value = "事由"
   xlsAnnuity.Range("A" & m_intColumn & ":" & "G" & m_intColumn).Select
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

'Add By Sindy 2022/1/25
Sub PrintDetail_Excel()
Dim m_j As Integer
Dim strField As String

m_intColumn = m_intColumn + 1
For m_j = 1 To 7
   strField = GetFieldStr(m_j, 64)
   xlsAnnuity.Range(strField & m_intColumn).Value = strTemp(m_j + 2)
Next m_j
End Sub

Sub StrMenu1()
Set Printer = Printers(Combo1.ListIndex)
Printer.EndDoc
Printer.Orientation = 2 '1.直印 2.橫印
'Printer.PaperSize = 9  'PDF

m_str = "SELECT SC01,SC02,a1.a0902,ST02,a3.ac03,a2.ac03,decode(SC03,'05','晉升','06','真除',SC03) " & _
             "FROM staff_change,staff,acc090 a1,allcode a2,allcode a3 " & _
             "WHERE SC01=ST01(+) " & _
             "and SC04=a1.a0901(+) " & _
             "and '01'=a2.ac01(+) " & _
             "and SC05=a2.ac02(+) " & _
             "and '02'=a3.ac01(+) " & _
             "and SC06=a3.ac02(+) " & _
             "and SC03 in ('05','06') " & m_StrSQL & _
             "order by SC04,SC01,SC02 "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        .MoveFirst
        
        iLine = 1
        strType = ""
        
        Do While Not .EOF
            For m_i = 1 To 10
                strTemp(m_i) = ""
            Next m_i
            
            strTemp(1) = CheckStr(m_rs.Fields(0)) '員工編號
            strTemp(2) = CheckStr(m_rs.Fields(1)) '異動日期
            strTemp(3) = CheckStr(m_rs.Fields(2)) '部門
            strTemp(4) = CheckStr(m_rs.Fields(3)) '姓名
            
            ' 查詢此筆資料的前一筆異動資料
            m_str2 = "SELECT a3.ac03,a2.ac03 " & _
                           "FROM staff_change,allcode a2,allcode a3 " & _
                           "WHERE SC01='" & strTemp(1) & "' " & _
                           "and SC02 = " & _
                           "(SELECT max(SC02) FROM staff_change " & _
                           "WHERE SC01='" & strTemp(1) & "' and SC02<'" & strTemp(2) & "') " & _
                           "and '01'=a2.ac01(+) and SC05=a2.ac02(+) " & _
                           "and '02'=a3.ac01(+) and SC06=a3.ac02(+) "
            If m_rs2.State = 1 Then m_rs2.Close
            m_rs2.CursorLocation = adUseClient
            m_rs2.Open m_str2, cnnConnection, adOpenStatic, adLockReadOnly
            If Not m_rs2.EOF And Not m_rs2.BOF Then
               strTemp(5) = CheckStr(m_rs2.Fields(0)) '原職位
               strTemp(6) = CheckStr(m_rs2.Fields(1)) '原職稱
            End If
            
            strTemp(7) = CheckStr(m_rs.Fields(4)) '現職
            strTemp(8) = CheckStr(m_rs.Fields(5)) '現職稱
            strTemp(9) = CheckStr(m_rs.Fields(6)) '原因
            
            If iLine > 34 Or iLine = 1 Then
               'If .AbsolutePosition <> .RecordCount Then
                   If strType <> "" Then Printer.NewPage
                   iLine = 1
                   PrintTitle '列印表頭
               'End If
            End If
            PrintDetail
            
            strType = strTemp(1)
            .MoveNext
        Loop
    End With
Else
    ShowNoData
    Exit Sub
End If
Printer.EndDoc
ShowPrintOk
End Sub

Sub GetPleft()
PLeft(1) = 500
PLeft(2) = 2500
PLeft(3) = 4000
PLeft(4) = 6000
PLeft(5) = 9500
PLeft(6) = 11500
PLeft(7) = 14500
End Sub

Sub PrintTitle()
GetPleft

Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False

Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(ChangeTStringToTDateString(txt1(0)) & " 起晉升、真除同仁名單") / 2)
Printer.CurrentY = iLine * 300
Printer.Print ChangeTStringToTDateString(txt1(0)) & " 起晉升、真除同仁名單"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "部　門"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "姓　名"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "原　職　位"
Printer.CurrentX = PLeft(4)
Printer.CurrentY = iLine * 300
Printer.Print "原　職　稱"
Printer.CurrentX = PLeft(5)
Printer.CurrentY = iLine * 300
Printer.Print "現　職"
Printer.CurrentX = PLeft(6)
Printer.CurrentY = iLine * 300
Printer.Print "現　職　稱"
Printer.CurrentX = PLeft(7)
Printer.CurrentY = iLine * 300
Printer.Print "事由"
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(205, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 7
   Printer.CurrentX = PLeft(m_j)
   Printer.CurrentY = iLine * 300
   Printer.Print strTemp(m_j + 2)
Next m_j
iLine = iLine + 1
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

   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2022/1/25
End Sub

Private Sub Form_Unload(Cancel As Integer)
Set frm160109 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0, 1
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If txt1(Index).Text <> "" Then
            If ChkDate(txt1(Index)) = False Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
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
         ' 檢查員工編號規則
         If txt1(Index).Text <> "" Then
            If ChkStaffID(txt1(Index)) Then
               Call txt1_GotFocus(Index)
               Cancel = True
               Exit Sub
            End If
         End If
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
      Case Else
   End Select
End Sub
