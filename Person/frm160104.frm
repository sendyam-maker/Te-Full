VERSION 5.00
Begin VB.Form frm160104 
   BorderStyle     =   1  '單線固定
   Caption         =   "勞、健、團保費名單列印"
   ClientHeight    =   3020
   ClientLeft      =   50
   ClientTop       =   330
   ClientWidth     =   4950
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3020
   ScaleWidth      =   4950
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   60
      TabIndex        =   12
      Top             =   2400
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   5
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   13
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "列印(&P)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   3000
      TabIndex        =   6
      Top             =   60
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   3945
      TabIndex        =   7
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   0
      Left            =   2040
      MaxLength       =   3
      TabIndex        =   1
      Top             =   1170
      Width           =   465
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2580
      MaxLength       =   3
      TabIndex        =   2
      Top             =   1170
      Width           =   465
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   2
      Left            =   2040
      MaxLength       =   6
      TabIndex        =   3
      Top             =   1530
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   3
      Left            =   2730
      MaxLength       =   6
      TabIndex        =   4
      Top             =   1530
      Width           =   615
   End
   Begin VB.TextBox txt1 
      Height          =   255
      Index           =   4
      Left            =   2040
      MaxLength       =   1
      TabIndex        =   0
      Top             =   810
      Width           =   195
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "部門代號："
      Height          =   180
      Left            =   1080
      TabIndex        =   11
      Top             =   1200
      Width           =   900
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "員工編號："
      Height          =   180
      Index           =   0
      Left            =   1080
      TabIndex        =   10
      Top             =   1560
      Width           =   900
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "名單類別："
      Height          =   180
      Left            =   1080
      TabIndex        =   9
      Top             =   840
      Width           =   900
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "(1.勞保 2.健保 3.團保)"
      Height          =   180
      Left            =   2310
      TabIndex        =   8
      Top             =   840
      Width           =   1695
   End
   Begin VB.Line Line1 
      X1              =   2280
      X2              =   3270
      Y1              =   1680
      Y2              =   1680
   End
   Begin VB.Line Line2 
      X1              =   2340
      X2              =   2700
      Y1              =   1290
      Y2              =   1290
   End
End
Attribute VB_Name = "frm160104"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/1/25 Form2.0已修改
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/17 SQLDate已檢查
'Memo By Sindy 2010/11/25 員工編號欄已修改
'Modify By Sindy 2010/7/21 日期欄已修改
'create by nickc 2008/03/26
Option Explicit

Dim m_rs As New ADODB.Recordset
Dim m_rs2 As New ADODB.Recordset
Dim m_str As String
Dim m_str2 As String
Dim m_StrSQL As String
Dim m_i As Integer
Dim PLeft(1 To 6) As Integer
Dim strTemp(1 To 6) As String
Dim strTempS(1 To 6) As String
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
        If txt1(4) = "" Then
            MsgBox "名單類別不可以空白！", vbInformation, "操作錯誤！"
            txt1(4).SetFocus
            Exit Sub
        End If
        
        Set Printer = Printers(Combo1.ListIndex)
        
        Screen.MousePointer = vbHourglass
        m_StrSQL = ""
        If txt1(0) <> "" Then
            m_StrSQL = m_StrSQL & " and st03>='" & txt1(0) & "' "
        End If
        If txt1(1) <> "" Then
            m_StrSQL = m_StrSQL & " and st03<='" & txt1(1) & "' "
        End If
        If txt1(2) <> "" Then
            m_StrSQL = m_StrSQL & " and st01>='" & txt1(2) & "' "
        End If
        If txt1(3) <> "" Then
            m_StrSQL = m_StrSQL & " and st01<='" & txt1(3) & "' "
        End If
        'm_StrSQL = m_StrSQL & " and st04='1' and (st31 is not null or st31>0) "
        m_StrSQL = m_StrSQL & " and st04='1' "
        'StrMenu Val(txt1(4))
        StrMenu_Excel Val(txt1(4))
        Screen.MousePointer = vbDefault
Case 1
        Unload Me
Case Else
End Select
End Sub

'Add By Sindy 2022/1/25
Sub StrMenu_Excel(oInt As Integer)
Dim strCompany As String 'Add By Sindy 2014/4/25

cmdok(0).Tag = "" 'Excel:要啟動Excel
m_intColumn = 0

'2009/1/16 modify by sonia 剔除F編號,且有薪資基本資料者(QPGMR)
'm_str = "select * from staff,acc090 where st03=a0901(+) " & m_StrSQL & " order by st03,st01 "
'Modify By Sindy 2014/4/25 +依公司別跳頁
'm_str = "select * from staff,acc090,salarydata where st01<'F' and st03=a0901(+) and st01=sd01(+) and sd01 is not null " & m_StrSQL & " order by st03,st01 "
'Modify By Sindy 2023/12/28 部門調整改抓ST93
m_str = "select * from staff,acc090NEW,salarydata,acc080  where st01<'F' and st93=a0921(+) and st01=sd01(+) and sd01 is not null and sd19=a0801(+) " & m_StrSQL & " order by sd19,st03,st01 "
'2009/1/16 end
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
   '設定使用者所選擇的印表機成預設印表機
   PUB_SetOsDefaultPrinter Combo1
   
   With m_rs
      .MoveFirst
      
      If cmdok(0).Tag = "" Then Call StartupExcel '要啟動Excel
      iLine = 0
      strType = ""
      For m_i = 1 To 6
          strTempS(m_i) = ""
      Next m_i
      
      Do While Not .EOF
         'Add By Sindy 2014/4/25
         If strCompany <> .Fields("A0802") Then
            iLine = 1
            '換頁
            If strCompany <> "" Then
               wksAnnuity.Range("A" & m_intColumn + 1).Select
               wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
            End If
            'PrintTitle oInt, .Fields("A0802") '列印表頭
            PrintTitle_Excel oInt, .Fields("A0802") '列印表頭
         Else
            iLine = iLine + 1
         End If
         '2014/4/25 END
         For m_i = 1 To 6
             strTemp(m_i) = ""
         Next m_i
         'Modify By Sindy 2023/12/28 部門調整改抓ST93
         strTemp(1) = CheckStr(.Fields("a0923"))
         '2023/12/28 END
         strTemp(2) = CheckStr(.Fields("st01"))
         strTemp(3) = CheckStr(.Fields("st02"))
         strCompany = "" & .Fields("A0802") 'Add By Sindy 2014/4/25
         Select Case oInt
         Case 1
                 strTemp(4) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st23"))))
                 strTemp(5) = CheckStr(.Fields("st26"))
                 strTemp(6) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st31"))))
         Case 2
                 strTemp(4) = CheckStr(.Fields("st26"))
         Case 3
                 strTemp(4) = CheckStr(.Fields("st26"))
                 strTemp(5) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st23"))))
                 strTemp(6) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st13"))))
         End Select
         If oInt = 2 Then
            '眷屬資料
            '2009/1/16 modify by sonia 只抓眷保眷屬
            'm_str2 = "select * from staff_relation where sr01='" & CheckStr(.Fields("st01")) & "' order by sr03 "
            'Modify by Morgan 2009/6/29
            'm_str2 = "select * from staff_relation where sr08 is null and sr01='" & CheckStr(.Fields("st01")) & "' order by sr03 "
            m_str2 = "select * from staff_relation where sr08='Y' and sr01='" & CheckStr(.Fields("st01")) & "' order by sr03 "
            If m_rs2.State = 1 Then m_rs2.Close
            m_rs2.CursorLocation = adUseClient
            m_rs2.Open m_str2, cnnConnection, adOpenStatic, adLockReadOnly
            If Not m_rs2.EOF And Not m_rs2.BOF Then
                m_rs2.MoveFirst
                iLine = iLine - 1 '前面先加了,扣回來
                Do While Not m_rs2.EOF
                    iLine = iLine + 1 '明細資料列 + 1
                    strTemp(5) = CheckStr(m_rs2.Fields("sr04"))
                    strTemp(6) = CheckStr(m_rs2.Fields("sr07"))
                    If strTemp(1) = strTempS(1) Or strTemp(1) = "" Then
                        strTemp(1) = ""
                        If strTemp(2) = strTempS(2) Or strTemp(2) = "" Then
                            strTemp(2) = ""
                            If strTemp(3) = strTempS(3) Or strTemp(3) = "" Then
                                strTemp(3) = ""
                                If strTemp(4) = strTempS(4) Or strTemp(4) = "" Then
                                    strTemp(4) = ""
                                    If strTemp(5) = strTempS(5) Then
                                        strTemp(5) = ""
                                        If strTemp(6) = strTempS(6) Then
                                            strTemp(6) = ""
                                        Else
                                            strTempS(6) = strTemp(6)
                                        End If
                                    Else
                                        strTempS(5) = strTemp(5)
                                        strTempS(6) = strTemp(6)
                                    End If
                                Else
                                    strTempS(4) = strTemp(4)
                                    strTempS(5) = strTemp(5)
                                    strTempS(6) = strTemp(6)
                                End If
                            Else
                                strTempS(3) = strTemp(3)
                                strTempS(4) = strTemp(4)
                                strTempS(5) = strTemp(5)
                                strTempS(6) = strTemp(6)
                            End If
                        Else
                            strTempS(2) = strTemp(2)
                            strTempS(3) = strTemp(3)
                            strTempS(4) = strTemp(4)
                            strTempS(5) = strTemp(5)
                            strTempS(6) = strTemp(6)
                        End If
                    Else
                        strTempS(1) = strTemp(1)
                        strTempS(2) = strTemp(2)
                        strTempS(3) = strTemp(3)
                        strTempS(4) = strTemp(4)
                        strTempS(5) = strTemp(5)
                        strTempS(6) = strTemp(6)
                    End If
                    If (iLine Mod 45 = 0) Then
                       '換頁
                       wksAnnuity.Range("A" & m_intColumn + 1).Select
                       wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
                       PrintTitle_Excel oInt, strCompany '列印表頭
                    End If
                    'PrintDetail
                    PrintDetail_Excel
                    m_rs2.MoveNext
                Loop
            Else
                If strTemp(1) = strTempS(1) Then
                    strTemp(1) = ""
                    If strTemp(2) = strTempS(2) Then
                        strTemp(2) = ""
                        If strTemp(3) = strTempS(3) Then
                            strTemp(3) = ""
                            If strTemp(4) = strTempS(4) Then
                                strTemp(4) = ""
                                If strTemp(5) = strTempS(5) Then
                                    strTemp(5) = ""
                                    If strTemp(6) = strTempS(6) Then
                                        strTemp(6) = ""
                                    Else
                                        strTempS(6) = strTemp(6)
                                    End If
                                Else
                                    strTempS(5) = strTemp(5)
                                    strTempS(6) = strTemp(6)
                                End If
                            Else
                                strTempS(4) = strTemp(4)
                                strTempS(5) = strTemp(5)
                                strTempS(6) = strTemp(6)
                            End If
                        Else
                            strTempS(3) = strTemp(3)
                            strTempS(4) = strTemp(4)
                            strTempS(5) = strTemp(5)
                            strTempS(6) = strTemp(6)
                        End If
                    Else
                        strTempS(2) = strTemp(2)
                        strTempS(3) = strTemp(3)
                        strTempS(4) = strTemp(4)
                        strTempS(5) = strTemp(5)
                        strTempS(6) = strTemp(6)
                    End If
                Else
                    strTempS(1) = strTemp(1)
                    strTempS(2) = strTemp(2)
                    strTempS(3) = strTemp(3)
                    strTempS(4) = strTemp(4)
                    strTempS(5) = strTemp(5)
                    strTempS(6) = strTemp(6)
                End If
                If (iLine Mod 45 = 0) Then
                     '換頁
                     wksAnnuity.Range("A" & m_intColumn + 1).Select
                     wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
                     PrintTitle_Excel oInt, strCompany '列印表頭
                End If
                'PrintDetail
                PrintDetail_Excel
            End If
         Else
            If strTemp(1) = strTempS(1) Then
                strTemp(1) = ""
                If strTemp(2) = strTempS(2) Then
                    strTemp(2) = ""
                    If strTemp(3) = strTempS(3) Then
                        strTemp(3) = ""
                        If strTemp(4) = strTempS(4) Then
                            strTemp(4) = ""
                            If strTemp(5) = strTempS(5) Then
                                strTemp(5) = ""
                                If strTemp(6) = strTempS(6) Then
                                    strTemp(6) = ""
                                Else
                                    strTempS(6) = strTemp(6)
                                End If
                            Else
                                strTempS(5) = strTemp(5)
                                strTempS(6) = strTemp(6)
                            End If
                        Else
                            strTempS(4) = strTemp(4)
                            strTempS(5) = strTemp(5)
                            strTempS(6) = strTemp(6)
                        End If
                    Else
                        strTempS(3) = strTemp(3)
                        strTempS(4) = strTemp(4)
                        strTempS(5) = strTemp(5)
                        strTempS(6) = strTemp(6)
                    End If
                Else
                    strTempS(2) = strTemp(2)
                    strTempS(3) = strTemp(3)
                    strTempS(4) = strTemp(4)
                    strTempS(5) = strTemp(5)
                    strTempS(6) = strTemp(6)
                End If
            Else
                strTempS(1) = strTemp(1)
                strTempS(2) = strTemp(2)
                strTempS(3) = strTemp(3)
                strTempS(4) = strTemp(4)
                strTempS(5) = strTemp(5)
                strTempS(6) = strTemp(6)
            End If
            If (iLine Mod 45 = 0) Then
               '換頁
               wksAnnuity.Range("A" & m_intColumn + 1).Select
               wksAnnuity.HPageBreaks.add Before:=wksAnnuity.Application.ActiveCell
               PrintTitle_Excel oInt, strCompany '列印表頭
            End If
            'PrintDetail
            PrintDetail_Excel
         End If
         strType = CheckStr(m_rs.Fields(0))
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

'Add By Sindy 2022/1/25
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
      wksAnnuity.Columns("A:A").ColumnWidth = 25
      wksAnnuity.Columns("B:B").ColumnWidth = 11
      wksAnnuity.Columns("C:C").ColumnWidth = 11
      wksAnnuity.Columns("D:D").ColumnWidth = 15
      wksAnnuity.Columns("E:E").ColumnWidth = 12
      wksAnnuity.Columns("F:F").ColumnWidth = 12
   
   '關閉Excel
   Else
      xlsAnnuity.ActiveSheet.PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
'      'Modify By Sindy 2022/1/19 列印標題
'      xlsAnnuity.ActiveSheet.PageSetup.PrintTitleRows = "$1:$4"
      
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
Sub PrintTitle_Excel(oInt As Integer, strCompany As String)
Dim oStr As String
   
   Select Case oInt
      Case 1
         oStr = "勞　保　名　單"
      Case 2
         oStr = "健　保　名　單"
      Case 3
         oStr = "團　保　名　單"
   End Select
   '標題
   m_intColumn = m_intColumn + 1
   xlsAnnuity.Range("A" & m_intColumn).Value = oStr
   xlsAnnuity.Range("A" & m_intColumn & ":" & "F" & m_intColumn).Select
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
   xlsAnnuity.Range("A" & m_intColumn).Value = "單位名稱：" & strCompany
   xlsAnnuity.Range("C" & m_intColumn).Value = "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
   xlsAnnuity.Range("C" & m_intColumn & ":" & "F" & m_intColumn).Select
   xlsAnnuity.Selection.MergeCells = True '合併儲存格
   xlsAnnuity.Selection.HorizontalAlignment = xlRight '靠右
   
   m_intColumn = m_intColumn + 2
   xlsAnnuity.Range("A" & m_intColumn).Value = "部　門　名　稱"
   xlsAnnuity.Range("B" & m_intColumn).Value = "員工編號"
   xlsAnnuity.Range("C" & m_intColumn).Value = "姓　名"
   Select Case oInt
      Case 1
           xlsAnnuity.Range("D" & m_intColumn).Value = "出生日期"
           xlsAnnuity.Range("E" & m_intColumn).Value = "身分證字號"
           xlsAnnuity.Range("F" & m_intColumn).Value = "加保日期"
      Case 2
           xlsAnnuity.Range("D" & m_intColumn).Value = "身分證字號"
           xlsAnnuity.Range("E" & m_intColumn).Value = "眷屬姓名"
           xlsAnnuity.Range("F" & m_intColumn).Value = "身分證字號"
      Case 3
           xlsAnnuity.Range("D" & m_intColumn).Value = "身分證字號"
           xlsAnnuity.Range("E" & m_intColumn).Value = "出生日期"
           xlsAnnuity.Range("F" & m_intColumn).Value = "到職日期"
   End Select
   xlsAnnuity.Range("A" & m_intColumn & ":" & "F" & m_intColumn).Select
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
For m_j = 1 To 6
   strField = GetFieldStr(m_j, 64)
   xlsAnnuity.Range(strField & m_intColumn).Value = strTemp(m_j)
   If m_j = 2 Then
      xlsAnnuity.Range("B" & m_intColumn & ":" & "B" & m_intColumn).Select
      xlsAnnuity.Selection.HorizontalAlignment = xlLeft '靠左
   End If
Next m_j
End Sub

Sub StrMenu(oInt As Integer)
Dim strCompany As String 'Add By Sindy 2014/4/25

Printer.Orientation = 1
'Printer.FontName = "標楷體"

'2009/1/16 modify by sonia 剔除F編號,且有薪資基本資料者(QPGMR)
'm_str = "select * from staff,acc090 where st03=a0901(+) " & m_StrSQL & " order by st03,st01 "
'Modify By Sindy 2014/4/25 +依公司別跳頁
'm_str = "select * from staff,acc090,salarydata where st01<'F' and st03=a0901(+) and st01=sd01(+) and sd01 is not null " & m_StrSQL & " order by st03,st01 "
'Modify By Sindy 2023/12/28 部門調整改抓ST93
m_str = "select * from staff,acc090NEW,salarydata,acc080  where st01<'F' and st93=a0921(+) and st01=sd01(+) and sd01 is not null and sd19=a0801(+) " & m_StrSQL & " order by sd19,st03,st01 "
'2009/1/16 end
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    With m_rs
        .MoveFirst
        
        iLine = 1
        strType = ""
        For m_i = 1 To 6
            strTempS(m_i) = ""
        Next m_i
        
        Do While Not .EOF
            'Add By Sindy 2014/4/25
            If strCompany <> "" And strCompany <> .Fields("A0802") Then
               Printer.NewPage
               PrintTitle oInt, .Fields("A0802") '列印表頭
            End If
            '2014/4/25 END
            For m_i = 1 To 6
                strTemp(m_i) = ""
            Next m_i
            'Modify By Sindy 2023/12/28 部門調整改抓ST93
            strTemp(1) = CheckStr(.Fields("a0923"))
            '2023/12/28 END
            strTemp(2) = CheckStr(.Fields("st01"))
            strTemp(3) = CheckStr(.Fields("st02"))
            strCompany = "" & .Fields("A0802") 'Add By Sindy 2014/4/25
            Select Case oInt
            Case 1
                    strTemp(4) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st23"))))
                    strTemp(5) = CheckStr(.Fields("st26"))
                    strTemp(6) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st31"))))
            Case 2
                    strTemp(4) = CheckStr(.Fields("st26"))
            Case 3
                    strTemp(4) = CheckStr(.Fields("st26"))
                    strTemp(5) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st23"))))
                    strTemp(6) = ChangeTStringToTDateString(TAIWANDATE(CheckStr(.Fields("st13"))))
            End Select
            If oInt = 2 Then
                '眷屬資料
                '2009/1/16 modify by sonia 只抓眷保眷屬
                'm_str2 = "select * from staff_relation where sr01='" & CheckStr(.Fields("st01")) & "' order by sr03 "
                'Modify by Morgan 2009/6/29
                'm_str2 = "select * from staff_relation where sr08 is null and sr01='" & CheckStr(.Fields("st01")) & "' order by sr03 "
                m_str2 = "select * from staff_relation where sr08='Y' and sr01='" & CheckStr(.Fields("st01")) & "' order by sr03 "
                If m_rs2.State = 1 Then m_rs2.Close
                m_rs2.CursorLocation = adUseClient
                m_rs2.Open m_str2, cnnConnection, adOpenStatic, adLockReadOnly
                If Not m_rs2.EOF And Not m_rs2.BOF Then
                    m_rs2.MoveFirst
                    Do While Not m_rs2.EOF
                        strTemp(5) = CheckStr(m_rs2.Fields("sr04"))
                        strTemp(6) = CheckStr(m_rs2.Fields("sr07"))
                        If strTemp(1) = strTempS(1) Or strTemp(1) = "" Then
                            strTemp(1) = ""
                            If strTemp(2) = strTempS(2) Or strTemp(2) = "" Then
                                strTemp(2) = ""
                                If strTemp(3) = strTempS(3) Or strTemp(3) = "" Then
                                    strTemp(3) = ""
                                    If strTemp(4) = strTempS(4) Or strTemp(4) = "" Then
                                        strTemp(4) = ""
                                        If strTemp(5) = strTempS(5) Then
                                            strTemp(5) = ""
                                            If strTemp(6) = strTempS(6) Then
                                                strTemp(6) = ""
                                            Else
                                                strTempS(6) = strTemp(6)
                                            End If
                                        Else
                                            strTempS(5) = strTemp(5)
                                            strTempS(6) = strTemp(6)
                                        End If
                                    Else
                                        strTempS(4) = strTemp(4)
                                        strTempS(5) = strTemp(5)
                                        strTempS(6) = strTemp(6)
                                    End If
                                Else
                                    strTempS(3) = strTemp(3)
                                    strTempS(4) = strTemp(4)
                                    strTempS(5) = strTemp(5)
                                    strTempS(6) = strTemp(6)
                                End If
                            Else
                                strTempS(2) = strTemp(2)
                                strTempS(3) = strTemp(3)
                                strTempS(4) = strTemp(4)
                                strTempS(5) = strTemp(5)
                                strTempS(6) = strTemp(6)
                            End If
                        Else
                            strTempS(1) = strTemp(1)
                            strTempS(2) = strTemp(2)
                            strTempS(3) = strTemp(3)
                            strTempS(4) = strTemp(4)
                            strTempS(5) = strTemp(5)
                            strTempS(6) = strTemp(6)
                        End If
                        If iLine > 48 Or iLine = 1 Then
                            'If .AbsolutePosition <> .RecordCount Then
                                If strType <> "" Then Printer.NewPage
                                PrintTitle oInt, strCompany '列印表頭
                            'End If
                        End If
                        PrintDetail
'                        If iLine >= 52 Then
'                            If .AbsolutePosition <> .RecordCount Then
'                                Printer.NewPage
'                                PrintTitle oInt
'                            End If
'                        End If
                        m_rs2.MoveNext
                    Loop
                Else
                    If strTemp(1) = strTempS(1) Then
                        strTemp(1) = ""
                        If strTemp(2) = strTempS(2) Then
                            strTemp(2) = ""
                            If strTemp(3) = strTempS(3) Then
                                strTemp(3) = ""
                                If strTemp(4) = strTempS(4) Then
                                    strTemp(4) = ""
                                    If strTemp(5) = strTempS(5) Then
                                        strTemp(5) = ""
                                        If strTemp(6) = strTempS(6) Then
                                            strTemp(6) = ""
                                        Else
                                            strTempS(6) = strTemp(6)
                                        End If
                                    Else
                                        strTempS(5) = strTemp(5)
                                        strTempS(6) = strTemp(6)
                                    End If
                                Else
                                    strTempS(4) = strTemp(4)
                                    strTempS(5) = strTemp(5)
                                    strTempS(6) = strTemp(6)
                                End If
                            Else
                                strTempS(3) = strTemp(3)
                                strTempS(4) = strTemp(4)
                                strTempS(5) = strTemp(5)
                                strTempS(6) = strTemp(6)
                            End If
                        Else
                            strTempS(2) = strTemp(2)
                            strTempS(3) = strTemp(3)
                            strTempS(4) = strTemp(4)
                            strTempS(5) = strTemp(5)
                            strTempS(6) = strTemp(6)
                        End If
                    Else
                        strTempS(1) = strTemp(1)
                        strTempS(2) = strTemp(2)
                        strTempS(3) = strTemp(3)
                        strTempS(4) = strTemp(4)
                        strTempS(5) = strTemp(5)
                        strTempS(6) = strTemp(6)
                    End If
                    If iLine > 48 Or iLine = 1 Then
                        'If .AbsolutePosition <> .RecordCount Then
                            If strType <> "" Then Printer.NewPage
                            PrintTitle oInt, strCompany '列印表頭
                        'End If
                    End If
                    PrintDetail
                End If
            Else
                If strTemp(1) = strTempS(1) Then
                    strTemp(1) = ""
                    If strTemp(2) = strTempS(2) Then
                        strTemp(2) = ""
                        If strTemp(3) = strTempS(3) Then
                            strTemp(3) = ""
                            If strTemp(4) = strTempS(4) Then
                                strTemp(4) = ""
                                If strTemp(5) = strTempS(5) Then
                                    strTemp(5) = ""
                                    If strTemp(6) = strTempS(6) Then
                                        strTemp(6) = ""
                                    Else
                                        strTempS(6) = strTemp(6)
                                    End If
                                Else
                                    strTempS(5) = strTemp(5)
                                    strTempS(6) = strTemp(6)
                                End If
                            Else
                                strTempS(4) = strTemp(4)
                                strTempS(5) = strTemp(5)
                                strTempS(6) = strTemp(6)
                            End If
                        Else
                            strTempS(3) = strTemp(3)
                            strTempS(4) = strTemp(4)
                            strTempS(5) = strTemp(5)
                            strTempS(6) = strTemp(6)
                        End If
                    Else
                        strTempS(2) = strTemp(2)
                        strTempS(3) = strTemp(3)
                        strTempS(4) = strTemp(4)
                        strTempS(5) = strTemp(5)
                        strTempS(6) = strTemp(6)
                    End If
                Else
                    strTempS(1) = strTemp(1)
                    strTempS(2) = strTemp(2)
                    strTempS(3) = strTemp(3)
                    strTempS(4) = strTemp(4)
                    strTempS(5) = strTemp(5)
                    strTempS(6) = strTemp(6)
                End If
                If iLine > 48 Or iLine = 1 Then
                   'If .AbsolutePosition <> .RecordCount Then
                       If strType <> "" Then Printer.NewPage
                       PrintTitle oInt, strCompany '列印表頭
                   'End If
                End If
                PrintDetail
            End If
            strType = CheckStr(m_rs.Fields(0))
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
PLeft(1) = 300
PLeft(2) = (Printer.ScaleWidth / 8) * 2
PLeft(3) = (Printer.ScaleWidth / 8) * 3
PLeft(4) = (Printer.ScaleWidth / 8) * 4
PLeft(5) = (Printer.ScaleWidth / 8) * 5
PLeft(6) = (Printer.ScaleWidth / 8) * 6
End Sub

Sub PrintTitle(oInt As Integer, strCompany As String)
Dim oStr As String
Select Case oInt
Case 1
        oStr = "勞　保　名　單"
Case 2
        oStr = "健　保　名　單"
Case 3
        oStr = "團　保　名　單"
End Select
GetPleft
Printer.Font.Size = 12
Printer.Font.Underline = False
Printer.FontBold = False
Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth(oStr) / 2)
Printer.CurrentY = 300
Printer.Print oStr
'Printer.Font.Size = 10
'Printer.Font.Underline = False
'Printer.FontBold = False
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 600
Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
Printer.CurrentX = PLeft(1)
Printer.CurrentY = 900
Printer.Print "單位名稱：" & strCompany 'Modify By Sindy 2014/4/25 "台一國際專利商標事務所"
Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
Printer.CurrentY = 900
Printer.Print "頁　　次：" & Printer.Page
iLine = 5
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print "部　門　名　稱"
Printer.CurrentX = PLeft(2)
Printer.CurrentY = iLine * 300
Printer.Print "員工編號"
Printer.CurrentX = PLeft(3)
Printer.CurrentY = iLine * 300
Printer.Print "姓　名"
Select Case oInt
Case 1
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = iLine * 300
        Printer.Print "出生日期"
        Printer.CurrentX = PLeft(5)
        Printer.CurrentY = iLine * 300
        Printer.Print "身分證字號"
        Printer.CurrentX = PLeft(6)
        Printer.CurrentY = iLine * 300
        Printer.Print "加保日期"
Case 2
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = iLine * 300
        Printer.Print "身分證字號"
        Printer.CurrentX = PLeft(5)
        Printer.CurrentY = iLine * 300
        Printer.Print "眷屬姓名"
        Printer.CurrentX = PLeft(6)
        Printer.CurrentY = iLine * 300
        Printer.Print "身分證字號"
Case 3
        Printer.CurrentX = PLeft(4)
        Printer.CurrentY = iLine * 300
        Printer.Print "身分證字號"
        Printer.CurrentX = PLeft(5)
        Printer.CurrentY = iLine * 300
        Printer.Print "出生日期"
        Printer.CurrentX = PLeft(6)
        Printer.CurrentY = iLine * 300
        Printer.Print "到職日期"
End Select
iLine = iLine + 1
Printer.CurrentX = PLeft(1)
Printer.CurrentY = iLine * 300
Printer.Print String(140, "-")
iLine = iLine + 1
End Sub

Sub PrintDetail()
Dim m_j As Integer
For m_j = 1 To 6
    Printer.CurrentX = PLeft(m_j)
    Printer.CurrentY = iLine * 300
    Printer.Print strTemp(m_j)
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
Set frm160104 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
    InverseTextBox txt1(Index)
    CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 4
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case 0, 1, 2, 3
         KeyAscii = UpperCase(KeyAscii)
      Case Else
   End Select
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         ElseIf Index = 1 Then
            If RunNick(txt1(Index - 1), txt1(Index)) Then
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
      Case 4
        If txt1(4) <> "" Then
            Select Case txt1(4)
            Case "1", "2", "3"
            Case Else
                MsgBox "報表類別只可以輸入 1 或 2 或 3！", vbInformation, "輸入錯誤！"
                Cancel = True
            End Select
        End If
      Case Else
   End Select
End Sub
