VERSION 5.00
Begin VB.Form frm170240 
   BorderStyle     =   1  '單線固定
   Caption         =   "年度特殊功績獎金"
   ClientHeight    =   3360
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5328
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   5328
   Begin VB.Frame Frame1 
      Caption         =   "設定"
      Height          =   600
      Left            =   180
      TabIndex        =   4
      Top             =   2700
      Visible         =   0   'False
      Width           =   4875
      Begin VB.ComboBox Combo1 
         Height          =   300
         Left            =   705
         Style           =   2  '單純下拉式
         TabIndex        =   3
         Top             =   180
         Width           =   3870
      End
      Begin VB.Label Label2 
         Caption         =   "印表機"
         Height          =   315
         Index           =   1
         Left            =   75
         TabIndex        =   5
         Top             =   240
         Width           =   765
      End
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   2340
      MaxLength       =   3
      TabIndex        =   0
      Top             =   1230
      Width           =   585
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "離開(&X)"
      Height          =   435
      Index           =   1
      Left            =   4215
      TabIndex        =   2
      Top             =   120
      Width           =   915
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "產生Excel檔(&E)"
      Default         =   -1  'True
      Height          =   435
      Index           =   0
      Left            =   2700
      TabIndex        =   1
      Top             =   120
      Width           =   1455
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "年度："
      Height          =   180
      Left            =   1770
      TabIndex        =   6
      Top             =   1290
      Width           =   540
   End
End
Attribute VB_Name = "frm170240"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sindy 2022/2/17 Form2.0已檢查 (無需修改的物件)
'Create By Sindy 2020/11/18
Option Explicit

Dim xlsSalesPoint As New Excel.Application
Dim wksaccrpt114 As New Worksheet
Dim i As Integer, j As Integer
Dim strFileName As String
Dim intSheets As Integer, strSheetsVal(1 To 10) As String


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         If txt1(0) = "" Then
            MsgBox "年度不可空白！", vbInformation, "操作錯誤！"
            If txt1(0) = "" Then txt1(0).SetFocus
            Exit Sub
         End If
         
         strExc(0) = "select count(*) from yearbonus where yb01=" & Val(txt1(0)) + 1911
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            If RsTemp.Fields(0) = 0 Then
               MsgBox "該年度(" & Val(txt1(0)) & ")尚無年終獎金資料！", vbExclamation + vbOKOnly
            End If
         End If
         
         Call Pub_ChkExcelPath 'Added by Lydia 2021/07/01 檢查xls資料夾的模組
         
         Screen.MousePointer = vbHourglass
         Call ExcelSave
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

'*************************************************
'轉成Excel檔案
'
'*************************************************
Private Sub ExcelSave()
Dim adoRs As New ADODB.Recordset
Dim adoRs2 As New ADODB.Recordset
Dim intRow As Integer, strSort As String
Dim strSheetsName As String

On Error GoTo flgErr
   
   strFileName = strExcelPath & Val(txt1(0)) & "年度特殊功績獎金.xls"
   If Dir(strFileName) <> MsgText(601) Then
      Kill strFileName
   End If
   
   For j = 1 To 10
      strSheetsVal(j) = ""
   Next j
   strExc(0) = "SELECT sd19" & _
               " FROM staff,salarydata" & _
               " WHERE st01=sd01(+) AND st03<>'P29' AND st01>'63' AND st01<'F' AND substr(st01,4,1)<>'9'" & _
               " AND ((st04='1' and st13<=" & Val(txt1(0)) + 1911 & "1231) or (st04<>'1' and st13<=" & Val(txt1(0)) + 1911 & "1231 and st51>" & Val(txt1(0)) + 1911 & "1231))" & _
               " AND sd19 is not null AND sd19<>'1'" & _
               " group by sd19"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      intSheets = RsTemp.RecordCount
      RsTemp.MoveFirst
      j = 0
      Do While Not RsTemp.EOF
         j = j + 1
         strSheetsVal(j) = RsTemp.Fields(0)
         RsTemp.MoveNext
      Loop
   End If
   xlsSalesPoint.SheetsInNewWorkbook = IIf(intSheets > 0, intSheets, 1) 'Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
   xlsSalesPoint.Workbooks.add
   For j = 1 To intSheets
      Set wksaccrpt114 = xlsSalesPoint.Worksheets(j)
      strSheetsName = A0802Query(strSheetsVal(j), True)
      wksaccrpt114.Name = strSheetsName
      wksaccrpt114.Columns("a:a").ColumnWidth = 11: wksaccrpt114.Range("a1").Value = "": wksaccrpt114.Range("a2").Value = "姓名"
      'Added by Morgan 2025/3/18
      wksaccrpt114.Columns("B:B").ColumnWidth = 11: wksaccrpt114.Range("B1").Value = "": wksaccrpt114.Range("B2").Value = "新部門"
      wksaccrpt114.Columns("C:C").ColumnWidth = 11: wksaccrpt114.Range("B1").Value = "": wksaccrpt114.Range("C2").Value = "職稱"
      'end 2025/3/18
      wksaccrpt114.Columns("D:D").ColumnWidth = 8: wksaccrpt114.Range("D1").Value = Val(txt1(0)) + 1911: wksaccrpt114.Range("D2").Value = "股數"
      wksaccrpt114.Columns("E:E").ColumnWidth = 8: wksaccrpt114.Range("E1").Value = "": wksaccrpt114.Range("E2").Value = "紅利"
      wksaccrpt114.Columns("F:F").ColumnWidth = 8: wksaccrpt114.Range("F1").Value = "": wksaccrpt114.Range("F2").Value = "特殊功績獎金"
      wksaccrpt114.Columns("G:G").ColumnWidth = 8: wksaccrpt114.Range("G1").Value = "": wksaccrpt114.Range("G2").Value = "合計"
      wksaccrpt114.Columns("H:H").ColumnWidth = 8: wksaccrpt114.Range("H1").Value = "": wksaccrpt114.Range("H2").Value = "部門建議金額"
      wksaccrpt114.Columns("I:I").ColumnWidth = 8: wksaccrpt114.Range("I1").Value = Val(txt1(0)) + 1911 - 1: wksaccrpt114.Range("I2").Value = "股數"
      wksaccrpt114.Columns("J:J").ColumnWidth = 8: wksaccrpt114.Range("J1").Value = "": wksaccrpt114.Range("J2").Value = "紅利"
      wksaccrpt114.Columns("K:K").ColumnWidth = 8: wksaccrpt114.Range("K1").Value = "": wksaccrpt114.Range("K2").Value = "特殊功績獎金"
      wksaccrpt114.Columns("L:L").ColumnWidth = 8: wksaccrpt114.Range("L1").Value = "": wksaccrpt114.Range("L2").Value = "合計"
      wksaccrpt114.Columns("M:M").ColumnWidth = 8: wksaccrpt114.Range("M1").Value = Val(txt1(0)) + 1911 - 2: wksaccrpt114.Range("M2").Value = "股數"
      wksaccrpt114.Columns("N:N").ColumnWidth = 8: wksaccrpt114.Range("N1").Value = "": wksaccrpt114.Range("N2").Value = "紅利"
      wksaccrpt114.Columns("O:O").ColumnWidth = 8: wksaccrpt114.Range("O1").Value = "": wksaccrpt114.Range("O2").Value = "特殊功績獎金"
      wksaccrpt114.Columns("P:P").ColumnWidth = 8: wksaccrpt114.Range("P1").Value = "": wksaccrpt114.Range("P2").Value = "合計"
      xlsSalesPoint.Sheets(strSheetsName).Activate
      wksaccrpt114.Range("F2").Select
      With xlsSalesPoint.Selection.Font
        '.Name = "新細明體"
        .Size = 10
      End With
      wksaccrpt114.Range("H2").Select
      With xlsSalesPoint.Selection.Font
        '.Name = "新細明體"
        .Size = 10
      End With
      wksaccrpt114.Range("K2").Select
      With xlsSalesPoint.Selection.Font
        '.Name = "新細明體"
        .Size = 10
      End With
      wksaccrpt114.Range("O2").Select
      With xlsSalesPoint.Selection.Font
        '.Name = "新細明體"
        .Size = 10
      End With
      wksaccrpt114.Range("A2:P2").Select
      With xlsSalesPoint.Selection
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlCenter
      End With
      '+顏色
      wksaccrpt114.Range("D1:H1").Interior.ColorIndex = 43 '淺綠
      wksaccrpt114.Range("I1:L1").Interior.ColorIndex = 20 '淺藍
      wksaccrpt114.Range("M1:P1").Interior.ColorIndex = 19 '淺黃
      '合併儲格
      With wksaccrpt114.Range("a1:a2")
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      'Added by Morgan 2025/3/18
      With wksaccrpt114.Range("B1:B2")
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      With wksaccrpt114.Range("C1:C2")
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      'end 2025/3/18
      With wksaccrpt114.Range("D1:H1")
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      With wksaccrpt114.Range("I1:L1")
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      With wksaccrpt114.Range("M1:P1")
          .HorizontalAlignment = xlCenter
          .VerticalAlignment = xlBottom
          .WrapText = False
          .Orientation = 0
          .AddIndent = False
          .ShrinkToFit = False
          .MergeCells = True
      End With
      'Modified by Morgan 2023/12/20 st03<>'P29'--> AND sd01 is not null
      'strExc(0) = "SELECT sd19,st03,st01,st02||decode(st04,'2','(職)','') as st02nm" & _
                  " FROM staff,salarydata" & _
                  " WHERE st01=sd01(+) AND st03<>'P29' AND st01>'63' AND st01<'F' AND substr(st01,4,1)<>'9'" & _
                  " AND ((st04='1' and st13<=" & Val(txt1(0)) + 1911 & "1231) or (st04<>'1' and st13<=" & Val(txt1(0)) + 1911 & "1231 and st51>" & Val(txt1(0)) + 1911 & "1231))" & _
                  " AND (sd19='" & strSheetsVal(j) & "'" & IIf(j = 1, " or sd19 is null", "") & _
                  IIf(strSheetsVal(j) = "2", " or sd19='1'", "") & ")" & _
                  " order by sd19,st03,st01"
      'Modified by Morgan 2025/3/18 +新部門,職稱
      strExc(0) = "SELECT sd19,st03,st01,st02||decode(st04,'2','(職)','') as st02nm,a0922 dept,ac03 tit" & _
                  " FROM staff,salarydata,acc090new,allcode" & _
                  " WHERE st01=sd01(+) AND sd01 is not null AND st01>'63' AND st01<'F' AND substr(st01,4,1)<>'9'" & _
                  " AND ((st04='1' and st13<=" & Val(txt1(0)) + 1911 & "1231) or (st04<>'1' and st13<=" & Val(txt1(0)) + 1911 & "1231 and st51>" & Val(txt1(0)) + 1911 & "1231))" & _
                  " AND (sd19='" & strSheetsVal(j) & "'" & IIf(j = 1, " or sd19 is null", "") & _
                  IIf(strSheetsVal(j) = "2", " or sd19='1'", "") & ")" & _
                  " and a0921(+)=st93 and ac02(+)=st20 and ac01(+)='01'" & _
                  " order by sd19,st03,st01"
      intI = 1
      Set adoRs = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         adoRs.MoveFirst
         intRow = 2
         Do While Not adoRs.EOF
            intRow = intRow + 1
            wksaccrpt114.Range("a" & intRow).Value = adoRs.Fields("st01") & adoRs.Fields("st02nm")
            'Added by Morgan 2025/3/18
            wksaccrpt114.Range("B" & intRow).Value = "" & adoRs.Fields("dept")
            wksaccrpt114.Range("C" & intRow).Value = "" & adoRs.Fields("tit")
            'end 2025/3/18
            strExc(0) = "select yb26,yb06,1 as sort from yearbonus where yb01=" & Val(txt1(0)) + 1911 & " and yb02='" & adoRs.Fields("st01") & "'" & _
                        " union select yb26,yb06,2 as sort from yearbonus where yb01=" & Val(txt1(0)) + 1911 - 1 & " and yb02='" & adoRs.Fields("st01") & "'" & _
                        " union select yb26,yb06,3 as sort from yearbonus where yb01=" & Val(txt1(0)) + 1911 - 2 & " and yb02='" & adoRs.Fields("st01") & "'" & _
                        " order by sort asc"
            intI = 1
            Set adoRs2 = ClsLawReadRstMsg(intI, strExc(0))
            strSort = ""
            If intI = 1 Then
               adoRs2.MoveFirst
               Do While Not adoRs2.EOF
                  strSort = strSort & "," & adoRs2.Fields("sort")
                  If adoRs2.Fields("sort") = 1 Then
                     wksaccrpt114.Range("E" & intRow).Value = adoRs2.Fields("yb26")
                     wksaccrpt114.Range("F" & intRow).Value = adoRs2.Fields("yb06")
                     wksaccrpt114.Range("G" & intRow).Formula = "=E" & intRow & "+F" & intRow
                  ElseIf adoRs2.Fields("sort") = 2 Then
                     wksaccrpt114.Range("J" & intRow).Value = adoRs2.Fields("yb26")
                     wksaccrpt114.Range("K" & intRow).Value = adoRs2.Fields("yb06")
                     wksaccrpt114.Range("L" & intRow).Formula = "=J" & intRow & "+K" & intRow
                  ElseIf adoRs2.Fields("sort") = 3 Then
                     wksaccrpt114.Range("N" & intRow).Value = adoRs2.Fields("yb26")
                     wksaccrpt114.Range("O" & intRow).Value = adoRs2.Fields("yb06")
                     wksaccrpt114.Range("P" & intRow).Formula = "=N" & intRow & "+O" & intRow
                  End If
                  adoRs2.MoveNext
               Loop
            End If
            If InStr(strSort, "1") = 0 Then
               wksaccrpt114.Range("E" & intRow).Value = 0
               wksaccrpt114.Range("F" & intRow).Value = 0
               wksaccrpt114.Range("G" & intRow).Formula = "=E" & intRow & "+F" & intRow
            End If
            If InStr(strSort, "2") = 0 Then
               wksaccrpt114.Range("J" & intRow).Value = 0
               wksaccrpt114.Range("K" & intRow).Value = 0
               wksaccrpt114.Range("L" & intRow).Formula = "=J" & intRow & "+K" & intRow
            End If
            If InStr(strSort, "3") = 0 Then
               wksaccrpt114.Range("N" & intRow).Value = 0
               wksaccrpt114.Range("O" & intRow).Value = 0
               wksaccrpt114.Range("P" & intRow).Formula = "=N" & intRow & "+O" & intRow
            End If
            adoRs.MoveNext
         Loop
         '凍結窗格
         wksaccrpt114.Range("D3").Select
         xlsSalesPoint.ActiveWindow.FreezePanes = True
         '格式化
         wksaccrpt114.Range("E3:G" & intRow).Select
         wksaccrpt114.Range("E3:G" & intRow).NumberFormatLocal = "#,##0_ "
         wksaccrpt114.Range("J3:L" & intRow).Select
         wksaccrpt114.Range("J3:L" & intRow).NumberFormatLocal = "#,##0_ "
         wksaccrpt114.Range("N3:P" & intRow).Select
         wksaccrpt114.Range("N3:P" & intRow).NumberFormatLocal = "#,##0_ "
         '框線
         wksaccrpt114.Range("a1:P" & intRow).Select
         With xlsSalesPoint.Selection.Borders(xlEdgeLeft)
             .LineStyle = xlContinuous
             .ColorIndex = xlAutomatic
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With xlsSalesPoint.Selection.Borders(xlEdgeTop)
             .LineStyle = xlContinuous
             .ColorIndex = xlAutomatic
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With xlsSalesPoint.Selection.Borders(xlEdgeBottom)
             .LineStyle = xlContinuous
             .ColorIndex = xlAutomatic
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With xlsSalesPoint.Selection.Borders(xlEdgeRight)
             .LineStyle = xlContinuous
             .ColorIndex = xlAutomatic
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With xlsSalesPoint.Selection.Borders(xlInsideVertical)
             .LineStyle = xlContinuous
             .ColorIndex = xlAutomatic
             .TintAndShade = 0
             .Weight = xlThin
         End With
         With xlsSalesPoint.Selection.Borders(xlInsideHorizontal)
             .LineStyle = xlContinuous
             .ColorIndex = xlAutomatic
             .TintAndShade = 0
             .Weight = xlThin
         End With
         wksaccrpt114.Range("a1").Select
      End If
      
      'Added by Morgan 2025/3/18
      '自動設定欄寬
      For intI = Asc("A") To Asc("P")
         wksaccrpt114.Columns(Chr(intI)).EntireColumn.AutoFit
      Next
      'end 2025/3/18
   Next j
   
   xlsSalesPoint.Sheets(1).Activate
   'xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName
   If Val(xlsSalesPoint.Version) < 12 Then
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   'Modify by Amy 2021/06/22 路徑改中文字顯示
   MsgBox "檔案已產生！電子檔位置：" & strExcelPathN & Replace(strFileName, strExcelPath, "")
   
flgErr:
   Set wksaccrpt114 = Nothing
   Set xlsSalesPoint = Nothing
   Set adoRs = Nothing
   Set adoRs2 = Nothing
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Sub Form_Load()
Dim strSystemKind As String
   
   MoveFormToCenter Me
   
   strSystemKind = GetSystemKindByNick
   PUB_SetPrinter Me.Name, Combo1
   
   txt1(0) = Left(strSrvDate(2), 3)
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170240 = Nothing
End Sub

Private Sub txt1_GotFocus(Index As Integer)
   InverseTextBox txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   Select Case Index
      Case 0
         KeyAscii = Pub_NumAscii(KeyAscii)
      Case Else
   End Select
End Sub
