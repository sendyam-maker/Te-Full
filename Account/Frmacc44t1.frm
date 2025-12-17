VERSION 5.00
Begin VB.Form Frmacc44t1 
   AutoRedraw      =   -1  'True
   Caption         =   "客戶年度未扣繳查詢"
   ClientHeight    =   3012
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   3900
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3012
   ScaleWidth      =   3900
   Begin VB.TextBox txtCmp 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   1
      Left            =   1920
      MaxLength       =   1
      TabIndex        =   1
      Text            =   "1"
      Top             =   500
      Width           =   345
   End
   Begin VB.TextBox txtCmp 
      Alignment       =   1  '靠右對齊
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Index           =   2
      Left            =   2520
      MaxLength       =   1
      TabIndex        =   2
      Text            =   "L"
      Top             =   500
      Width           =   390
   End
   Begin VB.TextBox txtYear 
      Height          =   290
      Left            =   1920
      MaxLength       =   3
      TabIndex        =   0
      Top             =   150
      Width           =   1000
   End
   Begin VB.CommandButton CmdExcel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel(&E)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   780
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   2436
      Width           =   2115
   End
   Begin VB.Label Label7 
      Caption         =   " 公 司 別 "
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   840
      TabIndex        =   7
      Top             =   504
      Width           =   996
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   2340
      TabIndex        =   6
      Top             =   480
      Width           =   120
   End
   Begin VB.Label Label2 
      Caption         =   "扣繳年度"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   288
      Left            =   840
      TabIndex        =   5
      Top             =   156
      Width           =   996
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "報表包含之未扣繳資料如下："
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   2800
      Left            =   180
      TabIndex        =   4
      Top             =   1000
      Width           =   5232
   End
   Begin VB.Image Image1 
      Height          =   135
      Left            =   0
      Top             =   1110
      Visible         =   0   'False
      Width           =   135
   End
End
Attribute VB_Name = "Frmacc44t1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create By Amy 2025/11/14
Option Explicit

Dim adoQ As New ADODB.Recordset
Dim i As Integer, intField As Integer, intTitleR As Integer
Dim strAllF As String, strAllW As String, strF() As String, arrWidth

Private Sub CmdExcel_Click()
   Dim strShowMsg As String, hLocalFile As Long
   
   If ChkForm = False Then
      Exit Sub
   End If
   
   Screen.MousePointer = vbHourglass
   If SaveExcel(strShowMsg) = False Then
      Screen.MousePointer = vbDefault
      Exit Sub
   Else
      If strShowMsg = "無資料產生" Then
         MsgBox strShowMsg, vbCritical
      Else
         If MsgBox("EXCEL檔案已產生！" & vbCrLf & vbCrLf & strExcelPath & vbCrLf & vbCrLf & "是否開啟資料夾？", vbYesNo + vbDefaultButton1 + vbInformation) = vbYes Then
            ShellExecute hLocalFile, "explore", strExcelPath, vbNullString, vbNullString, 1
         End If
      End If
   End If
   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
   strFormName = Name
   '表單初始化
   PUB_InitForm Me, 3950, 3450, strBackPicPath4
   '設定欄位
   strAllF = "收據抬頭,收據號碼,收據日期,收款日期,合併,服務費,規費,應扣額"
   strAllW = "22,11,11,11,6,11,11,11"
   strF = Split(strAllF, ",")
   arrWidth = Split(strAllW, ",")
   '說明
   Label1.Caption = Label1.Caption & vbCrLf & _
                                 "1.畫面扣繳年度單筆稅額超過2001" & vbCrLf & _
                                 "2.畫面扣繳年度可扣繳" & vbCrLf & _
                                 "　以上收據排除境外公司"
   
   If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
      MkDir strExcelPath
   End If
End Sub

Private Function ChkForm() As Boolean
   ChkForm = False
   
   '扣繳年度
   If txtYear = "" Then
      MsgBox "請輸入扣繳年度！", vbExclamation
      txtYear.SetFocus
      Exit Function
   Else
      If IsNumeric(txtYear) = False Then
         MsgBox "扣繳年度輸入錯誤！", vbExclamation
         txtYear.SetFocus
         Exit Function
      End If
   End If
   
   '公司別
   If txtCmp(1) <> "" Then
      If txtCmp(1) <> "1" And txtCmp(1) <> "L" Then
         MsgBox "公司別(起)輸入錯誤！", vbExclamation
         txtCmp(1).SetFocus
         Exit Function
      End If
   End If
   If txtCmp(2) <> "" Then
      If txtCmp(2) <> "1" And txtCmp(2) <> "L" Then
         MsgBox "公司別(迄)輸入錯誤！", vbExclamation
         txtCmp(2).SetFocus
         Exit Function
      End If
   End If
   
   ChkForm = True
End Function

Private Function SaveExcel(ByRef stShowMsg As String) As Boolean
   Dim xlsAp As New Excel.Application, Wks As New Worksheet, intQ As Integer, intRow As Integer, intWkPage As Integer
   Dim strQ As String, strWhr As String, strWhr2 As String, strWkName As String, strOldCmp As String, strFormat As String, strFileN As String
   Dim bolOpenXls As Boolean, strTp As String
On Error GoTo ErrHnd
   
   SaveExcel = False: intField = 65: intWkPage = 1: intRow = 1: stShowMsg = ""
   '檔名
   strFileN = txtYear & "年客戶年度未扣繳明細" & ACDate(ServerDate) & ServerTime
   '扣繳年度
   If txtYear <> "" Then
      strWhr = "And A0k16=" & Val(txtYear) & " "
      
   End If
   '公司別
   If txtCmp(1) <> "" Then
      strWhr = strWhr & "And a0k11>='" & txtCmp(1) & "' "
   End If
   If txtCmp(2) <> "" Then
      strWhr = strWhr & "And a0k11<='" & txtCmp(2) & "' "
   End If
   strWhr = strWhr & "And a0k11<>'J' "
   
   '分次收款時,收款日取最大者；考慮拆收據情形,是否合併改抓a0j07(參考frmacc44i0)
   strQ = "Select A0k04,A0k01,Sqldatet(A0K02+19110000) as a1K02,RecDate,a0j07,A0k06,A0k07,A1v04,A0k11 " & _
               "From Acc0k0,Acc0m0,Acc0L0,Acc1v0,Acc0j0,Customer " & _
                  ",(Select a0m02 as RecNo,Sqldatet(Max(a0L02)+19110000) as RecDate From Acc0k0,Acc0m0,Acc0l0 " & _
                     "Where a0m02(+)=a0k01 And a0l01(+)=a0m01 And a0m02 is not null " & strWhr & "Group by a0m02) A " & _
               "Where A0k05='2' And A1v04>2000 And a1v06=0 " & strWhr & "And A0k01=A1v02(+) " & _
               "And A0j01(+)=a1v01 And A0j13(+)=a1v02 And A0L01(+)=A0m01 And A0m02(+)=A0k01  And RecNo(+)=A0k01 " & _
               "And SubStr(a0k03,1,8)=cu01(+) And SubStr(a0k03,9,1)=cu02(+) And cu158 is null " & _
               "Order by A0k11,A0k04,A0k01 "
   intQ = 1
   Set adoQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      If adoQ.RecordCount > 0 Then
         adoQ.MoveFirst
         xlsAp.Visible = True
         xlsAp.SheetsInNewWorkbook = 3
         xlsAp.Workbooks.add
         bolOpenXls = True
         
         strWkName = Left(xlsAp.Worksheets(1).Name, Len(xlsAp.Worksheets(1).Name) - 1)
         Do While adoQ.EOF = False
            If strOldCmp <> "" & adoQ.Fields("A0K11") Then
              If strOldCmp <> "" Then
                  Call SetExcelEnd(xlsAp, Wks, strFileN, intRow)
                  Wks.Name = A0802Query(strOldCmp, True)
                  intWkPage = intWkPage + 1
                  intRow = 1
               End If
               Set Wks = xlsAp.Worksheets(strWkName & intWkPage)
               Wks.Activate
               Call SetTitle(xlsAp, Wks, IIf(strOldCmp = "", "" & adoQ.Fields("A0K11"), strOldCmp), intRow)
               intTitleR = intRow
               intRow = intRow + 1
            End If
            For i = LBound(strF) To UBound(strF)
               strFormat = ""
               Select Case i
                  Case GetColVal(strF, "收據抬頭")
                     strTp = "" & adoQ.Fields("A0k04")
                  Case GetColVal(strF, "收據號碼")
                     strTp = "" & adoQ.Fields("A0k01")
                  Case GetColVal(strF, "收據日期")
                     strTp = "" & adoQ.Fields("a1K02")
                  Case GetColVal(strF, "收款日期")
                     strTp = "" & adoQ.Fields("RecDate")
                  Case GetColVal(strF, "合併")
                     strTp = "" & adoQ.Fields("a0j07")
                  Case GetColVal(strF, "服務費")
                     strTp = "" & adoQ.Fields("A0k06")
                     strFormat = "#,##"
                  Case GetColVal(strF, "規費")
                     strTp = "" & adoQ.Fields("A0k07")
                     strFormat = "#,##"
                  Case GetColVal(strF, "應扣額")
                     strTp = "" & adoQ.Fields("A1v04")
                     strFormat = "#,##"
               End Select
               If strFormat <> "" Then
                  Wks.Range(Chr(i + intField) & intRow).NumberFormatLocal = strFormat
               End If
               Wks.Range(Chr(i + intField) & intRow).Value = strTp
            Next i
            intRow = intRow + 1
            strOldCmp = "" & adoQ.Fields("A0K11")
            adoQ.MoveNext
         Loop
         '最後設定
         Wks.Name = A0802Query(strOldCmp, True)
         Call SetExcelEnd(xlsAp, Wks, strFileN, intRow)
      End If
   Else
      stShowMsg = "無資料產生"
   End If
   SaveExcel = True
   
ErrHnd:
   If Err.Number <> 0 Then
      MsgBox Err.Description, , MsgText(5)
      'Resume
   End If
   If bolOpenXls = True Then
      If Val(xlsAp.Version) < 12 Then
         xlsAp.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN & MsgText(43), FileFormat:=-4143
      Else
         xlsAp.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN & ".xlsx", FileFormat:=51
      End If
      xlsAp.Workbooks.Close
      xlsAp.Quit
      Set Wks = Nothing
      Set xlsAp = Nothing
   End If
   Set adoQ = Nothing
End Function

Private Sub SetTitle(xlsApp As Excel.Application, Wks As Worksheet, ByVal stCmp As String, ByRef intRow As Integer)
   Dim ii As Integer, stTxt As String
     
   Wks.Range(Chr(intField) & intRow).Value = "客戶年度未扣繳明細"
   Wks.Range(Chr(intField) & intRow).Font.Size = 14
   Wks.Range(Chr(intField) & intRow).Font.Bold = True
   
   Wks.Range(Chr(intField) & intRow & ":" & Chr(UBound(strF) + intField) & intRow).Select
   With xlsApp.Selection
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlBottom
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .ShrinkToFit = False
      .MergeCells = True
   End With
    intRow = intRow + 1
   '條件
   stTxt = "扣繳年度：" & Format(txtYear, "000")
   Wks.Range(Chr(intField) & intRow).Value = stTxt
   Wks.Range(Chr(intField) & intRow & ":" & Chr(UBound(strF) + intField) & intRow).Select
   With xlsApp.Selection
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlBottom
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .ShrinkToFit = False
      .MergeCells = True
   End With
   intRow = intRow + 1
   
   stTxt = "公司別：" & IIf(stCmp = "2", "1", stCmp)
   Wks.Range(Chr(intField) & intRow).Value = stTxt
   Wks.Range(Chr(intField) & intRow & ":" & Chr(UBound(strF) + intField) & intRow).Select
   With xlsApp.Selection
      .HorizontalAlignment = xlCenter
      .VerticalAlignment = xlBottom
      .WrapText = False
      .Orientation = 0
      .AddIndent = False
      .ShrinkToFit = False
      .MergeCells = True
   End With
   intRow = intRow + 1
   
   '列印人員/日期
   stTxt = "列印人員：" & StaffQuery(strUserNum)
   Wks.Range(Chr(intField) & intRow).Value = stTxt
   stTxt = "列印日期：" & CFDate(ACDate(ServerDate))
   Wks.Range(Chr((intField + UBound(strF) - 2)) & intRow).Value = stTxt
   intRow = intRow + 1

   For ii = LBound(strF) To UBound(strF)
      stTxt = strF(ii)
      Wks.Range(Chr(intField + ii) & intRow).Value = stTxt
      Wks.Range(Chr(intField + ii) & intRow).Font.Bold = True
      Wks.Range(Chr(intField + ii) & intRow).ColumnWidth = Val(arrWidth(ii))
      Wks.Range(Chr(intField + ii) & intRow).HorizontalAlignment = xlCenter
   Next ii
End Sub

Private Sub SetExcelEnd(xlsApp As Excel.Application, Wks As Worksheet, stFileN As String, ByRef intRow As Integer)
  '內容字大小
   Wks.Range(Chr(intField) & intTitleR & ":" & Chr(intField + UBound(strF)) & intRow).Font.Size = 12
   '畫線
   Wks.Range(Chr(intField) & intRow - 1 & ":" & Chr(intField + UBound(strF)) & intRow - 1).Borders(xlEdgeBottom).LineStyle = xlContinuous
   '版面設定
   Wks.PageSetup.Orientation = xlPortrait '直印
   Wks.PageSetup.Zoom = 100 '縮放比例為100%,列印頁面水平置中
   Wks.PageSetup.HeaderMargin = xlsApp.Application.InchesToPoints(0) '頁首
   Wks.PageSetup.FooterMargin = xlsApp.Application.InchesToPoints(0) '頁尾
   Wks.PageSetup.TopMargin = xlsApp.InchesToPoints(0.3) '上
   Wks.PageSetup.BottomMargin = xlsApp.InchesToPoints(0.3) '下
   Wks.PageSetup.LeftMargin = xlsApp.InchesToPoints(0.1) '左邊界
   Wks.PageSetup.RightMargin = xlsApp.InchesToPoints(0.1) '右邊界
   Wks.PageSetup.PrintTitleRows = "$1:$" & intTitleR '表頭保留列
   Wks.PageSetup.CenterHorizontally = True '水平置中(版面設定->邊界->水平置中)

End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   
   Set Frmacc44t1 = Nothing
End Sub

Private Sub txtCmp_GotFocus(Index As Integer)
   TextInverse txtCmp(Index)
End Sub

Private Sub txtCmp_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub txtYear_GotFocus()
   TextInverse txtYear
End Sub
