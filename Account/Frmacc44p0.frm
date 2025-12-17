VERSION 5.00
Begin VB.Form Frmacc44p0 
   AutoRedraw      =   -1  'True
   Caption         =   "扣繳憑單明細表"
   ClientHeight    =   2595
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5460
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2595
   ScaleWidth      =   5460
   Begin VB.TextBox Text6 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   4
      Top             =   1350
      Width           =   612
   End
   Begin VB.TextBox Text4 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2760
      MaxLength       =   1
      TabIndex        =   3
      Top             =   990
      Width           =   612
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   372
      Left            =   360
      Style           =   1  '圖片外觀
      TabIndex        =   5
      Top             =   2040
      Width           =   4815
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   2280
      MaxLength       =   1
      TabIndex        =   2
      Top             =   630
      Width           =   612
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1320
      MaxLength       =   1
      TabIndex        =   1
      Top             =   630
      Width           =   612
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1320
      TabIndex        =   0
      Top             =   270
      Width           =   852
   End
   Begin VB.Label Label9 
      BackStyle       =   0  '透明
      Caption         =   "是否列印備註欄"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   12
      Top             =   1350
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  '透明
      Caption         =   "(Y/N)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   11
      Top             =   1350
      Width           =   735
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "(Y/N,空白表全部)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3480
      TabIndex        =   10
      Top             =   990
      Width           =   1770
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "扣繳憑單金額是否為零?"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   9
      Top             =   990
      Width           =   2415
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "-"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   13.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2040
      TabIndex        =   8
      Top             =   630
      Width           =   255
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   7
      Top             =   630
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "扣繳年度"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   360
      TabIndex        =   6
      Top             =   270
      Width           =   975
   End
End
Attribute VB_Name = "Frmacc44p0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit

'Public adoacc080 As New ADODB.Recordset
Public adoacc0w0 As New ADODB.Recordset
Public adoaccrpt421 As New ADODB.Recordset
'Dim dllaccrpt421(10) As Object
Dim intCounter As Integer


Private Sub Command1_Click()
   Dim strSql As String

   intCounter = 1
   If FormCheck = False Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Sub
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0801 >= '" & Text2 & "'"
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0801 <= '" & Text3 & "'"
   End If
   If Text4 = MsgText(602) Then
      strSql = strSql & " and a0w05 = 0"
   ElseIf Text4 = MsgText(603) Then
      strSql = strSql & " and a0w05 > 0"
   End If
   
   Screen.MousePointer = vbHourglass
   Accrpt421Delete
   If ProduceData Then
       Call ExcelSave
'      adoacc080.CursorLocation = adUseClient
'      'modify by sonia 2018/3/13
'      'adoacc080.Open "select distinct a0801, a0802 from acc0w0, acc080 where a0w04 = a0801 and (a0w15 is null)" & strSql & " order by a0801 asc", adoTaie, adOpenStatic, adLockReadOnly
'      adoacc080.Open "select distinct a0801, a0802 from acc0w0, acc080 where a0w01= " & Val(Text1) & " and a0w04 = a0801 and (a0w15 is null)" & strSql & " order by a0801 asc", adoTaie, adOpenStatic, adLockReadOnly
'      Do While adoacc080.EOF = False
''         Set dllaccrpt421(intCounter) = Nothing
''         Set dllaccrpt421(intCounter) = CreateObject("AccReport.ReportSelect")
''         RunReportDll
''         If (Text5 = MsgText(602)) Then
'            Text2Excel adoacc080.Fields("A0801")
''         End If
'         intCounter = intCounter + 1
'         adoacc080.MoveNext
'      Loop
'      adoacc080.Close
''      If (Text5 = MsgText(602)) Then
'         Me.Show
         MsgBox "Excel 匯出完成！", vbOKOnly
''      End If
      
   End If
   Screen.MousePointer = vbDefault
   FormClear
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

''*************************************************
''  Text 轉成 Excel 檔案
''
''*************************************************
'Sub Text2Excel(strCompNo As String)
'
'   Dim xlsSalesPoint As New Excel.Application
'   Dim wksaccrpt421 As New Worksheet
'   Dim sXlsPath As String, sFilePath As String
'   Dim sTMPPath As String
'
'   Frmacc0000.StatusBar1.Panels(1).Text = "匯出Excel檔案中..."
'
'   sXlsPath = Mid(strExcelPath, 1, Len(strExcelPath) - 1)
'   sFilePath = strExcelPath & Mid(ReportTitle(421), 6, 7) & strCompNo & "公司" & ACDate(ServerDate) & ServerTime & MsgText(43)
'   'Add By Sindy 2018/2/27
'   'sTMPPath = "C:\TAIE.TMP"
'   sTMPPath = strExcelPath & "TAIE.TMP"
'   '2018/2/27 END
'
'   '刪除舊檔
'   If Dir(sFilePath) = MsgText(601) Then
'      If Dir(sXlsPath, vbDirectory) = MsgText(601) Then
'         MkDir strExcelPath
'      End If
'   Else
'      Kill sFilePath
'   End If
'
'   If Dir(sTMPPath) = MsgText(601) Then Exit Sub
'
'   xlsSalesPoint.Workbooks.OpenText FileName:=sTMPPath, StartRow:=1, _
'        DataType:=xlFixedWidth, FieldInfo:=Array(Array(0, 2), Array(20, 2), Array(85, 1), Array(132, 2))
'   Set wksaccrpt421 = xlsSalesPoint.Worksheets(1)
'
'   'Add by Morgan 2005/5/9 改欄位名稱
'   'wksaccrpt421.Range("C14").FormulaR1C1 = "扣單稅額"
'   'wksaccrpt421.Range("D14").FormulaR1C1 = "給付總額"
'   wksaccrpt421.Columns("C:C").Select
'   xlsSalesPoint.Selection.NumberFormatLocal = "#,##0_ "
'   '3,5,8公司調整欄位
'   If strCompNo = "3" Or strCompNo = "5" Or strCompNo = "8" Then
'      wksaccrpt421.Columns("D:D").Select
'      xlsSalesPoint.Selection.HorizontalAlignment = xlRight
'      wksaccrpt421.Columns("C:C").Select
'      xlsSalesPoint.Selection.Cut
'      wksaccrpt421.Range("E1").Activate
'      xlsSalesPoint.ActiveSheet.Paste
'      wksaccrpt421.Columns("C:C").Select
'      xlsSalesPoint.Selection.Delete Shift:=xlToLeft
'   End If
'   '2005/5/9 end
'
'   wksaccrpt421.Columns("A:A").EntireColumn.AutoFit
'   wksaccrpt421.Columns("B:B").EntireColumn.AutoFit
'   wksaccrpt421.Columns("C:C").EntireColumn.AutoFit
'   wksaccrpt421.Columns("D:D").EntireColumn.AutoFit
'
''    wksaccrpt421.Columns("D:D").Select
''    xlsSalesPoint.Selection.NumberFormatLocal = "#,##0_ "
''    xlsSalesPoint.Selection.Cut
''    wksaccrpt421.Range("C4").Activate
''    xlsSalesPoint.Selection.Insert Shift:=xlToRight
'
'
''   xlsSalesPoint.ActiveWindow.ScrollRow = 3
''   wksaccrpt421.Rows("8:10").Select
''   xlsSalesPoint.Selection.ClearContents
''   xlsSalesPoint.Selection.Delete Shift:=xlUp
'   'Modify by Amy 2016/06/23 +判斷版本
'   If Val(xlsSalesPoint.Version) < 12 Then
''        xlsSalesPoint.ActiveWorkbook.SaveAs FileName:=sFilePath, FileFormat:= _
''        xlNormal, Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
''        , CreateBackup:=False
'        xlsSalesPoint.ActiveWorkbook.SaveAs FileName:=sFilePath, FileFormat:=-4143 _
'        , Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
'        , CreateBackup:=False
'   Else
'         xlsSalesPoint.ActiveWorkbook.SaveAs FileName:=sFilePath, FileFormat:=56 _
'        , Password:="", WriteResPassword:="", ReadOnlyRecommended:=False _
'        , CreateBackup:=False
'   End If
'   'end 2016/06/23
'   xlsSalesPoint.Workbooks.Close
'   xlsSalesPoint.Quit
'   Set xlsSalesPoint = Nothing
'   'Kill sTMPPath
'   Frmacc0000.StatusBar1.Panels(1).Text = "匯出Excel檔案完成"
'   StatusClear
'End Sub

'Add by Sindy 2020/5/6
'*************************************************
'  轉成Excel檔案
'
'*************************************************
Private Sub ExcelSave()
Dim strFilePath As String
Dim xlsSalesPoint As New Excel.Application
Dim wksTmp As New Worksheet
Dim lngCounter As Long
Dim intWorksheet As Integer
Dim strCompNo As String
Dim strVal As String
   
   strSql = "SELECT * FROM accrpt421 WHERE r42101='" & strUserNum & "'" & _
            " order by r42103,r42102"
   intI = 1
   Set adoRecordset = ClsLawReadRstMsg(intI, strSql)
   If intI = 0 Then
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   Else
      adoRecordset.MoveFirst
   End If
   
   Frmacc0000.StatusBar1.Panels(1).Text = "匯出Excel檔案中..."
   
   'Excel檔案路徑
   strFilePath = strExcelPath & Mid(ReportTitle(421), 6, 7) & ACDate(ServerDate) & ServerTime & MsgText(43)
   If Dir(strFilePath) = MsgText(601) Then
      If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
         MkDir strExcelPath
      End If
   Else
      Kill strFilePath
   End If
   
   xlsSalesPoint.SheetsInNewWorkbook = 5 '1
   xlsSalesPoint.Workbooks.add
NewSheet:
   intWorksheet = intWorksheet + 1
   Set wksTmp = xlsSalesPoint.Worksheets(intWorksheet)
   With wksTmp
      wksTmp.Select '切換工作表
      '欄寬
      .PageSetup.Orientation = xlPortrait  '直印
      .PageSetup.PrintTitleRows = "$1:$5"
      .PageSetup.CenterFooter = "第 &P 頁，共 &N 頁"
      .Columns("a:a").ColumnWidth = 11
      .Columns("b:b").ColumnWidth = 10 'Add By Sindy 2022/11/21
      .Columns("c:c").ColumnWidth = 30
      .Columns("d:d").ColumnWidth = 10
      .Columns("e:e").ColumnWidth = 10
      .Columns("f:f").ColumnWidth = 10
      
      '表頭
      .Range("b1").Value = "***  扣繳憑單明細表  ***"
      .Range("b3").Value = "扣繳年度：" & Text1
'      .Range("a5").Value = "公司別：" & Text1
'      .Range("b5").Value = "公司名稱：" & Text1
      
      .Range("a7").Value = "扣繳憑單編號"
      .Range("b7").Value = "客戶統編" 'Add By Sindy 2022/11/21
      .Range("c7").Value = "收據抬頭"
      .Range("d7").Value = "扣單稅額"
      .Range("e7").Value = "給付總額"
      If Text6 = "Y" Then
         .Range("f7").Value = "備註"
      End If
      
      .Range("a1:f1").Select
      With .Range("a1:f1")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlBottom
         .WrapText = False
         .Orientation = 0
         .AddIndent = False
         .ShrinkToFit = False
         .MergeCells = True
      End With
      .Range("a3:f3").Select
      With .Range("a3:f3")
         .HorizontalAlignment = xlCenter
         .VerticalAlignment = xlBottom
         .WrapText = False
         .Orientation = 0
         .AddIndent = False
         .ShrinkToFit = False
         .MergeCells = True
      End With
      
      lngCounter = 7
      Do While Not adoRecordset.EOF
         If lngCounter = 7 Then
            strCompNo = adoRecordset.Fields("r42103")
            '公司別資料
            .Range("a5").Value = "公司別：" & strCompNo
            .Range("b5").Value = "公司名稱：" & A0802Query(strCompNo)
            strVal = A0802Query(strCompNo, True)
            .Name = strCompNo 'strVal '工作表更名
         Else
            If strCompNo <> adoRecordset.Fields("r42103") Then
               '合計
               .Range("a" & lngCounter + 1).Value = "合計"
               .Range("c" & lngCounter + 1).Value = lngCounter - 7 '筆數
               .Range("d" & lngCounter + 1).Formula = "=sum(d8:d" & lngCounter & ")"
               .Range("e" & lngCounter + 1).Formula = "=sum(e8:e" & lngCounter & ")"
               .Range("d8:d" & lngCounter + 1).NumberFormatLocal = "#,##0"
               .Range("e8:e" & lngCounter + 1).NumberFormatLocal = "#,##0"
               GoTo NewSheet '不同公司放不同工作表
            End If
         End If
         lngCounter = lngCounter + 1
         .Range("a" & lngCounter).NumberFormatLocal = "@" '以使用者語言的字串傳回或設定 Variant 值，代表物件的格式代碼。
         .Range("a" & lngCounter).Value = "" & adoRecordset.Fields("r42102")
         .Range("b" & lngCounter).NumberFormatLocal = "@" '文字
         .Range("b" & lngCounter).Value = "" & adoRecordset.Fields("r42108") 'Add By Sindy 2022/11/21
         .Range("c" & lngCounter).Value = "" & adoRecordset.Fields("r42104")
         .Range("d" & lngCounter).Value = "" & adoRecordset.Fields("r42105")
         .Range("e" & lngCounter).Value = "" & adoRecordset.Fields("r42107")
         If Text6 = "Y" Then
            .Range("f" & lngCounter).Value = "" & adoRecordset.Fields("r42106")
         End If
         adoRecordset.MoveNext
      Loop
      '合計
      .Range("a" & lngCounter + 1).Value = "合計"
      .Range("c" & lngCounter + 1).Value = lngCounter - 7 '筆數
      .Range("d" & lngCounter + 1).Formula = "=sum(d8:d" & lngCounter & ")"
      .Range("e" & lngCounter + 1).Formula = "=sum(e8:e" & lngCounter & ")"
      .Range("d8:d" & lngCounter + 1).NumberFormatLocal = "#,##0"
      .Range("e8:e" & lngCounter + 1).NumberFormatLocal = "#,##0"
   End With
   
   '判斷版本
   If Val(xlsSalesPoint.Version) < 12 Then
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=-4143
   Else
        xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFilePath, FileFormat:=56
   End If
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   
   Frmacc0000.StatusBar1.Panels(1).Text = "匯出Excel檔案完成"
   StatusClear
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
'      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5570
   Me.Height = 3000
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   Text4 = MsgText(603)
'   Text5 = MsgText(602)
   Text6 = MsgText(603) 'MsgText(602)
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set Frmacc44p0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
Private Function ProduceData() As Boolean
Dim strSql As String
Dim m_CU11 As String 'Add By Sindy 2022/11/21
   
On Error GoTo Checking
   If Text1 <> MsgText(601) Then
      strSql = " and a0w01 = " & Val(Text1) & ""
   End If
   If Text2 <> MsgText(601) Then
      strSql = strSql & " and a0w04 >= '" & Text2 & "'"
   End If
   If Text3 <> MsgText(601) Then
      strSql = strSql & " and a0w04 <= '" & Text3 & "'"
   End If
   If Text4 = MsgText(602) Then
      strSql = strSql & " and a0w05 = 0"
   ElseIf Text4 = MsgText(603) Then
      strSql = strSql & " and a0w05 > 0"
   End If
'   If strSQL <> MsgText(601) Then
'      strSQL = " where " & Mid(strSQL, 5, Len(strSQL) - 4)
'   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt421.CursorLocation = adUseClient
   adoaccrpt421.Open "select * from accrpt421", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0w0.CursorLocation = adUseClient
   'Modify By Sindy 2020/5/5 + ,a0w16:給付總額
   'Modify By Sindy 2022/11/22 +,a0w03 custname
   adoacc0w0.Open "select a0w01, a0w02, RPAD(a0w03,30,'　') a0w03, a0w04, a0w05, substrb(a0w06,1,50) a0w06,a0w16,a0w03 custname from acc0w0 where a0w15 is null" & strSql & " order by a0w04 asc, a0w02 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0w0.RecordCount = 0 Then
      adoacc0w0.Close
      adoaccrpt421.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Function
   End If
   Do While adoacc0w0.EOF = False
      adoaccrpt421.AddNew
      adoaccrpt421.Fields("r42101").Value = strUserNum
      adoaccrpt421.Fields("r42102").Value = adoacc0w0.Fields("a0w02").Value
      If IsNull(adoacc0w0.Fields("a0w04").Value) Then '公司別
         adoaccrpt421.Fields("r42103").Value = Null
      Else
         adoaccrpt421.Fields("r42103").Value = adoacc0w0.Fields("a0w04").Value
      End If
      '收據抬頭
      If IsNull(adoacc0w0.Fields("a0w03").Value) Then
         adoaccrpt421.Fields("r42104").Value = Null
         adoaccrpt421.Fields("r42108").Value = Null 'Add By Sindy 2022/11/21 + 客戶統編
      Else
         adoaccrpt421.Fields("r42104").Value = adoacc0w0.Fields("a0w03").Value
         Call GetTitleCustData(Trim(adoacc0w0.Fields("custname").Value), "", "", , , _
                            , , , , , , , _
                            , , , , , , , , , , , , , , , m_CU11)
         adoaccrpt421.Fields("r42108").Value = m_CU11 'Add By Sindy 2022/11/21 + 客戶統編
      End If
      If IsNull(adoacc0w0.Fields("a0w05").Value) Then
         adoaccrpt421.Fields("r42105").Value = 0
      Else
         adoaccrpt421.Fields("r42105").Value = Val(adoacc0w0.Fields("a0w05").Value)
      End If
      'Add by Morgan 2004/3/29
      '備註
      'Modify by Morgan 2005/5/9 數字要格式化
      'adoaccrpt421.Fields("r42106").Value = "" & adoacc0w0.Fields("a0w06").Value
      If IsNumeric("" & adoacc0w0.Fields("a0w06").Value) Then
         adoaccrpt421.Fields("r42106").Value = Format(adoacc0w0.Fields("a0w06").Value, DDollar)
      Else
         adoaccrpt421.Fields("r42106").Value = "" & adoacc0w0.Fields("a0w06").Value
      End If
      '2005/5/9 end
      
      'Add By Sindy 2020/5/5 + 給付總額
      If IsNull(adoacc0w0.Fields("a0w16").Value) Then
         adoaccrpt421.Fields("r42107").Value = 0
      Else
         adoaccrpt421.Fields("r42107").Value = Val(adoacc0w0.Fields("a0w16").Value)
      End If
      '2020/5/5 END
      
      adoaccrpt421.UpdateBatch
      adoacc0w0.MoveNext
   Loop
   adoacc0w0.Close
   adoaccrpt421.Close
   
   'Add By Sindy 2017/5/15 收據抬頭（台北市777銀髮族協會）有數字，
   '後面欄位會亂掉，因為我們匯入Excel是截取固定長度
   Dim strCompText As String, strText As String
   adoacc0w0.CursorLocation = adUseClient
   adoacc0w0.Open "select R42104 from accrpt421 where R42101='" & strUserNum & "' group by R42104", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0w0.RecordCount > 0 Then
      adoacc0w0.MoveFirst
      Do While Not adoacc0w0.EOF
         strCompText = adoacc0w0.Fields("R42104")
         strText = PUB_ChangeZIPToSir(adoacc0w0.Fields("R42104"))
         If Trim(strCompText) <> strText Then
            strText = Left(strText & String(15, "　"), 15)
            strSql = "update accrpt421 set" & _
                     " R42104=" & CNULL(ChgSQL(strText)) & _
                     " where R42101='" & strUserNum & "' and R42104='" & adoacc0w0.Fields("R42104") & "'"
            cnnConnection.Execute strSql
         End If
         adoacc0w0.MoveNext
      Loop
   End If
   adoacc0w0.Close
   '2017/5/15 END
   
   StatusClear
   ProduceData = True
   
Checking:
   If Err.Number <> 0 Then MsgBox Err.Description, , MsgText(5)
   
End Function

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt421Delete()
   adoTaie.Execute "delete from accrpt421"
End Sub

''*************************************************
''  執行報表之 Dll
''
''*************************************************
'Private Sub RunReportDll()
'
'   'Modify by Morgan 2003/11/28
'   'dllaccrpt421(intCounter).Acc44p0 ReportTitle(421), Text1, adoacc080.Fields("a0801").Value, adoacc080.Fields("a0802").Value, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'
'   'Modify by Morgan 2004/3/29
'   '改傳陣列參數
''   If (Text5 = MsgText(602)) Then
''      dllaccrpt421(intCounter).Acc44p0 ReportTitle(421) & Chr(0), Text1, adoacc080.Fields("a0801").Value, adoacc080.Fields("a0802").Value, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
''   Else
''      dllaccrpt421(intCounter).Acc44p0 ReportTitle(421), Text1, adoacc080.Fields("a0801").Value, adoacc080.Fields("a0802").Value, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
''   End If
'   dllaccrpt421(intCounter).Acc44p0 ReportTitle(421), Text1 & Chr(23) & Text5 & Chr(23) & Text6, adoacc080.Fields("a0801").Value, adoacc080.Fields("a0802").Value, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'
'   'End
'End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
   Text1 = ""
   Text2 = ""
   Text3 = ""
   Text4 = MsgText(603)
'   Text5 = MsgText(602)
   Text6 = MsgText(602)
   Text1.SetFocus
End Sub

Private Sub Text4_GotFocus()
    TextInverse Me.Text4
End Sub

Private Sub Text4_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 78 And KeyAscii <> 89 And KeyAscii <> 32 Then
        KeyAscii = 0
    End If
End Sub

'Private Sub Text5_GotFocus()
'   TextInverse Text5
'End Sub
'
'Private Sub Text5_KeyPress(KeyAscii As Integer)
'   KeyAscii = UpperCase(KeyAscii)
'   If KeyAscii <> 8 And KeyAscii <> 78 And KeyAscii <> 89 Then
'       KeyAscii = 0
'   End If
'End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

Private Sub Text6_GotFocus()
    TextInverse Me.Text6
End Sub

Private Sub Text6_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
   If KeyAscii <> 8 And KeyAscii <> 78 And KeyAscii <> 89 Then
       KeyAscii = 0
   End If
End Sub
