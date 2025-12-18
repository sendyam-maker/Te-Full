VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form frm040212 
   BorderStyle     =   1  '單線固定
   Caption         =   "FCT審查報告來函超過15個工作天案件"
   ClientHeight    =   885
   ClientLeft      =   1200
   ClientTop       =   3900
   ClientWidth     =   4110
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   885
   ScaleWidth      =   4110
   Begin VB.CommandButton Cmdok 
      Caption         =   "結束(&X)"
      Height          =   350
      Index           =   1
      Left            =   3204
      TabIndex        =   3
      Top             =   10
      Width           =   800
   End
   Begin VB.CommandButton Cmdok 
      Caption         =   "確定(&O)"
      Default         =   -1  'True
      Height          =   350
      Index           =   0
      Left            =   2376
      TabIndex        =   2
      Top             =   10
      Width           =   800
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1290
      TabIndex        =   0
      Top             =   450
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin MSMask.MaskEdBox MaskEdBox2 
      Height          =   300
      Left            =   2820
      TabIndex        =   1
      Top             =   450
      Width           =   1125
      _ExtentX        =   1984
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.5
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2520
      X2              =   2730
      Y1              =   600
      Y2              =   600
   End
   Begin VB.Label lbl1 
      Caption         =   "來函日期："
      Height          =   180
      Index           =   1
      Left            =   105
      TabIndex        =   4
      Top             =   510
      Width           =   1110
   End
End
Attribute VB_Name = "frm040212"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Add by Amy 2021/07/01
Option Explicit

Dim RsQ As New ADODB.Recordset, strQ As String, intQ As Integer
Dim strField, intWidth '欄位名稱/大小
Dim i As Integer, intField As Integer, intCounter As Integer
Dim strFileName As String
 
Private Sub cmdOK_Click(Index As Integer)
    Select Case Index
        Case 0 '確定
            If FormCheck = False Then Exit Sub
            Screen.MousePointer = vbHourglass
            Call doQuery
            Screen.MousePointer = vbDefault
        Case 1 '結束
            Unload Me
        Case Else
    End Select
End Sub

Private Sub Form_Load()
    MoveFormToCenter Me
    FormClear
End Sub

Private Sub doQuery()
    Dim hLocalFile As Long
    Dim strAllField As String, strAllWidth As String, strDateS As String, strDateE As String
    
On Error GoTo ErrHnd

    strAllField = "本所案號,總收文號,承辦人,來函日,來函性質,發文日,工作天"
    strAllWidth = "16,10,9,10,18,10,10"
    
    strField = Split(strAllField, ",")
    intWidth = Split(strAllWidth, ",")
    
    strDateS = FCDate(MaskEdBox1) + 19110000
    strDateE = FCDate(MaskEdBox2) + 19110000
    strQ = "Select CP01,CP02,CP03,CP04,CP09,SubStr(ST02,1,3) as CP14N,SubStr(SqlDateW(CP05),1,10) CP05,CPM03,SubStr(SqlDateW(CP27),1,10) as CP27,CP27D,CP14 " & _
                "From Staff,CaseProperTyMap" & _
                ",(Select CP01,CP02,CP03,CP04,CP09,CP10,CP05,CP27,CP14,Count(*) CP27D From CaseProgress,WorkDay " & _
                        "Where CP01='FCT' And CP10 in ('1201','1202') And CP05>=" & strDateS & " " & _
                         "And WD01>=CP05 And WD01<=Decode(Nvl(CP27,0),0,To_Char(sysdate,'yyyymmdd'),Nvl(CP27,0)) " & _
                         "Group by CP01,CP02,CP03,CP04,CP09,CP10,CP05,CP27,CP14 " & _
                    "Union Select CP01,CP02,CP03,CP04,CP09,CP10,CP05,CP27,CP14,COUNT(*) CP27D From CaseProgress,WorkDay " & _
                        "Where CP01='FCT' And CP10 in ('1201','1202') And CP27>=" & strDateS & "  And CP27<=" & strDateE & " " & _
                        "And WD01>=CP05 And WD01<=Decode(Nvl(CP27,0),0,To_Char(sysdate,'yyyymmdd'),Nvl(CP27,0)) " & _
                        "Group by CP01,CP02,CP03,CP04,CP09,CP10,CP05,CP27,CP14 " & _
                " ) Where CP27D>15 And CP14=ST01(+) And CP01=CPM01(+) And CP10=CPM02(+) " & _
                   "Order By ST16,ST70,CP14,CP05,CP27D"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        RsQ.MoveFirst
        If SaveExcel = False Then
            Exit Sub
        ElseIf strFileName <> MsgText(601) Then
            ShellExecute hLocalFile, "open", strFileName, vbNullString, vbNullString, 1
        End If
    End If
    Exit Sub

ErrHnd:
    If Err.Number = 70 Then
        MsgBox ChgSQL(strFileName) & "檔案已開啟！", vbCritical
    Else
        MsgBox Err.Description, vbCritical
    End If
End Sub

Private Function SaveExcel() As Boolean
    Dim xlsApp As New Excel.Application
    Dim Wks As New Worksheet
    Dim intTitleRow As Integer, intFormat As String '抬頭列/格式:0-左/1-右/2-置中
    Dim strTemp(1) As String
    
On Error GoTo ErrHand1
    SaveExcel = False
    Call Pub_ChkExcelPath
    strFileName = "FCT審查報告來函超過15個工作天案件" & ACDate(ServerDate) & ServerTime
    If Dir(strExcelPath & strFileName & MsgText(43)) <> MsgText(601) Then
        Kill strExcelPath & strFileName & MsgText(43)
    End If
    If Dir(strExcelPath & strFileName & ".PDF") <> MsgText(601) Then
        Kill strExcelPath & strFileName & ".PDF"
    End If
    
    xlsApp.SheetsInNewWorkbook = 3 '工作表份數
    xlsApp.Workbooks.add
    Set Wks = xlsApp.Worksheets(1)
    'xlsApp.Visible = True
    
    intField = 65: intCounter = 1
    Call SetTitle(xlsApp, Wks)
    intTitleRow = intCounter
    
    intCounter = intCounter + 1
    Do While RsQ.EOF = False
        For i = LBound(strField) To UBound(strField)
            strTemp(0) = "": strTemp(1) = "": intFormat = 0 '靠左
            Select Case i
                Case GetValue("本所案號")
                    strTemp(0) = "" & RsQ.Fields("CP01") & "-" & RsQ.Fields("CP02") & "-" & RsQ.Fields("CP03") & "-" & RsQ.Fields("CP04")
                Case GetValue("總收文號")
                    strTemp(0) = "" & RsQ.Fields("CP09")
                Case GetValue("承辦人")
                    intFormat = 2 '置中
                    strTemp(0) = "" & RsQ.Fields("CP14N")
                Case GetValue("來函日")
                    intFormat = 2
                    strTemp(0) = "" & RsQ.Fields("CP05")
                    If strTemp(0) <> MsgText(601) Then
                        strTemp(1) = "yyyy/mm/dd"
                    End If
                Case GetValue("來函性質")
                    strTemp(0) = PUB_StrToStr_byVal("" & RsQ.Fields("CPM03"), 16)
                Case GetValue("發文日")
                    intFormat = 2
                    strTemp(0) = "" & RsQ.Fields("CP27")
                    If strTemp(0) <> MsgText(601) Then
                        strTemp(1) = "yyyy/mm/dd"
                    End If
                Case GetValue("工作天")
                    intFormat = 1 '靠右
                    strTemp(0) = "" & RsQ.Fields("CP27D")
            End Select
            
            '格式
            If strTemp(1) <> MsgText(601) Then
                Wks.Range(Chr(i + intField) & intCounter).NumberFormatLocal = strTemp(1)
            End If
            Wks.Range(Chr(i + intField) & intCounter).Value = strTemp(0)
            Select Case intFormat
                Case 0
                    Wks.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlLeft
                Case 1
                    Wks.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlRight
                Case 2
                    Wks.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter
            End Select
        Next i
        
        intCounter = intCounter + 1
        RsQ.MoveNext
    Loop
    
    Wks.PageSetup.PaperSize = xlPaperA4 'A4
    Wks.PageSetup.Orientation = xlPortrait '直印
    Wks.PageSetup.PrintTitleRows = "$1:$" & intTitleRow '標題列
    Wks.PageSetup.CenterFooter = "" & "&P"  '頁尾
    Wks.PageSetup.TopMargin = xlsApp.InchesToPoints(0.98) '2.5
    Wks.PageSetup.BottomMargin = xlsApp.InchesToPoints(0.51) '1
    Wks.PageSetup.HeaderMargin = xlsApp.InchesToPoints(0.39)
    Wks.PageSetup.LeftMargin = xlsApp.InchesToPoints(0.51) '1
    Wks.PageSetup.RightMargin = xlsApp.InchesToPoints(0.51) '1
    Wks.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
    
    '判斷版本2007
    If Val(xlsApp.Version) < 12 Then
        xlsApp.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strExcelPath & strFileName & ".pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        xlsApp.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName & ".xls", FileFormat:=56
    '版本2007以上
    Else
        xlsApp.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strExcelPath & strFileName & ".pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        xlsApp.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName & ".xls", FileFormat:=56
    End If
    xlsApp.Workbooks.Close
    xlsApp.Quit
    Kill strExcelPath & strFileName & ".xls"
    strFileName = strExcelPath & strFileName & ".pdf"
    
    SaveExcel = True
    Exit Function
    
ErrHand1:
    '判斷版本2007
    If Val(xlsApp.Version) < 12 Then
        xlsApp.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strExcelPath & strFileName & ".pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        xlsApp.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName & ".xls", FileFormat:=56
    '版本2007以上
    Else
        xlsApp.ActiveSheet.ExportAsFixedFormat Type:=0, FileName:=strExcelPath & strFileName & ".pdf", Quality:=0, IncludeDocProperties:=True, IgnorePrintAreas:=False, OpenAfterPublish:=False
        xlsApp.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName & ".xls", FileFormat:=56
    End If
    xlsApp.Workbooks.Close
    xlsApp.Quit
    Kill strExcelPath & strFileName & ".xls"
    If Err.Number <> 0 Then
        MsgBox "資料產生有誤(錯誤:" & Err.Description & ")", vbCritical
    End If
End Function

Private Function FormCheck() As Boolean
    Dim bCancel As Boolean
    
    FormCheck = False
    
    If MaskEdBox1 = MsgText(601) Or MaskEdBox1 = MsgText(29) Then
        MsgBox "來函起始日期不可空白！", , MsgText(5)
        MaskEdBox1.SetFocus
        Exit Function
    End If
    Call MaskEdBox1_Validate(bCancel)
    If bCancel = True Then
        Exit Function
    End If
    
    If MaskEdBox2 = MsgText(601) Or MaskEdBox2 = MsgText(29) Then
        MsgBox "來函迄止日期不可空白！", , MsgText(5)
        MaskEdBox2.SetFocus
        Exit Function
    End If
    Call MaskEdBox2_Validate(bCancel)
    If bCancel = True Then
        Exit Function
    End If
    If Val(FCDate(MaskEdBox1)) > Val(FCDate(MaskEdBox2)) Then
        MsgBox "來函迄止日期不可大於起日！", , MsgText(5)
        MaskEdBox2.SetFocus
        Exit Function
    End If
    FormCheck = True
End Function

Private Sub MaskEdBox1_Validate(Cancel As Boolean)
    Dim strDate As String
    If MaskEdBox1 = MsgText(601) Or MaskEdBox1 = MsgText(29) Then Exit Sub
    
    strDate = Format(DBDATE(MaskEdBox1), "####/##/##")
    If IsDate(strDate) = False Then
        MsgBox "來函起始日期格式錯誤！", , MsgText(5)
        Cancel = True
        Exit Sub
    End If
End Sub

Private Sub MaskEdBox2_Validate(Cancel As Boolean)
    Dim strDate As String
    If MaskEdBox2 = MsgText(601) Or MaskEdBox2 = MsgText(29) Then Exit Sub
    
    strDate = Format(DBDATE(MaskEdBox2), "####/##/##")
    If IsDate(strDate) = False Then
        MsgBox "來函迄止日期格式錯誤！", , MsgText(5)
        Cancel = True
        Exit Sub
    End If
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = 1 To UBound(strField)
       If UCase(strField(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Private Sub FormClear()
    MaskEdBox1.Mask = ""
    MaskEdBox1.Text = ""
    MaskEdBox1.Mask = DFormat
    MaskEdBox2.Mask = ""
    MaskEdBox2.Text = ""
    MaskEdBox2.Mask = DFormat
End Sub

Private Sub SetTitle(ByRef Xls As Excel.Application, ByRef Wks As Worksheet)
        
    '抬頭
    Wks.Range(Chr(intField) & intCounter).Value = "FCT審查報告來函超過15個工作天案件"
    Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strField) + intField) & intCounter).Select
    With Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strField) + intField) & intCounter)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
    End With
    intCounter = intCounter + 1
    
    '區間
    Wks.Range(Chr(intField) & intCounter).Value = "來函日期區間：" & MaskEdBox1 & " ~ " & MaskEdBox2
    Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strField) + intField) & intCounter).Select
    With Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strField) + intField) & intCounter)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
    End With
    intCounter = intCounter + 1
    
    '列印日期
    Wks.Range(Chr(intField) & intCounter).Value = "列印人員：" & StaffQuery(strUserNum)
    '列印人員
    Wks.Range(Chr(UBound(strField) - 1 + intField) & intCounter).Value = "列印日期：" & CFDate(ACDate(ServerDate))
    intCounter = intCounter + 1
    
    '說明
    Wks.Range(Chr(intField) & intCounter).Value = "PS：資料包含："
    intCounter = intCounter + 1
    Wks.Range(Chr(intField) & intCounter).Value = "1.  來函性質：1201審查報告、1202核駁前先行通知"
    intCounter = intCounter + 1
    Wks.Range(Chr(intField) & intCounter).Value = "2.  來函日期區間來函：已發文超過15個工作天、未發文但已超過15個工作天"
    intCounter = intCounter + 1
    Wks.Range(Chr(intField) & intCounter).Value = "3.  來函日期區間以前來函：於來函日期區間發文且超過15個工作天"
    intCounter = intCounter + 2
    
    For i = LBound(strField) To UBound(strField)
        Wks.Range(Chr(i + intField) & intCounter).Value = strField(i)
        Wks.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlLeft
        '格式
        If strField(i) = "承辦人" Or strField(i) = "來函日" Or strField(i) = "發文日" Then
            Wks.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter '置中
        ElseIf strField(i) = "工作天" Then
            Wks.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlRight '靠右
        End If
        '欄寬
        Wks.Columns(Chr(i + intField)).ColumnWidth = intWidth(i)
    Next i
    Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strField) + intField) & intCounter).Select
    Xls.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub
