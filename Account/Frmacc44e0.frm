VERSION 5.00
Begin VB.Form Frmacc44e0 
   AutoRedraw      =   -1  'True
   Caption         =   "資產負債比較表"
   ClientHeight    =   2265
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2265
   ScaleWidth      =   5160
   Begin VB.ComboBox CboCmp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1110
      TabIndex        =   0
      Top             =   150
      Width           =   3500
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
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
      Left            =   210
      Style           =   1  '圖片外觀
      TabIndex        =   10
      Top             =   3060
      Width           =   2346
   End
   Begin VB.CommandButton Cmd_Excel 
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
      Left            =   1380
      Style           =   1  '圖片外觀
      TabIndex        =   4
      Top             =   1740
      Width           =   2346
   End
   Begin VB.ComboBox Combo1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.25
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   345
      Left            =   1350
      Style           =   2  '單純下拉式
      TabIndex        =   8
      Top             =   2340
      Width           =   3450
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
      Height          =   300
      Left            =   2160
      TabIndex        =   3
      Top             =   1200
      Width           =   612
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
      Left            =   1110
      TabIndex        =   1
      Top             =   720
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
      Height          =   315
      Left            =   3480
      TabIndex        =   2
      Top             =   720
      Width           =   612
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
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
      Height          =   210
      Left            =   330
      TabIndex        =   11
      Top             =   210
      Width           =   675
   End
   Begin VB.Label Label10 
      BackStyle       =   0  '透明
      Caption         =   "印表機："
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
      Left            =   450
      TabIndex        =   9
      Top             =   2430
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "是否含子目(Y/N)"
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
      Left            =   330
      TabIndex        =   7
      Top             =   1230
      Width           =   1830
   End
   Begin VB.Label Label5 
      BackStyle       =   0  '透明
      Caption         =   "年度"
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
      Left            =   330
      TabIndex        =   6
      Top             =   750
      Width           =   615
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "截止月份"
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
      Left            =   2400
      TabIndex        =   5
      Top             =   750
      Width           =   975
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1680
      Visible         =   0   'False
      Width           =   132
   End
End
Attribute VB_Name = "Frmacc44e0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/12/1 memo by sonia 員工編號欄已修改
'Memo by Morgan2010/8/4 日期欄已修改
Option Explicit

Public adoacc040 As New ADODB.Recordset
Public adoacc010 As New ADODB.Recordset
Public adoaccrpt415 As New ADODB.Recordset
Dim lngCounter As Long
Dim douTotal1(5), douTotal2(5), douTotal3(5) As Double
Dim dllaccrpt415 As Object
Dim dou3222 As Double, douLast3222 As Double
Dim strPrinter As String 'Add By Sindy 2013/6/4
Dim strSql As String
Dim strSQL1 As String
'Add by Amy 2017/09/05
Dim i As Integer
Dim strF(), intWidth()
Dim intField As Integer, intCounter As Integer
Dim intTRow As Integer, intCol_R As Integer '抬頭列數/右邊抬頭欄起始
Dim bolAssetEnd As Boolean '資產資料結束
Dim strCmp As String, strCmpN As String 'Add by Sindy 2020/4/27


'Add by Sindy 2020/4/27
Private Sub SetCompN()
    strCmpN = "": strCmp = ""
    If Trim(CboCmp) <> MsgText(601) Then
        strCmp = CboCmp
        If InStr(strCmp, "　") > 0 Then
            strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
        End If
    End If
    strCmpN = GetAccReportCmpN(strCmp, True, True)
End Sub

Private Sub CboCmp_GotFocus()
    TextInverse CboCmp
End Sub

Private Sub CboCmp_KeyPress(KeyAscii As Integer)
    KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub CboCmp_Validate(Cancel As Boolean)
    Dim strCmp As String
    
    If Trim(CboCmp) = MsgText(601) Then Exit Sub
    
    strCmp = CboCmp
    If InStr(strCmp, "　") > 0 Then
        strCmp = Mid(strCmp, 1, Val(InStr(strCmp, "　")) - 1)
    End If
    If InStr(GetBookKeepCmp & 組合作帳公司 & ",", strCmp) = 0 Then
        MsgBox Label1 & MsgText(63), , MsgText(5)
        Cancel = True
        CboCmp.SetFocus
        Exit Sub
    ElseIf Len(Trim(CboCmp)) = 1 Then
        CboCmp = Trim(strCmp) & "　" & A0802Query(strCmp)
    End If
End Sub
'end 2020/4/27

'Add by Amy 2017/09/05 產生Excel
Private Sub Cmd_Excel_Click()
    If FormCheck = False Then
        Exit Sub
    End If
    
    Call SetCompN 'Add by Sindy 2020/4/27
    
    Screen.MousePointer = vbHourglass
    Accrpt415Delete
    ProduceData
   
    If adoaccrpt415.State = adStateOpen Then
        adoaccrpt415.Close
    End If
    
    adoaccrpt415.CursorLocation = adUseClient
    adoaccrpt415.Open "Select * From Accrpt415 Order by to_Number(R41502,'9999')", adoTaie, adOpenStatic, adLockReadOnly
    If adoaccrpt415.RecordCount <> 0 Then
        SaveExcel
    End If
    adoaccrpt415.Close
    Screen.MousePointer = vbDefault
    FormClear
    Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
End Sub

Private Sub SaveExcel()
    Dim xlsAgentPoint As New Excel.Application
    Dim wksrpt As New Worksheet
    Dim xlsFileName As String, strTmp(1) As String
    Dim intRow_L As Integer, intStart As Integer, intSum(1) As Integer  '左邊總列數/起始列/負債與業主權益加總列
    Dim bolSum As Boolean, bolTotal_R As Boolean
    Dim strStyle As String, strTotal_R(1 To 2) As String '儲存格格式/負債業主權益合計公式
    
    ReDim strF(17)
    ReDim intwith(17)
On Error GoTo ErrHand

    strF = Array("會計科目", "今年累計金額", "去年累計金額", "差異比率")
    intWidth = Array(19.25, 14, 14, 8.38) 'Modify by Amy 2018/01/16 調整欄位寬
    
'    strTmp(0) = "台一智權"
'    If Len(Trim(Text6)) > 0 Then
'        If Trim(Text6) = "1" Then
'            strTmp(0) = "台一"
'        Else
'            strTmp(0) = "智權"
'        End If
'    End If
'    xlsFileName = "資料負債比較表-" & strTmp(0) & ServerDate & MsgText(43)
    xlsFileName = "資料負債比較表-" & Replace(strCmpN, "/", "") & ServerDate & MsgText(43)
    If Dir(strExcelPath & xlsFileName) = MsgText(601) Then
       If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
       End If
    Else
       Kill strExcelPath & xlsFileName
    End If
    xlsAgentPoint.SheetsInNewWorkbook = 1 'Added by Lydia 2019/03/13 預設工作表數量
    xlsAgentPoint.Workbooks.add
    Set wksrpt = xlsAgentPoint.Worksheets(1)
    
    intCounter = 1: intField = 65: bolAssetEnd = False: bolSum = False: bolTotal_R = False
    Call SetField(wksrpt)
    intStart = intCounter
    
    adoaccrpt415.MoveFirst
    Do While adoaccrpt415.EOF = False
        If "" & adoaccrpt415.Fields("R41503") = "" Then
            '紙本列印換行用
        Else
            '資產
            If bolAssetEnd = False Then
                For i = LBound(strF) To UBound(strF)
                    strTmp(1) = "": strStyle = ""
                    strTmp(0) = i + intField
                    Select Case strTmp(0)
                        Case GetValue("會計科目")
                            strTmp(1) = "" & adoaccrpt415.Fields("R41503")
                            If InStr(strTmp(1), "資產合計") > 0 Then
                                intRow_L = intCounter '資產資料結束列
                                bolAssetEnd = True
                            End If
                            
                        Case GetValue("今年累計金額")
                            If IsNull(adoaccrpt415.Fields("R41504")) = False Then
                                strTmp(1) = Val("" & adoaccrpt415.Fields("R41504"))
                            End If
                            strStyle = "#,##0.00 ;[紅色]-#,##0.00"
                        Case GetValue("去年累計金額")
                            If IsNull(adoaccrpt415.Fields("R41505")) = False Then
                                strTmp(1) = Val("" & adoaccrpt415.Fields("R41505"))
                            End If
                            strStyle = "#,##0.00 ;[紅色]-#,##0.00"
                        Case GetValue("差異比率")
                            strTmp(1) = "=IF(" & Chr(GetValue("今年累計金額")) & intCounter & "<>""""," & Chr(GetValue("今年累計金額")) & intCounter & "/" & Chr(GetValue("去年累計金額")) & intCounter & "-1,"""")"
                            strStyle = "0.00%"
                    End Select
                    
                    If bolAssetEnd = False Then
                        wksrpt.Range(Chr(strTmp(0)) & intCounter).Value = strTmp(1)
                        If strStyle <> "" Then wksrpt.Range(Chr(strTmp(0)) & intCounter).NumberFormatLocal = strStyle
                    End If
                Next i
                intCounter = intCounter + 1
                If bolAssetEnd = True Then
                    intStart = intTRow + 1
                    intCounter = intStart
                End If
                
            '負債/業主權益
            Else
                For i = LBound(strF) To UBound(strF)
                    strTmp(1) = "": strStyle = ""
                    strTmp(0) = i + intCol_R
                    Select Case strTmp(0)
                        Case GetValue("會計科目", True)
                            strTmp(1) = "" & adoaccrpt415.Fields("R41503")
                            If InStr(strTmp(1), "負債小計") > 0 Or InStr(strTmp(1), "股東權益小計") > 0 Then
                                If InStr(strTmp(1), "負債小計") > 0 Then
                                    intSum(0) = intCounter
                                ElseIf InStr(strTmp(1), "股東權益小計") > 0 Then
                                    intSum(1) = intCounter
                                End If
                                bolSum = True
                            ElseIf InStr(strTmp(1), "負債與股東權益合計") > 0 Then
                                bolTotal_R = True
                            End If
                        Case GetValue("今年累計金額", True)
                            If IsNull(adoaccrpt415.Fields("R41504")) = False Then
                                strTmp(1) = Val("" & adoaccrpt415.Fields("R41504"))
                            End If
                            strStyle = "#,##0.00 ;[紅色]-#,##0.00"
                        Case GetValue("去年累計金額", True)
                            If IsNull(adoaccrpt415.Fields("R41505")) = False Then
                                strTmp(1) = Val("" & adoaccrpt415.Fields("R41505"))
                            End If
                            strStyle = "#,##0.00 ;[紅色]-#,##0.00"
                        Case GetValue("差異比率", True)
                            strTmp(1) = "=IF(" & Chr(GetValue("今年累計金額", True)) & intCounter & "<>""""," & Chr(GetValue("今年累計金額", True)) & intCounter & "/" & Chr(GetValue("去年累計金額", True)) & intCounter & "-1,"""")"
                            strStyle = "0.00%"
                    End Select
                    If bolTotal_R = True And Val(strTmp(0)) <> GetValue("會計科目", True) And Val(strTmp(0)) <> GetValue("差異比率", True) Then
                        strTotal_R(i) = "=" & Chr(strTmp(0)) & intSum(0) & "+" & Chr(strTmp(0)) & intSum(1)
                    ElseIf bolTotal_R = False And bolSum = True And Val(strTmp(0)) <> GetValue("會計科目", True) And Val(strTmp(0)) <> GetValue("差異比率", True) Then
                        strTmp(1) = "=Sum(" & Chr(strTmp(0)) & intStart & ":" & Chr(strTmp(0)) & intCounter - 1 & ")"
                    End If
                    If bolTotal_R = False Then
                        wksrpt.Range(Chr(strTmp(0)) & intCounter).Value = strTmp(1)
                        If strStyle <> "" Then wksrpt.Range(Chr(strTmp(0)) & intCounter).NumberFormatLocal = strStyle
                    End If
                Next i
                intCounter = intCounter + 1
                If bolSum = True Then
                    intCounter = intCounter + 1
                    intStart = intCounter
                    bolSum = False
                End If
            End If
            
            
        End If
        adoaccrpt415.MoveNext
    Loop
    
    '最後合計
    If intRow_L > intCounter Then
        intCounter = intRow_L + 2 '左邊資料比較多
    Else
        intCounter = intCounter + 2
    End If
    '左邊(資產)
    For i = LBound(strF) To UBound(strF)
        strStyle = ""
        If i + intField = GetValue("會計科目") Then
            strTmp(1) = ReportSum(22001)
        ElseIf i + intField = GetValue("差異比率") Then
            strTmp(1) = "=IF(" & Chr(GetValue("今年累計金額")) & intCounter & "<>""""," & Chr(GetValue("今年累計金額")) & intCounter & "/" & Chr(GetValue("去年累計金額")) & intCounter & "-1,"""")"
            strStyle = "0.00%"
        Else
            strTmp(1) = "=Sum(" & Chr(i + intField) & intTRow + 1 & ":" & Chr(i + intField) & intCounter - 1 & ")"
            strStyle = "#,##0.00 ;[紅色]-#,##0.00"
        End If
        wksrpt.Range(Chr(i + intField) & intCounter).Value = strTmp(1)
        If strStyle <> "" Then wksrpt.Range(Chr(i + intField) & intCounter).NumberFormatLocal = strStyle
    Next i
    '右邊
    For i = LBound(strF) To UBound(strF)
        strStyle = ""
        If i + intCol_R = GetValue("會計科目", True) Then
            strTmp(1) = ReportSum(23001)
        ElseIf i + intCol_R = GetValue("差異比率", True) Then
            strTmp(1) = "=IF(" & Chr(GetValue("今年累計金額", True)) & intCounter & "<>""""," & Chr(GetValue("今年累計金額", True)) & intCounter & "/" & Chr(GetValue("去年累計金額", True)) & intCounter & "-1,"""")"
            strStyle = "0.00%"
        Else
            strTmp(1) = strTotal_R(i)
            strStyle = "#,##0.00 ;[紅色]-#,##0.00"
        End If
        wksrpt.Range(Chr(i + intCol_R) & intCounter).Value = strTmp(1)
        If strStyle <> "" Then wksrpt.Range(Chr(i + intCol_R) & intCounter).NumberFormatLocal = strStyle
    Next i
    
    wksrpt.Range(Chr(intCol_R - 1) & ":" & Chr(intCol_R - 1)).ColumnWidth = 1
    wksrpt.Range(Chr(intField) & intTRow + 1 & ":" & Chr(intCol_R + UBound(strF)) & intCounter).Font.Size = 12
    'Add by Amy 2018/01/16 文字自動縮小(縮成符合欄位寬的大小)
    wksrpt.Range(Chr(intField) & intTRow + 1 & ":" & Chr(intCol_R + UBound(strF)) & intCounter).ShrinkToFit = True
    wksrpt.PageSetup.PaperSize = 9 '設定紙張 A4
    wksrpt.PageSetup.Orientation = xlPortrait 'Modify by Amy 2018/01/16 改直印-瑞婷
    wksrpt.PageSetup.PrintTitleRows = "$1:$" & intTRow
    'Modify by Amy 2018/01/16 左右邊界原0.5/加顯示格線/加縮放比列/加水平置
    wksrpt.PageSetup.LeftMargin = xlsAgentPoint.InchesToPoints(0) '左邊界
    wksrpt.PageSetup.RightMargin = xlsAgentPoint.InchesToPoints(0) '右邊界
    wksrpt.PageSetup.PrintGridlines = True '顯示格線
    wksrpt.PageSetup.Zoom = 75 '縮放比列
    wksrpt.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
    'end 2018/01/16
    
    '判斷版本
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set xlsAgentPoint = Nothing
    Set wksrpt = Nothing
    MsgBox "Excel檔案已產生！（檔案位置：" & strExcelPath & xlsFileName & "）"
    Exit Sub

ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(xlsAgentPoint.Version) < 12 Then
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=-4143
    Else
        xlsAgentPoint.Workbooks(1).SaveAs FileName:=strExcelPath & xlsFileName, FileFormat:=56
    End If
    xlsAgentPoint.Workbooks.Close
    xlsAgentPoint.Quit
    Set xlsAgentPoint = Nothing
    Set wksrpt = Nothing
End Sub

Private Sub SetField(ByRef Wks As Worksheet)
    Dim j As Integer, stTmp As String
    
    'Modify by Amy 2018/01/16 抬頭跨欄置中/字大小12-瑞婷
    Wks.Range(Chr(intField) & intCounter).Value = ReportTitle(415)
    Wks.Range(Chr(intField) & intCounter).Font.Size = 12
    Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strF) * 2 + 67) & intCounter).HorizontalAlignment = xlCenter
    Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strF) * 2 + 67) & intCounter).MergeCells = True
    intCounter = intCounter + 1
    Wks.Range(Chr(intField) & intCounter).Value = "公司別：" & strCmpN 'IIf(Text7 = "", "台一　專利商標/智權", Text7) Modify By Sindy 2020/4/27
    Wks.Range(Chr(intField) & intCounter).Font.Size = 12
    Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strF) * 2 + 67) & intCounter).HorizontalAlignment = xlCenter
    Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strF) * 2 + 67) & intCounter).MergeCells = True
    intCounter = intCounter + 1
    Wks.Range(Chr(intField) & intCounter).Value = "年度：" & Text3 & "　截止月份：" & Text1
    Wks.Range(Chr(intField) & intCounter).Font.Size = 12
    Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strF) * 2 + 67) & intCounter).HorizontalAlignment = xlCenter
    Wks.Range(Chr(intField) & intCounter & ":" & Chr(UBound(strF) * 2 + 67) & intCounter).MergeCells = True
    intCounter = intCounter + 1
    Wks.Range(Chr(intField) & intCounter).Value = "列印人員：" & StaffQuery(strUserNum)
    Wks.Range(Chr(intField) & intCounter).Font.Size = 12
    Wks.Range(Chr(intField + UBound(strF) * 2) & intCounter).Value = "列印日期：" & CFDate(ACDate(ServerDate))
    Wks.Range(Chr(intField + UBound(strF) * 2) & intCounter).Font.Size = 12
    intCounter = intCounter + 1
        
    '左方欄(資產)
    For i = LBound(strF) To UBound(strF)
        Wks.Range(Chr(i + intField) & intCounter).Value = strF(i)
        Wks.Range(Chr(i + intField) & intCounter).Font.Size = 12
        Wks.Columns(Chr(i + intField) & ":" & Chr(i + intField)).ColumnWidth = intWidth(i)
        Wks.Range(Chr(i + intField) & intCounter).HorizontalAlignment = xlCenter
    Next i
    
    intCol_R = UBound(strF) + intField + 2
    '右方欄(負債/業主權益)
    For i = LBound(strF) To UBound(strF)
        Wks.Range(Chr(intCol_R + i) & intCounter).Value = strF(i)
        Wks.Range(Chr(i + intField) & intCounter).Font.Size = 12
        Wks.Columns(Chr(intCol_R + i) & ":" & Chr(intCol_R + i)).ColumnWidth = intWidth(i)
        Wks.Range(Chr(intCol_R + i) & intCounter).HorizontalAlignment = xlCenter
    Next i
    'end 2018/01/16
    intTRow = intCounter
    intCounter = intCounter + 1
  
End Sub

Private Function GetValue(pFieldN As String, Optional ByVal bolCol_R As Boolean = False) As Integer
   Dim jj As Integer
 
    For jj = LBound(strF) To UBound(strF)
       If UCase(strF(jj)) = UCase(pFieldN) Then
          If bolCol_R = True Then
            GetValue = jj + intCol_R '右邊
          Else
            GetValue = jj + intField '左邊
          End If
          Exit For
       End If
    Next jj
End Function

'end 2016/09/05
'Modify By Sindy 2020/4/27 Mark.不使用 列印(&P) 按鈕
'Private Sub Command1_Click()
'   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
'      Exit Sub
'   End If
'   Screen.MousePointer = vbHourglass
'   Accrpt415Delete
'   ProduceData
'   PUB_SetOsDefaultPrinter Combo1  'Add By Sindy 2013/6/4
'   If adoaccrpt415.State = adStateOpen Then
'      adoaccrpt415.Close
'   End If
'   adoaccrpt415.CursorLocation = adUseClient
'   adoaccrpt415.Open "select * from accrpt415", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccrpt415.RecordCount <> 0 Then
'      '2014/2/20 modify by sonia
'      'dllaccrpt415.Acc44e0 ReportTitle(415), Text6, Text7, Text3, Text1, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'      dllaccrpt415.Acc44e0 ReportTitle(415), IIf(Text6 = "2", "J", Text6), IIf(Text7 = "", "台一　專利商標/智權", Text7), Text3, Text1, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'   End If
'   adoaccrpt415.Close
'   PUB_SetOsDefaultPrinter strPrinter  'Add By Sindy 2013/6/4
'   Screen.MousePointer = vbDefault
'   FormClear
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
'End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   Me.Icon = LoadPicture(strIcoPath)
   strFormName = Name
   Me.Width = 5250
   Me.Height = 2670 '3255 '2700
   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
   Image1 = LoadPicture(strBackPicPath4)
   sglWidth = Image1.Width
   sglHeight = Image1.Height
   For intX = 0 To Int(ScaleWidth / sglWidth)
       For intY = 0 To Int(ScaleHeight / sglHeight)
           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
       Next
   Next
   
   'Add by Sindy 2020/4/27 公司別改下拉
   CboCmp.AddItem "", 0
   Call Pub_SetCboCmp(CboCmp, True, False, False, , 1)
   'end 2020/4/27
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(101)
   Set dllaccrpt415 = CreateObject("AccReport.ReportSelect")
   PUB_SetPrinter Me.Name, Combo1, strPrinter 'Add By Sindy 2013/6/4
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   
   'Add By Sindy 2013/6/4
   If Me.Combo1.Text <> Me.Combo1.Tag Then
      PUB_UpdatePrintStartPoint strUserNum, Me.Name, Me.Combo1.Name, "0", "0", Me.Combo1.Text
   End If
   '2013/6/4 END

   Set dllaccrpt415 = Nothing
   Set Frmacc44e0 = Nothing
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

'Modify by Sindy 2020/4/27 公司別改下拉
'Private Sub Text6_Change()
'   '2014/2/20 modify by sonia
'   'If Text6 = MsgText(601) Then
'   '   Exit Sub
'   'End If
'   'Text7 = A0802Query(Text6)
'   Select Case Text6
'      Case "1"
'         Text7 = A0802Query(Text6)
'      Case "2"
'         Text7 = A0802Query("J")
'      Case ""
'         Text7 = "台一　專利商標/智權"
'   End Select
'   '2014/2/20 end
'End Sub
'
'Private Sub Text6_GotFocus()
'   TextInverse Text6
'End Sub
'
''2014/2/20 add by sonia
'Private Sub Text6_KeyPress(KeyAscii As Integer)
'   If KeyAscii <> 8 And KeyAscii <> Asc("1") And KeyAscii <> Asc("2") Then
'      KeyAscii = 0
'   End If
'End Sub
''2014/2/20 end

'*************************************************
'  產生報表資料
'
'*************************************************
Private Sub ProduceData()
Dim intCounter As Integer

On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   lngCounter = 0
   adoaccrpt415.CursorLocation = adUseClient
   adoaccrpt415.Open "select * from accrpt415", adoTaie, adOpenDynamic, adLockBatchOptimistic
'-------------------------------------------------
' 資產
'-------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '1' and a0101 < '2' AND A0101<>'1134' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/20 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '1' and a0101 < '2' AND A0101<>'1134' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   'modify by sonia 2018/8/22 取消AND A0101<>'1134'條件,一律用instr(a0102,'不用')=0
   'Modify By Sindy 2020/4/27
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '1' and a0101 < '2' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If strCmp <> "" Then
      If InStr(strCmp, "+") > 0 Then
         adoacc010.Open "select * from acc010 where a0101 >= '1' and a0101 < '2' and instr(a0102,'不用')=0 and (a0109 is null or a0109 In ('" & Replace(strCmp, "+", "','") & "')) order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
      Else
         adoacc010.Open "select * from acc010 where a0101 >= '1' and a0101 < '2' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & strCmp & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
      End If
   '2020/4/27 END
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '1' and a0101 < '2' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2014/2/20 end
   'end 2007/12/19
   Do While adoacc010.EOF = False
      If Text2 <> MsgText(602) Then
         If adoacc010.Fields("a0104").Value = "4" Then
            GoTo Check1
         End If
      End If
      Accrpt415Save
Check1:
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   PaintLine ReportSum(4)
   adoaccrpt415.UpdateBatch
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   'Modify By Cheng 2002/01/18
'   adoaccrpt415.Fields("r41503").Value = ReportSum(22)
   adoaccrpt415.Fields("r41503").Value = ReportSum(22001)
   Calculate "1", "199999"
   For intCounter = 3 To 5
      If IsNull(adoaccrpt415.Fields(intCounter).Value) Then
         douTotal1(intCounter) = 0
      Else
         douTotal1(intCounter) = Val(adoaccrpt415.Fields(intCounter).Value)
      End If
   Next intCounter
   adoaccrpt415.UpdateBatch
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   adoaccrpt415.UpdateBatch
'-------------------------------------------------
' 負債
'-------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '2' and a0101 < '3' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/20 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '2' and a0101 < '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2020/4/27
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '2' and a0101 < '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If strCmp <> "" Then
      If InStr(strCmp, "+") > 0 Then
         adoacc010.Open "select * from acc010 where a0101 >= '2' and a0101 < '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109 In ('" & Replace(strCmp, "+", "','") & "')) order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
      Else
         adoacc010.Open "select * from acc010 where a0101 >= '2' and a0101 < '3' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & strCmp & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
      End If
   '2020/4/27 END
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '2' and a0101 < '3' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2014/2/20 end
   'end 2007/12/19
   Do While adoacc010.EOF = False
      If Text2 <> MsgText(602) Then
         If adoacc010.Fields("a0104").Value = "4" Then
            GoTo Check2
         End If
      End If
      Accrpt415Save
Check2:
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   PaintLine ReportSum(4)
   adoaccrpt415.UpdateBatch
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   'Modify By Cheng 2002/01/18
'   adoaccrpt415.Fields("r41503").Value = ReportSum(10)
   adoaccrpt415.Fields("r41503").Value = ReportSum(10001)
   Calculate "2", "299999"
   For intCounter = 3 To 5
      If IsNull(adoaccrpt415.Fields(intCounter).Value) Then
         douTotal2(intCounter) = 0
      Else
         douTotal2(intCounter) = Val(adoaccrpt415.Fields(intCounter).Value)
      End If
   Next intCounter
   adoaccrpt415.UpdateBatch
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   adoaccrpt415.UpdateBatch
'-------------------------------------------------
' 股東權益
'-------------------------------------------------
   adoacc010.CursorLocation = adUseClient
   'Modify by Morgan 2007/12/19 科目有"不用"字眼的不顯示
   'adoacc010.Open "select * from acc010 where a0101 >= '3' and a0101 < '399999' order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   '2014/2/20 modify by sonia 加入a0109條件
   'adoacc010.Open "select * from acc010 where a0101 >= '3' and a0101 < '399999' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   'Modify By Sindy 2020/4/27
'   If Text6 <> "" Then
'      adoacc010.Open "select * from acc010 where a0101 >= '3' and a0101 < '399999' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & IIf(Text6 = "2", "J", Text6) & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   If strCmp <> "" Then
      If InStr(strCmp, "+") > 0 Then
         adoacc010.Open "select * from acc010 where a0101 >= '3' and a0101 < '399999' and instr(a0102,'不用')=0 and (a0109 is null or a0109 In ('" & Replace(strCmp, "+", "','") & "')) order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
      Else
         adoacc010.Open "select * from acc010 where a0101 >= '3' and a0101 < '399999' and instr(a0102,'不用')=0 and (a0109 is null or a0109='" & strCmp & "') order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
      End If
   '2020/4/27 END
   Else
      adoacc010.Open "select * from acc010 where a0101 >= '3' and a0101 < '399999' and instr(a0102,'不用')=0 order by a0101 asc", adoTaie, adOpenStatic, adLockReadOnly
   End If
   '2014/2/20 end
   'end 2007/12/19
   Do While adoacc010.EOF = False
      If Text2 <> MsgText(602) Then
         If adoacc010.Fields("a0104").Value = "4" Then
            GoTo Check3
         End If
      End If
      Accrpt415Save
Check3:
      adoacc010.MoveNext
   Loop
   adoacc010.Close
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   PaintLine ReportSum(4)
   adoaccrpt415.UpdateBatch
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   'Modify By Cheng 2002/01/18
'   adoaccrpt415.Fields("r41503").Value = ReportSum(12)
   adoaccrpt415.Fields("r41503").Value = ReportSum(12001)
   Calculate "3", "399999"
   For intCounter = 3 To 5
      If IsNull(adoaccrpt415.Fields(intCounter).Value) Then
         douTotal3(intCounter) = 0
      Else
         douTotal3(intCounter) = Val(adoaccrpt415.Fields(intCounter).Value)
      End If
   Next intCounter
   douTotal3(3) = douTotal3(3) + dou3222
   douTotal3(4) = douTotal3(4) + douLast3222
   douTotal3(5) = douTotal3(3) - douTotal3(4)
   If douTotal3(3) <> 0 Then
      adoaccrpt415.Fields("r41504").Value = douTotal3(3)
   End If
   If douTotal3(4) <> 0 Then
      adoaccrpt415.Fields("r41505").Value = douTotal3(4)
   End If
   If douTotal3(5) <> 0 Then
      adoaccrpt415.Fields("r41506").Value = douTotal3(5)
   End If
   adoaccrpt415.UpdateBatch
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   adoaccrpt415.UpdateBatch
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   PaintLine ReportSum(4)
   adoaccrpt415.UpdateBatch
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   'Modify By Cheng 2002/01/18
'   adoaccrpt415.Fields("r41503").Value = ReportSum(23)
   adoaccrpt415.Fields("r41503").Value = ReportSum(23001)
   For intCounter = 3 To 5
      If douTotal2(intCounter) + douTotal3(intCounter) = 0 Then
         adoaccrpt415.Fields(intCounter).Value = Null
      Else
         adoaccrpt415.Fields(intCounter).Value = douTotal2(intCounter) + douTotal3(intCounter)
      End If
   Next intCounter
   adoaccrpt415.UpdateBatch
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   PaintLine ReportSum(8)
   adoaccrpt415.UpdateBatch
   adoaccrpt415.Close
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt415Delete()
   adoTaie.Execute "delete from accrpt415"
End Sub

'*************************************************
'  儲存資料表(資產負債比較表暫存檔)
'
'*************************************************
Private Sub Accrpt415Save()
Dim intCounter As Integer
      
   adoaccrpt415.AddNew
   adoaccrpt415.Fields("r41501").Value = strUserNum
   adoaccrpt415.Fields("r41502").Value = Counter
   If IsNull(adoacc010.Fields("a0102").Value) Then
      adoaccrpt415.Fields("r41503").Value = Null
   Else
      adoaccrpt415.Fields("r41503").Value = adoacc010.Fields("a0102").Value
   End If
   Calculate adoacc010.Fields("a0101").Value, adoacc010.Fields("a0101").Value
   adoaccrpt415.UpdateBatch
End Sub

'*************************************************
'  計算累計金額
'
'*************************************************
Private Sub Calculate(strAccNo1 As String, strAccNo2 As String)
Dim douAmount As Double
Dim douLastAmount As Double

   strSql = "": strSQL1 = "" '2014/2/20 add by sonia
   
   If Text3 <> MsgText(601) Then
      strSql = " and a0401 = " & Val(Text3) & ""
      strSQL1 = " and a0401 = " & Val(Text3) & ""
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0402 = " & Val(Text1) & ""
      strSQL1 = strSQL1 & " and a0402 = " & Val(Text1) & ""
   End If
   
   'Modify By Sindy 2020/4/27
'   If Text6 <> MsgText(601) Then
'      '2014/2/20 modify by sonia
'      'strSql = strSql & " and a0403 = '" & Text6 & "'"
'      'strSQL1 = strSQL1 & " and a0403 = '" & Text6 & "'"
'      strSql = strSql & " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
'      strSQL1 = strSQL1 & " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
'      '2014/2/20 end
'   End If
   If strCmp <> MsgText(601) Then
      '2014/2/20 modify by sonia
      'strSql = strSql & " and a0403 = '" & Text6 & "'"
      'strSQL1 = strSQL1 & " and a0403 = '" & Text6 & "'"
      If InStr(strCmp, "+") > 0 Then
         strSql = strSql & " and a0403 In ('" & Replace(strCmp, "+", "','") & "')"
         strSQL1 = strSQL1 & " and a0403 In ('" & Replace(strCmp, "+", "','") & "')"
      Else
         strSql = strSql & " and a0403 = '" & strCmp & "'"
      strSQL1 = strSQL1 & " and a0403 = '" & strCmp & "'"
      End If
      '2014/2/20 end
   End If
   '2020/4/27 END
   
   If Text2 <> MsgText(602) Then
      If strAccNo1 <> MsgText(601) Then
         strSql = strSql & " and substr(a0405, 1, 4) >= '" & strAccNo1 & "'"
      End If
      If strAccNo2 <> MsgText(601) Then
         strSql = strSql & " and substr(a0405, 1, 4) <= '" & strAccNo2 & "'"
      End If
   Else
      If strAccNo1 <> MsgText(601) Then
         strSql = strSql & " and a0405 >= '" & strAccNo1 & "'"
      End If
      If strAccNo2 <> MsgText(601) Then
         strSql = strSql & " and a0405 <= '" & strAccNo2 & "'"
      End If
   End If
'   If strSQL <> MsgText(601) Then
'      strSQL = " where " & Mid(strSQL, 5, Len(strSQL) - 4)
'   End If
   If strAccNo1 = "3222" Then
      adoacc040.CursorLocation = adUseClient
      'modify by sonia 2016/8/8 於2007/5已分71科目,72科目,此程式未修改
      'adoacc040.Open "select sum(decode(substr(a0405, 1, 1), '4', a0408, '6', a0408 * -1, '7', a0408)) from acc040, acc010 where a0405 = a0101 and a0405 >= '4' and a0405 < '8' and a0404 = '" & MsgText(55) & "'" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
      adoacc040.Open "select nvl(sum(decode(substr(a0405, 1, 1), '4', a0408, '6', a0408 * -1, DECODE(SUBSTR(A0405,1,2),'71',a0408,A0408*-1))),0) from acc040, acc010 where a0405 = a0101 and a0405 >= '4' and a0405 < '8' and a0404 = '" & MsgText(55) & "'" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc040.RecordCount <> 0 Then
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt415.Fields("r41504").Value = Null
            douAmount = 0
            dou3222 = 0
         Else
            If Val(Text1) = 12 Then
               adoaccrpt415.Fields("r41504").Value = Null
               douAmount = 0
               dou3222 = 0
            Else
               adoaccrpt415.Fields("r41504").Value = Val(Format(adoacc040.Fields(0).Value, FAmount))
               douAmount = Val(Format(adoacc040.Fields(0).Value, FAmount))
               dou3222 = Val(Format(adoacc040.Fields(0).Value, FAmount))
            End If
         End If
      Else
         adoaccrpt415.Fields("r41504").Value = Null
         douAmount = 0
         dou3222 = 0
      End If
      adoacc040.Close
   Else
      adoacc040.CursorLocation = adUseClient
      adoacc040.Open "select sum(a0408) from acc040 where a0405 <> '3222' and a0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc040.RecordCount <> 0 Then
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt415.Fields("r41504").Value = Null
            douAmount = 0
         Else
            adoaccrpt415.Fields("r41504").Value = Val(Format(adoacc040.Fields(0).Value, FAmount))
            douAmount = Val(Format(adoacc040.Fields(0).Value, FAmount))
         End If
      Else
         adoaccrpt415.Fields("r41504").Value = Null
         douAmount = 0
      End If
      adoacc040.Close
   End If
   
   strSql = MsgText(601)
   strSQL1 = MsgText(601)
   If Text3 <> MsgText(601) Then
      strSql = " and a0401 = " & Val(Text3) - 1 & ""
      strSQL1 = " and a0401 = " & Val(Text3) - 1 & ""
   End If
   If Text1 <> MsgText(601) Then
      strSql = strSql & " and a0402 = " & Val(Text1) & ""
      strSQL1 = strSQL1 & " and a0402 = " & Val(Text1) & ""
   End If
   
   'Modify By Sindy 2020/4/27
'   If Text6 <> MsgText(601) Then
'      '2014/2/20 modify by sonia
'      'strSql = strSql & " and a0403 = '" & Text6 & "'"
'      'strSQL1 = strSQL1 & " and a0403 = '" & Text6 & "'"
'      strSql = strSql & " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
'      strSQL1 = strSQL1 & " and a0403 = '" & IIf(Text6 = "2", "J", "1") & "'"
'      '2014/2/20 end
'   End If
   If strCmp <> MsgText(601) Then
      '2014/2/20 modify by sonia
      'strSql = strSql & " and a0403 = '" & Text6 & "'"
      'strSQL1 = strSQL1 & " and a0403 = '" & Text6 & "'"
      If InStr(strCmp, "+") > 0 Then
         strSql = strSql & " and a0403 In ('" & Replace(strCmp, "+", "','") & "')"
         strSQL1 = strSQL1 & " and a0403 In ('" & Replace(strCmp, "+", "','") & "')"
      Else
         strSql = strSql & " and a0403 = '" & strCmp & "'"
         strSQL1 = strSQL1 & " and a0403 = '" & strCmp & "'"
      End If
      '2014/2/20 end
   End If
   '2020/4/27 END
   
   If Text2 <> MsgText(602) Then
      If strAccNo1 <> MsgText(601) Then
         strSql = strSql & " and substr(a0405, 1, 4) >= '" & strAccNo1 & "'"
      End If
      If strAccNo2 <> MsgText(601) Then
         strSql = strSql & " and substr(a0405, 1, 4) <= '" & strAccNo2 & "'"
      End If
   Else
      If strAccNo1 <> MsgText(601) Then
         strSql = strSql & " and a0405 >= '" & strAccNo1 & "'"
      End If
      If strAccNo2 <> MsgText(601) Then
         strSql = strSql & " and a0405 <= '" & strAccNo2 & "'"
      End If
   End If
'   If strSQL <> MsgText(601) Then
'      strSQL = " where " & Mid(strSQL, 5, Len(strSQL) - 4)
'   End If
   If strAccNo1 = "3222" Then
      adoacc040.CursorLocation = adUseClient
      '2014/2/20 modify by sonia J公司前一年無資料程式會錯
      'adoacc040.Open "select sum(decode(substr(a0101, 1, 1), '4', a0408, '6', a0408 * -1, '7', a0408)) from acc040, acc010 where a0405 = a0101 and a0405 >= '4' and a0405 < '8' and a0405 <> '3222' and a0404 = '" & MsgText(55) & "'" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
      'modify by sonia 2016/8/8 於2007/5已分71科目,72科目,此程式未修改
      'adoacc040.Open "select nvl(sum(decode(substr(a0101, 1, 1), '4', a0408, '6', a0408 * -1, '7', a0408)),0) from acc040, acc010 where a0405 = a0101 and a0405 >= '4' and a0405 < '8' and a0405 <> '3222' and a0404 = '" & MsgText(55) & "'" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
      adoacc040.Open "select nvl(sum(decode(substr(a0405, 1, 1), '4', a0408, '6', a0408 * -1, DECODE(SUBSTR(A0405,1,2),'71',a0408,A0408*-1))),0) from acc040, acc010 where a0405 = a0101 and a0405 >= '4' and a0405 < '8' and a0405 <> '3222' and a0404 = '" & MsgText(55) & "'" & strSQL1, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc040.RecordCount <> 0 Then
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt415.Fields("r41505").Value = Null
            douLastAmount = 0
            douLast3222 = 0
         Else
            If Val(Text1) = 12 Then
               adoaccrpt415.Fields("r41505").Value = Null
               douLastAmount = 0
               douLast3222 = 0
            Else
               adoaccrpt415.Fields("r41505").Value = Val(Format(adoacc040.Fields(0).Value, FAmount))
               douLastAmount = Val(Format(adoacc040.Fields(0).Value, FAmount))
               douLast3222 = Val(Format(adoacc040.Fields(0).Value, FAmount))
            End If
         End If
      Else
         adoaccrpt415.Fields("r41505").Value = Null
         douLastAmount = 0
         douLast3222 = 0
      End If
      If douAmount - douLastAmount = 0 Then
         adoaccrpt415.Fields("r41506").Value = Null
      Else
         adoaccrpt415.Fields("r41506").Value = douAmount - douLastAmount
      End If
      adoacc040.Close
   Else
      adoacc040.CursorLocation = adUseClient
      '2010/1/11 MODIFY BY SONIA 只抓TOT部門
      'adoacc040.Open "select sum(a0408) from acc040 where a0405 <> '3222'" & strSQL, adoTaie, adOpenStatic, adLockReadOnly
      '2014/2/20 modify by sonia J公司前一年無資料程式會錯
      'adoacc040.Open "select sum(a0408) from acc040 where a0405 <> '3222' and a0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      adoacc040.Open "select NVL(sum(a0408),0) from acc040 where a0405 <> '3222' and a0404 = '" & MsgText(55) & "'" & strSql, adoTaie, adOpenStatic, adLockReadOnly
      If adoacc040.RecordCount <> 0 Then
         If IsNull(adoacc040.Fields(0).Value) Or adoacc040.Fields(0).Value = 0 Then
            adoaccrpt415.Fields("r41505").Value = Null
            douLastAmount = 0
         Else
            adoaccrpt415.Fields("r41505").Value = Val(Format(adoacc040.Fields(0).Value, FAmount))
            douLastAmount = Val(Format(adoacc040.Fields(0).Value, FAmount))
         End If
      Else
         adoaccrpt415.Fields("r41505").Value = Null
         douLastAmount = 0
      End If
      If douAmount - douLastAmount = 0 Then
         adoaccrpt415.Fields("r41506").Value = Null
      Else
         adoaccrpt415.Fields("r41506").Value = douAmount - douLastAmount
      End If
      adoacc040.Close
   End If
End Sub

'*************************************************
'  畫線
'
'*************************************************
Private Sub PaintLine(strSign As String)
Dim intCounter As Integer

   For intCounter = 3 To 5
      adoaccrpt415.Fields(intCounter).Value = strSign
   Next intCounter
End Sub

'*************************************************
'  計數器
'
'*************************************************
Private Function Counter() As Long
   lngCounter = lngCounter + 1
   Counter = lngCounter
End Function

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
'   Text6 = ""
   Text3 = ""
   Text1 = ""
   Text2 = ""
'   Text6.SetFocus
   'Add By Sindy 2020/4/27
   CboCmp.ListIndex = -1
   CboCmp.SetFocus
   '2020/4/27 END
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
Dim bolCancel As Boolean

   'Add by Sindy 2020/4/27 +公司別判斷
   If Trim(CboCmp) <> MsgText(601) Then
      Call CboCmp_Validate(bolCancel)
      If bolCancel = True Then
          Exit Function
      End If
   End If
   'end 2020/4/27
   
   If Text3 = MsgText(601) Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Function
   End If
   If Text1 = MsgText(601) Then
      MsgBox MsgText(181), , MsgText(5)
      Exit Function
   End If
   
   FormCheck = True
End Function
