VERSION 5.00
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Begin VB.Form Frmacc14a0 
   AutoRedraw      =   -1  'True
   Caption         =   "暫收款明細表"
   ClientHeight    =   3060
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   5160
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   3060
   ScaleWidth      =   5160
   Begin VB.ComboBox CboComp 
      BeginProperty Font 
         Name            =   "標楷體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   324
      Left            =   1344
      TabIndex        =   0
      Text            =   "CboComp"
      Top             =   264
      Width           =   3520
   End
   Begin VB.CommandButton CmdExcel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Excel"
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
      Left            =   264
      Style           =   1  '圖片外觀
      TabIndex        =   15
      Top             =   2352
      Width           =   4692
   End
   Begin VB.TextBox Text5 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1344
      MaxLength       =   1
      TabIndex        =   6
      Top             =   1788
      Width           =   465
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "列印(&P)"
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
      Left            =   2376
      Style           =   1  '圖片外觀
      TabIndex        =   7
      Top             =   2520
      Visible         =   0   'False
      Width           =   4692
   End
   Begin VB.TextBox Text3 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   1344
      TabIndex        =   5
      Top             =   1368
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   1344
      TabIndex        =   3
      Top             =   1008
      Width           =   1572
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   3264
      TabIndex        =   4
      Top             =   1008
      Width           =   1572
   End
   Begin MSMask.MaskEdBox MaskEdBox1 
      Height          =   300
      Left            =   1344
      TabIndex        =   1
      Top             =   648
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
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
      Left            =   3264
      TabIndex        =   2
      Top             =   648
      Width           =   1572
      _ExtentX        =   2773
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      PromptChar      =   "_"
   End
   Begin VB.Label Label4 
      BackStyle       =   0  '透明
      Caption         =   "公司別"
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
      Left            =   408
      TabIndex        =   16
      Top             =   324
      Width           =   732
   End
   Begin VB.Label lblSalesName 
      BackStyle       =   0  '透明
      Caption         =   "智權人員名稱"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   10.8
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Left            =   2316
      TabIndex        =   14
      Top             =   1404
      Width           =   1356
   End
   Begin VB.Label Label6 
      BackStyle       =   0  '透明
      Caption         =   "沖銷狀況         (1.已沖 2.未沖)"
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
      Left            =   384
      TabIndex        =   13
      Top             =   1788
      Width           =   4488
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label3 
      BackStyle       =   0  '透明
      Caption         =   "智權人員"
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
      Left            =   384
      TabIndex        =   12
      Top             =   1368
      Width           =   972
   End
   Begin VB.Label Label5 
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
      Left            =   3024
      TabIndex        =   11
      Top             =   648
      Width           =   252
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "輸入日期"
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
      Left            =   384
      TabIndex        =   10
      Top             =   648
      Width           =   972
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "客戶編號"
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
      Left            =   384
      TabIndex        =   9
      Top             =   1008
      Width           =   972
   End
   Begin VB.Label Label7 
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
      Left            =   3024
      TabIndex        =   8
      Top             =   1008
      Width           =   252
   End
End
Attribute VB_Name = "Frmacc14a0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Memo By Sonia 2012/12/4 智權人員欄已修改
'2010/11/30 memo by sonia 員工編號欄已修改
'Memo by Morgan 2010/7/27 日期欄已修改
Option Explicit
Public adoacc0t0 As New ADODB.Recordset
Public adoaccrpt110 As New ADODB.Recordset
Dim dllaccrpt110 As Object
'Add by Amy 2022/08/09
Const lngMaxR As Long = 65534 '工作表最大列數
Dim strWkName As String, strQ As String, strFileName As String
Dim i As Integer, intField As Integer, intWksNo As Integer, intTitleR As Integer, lngR As Long
Dim strFieldN, intWidth

'Add by Amy 2022/08/09
Private Sub cmdExcel_Click()
    Dim strMsg As String
    
    If FormCheck = False Then
        MsgBox MsgText(181), , MsgText(5)
        Exit Sub
    End If
    Screen.MousePointer = vbHourglass
    Accrpt110Delete
    'Modified by Lydia 2024/11/28
    'ProduceData
    ProduceData_New
    
    'Modified by Lydia 2024/11/28
    'strQ = "Select * From Accrpt110 Where r11001='" & strUserNum & "'"
    strQ = "Select * From Accrpt110 Where r11001='" & strUserNum & "' order by r11002 asc , r11003 asc "
    If adoaccrpt110.State = adStateOpen Then adoaccrpt110.Close
    adoaccrpt110.CursorLocation = adUseClient
    adoaccrpt110.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoaccrpt110.RecordCount <> 0 Then
        If SaveExcel(strMsg) = False Then
            adoaccrpt110.Close
            Screen.MousePointer = vbDefault
            Exit Sub
        Else
            If strMsg <> MsgText(601) Then strMsg = strMsg & vbCrLf
            strMsg = strMsg & "資料已產生！" & "檔案存於 " & strExcelPath
            MsgBox strMsg
        End If
    End If
    adoaccrpt110.Close
    Screen.MousePointer = vbDefault
    FormClear
End Sub

Private Function SaveExcel(ByRef stMsg As String) As Boolean
    Dim Xls As New Excel.Application, Wks As New Worksheet
    Dim intAlign As Integer, strFormat As String, strVal As String, lngDataR As Long
On Error GoTo ErrHand
    
    'Modified by Lydia 2024/11/28
    'ReDim strFieldN(6)
    'ReDim intWidth(6)
    'strFieldN = Array("輸入日期", "暫收款單號", "客戶名稱", "智權人員", "暫收款金額", "處理方式", "處理日期")
    'intWidth = Array(10, 13, 60, 12, 13, 15, 10)
    If Text5 = "1" Then
       ReDim strFieldN(8)
       ReDim intWidth(8)
       strFieldN = Array("輸入日期", "暫收款單號", "客戶名稱", "智權人員", "暫收款金額", "處理方式", "處理日期", "公司別", "沖帳日期")
       intWidth = Array(10, 13, 60, 12, 13, 15, 10, 10, 10)
    Else
       ReDim strFieldN(7)
       ReDim intWidth(7)
       strFieldN = Array("輸入日期", "暫收款單號", "客戶名稱", "智權人員", "暫收款金額", "處理方式", "處理日期", "公司別")
       intWidth = Array(10, 13, 60, 12, 13, 15, 10, 10)
    End If
    'end 2024/11/28
    
    lngDataR = 1: lngR = 1: intField = 65: strWkName = ""
    strFileName = "暫收款明細表" & ServerDate & ServerTime & MsgText(43)
    If Dir(strExcelPath & strFileName) = MsgText(601) Then
        If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
            MkDir strExcelPath
        End If
    Else
        Kill strExcelPath & strFileName
    End If
    
    Xls.Visible = False
    Xls.SheetsInNewWorkbook = 3 '改設定(選項->一般->包括的工作表份數)
    Xls.Workbooks.add
    If strWkName = MsgText(601) Then
        strWkName = Left(Xls.Worksheets(1).Name, Len(Xls.Worksheets(1).Name) - 1)
        intWksNo = Right(Xls.Worksheets(1).Name, 1)
    End If
    Set Wks = Xls.Worksheets(strWkName & intWksNo)
    With adoaccrpt110
        Do While .EOF = False
            '超過一個工作表列數
            If lngDataR > lngMaxR And lngDataR Mod lngMaxR = 1 Then
                Call SetEndWks(Xls, Wks, 2)
                intWksNo = intWksNo + 1
                If intWksNo > 3 Then Xls.Worksheets.add After:=Wks
                Set Wks = Xls.Worksheets(strWkName & intWksNo)
                Wks.Activate
                lngR = 1
            End If
            If lngR = 1 Then
                Call SetTitle(Xls, Wks)
                intTitleR = lngR: lngR = lngR + 1
            End If
            
            For i = LBound(strFieldN) To UBound(strFieldN)
                strFormat = "": intAlign = 1 '靠左
                strVal = "" & .Fields(i + 1)
                Select Case i
                    'Modified by Lydia 2024/11/28 +GetValue("沖帳日期")
                    Case GetValue("輸入日期"), GetValue("處理日期"), GetValue("沖帳日期")
                        strVal = ChangeTStringToTDateString(strVal)
                    Case GetValue("暫收款金額")
                        strFormat = "#,##0"
                        intAlign = 2 '靠右
                End Select
                
                Wks.Range(Chr(intField + i) & lngR).Value = strVal
                '儲存格格式
                If strFormat <> MsgText(601) Then
                    Wks.Range(Chr(intField + i) & lngR).NumberFormatLocal = strFormat
                End If
                
                '對齊
                If intAlign = 1 Then
                    Wks.Range(Chr(intField + i) & lngR).HorizontalAlignment = xlLeft
                ElseIf intAlign = 2 Then
                    Wks.Range(Chr(intField + i) & lngR).HorizontalAlignment = xlRight
                Else
                    Wks.Range(Chr(intField + i) & lngR).HorizontalAlignment = xlCenter
                End If
            Next i
            lngR = lngR + 1
            lngDataR = lngDataR + 1
            .MoveNext
        Loop
    End With
    If intWksNo = 1 Then
        Call SetEndWks(Xls, Wks, 0, intWksNo)
    Else
        stMsg = "資料量過多所以區分不同工作表，請注意！"
        For i = intWksNo To 1 Step -1
            If i = intWksNo Then
                Call SetEndWks(Xls, Wks, 0, i)
            Else
                Set Wks = Xls.Worksheets(strWkName & i)
                Call SetEndWks(Xls, Wks, 1, i)
            End If
        Next i
    End If
    
    If Val(Xls.Version) < 12 Then
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    Xls.Workbooks.Close
    Xls.Quit
    
    SaveExcel = True
    Exit Function
    
ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If Val(Xls.Version) < 12 Then
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=-4143
    Else
        Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileName, FileFormat:=56
    End If
    Xls.Workbooks.Close
    Xls.Quit
End Function

Private Sub SetTitle(Xls As Excel.Application, ByRef Wks As Worksheet)
    Dim strTmp As String, j As Integer
    Dim intSpace As Single '補空白
    
    intSpace = 9
    'Modified by Lydia 2024/11/28
    'strTmp = "已沖"
    'If Text5 = "2" Then strTmp = "未沖"
    If Text5 = "1" Then
       strTmp = "已沖"
    Else
       strTmp = "未沖"
    End If
    'end 2024/11/28
    
    Wks.Range(Chr(intField) & lngR).Font.Size = 20
    Wks.Range(Chr(intField) & lngR).Font.Bold = True
    Wks.Range(Chr(intField) & lngR).Value = "***　暫收款明細表　*** (" & strTmp & ")"
    Wks.Range(Chr(intField) & lngR & ":" & Chr(intField + UBound(strFieldN)) & lngR).HorizontalAlignment = xlCenter
    Wks.Range(Chr(intField) & lngR & ":" & Chr(intField + UBound(strFieldN)) & lngR).MergeCells = True
    'Added by Lydia 2024/11/28
    lngR = lngR + 1: strTmp = " "
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).Font.Size = 10
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).Font.Name = "標楷體"
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).Value = "公司別：" & IIf(cboComp.Text <> "", cboComp.Text, String(22, " "))
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).HorizontalAlignment = xlCenter
    'end 2024/11/28
    
    lngR = lngR + 1: strTmp = " "
    If MaskEdBox1.Text <> MsgText(29) And MaskEdBox1.Text <> MsgText(601) Then
        strTmp = PUB_StrToStr(MaskEdBox1.Text, intSpace, True)
    End If
    strTmp = PUB_StrToStr(strTmp, intSpace, True)
    If MaskEdBox2.Text <> MsgText(29) And MaskEdBox2.Text <> MsgText(601) Then
        If strTmp = MsgText(601) Then strTmp = PUB_StrToStr(" ", intSpace, True)
        strTmp = strTmp & "~ " & PUB_StrToStr(MaskEdBox2.Text, intSpace, True)
    Else
        strTmp = strTmp & "~ " & PUB_StrToStr(" ", intSpace, True)
    End If
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).Font.Size = 10
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).Font.Name = "標楷體"
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).Value = "輸入日期：" & strTmp
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).HorizontalAlignment = xlCenter
    
    lngR = lngR + 1: strTmp = " "
    If Trim(Text1.Text) <> MsgText(601) And Text1.Text <> "X" Then
        strTmp = Text1.Text
    End If
    strTmp = PUB_StrToStr(strTmp, intSpace, True)
    If Trim(Text2.Text) <> MsgText(601) And Text2.Text <> "X" Then
        If strTmp = MsgText(601) Then strTmp = PUB_StrToStr(" ", intSpace, True)
        strTmp = strTmp & "~ " & PUB_StrToStr(Text2.Text, intSpace, True)
    Else
        strTmp = strTmp & "~ " & PUB_StrToStr(" ", intSpace, True)
    End If
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).Font.Size = 10
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).Font.Name = "標楷體"
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).Value = "客戶編號：" & strTmp
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).HorizontalAlignment = xlCenter
    
    lngR = lngR + 1: strTmp = " "
    If Trim(Text3.Text) = " " Then
        strTmp = Text3.Text & "(" & lblSalesName & ")"
    End If
    strTmp = PUB_StrToStr(strTmp, (intSpace + 1) * 2, True)
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).Font.Size = 10
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).Font.Name = "標楷體"
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).Value = "智權人員：" & strTmp
    Wks.Range(Chr(intField + GetValue("客戶名稱")) & lngR).HorizontalAlignment = xlCenter
    
    lngR = lngR + 1
    Wks.Range(Chr(intField) & lngR).Font.Size = 12
    Wks.Range(Chr(intField) & lngR).Value = "列印人員：" & StaffQuery(strUserNum)
    Wks.Range(Chr(intField + UBound(strFieldN)) & lngR).HorizontalAlignment = xlLeft
    Wks.Range(Chr(intField + UBound(strFieldN) - 1) & lngR).Font.Size = 12
    Wks.Range(Chr(intField + UBound(strFieldN) - 1) & lngR).Value = "列印日期：" & CFDate(ACDate(ServerDate))
    Wks.Range(Chr(intField + UBound(strFieldN) - 1) & lngR).HorizontalAlignment = xlLeft
    
    lngR = lngR + 1
    For j = LBound(strFieldN) To UBound(strFieldN)
        Wks.Range(Chr(intField + j) & lngR).Value = strFieldN(j)
        Wks.Columns(Chr(intField + j)).ColumnWidth = intWidth(j)
    Next j
    Wks.Range(Chr(intField) & lngR & ":" & Chr(intField + UBound(strFieldN)) & lngR).Font.Size = 12
    Wks.Range(Chr(intField) & lngR & ":" & Chr(intField + UBound(strFieldN)) & lngR).Font.Bold = True
    Wks.Range(Chr(intField) & lngR & ":" & Chr(intField + UBound(strFieldN)) & lngR).HorizontalAlignment = xlCenter
    '畫線
    Wks.Range(Chr(intField) & lngR & ":" & Chr(intField + UBound(strFieldN)) & lngR).Select
    Xls.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
    
    For jj = 1 To UBound(strFieldN)
        If UCase(strFieldN(jj)) = UCase(pFieldN) Then
            GetValue = jj
            Exit For
        End If
    Next jj
End Function
'intChoose:0-最後設定/1-只設定頁碼/2-先不設定頁碼
Private Sub SetEndWks(Xls As Excel.Application, ByRef Wks As Worksheet, ByVal intChoose As Integer, Optional ByVal intNowWks As Integer)
    If intChoose = 0 Then
        '畫線
        Wks.Range(Chr(intField + GetValue("暫收款金額") - 1) & lngR).Font.Bold = True
        Wks.Range(Chr(intField) & lngR - 1 & ":" & Chr(intField + UBound(strFieldN)) & lngR - 1).Select
        Xls.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        '合計
        Wks.Range(Chr(intField + GetValue("暫收款金額") - 1) & lngR).Value = "合　計"
        Wks.Range(Chr(intField + GetValue("暫收款金額") - 1) & lngR).HorizontalAlignment = xlCenter
        Wks.Range(Chr(intField + GetValue("暫收款金額")) & lngR).Value = _
                "=Sum(" & Chr(intField + GetValue("暫收款金額")) & intTitleR + 1 & ":" & Chr(intField + GetValue("暫收款金額")) & lngR - 1 & ")"
        Wks.Range(Chr(intField + GetValue("暫收款金額")) & lngR).HorizontalAlignment = xlRight
     End If
     If intChoose <> 1 Then
        Wks.PageSetup.PaperSize = xlPaperA4
        Wks.PageSetup.Orientation = xlLandscape '橫印
        Wks.PageSetup.PrintTitleRows = "$1:$" & intTitleR
        Wks.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
        Wks.PageSetup.LeftMargin = Xls.InchesToPoints(0.4) '邊界
        Wks.PageSetup.RightMargin = Xls.InchesToPoints(0.4)
        Wks.PageSetup.TopMargin = Xls.InchesToPoints(0.4)
        Wks.PageSetup.BottomMargin = Xls.InchesToPoints(0.8)
    End If
    If intChoose = 0 Or intChoose = 1 Then
        'Memo 工作表最大頁數=65534/31(A4 一頁顯示31筆)=2114(一個工作表最大頁數)
        strExc(1) = "第 &P+" & (intNowWks - 1) * 2114 & " 頁，共 &N+" & (intWksNo - 1) * 2114 & "頁"
        Wks.PageSetup.CenterFooter = strExc(1)
    End If
    
End Sub
'end 2022/08/09

'Mark by Amy 2022/08/09 不使用ACCRPT
Private Sub Command1_Click()
'   If FormCheck = False Then
'      MsgBox MsgText(181), , MsgText(5)
'      Exit Sub
'   End If
'   Screen.MousePointer = vbHourglass
'   Accrpt110Delete
'   ProduceData
'   If adoaccrpt110.State = adStateOpen Then
'      adoaccrpt110.Close
'   End If
'   adoaccrpt110.CursorLocation = adUseClient
'   adoaccrpt110.Open "select * from accrpt110 where r11001='" & strUserNum & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccrpt110.RecordCount <> 0 Then
'        'Modify By Cheng 2003/05/05
''      dllaccrpt110.Acc14a0 ReportTitle(110), MaskEdBox1.Text, MaskEdBox2.Text, Text1.Text, Text2.Text, Text3.Text, Text4.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'      'Modify by Morgan 2007/10/2 智權人員範圍改成一個
'      'dllaccrpt110.Acc14a0 ReportTitle(110) & IIf(Me.Text5.Text = "1", " (已沖)", IIf(Me.Text5.Text = "2", " (未沖)", "")), MaskEdBox1.Text, MaskEdBox2.Text, Text1.Text, Text2.Text, Text3.Text, Text4.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'      dllaccrpt110.Acc14a0 ReportTitle(110) & IIf(Me.Text5.Text = "1", " (已沖)", IIf(Me.Text5.Text = "2", " (未沖)", "")), MaskEdBox1.Text, MaskEdBox2.Text, Text1.Text, Text2.Text, Text3.Text, Text3.Text, StaffQuery(strUserNum), CFDate(ACDate(ServerDate))
'      'end 2007/10/2
'   End If
'   adoaccrpt110.Close
'   Screen.MousePointer = vbDefault
'   FormClear
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
End Sub

Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
   KeyEnter KeyCode
   If KeyCode <> vbKeyEscape Then
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   End If
End Sub

Private Sub Form_Load()
Dim intX As Integer
Dim intY As Integer
Dim sglWidth As Single
Dim sglHeight As Single
   
   'Modified by Lydia 2024/11/28 表單初始化
'   Me.Icon = LoadPicture(strIcoPath)
'   strFormName = Name
'   Me.Width = 5250
'
'   Me.Height = 2955
'   Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
'   Image1 = LoadPicture(strBackPicPath4)
'   sglWidth = Image1.Width
'   sglHeight = Image1.Height
'   For intX = 0 To Int(ScaleWidth / sglWidth)
'       For intY = 0 To Int(ScaleHeight / sglHeight)
'           PaintPicture Image1, intX * sglWidth, intY * sglHeight, sglWidth, sglHeight + 10, 0, 0
'       Next
'   Next
   PUB_InitForm Me, 5250, 3500, strBackPicPath4
   '預設公司別
   Call Pub_SetCboCmp(cboComp, False, False, False)
   'end 2024/11/28
   
   Text1 = "X"
   Text2 = "X"
   MaskEdBox1.Mask = DFormat
   MaskEdBox2.Mask = DFormat
   lblSalesName = ""
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(102)
   Set dllaccrpt110 = CreateObject("AccReport.ReportSelect")
End Sub

Private Sub Form_Unload(Cancel As Integer)
   strFormName = MsgText(601)
   KeyEnter vbKeyEscape
   MenuEnabled
   StatusClear
   Set dllaccrpt110 = Nothing
   Set Frmacc14a0 = Nothing
End Sub

Private Sub Text1_GotFocus()
   TextInverse Text1
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text1_Validate(Cancel As Boolean)
   If Len(Text1) = 6 Then
      Text1 = AfterZero(Text1)
   End If
End Sub

Private Sub Text2_GotFocus()
   TextInverse Text2
End Sub

Private Sub Text2_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

Private Sub Text2_Validate(Cancel As Boolean)
   If Len(Text2) = 6 Then
      Text2 = AfterZero(Text2)
   End If
End Sub

Private Sub Text3_Change()
   If Len(Text3) = 5 Then
      lblSalesName = StaffQuery(Text3)
   Else
      lblSalesName = MsgText(601)
   End If
End Sub

Private Sub Text3_GotFocus()
   TextInverse Text3
   CloseIme
End Sub

Private Sub Text3_KeyPress(KeyAscii As Integer)
   KeyAscii = UpperCase(KeyAscii)
End Sub

'*************************************************
'  產生報表資料
'
'*************************************************
'Mark by Lydia 2024/11/28 改新模組
'Private Sub ProduceData()
'Dim strSql As String
'
'On Error GoTo Checking
'   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
'      strSql = " and a0t03 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
'   End If
'   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
'      strSql = strSql & " and a0t03 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
'   End If
'   If Text1 <> MsgText(601) And Text1 <> "X" Then
'      strSql = strSql & " and a0t06 >= '" & Text1 & "'"
'   End If
'   If Text2 <> MsgText(601) And Text2 <> "X" Then
'      strSql = strSql & " and a0t06 <= '" & Text2 & "'"
'   End If
'   'Modify by Morgan 2007/10/2 智權人員範圍改成一個
'   'If Text3 <> MsgText(601) Then
'   '   strSQL = strSQL & " and a0t05 >= '" & Text3 & "'"
'   'End If
'   'If Text4 <> MsgText(601) Then
'   '   strSQL = strSQL & " and a0t05 <= '" & Text4 & "'"
'   'End If
'   If Text3 <> MsgText(601) Then
'      strSql = strSql & " and a0t05 = '" & Text3 & "'"
'   End If
'   'end 2007/10/2
'    'Add By Cheng 2004/01/13
'    '若非北所員工, 只能列印該所資料
'    If pub_strUserOffice <> "1" Then
'        strSql = strSql & " And ''||ST06='" & pub_strUserOffice & "' "
'    End If
'    'End
'    'Add By Cheng 2003/05/05
'    '沖銷狀況
'    'If Me.Text5.Text = "1" Then
'    '    strSQL = strSQL & " And a0t10 IS NOT NULL "
'    'ElseIf Me.Text5.Text = "2" Then
'    '    strSQL = strSQL & " And a0t10 IS NULL "
'    'End If
'    If Me.Text5.Text = "1" Then
'        '2007/10/30 modify by sonia 因J09400660轉國外收款沖到,故A1P02再加'F'
'        'modify by sonia 2021/8/23 很多暫收款財務處自行以總帳傳票沖銷,故人工上A0T10以區別,故加入A0T10為判斷條件
'        'strSql = strSql & " and a0t01 in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F') and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null)"
'        'modify by sonia 2024/5/14 a1p02加入'E'，因為a1p04='聯米企業股份有限公司971'有沖銷J09700301
'        strSql = strSql & " and (a0t10 is not null or a0t01 in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F', 'E') and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null))"
'    Else
'        '2007/10/29 MODIFY BY SONIA 應再剔除已做暫收款退費者, 10/30因J09400660轉國外收款沖到,故A1P02再加'F'
'        'strSQL = strSQL & " and a0t01 not in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W') and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null)"
'        'modify by sonia 2021/8/23 很多暫收款財務處自行以總帳傳票沖銷,故人工上A0T10以區別,故加入A0T10為判斷條件
'        'strSql = strSql & " and a0t01 not in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F') and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null) and a0t01 not in (select a0s02 from acc0s0 where substr(a0s02, 1, 1) = 'J')"
'        'modify by sonia 2024/5/14 a1p02加入'E'，因為a1p04='聯米企業股份有限公司971'有沖銷J09700301
'        strSql = strSql & " and a0t10 is null and a0t01 not in (select a1p23 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F', 'E') and a1p05 = '2401' and a1p07 <> 0 and a1p23 is not null) and a0t01 not in (select a0s02 from acc0s0 where substr(a0s02, 1, 1) = 'J')"
'    End If
'   If strSql <> MsgText(601) Then
'      strSql = " where " & Mid(strSql, 5, Len(strSql) - 4)
'   End If
'   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
'   adoaccrpt110.CursorLocation = adUseClient
'   adoaccrpt110.Open "select * from accrpt110 where r11001='" & strUserNum & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   adoacc0t0.CursorLocation = adUseClient
'    'Modify By Cheng 2004/01/13
''   adoacc0t0.Open "select * from acc0t0" & strSQL & " order by a0t03 asc, a0t01 asc", adoTaie, adOpenStatic, adLockReadOnly
'   adoacc0t0.Open "select * from acc0t0, Staff " & strSql & " And A0T05=ST01(+) order by a0t03 asc, a0t01 asc", adoTaie, adOpenStatic, adLockReadOnly
'    'End
'   If adoacc0t0.RecordCount = 0 Then
'      adoacc0t0.Close
'      adoaccrpt110.Close
'      MsgBox MsgText(28), , MsgText(5)
'      Exit Sub
'   End If
'   Do While adoacc0t0.EOF = False
'      adoaccrpt110.AddNew
'      adoaccrpt110.Fields("r11001").Value = strUserNum
'      If IsNull(adoacc0t0.Fields("a0t03").Value) Then
'         adoaccrpt110.Fields("r11002").Value = Null
'      Else
'         adoaccrpt110.Fields("r11002").Value = adoacc0t0.Fields("a0t03").Value
'      End If
'      adoaccrpt110.Fields("r11003").Value = adoacc0t0.Fields("a0t01").Value
'      If IsNull(adoacc0t0.Fields("a0t06").Value) Then
'         adoaccrpt110.Fields("r11004").Value = Null
'      Else
'        'Modify by Amy 2022/08/09 +StrToStr
'         adoaccrpt110.Fields("r11004").Value = StrToStr(adoacc0t0.Fields("a0t06").Value & " " & CustomerQuery(adoacc0t0.Fields("a0t06").Value, 1), 100)
'      End If
'      If IsNull(adoacc0t0.Fields("a0t05").Value) Then
'         adoaccrpt110.Fields("r11005").Value = Null
'      Else
'         adoaccrpt110.Fields("r11005").Value = StaffQuery(adoacc0t0.Fields("a0t05").Value)
'      End If
'      If IsNull(adoacc0t0.Fields("a0t08").Value) Then
'         adoaccrpt110.Fields("r11006").Value = 0
'      Else
'         adoaccrpt110.Fields("r11006").Value = adoacc0t0.Fields("a0t08").Value
'      End If
'      If IsNull(adoacc0t0.Fields("a0t02").Value) Then
'         adoaccrpt110.Fields("r11007").Value = Null
'      Else
'         Select Case adoacc0t0.Fields("a0t02").Value
'            Case "1"
'               adoaccrpt110.Fields("r11007").Value = ComboItem(201)
'            Case "2"
'               adoaccrpt110.Fields("r11007").Value = ComboItem(202)
'            Case "3"
'               adoaccrpt110.Fields("r11007").Value = ComboItem(203)
'         End Select
'      End If
'      If IsNull(adoacc0t0.Fields("a0t04").Value) Then
'         adoaccrpt110.Fields("r11008").Value = Null
'      Else
'         adoaccrpt110.Fields("r11008").Value = adoacc0t0.Fields("a0t04").Value
'      End If
'      adoaccrpt110.UpdateBatch
'      adoacc0t0.MoveNext
'   Loop
'   adoacc0t0.Close
'   adoaccrpt110.Close
'   adoTaie.Execute "delete from accrpt110 where r11002 is null and r11001='" & strUserNum & "'"
'   StatusClear
'Checking:
'   If Err.Number = 0 Then
'      Exit Sub
'   End If
'   MsgBox Err.Description, , MsgText(5)
'End Sub
'end 2024/11/28

'*************************************************
'  刪除報表資料
'
'*************************************************
Private Sub Accrpt110Delete()
   adoTaie.Execute "delete from accrpt110 where r11001='" & strUserNum & "'"
End Sub

'*************************************************
' 清除畫面
'
'*************************************************
Private Sub FormClear()
    MaskEdBox1.Mask = ""
    MaskEdBox1.Text = ""
    MaskEdBox1.Mask = DFormat
    MaskEdBox2.Mask = ""
    MaskEdBox2.Text = ""
    MaskEdBox2.Mask = DFormat
    Text1 = "X"
    Text2 = "X"
    Text3 = ""
    'Text4 = ""
    lblSalesName = ""
    'Add By Cheng 2003/05/05
    Me.Text5.Text = ""
    MaskEdBox1.SetFocus
    cboComp.Text = "" 'Added by Lydia 2024/11/28
End Sub

'*************************************************
'  畫面輸入檢查
'
'*************************************************
Public Function FormCheck() As Boolean
   If Text1 <> MsgText(601) And Text1 <> "X" Then
      FormCheck = True
      Exit Function
   End If
   If Text2 <> MsgText(601) And Text2 <> "X" Then
      FormCheck = True
      Exit Function
   End If
   
   If Text3 <> MsgText(601) Then
      FormCheck = True
      Exit Function
   End If
   'Remove by Morgan 2007/10/2 智權人員範圍改成一個
   'If Text4 <> MsgText(601) Then
   '   FormCheck = True
   '   Exit Function
   'End If
   'end 2007/10/2
   If MaskEdBox1.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   If MaskEdBox2.Text <> MsgText(29) Then
      FormCheck = True
      Exit Function
   End If
   FormCheck = False
End Function

Private Sub Text5_GotFocus()
    'Add By Cheng 2003/05/05
    TextInverse Me.Text5
End Sub

Private Sub Text5_KeyPress(KeyAscii As Integer)
    'Add By Cheng 2003/05/05
    KeyAscii = UpperCase(KeyAscii)
    If KeyAscii <> 8 And KeyAscii <> 49 And KeyAscii <> 50 Then
        KeyAscii = 0
    End If
End Sub

'Added by Lydia 2024/11/28
Private Sub CboComp_KeyPress(KeyAscii As Integer)
    KeyAscii = 0 ' 只可選(用單純下拉預設會錯)
End Sub

'Added by Lydia 2024/11/28 增加公司別
Private Sub ProduceData_New()
Dim strCon As String

On Error GoTo Checking
   If MaskEdBox1.Text <> MsgText(601) And MaskEdBox1.Text <> MsgText(29) Then
      strCon = " and a0t03 >= " & Val(FCDate(MaskEdBox1.Text)) & ""
   End If
   If MaskEdBox2.Text <> MsgText(601) And MaskEdBox2.Text <> MsgText(29) Then
      strCon = strCon & " and a0t03 <= " & Val(FCDate(MaskEdBox2.Text)) & ""
   End If
   If Text1 <> MsgText(601) And Text1 <> "X" Then
      strCon = strCon & " and a0t06 >= '" & Text1 & "'"
   End If
   If Text2 <> MsgText(601) And Text2 <> "X" Then
      strCon = strCon & " and a0t06 <= '" & Text2 & "'"
   End If

   If Text3 <> MsgText(601) Then
      strCon = strCon & " and a0t05 = '" & Text3 & "'"
   End If
   '若非北所員工, 只能列印該所資料
   If pub_strUserOffice <> "1" Then
      strCon = strCon & " And ''||ST06='" & pub_strUserOffice & "' "
   End If

   '公司別
   If Trim(cboComp.Text) <> "" Then
      strCon = strCon & " and a0t18='" & Left(Trim(cboComp.Text), 1) & "' "
   End If
   
   If strCon <> MsgText(601) Then
      strCon = " where " & Mid(strCon, 5, Len(strCon) - 4)
   End If
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(26)
   adoaccrpt110.CursorLocation = adUseClient
   adoaccrpt110.Open "select * from accrpt110 where r11001='" & strUserNum & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
   adoacc0t0.CursorLocation = adUseClient
   adoacc0t0.Open "select * from acc0t0, Staff " & strCon & " And A0T05=ST01(+) order by a0t03 asc, a0t01 asc", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc0t0.RecordCount = 0 Then
      adoacc0t0.Close
      adoaccrpt110.Close
      MsgBox MsgText(28), , MsgText(5)
      Exit Sub
   End If
   Do While adoacc0t0.EOF = False
      adoaccrpt110.AddNew
      adoaccrpt110.Fields("r11001").Value = strUserNum
      If IsNull(adoacc0t0.Fields("a0t03").Value) Then
         adoaccrpt110.Fields("r11002").Value = Null
      Else
         adoaccrpt110.Fields("r11002").Value = adoacc0t0.Fields("a0t03").Value
      End If
      adoaccrpt110.Fields("r11003").Value = adoacc0t0.Fields("a0t01").Value
      If IsNull(adoacc0t0.Fields("a0t06").Value) Then
         adoaccrpt110.Fields("r11004").Value = Null
      Else
         adoaccrpt110.Fields("r11004").Value = StrToStr(adoacc0t0.Fields("a0t06").Value & " " & CustomerQuery(adoacc0t0.Fields("a0t06").Value, 1), 100)
      End If
      If IsNull(adoacc0t0.Fields("a0t05").Value) Then
         adoaccrpt110.Fields("r11005").Value = Null
      Else
         adoaccrpt110.Fields("r11005").Value = StaffQuery(adoacc0t0.Fields("a0t05").Value)
      End If
      If IsNull(adoacc0t0.Fields("a0t08").Value) Then
         adoaccrpt110.Fields("r11006").Value = 0
      Else
         adoaccrpt110.Fields("r11006").Value = adoacc0t0.Fields("a0t08").Value
      End If
      If IsNull(adoacc0t0.Fields("a0t02").Value) Then
         adoaccrpt110.Fields("r11007").Value = Null
      Else
         Select Case adoacc0t0.Fields("a0t02").Value
            Case "1"
               adoaccrpt110.Fields("r11007").Value = ComboItem(201)
            Case "2"
               adoaccrpt110.Fields("r11007").Value = ComboItem(202)
            Case "3"
               adoaccrpt110.Fields("r11007").Value = ComboItem(203)
         End Select
      End If
      If IsNull(adoacc0t0.Fields("a0t04").Value) Then
         adoaccrpt110.Fields("r11008").Value = Null
      Else
         adoaccrpt110.Fields("r11008").Value = adoacc0t0.Fields("a0t04").Value
      End If
      '公司別
      If IsNull(adoacc0t0.Fields("a0t18").Value) Then
         adoaccrpt110.Fields("r11009").Value = Null
      Else
         adoaccrpt110.Fields("r11009").Value = adoacc0t0.Fields("a0t18").Value
      End If
      '沖帳日期r11010：若沖帳日期>輸入日期條件的止日時，則更新工作檔之沖帳日期為NULL；
                      '若沖銷狀況選「已沖」，則只抓工作檔中有沖帳日期的資料；若選「未沖」則只抓工作檔中沒有沖帳日期的資料。
      If "" & adoacc0t0.Fields("a0t10") <> "" Then
         strSql = "select a0205 as ndate from acc020 where a0202='" & adoacc0t0.Fields("a0t10") & "' and a0201='" & adoacc0t0.Fields("a0t18") & "' "
      Else
         'Modified by Lydia 2024/12/03 +,'0' as ord1
         strSql = "select a1p18 as ndate,'0' as ord1 from acc1p0 where a1p02 in ('A', 'Z', 'W', 'F', 'E') and a1p05 = '2401' and a1p07 <> 0 and a1p01='" & adoacc0t0.Fields("a0t18") & "' and a1p23='" & adoacc0t0.Fields("a0t01") & "' "
         'Added by Lydia 2024/12/03 判斷acc0s0
         strSql = strSql & " Union select a0s03 as ndate, '1' as ord1 from acc0s0 where a0s02='" & adoacc0t0.Fields("a0t01") & "' "
         strSql = strSql & " order by ord1 "
      End If
      intI = 1: strExc(1) = "": strExc(2) = ""
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Val("" & RsTemp.Fields("ndate")) > 0 Then
            If Replace(MaskEdBox2.Text, "/", "") <> "" And Replace(MaskEdBox2.Text, "/", "") < Val("" & RsTemp.Fields("ndate")) Then
               strExc(2) = Val("" & RsTemp.Fields("ndate"))
            Else
               strExc(1) = Val("" & RsTemp.Fields("ndate"))
            End If
         End If
      End If
      If strExc(1) <> "" Then
         adoaccrpt110.Fields("r11010").Value = strExc(1)
      Else
         adoaccrpt110.Fields("r11010").Value = Null
      End If
      
      'Modified by Lydia 2024/12/03 空白=未沖 => Or (Text5 = "" And strExc(1) <> "")
      If (Text5 = "1" And strExc(1) = "") Or (Text5 = "2" And strExc(1) <> "") Or (Trim(Text5) = "" And strExc(1) <> "") Then
         adoaccrpt110.Fields("r11002").Value = Null
      End If
      
      adoaccrpt110.UpdateBatch
      adoacc0t0.MoveNext
   Loop
   adoacc0t0.Close
   adoaccrpt110.Close
   adoTaie.Execute "delete from accrpt110 where r11002 is null and r11001='" & strUserNum & "'"
   StatusClear
Checking:
   If Err.Number = 0 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

