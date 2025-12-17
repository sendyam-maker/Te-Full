VERSION 5.00
Begin VB.Form Frmacc41l1 
   AutoRedraw      =   -1  'True
   Caption         =   "ACS待分潤明細表"
   ClientHeight    =   2148
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   6432
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   ScaleHeight     =   2148
   ScaleWidth      =   6432
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "11205年後ACS待分潤往來"
      ForeColor       =   &H00FF0000&
      Height          =   936
      Left            =   120
      TabIndex        =   6
      Top             =   1080
      Visible         =   0   'False
      Width           =   6200
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   0
         Left            =   1716
         MaxLength       =   6
         TabIndex        =   11
         Top             =   360
         Width           =   1212
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   1
         Left            =   2952
         MaxLength       =   1
         TabIndex        =   10
         Top             =   360
         Width           =   372
      End
      Begin VB.TextBox txtCode 
         Height          =   300
         Index           =   2
         Left            =   3348
         MaxLength       =   2
         TabIndex        =   9
         Top             =   360
         Width           =   492
      End
      Begin VB.TextBox txtSystem 
         Height          =   300
         Left            =   1080
         Locked          =   -1  'True
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "ACS"
         Top             =   360
         Width           =   612
      End
      Begin VB.CommandButton CmdTransactionDtl 
         BackColor       =   &H00C0C0C0&
         Caption         =   "  往來明細  Excel"
         BeginProperty Font 
            Name            =   "標楷體"
            Size            =   10.8
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   600
         Left            =   4500
         Style           =   1  '圖片外觀
         TabIndex        =   7
         Top             =   180
         Width           =   1560
      End
      Begin VB.Label Label15 
         BackStyle       =   0  '透明
         Caption         =   "本所案號"
         BeginProperty Font 
            Name            =   "新細明體"
            Size            =   10.2
            Charset         =   136
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Left            =   100
         TabIndex        =   12
         Top             =   360
         Width           =   1000
      End
   End
   Begin VB.CommandButton CmdExcel 
      BackColor       =   &H00C0FFC0&
      Caption         =   "   目前餘額   Excel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   765
      Left            =   4656
      Style           =   1  '圖片外觀
      TabIndex        =   3
      Top             =   120
      Width           =   1560
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(2)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   255
      Index           =   2
      Left            =   3330
      TabIndex        =   5
      Top             =   120
      Width           =   1100
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(0)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   0
      Left            =   2400
      TabIndex        =   4
      Top             =   120
      Width           =   900
   End
   Begin VB.Label Lbl1 
      BackStyle       =   0  '透明
      Caption         =   "Lbl(1)"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   11.4
         Charset         =   136
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   1
      Left            =   2400
      TabIndex        =   2
      Top             =   510
      Width           =   900
   End
   Begin VB.Label Label2 
      BackStyle       =   0  '透明
      Caption         =   "目前實績保留資料年月："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   1
      Top             =   495
      Width           =   2400
   End
   Begin VB.Image Image1 
      Height          =   132
      Left            =   0
      Top             =   960
      Visible         =   0   'False
      Width           =   132
   End
   Begin VB.Label Label1 
      BackStyle       =   0  '透明
      Caption         =   "目前智權點數輸入年月："
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   9.6
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   60
      TabIndex        =   0
      Top             =   120
      Width           =   2400
   End
End
Attribute VB_Name = "Frmacc41l1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Amy 2023/02/02
Option Explicit

Dim ado41L1 As New ADODB.Recordset
Dim bolOpenXls As Boolean, i As Integer
Dim strFileN As String, strAllF As String, strWidth As String, intField As Integer, intRow As Integer, intTitleR As Integer
Dim strField, intWidth
Dim bolHasAx210 As Boolean, strPreAxb(4 To 8) As String  '是否已過帳/系統-1個月傳票起迄
Dim strYM As String, strMaxSP01 As String, strA0b01 As String, strA0b05 As String '目前系統年月-1個月/目前智權點數輸入年月/目前過帳日/目前業績輸入關閉年月
Dim strReportN As String 'Add by Amy 2023/07/28

Private Sub CmdExcel_Click()
    'Modify by Amy 2023/07/28
    Screen.MousePointer = vbHourglass
    Frame1.Enabled = False
    strReportN = "ACS點數保留餘額明細"
    If ExcelSave = True Then
        MsgBox strReportN & " Excel檔案產生完成！（檔案位置：" & strExcelPath & strFileN & "）"
    End If
    Frame1.Enabled = True
    Screen.MousePointer = vbDefault
    'end 2023/07/28
End Sub

'Mark by Amy 2024/04/20 隱藏-婉莘
'Add by Amy 2023/07/28 往來明細
Private Sub CmdTransactionDtl_Click()
'   If Trim(txtCode(0)) <> MsgText(610) Then
'      If ChkCaseNo = False Then
'         txtCode(0).SetFocus
'         Exit Sub
'      End If
'   End If
'
'   Screen.MousePointer = vbHourglass
'   CmdExcel.Enabled = False
'   strReportN = "ACS點數保留往來明細"
'   If ExcelSave2 = True Then
'        MsgBox strReportN & " Excel檔案產生完成！（檔案位置：" & strExcelPath & strFileN & "）"
'        Call ClearText
'    End If
'   CmdExcel.Enabled = True
'   Screen.MousePointer = vbDefault
End Sub

Private Sub Form_Load()
    Dim intX As Integer, intY As Integer
    Dim sglWidth As Single, sglHeight As Single
    
    strFormName = Name
    Me.Width = 6675
    'Modify by Amy 2024/03/20 往來隱藏-婉莘
    '因目前往來只抓420101,辜11401月鑽系統漏洞於(產生轉票進傳票頁面)把49開頭的會科新增進傳票,導致往來資料不正確
    'Me.Height = 2712 'Modify by Amy 2023/07/28 原:1500
    Me.Height = 1512
    'end 2024/03/20
    Me.Move (lngWidth - Me.Width) / 2, (lngHeight - Me.Height) / 2
    Call ClearLabel
    Call ClearText 'Add by Amy 2023/07/28
    
    strYM = Mid(DBDATE(DateAdd("m", -1, Format(strSrvDate(1), "####/##/##"))), 1, 6) '系統前一個月
    strA0b01 = GetA0b01(strA0b05)
    Call bolAcc0b1(1, Val(strYM) - 191100, strPreAxb())
    '系統-1月,實績傳票是否已過帳
    If strPreAxb(4) <> MsgText(601) Then
        bolHasAx210 = Pub_ChkAxbPost(strPreAxb(4), strPreAxb(5))
    End If
    strMaxSP01 = GetMaxSP01(True)
    
    Lbl1(0).Caption = Val(strMaxSP01) - 191100
    Lbl1(1).Caption = strA0b05
    If bolHasAx210 = True Then
        Lbl1(2).Caption = "(已過帳)"
    End If
End Sub

Private Function ExcelSave() As Boolean
    Dim Xls As New Excel.Application, Wks As New Worksheet
    Dim strQ As String, strFileN As String, strWkName As String, strFormat As String, strD(1) As String
    Dim intQ As Integer, intPage As Integer
    Dim strTp As String
    
On Error GoTo ErrHand
    ExcelSave = False
    'Add by Amy 2023/07/28 從Form_Load搬過來
    strAllF = "傳票日期,客戶,案號,智權人員,目前餘額"
    strWidth = "10,40,13,15,20"
    strField = Split(strAllF, ",")
    intWidth = Split(strWidth, ",")
    'end 2023/07/28
    strD(0) = GetACSData("9", Me.Name, "", "")
    strD(1) = GetACSData("9.1", Me.Name, "", ",Acc020") '2492 第一筆傳票日
    
    strQ = "Select FirstDate,cu01||cu02||' '||Nvl(cu04,Decode(cu05,null,cu06,cu05||' '||cu88||' '||cu89||' '||cu90)) as CU,ax214,cu13||' '||st02 as Sales,Balance " & _
                "From (" & strD(0) & "),(" & strD(1) & "),LawCase,Customer,Staff " & _
                "Where ax214=CaseNo(+) And Substr(LC11,1,8)=cu01(+) And Decode(Substr(LC11,9,1),null,'0',Substr(LC11,9,1))=cu02(+) And cu13=st01(+) " & _
                "And SubStr(ax214, 1, length(ax214) - 9)=lc01 And lc02=SubStr(ax214, length(ax214)- 8, 6) And lc03=SubStr(ax214, length(ax214)- 2,1) And lc04=SubStr(ax214, length(ax214)- 1,length(ax214)) " & _
                "Order by ax214 "
    intQ = 1
    Set ado41L1 = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        intField = 65:  intRow = 1: intTitleR = 1: intPage = 1
        strFileN = strReportN & ServerDate & ServerTime & MsgText(43) 'Modify by Amy 2023/07/28 +strReportN
        If Dir(strExcelPath & strFileN) = MsgText(601) Then
            If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
                MkDir strExcelPath
            End If
        Else
            Kill strExcelPath & strFileN
        End If
        Xls.SheetsInNewWorkbook = 3
        Xls.Workbooks.add
        '工作表名稱改為中文
        If strWkName = MsgText(601) Then strWkName = Left(Xls.Worksheets(1).Name, Len(Xls.Worksheets(1).Name) - 1)
        Set Wks = Xls.Worksheets(strWkName & intPage)
        Wks.Activate
        Xls.Visible = True
        bolOpenXls = True
        Call SetField(Xls, Wks)
        '畫線
        Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + GetValue("目前餘額")) & intRow).Select
        Xls.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        
        intTitleR = intRow
        intRow = intRow + 1
    
        ado41L1.MoveFirst
        Do While ado41L1.EOF = False
            For i = LBound(strField) To UBound(strField)
                strFormat = ""
                Select Case i
                    Case GetValue("傳票日期")
                        strTp = "" & ado41L1.Fields("FirstDate")
                    Case GetValue("客戶")
                        strTp = "" & ado41L1.Fields("CU")
                    Case GetValue("案號")
                        strTp = "" & ado41L1.Fields("ax214")
                    Case GetValue("智權人員")
                        strTp = "" & ado41L1.Fields("Sales")
                    Case GetValue("目前餘額")
                        strTp = "" & ado41L1.Fields("Balance")
                        strFormat = "#,##0"
                End Select
                Wks.Range(Chr(intField + i) & intRow).Value = strTp
                If strFormat <> MsgText(601) Then
                    Wks.Range(Chr(intField + i) & intRow).NumberFormatLocal = strFormat
                End If
            Next i
            intRow = intRow + 1
            ado41L1.MoveNext
        Loop
        '加總
        Wks.Range(Chr(intField + GetValue("目前餘額")) & intRow).Value = "=Sum(" & Chr(intField + GetValue("目前餘額")) & intTitleR + 1 & ":" & Chr(intField + GetValue("目前餘額")) & intRow - 1 & ")"
        '畫線
        Wks.Range(Chr(intField) & intRow - 1 & ":" & Chr(intField + GetValue("目前餘額")) & intRow - 1).Select
        Xls.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
        'Excel 設定
        Wks.PageSetup.PaperSize = 9 'A4
        Wks.PageSetup.Orientation = xlPortrait '直印
        Wks.PageSetup.Zoom = 100
        Wks.PageSetup.LeftMargin = Xls.InchesToPoints(0.4)  '邊界
        Wks.PageSetup.RightMargin = Xls.InchesToPoints(0.4)
        Wks.PageSetup.TopMargin = Xls.InchesToPoints(0.4)
        Wks.PageSetup.BottomMargin = Xls.InchesToPoints(0.4)
        Wks.PageSetup.CenterHorizontally = True '版面設定->邊界->水平置中
        
        '判斷若版本2007以上改變存格式
        If Val(Xls.Version) < 12 Then
            Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
        Else
            Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
        End If
        Xls.Workbooks.Close
        Xls.Quit
        ExcelSave = True
    End If
    Exit Function

ErrHand:
    MsgBox Err.Description, , MsgText(5)
    If bolOpenXls = True Then
        If Val(Xls.Version) < 12 Then
            Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
        Else
            Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
        End If
        Xls.Workbooks.Close
        Xls.Quit
        Set Xls = Nothing
    End If
End Function

'Modify by Amy 2023/07/28 +Xls As Excel.Application
Private Sub SetField(Xls As Excel.Application, ByRef Wks As Worksheet)
   'Add by Amy 2023/07/28 +報表名
   Wks.Range(Chr(intField) & intRow).Value = strReportN
   Wks.Range(Chr(intField) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Select
  
    With Xls.Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .WrapText = False
        .Orientation = 0
        .AddIndent = False
        .ShrinkToFit = False
        .MergeCells = True
        .Font.Size = 18
        .Font.Bold = True
    End With
    intRow = intRow + 2
   
    Wks.Range(Chr(intField) & intRow).Value = "列印人員："
    Wks.Range(Chr(intField) & intRow).HorizontalAlignment = xlRight
    Wks.Range(Chr(intField + 1) & intRow).Value = StaffQuery(strUserNum)
    Wks.Range(Chr(intField + UBound(strField) - 2) & intRow).Value = "列印日期："
    Wks.Range(Chr(intField + UBound(strField) - 2) & intRow).HorizontalAlignment = xlRight
    Wks.Range(Chr(intField + UBound(strField) - 1) & intRow).Value = CFDate(ACDate(ServerDate))
    
    intRow = intRow + 1
    For i = LBound(strField) To UBound(strField)
        Wks.Range(Chr(intField + i) & intRow).Value = strField(i)
        Wks.Range(Chr(intField + i) & intRow).Font.Bold = True
        Wks.Columns(Chr(intField + i)).ColumnWidth = intWidth(i)
    Next i
    Wks.Range(Chr(intField + LBound(strField)) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).HorizontalAlignment = xlCenter
    '畫線
    Wks.Range(Chr(intField + LBound(strField)) & intRow & ":" & Chr(intField + UBound(strField)) & intRow).Select
    Xls.Selection.Borders(xlEdgeBottom).LineStyle = xlContinuous
End Sub

Private Function GetValue(pFieldN As String) As Integer
    Dim jj As Integer
 
    For jj = LBound(strField) To UBound(strField)
       If UCase(strField(jj)) = UCase(pFieldN) Then
          GetValue = jj
          Exit For
       End If
    Next jj
End Function

Private Sub ClearLabel()
    Dim objLbl As LABEL
    
    For Each objLbl In Lbl1
        objLbl.Caption = ""
    Next
End Sub

Private Sub Form_Unload(Cancel As Integer)
    strFormName = MsgText(601)
    KeyEnter vbKeyEscape
    MenuEnabled
    Set Frmacc41l1 = Nothing
End Sub

'Add by Amy 2023/07/28 確認是否為需分潤的案號
Private Function ChkCaseNo() As Boolean
   Dim rsA As New ADODB.Recordset, intA As Integer, strA As String, strA1 As String
 
   ChkCaseNo = False
   If txtCode(1) = MsgText(601) Then txtCode(1) = "0"
   If txtCode(2) = MsgText(601) Then txtCode(2) = "00"
   
   strA1 = " And ax214='" & txtSystem & txtCode(0) & txtCode(1) & txtCode(2) & "' "
   strA = GetACSData("9", Me.Name, "", ",Acc020", "And  a0205<1120501" & strA1) & _
   " Union " & GetACSData("8", Me.Name, "", ",ACC020", "And a0205>=1120501 And ax207>0" & strA1)
  
   intA = 1
   Set rsA = ClsLawReadRstMsg(intA, strA)
   If intA = 0 Then
      MsgBox "此案號無資料"
      Exit Function
   End If
   ChkCaseNo = True
End Function

Private Sub ClearText()
    Dim objTxt
    
    For Each objTxt In txtCode
        objTxt.Text = ""
    Next
End Sub

'往來明細
'Mark by Amy 2024/03/20 隱藏-婉莘
'因目前往來只抓420101,辜11401月鑽系統漏洞於(產生轉票進傳票頁面)把49開頭的會科新增進傳票,導致往來資料不正確
Private Function ExcelSave2() As Boolean
'   Dim Xls As New Excel.Application, Wks As New Worksheet, intQ As Integer, strQ As String, strWhere As String
'   Dim strCaseNo As String, strCaseNo_Old As String, strSalesP As String, strW2001P As String, strAx213P As String, strax212 As String
'   Dim strFileN As String, strWkName As String, strFormat As String, strFormula As String, strMinus As String
'   Dim intEnd As Integer, intPage As Integer, intIStartR As Integer, intBackColor As Integer, bolBlueColor As Boolean   'WorkSheet 頁數,收入起始列,儲存格底色,文字設藍色
'   Dim strTp(2) As String
'
'On Error GoTo ErrHand
'   ExcelSave2 = False
'
'   strAllF = "本所案號,智權人員,日期,收款點數,智權點數,顧服點數,智權/顧服,未分潤餘額,摘要,公司別,傳票號"
'   strWidth = "14,13.5,8,12,10,10,10,12,20,9,11"
'   strField = Split(strAllF, ",")
'   intWidth = Split(strWidth, ",")
'   intPage = 1
'
'   If Trim(txtCode(0)) <> MsgText(601) Then
'      strWhere = " And ax214='" & txtSystem & txtCode(0) & txtCode(1) & txtCode(2) & "' "
'   End If
'
'   '*** 11205月前(不含)「未分潤」的案子餘額 ***
'   strTp(1) = "Select 'I' as State,ax214,1120501 as a0205,ax208,Sum(ax207-ax206) as Balance,0 As Salesp,0 As W2001p,0 As Ax212p,'' as Ax201,'' as Ax202"
'   strTp(0) = GetACSData("9", Me.Name, "", ",Acc020", "And  a0205<1120501" & strWhere)
'   strTp(0) = Replace(UCase(strTp(0)), "SELECT AX214,SUM(AX207-AX206) AS BALANCE", strTp(1))
'   strTp(0) = Replace(UCase(strTp(0)), "GROUP BY AX214", "Group by ax214,ax208")
'
'   '*** 11205月以後「收款轉保留」的案子 ***
'   strTp(1) = "Select 'I' as State,ax214,a0205,ax208,Sum(ax207-ax206) as Balance,0 As Salesp,0 As W2001p,0 As Ax212p,'' as Ax201,'' as Ax202"
'   strTp(2) = GetACSData("8", Me.Name, "", ",ACC020", "And a0205>=1120501 And ax207>0" & strWhere)
'   strTp(2) = Replace(UCase(strTp(2)), "SELECT AX214,SUM(AX207-AX206) AS BALANCE", strTp(1))
'   strTp(2) = Replace(UCase(strTp(2)), "GROUP BY AX214", "Group by ax214,a0205,ax208")
'   strQ = strTp(0) & " Union " & strTp(2)
'
'   '*** 11205月以後「分潤」的案子 ***
'   strQ = strQ & " Union " & _
'                "Select 'S' as State,ax214,A0205,'' as Ax208,0 as Balance,Sum(Decode(Sort,1,Ax207,'')) As Salesp,Sum(Decode(Sort,2,Ax207,'')) As W2001P,Sum(Decode(Sort,3,Ax207,'')) As Ax212P,ax201,ax202 " & _
'                "From (" & GetACSData("5", Me.Name, "", ",Acc020,Staff", "And A0205>=1120501" & strWhere) & ") Group by ax214,a0205,ax201,ax202 "
'
'   strQ = strQ & " Order by ax214,state,a0205,ax201,ax202 "
'
'   intQ = 1
'   Set ado41L1 = ClsLawReadRstMsg(intQ, strQ)
'   If intQ = 1 Then
'      strFileN = strReportN & ServerDate & ServerTime & MsgText(43)
'      If Dir(strExcelPath & strFileN) = MsgText(601) Then
'         If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = MsgText(601) Then
'            MkDir strExcelPath
'         End If
'      Else
'         Kill strExcelPath & strFileN
'      End If
'      Xls.SheetsInNewWorkbook = 3
'      Xls.Workbooks.add
'      '工作表名稱改為中文
'      If strWkName = MsgText(601) Then strWkName = Left(Xls.Worksheets(1).Name, Len(Xls.Worksheets(1).Name) - 1)
'      Set Wks = Xls.Worksheets(strWkName & intPage)
'      Wks.Activate
'      Xls.Visible = True: bolOpenXls = True
'     intField = 65:  intRow = 1: intBackColor = 19
'      Call SetField(Xls, Wks)
'      intTitleR = intRow: intRow = intRow + 1: intIStartR = intRow
'
'      ado41L1.MoveFirst
'      Do While ado41L1.EOF = False
'         strMinus = ""
'         '不同案號,更新開始列,換色
'         If strCaseNo_Old <> ado41L1.Fields("ax214") Then
'            If strCaseNo_Old <> MsgText(601) And intBackColor = 19 Then
'               strTp(1) = Chr(intField + LBound(strField)) & intIStartR & ":" & Chr(intField + UBound(strField)) & intRow - 1
'               Wks.Range(strTp(1)).Interior.ColorIndex = 36  '設置儲存格填充色(黃)
'               intBackColor = 0
'            ElseIf intBackColor = 0 Then
'               intBackColor = 19
'            End If
'            intIStartR = intRow
'            '未分潤餘額 欄公式
'            strFormula = Chr(intField + GetValue("收款點數")) & intRow
'         Else
'            'State=S 分潤為減項
'            If ado41L1.Fields("State") = "S" Then strMinus = "*-1"
'            '前一列收款點數欄+目前列Sum(收款點數~智權/顧服)
'            strFormula = Chr(intField + GetValue("收款點數")) & intRow - 1 & _
'                "+Sum(" & Chr(intField + GetValue("收款點數")) & intRow & ":" & Chr(intField + GetValue("智權/顧服")) & intRow & ")" & strMinus
'         End If
'
'         intEnd = UBound(strField): strMinus = ""
'         'State=I=11205月前(不含)的未分潤的餘額 or 11205月以後收款轉保留的案子
'         If ado41L1.Fields("State") = "I" Then intEnd = GetValue("未分潤餘額")
'
'         For i = LBound(strField) To intEnd
'            strTp(0) = "": strTp(1) = "": strFormat = ""
'            Select Case i
'               Case GetValue("本所案號")
'                  strCaseNo = ado41L1.Fields("ax214")
'                  '案號不同才顯示
'                  If strCaseNo_Old <> strCaseNo Then
'                     strTp(0) = strCaseNo
'                  End If
'               Case GetValue("智權人員")
'                  '未分潤 or 收款轉保留的案子
'                  If ado41L1.Fields("State") = "I" Then
'                     strSalesP = "Y"
'                     '11205月前(不含)的未分潤的餘額,作帳不一致無法抓收款智權,故抓客戶檔目前智權
'                     If ado41L1.Fields("a0205") = "1120501" Then
'                        If Trim("" & ado41L1.Fields("ax208")) <> MsgText(601) Then Call GetCuSales(ado41L1.Fields("ax208"), strSalesP)
'                     '11205月以後收款轉保留的案子,抓4191 當月月底保留借方的對沖其他(顯示收款智權人員)
'                     Else
'                        Call GetACSData("3", Me.Name, Mid(ado41L1.Fields("a0205"), 1, 5), ",Acc020", "And ax206>0 And ax214='" & ado41L1.Fields("ax214") & "' ", strSalesP)
'                     End If
'                     strTp(0) = strSalesP & " " & GetPrjSalesNM(strSalesP)
'                  '分潤
'                  Else
'                     '與收款智權人員不同 , 顯示目前分潤傳票 智權人員(智權不是W2001且ax213是空)
'                     strTp(1) = Pub_GetField("Acc021", "ax201='" & ado41L1.Fields("ax201") & "' And ax202='" & ado41L1.Fields("ax202") & "' And ax209<>'W2001' And ax213 is Null ", "ax209")
'                     If strSalesP <> strTp(1) Then
'                        bolBlueColor = True
'                        strTp(0) = strTp(1) & " " & GetPrjSalesNM(strTp(1))
'                     End If
'                  End If
'               Case GetValue("日期")
'                  strTp(0) = "" & ado41L1.Fields("a0205")
'               Case GetValue("收款點數")
'                  If ado41L1.Fields("State") = "I" Then
'                     strTp(0) = Val(ado41L1.Fields("Balance"))
'                     strFormat = "#,##0"
'                  End If
'               Case GetValue("智權點數")
'                  If ado41L1.Fields("State") <> "I" Then
'                     strTp(0) = Val("" & ado41L1.Fields("SalesP"))
'                     strFormat = "#,##0"
'                  End If
'               Case GetValue("顧服點數")
'                  If ado41L1.Fields("State") <> "I" Then
'                     strTp(0) = Val("" & ado41L1.Fields("W2001P"))
'                     strFormat = "#,##0"
'                  End If
'               Case GetValue("智權/顧服")
'                  If ado41L1.Fields("State") <> "I" Then
'                     strTp(0) = Val("" & ado41L1.Fields("Ax212P"))
'                     strFormat = "#,##0"
'                  End If
'               Case GetValue("未分潤餘額")
'                  strTp(0) = "=" & strFormula
'                  strFormat = "#,##0"
'               Case GetValue("摘要")
'                  If ado41L1.Fields("State") <> "I" Then
'                     strTp(0) = Pub_GetField("Acc021", "ax201='" & ado41L1.Fields("ax201") & "' And ax202='" & ado41L1.Fields("ax202") & "' And ax209<>'W2001' And ax213 is Null ", "ax212")
'                  End If
'               Case GetValue("公司別")
'                  strTp(0) = "" & ado41L1.Fields("ax201")
'               Case GetValue("傳票號")
'                  strTp(0) = "" & ado41L1.Fields("ax202")
'            End Select
'            If strTp(0) <> MsgText(601) Then
'               Wks.Range(Chr(intField + i) & intRow).Value = strTp(0)
'            End If
'            If strFormat <> MsgText(601) Then
'               Wks.Range(Chr(intField + i) & intRow).NumberFormatLocal = "#,##0"
'            End If
'            If bolBlueColor = True Then
'               Wks.Range(Chr(intField + i) & intRow).Font.Color = vbBlue
'            End If
'         Next i
'         intRow = intRow + 1
'
'         strCaseNo_Old = strCaseNo
'         ado41L1.MoveNext
'      Loop
'   End If
'   '總計
'   Wks.Range(Chr(intField + GetValue("日期")) & intRow).Value = "合 計"
'   Wks.Range(Chr(intField + GetValue("日期")) & intRow).HorizontalAlignment = xlCenter
'   Wks.Range(Chr(intField + GetValue("日期")) & intRow).Font.Bold = True
'   For i = GetValue("收款點數") To GetValue("未分潤餘額")
'      Wks.Range(Chr(intField + i) & intRow).Value = "=Sum(" & Chr(intField + i) & intTitleR + 1 & ":" & Chr(intField + i) & intRow - 1 & ")"
'      Wks.Range(Chr(intField + i) & intRow).NumberFormatLocal = "#,##0"
'   Next i
'   '畫線
'   Wks.Range(Chr(intField + GetValue("收款點數")) & intRow & ":" & Chr(intField + GetValue("未分潤餘額")) & intRow).Select
'   Xls.Selection.Borders(xlEdgeTop).LineStyle = xlDouble
'   '設定表格內字型/字大小
'   Wks.Range(Chr(intField + LBound(strField)) & intTitleR & ":" & Chr(intField + UBound(strField)) & intRow).Font.Name = "標楷體"
'   Wks.Range(Chr(intField + LBound(strField)) & intTitleR & ":" & Chr(intField + UBound(strField)) & intRow).Font.Size = 11
'   'Excel 設定
'   Wks.PageSetup.PaperSize = 9 'A4
'   Wks.PageSetup.Orientation = xlLandscape '橫印
'   Wks.PageSetup.Zoom = 100
'   Wks.PageSetup.LeftMargin = Xls.InchesToPoints(0.4)  '邊界
'   Wks.PageSetup.RightMargin = Xls.InchesToPoints(0.4)
'   Wks.PageSetup.TopMargin = Xls.InchesToPoints(0.4)
'   Wks.PageSetup.BottomMargin = Xls.InchesToPoints(0.4)
'
'   If Val(Xls.Version) < 12 Then
'      Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
'   Else
'      Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
'   End If
'   Xls.Workbooks.Close
'   Xls.Quit
'   Set Xls = Nothing
'   ExcelSave2 = True
'   Exit Function
'
'ErrHand:
'    MsgBox Err.Description, , MsgText(5)
'    If bolOpenXls = True Then
'        If Val(Xls.Version) < 12 Then
'            Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=-4143
'        Else
'            Xls.Workbooks(1).SaveAs FileName:=strExcelPath & strFileN, FileFormat:=56
'        End If
'        Xls.Workbooks.Close
'        Xls.Quit
'        Set Xls = Nothing
'    End If
End Function

Private Sub txtCode_Validate(Index As Integer, Cancel As Boolean)
   If Index = 0 Then Exit Sub
    
    '第3及4碼案號未輸,補0
    If txtCode(Index - 1) <> MsgText(601) And txtCode(Index) = MsgText(601) Then
        If Index = 1 Then txtCode(Index) = "0"
        If Index = 2 Then txtCode(Index) = "00"
    End If
End Sub
