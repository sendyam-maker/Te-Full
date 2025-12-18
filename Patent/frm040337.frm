VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frm040337 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  '單線固定
   Caption         =   "資策會收到證書清單"
   ClientHeight    =   2100
   ClientLeft      =   36
   ClientTop       =   312
   ClientWidth     =   5676
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   5676
   Begin VB.CommandButton CmdPrt1 
      BackColor       =   &H00C0FFC0&
      Caption         =   "產生Excel(&E)"
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
      Left            =   390
      Style           =   1  '圖片外觀
      TabIndex        =   2
      Top             =   1380
      Width           =   4692
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "程式執行中請勿開啟Excel"
      BeginProperty Font 
         Name            =   "新細明體"
         Size            =   12
         Charset         =   136
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   1440
      TabIndex        =   5
      Top             =   180
      Width           =   2895
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   13
      Left            =   2940
      TabIndex        =   4
      Top             =   720
      Width           =   255
      VariousPropertyBits=   8388627
      Caption         =   "~"
      Size            =   "450;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   225
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   1
      Left            =   3180
      TabIndex        =   1
      Top             =   660
      Width           =   1200
      VariousPropertyBits=   679495707
      MaxLength       =   7
      Size            =   "2117;635"
      FontName        =   "新細明體"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.TextBox txtFM2 
      Height          =   360
      Index           =   0
      Left            =   1680
      TabIndex        =   0
      Top             =   660
      Width           =   1200
      VariousPropertyBits=   679495707
      MaxLength       =   7
      Size            =   "2117;635"
      FontName        =   "新細明體"
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
   End
   Begin MSForms.Label LblFM2 
      Height          =   225
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   735
      Width           =   1635
      VariousPropertyBits=   8388627
      Caption         =   "收到證書日期："
      Size            =   "2884;397"
      FontName        =   "新細明體-ExtB"
      FontEffects     =   1073741825
      FontHeight      =   195
      FontCharSet     =   136
      FontPitchAndFamily=   34
      FontWeight      =   700
   End
End
Attribute VB_Name = "frm040337"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Create by Lydia 2022/04/11
Option Explicit
Dim rsAD As New ADODB.Recordset
Dim intJ As Integer

Private Sub CmdPrt1_Click()
Dim stCon As String
Dim intP As Integer

   '畫面輸入檢查
   txtFM2(0) = Trim(txtFM2(0))
   txtFM2(1) = Trim(txtFM2(1))
   
   If txtFM2(0) <> "" And txtFM2(1) <> "" And txtFM2(0) > txtFM2(1) Then
       MsgBox "收到證書日期起值不可大於迄值！", vbCritical
       txtFM2(0).SetFocus
       txtFM2_GotFocus 0
       Exit Sub
   End If
            
   Screen.MousePointer = vbHourglass
   CmdPrt1.Enabled = False
   
   ClearQueryLog (Me.Name) '清除查詢印表記錄檔欄位
   
   '收到證書日期=來函日期
   If txtFM2(0) <> "" Then
       stCon = stCon & " AND CP05>=" & TransDate(txtFM2(0), 2)
   End If
   If txtFM2(1) <> "" Then
       stCon = stCon & " AND CP05<=" & TransDate(txtFM2(1), 2)
   End If
   pub_QL05 = pub_QL05 & ";" & LblFM2(0) & txtFM2(0) & "-" & txtFM2(1)
   
   'Modified by Lydia 2025/04/17 因為CFP-033339跑不出清單(指定國只剩UP，其他都閉卷)，發現條件不符合現況，所以修改
   '若為EPC案改抓指定國之年費期限---'Memo by Lydia 2025/04/16 找不到有相關文件和Email有提到此條件
   'strSql = "select cp01,cp02,cp03,cp04,cp05,pa48,na01,na60,pa14,pa15,pa22,pa24,pa25,pa72,pa73,pa16,na21 " & _
               "From caseprogress, patent, Nation " & _
               "where ((cp01 in ('P','CFP','FCP') and cp10='1603' and pa09<>'221') or (cp01='CFP' and cp10='224')) and pa09=na01(+) " & stCon & _
               " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
               "and instr(pa26||','||pa27||','||pa28||','||pa29||','||pa30,'X38805030') > 0 "
   'Modified by Morgan 2025/9/25
   strSql = "select cp01,cp02,cp03,cp04,cp05,pa48,na01,na60,pa14,pa15,pa22,pa24,pa25,pa72,pa73,pa16,na21,lastyear(pa72) lstyr " & _
               "From caseprogress, patent, Nation " & _
               "where cp01 in ('P','CFP','FCP') and cp10='1603' and pa09=na01(+) " & stCon & _
               " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
               "and instr(pa26||','||pa27||','||pa28||','||pa29||','||pa30,'X38805030') > 0 "
   strSql = strSql & " order by cp05,cp01,cp02"
    
   If rsAD.State = adStateOpen Then
      rsAD.Close
   End If
   rsAD.CursorLocation = adUseClient

   rsAD.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsAD.RecordCount <> 0 Then
        InsertQueryLog (rsAD.RecordCount)
        Call ProcExcelSave
   Else
        InsertQueryLog (0)
        ShowNoData
   End If
   rsAD.Close
   
   '執行完不清除條件
   CmdPrt1.Enabled = True
   Screen.MousePointer = vbDefault
   
End Sub

Private Sub Form_Load()

   MoveFormToCenter Me

   Call Pub_ChkExcelPath(strExcelPath)
   '「收到證書日期」的起迄條件，預設前月16日至本月15日
   txtFM2(0) = (Left(CompDate(1, -1, strSrvDate(1)), 6) - 191100) & "16"
   txtFM2(1) = Left(strSrvDate(1), 6) - 191100 & "15"
   
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm040337 = Nothing
End Sub

Private Sub txtFM2_GotFocus(Index As Integer)
    TextInverse txtFM2(Index)
End Sub

Private Sub txtFM2_KeyPress(Index As Integer, KeyAscii As MSForms.ReturnInteger)
     KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txtFM2_Validate(Index As Integer, Cancel As Boolean)
    
    Select Case Index
        Case 0, 1  '來函日期
           If PUB_CheckKeyInDate(txtFM2(Index)) = -1 Then
              GoTo EXITSUB
           End If
           
    End Select
    
    Exit Sub
    
EXITSUB:
    txtFM2(Index).SetFocus
    txtFM2_GotFocus Index
    Cancel = True
End Sub

'產生Excel檔案
Private Sub ProcExcelSave()
Dim xlsWDay As New Excel.Application
Dim wksWDay As New Worksheet
Dim strFileName As String '檔案名稱
Dim iRow As Integer, intCount As Integer
Dim strColName As String '欄位名稱
Dim strColW As String    '欄寬
Dim tmpArr1 As Variant, tmpArr2 As Variant
Dim xCols As Integer '行位置
Dim endX As String '最後一欄

On Error GoTo ErrHnd
   
   'Modified by Morgan 2025/9/25 資策會改欄位[下一年度年費繳納年度]->[已繳納年費之年度]--鈺華
   'strColName = "序號,收到證書日,資策會編號,事務所編號,國別,公告日,公告號,證書號,專利權始日,專利權止日,下一年度年費繳納年度,下一年度年費繳納期限（若為EPC案，此欄位為指定國之年費期限）,備註"
   strColName = "序號,收到證書日,資策會編號,事務所編號,國別,公告日,公告號,證書號,專利權始日,專利權止日,已繳納年費之年度,下一年度年費繳納期限（若為EPC案，此欄位為指定國之年費期限）,備註"
   strColW = "5,12,13.5,12,7,12.25,16,20,13,13,23.5,22,19"
   
   tmpArr1 = Split(strColName, ",")
   tmpArr2 = Split(strColW, ",")

    rsAD.MoveFirst
    strExc(1) = ""
   If txtFM2(0) <> "" Then strExc(1) = TransDate(txtFM2(0), 2) & strExc(1) & "~"
   If txtFM2(1) <> "" Then strExc(1) = strExc(1) & IIf(strExc(1) = "", "~", "") & TransDate(txtFM2(1), 2)
   strFileName = strExcelPath & Me.Caption & strExc(1) & MsgText(43)
   
   If Dir(strFileName) <> "" Then
      Kill strFileName
   End If
   xlsWDay.SheetsInNewWorkbook = 3
   xlsWDay.Workbooks.add
   xlsWDay.Visible = False '預設不顯示
   Set wksWDay = xlsWDay.Worksheets(1)
   xlsWDay.Sheets(1).Select '選擇工作表
   '設定抬頭
    iRow = 1
    xCols = Asc("A")
    For intI = 0 To UBound(tmpArr1)
        If Trim(tmpArr1(intI)) <> "" Then
           wksWDay.Range(Chr(xCols + intI) & iRow).Value = Trim(tmpArr1(intI))
           wksWDay.Range(Chr(xCols + intI) & iRow).HorizontalAlignment = xlCenter
           wksWDay.Range(Chr(xCols + intI) & ":" & Chr(xCols + intI)).ColumnWidth = Val(tmpArr2(intI))
           wksWDay.Range(Chr(xCols + intI) & ":" & Chr(xCols + intI)).HorizontalAlignment = xlCenter
           wksWDay.Range(Chr(xCols + intI) & iRow).Interior.ColorIndex = 19 '底色
           endX = Chr(xCols + intI)
        End If
    Next intI
    wksWDay.Range("1:1").RowHeight = 53
    wksWDay.Range("L1").WrapText = True '自動換列
    wksWDay.Range("2:2").Select
    xlsWDay.ActiveWindow.FreezePanes = True '凍結窗格
    wksWDay.Range("A1").Select
    wksWDay.Range("A:" & endX).Font.Size = 12
    
    Do While Not rsAD.EOF
        iRow = iRow + 1
        intCount = intCount + 1
        xCols = Asc("A")
        '序號
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = intCount
        End With
        xCols = xCols + 1
        '收到證書日
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = ChangeTStringToTDateString("" & rsAD.Fields("cp05"))
        End With
        xCols = xCols + 1
        '資策會編號
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = "" & rsAD.Fields("PA48")
        End With
        xCols = xCols + 1
        '事務所編號
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = rsAD.Fields("cp01") & "-" & rsAD.Fields("cp02") & IIf("" & rsAD.Fields("cp03") & rsAD.Fields("cp04") <> "000", "-" & rsAD.Fields("cp03") & "-" & rsAD.Fields("cp04"), "")
        End With
        xCols = xCols + 1
        '國別
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = "" & rsAD.Fields("NA60")
        End With
        xCols = xCols + 1
        '公告日
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = ChangeTStringToTDateString("" & rsAD.Fields("pa14"))
        End With
        xCols = xCols + 1
        '公告號
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = "" & rsAD.Fields("PA15")
        End With
        xCols = xCols + 1
        '證書號
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = "" & rsAD.Fields("PA22")
        End With
        xCols = xCols + 1
        '專利權始日
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = ChangeTStringToTDateString("" & rsAD.Fields("pa24"))
        End With
        xCols = xCols + 1
        '專利權止日
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = ChangeTStringToTDateString("" & rsAD.Fields("pa25"))
        End With
        xCols = xCols + 1
        '下一年度年費繳納年度: 模組取得, 但是期限是當期
        strExc(1) = "": strExc(2) = ""
        'Modified by Morgan 2025/9/25 資策會改欄位[下一年度年費繳納年度]->[已繳納年費之年度]--鈺華
        'strExc(1) = PUB_getNextPayYear(rsAD.Fields("CP01"), rsAD.Fields("CP02"), rsAD.Fields("CP03"), rsAD.Fields("CP04"), strExc(2))
        strExc(1) = "" & rsAD.Fields("lstyr")
        'end 2025/9/25
        '下一年度年費繳納年度=> 期限是當期,所以不使用
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = strExc(1)
        End With
        xCols = xCols + 1
       '下一年度年費繳納期限=>直接抓年費期限的法限
       'Modified by LYdia 2025/04/17 +607延展費
        strSql = "select min(np09) mdate from nextprogress where np06 is null and np07 in ('605','606','607') and np02='" & rsAD.Fields("CP01") & "' and np03='" & rsAD.Fields("CP02") & "' and np04='" & rsAD.Fields("CP03") & "' and np05='" & rsAD.Fields("CP04") & "' "
        strExc(2) = ""
        intI = 1
        Set RsTemp = ClsLawReadRstMsg(intI, strSql)
        If intI = 1 Then
            strExc(2) = "" & RsTemp.Fields("mdate")
        End If
        With wksWDay.Range(Chr(xCols) & iRow)
            .Value = ChangeTStringToTDateString(strExc(2))
        End With
        xCols = xCols + 1
        
       rsAD.MoveNext
    Loop

    xlsWDay.Sheets(1).Select '選擇工作表
   
   '判斷版本
   If Val(xlsWDay.Version) < 12 Then
        xlsWDay.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
        xlsWDay.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If

   xlsWDay.Workbooks.Close
   xlsWDay.Quit
   Set wksWDay = Nothing
   Set xlsWDay = Nothing
   MsgBox "Excel檔案產生完成！" & vbCrLf & "檔案位置：" & strExcelPathN
   Exit Sub

ErrHnd:

   MsgBox Err.Description
End Sub
