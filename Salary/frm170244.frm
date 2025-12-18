VERSION 5.00
Begin VB.Form frm170244 
   BorderStyle     =   1  '單線固定
   Caption         =   "智權人員薪點表"
   ClientHeight    =   2720
   ClientLeft      =   50
   ClientTop       =   340
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2720
   ScaleWidth      =   4860
   Begin VB.CommandButton cmdok 
      Caption         =   "結束(&X)"
      Height          =   375
      Index           =   1
      Left            =   3750
      TabIndex        =   3
      Top             =   90
      Width           =   975
   End
   Begin VB.CommandButton cmdok 
      Caption         =   "Excel(&E)"
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   2670
      TabIndex        =   2
      Top             =   90
      Width           =   975
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   0
      Left            =   1860
      MaxLength       =   5
      TabIndex        =   0
      Top             =   1080
      Width           =   780
   End
   Begin VB.TextBox txt1 
      Height          =   270
      Index           =   1
      Left            =   2850
      MaxLength       =   5
      TabIndex        =   1
      Top             =   1080
      Width           =   780
   End
   Begin VB.Label LblFileNote 
      AutoSize        =   -1  'True
      Caption         =   "備註: 電子檔存放位置= "
      ForeColor       =   &H00C00000&
      Height          =   180
      Left            =   210
      TabIndex        =   5
      Top             =   2430
      Width           =   1870
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "年　　月："
      Height          =   180
      Index           =   2
      Left            =   930
      TabIndex        =   4
      Top             =   1140
      Width           =   900
   End
   Begin VB.Line Line1 
      Index           =   0
      X1              =   2580
      X2              =   2895
      Y1              =   1200
      Y2              =   1200
   End
End
Attribute VB_Name = "frm170244"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Created by Sindy 2024/3/19
Option Explicit

Dim xlsSalesPoint As New Excel.Application
Dim wksheet As New Worksheet
Dim strYM(1 To 12) As String
Dim m_intYMcnt As Integer '抓的月數量 ex:11301~11303 = 3
Dim strPECol As String, strSMCol As String


Private Sub cmdok_Click(Index As Integer)
   Select Case Index
      Case 0
         Screen.MousePointer = vbHourglass
         If TxtValidate = True Then
            Me.Enabled = False
            PrintSheet
            Me.Enabled = True
         End If
         Screen.MousePointer = vbDefault
      Case 1
         Unload Me
   End Select
End Sub

Private Sub Form_Load()
   MoveFormToCenter Me
   LblFileNote.Caption = LblFileNote.Caption & strExcelPath
End Sub

Private Sub Form_Unload(Cancel As Integer)
   Set frm170244 = Nothing
End Sub

Private Function TxtValidate() As Boolean
   Dim bCancel As Boolean
   Dim oText As TextBox
   
   If txt1(0) = "" Then
      MsgBox "年月起不可空白 !"
      txt1(0).SetFocus
      Exit Function
   End If
   If txt1(1) = "" Then
      MsgBox "年月迄不可空白 !"
      txt1(1).SetFocus
      Exit Function
   End If
   If Mid(txt1(0), 1, Len(txt1(0)) - 2) <> Mid(txt1(1), 1, Len(txt1(1)) - 2) Then
      MsgBox "年月起迄需同年，期間限制1年 !"
      txt1(1).SetFocus
      Exit Function
   End If
   If txt1(1) >= Mid(strSrvDate(2), 1, Len(strSrvDate(2)) - 2) Then
      MsgBox "年月迄不可等於大於當月，因未發薪 !"
      txt1(1).SetFocus
      Exit Function
   End If
   
   TxtValidate = True
End Function

Private Sub PrintSheet()
Dim YM As String
Dim long_ii As Long, ii As Integer
   
   '清值
   For ii = 1 To 12
      strYM(ii) = ""
   Next ii
   ii = 0: strSql = ""
   For long_ii = txt1(0) To txt1(1)
      ii = ii + 1
      '民國年
      If Right(long_ii, 2) > 12 Then
         strExc(0) = long_ii
         long_ii = CLng(CStr(Left(strExc(0), Len(strExc(0)) - 2) + 1) & "01")
      End If
      strExc(10) = TransDate(GetLastDay(long_ii & "01"), 1)
      YM = Left(strExc(10), Len(strExc(10)) - 2)
      strYM(ii) = YM '記錄要統計的年月
      If ii = 12 Then Exit For
   Next long_ii
   If ii = 0 Then
      Exit Sub
   Else
      strSql = "delete from R170244 where ID='" & strUserNum & "'"
      cnnConnection.Execute strSql, intI
'      strSql = "drop table R170244"
'      cnnConnection.Execute strSql, intI
'      strSql = "create table R170244("
'      strSql = strSql + "ID varchar2(6),"
'      strSql = strSql + "RST01 varchar2(6),"
'      strSql = strSql + "RST02 varchar2(12),"
'      strSql = strSql + "RPE01 number(8),"
'      strSql = strSql + "RPE02 number(8),"
'      strSql = strSql + "RPE03 number(8),"
'      strSql = strSql + "RPE04 number(8),"
'      strSql = strSql + "RPE05 number(8),"
'      strSql = strSql + "RPE06 number(8),"
'      strSql = strSql + "RPE07 number(8),"
'      strSql = strSql + "RPE08 number(8),"
'      strSql = strSql + "RPE09 number(8),"
'      strSql = strSql + "RPE10 number(8),"
'      strSql = strSql + "RPE11 number(8),"
'      strSql = strSql + "RPE12 number(8),"
'      strSql = strSql + "RSM01 number(8),"
'      strSql = strSql + "RSM02 number(8),"
'      strSql = strSql + "RSM03 number(8),"
'      strSql = strSql + "RSM04 number(8),"
'      strSql = strSql + "RSM05 number(8),"
'      strSql = strSql + "RSM06 number(8),"
'      strSql = strSql + "RSM07 number(8),"
'      strSql = strSql + "RSM08 number(8),"
'      strSql = strSql + "RSM09 number(8),"
'      strSql = strSql + "RSM10 number(8),"
'      strSql = strSql + "RSM11 number(8),"
'      strSql = strSql + "RSM12 number(8),"
'      strSql = strSql + "Primary Key(ID,RST01)"
'      strSql = strSql + ")"
'      cnnConnection.Execute strSql
   
      '寫入資料
      strPECol = "": strSMCol = ""
      For ii = 1 To 12
         If strYM(ii) <> "" Then
            m_intYMcnt = ii
            strExc(10) = Right(strYM(ii), 2)
            strPECol = strPECol & ",RPE" & strExc(10)
            strSMCol = strSMCol & ",RSM" & strExc(10)
            '有設定目標的智權人員
            strSql = "insert into R170244(ID,RST01,RST02,RPE" & strExc(10) & ",RSM" & strExc(10) & ")" & _
                     " select '" & strUserNum & "',st01,st02,PE04,nvl(sm04,0)+nvl(sm05,0)+nvl(sm07,0)" & _
                     " From Performance, staff, SalaryMonth" & _
                     " where pe03=" & strYM(ii) + 191100 & " and pe02='TOT'" & _
                     " and pe01=st01(+) and substr(st03,1,1)='S' and substr(st01,1,1)>='6' and substr(st01,1,1)<'F'" & _
                     " and sm01(+)=pe01 and sm02(+)=pe03 and sm01 is not null" & _
                     " and pe01 not in(select rST01 from R170244 where ID='" & strUserNum & "' and rST01=pe01)"
            cnnConnection.Execute strSql, intI
            '尚無設定目標的智權人員
            strSql = "insert into R170244(ID,RST01,RST02,RPE" & strExc(10) & ",RSM" & strExc(10) & ")" & _
                     " select '" & strUserNum & "',st01,st02,0,nvl(sm04,0)+nvl(sm05,0)+nvl(sm07,0)" & _
                     " From staff, SalaryMonth" & _
                     " where substr(st03,1,1)='S' and substr(st01,1,1)>='6' and substr(st01,1,1)<'F' and st04='1'" & _
                     " and sm01(+)=st01 and sm02=" & strYM(ii) + 191100 & " and sm01 is not null" & _
                     " and sm01 not in(select rST01 from R170244 where ID='" & strUserNum & "' and rST01=sm01)"
            cnnConnection.Execute strSql, intI
            If ii > 1 Then
               strSql = "select st01,st02,PE04,nvl(sm04,0)+nvl(sm05,0)+nvl(sm07,0)" & _
                        " From Performance, staff, SalaryMonth" & _
                        " where pe03=" & strYM(ii) + 191100 & " and pe02='TOT'" & _
                        " and pe01=st01(+) and substr(st03,1,1)='S' and substr(st01,1,1)>='6' and substr(st01,1,1)<'F'" & _
                        " and sm01(+)=pe01 and sm02(+)=pe03 and sm01 is not null" & _
                        " and pe01 in(select rST01 from R170244 where ID='" & strUserNum & "' and rST01=pe01)"
               strSql = strSql & " union " & _
                        "select st01,st02,0,nvl(sm04,0)+nvl(sm05,0)+nvl(sm07,0)" & _
                        " From staff, SalaryMonth, Performance" & _
                        " where substr(st03,1,1)='S' and substr(st01,1,1)>='6' and substr(st01,1,1)<'F' and st04='1'" & _
                        " and sm01(+)=st01 and sm02=" & strYM(ii) + 191100 & " and sm01 is not null" & _
                        " and sm01 in(select rST01 from R170244 where ID='" & strUserNum & "' and rST01=sm01)" & _
                        " and pe03(+)=" & strYM(ii) + 191100 & " and pe02(+)='TOT' and pe01(+)=st01 and pe01 is null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  RsTemp.MoveFirst
                  Do While Not RsTemp.EOF
                     strSql = "update R170244 set RPE" & strExc(10) & "=" & RsTemp.Fields(2) & ",RSM" & strExc(10) & "=" & RsTemp.Fields(3) & _
                              " where ID='" & strUserNum & "'" & _
                              " and RST01='" & RsTemp.Fields("st01") & "'"
                     cnnConnection.Execute strSql
                     RsTemp.MoveNext
                  Loop
               End If
            End If
         Else
            Exit For
         End If
      Next ii
      
      ExcelSave
   End If
End Sub

'*************************************************
'  產出Excel檔案
'
'*************************************************
Private Sub ExcelSave()
Dim strFileName As String

On Error GoTo flgErr
   
   strFileName = strExcelPath & Val(txt1(0)) & "~" & Val(txt1(1)) & "智權人員薪點表.xls"
   If Dir(strFileName) <> MsgText(601) Then
      Kill strFileName
   End If
   xlsSalesPoint.SheetsInNewWorkbook = 1 'Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
   xlsSalesPoint.Workbooks.add
   
   If ExcelSave_Detail = False Then
      Exit Sub
   End If
   
   If Val(xlsSalesPoint.Version) < 12 Then
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
   Else
      xlsSalesPoint.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
   End If
   xlsSalesPoint.Workbooks.Close
   xlsSalesPoint.Quit
   Set wksheet = Nothing
   Set xlsSalesPoint = Nothing
   MsgBox "檔案已產生！電子檔位置：" & strFileName
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Sub

Private Function ExcelSave_Detail() As Boolean
Dim lngCounter As Integer, intCol As Integer, intStarRow As Integer
Dim j As Integer
Dim strTitleField As String
Dim strFieldPE As String, strFieldSM As String, strFirst_SM As String
Dim strAvg_PE As String, strAvg_SM As String
Dim strTp_PE As String, strTp_SM As String
Dim varTmp As Variant
   
On Error GoTo flgErr
   
   ExcelSave_Detail = False
   xlsSalesPoint.Visible = True 'Excel顯示出來
   Set wksheet = xlsSalesPoint.Worksheets(1)
   wksheet.Columns("a:a").ColumnWidth = 9
   wksheet.Columns("b:b").ColumnWidth = 9
   wksheet.Columns("c:c").ColumnWidth = 9
   wksheet.Range("a1").Value = "資料查詢期間 : " & txt1(0) & " ~ " & txt1(1)
   wksheet.Range("a2").Value = "離職或者新進人員的薪點須再確認!"
   wksheet.Range("a2").Font.Bold = True '粗體
   wksheet.Range("a2").Font.ColorIndex = 3 '紅字
   '標題
   wksheet.Range("A4").Value = "員工編號"
   wksheet.Range("B4").Value = "智權人員"
   wksheet.Range("C4").Value = "薪點"
   intCol = 2 '因為C欄位起頭,所以intCol預設為2
   '目標
   For j = 1 To m_intYMcnt
      intCol = intCol + 1
      strTitleField = GetFieldStr(intCol, Asc("A")) '65.A~90.Z 67=C :使用此函數需為A起頭算
      strFieldPE = strFieldPE & "," & strTitleField
      If j = 1 Then
         wksheet.Range(strTitleField & "3").Value = "目標"
      End If
      wksheet.Range(strTitleField & "4").Value = strYM(j)
      wksheet.Columns(strTitleField & ":" & strTitleField).ColumnWidth = 8
   Next j
   intCol = intCol + 1
   strTitleField = GetFieldStr(intCol, Asc("A"))
   strAvg_PE = strTitleField
   strFieldPE = strFieldPE & "," & strTitleField
   wksheet.Range(strTitleField & "4").Value = "平均目標"
   wksheet.Columns(strTitleField & ":" & strTitleField).ColumnWidth = 10
   '薪資
   For j = 1 To m_intYMcnt
      intCol = intCol + 1
      strTitleField = GetFieldStr(intCol, Asc("A"))
      strFieldSM = strFieldSM & "," & strTitleField
      If j = 1 Then
         wksheet.Range(strTitleField & "3").Value = "薪資"
         strFirst_SM = strTitleField '薪資第一個欄位
      End If
      wksheet.Range(strTitleField & "4").Value = strYM(j)
      wksheet.Columns(strTitleField & ":" & strTitleField).ColumnWidth = 8
   Next j
   intCol = intCol + 1
   strTitleField = GetFieldStr(intCol, Asc("A"))
   strAvg_SM = strTitleField
   strFieldSM = strFieldSM & "," & strTitleField
   wksheet.Range(strTitleField & "4").Value = "平均薪資"
   wksheet.Columns(strTitleField & ":" & strTitleField).ColumnWidth = 10
   
   intStarRow = 4 '起始列
   strSql = "select RST01,RST02" & strPECol & strSMCol & _
            " From R170244,staff where ID='" & strUserNum & "' and RST01=ST01(+)" & _
            " order by ST15 asc,ST01 asc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      lngCounter = intStarRow '起始列
      Do While Not RsTemp.EOF
         lngCounter = lngCounter + 1
         wksheet.Range("A" & lngCounter).Value = RsTemp.Fields(0)
         wksheet.Range("B" & lngCounter).Value = RsTemp.Fields(1)
         strTp_PE = "": strTp_SM = ""
         '目標
         varTmp = Split(strFieldPE, ",")
         For j = 1 To m_intYMcnt
            wksheet.Range(varTmp(j) & lngCounter).Value = RsTemp.Fields("RPE" & Right(strYM(j), 2))
            If j = 1 Then
               strTp_PE = "=ROUND(Sum(" & varTmp(j) & lngCounter & ":"
            End If
            If j = m_intYMcnt Then
               strTp_PE = strTp_PE & varTmp(j) & lngCounter & ")/" & m_intYMcnt & ",2)"
            End If
         Next j
         '薪資
         varTmp = Split(strFieldSM, ",")
         For j = 1 To m_intYMcnt
            intCol = intCol + 1
            wksheet.Range(varTmp(j) & lngCounter).Value = RsTemp.Fields("RSM" & Right(strYM(j), 2))
            If j = 1 Then
               strTp_SM = "=Sum(" & varTmp(j) & lngCounter & ":"
            End If
            If j = m_intYMcnt Then
               strTp_SM = strTp_SM & varTmp(j) & lngCounter & ")/" & m_intYMcnt
            End If
         Next j
         wksheet.Range(strAvg_PE & lngCounter).Value = strTp_PE '平均目標
         wksheet.Range(strAvg_SM & lngCounter).Value = strTp_SM '平均薪資
         '薪點 =ROUND(I4/130,2)
         wksheet.Range("C" & lngCounter).Value = "=ROUND(" & strAvg_SM & lngCounter & "/130,2)"
         RsTemp.MoveNext
      Loop
      '格式化
      wksheet.Range(strFirst_SM & intStarRow + 1 & ":" & strAvg_SM & lngCounter).NumberFormatLocal = "#,##0"
      wksheet.Range("C" & intStarRow - 1 & ":" & strAvg_SM & intStarRow).HorizontalAlignment = xlRight '靠右
      wksheet.Range("A1" & ":" & strAvg_SM & intStarRow).Font.Bold = True '粗體
      '設定底色
      'wksheet.Range("C" & intStarRow - 1 & ":C" & lngCounter).Interior.ColorIndex = 6 '淺黃色
      wksheet.Range("C" & intStarRow - 1 & ":C" & lngCounter).Font.Bold = True '粗體
      'wksheet.Range(strAvg_PE & intStarRow - 1 & ":" & strAvg_PE & lngCounter).Interior.ColorIndex = 7 '紫紅色
      wksheet.Range(strAvg_PE & intStarRow - 1 & ":" & strAvg_PE & lngCounter).Font.Bold = True '粗體
      'wksheet.Range(strAvg_SM & intStarRow - 1 & ":" & strAvg_SM & lngCounter).Interior.ColorIndex = 8 '淺藍色
      wksheet.Range(strAvg_SM & intStarRow - 1 & ":" & strAvg_SM & lngCounter).Font.Bold = True '粗體
   End If
   
   ExcelSave_Detail = True
   Exit Function
   
flgErr:
   If Err.Number <> 0 Then
      MsgBox Err.Description
   End If
End Function

Private Sub txt1_GotFocus(Index As Integer)
   TextInverse txt1(Index)
   CloseIme
End Sub

Private Sub txt1_KeyPress(Index As Integer, KeyAscii As Integer)
   KeyAscii = Pub_NumAscii(KeyAscii)
End Sub

Private Sub txt1_Validate(Index As Integer, Cancel As Boolean)
   Select Case Index
      Case 0, 1
         If txt1(Index) <> "" Then
            If ChkDate(txt1(Index) & "01") = False Then
               Cancel = True
            End If
         End If
         If Index = 0 Then
            If txt1(Index) <> "" And txt1(Index + 1) = "" Then
               txt1(Index + 1) = txt1(Index)
            End If
         Else
            If RunNick(txt1(Index - 1), txt1(Index)) Then
               Cancel = True
            End If
         End If
   End Select
End Sub
