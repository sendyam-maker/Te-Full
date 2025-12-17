Attribute VB_Name = "basDate"
'Memo By Sonia 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Public strSrvDate(1 To 2) As String '1 西元 '2 民國

'910628 Sieg
'智權人員部門別為F開頭，多印一張案件性質為901之接洽結案單
Public bol901 As Boolean


'*************************************************
'  傳回伺服器日期
'
'*************************************************
Public Function ServerDate() As Long
Dim adoSysDate As New ADODB.Recordset
   adoSysDate.CursorLocation = adUseClient
   adoSysDate.Open "select to_char(sysdate, 'YYYYMMDD') from dual", cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If adoSysDate.RecordCount <> 0 Then
      ServerDate = Val(adoSysDate.Fields(0).Value)
   End If
   adoSysDate.Close
End Function

'取得民國年
Public Function GetTaiwanThisYear() As String
'GetTaiwanThisYear = Format(Val(Left(ServerDate, 4)) - 1911)
   'Modify by Morgan 2011/2/24 改直接抓變數不要再連資料庫
   'GetTaiwanThisYear = Val(Left(ServerDate, 4)) - 1911
   GetTaiwanThisYear = Val(Left(strSrvDate(1), 4)) - 1911
End Function

'取得今日之民國日期
Public Function GetTaiwanTodayDate() As String
'Modify by Morgan 2011/2/24 改直接抓變數不要再連資料庫
'GetTaiwanTodayDate = ServerDate - 19110000
GetTaiwanTodayDate = strSrvDate(2)
End Function

'取得今日之西元日期
Public Function GetTodayDate() As String
'Modify by Morgan 2011/2/24 改直接抓變數不要再連資料庫
'GetTodayDate = ServerDate
GetTodayDate = strSrvDate(1)
End Function

'檢查日期是否為民國日期(六碼或七碼)
Public Function CheckIsTaiwanDate(strDate As String, Optional bolShow As Boolean = True) As Boolean
Dim intLen As String

'add by nickc 2007/06/23 若傳進來的變數有 / 時，直接告知錯誤   以免發生其他錯誤  EX:96/0620 若不控制，將會回傳 true
If InStr(1, strDate, "/") <> 0 Then
    CheckIsTaiwanDate = False
Else
    If Val(strDate) >= 111111 Then 'Added by Morgan 2012/2/22
      
      intLen = Len(strDate)
      If intLen = 6 Then
         strDate = Format(Val(Left(strDate, 2)) + 1911) + "/" + Mid(strDate, 3, 2) + "/" + Right(strDate, 2)
      ElseIf intLen = 7 Then
         strDate = Format(Val(Left(strDate, 3)) + 1911) + "/" + Mid(strDate, 4, 2) + "/" + Right(strDate, 2)
      End If
      CheckIsTaiwanDate = IsDate(strDate)
      'Add By Cheng 2003/12/10
      If CheckIsTaiwanDate = True Then
          '若日期年份超過系統年份+30年
          If Val(Left(strDate, 4)) > (Val(Left(strSrvDate(1), 4)) + 30) Then
              CheckIsTaiwanDate = False
          End If
      End If
   'add by sonia 2016/5/17配合法務系統之回執未回LA-003238
   ElseIf Val(strDate) = 110101 Then
      CheckIsTaiwanDate = True
   'end 2016/5/17
   End If
End If
If bolShow Then
   If CheckIsTaiwanDate = False Then
      ShowMsg MsgText(9003)
   End If
End If
End Function

'檢查日期是否為西元日期
Public Function CheckIsDate(strDate As String, Optional bolShow As Boolean = True) As Boolean
If Len(strDate) = 8 Then
   If Val(strDate) >= 19221111 Then 'Added by Morgan 2012/2/22
      strDate = Format(Left(strDate, 4)) + "/" + Mid(strDate, 5, 2) + "/" + Right(strDate, 2)
      CheckIsDate = IsDate(strDate)
   End If
End If
'Add By Cheng 2003/12/10
If CheckIsDate = True Then
    '若日期年份超過系統年份+30年
    If Val(Left(strDate, 4)) > (Val(Left(strSrvDate(1), 4)) + 30) Then
        CheckIsDate = False
    End If
End If
'End
If bolShow Then
   If CheckIsDate = False Then
      ShowMsg MsgText(9003)
   End If
End If
End Function

'轉換民國日期至有加/之格式
Public Function ChangeTStringToTDateString(ByRef strTString As String) As String
Dim intLen  As Integer

If strTString = "" Then
   ChangeTStringToTDateString = ""
   Exit Function
End If
intLen = Len(strTString)
'Modify By Sindy 2011/8/5
'If intLen = 5 Or intLen = 6 Or intLen = 7 Then
   ChangeTStringToTDateString = Left(strTString, intLen - 4) + "/" + Mid(strTString, intLen - 3, 2) + "/" + Right(strTString, 2)
'End If
'If intLen = 6 Then
'   ChangeTStringToTDateString = Left(strTString, 2) + "/" + Mid(strTString, 3, 2) + "/" + Right(strTString, 2)
'ElseIf intLen = 7 Then
'   ChangeTStringToTDateString = Left(strTString, 3) + "/" + Mid(strTString, 4, 2) + "/" + Right(strTString, 2)
'End If
End Function

'轉換民國日期至沒有加/之格式
Public Function ChangeTDateStringToTString(ByRef strTDateString As String) As String
Dim intIndex1 As Integer, intIndex2 As Integer

If strTDateString = "" Then
   ChangeTDateStringToTString = ""
   Exit Function
End If
intIndex1 = InStr(strTDateString, "/")
ChangeTDateStringToTString = Left(strTDateString, intIndex1 - 1)
intIndex2 = InStr(intIndex1 + 1, strTDateString, "/")
ChangeTDateStringToTString = ChangeTDateStringToTString + Format(Mid(strTDateString, intIndex1 + 1, intIndex2 - intIndex1 - 1), "00")
ChangeTDateStringToTString = ChangeTDateStringToTString + Format(Mid(strTDateString, intIndex2 + 1), "00")
End Function

'轉換西元有加/之格式至民國日期
'Modified by Morgan 2024/1/11 修正日期轉字串會因顯示格式而不同問題
'Public Function ChangeWDateStringToTString(ByRef strWDateString As String) As String
Public Function ChangeWDateStringToTString(oDate) As String
   Dim strWDateString As String
   'Modified by Morgan 2025/3/17
   'If TypeName(oDate) = "Date" Then
   '   strWDateString = Format(oDate, "YYYY/MM/DD")
   'Else
   '   strWDateString = oDate
   'End If
   strWDateString = PUB_ChgDateToWDateStr(oDate)
   'end 2025/3/17
'end 2024/1/11
   Dim intIndex1 As Integer, intIndex2 As Integer

If strWDateString = "" Then
   ChangeWDateStringToTString = ""
   Exit Function
End If
intIndex1 = InStr(strWDateString, "/")
ChangeWDateStringToTString = Format(Val(Left(strWDateString, intIndex1 - 1)) - 1911)
intIndex2 = InStr(intIndex1 + 1, strWDateString, "/")
ChangeWDateStringToTString = ChangeWDateStringToTString + Format(Mid(strWDateString, intIndex1 + 1, intIndex2 - intIndex1 - 1), "00") + Format(Mid(strWDateString, intIndex2 + 1), "00")
End Function

'轉換民國日期至西元有加/之格式
Public Function ChangeTStringToWDateString(ByRef strTString As String) As String
Dim intLen As Integer

If strTString = "" Then
   ChangeTStringToWDateString = ""
   Exit Function
End If
intLen = Len(strTString)
If intLen = 6 Then
   ChangeTStringToWDateString = Format(Val(Left(strTString, 2)) + 1911) + "/" + Mid(strTString, 3, 2) + "/" + Right(strTString, 2)
ElseIf intLen = 7 Then
    'Modify By Cheng 2004/02/12
'   ChangeTStringToWDateString = Format(Val(Left(strTString, 3)) + 1911) + "/" + Mid(strTString, 3, 2) + "/" + Right(strTString, 2)
   ChangeTStringToWDateString = Format(Val(Left(strTString, 3)) + 1911) + "/" + Mid(strTString, 4, 2) + "/" + Right(strTString, 2)
    'End
End If
End Function

'轉換西元有加/之格式至西元日期
'Modified by Morgan 2024/1/11 修正日期轉字串會因顯示格式而不同問題
'Public Function ChangeWDateStringToWString(ByRef strWDateString As String) As String
Public Function ChangeWDateStringToWString(oDate) As String
   Dim strWDateString As String
   'Modified by Morgan 2025/3/17
   'If TypeName(oDate) = "Date" Then
   '   strWDateString = Format(oDate, "YYYY/MM/DD")
   'Else
   '   strWDateString = oDate
   'End If
   strWDateString = PUB_ChgDateToWDateStr(oDate)
   'end 2025/3/17
'end 2024/1/11
Dim intIndex1 As Integer, intIndex2 As Integer

If strWDateString = "" Then
   ChangeWDateStringToWString = ""
   Exit Function
End If
intIndex1 = InStr(strWDateString, "/")
ChangeWDateStringToWString = Format(Left(strWDateString, intIndex1 - 1))
intIndex2 = InStr(intIndex1 + 1, strWDateString, "/")
ChangeWDateStringToWString = ChangeWDateStringToWString + Format(Mid(strWDateString, intIndex1 + 1, intIndex2 - intIndex1 - 1), "00") + Format(Mid(strWDateString, intIndex2 + 1), "00")
End Function

'轉換西元日期至西元有加/之格式
Public Function ChangeWStringToWDateString(ByRef strWString As String) As String
'Modify By Sindy 2012/5/23
'If strWString = "" Then
If Val(strWString) > 0 Then
   ChangeWStringToWDateString = Format(Left(strWString, 4)) + "/" + Mid(strWString, 5, 2) + "/" + Right(strWString, 2)
Else
   ChangeWStringToWDateString = ""
   Exit Function
End If
End Function

'轉換民國日期至西元之格式
Public Function ChangeTStringToWString(ByRef strTString As String) As String
'Modify By Sindy 2012/5/23
'If strTString <> "" Then
If Val(strTString) > 0 Then
   ChangeTStringToWString = Format(Val(strTString) + 19110000)
Else
   ChangeTStringToWString = ""
End If
End Function

'add by nick 2004/11/04
'轉換西元之格式到民國
Public Function ChangeWStringToTDateString(ByRef strTString As String) As String
'Modify By Sindy 2012/5/23
'If strTString <> "" Then
If Val(strTString) > 0 Then
   ChangeWStringToTDateString = ChangeTStringToTDateString(ChangeWStringToTString(strTString))
Else
   ChangeWStringToTDateString = ""
End If
End Function

'轉換西元日期至民國之格式
Public Function ChangeWStringToTString(ByRef strWString As String) As String
'Modify By Sindy 2012/5/23
'If strWString <> "" Then
If Val(strWString) > 0 Then
   ChangeWStringToTString = Format(Val(strWString) - 19110000)
Else
   ChangeWStringToTString = ""
End If
End Function

'顯示MsgBox
Public Sub ShowMsg(strMWord As String)
   'Modified by Morgan 2024/9/19 要能顯示Unicode
   'MsgBox strMWord + "!", vbCritical + vbOKOnly, MsgText(9001)
   MsgBoxU strMWord + "!", vbCritical + vbOKOnly, MsgText(9001)
   'end 2024/9/19
End Sub

Public Function ChgPatent(ByVal strTemp As String, Optional iSitu As Integer = 0) As String
 Dim strCase(1 To 4) As String
On Error GoTo ErrHand
   If iSitu = 0 Then
      If strTemp = "" Then GoTo ErrHand
      ChgCaseNo strTemp, strCase
      ChgPatent = "PA01='" & strCase(1) & "' AND PA02='" & strCase(2) & "' AND PA03='" & strCase(3) & "' AND PA04='" & strCase(4) & "'"
   Else
      ChgPatent = "DECODE(PA03||PA04,'000',PA01||'-'||PA02,PA01||'-'||PA02||'-'||PA03||'-'||PA04)"
   End If
   Exit Function
ErrHand:
   ChgPatent = "PA01 IS NULL AND PA02 IS NULL AND PA03 IS NULL AND PA04 IS NULL"
End Function

Public Function ChgLawcase(ByVal strTemp As String, Optional iSitu As Integer = 0) As String
 Dim strCase(1 To 4) As String
On Error GoTo ErrHand
   If iSitu = 0 Then
      If strTemp = "" Then GoTo ErrHand
      ChgCaseNo strTemp, strCase
      ChgLawcase = "LC01='" & strCase(1) & "' AND LC02='" & strCase(2) & "' AND LC03='" & strCase(3) & "' AND LC04='" & strCase(4) & "'"
   Else
      ChgLawcase = "DECODE(LC03||LC04,'000',LC01||'-'||LC02,LC01||'-'||LC02||'-'||LC03||'-'||LC04)"
   End If
   Exit Function
ErrHand:
   ChgLawcase = "LC01 IS NULL AND LC02 IS NULL AND LC03 IS NULL AND LC04 IS NULL"
End Function

Public Function ChgCaseprogress(ByVal strTemp As String, Optional iSitu As Integer = 0) As String
 Dim strCase(1 To 4) As String
On Error GoTo ErrHand
   If iSitu = 0 Then
      If strTemp = "" Then GoTo ErrHand
      ChgCaseNo strTemp, strCase
      ChgCaseprogress = "CP01='" & strCase(1) & "' AND CP02='" & strCase(2) & "' AND CP03='" & strCase(3) & "' AND CP04='" & strCase(4) & "'"
   Else
      ChgCaseprogress = "DECODE(CP03||CP04,'000',CP01||'-'||CP02,CP01||'-'||CP02||'-'||CP03||'-'||CP04)"
   End If
   Exit Function
ErrHand:
   ChgCaseprogress = "CP01 IS NULL AND CP02 IS NULL AND CP03 IS NULL AND CP04 IS NULL"
End Function

Public Function ChgMailRec(ByVal strTemp As String, Optional iSitu As Integer = 0) As String
 Dim strCase(1 To 4) As String
On Error GoTo ErrHand
   If iSitu = 0 Then
      If strTemp = "" Then GoTo ErrHand
      ChgCaseNo strTemp, strCase
      ChgMailRec = "MR12='" & strCase(1) & "' AND MR13='" & strCase(2) & "' AND MR14='" & strCase(3) & "' AND MR15='" & strCase(4) & "'"
   Else
      ChgMailRec = "DECODE(MR14||MR15,'000',MR12||'-'||MR13,MR12||'-'||MR13||'-'||MR14||'-'||MR15)"
   End If
   Exit Function
ErrHand:
   ChgMailRec = "MR12 IS NULL AND MR13 IS NULL AND MR14 IS NULL AND MR15 IS NULL"
End Function

Public Function ChgHirecase(ByVal strTemp As String, Optional iSitu As Integer = 0) As String
 Dim strCase(1 To 4) As String
On Error GoTo ErrHand
   If iSitu = 0 Then
      If strTemp = "" Then GoTo ErrHand
      ChgCaseNo strTemp, strCase
      ChgHirecase = "HC01='" & strCase(1) & "' AND HC02='" & strCase(2) & "' AND HC03='" & strCase(3) & "' AND HC04='" & strCase(4) & "'"
   Else
      ChgHirecase = "DECODE(HC03||HC04,'000',HC01||'-'||HC02,HC01||'-'||HC02||'-'||HC03||'-'||HC04)"
   End If
   Exit Function
ErrHand:
   ChgHirecase = "HC01 IS NULL AND HC02 IS NULL AND HC03 IS NULL AND HC04 IS NULL"
End Function

Public Function ChgService(ByVal strTemp As String, Optional iSitu As Integer = 0) As String
 Dim strCase(1 To 4) As String
On Error GoTo ErrHand
   If iSitu = 0 Then
      If strTemp = "" Then GoTo ErrHand
      ChgCaseNo strTemp, strCase
      ChgService = "SP01='" & strCase(1) & "' AND SP02='" & strCase(2) & "' AND SP03='" & strCase(3) & "' AND SP04='" & strCase(4) & "'"
   Else
      ChgService = "DECODE(SP03||SP04,'000',SP01||'-'||SP02,SP01||'-'||SP02||'-'||SP03||'-'||SP04)"
   End If
   Exit Function
ErrHand:
   ChgService = "SP01 IS NULL AND SP02 IS NULL AND SP03 IS NULL AND SP04 IS NULL"
End Function

Public Function ChgCaseMap(ByVal strTemp As String, Optional iSitu As Integer = 0, Optional iSitu1 As Integer = 0) As String
 Dim strCase(1 To 4) As String
On Error GoTo ErrHand
   If iSitu = 0 Then
      If strTemp = "" Then GoTo ErrHand
      ChgCaseNo strTemp, strCase
      If iSitu1 = 0 Then
         ChgCaseMap = "CM01='" & strCase(1) & "' AND CM02='" & strCase(2) & "' AND CM03='" & strCase(3) & "' AND CM04='" & strCase(4) & "'"
      Else
         ChgCaseMap = "CM05='" & strCase(1) & "' AND CM06='" & strCase(2) & "' AND CM07='" & strCase(3) & "' AND CM08='" & strCase(4) & "'"
      End If
   Else
      If iSitu1 = 0 Then
         ChgCaseMap = "DECODE(CM03||CM04,'000',CM01||'-'||CM02,CM01||'-'||CM02||'-'||CM03||'-'||CM04)"
      Else
         ChgCaseMap = "DECODE(CM07||CM08,'000',CM05||'-'||CM06,CM05||'-'||CM06||'-'||CM07||'-'||CM08)"
      End If
   End If
   Exit Function
ErrHand:
   If iSitu1 = 0 Then
      ChgCaseMap = "CM01 IS NULL AND CM02 IS NULL AND CM03 IS NULL AND CM04 IS NULL"
   Else
      ChgCaseMap = "CM05 IS NULL AND CM06 IS NULL AND CM07 IS NULL AND CM08 IS NULL"
   End If
End Function

Public Function ChgCustomer(ByVal strTemp As String) As String
 On Error GoTo ErrHand
   If strTemp = "" Then GoTo ErrHand
   If Len(strTemp) = 9 Then
      ChgCustomer = "CU01='" & Left(strTemp, 8) & "' AND CU02='" & Right(strTemp, 1) & "'"
   Else
      ChgCustomer = "CU01='" & strTemp & String(8 - Len(strTemp), "0") & "' AND CU02='0'"
   End If
   Exit Function
ErrHand:
   ChgCustomer = "CU01 IS NULL AND CU02 IS NULL"
End Function

Public Function ChgFagent(ByVal strTemp As String) As String
 On Error GoTo ErrHand
   If strTemp = "" Then GoTo ErrHand
   If Len(strTemp) = 9 Then
      ChgFagent = "FA01='" & Left(strTemp, 8) & "' AND FA02='" & Right(strTemp, 1) & "'"
   Else
      ChgFagent = "FA01='" & strTemp & String(8 - Len(strTemp), "0") & "' AND FA02='0'"
   End If
   Exit Function
ErrHand:
   ChgFagent = "FA01 IS NULL AND FA02 IS NULL"
End Function

Public Function ChgNextProgress(ByVal strTemp As String, Optional iSitu As Integer = 0) As String
 Dim strCase(1 To 4) As String
On Error GoTo ErrHand
   If iSitu = 0 Then
      If strTemp = "" Then GoTo ErrHand
      ChgCaseNo strTemp, strCase
      ChgNextProgress = "NP02='" & strCase(1) & "' AND NP03='" & strCase(2) & "' AND NP04='" & strCase(3) & "' AND NP05='" & strCase(4) & "'"
   Else
      ChgNextProgress = "DECODE(NP04||NP05,'000',NP02||'-'||NP03,NP02||'-'||NP03||'-'||NP04||'-'||NP05)"
   End If
   Exit Function
ErrHand:
   ChgNextProgress = "NP02 IS NULL AND NP03 IS NULL AND NP04 IS NULL AND NP05 IS NULL"
End Function

Public Function ChgPriDate(ByVal strTemp As String, Optional iSitu As Integer = 0) As String
 Dim strCase(1 To 4) As String
On Error GoTo ErrHand
   If iSitu = 0 Then
      If strTemp = "" Then GoTo ErrHand
      ChgCaseNo strTemp, strCase
      ChgPriDate = "PD01='" & strCase(1) & "' AND PD02='" & strCase(2) & "' AND PD03='" & strCase(3) & "' AND PD04='" & strCase(4) & "'"
   Else
      ChgPriDate = "DECODE(PD03||PD04,'000',PD01||'-'||PD02,PD01||'-'||PD02||'-'||PD03||'-'||PD04)"
   End If
   Exit Function
ErrHand:
   ChgPriDate = "PD01 IS NULL AND PD02 IS NULL AND PD03 IS NULL AND PD04 IS NULL"
End Function

Public Function ChgTradeMark(ByVal strTemp As String, Optional iSitu As Integer = 0) As String
 Dim strCase(1 To 4) As String
On Error GoTo ErrHand
   If iSitu = 0 Then
      If strTemp = "" Then GoTo ErrHand
      ChgCaseNo strTemp, strCase
      ChgTradeMark = "TM01='" & strCase(1) & "' AND TM02='" & strCase(2) & "' AND TM03='" & strCase(3) & "' AND TM04='" & strCase(4) & "'"
   Else
      ChgTradeMark = "DECODE(TM03||TM04,'000',TM01||'-'||TM02,TM01||'-'||TM02||'-'||TM03||'-'||TM04)"
   End If
   Exit Function
ErrHand:
   ChgTradeMark = "TM01 IS NULL AND TM02 IS NULL AND TM03 IS NULL AND TM04 IS NULL"
End Function

Public Sub ChgCaseNo(ByVal strTemp As String, ByRef strCase() As String)
 Dim i As Integer
   strTemp = UCase(strTemp)
   strTemp = Replace(strTemp, "-", "") 'Added by Lydia 2025/09/03
   i = 1
   Do While Not IsNumeric(Mid(strTemp, i, 1))
      i = i + 1
      'Add by Morgan 2005/6/20
      If i > 3 Then Exit Do
   Loop
   'Added by Lydia 2017/12/11 若輸入非案號，結果為負數會造成程式中斷
   If 8 + i - Len(strTemp) < 0 Then
        strCase(1) = Left(strTemp, i - 1)
        strCase(2) = "": strCase(3) = "": strCase(4) = ""
   Else
   'end 2017/12/11
        strTemp = strTemp & String(8 + i - Len(strTemp), "0")
        strCase(1) = Left(strTemp, i - 1)
        strCase(2) = Mid(strTemp, i, 6)
        strCase(3) = Mid(strTemp, i + 6, 1)
        strCase(4) = Right(strTemp, 2)
   End If  'end 2017/12/11
End Sub

Public Function CompDate(ByVal iSitu As Integer, ByVal iNum As Single, ByVal strTemp As String) As String
 Dim i As Integer, s As Single, strTmp As String
   If strTemp = "" Then CompDate = "": Exit Function
   If Len(strTemp) <> 8 Then strTemp = Format(Val(strTemp) + 19110000)
   Select Case iSitu
      Case 0 '年
         strTmp = Format(iNum)
         'Modify by Morgan 2005/7/22
         'CompDate = Format(Val(Left(strTemp, 4)) + iNum) & Mid(strTemp, 5, 2) & Right(strTemp, 2)
         CompDate = Format(DateAdd("yyyy", Int(iNum), DateSerial(Left(strTemp, 4), Mid(strTemp, 5, 2), Right(strTemp, 2))), "YYYYMMDD")
         i = InStr(strTmp, ".")
         If i > 0 Then
            s = (Val(iNum) - Int(iNum)) * 12
            CompDate = CompDate(1, s, CompDate)
         End If
      Case 1 '月
         CompDate = Format(DateAdd("M", iNum, DateSerial(Left(strTemp, 4), Mid(strTemp, 5, 2), Right(strTemp, 2))), "YYYYMMDD")
      Case 2 '日
         CompDate = Format(DateAdd("D", iNum, DateSerial(Left(strTemp, 4), Mid(strTemp, 5, 2), Right(strTemp, 2))), "YYYYMMDD")
   End Select
End Function

'Modify By Sindy 2017/7/28 + Optional ByVal bolShowMsg As Boolean = True
Public Function ChkDate(ByVal strTemp As String, Optional ByVal bolShowMsg As Boolean = True) As Boolean
 Dim i As Integer
On Error GoTo ErrHand
   ChkDate = False
   If Len(strTemp) = 6 Or Len(strTemp) = 7 Then '民國
      If Val(strTemp) >= 111111 Then 'Added by Morgan 2012/2/22
         i = Len(strTemp) - 4
         ChkDate = IsDate(Left(strTemp, i) + 1911 & "/" & Mid(strTemp, i + 1, 2) + "/" + Right(strTemp, 2))
      End If
   ElseIf Len(strTemp) = 8 Then '西元
      If Val(strTemp) >= 19221111 Then 'Added by Morgan 2012/2/22
         ChkDate = IsDate(Left(strTemp, 4) & "/" & Mid(strTemp, 5, 2) + "/" + Right(strTemp, 2))
      End If
   End If
    'Add By Cheng 2003/12/10
    If ChkDate = True Then
        '若日期年份超過系統年份+30年
        '2011/5/24 MODIFY BY SONIA 否則輸入超過30年的日期在PUB_DBDATE就已發生錯誤
        'If Val(Left(PUB_DBDATE(strTemp), 4)) > (Val(Left(strSrvDate(1), 4)) + 30) Then
        If Val(Left(TransDate(strTemp, 2), 4)) > (Val(Left(strSrvDate(1), 4)) + 30) Then
            ChkDate = False
        End If
    End If
    'End
   If ChkDate = False Then
      'Modify By Sindy 2017/7/28
      If bolShowMsg = True Then
      '2017/7/28 END
         MsgBox "日期錯誤，請重新輸入 !", vbCritical
      End If
   End If
   
   Exit Function
ErrHand:
   MsgBox "日期錯誤，請重新輸入 !", vbCritical
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'轉換日期 西元或民國年轉民國或西元 (iSitu=1 西元轉民國 iSitu=2 民國轉西元)
' Input : strTemp ==> 所傳入的日期
'         iSitu ==> 1 : 表轉成民國年
'                   2 : 表轉成西元年
' Output : 傳回轉換後的日期
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TransDate(ByVal strTemp As String, ByVal iSitu As Integer) As String
   Dim strDate As String
   If strTemp = "" Then TransDate = "": Exit Function
   ' 90.07.03 modify by louis 加日期自動轉換
   'Select Case iSitu
   '   Case 1
   '      TransDate = Format(Val(strTemp) - 19110000)
   '   Case 2
   '      TransDate = Format(Val(strTemp) + 19110000)
   'End Select
   TransDate = strTemp
   Select Case iSitu
      Case 1
         If Len(strTemp) = 8 Then TransDate = Format(Val(strTemp) - 19110000)
      Case 2
         If Len(strTemp) <> 8 Then TransDate = Format(Val(strTemp) + 19110000)
   End Select
End Function

'bolDate True 台灣 False 西元
Public Function SQLDate(ByVal strTemp As String, Optional bolDate As Boolean = True) As String
   If bolDate Then
      'Modfied by Lydia 2018/02/22 傳入日期有可能是0
      SQLDate = "DECODE(" & strTemp & ",'','','0','',SUBSTR(" & strTemp & ",1,4)-1911||'/'||SUBSTR(" & strTemp & ",5,2)||'/'||SUBSTR(" & strTemp & ",7,2))"
   Else
      'Modfied by Lydia 2018/02/22 傳入日期有可能是0
      SQLDate = "DECODE(" & strTemp & ",'','','0','',SUBSTR(" & strTemp & ",1,4)||'/'||SUBSTR(" & strTemp & ",5,2)||'/'||SUBSTR(" & strTemp & ",7,2))"
   End If
End Function

'取得某月份天數
Public Function PUB_GetMonthDays(intYear As Integer, intMonth As Integer) As Integer
Dim ii As Integer
PUB_GetMonthDays = 1
For ii = 1 To 31
   If Not IsDate(intYear & "/" & intMonth & "/" & ii) Then
      PUB_GetMonthDays = ii - 1
       Exit For
   End If
   PUB_GetMonthDays = ii
Next ii
End Function

Public Function PUB_CheckKeyInDate(obj As Object) As Integer
'obj : 傳入要檢查的控制項
On Error GoTo ErrorHandler

PUB_CheckKeyInDate = 0
With obj
   If Len(.Text) > 0 Then
      If IsNumeric(.Text) = True Then
         'Added by Morgan 2012/2/22
         If Val(.Text) < 111111 Then
            MsgBox "日期輸入錯誤!!!", vbExclamation
            PUB_CheckKeyInDate = -1
         Else
         'End 2012/2/22
            Select Case Len(.Text)
            Case 6 '6碼
               If Not IsDate(Format((Mid(.Text, 1, 2) + 1911), "0000") & "/" & Format(Mid(.Text, 3, 2), "00") & "/" & Format(Mid(.Text, 5, 2), "00")) Then
                  MsgBox "日期格式輸入錯誤!!!", vbExclamation
                  PUB_CheckKeyInDate = -1
               End If
            Case 7 '7碼
               If Not IsDate(Format((Mid(.Text, 1, 3) + 1911), "0000") & "/" & Format(Mid(.Text, 4, 2), "00") & "/" & Format(Mid(.Text, 6, 2), "00")) Then
                  MsgBox "日期格式輸入錯誤!!!", vbExclamation
                  PUB_CheckKeyInDate = -1
               End If
            Case Else
               MsgBox "日期格式輸入錯誤!!!", vbExclamation
               PUB_CheckKeyInDate = -1
            End Select
            
         End If
      Else
         MsgBox "日期格式輸入錯誤!!!", vbExclamation
         PUB_CheckKeyInDate = -1
      End If
   End If
End With
Exit Function
ErrorHandler:
   MsgBox "日期格式輸入錯誤!!!", vbExclamation
   PUB_CheckKeyInDate = -1
End Function

Public Function PUB_CheckKeyInYYMM(obj As Object) As Integer
'obj : 傳入要檢查的控制項
On Error GoTo ErrorHandler

PUB_CheckKeyInYYMM = 0
With obj
   If Len(.Text) > 0 Then
      If IsNumeric(.Text) = True Then
         'Added by Morgan 2012/2/22
         If Val(.Text) < 1111 Then
            MsgBox "年月輸入錯誤!!!", vbExclamation
            PUB_CheckKeyInYYMM = -1
         Else
         'End 2012/2/22
            Select Case Len(.Text)
            Case 4 '4碼
               If Not IsDate(Format((Mid(.Text, 1, 2) + 1911), "0000") & "/" & Format(Mid(.Text, 3, 2), "00") & "/01") Then
                  MsgBox "年月格式輸入錯誤!!!", vbExclamation
                  PUB_CheckKeyInYYMM = -1
               End If
            Case 5 '5碼
               If Not IsDate(Format((Mid(.Text, 1, 3) + 1911), "0000") & "/" & Format(Mid(.Text, 4, 2), "00") & "/01") Then
                  MsgBox "年月格式輸入錯誤!!!", vbExclamation
                  PUB_CheckKeyInYYMM = -1
               End If
            Case Else
               MsgBox "年月格式輸入錯誤!!!", vbExclamation
               PUB_CheckKeyInYYMM = -1
            End Select
         End If
      Else
         MsgBox "年月格式輸入錯誤!!!", vbExclamation
         PUB_CheckKeyInYYMM = -1
      End If
   End If
End With
Exit Function
ErrorHandler:
   MsgBox "日期格式輸入錯誤!!!", vbExclamation
   PUB_CheckKeyInYYMM = -1
End Function

'由法定期限計算本所期限
'Input strDate(1) 系統類別 strDate(2) 申請國家 strDate(3) 法定期限
'Output strDate(0) 本所期限
'Output strDate(4) 約定期限 Add By Sindy 2021/8/17
Public Sub GetCtrlDT(ByRef strDate() As String)
   Select Case strDate(1)
      Case "CFT"
         strDate(0) = CompDate(1, -1, strDate(3))
      'Modify By Cheng 2002/07/31
'      Case "CFP"
      Case "CFP", "CPS"
         strDate(0) = CompDate(2, -14, strDate(3))
      Case "T"
         If strDate(2) = "238" Then
            strDate(0) = CompDate(1, -1, strDate(3))
         Else
            strDate(0) = CompDate(2, -2, strDate(3))
         End If
      Case "P"
        '申請國家為台灣
         If strDate(2) = "000" Then
            '本所期限
            'Modified by Morgan 2014/10/27
            'strDate(0) = CompDate(2, -2, strDate(3))
            If strSrvDate(1) >= 台灣案所限新規則啟用日 Then
               strDate(0) = PUB_GetOurDeadline(strDate(3))
            Else
               strDate(0) = CompDate(2, -2, strDate(3))
            End If
            'end 2014/10/27
         'Added by Lydia 2025/10/29 改用模組
         'Modified by Lydia 2025/10/31 非台灣案(大陸、香港、澳門和PCT)都是相同算法---郭經理(電話)
         'ElseIf strDate(2) = "020" Then
         ElseIf InStr("020,013,044,056", strDate(2)) > 0 Then
            strDate(0) = PUB_GetPOurDeadline(strDate(3), "020")
         'end 2025/10/29
         Else
            strDate(0) = CompDate(1, -1, strDate(3))
            strDate(0) = CompDate(2, -5, strDate(0))
         End If
      Case Else
         'Added by Morgan 2014/11/7
         'Modified by Morgan 2014/11/20 外專改回舊規則
         If strDate(2) = "000" And strDate(1) <> "FCP" And strDate(1) <> "FG" Then
            strDate(0) = PUB_GetOurDeadline(strDate(3))
         'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
         ElseIf strSrvDate(1) >= 外專台灣案所限新規則啟用日 And strDate(2) = "000" And (strDate(1) = "FCP" Or strDate(1) = "FG") Then
            strDate(0) = PUB_GetFCPOurDeadline(strDate(3), 2)
            'Add By Sindy 2021/8/17 + , , strDate(4)
            'strDate(0) = PUB_GetFCPOurDeadline(strDate(3), 2, , strDate(4))
         'end 2019/7/11
         Else
         'end 2014/11/7
            strDate(0) = CompDate(2, -2, strDate(3))
         End If 'Added 2014/11/7
   End Select
End Sub

Public Function GetPA04(pa01 As String, pa02 As String, pa03 As String, PA09 As String) As String
 Dim strSql As String, adoRst As New ADODB.Recordset
   GetPA04 = ""
   strSql = "SELECT PA04 FROM PATENT WHERE PA01='" & pa01 & "' AND PA02='" & pa02 & "' AND PA03='" & pa03 & "' AND PA09='" & PA09 & "'"
   adoRst.CursorLocation = adUseClient
   adoRst.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If adoRst.RecordCount > 0 Then
      GetPA04 = adoRst.Fields("PA04")
   End If
   adoRst.Close
End Function

'Add By Cheng 2002/12/16
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得西元的日期
' Input : strDate ==> 輸入的日期
' Output : 傳回西元的日期 YYYYMMDD
' Description : 此功能會傳回西元日期, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PUB_DBDATE(ByVal strDate As String) As String
   Dim strTemp As String
   
   strDate = Replace(strDate, "-", "/") 'Added by Morgan 2025/3/18
   PUB_DBDATE = strDate
   ' 若日期中有/表示這是日期格式的字串如 YY/MM/DD 或 YYYY/MM/DD
   If InStr(1, strDate, "/") Then
      PUB_DBDATE = PUB_DBYEAR(strDate) & PUB_DBMONTH(strDate) & PUB_DBDAY(strDate)
   Else
      If Len(strDate) > 7 Then
         strTemp = strDate
         If CheckIsDate(strTemp, False) = True Then
            PUB_DBDATE = strDate
         End If
      Else
         strTemp = strDate
         If CheckIsTaiwanDate(strTemp, False) = True Then
            strTemp = strDate
            PUB_DBDATE = ChangeTStringToWString(strTemp)
         End If
      End If
   End If
End Function

'Add By Cheng 2002/12/16
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得西元的年份
' Input : strDate ==> 輸入的日期
' Output : 傳回西元的年份 YYYY
' Description : 此功能會傳回日期字串中的年, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PUB_DBYEAR(ByVal strDate As String) As String
   Dim nIndex As Integer
   Dim strTemp As String
   
   strDate = Replace(strDate, "-", "/") 'Added by Morgan 2025/3/18
   ' 若日期中有/表示這是日期格式的字串如 YY/MM/DD 或 YYYY/MM/DD
   PUB_DBYEAR = "0000"
   If InStr(1, strDate, "/") Then
      nIndex = InStr(1, strDate, "/")
      strTemp = Mid(strDate, 1, nIndex - 1)
      If Len(strTemp) < 4 Then
         strTemp = str(Val(strTemp) + 1911)
      End If
      PUB_DBYEAR = strTemp
   Else
      If Len(strDate) > 7 Then
         PUB_DBYEAR = Mid(strDate, 1, 4)
      Else
         strTemp = strDate
         PUB_DBYEAR = Mid(ChangeTStringToWString(strTemp), 1, 4)
      End If
   End If
End Function

'Add By Cheng 2002/12/16
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得西元的月份
' Input : strDate ==> 輸入的日期
' Output : 傳回西元的月份 MM
' Description : 此功能會傳回日期字串中的年, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PUB_DBMONTH(ByVal strDate As String) As String
   Dim nBegin As Integer
   Dim nEnd As Integer
   Dim strTemp As String
   
   strDate = Replace(strDate, "-", "/") 'Added by Morgan 2025/3/18
   PUB_DBMONTH = "00"
   ' 若日期中有/表示這是日期格式的字串如 YY/MM/DD 或 YYYY/MM/DD
   If InStr(1, strDate, "/") Then
      nBegin = InStr(1, strDate, "/")
      nEnd = InStr(nBegin + 1, strDate, "/")
      strTemp = Mid(strDate, nBegin + 1, nEnd - nBegin - 1)
      If Len(strTemp) = 1 Then
         PUB_DBMONTH = "0" & strTemp
      Else
         PUB_DBMONTH = strTemp
      End If
   Else
      If Len(strDate) > 7 Then
         PUB_DBMONTH = Mid(strDate, 5, 2)
      Else
         strTemp = strDate
         PUB_DBMONTH = Mid(ChangeTStringToWString(strTemp), 5, 2)
      End If
   End If
End Function

'Add By Cheng 2002/12/16
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得西元的日數
' Input : strDate ==> 輸入的日期
' Output : 傳回西元的日數 DD
' Description : 此功能會傳回日期字串中的年, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function PUB_DBDAY(ByVal strDate As String) As String
   Dim nTemp As Integer
   Dim nIndex As Integer
   Dim strTemp As String
   
   strDate = Replace(strDate, "-", "/") 'Added by Morgan 2025/3/18
   PUB_DBDAY = "00"
   ' 若日期中有/表示這是日期格式的字串如 YY/MM/DD 或 YYYY/MM/DD
   If InStr(1, strDate, "/") Then
      nTemp = InStr(1, strDate, "/")
      nIndex = InStr(nTemp + 1, strDate, "/")
      strTemp = Mid(strDate, nIndex + 1, Len(strDate) - nIndex)
      If Len(strTemp) = 1 Then
         PUB_DBDAY = "0" & strTemp
      Else
         PUB_DBDAY = strTemp
      End If
   Else
      If Len(strDate) > 7 Then
         PUB_DBDAY = Mid(strDate, 7, 2)
      Else
         strTemp = strDate
         PUB_DBDAY = Mid(ChangeTStringToWString(strTemp), 7, 2)
      End If
   End If
End Function

'Add by Morgan 2004/12/21 月底轉換
'stDate1：運算前日期 stDate2：運算後日期
Public Function PUB_LastDayConvert(ByVal stDate1 As String, ByRef stDate2 As String)
   Dim stDate3 As String
   Dim stYear As String, stMonth As String, StDay As String
   stDate3 = DateAdd("M", 1, ChangeWStringToWDateString(stDate1))
   stDate3 = Format(DateAdd("D", -1 * Day(stDate3), stDate3), "YYYYMMDD")
   If stDate1 = stDate3 Then
      stDate3 = DateAdd("M", 1, ChangeWStringToWDateString(stDate2))
      stDate2 = Format(DateAdd("D", -1 * Day(stDate3), stDate3), "YYYYMMDD")
   End If
End Function

'*************************************************
'  西元轉民國年
'
'*************************************************
Public Function ACDate(InputDate As String) As String
   If (Val(Mid(InputDate, 1, 4)) - 1911) >= 100 Then
'      ACDate = Trim(Str(Val(Mid(InputDate, 1, 4)) - 1911)) & "/" & Mid(InputDate, 5, 2) & "/" & Mid(InputDate, 7, 2)
      ACDate = Trim(str(Val(Mid(InputDate, 1, 4)) - 1911)) & Mid(InputDate, 5, 4)
   Else
'      ACDate = "0" & Trim(Str(Val(Mid(InputDate, 1, 4)) - 1911)) & "/" & Mid(InputDate, 5, 2) & "/" & Mid(InputDate, 7, 2)
      ACDate = "0" & Trim(str(Val(Mid(InputDate, 1, 4)) - 1911)) & Mid(InputDate, 5, 4)
   End If
End Function

'*************************************************
'  將日期轉換為顯示格式
'
'*************************************************
Public Function CFDate(InputDate As String) As String
   If Len(InputDate) > 6 Then
      CFDate = Mid(InputDate, 1, 3) & "/" & Mid(InputDate, 4, 2) & "/" & Mid(InputDate, 6, 2)
   Else
      CFDate = "0" & Mid(InputDate, 1, 2) & "/" & Mid(InputDate, 3, 2) & "/" & Mid(InputDate, 5, 2)
   End If
End Function

'*************************************************
'  民國轉西元年
'
'*************************************************
Public Function CADate(InputDate As String) As String
'   CADate = Trim(Str(Val(Mid(InputDate, 1, 3)) + 1911)) & Mid(InputDate, 5, 2) & Mid(InputDate, 8, 2)
   If Len(InputDate) > 6 Then
      CADate = Trim(str(Val(Mid(InputDate, 1, 3)) + 1911)) & Mid(InputDate, 4, 4)
   Else
      CADate = Trim(str(Val(Mid(InputDate, 1, 2)) + 1911)) & Mid(InputDate, 3, 4)
   End If
End Function

'*************************************************
'  將日期轉換為顯示格式 (西元)
'
'*************************************************
Public Function AFDate(InputDate As String) As String
   AFDate = Mid(InputDate, 1, 4) & "/" & Mid(InputDate, 5, 2) & "/" & Mid(InputDate, 7, 2)
End Function

'*************************************************
'  將日期轉換為儲存格式
'
'*************************************************
Public Function FCDate(InputDate As String) As String
   If Len(InputDate) > 8 Then
      FCDate = Mid(InputDate, 1, 3) & Mid(InputDate, 5, 2) & Mid(InputDate, 8, 2)
   Else
      FCDate = Mid(InputDate, 1, 2) & Mid(InputDate, 4, 2) & Mid(InputDate, 7, 2)
   End If
End Function

'*************************************************
'  傳回伺服器時間
'
'*************************************************
Public Function ServerTime() As Long
Dim adoSysTime As New ADODB.Recordset
   adoSysTime.CursorLocation = adUseClient
   adoSysTime.Open "select to_char(sysdate, 'HH24MISS') from dual", cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If adoSysTime.RecordCount <> 0 Then
      ServerTime = Val(adoSysTime.Fields(0).Value)
   End If
   adoSysTime.Close
End Function

'*************************************************
'  日期檢核
'
'*************************************************
'Modify by Morgan 2008/10/1 改寫成民國西元有無斜線皆通用
Public Function DateCheck(p_strDate As String) As String
   Dim strDate As String
   strDate = Replace(p_strDate, "/", "")
   strDate = DBDATE(strDate)
   'Added by Morgan 2012/2/22
   If Val(strDate) < 19221111 Then
      DateCheck = MsgText(603)
   Else
      strDate = Format(strDate, "####/##/##")
      If IsDate(strDate) = False Then
         DateCheck = MsgText(603)
         Exit Function
      End If
      DateCheck = MsgText(602)
   End If
End Function

'*************************************************
'  計算天數
'
'*************************************************
Public Function CDays(StartDate As String, EndDate As String) As Long
Dim Sdate As String, Edate As String

   'Modified by Morgan 2024/10/28
   'Sdate = Val(StartDate) + 19110000
   'Edate = Val(EndDate) + 19110000
   Sdate = DBDATE(StartDate)
   Edate = DBDATE(EndDate)
   'end 2024/10/28
   Sdate = AFDate(Sdate)
   Edate = AFDate(Edate)
   CDays = DateDiff("d", Sdate, Edate)
End Function

'*************************************************
'  取日期月份的最後1天的日期
'
'*************************************************
'Added by Morgan 2013/1/28
Public Function GetLastDay(pDate As String) As String
   Dim stDate As String
   stDate = DBDATE(pDate)
   stDate = CompDate(1, 1, Left(stDate, 6) & "01")
   stDate = CompDate(2, -1, stDate)
   GetLastDay = stDate
End Function

'Modified by Morgan 2020/2/13 從basFunction移來
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 增加月數或減少月數
' Input : strDate ==> 原始的日期
'         nIncrement ==> 加減的月數 (正數表加, 負數表減)
' Output : 傳回經過運算後的日期
' Remark : 當加或減某一月數後的日期, 若其非正確的日期,
'          則以當月的最後一天為正確日期
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function AddMonth(ByVal strDate As String, ByVal nIncrement As Integer) As String
   Dim nYear As Integer
   Dim nMonth As Integer
   Dim nDay As Integer
   Dim nYearCount As Integer
   Dim nMonthCount As Integer
   Dim strNewDate As String
   
   nYear = CInt(DBYEAR(DBDATE(strDate)))
   nMonth = CInt(DBMONTH(DBDATE(strDate)))
   nDay = CInt(DBDAY(DBDATE(strDate)))
  
   nMonth = nMonth + nIncrement
   If nMonth >= 0 Then
      '911021 NICK 大於12時再減或除
      If nMonth > 12 Then
         'Modified by Morgan 2020/9/22 bug
         'nYearCount = nMonth / 12
         nYearCount = nMonth \ 12
         'end 2020/9/22
         nMonth = nMonth Mod 12
         nYear = nYear + nYearCount
      End If
   Else
      Do While nMonth < 0
         nMonth = nMonth + 12
         nYear = nYear - 1
      Loop
   End If
   
   Do While True
      Dim strYear As String
      Dim strMonth As String
      Dim strDay As String
      strYear = String(4 - Len(CStr(nYear)), "0") & CStr(nYear)
      strMonth = String(2 - Len(CStr(nMonth)), "0") & CStr(nMonth)
      strDay = String(2 - Len(CStr(nDay)), "0") & CStr(nDay)
      strNewDate = strYear & "/" & strMonth & "/" & strDay
      If Not IsDate(strNewDate) Then
         nDay = nDay - 1
      Else
         Exit Do
      End If
   Loop
   
   AddMonth = DBDATE(strNewDate)
End Function

'Added by Morgan 2023/6/28
Public Function PUB_LockRsvDN() As Boolean
   Dim iRec As Integer
   adoTaie.Execute "update PrintStartPoint set psp03=psp03 where psp01 = '" & strUserNum & "' and psp02='預留請款單號'", iRec
   If iRec > 0 Then
      PUB_LockRsvDN = True
   End If
End Function
'Added by Morgan 2023/6/28
'預留請款單號
Public Function PUB_RsvDN(pNum As Integer, pNo1 As String, pNo2 As String) As Boolean
   Dim iRec As Integer
   
   adoTaie.BeginTrans
   
On Error GoTo ErrHnd
   
   'Added by Morgan 2023/8/9
   If PUB_LockRsvDN() = True Then
      adoTaie.RollbackTrans
      MsgBox "目前已有預留請款單號，請確認！", vbExclamation
      Exit Function
   End If
   'end 2023/8/9
   
   adoTaie.Execute "update acc1r0 set a1r04 = a1r04 where a1r01 = 'X'"
   pNo1 = AccAutoNo("X", 5)
   If AccSaveAutoNo("X", Right(strExc(1), 5)) = "Y" Then
      If pNum > 1 Then
         pNo2 = "X" & (Right(pNo1, 8) - 1 + pNum)
         adoTaie.Execute "update acc1r0 set a1r04 = a1r04+" & (pNum - 1) & " where a1r01 = 'X'"
      Else
         pNo2 = pNo1
      End If
      adoTaie.Execute "delete PrintStartPoint where psp01 = '" & strUserNum & "' and psp02='預留請款單號'", iRec
      adoTaie.Execute "insert into PrintStartPoint(psp01,psp02,psp03,psp06) values('" & strUserNum & "','預留請款單號','" & pNo1 & "','" & pNo2 & "')", iRec
   End If
   adoTaie.CommitTrans
   PUB_RsvDN = True
   Exit Function
   
ErrHnd:
   adoTaie.RollbackTrans
   MsgBox Err.Description, vbCritical
   
End Function
'讀取預留請款單號
Public Function PUB_GetRsvDN(pNo1 As String, pNo2 As String) As Boolean
   Dim intQ As Integer, strQ As String
   Dim rstQ As ADODB.Recordset
   
   strQ = "select psp03,psp06 from PrintStartPoint where psp01='" & strUserNum & "' and psp02='預留請款單號' and psp06>=psp03"
   intQ = 1
   Set rstQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      pNo1 = rstQ("psp03")
      pNo2 = rstQ("psp06")
      PUB_GetRsvDN = True
   End If
   Set rstQ = Nothing
End Function
'更新預留請款單號
Public Function PUB_UpdRsvDN(pNo As String) As Boolean
   Dim iRec As Integer
   adoTaie.Execute "delete PrintStartPoint where psp01='" & strUserNum & "' and psp02='預留請款單號' and psp03=psp06 and psp03='" & pNo & "'", iRec
   If iRec = 0 Then
      adoTaie.Execute "update PrintStartPoint set psp03='X'||(substr(psp03, 2) + 1) where psp01='" & strUserNum & "' and psp02='預留請款單號' and psp06>psp03 and psp03='" & pNo & "'", iRec
   End If
   If iRec = 1 Then
      PUB_UpdRsvDN = True
   End If
End Function
'end 2023/6/28

'Add By Cheng 2002/11/01
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得西元的日期
' Input : strDate ==> 輸入的日期
' Output : 傳回西元的日期 YYYYMMDD
' Description : 此功能會傳回西元日期, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modified by Morgan 2024/1/11 修正日期轉字串會因顯示格式而不同問題
'Public Function DBDATE(ByVal strDate As String) As String
Public Function DBDATE(oDate) As String
   Dim strDate As String
   'Modified by Morgan 2025/3/17
   'If TypeName(oDate) = "Date" Then
   '   strDate = Format(oDate, "YYYY/MM/DD")
   'Else
   '   strDate = oDate
   'End If
   strDate = PUB_ChgDateToWDateStr(oDate)
   'end 2025/3/17
'end 2024/1/11
   Dim strTemp As String
   DBDATE = strDate
   ' 若日期中有/表示這是日期格式的字串如 YY/MM/DD 或 YYYY/MM/DD
   If InStr(1, strDate, "/") Then
      DBDATE = DBYEAR(strDate) & DBMONTH(strDate) & DBDAY(strDate)
   Else
      If Len(strDate) > 7 Then
         strTemp = strDate
         If CheckIsDate(strTemp, False) = True Then
            DBDATE = strDate
         End If
      Else
         strTemp = strDate
         If CheckIsTaiwanDate(strTemp, False) = True Then
            strTemp = strDate
            DBDATE = ChangeTStringToWString(strTemp)
         End If
      End If
   End If
End Function

'Added by Morgan 2025/3/17
Public Function PUB_ChgDateToWDateStr(oDate) As String
   If TypeName(oDate) = "Date" Then
      '不能直接用format轉格式，因當作業系統的顯示格式的日期分隔符號為"-"時，format就算指定用"/"也仍會回傳"-"
      PUB_ChgDateToWDateStr = Year(oDate) & "/" & Format(Month(oDate), "00") & "/" & Format(Day(oDate), "00")
   Else
      PUB_ChgDateToWDateStr = Replace(oDate, "-", "/")
   End If
End Function
