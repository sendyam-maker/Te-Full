Attribute VB_Name = "acc_fun"
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

'*************************************************
'  代理人名稱查詢
'
'*************************************************
Public Function FagentQuery(InputNo As String, InputSelect As Integer) As String
Dim adofagent As New ADODB.Recordset
   adofagent.CursorLocation = adUseClient
'   adofagent.Open "select * from fagent where fa01 = '" & Mid(InputNo, 1, 8) & "' and fa02 = '" & IIf(Mid(InputNo, 9, 1) = "", "0", Mid(InputNo, 9, 1)) & "'", adoTaie, adOpenStatic, adLockReadOnly
   adofagent.Open "select * from fagent where fa01 = '" & ChgSQL(Mid(InputNo, 1, 8)) & "' and fa02 = '" & ChgSQL(IIf(Mid(InputNo, 9, 1) = "", "0", Mid(InputNo, 9, 1))) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adofagent.RecordCount <> 0 Then
      adofagent.MoveFirst
      Select Case InputSelect
         Case 1
            If IsNull(adofagent.Fields("fa04").Value) Then
               FagentQuery = MsgText(601)
            Else
               FagentQuery = adofagent.Fields("fa04").Value
            End If
         Case 2
            If IsNull(adofagent.Fields("fa05").Value) Then
               FagentQuery = MsgText(601)
            Else
               FagentQuery = adofagent.Fields("fa05").Value
            End If
         Case 3
            If IsNull(adofagent.Fields("fa06").Value) Then
               FagentQuery = MsgText(601)
            Else
               FagentQuery = adofagent.Fields("fa06").Value
            End If
      End Select
   Else
      FagentQuery = MsgText(601)
   End If
   adofagent.Close
End Function

'Add By Cheng 2003/07/01
'*************************************************
'  代理人名稱查詢
'
'*************************************************
Public Function FagentQuery_1(InputNo As String, InputSelect As Integer) As String
Dim adofagent As New ADODB.Recordset
   adofagent.CursorLocation = adUseClient
   adofagent.Open "select * from fagent where fa01 = '" & Mid(InputNo, 1, 8) & "' and fa02 = '" & IIf(Mid(InputNo, 9, 1) = "", "0", Mid(InputNo, 9, 1)) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adofagent.RecordCount <> 0 Then
      adofagent.MoveFirst
      Select Case InputSelect
         Case 2
            If IsNull(adofagent.Fields("fa05").Value) Then
               FagentQuery_1 = MsgText(601)
            Else
               FagentQuery_1 = Trim(adofagent.Fields("fa05").Value & " " & adofagent.Fields("fa63").Value & " " & adofagent.Fields("fa64").Value & " " & adofagent.Fields("fa65").Value)
            End If
         Case 3
            If IsNull(adofagent.Fields("fa04").Value) Then
               FagentQuery_1 = MsgText(601)
            Else
               FagentQuery_1 = adofagent.Fields("fa04").Value
            End If
      End Select
   Else
      FagentQuery_1 = MsgText(601)
   End If
   adofagent.Close
End Function

'*************************************************
'  延遲系統作業時間
'
'*************************************************
'Remove by Morgan 2011/8/9 改用API
'Public Sub Sleep(InputPauseTime As Integer)
'Dim sglStart As Single
'
'      sglStart = Timer
'      Do While Timer < sglStart + InputPauseTime
'         DoEvents
'      Loop
'End Sub

'Modify By Sindy 2014/5/27 Mark統一使用basQuery中的函數
''*************************************************
''  電腦自動給號(傳票號碼)
''
''*************************************************
'Public Function AccAutoNo(InputItem As String, InputLength As Integer, Optional intYear As Integer = 0, Optional intMonth As Integer = 0) As String
'Dim adoaccnum As New ADODB.Recordset
'Dim strItem As String, strYes As String
'   If Len(InputItem) > 1 Then
'      strItem = Mid(InputItem, 2, 1)
'   Else
'      strItem = InputItem
'   End If
'   adoaccnum.CursorLocation = adUseClient
'   If intYear <> 0 Then
'      adoaccnum.Open "select * from acc1r0 where a1r01 = '" & InputItem & "' and a1r02 = " & (intYear + 1911) & " and a1r03 = " & intMonth & "", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccnum.RecordCount = 0 Then
'         AccAutoNo = strItem & IIf(intYear < 100, "0" & intYear, intYear) & IIf(intMonth < 10, "0" & intMonth, intMonth) & ZeroBeforeNo("0", InputLength)
'      Else
'         AccAutoNo = strItem & IIf(intYear < 100, "0" & intYear, intYear) & IIf(intMonth < 10, "0" & intMonth, intMonth) & ZeroBeforeNo(str(adoaccnum.Fields("a1r04").Value), InputLength)
'      End If
'   Else
'      adoaccnum.Open "select * from acc1r0 where a1r01 = '" & InputItem & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccnum.RecordCount = 0 Then
'         AccAutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & ZeroBeforeNo("0", InputLength)
'      Else
'         If adoaccnum.Fields("a1r02").Value <> Val(Mid(ServerDate, 1, 4)) Then
'            AccAutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & ZeroBeforeNo("0", InputLength)
'         Else
'            AccAutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & ZeroBeforeNo(str(adoaccnum.Fields("a1r04").Value), InputLength)
'         End If
'      End If
'   End If
'   adoaccnum.Close
'End Function
''*************************************************
''  電腦給號存檔(傳票號碼)
''
''*************************************************
'Public Function AccSaveAutoNo(InputItem As String, InputNo As String, Optional intYear As Integer = 0, Optional intMonth As Integer = 0) As String
'Dim adoaccnum As New ADODB.Recordset
'
'   adoaccnum.CursorLocation = adUseClient
'   If intYear <> 0 Then
'      adoaccnum.Open "select * from acc1r0 where a1r01 = '" & InputItem & "' and a1r02 = " & (intYear + 1911) & " and a1r03 = " & intMonth & "", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccnum.RecordCount = 0 Then
'         adoTaie.Execute "insert into acc1r0 (a1r01, a1r02, a1r03, a1r04) values ('" & InputItem & "', '" & Mid(ServerDate, 1, 4) & "', '" & Mid(ServerDate, 5, 2) & "', '" & InputNo & "')"
'      Else
'         adoTaie.Execute "UPDATE ACC1R0 SET A1R04 = " & InputNo & " WHERE A1R01 = '" & InputItem & "' and a1r02 = " & (intYear + 1911) & " and a1r03 = " & intMonth & ""
'      End If
'   Else
'      adoaccnum.Open "select * from acc1r0 where a1r01 = '" & InputItem & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccnum.RecordCount = 0 Then
'         adoTaie.Execute "insert into acc1r0 (a1r01, a1r02, a1r03, a1r04) values ('" & InputItem & "', '" & Mid(ServerDate, 1, 4) & "', '" & Mid(ServerDate, 5, 2) & "', '" & InputNo & "')"
'      Else
'         adoTaie.Execute "UPDATE ACC1R0 SET A1R01 = '" & InputItem & "', A1R02 = '" & Mid(ServerDate, 1, 4) & "', A1R03 = '" & Mid(ServerDate, 5, 2) & "', A1R04 = '" & InputNo & "' WHERE A1R01 = '" & InputItem & "'"
'      End If
'   End If
'   AccSaveAutoNo = MsgText(602)
'   adoaccnum.Close
'End Function

'*************************************************
'  計算天數
'
'*************************************************
Public Function CalculateDays(strStartDate As String, strEndDate As String) As Long
   CalculateDays = CDate(Mid(strEndDate, 1, 4) & "/" & Mid(strEndDate, 5, 2) & "/" & Mid(strEndDate, 7, 2)) - CDate(Mid(strStartDate, 1, 4) & "/" & Mid(strStartDate, 5, 2) & "/" & Mid(strStartDate, 7, 2))
End Function

'*************************************************
'  訊息顯示
'
'*************************************************
Public Sub MessageShow(strInputMsg As String)
   MsgBox MsgText(45) & strInputMsg & MsgText(46), , MsgText(5)
End Sub

''*************************************************
''  檢核資料是否存在
''
''*************************************************
'Public Function ExistCheck(strTable As String, strField As String, strValue As String, strError As String, Optional bolMsg As Boolean = True) As Boolean
'Dim adocheck As New ADODB.Recordset
'
'   strValue = Replace(strValue, "'", "''")
'   adocheck.CursorLocation = adUseClient
''   adocheck.Open "select " & strField & " from " & strTable & " where " & strField & " = " & "'" & strValue & "'", adoTaie, adOpenStatic, adLockReadOnly
'   adocheck.Open "select " & strField & " from " & strTable & " where " & strField & " = " & "'" & ChgSQL(strValue) & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adocheck.RecordCount = 0 Then
'      If bolMsg Then
'         MessageShow strError
'      End If
'      ExistCheck = False
'   Else
'      ExistCheck = True
'   End If
'   adocheck.Close
'End Function

'*************************************************
'  統一編號檢核
'
'*************************************************
Public Function UnionCode(strCode As String) As String
Dim intTotal As Integer

   intTotal = Val(Mid(Trim(str(Val(Mid(strCode, 2, 1)) * 2)), 1, 1)) + Val(Mid(Trim(str(Val(Mid(strCode, 2, 1)) * 2)), 2, 1)) + Val(Mid(strCode, 1, 1)) _
            + Val(Mid(Trim(str(Val(Mid(strCode, 4, 1)) * 2)), 1, 1)) + Val(Mid(Trim(str(Val(Mid(strCode, 4, 1)) * 2)), 2, 1)) + Val(Mid(strCode, 3, 1)) _
            + Val(Mid(Trim(str(Val(Mid(strCode, 6, 1)) * 2)), 1, 1)) + Val(Mid(Trim(str(Val(Mid(strCode, 6, 1)) * 2)), 2, 1)) + Val(Mid(strCode, 5, 1)) _
            + Val(Mid(Trim(str(Val(Mid(strCode, 7, 1)) * 4)), 1, 1)) + Val(Mid(Trim(str(Val(Mid(strCode, 7, 1)) * 4)), 2, 1)) + Val(Mid(strCode, 8, 1))
   If intTotal = 0 Then
      UnionCode = MsgText(603)
      Exit Function
   End If
   If (intTotal / 10) = Int(intTotal / 10) Then
      UnionCode = MsgText(602)
   Else
      If Mid(strCode, 7, 1) = "7" Then
         intTotal = Val(Mid(Trim(str(Val(Mid(strCode, 2, 1)) * 2)), 1, 1)) + Val(Mid(Trim(str(Val(Mid(strCode, 2, 1)) * 2)), 2, 1)) + Val(Mid(strCode, 1, 1)) _
                  + Val(Mid(Trim(str(Val(Mid(strCode, 4, 1)) * 2)), 1, 1)) + Val(Mid(Trim(str(Val(Mid(strCode, 4, 1)) * 2)), 2, 1)) + Val(Mid(strCode, 3, 1)) _
                  + Val(Mid(Trim(str(Val(Mid(strCode, 6, 1)) * 2)), 1, 1)) + Val(Mid(Trim(str(Val(Mid(strCode, 6, 1)) * 2)), 2, 1)) + Val(Mid(strCode, 5, 1)) _
                  + Val(Mid(Trim(Val(Mid(Trim(str(Val(Mid(strCode, 7, 1)) * 4)), 1, 1)) + Val(Mid(Trim(str(Val(Mid(strCode, 7, 1)) * 4)), 2, 1))), 1, 1)) + Val(Mid(strCode, 8, 1))
         If (intTotal / 10) = Int(intTotal / 10) Then
            UnionCode = MsgText(602)
         Else
            UnionCode = MsgText(603)
         End If
      Else
         UnionCode = MsgText(603)
      End If
   End If
End Function

'Modify By Sindy 2014/5/27 Mark統一使用basQuery中的函數
''*************************************************
''  刪除資料檢核
''
''*************************************************
'Public Function DeleteCheck(strSql As String) As String
'Dim adoDeleteCheck As New ADODB.Recordset
'
'   adoDeleteCheck.CursorLocation = adUseClient
'   adoDeleteCheck.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoDeleteCheck.RecordCount = 0 Then
'      MsgBox MsgText(28), , MsgText(5)
'      DeleteCheck = MsgText(603)
'   Else
'      DeleteCheck = MsgText(602)
'   End If
'   adoDeleteCheck.Close
'End Function

'*************************************************
'  檢核資料是否存在
'
'*************************************************
Public Function CheckData(strSql As String, strMsg As String) As Boolean
Dim adocheck As New ADODB.Recordset

   adocheck.CursorLocation = adUseClient
   adocheck.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount = 0 Then
      MsgBox MsgText(28) & strMsg, , MsgText(5)
      CheckData = False
   Else
      CheckData = True
   End If
   adocheck.Close
End Function

'*************************************************
'  檢核記錄是否異動
'
'*************************************************
Public Function CheckRecord(strSql As String, lngDate As Long, lngTime As Long) As Boolean
Dim adocheck As New ADODB.Recordset
Dim strConfirm As Variant

   adocheck.CursorLocation = adUseClient
   adocheck.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount <> 0 Then
      If lngDate = IIf(IsNull(adocheck.Fields(0).Value), 0, adocheck.Fields(0).Value) Then
         If lngTime = IIf(IsNull(adocheck.Fields(1).Value), 0, adocheck.Fields(1).Value) Then
            CheckRecord = True
            adocheck.Close
            Exit Function
         End If
      End If
   End If
   strConfirm = MsgBox(MsgText(66), vbYesNo, MsgText(5))
   If strConfirm = vbYes Then
      CheckRecord = True
   Else
      CheckRecord = False
   End If
   adocheck.Close
End Function

'*************************************************
'  取得序號
'
'*************************************************
Public Function GetSerialNo(strSql As String, intLength As Integer) As String
Dim adogetno As New ADODB.Recordset

   adogetno.CursorLocation = adUseServer
   adogetno.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
   If adogetno.RecordCount <> 0 Then
      If IsNull(adogetno.Fields(0).Value) Then
         GetSerialNo = ZeroBeforeNo("0", intLength)
      Else
         GetSerialNo = ZeroBeforeNo(adogetno.Fields(0).Value, intLength)
      End If
   Else
      GetSerialNo = ZeroBeforeNo("0", intLength)
   End If
   adogetno.Close
End Function

'*************************************************
'  取得上期科目結餘
'
'*************************************************
Public Function GetLastMonthBalance(intYear As Integer, intMonth As Integer, strCompany As String, strDepart As String, strAccNo As String) As Double
Dim adogetbalance As New ADODB.Recordset

   adogetbalance.CursorLocation = adUseClient
   adogetbalance.Open "select a0408 from acc040 where to_number(a0401||decode(length(a0402), 1, '0'||a0402, a0402)) in (select max(to_number(a0401||decode(length(a0402), 1, '0'||a0402, a0402))) from acc040 where to_number(a0401||decode(length(a0402), 1, '0'||a0402, a0402)) < " & Val(intYear & IIf(intMonth < 10, "0" & intMonth, intMonth)) & " and a0403 = '" & strCompany & "' and a0404 = '" & strDepart & "' and a0405 = '" & strAccNo & "') and a0403 = '" & strCompany & "' and a0404 = '" & strDepart & "' and a0405 = '" & strAccNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adogetbalance.RecordCount <> 0 Then
      If IsNull(adogetbalance.Fields(0).Value) Then
         GetLastMonthBalance = 0
      Else
         GetLastMonthBalance = Val(adogetbalance.Fields(0).Value)
      End If
   Else
      GetLastMonthBalance = 0
   End If
   adogetbalance.Close
End Function

'*************************************************
'  取得借/貸方別
'
'*************************************************
Public Function GetDebitCredit(strAccNo As String) As String
Dim adoget As New ADODB.Recordset

   adoget.CursorLocation = adUseClient
   adoget.Open "select * from acc010 where a0101 = '" & strAccNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoget.RecordCount <> 0 Then
      If IsNull(adoget.Fields("a0103").Value) Then
         GetDebitCredit = MsgText(601)
      Else
         GetDebitCredit = adoget.Fields("a0103").Value
      End If
   Else
      GetDebitCredit = MsgText(601)
   End If
   adoget.Close
End Function
'*************************************************
'  國家名稱查詢
'
'*************************************************
Public Function NationQuery(InputNo As String, InputSelect As Integer) As String
Dim adonation As New ADODB.Recordset
   adonation.CursorLocation = adUseClient
   adonation.Open "select * from nation where na01 = '" & InputNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adonation.RecordCount <> 0 Then
      adonation.MoveFirst
      Select Case InputSelect
         Case 1
            If IsNull(adonation.Fields("na03").Value) Then
               NationQuery = MsgText(601)
            Else
               NationQuery = adonation.Fields("na03").Value
            End If
         Case 2
            If IsNull(adonation.Fields("na04").Value) Then
               NationQuery = MsgText(601)
            Else
               NationQuery = adonation.Fields("na04").Value
            End If
      End Select
   Else
      NationQuery = MsgText(601)
   End If
   adonation.Close
End Function

'edit by nickc 2007/02/09 重複
'*************************************************
'  反白
'
'*************************************************
'Public Sub TextInverse(ByRef txtTemp As TextBox)
'   txtTemp.SelStart = 0
'   txtTemp.SelLength = Len(txtTemp.Text)
'End Sub

'*************************************************
'  計算各部門收入比率(管理部門費用)
'
'*************************************************
Public Function DeptPercentM(intYear As Integer, intMonth As Integer, strCompany As String, strDept As String) As Double
Dim adoacc040 As New ADODB.Recordset
Dim adoquery As New ADODB.Recordset

   If intMonth = 1 Then
      intYear = intYear - 1
      intMonth = 12
   Else
      intMonth = intMonth - 1
   End If
   adoacc040.CursorLocation = adUseClient
   adoacc040.Open "select sum(a0408) from acc040 where a0401 = " & intYear & " and a0402 = " & intMonth & " and a0403 = '" & strCompany & "' and a0404 not in ('TOT', 'M', 'SAL') and substr(a0405, 1, 1) = '4'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc040.RecordCount <> 0 Then
      If IsNull(adoacc040.Fields(0).Value) = False And adoacc040.Fields(0).Value <> 0 Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select sum(a0408) from acc040 where a0401 = " & intYear & " and a0402 = " & intMonth & " and a0403 = '" & strCompany & "' and a0404 = '" & strDept & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) = False Then
               DeptPercentM = Val(Format(Val(adoquery.Fields(0).Value) / Val(adoacc040.Fields(0).Value) * 100, FAmount))
               adoquery.Close
               adoacc040.Close
               Exit Function
            End If
         End If
         adoquery.Close
      End If
   End If
   DeptPercentM = 0
   adoacc040.Close
End Function

'*************************************************
'  計算各部門收入比率(智權部門費用)
'
'*************************************************
Public Function DeptPercentS(intYear As Integer, intMonth As Integer, strCompany As String, strDept As String) As Double
Dim adoacc040 As New ADODB.Recordset
Dim adoquery As New ADODB.Recordset

   If intMonth = 1 Then
      intYear = intYear - 1
      intMonth = 12
   Else
      intMonth = intMonth - 1
   End If
   adoacc040.CursorLocation = adUseClient
   adoacc040.Open "select sum(a0408) from acc040 where a0401 = " & intYear & " and a0402 = " & intMonth & " and a0403 = '" & strCompany & "' and a0404 not in ('TOT', 'M', 'SAL', 'FCL', 'FCP', 'FCT', 'FL') and substr(a0405, 1, 1) = '4'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc040.RecordCount <> 0 Then
      If IsNull(adoacc040.Fields(0).Value) = False And adoacc040.Fields(0).Value <> 0 Then
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select sum(a0408) from acc040 where a0401 = " & intYear & " and a0402 = " & intMonth & " and a0403 = '" & strCompany & "' and a0404 = '" & strDept & "'", adoTaie, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 Then
            If IsNull(adoquery.Fields(0).Value) = False Then
               DeptPercentS = Val(Format(Val(adoquery.Fields(0).Value) / Val(adoacc040.Fields(0).Value) * 100, FAmount))
               adoquery.Close
               adoacc040.Close
               Exit Function
            End If
         End If
         adoquery.Close
      End If
   End If
   DeptPercentS = 0
   adoacc040.Close
End Function

'*************************************************
'  檢核是否輸入部門別
'
'*************************************************
Public Function CheckDept(strAccNo As String, strDept As String) As Boolean
Dim adoacc010 As New ADODB.Recordset

   If Mid(strAccNo, 1, 1) = "4" Then
      If strDept = MsgText(55) Or strDept = "" Then
         CheckDept = False
         Exit Function
      End If
   End If
   If Mid(strAccNo, 1, 1) = "6" Then
      adoacc010.CursorLocation = adUseClient
      adoacc010.Open "select a0105 from acc010 where a0101 = '" & strAccNo & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc010.RecordCount <> 0 Then
         If IsNull(adoacc010.Fields(0).Value) Then
'            If strDept = MsgText(55) Or strDept = "" Then
'               CheckDept = False
'            Else
               CheckDept = True
'            End If
         Else
            If adoacc010.Fields(0).Value = "" Then
               If strDept = MsgText(55) Or strDept = "" Then
                  CheckDept = False
               Else
                  CheckDept = True
               End If
            Else
               CheckDept = True
            End If
         End If
      Else
         If strDept = MsgText(55) Or strDept = "" Then
            CheckDept = False
         Else
            CheckDept = True
         End If
      End If
      adoacc010.Close
   Else
      CheckDept = True
   End If
End Function
'Remove by Lydia 2016/04/14 已存在於basQuery
'*************************************************
'  將數字轉換成中文數字
'
'*************************************************
'Public Function ChangeNumber(strInputValue As String) As String
'Dim intCounter As Integer
'Dim bolZero As Boolean
'
'   bolZero = False
'   For intCounter = 1 To Len(strInputValue)
'      Select Case intCounter
'         Case 1
'            If Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1)) <> 0 Then
'               ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1)))
'            End If
'         Case 2
'            If Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1)) <> 0 Then
'               ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(10) & ChangeNumber
'            Else
'               If Len(strInputValue) > 2 Then
'                  bolZero = True
'               End If
'            End If
'         Case 3
'            If Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1)) <> 0 Then
'               If bolZero Then
'                  If Mid(strInputValue, Len(strInputValue) - 1, 2) = "00" Then
'                     ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(11) & ChangeNumber
'                  Else
'                     ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(11) & ShowNumberWord(0) & ChangeNumber
'                  End If
'                  bolZero = False
'               Else
'                  ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(11) & ChangeNumber
'               End If
'            Else
'               If Len(strInputValue) > 3 Then
'                  bolZero = True
'               End If
'            End If
'         Case 4
'            If Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1)) <> 0 Then
'               If bolZero Then
'                  If Mid(strInputValue, Len(strInputValue) - 2, 3) = "000" Then
'                     ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(12) & ChangeNumber
'                  Else
'                     ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(12) & ShowNumberWord(0) & ChangeNumber
'                  End If
'                  bolZero = False
'               Else
'                  ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(12) & ChangeNumber
'               End If
'            Else
'               If Len(strInputValue) > 4 Then
'                  bolZero = True
'               End If
'            End If
'         Case 5
'            If Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1)) <> 0 Then
'               If bolZero Then
'                  If Mid(strInputValue, Len(strInputValue) - 3, 4) = "0000" Then
'                     ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(13) & ChangeNumber
'                  Else
'                     ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(13) & ShowNumberWord(0) & ChangeNumber
'                  End If
'                  bolZero = False
'               Else
'                  ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(13) & ChangeNumber
'               End If
'            Else
'               If Len(strInputValue) > 5 Then
'                  If Len(strInputValue) > 8 Then
'                     If Mid(strInputValue, Len(strInputValue) - 6, 3) = "000" Then
'                        If Mid(strInputValue, Len(strInputValue) - 3, 4) = "0000" Then
'                           ChangeNumber = ChangeNumber
'                        Else
'                           ChangeNumber = ShowNumberWord(0) & ChangeNumber
'                        End If
'                     Else
'                        ChangeNumber = ShowNumberWord(13) & ShowNumberWord(0) & ChangeNumber
'                     End If
'                  Else
'                     If bolZero Then
'                        If Mid(strInputValue, Len(strInputValue) - 3, 4) = "0000" Then
'                           ChangeNumber = ShowNumberWord(13) & ChangeNumber
'                        Else
'                           ChangeNumber = ShowNumberWord(13) & ShowNumberWord(0) & ChangeNumber
'                        End If
'                        bolZero = False
'                     Else
'                        ChangeNumber = ShowNumberWord(13) & ChangeNumber
'                        bolZero = True
'                     End If
'                  End If
'               End If
'            End If
'         Case 6
'            If Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1)) <> 0 Then
'               ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(10) & ChangeNumber
'            Else
'               If Len(strInputValue) > 6 Then
'                  bolZero = True
'               End If
'            End If
'         Case 7
'            If Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1)) <> 0 Then
'               If bolZero Then
'                  If Mid(strInputValue, Len(strInputValue) - 5, 2) = "00" Then
'                     ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(11) & ChangeNumber
'                  Else
'                     ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(11) & ShowNumberWord(0) & ChangeNumber
'                  End If
'                  bolZero = False
'               Else
'                  ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(11) & ChangeNumber
'               End If
'            Else
'               If Len(strInputValue) > 7 Then
'                  bolZero = True
'               End If
'            End If
'         Case 8
'            If Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1)) <> 0 Then
'               If bolZero Then
'                  If Mid(strInputValue, Len(strInputValue) - 6, 3) = "000" Then
'                     ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(12) & ChangeNumber
'                  Else
'                     ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(12) & ShowNumberWord(0) & ChangeNumber
'                  End If
'                  bolZero = False
'               Else
'                  ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(12) & ChangeNumber
'               End If
'            End If
'         Case 9
'            If Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1)) <> 0 Then
'               ChangeNumber = ShowNumberWord(Val(Mid(strInputValue, Len(strInputValue) - intCounter + 1, 1))) & ShowNumberWord(14) & ChangeNumber
'            Else
'               ChangeNumber = ShowNumberWord(14) & ChangeNumber
'            End If
'      End Select
'   Next intCounter
'   ChangeNumber = ChangeNumber & ShowNumberWord(20)
'End Function

'*************************************************
'  案件性質名稱查詢
'
'*************************************************
Public Function PropertyQuery(strNo As String, strType As String) As String
Dim adoquery As New ADODB.Recordset

   adoquery.CursorLocation = adUseClient
   adoquery.Open "select nvl(cpm03, nvl(cpm04, cpm10)) from casepropertymap where cpm01 = '" & strNo & "' and cpm02 = '" & strType & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If IsNull(adoquery.Fields(0).Value) Then
         PropertyQuery = MsgText(601)
      Else
         PropertyQuery = adoquery.Fields(0).Value
      End If
   Else
      PropertyQuery = MsgText(601)
   End If
   adoquery.Close
End Function

'*************************************************
'  員工部門查詢
'
'*************************************************
Public Function StaffDeptQuery(InputNo As String) As String
Dim adostaff As New ADODB.Recordset
   adostaff.CursorLocation = adUseClient
   '2011/4/12 modify by sonia 改抓st15
   'adostaff.Open "select st03 from staff where st01 = '" & InputNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   adostaff.Open "select st15 from staff where st01 = '" & InputNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adostaff.RecordCount <> 0 Then
      If IsNull(adostaff.Fields("st15").Value) Then
         StaffDeptQuery = MsgText(601)
      Else
         StaffDeptQuery = adostaff.Fields("st15").Value
      End If
   Else
      StaffDeptQuery = MsgText(601)
   End If
   adostaff.Close
End Function

'*************************************************
'  智權人員所屬所別查詢
'
'*************************************************
Public Function StaffArea(InputNo As String) As String
Dim adostaff As New ADODB.Recordset
   adostaff.CursorLocation = adUseClient
   adostaff.Open "select st06 from staff where st01 = '" & InputNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adostaff.RecordCount <> 0 Then
      If IsNull(adostaff.Fields("st06").Value) Then
         StaffArea = MsgText(601)
      Else
         StaffArea = adostaff.Fields("st06").Value
      End If
   Else
      StaffArea = MsgText(601)
   End If
   adostaff.Close
End Function

'*************************************************
'  將數字轉換成英文數字
'
'*************************************************
Public Function ChangeNumberE(strInputValue As String) As String
Dim intCounter As Integer
Dim strNumber As String
Dim strWord As String
   
   strNumber = Format(strInputValue, "0.00")
   For intCounter = 1 To Len(strNumber)
      Select Case intCounter
         Case 1
            If Val(Mid(strNumber, Len(strNumber) - 1, 2)) <> 0 Then
               If Val(Mid(strNumber, Len(strNumber) - 1, 1)) <> 1 Then
                  If Val(Mid(strNumber, Len(strNumber), 1)) <> 0 Then
                     strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber), 1)) & strWord
                  End If
               End If
            End If
            strWord = strWord & " " & ShowNumber(107)
         Case 2
            If Val(Mid(strNumber, Len(strNumber) - 1, 1)) = 1 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 1, 2)) & strWord
            Else
               If Val(Mid(strNumber, Len(strNumber) - 1, 1)) > 1 Then
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 1, 1) & "0") & strWord
               End If
            End If
            If strWord <> (" " & ShowNumber(107)) Then
               strWord = " " & ShowNumber(105) & " " & ShowNumber(99) & " " & strWord
            End If
         Case 4
            If Len(strNumber) = 4 Then
               If Val(Mid(strNumber, Len(strNumber) - 3, 1)) <> 0 Then
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 3, 1)) & " " & strWord
               End If
            Else
               If Val(Mid(strNumber, Len(strNumber) - 3, 1)) >= 1 Then
                  If Val(Mid(strNumber, Len(strNumber) - 4, 1)) <> 1 Then
                     strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 3, 1)) & " " & strWord
                  End If
               End If
            End If
         Case 5
            If Val(Mid(strNumber, Len(strNumber) - 4, 1)) = 1 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 4, 2)) & strWord
            Else
               If Val(Mid(strNumber, Len(strNumber) - 4, 1)) > 1 Then
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 4, 1) & "0") & strWord
               End If
            End If
         Case 6
            If Val(Mid(strNumber, Len(strNumber) - 5, 1)) <> 0 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 5, 1)) & " " & ShowNumber(100) & strWord
            End If
         Case 7
            If Len(strNumber) = 7 Then
               If Val(Mid(strNumber, Len(strNumber) - 6, 1)) <> 0 Then
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 6, 1)) & " " & ShowNumber(101) & strWord
               Else
                  strWord = " " & ShowNumber(101) & strWord
               End If
            Else
               If Val(Mid(strNumber, Len(strNumber) - 7, 1)) <= 1 Then
                  If Len(strNumber) > 8 Then
                     If Val(Mid(strNumber, Len(strNumber) - 8, 3)) <> 0 Then
                        strWord = " " & ShowNumber(101) & strWord
                     End If
                  Else
                     strWord = " " & ShowNumber(101) & strWord
                  End If
               Else
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 6, 1)) & " " & ShowNumber(101) & strWord
               End If
            End If
         Case 8
            If Val(Mid(strNumber, Len(strNumber) - 7, 1)) = 1 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 7, 2)) & strWord
            Else
               If Val(Mid(strNumber, Len(strNumber) - 7, 1)) > 1 Then
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 7, 1) & "0") & strWord
               End If
            End If
         Case 9
            If Val(Mid(strNumber, Len(strNumber) - 8, 1)) <> 0 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 8, 1)) & " " & ShowNumber(100) & strWord
            End If
         Case 10
            If Len(strNumber) = 10 Then
               If Val(Mid(strNumber, Len(strNumber) - 9, 1)) <> 0 Then
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 9, 1)) & " " & ShowNumber(102) & strWord
               Else
                  strWord = " " & ShowNumber(102) & strWord
               End If
            Else
               If Val(Mid(strNumber, Len(strNumber) - 10, 1)) = 1 Then
                  strWord = " " & ShowNumber(102) & strWord
               Else
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 9, 1)) & " " & ShowNumber(102) & strWord
               End If
            End If
         Case 11
            If Val(Mid(strNumber, Len(strNumber) - 10, 1)) = 1 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 10, 2)) & strWord
            Else
               If Val(Mid(strNumber, Len(strNumber) - 10, 1)) > 1 Then
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 10, 1) & "0") & strWord
               End If
            End If
         Case 12
            If Val(Mid(strNumber, Len(strNumber) - 11, 1)) <> 0 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 11, 1)) & " " & ShowNumber(100) & strWord
            End If
         Case 13
            If Val(Mid(strNumber, Len(strNumber) - 12, 1)) <> 0 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 12, 1)) & " " & ShowNumber(101) & strWord
            End If
         Case 14
            If Len(strNumber) = 14 Then
               If Val(Mid(strNumber, Len(strNumber) - 13, 1)) <> 0 Then
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 13, 1)) & " " & ShowNumber(103) & strWord
               Else
                  strWord = " " & ShowNumber(103) & strWord
               End If
            Else
               If Val(Mid(strNumber, Len(strNumber) - 14, 1)) = 1 Then
                  strWord = " " & ShowNumber(103) & strWord
               Else
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 13, 1)) & " " & ShowNumber(103) & strWord
               End If
            End If
         Case 15
            If Val(Mid(strNumber, Len(strNumber) - 14, 1)) = 1 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 14, 2)) & strWord
            Else
               If Val(Mid(strNumber, Len(strNumber) - 14, 1)) > 1 Then
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 14, 1) & "0") & strWord
               End If
            End If
         Case 16
            If Val(Mid(strNumber, Len(strNumber) - 15, 1)) <> 0 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 15, 1)) & " " & ShowNumber(100) & strWord
            End If
         Case 17
            If Val(Mid(strNumber, Len(strNumber) - 16, 1)) <> 0 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 16, 1)) & " " & ShowNumber(101) & strWord
            End If
         Case 18
            If Len(strNumber) = 18 Then
               If Val(Mid(strNumber, Len(strNumber) - 17, 1)) <> 0 Then
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 17, 1)) & " " & ShowNumber(104) & strWord
               Else
                  strWord = " " & ShowNumber(104) & strWord
               End If
            Else
               If Val(Mid(strNumber, Len(strNumber) - 18, 1)) = 1 Then
                  strWord = " " & ShowNumber(104) & strWord
               Else
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 17, 1)) & " " & ShowNumber(104) & strWord
               End If
            End If
         Case 19
            If Val(Mid(strNumber, Len(strNumber) - 18, 1)) = 1 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 18, 2)) & strWord
            Else
               If Val(Mid(strNumber, Len(strNumber) - 18, 1)) > 1 Then
                  strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 18, 1) & "0") & strWord
               End If
            End If
         Case 20
            If Val(Mid(strNumber, Len(strNumber) - 19, 1)) <> 0 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 19, 1)) & " " & ShowNumber(100) & strWord
            End If
         Case 21
            If Val(Mid(strNumber, Len(strNumber) - 20, 1)) <> 0 Then
               strWord = " " & ShowNumber(Mid(strNumber, Len(strNumber) - 20, 1)) & " " & ShowNumber(101) & strWord
            End If
      End Select
   Next intCounter
   If Mid(strWord, 1, 4) = (" " & ShowNumber(105)) Then
      strWord = Mid(strWord, 5, Len(strWord) - 4)
   Else
      strWord = ShowNumber(108) & " " & strWord
   End If
   ChangeNumberE = strWord
End Function

'*************************************************
'  本所案號後補零
'
'*************************************************
Public Function CaseNoZero(strInputValue As String) As String
   If IsNumeric(Mid(strInputValue, 3, 1)) = False Then
      Select Case Len(strInputValue)
         Case 9
            CaseNoZero = strInputValue & "000"
         Case 10
            CaseNoZero = strInputValue & "00"
         Case Else
            CaseNoZero = strInputValue
      End Select
   Else
      If IsNumeric(Mid(strInputValue, 2, 1)) = False Then
         Select Case Len(strInputValue)
            Case 8
               CaseNoZero = strInputValue & "000"
            Case 9
               CaseNoZero = strInputValue & "00"
            Case Else
               CaseNoZero = strInputValue
         End Select
      Else
         Select Case Len(strInputValue)
            Case 7
               CaseNoZero = strInputValue & "000"
            Case 8
               CaseNoZero = strInputValue & "00"
            Case Else
               CaseNoZero = strInputValue
         End Select
      End If
   End If
End Function

'*************************************************
'  取得上期實際收入合計
'
'*************************************************
Public Function GetLastMonthIncome(intYear As Integer, intMonth As Integer, strCompany As String, strDepart As String) As Double
Dim adogetIncome As New ADODB.Recordset

   adogetIncome.CursorLocation = adUseClient
   adogetIncome.Open "select sum(a0408) from acc040 where to_number(a0401||decode(length(a0402), 1, '0'||a0402, a0402)) in (select max(to_number(a0401||decode(length(a0402), 1, '0'||a0402, a0402))) from acc040 where to_number(a0401||decode(length(a0402), 1, '0'||a0402, a0402)) < " & Val(intYear & IIf(intMonth < 10, "0" & intMonth, intMonth)) & " and a0403 = '" & strCompany & "' and a0404 = '" & strDepart & "' and substr(a0405, 1, 1) = '4') and a0403 = '" & strCompany & "' and a0404 = '" & strDepart & "' and substr(a0405, 1, 1) = '4'", adoTaie, adOpenStatic, adLockReadOnly
   If adogetIncome.RecordCount <> 0 Then
      If IsNull(adogetIncome.Fields(0).Value) Then
         GetLastMonthIncome = 0
      Else
         GetLastMonthIncome = Val(adogetIncome.Fields(0).Value)
      End If
   Else
      GetLastMonthIncome = 0
   End If
   adogetIncome.Close
End Function

'*************************************************
'  取得當期實際收入合計
'
'*************************************************
Public Function GetMonthIncome(intYear As Integer, intMonth As Integer, strCompany As String, strDepart As String) As Double
Dim adogetIncome As New ADODB.Recordset

   adogetIncome.CursorLocation = adUseClient
   adogetIncome.Open "select sum(a0408) from acc040 where a0401 = " & intYear & " and a0402 = " & intMonth & " and a0403 = '" & strCompany & "' and a0404 = '" & strDepart & "' and substr(a0405, 1, 1) = '4'", adoTaie, adOpenStatic, adLockReadOnly
   If adogetIncome.RecordCount <> 0 Then
      If IsNull(adogetIncome.Fields(0).Value) Then
         GetMonthIncome = 0
      Else
         GetMonthIncome = Val(adogetIncome.Fields(0).Value)
      End If
   Else
      GetMonthIncome = 0
   End If
   adogetIncome.Close
End Function

'*************************************************
'  案件性質名稱查詢
'
'*************************************************
Public Function CasePropertyQuery(strSys As String, strProperty As String) As String
Dim adoquery As New ADODB.Recordset

   adoquery.CursorLocation = adUseClient
   adoquery.Open "select * from casepropertymap where cpm01 = '" & strSys & "' and cpm02 = '" & strProperty & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      CasePropertyQuery = ""
   Else
      CasePropertyQuery = ""
   End If
   adoquery.Close
End Function

'*************************************************
'  自動給過去年度之流水號
'
'*************************************************
Public Function UpdateNo(strTable As String, strField As String, intLength As Integer, strDate As String, Optional strDocNo As String = "") As String
Dim adoaccnum As New ADODB.Recordset

On Error GoTo Checking
   adoaccnum.CursorLocation = adUseClient
   adoaccnum.Open "select nvl(max(" & strField & "), 0) from " & strTable & " where substr(" & strField & ", 2, 3) = '" & Mid(strDate, 1, 3) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccnum.RecordCount <> 0 Then
      If IsNull(adoaccnum.Fields(0).Value) = False Then
         If adoaccnum.Fields(0).Value = "0" Then
            UpdateNo = strDocNo & Mid(strDate, 1, 3) & ZeroBeforeNo(Mid(adoaccnum.Fields(0).Value, 5, intLength), intLength)
         Else
            UpdateNo = Mid(adoaccnum.Fields(0).Value, 1, 4) & ZeroBeforeNo(Mid(adoaccnum.Fields(0).Value, 5, intLength), intLength)
         End If
'         strYes = SaveAutoNo(Mid(adoaccnum.Fields(0).Value, 1, 1), Mid(UpdateNo, 5, intLength))
      End If
   End If
   adoaccnum.Close
Checking:
   If Err.Number = 0 Then
      Exit Function
   End If
   MsgBox Err.Description, , MsgText(5)
End Function
'edit by nickc 2007/02/09 重複
'轉換字串以塞入SQL語法
'Public Function CNULL(ByRef strNULL As String) As String
'If strNULL = "" Then
'   CNULL = "NULL"
'Else
'   CNULL = "'" + strNULL + "'"
'End If
'End Function

'*************************************************
'  顯示加天數後之日期
'
'*************************************************
Public Function ShowDate(strValue As String, LngDays As Long) As String
Dim datDate As Date

On Error GoTo Checking
   datDate = AFDate(Val(strValue) + 19110000)
   ShowDate = CFDate(ACDate(Format(datDate + LngDays, "YYYYMMDD")))
Checking:
   If Err.Number = 0 Then
      Exit Function
   End If
   MsgBox Err.Description, , MsgText(5)
End Function

'*************************************************
'  依科目帶出智權人員編號
'
'*************************************************
'modify by sonia 2021/1/18 +strCaseNo
Public Function AccNoToSalesNo(strAccNo As String, strCaseNo As String) As String
'add by sonia 2021/1/18 以本所案號判別FCP,FCT英日文組
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim stCP13 As String
Dim strCaseNo1 As String, strCaseNo2 As String, StrCaseNo3 As String, strCaseNo4 As String

   stCP13 = ""
   If strCaseNo <> "" Then
      strCaseNo1 = Mid(strCaseNo, 1, Len(strCaseNo) - 9)
      strCaseNo2 = Mid(strCaseNo, Len(strCaseNo) - 8, 6)
      StrCaseNo3 = Mid(strCaseNo, Len(strCaseNo) - 2, 1)
      strCaseNo4 = Mid(strCaseNo, Len(strCaseNo) - 1, 2)
      stCP13 = PUB_GetAKindSalesNo(strCaseNo1, strCaseNo2, StrCaseNo3, strCaseNo4)
   End If
'end 2021/1/18
   
   Select Case Left(strAccNo, 4)
      Case "4171"
         'modify by sonia 2021/1/18 依本所案號判別組別
         'AccNoToSalesNo = "F4102"
         If stCP13 = "" Then
            AccNoToSalesNo = "F4102"
         Else
            StrSQLa = "Select ST16 From staff Where st01='" & stCP13 & "'"
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               If "" & rsA.Fields(0).Value = "2" Then
                  AccNoToSalesNo = "F4105"
               Else
                  AccNoToSalesNo = "F4104"
               End If
            End If
         End If
         'end 2021/1/18
      Case "4172"
         'modify by sonia 2021/1/18 依本所案號判別組別
         'AccNoToSalesNo = "F4103"
         If stCP13 = "" Then
            AccNoToSalesNo = "F4103"
         Else
            StrSQLa = "Select ST16 From staff Where st01='" & stCP13 & "'"
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               If "" & rsA.Fields(0).Value = "4" Then
                  AccNoToSalesNo = "F4107"
               Else
                  AccNoToSalesNo = "F4106"
               End If
            End If
         End If
         'end 2021/1/18
      Case "4161"
         AccNoToSalesNo = "F4101"
      Case Else
         AccNoToSalesNo = ""
   End Select
End Function

'cancel by sonia 2021/1/18 Casher系統未使用
''*************************************************
''  依智權人員帶出做帳智權人員-->國外收款
''  92.10.9 add by sonia
''*************************************************
'Public Function SalesNoToAccSales(strSalesNo As String, strAccNo As String) As String
'Dim adoquery As New ADODB.Recordset
'
'   adoquery.CursorLocation = adUseClient
'   adoquery.Open "select ST15 from staff where st01 = '" & strSalesNo & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoquery.RecordCount <> 0 Then
'      If adoquery.Fields(0).Value >= "F" And adoquery.Fields(0).Value <= "G" Then
'         Select Case strAccNo
'            Case "4171"
'               SalesNoToAccSales = "F4102"
'            Case "417201", "417202"
'               SalesNoToAccSales = "F4103"
'            Case "416101", "416102"
'               SalesNoToAccSales = "F4101"
'            Case Else
'               If adoquery.Fields(0).Value >= "F1" And adoquery.Fields(0).Value <= "F19" Then
'                  SalesNoToAccSales = "F4103"
'               ElseIf adoquery.Fields(0).Value >= "F2" And adoquery.Fields(0).Value <= "F29" Then
'                  SalesNoToAccSales = "F4102"
'               ElseIf adoquery.Fields(0).Value >= "F3" And adoquery.Fields(0).Value <= "F49" Then
'                  SalesNoToAccSales = "F4101"
'               Else
'                  SalesNoToAccSales = ""
'               End If
'         End Select
'      ElseIf adoquery.Fields(0).Value < "S" Then
'         SalesNoToAccSales = "M0100"
'      ElseIf adoquery.Fields(0).Value >= "T" Then
'         SalesNoToAccSales = "M0100"
'      Else
'         SalesNoToAccSales = strSalesNo
'      End If
'   Else
'      SalesNoToAccSales = ""
'   End If
'   adoquery.Close
'
'End Function
'end 2021/1/18

'cancel by sonia 2021/1/18 Casher系統未使用
''*************************************************
''  依智權人員帶出做帳智權人員-->國內收款
''  92.12.22 add by sonia
''*************************************************
'Public Function SalesNoToSales(strSalesNo As String, strAccNo As String) As String
'Dim adoquery As New ADODB.Recordset
'
'   adoquery.CursorLocation = adUseClient
'   adoquery.Open "select ST15 from staff where st01 = '" & strSalesNo & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoquery.RecordCount <> 0 Then
'      If adoquery.Fields(0).Value >= "F" And adoquery.Fields(0).Value <= "G" Then
'         Select Case strAccNo
'            Case "4171"
'               SalesNoToSales = "F4102"
'            Case "417201", "417202"
'               SalesNoToSales = "F4103"
'            Case "416101", "416102"
'               SalesNoToSales = "F4101"
'            Case Else
'               If adoquery.Fields(0).Value >= "F1" And adoquery.Fields(0).Value <= "F19" Then
'                  SalesNoToSales = "F4103"
'               ElseIf adoquery.Fields(0).Value >= "F2" And adoquery.Fields(0).Value <= "F29" Then
'                  SalesNoToSales = "F4102"
'               ElseIf adoquery.Fields(0).Value >= "F3" And adoquery.Fields(0).Value <= "F49" Then
'                  SalesNoToSales = "F4101"
'               Else
'                  SalesNoToSales = ""
'               End If
'         End Select
'      Else
'         SalesNoToSales = strSalesNo
'      End If
'   Else
'      SalesNoToSales = ""
'   End If
'   adoquery.Close
'
'End Function
'end 2021/1/18
'
'Removed by Morgan 2014/3/10 整合同名函數到 basQuery
''*************************************************
''  電腦自動給號
''
''*************************************************
'Public Function AutoNo(InputItem As String, InputLength As Integer, Optional intTrans As Integer = 0) As String
'Dim adoaccnum As New ADODB.Recordset
'Dim strItem As String, strYes As String
'
'   If intTrans = 0 Then
'      adoTaie.BeginTrans
'   End If
'   adoTaie.Execute "update autonumber set au03 = au03 where au01 = '" & InputItem & "'"
'   If Len(InputItem) > 1 Then
'      strItem = Mid(InputItem, 2, 1)
'   Else
'      strItem = InputItem
'   End If
'   adoaccnum.CursorLocation = adUseClient
'   adoaccnum.Open "select * from autonumber where au01 = '" & InputItem & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccnum.RecordCount = 0 Then
'      If InputItem = "E" Then
'         AutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & ZeroBeforeNo("2000", InputLength)
'      Else
'         AutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & ZeroBeforeNo("0", InputLength)
'      End If
'   Else
'      If adoaccnum.Fields("au02").Value <> Val(Mid(ServerDate, 1, 4)) Then
'         If InputItem = "E" Then
'            AutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & ZeroBeforeNo("2000", InputLength)
'         Else
'            AutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & ZeroBeforeNo("0", InputLength)
'         End If
'      Else
'         AutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & ZeroBeforeNo(str(adoaccnum.Fields("au03").Value), InputLength)
'      End If
'   End If
'   If Len(InputItem) = 1 Then
'      strYes = SaveAutoNo(InputItem, Mid(AutoNo, 5, InputLength))
'   End If
'   adoaccnum.Close
'   If intTrans = 0 Then
'      adoTaie.CommitTrans
'   End If
'End Function
'
''*************************************************
''  電腦給號存檔
''
''*************************************************
'Public Function SaveAutoNo(InputItem As String, InputNo As String) As String
'Dim adoaccnum As New ADODB.Recordset
''   adoTaie.Execute "UPDATE AUTONUMBER SET AU01 = '" & InputItem & "', AU02 = '" & Mid(ServerDate, 1, 4) & "', AU03 = '" & InputNo & "' WHERE AU01 = '" & InputItem & "'"
'   adoaccnum.CursorLocation = adUseClient
'   adoaccnum.Open "select * from autonumber where au01 = '" & InputItem & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   If adoaccnum.RecordCount = 0 Then
'      adoaccnum.AddNew
'      adoaccnum.Fields("au01").Value = InputItem
'   End If
'   adoaccnum.Fields("au02").Value = Mid(ServerDate, 1, 4)
'   adoaccnum.Fields("au03").Value = InputNo
'   adoaccnum.UpdateBatch
'   adoaccnum.Close
'   SaveAutoNo = MsgText(602)
'End Function

'*************************************************
'  公司名稱查詢
'
'*************************************************
'Remove by Lydia 2020/03/26 以basUpdate為準
'Public Function A0802Query(InputNo As String) As String
'Dim adoacc080 As New ADODB.Recordset
'   adoacc080.CursorLocation = adUseClient
'   adoacc080.Open "select * from acc080 where a0801 = '" & InputNo & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoacc080.RecordCount <> 0 Then
'      If IsNull(adoacc080.Fields("a0802").Value) Then
'         A0802Query = MsgText(601)
'      Else
'         A0802Query = adoacc080.Fields("a0802").Value
'      End If
'   Else
'      A0802Query = MsgText(601)
'   End If
'   adoacc080.Close
'End Function

'Copy from aacc_fun by Morgan 2013/5/8
Public Function PUB_GetStaffState(p_ST01 As String, p_ST02 As String, Optional p_bolMsg As Boolean) As Integer
Dim stSQL As String, ii As Integer
   
   stSQL = "select st02,st04 from staff where st01='" & p_ST01 & "'"
   ii = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(ii, stSQL)
   Set RsTemp = ClsLawReadRstMsg(ii, stSQL)
   If ii = 1 Then
      p_ST02 = "" & RsTemp.Fields("st02")
      PUB_GetStaffState = Val("" & RsTemp.Fields("st04"))
      If p_bolMsg = True Then
         If PUB_GetStaffState = 2 Then
            MsgBox "員工已離職！", vbExclamation
         End If
      End If
   Else
      PUB_GetStaffState = 0
      If p_bolMsg = True Then
         MsgBox "員工不存在！", vbCritical
      End If
   End If
End Function

'Add By Sindy 2014/10/20
Public Sub PUB_AccSettingMail(strTemplatePath As String, strFilePathName As String, _
                              strSubject As String, strContent As String, strTo As String)
Dim adoRst As ADODB.Recordset
Dim objOutLook As Object
Dim objMail As Object
Dim strEmp As String, strEMP_Tel As String
Dim strOldContent As String
Dim ArrStr() As String, ii As Integer
   
   strExc(0) = "select st02,ed01" & _
               " from staff,ExtensionData" & _
               " where ST01=ED02(+)" & _
               " and st01='" & Pub_GetSpecMan("財務處總帳人員") & "'"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strEmp = adoRst.Fields("st02")
      strEMP_Tel = "" & adoRst.Fields("ed01")
   End If
   
   '呼叫新郵件：
   Set objOutLook = CreateObject("Outlook.Application")
   If Dir(strTemplatePath) <> "" Then
      Set objMail = objOutLook.CreateItemFromTemplate(strTemplatePath)
   Else
      Set objMail = objOutLook.CreateItem(0)
   End If
   strOldContent = objMail.HTMLBody '郵件原內文
   strOldContent = Replace(Replace(strOldContent, "&lt;EMP&gt;", strEmp), "EMPTEL", strEMP_Tel)
   '副本.cc
   '收件者.To
   objMail.To = strTo
   '主旨.Subject
   objMail.Subject = strSubject
   '加附件
   If strFilePathName <> "" Then
      ArrStr = Split(strFilePathName, ";")
      For ii = 0 To UBound(ArrStr)
         objMail.Attachments.add ArrStr(ii)
      Next ii
   End If
   '內文.Body
   '轉HTML格式
   strContent = Replace(strContent, vbCrLf, "<BR>")
   strContent = Replace(strContent, "  ", "&nbsp;&nbsp;")
   objMail.HTMLBody = "<FONT FACE=""Times New Roman"">" & strContent & "<BR>" & strOldContent & "</FONT>"
   objMail.Display
   
   Set objMail = Nothing
   Set objOutLook = Nothing
   Set adoRst = Nothing
End Sub

'Added by Morgan 2014/1/17
'Modified by Morgan 2015/5/7 從銷退移來共用
'更新收據結清與介紹獎金發放日期
Public Sub PUB_UpdateReceiptStatus(pNo As String)
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select nvl(Amt2,0) Amt2,nvl(Amt3,0) Amt3,nvl(a0k06,0)+nvl(a0k07,0) Amt1,a0k34,a0k36,a0k37 from acc0k0,(select a1u02, sum(nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0)) Amt2,sum(nvl(a1u07,0)+nvl(a1u09,0)) Amt3 from acc1u0 where a1u02='" & pNo & "' group by a1u02) x where a0k01='" & pNo & "' and a0k01=a1u02(+)"
   Set rsQuery = Nothing
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      '全銷(應收=已銷)
      If rsQuery.Fields("Amt1") = rsQuery.Fields("Amt3") Then
         stSQL = "update acc0k0 set a0k36=null, a0k37='N' where a0k01='" & pNo & "'"
         adoTaie.Execute stSQL, intR
      '全收(應收-已銷=已收)
      ElseIf rsQuery.Fields("Amt1") - rsQuery.Fields("Amt3") = rsQuery.Fields("Amt2") Then
         stSQL = "update acc0k0 set a0k36=decode(a0k34,null,null,nvl(a0k36," & strSrvDate(2) & ")), a0k37='Y' where a0k01='" & pNo & "'"
         adoTaie.Execute stSQL, intR
      Else
         stSQL = "update acc0k0 set  a0k36=NULL, a0k37=NULL where a0k01='" & pNo & "'"
         adoTaie.Execute stSQL, intR
      End If
   End If
   Set rsQuery = Nothing
End Sub

'Modify By Sindy 2015/8/6 從Frmacc1151搬至此處,變共用func
'Modify By Sindy 2017/2/14 + Optional ByVal bolShowMsg As Boolean = True
Public Function PUB_ChkIsPerson(pNo As String, Optional ByVal bolShowMsg As Boolean = True) As Boolean
Dim m_CU158 As String 'Add By Sindy 2015/12/10 檢查是否為境外公司
   
   PUB_ChkIsPerson = False
   strExc(0) = "select a0k05,a0k11 from acc0k0 where a0k01='" & pNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      If RsTemp("a0k05") = "1" Then
         If bolShowMsg = True Then
            MsgBox "本收據屬個人不能扣繳!!", vbExclamation, "扣繳檢查"
         End If
         PUB_ChkIsPerson = True
      'Added by Morgan 2013/12/26
      ElseIf RsTemp("a0k11") = "J" Then
         If bolShowMsg = True Then
            MsgBox "智權公司不能扣繳!!", vbExclamation, "扣繳檢查"
         End If
         PUB_ChkIsPerson = True
      'end 2013/12/26
      End If
   End If
   'Add By Sindy 2015/12/10 檢查是否為境外公司
   Call GetTitleCustData("", "", pNo, , , , , , , , , , , , m_CU158)
   If m_CU158 = "Y" Then
      If bolShowMsg = True Then
         MsgBox "境外公司不能扣繳!!", vbExclamation, "扣繳檢查"
      End If
      PUB_ChkIsPerson = True
   End If
   '2015/12/10 END
End Function

'Add by Morgan 2005/12/23
'** 檢查傳票是否已過帳
Public Function PUB_CheckPosted(p_a1p22 As String, Optional p_Msg As Boolean = True, Optional p_a1p01 As String = "1") As Boolean

On Error GoTo ErrHnd
   
   strSql = "select ax210 from acc021" & _
      " where ax201 = " & CNULL(p_a1p01) & " and ax202 = " & CNULL(p_a1p22) & " and ax210 is not null"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         If p_Msg Then
            MsgBox MsgText(155), , MsgText(5)
         End If
         PUB_CheckPosted = True
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Function

'Add by Amy 2014/09/26 輸入日期與系統日期比較 適用於入帳日、傳票日
'strCompyNo 公司別(A0201) /WorkDay(需判斷日期)/strMsg(回傳訊息)
Public Function ChkWorkData(strCompyNo As String, WorkDay As String, strMsg As String) As Boolean
    Dim strQuery As String
    Dim intQ As Integer, intChoose As Integer
    Dim RsQ As New ADODB.Recordset
    
    ChkWorkData = False
    strMsg = ""
    If Mid(WorkDay, 1, 6) = Mid(strSrvDate(1), 1, 6) Then
        '作業月=系統月,取當月最大傳票日
        intChoose = 1
        strQuery = "Select Max(A0205)+19110000 From Acc020 Where A0201='" & strCompyNo & "' And A0205 Between " & Mid(WorkDay, 1, 6) - 191100 & "00" & " And " & Mid(WorkDay, 1, 6) - 191100 & "31"
    ElseIf Mid(WorkDay, 1, 6) < Mid(strSrvDate(1), 1, 6) Then
        '作業月<系統月,取該月最大工作日
        intChoose = 2
        strQuery = "Select Max(WD01) From WorkDay Where WD01 Between " & Mid(WorkDay, 1, 6) & "00" & " And " & Mid(WorkDay, 1, 6) & "31"
    Else
        '作業月>系統月,取該月第一個工作日
        intChoose = 3
        strQuery = "Select Min(WD01) From WorkDay Where WD01 Between " & Mid(WorkDay, 1, 6) & "00" & " And " & Mid(WorkDay, 1, 6) & "31"
    End If
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQuery)
    If intQ = 1 Then
        Select Case intChoose
            Case 1
                If RsQ.Fields(0) > WorkDay Then
                    strMsg = "不可小於當月最大傳票日!!"
                    Exit Function
                End If
            Case 2
               If RsQ.Fields(0) <> WorkDay Then
                    strMsg = "必需為該月最大工作日!!"
                    Exit Function
                End If
            Case 3
                If RsQ.Fields(0) <> WorkDay Then
                    strMsg = "必需為該月第一個工作日!!"
                    Exit Function
                End If
        End Select
    Else
        strMsg = "有誤請洽電腦中心!!"
        Exit Function
    End If
    ChkWorkData = True
End Function

'Add by Morgan 2007/2/2
'檢查會計科目
'參數 p_AccNo:科目,p_bolMsg:是否彈錯誤訊息,p_Dept:部門
'回傳 0:正確, 1:科目錯, 2:部門錯, 3:智權人員
Public Function PUB_AccNoGood(p_AccNo As String, p_Dept As String, Optional p_Sales As String, Optional p_bolMsg As Boolean = True) As Integer
Dim strDept As String, strSql As String
Dim rstAdo As ADODB.Recordset, iRtn As Integer
   
   PUB_AccNoGood = 0
   
   'Add by Morgan 2007/10/2 若科目為4碼時檢查是否還有子科目
   If Len(p_AccNo) = 4 Then
      strSql = "select 1 from acc010 where substr(a0101,1,4)='" & p_AccNo & "' and length(a0101)>4 and rownum<2"
      iRtn = 1
      Set rstAdo = ClsLawReadRstMsg(iRtn, strSql)
      If iRtn = 1 Then
         PUB_AccNoGood = 1
         If p_bolMsg = True Then
            MsgBox "【" & p_AccNo & "】 為主科目，不可使用！"
         End If
         Set rstAdo = Nothing
         Exit Function
      End If
   End If
   'end 2007/10/2
   
   'Add by Morgan 2007/10/5 檢查部門代號是否為利潤中心部門
   If p_Dept <> "" Then
      strSql = "select 1 from acc090 where a0901='" & p_Dept & "' and a0904='Y'"
      iRtn = 1
      Set rstAdo = ClsLawReadRstMsg(iRtn, strSql)
      If iRtn <> 1 Then
         PUB_AccNoGood = 2
         If p_bolMsg = True Then
            MsgBox "部門代號【" & p_Dept & "】並非利潤中心部門，不可使用！"
         End If
         Set rstAdo = Nothing
         Exit Function
      End If
   End If
   'end 2007/10/5
   
   If Left(p_AccNo, 1) = "4" Then
'modify by sonia 2015/12/30 配合各部門之法務收入重新整理
'      If (p_AccNo >= "410101" And p_AccNo <= "410108") Or p_AccNo = "417202" Then
'         strDept = "T"
'      ElseIf (p_AccNo >= "411101" And p_AccNo <= "411105") Then
'         strDept = "P"
'      ElseIf p_AccNo = "4121" Then
'         strDept = "CFT"
'      ElseIf p_AccNo = "4131" Then
'         strDept = "CFP"
'      ElseIf (p_AccNo >= "414101" And p_AccNo <= "414102") Or (p_AccNo >= "418101" And p_AccNo <= "418102") Then
'         strDept = "L"
'      ElseIf p_AccNo = "416101" Then
'         strDept = "FCL"
'      ElseIf p_AccNo = "416102" Then
'         '2007/10/5 modify by sonia CFL非利潤中心部門
'         'strDept = "CFL"
'         strDept = "FCL"
'      '2009/4/17 MODIFY BY SONIA
'      'ElseIf p_AccNo = "4171" Then
'      ElseIf (p_AccNo >= "417101" And p_AccNo <= "417102") Then
'         strDept = "FCP"
'      ElseIf p_AccNo = "417201" Then
'         strDept = "FCT"
'      Else
'         strDept = p_Dept
'      End If
      If (p_AccNo >= "410101" And p_AccNo <= "410110" And p_AccNo <> "410109") Or p_AccNo = "417202" Then
         strDept = "T"
      ElseIf (p_AccNo >= "411101" And p_AccNo <= "411110" And p_AccNo <> "411106") Then
         strDept = "P"
      ElseIf (p_AccNo >= "4121" And p_AccNo <= "412110") Then
         strDept = "CFT"
      ElseIf (p_AccNo >= "4131" And p_AccNo <= "413110") Then
         strDept = "CFP"
      ElseIf (p_AccNo >= "414101" And p_AccNo <= "414110") Or (p_AccNo >= "418101" And p_AccNo <= "418110") Then
         strDept = "L"
      ElseIf (p_AccNo >= "416101" And p_AccNo <= "416110") Then
         strDept = "FCL"
      ElseIf (p_AccNo >= "417101" And p_AccNo <= "417110") Then
         strDept = "FCP"
      ElseIf (p_AccNo >= "417201" And p_AccNo <= "417210") Then
         strDept = "FCT"
      Else
         strDept = p_Dept
      End If
'end 2015/12/30

      If strDept <> p_Dept Then
         PUB_AccNoGood = 2
         If p_bolMsg = True Then
            MsgBox "【" & p_AccNo & "】 的部門必須為【" & strDept & "】，不可為【" & IIf(p_Dept = "", "空白", p_Dept) & "】！"
         End If
      End If
      If PUB_AccNoGood = 0 Then
         If p_Sales = "" Then
            PUB_AccNoGood = 3
            If p_bolMsg = True Then
               MsgBox "【" & p_AccNo & "】 的智權人員不可空白！"
            End If
         End If
      End If
   End If
   
   '2013/1/10 add by sonia 結餘保留科目一定要輸對沖(業)
   If Left(p_AccNo, 4) = "2491" Then
      If p_Sales = "" Then
         PUB_AccNoGood = 3
         If p_bolMsg = True Then
            MsgBox "【" & p_AccNo & "】 的智權人員不可空白！"
         End If
      End If
   End If
   '2013/1/10 end
End Function

'add by sonia 2015/12/30
'檢查民國105年起法務收入科目不可使用
'參數 p_AccNo:科目,p_Date:傳票日期,p_bolMsg:是否彈錯誤訊息
'回傳 0:正確, 1:錯
Public Function PUB_AccNoEnable(p_AccNo As String, p_Date As String, Optional p_bolMsg As Boolean = True) As Integer
   
   PUB_AccNoEnable = 0
   
   If Left(p_AccNo, 4) <> "4141" And Left(p_AccNo, 4) <> "4161" And Left(p_AccNo, 4) <> "4181" Then
      Exit Function
   ElseIf Val(p_Date) < 1050101 Then
      Exit Function
   Else
      PUB_AccNoEnable = 1
      If p_bolMsg = True Then
         MsgBox "民國 105 年起法務收入科目不可再使用！"
      End If
   End If
   
End Function
'end 2015/12/30

'Add By Sindy 2022/12/2 檢視檢洽單
Public Sub PUB_QueryFrm090801_Q(strCP140 As String, frmMe As Object, Optional bolLocked As Boolean = False)
Dim frmTmp As Form
Dim formCnt As Integer
   
   If strCP140 = "" Then Exit Sub
   
   If strSrvDate(1) >= 接洽單電子收文啟用日 Then
      '查詢接洽記錄單
      For formCnt = 0 To Forms.Count - 1
         If UCase(Forms(formCnt).Name) = UCase("frm090801_Q") Then
            If Forms(formCnt).Text5 = strCP140 Then
               Forms(formCnt).Show
               Forms(formCnt).ZOrder
               Exit Sub
            End If
         End If
      Next
'         If PUB_CheckFormExist("frm090801_Q") = True Then
'            Unload frm090801_Q
'         End If
      Set frmTmp = New frm090801_Q
      With frmTmp
         .SetParent frmMe
         .bolIsTmp = True
         .m_blnCallPrint = True
         .Text5 = strCP140
         Call .cmdOK_Click(4)
         'Add By Sindy 2023/1/12
         If bolLocked = True Then
            .Show vbModal
         Else
         '2023/1/12 END
            .Show
         End If
      End With
      Set frmTmp = Nothing
   End If
End Sub

'add by sonia 2023/11/13
Public Function PUB_GetShortName(ByVal p_SN02 As String) As String
   
   strSql = "select sn01 from salesno where sn02 = '" & p_SN02 & "'"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, adoTaie, adOpenForwardOnly, adLockReadOnly
      If .RecordCount > 0 Then
         PUB_GetShortName = "" & .Fields(0)
      End If
   End With
End Function
'end 2023/11/13

