Attribute VB_Name = "aacc_fun"
'Memo by Morgan2010/8/19 日期欄已修改
'Modified by Morgan 2012/1/9 ACDate(ServerDate) -> strSrvDate(2), ServerDate -> strSrvDate(1)
Option Explicit


'*************************************************
'  流水號之前補零
'
'*************************************************
Public Function ZeroBeforeNo(strInputValue As String, intInputLength As Integer) As String
Dim intCounter As Integer
   
   For intCounter = 1 To (intInputLength - Len(Trim(str(Val(strInputValue) + 1))))
      ZeroBeforeNo = ZeroBeforeNo & Mid(MsgText(12), 1, 1)
   Next intCounter
   ZeroBeforeNo = ZeroBeforeNo & (Val(strInputValue) + 1)
End Function

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
'         AutoNo = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo("2000", InputLength)
'      Else
'         AutoNo = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo("0", InputLength)
'      End If
'   Else
'      If adoaccnum.Fields("au02").Value <> Val(Mid(strSrvDate(1), 1, 4)) Then
'         If InputItem = "E" Then
'            AutoNo = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo("2000", InputLength)
'         Else
'            'Modified by Morgan 2012/1/9 收款單號有可能跨年預先使用
'            'AutoNo = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo("0", InputLength)
'            If InputItem = "F" Then
'               AutoNo = UpdateNo("acc0l0", "a0l01", 5, strSrvDate(2), InputItem)
'               If AutoNo = "" Then
'                  AutoNo = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo("0", InputLength)
'               End If
'            Else
'               AutoNo = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo("0", InputLength)
'            End If
'            'end 2012/1/9
'         End If
'      Else
'         AutoNo = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo(str(adoaccnum.Fields("au03").Value), InputLength)
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
'
''   adoTaie.Execute "UPDATE AUTONUMBER SET AU01 = '" & InputItem & "', AU02 = '" & Mid(strSrvDate(1), 1, 4) & "', AU03 = '" & InputNo & "' WHERE AU01 = '" & InputItem & "'"
'   adoaccnum.CursorLocation = adUseClient
'   adoaccnum.Open "select * from autonumber where au01 = '" & InputItem & "'", adoTaie, adOpenDynamic, adLockBatchOptimistic
'   If adoaccnum.RecordCount = 0 Then
'      adoaccnum.AddNew
'      adoaccnum.Fields("au01").Value = InputItem
'   End If
'   adoaccnum.Fields("au02").Value = Mid(strSrvDate(1), 1, 4)
'   adoaccnum.Fields("au03").Value = InputNo
'   adoaccnum.UpdateBatch
'   adoaccnum.Close
'   SaveAutoNo = MsgText(602)
'End Function
'end 2014/3/10

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

'Modify By Sindy 2014/5/27 Mark統一Move到basQuery中
''*************************************************
''  電腦自動給號(傳票號碼)
''
''*************************************************
'Public Function AccAutoNo(InputItem As String, InputLength As Integer, Optional intYear As Integer = 0, Optional intMonth As Integer = 0) As String
'Dim adoaccnum As New ADODB.Recordset
'Dim strItem As String, strYes As String
'
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
'         AccAutoNo = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo("0", InputLength)
'      Else
'         If adoaccnum.Fields("a1r02").Value <> Val(Mid(strSrvDate(1), 1, 4)) Then
'            AccAutoNo = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo("0", InputLength)
'         Else
'            AccAutoNo = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo(str(adoaccnum.Fields("a1r04").Value), InputLength)
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
'         '2012/3/30 MODIFY BY SONIA  3/30輸4月傳票會錯
'         'adoTaie.Execute "insert into acc1r0 (a1r01, a1r02, a1r03, a1r04) values ('" & InputItem & "', '" & Mid(strSrvDate(1), 1, 4) & "', '" & Mid(strSrvDate(1), 5, 2) & "', '" & InputNo & "')"
'         '2012/12/19 modify by sonia 101/12/19輸102/1/2會錯
'         'adoTaie.Execute "insert into acc1r0 (a1r01, a1r02, a1r03, a1r04) values ('" & InputItem & "', '" & Mid(strSrvDate(1), 1, 4) & "', '" & intMonth & "', '" & InputNo & "')"
'         adoTaie.Execute "insert into acc1r0 (a1r01, a1r02, a1r03, a1r04) values ('" & InputItem & "', '" & (intYear + 1911) & "', '" & intMonth & "', '" & InputNo & "')"
'      Else
'         adoTaie.Execute "UPDATE ACC1R0 SET A1R04 = " & InputNo & " WHERE A1R01 = '" & InputItem & "' and a1r02 = " & (intYear + 1911) & " and a1r03 = " & intMonth & ""
'      End If
'   Else
'      adoaccnum.Open "select * from acc1r0 where a1r01 = '" & InputItem & "'", adoTaie, adOpenStatic, adLockReadOnly
'      If adoaccnum.RecordCount = 0 Then
'         adoTaie.Execute "insert into acc1r0 (a1r01, a1r02, a1r03, a1r04) values ('" & InputItem & "', '" & Mid(strSrvDate(1), 1, 4) & "', '" & Mid(strSrvDate(1), 5, 2) & "', '" & InputNo & "')"
'      Else
'         adoTaie.Execute "UPDATE ACC1R0 SET A1R01 = '" & InputItem & "', A1R02 = '" & Mid(strSrvDate(1), 1, 4) & "', A1R03 = '" & Mid(strSrvDate(1), 5, 2) & "', A1R04 = '" & InputNo & "' WHERE A1R01 = '" & InputItem & "'"
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

'*************************************************
'  檢核資料是否存在
'
'*************************************************
Public Function ExistCheck(strTable As String, strField As String, strValue As String, strError As String, Optional bolMsg As Boolean = True) As Boolean
Dim adocheck As New ADODB.Recordset

   strValue = Replace(strValue, "'", "''")
   adocheck.CursorLocation = adUseClient
'   adocheck.Open "select " & strField & " from " & strTable & " where " & strField & " = " & "'" & strValue & "'", adoTaie, adOpenStatic, adLockReadOnly
   adocheck.Open "select " & strField & " from " & strTable & " where " & strField & " = " & "'" & ChgSQL(strValue) & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount = 0 Then
      If bolMsg Then
         MessageShow strError
      End If
      ExistCheck = False
   Else
      ExistCheck = True
   End If
   adocheck.Close
End Function

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

'Modify By Sindy 2014/5/27 Mark統一Move到basQuery中
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

'取得中英混雜字串之長度
Public Function GetTextLength_1(ByRef strTemp As String) As Integer
   GetTextLength_1 = LenB(StrConv(strTemp, vbFromUnicode))
End Function

'判斷中英混雜字串之長度是否有超過最大長度
Public Function CheckLengthIsOK_1(ByRef strTemp As String, ByRef intTemp As Integer) As Boolean
   If GetTextLength_1(strTemp) > intTemp Then
      Beep
      ShowMsg "輸入之資料過長，超過" & Format(intTemp) & "個字（註：中文算兩個字)"
      CheckLengthIsOK_1 = False
   Else
      CheckLengthIsOK_1 = True
   End If
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
'Modify by Amy 2021/05/17 +表單名
'Modify by Amy 2023/06/14 +stMsg
Public Function CheckDept(strAccNo As String, strDept As String, Optional stFormN As String = "", Optional ByRef stMsg As String) As Boolean
Dim adoacc010 As New ADODB.Recordset
   
   'Modify by  Amy 2023/06/14 +stMsg
   stMsg = ""
   If Mid(strAccNo, 1, 1) = "4" Then
      'Modify by Amy 2021/05/17 傳票會計科49開頭部門別只能輸TOT-瑞婷,秀玲:只能改傳票輸入,其他表單不可改
      If UCase(stFormN) = "FRMACC4120" And Mid(strAccNo, 1, 2) = "49" Then
         'Modify by Amy 2023/07/05 490101之部門別[只能]輸TOT;490102之部門別[不能]輸TOT
         If strAccNo = "490101" And strDept <> MsgText(55) Then
            stMsg = "【" & strAccNo & "】科目部門只能輸 " & MsgText(55)
            CheckDept = False
            Exit Function
         ElseIf strAccNo = "490102" And strDept = MsgText(55) Then
            stMsg = "【" & strAccNo & "】科目部門不能輸 " & MsgText(55)
            CheckDept = False
            Exit Function
         End If
      ElseIf strDept = MsgText(55) Or strDept = "" Then
         CheckDept = False
         Exit Function
      End If
   End If
   'add by sonia 2016/1/27 105年起不可再使用CFL,FCL,LA部門
   'cancel by sonia 2020/5/29 法律所成立再啟用CFL,FCL, LA以非利潤中心部門控制
   'If strDept = "CFL" Or strDept = "FCL" Or strDept = "LA" Then
   '   CheckDept = False
   '   Exit Function
   'End If
   'end 2020/5/29
   'end 2016/1/27
   'add by sonia 2016/3/23 105年起規費科目部門只可為TOT
   If Mid(strAccNo, 1, 4) = "2201" Then
      If strDept <> MsgText(55) And strDept <> "" Then
         CheckDept = False
         Exit Function
      End If
   End If
   If Mid(strAccNo, 1, 1) = "6" Then
      adoacc010.CursorLocation = adUseClient
      adoacc010.Open "select a0105 from acc010 where a0101 = '" & strAccNo & "'", adoTaie, adOpenStatic, adLockReadOnly
      If adoacc010.RecordCount <> 0 Then
         If IsNull(adoacc010.Fields(0).Value) Then
            'Modify by Morgan 2004/11/17 沒有設定分攤類別的費用一定要輸入非本所的部門(原來拿掉，現在又還原)
            If strDept = MsgText(55) Or strDept = "" Then
               CheckDept = False
            Else
               CheckDept = True
            End If
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

''將From移至畫面之中心
'Public Sub MoveFormToCenter(ByRef frmTemp As Form)
'Dim intX  As Integer, intY As Integer
'
'   If frmTemp.MDIChild Then
'      intX = (Frmacc0000.ScaleWidth - frmTemp.Width) / 2
'      intY = (Frmacc0000.ScaleHeight - frmTemp.Height) / 2
'   Else
'      intX = (Screen.Width - frmTemp.Width) / 2
'      intY = (Screen.Height - frmTemp.Height) / 2
'   End If
'   frmTemp.Move intX, intY
'End Sub

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
'Removed by Morgan 2014/3/10 整合同名函數到 basQuery
''*************************************************
''  自動給過去年度之流水號
''
''*************************************************
'Public Function UpdateNo(strTable As String, strField As String, intLength As Integer, strDate As String, Optional strDocNo As String = "") As String
'Dim adoaccnum As New ADODB.Recordset
'
'On Error GoTo Checking
'   adoaccnum.CursorLocation = adUseClient
'   adoaccnum.Open "select nvl(max(" & strField & "), 0) from " & strTable & " where substr(" & strField & ", 2, 3) = '" & Mid(strDate, 1, 3) & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccnum.RecordCount <> 0 Then
'      If IsNull(adoaccnum.Fields(0).Value) = False Then
'         If adoaccnum.Fields(0).Value = "0" Then
'            UpdateNo = strDocNo & Mid(strDate, 1, 3) & ZeroBeforeNo(Mid(adoaccnum.Fields(0).Value, 5, intLength), intLength)
'         Else
'            UpdateNo = Mid(adoaccnum.Fields(0).Value, 1, 4) & ZeroBeforeNo(Mid(adoaccnum.Fields(0).Value, 5, intLength), intLength)
'         End If
''         strYes = SaveAutoNo(Mid(adoaccnum.Fields(0).Value, 1, 1), Mid(UpdateNo, 5, intLength))
'      End If
'   End If
'   adoaccnum.Close
'Checking:
'   If Err.Number = 0 Then
'      Exit Function
'   End If
'   MsgBox Err.Description, , MsgText(5)
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
Public Function AccNoToSalesNo(strAccNo As String, Optional strCaseNo As String) As String
'add by sonia 2021/1/18 以本所案號判別FCP,FCT英日文組
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim stCP13 As String
Dim strCaseNo1 As String, strCaseNo2 As String, StrCaseNo3 As String, strCaseNo4 As String

   stCP13 = ""
   If strCaseNo <> "" Then
      strCaseNo = CaseNoZero(strCaseNo) 'Add by Amy 2024/08/01 避免完整滿造成錯誤
      strCaseNo1 = Mid(strCaseNo, 1, Len(strCaseNo) - 9)
      strCaseNo2 = Mid(strCaseNo, Len(strCaseNo) - 8, 6)
      StrCaseNo3 = Mid(strCaseNo, Len(strCaseNo) - 2, 1)
      strCaseNo4 = Mid(strCaseNo, Len(strCaseNo) - 1, 2)
      stCP13 = PUB_GetAKindSalesNo(strCaseNo1, strCaseNo2, StrCaseNo3, strCaseNo4)
'      '若為法務案再抓該案最新之案源介紹人
'      If InStr(strCaseNo1, "L") > 0 Then
'         StrSQLa = "Select ST16 From staff Where st01='" & stCP13 & "'"
'         rsA.CursorLocation = adUseClient
'         rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount > 0 Then
'         End If
'      End If
   End If
'end 2021/1/18
   
   'modify by sonia 2020/12/30 改取4碼
   'Select Case strAccNo
   Select Case Left(strAccNo, 4)
      '2009/4/17 MODIFY BY SONIA
      'Case "4171"
      'modify by sonia 2016/8/1 +417104,417105,417109
      'modify by sonia 2020/5/21 參考SalesNoToAccSales +417103
      'modify by sonia 2020/12/30 改取4碼
      'Case "417101", "417102", "417104", "417105", "417109", "417103"
      Case "4171"
         'modify by sonia 2021/1/18
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
      'modify by sonia 2020/5/21 參考SalesNoToAccSales +417203
      'modify by sonia 2020/12/30 改取4碼
      'Case "417201", "417202", "417203"
      Case "4172"
         'modify by sonia 2021/1/18
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
      'modify by sonia 2020/12/30 改取4碼
      'Case "416101", "416102"
      Case "4161"
         AccNoToSalesNo = "F4101"
      Case Else
         AccNoToSalesNo = ""
   End Select
End Function

'*************************************************
'  依智權人員帶出做帳智權人員-->國外收款
'  92.10.9 add by sonia
'*************************************************
'Modify By Cheng 2004/05/11
'加參數--本所案號
'Public Function SalesNoToAccSales(strSalesNo As String, strAccNo As String) As String
'modify by sonia 2021/3/12 加參數--日期2021/4起商標MCT小組收款不掛M0100改掛P2005
Public Function SalesNoToAccSales(strSalesNo As String, strAccNo As String, ByVal strCaseNo As String, ByVal strDate As String) As String
Dim adoquery As New ADODB.Recordset
Dim blnNoFOrS As Boolean '非國外部人員也非國內智權人員
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strSysKind As String  'add by sonia 2017/8/22
   
   strSysKind = Mid(strCaseNo, 1, Len(strCaseNo) - 9)  'add by sonia 2017/8/22
   blnNoFOrS = False
   'add by sonia 2022/2/16 F4102,F4103改抓案件智權人員(A1p04='A11100288Y34181B10')
   If strSalesNo = "F4102" Then
      strSalesNo = PUB_GetFCPSalesNo(Mid(strCaseNo, 1, Len(strCaseNo) - 9), Mid(strCaseNo, Len(strCaseNo) - 8, 6), Mid(strCaseNo, Len(strCaseNo) - 2, 1), Mid(strCaseNo, Len(strCaseNo) - 1, 2))
   ElseIf strSalesNo = "F4103" Then
      strSalesNo = PUB_GetFCTSalesNo(Mid(strCaseNo, 1, Len(strCaseNo) - 9), Mid(strCaseNo, Len(strCaseNo) - 8, 6), Mid(strCaseNo, Len(strCaseNo) - 2, 1), Mid(strCaseNo, Len(strCaseNo) - 1, 2))
   End If
   'end 2022/2/16
   adoquery.CursorLocation = adUseClient
   'modify by sonia 2020/12/30 +st16非日文組都算英文組
   adoquery.Open "select ST15,ST16 from staff where st01 = '" & strSalesNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If adoquery.Fields(0).Value >= "F" And adoquery.Fields(0).Value <= "G" Then
         blnNoFOrS = False
         'modify by sonia 2020/12/30 改取4碼
         'Select Case strAccNo
         Select Case Left(strAccNo, 4)
            '2009/4/17 MODIFY BY SONIA
            'Case "4171"
            'modify by sonia 2016/8/1 +417104,417105,417109
            'modify by sonia 2017/8/18 +417103 (X10610486)
            'modify by sonia 2020/12/30 改取4碼
            'Case "417101", "417102", "417104", "417105", "417109", "417103"
            Case "4171"
               'modify by sonia 2020/12/30
               'SalesNoToAccSales = "F4102"
               If "" & adoquery.Fields(1).Value = "2" Then
                  SalesNoToAccSales = "F4105"
               Else
                  If "" & adoquery.Fields(1).Value = "" Then
                     SalesNoToAccSales = "F4102"
                  Else
                     SalesNoToAccSales = "F4104"
                  End If
               End If
               'end 2020/12/30
            'modify by sonia 2017/8/18 +417203
            'modify by sonia 2020/12/30 改取4碼
            'Case "417201", "417202", "417203", "417203"
            Case "4172"
               'modify by sonia 2020/12/30
               'SalesNoToAccSales = "F4103"
               If "" & adoquery.Fields(1).Value = "4" Then
                  SalesNoToAccSales = "F4107"
               Else
                  If "" & adoquery.Fields(1).Value = "" Then
                     SalesNoToAccSales = "F4103"
                  Else
                     SalesNoToAccSales = "F4106"
                  End If
               End If
               'end 2020/12/30
            'modify by sonia 2020/12/30 改取4碼
            'Case "416101", "416102"
            Case "4161"
               SalesNoToAccSales = "F4101"
               'add by sonia 2017/8/22 M10604167之X10610486
               Select Case strSysKind
                  Case "P", "PS", "CFP", "CPS", "FCP", "FG"  '專利
                     'modify by sonia 2020/12/30
                     'SalesNoToAccSales = "F4102"
                     If "" & adoquery.Fields(1).Value = "2" Then
                        SalesNoToAccSales = "F4105"
                     Else
                        SalesNoToAccSales = "F4104"
                     End If
                     'end 2020/12/30
                  Case Else   '商標
                     'modify by sonia 2020/12/30
                     'SalesNoToAccSales = "F4103"
                     If "" & adoquery.Fields(1).Value = "4" Then
                        SalesNoToAccSales = "F4107"
                     Else
                        SalesNoToAccSales = "F4106"
                     End If
                     'end 2020/12/30
               End Select
               'end 2017/8/22
            Case Else
               If adoquery.Fields(0).Value >= "F1" And adoquery.Fields(0).Value <= "F19" Then
                  'modify by sonia 2020/12/30
                  'SalesNoToAccSales = "F4103"
                  If "" & adoquery.Fields(1).Value = "4" Then
                     SalesNoToAccSales = "F4107"
                  Else
                     If adoquery.Fields(1).Value = "" Then
                        SalesNoToAccSales = "F4103"
                     Else
                        SalesNoToAccSales = "F4106"
                     End If
                  End If
                  'end 2020/12/30
               ElseIf adoquery.Fields(0).Value >= "F2" And adoquery.Fields(0).Value <= "F29" Then
                  'modify by sonia 2020/12/30
                  'SalesNoToAccSales = "F4102"
                  If "" & adoquery.Fields(1).Value = "2" Then
                     SalesNoToAccSales = "F4105"
                  Else
                     If adoquery.Fields(1).Value = "" Then
                        SalesNoToAccSales = "F4102"
                     Else
                        SalesNoToAccSales = "F4104"
                     End If
                  End If
                  'end 2020/12/30
               ElseIf adoquery.Fields(0).Value >= "F3" And adoquery.Fields(0).Value <= "F49" Then
                  SalesNoToAccSales = "F4101"
               Else
                  SalesNoToAccSales = ""
               End If
         End Select
      'add by sonia 2021/3/12
      ElseIf adoquery.Fields(0).Value >= "P2" And adoquery.Fields(0).Value <= "P29" And strDate > "1100400" Then
         blnNoFOrS = False
         SalesNoToAccSales = "P2005"
      'end 2021/3/12
      '2007/10/17 add by sonia 巨京商標96029,96030
      'Modify by Morgan 2010/6/4 +巨京專利96031,96032
      ElseIf strSalesNo >= "96029" And strSalesNo <= "96032" Then
         blnNoFOrS = False
         SalesNoToAccSales = strSalesNo
      ElseIf adoquery.Fields(0).Value < "S" Then
         blnNoFOrS = True
         SalesNoToAccSales = "M0100"
      ElseIf adoquery.Fields(0).Value >= "T" Then
         blnNoFOrS = True
         SalesNoToAccSales = "M0100"
      Else
         blnNoFOrS = False
         SalesNoToAccSales = strSalesNo
      End If
   Else
      SalesNoToAccSales = ""
   End If
   adoquery.Close
   
   '93.8.17 ADD BY SONIA 智權人員為郭雅娟 79075 時,固定為 M0100
   If strSalesNo = "79075" Then
      'modify by sonia 2022/3/15 日期2022/4起79075收款不掛M0100改掛P1005
      'SalesNoToAccSales = "M0100"
      If strDate > "1110400" Then
         SalesNoToAccSales = "P1005"
      End If
      'end 2022/3/15
  Else
   '93.8.17 END
      'Add By Cheng 2004/05/11
      '若非國外部人員也非國內智權人員
      If blnNoFOrS = True Then
          StrSQLa = "Select PA75 From Patent Where " & ChgPatent(strCaseNo)
          StrSQLa = StrSQLa & " Union Select TM44 From Trademark Where " & ChgTradeMark(strCaseNo)
          StrSQLa = StrSQLa & " Union Select LC22 From Lawcase Where " & ChgLawcase(strCaseNo)
          StrSQLa = StrSQLa & " Union Select '' From Hirecase Where " & ChgHirecase(strCaseNo)
          StrSQLa = StrSQLa & " Union Select SP26 From Servicepractice Where " & ChgService(strCaseNo)
          rsA.CursorLocation = adUseClient
          rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
          If rsA.RecordCount > 0 Then
              '若有FC代理人
              If "" & rsA.Fields(0).Value <> "" Then
                  SalesNoToAccSales = "M0100"
              Else
                  SalesNoToAccSales = strSalesNo
              End If
          Else
              SalesNoToAccSales = strSalesNo
          End If
          If rsA.State <> adStateClosed Then rsA.Close
          Set rsA = Nothing
      End If
   End If
End Function

'*************************************************
'  依智權人員帶出做帳智權人員-->國內收款
'  92.12.22 add by sonia
'*************************************************
Public Function SalesNoToSales(strSalesNo As String, strAccNo As String) As String
Dim adoquery As New ADODB.Recordset
   
   adoquery.CursorLocation = adUseClient
   'modify by sonia 2022/1/22 +st16非日文組都算英文組
   adoquery.Open "select ST15,ST16 from staff where st01 = '" & strSalesNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoquery.RecordCount <> 0 Then
      If adoquery.Fields(0).Value >= "F" And adoquery.Fields(0).Value <= "G" Then
         'modify by sonia 2022/1/22 改取4碼
         'Select Case strAccNo
         Select Case Left(strAccNo, 4)
            '2009/4/17 MODIFY BY SONIA
            'Case "4171"
            'modify by sonia 2016/8/1 +417104,417105,417109
            'modify by sonia 2022/1/22 改取4碼
            'Case "417101", "417102", "417104", "417105", "417109"
            Case "4171"
               'modify by sonia 2022/1/22
               'SalesNoToSales = "F4102"
               If "" & adoquery.Fields(1).Value = "2" Then
                  SalesNoToSales = "F4105"
               Else
                  If "" & adoquery.Fields(1).Value = "" Then
                     SalesNoToSales = "F4102"
                  Else
                     SalesNoToSales = "F4104"
                  End If
               End If
               'end 2022/1/22
            'modify by sonia 2022/1/22 改取4碼
            'Case "417201", "417202"
            Case "4172"
               'modify by sonia 2022/1/22
               'SalesNoToSales = "F4103"
               If "" & adoquery.Fields(1).Value = "4" Then
                  SalesNoToSales = "F4107"
               Else
                  If "" & adoquery.Fields(1).Value = "" Then
                     SalesNoToSales = "F4103"
                  Else
                     SalesNoToSales = "F4106"
                  End If
               End If
               'end 2022/1/22
            'cancel by sonia 2016/9/14 已不使用416101,416102 (F10509364之1203智權人員會帶成F4101)
            'Case "416101", "416102"
            '   SalesNoToSales = "F4101"
            Case Else
               If adoquery.Fields(0).Value >= "F1" And adoquery.Fields(0).Value <= "F19" Then
                  'modify by sonia 2022/1/22
                  'SalesNoToSales = "F4103"
                  If "" & adoquery.Fields(1).Value = "4" Then
                     SalesNoToSales = "F4107"
                  Else
                     If adoquery.Fields(1).Value = "" Then
                        SalesNoToSales = "F4103"
                     Else
                        SalesNoToSales = "F4106"
                     End If
                  End If
                  'end 2022/1/22
               ElseIf adoquery.Fields(0).Value >= "F2" And adoquery.Fields(0).Value <= "F29" Then
                  'modify by sonia 2022/1/22
                  'SalesNoToSales = "F4102"
                  If "" & adoquery.Fields(1).Value = "2" Then
                     SalesNoToSales = "F4105"
                  Else
                     If adoquery.Fields(1).Value = "" Then
                        SalesNoToSales = "F4102"
                     Else
                        SalesNoToSales = "F4104"
                     End If
                  End If
                  'end 2022/1/22
               ElseIf adoquery.Fields(0).Value >= "F3" And adoquery.Fields(0).Value <= "F49" Then
                  'modify by sonia 2017/8/23 外法改M0100(F10608407)
                  'SalesNoToSales = "F4101"
                  SalesNoToSales = "M0100"
               Else
                  SalesNoToSales = ""
               End If
         End Select
      Else
         SalesNoToSales = strSalesNo
      End If
   Else
      SalesNoToSales = ""
   End If
   adoquery.Close
   'add by sonia 2022/1/22 智權人員為郭雅娟 79075 時,固定為 M0100  客戶X83843溢泰(南京)開國內收據
   If strSalesNo = "79075" Then
      'modify by sonia 2022/6/6 由M0100改為P1005
      SalesNoToSales = "P1005"
   End If
   'end 2022/1/22
End Function

'add by sonia 2021/1/27 國外部業務對沖與科目檢查 國外收款單M10905805第010項次輸錯業務對沖
Public Function SalesNoCheckAccNo(strAccNo As String, strSalesNo As String) As Boolean
   SalesNoCheckAccNo = True
   Select Case Left(strAccNo, 4)
      Case "4171"  'FCP
         If strSalesNo <> "F4102" And strSalesNo <> "F4104" And strSalesNo <> "F4105" Then
            MsgBox "ＦＣＰ收入科目，業務非該部門人員，請注意！", , MsgText(5)
            SalesNoCheckAccNo = False
         End If
      Case "4172"  'FCT
         If strSalesNo <> "F4103" And strSalesNo <> "F4106" And strSalesNo <> "F4107" Then
            MsgBox "ＦＣＴ收入科目，業務非該部門人員，請注意！", , MsgText(5)
            SalesNoCheckAccNo = False
         End If
   End Select
End Function

'Add By Cheng 2004/03/09
'結匯明細匯總表(印完付款明細草稿後產生的)
Public Sub PUB_ExcelSave()
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
Dim strCompany As String '公司別
Dim strKind As String '匯款方式
Dim xlsSalesPoint As New Excel.Application
Dim wktmp As New Worksheet
Dim lngCounter As Long
Dim lngLocation As Long
Dim lngRow As Long
Dim ii As Long
Dim lngMaxCounter As Long '最大列數
Dim lngMaxLocation As Long '最大欄數

   StrSQLa = "Select R21805 From ACCRPT218 Group By R21805 Order By 1"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      If Dir(strExcelPath & "結匯明細匯總表" & ACDate(strSrvDate(1)) & ServerTime & MsgText(43)) = "" Then
         If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = "" Then
            MkDir strExcelPath
         End If
      Else
         Kill strExcelPath & "結匯明細匯總表" & ACDate(strSrvDate(1)) & ServerTime & MsgText(43)
      End If
      xlsSalesPoint.Workbooks.add
      Set wktmp = xlsSalesPoint.Worksheets(1)
      wktmp.Range("a1").Value = "結匯明細匯總表"
      lngCounter = 1: lngLocation = 1
      '預設公司別為1匯款方式1票匯
      strCompany = "1": strKind = "1"
ReDo:
      Select Case strCompany & strKind
      Case "11"
          StrSqlB = "Select R21808, a1811, R21803, R21805, Sum(R21806) From ACCRPT218, ACC180 Where R21802=A1801(+) And R21808='1' And A1811='1' Group By R21808, R21803, A1811, R21805 Order By 1, 2, 3, 4 "
      Case "12"
          StrSqlB = "Select R21808, a1811, R21803, R21805, Sum(R21806) From ACCRPT218, ACC180 Where R21802=A1801(+) And R21808='1' And A1811='2' Group By R21808, R21803, A1811, R21805 Order By 1, 2, 3, 4 "
      Case "21"
          StrSqlB = "Select R21808, a1811, R21803, R21805, Sum(R21806) From ACCRPT218, ACC180 Where R21802=A1801(+) And R21808='2' And A1811='1' Group By R21808, R21803, A1811, R21805 Order By 1, 2, 3, 4 "
      Case "22"
          StrSqlB = "Select R21808, a1811, R21803, R21805, Sum(R21806) From ACCRPT218, ACC180 Where R21802=A1801(+) And R21808='2' And A1811='2' Group By R21808, R21803, A1811, R21805 Order By 1, 2, 3, 4 "
      Case Else
         If lngMaxLocation = 0 Then lngMaxLocation = 1 ' 'Modified by Lydia 2015/02/25 避免零值出錯
         wktmp.Range(wktmp.Cells(1, 1), wktmp.Cells(1, lngMaxLocation)).Select
         With xlsSalesPoint.Selection
             .HorizontalAlignment = xlCenter
             .VerticalAlignment = xlBottom
             .WrapText = False
             .Orientation = 0
             .AddIndent = False
             .ShrinkToFit = False
             .MergeCells = True
         End With
         wktmp.Range(wktmp.Cells(1, 1), wktmp.Cells(lngCounter, lngMaxLocation)).Select
         With xlsSalesPoint.Selection
             .Columns.AutoFit
         End With
         xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExcelPath & "結匯明細匯總表" & ACDate(strSrvDate(1)) & ServerTime & MsgText(43)
         xlsSalesPoint.Workbooks.Close
         xlsSalesPoint.Quit
         Set xlsSalesPoint = Nothing
         If rsB.State <> adStateClosed Then rsB.Close
         Set rsB = Nothing
         If rsA.State <> adStateClosed Then rsA.Close
         Set rsA = Nothing
         Exit Sub
      End Select
      rsB.CursorLocation = adUseClient
      rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
      If rsB.RecordCount > 0 Then
         lngCounter = lngCounter + 1
         wktmp.Cells(lngCounter, 1).Value = strCompany & " 公司"
         wktmp.Cells(lngCounter, 2).Value = IIf(strKind = "1", "票匯", "電匯")
         lngCounter = lngCounter + 1
         '幣別
         lngRow = lngCounter
         lngLocation = 1
         rsA.MoveFirst
         While Not rsA.EOF
             lngLocation = lngLocation + 1
             wktmp.Cells(lngCounter, lngLocation).Value = "" & rsA.Fields(0).Value
             rsA.MoveNext
         Wend
         lngMaxLocation = lngLocation
         lngMaxCounter = 0
         While Not rsB.EOF
             ii = 2
             Do While "" & rsB.Fields(3).Value <> wktmp.Cells(lngRow, ii)
                 ii = ii + 1
             Loop
             lngCounter = lngRow
             Do While wktmp.Cells(lngCounter, ii) <> ""
                 lngCounter = lngCounter + 1
             Loop
             If lngCounter > lngMaxCounter Then lngMaxCounter = lngCounter
             wktmp.Cells(lngCounter, ii).Value = "" & rsB.Fields(4).Value
             rsB.MoveNext
         Wend
         lngCounter = lngMaxCounter + 1
         If rsB.State <> adStateClosed Then rsB.Close
         Set rsB = Nothing
         '匯款方式小計
         Select Case strCompany & strKind
         Case "11"
             StrSqlB = "Select R21808, a1811, R21805, Sum(R21806) From ACCRPT218, ACC180 Where R21802=A1801(+) And R21808='1' And A1811='1' Group By R21808, A1811, R21805 Order By 1, 2, 3 "
         Case "12"
             StrSqlB = "Select R21808, a1811, R21805, Sum(R21806) From ACCRPT218, ACC180 Where R21802=A1801(+) And R21808='1' And A1811='2' Group By R21808, A1811, R21805 Order By 1, 2, 3 "
         Case "21"
             StrSqlB = "Select R21808, a1811, R21805, Sum(R21806) From ACCRPT218, ACC180 Where R21802=A1801(+) And R21808='2' And A1811='1' Group By R21808, A1811, R21805 Order By 1, 2, 3 "
         Case "22"
             StrSqlB = "Select R21808, a1811, R21805, Sum(R21806) From ACCRPT218, ACC180 Where R21802=A1801(+) And R21808='2' And A1811='2' Group By R21808, A1811, R21805 Order By 1, 2, 3 "
         End Select
         rsB.CursorLocation = adUseClient
         rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
         If rsB.RecordCount > 0 Then
             wktmp.Cells(lngCounter, 1).Value = IIf(strKind = "1", "票匯小計", "電匯小計")
             While Not rsB.EOF
                 ii = 2
                 Do While "" & rsB.Fields(2).Value <> wktmp.Cells(lngRow, ii)
                     ii = ii + 1
                 Loop
                 wktmp.Cells(lngCounter, ii).Value = "" & rsB.Fields(3).Value
                 rsB.MoveNext
             Wend
         End If
         lngCounter = lngCounter + 1
      End If
      If rsB.State <> adStateClosed Then rsB.Close
      Set rsB = Nothing
      '公司別小計
      If strKind = "2" Then
          If strCompany = "1" Then
              StrSqlB = "Select R21808, R21805, Sum(R21806) From ACCRPT218, ACC180 Where R21802=A1801(+) And R21808='1' Group By R21808, R21805 Order By 1, 2 "
          ElseIf strCompany = "2" Then
              StrSqlB = "Select R21808, R21805, Sum(R21806) From ACCRPT218, ACC180 Where R21802=A1801(+) And R21808='2' Group By R21808, R21805 Order By 1, 2 "
          End If
          rsB.CursorLocation = adUseClient
          rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
          If rsB.RecordCount > 0 Then
              wktmp.Cells(lngCounter, 1).Value = strCompany & " 公司小計"
              While Not rsB.EOF
                  ii = 2
                  Do While "" & rsB.Fields(1).Value <> wktmp.Cells(lngRow, ii)
                      ii = ii + 1
                  Loop
                  wktmp.Cells(lngCounter, ii).Value = "" & rsB.Fields(2).Value
                  rsB.MoveNext
              Wend
              lngCounter = lngCounter + 1
          End If
          If rsB.State <> adStateClosed Then rsB.Close
          Set rsB = Nothing
      End If
      If strCompany & strKind = "11" Then
          strCompany = "1": strKind = "2"
          GoTo ReDo
      ElseIf strCompany & strKind = "12" Then
          strCompany = "2": strKind = "1"
          GoTo ReDo
      ElseIf strCompany & strKind = "21" Then
          strCompany = "2": strKind = "2"
          GoTo ReDo
      ElseIf strCompany & strKind = "22" Then
          strCompany = "": strKind = ""
          GoTo ReDo
      End If
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing

End Sub

Public Sub PUB_InitCombo(ByRef p_Combo As Object)
   p_Combo.Clear
   p_Combo.AddItem "專利申請"
   p_Combo.AddItem "商標申請"
   p_Combo.AddItem "國外專利申請"
   p_Combo.AddItem "國外商標申請"
   p_Combo.AddItem "國內外專利申請"
   p_Combo.AddItem "國內外商標申請"
   p_Combo.AddItem "發明專利申請"
   p_Combo.AddItem "新型專利申請"
   p_Combo.AddItem "設計專利申請"
   p_Combo.AddItem "資料檢索"
End Sub

'Add by Morgan 2005/11/30
Public Sub PUB_SetToolBar(ByVal p_iState As Integer)
   
   Select Case p_iState
      Case 1
         tool1_enabled
      Case 2
         tool2_enabled
      Case 3
         tool3_enabled
      Case 4
         tool4_enabled
      Case 5
         tool5_enabled
      Case 6
         tool6_enabled
      Case 7
         tool7_enabled
      Case 8
         tool8_enabled
      Case 9
         tool9_enabled
      Case 10
         tool10_enabled
      Case 11
         tool11_enabled
      Case 12
         tool12_enabled
      Case 13
         tool13_enabled
      Case 14
         tool14_enabled
      Case 15
         tool17_enabled
      Case 16
         tool18_enabled
         
   End Select
End Sub

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

'2005/11/4 ADD BY SONIA
'取得國內收據抬頭
Public Function GetA0K04(strCaseNo As String, strCP09 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   GetA0K04 = ""
   If strCP09 <> "" Then
      '2007/8/22 MODIFY BY SONIA 國內收據(CP60為'E'字頭)抓收據抬頭,國外請款單(CP60非'E'字頭)抓客戶名稱,
      '                          CP60為空者先抓該案母號有收據且收文日最大的收據抬頭,若都無國內收據則抓客戶名稱
      '2021/7/19 再改若本身無收據，則先抓相關總收文號cp43的收據，之後再考慮母號有收據且收文日最大,若都無國內收據則抓客戶名稱
      'strSQLA = "select A0K04 from caseprogress, acc0k0 where cp60 = a0k01 and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' AND SUBSTR(CP60,1,1)='E' union " & _
      '          "select CU04  from caseprogress, PATENT,CUSTOMER WHERE CP01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and CP02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and CP03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and CP04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND (SUBSTR(CP60,1,1)<>'E' OR CP60 IS NULL) and SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) UNION " & _
      '          "select CU04  from caseprogress, TRADEMARK,CUSTOMER WHERE CP01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and CP02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and CP03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and CP04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' AND CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND (SUBSTR(CP60,1,1)<>'E' OR CP60 IS NULL) and SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) UNION " & _
      '          "select CU04  from caseprogress, SERVICEPRACTICE,CUSTOMER WHERE CP01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and CP02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and CP03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and CP04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' AND CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04 AND (SUBSTR(CP60,1,1)<>'E' OR CP60 IS NULL) and SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) UNION " & _
      '          "select CU04  from caseprogress, LAWCASE,CUSTOMER WHERE CP01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and CP02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and CP03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and CP04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND (SUBSTR(CP60,1,1)<>'E' OR CP60 IS NULL) and SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) UNION " & _
      '          "select CU04  from caseprogress, HIRECASE,CUSTOMER WHERE CP01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and CP02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and CP03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and CP04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND (SUBSTR(CP60,1,1)<>'E' OR CP60 IS NULL) and SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) "
      'rsA.CursorLocation = adUseClient
      'rsA.Open strSQLA, adoTaie, adOpenStatic, adLockReadOnly
      'If rsA.RecordCount > 0 Then
      '   GetA0K04 = "" & rsA.Fields(0).Value
      'End If
      'Modified by Morgan 2011/12/23 考慮拆收據情形
      'StrSQLa = "select A0K04 from caseprogress, acc0k0 where cp60 = a0k01 and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' AND SUBSTR(CP60,1,1)='E' "
      StrSQLa = "select A0K04 from acc0j0, acc0k0 where  a0k01(+)=a0j13 and a0j02='" & strCaseNo & "' and a0j01 = '" & strCP09 & "' AND SUBSTR(a0j13,1,1)='E' "
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         GetA0K04 = "" & rsA.Fields(0).Value
      Else
         If rsA.State <> adStateClosed Then rsA.Close
         'add by sonia 2021/7/19 再改若本身無收據，則先抓相關總收文號cp43的收據，之後再考慮母號有收據且收文日最大,若都無國內收據則抓客戶名稱
         StrSQLa = "select A0K04 from caseprogress c1,caseprogress c2,acc0j0,acc0k0 where c1.cp09='" & strCP09 & "' and c1.cp43=c2.cp09(+) AND SUBSTR(c2.CP60,1,1)='E' and a0j01(+)=c2.cp09 and a0k01(+)=a0j13 "
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            GetA0K04 = "" & rsA.Fields(0).Value
         End If
         'end 2021/7/19
         '先抓該案母號有收據且收文日最大的收據抬頭
         'Modified by Morgan 2011/12/23 考慮拆收據情形
         'StrSQLa = "select SUBSTR(MAX(CP05||A0K04), 9) from caseprogress, acc0k0 where cp60 = a0k01 and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '00' AND SUBSTR(CP60,1,1)='E' "
         If GetA0K04 = "" Then    'add by sonia 2021/7/19
            If rsA.State <> adStateClosed Then rsA.Close
            StrSQLa = "select SUBSTR(MAX(CP05||A0K04), 9) from caseprogress,acc0j0,acc0k0 where a0j01(+)=cp09 and a0k01(+)=a0j13 and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '00' AND SUBSTR(CP60,1,1)='E' "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               GetA0K04 = "" & rsA.Fields(0).Value
            End If
         End If   'add by sonia 2021/7/19
         If GetA0K04 = "" Then
            If rsA.State <> adStateClosed Then rsA.Close
            '2011/12/7 MODIFY BY SONIA SUBSTR(CP60,1,1)<>'E'改為(CP60 IS NULL OR SUBSTR(CP60,1,1)<>'E') FCP044259000
            StrSQLa = "select CU04 from caseprogress, PATENT,CUSTOMER WHERE CP01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and CP02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and CP03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and CP04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' AND CP01=PA01 AND CP02=PA02 AND CP03=PA03 AND CP04=PA04 AND (CP60 IS NULL OR SUBSTR(CP60,1,1)<>'E') and SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) UNION " & _
                      "select CU04 from caseprogress, TRADEMARK,CUSTOMER WHERE CP01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and CP02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and CP03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and CP04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' AND CP01=TM01 AND CP02=TM02 AND CP03=TM03 AND CP04=TM04 AND (CP60 IS NULL OR SUBSTR(CP60,1,1)<>'E') and SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) UNION " & _
                      "select CU04 from caseprogress, SERVICEPRACTICE,CUSTOMER WHERE CP01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and CP02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and CP03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and CP04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' AND CP01=SP01 AND CP02=SP02 AND CP03=SP03 AND CP04=SP04 AND (CP60 IS NULL OR SUBSTR(CP60,1,1)<>'E') and SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) UNION " & _
                      "select CU04 from caseprogress, LAWCASE,CUSTOMER WHERE CP01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and CP02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and CP03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and CP04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' AND CP01=LC01 AND CP02=LC02 AND CP03=LC03 AND CP04=LC04 AND (CP60 IS NULL OR SUBSTR(CP60,1,1)<>'E') and SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) UNION " & _
                      "select CU04 from caseprogress, HIRECASE,CUSTOMER WHERE CP01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and CP02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and CP03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and CP04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' AND CP01=HC01 AND CP02=HC02 AND CP03=HC03 AND CP04=HC04 AND (CP60 IS NULL OR SUBSTR(CP60,1,1)<>'E') and SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
               GetA0K04 = "" & rsA.Fields(0).Value
            End If
         End If
      End If
      '2007/8/22 END
      If rsA.State <> adStateClosed Then rsA.Close
      Set rsA = Nothing
   End If

End Function

'Add By Cheng 2003/07/23
'取得國內收款日
'2011/8/19 modify by sonia加傳收款金額strDomAmt1(有format),strDomAmt2(數字),
Public Function GetA1l02(strCaseNo As String, strCP09 As String, strDomAmt1 As String, Optional strDomAmt2 As Double) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

   GetA1l02 = ""
   strDomAmt1 = "0": strDomAmt2 = 0
   If strCP09 <> "" Then
      '2011/8/19 add by sonia 改同GetA0K04規則,CP60為空者改抓該案母號有收據且收文日最大的收據的國內收款日
      '2021/7/19 再改CP60為空者先抓其相關總收文號之CP60,再無才改抓該案母號有收據且收文日最大的收據的國內收款日
      StrSQLa = "select * from caseprogress where cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' and cp60 is null"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
         'add by sonia 2023/11/27 FMP領證及年費因拿到收據才請款，故結匯時還沒有請款，所以摘要不抓前一筆收款資訊
         If "" & rsA.Fields("CP01") = "P" And ("" & rsA.Fields("CP10") = "601" Or "" & rsA.Fields("CP10") = "605") Then
            Exit Function
         End If
         'end 2023/11/27
         'modify by sonia 2021/7/19 CP60為空者先抓其相關總收文號之CP60,再無才改抓該案母號有收據且收文日最大的收據的國內收款日
         'strCP09 = ""
         StrSQLa = "select c2.cp60,c2.cp09 from caseprogress c1,caseprogress c2 where c1.cp09 = '" & strCP09 & "' and c1.cp43=c2.cp09(+) "
         If rsA.State <> adStateClosed Then rsA.Close
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            If "" & rsA.Fields(0).Value = "" Then
               strCP09 = ""
            Else
               strCP09 = "" & rsA.Fields(1).Value
            End If
         End If
         'end 2021/7/19
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      If strCP09 = "" Then
         StrSQLa = "select SUBSTR(MAX(CP05||cp09), 9) from caseprogress where cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '00' AND SUBSTR(CP60,1,1)='E' "
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            strCP09 = "" & rsA.Fields(0).Value
         End If
         If rsA.State <> adStateClosed Then rsA.Close
      End If
      '2011/8/19 end
      
      'Modified by Morgan 2011/12/23 考慮拆收據情形要抓該收文號的最後收款日及所有收據的總收款金額
      'StrSQLa = "select a0l02, nvl(a0k17, 0)+nvl(a0k18, 0) as Amount from caseprogress, acc0k0, acc0m0, acc0l0, acc1u0 where cp60 = a0k01 and a0k01 = a0m02 and a0m01 = a0l01 and cp60 = a1u02 (+) and cp09 = a1u03 (+) and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "'"
      'StrSQLa = StrSQLa & " union select a0y02 as a0l02, nvl(a1k30, 0) as Amount from caseprogress, acc1k0, acc0z0, acc0y0 where cp60 = a1k01 and a1k01 = a0z02 and a0z01 = a0y01 and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "'"
      'Modified by Morgan 2012/3/27 考慮一收據多案號情形,改抓該收文號的最後收款日及所有收據且相同本所案號的總收款金額 Ex.收文號='AA0053505';另一併扣除退費金額
      'StrSQLa = "select max(a0l02) a0l02, sum(nvl(a0k17, 0)+nvl(a0k18, 0)) as Amount from (select a0j13,max(a0l02) a0l02 from acc0j0,acc0m0,acc0l0 where a0j01 = '" & strCP09 & "' and a0m02(+)=a0j13 and a0l01(+)=a0m01 group by a0j13),acc0k0 where a0k01(+)=a0j13"
      '2012/10/17 modify by sonia 若申請國家為台灣時不抓規費欄(A10100995Y52268000之FCP045645000)
      'StrSQLa = "select max(a0l02) a0l02,sum(nvl(a1u04,0)+nvl(a1u05,0)-nvl(a1u08,0)-nvl(a1u10,0)) Amount" & _
         " from acc0j0 a,acc1u0 c,acc0l0 d where (a0j13,a0j02) in (select b.a0j13,b.a0j02 from acc0j0 b where b.a0j01='" & strCP09 & "') and a1u03(+)=a0j01 and a0l01(+)=a1u01"
      'StrSQLa = StrSQLa & " union select a0y02 as a0l02, nvl(a1k30, 0) as Amount from caseprogress, acc1k0, acc0z0, acc0y0 where cp60 = a1k01(+) and a1k01 = a0z02(+) and a0z01 = a0y01(+)  and cp09 = '" & strCP09 & "' and a1k01 is not null"
      StrSQLa = "select max(a0l02) a0l02,sum(nvl(a1u04,0)+decode(a0j04,'000',0,nvl(a1u05,0))-nvl(a1u08,0)-decode(a0j04,'000',0,nvl(a1u10,0))) Amount" & _
                " from acc0j0 a,acc1u0 c,acc0l0 d where (a0j13,a0j02) in (select b.a0j13,b.a0j02 from acc0j0 b where b.a0j01='" & strCP09 & "') and a1u03(+)=a0j01 and a1u02(+)=a0j13 and a0l01(+)=a1u01"
      StrSQLa = StrSQLa & " union select a0y02 as a0l02, decode(nvl(a1k30,0),0,0,decode('" & GetPrjNation1(Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "-" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "-" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "-" & Mid(strCaseNo, Len(strCaseNo) - 1, 2)) & "','000',nvl(a1k30,0)-nvl(a1k09,0),nvl(a1k30, 0))) as Amount from caseprogress, acc1k0, acc0z0, acc0y0 where cp60 = a1k01(+) and a1k01 = a0z02(+) and a0z01 = a0y01(+)  and cp09 = '" & strCP09 & "' and a1k01 is not null"
      '2012/10/17 end
      
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
      If rsA.RecordCount > 0 Then
          GetA1l02 = "" & rsA.Fields(0).Value
          strDomAmt1 = Format(Val("" & rsA.Fields(1).Value), "#,##0")
          strDomAmt2 = Val("" & rsA.Fields(1).Value)
      End If
      If rsA.State <> adStateClosed Then rsA.Close
      
      'add by sonia 2013/7/22 辜說未收款則抓應收金額,但只抓國內收據不抓國外請款單
      If GetA1l02 = "" Then
         StrSQLa = "select sum(nvl(a0j09,0)+decode(a0j04,'000',0,nvl(a0j10,0))-nvl(a1u07,0)-decode(a0j04,'000',0,nvl(a1u09,0))) Amount" & _
                   " from acc0j0 a,acc1u0 c where (a0j13,a0j02) in (select b.a0j13,b.a0j02 from acc0j0 b where b.a0j01='" & strCP09 & "') and a1u03(+)=a0j01 and a1u02(+)=a0j13"
         rsA.CursorLocation = adUseClient
         rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
         If rsA.RecordCount > 0 Then
            strDomAmt1 = Format(Val("" & rsA.Fields(0).Value), "#,##0")
            strDomAmt2 = Val("" & rsA.Fields(0).Value)
         End If
         If rsA.State <> adStateClosed Then rsA.Close
      End If
      '2013/7/22 end

'Removed by Morgan 2012/3/27 併到上面程式一次抓
'      '2011/12/13 add by sonia 應扣除國內整張收據銷退部分 CFP-024054,但CFP領證費的退公開費不扣除
'      'Modified by Morgan 2011/12/23 考慮拆收據情形
'      'StrSQLa = "select sum(nvl(a1u08,0)+nvl(a1u10,0)) as Amount from caseprogress,acc1u0,acc0s0 " & _
'                "where cp60 = a1u02 (+) and cp01 = '" & Mid(strCaseNo, 1, Len(strCaseNo) - 9) & "' and cp02 = '" & Mid(strCaseNo, Len(strCaseNo) - 8, 6) & "' and cp03 = '" & Mid(strCaseNo, Len(strCaseNo) - 2, 1) & "' and cp04 = '" & Mid(strCaseNo, Len(strCaseNo) - 1, 2) & "' and cp09 = '" & strCP09 & "' " & _
'                "and a1u01=a0s01(+) and a0s24 is null"
'      StrSQLa = "select sum(nvl(a1u08,0)+nvl(a1u10,0)) as Amount from caseprogress,acc1u0,acc0s0 " & _
'                "where cp09 = '" & strCP09 & "' and a1u03(+)=cp09 " & _
'                "and a1u01=a0s01(+) and a0s24 is null"
'      rsA.CursorLocation = adUseClient
'      rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
'      If rsA.RecordCount > 0 Then
'          strDomAmt2 = strDomAmt2 - Val("" & rsA.Fields(0).Value)
'          strDomAmt1 = Format(Val(strDomAmt2), "#,##0")
'      End If
'      If rsA.State <> adStateClosed Then rsA.Close
'      '2011/12/13 end
'      Set rsA = Nothing

   End If
   
End Function

'Add by Morgan 2006/7/21
'設定常用帳戶
Public Sub PUB_SetAccount(ByRef p_Combo1 As Object)
   Dim jj As Integer, stAllNo As String, stName As String, arrAccNo, arrName
   
   p_Combo1.Clear
   'Modify by Amy 2023/05/23 拿掉0149950/0149980/0013472/033169800/0229069/7118730-瑞婷
   stAllNo = "1756680,0202330,1756890,1607750,0236819"
   stName = "智慧瑞興,智慧華銀,法律瑞興,智權瑞興,智權華銀"
   
   stAllNo = stAllNo & ",0172130,1756650,1305688,1911,1912,1913"
   stName = stName & ",林大熙瑞興,智慧瑞興甲存,台銀乙存,中所往來,南所往來,高所往來"
   
   arrAccNo = Split(stAllNo, ",")
   arrName = Split(stName, ",")
   For jj = LBound(arrAccNo) To UBound(arrAccNo)
        p_Combo1.AddItem arrAccNo(jj) & " (" & arrName(jj) & ")"
   Next jj
   Exit Sub
   'end 2023/05/23
   
   'Memo by Amy 2023/05/23 以下不使用
   p_Combo1.AddItem "0149950"     '2010/6/21 ADD BY SONIA 瑞婷
   p_Combo1.AddItem "0149980"
   p_Combo1.AddItem "0202330"
   p_Combo1.AddItem "1911"
   p_Combo1.AddItem "1912"
   p_Combo1.AddItem "1913"
   p_Combo1.AddItem "0013472"
   p_Combo1.AddItem "033169800"
   p_Combo1.AddItem "0229069"      '2010/6/18 ADD BY SONIA 瑞婷
   p_Combo1.AddItem "7118730"      'Add by Morgan 2011/2/16
   p_Combo1.AddItem "1305688"      'Add by Morgan 2011/2/21
   p_Combo1.AddItem "1607750"      '2014/1/10 add by sonia 智權公司
   p_Combo1.AddItem "0236819"      '2015/5/12 add by sonia 智權公司
   p_Combo1.AddItem "0172130"      '2017/7/11 ADD BY SONIA 辜
   p_Combo1.AddItem "1756680"      '2020/4/7 ADD BY SONIA 智慧所
   p_Combo1.AddItem "1756890"      '2020/4/7 ADD BY SONIA 法律所
   p_Combo1.AddItem "1756650"      '2020/6/18 ADD BY SONIA 智慧所
End Sub
'Add by Morgan 2006/11/10
'退費收訖憑單
Public Sub PUB_PrintReceipt3(p_ado0e0 As ADODB.Recordset, p_ado0s0 As ADODB.Recordset, p_Yo As Long, p_PageNo As Long, p_Amount As Long, Optional p_RetNo As String, Optional p_PDate As String)
Dim strCompName As String, strCaseDesc As String, iCount As Integer
   
   With p_ado0s0
      'modify by sonia 2020/4/23 辜說1公司收據銷退自4/1起回執改印智慧所(4/20退X75235)
      'strCompName = A0802Query("" & .Fields("a0k11"))
      strCompName = A0802Query("" & IIf(.Fields("a0k11") = "1", "2", .Fields("a0k11")))
      strCaseDesc = PUB_GetCaseInfo(.Fields("a0k01"))
      If p_PageNo = 0 Then
         p_PageNo = 1
      Else
         If p_Yo > 0 Then
            Printer.NewPage
            p_PageNo = p_PageNo + 1
            p_Yo = 0
         Else
            p_Yo = Printer.Height / 2
         End If
      End If
      
      Printer.FontName = "標楷體"
      Printer.FontSize = 12
      Printer.CurrentX = 200
      Printer.CurrentY = p_Yo + 0
      Printer.Print "請 沿 此 虛 線 撕 下 寄 回"
      
      Printer.FontSize = 24
      Printer.CurrentX = 3800
      Printer.CurrentY = p_Yo + 1000
      Printer.Font.Underline = True
      Printer.Print "退費收訖憑單"
      Printer.Font.Underline = False
      
      Printer.FontSize = 14
      '回執單號
      If p_RetNo = "" Then
         p_RetNo = AutoNo("H", 5)
         strExc(1) = "INSERT INTO ACC250(A2501,A2502,A2503,A2504,A2505,A2506,A2513) VALUES('" & p_RetNo & "','3','" & p_ado0e0("a0q03") & "'," & p_Amount & ",'" & .Fields("a0o01") & "','" & strUserNum & "','" & ChgSQL("" & p_ado0e0("a0q05")) & "')"
         adoTaie.Execute strExc(1)
      End If
      
      Printer.CurrentX = 750
      Printer.CurrentY = p_Yo + 1500
      Printer.Print p_RetNo
      
      Printer.CurrentX = 7500
      Printer.CurrentY = p_Yo + 1500
      If p_PDate <> "" Then
         'Modify by Amy 2014/06/30 +iif 因格式可能已經轉好才帶入
         Printer.Print "日期  " & IIf(InStr(p_PDate, "/") > 0, p_PDate, CFDate(TransDate(p_PDate, 1)))
      Else
         Printer.Print "日期  " & CFDate(strSrvDate(2))
      End If
      
      Printer.CurrentX = 2400
      Printer.CurrentY = p_Yo + 2100
      Printer.Print "茲收到　" & strCompName
   

      Printer.CurrentX = 750
      Printer.CurrentY = p_Yo + 2700
      Printer.Print "新台幣"
      
      '加底線
      Printer.Font.Underline = True
      Printer.CurrentX = 2000
      Printer.CurrentY = p_Yo + 2700
      Printer.Print ChangeNumber(str(p_Amount))
      
      Printer.CurrentX = 7500
      Printer.CurrentY = p_Yo + 2700
      Printer.Print "NTD" & Format(p_Amount, DDollar) 'NT$
      Printer.Font.Underline = False
      
      Printer.CurrentX = 750
      Printer.CurrentY = p_Yo + 3300
      Printer.Print "上述款項係"
      
      '加底線
      strCaseDesc = "退回 " & strCaseDesc & " 款項"
      Printer.Font.Underline = True
      'Modify by Morgan 2011/7/29 考慮超過兩行
      iCount = 0
      Do While strCaseDesc <> ""
         Printer.CurrentX = 750 + Printer.TextWidth(String(6, "　"))
         Printer.CurrentY = p_Yo + 3300 + iCount * 300
         intI = getCutPos(strCaseDesc, Printer.TextWidth(String(25, "　")))
         If intI = 0 Then
            Printer.Print strCaseDesc
            strCaseDesc = ""
         Else
            Printer.Print Left(strCaseDesc, intI)
            iCount = iCount + 1
            strCaseDesc = Mid(strCaseDesc, intI + 1)
         End If
      Loop
      
      Printer.FontSize = 18
      Printer.CurrentX = 4200
      Printer.CurrentY = p_Yo + 5000
      'Modify by Morgan 2008/1/17 改印收據抬頭
      'Printer.Print p_ado0e0("a0q05")
      Printer.Print "" & .Fields("a0k04")
      'end 2008/1/17
      Printer.Font.Underline = False
      Printer.FontSize = 12
      Printer.CurrentX = 750
      Printer.CurrentY = p_Yo + 5500
      Printer.Print "敬請於(退費收訖憑單)上簽蓋　台端(貴公司)之收款章；並退回本所　謝謝！"
      
      Printer.Line (600, p_Yo + 1800)-(600 + 9650, p_Yo + 1800)
      Printer.Line (600, p_Yo + 1800 + 4200)-(600 + 9650, p_Yo + 1800 + 4200)
      Printer.Line (600, p_Yo + 1800)-(600, p_Yo + 1800 + 4200)
      Printer.Line (600 + 9650, p_Yo + 1800)-(600 + 9650, p_Yo + 1800 + 4200)
   
      '2015/8/18 ADD BY SONIA
      Printer.FontSize = 12
      Printer.CurrentX = 6500
      Printer.CurrentY = p_Yo + 6100
      'Modified by Morgan 2022/12/6 --瑞婷
      'Printer.Print "請蓋章後寄回或傳真(02)25068147"
      Printer.Print "請蓋章後寄回或傳真(02)25011666"
      'end 2022/12/6
      '2015/8/18 END
   End With
End Sub

'Add by Morgan 2007/4/12
'退費收訖憑單
'Modify by Morgan 2009/5/8 +p_Choice:2=轉帳同意書
Public Sub PUB_PrintReceipt4(p_adoQuery As ADODB.Recordset, p_Yo As Long, p_PageNo As Long, Optional p_RetNo As String, Optional p_PDate As String, Optional p_Choice As String)
Dim strCompName As String, strAmount As String, strCaseDesc As String, lngX As Long, lngY As Long, lngMax As Long
   
   With p_adoQuery
      strCompName = A0802Query("" & .Fields("a0k11"))
      If p_PageNo = 0 Then
         p_PageNo = 1
      Else
         If p_Yo > 0 Then
            Printer.NewPage
            p_PageNo = p_PageNo + 1
            p_Yo = 0
         Else
            p_Yo = Printer.Height / 2
         End If
      End If
      
      Printer.FontName = "標楷體"
      Printer.FontSize = 12
      Printer.CurrentX = 200
      Printer.CurrentY = p_Yo + 0
      Printer.Print "請 沿 此 虛 線 撕 下 寄 回"
      
      Printer.FontSize = 24
      Printer.CurrentX = 3800
      Printer.CurrentY = p_Yo + 1000
      Printer.Font.Underline = True
      If p_Choice = "6" Then
         Printer.Print "退費轉帳同意書"
      Else
         Printer.Print "退費收訖憑單"
      End If
      Printer.Font.Underline = False
      
      Printer.FontSize = 14
      '回執單號
      If p_RetNo = "" Then
         p_RetNo = AutoNo("H", 5)
         strExc(1) = "INSERT INTO ACC250(A2501,A2502,A2503,A2504,A2505,A2506,A2513,A2514) VALUES('" & p_RetNo & "','4','" & .Fields("a2503") & "'," & Val("" & .Fields("a2504")) & ",'" & .Fields("a2505") & "','" & strUserNum & "','" & ChgSQL(.Fields("a2513")) & "','" & ChgSQL(.Fields("a2514")) & "')"
         adoTaie.Execute strExc(1)
      End If
      
      Printer.CurrentX = 750
      Printer.CurrentY = p_Yo + 1500
      Printer.Print p_RetNo
      
      Printer.CurrentX = 7500
      Printer.CurrentY = p_Yo + 1500
      If p_PDate <> "" Then
         'Modify by Amy 2014/06/30 +iif 因格式可能已經轉好才帶入
         Printer.Print "日期  " & IIf(InStr(p_PDate, "/") > 0, p_PDate, CFDate(TransDate(p_PDate, 1)))
      Else
         Printer.Print "日期  " & CFDate(strSrvDate(2))
      End If
      
      Printer.CurrentX = 2400
      Printer.CurrentY = p_Yo + 2100
      If p_Choice = "2" Then
         Printer.Print "茲同意　" & strCompName
      Else
         Printer.Print "茲收到　" & strCompName
      End If
   

      Printer.CurrentX = 750
      Printer.CurrentY = p_Yo + 2700
      Printer.Print "新台幣"
      
      strAmount = Val("" & .Fields("a2504"))
      '加底線
      Printer.Font.Underline = True
      Printer.CurrentX = 2000
      Printer.CurrentY = p_Yo + 2700
      Printer.Print ChangeNumber(strAmount)
      
      Printer.CurrentX = 7500
      Printer.CurrentY = p_Yo + 2700
      Printer.Print "NTD" & Format(strAmount, DDollar) 'NT$
      Printer.Font.Underline = False
      
      Printer.CurrentX = 750
      Printer.CurrentY = p_Yo + 3300
      Printer.Print "上述款項係"
      
      
      strCaseDesc = "" & .Fields("a2514")
      Printer.Font.Underline = True
      lngX = 750 + Printer.TextWidth(String(6, "　"))
      lngY = p_Yo + 3300
      lngMax = Printer.TextWidth(String(26, "　"))
      intI = getCutPos(strCaseDesc, lngMax)
      Do While intI > 0
         Printer.CurrentX = lngX
         Printer.CurrentY = lngY
         Printer.Print Left(strCaseDesc, intI)
         strCaseDesc = Mid(strCaseDesc, intI + 1)
         intI = getCutPos(strCaseDesc, lngMax)
         lngY = lngY + 300
      Loop
      Printer.CurrentX = lngX
      Printer.CurrentY = lngY
      Printer.Print strCaseDesc
      
      Printer.FontSize = 18
      Printer.CurrentX = 4200
      Printer.CurrentY = p_Yo + 5000
      Printer.Print "" & .Fields("a2513")
      Printer.Font.Underline = False
      Printer.FontSize = 12
      Printer.CurrentX = 750
      Printer.CurrentY = p_Yo + 5500
      Printer.Print "敬請於(退費收訖憑單)上簽蓋　台端(貴公司)之收款章；並退回本所　謝謝！"
      
      Printer.Line (600, p_Yo + 1800)-(600 + 9650, p_Yo + 1800)
      Printer.Line (600, p_Yo + 1800 + 4200)-(600 + 9650, p_Yo + 1800 + 4200)
      Printer.Line (600, p_Yo + 1800)-(600, p_Yo + 1800 + 4200)
      Printer.Line (600 + 9650, p_Yo + 1800)-(600 + 9650, p_Yo + 1800 + 4200)
   End With
   
End Sub

'Add by Morgan 2006/11/10
Private Function getCutPos(p_Desc As String, p_lWidth As Long) As Integer
Dim i As Integer
   
   For i = 1 To Len(p_Desc)
      If Printer.TextWidth(Left(p_Desc, i)) > p_lWidth Then
         getCutPos = i - 1
         Exit For
      End If
   Next
End Function

'Add by Morgan 2006/11/10
'票據受領收據
Public Sub PUB_PrintReceipt1(p_ado0e0 As ADODB.Recordset, p_Yo As Long, p_PageNo As Long, Optional p_RetNo As String, Optional p_PDate As String)
Dim strAmount As String
   
   Printer.Font = "新細明體"
   Printer.FontSize = 12
   If p_PageNo = 0 Then
      p_PageNo = 1
   Else
      If p_Yo > 0 Then
         Printer.NewPage
         p_PageNo = p_PageNo + 1
         p_Yo = 0
      Else
         p_Yo = Printer.Height / 2
      End If
   End If
   With p_ado0e0
      strAmount = "$" & Format(.Fields("a0e11"), DDollar) & "**"
      Printer.CurrentX = 200
      Printer.CurrentY = p_Yo + 0
      Printer.Print ReportSum(50)
      Printer.CurrentX = 200
      Printer.CurrentY = p_Yo + 2500
      Printer.Print A0802Query(.Fields("a1p01")) & ReportSum(43)
      Printer.CurrentX = 1000
      Printer.CurrentY = p_Yo + 4000
      Printer.Print ReportSum(37) & A0g02Query(.Fields("a0e01"))
      Printer.CurrentX = 1000
      Printer.CurrentY = p_Yo + 4300
      Printer.Print ReportSum(38) & .Fields("a0e07")
      Printer.CurrentX = 1000
      Printer.CurrentY = p_Yo + 4600
      Printer.Print ReportSum(39) & .Fields("a0e02")
      Printer.CurrentX = 1000
      Printer.CurrentY = p_Yo + 4900
      Printer.Print ReportSum(40) & CFDate(.Fields("a0e10"))
      
      Printer.CurrentX = 1000
      Printer.CurrentY = p_Yo + 5200
      Printer.Print ReportSum(41) & strAmount
      
      Printer.CurrentX = 1000
      Printer.CurrentY = p_Yo + 6000
      Select Case .Fields("a1p26")
         Case "1"
            Printer.Print ReportSum(42) & ComboItem(111)
         Case "2"
            Printer.Print ReportSum(42) & ComboItem(112)
         Case "3"
            Printer.Print ReportSum(42) & ComboItem(113)
         Case "4"
            Printer.Print ReportSum(42) & ComboItem(114)
         Case "5"
            Printer.Print ReportSum(42) & ComboItem(115)
         Case "6"
            Printer.Print ReportSum(42) & ComboItem(116)
         Case "7"
            'Modify by Morgan 2006/11/2
            'Printer.Print ReportSum(42) & ComboItem(117)
            If Not IsNull(.Fields("a0q18")) Then
               Printer.Print ReportSum(42) & .Fields("a0q18")
            Else
               Printer.Print ReportSum(42) & ComboItem(117)
            End If
      End Select
            
      Printer.FontSize = 14
      Printer.CurrentX = 3800
      Printer.CurrentY = p_Yo + 1300
      Printer.Print ReportTitle(1114)
      Printer.FontSize = 12
      
      '回執單號
      If p_RetNo = "" Then
         p_RetNo = AutoNo("H", 5)
         strExc(1) = "INSERT INTO ACC250(A2501,A2502,A2503,A2504,A2505,A2506,A2513) select '" & p_RetNo & "','1','" & .Fields("a0q03") & "'," & Val("" & .Fields("a0e11")) & ",a0o01,'" & strUserNum & "','" & ChgSQL("" & .Fields("a0q05")) & "' from acc0o0 where a0o03='" & .Fields("a0q03") & "' and a0o11=" & .Fields("a0q01") & " and rownum<2"
         adoTaie.Execute strExc(1)
      End If
      
      Printer.CurrentX = 750
      Printer.CurrentY = p_Yo + 1500
      Printer.Print p_RetNo
            
      Printer.CurrentX = 7500
      Printer.CurrentY = p_Yo + 1500
      If p_PDate <> "" Then
         'Modify by Amy 2014/06/30 +iif 因格式可能已經轉好才帶入
         Printer.Print "日期  " & IIf(InStr(p_PDate, "/") > 0, p_PDate, CFDate(TransDate(p_PDate, 1)))
      Else
         Printer.Print "日期  " & CFDate(strSrvDate(2))
      End If
      
      Printer.Line (200, p_Yo + 1750)-(10000, p_Yo + 1750)
      Printer.CurrentX = 200
      Printer.CurrentY = p_Yo + 2900
      Printer.Print ReportSum(51)
      Printer.CurrentX = 1000
      Printer.CurrentY = p_Yo + 3500
      Printer.Print "    票           據           內            容"
      Printer.CurrentX = 5100
      Printer.CurrentY = p_Yo + 3500
      Printer.Print " 蓋                                      章"
      Printer.Line (700, p_Yo + 3400)-(8500, p_Yo + 3400)
      Printer.Line (700, p_Yo + 3800)-(8500, p_Yo + 3800)
      Printer.Line (700, p_Yo + 5900)-(8500, p_Yo + 5900)
      Printer.Line (700, p_Yo + 6300)-(8500, p_Yo + 6300)
      Printer.Line (700, p_Yo + 3400)-(700, p_Yo + 6300)
      Printer.Line (5000, p_Yo + 3400)-(5000, p_Yo + 5900)
      Printer.Line (8500, p_Yo + 3400)-(8500, p_Yo + 6300)
      Printer.CurrentX = 5100
      Printer.CurrentY = p_Yo + 6500
      Printer.Print "具領人：" & .Fields("a0q05")
   
      '2015/8/19 ADD BY SONIA
      Printer.FontSize = 12
      Printer.CurrentX = 5100
      Printer.CurrentY = p_Yo + 6800
      'Modified by Morgan 2022/12/6 --瑞婷
      'Printer.Print "請蓋章後寄回或傳真(02)25068147"
      Printer.Print "請蓋章後寄回或傳真(02)25011666"
      'end 2022/12/6
      '2015/8/19 END
   End With
   
End Sub
'Add by Morgan 2006/11/10
Public Function PUB_GetCaseInfo(p_A0k01 As String) As String
   
   'Modify by Morgan 2009/6/30 +判斷是否退公開費
   'Modify by Morgan 2011/2/18 +判斷抓進度檔有退費的程序
   'Modified by Morgan 2011/12/26 取消 a0j03,a0j20,a0j21
   'strExc(0) = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04 cno,a0j02,decode(a0s24,'Y','公開費',a0j20) a0j20,a0j21" & _
      " from acc0j0,caseprogress,acc0k0,acc0s0" & _
      " where a0j13='" & p_A0k01 & "' and cp09(+)=a0j01 and cp78>0" & _
      " and a0k01=a0j13 and a0s01(+)=a0k10" & _
      " order by a0j03"
   strExc(0) = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04 cno,a0j02,decode(a0s24,'Y','公開費',getcp10desc(cp01,cp10,a0j04)) CP10N ,na03" & _
      " from acc0j0,caseprogress,acc0k0,acc0s0,nation" & _
      " where a0j13='" & p_A0k01 & "' and cp09(+)=a0j01 and cp78>0" & _
      " and a0k01=a0j13 and a0s01(+)=a0k10 and na01(+)=a0j04" & _
      " order by cp10"
      
   intI = 1
   'edit by nickc 2007/02/07 不用 dll 了
   'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      'Modified by Morgan 2011/12/26 取消 a0j03,a0j20,a0j21
      'PUB_GetCaseInfo = "" & IIf(Right(RsTemp("a0j02"), 3) = "000", Left(RsTemp("a0j02"), Len(RsTemp("a0j02")) - 3), RsTemp("a0j02")) & " " & GetPrjName("" & RsTemp("cno")) & RsTemp("a0j21") & RsTemp("a0j20")
      PUB_GetCaseInfo = "" & IIf(Right(RsTemp("a0j02"), 3) = "000", Left(RsTemp("a0j02"), Len(RsTemp("a0j02")) - 3), RsTemp("a0j02")) & " " & GetPrjName("" & RsTemp("cno")) & RsTemp("na03") & RsTemp("CP10N")
      If RsTemp.RecordCount > 1 Then
         PUB_GetCaseInfo = PUB_GetCaseInfo & "等..."
      End If
   End If
End Function
'Add by Morgan 2006/11/6
'繳款通知書回執
Public Sub PUB_PrintReceipt2(p_adoQuery As ADODB.Recordset, p_Yo As Long, p_PageNo As Long, Optional p_RetNo As String, Optional p_PDate As String)
Dim strAmount As String
   
   Printer.Font = "新細明體"
   Printer.FontSize = 12
   If p_PageNo = 0 Then
      p_PageNo = 1
   Else
      If p_Yo > 0 Then
         Printer.NewPage
         p_PageNo = p_PageNo + 1
         p_Yo = 0
      Else
         p_Yo = Printer.Height / 2
      End If
   End If
   
   With p_adoQuery
      '回執單號
      If p_RetNo = "" Then
         p_RetNo = AutoNo("H", 5)
         strExc(1) = "INSERT INTO ACC250(A2501,A2502,A2503,A2504,A2506,A2512,A2513) values( '" & p_RetNo & "','2','" & .Fields("T0701") & "'," & Val("" & .Fields("T0705")) & ",'" & strUserNum & "'," & Val("" & .Fields("T0706")) & ",'" & ChgSQL(.Fields("T0709")) & "')"
         adoTaie.Execute strExc(1)
      End If
      
      Printer.FontSize = 12
      Printer.CurrentX = 200
      Printer.CurrentY = p_Yo + 0
      Printer.Print ReportSum(50)
      
      Printer.CurrentX = 200
      Printer.CurrentY = p_Yo + 1500
      Printer.Print p_RetNo
      
      Printer.FontSize = 14
      Printer.CurrentX = 3800
      Printer.CurrentY = p_Yo + 1500
      Printer.Print ReportTitle(115)
      
      Printer.FontSize = 12
      Printer.CurrentX = 7500
      Printer.CurrentY = p_Yo + 1500
      If p_PDate <> "" Then
         'Modify by Amy 2014/06/30 +iif 因格式可能已經轉好才帶入
         Printer.Print "日期  " & IIf(InStr(p_PDate, "/") > 0, p_PDate, CFDate(TransDate(p_PDate, 1)))
      Else
         Printer.Print "日期  " & CFDate(strSrvDate(2))
      End If
      
      Printer.Line (200, p_Yo + 2000)-(11500, p_Yo + 2000)
      
      Printer.CurrentX = 200
      Printer.CurrentY = p_Yo + 3000
      Printer.Print A0802Query("1") & ReportSum(43)
      
      Printer.CurrentX = 700
      Printer.CurrentY = p_Yo + 4000
      Printer.Print ReportSum(54)
      
      Printer.CurrentX = 8300
      Printer.CurrentY = p_Yo + 4000
      Printer.Print IIf(IsNull(.Fields("t0706")), 0, .Fields("t0706"))
      
      Printer.CurrentX = 3000
      Printer.CurrentY = p_Yo + 5000
      Printer.Print ReportSum(55)
      
      strAmount = "$" & Format(IIf(IsNull(.Fields("t0705")), 0, .Fields("t0705")), DDollar) & "**"
      Printer.CurrentX = 3000 - Printer.TextWidth(strAmount)
      Printer.CurrentY = p_Yo + 5000
      Printer.Print strAmount
      
      Printer.CurrentX = 5600
      Printer.CurrentY = p_Yo + 6500
      strExc(1) = "簽　收　人："
      Printer.Print strExc(1)
      Pub_SmartPrint "" & .Fields("t0709"), 5600 + Printer.TextWidth(strExc(1)), p_Yo + 6500, 65
      
   End With
End Sub
'Add by Morgan 2007/2/2
'檢查會計科目
'參數 p_AccNo:科目,p_bolMsg:是否彈錯誤訊息,p_Dept:部門
'回傳 0:正確, 1:科目錯, 2:部門錯, 3:智權人員
'Modify by Amy 2021/03/08 +p_A0202 傳票號
Public Function PUB_AccNoGood(p_AccNo As String, p_Dept As String, Optional p_Sales As String, Optional p_bolMsg As Boolean = True, Optional p_AX202 As String = "") As Integer
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
      'modify by sonia 2021/8/23 取消417202，因為有可能是FCT
      'If (p_AccNo >= "410101" And p_AccNo <= "410110" And p_AccNo <> "410109") Or p_AccNo = "417202" Then
      If (p_AccNo >= "410101" And p_AccNo <= "410110" And p_AccNo <> "410109") Then
         strDept = "T"
      ElseIf (p_AccNo >= "411101" And p_AccNo <= "411110" And p_AccNo <> "411106") Then
         strDept = "P"
      ElseIf (p_AccNo >= "4121" And p_AccNo <= "412110") Then
         strDept = "CFT"
      ElseIf (p_AccNo >= "4131" And p_AccNo <= "413110") Then
         strDept = "CFP"
      'cancel by sonia 2020/4/14 F10903266智慧所介紹法律所之法務收入部分列在SAL(婧)
      'ElseIf (p_AccNo >= "414101" And p_AccNo <= "414110") Or (p_AccNo >= "418101" And p_AccNo <= "418110") Then
      '   strDept = "L"
      'ElseIf (p_AccNo >= "416101" And p_AccNo <= "416110") Then
      '   strDept = "FCL"
      'end 2020/4/14
      ElseIf (p_AccNo >= "417101" And p_AccNo <= "417110") Then
         strDept = "FCP"
      'modify by sonia 2021/10/22 417202會有T或FCT故加入 And p_AccNo <> "417202條件
      ElseIf (p_AccNo >= "417201" And p_AccNo <= "417210" And p_AccNo <> "417202") Then
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
         'Add by Amy 2021/03/08 110年後不可使用F4102/03
         ElseIf p_AX202 <> MsgText(601) Then
            If Left(p_AX202, 4) >= "D110" And (p_Sales = "F4102" Or p_Sales = "F4103") Then
                PUB_AccNoGood = 3
                If p_bolMsg = True Then
                   MsgBox "110年以後智權人員不可使用【" & p_Sales & "】！"
                End If
            End If
         'end 2021/03/08
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
      'Add by Amy 2021/03/08 110年後不可使用F4102/03
      ElseIf p_AX202 <> MsgText(601) Then
         If Left(p_AX202, 4) >= "D110" And (p_Sales = "F4102" Or p_Sales = "F4103") Then
            PUB_AccNoGood = 3
            If p_bolMsg = True Then
                MsgBox "110年以後智權人員不可使用【" & p_Sales & "】！"
            End If
         End If
      'end 2021/03/08
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
   Exit Function 'Added by Morga 2020/4/13 法律所要用,取消檢查,改秀玲每日檢查舊收據資料
   
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

'add by sonia 2015/12/31
'依收入抓所屬專業部門
'參數 p_AccNo:科目,p_Dept:部門
Public Function PUB_GETAccNODept(p_AccNo As String, p_Dept As String) As String
   
   PUB_GETAccNODept = p_Dept
   
   If Left(p_AccNo, 1) <> "4" Then
      Exit Function
   End If
   
   If (p_AccNo >= "410101" And p_AccNo <= "410110" And p_AccNo <> "410109") Or p_AccNo = "417202" Then
      PUB_GETAccNODept = "T"
   ElseIf (p_AccNo >= "415101" And p_AccNo <= "415110") Then
      PUB_GETAccNODept = "T"
   ElseIf p_AccNo = "417202" Then
      PUB_GETAccNODept = "T"
   ElseIf (p_AccNo >= "411101" And p_AccNo <= "411110" And p_AccNo <> "411106") Then
      PUB_GETAccNODept = "P"
   ElseIf (p_AccNo >= "4121" And p_AccNo <= "412110") Then
      PUB_GETAccNODept = "CFT"
   ElseIf (p_AccNo >= "4131" And p_AccNo <= "413110") Then
      PUB_GETAccNODept = "CFP"
   ElseIf (p_AccNo >= "414101" And p_AccNo <= "414110") Or (p_AccNo >= "418101" And p_AccNo <= "418110") Then
      PUB_GETAccNODept = "L"
   ElseIf (p_AccNo >= "416101" And p_AccNo <= "416129") Then
      'modify by sonia 2020/5/29
      'PUB_GETAccNODept = "FCL"
      If p_AccNo = "416101" Or p_AccNo = "416111" Or p_AccNo = "416112" Then
         PUB_GETAccNODept = "FCL"
      Else
         PUB_GETAccNODept = "CFL"
      End If
      'end 2020/5/29
   ElseIf (p_AccNo >= "417101" And p_AccNo <= "417110") Then
      PUB_GETAccNODept = "FCP"
   ElseIf (p_AccNo >= "417201" And p_AccNo <= "417210" And p_AccNo <> "417202") Then
      PUB_GETAccNODept = "FCT"
   'add by sonia 2020/1/7
   ElseIf (p_AccNo >= "4201" And p_AccNo <= "420110") Then
      'modify by sonia 2020/10/13
      'PUB_GETAccNODept = "SAL"
      PUB_GETAccNODept = "W"
   Else
      PUB_GETAccNODept = ""
      MsgBox "收入科目 " & p_AccNo & " 未設定部門！請自行補輸，並通知電腦中心修改程式！", vbExclamation
   'end 2020/1/7
   End If
   
End Function
'end 2015/12/31

'Add by Morgan 2007/2/5
'檢查員工狀態
'參數 p_ST01:員工編號,p_ST02:員工名稱
'回傳 0:不存在, 1:正常, 2:離職
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

'Modify by Amy 2022/03/28 +stFormN:表單名/stCmp:公司別/bolHalfSheet:是半張紙
'Modified by Lydia 2024/03/12 +pSaveDoc: 存檔位置
Public Sub PUB_PrintReceipt(ByRef p_Acc250 As ADODB.Recordset, p_lngYo As Long, p_lngPageNo As Long, Optional ByVal stFormN As String = "", Optional ByVal stCmp As String = "", _
      Optional ByVal bolHalfSheet As Boolean = False, Optional ByRef pSaveDoc As String = "")
Dim p_adoquery1 As New ADODB.Recordset
Dim stA2504 As String 'Add by Amy 2022/03/29

   With p_Acc250
      'Modify by Amy 2022/03/29 改共用function開Word畫表格印
      Select Case .Fields("A2502")
         Case "1"
            'Modify by Morgan 2011/9/28 不必限制領款方式--瑞婷
            'strExc(0) = "select a0q18,a0e11,a1p01,a0e01,a0e07,a0e02,a0e10,a1p26,a0q05 from acc0o0, acc0q0, acc1p0, acc0e0 where a0o01='" & .Fields("A2505") & "' and a0q01=a0o11 and a0q03=a0o03 and a1p04 = a0q17 and a1p10 = a0e01 and a1p09 = a0e02  and a1p02 = 'C' AND A1P24 NOT IN ('1', '4')"
            'Modify by Amy 2022/03/29 bug-加 And A1p01=A0e23 And a1p11=a0e07,資料會抓很久不出來  (因 2014/01/27 +公司別/2020/07/06 +a0e07 因改為key,參看frmacc14b0未加到的部分)
            'strExc(0) = "select a0q18,a0e11,a1p01,a0e01,a0e07,a0e02,a0e10,a1p26,a0q05 from acc0o0, acc0q0, acc1p0, acc0e0 where a0o01='" & .Fields("A2505") & "' and a0q01=a0o11 and a0q03=a0o03 and a1p04 = a0q17 and a1p10 = a0e01 and a1p09 = a0e02  and a1p02 = 'C'"
            'Modify by Amy 2022/03/29 +a0q03,a0q01,a0q17
            strExc(0) = "select a0q18,a0e11,a1p01,a0e01,a0e07,a0e02,a0e10,a1p26,a0q05,a0q03,a0q01,a0q17 from acc0o0, acc0q0, acc1p0, acc0e0 where a0o01='" & .Fields("A2505") & "' and a0q01=a0o11 and a0q03=a0o03 and a1p04 = a0q17 And A1p01=A0e23 And a1p11=a0e07 and a1p10 = a0e01 and a1p09 = a0e02  and a1p02 = 'C'"
            intI = 1
            'edit by nickc 2007/02/07 不用 dll 了
            'Set p_adoquery1 = objLawDll.ReadRstMsg(intI, strExc(0))
            Set p_adoquery1 = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'PUB_PrintReceipt1 p_adoquery1, p_lngYo, p_lngPageNo, .Fields("A2501"), .Fields("A2518")
               'Modified by Lydia 2024/03/12 +pSaveDoc
               PUB_PrintReceipt_Doc stFormN, .Fields("A2502"), p_adoquery1, stCmp, .Fields("A2501"), .Fields("A2518"), , True, bolHalfSheet, pSaveDoc
            Else
               MsgBox "無法讀取相關資料！"
            End If
         'Modify by Sindy 2022/4/13 辜說沒有在用,可以刪掉;瑞婷:不用了
'         '2=繳款書
'         Case "2"
'            strExc(0) = "select " & Val("" & .Fields("A2504")) & " as t0705," & Val("" & .Fields("A2512")) & " as t0706,'" & .Fields("A2513") & "' as t0709 from dual"
'            intI = 1
'            'edit by nickc 2007/02/07 不用 dll 了
'            'Set p_adoquery1 = objLawDll.ReadRstMsg(intI, strExc(0))
'            Set p_adoquery1 = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               PUB_PrintReceipt2 p_adoquery1, p_lngYo, p_lngPageNo, .Fields("A2501"), .Fields("A2518")
'            Else
'               MsgBox "無法讀取相關資料！"
'            End If
         Case "3"
            'Modify by Morgan 2011/1/18 不必限制領款方式--瑞婷
            'strExc(0) = "select a0k11,a0k01,a0q05,a0k04 from acc0o0, acc0q0, acc1p0, acc0e0, acc0s0, acc0k0 where a0o01='" & .Fields("A2505") & "' and a0q01=a0o11 and a0q03=a0o03 and a1p04 = a0q17 and a1p10 = a0e01 and a1p09 = a0e02  and a1p02 = 'C' AND A1P24 NOT IN ('1', '4') and a0s01(+)=a0o09 and a0k01(+)=a0s02"
            'Modify by Amy 2022/03/29 +a2504
            stA2504 = Replace(Trim(.Fields("a2504")), ",", "")
            strExc(0) = "select a0k11,a0k01,a0q05,a0k04," & stA2504 & " as a2504 from acc0o0, acc0q0, acc0s0, acc0k0 where a0o01='" & .Fields("A2505") & "' and a0q01(+)=a0o11 and a0q03(+)=a0o03 and a0s01(+)=a0o09 and a0k01(+)=a0s02"
            intI = 1
            'edit by nickc 2007/02/07 不用 dll 了
            'Set p_adoquery1 = objLawDll.ReadRstMsg(intI, strExc(0))
            Set p_adoquery1 = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'PUB_PrintReceipt3 p_adoquery1, p_adoquery1, p_lngYo, p_lngPageNo, .Fields("A2504"), .Fields("A2501"), .Fields("A2518")
               'Modified by Lydia 2024/03/12 +pSaveDoc
               PUB_PrintReceipt_Doc stFormN, .Fields("A2502"), p_adoquery1, stCmp, .Fields("A2501"), .Fields("A2518"), , True, bolHalfSheet, pSaveDoc
            Else
               MsgBox "無法讀取相關資料！"
            End If
         Case "4", "6"
            'Modify By Sindy 2022/4/12 + ,a2503,a2505,a2520
            strExc(0) = "select a0k11,a2504,a2513,a2514,a2503,a2505,a2520 from acc250,acc0s0,acc0k0 where a2501='" & .Fields("A2501") & "' and a2505 in (a0s10,a0s23) and a0k01=a0s02"
            intI = 1
            'edit by nickc 2007/02/07 不用 dll 了
            'Set p_adoquery1 = objLawDll.ReadRstMsg(intI, strExc(0))
            Set p_adoquery1 = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'PUB_PrintReceipt4 p_adoquery1, p_lngYo, p_lngPageNo, .Fields("A2501"), , .Fields("A2502")
               'Modified by Lydia 2024/03/12 +pSaveDoc
               PUB_PrintReceipt_Doc stFormN, .Fields("A2502"), p_adoquery1, stCmp, .Fields("A2501"), .Fields("A2518"), , True, bolHalfSheet, pSaveDoc
            Else
               MsgBox "無法讀取相關資料！"
            End If
         'Add by Morgan 2007/6/6
         Case "5"
            'PUB_PrintReceipt5 .Fields("A2501"), p_lngYo, p_lngPageNo
            'Modified by Lydia 2024/03/12 +pSaveDoc
            PUB_PrintReceipt_Doc stFormN, .Fields("A2502"), p_adoquery1, stCmp, .Fields("A2501"), .Fields("A2518"), , True, bolHalfSheet, pSaveDoc
      End Select
      'end 2022/03/29
   End With
End Sub



'*************************************************
'  會計科目名稱查詢
'
'*************************************************
Public Function A0102Query(InputNo As String) As String
Dim adoacc010 As New ADODB.Recordset
   
   adoacc010.CursorLocation = adUseClient
   adoacc010.Open "select * from acc010 where a0101 = '" & InputNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc010.RecordCount <> 0 Then
      If IsNull(adoacc010.Fields("a0102").Value) Then
         A0102Query = MsgText(601)
      Else
         A0102Query = adoacc010.Fields("a0102").Value
      End If
   Else
      A0102Query = MsgText(601)
   End If
   adoacc010.Close
End Function

'Mark by Amy 2015/09/02 改至basQery 因案件的basFunction也有但有改,此未更新
''Add by Morgan 2008/11/13
''轉員工編號為名稱並加到ListBox
'Public Sub PUB_SetUserList(p_ListBox As ListBox, p_stNums As String)
'Dim arrID, stSQL As String, intR As Integer, rstTmp As ADODB.Recordset
'
'   p_ListBox.Clear
'   If p_stNums <> "" Then
'      stSQL = "select st01,st02 from staff where instr('" & p_stNums & "',st01)>0"
'      intR = 1
'      Set rstTmp = ClsLawReadRstMsg(intR, stSQL)
'      If intR = 1 Then
'         arrID = Split(p_stNums, ",")
'         With rstTmp
'         '照原順序排
'         For intI = UBound(arrID) To LBound(arrID) Step -1
'            .MoveFirst
'            Do While Not .EOF
'               If .Fields("st01") = arrID(intI) Then
'                  p_ListBox.AddItem "" & .Fields(1), 0
'                  p_ListBox.ItemData(0) = .Fields(0) '員工編號
'                  .MoveLast
'               End If
'               .MoveNext
'            Loop
'         Next
'         End With
'      End If
'   End If
'   Set rstTmp = Nothing
'End Sub
'end 2015/09/02

'Add By Cheng 2003/07/22
'2013/6/26 自Frmacc2170及Frmacc21b0抽出來並改為共用
'取得客戶/代理人為電匯或票匯
'Modified by Lydia 2015/10/06 +A1718
'Memo by Lydia 2022/10/05 若規則有變更，請加註文件：\\LINUX\PolyCOM\TaieNew\電腦中心日常工作\結匯-預設匯款方式(a1811和媒體檔).doc
Public Function GetTermOfPayment(strA1501 As String, strCurrency As String, Optional strCompType As String, Optional strA1718 As String) As String
Dim StrSQLa As String, strTmp1 As String
Dim rsA As New ADODB.Recordset
Dim m_A2201 As String   'add by sonia 2013/6/26
Dim strCU10 As String 'Added by Lydia 2015/06/12
Dim m_A2219 As String, m_A2217 As String 'Added by Lydia 2015/07/22
Dim m_Rname As String 'Added by Lydia 2017/09/06 受款銀行名稱
Dim m_A2220 As String 'Added by Lydia 2017/09/20 CNAPS(大陸匯款)
Dim m_A2210 As String 'Added by Lydia 2025/05/23 受款銀行代號種類
Dim m_A2224 As String, m_A2225 As String  'Added by Lydia 2025/07/21 受款人地址城市、受款人地址國家代號

   '預設為票匯
   GetTermOfPayment = "1"
   
'modify by sonia 2013/6/26 付款對象固定改為另一家, 故電匯或票匯也要改抓付款對象
'   StrSQLa = "Select A2212 From Acc150, Acc220 Where A1503=A2201 And A1501='" & strA1501 & "' And A2202='" & strCurrency & "' "
'   'Add by Morgan 2006/12/8
'   StrSQLa = StrSQLa & " union Select A2212 From Acc160, Acc220 Where A1603=A2201 And A1601='" & strA1501 & "' And A2202='" & strCurrency & "' "
'   '2012/5/10 ADD BY SONIA 加國外暫收款退費O10100028
'   StrSQLa = StrSQLa & " union Select A2212 From Acc130, Acc220 Where A1304=A2201 And A1301='" & strA1501 & "' And A2202='" & strCurrency & "' "
'   rsA.CursorLocation = adUseClient
'   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'   '若有資料
'   If rsA.RecordCount > 0 Then
'       GetTermOfPayment = "" & rsA.Fields(0).Value
'   End If
'   If rsA.State <> adStateClosed Then rsA.Close
'   Set rsA = Nothing
   
   m_A2201 = ""
   'modify by sonia 2013/8/8 加入以代理人輸入之其他結匯 Acc170
   StrSQLa = "       Select A1503 From Acc150 Where A1501='" & strA1501 & "' " & _
             " union Select A1603 From Acc160 Where A1601='" & strA1501 & "' " & _
             " union Select A1705 From Acc170 Where A1702='" & strA1501 & "' " & _
             " union Select A1304 From Acc130 Where A1301='" & strA1501 & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   '若有資料
   If rsA.RecordCount > 0 Then
       m_A2201 = "" & rsA.Fields(0).Value
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   
   '除暫收款退費ACC130外(表單編號O)都要考慮付款對象
   If Left(strA1501, 1) <> "O" Then
      Select Case Mid(m_A2201, 1, 6)
         '2014/11/24 modify by sonia 婉莘再加Y54053
         'modify by sonia 2016/4/28 陳經理再加Y54052
         'modify by sonia 2017/8/1  陳經理再加Y54715
         'modify by sonia 2020/1/16 陳經理再加Y55351
         'modify by sonia 2024/8/12 再加Y56014
         'modify by sonia 2025/3/4  再加Y56137
         'modify by sonia 2025/8/29 再加Y56167000及Y56167B10
         Case "Y20908", "Y20915", "Y20919", "Y20929", "Y20934", "Y34282", "Y22247", "Y30249", "Y51368", "Y51523", "Y52243", "Y20076", "Y20339", "Y54053", "Y54052", "Y54715", "Y55351", "Y56014", "Y56137", "Y56167"
            m_A2201 = "Y20076000"
         Case "Y45778", "Y53120", "Y20917", "Y53117", "Y53118", "Y53119", "Y53122", "Y53123", "Y53121", "Y51352"
            m_A2201 = "Y20917000"
         Case "Y49419", "Y53188"
            m_A2201 = "Y53188000"
         'add by sonia 2018/1/26
         Case Else
            If m_A2201 = "Y20026020" Then m_A2201 = "Y20026000"
            'If m_A2201 = "Y20284000" Then m_A2201 = "Y20284010"   'add by sonia 2018/5/21
            If m_A2201 = "Y45878000" Then m_A2201 = "Y55253000"   'add by sonia 2019/5/28 婉莘
            If m_A2201 = "Y51333010" Then m_A2201 = "Y51333000"   'add by sonia 2019/6/24 婉莘
            If m_A2201 = "Y52754020" Then m_A2201 = "Y52754010"   'add by sonia 2022/1/20 婉莘
         'end 2018/1/26
      End Select
   End If
   
   'Modified by Lydia 2015/07/22
   'StrSQLa = "Select A2212 From Acc220 Where A2201='" & m_A2201 & "' And A2202='" & strCurrency & "' "
   'Memo by Lydia 2016/05/11 A2219若有變更,有歷史記錄在案件基本資料維護(dml_log)
   StrSQLa = "Select * From Acc220 Where A2201='" & m_A2201 & "' And A2202='" & strCurrency & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   '若有資料
   If rsA.RecordCount > 0 Then
       'Modified by Lydia 2015/07/22
       'GetTermOfPayment = "" & rsA.Fields(0).Value
       GetTermOfPayment = "" & rsA.Fields("A2212").Value '匯款方式
       m_A2219 = "" & rsA.Fields("A2219").Value '手續費方式
       m_A2217 = "" & rsA.Fields("A2217").Value '受款銀行國籍
       m_Rname = "" & rsA.Fields("A2208") & rsA.Fields("A2209") 'Added by Lydia 2017/09/06 受款銀行名稱
       m_A2220 = "" & rsA.Fields("A2220") 'Added by Lydia 2017/09/15 CNAPS(大陸匯款)
       m_A2210 = "" & rsA.Fields("A2210")   'Added by Lydia 2025/05/23 受款銀行代號種類
       'Added by Lydia 2025/07/21
       m_A2224 = "" & rsA.Fields("A2224") '受款人地址城市
       m_A2225 = "" & rsA.Fields("A2225") '受款人地址國家代號
       'end 2025/07/21
   End If
   
   'Added by Lydia 2015/06/12 台銀結匯若國籍是大陸,並且是付款方式是電匯,改為3.台銀電匯紙本
   '方便財務室直接用非台銀水單列印frmacc2430的"列印票匯",套表列印
   'Modified by Lydia 2017/09/06 新增"華銀結匯媒體",判斷4.華銀電匯紙本與台銀的方式一致
   'If Len(strCompType) > 0 And strCompType <> "J" And GetTermOfPayment = "2" Then
   If Len(strCompType) > 0 And GetTermOfPayment = "2" Then
        If Left(m_A2201, 1) = "X" Then
           strCU10 = GetPrjNationNumber1(m_A2201)
        Else
           strCU10 = GetPrjNationNumber(m_A2201)
        End If
        'Added by Lydia 2015/07/22 匯款地在"白俄羅斯"以及匯款手續費方式為"71:BEN", 不能歸入台銀整批匯款, 需被排除.
        'If strCU10 = "020" Then GetTermOfPayment = "3"
        'Modified by Lydia 2016/04/13 國籍取前3碼
        strCU10 = Left(strCU10, 3)
        'Modified by Lydia 2017/09/06
        'If InStr("020,243", strCU10) > 0 Or (InStr("243", m_A2217) > 0 And Len(m_A2217) > 0) Or m_A2219 = "71:BEN" Then
        If m_A2217 <> "" Then
            strTmp1 = m_A2217
        Else
            strTmp1 = strCU10
        End If
        ' 1.變更為國籍是大陸, 且匯款人銀行資料(OR09=A2208+A2209)有中文 ; 2.匯款地在"白俄羅斯" ; 3.匯款手續費方式為"71:BEN"
        If (strTmp1 = "020" And PUB_CheckStrNEC(m_Rname) = True) Or (strTmp1 = "243") Or UCase(m_A2219) = "71:BEN" Then
        'end 2017/09/06
           GetTermOfPayment = "3"
        End If
        'Added by Lydia 2017/09/15 4.國別為大陸,幣別為RMB,但是沒有CNAPS的資料,改為電匯紙本
        'Remove by Lydia 2020/12/03 自2021.1.1 CNAPS系統不再處理國外匯款：人民幣跨境匯款不需再輸入CNAPS號碼，應使用SWIFT BIC
        '                                          =>不需要特別處理的意思，所以日後將以客戶代理人匯款銀行資料之匯款方式為準。
        'If strTmp1 = "020" And strCurrency = "RMB" And Len(m_A2220) = 0 Then
        '   GetTermOfPayment = "3"
        'End If
        'end 2020/12/03
        
        'Added by Lydia 2019/08/06 越南地區(銀行款銀行國籍)結匯必須紙本結匯
        If Left(m_A2217, 3) = "042" Then
           GetTermOfPayment = "3"
        End If
   End If
   'Added by Lydia 2015/10/06 有代為結匯之客戶編號,一律用台銀電匯紙本
   If Len(strA1718) > 0 And Mid(strA1718, 1, 1) = "X" Then
      GetTermOfPayment = "3"
   End If
   'end 2015/10/06
   
   'Added by Lydia 2025/05/23 配合台銀114年8月新的TXT媒體格式，非Swift Code需要分別輸入銀行地址+城市+國家，將銀行資料不是SwiftCode的代理人，調整“以紙本結匯”
   'Move by Lydia 2025/06/25 從If Left(m_A2217, 3) = "042" Then下方移過來；因為華銀格式也從MT格式(不需城市+國家)改成MX格式，所以也將銀行資料不是SwiftCode的代理人調整"以紙本結匯"
   If InStr(UCase(m_A2210) & ",", "SWIFT CODE") = 0 Then
      GetTermOfPayment = "3"
   End If
   'end 2025/05/23
   'Added by Lydia 2025/07/21 參考華銀提供資訊" 經洽國金部，目前有部份國家在受款人地址、城市、國家為必填，分別為加拿大、澳洲、埃及(台銀:澳洲/紐西蘭/加拿大/南非)"，所以判斷這些國家若缺少資料也調整"以紙本結匯"。
   If Trim(m_A2224) = "" Or Trim(m_A2225) = "" Then   '受款人地址城市、受款人地址國家
      strTmp1 = "加拿大102、澳洲015、紐西蘭016、埃及303、南非301"
      If InStr(strTmp1, strCU10) > 0 Or (Trim(m_A2225) <> "" And InStr(strTmp1, m_A2225) > 0) Then
         GetTermOfPayment = "3"
      End If
   End If
   'end 2025/07/21
   
   'Added by Lydia 2017/09/06 新增"華銀結匯媒體",判斷4.華銀電匯紙本與台銀的方式一致
   If GetTermOfPayment = "3" And strCompType = "J" Then
      GetTermOfPayment = "4"
   End If
   'end 2017/09/06
   
   'Added by Lydia 2022/12/28 設定Y55161固定為紙本結匯：因代理人付款指示特殊，台銀要求紙本結匯
   If m_A2201 = "Y55161000" And strCompType <> "J" Then
      GetTermOfPayment = "1"
   End If
   
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
'2013/6/26 end
End Function

'Add by Amy 2013/12/17 抓取各案件基本檔特殊出名公司,回傳 公司別代號及名稱
'strCompNo:回傳公司代號/intLength:名稱字數限制(0則取全部)
Public Function GetSpecialComp(ByVal strNo1, ByVal strNo2, ByVal strNo3, ByVal strNo4, ByRef strCompNo As String, Optional intLength As Integer = 0) As String
    Dim stSQL As String, intR As Integer
    Dim Rs As ADODB.Recordset
    
    GetSpecialComp = ""
    Select Case strNo1
        Case "CFP", "FCP", "P"   '專利
            stSQL = "Select Decode(PA161,null,'1',Decode(PA161,'T','1',PA161)) From Patent " & _
                       "Where PA01='" & strNo1 & "' And PA02='" & strNo2 & "' And PA03='" & strNo3 & "' And PA04='" & strNo4 & "' "
                       
        Case "CFT", "FCT", "T", "TF"    '商標
            stSQL = "Select Decode(TM130,null,'1',Decode(TM130,'T','1',TM130)) From TradeMark " & _
                         "Where TM01='" & strNo1 & "' And TM02='" & strNo2 & "' And TM03='" & strNo3 & "' And TM04='" & strNo4 & "' "
                       
        Case "CFL", "FCL", "L", "LIN"  '法務
            'Modify by Amy 2020/04/14 原:Decode(LC48,null,'1',Decode(LC48,'T','1',LC48))
            stSQL = "Select 'L' From LawCase " & _
                         "Where LC01='" & strNo1 & "' And LC02='" & strNo2 & "' And LC03='" & strNo3 & "' And LC04='" & strNo4 & "' "
                         
        Case "LA"   '顧問-無特殊出名人資料
                         
        Case Else     '服務
            stSQL = "Select Decode(SP85,null,'1',Decode(SP85,'T','1',SP85)) From ServicePractice " & _
                         "Where SP01='" & strNo1 & "' And SP02='" & strNo2 & "' And SP03='" & strNo3 & "' And SP04='" & strNo4 & "' "
    End Select
   
   'Modify by Amy 2020/04/14 1公司顯示2公司簡稱
   'stSQL = "Select A0801,A0802 From Acc080 Where A0801=( " & stSQL & ") "
   stSQL = "Select A0801,Decode(a0801,'1',BName,a0820) A0820 From Acc080,(Select '1' ANo,a0820 BName From Acc080 Where a0801='2' ) " & _
                "Where A0801=( " & stSQL & ") And A0801=ANo(+) "
   intR = 1
   Set Rs = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
        strCompNo = Rs.Fields("A0801")
        GetSpecialComp = Rs.Fields("A0801") & "-" & IIf(intLength > 0, Left(Rs.Fields("A0820"), intLength), Rs.Fields("A0820"))
   End If
   'end 2020/04/14
End Function
'end 2013/12/17

'Added by Morgan 2014/1/3
'Modified by Morgan 2014/1/8 +科目檢查,科目名稱
'檢查會計科目的公司別是否正確
'Input: pAccCode=科目代碼,pCompNo=公司別
'Output: pAccName=科目名稱
Public Function PUB_CheckCompany(pAccCode As String, Optional pCompNo As String, Optional pAccName As String) As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
      
   stSQL = "select a0102,a0109 from acc010 where a0101='" & pAccCode & "'"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      pAccName = "" & rsQuery("a0102")
      If Not IsNull(rsQuery("a0109")) Then
         If rsQuery("a0109") <> pCompNo Then
            MsgBox "會計科目【" & pAccCode & " " & rsQuery("a0102") & "】之公司別只能為【" & rsQuery("a0109") & "】!!", vbExclamation, MsgText(5)
            Exit Function
         End If
      End If
   Else
      MsgBox "會計科目【" & pAccCode & "】不存在!!", vbExclamation, MsgText(5)
      Exit Function
   End If
   PUB_CheckCompany = True
   Set rsQuery = Nothing
End Function

'Add by Amy 2014/07/16
'判斷傳票資料借方會計科目為1211(進項稅額)
Public Function CheckIs1211(stAx201 As String, stAx202 As String) As Boolean
    Dim strQuery As String, intQ As Integer
    Dim adoquery As New ADODB.Recordset
    strQuery = "Select 1 From Acc021 Where Ax201='" & stAx201 & "' And Ax202='" & stAx202 & "' And Ax205='1211' And Ax206 >0"
    intQ = 1
    Set adoquery = ClsLawReadRstMsg(intQ, strQuery)
    If intQ = 1 Then
        CheckIs1211 = True
    Else
        CheckIs1211 = False
    End If
End Function
'判斷是否有A1p04
Public Function CheckExistA1p04(stA1p01 As String, stA1P22 As String) As Boolean
    Dim strQuery As String, intQ As Integer
    Dim adoquery As New ADODB.Recordset
    strQuery = "Select A1P04 From Acc1p0 Where A1P01='" & stA1p01 & "' And A1P22='" & stA1P22 & "' "
    intQ = 1
    Set adoquery = ClsLawReadRstMsg(intQ, strQuery)
    If intQ = 1 Then
        If IsNull(adoquery.Fields("A1P04")) Then
            CheckExistA1p04 = False
        Else
            CheckExistA1p04 = True
        End If
    Else
        CheckExistA1p04 = False
    End If
End Function
'end 2014/07/16

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
    'Modify by Amy 2024/11/01 +北所工作日
    ElseIf Mid(WorkDay, 1, 6) < Mid(strSrvDate(1), 1, 6) Then
        '作業月<系統月,取該月最大工作日
        intChoose = 2
        strQuery = "Select Max(WD01) From WorkDay Where WD01 Between " & Mid(WorkDay, 1, 6) & "00" & " And " & Mid(WorkDay, 1, 6) & "31 And WD02 is null "
    Else
        '作業月>系統月,取該月第一個工作日
        intChoose = 3
        strQuery = "Select Min(WD01) From WorkDay Where WD01 Between " & Mid(WorkDay, 1, 6) & "00" & " And " & Mid(WorkDay, 1, 6) & "31 And WD02 is null "
    End If
    'end 2024/11/01
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

'判斷是否有A1p22
Public Function CheckExistA1p22(stA1p01 As String, stA1p02 As String, stA1P04 As String) As Boolean
    Dim strQuery As String, intQ As Integer
    Dim adoquery As New ADODB.Recordset
    'Modified by Morgan 2022/6/29 公司別空白時不過濾(收款可能多公司),另增加a1p22條件以避免新增分錄的公司別沒傳票而判斷錯誤
    'strQuery = "Select A1P22 From Acc1p0 Where A1P01='" & stA1p01 & "' And A1P02='" & stA1p02 & "' And A1P04='" & stA1P04 & "' "
    strQuery = "Select A1P22 From Acc1p0 Where " & IIf(stA1p01 <> "", "A1P01='" & stA1p01 & "' And", "") & " A1P02='" & stA1p02 & "' And A1P04='" & stA1P04 & "' and a1p22 is not null"
    'end 2022/6/29
    intQ = 1
    Set adoquery = ClsLawReadRstMsg(intQ, strQuery)
    If intQ = 1 Then
        If IsNull(adoquery.Fields("A1P22")) Then
            CheckExistA1p22 = False
        Else
            CheckExistA1p22 = True
        End If
    Else
        CheckExistA1p22 = False
    End If
End Function
'end 2014/09/26

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

'Added by Lydia 2015/02/24 幣別為主,改為幣別為主
'結匯明細匯總表(印完付款明細草稿後產生的)
Public Sub PUB_ExcelSave2(ByVal FrmType As String, Optional ByVal tmpStr As String, Optional ByVal tmpCo As String)
Dim StrSQLa As String, StrSqlB As String
Dim rsA As New ADODB.Recordset, rsB As New ADODB.Recordset, rsCopy As New ADODB.Recordset
Dim xlsSalesPoint As New Excel.Application
Dim wktmp As New Worksheet
Dim lngCounter As Long, lngLocation As Long
Dim ii As Long
Dim lngMaxCounter As Long, lngMaxCounter2 As Long '最大列數 , 2匯率
Dim lngMaxLocation As Long '最大欄數
Dim midTitle As String, strFName As String
Dim startX As Integer, startY As Integer, sX2 As Integer, sY2 As Integer '第1頁金額資料起始位置,第2頁
Dim idx As Long
Dim startUSD As Integer  '非J公司的USD抬頭起始位置
Dim idJ As Integer 'J公司的資料明細欄數
Dim bolCoJ As Boolean, Tsht1 As String, Tsht2 As String '智權公司用
Dim userSql As String  'accrpt218鎖定使用者
Dim mSq As String
Dim strSavePath As String 'Added by Lydia 2017/09/14 存檔資料夾
Dim AddJ As Integer 'Added by Lydia 2017/09/30 智權用於調整分區整批或紙本的起始列
'Added by Lydia 2017/09/30
Dim strGrp As String, strCurrList As String, pGrp As String
Dim tmpArr As Variant, intA As Integer
Dim yCurr As Integer '幣別抬頭欄(Y列)
Dim yTot As Integer  '合計欄(Y列)
Dim strTmp(1 To 4) As String 'Added by Lydia 2022/12/28

'Modified by Lydia 2015/4/17 台銀媒體清單與水單清單一致
'   If FrmType = "Frmacc24m0" Then '台銀結匯水單
'        intI = 1:   userSql = "select count(*) from accrpt218 where R21801='" & strUserNum & "' "
'        Set RsTemp = ClsLawReadRstMsg(intI, userSql)
'        If RsTemp.Fields(0) = 0 Then
'           userSql = ""
'        Else
'           userSql = "and R21801='" & strUserNum & "' "
'        End If
''Modified by Lydia 2015/03/20 "結匯明細彙總表"是所有資料的彙總,所以不剔除020
'        'Added by Lydia 2015/03/19 台銀水單不接受中文資料,因此排除代理人國別為020
'        '先整理出條件範圍內代理或客戶之國別,暫存TB
''        strExc(10) = "select R21802,R21803,R21805,NA01 From ACCRPT218,acc180,acc190,fagent,nation where R21802=A1801(+) and a1801=a1901(+) and R21808<>'J' " & _
''                     userSql & "and a1803>'Y' and substr(a1803,1,8)=fa01(+) and substr(a1803,9,1)=fa02(+) and fa10=na01(+) group by R21802,R21803,R21805,NA01 " & _
''                     "union select R21802,R21803,R21805,NA01 From ACCRPT218,acc180,acc190,customer,nation where R21802=A1801(+) and a1801=a1901(+) and R21808<>'J' " & _
''                     userSql & "and a1803<'Y' and substr(a1803,1,8)=CU01(+) and substr(a1803,9,1)=CU02(+) and cu10=na01(+) group by R21802,R21803,R21805,NA01 "
''        intI = 1:    Set RsTemp = ClsLawReadRstMsg(intI, strExc(10))
''        Set rsCopy = PUB_CreateRecordset(RsTemp, , , , "AACC_FUN_EXCELSAVE2", mSq)
''        mSq = "and FormName='AACC_FUN_EXCELSAVE2' And ID='" & strUserNum & "' and seqno='" & mSq & "' and A1801=R001 and A1803=R002 and R004<>'020' "
'      'StrSQLa = "Select decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',R21805) " & _
'                "From ACCRPT218,acc180,acc190,RDataFactory where R21802=A1801(+) and a1801=a1901(+) and R21808<>'J' " & userSql & mSq & _
'                "group by decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',R21805) order by 1 "
'      'Modified by Lydia 2015/03/23 以付款單為抓資料基準
'      StrSQLa = "Select decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',R21805) " & _
'                "From ACCRPT218,acc180,acc190 where R21802=A1801 and a1801=a1901(+) and R21808<>'J' " & userSql & _
'                "group by decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',R21805) order by 1 "
'
'      strFName = "結匯明細匯總表"
'   Else
      '避免2次查無資料訊息
      If tmpStr = MsgText(28) Then Exit Sub
      'Modified by Lydia 2017/09/14 變更名稱
      If FrmType = "Frmacc24m0" Then '媒體檔作業
         strFName = "結匯明細匯總表"
      Else
         strFName = "結匯水單匯總表" '紙本作業
      End If
         
      If tmpCo <> "J" Then
         'Modified by Lydia 2015/4/17 +國別判斷
         'StrSQLa = "Select decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903) From acc180,acc190 " & _
                "where A1801=a1901(+) and length(a1811)>0 and a1908 is null " & TmpStr & _
                "group by decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903) order by 1 "
         StrSQLa = "Select decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903) From acc180,acc190,fagent,nation " & _
                   "where A1801=a1901(+) and length(a1811)>0 and a1908 is null and substr(a1803,1,8)=fa01(+) and substr(a1803,9,1)=fa02(+) and fa10=na01(+) "
         StrSQLa = StrSQLa & tmpStr & " group by decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903) "
         StrSQLa = StrSQLa & "Union Select decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903) From acc180,acc190,customer,nation " & _
                   "where A1801=a1901(+) and length(a1811)>0 and a1908 is null and substr(a1803,1,8)=cu01(+) and substr(a1803,9,1)=cu02(+) and cu10=na01(+) "
         StrSQLa = StrSQLa & tmpStr & " group by decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903) order by 1 "
         'Modified by Lydia 2017/09/14
         'If FrmType = "Frmacc24m0" Then '台銀結匯水單
         '   strFName = "結匯明細匯總表"
         'Else
         '   strFName = "水單結匯清單_台一"
         'End If
         strFName = strFName & "_台一"
         'end 2017/09/14
      Else
         'Modified by Lydia 2017/09/30 華銀整批媒體RMB改CNY (a1903=>  DECODE(A1903,'RMB','" & J_RMB & "',A1903) A1903 )
         StrSQLa = "Select DECODE(A1903,'RMB','" & J_RMB & "',A1903) A1903 From acc180,acc190 " & _
                "where A1801=a1901(+) and length(a1811)>0 and a1908 is null " & tmpStr & _
                "group by a1903 order by 1 "
         'Modified by Lydia 2017/09/14
         'strFName = "水單結匯清單_智權": bolCoJ = True
         bolCoJ = True
         strFName = strFName & "_智權"
         'end 2017/09/14
      End If
'   End If
   
'跳過抬頭欄位
If bolCoJ Then
    'Modified by Lydia 2017/09/30 第一頁明細抬頭多加合計和空白一行
    'startX = 1:     startY = 8
    startX = 1:     startY = 10
    sX2 = 1:        sY2 = 2
    idJ = 6
    'Added by Lydia 2017/09/30
    AddJ = 0
    yCurr = 4
    yTot = startY - 4
Else
    'Modified by Lydia 2017/09/27 抬頭空白多加一行
    'startX = 1:     startY = 5
    startX = 1:     startY = 6
End If
      
'Modified by Lydia 2017/09/28 不用+子資料夾
'strSavePath = strExcelPath & "結匯水單匯總表" 'Added by Lydia 2017/09/14
strSavePath = strExcelPath
strGrp = "" 'Added by Lydia 2017/09/30

   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      'Modified by Lydia 2017/09/14 改路徑
      'If Dir(strExcelPath & strFName & ACDate(strSrvDate(1)) & ServerTime & MsgText(43)) = "" Then
      '   If Dir(Mid(strExcelPath, 1, Len(strExcelPath) - 1), vbDirectory) = "" Then
      '      MkDir strExcelPath
      '   End If
      'Else
      '   Kill strExcelPath & strFName & ACDate(strSrvDate(1)) & ServerTime & MsgText(43)
      If Dir(strSavePath & "\" & strFName & ACDate(strSrvDate(1)) & ServerTime & MsgText(43)) = "" Then
         If Dir(strSavePath, vbDirectory) = "" Then
            MkDir strSavePath
         End If
      Else
         Kill strSavePath & "\" & strFName & ACDate(strSrvDate(1)) & ServerTime & MsgText(43)
      'end 2017/09/14
      End If
      'Added by Lydia 2019/02/22 Office2013建立excel檔案的工作表不一定存在,一開始預設工作表數量
      If bolCoJ = True Then
           xlsSalesPoint.SheetsInNewWorkbook = 2
      Else
           xlsSalesPoint.SheetsInNewWorkbook = 1
      End If
      'end 2019/02/22
      xlsSalesPoint.Workbooks.add
      
      Set wktmp = xlsSalesPoint.Worksheets(1)
      wktmp.Range("A1").Value = strFName '抬頭
      If bolCoJ Then
        'Ｊ公司　匯率頁輸入
        Tsht1 = "明細": Tsht2 = "匯率"
        wktmp.Name = Tsht1
        Set wktmp = xlsSalesPoint.Worksheets(2)
        wktmp.Name = Tsht2
        wktmp.Range("a1").Value = "幣別／匯款方式"
        lngCounter = 1
        rsA.MoveFirst
           While Not rsA.EOF
               lngCounter = lngCounter + 1
               wktmp.Cells(lngCounter, 1).Value = "" & rsA.Fields(0).Value
               rsA.MoveNext
               If rsA.EOF Then
                  'Added by Lydia 2018/10/30 華銀調漲手續費
                  'Modified by Lydia 2022/12/28 華銀手續費調整，折扣50元
                  'If strSrvDate(2) >= "1071101" Then
                  '      lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "check": wktmp.Cells(lngCounter, 2).Value = "400"
                  '      lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "wire": wktmp.Cells(lngCounter, 2).Value = "400"
                  '      lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "大陸中文": wktmp.Cells(lngCounter, 2).Value = "600"
                  '      lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "足額到行": wktmp.Cells(lngCounter, 2).Value = "700"
                  'Else
                  ''end 2018/10/30
                  '      lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "check": wktmp.Cells(lngCounter, 2).Value = "300"
                  '      lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "wire": wktmp.Cells(lngCounter, 2).Value = "320"
                  '      lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "大陸中文": wktmp.Cells(lngCounter, 2).Value = "600"
                  '      lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "足額到行": wktmp.Cells(lngCounter, 2).Value = "540"
                  'End If
                  strTmp(1) = "350":   strTmp(2) = "350": strTmp(3) = "550": strTmp(4) = "650"
                  lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "check": wktmp.Cells(lngCounter, 2).Value = strTmp(1)
                  lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "wire": wktmp.Cells(lngCounter, 2).Value = strTmp(2)
                  lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "大陸中文": wktmp.Cells(lngCounter, 2).Value = strTmp(3)
                  lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "足額到行": wktmp.Cells(lngCounter, 2).Value = strTmp(4)
                  'end 2022/12/28
                  
                  'Added by Lydia 2016/08/02 +備註
                  ii = lngCounter
                  lngCounter = lngCounter + 2: wktmp.Cells(lngCounter, 1).Value = "備註："
                  lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "足額到行_ 匯款金額會全額到對方銀行"
                  lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "大陸中文_ 大陸地區的匯款，但僅提供中文抬頭匯款。"
                  lngCounter = lngCounter + 2: wktmp.Cells(lngCounter, 1).Value = "之前大陸地區匯款被要求""英文名稱""必須提供在匯款資訊內，"
                  lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "而提供英文名稱後，就可以比較便宜的手續費方式，讓代理人收到相同的匯款金額。"
                  lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "故，大陸中文匯款方式，改以足額到行匯款。"
                  'end 2016/08/02
                  'Added by Lydia 2019/03/13 +備註
                  lngCounter = lngCounter + 2: wktmp.Cells(lngCounter, 1).Value = "結匯台幣超過20萬, 手續費會增加, 計算公式: 匯款金額* 匯率* 5/10000"
                  wktmp.Range("A" & lngCounter & ":" & "G" & lngCounter + 1).Interior.ColorIndex = 6 '底色-黃色
                  lngCounter = lngCounter + 1: wktmp.Cells(lngCounter, 1).Value = "(最低NT100, 最高NT800)"
                  lngCounter = lngCounter + 2: wktmp.Cells(lngCounter, 1).Value = "結匯超過台幣50萬就需要填寫交易申報書."
                  'end 2019/03/13
               End If
           Wend
        'Modified by Lydia 2016/08/02
        'wktmp.Range(wktmp.Cells(sY2, sX2 + 1), wktmp.Cells(lngCounter - 4, sX2 + 1)).Value = "0"
        wktmp.Range(wktmp.Cells(sY2, sX2 + 1), wktmp.Cells(ii - 4, sX2 + 1)).Value = "0"
         lngMaxCounter2 = lngCounter + 2
      Else
        '非Ｊ公司 幣別欄設定
        lngCounter = startY - 1: lngLocation = startX - 1
        rsA.MoveFirst
           While Not rsA.EOF
               lngLocation = lngLocation + 1
               wktmp.Cells(lngCounter, lngLocation).Value = "" & rsA.Fields(0).Value '列位置 1.抬頭 2.件數小計 3.幣別
               'Added by Lydia 2018/02/07 婉莘要求NTD-USD格子的上方加註"費用外加"
               If "" & rsA.Fields(0).Value = "NTD-USD" Then
                    wktmp.Cells(lngCounter - 2, lngLocation).Value = "費用外加"
               End If
               'end 2018/02/07
               If Trim(rsA.Fields(0).Value) = "USD" Then
                  '美金分:(以新臺幣結購),1.票匯,2.電匯
                 ' lngLocation = lngLocation
                  startUSD = lngLocation
                  wktmp.Cells(lngCounter, lngLocation).Value = "" & Trim(rsA.Fields(0).Value) & "(以新臺幣結購)"
                  lngLocation = lngLocation + 1
                  'Modified by Lydia 2017/09/26 分1,2公司和美金,票匯
                  'wktmp.Cells(lngCounter, lngLocation).Value = "" & Trim(rsA.Fields(0).Value)
                  'lngLocation = lngLocation + 1
                  'wktmp.Cells(lngCounter, lngLocation).Value = "" & Trim(rsA.Fields(0).Value) & "(票)"
                  'Remove by Lydia 2020/09/02  (9/1起法律所L併入2公司,拿掉1公司) 不需要再分 商 專
                  'wktmp.Cells(lngCounter, lngLocation).Value = "" & Trim(rsA.Fields(0).Value) & "21" '商USD
                  'lngLocation = lngLocation + 1
                  'wktmp.Cells(lngCounter, lngLocation).Value = "" & Trim(rsA.Fields(0).Value) & "22" '專USD
                  'lngLocation = lngLocation + 1
                  'wktmp.Cells(lngCounter, lngLocation).Value = "" & Trim(rsA.Fields(0).Value) & "11" '商USD(票)
                  'lngLocation = lngLocation + 1
                  'wktmp.Cells(lngCounter, lngLocation).Value = "" & Trim(rsA.Fields(0).Value) & "12" '專USD(票)
                  ''end 2017/09/26
                  wktmp.Cells(lngCounter, lngLocation).Value = "" & Trim(rsA.Fields(0).Value)
                  lngLocation = lngLocation + 1
                  wktmp.Cells(lngCounter, lngLocation).Value = "" & Trim(rsA.Fields(0).Value) & "(票)"
                  'end 2020/09/02
               End If
               rsA.MoveNext
           Wend
        lngMaxLocation = lngLocation '幣別數＝'最大欄數
        lngMaxCounter = startY - 1   '最大資料列數
      End If
      
'Modified by Lydia 2015/4/17 台銀媒體清單與水單清單一致
'      'Accrpt218 = acc180+acc190
'      'R21802=A1901, R21803=A1803, R21804=A1902, R21805=A1903, R21806=A1904, R21807=A1907, R21808=A1917
'      If FrmType = "Frmacc24m0" Then '台銀結匯水單
'
'         'Accrpt218的R21808=A1917, 與現在的A1917可能不一致
''         StrSQLa = "Select a1901,a1812,R21808 as X101, ' ' as aKind, R21803 as X102,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',R21805) X103,sum(R21806) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1 " & _
''                   "From ACCRPT218,ACC180,ACC190 Where R21802=A1801(+) and R21802=a1901(+) and R21804=a1902(+) and substr(R21803,1,9)=substr(a1803,1,9) and R21805<>'USD' and length(a1811)>0 " & _
''                   "Group By a1901,a1812,R21808, R21803,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',R21805),decode(a1812,'Y','1',decode(a1903,'USD','2','1')) " & _
''                   "Union Select a1901,a1812,R21808 as X101, a1811 as aKind, R21803 as X102,R21805 as X103,sum(R21806) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1 " & _
''                   "From ACCRPT218,ACC180,ACC190 Where R21802=A1801(+) and R21802=a1901(+) and R21804=a1902(+) and substr(R21803,1,9)=substr(a1803,1,9) and R21805='USD' and length(a1811) > 0 " & _
''                   " Group By a1901,a1812,R21808,a1811,R21803,R21805,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) "
'          'Modified by Lydia 2015/03/20 "結匯明細彙總表"是所有資料的彙總,所以不剔除020
' '排除代理人國別為020
''         StrSQLa = "Select a1901,a1812,a1917 as X101, ' ' as aKind, R21803 as X102,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',R21805) X103,sum(R21806) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1 " & _
''                   "From ACCRPT218,ACC180,ACC190,RDataFactory Where R21802=A1801(+) and R21808<>'J' " & userSql & " and R21802=a1901(+) and R21804=a1902(+) and substr(R21803,1,9)=substr(a1803,1,9) and R21805<>'USD' and length(a1811)>0 " & mSq & _
''                   "Group By a1901,a1812,a1917, R21803,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',R21805),decode(a1812,'Y','1',decode(a1903,'USD','2','1')) " & _
''                   "Union Select a1901,a1812,a1917 as X101,a1811 as aKind,R21803 as X102,R21805 as X103,sum(R21806) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1 " & _
''                   "From ACCRPT218,ACC180,ACC190,RDataFactory Where R21802=A1801(+) and R21808<>'J' " & userSql & " and R21802=a1901(+) and R21804=a1902(+) and substr(R21803,1,9)=substr(a1803,1,9) and R21805='USD' and length(a1811)>0 " & mSq & _
''                   " Group By a1901,a1812,a1917,a1811,R21803,R21805,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) "
'         StrSQLa = "Select a1901,a1812,a1917 as X101, ' ' as aKind, R21803 as X102,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',R21805) X103,sum(R21806) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1 " & _
'                   "From ACCRPT218,ACC180,ACC190 Where R21802=A1801(+) and R21808<>'J' " & userSql & " and R21802=a1901(+) and R21804=a1902(+) and substr(R21803,1,9)=substr(a1803,1,9) and R21805<>'USD' and length(a1811)>0 " & _
'                   "Group By a1901,a1812,a1917, R21803,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',R21805),decode(a1812,'Y','1',decode(a1903,'USD','2','1')) " & _
'                   "Union Select a1901,a1812,a1917 as X101, a1811 as aKind, R21803 as X102,R21805 as X103,sum(R21806) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1 " & _
'                   "From ACCRPT218,ACC180,ACC190 Where R21802=A1801(+) and R21808<>'J' " & userSql & " and R21802=a1901(+) and R21804=a1902(+) and substr(R21803,1,9)=substr(a1803,1,9) and R21805='USD' and length(a1811)>0 " & _
'                   " Group By a1901,a1812,a1917,a1811,R21803,R21805,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) "
'
'      Else
      'Modified by Lydia 2015/4/17 +國別判斷
'         StrSQLa = "Select a1901,a1812,a1917 as X101, ' ' as aKind, a1803 as X102,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903) X103,sum(a1904) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1 " & _
'                   "From ACC180,ACC190 Where A1801=a1901(+) and a1903<>'USD' and length(a1811)>0 and a1908 is null" & TmpStr & _
'                   " Group By a1901,a1812,a1917,a1803,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903),decode(a1812,'Y','1',decode(a1903,'USD','2','1')) " & _
'                   "Union Select a1901,a1812,a1917 as X101, a1811 as aKind, a1803 as X102,a1903 as X103,sum(a1904) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1 " & _
'                   "From ACC180,ACC190 Where A1801=a1901(+) and a1903='USD' and length(a1811)>0 and a1908 is null" & TmpStr & _
'                   " Group By a1901,a1812,a1917,a1811, a1803,a1903,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) "
         'Modified by Lydia 2017/03/23 + A1811匯款方式
         'Modified by Lydia 2017/09/30 代入國別和幣別(strexc(3),strexc(4))
         'StrSQLa = "Select a1901,a1812,a1917 as X101, ' ' as aKind, a1803 as X102,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903) X103,sum(a1904) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1,a1811 " & _
                   "From ACC180,ACC190,FAGENT,NATION Where A1801=a1901(+) and a1903<>'USD' and length(a1811)>0 and a1908 is null and substr(a1803,1,8)=fa01(+) and substr(a1803,9,1)=fa02(+) and fa10=na01(+) " & tmpStr & _
                   " Group By a1901,a1812,a1917,a1803,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903),decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1811 "
         'StrSQLa = StrSQLa & "Union Select a1901,a1812,a1917 as X101, ' ' as aKind, a1803 as X102,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903) X103,sum(a1904) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1,a1811 " & _
                   "From ACC180,ACC190,customer,NATION Where A1801=a1901(+) and a1903<>'USD' and length(a1811)>0 and a1908 is null and substr(a1803,1,8)=cu01(+) and substr(a1803,9,1)=cu02(+) and CU10=na01(+) " & tmpStr & _
                   " Group By a1901,a1812,a1917,a1803,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903),decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1811 "
         'StrSQLa = StrSQLa & "Union Select a1901,a1812,a1917 as X101, a1811 as aKind, a1803 as X102,a1903 as X103,sum(a1904) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1,a1811 " & _
                   "From ACC180,ACC190,FAGENT,NATION Where A1801=a1901(+) and a1903='USD' and length(a1811)>0 and a1908 is null and substr(a1803,1,8)=fa01(+) and substr(a1803,9,1)=fa02(+) and fa10=na01(+) " & tmpStr & _
                   " Group By a1901,a1812,a1917,a1811, a1803,a1903,decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1811 "
         'StrSQLa = StrSQLa & "Union Select a1901,a1812,a1917 as X101, a1811 as aKind, a1803 as X102,a1903 as X103,sum(a1904) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1,a1811 " & _
                   "From ACC180,ACC190,customer,NATION Where A1801=a1901(+) and a1903='USD' and length(a1811)>0 and a1908 is null and substr(a1803,1,8)=cu01(+) and substr(a1803,9,1)=cu02(+) and cu10=na01(+) " & tmpStr & _
                   " Group By a1901,a1812,a1917,a1811, a1803,a1903,decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1811 "
         'END 2017/03/20
         If bolCoJ = True Then
            strExc(3) = ",SUBSTR(NA01,1,3) NA01 ,DECODE(A1903,'RMB','" & J_RMB & "',A1903) JCURR "
            strExc(4) = ",SUBSTR(NA01,1,3) ,DECODE(A1903,'RMB','" & J_RMB & "',A1903) "
         Else
            strExc(3) = ",SUBSTR(NA01,1,3) NA01 ,A1903 JCURR "
            strExc(4) = ",SUBSTR(NA01,1,3) ,A1903 "
         End If
         StrSQLa = "Select a1901,a1812,a1917 as X101, ' ' as aKind, a1803 as X102,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903) X103," & _
                   "sum(a1904) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1,a1811 " & strExc(3) & _
                   " From ACC180,ACC190,FAGENT,CUSTOMER,NATION Where A1801=a1901(+) and a1903<>'USD' and length(a1811)>0 and a1908 is null " & _
                   "AND SUBSTR(A1803,1,8)=FA01(+) AND SUBSTR(A1803,9,1)=FA02(+) AND SUBSTR(A1803,1,8)=CU01(+) AND SUBSTR(A1803,9,1)=CU02(+) AND NVL(FA10,CU10)=NA01 " & tmpStr & _
                   " Group By a1901,a1812,a1917,a1803,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903),decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1811 " & strExc(4)
         StrSQLa = StrSQLa & " Union Select a1901,a1812,a1917 as X101, A1811 as aKind, a1803 as X102,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903) X103," & _
                   "sum(a1904) X104,decode(a1812,'Y','1',decode(a1903,'USD','2','1')) flag1,a1811 " & strExc(3) & _
                   " From ACC180,ACC190,FAGENT,CUSTOMER,NATION Where A1801=a1901(+) and a1903='USD' and length(a1811)>0 and a1908 is null " & _
                   "AND SUBSTR(A1803,1,8)=FA01(+) AND SUBSTR(A1803,9,1)=FA02(+) AND SUBSTR(A1803,1,8)=CU01(+) AND SUBSTR(A1803,9,1)=CU02(+) AND NVL(FA10,CU10)=NA01 " & tmpStr & _
                   " Group By a1901,a1812,a1917,a1803,decode(substr(a1902,1,1)||a1903,'ONTD','NTD-USD',a1903),decode(a1812,'Y','1',decode(a1903,'USD','2','1')),a1811 " & strExc(4)
         'END 2017/03/20
'      End If
      '與報表可能有差異,例如:104/03/05~104/03/06 accrpt218 有W10400499,無W10400503(B單號)
         '獨立水單
        'Modified by Lydia 2017/03/23 抓受款行資料國籍
       'StrSqlB = "select x1.X101,x1.akind,x1.x102,x1.x103,x1.x104 as sum1,x1.flag1 from (" & StrSQLa & ") X1 where x1.a1812='Y' " ' and x1.X103='USD' "
       'Modified by Lydia 2017/04/07 改成紙本(非台銀電匯)的匯款排前面 decode(X1.a1811,2,decode(nvl(fa10,cu10),'020',1,0),1) => decode(X1.a1811,2,decode(nvl(fa10,cu10),'020',0,1),0)
       'Modified by Lydia 2017/09/26 PKIND排序 (0-台銀電匯紙本,1-華銀電匯紙本,2-其他),UKIND排序(非美金預設為0,USD的匯款方式排序:1.USD(含電匯紙本)->2.票匯)
       'StrSqlB = "select x1.X101,x1.akind,x1.x102,x1.x103,x1.x104 as sum1,x1.flag1,decode(X1.a1811,2,decode(nvl(fa10,cu10),'020',0,1),0) PKIND " & _
                 "from (" & StrSQLa & ") X1,acc220,fagent,customer where x1.a1812='Y' and x1.x102=a2201(+) and x1.x103=a2202(+) " & _
                 "and substr(x1.x102,1,8)=fa01(+) and substr(x1.x102,9,1)=fa02(+) and substr(x1.x102,1,8)=cu01(+) and substr(x1.x102,9,1)=cu02(+) "
       'Modified by Lydia 2017/09/30 +手續費A2219,華銀幣別JCURR,國別NVL(A2217,X1.NA01) NA01
       'StrSqlB = "select x1.X101,x1.akind,x1.x102,x1.x103,x1.x104 as sum1,x1.flag1,decode(X1.a1811,3,0,4,1,2) PKIND,decode(X1.X103,'USD',DECODE(X1.A1811,1,2,1),0) UKIND,X1.A1811 " & _
                 "from (" & StrSQLa & ") X1,acc220,fagent,customer where x1.a1812='Y' and x1.x102=a2201(+) and x1.x103=a2202(+) " & _
                 "and substr(x1.x102,1,8)=fa01(+) and substr(x1.x102,9,1)=fa02(+) and substr(x1.x102,1,8)=cu01(+) and substr(x1.x102,9,1)=cu02(+) "
       StrSqlB = "select x1.X101,x1.akind,x1.x102,x1.x103,x1.x104 as sum1,x1.flag1,decode(X1.A1811,3,0,4,1,2) PKIND," & _
                 "decode(X1.X103,'USD',DECODE(X1.A1811,1,2,1),0) UKIND,X1.A1811,A2219,X1.JCURR,NVL(A2217,X1.NA01) NA01 " & _
                 "from (" & StrSQLa & ") X1,acc220 where x1.a1812='Y' and x1.x102=a2201(+) and DECODE(X1.X103,'" & J_RMB & "','RMB',X1.X103)=a2202(+) "

         '非獨立水單,合併計算
         '收據公司別=a1917 as X101
       'Modified by Lydia 2016/08/09 改成Union all 合集 (W10501431和W10501432內容相同)
       'StrSqlB = StrSqlB & "union select x2.X101,x2.akind,x2.x102,x2.x103,sum(x2.x104) sum1,x2.flag1 from (" & StrSQLa & ") X2 where x2.a1812 is null " & _
                "group by x2.X101,x2.akind,x2.x102,x2.x103,x2.flag1 "
       'Modified by Lydia 2017/03/20 增加A2219(Acc220),凡手續費為71:OUR集中在幣別群組的前方
       'StrSqlB = StrSqlB & "union all select x2.X101,x2.akind,x2.x102,x2.x103,sum(x2.x104) sum1,x2.flag1 from (" & StrSQLa & ") X2 where x2.a1812 is null " & _
                "group by x2.X101,x2.akind,x2.x102,x2.x103,x2.flag1 "
        'Modified by Lydia 2017/03/23 依代理人或客戶的國籍判斷是否為台銀電匯
        'Modified by Lydia 2017/04/07 改成紙本(非台銀電匯)的匯款排前面 decode(X2.a1811,2,decode(nvl(fa10,cu10),'020',1,0),1) => decode(X2.a1811,2,decode(nvl(fa10,cu10),'020',0,1),0)
       'Modified by Lydia 2017/09/26 PKIND排序 (0-台銀電匯紙本,1-華銀電匯紙本,2-其他),UKIND排序(非美金預設為0,USD的匯款方式排序:1.USD(含電匯紙本)->2.票匯)
       'StrSqlB = StrSqlB & "union all select x2.X101,x2.akind,x2.x102,x2.x103,sum(x2.x104) sum1,x2.flag1,decode(X2.a1811,2,decode(nvl(fa10,cu10),'020',0,1),0) PKIND " & _
                "from (" & StrSQLa & ") X2,acc220,fagent,customer where x2.a1812 is null and x2.x102=a2201(+) and x2.x103=a2202(+) " & _
                "and substr(x2.x102,1,8)=fa01(+) and substr(x2.x102,9,1)=fa02(+) and substr(x2.x102,1,8)=cu01(+) and substr(x2.x102,9,1)=cu02(+) " & _
                "group by x2.X101,x2.akind,x2.x102,x2.x103,x2.flag1,decode(X2.a1811,2,decode(nvl(fa10,cu10),'020',0,1),0) "
       'Modified by Lydia 2017/09/30 +手續費A2219, 華銀幣別X2.JCURR,國別NVL(A2217,X2.NA01) NA01
       'StrSqlB = StrSqlB & "union all select x2.X101,x2.akind,x2.x102,x2.x103,sum(x2.x104) sum1,x2.flag1,decode(X2.a1811,3,0,4,1,2) PKIND,decode(X2.X103,'USD',DECODE(X2.A1811,1,2,1),0) UKIND,X2.A1811 " & _
                "from (" & StrSQLa & ") X2,acc220,fagent,customer where x2.a1812 is null and x2.x102=a2201(+) and x2.x103=a2202(+) " & _
                "and substr(x2.x102,1,8)=fa01(+) and substr(x2.x102,9,1)=fa02(+) and substr(x2.x102,1,8)=cu01(+) and substr(x2.x102,9,1)=cu02(+) " & _
                "group by x2.X101,x2.akind,x2.x102,x2.x103,x2.flag1,decode(X2.a1811,3,0,4,1,2),decode(X2.X103,'USD',DECODE(X2.A1811,1,2,1),0),X2.A1811 "
       StrSqlB = StrSqlB & "union all select x2.X101,x2.akind,x2.x102,x2.x103,sum(x2.x104) sum1,x2.flag1,decode(X2.A1811,3,0,4,1,2) PKIND," & _
                "decode(X2.X103,'USD',DECODE(X2.A1811,1,2,1),0) UKIND,X2.A1811,A2219,X2.JCURR,NVL(A2217,X2.NA01) NA01 " & _
                "from (" & StrSQLa & ") X2,acc220 where x2.a1812 is null and x2.x102=a2201(+) and DECODE(X2.X103,'" & J_RMB & "','RMB',X2.X103)=a2202(+) " & _
                "group by x2.X101,x2.akind,x2.x102,x2.x103,x2.flag1,decode(X2.a1811,3,0,4,1,2)," & _
                "decode(X2.X103,'USD',DECODE(X2.A1811,1,2,1),0),X2.A1811,A2219,X2.JCURR,NVL(A2217,X2.NA01) "
        'Memo by Lydia 2017/09/30 台銀結匯媒體(1040320)的判斷欄位名稱和註解
        'X101 收據公司別
        'Flag1   獨立水單&以新台幣結購
        'aKind   非USD=空白,USD=A1811 (1:票匯 2:電匯)
        'X102 代理人
        'X103 幣別
        'X104 金額
        'JCURR 華銀幣別 'Added by Lydia 2017/09/30
        'end 2017/09/30
        
       If bolCoJ Then '智權
           'Modified by Lydia 2017/09/30 智權分整批結匯和紙本結匯((1.票匯+4.華銀紙本結匯)
           'StrSqlB = StrSqlB & "Order By 4 asc,3 asc,2 desc,1 asc " '排序:幣別、代理人、匯款方式、收據公司別
           StrSqlB = Replace(StrSqlB, "1811,3,0,4,1,2", "1811,2,0,1") '將PKIND改成整批結匯和紙本結匯的排序
           StrSqlB = StrSqlB & "Order By PKIND asc, JCURR asc, X102 asc, A1811 asc" '排序:PKIND排序、華銀幣別、代理人、匯款方式
           'end 2017/09/30
       Else           '非智權
           'Modified by Lydia 2017/03/23 改成台銀電匯集中在幣別群組的前方
           'StrSqlB = StrSqlB & "Order By 4 asc,6 asc,2 desc,1 asc,3 asc " '排序:幣別、獨立水單以台幣結匯、匯款方式、收據公司別、代理人
           'Modified by Lydia 2017/05/08 幣別(付款方式)要優先於是否為紙本
           'StrSqlB = StrSqlB & "Order By 4 asc,6 asc,PKIND asc,2 desc,1 asc,3 asc " '排序:幣別、獨立水單以台幣結匯、是否為台銀電匯、匯款方式、收據公司別、代理人
           'Modified by Lydia 2017/09/26 台銀要分1,2公司的美金和票匯
           'StrSqlB = StrSqlB & "Order By 4 asc,6 asc,2 desc,PKIND asc,1 asc,3 asc "  '排序:幣別、獨立水單以台幣結匯、匯款方式、是否為台銀電匯、收據公司別、代理人
           StrSqlB = StrSqlB & "Order By X103 asc,flag1 asc,UKIND asc,PKIND asc,aKind desc,X101 asc,X102 asc "  '排序:幣別、獨立水單以台幣結匯、UKIND排序、PKIND排序、匯款方式akind(USD才有值)、收據公司別、代理人
       End If
        
      rsB.CursorLocation = adUseClient
      rsB.Open StrSqlB, adoTaie, adOpenStatic, adLockReadOnly
        If bolCoJ Then
          'Ｊ公司　資料輸入
'======================
            Set wktmp = xlsSalesPoint.Worksheets(1)

            'Remove by Lydia 2017/09/30 改到迴圈內
            '幣別    匯率    金額    匯款方式    手續費  應付台幣
            'strExc(3) = "幣別": strExc(4) = "匯率": strExc(5) = "金額": strExc(6) = "匯款方式": strExc(7) = "手續費": strExc(8) = "應付台幣"
            'For ii = 1 To idJ
            '    wktmp.Range(Chr(ii + 64) & startY - 1).Value = strExc(ii + 2)
            'Next ii
            'end 2017/09/30
            
            If rsB.RecordCount > 0 Then
                lngMaxCounter = startY
                lngMaxLocation = startX - 1
                lngCounter = lngMaxCounter - 1
                StrSQLa = ""
                While Not rsB.EOF
                    'Added by Lydia 2017/09/30 顯示欄位抬頭
                    '幣別    匯率    金額    匯款方式    手續費  應付台幣
                    If strGrp <> Trim("" & rsB.Fields("PKIND")) Then
                        strExc(3) = "幣別": strExc(4) = "匯率": strExc(5) = "金額": strExc(6) = "匯款方式": strExc(7) = "手續費": strExc(8) = "應付台幣"
                        For ii = 1 To idJ
                          If ii = 1 Then
                            If strGrp <> "" Then '調整分區整批或紙本的起始列
                                '記錄終止列
                                pGrp = pGrp & "/" & lngMaxCounter & ";"
                                
                                AddJ = lngMaxCounter + 5
                                lngCounter = AddJ - 1
                                
                                '記錄整批或紙本:起始列
                                pGrp = pGrp & Trim(rsB.Fields("PKIND")) & ":" & AddJ
                            Else
                                AddJ = startY
                                
                                '記錄整批或紙本:起始列
                                pGrp = pGrp & Trim(rsB.Fields("PKIND")) & ":" & AddJ
                            End If
                          End If
                          wktmp.Range(Chr(ii + 64) & AddJ - 1).Value = strExc(ii + 2)
                        Next ii
                        strGrp = Trim("" & rsB.Fields("PKIND"))
                    End If
                    'end 2017/09/30
                    
                    'Modified by Lydia 2017/09/30
                    'If midTitle <> rsB.Fields("X103") Then
                    '   midTitle =Trim( "" & rsB.Fields("X103"))
                    If strCurrList = "" Or (strCurrList <> "" And InStr(strCurrList, Trim("" & rsB.Fields("JCURR"))) = 0) Then
                       midTitle = Trim("" & rsB.Fields("JCURR"))
                       strCurrList = strCurrList & midTitle & ","
                    'end 2017/09/30
                       lngMaxLocation = lngMaxLocation + 1
                       'Modified by Lydia 2017/09/30
                       'wktmp.Cells(startY - 4, lngMaxLocation).Value = midTitle
                       wktmp.Cells(yCurr, lngMaxLocation).Value = midTitle
                    End If
                    
                    lngCounter = lngCounter + 1
                    For ii = 1 To idJ
                       strExc(ii) = Chr((startX - 1) + 64 + ii) & lngCounter
                    Next ii
                    'Modified by Lydia 2017/09/30 X103=>JCURR
                    wktmp.Range(strExc(1)).Value = Trim("" & rsB.Fields("JCURR"))  '幣別
                    '匯率=VLOOKUP(A6,匯率!$A$2:$B$10,2,FALSE)
                    wktmp.Range(strExc(2)).Formula = "=VLOOKUP(" & strExc(1) & "," & Tsht2 & "!$" & Chr(sX2 + 64) & "$" & sY2 & ":$" & Chr(sX2 + 65) & "$" & lngMaxCounter2 & ",2,FALSE)"
                    wktmp.Range(strExc(3)).Value = "" & rsB.Fields("sum1")   '金額
                    wktmp.Range(strExc(3)).NumberFormatLocal = "#,##0.00"

                    '匯款方式
                    'Added by Lydia 2017/09/30 手續費=OUR,收款地為美國以外, 帶出"足額到行"
                    If strGrp = "0" Then '華銀:整批
                       'Modified by Lydia 2019/06/25 華銀: 手續費=OUR,帶出"足額到行"
                       'If Trim("" & rsB.Fields("A2219")) = "71:OUR" And Mid("" & rsB.Fields("NA01"), 1, 3) <> "101" Then
                       If Trim("" & rsB.Fields("A2219")) = "71:OUR" Then
                          wktmp.Range(strExc(4)).Value = "足額到行"
                       Else
                          wktmp.Range(strExc(4)).Value = "wire"
                       End If
                    Else
                       wktmp.Range(strExc(4)).Value = "check"
                    End If
                    '手續費=IF(LEN(D6)=0,0,VLOOKUP(D6,匯率!$A$2:$B$10,2,FALSE))
                    wktmp.Range(strExc(5)).Formula = "=IF(LEN(" & strExc(4) & ")=0,0,VLOOKUP(" & strExc(4) & "," & Tsht2 & "!$" & Chr(sX2 + 64) & "$" & sY2 & ":$" & Chr(sX2 + 65) & "$" & lngMaxCounter2 & ",2,FALSE))"
                    '應付台幣
                    wktmp.Range(strExc(6)).Formula = "=ROUND(" & strExc(3) & "*" & strExc(2) & ",0)+" & strExc(5)
                    wktmp.Range(strExc(6)).NumberFormatLocal = "#,##0"
                    If lngCounter > lngMaxCounter Then lngMaxCounter = lngCounter
                    rsB.MoveNext
                Wend
                
                'Modified by Lydia 2017/09/30 處理格式並計算計數和合計
'                '+3行空白列
'                lngMaxCounter = lngMaxCounter + 3
'
'                '幣別合計
'                '=SUMIF($A$6:$A$98,"USD",$C$6:$C$98)
'                idx = startY - 5
'                For ii = startX To lngMaxLocation
'                    wktmp.Range(Chr(ii + 64) & idx).Formula = "=SUMIF($" & Chr(startX + 64) & startY & ":$" & Chr(startX + 64) & lngMaxCounter & "," & _
'                                                                Chr(ii + 64) & idx + 1 & ",$" & Chr(startX + 66) & "$" & startY & ":$" & Chr(startX + 66) & "$" & lngMaxCounter & ")"
'                    wktmp.Range(Chr(ii + 64) & idx).NumberFormatLocal = "#,##0.00"
'                    wktmp.Range(Chr(ii + 64) & idx).HorizontalAlignment = xlCenter
'                Next ii
'                '件數合計
'                ii = startX + 2
'                wktmp.Range(Chr(ii + 64) & startY - 2).Formula = "=COUNT(" & Chr(ii + 64) & startY & ":" & Chr(ii + 64) & lngMaxCounter & ")"
'                wktmp.Range(Chr(ii + 64) & startY - 2).NumberFormatLocal = "#,##0件"
'                '台幣合計
'                wktmp.Range(Chr(idJ + 64) & startY - 2).Formula = "=SUM(" & Chr(idJ + 64) & startY & ":" & Chr(idJ + 64) & lngMaxCounter & ")"
'                wktmp.Range(wktmp.Cells(idx + 1, startX), wktmp.Cells(idx + 1, lngMaxLocation)).HorizontalAlignment = xlCenter
'                wktmp.Range(wktmp.Cells(1, 1), wktmp.Cells(lngMaxCounter + 10, IIf(lngMaxLocation < 6, 8, lngMaxLocation + 2))).ColumnWidth = 10.5
'                '資料抬頭粗體
'                wktmp.Range(wktmp.Cells(startY - 2, 1), wktmp.Cells(startY - 1, IIf(lngMaxLocation < 6, 8, lngMaxLocation + 2))).Font.Bold = True
'                '資料欄位畫格線 (點線)
'                wktmp.Range(wktmp.Cells(startY - 1, startX), wktmp.Cells(lngMaxCounter, idJ)).Borders.LineStyle = xlContinuous
'                wktmp.Range(wktmp.Cells(startY - 1, startX), wktmp.Cells(lngMaxCounter, idJ)).Borders.Weight = xlHairline
'                wktmp.Range(wktmp.Cells(startY - 5, startX), wktmp.Cells(startY - 4, lngMaxLocation)).Borders.LineStyle = xlContinuous
'                wktmp.Range(wktmp.Cells(startY - 5, startX), wktmp.Cells(startY - 4, lngMaxLocation)).Borders.Weight = xlHairline
                pGrp = pGrp & "/" & lngMaxCounter & ";" '記錄終止列
                tmpArr = Empty
                tmpArr = Split(pGrp, ";")
                For intA = 0 To UBound(tmpArr)
                    strExc(1) = Trim(tmpArr(intA))
                    If strExc(1) <> "" Then
                       '記錄整批或紙本:起始列/終止列
                       '分區
                       strExc(5) = Mid(strExc(1), 1, InStr(strExc(1), ":") - 1)
                       strExc(1) = Mid(strExc(1), InStr(strExc(1), ":") + 1)
                       '起始列
                       strExc(6) = Mid(strExc(1), 1, InStr(strExc(1), "/") - 1)
                       '終止列
                       strExc(7) = Mid(strExc(1), InStr(strExc(1), "/") + 1)
                       
                       lngMaxCounter = Val(strExc(7)) + 1 '加一列空白
                        
                       '幣別合計
                       '=SUMIF($A$6:$A$98,"USD",$C$6:$C$98)
                        For ii = startX To lngMaxLocation
                            If wktmp.Range(Chr(startX + 64) & yTot).Value = "" Then '尚未有合計抬頭(第1次)
                                wktmp.Range(Chr(ii + 64) & yCurr - 1).Formula = "=SUMIF($" & Chr(startX + 64) & Val(strExc(6)) & ":$" & Chr(startX + 64) & lngMaxCounter & "," & _
                                                                            Chr(ii + 64) & yCurr & ",$" & Chr(startX + 66) & "$" & Val(strExc(6)) & ":$" & Chr(startX + 66) & "$" & lngMaxCounter & ")"
                                wktmp.Range(Chr(ii + 64) & yCurr - 1).NumberFormatLocal = "#,##0.00"
                                wktmp.Range(Chr(ii + 64) & yCurr - 1).HorizontalAlignment = xlCenter
                            Else '+分區
                                wktmp.Range(Chr(ii + 64) & yCurr - 1).Formula = wktmp.Range(Chr(ii + 64) & yCurr - 1).Formula & " + SUMIF($" & Chr(startX + 64) & Val(strExc(6)) & ":$" & Chr(startX + 64) & lngMaxCounter & "," & _
                                                                            Chr(ii + 64) & yCurr & ",$" & Chr(startX + 66) & "$" & Val(strExc(6)) & ":$" & Chr(startX + 66) & "$" & lngMaxCounter & ")"
                            End If
                        Next ii
                        
                        '件數(分區)
                        ii = startX + 2
                        wktmp.Range(Chr(startX + 64) & Val(strExc(6)) - 2).Value = IIf(strExc(5) = "0", "整批結匯", "紙本結匯")
                        wktmp.Range(Chr(ii + 64) & Val(strExc(6)) - 2).Formula = "=COUNT(" & Chr(ii + 64) & Val(strExc(6)) & ":" & Chr(ii + 64) & lngMaxCounter & ")"
                        wktmp.Range(Chr(ii + 64) & Val(strExc(6)) - 2).NumberFormatLocal = "#,##0件"
                        '台幣(分區)
                        wktmp.Range(Chr(idJ + 64) & Val(strExc(6)) - 2).Formula = "=SUM(" & Chr(idJ + 64) & Val(strExc(6)) & ":" & Chr(idJ + 64) & lngMaxCounter & ")"
                        '合計
                        If wktmp.Range(Chr(startX + 64) & yTot).Value = "" Then
                           wktmp.Range(Chr(startX + 64) & yTot).Value = "合計"
                           wktmp.Range(Chr(startX + 65) & yTot).Formula = "=" & Chr(ii + 64) & Val(strExc(6)) - 2
                           wktmp.Range(Chr(startX + 65) & yTot).NumberFormatLocal = "#,##0件"
                           wktmp.Range(Chr(startX + 66) & yTot).Formula = "=" & Chr(idJ + 64) & Val(strExc(6)) - 2
                           '格式：粗體
                           wktmp.Range(wktmp.Cells(yTot, 1), wktmp.Cells(yTot, IIf(lngMaxLocation < 6, 8, lngMaxLocation + 2))).Font.Bold = True
                           '保留-畫格線
                           'wktmp.Range(wktmp.Cells(yTot, 1), wktmp.Cells(yTot, 3)).Borders.LineStyle = xlContinuous
                           'wktmp.Range(wktmp.Cells(yTot, 1), wktmp.Cells(yTot, 3)).Borders.Weight = xlHairline
                        Else
                           wktmp.Range(Chr(startX + 65) & yTot).Formula = wktmp.Range(Chr(startX + 65) & yTot).Formula & " + " & Chr(ii + 64) & Val(strExc(6)) - 2
                           wktmp.Range(Chr(startX + 66) & yTot).Formula = wktmp.Range(Chr(startX + 66) & yTot).Formula & " + " & Chr(idJ + 64) & Val(strExc(6)) - 2
                        End If
                        '資料抬頭粗體
                        wktmp.Range(wktmp.Cells(Val(strExc(6)) - 2, 1), wktmp.Cells(Val(strExc(6)) - 1, IIf(lngMaxLocation < 6, 8, lngMaxLocation + 2))).Font.Bold = True
                        '資料欄位畫格線 (點線)
                        wktmp.Range(wktmp.Cells(Val(strExc(6)) - 2, startX), wktmp.Cells(lngMaxCounter, idJ)).Borders.LineStyle = xlContinuous
                        wktmp.Range(wktmp.Cells(Val(strExc(6)) - 2, startX), wktmp.Cells(lngMaxCounter, idJ)).Borders.Weight = xlHairline
                    End If
                Next intA
                '幣別-金額小計:置中、畫格線 (點線)
                wktmp.Range(wktmp.Cells(yCurr - 1, startX), wktmp.Cells(yCurr, lngMaxLocation)).HorizontalAlignment = xlCenter
                wktmp.Range(wktmp.Cells(yCurr - 1, startX), wktmp.Cells(yCurr, lngMaxLocation)).Borders.LineStyle = xlContinuous
                wktmp.Range(wktmp.Cells(yCurr - 1, startX), wktmp.Cells(yCurr, lngMaxLocation)).Borders.Weight = xlHairline
                '全部欄寬
                wktmp.Range(wktmp.Cells(1, 1), wktmp.Cells(lngMaxCounter + 10, IIf(lngMaxLocation < 6, 8, lngMaxLocation + 2))).ColumnWidth = 10.5
                'end 2017/09/30
                
                Set wktmp = xlsSalesPoint.Worksheets(2)
                wktmp.Range(wktmp.Cells(1, 1), wktmp.Cells(lngMaxCounter2, 5)).ColumnWidth = 9
                Set wktmp = xlsSalesPoint.Worksheets(1)
            End If
'======================
        Else
          '非Ｊ公司
'---------------------------------------------
          If rsB.RecordCount > 0 Then
             lngCounter = lngMaxCounter
             lngLocation = startX
             StrSQLa = ""
             While Not rsB.EOF
                 If rsB.Fields("X103") <> "USD" Then
                    midTitle = Trim("" & rsB.Fields("X103"))
                 Else
                    '美金分:(以新臺幣結購),1.票匯,2.電匯
                    If rsB.Fields("flag1") = "1" Then
                            midTitle = Trim("" & rsB.Fields("X103")) & "(以新臺幣結購)"
                            'Modified by Lydia 2015/05/05 + 台銀電匯紙本
                    'Modified by Lydia 2017/09/26 遇到3-台銀電匯紙本、4-華銀電匯紙本和5-台銀合併結匯,改為1,2公司電匯
                    'ElseIf rsB.Fields("aKind") = "2" Or rsB.Fields("aKind") = "3" Then
                    '        midTitle =Trim( "" & rsB.Fields("X103")) '電匯
                    'Else
                    '        midTitle =Trim( "" & rsB.Fields("X103")) & "(票)"
                    ElseIf InStr("2,3,4,5", rsB.Fields("aKind")) > 0 Then
                           'Modified by Lydia 2020/09/02 (9/1起法律所L併入2公司,拿掉1公司) 不需要再分 商 專
                           'midTitle = Trim("" & rsB.Fields("X103")) & "2" & Trim(rsB.Fields("X101")) '含商USD、專USD
                           midTitle = Trim("" & rsB.Fields("X103")) '電匯
                    Else
                           'Modified by Lydia 2020/09/02 (9/1起法律所L併入2公司,拿掉1公司) 不需要再分 商 專
                           'midTitle = Trim("" & rsB.Fields("X103")) & "1" & Trim(rsB.Fields("X101"))  '含商USD(票)、專USD(票)
                           midTitle = Trim("" & rsB.Fields("X103")) & "(票)"
                    'end 2017/09/26
                    End If
                 End If
                 
                 '尋找幣別的行位置
                 If StrSQLa <> midTitle Then
                    StrSQLa = midTitle: lngCounter = startY - 1: lngLocation = startX
                    'Modified by Lydia 2017/09/26 + .value
                    Do While midTitle <> wktmp.Cells(lngCounter, lngLocation).Value
                        lngLocation = lngLocation + 1
                        If lngLocation > lngMaxLocation Then Exit Do
                    Loop
                 End If

                lngCounter = lngCounter + 1
                If lngCounter > lngMaxCounter Then lngMaxCounter = lngCounter
                If wktmp.Cells(lngCounter, lngLocation).Value = "" Then '逐筆比對寫入
                   wktmp.Cells(lngCounter, lngLocation).Value = "" & rsB.Fields("sum1")
                Else
                   'Modified by Lydia 2017/09/26
                   'Do While wktmp.Cells(lngCounter, lngLocation) = ""
                   Do While wktmp.Cells(lngCounter, lngLocation).Value <> ""
                   'end 2017/09/26
                       lngCounter = lngCounter + 1
                       If lngCounter > lngMaxCounter Then Exit Do
                   Loop
                   If lngCounter > lngMaxCounter Then lngMaxCounter = lngCounter
                   wktmp.Cells(lngCounter, lngLocation).Value = "" & rsB.Fields("sum1")
                End If
                'Modified by Lydia 2017/03/20 改到最後統一格式
                'wktmp.Range(Chr(lngLocation + 64) & lngCounter).NumberFormatLocal = "#,##0.00"
                'Added by Lydia 2017/03/20 Pkind=0 (字體加粗和斜體)
                'Modified by Lydia 2020/12/03 +票匯
                'If "" & rsB.Fields("PKIND") = "0" Then
                If "" & rsB.Fields("PKIND") = "0" Or "" & rsB.Fields("A1811") = "1" Then
                   wktmp.Range(Chr(lngLocation + 64) & lngCounter).Font.Bold = True
                   wktmp.Range(Chr(lngLocation + 64) & lngCounter).Font.Italic = True
                End If
                'end 2017/03/20
                 rsB.MoveNext
             Wend
             'Added by Lydia 2017/03/20 空白3行,也要設定格式
             wktmp.Range(Chr(Asc("A") + startX - 1) & startY & ":" & Chr(Asc("A") + lngMaxLocation - 1) & lngMaxCounter + 3).NumberFormatLocal = "#,##0.00"
          End If 'If rsB.RecordCount > 0 Then
          
          '加上件數小計(line-2),金額小計,匯率,換算
          lngCounter = startY - 2: lngLocation = startX
          'Modified by Lydia 2015/03/23 將空白列也列入加總
          lngMaxCounter = lngMaxCounter + 3
          For ii = lngLocation To lngMaxLocation
              '件數小計(line-2)
              wktmp.Range(Chr(ii + 64) & lngCounter).Formula = "=COUNT(" & Chr(ii + 64) & startY & ":" & Chr(ii + 64) & lngMaxCounter & ")"
              wktmp.Range(Chr(ii + 64) & lngCounter).NumberFormatLocal = "#,##0件"
              '金額小計
              idx = lngMaxCounter + 1 'idx = lngMaxCounter + 4
              wktmp.Range(Chr(ii + 64) & idx).Formula = "=SUM(" & Chr(ii + 64) & startY & ":" & Chr(ii + 64) & lngMaxCounter & ")"
              wktmp.Range(Chr(ii + 64) & idx).NumberFormatLocal = "#,##0.00"
              wktmp.Range(Chr(ii + 64) & idx).Font.Color = &HFF&  '紅字
              '匯率=>預設代0
              idx = idx + 1:   strExc(0) = "0"
    '          strExc(0) = PUB_GetUSXRate_1(strSrvDate(2), Mid(wktmp.Cells(startY - 1, ii).Value, 1, 3))
    '          If Len(strExc(0)) > 0 And Val(strExc(0)) > 0 Then
               'USD不必輸匯率, 不必計算明細台幣, 不可計入台幣合計
               'Modified by Lydia 2015/04/13 判斷是否有美金結匯
               If ii > startUSD And startUSD > 0 Then
                   wktmp.Range(Chr(ii + 64) & idx + 1).Font.Color = &HFF0000     '藍字
               Else
               '非USD
'                   If ii = lngMaxLocation - 2 Then
'                      wktmp.Range(Chr(ii + 64) & idx).Value = "1"
'                   Else
                   wktmp.Range(Chr(ii + 64) & idx).Value = strExc(0)
'                   End If
                   wktmp.Range(Chr(ii + 64) & idx).NumberFormatLocal = "#,##0.000"
                   '換算
                   wktmp.Range(Chr(ii + 64) & idx + 1).Formula = "=ROUND(IF(" & Chr(ii + 64) & idx & " > 0," & Chr(ii + 64) & idx - 1 & "*" & Chr(ii + 64) & idx & "," & Chr(ii + 64) & idx - 1 & "),0)"
                   wktmp.Range(Chr(ii + 64) & idx + 1).NumberFormatLocal = "#,##0"
                   wktmp.Range(Chr(ii + 64) & idx + 1).Font.Color = &HFF0000     '藍字
               End If
    '          End If
          Next ii
          '左邊title -不需要
'            wktmp.Range("A" & startY - 2).Value = "件數"
'            wktmp.Range("A" & startY - 1).Value = "幣別"
'            wktmp.Range("A" & startY).Value = "明細"
'            idx = lngMaxCounter + 4
'            wktmp.Range("A" & idx).Value = "明細小計": idx = idx + 1
'            wktmp.Range("A" & idx).Value = "匯率": idx = idx + 1
'            wktmp.Range("A" & idx).Value = "明細台幣": idx = idx + 1
'            wktmp.Range("A" & idx).Value = "台幣合計"
'            wktmp.Range("B" & idx).Formula = "=SUM(B" & idx - 1 & ":" & Chr(lngMaxLocation + 64) & idx - 1 & ")"
           '-------------
           
            'Modified by Lydia 2017/09/27 合計和明細間的空行
            'idx = lngMaxCounter + 4
            idx = lngMaxCounter + 5
            
            'USD不必輸匯率, 不必計算明細台幣, 不可計入台幣合計
            wktmp.Range(Chr(startX + 64) & idx).Value = "台幣合計"
            'wktmp.Range(Chr(startX + 65) & idx).Formula = "=SUM(" & Chr(startX + 64) & idx - 1 & ":" & Chr(lngMaxLocation + 64 - 2) & idx - 1 & ")"
            If startUSD = 0 Then
               'Modified by Lydia 2017/09/27
               'wktmp.Range(Chr(startX + 65) & idx).Formula = "=SUM(" & Chr(startX + 64) & idx - 1 & ":" & Chr(lngMaxLocation + 64) & idx - 1 & ")"
               wktmp.Range(Chr(startX + 65) & idx).Formula = "=SUM(" & Chr(startX + 64) & idx - 2 & ":" & Chr(lngMaxLocation + 64) & idx - 2 & ")"
            Else
               'Modified by Lydia 2017/09/27
               'wktmp.Range(Chr(startX + 65) & idx).Formula = "=SUM(" & Chr(startX + 64) & idx - 1 & ":" & Chr(startUSD + 64) & idx - 1 & ")"
               wktmp.Range(Chr(startX + 65) & idx).Formula = "=SUM(" & Chr(startX + 64) & idx - 2 & ":" & Chr(startUSD + 64) & idx - 2 & ")"
            End If
            wktmp.Range(Chr(startX + 65) & idx).NumberFormatLocal = "#,##0.00"
            idx = idx + 1
            
            '美金小計,改USD抬頭
            If startUSD > 0 Then
                'Added by Lydia 2017/09/26 美金/商,美金/專 (小計)
                'Remove by Lydia 2020/09/02 (9/1起法律所L併入2公司,拿掉1公司) 不需要再分 商 專
                'wktmp.Range(Chr(startX + 64) & idx).Value = "美金/商"
                'wktmp.Range(Chr(startX + 65) & idx).Formula = "=SUM(" & Chr(startUSD + 1 + 64) & lngMaxCounter + 1 & ":" & Chr(startUSD + 1 + 64) & lngMaxCounter + 1 & ")" & _
                                                               "+ SUM(" & Chr(startUSD + 3 + 64) & lngMaxCounter + 1 & ":" & Chr(startUSD + 3 + 64) & lngMaxCounter + 1 & ")"
                'idx = idx + 1
                'wktmp.Range(Chr(startX + 64) & idx).Value = "美金/專"
                'wktmp.Range(Chr(startX + 65) & idx).Formula = "=SUM(" & Chr(startUSD + 2 + 64) & lngMaxCounter + 1 & ":" & Chr(startUSD + 2 + 64) & lngMaxCounter + 1 & ")" & _
                                                               "+ SUM(" & Chr(startUSD + 4 + 64) & lngMaxCounter + 1 & ":" & Chr(startUSD + 4 + 64) & lngMaxCounter + 1 & ")"
                'wktmp.Range(Chr(startX + 65) & idx).Borders(xlEdgeBottom).LineStyle = 1
                'idx = idx + 1
                ''end 2017/09/26
                'end 2020/09/02
                wktmp.Range(Chr(startX + 64) & idx).Value = "美金合計"
                'Modified by Lydia 2020/09/02
                'wktmp.Range(Chr(startX + 65) & idx).Formula = "=SUM(" & Chr(startUSD + 1 + 64) & lngMaxCounter + 1 & ":" & Chr(lngMaxLocation + 64) & lngMaxCounter + 1 & ")"
                wktmp.Range(Chr(startX + 65) & idx).Formula = "=" & Chr(startUSD + 1 + 64) & lngMaxCounter + 1 & "+" & Chr(startUSD + 2 + 64) & lngMaxCounter + 1
                
                wktmp.Range(Chr(startX + 65) & idx).NumberFormatLocal = "#,##0.00"
                wktmp.Range(Chr(startUSD + 64) & startY - 1).Value = "USD"
                wktmp.Range(Chr(startUSD + 64) & startY - 3).Value = "(以新臺幣結購)"
                wktmp.Range(Chr(startUSD + 64) & startY - 3).Font.Size = 10
                'Added by Lydia 2017/09/27 縮小字型符合欄寬
                wktmp.Range(Chr(startUSD + 64) & startY - 3).ShrinkToFit = True
                'Added by Lydia 2017/09/26 美金/商,美金/專 ->改抬頭
                'Remove by Lydia 2020/09/02 (9/1起法律所L併入2公司,拿掉1公司) 不需要再分 商 專
                'wktmp.Range(Chr(startUSD + 65) & startY - 1).Value = "USD"
                'wktmp.Range(Chr(startUSD + 65) & startY - 3).Value = "商"
                'wktmp.Range(Chr(startUSD + 66) & startY - 1).Value = "USD"
                'wktmp.Range(Chr(startUSD + 66) & startY - 3).Value = "專"
                'wktmp.Range(Chr(startUSD + 67) & startY - 1).Value = "USD(票)"
                'wktmp.Range(Chr(startUSD + 67) & startY - 3).Value = "商"
                'wktmp.Range(Chr(startUSD + 68) & startY - 1).Value = "USD(票)"
                'wktmp.Range(Chr(startUSD + 68) & startY - 3).Value = "專"
                'end 2020/09/02
                '說明-置中
                wktmp.Range(wktmp.Cells(startY - 3, startX), wktmp.Cells(startY - 3, lngMaxLocation)).HorizontalAlignment = xlCenter
                'end 2017/09/26
            End If
            '合計總件數
            wktmp.Range("C1").Formula = "=SUM(" & Chr(startX + 64) & startY - 2 & ":" & Chr(lngMaxLocation + 64) & startY - 2 & ")"
            wktmp.Range("C1").NumberFormatLocal = "#,##0件"
            '合併表格
'            wktmp.Range(wktmp.Cells(1, 1), wktmp.Cells(1, lngMaxLocation)).Select
'            With xlsSalesPoint.Selection
'                .HorizontalAlignment = xlCenter
'                .VerticalAlignment = xlBottom
'                .WrapText = False
'                .Orientation = 0
'                .AddIndent = False
'                .ShrinkToFit = False
'                .MergeCells = True
'            End With
'            wktmp.Range(wktmp.Cells(1, 1), wktmp.Cells(1, lngMaxLocation)).Select
'            With xlsSalesPoint.Selection
'                .Columns.AutoFit
'            End With
            wktmp.Range(wktmp.Cells(startY - 1, startX), wktmp.Cells(startY - 1, lngMaxLocation)).HorizontalAlignment = xlCenter
           ' wktmp.Range(wktmp.Cells(1, 1), wktmp.Cells(lngMaxCounter + 10, lngMaxLocation)).ColumnWidth = 10.5
            wktmp.Range(wktmp.Cells(1, 1), wktmp.Cells(lngMaxCounter + 7, lngMaxLocation)).ColumnWidth = 10.5
            '資料欄位畫格線 (點線)
            wktmp.Range(wktmp.Cells(startY - 2, startX), wktmp.Cells(lngMaxCounter + 3, lngMaxLocation)).Borders.LineStyle = xlContinuous
            wktmp.Range(wktmp.Cells(startY - 2, startX), wktmp.Cells(lngMaxCounter + 3, lngMaxLocation)).Borders.Weight = xlHairline
            'end 'Modified by Lydia 2015/03/23
        End If
'---------------------------------------------
FileToEnd:
        If FrmType = "Frmacc24m0" Then '台銀結匯水單
           'Modified by Lydia 2017/09/14 改路徑
           'strExc(0) = strExcelPath & strFName & ACDate(strSrvDate(1)) & ServerTime & "(含票匯)" & MsgText(43)
           'Modified by Lydia 2017/09/28
           strExc(0) = strSavePath & IIf(Right(strSavePath, 1) <> "\", "\", "") & strFName & ACDate(strSrvDate(1)) & ServerTime & "(含票匯)" & MsgText(43)
        Else
           'Modified by Lydia 2017/09/14 改路徑
           'strExc(0) = strExcelPath & strFName & ACDate(strSrvDate(1)) & ServerTime & MsgText(43)
           strExc(0) = strSavePath & IIf(Right(strSavePath, 1) <> "\", "\", "") & strFName & ACDate(strSrvDate(1)) & ServerTime & MsgText(43)
        End If
        'Modified by Lydia 2017/09/26 判斷版本
        'xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExc(0)
        If Val(xlsSalesPoint.Version) < 12 Then
            xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExc(0), FileFormat:=-4143
        Else
            xlsSalesPoint.Workbooks(1).SaveAs FileName:=strExc(0), FileFormat:=56
        End If
        'end 2017/09/26
        
        xlsSalesPoint.Workbooks.Close
        xlsSalesPoint.Quit
        'Modified by Lydia 2015/4/17 應婉莘要求,不產生訊息
        'MsgBox "產生" & strExc(0) & " !", vbInformation
        Set xlsSalesPoint = Nothing
        If rsB.State <> adStateClosed Then rsB.Close
        Set rsB = Nothing
        If rsA.State <> adStateClosed Then rsA.Close
        Set rsA = Nothing
   Else
        'Modified by Lydia 2018/02/22 區別訊息
        'MsgBox MsgText(28)
        MsgBox "查無資料，Excel檔無法產生 !!"
   End If  'If rsA.RecordCount > 0 Then

   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing

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

'ADD BY SONIA 2016/1/4 法務收入拆至各部門
'因民國105年起法務收入科目不再使用,所有法務收入要轉其他專業部科目
Public Function InsertLawACC1P0(ByVal strA1P01 As String, ByVal stra1p02 As String, ByVal strA1P03 As String, ByVal strA1P04 As String, ByVal strA1p05 As String, _
                                ByVal strA1P06 As String, ByVal strA1P07 As String, ByVal strA1p08 As String, ByVal strA1P09 As String, ByVal strA1P10 As String, _
                                ByVal strA1P11 As String, ByVal strA1P12 As String, ByVal strA1P13 As String, ByVal strA1p14 As String, ByVal strA1P15 As String, _
                                ByVal strA1P16 As String, ByVal strA1P17 As String, ByVal strA1P18 As String, ByVal strA1P19 As String, ByVal strA1P20 As String, _
                                ByVal strA1P21 As String, ByVal stra1p22 As String, ByVal strA1P23 As String, ByVal strA1P24 As String, ByVal strA1P25 As String, _
                                ByVal strA1P26 As String, ByVal stra1p27 As String, ByVal strA1P30 As String, ByVal strA1P31 As String, Optional ByVal strCP09 As String = "")
Dim strSql As String, strSysKind As String
Dim rstAdo As ADODB.Recordset, iRtn As Integer
Dim strAMT1D As String, strAMT2D As String, m_AMT1D As String  '借方金額 A1P07
Dim strAMT1C As String, strAMT2C As String, m_AMT1C As String  '貸方金額 A1P08
'add by sonia 2018/11/19
Dim strAMT2492 As String                                       '點數保留2192金額
Dim strCP16 As String
'end 2018/11/19
Dim strAccNo2 As String
Dim strSerialNo As String '分錄序次
Dim m_A1P16 As String     '2016/3/2 add by sonia
Dim m_CU10 As String      'add by sonia 2017/8/23
Dim w_A1P16 As String     'add by sonia 2021/1/21
Dim strA1P05_New As String 'Added by Morgan 2022/7/5

   strSysKind = Mid(strA1P17, 1, Len(strA1P17) - 9)
   strAMT1D = Val(strA1P07)
   strAMT1C = Val(strA1p08)
   strAMT2D = 0
   strAMT2C = 0
   strAccNo2 = ""
   If stra1p22 = "null" Then stra1p22 = ""
   If stra1p27 = "null" Then stra1p27 = ""
   
   Select Case strSysKind
      Case "L", "FCL", "LIN", "CFL"
         '讀法務之案件屬性決定科目
         'modify by sonia 2017/8/23 +CU10以便後面非專利商標或是同時與專利商標著作權案件之科目
         strSql = "select LC47,nvl(CU10,fa10) cu10 from lawcase,customer,fagent where lc01 = '" & Mid(strA1P17, 1, Len(strA1P17) - 9) & "' and lc02 = '" & Mid(strA1P17, Len(strA1P17) - 8, 6) & "' and lc03 = '" & Mid(strA1P17, Len(strA1P17) - 2, 1) & "' and lc04 = '" & Mid(strA1P17, Len(strA1P17) - 1, 2) & "' and substr(lc11,1,8)=cu01(+) and substr(lc11,9,1)=cu02(+) and substr(lc22,1,8)=fa01(+) and substr(lc22,9,1)=fa02(+) "
         iRtn = 1
         Set rstAdo = ClsLawReadRstMsg(iRtn, strSql)
         If iRtn = 1 Then
            '僅與專利有關
            If InStr("" & rstAdo.Fields("LC47"), "專利") > 0 And InStr("" & rstAdo.Fields("LC47"), "商標") = 0 And InStr("" & rstAdo.Fields("LC47"), "著作權") = 0 Then
               Select Case strSysKind
                  Case "L"
                     strA1p05 = "411107"
                  Case "FCL", "LIN"
                     strA1p05 = "417103"
                  Case "CFL"
                     'modify by sonia 2020/7/3 回歸CFP收入-法務(F10906203)
                     'strA1p05 = "412102"   'CFL全數轉CFT法務收入
                     strA1p05 = "413102"
               End Select
            '僅與商標或著作權有關
            ElseIf InStr("" & rstAdo.Fields("LC47"), "專利") = 0 And (InStr("" & rstAdo.Fields("LC47"), "商標") > 0 Or InStr("" & rstAdo.Fields("LC47"), "著作權") > 0) Then
               Select Case strSysKind
                  Case "L"
                     strA1p05 = "410110"
                  Case "FCL", "LIN"
                     strA1p05 = "417203"
                  Case "CFL"
                     strA1p05 = "412102"   'CFL全數轉CFT法務收入
               End Select
            '非專利商標或是同時與專利商標著作權有關
            Else
               'modify by sonia 2017/8/23 F10608407之L-5747(外專外商收文以外者,改以申請人國籍決定科目及比例
'               Select Case strSysKind
'                  '依比例分
'                  Case "L"
'                     strA1p05 = "411107"
'                     strAccNo2 = "410110"
'                     strAMT1D = Val(strA1P07) * 60 / 100
'                     strAMT1C = Val(strA1p08) * 60 / 100
'                     strAMT2D = Val(strA1P07) - Val(strAMT1D)
'                     strAMT2C = Val(strA1p08) - Val(strAMT1C)
'                  '依收文人員之部門,非外專外商則依比例分
'                  Case "FCL", "LIN"
'                     strSql = "select cp12 from caseprogress where cp09 = '" & strCP09 & "'"
'                     iRtn = 1
'                     Set rstAdo = ClsLawReadRstMsg(iRtn, strSql)
'                     If iRtn = 1 Then
'                        Select Case Left(rstAdo.Fields("cp12"), 2)
'                           Case "F1"  '外商
'                              strA1p05 = "417203"
'                           Case "F2"  '外專
'                              strA1p05 = "417103"
'                           Case Else
'                              strA1p05 = "417103"
'                              strAccNo2 = "417203"
'                              strAMT1D = Val(strA1P07) * 550 / 700
'                              strAMT1C = Val(strA1p08) * 150 / 700
'                              strAMT2D = Val(strA1P07) - Val(strAMT1D)
'                              strAMT2C = Val(strA1p08) - Val(strAMT1C)
'                        End Select
'                     Else
'                        MsgBox "法務案收文號錯誤 ! " + strA1P17 + "(" + strCP09 + ")", vbInformation
'                     End If
'                  Case "CFL"
'                     strA1p05 = "412102"   'CFL全數轉CFT法務收入
'               End Select
               m_CU10 = "" & rstAdo.Fields("CU10")
               '先判斷收文智權人員之部門
               strSql = "select cp12 from caseprogress where cp09 = '" & strCP09 & "'"
               iRtn = 1
               Set rstAdo = ClsLawReadRstMsg(iRtn, strSql)
               If iRtn = 1 Then
                  Select Case Left(rstAdo.Fields("cp12"), 2)
                     Case "F1"  '外商
                        strA1p05 = "417203"
                     Case "F2"  '外專
                        strA1p05 = "417103"
                     Case Else
                        '再以申請人國籍決定科目及比例
                        If m_CU10 < "010" Then                         'CCP,CCT之法務
                           strA1p05 = "411107"
                           strAccNo2 = "410110"
                           strAMT1D = Val(strA1P07) * 60 / 100
                           strAMT1C = Val(strA1p08) * 60 / 100
                           strAMT2D = Val(strA1P07) - Val(strAMT1D)
                           strAMT2C = Val(strA1p08) - Val(strAMT1C)
                        Else                                           'FCP,FCT之法務
                           strA1p05 = "417103"
                           strAccNo2 = "417203"
                           strAMT1D = Val(strA1P07) * 550 / 700
                           strAMT1C = Val(strA1p08) * 150 / 700
                           strAMT2D = Val(strA1P07) - Val(strAMT1D)
                           strAMT2C = Val(strA1p08) - Val(strAMT1C)
                        End If
                  End Select
               End If
               'end 2017/8/23
            End If
            'add by sonia 2021/1/21 法務案件抓案源介紹人，無案源抓收文智權人員
            w_A1P16 = ""
            If strCP09 <> "" Then
               strSql = "select nvl(substr(los04,1,5),cp13) cp13 from caseprogress,lawofficesource where cp09=los06(+) and cp09= '" & strCP09 & "'"
               iRtn = 1
               Set rstAdo = ClsLawReadRstMsg(iRtn, strSql)
               If iRtn = 1 Then
                  w_A1P16 = rstAdo.Fields("cp13")
               End If
            End If
            'end 2021/1/21
         Else
            MsgBox "法務案號錯誤 ! " + strA1P17, vbInformation
         End If
      Case "LA"
         '顧問聘任CP10='0'時專利1/3,商標1/3,剩餘1/3由專利商標5:5拆,非聘任由專利商標依比例拆
         'modify by sonia 2018/11/19 +CP16,CP53,CP54以判斷顧問聘任簽約年度year
         strSql = "select cp10,cp16,trunc(months_between(to_date(cp54,'yyyy/mm/dd'),to_date(cp53,'yyyy/mm/dd'))/12)+1 year from caseprogress where cp09 = '" & strCP09 & "'"
         iRtn = 1
         Set rstAdo = ClsLawReadRstMsg(iRtn, strSql)
         If iRtn = 1 Then
            '聘任
            If rstAdo.Fields("cp10") = "0" Then
               'add by sonia 2018/11/19  顧問聘任簽約多年僅第一年做收入,其他做2492點數保留
               strCP16 = ""
               If rstAdo.Fields("year") <= 1 Then
                  strAMT2492 = 0
               Else
                  If strA1p08 < Val(rstAdo.Fields("cp16")) Then strCP16 = "Y"   '顧問聘任簽約多年且為部分收款
                  strAMT2492 = (Val(strA1p08) \ rstAdo.Fields("year")) * (rstAdo.Fields("year") - 1)
                  strA1P07 = Val(strA1P07) \ rstAdo.Fields("year")
                  strA1p08 = Val(strA1p08) \ rstAdo.Fields("year")
               End If
               'end 2018/11/19
               'add by sonia 2016/1/20 先insert專利顧問411102及商標顧問410102(取整數)
               'Modified by Morgan 2023/8/16 顧問收款不再分至(其他各項收入)，直接專利/商標各半即可--瑞婷
               'strAMT1D = Val(strA1P07) \ 3
               'strAMT1C = Val(strA1p08) \ 3
               strAMT1D = Val(strA1P07) \ 2
               strAMT1C = Val(strA1p08) \ 2
               'end 2023/8/16
               m_AMT1D = strAMT1D
               m_AMT1C = strAMT1C
               strA1p05 = "411102"
               strA1P06 = PUB_GETAccNODept(strA1p05, strA1P06)
               strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
                  " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strA1P03 & "', '" & strA1P04 & "', '" & strA1p05 & "', '" & strA1P06 & "', '" & strAMT1D & "', '" & strAMT1C & "', '" & strA1P09 & "', '" & strA1P10 & "', '" & strA1P11 & "', '" & strA1P12 & "', '" & strA1P13 & "', '" & strA1p14 & "', '" & strA1P15 & "', " & _
                           "'" & strA1P16 & "', '" & strA1P17 & "', '" & strA1P18 & "', '" & strA1P19 & "', '" & strA1P20 & "', '" & strA1P21 & "', '" & stra1p22 & "', '" & strA1P23 & "', '" & strA1P24 & "', '" & strA1P25 & "', '" & strA1P26 & "', '" & stra1p27 & "', '" & strA1P30 & "', '" & strA1P31 & "')"
               adoTaie.Execute strSql, intI
               
               'Added by Morgan 2023/8/16
               strAMT1D = Val(strA1P07) - strAMT1D
               strAMT1C = Val(strA1p08) - strAMT1C
               
               strA1p05 = "410102"
               strA1P06 = PUB_GETAccNODept(strA1p05, strA1P06)
               strA1P03 = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = '" & stra1p02 & "' and a1p04 = '" & strA1P04 & "'", 3)
               strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
                  " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strA1P03 & "', '" & strA1P04 & "', '" & strA1p05 & "', '" & strA1P06 & "', '" & strAMT1D & "', '" & strAMT1C & "', '" & strA1P09 & "', '" & strA1P10 & "', '" & strA1P11 & "', '" & strA1P12 & "', '" & strA1P13 & "', '" & strA1p14 & "', '" & strA1P15 & "', " & _
                           "'" & strA1P16 & "', '" & strA1P17 & "', '" & strA1P18 & "', '" & strA1P19 & "', '" & strA1P20 & "', '" & strA1P21 & "', '" & stra1p22 & "', '" & strA1P23 & "', '" & strA1P24 & "', '" & strA1P25 & "', '" & strA1P26 & "', '" & stra1p27 & "', '" & strA1P30 & "', '" & strA1P31 & "')"
               adoTaie.Execute strSql, intI
               
               
               'Modified by Morgan 2023/8/16
               'strA1P03 = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = '" & stra1p02 & "' and a1p04 = '" & strA1P04 & "'", 3)
               ''end 2016/1/20
               'strA1p05 = "411107"
               'strAccNo2 = "410110"
               'strAMT1D = (Val(strA1P07) - (m_AMT1D * 2)) \ 2
               'strAMT1C = (Val(strA1p08) - (m_AMT1C * 2)) \ 2
               'strAMT2D = Val(strA1P07) - (m_AMT1D * 2) - Val(strAMT1D)
               'strAMT2C = Val(strA1p08) - (m_AMT1C * 2) - Val(strAMT1C)
               strAMT1D = ""
               strAMT1C = ""
               strAMT2D = ""
               strAMT2C = ""
               'end 2023/8/16
            '非聘任
            Else
               strA1p05 = "411107"
               strAccNo2 = "410110"
               strAMT1D = Val(strA1P07) * 60 / 100
               strAMT1C = Val(strA1p08) * 60 / 100
               strAMT2D = Val(strA1P07) - Val(strAMT1D)
               strAMT2C = Val(strA1p08) - Val(strAMT1C)
            End If
         End If
      Case "P", "PS"
         strA1p05 = "411107"
      Case "CFP", "CPS"
         'modify by sonia 2020/7/3 回歸CFP收入-法務(F10906203)
         'strA1p05 = "412102"   'CFL全數轉CFT法務收入
         strA1p05 = "413102"
      Case "FCP", "FG"
         strA1p05 = "417103"
      Case "FCT", "S"
         strA1p05 = "417203"
      Case "CFT"
         strA1p05 = "412102"
      Case Else   'T*
         strA1p05 = "410110"
   End Select
   
   
If Val(strAMT1D) > 0 Or Val(strAMT1C) > 0 Then 'Added by Morgan 2023/8/16
   
   strA1P06 = PUB_GETAccNODept(strA1p05, strA1P06)
   'modify by sonia 2016/3/2 對沖業務也不可為F4101投資法務
   'strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
      " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strA1P03 & "', '" & strA1P04 & "', '" & strA1p05 & "', '" & strA1P06 & "', '" & strAMT1D & "', '" & strAMT1C & "', '" & strA1P09 & "', '" & strA1P10 & "', '" & strA1P11 & "', '" & strA1P12 & "', '" & strA1P13 & "', '" & strA1p14 & "', '" & strA1P15 & "', " & _
               "'" & strA1P16 & "', '" & strA1P17 & "', '" & strA1P18 & "', '" & strA1P19 & "', '" & strA1P20 & "', '" & strA1P21 & "', '" & stra1p22 & "', '" & strA1P23 & "', '" & strA1P24 & "', '" & strA1P25 & "', '" & strA1P26 & "', '" & stra1p27 & "', '" & strA1P30 & "', '" & strA1P31 & "')"
   m_A1P16 = strA1P16
   w_A1P16 = strA1P16 'Added by Morgan 2021/2/5
   If m_A1P16 = "F4101" Then
      Select Case strA1p05
         Case "417103"             'FCP法務
            'modify by sonia 2021/1/21 F4102改F4104或F4105
            'm_A1P16 = "F4102"
            If w_A1P16 = "" Then w_A1P16 = "F4102"
            'end 2021/1/21
         Case "412102", "417203"   'CFT及FCT法務
            'modify by sonia 2021/1/21 F4103改F4106或F4107
            'm_A1P16 = "F4103"
            If w_A1P16 = "" Then w_A1P16 = "F4103"
            'end 2021/1/21
      End Select
   End If
   
   'modify by sonia 2021/1/21 m_A1P16改用SalesNoToAccSales(w_A1P16, strA1p05, strA1P17)否則會影響下面判斷
   'strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
      " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strA1P03 & "', '" & strA1P04 & "', '" & strA1p05 & "', '" & strA1P06 & "', '" & strAMT1D & "', '" & strAMT1C & "', '" & strA1P09 & "', '" & strA1P10 & "', '" & strA1P11 & "', '" & strA1P12 & "', '" & strA1P13 & "', '" & strA1p14 & "', '" & strA1P15 & "', " & _
               "'" & m_A1P16 & "', '" & strA1P17 & "', '" & strA1P18 & "', '" & strA1P19 & "', '" & strA1P20 & "', '" & strA1P21 & "', '" & stra1p22 & "', '" & strA1P23 & "', '" & strA1P24 & "', '" & strA1P25 & "', '" & strA1P26 & "', '" & stra1p27 & "', '" & strA1P30 & "', '" & strA1P31 & "')"
   'modify by sonia 2021/3/12 SalesNoToAccSales加傳日期
   'Modified by Morgan 2022/7/5
   'strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
      " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strA1P03 & "', '" & strA1P04 & "', '" & stra1p05 & "', '" & strA1P06 & "', '" & strAMT1D & "', '" & strAMT1C & "', '" & strA1P09 & "', '" & strA1P10 & "', '" & strA1P11 & "', '" & strA1P12 & "', '" & strA1P13 & "', '" & strA1p14 & "', '" & strA1P15 & "', " & _
               "'" & SalesNoToAccSales(w_A1P16, stra1p05, strA1P17, strA1P18) & "', '" & strA1P17 & "', '" & strA1P18 & "', '" & strA1P19 & "', '" & strA1P20 & "', '" & strA1P21 & "', '" & stra1p22 & "', '" & strA1P23 & "', '" & strA1P24 & "', '" & strA1P25 & "', '" & strA1P26 & "', '" & stra1p27 & "', '" & strA1P30 & "', '" & strA1P31 & "')"
   strA1P05_New = PUB_ConvAccNo(strA1P01, strA1p05)
   strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
      " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strA1P03 & "', '" & strA1P04 & "', '" & strA1P05_New & "', '" & strA1P06 & "', '" & strAMT1D & "', '" & strAMT1C & "', '" & strA1P09 & "', '" & strA1P10 & "', '" & strA1P11 & "', '" & strA1P12 & "', '" & strA1P13 & "', '" & strA1p14 & "', '" & strA1P15 & "', " & _
               "'" & SalesNoToAccSales(w_A1P16, strA1p05, strA1P17, strA1P18) & "', '" & strA1P17 & "', '" & strA1P18 & "', '" & strA1P19 & "', '" & strA1P20 & "', '" & strA1P21 & "', '" & stra1p22 & "', '" & strA1P23 & "', '" & strA1P24 & "', '" & strA1P25 & "', '" & strA1P26 & "', '" & stra1p27 & "', '" & strA1P30 & "', '" & strA1P31 & "')"
   'end 2022/7/5
   'END 2016/3/2
   adoTaie.Execute strSql, intI

'Added by Morgan 2023/8/16
End If

If Val(strAMT2D) > 0 Or Val(strAMT2C) > 0 Then
'end 2023/8/16

   If strAccNo2 <> "" Then
      strA1P06 = PUB_GETAccNODept(strAccNo2, strA1P06)
      strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = '" & stra1p02 & "' and a1p04 = '" & strA1P04 & "'", 3)
      'modify by sonia 2016/3/2 對沖業務也不可為F4101投資法務
      'strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
         " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strSerialNo & "', '" & strA1P04 & "', '" & strAccNo2 & "', '" & strA1P06 & "', '" & strAMT2D & "', '" & strAMT2C & "', '" & strA1P09 & "', '" & strA1P10 & "', '" & strA1P11 & "', '" & strA1P12 & "', '" & strA1P13 & "', '" & strA1p14 & "', '" & strA1P15 & "', " & _
                  "'" & strA1P16 & "', '" & strA1P17 & "', '" & strA1P18 & "', '" & strA1P19 & "', '" & strA1P20 & "', '" & strA1P21 & "', '" & stra1p22 & "', '" & strA1P23 & "', '" & strA1P24 & "', '" & strA1P25 & "', '" & strA1P26 & "', '" & stra1p27 & "', '" & strA1P30 & "', '" & strA1P31 & "')"
      m_A1P16 = strA1P16
      If m_A1P16 = "F4101" Then
         Select Case strAccNo2
            Case "417103"             'FCP法務
               'modify by sonia 2021/1/21 F4102改F4104或F4105
               'm_A1P16 = "F4102"
               If w_A1P16 = "" Then w_A1P16 = "F4102"
               'modify by sonia 2021/3/12 加傳日期
               m_A1P16 = SalesNoToAccSales(w_A1P16, strAccNo2, strA1P17, strA1P18)
               'end 2021/1/21
            Case "412102", "417203"   'CFT及FCT法務
               'modify by sonia 2021/1/21 F4103改F4106或F4107
               'm_A1P16 = "F4103"
               If w_A1P16 = "" Then w_A1P16 = "F4103"
               'modify by sonia 2021/3/12 加傳日期
               m_A1P16 = SalesNoToAccSales(w_A1P16, strAccNo2, strA1P17, strA1P18)
               'end 2021/1/21
         End Select
      End If
      'Modified by Morgan 2022/7/5
      'strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
         " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strSerialNo & "', '" & strA1P04 & "', '" & strAccNo2 & "', '" & strA1P06 & "', '" & strAMT2D & "', '" & strAMT2C & "', '" & strA1P09 & "', '" & strA1P10 & "', '" & strA1P11 & "', '" & strA1P12 & "', '" & strA1P13 & "', '" & strA1p14 & "', '" & strA1P15 & "', " & _
                  "'" & m_A1P16 & "', '" & strA1P17 & "', '" & strA1P18 & "', '" & strA1P19 & "', '" & strA1P20 & "', '" & strA1P21 & "', '" & stra1p22 & "', '" & strA1P23 & "', '" & strA1P24 & "', '" & strA1P25 & "', '" & strA1P26 & "', '" & stra1p27 & "', '" & strA1P30 & "', '" & strA1P31 & "')"
      strA1P05_New = PUB_ConvAccNo(strA1P01, strAccNo2)
      strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
         " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strSerialNo & "', '" & strA1P04 & "', '" & strA1P05_New & "', '" & strA1P06 & "', '" & strAMT2D & "', '" & strAMT2C & "', '" & strA1P09 & "', '" & strA1P10 & "', '" & strA1P11 & "', '" & strA1P12 & "', '" & strA1P13 & "', '" & strA1p14 & "', '" & strA1P15 & "', " & _
                  "'" & m_A1P16 & "', '" & strA1P17 & "', '" & strA1P18 & "', '" & strA1P19 & "', '" & strA1P20 & "', '" & strA1P21 & "', '" & stra1p22 & "', '" & strA1P23 & "', '" & strA1P24 & "', '" & strA1P25 & "', '" & strA1P26 & "', '" & stra1p27 & "', '" & strA1P30 & "', '" & strA1P31 & "')"
   'END 2016/3/2
      adoTaie.Execute strSql, intI
   End If
   
End If 'Added by Morgan 2023/8/16

   'add by sonia 2018/11/19 '顧問聘任簽約多年僅第一年做收入,其他做2492點數保留(部門TOT)
   If Val(strAMT2492) > 0 Then
      strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = '" & stra1p02 & "' and a1p04 = '" & strA1P04 & "'", 3)
      strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
         " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strSerialNo & "', '" & strA1P04 & "', '2492', 'TOT', '0', '" & strAMT2492 & "', '" & strA1P09 & "', '" & strA1P10 & "', '" & strA1P11 & "', '" & strA1P12 & "', '" & strA1P13 & "', '" & strA1p14 & "', '" & strA1P15 & "', " & _
                  "'" & strA1P16 & "', '" & strA1P17 & "', '" & strA1P18 & "', '" & strA1P19 & "', '" & strA1P20 & "', '" & strA1P21 & "', '" & stra1p22 & "', '" & strA1P23 & "', '" & strA1P24 & "', '" & strA1P25 & "', '" & strA1P26 & "', '" & stra1p27 & "', '" & strA1P30 & "', '" & strA1P31 & "')"
      adoTaie.Execute strSql, intI
      If strCP16 = "Y" Then  '多年聘任簽約且為部分收款
         MsgBox "顧問聘任簽約多年且為部分收款，請自行調整各收入科目金額！", vbInformation
      End If
   End If
   'end 2018/11/19
   
End Function
'END 2016/1/4

'Added by Morgan 2022/7/5
'轉換1公司科目
Private Function PUB_ConvAccNo(pA1P01 As String, pA1P05 As String) As String
   PUB_ConvAccNo = pA1P05
   '1公司的下列科目，一律改為4901專業其他收入
   '410110 商標收入-CCT法務
   '411107 專利收入 -CCP法務
   '412102 CFT收入 -法務
   '413102 CFP收入 -法務
   '417103 FCP收入 -法務
   '417203 FCT收入 -法務
   If pA1P01 = "1" Then
      Select Case pA1P05
      Case "410110", "411107", "412102", "413102", "417103", "417203"
         PUB_ConvAccNo = "490102"
      End Select
   End If
End Function

'ADD BY SONIA 2016/8/31 2016/9/1起FCP收入再細分科目
Public Function InsertFCPACC1P0(ByVal strA1P01 As String, ByVal stra1p02 As String, ByVal strA1P03 As String, ByVal strA1P04 As String, ByVal strA1p05 As String, _
                                ByVal strA1P06 As String, ByVal strA1P07 As String, ByVal strA1p08 As String, ByVal strA1P09 As String, ByVal strA1P10 As String, _
                                ByVal strA1P11 As String, ByVal strA1P12 As String, ByVal strA1P13 As String, ByVal strA1p14 As String, ByVal strA1P15 As String, _
                                ByVal strA1P16 As String, ByVal strA1P17 As String, ByVal strA1P18 As String, ByVal strA1P19 As String, ByVal strA1P20 As String, _
                                ByVal strA1P21 As String, ByVal stra1p22 As String, ByVal strA1P23 As String, ByVal strA1P24 As String, ByVal strA1P25 As String, _
                                ByVal strA1P26 As String, ByVal stra1p27 As String, ByVal strA1P30 As String, ByVal strA1P31 As String, ByVal strA0Z02 As String, ByVal strNation As String)
Dim strSql As String
Dim rstAdo As ADODB.Recordset, iRtn As Integer
Dim strAMT1C As String, m_AMT1C As String  '貸方金額 A1P08
Dim strAccNo As String
Dim strSerialNo As String '分錄序次
Dim m_A1P16 As String

   If stra1p22 = "null" Then stra1p22 = ""
   If stra1p27 = "null" Then stra1p27 = ""
   
   'modify by sonia 2016/9/14 M10504587之X10509441要扣除折扣A1L07,X10511506要扣除417103FCP收入-法務
   'modify by sonia 2016/11/30 +cpm03或cpm04
   strSql = "SELECT decode('" & strNation & "','000',cpm11,cpm24) CPM11,sum(A1L05-nvl(A1L07,0)) A1L05,decode('" & strNation & "','000',cpm03,cpm04) Property from ACC1L0,casepropertymap " & _
            "where A1L01='" & strA0Z02 & "' and substr(A1L04,-2) not in ('99','98') and A1L03=cpm01(+) and A1L04=cpm02(+) and cpm11||cpm24 is not null " & _
            "and decode('" & strNation & "','000',cpm11,cpm24)<>'" & strA1p05 & "'and decode('" & strNation & "','000',cpm11,cpm24)<>'417103' group by decode('" & strNation & "','000',cpm11,cpm24),decode('" & strNation & "','000',cpm03,cpm04)"
   iRtn = 1
   Set rstAdo = ClsLawReadRstMsg(iRtn, strSql)
   If iRtn = 1 Then
      While Not rstAdo.EOF
         m_AMT1C = Val("" & rstAdo.Fields("A1L05"))
         strAccNo = rstAdo.Fields("CPM11")
         strA1P06 = PUB_GETAccNODept(strAccNo, strA1P06)
         'add by sonia 2016/11/30 重組摘要欄A1P14,婉莘說要帶正確案件性質D105102072
         strA1p14 = strA1P17 & "/" & rstAdo.Fields("Property")
         'end 2016/11/30
         strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
            " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strA1P03 & "', '" & strA1P04 & "', '" & strAccNo & "', '" & strA1P06 & "', 0, '" & m_AMT1C & "', '" & strA1P09 & "', '" & strA1P10 & "', '" & strA1P11 & "', '" & strA1P12 & "', '" & strA1P13 & "', '" & strA1p14 & "', '" & strA1P15 & "', " & _
                     "'" & strA1P16 & "', '" & strA1P17 & "', '" & strA1P18 & "', '" & strA1P19 & "', '" & strA1P20 & "', '" & strA1P21 & "', '" & stra1p22 & "', '" & strA1P23 & "', '" & strA1P24 & "', '" & strA1P25 & "', '" & strA1P26 & "', '" & stra1p27 & "', '" & strA1P30 & "', '" & strA1P31 & "')"
         adoTaie.Execute strSql, intI
         strA1P03 = GetSerialNo("select max(a1p03) from acc1p0 where a1p02 = '" & stra1p02 & "' and a1p04 = '" & strA1P04 & "'", 3)
         strA1p08 = Val(strA1p08) - m_AMT1C
         rstAdo.MoveNext
      Wend
   End If
   
   '餘額放原科目
   strA1P06 = PUB_GETAccNODept(strA1p05, strA1P06)
   strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
      " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strA1P03 & "', '" & strA1P04 & "', '" & strA1p05 & "', '" & strA1P06 & "', '" & strA1P07 & "', '" & strA1p08 & "', '" & strA1P09 & "', '" & strA1P10 & "', '" & strA1P11 & "', '" & strA1P12 & "', '" & strA1P13 & "', '" & strA1p14 & "', '" & strA1P15 & "', " & _
               "'" & strA1P16 & "', '" & strA1P17 & "', '" & strA1P18 & "', '" & strA1P19 & "', '" & strA1P20 & "', '" & strA1P21 & "', '" & stra1p22 & "', '" & strA1P23 & "', '" & strA1P24 & "', '" & strA1P25 & "', '" & strA1P26 & "', '" & stra1p27 & "', '" & strA1P30 & "', '" & strA1P31 & "')"
   adoTaie.Execute strSql, intI
   
End Function
'END 2016/8/31

'add by sonia 2016/9/19
Public Function UpdateFCPACC1P0(ByVal strA1P01 As String, ByVal stra1p02 As String, ByVal strA1P03 As String, ByVal strA1P04 As String, ByVal strA1p05 As String, _
                                ByVal strA0Z02 As String, ByVal strNation As String, Optional ByVal strA1K35 As String = "")
Dim strSql As String
Dim rstAdo As ADODB.Recordset, iRtn As Integer
Dim strA1P06 As String  '部門
Dim strA1p08 As String  '貸方金額
Dim strAccNo As String
Dim strSerialNo As String '分錄序次
Dim m_AMT1C As String
Dim strA1p14 As String    '摘要  add by sonia 2016/11/30
Dim rstAdo2 As ADODB.Recordset 'Added by Morgan 2022/10/20

'Modified by Morgan 2022/10/20 RsTemp->rstAdo2,a1p02='F'->a1p02='" & stra1p02 & "'

   '先抓原資料
   strExc(0) = "SELECT * from ACC1p0 where a1p01='1' and a1p02='" & stra1p02 & "' and a1p03='" & strA1P03 & "' and a1p04='" & strItemNo & "'"
   intI = 1
   Set rstAdo2 = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strA1p08 = Val("" & rstAdo2.Fields("A1P08"))
   Else
      Exit Function
   End If
   
   'modify by sonia 2016/9/14 M10504587之X10509441要扣除折扣A1L07,X10511506要扣除417103FCP收入-法務
   'modify by sonia 2016/11/30 +cpm03或cpm04
   strSql = "SELECT decode('" & strNation & "','000',cpm11,cpm24) CPM11,sum(A1L05-nvl(A1L07,0)) A1L05,decode('" & strNation & "','000',cpm03,cpm04) Property from ACC1L0,casepropertymap " & _
            "where A1L01='" & strA0Z02 & "' and substr(A1L04,-2) not in ('99','98') and A1L03=cpm01(+) and A1L04=cpm02(+) and cpm11||cpm24 is not null " & _
            "and decode('" & strNation & "','000',cpm11,cpm24)<>'" & strA1p05 & "'and decode('" & strNation & "','000',cpm11,cpm24)<>'417103' group by decode('" & strNation & "','000',cpm11,cpm24),decode('" & strNation & "','000',cpm03,cpm04)"
   iRtn = 1
   Set rstAdo = ClsLawReadRstMsg(iRtn, strSql)
   If iRtn = 1 Then
      While Not rstAdo.EOF
         m_AMT1C = Val("" & rstAdo.Fields("A1L05"))
         strAccNo = rstAdo.Fields("CPM11")
         strA1P06 = PUB_GETAccNODept(strAccNo, strA1P06)
         strSerialNo = GetSerialNo("select max(a1p03) from acc1p0 where a1p01='1' and a1p02 = '" & stra1p02 & "' and a1p04 = '" & strA1P04 & "'", 3)
         'add by sonia 2016/11/30 重組摘要欄A1P14,婉莘說要帶正確案件性質D105102072
         strA1p14 = rstAdo2.Fields("A1P17") & "/" & rstAdo.Fields("Property")
         If strA1K35 <> "" Then strA1p14 = Mid(strA1K35, 1, 6) & "/" & strA1p14  'add by sonia 2023/1/18
         'end 2016/11/30
         'modify by sonia 2016/11/30 rstAdo2.Fields("A1P14")改為strA1p14
         strSql = "insert into acc1p0 (a1p01, a1p02, a1p03, a1p04, a1p05, a1p06, a1p07, a1p08, a1p09, a1p10, a1p11, a1p12, a1p13, a1p14, a1p15, a1p16, a1p17, a1p18, a1p19, a1p20, a1p21, a1p22, a1p23, a1p24, a1p25, a1p26, a1p27,a1p30,a1p31)" & _
            " values ('" & strA1P01 & "', '" & stra1p02 & "', '" & strSerialNo & "', '" & strA1P04 & "', '" & strAccNo & "', '" & strA1P06 & "', 0, '" & m_AMT1C & "', '" & "" & rstAdo2.Fields("A1P09") & "', '" & "" & rstAdo2.Fields("A1P10") & "', '" & "" & rstAdo2.Fields("A1P11") & "', '" & "" & rstAdo2.Fields("A1P12") & "', '" & "" & rstAdo2.Fields("A1P13") & "', '" & "" & strA1p14 & "', '" & "" & rstAdo2.Fields("A1P15") & "', " & _
                     "'" & "" & rstAdo2.Fields("A1P16") & "', '" & "" & rstAdo2.Fields("A1P17") & "', '" & "" & rstAdo2.Fields("A1P18") & "', '" & "" & rstAdo2.Fields("A1P19") & "', '" & "" & rstAdo2.Fields("A1P20") & "', '" & "" & rstAdo2.Fields("A1P21") & "', '" & "" & rstAdo2.Fields("A1P22") & "', '" & "" & rstAdo2.Fields("A1P23") & "', '" & "" & rstAdo2.Fields("A1P24") & "', '" & "" & rstAdo2.Fields("A1P25") & "', '" & "" & rstAdo2.Fields("A1P26") & "', '" & "" & rstAdo2.Fields("A1P27") & "', '" & "" & rstAdo2.Fields("A1P30") & "', '" & "" & rstAdo2.Fields("A1P31") & "')"
         adoTaie.Execute strSql, intI
         strA1p08 = Val(strA1p08) - m_AMT1C
         rstAdo.MoveNext
      Wend
   End If
   
   'add by sonia 2017/11/20 FCP及FMP案B類收文927其他翻譯且承辦人為外翻編號且相關總收文號為C類之結匯金額,已在結匯及請款單扣點數,此處收款要再加回來M10605730(X10617049(FCP-047593)
   strSql = "select a1p07,a1w01,a1w02,cp60,cp61 from acc1w0,caseprogress,acc1p0 where a1w01='" & strA0Z02 & "' and substr(a1w02,1,1)='B' and a1w02=cp09(+) " & _
               "and cp01 in ('P','FCP') and cp10='927' and substr(cp14,1,1)='F' and substr(cp43,1,1)='C' and cp61||a1w02=a1p23 and a1p07>0"
   iRtn = 1
   Set rstAdo = ClsLawReadRstMsg(iRtn, strSql)
   If iRtn = 1 Then
      strA1p08 = Val(strA1p08) + Val("" & rstAdo.Fields("a1p07"))
   End If
   'end 2017/11/20
   
   '餘額放原項次
   'modify by sonia 2017/3/10 M10600765餘額為0要刪除,M10600846餘額為負數者由財務處自行調整
   'strSql = "update acc1p0 set a1p08=" & strA1p08 & " where a1p01='1' and a1p02='F' and a1p03='" & strA1P03 & "' and a1p04='" & strItemNo & "'"
   If strA1p08 = 0 Then
      strSql = "delete acc1p0 where a1p01='1' and a1p02='" & stra1p02 & "' and a1p03='" & strA1P03 & "' and a1p04='" & strItemNo & "'"
   Else
      strSql = "update acc1p0 set a1p08=" & strA1p08 & " where a1p01='1' and a1p02='" & stra1p02 & "' and a1p03='" & strA1P03 & "' and a1p04='" & strItemNo & "'"
      'add by sonia 2017/3/10 M10600846餘額為負數者則彈訊息由財務處自行調整
      If strA1p08 < 0 Then
         MsgBox "FCP收入科目有負數產生, 請自行調整！" & Err.Description, vbCritical
      End If
      'end 2017/3/10
   End If
   'end 2017/3/10
   adoTaie.Execute strSql, intI
   
End Function
'end 2016/9/19

'Modify By Sindy 2016/11/3 同basPublic的PUB_PrintFontIntoBox
'*********************************************************************************************
'*********************************************************************************************
'放字進去框框中
'oStr 要放的字    要分行就要加 |
'oLeft, oUp  框框左上角
'oRight, 0Down 框框右下角
'IsCenter 是否置中
'IsAvgFont 是否自平均分配  與 IsCenterW = true  同時使用 無效
'Sub PrintFontIntoBox(ByVal oStr As String, oLeft As Integer, oUp As Integer, oRight As Integer, oDown As Integer, Optional IsCenterH As Boolean = True, Optional IsCenterW As Boolean = True)
'Modify By Sindy 2015/10/16 +Optional intFontSize As Integer = 0
'Remove by Lydia 2020/08/21 已改到basLetter
'Public Sub PUB_PrintFontIntoBox(ByVal oStr As String, oLeft As Integer, oUp As Integer, oRight As Integer, oDown As Integer, _
'                                Optional IsCenterH As Boolean = True, Optional IsCenterW As Boolean = True, _
'                                Optional IsAvgFont As Boolean = False, Optional intFontSize As Integer = 0)
'Dim BoxHeight As Integer
'Dim BoxWidth As Integer
'Dim FontHeight As Integer
'Dim FontWidth As Double
'Dim ArrStr As Variant
'Dim oIntI As Integer
'Dim oIntJ As Integer
'Dim FontTop As Integer
'Dim FontLeft As Integer
'Dim FontAllHeight As Integer
'Dim SingleFontWidth As Integer
'Dim CalFontWidth As Integer
'Dim SingleFont As String
'Dim TmpFont As String         '暫存的單字
'Dim TmpAllFont As String         '暫存的整格字
'Dim TmpLineFont As String
''add by nickc 2007/03/01
'Dim TmpPrtWd As Integer
'
'   'add by nickc 2005/09/09
'   oStr = Replace(oStr, vbCrLf, "|")
'   '先去跳行符號
'   oStr = Replace(Replace(oStr, Chr(13), ""), Chr(10), "")
'   BoxHeight = oDown - oUp
'   BoxWidth = oRight - oLeft
'   FontHeight = Printer.TextHeight(Mid(oStr, 1, 1))
'   FontWidth = Printer.TextWidth(Mid(oStr, 1, 1))
'   ArrStr = Split(Replace(oStr, vbCrLf, ""), "|")
'   '檢查若是超過長度，自動跳行
'   For oIntI = 0 To UBound(ArrStr)
'      '超過
'      TmpFont = ArrStr(oIntI)
'      If Left(Trim(ArrStr(oIntI)), 1) = "□" Then
'         'Modify By Sindy 2015/10/16
'         If intFontSize > 0 Then
'            Printer.Font.Size = intFontSize
'         Else
'         '2015/10/16 END
'            Printer.Font.Size = 14 '9
'         End If
'      Else
'         'Modify By Sindy 2015/10/16
'         If intFontSize > 0 Then
'            Printer.Font.Size = intFontSize
'         Else
'         '2015/10/16 END
'            Printer.Font.Size = 14
'         End If
'      End If
'      If TmpFont <> "" Then
'         FontHeight = Printer.TextHeight(Mid(Trim(ArrStr(oIntI)), 1, 1))
'         FontWidth = Printer.TextWidth(ArrStr(oIntI))
'         If Len(TmpFont) > (BoxWidth / FontWidth) Then
'            TmpAllFont = ""
'            SingleFont = ""
'            CalFontWidth = 0
'            SingleFontWidth = 0
'            TmpFont = ""
'            TmpLineFont = ""
'            TmpFont = ArrStr(oIntI)
'            Do While Not Len(TmpFont) = 0
'               SingleFont = PUB_GetOneFont(TmpFont)
'               SingleFontWidth = Printer.TextWidth(SingleFont)
'               If Printer.TextWidth(TmpLineFont) + SingleFontWidth > BoxWidth Then
'                  TmpAllFont = TmpAllFont & TmpLineFont & "|"
'                  TmpLineFont = ""
'                  CalFontWidth = 0
'               End If
'               TmpLineFont = TmpLineFont & SingleFont
'            Loop
'            If TmpLineFont <> "" Then
'               TmpAllFont = TmpAllFont & TmpLineFont
'            End If
'            If TmpAllFont <> "" Then
'               ArrStr(oIntI) = TmpAllFont
'            End If
'         End If
'      End If
'   Next oIntI
'   oStr = Join(ArrStr, "|")
'   ArrStr = Split(oStr, "|")
'   If IsCenterH = True Then
'      FontAllHeight = (Val(UBound(ArrStr)) + 1) * FontHeight
'      If FontAllHeight > BoxHeight Then FontAllHeight = BoxHeight
'      FontTop = ((BoxHeight - FontAllHeight) / 2) + oUp
'   Else
'      FontTop = oUp
'   End If
'   For oIntI = 0 To UBound(ArrStr)
'      If (FontTop + (FontHeight * (oIntI + 1))) < oDown Then
'         'FontTop = FontTop + (FontHeight * oIntI)
'         If IsCenterW = True Then
'            FontWidth = Printer.TextWidth(ArrStr(oIntI))
'            FontLeft = ((BoxWidth - FontWidth) / 2) + oLeft
'            'add by nickc 2007/03/08 遇到第一個是 □ 改縮小
'            If Left(Trim(ArrStr(oIntI)), 1) = "□" Then
'               'Modify By Sindy 2015/10/16
'               If intFontSize > 0 Then
'                  Printer.Font.Size = intFontSize
'               Else
'               '2015/10/16 END
'                  Printer.Font.Size = 14 '9
'               End If
'   '            FontHeight = Printer.TextHeight(Mid(oStr, 1, 1))
'   '            FontWidth = Printer.TextWidth(ArrStr(oIntI))
'   '            FontLeft = ((BoxWidth - FontWidth) / 2) + oLeft
'            End If
'            Printer.CurrentX = FontLeft
'            Printer.CurrentY = FontTop + (FontHeight * oIntI)
'            Printer.Print ArrStr(oIntI)
'            'add by nickc 2007/03/08 遇到第一個是 □ 改縮小
'            If Left(Trim(ArrStr(oIntI)), 1) = "□" Then
'               'Modify By Sindy 2015/10/16
'               If intFontSize > 0 Then
'                  Printer.Font.Size = intFontSize
'               Else
'               '2015/10/16 END
'                  Printer.Font.Size = 14
'               End If
'            End If
'         ElseIf IsAvgFont = True Then
'            'oLeft = BoxWidth \ Len(ArrStr(0))
'
'            For oIntJ = 1 To Len(ArrStr(oIntI))
'               TmpFont = Mid(ArrStr(oIntI), oIntJ, 1)
'               FontWidth = Printer.TextWidth(TmpFont)
'               TmpPrtWd = ((BoxWidth - FontWidth) \ (Len(ArrStr(oIntI)) - 1))
'               FontLeft = (((BoxWidth - FontWidth) \ (Len(ArrStr(oIntI)) - 1)) * (oIntJ - 1)) + oLeft
'               Printer.CurrentX = FontLeft
'               Printer.CurrentY = FontTop + (FontHeight * oIntI)
'               Printer.Print TmpFont
'            Next oIntJ
'         Else
'            FontLeft = oLeft
'            'add by nickc 2007/03/08 遇到第一個是 □ 改縮小
'            If Left(Trim(ArrStr(oIntI)), 1) = "□" Then
'               'Modify By Sindy 2015/10/16
'               If intFontSize > 0 Then
'                  Printer.Font.Size = intFontSize
'               Else
'               '2015/10/16 END
'                  Printer.Font.Size = 14 '9
'               End If
'   '            FontHeight = Printer.TextHeight(Mid(oStr, 1, 1))
'   '            FontWidth = Printer.TextWidth(ArrStr(oIntI))
'   '            FontLeft = ((BoxWidth - FontWidth) / 2) + oLeft
'            End If
'            Printer.CurrentX = FontLeft
'            Printer.CurrentY = FontTop + (FontHeight * oIntI)
'            Printer.Print ArrStr(oIntI)
'            'add by nickc 2007/03/08 遇到第一個是 □ 改縮小
'            If Left(Trim(ArrStr(oIntI)), 1) = "□" Then
'               'Modify By Sindy 2015/10/16
'               If intFontSize > 0 Then
'                  Printer.Font.Size = intFontSize
'               Else
'               '2015/10/16 END
'                  Printer.Font.Size = 14
'               End If
'            End If
'         End If
'      End If
'   Next oIntI
'End Sub

'Remove by Lydia 2020/08/21 已改到basLetter
'Public Function PUB_GetOneFont(ByRef oStr As String) As String
'Dim i As Integer
'
'   PUB_GetOneFont = ""
'   If Asc(Mid(oStr, 1, 1)) < 0 Or Asc(Mid(oStr, 1, 1)) > 256 Then
'      '雙位元組
'      PUB_GetOneFont = Mid(oStr, 1, 1)
'      oStr = Mid(oStr, 2)
'      Exit Function
'   Else
'      Select Case Mid(oStr, 1, 1)
'      '符號，或特殊字
'      Case ",", " ", ":", ";", "!"
'         PUB_GetOneFont = Mid(oStr, 1, 1)
'         oStr = Mid(oStr, 2)
'      '單位元組
'      Case Else
'         For i = 1 To Len(oStr)
'            If Asc(Mid(oStr, i, 1)) < 0 Or Asc(Mid(oStr, i, 1)) > 256 Then
'               Exit For
'            Else
'               Select Case Mid(oStr, i, 1)
'               '符號，或特殊字
'               Case ",", " ", ":", ";", "!"
'                  Exit For
'               Case Else
'                  If Asc(Mid(oStr, i, 1)) = 13 Or Asc(Mid(oStr, i, 1)) = 10 Then
'                     oStr = Mid(oStr, 2)
'                     Exit For
'                  Else
'                     PUB_GetOneFont = PUB_GetOneFont & Mid(oStr, i, 1)
'                  End If
'               End Select
'            End If
'         Next i
'         oStr = Mid(oStr, Len(PUB_GetOneFont) + 1)
'      End Select
'   End If
'End Function
''*********************************************************************************************
''*********************************************************************************************
'end 2020/08/21

'AddBd by Lydia 2017/02/23 抓每月固定傳票已攤的傳票金額
'Modified by Lydia 2017/05/22 +指定會計科目 sAccno
Public Function PUB_SumA1PtoU(ByVal AP01 As String, ByVal AP04_1 As String, Optional ByVal ExDate As String = "", Optional ByVal Bdate As String = "", Optional ByVal sAccno As String = "") As String
'AP04_1:每月固定傳票流水號
'Bdate:有效期間起始
'ExDate:上次處理日期
Dim rsA1 As New ADODB.Recordset
Dim strExDate As String, strBdate As String
Dim inA As Integer

    PUB_SumA1PtoU = ""
    
    If AP01 = "" Or AP04_1 = "" Then Exit Function
    
    If ExDate = "" Or Bdate = "" Then
       strSql = "select * from acc0d1 where axd01='" & AP01 & "' and axd02='" & AP04_1 & "' "
       inA = 1
       Set rsA1 = ClsLawReadRstMsg(inA, strSql)
       If inA = 1 Then
          strBdate = "" & rsA1.Fields("axd11")
          strExDate = "" & rsA1.Fields("axd04")
       Else
          Exit Function
       End If
    Else
       strBdate = Replace(Replace(Bdate, "/", ""), "_", "")
       strExDate = Replace(Replace(ExDate, "/", ""), "_", "")
    End If
    
    '以公司別+流水號抓ACC1P0之A1P01='公司別' AND A1P02='U' AND A1P04>=流水號||有效期間起始年月 AND A1P04<=流水號||上次處理日期的借方總額SUM(A1P07)
    'Modified by Lydia 2017/05/22 +index
    'strSql = "select nvl(sum(a1p07),0) s1 from acc1p0 where a1p01='" & AP01 & "' and a1p02='U' and A1P04>='" & AP04_1 & IIf(Trim(strBdate) <> "", strBdate, "00000") & "' AND A1P04<='" & AP04_1 & IIf(Trim(strExDate) <> "", strExDate, "00000") & "' "
    strSql = "select /*+ INDEX(ACC1P0 IDXA1P020405) */ sum(a1p07) s1 from acc1p0 where a1p01='" & AP01 & "' and a1p02='U' and A1P04>='" & AP04_1 & IIf(Trim(strBdate) <> "", strBdate, "00000") & "' AND A1P04<='" & AP04_1 & IIf(Trim(strExDate) <> "", strExDate, "00000") & "' "
    'Added by Lydia 2017/05/22 指定會計科目
    If sAccno <> "" Then strSql = strSql & " and substr(a1p05,1," & Len(sAccno) & ")='" & sAccno & "' "
    
    inA = 1
    Set rsA1 = ClsLawReadRstMsg(inA, strSql)
    If inA = 1 Then
       'Modified by Lydia 2017/05/22
       'PUB_SumA1PtoU = rsA1(0)
       PUB_SumA1PtoU = Val("" & rsA1(0))
    End If
    Set rsA1 = Nothing
End Function

'Added by Lydia 2017/03/17 取得財產目錄的類別
Public Function Pub_GetA2b02Name(Optional ByVal iTyp As String = "") As String
    
    If iTyp = "" Then '取得類別數
       Pub_GetA2b02Name = "4"
    Else
       Select Case iTyp
          Case "1": Pub_GetA2b02Name = "交通運輸設備"
          Case "2": Pub_GetA2b02Name = "生財器具"
          Case "3": Pub_GetA2b02Name = "電腦硬體"
          Case "4": Pub_GetA2b02Name = "電腦軟體"
       End Select
    End If
End Function

'Added by Lydia 2017/03/17 取得財產目錄的所別
Public Function Pub_GetA2b03Sname(Optional ByVal iTyp As String = "") As String
    
    If iTyp = "" Then '取得所別數
       Pub_GetA2b03Sname = "5"
    Else
       Select Case iTyp
          Case "1": Pub_GetA2b03Sname = "北所"
          Case "2": Pub_GetA2b03Sname = "中所"
          Case "3": Pub_GetA2b03Sname = "南所"
          Case "4": Pub_GetA2b03Sname = "高所"
          Case "5": Pub_GetA2b03Sname = "其他"
       End Select
    End If
End Function

'Add by Amy 2017/04/14
'傳入年月判斷智權點數傳票是否已過帳
'stStartNo:傳票起始號
'stEndNo:傳票截止號
'stTranNo:指定傳票號
Public Function Pub_ChkAxbPost(ByVal stStartNo As String, ByVal stEndNo As String, Optional ByVal stTranNo As String = "") As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, strWhere As String
    Dim intQ As Integer
    
    Pub_ChkAxbPost = False
    
    If stTranNo <> MsgText(601) Then
        strWhere = " Or ax202='" & stTranNo & "' "
    End If
    strWhere = " And ax202>='" & stStartNo & "' And ax202<='" & stEndNo & "' " & strWhere
    'Modify by Amy 2017/10/18 改排序
    strQ = "Select Distinct ax210 From Acc021 Where ax201='1' " & strWhere & " Order by  ax210 Desc"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        If IsNull(RsQ.Fields("ax210")) = False Then Pub_ChkAxbPost = True
    End If
    
    RsQ.Close
    
End Function

'取得內帳系統參數資料
'stField:欄位名稱
'stCmp:公司別
Public Function Pub_GetAcc0b0(ByVal stField As String, ByVal stCmp As String) As String
    Dim adoquery As New ADODB.Recordset
    Dim stQ As String
    
    Pub_GetAcc0b0 = ""
    
    stQ = "Select " & stField & " From Acc0b0 Where a0b04 = '" & stCmp & "' "
    If adoquery.State = adStateOpen Then adoquery.Close
    adoquery.CursorLocation = adUseClient
    adoquery.Open stQ, adoTaie, adOpenStatic, adLockReadOnly
    If adoquery.RecordCount <> 0 Then
        Pub_GetAcc0b0 = "" & adoquery.Fields(0)
    End If
    adoquery.Close
End Function

'取得所別
'Modify by Amy 2017/04/24 +bolShort參數
Public Function PUB_GetZone(ByVal stDept As String, Optional ByVal bolShort As Boolean = False) As String
    Select Case Mid(stDept, 2, 1)
        Case 1
            If bolShort = True Then
                PUB_GetZone = "北"
            Else
                PUB_GetZone = "台北所"
            End If
        Case 2
            If bolShort = True Then
                PUB_GetZone = "中"
            Else
                PUB_GetZone = "台中所"
            End If
        Case 3
            If bolShort = True Then
                PUB_GetZone = "南"
            Else
                PUB_GetZone = "台南所"
            End If
        Case 4
            If bolShort = True Then
                PUB_GetZone = "高"
            Else
                PUB_GetZone = "高雄所"
            End If
        Case Else
            PUB_GetZone = ""
    End Select
End Function

'Added by Lydia 2017/09/22 更新付款單的匯款方式=>5.台銀合併結匯
Public Sub PUB_UpdateA1811toType(ByVal strNoList As String)
Dim strUpd As String
Dim intJ As Integer
Dim rsWD As New ADODB.Recordset
'Added by Lydia 2019/08/06
Dim strGrp As String, strCom As String
Dim strA1810 As String, strA1718 As String
Dim bolUpd As Boolean

On Error GoTo ErrHandle

   '判斷同一批的1,2公司電匯符合"同收據公司別＋同受款人(考慮A1718代為結匯之客戶編號)＋同幣別＋非獨立水單A1812"，將其匯款方式改成5-台銀合併結匯
   strSql = "select a1803,a1903,a1917,a1810,a1718, count(distinct a1801) as TCounter From acc180, acc190, acc170 " & _
            "Where a1801 in (" & GetAddStr(strNoList) & ") and a1801=a1901 And a1902=a1702(+) and a1917<>'J' and a1811='2' and nvl(a1812,'N') <> 'Y' " & _
            "group by a1803, a1903, a1917,a1810,a1718 having count(distinct a1801) > 1 "
   strSql = strSql & "order by a1917,a1903,a1803 "
   intJ = 1
   Set rsWD = ClsLawReadRstMsg(intJ, strSql)
   If intJ = 1 Then
      With rsWD
          .MoveFirst
          cnnConnection.BeginTrans
          Do While Not .EOF
              strUpd = ""
              If "" & .Fields("a1810") <> "" Then strUpd = strUpd & " and a1810='" & .Fields("a1810") & "' "
              If "" & .Fields("a1718") <> "" Then strUpd = strUpd & " and a1718='" & .Fields("a1718") & "' "
              
              If "" & .Fields("a1917") = "1" Or "" & .Fields("a1917") = "2" Then   'Added by Lydia 2020/08/31 排除L公司=> 限1,2公司
                  strUpd = "update acc180 set a1811='5' where a1801 in (select a1801 from acc180,acc190 Where a1801 = a1901 And a1801 in (" & GetAddStr(strNoList) & _
                           ") and a1803='" & .Fields("a1803") & "' and a1903='" & .Fields("a1903") & "' and a1917='" & .Fields("a1917") & "' " & strUpd & _
                           "and a1917<>'J' and a1811='2' group by a1801)"
                  cnnConnection.Execute strUpd, intJ
              End If 'Added by Lydia 2020/08/31
             .MoveNext
          Loop
          cnnConnection.CommitTrans
      End With
   End If
   'Added by Lydia 2019/08/06 代理人若同時有在1 2公司結匯(電匯+合併結匯), 則直接合併結匯. 條件如下:
                                           '1.合併至金額較大的那個公司
                                           '2.付款明細需備註原始歸類的公司別
   strSql = "select a1803,a1903,a1917,a1810,a1718,a1801,sum(a1904) uamt From acc180, acc190, acc170 " & _
            "Where a1801 in (" & GetAddStr(strNoList) & ") and a1801=a1901 And a1902=a1702(+) and a1917<>'J' and a1811 in ('2','5') and nvl(a1812,'N') <> 'Y' " & _
            "Group By A1803, A1903, A1917,A1810,A1718,a1801 "
   strSql = strSql & "order by a1803,a1903,uamt desc,a1801 "
   intJ = 1
   Set rsWD = ClsLawReadRstMsg(intJ, strSql)
   If intJ = 1 Then
      With rsWD
          .MoveFirst
          Do While Not .EOF
              'Added by Lydia 2020/08/31 L公司都歸類在2公司 (整批電匯) ,因為L公司目前沒有外幣戶
              If "" & .Fields("a1917") = "L" Then
                    If bolUpd = False Then
                        cnnConnection.BeginTrans
                        bolUpd = True
                    End If
                    strUpd = "update acc190 set a1917='2' where a1901='" & .Fields("a1801") & "' and a1917='L' "
                    cnnConnection.Execute strUpd, intJ
                    'Remove by Lydia 2020/10/27 先取消變更公司別
                    'If intJ > 0 Then
                    '    strUpd = "update acc180 set a1811='2' ,a1813='原公司別：" & .Fields("a1917") & "'||decode(a1813,null,'',';'||a1813) where a1801='" & .Fields("a1801") & "' "
                    '    cnnConnection.Execute strUpd, intJ
                    'End If
                    'end 2020/10/27
                    strCom = "" & .Fields("a1917") '公司別(第一張)
                    strGrp = ""  '單號+代理人+幣別=>清空,避免同一代理人+幣別有1,2,L公司
                    
              ElseIf "" & .Fields("a1917") = "1" Or "" & .Fields("a1917") = "2" Then   '限1,2公司
              'end 2020/08/31
                    If Mid(strGrp, 10) = "" & .Fields("a1803") & .Fields("a1903") Then '代理人+幣別
                          If bolUpd = False Then
                              cnnConnection.BeginTrans
                              bolUpd = True
                          End If
                          If strCom <> "" & .Fields("a1917") Then
                              strUpd = ""
                              If "" & .Fields("a1810") <> "" Then strUpd = strUpd & " and a1810='" & .Fields("a1810") & "' " '代理人名稱(A1810)
                              If "" & .Fields("a1718") <> "" Then strUpd = strUpd & " and a1718='" & .Fields("a1718") & "' " '代為結匯之客戶編號
                              '上一張
                              strUpd = "update acc180 set a1811='5' where a1801='" & Mid(strGrp, 1, 9) & "' "
                              cnnConnection.Execute strUpd, intJ
                              '現在(變更公司別)
                              strUpd = "update acc180 set a1811='5' ,a1813='原公司別：" & .Fields("a1917") & "'||decode(a1813,null,'',';'||a1813) where a1801='" & .Fields("a1801") & "' "
                              cnnConnection.Execute strUpd, intJ
                              If intJ > 0 Then
                                  strUpd = "update acc190 set a1917='" & strCom & "' where a1901='" & .Fields("a1801") & "' "
                                  cnnConnection.Execute strUpd, intJ
                              End If
                          Else '不用變更公司別
                              '上一張
                              strUpd = "update acc180 set a1811='5' where a1801='" & Mid(strGrp, 1, 9) & "' "
                              cnnConnection.Execute strUpd, intJ
                              '現在
                              strUpd = "update acc180 set a1811='5' where a1801='" & .Fields("a1801") & "' "
                              cnnConnection.Execute strUpd, intJ
                          End If
                    Else
                          strCom = "" & .Fields("a1917") '公司別(第一張)
                    End If
                    strGrp = "" & .Fields("a1801") & .Fields("a1803") & .Fields("a1903")  '單號+代理人+幣別
              End If 'Added by Lydia 2020/08/17
             .MoveNext
          Loop
          If bolUpd = True Then
             cnnConnection.CommitTrans
          End If
      End With
   End If
   
   Exit Sub
   
ErrHandle:
   If strUpd <> "" Then cnnConnection.RollbackTrans
End Sub

'Add by Amy 2017/09/29 傳入公司(stCmp)及傳票年月(stYM),取其公司別最大傳票日
Public Function Pub_GetMaxA0205(ByVal stCmp As String, ByVal stYM As String) As String
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    Pub_GetMaxA0205 = ""
    
    strQ = "Select Max(A0205) as A0205 From Acc020 " & _
            "Where A0201='" & stCmp & "' And SubStr(A0202,1,6)='D" & stYM & "' " & _
            "Having Max(A0205) is not null "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        Pub_GetMaxA0205 = "" & RsQ.Fields("A0205")
    End If
    RsQ.Close
End Function

'add by sonia 2017/12/8 傳入公司(stCmp)及傳票年月(stYM),取其公司別最大傳票日
Public Function CheckAX210(ByVal stCmp As String, ByVal stYM_F As String, ByVal stYM_T As String) As Boolean
Dim adoacc021 As New ADODB.Recordset
Dim s_msg As String
Dim strQ As String 'Add by Amy 2020/08/12
    
   CheckAX210 = False
   s_msg = ""
   'Moidfy by Amy 2020/08/12 改公司別
   If stCmp <> "" Then
        If InStr(stCmp, "+") > 0 Then
            strQ = " And a0201 in('" & Replace(stCmp, "+", "','") & "')"
        Else
            strQ = " And a0201='" & stCmp & "' "
        End If
    End If
'   If stCmp <> "" Then
'      adoacc021.Open "select distinct a0205 from acc021, acc020 where a0201=ax201(+) and a0202=ax202(+) and a0201 = '" & IIf(stCmp = "2", "J", stCmp) & "' and a0205 >= " & Val(stYM_F) & "01" & " and a0205 <= " & Val(stYM_T) & "31" & " and nvl(ax210,0)=0", adoTaie, adOpenStatic, adLockReadOnly
'   Else
'      adoacc021.Open "select distinct a0205 from acc021, acc020 where a0201=ax201(+) and a0202=ax202(+) and a0205 >= " & Val(stYM_F) & "01" & " and a0205 <= " & Val(stYM_T) & "31" & " and nvl(ax210,0)=0", adoTaie, adOpenStatic, adLockReadOnly
'   End If
    strQ = "select distinct a0205 from acc021, acc020 where a0201=ax201(+) and a0202=ax202(+)  " & strQ & " and a0205 >= " & Val(stYM_F) & "01" & " and a0205 <= " & Val(stYM_T) & "31" & " and nvl(ax210,0)=0 "
    adoacc021.CursorLocation = adUseClient
    adoacc021.Open strQ, adoTaie, adOpenStatic, adLockReadOnly
   'end 2020/08/12
   If adoacc021.RecordCount > 0 Then
      adoacc021.MoveFirst
      Do While Not adoacc021.EOF
         If s_msg = "" Then s_msg = "尚有未過帳傳票日期："
         s_msg = s_msg & adoacc021.Fields(0) & "；"
         adoacc021.MoveNext
      Loop
      If s_msg <> "" Then
         CheckAX210 = True
         MsgBox s_msg
      End If
      adoacc021.Close
      Exit Function
   End If
   adoacc021.Close

End Function
'end 2017/12/8

'Added by Morgan 2018/12/13
'211準備程序、212言詞辯論若串的相關總收文號(C類來函)的相關總收文號(A類)為訴願程序，在發文後不扣智權同仁點數。 Ex:P116209
Public Function PUB_ChkNoXFee(pCP09 As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select b.cp09 from caseprogress a,caseprogress b,caseprogress c" & _
      " where a.cp09='" & pCP09 & "' and a.cp43>'C' and b.cp09(+)=a.cp43 and c.cp09(+)=b.cp43 and c.cp10='501'"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      PUB_ChkNoXFee = True
   End If
   Set rsQuery = Nothing
End Function

'Modify by Amy 2019/09/16 從frmacc1172搬至aacc_fu,+stA4504 並改一年內發票號碼不可重覆
'Add by Amy 2014/03/10 同一個付款單號的發票號碼不可重覆
Public Function ChkA4504(ByVal stA4504 As String) As Boolean
    Dim stQuery As String
    Dim intR As Integer
    Dim rsQuery As New ADODB.Recordset
    
    ChkA4504 = False
    'Modify by Amy 2019/09/16 改一年內發票號碼不可重覆
    'stQuery = "Select 'Y' From Acc450 Where A4501 = '" & strA4501 & "' And A4504='" & Text2 & "' "
    stQuery = "Select 'Y' From Acc450 Where A4504='" & stA4504 & "' And A4503>=" & strSrvDate(1) - 19120000 & " And A4503<=" & strSrvDate(2)
    intR = 1
    Set rsQuery = ClsLawReadRstMsg(intR, stQuery)
    If intR = 1 Then
        ChkA4504 = True
    End If
    Set rsQuery = Nothing
End Function
'end 2014/03/10

'Modify By Sindy 2014/8/4 因為此作業Account和Promoter等都有呼叫到,以防在Account需要加入一堆Form,因此抽出來至Func
'不可以放basFunction因AutoBatchDay有此Func
Public Sub frm210132_SubPubShowNextData(cmdState As Integer, ByRef oForm As Form)
Dim i As Integer, j As Integer
Dim StrTag As String

   'frm210132 未列印收據/請款單查詢(智權部也有使用) : 會呼叫此函數
End Sub

'Add by Amy 2022/03/28 付款通知單-退費收訖憑單-票據受領收據 A4中一刀套印,從Frmacc14b0 (使用時不可開啟Word)
'strChoose:1-票據受領收據/3-退費收訖憑單(票據)/3.1-退費收訖憑單(非票據)/4-銷帳轉帳收退費/5-翻譯費證明單/6-銷帳轉暫收轉帳
'stCmpNo:公司別
'p_RetNo:回執單號
'p_PDate:日期
'strShowMsg:回傳訊息
'bolOnlyReceipt:只印回執表格(不印付款通知單)
'bolHalfSheet:只放半張紙
'Modified by Lydia 2024/03/12 pSaveDoc: 存檔位置
Public Function PUB_PrintReceipt_Doc(ByVal strFormN As String, ByVal strChoose As String, p_ado0e0 As ADODB.Recordset, stCmpNo As String, p_RetNo As String, _
  Optional ByVal p_PDate As String = "", Optional ByRef strShowMsg As String = "", Optional ByVal bolOnlyReceipt As Boolean = False, Optional ByVal bolHalfSheet As Boolean = False, _
  Optional ByRef pSaveDoc As String = "") As Boolean
    Dim oTable As Word.Table
    Dim RsQ As New ADODB.Recordset, strIns As String, intQ As Integer, intRun As Integer, intRow As String
    Dim bVisible As Boolean, m_WordLeft As Long, m_WordTop As Long, bolPayNotice As Boolean
    Dim lngAmount As Long, lngAmt As Long
    Dim strFileName As String, strQ As String, strTp(2) As String
    Dim str0E0Data(5) As String, strTitleDesc(1) As String, strCompName As String, strDesc As String, strDate As String
    Dim strReceTitle As String, strDocTitle As String, intLen As Integer, intFontSize As Integer 'Add by Amy 2022/03/29
        
On Error GoTo ErrHnd

    PUB_PrintReceipt_Doc = False: intRun = 0
    bolPayNotice = True '是否印付款通知單(只印一次)
    If bolOnlyReceipt = True Then bolPayNotice = False
    
    '5-翻譯費證明單
    If strChoose = "5" Then
        '依 PUB_PrintReceipt5 修改
        strQ = "Select * From Acc250, Acc0i0 Where a2501='" & p_RetNo & "' And a0i01(+)=a2503"
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, strQ)
        If intQ = 1 Then
            strDocTitle = "證　明　單"
            strDesc = "" & RsQ.Fields("a2514") & " 翻譯費"
            lngAmt = "" & RsQ.Fields("a2504")
        End If
    '4-銷帳轉帳收退費/轉帳
    ElseIf strChoose = "4" Or strChoose = "6" Then
        Set RsQ = p_ado0e0.Clone
        strDesc = "" & RsQ.Fields("a2514")
        If strChoose = "4" Then
            strDocTitle = "退費收訖憑單"
        Else
            strDocTitle = "退費轉帳同意書"
        End If
    '3-案件未辦退
    ElseIf Left(strChoose, 1) = "3" Then
        strDocTitle = "退費收訖憑單"
        If UCase(strFormN) = UCase("Frmacc14b0") Then
            'Memo by Amy 需抓退費明細資料(可能會有多張收據)-從Frmacc14b0搬過來
            strQ = "Select a0k01,a0k11,a0s06,a0s07,a0o01,a0s17,a0k04 From Acc0o0, Acc0s0, Acc0k0 " & _
                      "Where a0s01(+)=a0o09 And a0k01(+)=a0s02 And a0o09 is not null And SubStr(a0s02,1,1)='E' " & _
                      "And a0o07='" & p_ado0e0.Fields("A0q19") & "' And a0o03='" & p_ado0e0.Fields("a0q03") & "' " & _
                      "And a0o11=" & p_ado0e0.Fields("a0q01") & " Order by a0s02"
            intQ = 1
            Set RsQ = ClsLawReadRstMsg(intQ, strQ)
            If intQ = 1 Then
               RsQ.MoveFirst
               strDesc = PUB_GetCaseInfo("" & RsQ.Fields("a0k01"))  '退費明細
            End If
            '退費收訖憑單(非票據)
            If strChoose = "3.1" Then
               lngAmount = Val("" & p_ado0e0.Fields("a0q06").Value)
            Else
               lngAmount = Val("" & p_ado0e0.Fields("a0e11").Value)
            End If
        ElseIf UCase(strFormN) = UCase("Frmacc12a0") Then
            Set RsQ = p_ado0e0.Clone
            strDesc = PUB_GetCaseInfo("" & RsQ.Fields("a0k01"))  '退費明細
        'Modify By Sindy 2022/4/11
        Else 'If UCase(strFormN) = UCase("Frmacc11i0") Then
            'Memo by Amy 從Frmacc11i0 搬過來
            'Modify By Sindy 2022/4/12 + ," & Val("" & p_ado0e0.Fields("a2504").Value) & " as A2504,a0q03,a0o01
            strQ = "Select a0k11,a0k01,a0q05,a0k04," & Val("" & p_ado0e0.Fields("a2504").Value) & " as A2504,a0q03,a0o01 From Acc0o0, Acc0q0, Acc0s0, Acc0k0 " & _
                     "Where a0o01='" & RsTemp.Fields("A2505") & "' And a0q01(+)=a0o11 And a0q03(+)=a0o03 " & _
                     "And a0s01(+)=a0o09 And a0k01(+)=a0s02 "
            intQ = 1
            Set RsQ = ClsLawReadRstMsg(intQ, strQ)
            If intQ = 1 Then
               RsQ.MoveFirst
               strDesc = PUB_GetCaseInfo("" & RsQ.Fields("a0k01"))  '退費明細
            End If
'        Else
'           Set RsQ = p_ado0e0.Clone
        '2022/4/11 END
        End If
    End If
    
    '公司名稱
    If strChoose = "3" And UCase(strFormN) = UCase("Frmacc12a0") Then
        'PUB_PrintReceipt3 2020/4/23 辜說1公司收據銷退自4/1起回執改印智慧所(4/20退X75235)
        strCompName = A0802Query("" & IIf(RsQ.Fields("a0k11") = "1", "2", RsQ.Fields("a0k11")))
    '翻譯費 證明單
    ElseIf strChoose = "5" Then
        '依 PUB_PrintReceipt5
        strCompName = A0802Query("2")
    ElseIf stCmpNo <> "" Then
        strCompName = A0802Query(stCmpNo)
    'Add By Sindy 2022/4/12
    Else
      strCompName = A0802Query("" & RsQ.Fields("a0k11"))
      '2022/4/12 END
    End If
    
    '付款通知單-變數設定
    If bolPayNotice = True Then
        '付款通知單-抬頭及說明
        strTitleDesc(0) = Replace(ReportSum(43), ":", "：") & vbCrLf '台 鑒
        strTitleDesc(0) = strTitleDesc(0) & String(2, "　") & Replace(ReportSum(44), "  ", " ") '茲寄上
        If Left(strChoose, 1) = "3" Then
            strTitleDesc(0) = Replace(strTitleDesc(0), "票 據 受 領 收 據", "退 費 收 訖 憑 單")
        Else
            strDesc = ReportSum(46) & Replace(ReportSum(47), "  ", " ") '特別說明
        End If
        strTitleDesc(0) = strTitleDesc(0) & ReportSum(45) & vbCrLf
    End If
    
    'Modify By Sindy 2022/4/12 + Or strChoose = "1"
    If (UCase(strFormN) = UCase("Frmacc14b0") And strChoose <> "3.1") Or (UCase(strFormN) = UCase("Frmacc12a0") And strChoose = "1") Or strChoose = "1" Then
        '付款行庫
        str0E0Data(1) = Replace(ReportSum(37), ":", " ：") & A0g02Query(p_ado0e0.Fields("a0e01")) & vbCrLf
        '付款帳號
        str0E0Data(2) = Replace(ReportSum(38), ":", " ：") & p_ado0e0.Fields("a0e07") & vbCrLf
        '支票號碼
        str0E0Data(3) = Replace(ReportSum(39), ":", " ：") & p_ado0e0.Fields("a0e02") & vbCrLf
        '到 期 日
        str0E0Data(4) = Replace(ReportSum(40), ":   ", " ： ") & CFDate(p_ado0e0.Fields("a0e10")) & vbCrLf
        '金額
        str0E0Data(5) = Replace(ReportSum(41), ":", " ：") & "$" & Format(p_ado0e0.Fields("a0e11"), DDollar) & "**" & vbCrLf
        '備註
        str0E0Data(0) = Replace(ReportSum(42), ":", " ：")
        Select Case "" & p_ado0e0.Fields("a1p26")
            Case "1"
                str0E0Data(0) = str0E0Data(0) & ComboItem(111) '1--核駁退費
            Case "2"
                str0E0Data(0) = str0E0Data(0) & ComboItem(112) '2--溢收款
            Case "3"
                str0E0Data(0) = str0E0Data(0) & ComboItem(113) '3--案件未辦退費
            Case "4"
                str0E0Data(0) = str0E0Data(0) & ComboItem(114) '4--扣繳
            Case "5"
                str0E0Data(0) = str0E0Data(0) & ComboItem(115) '5--貨款
            Case "6"
                str0E0Data(0) = str0E0Data(0) & ComboItem(116) '6--稅款繳款書
            Case "7"
                If Not IsNull(p_ado0e0("a0q18")) Then
                    str0E0Data(0) = str0E0Data(0) & p_ado0e0.Fields("a0q18")
                Else
                    str0E0Data(0) = str0E0Data(0) & ComboItem(117) '7--其他
                End If
        End Select
        str0E0Data(0) = str0E0Data(0)
    End If
    
    '回執表格內文變數設定
    If Left(strChoose, 1) = "3" Or UCase(strFormN) = UCase("Frmacc12a0") Then
        strTitleDesc(1) = "" & strCompName
    Else
        strTitleDesc(1) = strCompName & Replace(ReportSum(43), ":", "：") & vbCrLf '台 鑒
        strTitleDesc(1) = strTitleDesc(1) & String(2, "　") & Replace(ReportSum(51), "  ", " ") '茲 收 到
    End If
   
    '開Word
    'Modified by Lydia 2024/03/12
    'strFileName = "$$退費收訖憑單-票據受領收據.doc"
    'If Dir(App.path & "\" & strUserNum & "\" & strFileName) <> "" Then
    '    Kill App.path & "\" & strUserNum & "\" & strFileName
    'End If
    If pSaveDoc <> "" Then
       strFileName = pSaveDoc '來源程式已先檢查
    Else
       strFileName = "$$退費收訖憑單-票據受領收據.doc"
       If Dir(App.path & "\" & strUserNum & "\" & strFileName) <> "" Then
          Kill App.path & "\" & strUserNum & "\" & strFileName
       End If
       strFileName = App.path & "\" & strUserNum & "\" & strFileName
    End If
    'end 2024/03/12


    If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Function
    
    'Printer.PaperSize = 9 '印表機 紙張設 A4 Mark by Amy 2022/03/31 財務無效
    With g_WordAp
        .Documents.add
        .Application.WindowState = wdWindowStateMinimize
        'Modified by Lydia 2024/03/12
        '.ActiveDocument.SaveAs App.path & "\" & strUserNum & "\" & strFileName
        .ActiveDocument.SaveAs strFileName
        .Selection.PageSetup.PaperSize = wdPaperA4 'Add by Amy 2022/03/31
        With .ActiveDocument.PageSetup
            .Orientation = wdOrientPortrait '直印
            .TopMargin = g_WordAp.CentimetersToPoints(1)
            .BottomMargin = g_WordAp.CentimetersToPoints(1)
            .LeftMargin = g_WordAp.CentimetersToPoints(1)
            .RightMargin = g_WordAp.CentimetersToPoints(1)
            .Gutter = g_WordAp.CentimetersToPoints(0)
            .HeaderDistance = g_WordAp.CentimetersToPoints(0)
            .FooterDistance = g_WordAp.CentimetersToPoints(0)
        End With
        
        .Selection.Tables.add Range:=.Selection.Range, NumRows:=2, NumColumns:=1
        .ActiveDocument.Tables(1).Rows(1).Cells(1).Select
        .Selection.Cells.SetHeight RowHeight:=.CentimetersToPoints(14.2), HeightRule:=wdRowHeightExactly
         
        .ActiveDocument.Tables(1).Rows(2).Cells(1).Select
        .Selection.Cells.SetHeight RowHeight:=.CentimetersToPoints(12), HeightRule:=wdRowHeightExactly
        
        '**** 付款通知單 ****
        If bolPayNotice = True Then
            '抬頭
            .ActiveDocument.Tables(1).Rows(1).Cells(1).Select
            .Selection.ParagraphFormat.SpaceBefore = 0
            .Selection.ParagraphFormat.SpaceAfter = 0
            .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly '固定行高
            .Selection.ParagraphFormat.LineSpacing = 18
            .Selection.Font.Name = "新細明體"
            .Selection.Font.Size = 14
            .Selection.TypeText strCompName  '公司名稱
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
            .Selection.TypeParagraph '換行
            
            .Selection.TypeText ReportTitle(1111) '付款通知單
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
            .Selection.TypeParagraph '換行
            .Selection.Font.Size = 12
            .Selection.Font.Underline = wdUnderlineSingle '文字底線
            .Selection.TypeText ReportSum(35) & Replace(CFDate(strSrvDate(2)), ":", "：") & String(116, " ") '製表日期
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
            .Selection.Font.Underline = wdUnderlineNone '文字無底線
            .Selection.TypeParagraph '換行
            .Selection.TypeParagraph '換行
            '內文
            .Selection.TypeText "" & p_ado0e0.Fields("a0q05") & strTitleDesc(0)
            .Selection.TypeParagraph '換行
            .Selection.TypeText String(2, "　") & str0E0Data(1)  '付款行庫
            .Selection.TypeText String(2, "　") & str0E0Data(2)  '付款帳號
            .Selection.TypeText String(2, "　") & str0E0Data(3)  '支票號碼
            .Selection.TypeText String(2, "　") & str0E0Data(4)  '到 期 日
            .Selection.TypeText String(2, "　") & str0E0Data(5)  '金額
            .Selection.TypeText String(2, "　") & str0E0Data(0)  '備註
            .Selection.TypeParagraph '換行
            .Selection.TypeParagraph '換行
            strTp(1) = strDesc
            If Left(strChoose, 1) = "3" Then
                 strTp(1) = "退 費 明 細 ：" & strTp(1)
            End If
            .Selection.TypeText strTp(1) '退費明細/特別說明
            bolPayNotice = False
            intRun = intRun + 1
        End If
        '**** End 付款通知單 ****
        
        '判斷跳上表 or 下表
        If intRun = 0 Then
            .ActiveDocument.Tables(1).Rows(1).Cells(1).Select
        Else
            .ActiveDocument.Tables(1).Rows(2).Cells(1).Select
        End If

        If strChoose = "5" Then
'*** 5-翻譯費證明單 ***
            If p_PDate = MsgText(601) Then
                strDate = "日期  " & CFDate(strSrvDate(2))
            Else
                strDate = "日期  " & IIf(InStr(p_PDate, "/") > 0, p_PDate, CFDate(TransDate(p_PDate, 1)))
            End If
            .Selection.Font.Name = "標楷體"
            .Selection.Font.Size = 12
            .Selection.TypeText ReportSum(50)  '請 沿 此 虛 線 撕 下 寄 回
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
            .Selection.TypeParagraph '換行
            .Selection.Font.Size = 24
            .Selection.Font.Underline = wdUnderlineSingle '文字底線
            .Selection.TypeText strDocTitle '證明單
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
            .Selection.Font.Underline = wdUnderlineNone '文字無底線
            .Selection.TypeParagraph '換行
            
            .Selection.Font.Size = 14
            '畫表格
            Set oTable = g_WordAp.Selection.Tables.add(Range:=g_WordAp.Selection.Range, NumRows:=9, NumColumns:=1)
            With oTable
                .Select
                .Rows.SetLeftIndent LeftIndent:=g_WordAp.CentimetersToPoints(0.5), RulerStyle:=wdAdjustNone '設定表格中列的縮排
                .Columns(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(17), RulerStyle:=wdAdjustProportional
                
                intRow = 1
                .Rows(intRow).Cells(1).Select
                g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                g_WordAp.Selection.TypeText p_RetNo & String(20, "　") & strDate   '日期
                g_WordAp.Selection.Cells.VerticalAlignment = wdCellAlignVerticalBottom '垂直靠下
                
                intRow = intRow + 1
                .Rows(intRow).Cells(1).Select
                g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                g_WordAp.Selection.TypeText "茲收到　" & strCompName
                g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
                '格線
                .Rows(intRow).Borders(wdBorderTop).LineStyle = wdLineStyleSingle '實心線
                .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                
                strTp(0) = ChangeNumber(str(lngAmt))
                If Len(strTp(0)) > 20 Then
                    strTp(1) = "　 "
                Else
                    strTp(1) = String(Val(20 - Len(strTp(0))), "　")
                End If
                intRow = intRow + 1
                .Rows(intRow).Cells(1).Select
                g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                g_WordAp.Selection.TypeText "新台幣　"
                g_WordAp.Selection.Font.Underline = wdUnderlineSingle '文字底線
                g_WordAp.Selection.TypeText strTp(0)
                g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                g_WordAp.Selection.TypeText strTp(1)
                g_WordAp.Selection.Font.Underline = wdUnderlineSingle '文字底線
                g_WordAp.Selection.TypeText "NTD" & Format(lngAmt, DDollar)
                g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                '格線
                .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle '實心線
                .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                
                '*** 上述款項說明 ***
                strTp(1) = "": strTp(2) = ""
                strTp(0) = strDesc
                If LenB(strTp(0)) >= 25 Then
                    strTp(2) = strTp(0)
                    strTp(0) = PUB_StrToStr(strTp(2), 50)
                    strTp(2) = Replace(strTp(2), strTp(0), "")
                    If strTp(2) <> MsgText(601) Then
                        strTp(1) = PUB_StrToStr(strTp(2), 50)
                    End If
                    strTp(2) = Replace(strTp(2), strTp(1), "")
                    If strTp(2) <> MsgText(601) Then
                        'Modify by Amy 2022/04/13 和婧瑄確認若無法完整顯示,可用...翻譯費顯示
                        strTp(1) = Mid(strTp(1), 1, Len(strTp(1)) - 7) & "...翻譯費"
                    End If
                End If
                intRow = intRow + 1
                .Rows(intRow).Cells(1).Select
                g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                g_WordAp.Selection.TypeText "上述款項系　"
                g_WordAp.Selection.Font.Underline = wdUnderlineSingle '文字底線
                g_WordAp.Selection.TypeText strTp(0)
                g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                g_WordAp.Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop '垂直靠上
                g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                '格線
                .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle '實心線
                .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                
                intRow = intRow + 1
                .Rows(intRow).Cells(1).Select
                g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                g_WordAp.Selection.TypeText "　　　　　　"
                g_WordAp.Selection.Font.Underline = wdUnderlineSingle '文字底線
                g_WordAp.Selection.TypeText strTp(1)
                g_WordAp.Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop '垂直靠上
                g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                '格線
                .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle '實心線
                .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                '*** End 上述款項說明 ***
                
                strTp(0) = "　" & RsQ.Fields("a2513") & "　"
                strTp(1) = "　　簽章："
                intRow = intRow + 1
                .Rows(intRow).Cells(1).Select
                g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                g_WordAp.Selection.TypeText "　　　　　　具　領　人："
                g_WordAp.Selection.Font.Underline = wdUnderlineSingle '文字底線
                g_WordAp.Selection.TypeText strTp(0)
                g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                g_WordAp.Selection.TypeText String(2, "　") & strTp(1)
                g_WordAp.Selection.Font.Underline = wdUnderlineSingle '文字底線
                g_WordAp.Selection.TypeText "　　　　　"
                g_WordAp.Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop '垂直靠上
                g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                '格線
                .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle '實心線
                .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                
                '*** 地址 ***
                intRow = intRow + 1
                .Rows(intRow).Cells(1).Select
                g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                g_WordAp.Selection.TypeText "　　　　　　地　　　址："
                
                strTp(1) = "": strTp(2) = ""
                'Modify By Sindy 2022/4/12 + Replace ex:(新北市永和區竹林路６０-５號　　１４樓　　　　　　　　　　　　　　)
                strTp(0) = "" & RsQ.Fields("a0i04") & Trim(Replace(RsQ.Fields("a0i03"), "　", ""))
                intLen = 22
                If LenB(strTp(0)) > intLen * 4 Then
                    intLen = 36
                    intFontSize = 10
                End If
                If LenB(strTp(0)) >= intLen Then
                    strTp(2) = strTp(0)
                    strTp(0) = PUB_StrToStr(strTp(2), intLen * 2)
                    strTp(1) = String(12, "　") & Replace(strTp(2), strTp(0), "")
                End If
                If intFontSize <> 0 Then
                    g_WordAp.Selection.Font.Size = intFontSize
                End If
                g_WordAp.Selection.TypeText strTp(0)
                g_WordAp.Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop '垂直靠上
                g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                '格線
                .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle '實心線
                .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                
                intRow = intRow + 1
                .Rows(intRow).Cells(1).Select
                g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                If intFontSize <> 0 Then
                    g_WordAp.Selection.Font.Size = intFontSize
                End If
                g_WordAp.Selection.TypeText strTp(1)
                g_WordAp.Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop '垂直靠上
                g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                '格線
                .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle '實心線
                .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                '*** End 地址 ***
                
                g_WordAp.Selection.Font.Size = 14
                intRow = intRow + 1
                .Rows(intRow).Cells(1).Select
                g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                '格線
                .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle '實心線
                .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                
                g_WordAp.Selection.Font.Size = 12
                g_WordAp.Selection.TypeText "敬請於證明單上簽名或蓋章並退回本所　謝謝！"
                g_WordAp.Selection.Cells.VerticalAlignment = wdCellAlignVerticalBottom '垂直靠下
                g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                '格線
                .Rows(intRow).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle '實心線
                .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                '列印
                If pSaveDoc = "" Then 'Added by Lydia 2024/03/12
                   g_WordAp.PrintOut Background:=False, Copies:=1, Collate:=True
                End If
            End With
'*** End 5-翻譯費證明單 ***
        ElseIf Left(strChoose, 1) = "3" Or strChoose = "4" Or strChoose = "6" Then
'*** 3-案件未辦退費=銷退付款 / 4-銷帳轉帳收退費 / 6-銷帳轉暫收轉帳 ***
            Do While RsQ.EOF = False
                '銷帳轉帳收退費 / 銷帳轉暫收轉帳
                If strChoose = "4" Or strChoose = "6" Then
                    strReceTitle = "" & RsQ.Fields("a2513")
                    lngAmt = "" & RsQ.Fields("a2504")
                '案件未辦退費=銷退付款
                Else
                    strReceTitle = "" & RsQ.Fields("a0k04")
                    If UCase(strFormN) = UCase("Frmacc14b0") Then
                        lngAmt = Val("" & RsQ.Fields("a0s06")) + Val("" & RsQ.Fields("a0s07"))
                    Else
                        lngAmt = RsQ.Fields("a2504")
                    End If
                End If
                
                '回執單號
                If p_RetNo = "" Then
                    If Left(strChoose, 1) = "3" And UCase(strFormN) = UCase("Frmacc14b0") Then
                        'Memo by Amy 因發現每次從frmacc14b0列印都會產生新的回執單號,故增加判斷,查詢G單號是否已存在,抓舊號 (也為可重印地址條)-與Morgan及瑞婷討論後之修改方式
                        '增加 A2520 寫入「收據編號」
                        p_PDate = ""
                        p_RetNo = GetA2501("3", "" & RsQ.Fields("a0o01"), "" & RsQ.Fields("a0k01"), p_PDate)
                        If p_RetNo = "" Then
                            p_RetNo = AutoNo("H", 5)
                            strIns = "Values('" & p_RetNo & "','3','" & p_ado0e0.Fields("a0q03") & "'," & lngAmt & ",'" & RsQ.Fields("a0o01") & "','" & strUserNum & "','" & ChgSQL("" & p_ado0e0.Fields("a0q05")) & "','" & RsQ.Fields("a0k01") & "')"
                            strIns = "Insert Into Acc250(A2501,A2502,A2503,A2504,A2505,A2506,A2513,A2520) " & strIns
                            adoTaie.Execute strIns
                        End If
                    'Modify By Sindy 2022/4/12 Mark: If UCase(strFormN) = UCase("Frmacc12a0") Then
                    Else 'If UCase(strFormN) = UCase("Frmacc12a0") Then
                        '依 PUB_PrintReceipt3/4 修改
                        p_RetNo = AutoNo("H", 5)
                        If strChoose = "4" Or strChoose = "6" Then
                            strIns = "Values('" & p_RetNo & "','" & strChoose & "','" & RsQ.Fields("a2503") & "'," & lngAmt & ",'" & RsQ.Fields("a2505") & "','" & strUserNum & "','" & ChgSQL(RsQ.Fields("a2513")) & "','" & ChgSQL(RsQ.Fields("a2514")) & "','" & RsQ.Fields("a2520") & "')"
                            strIns = "Insert Into Acc250(A2501,A2502,A2503,A2504,A2505,A2506,A2513,A2514,A2520) " & strIns
                        Else
                            strIns = "Values('" & p_RetNo & "','3','" & RsQ.Fields("a0q03") & "'," & lngAmt & ",'" & RsQ.Fields("a0o01") & "','" & strUserNum & "','" & ChgSQL("" & RsQ.Fields("a0q05")) & "','" & RsQ.Fields("a0k01") & "')"
                            strIns = "Insert Into Acc250(A2501,A2502,A2503,A2504,A2505,A2506,A2513,A2520) " & strIns
                        End If
                        adoTaie.Execute strIns
                    End If
                End If
                If p_PDate = MsgText(601) Then
                    strDate = "日期  " & CFDate(strSrvDate(2))
                Else
                    strDate = "日期  " & IIf(InStr(p_PDate, "/") > 0, p_PDate, CFDate(TransDate(p_PDate, 1)))
                End If
                If UCase(strFormN) = UCase("Frmacc14b0") Then
                    If RsQ.AbsolutePosition = RsQ.RecordCount Then
                        If lngAmt - Val("" & RsQ.Fields("a0s17")) <> lngAmount Then
                            strShowMsg = strShowMsg & vbCrLf & "<" & p_ado0e0.Fields("a0q03") & ">" & p_ado0e0.Fields("a0q05")
                        End If
                    End If
                End If
                .Selection.Font.Name = "標楷體"
                .Selection.Font.Size = 12
                .Selection.TypeText ReportSum(50)  '請 沿 此 虛 線 撕 下 寄 回
                .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                .Selection.TypeParagraph '換行
                .Selection.Font.Size = 24
                .Selection.Font.Underline = wdUnderlineSingle '文字底線
                .Selection.TypeText strDocTitle
                .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
                .Selection.Font.Underline = wdUnderlineNone '文字無底線
                .Selection.TypeParagraph '換行
                
                .Selection.Font.Size = 14
                '畫表格
                Set oTable = g_WordAp.Selection.Tables.add(Range:=g_WordAp.Selection.Range, NumRows:=9, NumColumns:=1)
                With oTable
                    .Select
                    .Rows.SetLeftIndent LeftIndent:=g_WordAp.CentimetersToPoints(0.5), RulerStyle:=wdAdjustNone '設定表格中列的縮排
                    .Columns(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(17), RulerStyle:=wdAdjustProportional
                    
                    intRow = 1
                    .Rows(intRow).Cells(1).Select
                    g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                    g_WordAp.Selection.TypeText p_RetNo & String(20, "　") & strDate   '日期
                    g_WordAp.Selection.Cells.VerticalAlignment = wdCellAlignVerticalBottom '垂直靠下
                    
                    intRow = intRow + 1
                    .Rows(intRow).Cells(1).Select
                    g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                    g_WordAp.Selection.TypeText "茲收到　" & strCompName
                    g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
                    '格線
                    .Rows(intRow).Borders(wdBorderTop).LineStyle = wdLineStyleSingle '實心線
                    .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                    .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                    
                    strTp(0) = ChangeNumber(str(lngAmt))
                    If Len(strTp(0)) > 20 Then
                        strTp(1) = "　 "
                    Else
                        strTp(1) = String(Val(20 - Len(strTp(0))), "　")
                    End If
                    intRow = intRow + 1
                    .Rows(intRow).Cells(1).Select
                    g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                    g_WordAp.Selection.TypeText "新台幣　"
                    g_WordAp.Selection.Font.Underline = wdUnderlineSingle '文字底線
                    g_WordAp.Selection.TypeText strTp(0)
                    g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                    g_WordAp.Selection.TypeText strTp(1)
                    g_WordAp.Selection.Font.Underline = wdUnderlineSingle '文字底線
                    g_WordAp.Selection.TypeText "NTD" & Format(lngAmt, DDollar)
                    g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                    g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                    '格線
                    .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle '實心線
                    .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                    
                    '*** 上述款項說明 ***
                    strTp(0) = "": strTp(1) = "": strTp(2) = ""
                    If Left(strChoose, 1) = "3" Then
                        strDesc = PUB_GetCaseInfo("" & RsQ.Fields("a0k01"))  '退費明細 'Added by Morgan 2023/6/1
                        strTp(0) = "退回 " & strDesc & " 款項"
                    Else
                        strTp(0) = strDesc
                    End If
                    If LenB(strTp(0)) >= 25 Then
                        strTp(2) = strTp(0)
                        strTp(0) = PUB_StrToStr(strTp(2), 50)
                        strTp(2) = Replace(strTp(2), strTp(0), "")
                        If strTp(2) <> MsgText(601) Then
                            strTp(1) = PUB_StrToStr(strTp(2), 50)
                        End If
                        strTp(2) = Replace(strTp(2), strTp(1), "")
                    End If
                    intRow = intRow + 1
                    .Rows(intRow).Cells(1).Select
                    g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                    g_WordAp.Selection.TypeText "上述款項系　"
                    g_WordAp.Selection.Font.Underline = wdUnderlineSingle '文字底線
                    g_WordAp.Selection.TypeText strTp(0)
                    g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                    g_WordAp.Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop '垂直靠上
                    g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                    '格線
                    .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle '實心線
                    .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                    
                    intRow = intRow + 1
                    .Rows(intRow).Cells(1).Select
                    g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                    g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                    g_WordAp.Selection.TypeText "　　　　　　"
                    g_WordAp.Selection.Font.Underline = wdUnderlineSingle '文字底線
                    g_WordAp.Selection.TypeText strTp(1)
                    g_WordAp.Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop '垂直靠上
                    g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                    '格線
                    .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle '實心線
                    .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                    
                    intRow = intRow + 1
                    .Rows(intRow).Cells(1).Select
                    g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1), HeightRule:=wdRowHeightExactly
                    g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                    g_WordAp.Selection.TypeText "　　　　　　"
                    g_WordAp.Selection.Font.Underline = wdUnderlineSingle '文字底線
                    g_WordAp.Selection.TypeText strTp(2)
                    g_WordAp.Selection.Cells.VerticalAlignment = wdCellAlignVerticalTop '垂直靠上
                    g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                    g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                    '格線
                    .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle '實心線
                    .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                    '*** End 上述款項說明 ***
                    
                    strTp(0) = strReceTitle
                    intRow = intRow + 1
                    .Rows(intRow).Cells(1).Select
                    g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(1.5), HeightRule:=wdRowHeightExactly
                    g_WordAp.Selection.Font.Size = 18
                    g_WordAp.Selection.Font.Underline = wdUnderlineSingle '文字底線
                    g_WordAp.Selection.TypeText strTp(0)
                    g_WordAp.Selection.Font.Underline = wdUnderlineNone '文字無底線
                    g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
                    '格線
                    .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle '實心線
                    .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                    
                    intRow = intRow + 1
                    .Rows(intRow).Cells(1).Select
                    g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(0.7), HeightRule:=wdRowHeightExactly
                    g_WordAp.Selection.Font.Size = 12
                    g_WordAp.Selection.TypeText "敬請於(退費收訖憑單)上簽蓋　台端(貴公司)之收款章；並退回本所　謝謝！"
                    g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
                    '格線
                    .Rows(intRow).Borders(wdBorderBottom).LineStyle = wdLineStyleSingle '實心線
                    .Rows(intRow).Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                    .Rows(intRow).Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                    
                    intRow = intRow + 1
                    .Rows(intRow).Cells(1).Select
                    g_WordAp.Selection.Cells.SetHeight RowHeight:=g_WordAp.CentimetersToPoints(0.7), HeightRule:=wdRowHeightExactly
                    g_WordAp.Selection.Font.Size = 12
                    'Modified by Morgan 2022/12/6 --瑞婷
                    'g_WordAp.Selection.TypeText "請蓋章後寄回或傳真(02)25068147"
                    g_WordAp.Selection.TypeText "請蓋章後寄回或傳真(02)25011666"
                    'end 2022/12/6
                    g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight '靠右
                End With
                If bolHalfSheet = False Then intRun = intRun + 1 '印A4紙
                
                '印半張紙
                If bolHalfSheet = True Then
                    .ActiveDocument.Tables(1).Rows(2).Cells(1).Select
                    .Selection.Delete Unit:=wdCharacter, Count:=1
                    If pSaveDoc = "" Then 'Added by Lydia 2024/03/12
                       .PrintOut Background:=False, Copies:=1, Collate:=True
                    End If
                '判斷下表是否已有資料
                ElseIf intRun = 2 Then
                    With g_WordAp
                        If pSaveDoc = "" Then 'Added by Lydia 2024/03/12
                          .PrintOut Background:=False, Copies:=1, Collate:=True
                        End If
                        '印完刪上下表內容
                        .ActiveDocument.Tables(1).Select
                        .Selection.Delete Unit:=wdCharacter, Count:=1
                        '跳至上表
                        .ActiveDocument.Tables(1).Rows(1).Cells(1).Select
                        '非最後一筆才設0,才不會最後一筆又印空白頁
                        If RsQ.AbsolutePosition <> RsQ.RecordCount Then intRun = 0
                    End With
                Else
                    .ActiveDocument.Tables(1).Rows(2).Cells(1).Select
                End If
                p_RetNo = ""
                If Left(strChoose, 1) = "3" And UCase(strFormN) = UCase("Frmacc14b0") Then
                    lngAmount = lngAmount - (lngAmt - Val("" & RsQ.Fields("a0s17")))
                End If
                RsQ.MoveNext
            Loop
            '印A4最後一筆
            If bolHalfSheet = False And intRun <> 2 Then
                g_WordAp.Selection.Delete Unit:=wdCharacter, Count:=1
                If pSaveDoc = "" Then 'Added by Lydia 2024/03/12
                   g_WordAp.PrintOut Background:=False, Copies:=1, Collate:=True
                End If
            End If
'*** End 3-案件未辦退費=銷退付款 / 4-銷帳轉帳收退費 / 6-銷帳轉暫收轉帳  ***
        ElseIf strChoose = "1" Then
'*** 1-扣繳 ***
            If p_RetNo = "" Then
                '回執單號 (Memo 目前從frmacc14b0/PUB_PrintReceipt1 搬過來修改,查詢G單號是否已存在,抓舊號-瑞婷)
                '增加 A2520 寫入「國內付款單號」
                p_PDate = ""
                strTp(0) = GetA0o01("" & p_ado0e0.Fields("a0q03"), "" & p_ado0e0.Fields("a0q01"))
                p_RetNo = GetA2501("1", strTp(0), "" & p_ado0e0.Fields("a0q17"), p_PDate)
                If p_RetNo = "" Then
                    p_RetNo = AutoNo("H", 5)
                    strIns = "Values('" & p_RetNo & "','1','" & p_ado0e0.Fields("a0q03") & "'," & Val("" & p_ado0e0.Fields("a0e11")) & ",'" & strTp(0) & "','" & strUserNum & "','" & ChgSQL("" & p_ado0e0.Fields("a0q05")) & "','" & p_ado0e0.Fields("a0q17") & "' )"
                    strIns = "Insert Into Acc250(A2501,A2502,A2503,A2504,A2505,A2506,A2513,A2520) " & strIns
                    adoTaie.Execute strIns
                End If
            End If
            
            If p_PDate = MsgText(601) Then
                strDate = "日期  " & CFDate(strSrvDate(2))
            Else
                strDate = "日期  " & IIf(InStr(p_PDate, "/") > 0, p_PDate, CFDate(TransDate(p_PDate, 1)))
            End If
            .Selection.Font.Name = "新細明體"
            .Selection.Font.Size = 12
            .Selection.TypeText ReportSum(50)  '請 沿 此 虛 線 撕 下 寄 回
            .Selection.TypeParagraph '換行
            .Selection.Font.Size = 14
            .Selection.TypeText ReportTitle(1114) '票據受領收據
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
            .Selection.TypeParagraph '換行
            .Selection.Font.Size = 12
            .Selection.Font.Underline = wdUnderlineSingle '文字底線
            .Selection.TypeText String(2, "　") & p_RetNo & String(25, "　") & strDate & String(14, "　")   '日期
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft '靠左
            .Selection.Font.Underline = wdUnderlineNone '文字無底線
            .Selection.TypeParagraph '換行
            .Selection.TypeParagraph '換行
            .Selection.TypeParagraph 'Add by Amy 2022/04/14 瑞婷說再加一行空白
            '內文
            .Selection.TypeText A0802Query("" & p_ado0e0.Fields("a1p01")) & Replace(ReportSum(43), ":", "：")  '台 鑒
            .Selection.TypeParagraph '換行
            .Selection.TypeText Replace(ReportSum(51), "  ", " ") '茲收到...
            .Selection.TypeParagraph '換行
            '畫表格
            Set oTable = g_WordAp.Selection.Tables.add(Range:=g_WordAp.Selection.Range, NumRows:=5, NumColumns:=2)
            With oTable
                .Select
                .Rows.SetLeftIndent LeftIndent:=38.3, RulerStyle:=wdAdjustNone '設定表格中列的縮排
                .Columns(1).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(7.5), RulerStyle:=wdAdjustProportional
                .Columns(2).SetWidth ColumnWidth:=g_WordAp.CentimetersToPoints(6), RulerStyle:=wdAdjustProportional
                '格線
                .Borders(wdBorderTop).LineStyle = wdLineStyleSingle '實心線
                .Borders(wdBorderBottom).LineStyle = wdLineStyleSingle
                .Borders(wdBorderLeft).LineStyle = wdLineStyleSingle
                .Borders(wdBorderRight).LineStyle = wdLineStyleSingle
                .Borders(wdBorderVertical).LineStyle = wdLineStyleSingle
                .Borders(wdBorderHorizontal).LineStyle = wdLineStyleSingle
                '欄位名稱
                .Rows(1).Cells(1).Select
                g_WordAp.Selection.TypeText "票       據       內        容"
                g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
                .Rows(1).Cells(2).Select
                g_WordAp.Selection.TypeText "蓋　　　　　章"
                g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter '置中
                '第二列-票據內容
                .Rows(2).Cells(1).Select
                g_WordAp.Selection.TypeText String(1, "　") & str0E0Data(1)  '付款行庫
                g_WordAp.Selection.TypeText String(1, "　") & str0E0Data(2)  '付款帳號
                g_WordAp.Selection.TypeText String(1, "　") & str0E0Data(3)  '支票號碼
                g_WordAp.Selection.TypeText String(1, "　") & str0E0Data(4)  '到 期 日
                g_WordAp.Selection.TypeText String(1, "　") & str0E0Data(5)  '金額
                '第三列-票據內容
                .Rows(3).Cells(1).Select
                .Rows(3).Cells.Merge
                g_WordAp.Selection.TypeText String(1, "　") & str0E0Data(0)
                '第四列-具領人
                .Rows(4).Cells(1).Select
                .Rows(4).Cells.Merge
                g_WordAp.Selection.TypeText "具領人：" & p_ado0e0.Fields("a0q05")
                g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight '靠右
                .Rows(4).Borders(wdBorderLeft).LineStyle = wdLineStyleNone '無線條
                .Rows(4).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                '第五列
                .Rows(5).Cells(1).Select
                .Rows(5).Cells.Merge
                'Modified by Morgan 2022/12/6 --瑞婷
                'g_WordAp.Selection.TypeText "請蓋章後寄回或傳真(02)25068147"
                g_WordAp.Selection.TypeText "請蓋章後寄回或傳真(02)25011666"
                'end 2022/12/6
                g_WordAp.Selection.ParagraphFormat.Alignment = wdAlignParagraphRight '靠右
                g_WordAp.Selection.MoveUp Unit:=wdLine, Count:=2, Extend:=wdExtend
                .Rows(5).Borders(wdBorderTop).LineStyle = wdLineStyleNone '無線條
                .Rows(5).Borders(wdBorderBottom).LineStyle = wdLineStyleNone
                .Rows(5).Borders(wdBorderLeft).LineStyle = wdLineStyleNone
                .Rows(5).Borders(wdBorderRight).LineStyle = wdLineStyleNone
                '列印
                If pSaveDoc = "" Then 'Added by Lydia 2024/03/12
                   g_WordAp.PrintOut Background:=False, Copies:=1, Collate:=True
                End If
            End With
'*** End 1-扣繳 ***
        End If
        'Added by Lydia 2024/03/12
        If pSaveDoc <> "" Then
           .ActiveDocument.Close wdSaveChanges
        Else
        'end 2024/03/12
           .ActiveDocument.Close wdDoNotSaveChanges
        End If
        .Quit wdDoNotSaveChanges
        Set g_WordAp = Nothing
    End With 'g_WordAp
   
    PUB_PrintReceipt_Doc = True
    Exit Function
    
ErrHnd:
    g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
    g_WordAp.Quit wdDoNotSaveChanges
    Set g_WordAp = Nothing
    Set RsQ = Nothing
    pSaveDoc = "" 'Added by Lydia 2024/03/12
    
    MsgBox Err.Description, vbCritical
End Function
'抓取客戶應付單號
Public Function GetA0o01(ByVal stA0o03 As String, ByVal stA0o11 As String) As String
    Dim rsA As New ADODB.Recordset
    Dim strA As String, intA As Integer
    
    GetA0o01 = ""
    strA = "Select Max(a0o01) a0o01 From Acc0o0 Where A0o03='" & stA0o03 & "' And A0o11=" & stA0o11
    
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, strA)
    If intA = 1 Then
        GetA0o01 = "" & rsA.Fields("a0o01")
    End If
    Set rsA = Nothing
End Function
'抓取回執記錄編號
Public Function GetA2501(ByVal stA2502 As String, ByVal stA2505 As String, ByVal stA2520 As String, Optional ByRef stDate As String = "") As String
    Dim rsA As New ADODB.Recordset
    Dim strA As String, intA As Integer
    
    GetA2501 = ""
    strA = "Select Distinct A2501,A2518 " & _
                "From Acc250 Where a2502='" & stA2502 & "' And a2505='" & stA2505 & "' And a2520='" & stA2520 & "' " & _
                "Order by A2501,A2518 Desc"
    intA = 1
    Set rsA = ClsLawReadRstMsg(intA, strA)
    If intA = 1 Then
        GetA2501 = "" & rsA.Fields("A2501")
        stDate = "" & rsA.Fields("A2518")
    End If
    Set rsA = Nothing
End Function
'end 2022/03/28

'Add By Sindy 2022/9/30 因為Account有引用 basFlow 會連帶需要引用到 Service1
'但因接洽單電子收文就會呼叫到一些案件系統函數, 所以才建此虛函數
Public Function PUB_AutoRecvCRLMain(strSys As String, strCRL01 As String) As Boolean
End Function

'Add by Amy 2023/04/18 確認 ACS分潤當月合計是否與期末金額相符
'stDate:民國年 前5碼(YYYMM)
Public Function ChkACSIncomeAndEndAmt(ByVal stFormN As String, ByVal stDate As String, Optional ByVal stAxb17 As String = "", Optional ByRef stMsg As String) As Boolean
    Dim strACSAmt As String, strACS2492Amt As String 'ACS 收入金額/保留金額
    Dim strAxb17(0) As String, strFixBack As String
    
    ChkACSIncomeAndEndAmt = False
    Select Case UCase(stFormN)
        Case "FRMACC4320" '過帳及分攤作業
            strFixBack = ",不可過帳 ！"
        'Mark by Amy 2023/06/05 不檢查
'        Case "FRMACC43C0" '每月業績開放/關閉輸入
'            strFixBack = ",不可關閉 ！"
    End Select
    
    If stAxb17 = MsgText(601) Then
        Call bolAcc0b1(9, stDate, strAxb17())
        stAxb17 = strAxb17(0)
    End If
    
    strACSAmt = "Y"
    Call GetACSData("0", "ChkACSIncomeAndEndAmt", stDate, ",Acc020", , strACSAmt)
    'Modify by  Amy 2023/06/05 不需管是否已產生期末傳票(stAxb17是否有值),一律以當月strACSAmt和strACS2492Amt比即可-秀玲
'    If Val(strACSAmt) > 0 And stAxb17 = MsgText(601) Then
'        stMsg = stDate & "月 ACS需分潤案有收款資料" & vbCrLf & _
'                    "但未產生「期末實績保留傳票」" & strFixBack
'        Exit Function
'    ElseIf stAxb17 <> MsgText(601) Then
        '確認所有ACS需分潤案 當月收入總金額 是否=當月2492 貸方總金額
        'Memo 11203月已經過帳,報表已完成,不會有重新過帳的問題,故11203月會不和不用管-秀玲(因 2492有做調整分錄,故會收入與 2492會不和)
        strACS2492Amt = "Y"
        'Modify by Amy 2023/07/31 改GetACSData
        Call GetACSData("8", "ChkACSIncomeAndEndAmt", stDate, ",Acc020", , strACS2492Amt)
        If Val(strACSAmt) <> Val(strACS2492Amt) Then
            stMsg = stDate & "月 ACS需分潤案有收款資料" & vbCrLf & _
                                "與目前「期末實績保留傳票」金額不符" & strFixBack
            Exit Function
        End If
'    End If
    'end 2023/06/05
    ChkACSIncomeAndEndAmt = True
End Function

'Added by Lydia 2017/09/28 檢查華銀媒體檔的欄位字首
'Move by Lydia 2025/07/21 從frmacc24m0移過來，並且改成共用模組
Public Function CheckFstSpec(ByVal pStr As String) As String
Dim Str01 As String
'Added by Lydia 2017/09/30
Dim inX As Integer, inJ As Integer
Dim nStr01 As String
Dim strSpecList As String
Dim bolOK As Boolean
    inX = 36 '指定位數不可為特殊字元
   'Swift文數字規定ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890+-/()?.,' ,第1個字母不可放-或/
    strSpecList = "ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890+()?.,'" '可接受的字元,排除「-」或「/」
'end 2017/09/30
      
   CheckFstSpec = pStr
   If Len(Trim(pStr)) > 0 Then
      'Modified by Lydia 2017/09/30 華銀經過實際測試, 前面加上空格是不可行
      '所以改成,若第36個字元為「-」時，第34個字後方塞空白，第35個字自動往後到第36個字元　ex.104-0032 =>10 4-0032
      'Str01 = UCase(Mid(pStr, 1, 1))
      ''Swift文數字規定ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890+-/()?.,' ,第1個字母不可放-或/
      'If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ1234567890+()?.,'", Str01) = 0 And Str01 <> " " Then
      '   CheckFstSpec = " " & pStr
      'End If
      '----------------
      Str01 = UCase(Mid(pStr, inX, 1))
      '若空白後,接特殊字元
      If Str01 = " " And Trim(Mid(pStr, inX + 1)) <> "" Then
         Str01 = UCase(Mid(pStr, inX + 1, 1))
         inX = inX + 1
      ElseIf Str01 <> " " And UCase(Mid(pStr, inX - 1, 1)) = " " Then
         inX = inX - 1
      End If
      
      '檢查指定位數不可為特殊字元
      If InStr(strSpecList, Str01) = 0 And Str01 <> " " Then
         bolOK = False
         '往前判斷字元後塞空白,令指定位數不可為特殊字元
         For inJ = inX - 1 To 2 Step -1
             nStr01 = Mid(CheckFstSpec, inJ, 1)
             If InStr(strSpecList, nStr01) > 0 And nStr01 <> "" Then
                bolOK = True
                Exit For
             End If
         Next inJ
         If bolOK = True Then
            CheckFstSpec = Mid(pStr, 1, inJ - 1) & String(IIf(inX = 36, 1, 2), " ") & Trim(Mid(pStr, inJ))
         End If
      End If
      'end 2017/09/30
   End If
End Function
