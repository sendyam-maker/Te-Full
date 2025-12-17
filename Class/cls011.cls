VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls011"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
'讀取解除期限資料
Public Function ReadRelieveDeadlineRst(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String) As ADODB.Recordset
Dim strSQL As String, rsRecordset As New ADODB.Recordset
Dim objPublicData As Object

On Error GoTo Err
Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
strSQL = "select " & SQLDate("cp05") & " 來函收文日,decode(np07," + CNULL(大陸國家代號) + ",cpm04,cpm03) 下一程序," & SQLDate("np08") & " 本所期限," & SQLDate("np09") & " 法定期限,staff.st02 承辦人,staff1.st02 業務員,np13 機關文號,np14 相關人,np01,np07,dbms_rowid.rowid_to_restricted(nextprogress.rowid,0) rid from caseprogress,nextprogress,casepropertymap,staff,staff staff1 where np02=cpm01 and np07=cpm02 and cp14=staff.st01(+) and np10=staff1.st01(+) and cp01=np02 and cp02=np03 and cp03=np04 and cp04=np05 and np06 is null and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and cp09=np01"
Set ReadRelieveDeadlineRst = objPublicData.ReadRst(strSQL)
Set objPublicData = Nothing
Exit Function
Err:
Set objPublicData = Nothing
ErrorLog
End Function
'讀取取消收文資料
Public Function ReadCancelReceivedDayRst(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String) As ADODB.Recordset
Dim strSQL As String, rsRecordset As New ADODB.Recordset
Dim objPublicData As Object
'Add By Cheng 2002/01/14
Dim rs1 As New ADODB.Recordset

On Error GoTo Err
Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
'Add By Cheng 2002/01/14
'若案件性質為"不續辦"或"閉卷"的資料不抓(各系統之代號不同)
If rs1.State <> adStateClosed Then rs1.Close
Set rs1 = Nothing
rs1.CursorLocation = adUseClient
rs1.Open "Select SK02 From SystemKind Where SK01 = '" & "" & strCode1 & "'", _
         cnnConnection, adOpenStatic, adLockReadOnly
If rs1.RecordCount > 0 Then
   Select Case rs1.Fields(0).Value
   '專利或專利服務
   Case "1", "5"
      '911024 nick 取消收文應該為 null ，只有取消收文會用到
      'strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp57") & " 取消收文日 " & _
               " from caseprogress,casepropertymap,staff,staff staff1 " & _
               " where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp27 is null and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" & _
               " And (CP10 <> '907' And CP10 <> '913')"
      '92.1.12 MODIFY BY SONIA 取消發文日的限制
      'strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp57") & " 取消收文日 " & _
      '         " from caseprogress,casepropertymap,staff,staff staff1 " & _
      '         " where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp27 is null and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" & _
      '         " And (CP10 <> '907' And CP10 <> '913') and cp57 is null "
      '92.7.29 MODIFY BY SONIA 增加顯示發文日
      'strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp57") & " 取消收文日 " & _
      '         " from caseprogress,casepropertymap,staff,staff staff1 " & _
      '         " where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" & _
      '         " And (CP10 <> '907' And CP10 <> '913') and cp57 is null "
      strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp27") & " 發文日," & SQLDate("cp57") & " 取消收文日 " & _
               " from caseprogress,casepropertymap,staff,staff staff1 " & _
               " where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" & _
               " And (CP10 <> '907' And CP10 <> '913') and cp57 is null "
      '92.7.29 END
      '92.1.12 END
   '商標或商標服務
   Case "2", "6"
      '911024 nick 取消收文應該為 null ，只有取消收文會用到
      'strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp57") & " 取消收文日 " & _
               " from caseprogress,casepropertymap,staff,staff staff1 " & _
               " where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp27 is null and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" & _
               " And (CP10 <> '703' And CP10 <> '704')"
      '92.1.12 MODIFY BY SONIA 取消發文日的限制
      'strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp57") & " 取消收文日 " & _
      '         " from caseprogress,casepropertymap,staff,staff staff1 " & _
      '         " where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp27 is null and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" & _
      '         " And (CP10 <> '703' And CP10 <> '704') and cp57 is null "
      '92.7.29 MODIFY BY SONIA 增加顯示發文日
      'strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp57") & " 取消收文日 " & _
      '         " from caseprogress,casepropertymap,staff,staff staff1 " & _
      '         " where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" & _
      '         " And (CP10 <> '703' And CP10 <> '704') and cp57 is null "
      strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp27") & " 發文日," & SQLDate("cp57") & " 取消收文日 " & _
               " from caseprogress,casepropertymap,staff,staff staff1 " & _
               " where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" & _
               " And (CP10 <> '703' And CP10 <> '704') and cp57 is null "
      '92.7.29 END
      '92.1.12 END
   '法務顧問或法務顧問服務
   Case Else
      '911024 nick 取消收文應該為 null ，只有取消收文會用到
      'strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp57") & " 取消收文日 " & _
               " from caseprogress,casepropertymap,staff,staff staff1 " & _
               " where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp27 is null and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" & _
               " And (CP10 <> '998' And CP10 <> '999')"
      '92.1.12 MODIFY BY SONIA 取消發文日的限制
      'strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp57") & " 取消收文日 " & _
      '         " from caseprogress,casepropertymap,staff,staff staff1 " & _
      '         " where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp27 is null and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" & _
      '         " And (CP10 <> '998' And CP10 <> '999') and cp57 is null "
      '92.7.29 MODIFY BY SONIA 增加顯示發文日
      'strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp57") & " 取消收文日 " & _
      '         " from caseprogress,casepropertymap,staff,staff staff1 " & _
      '         " where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" & _
      '         " And (CP10 <> '998' And CP10 <> '999') and cp57 is null "
      strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp27") & " 發文日," & SQLDate("cp57") & " 取消收文日 " & _
               " from caseprogress,casepropertymap,staff,staff staff1 " & _
               " where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")" & _
               " And (CP10 <> '998' And CP10 <> '999') and cp57 is null "
      '92.7.29 END
      '92.1.12 END
   End Select
Else
   '911024 nick 取消收文應該為 null ，只有取消收文會用到
   'strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp57") & " 取消收文日 from caseprogress,casepropertymap,staff,staff staff1 where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp27 is null and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")"
   '92.1.12 MODIFY BY SONIA 取消發文日的限制
   'strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp57") & " 取消收文日 from caseprogress,casepropertymap,staff,staff staff1 where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp27 is null and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")  and cp57 is null "
  '92.7.29 MODIFY BY SONIA 增加顯示發文日
   'strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp57") & " 取消收文日 from caseprogress,casepropertymap,staff,staff staff1 where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")  and cp57 is null "
   strSQL = "select cp09 收文號," & SQLDate("cp05") & " 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人," & SQLDate("cp06") & " 本所期限," & SQLDate("cp07") & " 法定期限," & SQLDate("cp27") & " 發文日," & SQLDate("cp57") & " 取消收文日 from caseprogress,casepropertymap,staff,staff staff1 where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")  and cp57 is null "
   '92.7.29 END
   '92.1.12 END
End If
If rs1.State <> adStateClosed Then rs1.Close
Set rs1 = Nothing

'strSQL = "select cp09 收文號,substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 收文日,decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) 案件性質,staff.st02 業務員,staff1.st02 承辦人,decode(cp06,null,'',substr(cp06,1,4)-1911||'/'||substr(cp06,5,2)||'/'||substr(cp06,7,2)) 本所期限,decode(cp07,null,'',substr(cp07,1,4)-1911||'/'||substr(cp07,5,2)||'/'||substr(cp07,7,2)) 法定期限,decode(cp57,null,'',substr(cp57,1,4)-1911||'/'||substr(cp57,5,2)||'/'||substr(cp57,7,2)) 取消收文日 from caseprogress,casepropertymap,staff,staff staff1 where cp01=cpm01 and cp10=cpm02 and cp13=staff.st01(+) and cp14=staff1.st01(+) and cp27 is null and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + " and substr(cp09,1,1) in (" + CNULL(接洽記錄單) + "," + CNULL(內部收文) + ")"

Set ReadCancelReceivedDayRst = objPublicData.ReadRst(strSQL)
Set objPublicData = Nothing
Exit Function
Err:
Set objPublicData = Nothing
ErrorLog
End Function
'判斷是否有子案或相關卷號
Public Function CheckChildCaseOrCaseRelation(ByRef field() As String) As Integer
Dim strSQL As String, rsRecordset As New ADODB.Recordset

On Error GoTo Err
CheckChildCaseOrCaseRelation = -1
Select Case CheckChildCase(field())
             Case 1
                        CheckChildCaseOrCaseRelation = 1
                        Exit Function
             Case -1
                        Exit Function
End Select
Select Case CheckCaseRelation(field())
             Case 1
                        CheckChildCaseOrCaseRelation = 2
                        Exit Function
             Case -1
                        Exit Function
End Select
CheckChildCaseOrCaseRelation = 0
Exit Function
Err:
CheckChildCaseOrCaseRelation = -2
ErrorLog
End Function
'判斷是否有相關卷號
Public Function CheckCaseRelation(ByRef field() As String) As Integer
Dim strSQL As String, rsRecordset As New ADODB.Recordset

On Error GoTo Err
strSQL = "select cr01 from caserelation where cr01=" + CNULL(field(1)) + " and cr02=" + CNULL(field(2)) + " and cr03=" + CNULL(field(3)) + " and cr04=" + CNULL(field(4))
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSQL, cnnConnection
If rsRecordset.RecordCount > 0 Then
   CheckCaseRelation = 1
Else
   CheckCaseRelation = 0
End If
Exit Function
Err:
CheckCaseRelation = -1
ErrorLog
End Function
'判斷是否有子案
Public Function CheckChildCase(ByRef field() As String) As Integer
Dim strSQL As String, rsRecordset As New ADODB.Recordset
Dim objPublicData As Object, intCaseKind As Integer

On Error GoTo Err
Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
If objPublicData.GetSystemKind(field(1), intCaseKind) Then
   Select Case intCaseKind
                Case 專利
                           strSQL = "select pa01 from patent where pa01=" + CNULL(field(1)) + " and pa02=" + CNULL(field(2)) + " and pa03<>'0'"
                Case 商標
                           strSQL = "select tm01 from trademark where tm27=" + CNULL(15)
                Case 法務
                           strSQL = "select lc01 from lawcase where lc01=" + CNULL(field(1)) + " and lc02=" + CNULL(field(2)) + " and lc03<>'0'"
                Case 顧問
                           strSQL = "select hc01 from hirecase where hc01=" + CNULL(field(1)) + " and hc02=" + CNULL(field(2)) + " and hc03<>'0'"
                Case Else
                           strSQL = "select sp01 from servicepractice where sp01=" + CNULL(field(1)) + " and sp02=" + CNULL(field(2)) + " and sp03<>'0'"
   End Select
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSQL, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      CheckChildCase = 1
   Else
      CheckChildCase = 0
   End If
End If
Set objPublicData = Nothing
Exit Function
Err:
CheckChildCase = -1
Set objPublicData = Nothing
ErrorLog
End Function
'讀取解除期限原因
Public Function ReadReasonOfRelief(ByRef strReasonNo() As String, ByRef strReasonName() As String) As Integer
Dim strSQL As String, rsRecordset As New ADODB.Recordset, i As Integer

On Error GoTo Err
strSQL = "select ror01,ror02 from reasonofrelief"
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSQL, cnnConnection
If rsRecordset.RecordCount > 0 Then
   Do While Not rsRecordset.EOF
         ReDim Preserve strReasonNo(i) As String
         ReDim Preserve strReasonName(i) As String
         strReasonNo(i) = rsRecordset.Fields(0)
         strReasonName(i) = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
         i = i + 1
         rsRecordset.MoveNext
   Loop
   ReadReasonOfRelief = 1
Else
   ReadReasonOfRelief = 0
End If
Exit Function
Err:
ShowMsg MsgText(8001)
ReadReasonOfRelief = -1
ErrorLog
End Function
'解除期限存檔
Public Function SaveRelieveDeadlineData(ByRef intCaseKind As Integer, ByRef intWhere As Integer, ByRef cp() As String, ByRef field() As String, ByRef strRowID As String, ByRef strClose As String, ByRef strDate1 As String, ByRef strDate2 As String, ByRef strMemo As String) As Boolean
Dim strReceiveCode As String, strTemp As String, strNextProgress(0) As String
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo err1
cnnConnection.BeginTrans
Dim objPublicData As Object
Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
If objPublicData.SaveAllData(cp(), field(), intCaseKind, intWhere) = False Then GoTo Err
If strClose = "Y" Then
   Dim strDataTemp(1 To T_CP) As String
   strDataTemp(1) = field(1)
   strDataTemp(2) = field(2)
   strDataTemp(3) = field(3)
   strDataTemp(4) = field(4)
   strDataTemp(5) = GetTodayDate
   strDataTemp(27) = GetTodayDate
   strDataTemp(9) = 內部收文 + GetTaiwanThisYear
   strDataTemp(10) = 不續辦
   strDataTemp(13) = cp(13)
   strDataTemp(12) = cp(12)
   strDataTemp(20) = "N"
   strDataTemp(32) = "N"
   strDataTemp(43) = cp(9)
   If objPublicData.SaveNewCaseProgressDatabase(內部收文, strDataTemp(), intWhere) = False Then GoTo Err
End If
strNextProgress(0) = strRowID
If objPublicData.SetNextProgressGoOn(strNextProgress(), True, False) = False Then GoTo Err
If objPublicData.SetNextProgressDate(strNextProgress(), strDate1, strDate2) = False Then GoTo Err
If strDate1 <> "" Then
   Dim np(1 To T_NP) As String
   If objPublicData.ReadNextProgressData(np(), strRowID) = False Then GoTo Err
   cp(9) = np(1)
   cp(13) = np(10)
   If objPublicData.SaveNewNextProgressDatabase(intWhere, cp(), strDate1, strDate2, np(7), strMemo, np(13), np(14)) = False Then GoTo Err
End If
Set objPublicData = Nothing
SaveRelieveDeadlineData = True
Err:
If SaveRelieveDeadlineData Then
        'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
Else
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.RollbackTrans
    End If
End If
Exit Function
err1:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
ErrorLog
End Function
'取消收文存檔
Public Function SaveCancelReceivedDayData(ByRef intCaseKind As Integer, ByRef intWhere As Integer, ByRef cp() As String, ByRef field() As String, ByRef strClose As String, ByRef strDate1 As String, ByRef strDate2 As String, ByRef strMemo As String) As Boolean
Dim strReceiveCode As String, strTemp As String, strNextProgress(0) As String
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo err1
cnnConnection.BeginTrans
Dim objPublicData As Object
Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
If objPublicData.SaveAllData(cp(), field(), intCaseKind, intWhere) = False Then GoTo Err
If strClose = "Y" Then
   Dim strDataTemp(1 To T_CP) As String
   strDataTemp(1) = field(1)
   strDataTemp(2) = field(2)
   strDataTemp(3) = field(3)
   strDataTemp(4) = field(4)
   strDataTemp(5) = GetTodayDate
   strDataTemp(27) = GetTodayDate
   strDataTemp(9) = 內部收文 + GetTaiwanThisYear
   strDataTemp(10) = 不續辦
   strDataTemp(13) = cp(13)
   strDataTemp(12) = cp(12)
   strDataTemp(20) = "N"
   strDataTemp(32) = "N"
   strDataTemp(43) = cp(9)
   If objPublicData.SaveNewCaseProgressDatabase(內部收文, strDataTemp(), intWhere) = False Then GoTo Err
End If
If strDate1 <> "" Then
   If cp(40) <> "" Then
      strTemp = cp(40)
   ElseIf cp(41) <> "" Then
      strTemp = cp(41)
   Else
      strTemp = cp(42)
   End If
   If objPublicData.SaveNewNextProgressDatabase(intWhere, cp(), strDate1, strDate2, cp(10), strMemo, , strTemp) = False Then GoTo Err
End If
Set objPublicData = Nothing
SaveCancelReceivedDayData = True
Err:
If SaveCancelReceivedDayData Then
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
Else
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.RollbackTrans
    End If
End If
Exit Function
err1:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
ErrorLog
End Function
'讀取閉卷申請人或代理人資料
Public Function ReadCloseCaseRst(ByRef strNo As String, ByRef strSystemKind As String) As ADODB.Recordset
Dim strSQL As String, rsRecordset As New ADODB.Recordset
Dim objPublicData As Object, strSQLS(4) As String, strSystem(4) As String, i As Integer

On Error GoTo Err
Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
If objPublicData.AnalysisSystemKindString(strSystemKind, strSystem()) = False Then GoTo err1
strNo = ChangeCustomerL(strNo)
If Left(strNo, 1) = 代理人編號 Then
   strNo = Left(strNo, 8)
End If
If strSystem(0) <> "" Then
   strSQLS(0) = "select pa01 s01,pa02 s02,pa02 s03,replace(pa03,'0') s04,replace(pa04,'00') s05,nvl(pa05,nvl(pa06,pa07)) s06,na03 s07,pa11 s08,pa01 s09,pa02 s10,pa03 s11,pa04 s12 from patent,nation where pa09=na01 "
   If Left(strNo, 1) = 客戶編號 Then
      strSQLS(0) = strSQLS(0) + "and (pa26=" + CNULL(strNo) + " or pa27=" + CNULL(strNo) + " or pa28=" + CNULL(strNo) + " or pa29=" + CNULL(strNo) + " or pa30=" + CNULL(strNo) + ")"
   Else
      strSQLS(0) = strSQLS(0) + "and pa75=" + CNULL(strNo)
   End If
   strSQLS(0) = strSQLS(0) + " and pa01 in (" + strSystem(0) + ")"
End If
If strSystem(1) <> "" Then
   strSQLS(1) = "select tm01 s01,decode(tm01," + CNULL(馬德里案) + ",substr(tm02,1,5),tm02) s02,decode(tm01," + CNULL(馬德里案) + ",replace(substr(tm02,6,1),'0'),tm02) s03,replace(tm03,'0') s04,replace(tm04,'00') s05,nvl(tm05,nvl(tm06,tm07)) s06,na03 s07,tm12 s08,tm01 s09,tm02 s10,tm03 s11,tm04 s12 from trademark,nation where tm10=na01 "
   If Left(strNo, 1) = 客戶編號 Then
      strSQLS(1) = strSQLS(1) + "and tm23=" + CNULL(strNo)
   Else
      strSQLS(1) = strSQLS(1) + "and tm44=" + CNULL(strNo)
   End If
   strSQLS(1) = strSQLS(1) + " and tm01 in (" + strSystem(1) + ")"
End If
If strSystem(2) <> "" Then
   strSQLS(2) = "select lc01 s01,lc02 s02,lc02 s03,replace(lc03,'0') s04,replace(lc04,'00') s05,nvl(lc05,nvl(lc06,lc07)) s06,na03 s07,'' s08,lc01 s09,lc02 s10,lc03 s11,lc04 s12 from lawcase,nation where lc15=na01 "
   If Left(strNo, 1) = 客戶編號 Then
      strSQLS(2) = strSQLS(2) + "and lc11=" + CNULL(strNo)
   Else
      strSQLS(2) = strSQLS(2) + "and lc22=" + CNULL(strNo)
   End If
   strSQLS(2) = strSQLS(2) + " and lc01 in (" + strSystem(2) + ")"
End If
If strSystem(3) <> "" Then
   strSQLS(3) = "select hc01 s01,hc02 s02,hc02 s03,replace(hc03,'0') s04,replace(hc04,'00') s05,hc06 s06,'' s07,'' s08,hc01 s09,hc02 s10,hc03 s11,hc04 s12 from hirecase where hc05=" + CNULL(strNo) + " and hc01 in (" + strSystem(3) + ")"
End If
If strSystem(4) <> "" Then
   strSQLS(4) = "select sp01 s01,sp02 s02,sp02 s03,replace(sp03,'0') s04,replace(sp04,'00') s05,nvl(sp05,nvl(sp06,sp07)) s06,na03 s07,sp11 s08,sp01 s09,sp02 s10,sp03 s11,sp04 s12 from servicepractice,nation where sp09=na01 "
   If Left(strNo, 1) = 客戶編號 Then
      strSQLS(4) = strSQLS(4) + "and (sp08=" + CNULL(strNo) + " or sp58=" + CNULL(strNo) + " or sp59=" + CNULL(strNo) + ")"
   Else
      strSQLS(4) = strSQLS(4) + "and sp26=" + CNULL(strNo)
   End If
   strSQLS(4) = strSQLS(4) + " and sp01 in (" + strSystem(4) + ")"
End If
strSQL = "select '' ˇ,s01 本所案號,s02 本所案號,s03 本所案號,s04 本所案號,s05 本所案號,s06 案件名稱,s07 申請國家,s08 申請案號,s09,s10,s11,s12 from ("
For i = 0 To 4
       If strSQLS(i) <> "" Then
          strSQL = strSQL + strSQLS(i) + " union "
       End If
Next
strSQL = Left(strSQL, Len(strSQL) - 7)
strSQL = strSQL + ") order by s09,s10,s11,s12"
Set ReadCloseCaseRst = objPublicData.ReadRst(strSQL)
err1:
Set objPublicData = Nothing
Exit Function
Err:
Set objPublicData = Nothing
ErrorLog
End Function
'讀取閉卷之下一程序及案件進度檔資料
Public Function ReadCloseCaseDRst(ByRef intWhere As Integer, ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String) As ADODB.Recordset
Dim strSQL As String, rsRecordset As New ADODB.Recordset
Dim objPublicData As Object, strSQL1 As String, strSQL2 As String, i As Integer

On Error GoTo Err
Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
strSQL1 = "select decode(np07," + CNULL(大陸國家代號) + ",cpm04,cpm03) s01," & SQLDate("np08") & " s02," & SQLDate("np09") & " s03,np13 s04,np14 s05,'' s06 from nextprogress,casepropertymap where np02=cpm01 and np07=cpm02 and np02=" + CNULL(strCode1) + " and np03=" + CNULL(strCode2) + " and np04=" + CNULL(strCode3) + " and np05=" + CNULL(strCode4)
strSQL2 = "select decode(cp10," + CNULL(大陸國家代號) + ",cpm04,cpm03) s01," & SQLDate("cp06") & " s02," & SQLDate("cp07") & " s03,cp08 s04,nvl(cp40,nvl(cp50,cp55)) s05,cp09 s06 from caseprogress,casepropertymap where cp01=cpm01 and cp10=cpm02 and cp01=" + CNULL(strCode1) + " and cp02=" + CNULL(strCode2) + " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4)
strSQL = "select s01 案件性質,s02 本所期限,s03 法定期限,s04 機關文號,s05 相關人,s06 收文號 from (" + strSQL1 + " union " + strSQL2 + ") order by s02,s06"
Set ReadCloseCaseDRst = objPublicData.ReadRst(strSQL)
err1:
Set objPublicData = Nothing
Exit Function
Err:
Set objPublicData = Nothing
ErrorLog
End Function
'閉卷存檔
Public Function SaveCloseCaseData(ByRef intCaseKind As Integer, ByRef intWhere As Integer, ByRef cp() As String, ByRef field() As String) As Boolean
Dim strReceiveCode As String, strTemp As String, strNextProgress(0) As String
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo err1
cnnConnection.BeginTrans
Dim objPublicData As Object
Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
If objPublicData.SaveAllData(cp(), field(), intCaseKind, intWhere) = False Then GoTo Err
Dim strDataTemp(1 To T_CP) As String
strDataTemp(1) = field(1)
strDataTemp(2) = field(2)
strDataTemp(3) = field(3)
strDataTemp(4) = field(4)
strDataTemp(5) = GetTodayDate
strDataTemp(27) = GetTodayDate
strDataTemp(9) = 內部收文 + GetTaiwanThisYear
strDataTemp(10) = 不續辦
strDataTemp(20) = "N"
strDataTemp(32) = "N"
If objPublicData.SaveNewCaseProgressDatabase(內部收文, strDataTemp(), intWhere) = False Then GoTo Err
Set objPublicData = Nothing
SaveCloseCaseData = True
Err:
If SaveCloseCaseData Then
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.CommitTrans
    End If
Else
    'Modify By Cheng 2002/11/14
    If BolTransOk Then
        cnnConnection.RollbackTrans
    End If
End If
Exit Function
err1:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
ErrorLog
End Function
'讀取相關卷號檔
Public Function ReadCaseRelationData(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String, ByRef strRelation() As String) As Integer
Dim strSQL As String, rsRecordset As New ADODB.Recordset, i As Integer, j As Integer
On Error GoTo Err
   strSQL = "select cr05,cr06,cr07,cr08 from caserelation where cr01=" + CNULL(strCode1) + " and cr02=" + CNULL(strCode2) + " and cr03=" + CNULL(strCode3) + " and cr04=" + CNULL(strCode4)
   rsRecordset.Open strSQL, cnnConnection
   ReadCaseRelationData = 0
   Do While Not rsRecordset.EOF
      ReadCaseRelationData = 1
         ReDim Preserve strRelation(3, j)
         For i = 0 To 3
                strRelation(i, j) = rsRecordset.Fields(i)
         Next
         rsRecordset.MoveNext
         j = j + 1
   Loop
   Exit Function
Err:
ReadCaseRelationData = -1
ErrorLog
End Function
'相關卷號檔存檔
Public Function SaveCaseRelationData(ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef strCode4 As String, ByRef strRelation() As String) As Boolean
Dim strSQL As String, rsRecordset As New ADODB.Recordset, i As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans
strSQL = "delete from caserelation where cr01=" + CNULL(strCode1) + " and cr02=" + CNULL(strCode2) + " and cr03=" + CNULL(strCode3) + " and cr04=" + CNULL(strCode4)
cnnConnection.Execute strSQL
For i = 0 To UBound(strRelation, 2)
       strSQL = "insert into caserelation values(" + CNULL(strCode1) + "," + CNULL(strCode2) + "," + CNULL(strCode3) + "," + CNULL(strCode4) + _
             "," + CNULL(strRelation(0, i)) + "," + CNULL(strRelation(1, i)) + "," + CNULL(strRelation(2, i)) + "," + CNULL(strRelation(3, i)) + ")"
       cnnConnection.Execute strSQL
Next
For i = 0 To UBound(strRelation, 2)
       strSQL = "delete from caserelation where cr01=" + CNULL(strRelation(0, i)) + " and cr02=" + CNULL(strRelation(1, i)) + " and cr03=" + CNULL(strRelation(2, i)) + " and cr04=" + CNULL(strRelation(3, i)) + " and cr05=" + CNULL(strCode1) + " and cr06=" + CNULL(strCode2) + " and cr07=" + CNULL(strCode3) + " and cr08=" + CNULL(strCode4)
       cnnConnection.Execute strSQL
Next
For i = 0 To UBound(strRelation, 2)
       strSQL = "insert into caserelation values(" + CNULL(strRelation(0, i)) + "," + CNULL(strRelation(1, i)) + "," + CNULL(strRelation(2, i)) + "," + CNULL(strRelation(3, i)) + "," + _
            CNULL(strCode1) + "," + CNULL(strCode2) + "," + CNULL(strCode3) + "," + CNULL(strCode4) + ")"
       cnnConnection.Execute strSQL
Next
SaveCaseRelationData = True
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
ErrorLog
End Function
Public Property Set Connection(ByVal vNewValue As ADODB.Connection)
Set cnnConnection = vNewValue
End Property
