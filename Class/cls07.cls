VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls07"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim strSQL As String
Public Function GetRs001(ByRef i As String, ByRef sstr As String, ByRef str1 As String, ByRef str2 As String, ByRef str3 As String, _
       ByRef hstr As String, ByRef hstr1 As String, ByRef hstr2 As String, ByRef ustr3 As String) As ADODB.Recordset
On Error GoTo err
Select Case i
Case 1
 strSQL = " SELECT ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日," + _
     " cp09 AS 收文號,decode(cp01 || CP10, cpm01 || CPM02, CPM03, CPM04) AS 案件性質,CP01  AS 本所案號," + _
     " CP02 AS 本所案號,CP03 AS 本所案號,CP04 AS 本所案號, decode(LC11, CU01 || CU02," + _
     " NVL(CU04, NVL(CU05, CU06))) AS 當事人,decode(CP13,S1.ST01,S1.ST02) AS 業務員," + _
     " decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(CP29,S3.ST01,S3.ST02) AS 法務人員" + _
     " FROM " + sstr + ",STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,CUSTOMER" + _
     " WHERE " + ustr3 + "" + str1 + "" + str2 + "" + str3 + "  AND  lc01||lc02||lc03||lc04=cp01||cp02||cp03||cp04 " + _
     " and substr(lc11,1,8)=cu01(+) and substr(lc11,9,9)=cu02(+) AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND " + _
     " cp01=cpm01(+) and cp10=cpm02(+) union SELECT ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || " + _
     " SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日, cp09 AS 收文號," + _
     " decode(cp01 || CP10, cpm01 || CPM02, CPM03, CPM04) AS 案件性質,CP01  AS 本所案號," + _
     " CP02 AS 本所案號,CP03 AS 本所案號,CP04 AS 本所案號, decode(HC05, CU01 || CU02," + _
     " NVL(CU04, NVL(CU05, CU06))) AS 當事人,decode(CP13,S1.ST01,S1.ST02) AS 業務員," + _
     " decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(CP29,S3.ST01,S3.ST02) AS 法務人員" + _
     " from " + hstr + ",STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,CUSTOMER" + _
     " where " + ustr3 + "" + hstr1 + "" + hstr2 + "" + str3 + " AND  hc01||hc02||hc03||hc04=cp01||cp02||cp03||cp04" + _
     " and substr(hc05,1,8)=cu01(+) and substr(hc05,9,9)=cu02(+) AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and" + _
     " cp29 = s3.st01(+) AND cp01=cpm01(+) and cp10=cpm02(+) "
     
Case 2
  strSQL = " SELECT ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日," + _
     " cp09 AS 收文號,decode(cp01 || CP10, cpm01 || CPM02, CPM03, CPM04) AS 案件性質,CP01  AS 本所案號," + _
     " CP02 AS 本所案號,CP03 AS 本所案號,CP04 AS 本所案號, decode(LC11, CU01 || CU02," + _
     " NVL(CU04, NVL(CU05, CU06))) AS 當事人,decode(CP13,S1.ST01,S1.ST02) AS 業務員," + _
     " decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(CP29,S3.ST01,S3.ST02) AS 法務人員" + _
     " from " + sstr + ",STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,CUSTOMER" + _
     " WHERE " + ustr3 + "" + str1 + "" + str2 + "" + str3 + "  AND substr(lc11,1,8)=cu01(+) and " + _
     " substr(lc11,9,9)=cu02(+)  AND cp13 = s1.st01(+)  AND cp01||cp02||cp03||cp04=lc01||lc02||lc03||lc04 " + _
     " and cp14 = s2.st01(+) and cp29 = s3.st01(+) AND cp01=cpm01(+) and cp10=cpm02(+)  "

End Select
Dim objPublicData As New clsPublicData
Set GetRs001 = objPublicData.ReadRstD(strSQL)
Set objPublicData = Nothing
Exit Function

err:
ErrorLog
End Function
Public Function GetData002(ByRef i As Integer, ByRef strtxt As String, ByRef cp01 As String, ByRef cp02 As String, ByRef cp03 As String, _
       ByRef cp04 As String, ByRef cp05 As String, ByRef cp06 As String, ByRef cp07 As String, ByRef cp09 As String, _
       ByRef cp10 As String, ByRef cp13 As String, ByRef cp14 As String, ByRef cp16 As String, ByRef cp18 As String, _
       ByRef cp19 As String, ByRef cp21 As String, ByRef cp26 As String, ByRef cp29 As String, ByRef cp43 As String, ByRef cp49 As String, _
       ByRef cp57 As String, ByRef cp64 As String, ByRef lc05 As String, ByRef lc06 As String, ByRef lc07 As String, _
       ByRef lc08 As String, ByRef lc11 As String, ByRef lc13 As String, ByRef lc16 As String, ByRef lc27 As String) As Boolean
       
Dim rs As New ADODB.Recordset

On Error GoTo err
Select Case i
Case 0
    strSQL = "select cp01, cp02, cp03, cp04, cp05, cp06, cp07,cp09 , cp10, cp13, cp14, cp16, cp18, cp19,cp21, cp26," + _
        " cp29, cp43, cp49, cp57, cp64, lc05, lc06, lc07, lc08, lc11, lc13, lc16, lc27 from lawcase, caseprogress where" + _
        " lc01||lc02||lc03||lc04=cp01||cp02||cp03||cp04 and cp09='" + strtxt + "' order by lc01||lc02||lc03||lc04"
Case 1
    strSQL = "select cp01, cp02, cp03, cp04, cp05, cp06, cp07,cp09 , cp10, cp13, cp14, cp16, cp18, cp19,cp21, cp26," + _
        " cp29, cp43, cp49, cp57, cp64, hc05,hc06,hc07,hc09,hc12 from hirecase, caseprogress where" + _
        " hc01||hc02||hc03||hc04=cp01||cp02||cp03||cp04 and cp09='" + strtxt + "' order by  hc01||hc02||hc03||hc04"
End Select

rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.Open strSQL, cnnConnection
If Not (rs.BOF And rs.EOF) Then

Select Case i
Case 0
    lc05 = IIf(IsNull(rs.Fields!lc05), "", rs.Fields!lc05)
    lc06 = IIf(IsNull(rs.Fields!lc06), "", rs.Fields!lc06)
    lc07 = IIf(IsNull(rs.Fields!lc07), "", rs.Fields!lc07)
    lc08 = IIf(IsNull(rs.Fields!lc08), "", rs.Fields!lc08)
    lc11 = IIf(IsNull(rs.Fields!lc11), "", ChangeCustomerS(rs.Fields!lc11))
    lc13 = IIf(IsNull(rs.Fields!lc13), "", rs.Fields!lc13)
    lc16 = IIf(IsNull(rs.Fields!lc16), "", rs.Fields!lc16)
    lc27 = IIf(IsNull(rs.Fields!lc27), "", rs.Fields!lc27)
Case 1
    lc11 = IIf(IsNull(rs.Fields!hc05), "", ChangeCustomerS(rs.Fields!hc05)) 'lc11=>hc05
    lc05 = IIf(IsNull(rs.Fields!hc06), "", rs.Fields!hc06)   'lc05=>hc06
    lc16 = IIf(IsNull(rs.Fields!hc07), "", rs.Fields!hc07)  'lc16=>hc07
    lc08 = IIf(IsNull(rs.Fields!hc09), "", rs.Fields!hc09) 'lc08=>hc09
    lc27 = IIf(IsNull(rs.Fields!hc12), "", rs.Fields!hc12) 'lc27=>hc12
End Select
    
    cp01 = IIf(IsNull(rs.Fields!cp01), "", rs.Fields!cp01)
    cp02 = IIf(IsNull(rs.Fields!cp02), "", rs.Fields!cp02)
    cp03 = IIf(IsNull(rs.Fields!cp03), "", rs.Fields!cp03)
    cp04 = IIf(IsNull(rs.Fields!cp04), "", rs.Fields!cp04)
    cp05 = IIf(IsNull(rs.Fields!cp05), "", rs.Fields!cp05)
    cp06 = IIf(IsNull(rs.Fields!cp06), "", rs.Fields!cp06)
    cp07 = IIf(IsNull(rs.Fields!cp07), "", rs.Fields!cp07)
    cp09 = IIf(IsNull(rs.Fields!cp09), "", rs.Fields!cp09)
    cp10 = IIf(IsNull(rs.Fields!cp10), "", rs.Fields!cp10)
    cp13 = IIf(IsNull(rs.Fields!cp13), "", rs.Fields!cp13)
    cp14 = IIf(IsNull(rs.Fields!cp14), "", rs.Fields!cp14)
    cp16 = IIf(IsNull(rs.Fields!cp16), "", rs.Fields!cp16)
    cp18 = IIf(IsNull(rs.Fields!cp18), "", rs.Fields!cp18)
    cp19 = IIf(IsNull(rs.Fields!cp19), "", rs.Fields!cp19)
    cp26 = IIf(IsNull(rs.Fields!cp26), "", rs.Fields!cp26)
    cp29 = IIf(IsNull(rs.Fields!cp29), "", rs.Fields!cp29)
    cp43 = IIf(IsNull(rs.Fields!cp43), "", rs.Fields!cp43)
    cp49 = IIf(IsNull(rs.Fields!cp49), "", rs.Fields!cp49)
    cp57 = IIf(IsNull(rs.Fields!cp57), "", rs.Fields!cp57)
    cp64 = IIf(IsNull(rs.Fields!cp64), "", rs.Fields!cp64)
GetData002 = True

End If
If rs.RecordCount < 1 Then MsgBox "資料庫無資料", vbCritical

Exit Function
err:
GetData002 = False
ErrorLog
End Function
Public Function GetRs002(ByRef strlc As String) As ADODB.Recordset
On Error GoTo err
Dim rs As New ADODB.Recordset

  strSQL = "select  decode(np02 || np07, cpm01 || CPM02, CPM03, CPM04)下一程序," + _
                  " decode(np08,null,'',SUBSTR(np08, 1, 4)- 1911 || '/' || SUBSTR(np08, 5, 2)|| '/' || SUBSTR(np08, 7, 2)) 本所期限," + _
                  " decode(np09,null,'',SUBSTR(np09, 1, 4)- 1911 || '/' || SUBSTR(np09, 5, 2)|| '/' || SUBSTR(np09, 7, 2))  法定期限," + _
                  " np13 機關文號,np14 相關人, decode(np11,null,'',SUBSTR(np11, 1, 4)- 1911 || '/' || SUBSTR(np11, 5, 2)|| '/' || " + _
                  " SUBSTR(np11, 7, 2)) 解除期限日期,np06 ,np01,np07,np16,np17,np18 from nextprogress,CASEPROPERTYMAP where " + _
                  " np02||np03||np04||np05=" + CNULL(strlc) + " and np02=cpm01(+) and np07=cpm02(+) and (np06='N' or np06 is null)"


rs.CursorLocation = adUseClient
rs.CursorType = adOpenDynamic
rs.Open strSQL, cnnConnection
Set GetRs002 = rs
Set rs = Nothing

Exit Function
err:
ErrorLog
End Function
Public Function Update002(ByRef i As Integer, ByRef blnIsChange As Boolean, _
     ByRef cp09 As String, ByRef cp05 As String, ByRef lc11 As String, ByRef lc05 As String, _
     ByRef lc06 As String, ByRef lc07 As String, ByRef lc13 As String, ByRef lc16 As String, _
     ByRef cp14 As String, ByRef cp29 As String, ByRef cp10 As String, ByRef cp13 As String, _
     ByRef cp06 As String, ByRef cp07 As String, ByRef cp21 As String, ByRef cp26 As String, _
     ByRef cp43 As String, ByRef cp49 As String, ByRef lc01 As String, ByRef lc02 As String, _
     ByRef lc03 As String, ByRef lc04 As String, ByRef lc08 As String, ByRef cp64 As String, _
     ByRef lc27 As String) As Boolean
Dim strlc As String
On Error GoTo err
cnnConnection.BeginTrans
  strlc = lc01 + lc02 + lc03 + lc04
 
 Select Case i
Case 0
   strSQL = "update lawcase set lc05=" + CNULL(lc05) + ",lc06=" + CNULL(lc06) + ",lc07=" + CNULL(lc07) + _
       ", lc08 = " + CNULL(lc08) + ",lc11=" + CNULL(ChangeCustomerL(lc11)) + ",lc13=" + CNULL(lc13) + ",lc16=" + CNULL(lc16) + _
       ", lc27=" + CNULL(lc27) + " where lc01||lc02||lc03||lc04=" + CNULL(strlc) + ""
 Case 1
   strSQL = "update hirecase set hc05=" + CNULL(ChangeCustomerL(lc11)) + ",hc06=" + CNULL(lc05) + ",hc07=" + CNULL(lc16) + _
                   " ,hc09=" + CNULL(lc08) + " ,hc12=" + CNULL(lc27) + " where hc01||hc02||hc03||hc04=" + CNULL(strlc) + ""
 End Select
cnnConnection.Execute strSQL

If blnIsChange Then
        strSQL = "update caseprogress set cp01=" + CNULL(lc01) + ",cp02=" + CNULL(lc02) + ",cp03=" + CNULL(lc03) + ",cp04=" + CNULL(lc04) + _
        ",cp05=" + CNULL(cp05) + ",cp06=" + CNULL(cp06) + ", cp07=" + CNULL(cp07) + ",cp10=" + CNULL(cp10) + _
        ",cp13=" + CNULL(cp13) + ",cp14=" + CNULL(cp14) + ",cp21=" + CNULL(cp21) + ",cp26=" + CNULL(cp26) + _
        ", cp29=" + CNULL(cp29) + ",cp43=" + CNULL(cp43) + ",cp64=" + CNULL(cp64) + ",cp49=" + CNULL(cp49) + " where cp09=" + CNULL(cp09) + ""
        cnnConnection.Execute strSQL
Else
        strSQL = "update caseprogress set cp05=" + CNULL(cp05) + ",cp06=" + CNULL(cp06) + ", cp07=" + CNULL(cp07) + ",cp10=" + CNULL(cp10) + _
            ",cp13=" + CNULL(cp13) + ",cp14=" + CNULL(cp14) + ",cp21=" + CNULL(cp21) + ",cp26=" + CNULL(cp26) + ", cp29=" + CNULL(cp29) + _
            ",cp43=" + CNULL(cp43) + ",cp49=" + CNULL(cp49) + ", cp64=" + CNULL(cp64) + " where cp09=" + CNULL(cp09) + ""
         cnnConnection.Execute strSQL
End If
    
 
cnnConnection.CommitTrans
Update002 = True

Exit Function
err:
Update002 = False
cnnConnection.RollbackTrans
ErrorLog
End Function
Public Function IsExistCaseNum(ByRef strlc As String, ByRef lc05 As String, ByRef lc06 As String, ByRef lc07 As String, _
                  ByRef lc11 As String) As Boolean
Dim strSQL As String, rsRecordset As New ADODB.Recordset

On Error GoTo err
strSQL = "select lc05,lc06,lc07,lc11 from lawcase where lc01||lc02||lc03||lc04=" + CNULL(strlc) + " "
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSQL, cnnConnection
If rsRecordset.RecordCount > 0 Then
   lc05 = IIf(IsNull(rsRecordset.Fields!lc05), "", rsRecordset.Fields!lc05)
   lc06 = IIf(IsNull(rsRecordset.Fields!lc06), "", rsRecordset.Fields!lc06)
   lc07 = IIf(IsNull(rsRecordset.Fields!lc07), "", rsRecordset.Fields!lc07)
   lc11 = IIf(IsNull(rsRecordset.Fields!lc11), "", rsRecordset.Fields!lc11)
   IsExistCaseNum = True
Else
   MsgBox "此本所案號不存在", vbCritical
   IsExistCaseNum = False
End If
rsRecordset.Close
Exit Function
err:
ErrorLog
End Function
Public Function UpdateNextProgress(ByRef strlc As String, ByRef np01 As String, ByRef np06 As String, _
          ByRef np07 As String, ByRef np08 As String, ByRef np16 As String, ByRef np17 As String, ByRef np18 As String) As Boolean
On Error GoTo err
  strSQL = "update nextprogress set np06=" + CNULL(np06) + " where np01=" + CNULL(np01) + " and np02||np03||np04||np05=" + CNULL(strlc) + _
       " and np07=" + CNULL(np07) + " and np08=" + CNULL(np08) + " and np16=" + CNULL(np16) + " and np17=" + CNULL(np17) + " and np18=" + CNULL(np18) + " and np06 is null "

cnnConnection.Execute strSQL
UpdateNextProgress = True
Exit Function
err:
UpdateNextProgress = False
ErrorLog
End Function
Public Function GetData0022(ByRef strlc As String, ByRef cp09 As String, ByRef cp40 As String, ByRef cp41 As String, ByRef cp42 As String) As Boolean
Dim rs As New ADODB.Recordset
On Error GoTo err
strSQL = "select cp40,cp41,cp42  from caseprogress where" + _
    " cp01||cp02||cp03||cp04=" + CNULL(strlc) + " and cp09=" + CNULL(cp09) + ""
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.Open strSQL, cnnConnection
If rs.RecordCount > 0 Then
   cp40 = IIf(IsNull(rs.Fields!cp40), "", rs.Fields!cp40)
   cp41 = IIf(IsNull(rs.Fields!cp41), "", rs.Fields!cp41)
   cp42 = IIf(IsNull(rs.Fields!cp42), "", rs.Fields!cp42)
End If
 GetData0022 = True
 Exit Function
err:
ErrorLog
End Function
Public Function UpdateData0022(ByRef strlc As String, ByRef cp09 As String, ByRef cp40 As String, ByRef cp41 As String, ByRef cp42 As String) As Boolean
Dim rs As New ADODB.Recordset
On Error GoTo err
strSQL = "update caseprogress set cp40=" + CNULL(cp40) + ",cp41=" + CNULL(cp41) + _
            ",cp42=" + CNULL(cp42) + "  where cp01||cp02||cp03||cp04=" + CNULL(strlc) + " and cp09=" + CNULL(cp09) + ""
cnnConnection.Execute strSQL
UpdateData0022 = True
Exit Function
err:
UpdateData0022 = False
ErrorLog
End Function
Public Function GetRs0023(ByRef strlc As String, ByRef cp09 As String) As ADODB.Recordset
On Error GoTo err
If Left(strlc, 2) <> "LA" Then
strSQL = "select '' v,cp09 收文號,decode(cp05,null,'',SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2))" + _
                 " AS 收文日, decode(lc15,020 , CPM03, CPM04) AS 案件性質, decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 " + _
                " || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)) as 發文日,cp19 後金 ,decode(cp24,1,'淮/勝',2,'駁/敗') as 結果 ,cp50 相關人 from caseprogress," + _
                " CASEPROPERTYMAP ,lawcase where cp01||cp02||cp03||cp04=" + CNULL(strlc) + " and cp01=cpm01 and cp10=cpm02  " + _
                " and cp01||cp02||cp03||cp04=lc01||lc02||lc03||lc04 and cp09<>" + CNULL(cp09) + ""
Else
strSQL = "select '' v,cp09 收文號,decode(cp05,null,'',SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2))" + _
                 " AS 收文日,  CPM03 AS 案件性質, decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 " + _
                " || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)) as 發文日,cp19 後金 ,decode(cp24,1,'淮/勝',2,'駁/敗') as 結果 ,cp50 相關人 from caseprogress," + _
                " CASEPROPERTYMAP ,hirecase  where cp01||cp02||cp03||cp04=" + CNULL(strlc) + " and cp01=cpm01 and cp10=cpm02 " + _
                " and cp01||cp02||cp03||cp04=hc01||hc02||hc03||hc04 and cp09<>" + CNULL(cp09) + ""
End If
Dim objPublicData As New clsPublicData
Set GetRs0023 = objPublicData.ReadRstD(strSQL)
Set objPublicData = Nothing
Exit Function

err:
ErrorLog
End Function
Public Function GetData003(ByRef strtxt As String, ByRef cp01 As String, ByRef cp02 As String, ByRef cp03 As String, _
     ByRef cp04 As String, ByRef cp05 As String, ByRef cp09 As String, ByRef cp10 As String, ByRef cp13 As String, _
     ByRef cp16 As String, ByRef cp18 As String, ByRef cp53 As String, ByRef cp54 As String, ByRef hc05 As String, ByRef hc06 As String, ByRef hc07 As String, _
     ByRef hc12 As String) As Boolean

Dim rs As New ADODB.Recordset
On Error GoTo err
strSQL = "select cp01, cp02, cp03, cp04, cp05,cp09,cp10, cp13," + _
     " cp16 , cp18, cp53, cp54,hc05, hc06, hc07, hc12 from caseprogress,hirecase" + _
     " where cp09=" + CNULL(strtxt) + " and cp01||cp02||cp03||cp04=hc01||hc02||hc03||hc04"
     
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.Open strSQL, cnnConnection
If Not (rs.BOF And rs.EOF) Then
    cp01 = IIf(IsNull(rs.Fields!cp01), "", rs.Fields!cp01)
    cp02 = IIf(IsNull(rs.Fields!cp02), "", rs.Fields!cp02)
    cp03 = IIf(IsNull(rs.Fields!cp03), "", rs.Fields!cp03)
    cp04 = IIf(IsNull(rs.Fields!cp04), "", rs.Fields!cp04)
    cp05 = IIf(IsNull(rs.Fields!cp05), "", rs.Fields!cp05)
    cp10 = IIf(IsNull(rs.Fields!cp10), "", rs.Fields!cp10)
    cp13 = IIf(IsNull(rs.Fields!cp13), "", rs.Fields!cp13)
    cp16 = IIf(IsNull(rs.Fields!cp16), "", rs.Fields!cp16)
    cp18 = IIf(IsNull(rs.Fields!cp18), "", rs.Fields!cp18)
    cp53 = IIf(IsNull(rs.Fields!cp53), "", rs.Fields!cp53)
    cp54 = IIf(IsNull(rs.Fields!cp54), "", rs.Fields!cp54)
    hc05 = IIf(IsNull(rs.Fields!hc05), "", ChangeCustomerS(rs.Fields!hc05))
    hc06 = IIf(IsNull(rs.Fields!hc06), "", rs.Fields!hc06)
    hc07 = IIf(IsNull(rs.Fields!hc07), "", rs.Fields!hc07)
    hc12 = IIf(IsNull(rs.Fields!hc12), "", rs.Fields!hc12)

End If
GetData003 = True
Exit Function
err:
GetData003 = False
ErrorLog

End Function
Public Function GetRs003(ByRef strCaseNum As String, ByRef strnum As String, _
       ByRef strproperty As String, ByRef cp53 As String, ByRef cp54 As String) As Boolean

Dim rs As New ADODB.Recordset, i As Integer, j As Integer
On Error GoTo err
  
  strSQL = "select cp53,cp54,cp09 from caseprogress, casepropertymap where cp01||cp02||cp03||cp04=" + CNULL(strCaseNum) + _
           " and cpm03=" + CNULL(strproperty) + "and  cpm01||cpm02=cp01||cp10 order by cp05 "

rs.CursorLocation = adUseClient
rs.CursorType = adOpenDynamic
rs.Open strSQL, cnnConnection
If rs.RecordCount > 0 Then
   rs.Find "cp09='" + strnum + "'", 0, adSearchForward
   i = rs.AbsolutePosition
End If
If i > 1 Then
   rs.AbsolutePosition = i - 1
   
   cp53 = IIf(IsNull(rs.Fields!cp53), "", rs.Fields!cp53)
   cp54 = IIf(IsNull(rs.Fields!cp54), "", rs.Fields!cp54)
   GetRs003 = True
   
End If
Set rs = Nothing
Exit Function
err:
GetRs003 = False
ErrorLog

End Function
Public Function Update003(ByRef strlc As String, ByRef hc05 As String, ByRef hc06 As String, ByRef hc07 As String, _
        ByRef hc12 As String, ByRef cp05 As String, ByRef cp09 As String, ByRef cp13 As String, ByRef cp53 As String, _
        ByRef cp54 As String) As Boolean
        
On Error GoTo err
cnnConnection.BeginTrans
strSQL = "update hirecase set hc05=" + CNULL(ChangeCustomerL(hc05)) + ",hc06=" + CNULL(hc06) + _
           " ,hc07=" + CNULL(hc07) + " ,hc12=" + CNULL(hc12) + " where hc01||hc02||hc03||hc04=" + CNULL(ChangeCaseNumL(strlc)) + ""
cnnConnection.Execute strSQL


strSQL = "update caseprogress set cp05=" + CNULL(cp05) + ",cp13=" + CNULL(cp13) + ",cp53=" + CNULL(cp53) + _
           ",cp54=" + CNULL(cp54) + " where cp09=" + CNULL(cp09) + ""
cnnConnection.Execute strSQL


Update003 = True
cnnConnection.CommitTrans

Exit Function

err:
Update003 = False
 cnnConnection.RollbackTrans
 ErrorLog
End Function
Public Function GetCountCase(ByRef strlc As String, ByRef strproperty As String) As Boolean
Dim rs As New ADODB.Recordset
On Error GoTo err
strSQL = "select cp09 from caseprogress where cp01|| cp02||cp03||cp04=" + CNULL(strlc) + _
       " and cp10=" + CNULL(strproperty) + " and cp26 is null"

rs.CursorLocation = adUseClient
rs.CursorType = adOpenDynamic
rs.Open strSQL, cnnConnection

If rs.RecordCount > 1 Then GetCountCase = True Else GetCountCase = False
  Exit Function
err:
ErrorLog

End Function
Public Function GetRs004(ByRef i As Integer, ByRef strtxt As String) As ADODB.Recordset

Select Case i
Case 0
   strSQL = "select cp05,cp06,cp07,cp08,cp09,cp10,cp01,cp02,cp03,cp04,cp13,cp14,cp18,cp27,cp29,cp64,cp71," + _
            " cp50,lc05,lc06,lc07,lc09,lc11,lc15,lc16,lc27  from caseprogress,lawcase where cp09 in (" + strtxt + ") and " + _
            " lc01||lc02||lc03||lc04=cp01||cp02||cp03||cp04  and cp27 is null  and cp57 is null"
Case 1
  strSQL = "select cp05,cp06,cp07,cp08,cp09,cp10,cp01,cp02,cp03,cp04,cp13,cp14,cp18,cp27,cp29,cp64,cp71," + _
            " cp50,hc06,hc05 as lc11,hc07 as lc16,hc12 as lc27  from caseprogress,hirecase where cp09 in (" + strtxt + ") and " + _
            " hc01||hc02||hc03||hc04=cp01||cp02||cp03||cp04  and cp27 is null and cp57 is null  "

Case 2
   strSQL = "select CP05, cp09, CP10,cp13,cp14,cp29,cp01,cp02,cp03,cp04,lc05,lc06,lc07,lc11 from lawcase,caseprogress" + _
           " where cp01|| cp02||cp03||cp04='" + strtxt + "' and substr(cp09,1,1)<>'C' and lc01||lc02||lc03||lc04=cp01||cp02||cp03||cp04 and cp27 is null and cp57 is null"
 Case 3
  strSQL = "select CP05, cp09 ,CP10,cp13,cp14,cp29,cp01,cp02,cp03,cp04,hc06,hc05 from hirecase,caseprogress" + _
           " where cp01|| cp02||cp03||cp04='" + strtxt + "' and substr(cp09,1,1)<>'C' and hc01||hc02||hc03||hc04=cp01||cp02||cp03||cp04 and cp27 is null and cp57 is null"
 
End Select

Dim objPublicData As New clsPublicData
Set GetRs004 = objPublicData.ReadRstD(strSQL)
Set objPublicData = Nothing
Exit Function

err:
ErrorLog

End Function
Public Function GetCaseFee(ByRef strKind As String, ByRef strRelNation As String, ByRef strproperty As String, ByRef dblDay As Double) As Boolean
Dim rs As New ADODB.Recordset

strSQL = "select cf04 from casefee where cf01='" + strKind + "' and cf02='" + strRelNation + "' and cf03='" + strproperty + "'"
rs.CursorLocation = adUseClient
rs.CursorType = adOpenDynamic
rs.Open strSQL, cnnConnection
If rs.RecordCount > 0 Then
   dblDay = rs.Fields(0)
   GetCaseFee = True
End If
Exit Function
err:
GetCaseFee = False
ErrorLog
End Function
Public Function GetNextProgress(ByRef cp09 As String, ByRef cp As String, ByRef strproperty As String, ByRef strPropertymemo As String) As Boolean
Dim rs As New ADODB.Recordset
On Error GoTo err
strSQL = "select np07,np15 from nextprogress where np01='" + cp09 + "' and np02||np03||np04||np05='" + cp + "'"
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.Open strSQL, cnnConnection
If rs.RecordCount > 0 Then
   strproperty = rs.Fields(0)
   strPropertymemo = IIf(IsNull(rs.Fields(1)), "", rs.Fields(1))
   GetNextProgress = True
End If
Exit Function
err:
GetNextProgress = False
ErrorLog
End Function
Public Function GetData005(ByRef strtxt As String, ByRef cp01 As String, ByRef cp02 As String, _
       ByRef cp03 As String, ByRef cp04 As String, ByRef cp05 As String, ByRef cp06 As String, _
       ByRef cp07 As String, ByRef cp08 As String, ByRef cp09 As String, ByRef cp10 As String, _
       ByRef cp13 As String, ByRef cp14 As String, ByRef cp18 As String, ByRef cp27 As String, _
       ByRef cp41 As String, ByRef lc05 As String, ByRef lc09 As String, ByRef lc13 As String, _
       ByRef lc15 As String, ByRef lc27 As String) As Boolean
Dim rs As New ADODB.Recordset

On Error GoTo err
strSQL = "select cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10," + _
       "cp13,cp14,cp18,cp27,cp41,lc05,lc09,lc13,lc15,lc27 from lawcase,caseprogress" + _
       " where lc01||lc02||lc03||lc04=cp01||cp02||cp03||cp04 and cp09='" + strtxt + "'"


rs.CursorLocation = adUseServer
rs.CursorType = adOpenStatic
rs.Open strSQL, cnnConnection
If Not (rs.BOF And rs.EOF) Then
    lc05 = IIf(IsNull(rs.Fields!lc05), "", rs.Fields!lc05)
    lc09 = IIf(IsNull(rs.Fields!lc09), "", rs.Fields!lc09)
    lc13 = IIf(IsNull(rs.Fields!lc13), "", rs.Fields!lc13)
    lc15 = IIf(IsNull(rs.Fields!lc15), "", rs.Fields!lc15)
    lc27 = IIf(IsNull(rs.Fields!lc27), "", rs.Fields!lc27)
    cp01 = IIf(IsNull(rs.Fields!cp01), "", rs.Fields!cp01)
    cp02 = IIf(IsNull(rs.Fields!cp02), "", rs.Fields!cp02)
    cp03 = IIf(IsNull(rs.Fields!cp03), "", rs.Fields!cp03)
    cp04 = IIf(IsNull(rs.Fields!cp04), "", rs.Fields!cp04)
    cp05 = IIf(IsNull(rs.Fields!cp05), "", rs.Fields!cp05)
    cp06 = IIf(IsNull(rs.Fields!cp06), "", rs.Fields!cp06)
    cp07 = IIf(IsNull(rs.Fields!cp07), "", rs.Fields!cp07)
    cp08 = IIf(IsNull(rs.Fields!cp08), "", rs.Fields!cp08)
    cp09 = IIf(IsNull(rs.Fields!cp09), "", rs.Fields!cp09)
    cp10 = IIf(IsNull(rs.Fields!cp10), "", rs.Fields!cp10)
    cp13 = IIf(IsNull(rs.Fields!cp13), "", rs.Fields!cp13)
    cp14 = IIf(IsNull(rs.Fields!cp14), "", rs.Fields!cp14)
    cp18 = IIf(IsNull(rs.Fields!cp18), "", rs.Fields!cp18)
    cp27 = IIf(IsNull(rs.Fields!cp27), "", rs.Fields!cp27)
    cp41 = IIf(IsNull(rs.Fields!cp41), "", rs.Fields!cp41)
    GetData005 = True
Else
    GetData005 = False
End If
Exit Function
err:
ErrorLog

End Function
Public Function Update006(ByRef cp27 As String, ByRef cp71 As String, ByRef np07 As String, ByRef cp06 As String, _
            ByRef cp07 As String, ByRef np15 As String, ByRef lc11 As String, ByRef lc16 As String, ByRef cp50 As String, _
            ByRef cp64 As String, ByRef lc27 As String, ByRef cp09 As String, ByRef lc01 As String, ByRef lc02 As String, _
            ByRef lc03 As String, ByRef lc04 As String, ByRef cp13 As String, ByRef cp14 As String, ByRef cp10 As String, _
            ByRef cp05 As String, ByRef TDate As String, ByRef i As Integer, ByRef Dname As String, ByRef CDT As String, ByRef CTM As String) As Boolean
On Error GoTo err
Dim strlc As String, j As Integer, strcp50 As Variant, strNewNum As String, strPaperNum As String
Dim objPublicData As New clsPublicData

strlc = lc01 + lc02 + lc03 + lc04
cnnConnection.BeginTrans
    strSQL = "update caseprogress set cp27=" + CNULL(cp27) + ",cp71=" + CNULL(cp71) + ",cp06=" + CNULL(cp06) + _
                ",cp07=" + CNULL(cp07) + ",cp64=" + CNULL(cp64) + " where cp09='" + cp09 + "'"
                Debug.Print strSQL
    cnnConnection.Execute strSQL
  If lc01 <> "LA" Then
    strSQL = "update lawcase set lc11=" + CNULL(ChangeCustomerL(lc11)) + ",lc16=" + CNULL(lc16) + ",lc27=" + CNULL(lc27) + _
                " where lc01||lc02||lc03||lc04=" + CNULL(strlc) + ""
  Else
    strSQL = "update hirecase set hc05=" + CNULL(ChangeCustomerL(lc11)) + ",hc07=" + CNULL(lc16) + ",hc12=" + CNULL(lc27) + _
                " where hc01||hc02||hc03||hc04=" + CNULL(strlc) + ""
  End If
    Debug.Print strSQL
                
    cnnConnection.Execute strSQL
    If i = 1 Then
        strSQL = "insert into nextprogress(np01,np02,np03,np04,np05, np07,np08,np09,np10,np16,np17,np18) values(" + CNULL(cp09) + _
                    "," + CNULL(lc01) + "," + CNULL(lc02) + "," + CNULL(lc03) + "," + CNULL(lc04) + "," + CNULL(np07) + _
                    "," + CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(cp13) + "," + CNULL(Dname) + "," + CNULL(CDT) + "," + CNULL(CTM) + ") "
                    Debug.Print strSQL
        cnnConnection.Execute strSQL
    End If
    
    If cp50 <> "" Or Not IsNull(cp50) Then

       strcp50 = Split(cp50, ",")
       For j = 0 To UBound(strcp50) - 1
         If objPublicData.GetAutoNumber("B", strNewNum, 1, 1) Then
             strPaperNum = "B" + GetTaiwanThisYear() + strNewNum
         End If
        
       strSQL = "insert into caseprogress(cp09,cp01,cp02,cp03,cp04,cp50,cp10,cp13,cp14,cp05,cp43) values('" + strPaperNum + _
         "','" + lc01 + "','" + lc02 + "','" + lc03 + "','" + lc04 + "','" + strcp50(j) + "','" + cp10 + "','" + cp13 + "','" + cp14 + "','" + TDate + "','" + cp09 + "')"
            Debug.Print strSQL
       cnnConnection.Execute strSQL
         
      Next
   End If
    
    cnnConnection.CommitTrans
    Update006 = True
    Exit Function
err:
Update006 = False
cnnConnection.RollbackTrans
ErrorLog
End Function
Public Function GetData007(ByRef strcp09 As String, ByRef strName As Variant) As Boolean
Dim rs As New ADODB.Recordset, strTemp() As String, i As Integer
On Error GoTo err
strSQL = "select cp50 from caseprogress where cp43='" + strcp09 + "'"
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.Open strSQL, cnnConnection

If Not (rs.EOF And rs.BOF) Then
    rs.MoveFirst
    Do
    If Not IsNull(rs.Fields(0)) Then
          ReDim Preserve strTemp(i)
          strTemp(i) = rs.Fields(0)
       i = i + 1
    End If
   rs.MoveNext
   Loop Until rs.EOF
   strName = strTemp
    GetData007 = True
    Set rs = Nothing
    
End If
Exit Function
err:
GetData007 = False
ErrorLog
   
End Function
Public Function GetRS008(ByRef i As Integer, ByRef strtxt As String) As ADODB.Recordset
 
 On Error GoTo err
 Select Case i
 Case 0
 strSQL = "select ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日," + _
     " cp09 AS 收文號,decode(cp01 || CP10, cpm01 || CPM02, CPM03, CPM04) AS 案件性質,decode(CP13," + _
     " S1.ST01,S1.ST02) AS 業務員, decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(CP29,S3.ST01,S3.ST02) AS 法務人員," + _
     " decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)) as 發文日," + _
     " decode(cp71,or01,or02) as 法院,lc05,lc06,lc07,lc11  from caseprogress, lawcase," + _
     " STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,organization where cp01||cp02||cp03||cp04='" + strtxt + "' and " + _
     " lc01||lc02||lc03||lc04=cp01||cp02||cp03||cp04 AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND " + _
     " cp01||cp10=cpm01||cpm02 and cp71=or01(+) and  substr(cp09,1,1)<>'C' and cp27 is not null"
 Case 1
 strSQL = "select ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日," + _
     " cp09 AS 收文號,decode(cp01 || CP10, cpm01 || CPM02, CPM03, CPM04) AS 案件性質,decode(CP13," + _
     " S1.ST01,S1.ST02) AS 業務員, decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(CP29,S3.ST01,S3.ST02) AS 法務人員," + _
     " decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)) as 發文日," + _
     " decode(cp71,or01,or02) as 法院,hc05,hc06 from caseprogress, hirecase,STAFF S1,STAFF S2,STAFF S3, " + _
     " CASEPROPERTYMAP,organization where cp01||cp02||cp03||cp04='" + strtxt + "' and hc01||hc02||hc03||hc04=cp01||cp02||cp03||cp04 " + _
     " AND cp13 = s1.st01(+) AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND cp01||cp10=cpm01||cpm02 and cp71=or01(+) " + _
     " and substr(cp09,1,1)<>'C' and cp27 is not null order by cp05"
     
End Select
Dim objPublicData As New clsPublicData
Set GetRS008 = objPublicData.ReadRstD(strSQL)
Set objPublicData = Nothing
Exit Function

err:
ErrorLog

End Function
Public Function GetRs009(ByRef i As Integer, ByRef strtxt As String) As ADODB.Recordset

On Error GoTo err
Select Case i
Case 1
strSQL = "select cp09,cp10,cp46,cp25,cp06,cp07,cp71,cp30,cp08,cp13," + _
          " cp14,cp48,cp29,cp64 ,lc15,lc27,lc08 from caseprogress,lawcase where cp09='" + strtxt + _
          "' and cp01||cp02||cp03||cp04=lc01||lc02||lc03||lc04"
Case 2
strSQL = "select cp09,cp10,cp46,cp25,cp06,cp07,cp71,cp30,cp08,cp13," + _
          " cp14,cp48,cp29,cp64 ,hc12,hc09 from caseprogress,hirecase where cp09='" + strtxt + _
          "' and cp01||cp02||cp03||cp04=hc01||hc02||hc03||hc04"
  
 End Select
Dim objPublicData As New clsPublicData
Set GetRs009 = objPublicData.ReadRst(strSQL)
Set objPublicData = Nothing
Exit Function

err:
ErrorLog

End Function
Public Function GetRs015(ByRef i As Integer, ByRef strtxt As String) As ADODB.Recordset

On Error GoTo err
Select Case i
Case 1
strSQL = "select cp09,cp10,cp46,cp25,cp06,cp07,cp71,cp30,cp08,cp13," + _
          " cp14,cp48,cp29,cp64 ,lc15,lc27,lc08 from caseprogress,lawcase where " + _
        " cp09 in (" + strtxt + ") and cp01||cp02||cp03||cp04=lc01||lc02||lc03||lc04"
Case 2
strSQL = "select cp09,cp10,cp46,cp25,cp06,cp07,cp71,cp30,cp08,cp13," + _
          " cp14,cp48,cp29,cp64 ,hc12,hc09 from caseprogress,hirecase where " + _
          " cp09 in (" + strtxt + ") and cp01||cp02||cp03||cp04=hc01||hc02||hc03||hc04"
  
 End Select
Dim objPublicData As New clsPublicData
Set GetRs015 = objPublicData.ReadRst(strSQL)
Set objPublicData = Nothing
Exit Function

err:
ErrorLog

End Function

Public Function ChkMRec(ByRef mr02 As String, ByRef strlc As String, ByRef mr16 As String, ByRef mr17 As String) As Boolean
Dim rs As New ADODB.Recordset
On Error GoTo err
strSQL = " select mr16,mr17 from mailrec where mr02=" + CNULL(mr02) + " and mr12||mr13||mr14||mr15=" + CNULL(strlc) + ""
rs.CursorLocation = adUseClient
rs.CursorType = adOpenStatic
rs.Open strSQL, cnnConnection
If rs.RecordCount > 0 Then
    mr16 = IIf(IsNull(rs.Fields!mr16), "", rs.Fields!mr16)
    mr17 = IIf(IsNull(rs.Fields!mr17), "", rs.Fields!mr17)
    ChkMRec = True
Else
  ChkMRec = False
End If

Exit Function
err:
ErrorLog
End Function
Public Function Update009(ByRef Ccp09 As String, ByRef cp01 As String, ByRef cp02 As String, ByRef cp03 As String, ByRef cp04 As String, ByRef cp05 As String, _
       ByRef cp12 As String, ByRef cp13 As String, ByRef cp43 As String, ByRef cp25 As String, ByRef cp10 As String, ByRef np07 As String, ByRef cp06 As String _
       , ByRef cp07 As String, ByRef cp71 As String, ByRef cp30 As String, ByRef cp35 As String, ByRef cp08 As String, ByRef cp14 As String _
       , ByRef cp48 As String, ByRef cp29 As String, ByRef cp64 As String, ByRef mr09 As String) As Boolean

Dim strTemp As String
On Error GoTo err
cnnConnection.BeginTrans
    strSQL = "insert into caseprogress(cp09,cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp12,cp13,cp32,cp43,cp26," + _
            " cp10,cp08,cp35) values(" + CNULL(Ccp09) + "," + CNULL(cp01) + "," + CNULL(cp02) + "," + CNULL(cp03) + _
            "," + CNULL(cp04) + "," + CNULL(cp05) + "," + CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(cp12) + _
            "," + CNULL(cp13) + ",'N'," + CNULL(cp43) + ",'N'," + CNULL(cp10) + "," + CNULL(cp08) + "," + CNULL(cp35) + ")"
    cnnConnection.Execute strSQL
    
    If np07 <> "" Or Not IsNull(np07) Then
    strSQL = "insert into nextprogress(np01,np02,np03,np04,np05,np07,np10,np13,np15) values(" + CNULL(Ccp09) + "," + CNULL(cp01) + "," + CNULL(cp02) + _
         "," + CNULL(cp03) + "," + CNULL(cp04) + "," + CNULL(np07) + "," + CNULL(cp13) + "," + CNULL(cp08) + "," + CNULL(cp64) + ")"
    cnnConnection.Execute strSQL
    End If
    If cp25 <> "" Or IsNull(cp25) Then
      strSQL = "update caseprogress set cp25=" + CNULL(cp25) + " where cp09='" + cp43 + "' "
      cnnConnection.Execute strSQL
    End If
Update009 = True
cnnConnection.CommitTrans

Exit Function
err:
    Update009 = False
    cnnConnection.RollbackTrans
    ErrorLog
End Function
Public Function GetRS010(ByRef i As Integer, ByRef strtxt As String) As ADODB.Recordset
 
 On Error GoTo err
 Select Case i
 Case 0
 strSQL = "select ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日," + _
     " cp09 AS 收文號,decode(cp01 || CP10, cpm01 || CPM02, CPM03, CPM04) AS 案件性質,decode(CP13," + _
     " S1.ST01,S1.ST02) AS 業務員, decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(CP29,S3.ST01,S3.ST02) AS 法務人員," + _
     " decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)) as 發文日," + _
     " decode(cp71,or01,or02) as 機關代號,cp50 as 收件人,lc05,lc06,lc07,lc11  from caseprogress, lawcase," + _
     " STAFF S1,STAFF S2,STAFF S3, CASEPROPERTYMAP,organization where cp01||cp02||cp03||cp04='" + strtxt + "' and" + _
     " lc01||lc02||lc03||lc04=cp01||cp02||cp03||cp04 AND cp13 = s1.st01(+)  AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND " + _
     " cp01||cp10=cpm01||cpm02 and cp71=or01(+) and substr(cp09,1,1)<>'C' and cp27 is not null"
 Case 1
 strSQL = "select ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日," + _
     " cp09 AS 收文號,decode(cp01 || CP10, cpm01 || CPM02, CPM03, CPM04) AS 案件性質,decode(CP13," + _
     " S1.ST01,S1.ST02) AS 業務員, decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(CP29,S3.ST01,S3.ST02) AS 法務人員," + _
     " decode(cp27,null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)) as 發文日," + _
     " decode(cp71,or01,or02) as 機關代號,cp50 as 收件人,hc05,hc06 from caseprogress, hirecase,STAFF S1," + _
     " STAFF S2,STAFF S3, CASEPROPERTYMAP,organization  where cp01||cp02||cp03||cp04='" + strtxt + "' and" + _
     " hc01||hc02||hc03||hc04=cp01||cp02||cp03||cp04 AND cp13 = s1.st01(+) and substr(cp09,1,1)<>'C' " + _
     " AND cp14 = s2.st01(+) and cp29 = s3.st01(+) AND cp01||cp10=cpm01||cpm02 and cp71=or01(+) and cp27 is not null"
End Select
Dim objPublicData As New clsPublicData
Set GetRS010 = objPublicData.ReadRstD(strSQL)
Set objPublicData = Nothing

Exit Function
err:
ErrorLog
End Function
Public Function GetRs011(ByRef strtxt As String) As ADODB.Recordset
 
 On Error GoTo err
 strSQL = "select cp09 ,cp10 ,cp27,cp46, cp50  from caseprogress where cp09='" + strtxt + "'"

Dim objPublicData As New clsPublicData
Set GetRs011 = objPublicData.ReadRst(strSQL)
Set objPublicData = Nothing
Exit Function
err:
ErrorLog
End Function
Public Function Update011(ByRef cp09 As String, ByRef cp46 As String) As Boolean
On Error GoTo err
 
 strSQL = "update caseprogress set cp46=" + cp46 + " where cp09='" + cp09 + "'"
 cnnConnection.Execute strSQL
 Update011 = True
 Exit Function
err:
Update011 = False
ErrorLog
End Function
Public Function GetRS013(ByRef cp09 As String) As ADODB.Recordset

On Error GoTo err
  strSQL = "select cdp05,cp35,cp30,cdp06,cdp03,cdp04,cdp02,cdp07,cdp08,cdp09,cp09," + _
       " cp46 ,cp49,cp12,cp13,cp64 from caseprogress,courtyardperiod where cp43 " + _
       " in (" + cp09 + ")  and cp09=cdp01  and  cdp03=cp06 and cp71=cdp05 order by cdp03 desc"
  
Dim objPublicData As New clsPublicData
Set GetRS013 = objPublicData.ReadRst(strSQL)
Set objPublicData = Nothing
Exit Function
err:
ErrorLog
End Function
Public Function Update013(ByRef cp09 As String, ByRef cp01 As String, ByRef cp02 As String, ByRef cp03 As String, ByRef cp04 As String, _
        ByRef cp12 As String, ByRef cp13 As String, ByRef cp43 As String, ByRef cp05 As String, ByRef cp71 As String, ByRef cp35 As String, ByRef cp30 As String, _
        ByRef cdp06 As String, ByRef cdp03 As String, ByRef cdp04 As String, ByRef cdp02 As String, ByRef cdp08 As String, ByRef cdp09 As String, _
        ByRef cp64 As String, ByRef cid As String, ByRef Cd As String, ByRef Ctime As String) As Boolean
Dim cdp05 As String
On Error GoTo err
cnnConnection.BeginTrans
    strSQL = "insert into caseprogress(cp09,cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp48,cp12,cp13,cp32,cp43,cp26," + _
            " cp10,cp71,cp35) values(" + CNULL(cp09) + "," + CNULL(cp01) + "," + CNULL(cp02) + "," + CNULL(cp03) + "," + CNULL(cp04) + _
            "," + CNULL(cp05) + "," + CNULL(cdp03) + "," + CNULL(cdp03) + "," + CNULL(cdp03) + "," + CNULL(cp12) + "," + CNULL(cp13) + _
            ",'N'," + CNULL(cp43) + ",'N'," + CNULL(通知開庭) + "," + CNULL(cp71) + "," + CNULL(cp35) + ")"
cnnConnection.Execute strSQL
    
    strSQL = "insert into courtyardperiod(cdp01,cdp02,cdp03,cdp04,cdp05,cdp06,cdp07,cdp08,cdp09,cdp10,cdp11,cdp12) values " + _
            " (" + CNULL(cp09) + "," + CNULL(cdp02) + "," + CNULL(cdp03) + "," + CNULL(cdp04) + "," + CNULL(cp71) + _
            "," + CNULL(cdp06) + "," + CNULL(cp64) + "," + CNULL(cdp08) + "," + CNULL(cdp09) + "," + CNULL(cid) + "," + CNULL(Cd) + "," + CNULL(Ctime) + ")"
cnnConnection.Execute strSQL
cnnConnection.CommitTrans
Update013 = True
Exit Function
err:
cnnConnection.RollbackTrans
ErrorLog
End Function
Public Function GetData015(ByRef strtxt As String, ByRef lc01 As String, ByRef lc02 As String, _
    ByRef lc03 As String, ByRef lc04 As String, ByRef cdp06 As String, ByRef cdp07 As String, _
    ByRef cdp08 As String, ByRef cdp09 As String, ByRef cdp10 As String, ByRef lc09 As String, ByRef lc05 As String, _
    ByRef lc15 As String, ByRef mr02 As String, ByRef mr07 As String, ByRef mr08 As String) As Boolean
 Dim rs As New ADODB.Recordset
 
 strSQL = "select lc01,lc02,lc03,lc04,cdp06,cdp07,cdp08,cdp09,cdp10,lc09,lc05,mr02,mr07,mr08 from " + _
          "mailrec,courtyardperiod,lawcase where mr01='" + strtxt + "' and mr09||mr10||mr11||mr12=lc01||lc02||lc03||lc04" + _
          "mr09||mr10||mr11||mr12=cdp01||cdp02||cdp03||cdp04"
          
rs.CursorLocation = adUseServer
rs.CursorType = adOpenStatic
rs.Open strSQL, cnnConnection
If Not (rs.BOF And rs.EOF) Then
   lc01 = IIf(IsNull(rs.Fields!lc01), "", rs.Fields!lc01)
   lc02 = IIf(IsNull(rs.Fields!lc02), "", rs.Fields!lc02)
   lc03 = IIf(IsNull(rs.Fields!lc03), "", rs.Fields!lc03)
   lc04 = IIf(IsNull(rs.Fields!lc04), "", rs.Fields!lc04)
   lc05 = IIf(IsNull(rs.Fields!lc05), "", rs.Fields!lc05)
   lc09 = IIf(IsNull(rs.Fields!lc09), "", rs.Fields!lc09)
   mr02 = IIf(IsNull(rs.Fields!mr02), "", rs.Fields!mr02)
   mr07 = IIf(IsNull(rs.Fields!mr07), "", rs.Fields!mr07)
   mr08 = IIf(IsNull(rs.Fields!mr08), "", rs.Fields!mr08)
   cdp06 = IIf(IsNull(rs.Fields!cdp06), "", rs.Fields!cdp06)
   cdp07 = IIf(IsNull(rs.Fields!cdp07), "", rs.Fields!cdp07)
   cdp08 = IIf(IsNull(rs.Fields!cdp08), "", rs.Fields!cdp08)
   cdp09 = IIf(IsNull(rs.Fields!cdp09), "", rs.Fields!cdp09)
   cdp10 = IIf(IsNull(rs.Fields!cdp10), "", rs.Fields!cdp10)
   GetData015 = True
Else
   MsgBox "無此資料"
   GetData015 = False
End If
End Function
Public Function GetRS014(ByRef i As Integer, ByRef strtxt As String) As ADODB.Recordset
 
 On Error GoTo err
 Select Case i
 Case 0
 strSQL = "select ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日," + _
     " cp09 AS 收文號,decode(cp01 || CP10, cpm01 || CPM02, CPM03, CPM04) AS 案件性質,decode(CP13," + _
     " S1.ST01,S1.ST02) AS 業務員, decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(cp27," + _
     " null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)) as 發文日," + _
     " decode(cp71,or01,or02) as 機關名稱,decode(cp24,1,'淮/勝',2,'駁/敗') as 結果," + _
     " decode(cp44,fa01||fa02,nvl(fa04,nvl(fa05,fa06))) as 代理人,lc05,lc06,lc07,lc11  from caseprogress, lawcase," + _
     " STAFF S1,STAFF S2, fagent,CASEPROPERTYMAP,organization where cp01||cp02||cp03||cp04='" + strtxt + "' and " + _
     " lc01||lc02||lc03||lc04=cp01||cp02||cp03||cp04 AND cp13 = s1.st01(+) AND cp14 = s2.st01(+)  AND " + _
     " cp01||cp10=cpm01||cpm02 and cp71=or01(+) and  substr(cp44,1,8)=fa01(+) " + _
     " and substr(cp44,9,9)=fa02(+) and substr(cp09,1,1)<>'C' and cp27 is not null"
 Case 1
 strSQL = "select ' ' v,SUBSTR(CP05, 1, 4)- 1911 || '/' || SUBSTR(CP05, 5, 2)|| '/' || SUBSTR(CP05, 7, 2) AS 收文日," + _
     " cp09 AS 收文號,decode(cp01 || CP10, cpm01 || CPM02, CPM03, CPM04) AS 案件性質,decode(CP13," + _
     " S1.ST01,S1.ST02) AS 業務員, decode(CP14,S2.ST01,S2.ST02) AS 承辦人,decode(cp27," + _
     " null,'',SUBSTR(CP27, 1, 4)- 1911 || '/' || SUBSTR(CP27, 5, 2)|| '/' || SUBSTR(CP27, 7, 2)) as 發文日," + _
     " decode(cp71,or01,or02) as 機關名稱,decode(cp24,1,'淮/勝',2,'駁/敗') as 結果," + _
     " decode(cp44,fa01||fa02,nvl(fa04,nvl(fa05,fa06))) as 代理人,hc05,hc06 from caseprogress, hirecase,STAFF S1,STAFF S2 ," + _
     " CASEPROPERTYMAP,fagent,organization where cp01||cp02||cp03||cp04='" + strtxt + "' and hc01||hc02||hc03||hc04=cp01||cp02||cp03||cp04 " + _
     " AND cp13 = s1.st01(+) AND cp14 = s2.st01(+)  AND cp01||cp10=cpm01||cpm02 and cp71=or01(+) " + _
     " and substr(cp44,1,8)=fa01(+) and substr(cp44,9,9)=fa02(+) and substr(cp09,1,1)<>'C' and cp27 is not null"
     
End Select
Dim objPublicData As New clsPublicData
Set GetRS014 = objPublicData.ReadRstD(strSQL)
Set objPublicData = Nothing
Exit Function

err:
ErrorLog
End Function
Public Function Update015(ByRef Ccp09 As String, ByRef cp01 As String, ByRef cp02 As String, ByRef cp03 As String, ByRef cp04 As String, ByRef cp05 As String, _
       ByRef cp12 As String, ByRef cp13 As String, ByRef cp43 As String, ByRef cp10 As String, ByRef np07 As String, ByRef cp06 As String _
       , ByRef cp07 As String, ByRef cp71 As String, ByRef cp08 As String, ByRef cp14 As String _
       , ByRef cp48 As String, ByRef cp29 As String, ByRef lc08 As String, ByRef cp64 As String, ByRef cid As String, ByRef dt As String, ByRef tm As String) As Boolean
Dim cp35 As String
Dim strTemp As String
On Error GoTo err
cnnConnection.BeginTrans
    strSQL = "insert into caseprogress(cp09,cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp12,cp13,cp20,cp32,cp43,cp26," + _
            " cp10,cp08,cp65,cp66,cp67) values(" + CNULL(Ccp09) + "," + CNULL(cp01) + "," + CNULL(cp02) + "," + CNULL(cp03) + _
            "," + CNULL(cp04) + "," + CNULL(cp05) + "," + CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(cp12) + _
            "," + CNULL(cp13) + ",'N','N'," + CNULL(cp43) + ",'N'," + CNULL(cp10) + "," + CNULL(cp08) + "," + CNULL(cid) + "," + CNULL(dt) + "," + CNULL(tm) + ")"
         cnnConnection.Execute strSQL
    
    strSQL = "insert into nextprogress(np01,np02,np03,np04,np05,np07,np08,np09,np10,np13,np15,np16,np17,np18) values(" + CNULL(Ccp09) + "," + CNULL(cp01) + "," + CNULL(cp02) + _
         "," + CNULL(cp03) + "," + CNULL(cp04) + "," + CNULL(np07) + "," + CNULL(cp06) + "," + CNULL(cp07) + "," + CNULL(cp13) + "," + CNULL(cp08) + "," + CNULL(cp64) + "," + CNULL(cid) + "," + CNULL(dt) + "," + CNULL(tm) + ")"
        cnnConnection.Execute strSQL
    
    strSQL = "update lawcase set lc08=" + CNULL(lc08) + " where lc01=" + CNULL(cp01) + " and lc02=" + CNULL(cp02) + " and " + _
                " lc03=" + CNULL(cp03) + " and lc04=" + CNULL(cp04) + ""
    
    Debug.Print strSQL
    cnnConnection.Execute strSQL
Update015 = True
cnnConnection.CommitTrans

Exit Function
err:
    Update015 = False
    cnnConnection.RollbackTrans
    ErrorLog
End Function

