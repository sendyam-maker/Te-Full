VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cls08"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim strSQL As String
Public Function GetRs201(ByRef i As Integer, ByRef strMerger As String) As ADODB.Recordset
Dim strSQL1 As String, strSQL2 As String
On Error GoTo err
Select Case i
Case 1
strSQL = "select  cp01||'-'||cp02||decode(cp03,'0','','-'||cp03)||decode(cp04,'00','','-'||cp04)  本所案號,nvl(lc05,nvl(lc06,lc07)) 案件名稱,decode(cp46,null,'',substr(cp46,1,4)-1911" + _
       " ||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2)) 收受日,decode(cdp02,st01,st02) 開庭人員,decode(cdp05,or01,or02) " + _
       " 法院,cp35 法院案號,substr(cdp03,1,4)-1911||'/'||substr(cdp03,5,2)||'/'||substr(cdp03,7,2) 開庭日期," + _
       " cdp04 開庭時間,decode(cdp06,'1','偵查','2','審查','3','辯論') 開庭性質, cdp08 法官,cdp09 檢察官,cdp07 備註  " + _
       " from lawcase,caseprogress,courtyardperiod,staff,organization  where cp01||cp02||cp03||cp04=lc01||lc02||lc03||lc04" + _
       " and  cdp02=st01(+) and cdp05=or01(+) and  cp09=cdp01  and  " + strMerger + "  "
Case 2
strSQL = "select  cp01||'-'||cp02||decode(cp03,'0','','-'||cp03)||decode(cp04,'00','','-'||cp04)  本所案號,hc06 案件名稱,decode(cp46,null,'',substr(cp46,1,4)-1911" + _
       " ||'/'||substr(cp46,5,2)||'/'||substr(cp46,7,2)) 收受日,decode(cdp02,st01,st02) 開庭人員,decode(cdp05,or01,or02) " + _
       " 法院,cp35 法院案號,substr(cdp03,1,4)-1911||'/'||substr(cdp03,5,2)||'/'||substr(cdp03,7,2) 開庭日期," + _
       " cdp04 開庭時間,decode(cdp06,'1','偵查','2','審查','3','辯論') 開庭性質, cdp08 法官,cdp09 檢察官,cdp07 備註  " + _
       " from hirecase,caseprogress,courtyardperiod,staff,organization  where cp01||cp02||cp03||cp04=hc01||hc02||hc03||hc04" + _
       " and  cdp02=st01(+) and cdp05=or01(+) and  cp09=cdp01  and  " + strMerger + "  "
End Select
Dim objPublicData As New clsPublicData
Set GetRs201 = objPublicData.ReadRstD(strSQL)
Set objPublicData = Nothing
Exit Function
err:
ErrorLog
End Function
Public Function GetRs203(ByRef strMerger1 As String, ByRef strMerger2 As String) As ADODB.Recordset
Dim strSQL1 As String, strSQL2 As String
On Error GoTo err

strSQL1 = "select  cp01||'-'||cp02||decode(cp03,'0','','-'||cp03)||decode(cp04,'00','','-'||cp04)  本所案號,substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 收文日," + _
       " decode(lc11,cu01||cu02,nvl(cu04,nvl(cu05,cu06))) 當事人,decode(cp10,cpm02,cpm03,cpm04) 案件性質,decode(cp14,st01,st02,-1) 承辦人, decode(cp29,st01,st02) 法務人員," + _
       " decode(cp27,null,'',substr(cp27,1,4)-1911||'/'||substr(cp27,5,2)||'/'||substr(cp27,7,2)) 發文日  from lawcase,caseprogress,staff,casepropertymap,customer  where cp01||cp02||cp03||cp04=lc01||lc02||lc03||lc04" + _
       " and cpm01=cp01 and cp14=st01 and substr(lc11,1,8)=cu01(+) and substr(lc11,9,9)=cu02(+)   and cp10=cpm02  and " + strMerger1 + "" + strMerger2 + " "
         
strSQL2 = "select cp01||'-'||cp02||decode(cp03,'0','','-'||cp03)||decode(cp04,'00','','-'||cp04)  本所案號,substr(cp05,1,4)-1911||'/'||substr(cp05,5,2)||'/'||substr(cp05,7,2) 收文日," + _
      " decode(hc05,cu01||cu02,nvl(cu04,nvl(cu05,cu06))) 當事人,decode(cp10,cpm02,cpm03,cpm04) 案件性質,decode(cp14,st01,st02,-1) 承辦人,decode(cp29,st01,st02) 法務人員," + _
      " decode(cp27,null,'',substr(cp27,1,4)-1911||'/'||substr(cp27,5,2)||'/'||substr(cp27,7,2)) 發文日 from hirecase,caseprogress,staff,casepropertymap,customer" + _
      " where cp01||cp02||cp03||cp04=hc01||hc02||hc03||hc04 and cpm01=cp01 and cp14=st01 and substr(hc05,1,8)=cu01(+) and substr(hc05,9,9)=cu02(+)   and cp10=cpm02 and " + strMerger1 + "" + strMerger2 + " "
      
strSQL = strSQL1 + " Union " + strSQL2
Dim objPublicData As New clsPublicData
Set GetRs203 = objPublicData.ReadRstD(strSQL)
Set objPublicData = Nothing
Exit Function

err:
ErrorLog
End Function
Public Function GetRS202(ByRef i As Integer, ByRef txt1 As String, ByRef txt2 As String) As ADODB.Recordset
  On Error GoTo err
  Select Case i
   Case 0
    
    strSQL = "select mr09 mr10,mr11,mr12, lc05,cdp06,mr07,mr08,cdp07,cdp08," + _
         " cdp10 from courtyardperiod,mailrec,lawcase where mr09||mr10||mr11||mr12='" + txt1 + _
         "' and mr09||mr10||mr11||mr12=lc01||lc02||lc03||lc04 "
   Case 1
    strSQL = "select mr09 mr10,mr11,mr12, lc05,cdp06,mr07,mr08,cdp07,cdp08," + _
         " cdp10 from courtyardperiod,mailrec,lawcase where mr08='" + txt1 + _
         "' and mr09||mr10||mr11||mr12=lc01||lc02||lc03||lc04 "
   
   Case 2
    strSQL = "select mr09 mr10,mr11,mr12, lc05,cdp06,mr07,mr08,cdp07,cdp08," + _
         " cdp10 from courtyardperiod,mailrec,lawcase where cdp06='" + txt1 + _
         "' and mr09||mr10||mr11||mr12=lc01||lc02||lc03||lc04 "
 
   Case 3
    strSQL = "select mr09 mr10,mr11,mr12, lc05,cdp06,mr07,mr08,cdp07,cdp08," + _
         " cdp10 from courtyardperiod,mailrec,lawcase where cdp07 between '" + txt1 + _
         "' and '" + txt2 + "' and mr09||mr10||mr11||mr12=lc01||lc02||lc03||lc04 "
    
  End Select
Dim objPublicData As New clsPublicData
Set GetRS202 = objPublicData.ReadRstD(strSQL)
Set objPublicData = Nothing
Exit Function
err:
ErrorLog

End Function
Public Function GetRs204(ByRef law1 As String, ByRef law2 As String, ByRef day1 As String, ByRef day2 As String)
On Error GoTo err
strSQL = "select cp01 cp02,cp03,cp04,cp05,lc16,cp14,cp28,cp13,cp27 from lawcase,caseprogress " + _
         " where cp28 between '" + law1 + "' and '" + law2 + "' and cp05 between '" + day1 + _
         "' and '" + day2 + "' and lc01||lc02||lc03||lc04=cp01||cp02||cp03||cp04"

Dim objPublicData As New clsPublicData
Set GetRs204 = objPublicData.ReadRstD(strSQL)
Set objPublicData = Nothing
Exit Function
err:
ErrorLog

End Function
Public Function getRS206(ByRef lw1 As String, ByRef lw2 As String) As ADODB.Recordset

On Error GoTo err
 strSQL = "select lw01, lw02,lw03 from law " + _
       "  where lw01 Between '" + lw1 + "' and '" + lw2 + "'"

Dim objPublicData As New clsPublicData
Set getRS206 = objPublicData.ReadRstD(strSQL)
Set objPublicData = Nothing
Exit Function

err:
ErrorLog
End Function
Public Function getRS208(ByRef num1 As String, ByRef num2 As String, ByRef name As String, ByRef cty1 As String, ByRef cty2 As String) As ADODB.Recordset

On Error GoTo err
 strSQL = "select ecd01, ecd02,ecd03,ecd04,ecd05 from law " + _
       "  where ecd02 Between '" + num1 + "' and '" + num2 + "' and ecd03='" + name + _
       "' ecd05 between '" + cty1 + "' and '" + cty2 + "' "

Dim objPublicData As New clsPublicData
Set getRS208 = objPublicData.ReadRstD(strSQL)
Set objPublicData = Nothing
Exit Function

err:
ErrorLog
End Function

