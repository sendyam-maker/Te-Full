VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsla001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim strSQL As String
Public Function GetRS309() As ADODB.Recordset
On Error GoTo Err
 strSQL = "select  a0901,nvl(a0902,a0903) from acc090 "

Dim objPublicData As New clsPublicData
Set GetRS309 = objPublicData.ReadRst(strSQL)
Set objPublicData = Nothing
Exit Function

Err:
ErrorLog
End Function
Public Function GetRs310(ByRef strMerge As String) As ADODB.Recordset
On Error GoTo Err

strSQL = "select '' v ,decode(CP13,ST01, ST02)  as 業務員,hc05 客戶編號, " + _
          " decode(hc05, CU01 || CU02,NVL(CU04, NVL(CU05, CU06))) as 客戶名稱," + _
          "DECODE(CP54,NULL,'',SUBSTR(CP54,1,4)-1911||'/'||SUBSTR(CP54,5,2)||'/'||SUBSTR(CP54,7,2)) as 聘任到期日 " + _
          " from caseprogress,hirecase,customer,staff where substr(hc05,1,8)=cu01(+) and substr(hc05,9,1)=cu02(+) " + _
          " and hc01=cp01 and hc02=cp02 and hc03=cp03 and hc04=cp04 and cp13=st01(+) and  " + strMerge + ""

Dim objPublicData As New clsPublicData
Set GetRs310 = objPublicData.ReadRst(strSQL)
Set objPublicData = Nothing
Exit Function

Err:
ErrorLog
End Function
Public Function GetRs502(ByRef i As Integer) As ADODB.Recordset

On Error GoTo Err
If i = 1 Then
strSQL = "select lc11,lc05,lc06,lc07,lc08,lc16,lc09,lc10,lc15,lc17," + _
         " lc22,lc23,lc13,lc14,lc26,lc25,lc24,lc12,lc21,lc18,lc19,lc20," + _
         " lc27,lc34,lc28,lc29,lc30,lc31,lc32,lc33,lc01,lc02,lc03,lc04 from " + _
         " lawcase  where lc01='L' or lc01='FCL' order by lc01||lc02||lc03||lc04"
Else
strSQL = "select lc11,lc05,lc06,lc07,lc08,lc16,lc09,lc10,lc15,lc17,lc22," + _
         " lc23,lc13,lc14,lc26,lc25,lc24,lc12,lc21,lc18,lc19,lc20,lc27,lc34,lc28," + _
         " lc29,lc30,lc31,lc32,lc33,lc01,lc02,lc03,lc04 from lawcase  where lc01='CFL'"
End If
Dim objPublicData As New clsPublicData
Set GetRs502 = objPublicData.ReadRst(strSQL)
Set objPublicData = Nothing
Exit Function

Err:
ErrorLog
End Function
Public Function GetData502(ByRef i As String, ByRef strlc As String, ByRef lc05 As String, ByRef lc06 As String, _
               ByRef lc07 As String, ByRef lc08 As String, ByRef lc09 As String, ByRef lc10 As String, _
               ByRef lc11 As String, ByRef lc12 As String, ByRef lc13 As String, ByRef lc14 As String, _
               ByRef lc15 As String, ByRef lc16 As String, ByRef lc17 As String, ByRef lc18 As String, _
               ByRef lc19 As String, ByRef lc20 As String, ByRef lc21 As String, ByRef lc22 As String, _
               ByRef lc23 As String, ByRef lc24 As String, ByRef lc25 As String, ByRef lc26 As String, _
               ByRef lc27 As String, ByRef lc28 As String, ByRef lc29 As String, ByRef lc30 As String, _
               ByRef lc31 As String, ByRef lc32 As String, ByRef lc33 As String) As Boolean
Dim rst As New ADODB.Recordset

On Error GoTo Err
    strSQL = "select lc01,lc02,lc03,lc04,lc05,lc06,lc07,lc08,lc09,lc10," + _
             " lc11,lc12,lc13,lc14,lc15,lc16,lc17,lc18,lc19,lc20,lc21,lc22,lc23," + _
             " lc24,lc25,lc26,lc27,lc28,lc29,lc30,lc31,lc32,lc33 from lawcase where " + ChgLawcase(strlc)

rst.CursorLocation = adUseClient
rst.CursorType = adOpenStatic
rst.Open strSQL, cnnConnection
If rst.RecordCount > 0 Then
    rst.MoveFirst
    lc05 = IIf(IsNull(rst.Fields!lc05), "", rst.Fields!lc05)
    lc06 = IIf(IsNull(rst.Fields!lc06), "", rst.Fields!lc06)
    lc07 = IIf(IsNull(rst.Fields!lc07), "", rst.Fields!lc07)
    lc08 = IIf(IsNull(rst.Fields!lc08), "", rst.Fields!lc08)
    lc09 = IIf(IsNull(rst.Fields!lc09), "", rst.Fields!lc09)
    lc10 = IIf(IsNull(rst.Fields!lc10), "", rst.Fields!lc10)
    lc11 = IIf(IsNull(rst.Fields!lc11), "", rst.Fields!lc11)
    lc12 = IIf(IsNull(rst.Fields!lc12), "", rst.Fields!lc12)
    lc13 = IIf(IsNull(rst.Fields!lc13), "", rst.Fields!lc13)
    lc14 = IIf(IsNull(rst.Fields!lc14), "", rst.Fields!lc14)
    lc15 = IIf(IsNull(rst.Fields!lc15), "", rst.Fields!lc15)
    lc16 = IIf(IsNull(rst.Fields!lc16), "", rst.Fields!lc16)
    lc17 = IIf(IsNull(rst.Fields!lc17), "", rst.Fields!lc17)
    lc18 = IIf(IsNull(rst.Fields!lc18), "", rst.Fields!lc18)
    lc19 = IIf(IsNull(rst.Fields!lc19), "", rst.Fields!lc19)
    lc20 = IIf(IsNull(rst.Fields!lc20), "", rst.Fields!lc20)
    lc21 = IIf(IsNull(rst.Fields!lc21), "", rst.Fields!lc21)
    lc22 = IIf(IsNull(rst.Fields!lc22), "", rst.Fields!lc22)
    lc23 = IIf(IsNull(rst.Fields!lc23), "", rst.Fields!lc23)
    lc24 = IIf(IsNull(rst.Fields!lc24), "", rst.Fields!lc24)
    lc25 = IIf(IsNull(rst.Fields!lc25), "", rst.Fields!lc25)
    lc26 = IIf(IsNull(rst.Fields!lc26), "", rst.Fields!lc26)
    lc27 = IIf(IsNull(rst.Fields!lc27), "", rst.Fields!lc27)
    lc28 = IIf(IsNull(rst.Fields!lc28), "", rst.Fields!lc28)
    lc29 = IIf(IsNull(rst.Fields!lc29), "", rst.Fields!lc29)
    lc30 = IIf(IsNull(rst.Fields!lc30), "", rst.Fields!lc30)
    lc31 = IIf(IsNull(rst.Fields!lc31), "", rst.Fields!lc31)
    lc32 = IIf(IsNull(rst.Fields!lc32), "", rst.Fields!lc32)
    lc33 = IIf(IsNull(rst.Fields!lc33), "", rst.Fields!lc33)

    GetData502 = True
    
Else
 MsgBox "資料庫無此資料"
 GetData502 = False
 Exit Function
End If
Err:
ErrorLog
End Function
Public Function Update502(ByRef strlc As String, ByRef lc05 As String, ByRef lc06 As String, ByRef lc07 As String, _
       ByRef lc08 As String, ByRef lc09 As String, ByRef lc10 As String, ByRef lc11 As String, _
       ByRef lc12 As String, ByRef lc13 As String, ByRef lc14 As String, ByRef lc15 As String, _
       ByRef lc16 As String, ByRef lc17 As String, ByRef lc18 As String, ByRef lc19 As String, _
       ByRef lc20 As String, ByRef lc21 As String, ByRef lc22 As String, ByRef lc23 As String, _
       ByRef lc24 As String, ByRef lc25 As String, ByRef lc26 As String, ByRef lc27 As String, _
       ByRef lc34 As String, ByRef lc31 As String, ByRef lc32 As String, ByRef lc33 As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans
strSQL = " update lawcase set lc05=" + CNULL(lc05) + ",lc06=" + CNULL(lc06) + ",lc07=" + CNULL(lc07) + _
         ",lc08=" + CNULL(lc08) + ",lc09=" + CNULL(lc09) + ",lc10=" + CNULL(lc10) + ",lc11=" + CNULL(lc11) + _
         ",lc12=" + CNULL(lc12) + ",lc13=" + CNULL(lc13) + ",lc14=" + CNULL(lc14) + ",lc15=" + CNULL(lc15) + _
         ",lc16=" + CNULL(lc16) + ",lc17=" + CNULL(lc17) + ",lc18=" + CNULL(lc18) + ",lc19=" + CNULL(lc19) + _
         ",lc20=" + CNULL(lc20) + ",lc21=" + CNULL(lc21) + ",lc22=" + CNULL(lc22) + ",lc23=" + CNULL(lc23) + _
         ",lc24=" + CNULL(lc24) + ",lc25=" + CNULL(lc25) + ",lc26=" + CNULL(lc26) + ",lc27=" + CNULL(lc27) + _
         ",lc34=" + CNULL(lc34) + ",lc31=" + CNULL(lc31) + ",lc32=" + CNULL(lc32) + ",lc33=" + CNULL(lc33) + _
         " where " + ChgLawcase(strlc)

cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Update502 = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Update502 = False
ErrorLog
End Function
Public Function Addnew502(ByRef lc01 As String, ByRef lc02 As String, ByRef lc03 As String, _
       ByRef lc04 As String, ByRef lc05 As String, ByRef lc06 As String, ByRef lc07 As String, _
       ByRef lc08 As String, ByRef lc09 As String, ByRef lc10 As String, ByRef lc11 As String, _
       ByRef lc12 As String, ByRef lc13 As String, ByRef lc14 As String, ByRef lc15 As String, _
       ByRef lc16 As String, ByRef lc17 As String, ByRef lc18 As String, ByRef lc19 As String, _
       ByRef lc20 As String, ByRef lc21 As String, ByRef lc22 As String, ByRef lc23 As String, _
       ByRef lc24 As String, ByRef lc25 As String, ByRef lc26 As String, ByRef lc27 As String, _
       ByRef lc34 As String, ByRef lc28 As String, ByRef lc29 As String, ByRef lc30 As String) As Boolean
       
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
       
On Error GoTo Err
cnnConnection.BeginTrans

strSQL = "insert into lawcase(lc01,lc02 ,lc03,lc04 ,lc05,lc06,lc07,lc08,lc09,lc10,lc11,lc12 , lc13 ," + _
    " lc14,lc15,lc16,lc17,lc18,lc19,lc20,lc21,lc22,lc23,lc24,lc25,lc26 ,lc27,lc34,lc28,lc29,lc30) values(" + CNULL(lc01) + "," + _
    "" + CNULL(lc02) + "," + CNULL(lc03) + "," + CNULL(lc04) + "," + CNULL(lc05) + "," + CNULL(lc06) + "," + CNULL(lc07) + _
    "," + CNULL(lc08) + "," + CNULL(lc09) + "," + CNULL(lc10) + "," + CNULL(lc11) + "," + CNULL(lc12) + "," + CNULL(lc13) + _
    "," + CNULL(lc14) + "," + CNULL(lc15) + "," + CNULL(lc16) + "," + CNULL(lc17) + "," + CNULL(lc18) + _
    "," + CNULL(lc19) + " ," + CNULL(lc20) + "," + CNULL(lc21) + "," + CNULL(lc22) + "," + CNULL(lc23) + "," + CNULL(lc24) + _
    "," + CNULL(lc25) + "," + CNULL(lc26) + "," + CNULL(lc27) + "," + CNULL(lc34) + "," + CNULL(lc28) + "," + CNULL(lc29) + "," + CNULL(lc30) + ")"

cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If
Addnew502 = True


Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans

Addnew502 = False
ErrorLog

End Function
Public Function Delete502(ByRef strlc As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans

strSQL = "delete lawcase where " & ChgLawcase(strlc)
cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Delete502 = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Delete502 = False
ErrorLog
End Function
Public Function GetReasonOfRelief(ByRef strnum As String, ByRef strResult As String) As Boolean
Dim Rs As New ADODB.Recordset
On Error GoTo Err
   
   strSQL = "select ror02 from reasonofrelief where ror01='" + strnum + "'"

Rs.CursorLocation = adUseClient
Rs.CursorType = adOpenDynamic
Rs.Open strSQL, cnnConnection
If Not (Rs.BOF And Rs.EOF) Then
    strResult = IIf(IsNull(Rs.Fields(0)), "", Rs.Fields(0))
    GetReasonOfRelief = True
Else
    ShowMsg MsgText(9050)
End If
Exit Function
Err:
GetReasonOfRelief = False
ErrorLog
End Function
Public Function GetRs504(ByRef strFields As String, ByRef strcp As String, ByRef strTbl As String, ByRef strmail As String) As ADODB.Recordset
On Error GoTo Err

strSQL = "select ' ' v," & SQLDate("CP05") & " as 收文日,cp09 as 收文號,nvl(CPM03,CPM04) AS 案件性質,S2.ST02 AS 承辦人," & _
   "S1.ST02 AS 業務員," & SQLDate("CP27") & " as 發文日,decode(cp24,1,'淮/勝',2,'駁/敗') as 結果," & _
   SQLDate("CP57") & " as 取消收文日,cp44 as 代理人, " + strFields + _
   " from staff s1,staff s2,fagent,customer,casepropertymap," + strTbl + " where " & strcp & " AND " & strmail & " AND " & _
   "cp13 =S1.st01(+) AND cp14 =S2.st01(+) AND  cp01= cpm01(+) and cp10=cpm02(+) AND " + _
   "substr(cp44,1,8)=fa01(+) and substr(cp44,8,9)=fa02(+) order by cp09"

Dim objPublicData As New clsPublicData
Set GetRs504 = objPublicData.ReadRst(strSQL, True)
Set objPublicData = Nothing
Exit Function
Err:
ErrorLog
End Function
Public Function GetRs505(ByRef strlc As String, ByRef cp09 As String) As ADODB.Recordset

On Error GoTo Err

strSQL = "select cp05,cp10,cp06,cp07,cp12,cp13,cp14,cp48,cp15,cp29,cp27,cp28,cp44,cp45,cp23," + _
   "cp24,cp25,cp35,cp08,cp43,cp57,cp58,cp26,cp31,cp21,cp22,cp64,cp16,cp17,cp18,cp19,cp33," + _
   "cp34,cp46,cp47,cp32,cp20,cp11,cp71,cp59,cp30,cp60,cp61,cp62,cp63,cp49,cp55,cp56,cp36," & _
   "cp37,cp38,cp39,cp40,cp41,cp42,cp50,cp53,cp54,cp72,cp09,cp65,cp66,cp67,cp68,cp69,cp70 " & _
   "from caseprogress where " + ChgCaseprogress(strlc) & " and cp09 in (" + cp09 + ") order by cp09"
Dim objPublicData As New clsPublicData
Set GetRs505 = objPublicData.ReadRst(strSQL, True)
Set objPublicData = Nothing
Exit Function

Err:
ErrorLog
End Function
Public Function AddNew505(ByRef strMerge As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
                
On Error GoTo Err
cnnConnection.BeginTrans

strSQL = "insert into caseprogress (cp05,cp10,cp06,cp07,cp12,cp13,cp14,cp48,cp15,cp29,cp27," + _
   "cp28,cp44,cp45,cp23,cp24,cp25,cp35,cp08,cp43,cp57,cp58,cp26,cp31,cp21,cp22,cp64," + _
   "cp16,cp17,cp18,cp19,cp33,cp34,cp46,cp47,cp32,cp20,cp11,cp71,cp59,cp30,cp60,cp61,cp62," + _
   "cp63,cp49,cp55,cp56,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp50,cp53,cp54,cp72,cp09,cp01,cp02,cp03,cp04)" + _
   "values (" + strMerge + ")"
             
cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

AddNew505 = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
AddNew505 = False
ErrorLog
End Function
Public Function Update505(ByRef strMerge As String)
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
 
On Error GoTo Err
cnnConnection.BeginTrans
strSQL = " update caseprogress set " + strMerge + ""

cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Update505 = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Update505 = False
ErrorLog
End Function
Public Function Delete505(ByRef cp09 As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans

strSQL = "delete caseprogress where cp09=" + CNULL(cp09) + " "
cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Delete505 = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Delete505 = False
ErrorLog
End Function
Public Function GetRS506() As ADODB.Recordset
On Error GoTo Err
strSQL = "select hc01,hc02,hc03,hc04,hc05,hc06,hc07,hc08,hc09," + _
       "hc10,hc11,hc12,hc12,hc13,hc14,hc15,hc16,hc17,hc18 " + _
       " from hirecase order by hc01||hc02||hc03||hc04 "

Dim objPublicData As New clsPublicData
Set GetRS506 = objPublicData.ReadRst(strSQL)
Set objPublicData = Nothing
Exit Function
       
Err:
ErrorLog

End Function
Public Function Addnew506(ByRef hc01 As String, ByRef hc02 As String, ByRef hc03 As String, ByRef hc04 As String, ByRef hc05 As String, ByRef hc06 As String, hc07 As String, hc08 As String, hc09 As String, _
       ByRef hc10 As String, ByRef hc11 As String, ByRef hc12 As String, ByRef hc13 As String, ByRef hc14 As String, ByRef hc15 As String, ByRef hc19 As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans

strSQL = "insert into hirecase(hc01,hc02,hc03,hc04,hc05,hc06,hc07,hc08,hc09," + _
       "hc10,hc11,hc12,hc13,hc14,hc15,hc19) values (" + CNULL(hc01) + "," + _
    "" + CNULL(hc02) + "," + CNULL(hc03) + "," + CNULL(hc04) + "," + CNULL(hc05) + "," + CNULL(hc06) + "," + CNULL(hc07) + _
    "," + CNULL(hc08) + "," + CNULL(hc09) + "," + CNULL(hc10) + "," + CNULL(hc11) + "," + CNULL(hc12) + "," + CNULL(hc13) + _
    "," + CNULL(hc14) + "," + CNULL(hc15) + "," + CNULL(hc19) + ")"

cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Addnew506 = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Addnew506 = False
ErrorLog
End Function
Public Function Update506(ByRef strlc As String, ByRef hc05 As String, ByRef hc06 As String, ByRef hc07 As String, _
       ByRef hc08 As String, ByRef hc09 As String, ByRef hc10 As String, ByRef hc11 As String, ByRef hc12 As String, _
       ByRef hc16 As String, ByRef hc17 As String, ByRef hc18 As String) As Boolean

'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans

strSQL = " update hirecase set hc05=" + CNULL(hc05) + ",hc06=" + CNULL(hc06) + ",hc07=" + CNULL(hc07) + _
         ",hc08=" + CNULL(hc08) + ",hc09=" + CNULL(hc09) + ",hc10=" + CNULL(hc10) + ",hc11=" + CNULL(hc11) + _
         ",hc12=" + CNULL(hc12) + ",hc16=" + CNULL(hc16) + ",hc17=" + CNULL(hc17) + ",hc18=" + CNULL(hc18) + _
         " where " + ChgHirecase(strlc)

cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Update506 = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Update506 = False
ErrorLog
End Function
Public Function Delete506(ByRef hc As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans

strSQL = "delete hirecase where " + ChgHirecase(hc)
cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Delete506 = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Delete506 = False
ErrorLog
End Function
Public Function Get506HireDate(ByRef cp As String, ByRef strFirstDay As String, _
              ByRef cp53 As String, ByRef cp54 As String) As Boolean
Dim rst As New ADODB.Recordset

On Error GoTo Err
strSQL = "select cp53,cp54 from caseprogress where " & ChgCaseprogress(cp) & " order by cp05 asc"

rst.CursorLocation = adUseClient
rst.CursorType = adOpenStatic
rst.Open strSQL, cnnConnection
If Not (rst.EOF And rst.BOF) Then
    rst.MoveFirst
    strFirstDay = IIf(IsNull(rst.Fields!cp53), 0, rst.Fields!cp53)
    rst.MoveLast
    cp53 = IIf(IsNull(rst.Fields!cp53), 0, rst.Fields!cp53)
    rst.MoveLast
    cp54 = IIf(IsNull(rst.Fields!cp54), 0, rst.Fields!cp54)
    Get506HireDate = True
   
End If
 Set rst = Nothing
Exit Function
Err:
Get506HireDate = False
ErrorLog
End Function
Public Function GetRs507(ByRef strlc As String, ByRef strSysKind As String) As ADODB.Recordset
 Dim strCaseName As String, i As Integer, strTmp As String
 Dim objPublicData As New clsPublicData
On Error GoTo Err

If objPublicData.GetSystemKind(strSysKind, i, strCaseName) Then
 strSQL = "select distinct  ''  v, np01 收文號," & ChgNextProgress("", 1) & " 本所案號," + _
           "nvl(CPM03,CPM04) 下一程序," & SQLDate("NP08") & " 本所期限," + _
           SQLDate("NP09") & " 法定期限,ST02 業務員,np06 是否續辦,np16,np17,np18,"
strTmp = "and cp09=np01 and np02=cpm01(+) and np07=cpm02(+) and np10=st01(+) order by np01,本所案號,本所期限,法定期限"
Select Case i
Case 1
   strSQL = strSQL + "pa05,pa06,pa07,np07,np22  from patent,nextprogress ,casepropertymap,staff ,caseprogress " + _
     " where " + ChgPatent(strlc) + " and " & ChgNextProgress(strlc) & strTmp
Case 2
   strSQL = strSQL + "tm05,tm06,tm07,,np07,np22 from trademark,nextprogress ,casepropertymap,staff,caseprogress where " + _
      ChgTradeMark(strlc) + " and " & ChgNextProgress(strlc) & strTmp

Case 3
    strSQL = strSQL + "lc05,lc06,lc07,np07,np22 from lawcase,nextprogress,casepropertymap,staff,caseprogress " + _
      "where " + ChgLawcase(strlc) + " and " & ChgNextProgress(strlc) & strTmp

Case 4
    strSQL = strSQL + "hc05,hc06,'',np07,np22 from hirecase,nextprogress,casepropertymap,staff,caseprogress" + _
      "  where " + ChgHirecase(strlc) + " and " & ChgNextProgress(strlc) & strTmp
Case 5, 6, 7, 8
    strSQL = strSQL + "sp05,sp06,sp07,np07,np22 from servicepractice,nextprogress,casepropertymap,staff,caseprogress " + _
      "  where " + ChgService(strlc) + " and " & ChgNextProgress(strlc) & strTmp
End Select
Set GetRs507 = objPublicData.ReadRst(strSQL, True)
Set objPublicData = Nothing
Else
    Set GetRs507 = Nothing
End If
Exit Function
Err:
ErrorLog
End Function
Public Function GetRs508(ByRef strkey1 As String, ByRef strkey2 As String, ByRef strcp As String, ByRef strKind As String) As ADODB.Recordset
Dim strCaseName As String, i As Integer
Dim objPublicData As New clsPublicData
On Error GoTo Err

If objPublicData.GetSystemKind(strKind, i, strCaseName) Then
 strSQL = "select np01,np02,np03,np04,np05,np06,np07,np08,np09,np10,np11,np12,np13," + _
      " np14,np15,np16,np17,np18,np19,np20,np21,cp05,cp14,cp27 ,"
Select Case i
Case "1"
   strSQL = strSQL + "pa05,pa06,pa07,nvl(pa26,nvl(pa27,nvl(pa28,nvl(pa29,pa30))))  from patent,caseprogress," + _
       " nextprogress where np01=cp09 and np02=pa01 and np03=pa02 and np04=pa03 and np05=pa04"
                  
 
Case "2"
   strSQL = strSQL + "tm05,tm06,tm07,tm23  from trademark,nextprogress,caseprogress where " + _
      "  np01=cp09  and np02=tm01 and np03=tm02 and np04=tm03 and np05=tm04"

                   
                  
Case "3"
   strSQL = strSQL + "lc05,lc06,lc07,lc11 from lawcase,caseprogress,nextprogress where " + _
      "  np01=cp09  and np02=lc01 and np03=lc02 and np04=lc03 and np05=lc04"

Case "4"
   strSQL = strSQL + "hc05,hc06 from hirecase,nextprogress,caseprogress  where  " + _
      "  np01=cp09  np02=hc01 and np03=hc02 and np04=hc03 and np05=hc04"
         
                 
End Select
   strSQL = strSQL + " and np01 in (" + strcp + ")" + " and np07 in (" + strkey1 + ")" + " and np22 in (" + strkey2 + ")" + _
      " order by np01,np02||np03||np04||np05"
                
Set GetRs508 = objPublicData.ReadRst(strSQL, True)
Set objPublicData = Nothing
End If
Exit Function
Err:
ErrorLog
End Function
Public Function AddNew508(ByRef np01 As String, ByRef np02 As String, ByRef np03 As String, ByRef np04 As String, _
         ByRef np05 As String, ByRef np06 As String, ByRef NP07 As String, ByRef np08 As String, ByRef np09 As String, _
         ByRef np10 As String, ByRef np11 As String, ByRef np12 As String, ByRef np13 As String, ByRef np14 As String, _
         ByRef np15 As String, ByRef np22 As String) As Boolean

'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans
   'Modify By Cheng 2002/09/25
' strSQL = "insert into nextprogress(np01,np02,np03,np04,np05,np06,np07,np08,np09,np10," + _
'   " np11,np12,np13,np14,np15,np22) values(" + CNULL(np01) + "," + CNULL(np02) + _
'   "," + CNULL(np03) + "," + CNULL(np04) + "," + CNULL(np05) + "," + CNULL(np06) + "," + CNULL(NP07) + _
'   "," + CNULL(np08) + "," + CNULL(np09) + "," + CNULL(np10) + "," + CNULL(np11) + "," + CNULL(np12) + _
'   "," + CNULL(np13) + "," + CNULL(np14) + "," + CNULL(np15) + "," + CNULL(np22) + ")"
 strSQL = "insert into nextprogress(np01,np02,np03,np04,np05,np06,np07,np08,np09,np10," + _
   " np11,np12,np13,np14,np15,np22) values(" + CNULL(np01) + "," + CNULL(np02) + _
   "," + CNULL(np03) + "," + CNULL(np04) + "," + CNULL(np05) + "," + CNULL(np06) + "," + CNULL(NP07) + _
   "," + CNULL(np08) + "," + CNULL(np09) + "," + CNULL(np10) + "," + CNULL(np11) + "," + CNULL(np12) + _
   "," + CNULL(np13) + "," + CNULL(ChgSQL(np14)) + "," + CNULL(np15) + "," + CNULL(np22) + ")"
          
cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

AddNew508 = True

Exit Function

Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
AddNew508 = False
ErrorLog
End Function
Public Function Update508(ByRef np06 As String, _
         ByRef np08 As String, ByRef np09 As String, ByRef np10 As String, ByRef np11 As String, ByRef np12 As String, _
         ByRef np13 As String, ByRef np14 As String, ByRef np15 As String, _
         ByRef np01 As Variant, ByRef NP07 As Variant, ByRef np22 As Variant) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans
'Modify By Cheng 2002/09/25
'strSQL = "update nextprogress set np06=" + CNULL(np06) + ",np08=" + CNULL(np08) + _
'   " ,np09=" + CNULL(np09) + ",np10=" + CNULL(np10) + ",np11=" + CNULL(np11) + ",np12=" + CNULL(np12) + _
'   ", np13=" + CNULL(np13) + ",np14=" + CNULL(np14) + ",np15=" + CNULL(np15) + _
'   " where np01=" + np01 + " and np07=" + NP07 + " and np22 =" + np22
strSQL = "update nextprogress set np06=" + CNULL(np06) + ",np08=" + CNULL(np08) + _
   " ,np09=" + CNULL(np09) + ",np10=" + CNULL(np10) + ",np11=" + CNULL(np11) + ",np12=" + CNULL(np12) + _
   ", np13=" + CNULL(np13) + ",np14=" + CNULL(ChgSQL(np14)) + ",np15=" + CNULL(np15) + _
   " where np01=" + np01 + " and np07=" + NP07 + " and np22 =" + np22
        
cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Update508 = True

Exit Function

Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Update508 = False
ErrorLog
End Function
Public Function Delete508(ByRef Stra As Variant, ByRef strb As Variant, ByRef strc As Variant) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans
strSQL = "delete nextprogress where np01=" + Stra + " and np07=" + strb + " and np22=" + strc

cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Delete508 = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Delete508 = False
ErrorLog
End Function

Public Function GetRS510() As ADODB.Recordset
On Error GoTo Err
strSQL = "select or01,or02,or03,or04,or05,or06 from organization "

Dim objPublicData As New clsPublicData
Set GetRS510 = objPublicData.ReadRst(strSQL, True)
Set objPublicData = Nothing
Exit Function
       
Err:
ErrorLog

End Function
Public Function Addnew510(ByRef or01 As String, ByRef or02 As String, ByRef or03 As String, _
     ByRef or04 As String, ByRef or05 As String, ByRef or06 As String) As Boolean

'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans

strSQL = "insert into organization(or01,or02,or03,or04,or05,or06) values(" + CNULL(or01) + "," + _
    "" + CNULL(or02) + "," + CNULL(or03) + "," + CNULL(or04) + "," + CNULL(or05) + "," + CNULL(or06) + ")"

cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Addnew510 = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Addnew510 = False
ErrorLog

End Function
Public Function Update510(ByRef or01 As String, ByRef or02 As String, ByRef or03 As String, _
      ByRef or04 As String, ByRef or05 As String, ByRef or06 As String) As Boolean

'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans

strSQL = " update organization set or02=" + CNULL(or02) + ",or03=" + CNULL(or03) + ",or04=" + CNULL(or04) + _
   ",or05=" + CNULL(or05) + ",or06=" + CNULL(or06) + " where or01=" + CNULL(or01) + ""

cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Update510 = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Update510 = False
ErrorLog
End Function
Public Function Delete510(ByRef hc As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans

strSQL = "delete lawcase where hc01=" + CNULL(hc) + " "
cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Delete510 = True
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Delete510 = False
ErrorLog
End Function
Public Function GetRS511(ByRef i As Integer, ByRef txt As String) As ADODB.Recordset

On Error GoTo Err
Select Case i
Case 1
strSQL = "select '' v,decode(cp05,null,'',SUBSTR(CP05, 1, 4) - 1911 || '/' || SUBSTR(CP05, 5, 2) || '/' ||" + _
        " SUBSTR(CP05, 7, 2))  as 收受日,cdp01 as 收文號,decode(cdp03,null," + _
        " '',SUBSTR(cdp03, 1, 4) - 1911 || '/' || SUBSTR(cdp03, 5, 2) || '/' || SUBSTR(cdp03, 7, 2)) " + _
        " as 開庭日期,cdp04 as 時間, decode(cdp05,or01,or02) 法院,decode(cdp02, ST01, ST02) 開庭人員 ," + _
        " decode(cdp06,'1','偵查','2','審查','3','辯論') 開庭性質,lc05,lc06,lc07,lc11 from" + _
         " lawcase,caseprogress,courtyardperiod,organization,staff where " & ChgLawcase(txt) & _
         " and cp01||cp02||cp03||cp04=lc01||lc02||lc03||lc04 and cp09=cdp01 and cdp05=or01(+) and cdp02=st01(+)"
Case 2
strSQL = "select '' v,decode(cp05,null,'',SUBSTR(CP05, 1, 4) - 1911 || '/' || SUBSTR(CP05, 5, 2) || '/' ||" + _
        " SUBSTR(CP05, 7, 2))  as 收受日,cdp01 as 收文號,decode(cdp03,null,''," + _
        " SUBSTR(cdp03, 1, 4) - 1911 || '/' || SUBSTR(cdp03, 5, 2) || '/' || SUBSTR(cdp03, 7, 2)) " + _
        " as 開庭日期,cdp04 as 時間, decode(cdp05,or01,or02) 法院,decode(cdp02, ST01, ST02) 開庭人員 ," + _
        " decode(cdp06,'1','偵查','2','審查','3','辯論') 開庭性質,hc05,hc06 from" + _
         " caseprogress,hirecase where " + ChgHirecase(txt) + _
         " and cp01||cp02||cp03||cp04=hc01||hc02||hc03||hc04 and cp09=cdp01 and cdp05=or01(+) and cdp02=st01(+)"
End Select

Dim objPublicData As New clsPublicData
Set GetRS511 = objPublicData.ReadRst(strSQL, True)
Set objPublicData = Nothing
Exit Function

Err:
ErrorLog
End Function
Public Function GetRS512(ByRef strlc As String, ByRef cp09 As String) As ADODB.Recordset
On Error GoTo Err

 strSQL = "select cp05,cdp05,cp13,cp35,cp30,cdp01,cdp02,cdp03,cdp04,cdp06,cdp07,cdp08,cdp09,lc05,lc06,lc07," + _
         "lc11 from caseprogress,courtyardperiod,lawcase where cdp01 in (" + cp09 + ") and cdp01=cp09 and " + _
         ChgLawcase(strlc) + " and lc01||lc02||lc03||lc04=cp01||cp02||cp03||cp04 order by cdp01"

Dim objPublicData As New clsPublicData
Set GetRS512 = objPublicData.ReadRst(strSQL, True)
Set objPublicData = Nothing
Exit Function
       
Err:
ErrorLog

End Function
Public Function Update512(ByRef cdp01 As String, ByRef cdp02 As String, ByRef cdp03 As String, ByRef cdp04 As String, _
          ByRef cdp06 As String, ByRef cdp07 As String, ByRef cdp08 As String, ByRef cdp09 As String, ByRef cp05 As String, _
          ByRef cp71 As String, ByRef cp30 As String, ByRef cp35 As String) As Boolean

'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans

strSQL = " update courtyardperiod set cdp02=" + CNULL(cdp02) + ",cdp03=" + CNULL(cdp03) + ",cdp04=" + CNULL(cdp04) + _
   ",cdp06=" + CNULL(cdp06) + ",cdp07=" + CNULL(cdp07) + ",cdp08=" + CNULL(cdp08) + ",cdp09=" + CNULL(cdp09) + _
   " where CDP01=" + CNULL(cdp01) + ""
cnnConnection.Execute strSQL
         
strSQL = "update caseprogress set cp05=" + CNULL(cp05) + ",cp71=" + CNULL(cp71) + ",cp30=" + CNULL(cp30) + _
       ",cp35=" + CNULL(cp35) + " where cp09=" + CNULL(cdp01) + ""

cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Update512 = True
Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Update512 = False
ErrorLog
End Function
Public Function Delete512(ByRef cdp01 As String) As Boolean
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True

On Error GoTo Err
cnnConnection.BeginTrans

strSQL = "delete courtyardperiod where cdp01=" + CNULL(cdp01) + " "
cnnConnection.Execute strSQL
'Modify By Cheng 2002/11/14
If BolTransOk Then
    cnnConnection.CommitTrans
End If

Delete512 = True
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

cnnConnection.RollbackTrans
Delete512 = False
ErrorLog
End Function
Public Function GetCaseProperty(ByRef kind As Integer, ByRef strRecieveKind As String, ByRef strproperty As String, ByRef strPropertyName As String) As Boolean
Dim strSQL As String, rsRecordset As New ADODB.Recordset

On Error GoTo Err
Select Case kind
Case 1
 strSQL = "select cpm03 from casepropertymap where cpm01='" + strRecieveKind + "' and cpm02='" + strproperty + "' "
Case 2
 strSQL = "select cpm04 from casepropertymap where cpm01='" + strRecieveKind + "' and cpm02='" + strproperty + "' "
End Select

rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSQL, cnnConnection
If rsRecordset.RecordCount > 0 Then
   strPropertyName = rsRecordset.Fields(0)
   GetCaseProperty = True
Else
   MsgBox "案件性質代碼錯誤!!", vbCritical
End If
rsRecordset.Close
Exit Function
Err:
ErrorLog
End Function
Public Function GetStaffDeptName(ByRef strSDNum As String, ByRef strSDName As String) As Boolean
Dim strSQL As String, rsRecordset As New ADODB.Recordset
On Error GoTo Err
strSQL = "select a0902 from acc090 where a0901='" + strSDNum + "'"
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSQL, cnnConnection
If rsRecordset.BOF And rsRecordset.EOF Then
   MsgBox "部門代號錯誤"
   GetStaffDeptName = False
Else
  strSDName = rsRecordset.Fields(0)
  GetStaffDeptName = True
End If

Exit Function
Err:
  GetStaffDeptName = False
ErrorLog
End Function
'取得國外代理人名稱
Public Function GetFAgent(ByRef AgentNum As String, ByRef AgentName As String) As Boolean
Dim strSQL As String, rsRecordset As New ADODB.Recordset

On Error GoTo Err
strSQL = "select nvl(fa04,nvl(fa05,fa06)) from fagent where " + ChgFagent(AgentNum)
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSQL, cnnConnection
If rsRecordset.BOF And rsRecordset.EOF Then
   MsgBox "代理人代號錯誤"
   GetFAgent = False
Else
  AgentName = rsRecordset.Fields(0)
  GetFAgent = True
End If
Exit Function
Err:
ErrorLog
End Function
'取得機關名稱
Public Function GetGovName(ByRef strGov As String, ByRef strGovName As String) As Boolean
Dim strSQL As String, Rs As New ADODB.Recordset

On Error GoTo Err

strSQL = "select or02 from organization where or01='" + strGov + "'"

Rs.CursorLocation = adUseClient
Rs.CursorType = adOpenStatic
Rs.Open strSQL, cnnConnection
If Not (Rs.BOF And Rs.EOF) Then
   strGovName = Rs.Fields!or02
   GetGovName = True
Else
   MsgBox "無此資料"
   GetGovName = False
End If

Exit Function
Err:
ErrorLog

End Function
Public Function GetRelation(ByRef strlc As String, ByRef strRelation As String) As Boolean
Dim strSQL As String, rsRecordset As New ADODB.Recordset
On Error GoTo Err
strSQL = "select cp09 from caseprogress where " + ChgCaseprogress(strlc) + " and cp09='" + strRelation + "'"

rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSQL, cnnConnection
If rsRecordset.BOF And rsRecordset.EOF Then
   MsgBox "相關總收文號錯誤", vbCritical
   GetRelation = False
Else
  GetRelation = True
End If

Exit Function
Err:
  GetRelation = False
ErrorLog
End Function
