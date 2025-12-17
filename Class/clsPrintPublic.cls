VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsPrintPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Modify By Cheng 2002/12/05
'多加傳入參數聯絡單備註, 申請國家, 申請人
'Public Function PrintEmail(ByRef intCaseKind As Integer, ByRef intWhere As Integer, ByRef strReceiveCode As String, ByRef strUserName As String) As Boolean
Public Function PrintEmail(ByRef intCaseKind As Integer, ByRef intWhere As Integer, ByRef strReceiveCode As String, ByRef strUserName As String, _
    Optional ByRef strNote As String = "", Optional ByRef strNation As String = "", Optional ByRef strCustomer As String = "") As Boolean
 On Error GoTo Err

   DataEnvPublic.Commands(1).CommandText = GetSQL(intCaseKind, intWhere, strReceiveCode)
   DataEnvPublic.cmdContact
   datrptPublic.Sections(1).Controls("rptlblUserName").Caption = strUserName
   'Add By Cheng 2002/12/05
    '若有傳入聯絡單備註
   If strNote <> "" Then
       datrptPublic.Sections(1).Controls("Label15").Caption = strNote
   End If
   '申請國家
    datrptPublic.Sections(1).Controls("lblNation").Caption = strNation
   '申請人
    datrptPublic.Sections(1).Controls("lblCustomer").Caption = strCustomer
   
    'Modify By Cheng 2002/12/05
   datrptPublic.PrintReport
   Unload datrptPublic
   DataEnvPublic.rscmdContact.Close
   PrintEmail = True
   Exit Function
Err:
   ErrorLog
End Function

Private Sub Class_Initialize()
   Set cnnConnection = DataEnvPublic.Connection
End Sub

Private Sub Class_Terminate()
   cnnConnection.Close
   Set DataEnvPublic = Nothing
End Sub

'Public Function PrtFCPMail(ByVal CP09 As String, ByVal strUserName As String) As Boolean
Public Function PrtFCPMail(ByVal CP09 As String, ByVal strUserName As String, Optional ByVal bShow As Boolean = True) As Boolean
 Dim strTmp As String
'Add By Cheng 2002/12/05
 On Error GoTo Err
   
   DataEnvPublic.Commands(2).CommandText = "select st03 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
      "pa07 a06,cu05||cu88 a07,FA05||FA63 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
      "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN WHERE cp09='" & CP09 & "' AND " & _
      "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
      "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
      "PA11=TPB01(+)"
   DataEnvPublic.cmd060302
   'datrpt060302.Orientation = rptOrientPortrait
   datrpt060302.Sections(1).Controls("rptlblUserName").Caption = strUserName
    'Add By Cheng 2002/12/20
    Screen.MousePointer = vbDefault
   ' 90.08.16 modify by louis
   If bShow = True Then
      datrpt060302.Show vbModal
   End If
   datrpt060302.PrintReport
   Unload datrpt060302
   DataEnvPublic.rscmdContact.Close
   PrtFCPMail = True
   Exit Function
Err:
   ErrorLog
End Function

'Add By Cheng 2003/01/09
Public Function PrtPMail(ByVal CP09 As String, ByVal strUserName As String, Optional ByVal bShow As Boolean = True) As Boolean
 Dim strTmp As String
 On Error GoTo Err
   
   DataEnvPublic.Commands(2).CommandText = "select st03 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
      "pa07 a06,cu04 a07,FA04 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
      "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN WHERE cp09='" & CP09 & "' AND " & _
      "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
      "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
      "PA11=TPB01(+)"
   DataEnvPublic.cmd060302
   'datrpt060302.Orientation = rptOrientPortrait
   datrpt060302.Sections(1).Controls("rptlblUserName").Caption = strUserName
    'Add By Cheng 2002/12/20
    Screen.MousePointer = vbDefault
   ' 90.08.16 modify by louis
   If bShow = True Then
      datrpt060302.Show vbModal
   End If
   datrpt060302.PrintReport
   Unload datrpt060302
   DataEnvPublic.rscmdContact.Close
   PrtPMail = True
   Exit Function
Err:
   ErrorLog
End Function

'Add By Cheng 2002/06/26
Public Function PrtMail(ByVal CP09 As String, ByVal strUserName As String, Optional ByVal bShow As Boolean = True, Optional strNote As String, Optional strDept As String, Optional strDate As String) As Boolean
 Dim strTmp As String
'Add By Cheng 2002/12/05
On Error GoTo Err

   DataEnvPublic.Commands(2).CommandText = "select A0902 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
      "pa07 a06,cu04 a07,FA05||FA63 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
      "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN,ACC090 WHERE ST03=A0901 AND cp09='" & CP09 & "' AND " & _
      "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
      "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
      "PA11=TPB01(+)"
   DataEnvPublic.cmd060302
   datrptPublicA.Sections(1).Controls("rptlblUserName").Caption = strUserName
   datrptPublicA.Sections(1).Controls("Label15").Caption = strNote
   datrptPublicA.Sections(1).Controls("Label14").Caption = strDept
   datrptPublicA.Sections(1).Controls("rptlblDate").Caption = strDate
    'Add By Cheng 2002/12/20
    Screen.MousePointer = vbDefault
   If bShow = True Then
      datrptPublicA.Show vbModal
   Else
      datrptPublicA.PrintReport
   End If
   Unload datrptPublicA
   DataEnvPublic.rscmd060302.Close
   PrtMail = True
   Exit Function
Err:
   ErrorLog
End Function

'Add By Cheng 2002/10/14
Public Function PrtMail_1(ByVal strCPNO01 As String, ByVal strCPNO02 As String, ByVal strCPNO03 As String, ByVal strCPNO04 As String, _
      ByVal strUserName As String, Optional ByVal bShow As Boolean = True, Optional strNote As String, Optional strDept As String, Optional strDate As String, Optional strSaleZone As String, Optional strSaleName As String, _
      Optional strCPName1 As String, Optional strCPName2 As String, Optional strCPName3 As String) As Boolean
   Dim strTmp As String
   Dim strSQLA As String
     
'Add By Cheng 2002/12/05
On Error GoTo Err
     
    'Modify By Cheng 2002/12/05
'   strSQLA = "select CU04 FROM PATENT,CUSTOMER " & _
'      "WHERE PA01='" & strCPNO01 & "' AND PA02='" & strCPNO02 & "' AND PA03='" & strCPNO03 & "' AND PA04='" & strCPNO04 & "' " & _
'      "AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & _
'      " UNION select CU04 FROM TRADEMARK,CUSTOMER " & _
'      "WHERE TM01='" & strCPNO01 & "' AND TM02='" & strCPNO02 & "' AND TM03='" & strCPNO03 & "' AND TM04='" & strCPNO04 & "' " & _
'      "AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) " & _
'      " UNION select CU04 FROM LAWCASE,CUSTOMER " & _
'      "WHERE LC01='" & strCPNO01 & "' AND LC02='" & strCPNO02 & "' AND LC03='" & strCPNO03 & "' AND LC04='" & strCPNO04 & "' " & _
'      "AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) " & _
'      " UNION select CU04 FROM HIRECASE,CUSTOMER " & _
'      "WHERE HC01='" & strCPNO01 & "' AND HC02='" & strCPNO02 & "' AND HC03='" & strCPNO03 & "' AND HC04='" & strCPNO04 & "' " & _
'      "AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) " & _
'      " UNION select CU04 FROM SERVICEPRACTICE,CUSTOMER " & _
'      "WHERE SP01='" & strCPNO01 & "' AND SP02='" & strCPNO02 & "' AND SP03='" & strCPNO03 & "' AND SP04='" & strCPNO04 & "' " & _
'      "AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) "
   strSQLA = "select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03 FROM PATENT,CUSTOMER,NATION " & _
      "WHERE PA01='" & strCPNO01 & "' AND PA02='" & strCPNO02 & "' AND PA03='" & strCPNO03 & "' AND PA04='" & strCPNO04 & "' " & _
      "AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND PA09=NA01(+) " & _
      " UNION select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03 FROM TRADEMARK,CUSTOMER,NATION " & _
      "WHERE TM01='" & strCPNO01 & "' AND TM02='" & strCPNO02 & "' AND TM03='" & strCPNO03 & "' AND TM04='" & strCPNO04 & "' " & _
      "AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND TM10=NA01(+) " & _
      " UNION select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03 FROM LAWCASE,CUSTOMER,NATION " & _
      "WHERE LC01='" & strCPNO01 & "' AND LC02='" & strCPNO02 & "' AND LC03='" & strCPNO03 & "' AND LC04='" & strCPNO04 & "' " & _
      "AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND LC15=NA01(+) " & _
      " UNION select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03 FROM HIRECASE,CUSTOMER,NATION " & _
      "WHERE HC01='" & strCPNO01 & "' AND HC02='" & strCPNO02 & "' AND HC03='" & strCPNO03 & "' AND HC04='" & strCPNO04 & "' " & _
      "AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) AND '000' = NA01(+) " & _
      " UNION select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03 FROM SERVICEPRACTICE,CUSTOMER,NATION " & _
      "WHERE SP01='" & strCPNO01 & "' AND SP02='" & strCPNO02 & "' AND SP03='" & strCPNO03 & "' AND SP04='" & strCPNO04 & "' " & _
      "AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND SP09=NA03(+) "
   
   DataEnvPublic.Commands("cmd1106").CommandText = strSQLA
   DataEnvPublic.cmd1106
   datrptPublicB.Sections(1).Controls("lblA01").Caption = strSaleZone
   datrptPublicB.Sections(1).Controls("lblA02").Caption = strSaleName
   datrptPublicB.Sections(1).Controls("lblA03").Caption = strCPNO01 & "-" & strCPNO02 & "-" & strCPNO03 & "-" & strCPNO04
   datrptPublicB.Sections(1).Controls("lblA04").Caption = strCPName1
   datrptPublicB.Sections(1).Controls("lblA05").Caption = strCPName2
   datrptPublicB.Sections(1).Controls("lblA06").Caption = strCPName3
   datrptPublicB.Sections(1).Controls("rptlblUserName").Caption = strUserName
   datrptPublicB.Sections(1).Controls("Label15").Caption = strNote
   datrptPublicB.Sections(1).Controls("Label14").Caption = strDept
   datrptPublicB.Sections(1).Controls("rptlblDate").Caption = strDate
    'Add By Cheng 2002/12/20
    Screen.MousePointer = vbDefault
   If bShow = True Then
      datrptPublicB.Show vbModal
   Else
      datrptPublicB.PrintReport
   End If
   Unload datrptPublicB
   DataEnvPublic.rscmd1106.Close
   PrtMail_1 = True
   Exit Function
Err:
   ErrorLog
End Function

'Add By Cheng 2002/06/26
'Modify By Cheng 2002/12/20
'Public Function PrtMail_2(ByVal CP09 As String, ByVal strUserName As String, Optional ByVal bShow As Boolean = True, Optional strNote As String, Optional strDept As String, Optional strDate As String) As Boolean
Public Function PrtMail_2(ByVal CP09 As String, ByVal strUserName As String, Optional ByVal bShow As Boolean = True, Optional strNote As String, Optional strDept As String, Optional strDate As String, Optional strDate1 As String) As Boolean
 Dim strTmp As String
'Add By Cheng 2002/12/05
On Error GoTo Err

   DataEnvPublic.Commands(2).CommandText = "select A0902 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
      "pa07 a06,cu04 a07,FA05||FA63 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
      "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN,ACC090 WHERE ST03=A0901 AND cp09='" & CP09 & "' AND " & _
      "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
      "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
      "PA11=TPB01(+)"
   DataEnvPublic.cmd060302
   datrptPublicC.Sections(1).Controls("rptlblUserName").Caption = strUserName
   datrptPublicC.Sections(1).Controls("Label15").Caption = strNote
   datrptPublicC.Sections(1).Controls("Label14").Caption = strDept
   datrptPublicC.Sections(1).Controls("rptlblDate").Caption = strDate
   datrptPublicC.Sections(1).Controls("rptlblDate1").Caption = strDate1
    'Add By Cheng 2002/12/18
   datrptPublicC.Sections(1).Controls("rptlblUserName_1").Caption = strUserName
   datrptPublicC.Sections(1).Controls("Label15_1").Caption = strNote
   datrptPublicC.Sections(1).Controls("Label14_1").Caption = strDept
   datrptPublicC.Sections(1).Controls("rptlblDate_1").Caption = strDate
   datrptPublicC.Sections(1).Controls("rptlblDate1_1").Caption = strDate1
    'Add By Cheng 2002/12/20
    Screen.MousePointer = vbDefault
   If bShow = True Then
      datrptPublicC.Show vbModal
   Else
      datrptPublicC.PrintReport
   End If
   Unload datrptPublicC
   DataEnvPublic.rscmd060302.Close
   PrtMail_2 = True
   Exit Function
Err:
   ErrorLog
End Function

'Add By Cheng 2002/12/20
Public Function PrtMail_3(ByVal CP09 As String, ByVal strUserName As String, Optional ByVal bShow As Boolean = True, Optional strNote As String, Optional strDept As String, Optional strDate As String, Optional strDate1 As String) As Boolean
 Dim strTmp As String
On Error GoTo Err

   DataEnvPublic.Commands(2).CommandText = "select A0902 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
      "pa07 a06,cu04 a07,FA05||FA63 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
      "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN,ACC090 WHERE ST03=A0901 AND cp09='" & CP09 & "' AND " & _
      "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
      "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
      "PA11=TPB01(+)"
   DataEnvPublic.cmd060302
   datrptPublicD.Sections(1).Controls("rptlblUserName").Caption = strUserName
   datrptPublicD.Sections(1).Controls("Label15").Caption = strNote
   datrptPublicD.Sections(1).Controls("Label14").Caption = strDept
   datrptPublicD.Sections(1).Controls("rptlblDate").Caption = strDate
   datrptPublicD.Sections(1).Controls("rptlblDate1").Caption = strDate1
    'Add By Cheng 2002/12/20
    Screen.MousePointer = vbDefault
   If bShow = True Then
      datrptPublicD.Show vbModal
   Else
      datrptPublicD.PrintReport
   End If
   Unload datrptPublicD
   DataEnvPublic.rscmd060302.Close
   PrtMail_3 = True
   Exit Function
Err:
   ErrorLog
End Function

