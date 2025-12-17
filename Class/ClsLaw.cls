VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsLaw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim StrQty As String
'取得員工名稱

'取得機關名稱
Public Function GetGovName(ByRef strGov As String, ByRef strGovName As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   strGovName = ""
   StrQty = "select or02 from organization where or01='" + strGov + "'"
   RsTemp.Open StrQty, cnnConnection
   GetGovName = False
   Do While Not RsTemp.EOF
      If Not IsNull(RsTemp.Fields(0).Value) Then strGovName = RsTemp.Fields(0).Value
      GetGovName = True
      Exit Do
   Loop
   RsTemp.Close
   If GetGovName = False Then MsgBox "無此機關代號資料"
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'取得部門名稱
Public Function GetStaffDeptName(ByRef strSDNum As String, ByRef strSDName As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   strSDName = ""
   StrQty = "select a0902 from acc090 where a0901='" + strSDNum + "'"
   RsTemp.Open StrQty, cnnConnection
   GetStaffDeptName = False
   Do While Not RsTemp.EOF
     If Not IsNull(RsTemp.Fields(0).Value) Then strSDName = RsTemp.Fields(0).Value
     GetStaffDeptName = True
     Exit Do
   Loop
   RsTemp.Close
   If GetStaffDeptName = False Then MsgBox "部門代號錯誤", vbCritical
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function GetExpand(ByVal strNum As String, ByRef StrNam As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   StrNam = ""
   StrQty = "SELECT ECA02 FROM EXPANDCUSATTR WHERE ECA01 ='" & strNum & "'"
   RsTemp.Open StrQty, cnnConnection
   GetExpand = False
   Do While Not RsTemp.EOF
      If Not IsNull(RsTemp.Fields(0).Value) Then StrNam = RsTemp.Fields(0).Value
      GetExpand = True
      Exit Do
   Loop
   RsTemp.Close
   If GetExpand = False Then MsgBox "屬性代號錯誤", vbCritical
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function GetRelation(ByRef strlc As String, ByRef strRelcp09 As String, ByRef strRelation As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   StrQty = "select cp09 from caseprogress where " & ChgCaseprogress(strlc) & " and cp09='" + strRelation + "'"
   RsTemp.Open StrQty, cnnConnection
   GetRelation = False
   Do While Not RsTemp.EOF
      GetRelation = True
      Exit Do
   Loop
   RsTemp.Close
   If GetRelation = False Then MsgBox "非此案之其他收文號", vbCritical
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function ChkCaseNum(ByVal cp01 As String, ByVal cp02 As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   StrQty = "select au03 from autonumber where au01=" + CNULL(cp01) + " order by  au03 desc"
   RsTemp.Open StrQty, cnnConnection
   ChkCaseNum = True
   Do While Not RsTemp.EOF
      If Val(RsTemp.Fields!au03) - Val(cp02) >= 0 Then ChkCaseNum = False
      Exit Do
   Loop
   RsTemp.Close
   If ChkCaseNum = True Then MsgBox "輸入之流水號大於目前最大之流水號", vbCritical
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'cp(1),cp(2),cp(3),cp(4) 本所案號
'cp(5) 案件性質代號
'cp(6) 案件性質名稱
'cp(7) 收文日(民國年)
'cp(8) 總收文號
'911118 nick 新增原來申請人
'cp(9)  原來申請人
Public Function ChkSameCase(ByRef cp() As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
 Dim strID As String, strName As String
 Dim strTmp As String
 Dim i As Integer
On Error GoTo Err
   ChkSameCase = True
   StrQty = "SELECT COUNT(*) FROM CASEPROGRESS WHERE CP01='" & cp(1) & "' AND " & _
      "CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND " & _
      "CP05='" & TransDate(cp(7), 2) & "' AND CP10='" & cp(5) & "'"
   RsTemp.Open StrQty, cnnConnection
   
   If RsTemp.Fields(0) > 0 Then
      MsgBox "該案件當日相同案件性質有一筆以上，請自行確認 !" & vbCrLf & _
         "本所案號 : " & cp(1) & cp(2) & cp(3) & cp(4) & vbCrLf & _
         "案件性質 : " & cp(6) & vbCrLf & _
         "收文日 : " & cp(7) & vbCrLf, vbInformation
   End If
   RsTemp.Close
   
   StrQty = "select sk02 from systemkind where sk01='" & cp(1) & "'"
   RsTemp.Open StrQty, cnnConnection
   i = RsTemp.Fields(0)
   RsTemp.Close
   
   Select Case i
      Case 1
         StrQty = "SELECT PA26 FROM PATENT WHERE PA01='" & cp(1) & "' AND " & _
            "PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'"
      Case 2
         StrQty = "SELECT TM23 FROM TRADEMARK WHERE TM01='" & cp(1) & "' AND " & _
            "TM02='" & cp(2) & "' AND TM03='" & cp(3) & "' AND TM04='" & cp(4) & "'"
      Case 3
         StrQty = "SELECT LC11 FROM LAWCASE WHERE LC01='" & cp(1) & "' AND " & _
            "LC02='" & cp(2) & "' AND LC03='" & cp(3) & "' AND LC04='" & cp(4) & "'"
      Case 4
         StrQty = "SELECT HC05 FROM HIRECASE WHERE HC01='" & cp(1) & "' AND " & _
            "HC02='" & cp(2) & "' AND HC03='" & cp(3) & "' AND HC04='" & cp(4) & "'"
      Case Else
         StrQty = "SELECT SP08 FROM SERVICEPRACTICE WHERE SP01='" & cp(1) & "' AND " & _
            "SP02='" & cp(2) & "' AND SP03='" & cp(3) & "' AND SP04='" & cp(4) & "'"
   End Select
      
   strID = ""
   RsTemp.Open StrQty, cnnConnection
   If Not RsTemp.EOF And Not RsTemp.BOF Then
      If Not IsNull(RsTemp.Fields(0)) Then
         strID = RsTemp.Fields(0)
      End If
      RsTemp.Close
   Else
      RsTemp.Close
      
      StrQty = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP09='" & cp(8) & "'"
      RsTemp.Open StrQty, cnnConnection
      
      Select Case i
         Case 1
            StrQty = "SELECT PA26 FROM PATENT WHERE PA01='" & RsTemp.Fields(0) & "' AND " & _
               "PA02='" & RsTemp.Fields(1) & "' AND PA03='" & RsTemp.Fields(2) & "' AND PA04='" & RsTemp.Fields(3) & "'"
         Case 2
            StrQty = "SELECT TM23 FROM TRADEMARK WHERE TM01='" & RsTemp.Fields(0) & "' AND " & _
               "TM02='" & RsTemp.Fields(1) & "' AND TM03='" & RsTemp.Fields(2) & "' AND TM04='" & RsTemp.Fields(3) & "'"
         Case 3
            StrQty = "SELECT LC11 FROM LAWCASE WHERE LC01='" & RsTemp.Fields(0) & "' AND " & _
               "LC02='" & RsTemp.Fields(1) & "' AND LC03='" & RsTemp.Fields(2) & "' AND LC04='" & RsTemp.Fields(3) & "'"
         Case 4
            StrQty = "SELECT HC05 FROM HIRECASE WHERE HC01='" & RsTemp.Fields(0) & "' AND " & _
               "HC02='" & RsTemp.Fields(1) & "' AND HC03='" & RsTemp.Fields(2) & "' AND HC04='" & RsTemp.Fields(3) & "'"
         Case Else
            StrQty = "SELECT SP08 FROM SERVICEPRACTICE WHERE SP01='" & RsTemp.Fields(0) & "' AND " & _
               "SP02='" & RsTemp.Fields(1) & "' AND SP03='" & RsTemp.Fields(2) & "' AND SP04='" & RsTemp.Fields(3) & "'"
      End Select
      
      RsTemp.Close
      
      strID = ""
      RsTemp.Open StrQty, cnnConnection
      If Not RsTemp.EOF And Not RsTemp.BOF Then
         If Not IsNull(RsTemp.Fields(0)) Then
            strID = RsTemp.Fields(0)
         End If
         RsTemp.Close
      End If
   End If
   
   Dim strTmp1 As String
   If GetCusCAJnam(strID, strName, strTmp, strTmp1) Then
      
   End If
   
   strTmp1 = ""
   StrQty = "SELECT CP60 FROM CASEPROGRESS WHERE CP09='" & cp(8) & "'"
   RsTemp.Open StrQty, cnnConnection
   If Not RsTemp.EOF And Not RsTemp.BOF Then
      If Not IsNull(RsTemp.Fields(0)) Then
         
         strTmp1 = RsTemp.Fields(0)
         RsTemp.Close
         
         '911118 nick 加入  若同一收據，但申請人不同，有兩筆以上時，不能修改
         'StrQty = "SELECT COUNT(DISTINCT A0J02) FROM ACC0J0 WHERE A0J13='" & strTmp1 & "'"
         StrQty = "SELECT COUNT(DISTINCT A0J02) FROM ACC0J0 WHERE A0J13='" & strTmp1 & "' and a0j11<>'" & cp(9) & "' "
         RsTemp.Open StrQty, cnnConnection
         If RsTemp.Fields(0) > 1 Then
            MsgBox "同一收據有其他案號收文資料，不可轉本所案號 !", vbInformation
            ChkSameCase = False
         Else
      
            StrQty = "UPDATE ACC0J0 SET A0J02='" & cp(1) & cp(2) & cp(3) & cp(4) & "',A0J11=" & CNULL(strID) & " WHERE A0J13=" & CNULL(strTmp1)
            cnnConnection.Execute StrQty
            '92.3.8 MODIFY BY SONIA 不改收據抬頭
            'StrQty = "UPDATE ACC0K0 SET A0K03=" & CNULL(strID) & ",A0K04=" & CNULL(strName) & " WHERE A0K01=" & CNULL(strTmp1)
            StrQty = "UPDATE ACC0K0 SET A0K03=" & CNULL(strID) & " WHERE A0K01=" & CNULL(strTmp1)
            '92.3.8 END
            cnnConnection.Execute StrQty
         End If
         RsTemp.Close
      End If
   End If
   
   Exit Function
Err:
   ChkSameCase = False
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'cp(1),cp(2),cp(3),cp(4) 本所案號
'cp(5) 收據號碼
'cp(6) 客戶編號
'cp(7) 收據抬頭
'911118 nick 新增原來申請人
'cp(8) 原來申請人
Public Function UpdAcc0k0(ByRef cp() As String, Optional bolUpd As Boolean = False, _
   Optional frm040104_9 As Boolean = False) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   UpdAcc0k0 = False
   cp(6) = ChangeCustomerL(cp(6))
   '911118 nick 加入  若同一收據，但申請人不同，有兩筆以上時，不能修改
   'StrQty = "SELECT COUNT(DISTINCT A0J02) FROM ACC0J0 WHERE A0J13='" & cp(5) & "'"
   StrQty = "SELECT COUNT(DISTINCT A0J02) FROM ACC0J0 WHERE A0J13='" & cp(5) & "' and a0j11<>'" & cp(8) & "' "
   RsTemp.Open StrQty, cnnConnection
   If RsTemp.Fields(0) > 1 Then
      MsgBox "同一收據有其他案號收文資料，不可修改申請人 !", vbInformation
   Else
      If bolUpd Then
         '92.3.9 cancel by sonia 都不更新收據抬頭
         'If Not frm040104_9 Then
         '   StrQty = "UPDATE ACC0K0 SET A0K03=" & CNULL(cp(6)) & ",A0K04=" & CNULL(cp(7)) & _
         '      " WHERE A0K01=" & CNULL(cp(5))
         'Else
            'frm040104_9 不更新收據抬頭
            StrQty = "UPDATE ACC0K0 SET A0K03=" & CNULL(cp(6)) & _
               " WHERE A0K01=" & CNULL(cp(5))
         'End If
         '92.3.9 end
         cnnConnection.Execute StrQty
         
         StrQty = "UPDATE ACC0J0 SET A0J11=" & CNULL(cp(6)) & " WHERE A0J13='" & cp(5) & "'"
         cnnConnection.Execute StrQty
      End If
      UpdAcc0k0 = True
   End If
   RsTemp.Close
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function CNULL(ByRef strNULL As String) As String
   If strNULL = "" Then
      CNULL = "NULL"
   Else
      CNULL = "'" + strNULL + "'"
   End If
End Function

Public Function CheckIsExistCaseNum(ByRef intKind As Integer, ByRef strlc As String, ByRef msg As String) As Boolean
 Dim rsRecordset As New ADODB.Recordset
On Error GoTo Err
   Select Case intKind
   Case 1
      StrQty = "select lc11 from lawcase where " & ChgLawcase(strlc)
   Case 2
      StrQty = "select hc05  from hirecase where " & ChgHirecase(strlc)
   End Select
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open StrQty, cnnConnection
   If rsRecordset.RecordCount > 0 Then
      msg = "此本所案號" + strlc + "已存在"
      CheckIsExistCaseNum = False
   Else
      msg = "此本所案號" + strlc + "不存在"
      CheckIsExistCaseNum = True
   End If
   rsRecordset.Close
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function GetStaffArea(ByRef strSDNum As String, ByRef strSDName As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   strSDName = ""
   StrQty = "select st03 from staff  where st01='" + strSDNum + "'"
   GetStaffArea = False
   RsTemp.Open StrQty, cnnConnection
   Do While Not RsTemp.EOF
      If Not IsNull(RsTemp.Fields(0).Value) Then strSDName = RsTemp.Fields(0).Value
      GetStaffArea = True
      Exit Do
   Loop
   RsTemp.Close
   If GetStaffArea = False Then MsgBox "員工代號不存在"
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function GetReasonOfRelief(ByRef strNum As String, ByRef strResult As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   strResult = ""
   StrQty = "select ror02 from reasonofrelief where ror01='" + strNum + "'"
   GetReasonOfRelief = False
   RsTemp.Open StrQty, cnnConnection
   Do While Not RsTemp.EOF
      If Not IsNull(RsTemp.Fields(0).Value) Then strResult = RsTemp.Fields(0).Value
      GetReasonOfRelief = True
      Exit Do
   Loop
   RsTemp.Close
   If GetReasonOfRelief = False Then MsgBox "解除期限原因代碼錯誤.", vbCritical
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function ChkMRec(mr02 As String, strlc As String, mr16 As String, mr17 As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   StrQty = "select mr16,mr17 from mailrec where " & ChgMailRec(strlc) & " AND mr02=" + CNULL(mr02)
   ChkMRec = False
   mr16 = ""
   mr17 = ""
   RsTemp.Open StrQty, cnnConnection
   Do While Not RsTemp.EOF
      If IsNull(RsTemp.Fields(0)) = False Then mr16 = RsTemp.Fields(0)
      If IsNull(RsTemp.Fields(1)) = False Then mr17 = RsTemp.Fields(1)
      ChkMRec = True
      Exit Do
   Loop
   RsTemp.Close
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function Update013(ByRef CP09 As String, ByRef cp01 As String, ByRef cp02 As String, ByRef cp03 As String, ByRef cp04 As String, _
   ByRef cp12 As String, ByRef cp13 As String, ByRef cp43 As String, ByRef cp05 As String, ByRef cp71 As String, ByRef cp35 As String, ByRef cp30 As String, _
   ByRef cdp06 As String, ByRef cdp03 As String, ByRef cdp04 As String, ByRef cdp02 As String, ByRef cdp08 As String, ByRef cdp09 As String, _
   ByRef cp64 As String, ByRef cid As String, ByRef Cd As String, ByRef Ctime As String) As Boolean
 Dim cdp05 As String, strS(1 To 2) As String
   strS(1) = "INSERT INTO CASEPROGRESS (cp09,cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp48,cp12,cp13,cp32,cp43,cp26," + _
      " cp10,cp71,cp35) values(" + CNULL(CP09) + "," + CNULL(cp01) + "," + CNULL(cp02) + "," + CNULL(cp03) + "," + CNULL(cp04) + _
      "," + CNULL(cp05) + "," + CNULL(cdp03) + "," + CNULL(cdp03) + "," + CNULL(cdp03) + "," + CNULL(cp12) + "," + CNULL(cp13) + _
      ",'N'," + CNULL(cp43) + ",'N'," + CNULL(通知開庭) + "," + CNULL(cp71) + "," + CNULL(cp35) + ")"
   strS(2) = "insert into courtyardperiod(cdp01,cdp02,cdp03,cdp04,cdp05,cdp06,cdp07,cdp08,cdp09,cdp10,cdp11,cdp12) values " + _
      " (" + CNULL(CP09) + "," + CNULL(cdp02) + "," + CNULL(cdp03) + "," + CNULL(cdp04) + "," + CNULL(cp71) + _
      "," + CNULL(cdp06) + "," + CNULL(cp64) + "," + CNULL(cdp08) + "," + CNULL(cdp09) + "," + CNULL(cid) + "," + CNULL(Cd) + "," + CNULL(Ctime) + ")"
   If ExecSQL(2, strS) = True Then Update013 = True
End Function

Public Function GetCaseFee(ByRef strKind As String, ByRef strRelNation As String, ByRef strproperty As String, ByRef dblDay As Double) As Boolean
 Dim RsTemp As New ADODB.Recordset
   dblDay = 0
   StrQty = "select cf04 from casefee where cf01='" + strKind + "' and cf02='" + strRelNation + "' and cf03='" + strproperty + "'"
   GetCaseFee = False
   RsTemp.Open StrQty, cnnConnection
   Do While Not RsTemp.EOF
      If Not IsNull(RsTemp.Fields(0)) Then dblDay = RsTemp.Fields(0)
      GetCaseFee = True
      Exit Do
   Loop
   RsTemp.Close
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'傳進西元年，傳出西元年
Public Function GetCaseFeeDelay(ByVal CF01 As String, ByVal CF02 As String, ByVal CF03 As String, ByRef CFOther() As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   GetCaseFeeDelay = True
   CFOther(1) = CFOther(0) '法定
   CFOther(2) = CFOther(0) '本所
   StrQty = "SELECT CF22,CF25,CF27 FROM CASEFEE WHERE CF01='" & CF01 & "' AND CF02='" & CF02 & "' AND CF03='" & CF03 & "'"
   RsTemp.Open StrQty, cnnConnection
   Do While Not RsTemp.EOF
      If IsNull(RsTemp.Fields(0)) Or RsTemp.Fields(0) = 0 Then
         If Not IsNull(RsTemp.Fields(1)) Then
         '月
            CFOther(1) = CompDate(1, RsTemp.Fields(1), CFOther(0))
            If RsTemp.Fields("CF27") = "1" Then CFOther(1) = CompDate(2, -1, CFOther(1))
            
            Select Case CF01
               Case "CFT"
                  CFOther(2) = CompDate(1, -1, CFOther(1))
               Case "CFP"
                  CFOther(2) = CompDate(2, -14, CFOther(1))
               Case "T"
                  If CF02 = "238" Then
                     CFOther(2) = CompDate(1, -1, CFOther(1))
                  Else
                     If RsTemp.Fields(1) >= 2 Then
                        CFOther(2) = CompDate(2, -4, CFOther(1))
                     Else
                        CFOther(2) = CompDate(2, -2, CFOther(1))
                     End If
                  End If
               Case "P"
                  If CF02 = "000" Then
                     If RsTemp.Fields(1) >= 2 Then
                        CFOther(2) = CompDate(2, -4, CFOther(1))
                     Else
                        CFOther(2) = CompDate(2, -2, CFOther(1))
                     End If
                  Else
                     CFOther(2) = CompDate(2, -10, CFOther(1))
                  End If
               Case Else
                  If RsTemp.Fields(1) >= 2 Then
                     CFOther(2) = CompDate(2, -4, CFOther(1))
                  Else
                     CFOther(2) = CompDate(2, -2, CFOther(1))
                  End If
            End Select
            
         End If
      Else
         If Not IsNull(RsTemp.Fields(0)) Then
         '日
            CFOther(1) = CompDate(2, RsTemp.Fields(0), CFOther(0))
            If RsTemp.Fields("CF27") = "1" Then CFOther(1) = CompDate(2, -1, CFOther(1))
            
            Select Case CF01
               Case "CFT"
                  CFOther(2) = CompDate(1, -1, CFOther(1))
               Case "CFP"
                  CFOther(2) = CompDate(2, -14, CFOther(1))
               Case "T"
                  If CF02 = "238" Then
                     CFOther(2) = CompDate(1, -1, CFOther(1))
                  Else
                     If RsTemp.Fields(0) >= 60 Then
                        CFOther(2) = CompDate(2, -4, CFOther(1))
                     Else
                        CFOther(2) = CompDate(2, -2, CFOther(1))
                     End If
                  End If
               Case "P"
                  If CF02 = "000" Then
                     If RsTemp.Fields(0) >= 60 Then
                        CFOther(2) = CompDate(2, -4, CFOther(1))
                     Else
                        CFOther(2) = CompDate(2, -2, CFOther(1))
                     End If
                  Else
                     CFOther(2) = CompDate(2, -10, CFOther(1))
                  End If
               Case Else
                  If RsTemp.Fields(0) >= 60 Then
                     CFOther(2) = CompDate(2, -4, CFOther(1))
                  Else
                     CFOther(2) = CompDate(2, -2, CFOther(1))
                  End If
            End Select
         End If
      End If
      GetCaseFeeDelay = True
      Exit Do
   Loop
   RsTemp.Close
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function GetNextProgress(ByRef CP09 As String, ByRef cp As String, ByRef strproperty As String, ByRef strPropertymemo As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   strproperty = ""
   strPropertymemo = ""
   StrQty = "select np07,np15 from nextprogress where np01='" + CP09 + "' and " & ChgNextProgress(cp)
   RsTemp.Open StrQty, cnnConnection
   Do While Not RsTemp.EOF
      If Not IsNull(RsTemp.Fields(0)) Then strproperty = RsTemp.Fields(0)
      If Not IsNull(RsTemp.Fields(1)) Then strPropertymemo = RsTemp.Fields(1)
      GetNextProgress = True
   Loop
   RsTemp.Close
   Exit Function
Err:
   GetNextProgress = False
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function GetMax() As Long
 Dim RsTemp As New ADODB.Recordset
   StrQty = "SELECT MAX(NP22) FROM NEXTPROGRESS"
   RsTemp.Open StrQty, cnnConnection
   GetMax = 1
   Do While Not RsTemp.EOF
      GetMax = RsTemp.Fields(0).Value + 1
      Exit Do
   Loop
   RsTemp.Close
End Function

'動態集之recordset Input 0:有訊息 1:無訊息 Output 0:無資料 1:有資料 2:錯誤
Public Function ReadRstMsg(ByRef i As Integer, ByRef strSQL As String, Optional Server As Boolean = False, Optional ErrShow As Boolean = False) As ADODB.Recordset
 
 Dim Rc As New ADODB.Recordset
On Error GoTo Err
   If Server = False Then
      Rc.CursorLocation = adUseClient
      
      'Add By Cheng 2002/01/02
      'nick建議會加速查詢速度
'      Rc.CursorType = adOpenDynamic
      Rc.CursorType = adOpenStatic
      Rc.LockType = adLockReadOnly
      Rc.Open strSQL, cnnConnection
      If Rc.RecordCount > 0 Then
         Set ReadRstMsg = Rc
         i = 1
      Else
         If i = 0 Then MsgBox "資料庫無資料 !", vbInformation
         Set ReadRstMsg = Rc
         i = 0
      End If
   Else
      Rc.CursorLocation = adUseServer
      
      'Add By Cheng 2002/01/02
      'nick建議會加速查詢速度
'      Rc.CursorType = adOpenDynamic
      Rc.CursorType = adOpenStatic
      Rc.LockType = adLockReadOnly
      
      Rc.Open strSQL, cnnConnection
      If Rc.EOF And Rc.BOF Then
         If i = 0 Then MsgBox "資料庫無資料 !", vbInformation
         Set ReadRstMsg = Rc
         i = 0
      Else
         Set ReadRstMsg = Rc
         i = 1
      End If
   End If
   Exit Function
Err:
   i = 2
   If ErrShow = False Then MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function ExecSQL(ByVal j As Integer, ByRef Qty() As String) As Boolean
 Dim i As Integer
'Add By Cheng 2002/11/14
Dim BolTransOk As Boolean
BolTransOk = True
 
On Error GoTo Err
   If j > 0 Then
      cnnConnection.BeginTrans
      For i = 1 To j
         If Qty(i) <> "" Then cnnConnection.Execute Qty(i)
      Next
        'Modify By Cheng 2002/11/14
        If BolTransOk Then
            cnnConnection.CommitTrans
        End If
   End If
   ExecSQL = True
   Exit Function
Err:
    'Add By Cheng 2002/11/14
    If Err.Number = -2147168237 Then
       BolTransOk = False
       Resume Next
    End If

   cnnConnection.RollbackTrans
   ExecSQL = False
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function ChangeCustomerS(ByRef strTemp As String) As String
   If strTemp <> "" Then
      ChangeCustomerS = IIf(Right(strTemp, 3) = "000", Mid(strTemp, 1, 6), IIf(Right(strTemp, 1) = "0", Mid(strTemp, 1, 8), strTemp))
   Else
      ChangeCustomerS = ""
   End If
End Function

'傳回長的客戶代號
Public Function ChangeCustomerL(ByRef strTemp As String) As String
   If strTemp <> "" Then ChangeCustomerL = strTemp + String(9 - Len(strTemp), "0")
End Function

'取得代理人名稱
Public Function LawGetName(ByRef strNum As String, ByRef strName As String, Optional ByVal intSituE As Integer = 0) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   strName = ""
   strNum = ChangeCustomerL(strNum)
   If Left(strNum, 1) = 代理人編號 Then
      StrQty = "SELECT NVL(FA05,NVL(FA04,FA06)) FROM FAGENT WHERE " & ChgFagent(strNum)
   ElseIf Left(strNum, 1) = 客戶編號 Then
      If intSituE = 1 Then
         StrQty = "SELECT NVL(CU05,NVL(CU04,CU06)) FROM CUSTOMER WHERE " & ChgCustomer(strNum)
      Else
         StrQty = "SELECT NVL(CU04,NVL(CU05,CU06)) FROM CUSTOMER WHERE " & ChgCustomer(strNum)
      End If
   Else
      MsgBox "編號第一碼錯誤，請重新輸入 !", vbCritical
      Exit Function
   End If
   RsTemp.Open StrQty, cnnConnection
   LawGetName = False
   Do While Not RsTemp.EOF
      If Not IsNull(RsTemp.Fields(0).Value) Then strName = RsTemp.Fields(0).Value
      LawGetName = True
      Exit Do
   Loop
   RsTemp.Close
   strNum = ChangeCustomerS(strNum)
   If LawGetName = False Then MsgBox "編號錯誤，請重新輸入 !", vbCritical
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function GetInvLetter(ByVal strIn As String, ByRef strName() As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
 Dim i As Integer
On Error GoTo Err
   GetInvLetter = False
   For i = 0 To 4
      strName(i) = ""
   Next
   strIn = strIn & String(10 - Len(strIn), "0")
   StrQty = "SELECT IN03,IN04,IN05,IN07,NA03 FROM INVENTOR,NATION WHERE in01='" & Left(strIn, 8) & "' AND IN02='" & Right(strIn, 2) & "' AND IN11=NA01"
   RsTemp.Open StrQty, cnnConnection
   Do While Not RsTemp.EOF
      For i = 0 To 4
         If Not IsNull(RsTemp.Fields(i)) Then strName(i) = RsTemp.Fields(i)
      Next
      GetInvLetter = True
      Exit Do
   Loop
   RsTemp.Close
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function GetCusLetter(ByVal strNum As String, ByRef strNA03 As String, ByRef strCU11 As String, _
   ByRef strCU23 As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   GetCusLetter = False
   strNA03 = ""
   strCU11 = ""
   strCU23 = ""
   StrQty = "SELECT NA03,CU11,CU23 FROM CUSTOMER,NATION WHERE " & ChgCustomer(strNum) & " AND CU10=NA01"
   RsTemp.Open StrQty, cnnConnection
   With RsTemp
      Do While Not .EOF
         If Not IsNull(.Fields(0).Value) Then strNA03 = .Fields(0).Value
         If Not IsNull(.Fields(1).Value) Then strCU11 = .Fields(1).Value
         If Not IsNull(.Fields(2).Value) Then strCU23 = .Fields(2).Value
         GetCusLetter = True
         Exit Do
      Loop
      .Close
   End With
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Function GetCusCAJnam(ByVal strNum As String, ByRef StrCnam As String, _
   ByRef StrAnam As String, ByRef StrJnam As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   StrCnam = ""
   StrAnam = ""
   StrJnam = ""
   If Left(strNum, 1) = 代理人編號 Then
      StrQty = "SELECT FA04,FA05||' '||FA63||' '||FA64||' '||FA65,FA06 FROM FAGENT WHERE " & ChgFagent(strNum)
   ElseIf Left(strNum, 1) = 客戶編號 Then
      StrQty = "SELECT CU04,CU05||' '||CU88||' '||CU89||' '||CU90,CU06 FROM CUSTOMER WHERE " & ChgCustomer(strNum)
   Else
      MsgBox "編號第一碼錯誤，請重新輸入 !", vbCritical
      Exit Function
   End If
   RsTemp.Open StrQty, cnnConnection
   GetCusCAJnam = False
   With RsTemp
      Do While Not .EOF
         If Not IsNull(.Fields(0).Value) Then StrCnam = .Fields(0).Value
         If Not IsNull(.Fields(1).Value) Then StrAnam = .Fields(1).Value
         If Not IsNull(.Fields(2).Value) Then StrJnam = .Fields(2).Value
         GetCusCAJnam = True
         Exit Do
      Loop
      .Close
   End With
   If GetCusCAJnam = False Then MsgBox "編號錯誤，請重新輸入 !", vbCritical
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'讀取客戶發明人檔
Public Function GetInventor(ByVal strIn As String, ByRef strName() As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
 Dim i As Integer
On Error GoTo Err
   GetInventor = False
   For i = 0 To 4
      strName(i + 1) = ""
   Next
   ' 90.08.22 modify by louis
   If Len(strIn) < 10 Then
      strIn = strIn & String(10 - Len(strIn), "0")
   End If
   StrQty = "SELECT IN04,IN05,IN06,IN07,IN08 FROM INVENTOR WHERE in01='" & Left(strIn, 8) & "' AND IN02='" & Right(strIn, 2) & "'"
   RsTemp.Open StrQty, cnnConnection
   Do While Not RsTemp.EOF
      For i = 0 To 4
         If Not IsNull(RsTemp.Fields(i)) Then strName(i + 1) = RsTemp.Fields(i)
      Next
      GetInventor = True
      Exit Do
   Loop
   RsTemp.Close
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'取得下次繳費日,利用array傳值,array(1-4)為本所案號,
Public Function GetNextPayDate(ByRef pa() As String, ByRef strName As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
On Error GoTo Err
   GetNextPayDate = False
   strName = ""
   StrQty = "SELECT MAX(NP09) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
      " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "') And NP08 IS NOT NULL AND NP06 IS NULL"
   RsTemp.Open StrQty, cnnConnection
   Do While Not RsTemp.EOF
      If Not IsNull(RsTemp.Fields(0)) Then strName = RsTemp.Fields(0)
      GetNextPayDate = True
      Exit Do
   Loop
   RsTemp.Close
   Exit Function
Err:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

Public Property Set Connection(ByVal vNewValue As ADODB.Connection)
   Set cnnConnection = vNewValue
End Property
