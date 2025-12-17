Attribute VB_Name = "basAutoBatch"
'Memo by Lydia 2022/05/30 (已檢查)整理frm880005改用寄信模組
'Memo By Morgan 2012/12/12 智權人員欄已修改
'Memo by Morgan2010/8/10 日期欄已修改
Option Explicit

Public Conn As New ADODB.Connection
Public rs As New ADODB.Recordset
'Modified by Morgan 2017/8/14 O12大小寫有分
'Public Const CnStr As String = "Provider=MSDAORA.1;Password=pgmpwd;User ID=pgmid;Data Source=m51con;Persist Security Info=True"
'Removed by Morgan 2022/3/23 不再使用
'Public Const CnStr As String = "Provider=MSDAORA.1;Password=PGMPWD;User ID=PGMID;Data Source=m51con;Persist Security Info=True"

'Global strSQL As String
'Public strSrvDate(1 To 2) As String '1 西元 '2 民國

'Public sBodyText As String
'Public Const cBoundaryA As String = "Boundary_Taie_A"
'Public Const cBoundaryB As String = "Boundary_Taie_B"
Public Const cAST As String = "*"
'Public Const cDASH2 As String = "--"
Public Const cDASH As String = "-"
'Public Const cDOT As String = "."
'Public Const cSEMIC As String = ";"
Public Const LEN1024 As Long = 1024
Public Const cSRC = "SRC="
Public Const R_BRACKET = ">"
Public Const cSPACE = " "
Public Const cEQUAL = "="
'Public AdoRecordSet3 As New ADODB.Recordset 'Add By Sindy 2009/05/26
'Global cnnConnection As ADODB.Connection 'Add By Sindy 2009/05/26

'Copy By Sindy 98/04/02
'add by nickc 2006/05/24
'Type BITMAP '14 bytes
'    bmType As Long          '=0
'    bmWidth As Long         '點陣圖寬度(單位：像素)
'    bmHeight As Long        '點陣圖高度(單位：像素)
'    bmWidthBytes As Long    'the number of bytes in each scan line (word aligned)
'    bmPlanes As Integer     '=1
'    bmBitsPixel As Integer  '每個像素以幾位元儲存
'    bmBits As Long          '指標。指向點陣資料陣列
'End Type
Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
'98/04/02 End

Public Const KEY_QUERY_VALUE = &H1
Public Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const WM_LBUTTONUP = &H202

Public Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueExString Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, ByVal lpData As String, lpcbData As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (xDest As Any, xSource As Any, ByVal nbytes As Long)
Public Declare Function CoCreateGuid Lib "ole32.dll" (pGuid As GUID) As Long
Public Declare Function StringFromGUID2 Lib "ole32.dll" (pGuid As GUID, ByVal PointerToString As Long, ByVal MaxLength As Long) As Long
Public BooPlain As Boolean
'Public APics() As String
'Public pub_HostName As String
Private Declare Function gethostname Lib "WSOCK32.DLL" (ByVal Name As String, ByVal namelen As Integer) As Integer
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Private Const MAX_PATH = 260


''讀取電腦名稱
'Public Function PUB_ReadHostName() As String
'   Dim stHostName As String
'   Dim dwLength As Integer
'   dwLength = 256
'   stHostName = String(dwLength, Chr(0))
'   gethostname stHostName, Len(stHostName)
'   PUB_ReadHostName = Replace(stHostName, Chr(0), "")
'End Function

'Public Function ServerDate() As Long
'   Dim adoSysDate As New ADODB.Recordset
'   adoSysDate.CursorLocation = adUseClient
'   adoSysDate.Open "select to_char(sysdate, 'YYYYMMDD') from dual", Conn, adOpenStatic, adLockReadOnly
'   If adoSysDate.RecordCount <> 0 Then
'      ServerDate = Val(adoSysDate.Fields(0).Value)
'   End If
'   adoSysDate.Close
'   Set adoSysDate = Nothing
'End Function

'Public Function PUB_GetDbTerminal() As String
'   Dim strSql As String
'   Dim AdoRecordSet3 As New ADODB.Recordset
'
'On Error GoTo ErrHnd
'   strSql = "select TERMINAL FROM V$SESSION where SID=1"
'   With AdoRecordSet3
'      .CursorLocation = adUseClient
'      .Open strSql, Conn, adOpenForwardOnly, adLockReadOnly
'      If .RecordCount > 0 Then
'         PUB_GetDbTerminal = "(" & .Fields(0) & ")"
'      End If
'   End With
'ErrHnd:
'   If Err.Number <> 0 Then
'      MsgBox Err.Description, vbCritical
'   End If
'   Set AdoRecordSet3 = Nothing
'End Function

''關閉連線
'Public Sub CheckOC()            'NICK
'If adoRecordset.State <> 0 Then
'   adoRecordset.Close
'End If
'End Sub

''取得員工代收郵件人員
'Public Function PUB_GetST14(ByVal p_User As String) As String
'   Dim strSql As String
'   Dim AdoRecordSet3 As New ADODB.Recordset
'   strSql = "select st14 from staff where st01='" & p_User & "'"
'
'On Error GoTo ErrHnd
'
'   With AdoRecordSet3
'      .CursorLocation = adUseClient
'      .Open strSql, Conn, adOpenForwardOnly, adLockReadOnly
'      If .RecordCount > 0 Then
'         PUB_GetST14 = "" & .Fields(0)
'      End If
'   End With
'
'ErrHnd:
'   If Err.Number <> 0 Then
'      MsgBox Err.Description, vbCritical
'   End If
'   Set AdoRecordSet3 = Nothing
'End Function

'Public Function SQLDate(ByVal strTemp As String, Optional bolDate As Boolean = True) As String
'   If bolDate Then
'      SQLDate = "DECODE(" & strTemp & ",'','',SUBSTR(" & strTemp & ",1,4)-1911||'/'||SUBSTR(" & strTemp & ",5,2)||'/'||SUBSTR(" & strTemp & ",7,2))"
'   Else
'      SQLDate = "DECODE(" & strTemp & ",'','',SUBSTR(" & strTemp & ",1,4)||'/'||SUBSTR(" & strTemp & ",5,2)||'/'||SUBSTR(" & strTemp & ",7,2))"
'   End If
'End Function

''將不同型態轉為字串
'Public Function CheckStr(ByRef Strindex As Variant) As String
'If IsNull(Strindex) Then
'    CheckStr = ""
'Else
'    Select Case VarType(Strindex)
'    Case 2, 3, 4, 5
'         CheckStr = Trim(str(Strindex)) 'LTrim()
'    Case Else
'         CheckStr = Trim(Strindex)
'    End Select
'End If
'End Function

''取的新字串長度
'Public Function StrToStr(ByRef Strindex As String, ByRef StrIndex2 As Single) As String
'StrToStr = StrConv(MidB(StrConv(Strindex, vbFromUnicode), 1, StrIndex2 * 2), vbUnicode)
'End Function

''轉換民國日期至西元有加/之格式
'Public Function ChangeTStringToWDateString(ByRef strTString As String) As String
'Dim intLen As Integer
'
'If strTString = "" Then
'   ChangeTStringToWDateString = ""
'   Exit Function
'End If
'intLen = Len(strTString)
'If intLen = 6 Then
'   ChangeTStringToWDateString = Format(Val(Left(strTString, 2)) + 1911) + "/" + Mid(strTString, 3, 2) + "/" + Right(strTString, 2)
'ElseIf intLen = 7 Then
'    'Modify By Cheng 2004/02/12
''   ChangeTStringToWDateString = Format(Val(Left(strTString, 3)) + 1911) + "/" + Mid(strTString, 3, 2) + "/" + Right(strTString, 2)
'   ChangeTStringToWDateString = Format(Val(Left(strTString, 3)) + 1911) + "/" + Mid(strTString, 4, 2) + "/" + Right(strTString, 2)
'    'End
'End If
'End Function

''轉換西元有加/之格式至西元日期
'Public Function ChangeWDateStringToWString(ByRef strWDateString As String) As String
'Dim intIndex1 As Integer, intIndex2 As Integer
'
'If strWDateString = "" Then
'   ChangeWDateStringToWString = ""
'   Exit Function
'End If
'intIndex1 = InStr(strWDateString, "/")
'ChangeWDateStringToWString = Format(Left(strWDateString, intIndex1 - 1))
'intIndex2 = InStr(intIndex1 + 1, strWDateString, "/")
'ChangeWDateStringToWString = ChangeWDateStringToWString + Format(Mid(strWDateString, intIndex1 + 1, intIndex2 - intIndex1 - 1), "00") + Format(Mid(strWDateString, intIndex2 + 1), "00")
'End Function

''轉換西元日期至西元有加/之格式
'Public Function ChangeWStringToWDateString(ByRef strWString As String) As String
'
'If strWString = "" Then
'   ChangeWStringToWDateString = ""
'   Exit Function
'End If
'ChangeWStringToWDateString = Format(Left(strWString, 4)) + "/" + Mid(strWString, 5, 2) + "/" + Right(strWString, 2)
'End Function

''轉換西元之格式到民國
'Public Function ChangeWStringToTDateString(ByRef strTString As String) As String
'If strTString <> "" Then
'   ChangeWStringToTDateString = ChangeTStringToTDateString(ChangeWStringToTString(strTString))
'Else
'   ChangeWStringToTDateString = ""
'End If
'End Function

''轉換民國日期至有加/之格式
'Public Function ChangeTStringToTDateString(ByRef strTString As String) As String
'Dim intLen  As Integer
'
'If strTString = "" Then
'   ChangeTStringToTDateString = ""
'   Exit Function
'End If
'intLen = Len(strTString)
'If intLen = 6 Then
'   ChangeTStringToTDateString = Left(strTString, 2) + "/" + Mid(strTString, 3, 2) + "/" + Right(strTString, 2)
'ElseIf intLen = 7 Then
'   ChangeTStringToTDateString = Left(strTString, 3) + "/" + Mid(strTString, 4, 2) + "/" + Right(strTString, 2)
'End If
'End Function

''轉換西元日期至民國之格式
'Public Function ChangeWStringToTString(ByRef strWString As String) As String
'If strWString <> "" Then
'   ChangeWStringToTString = Format(Val(strWString) - 19110000)
'Else
'   ChangeWStringToTString = ""
'End If
'End Function
''Add by Morgan 2008/6/9
'Public Function ChkWorkDay(ByVal strTemp As String) As Boolean
'
'   If Conn.State = adStateClosed Then
'      Conn.ConnectionString = CnStr
'      Conn.Open
'   End If
'
'   Dim rsTemp1 As New ADODB.Recordset
'   Dim stSQL As String
'
'   ChkWorkDay = False
'   stSQL = "SELECT * FROM WORKDAY WHERE WD01=" & strTemp
'   rsTemp1.CursorLocation = adUseClient
'   rsTemp1.Open stSQL, Conn, adOpenForwardOnly, adLockReadOnly
'   If rsTemp1.RecordCount > 0 Then
'      ChkWorkDay = True
'   End If
'   Set rsTemp1 = Nothing
'
'End Function

''strStart 大 strEnd 小
'Public Function GetWorkDay(ByVal strStart As String, ByVal strEnd As String) As Integer
'Dim RsNick2 As New ADODB.Recordset
'Dim strExc As String
'Set RsNick2 = New ADODB.Recordset
'
'   strExc = "SELECT COUNT(*) FROM WORKDAY WHERE WD01 BETWEEN " & strEnd & " AND " & strStart
'   RsNick2.CursorLocation = adUseClient
'   RsNick2.Open strExc, Conn, adOpenStatic, adLockReadOnly
'   If IsNull(RsNick2.Fields(0).Value) Then
'      GetWorkDay = 0
'   Else
'      If RsNick2.Fields(0) = 0 Then
'         GetWorkDay = 0
'      Else
'            GetWorkDay = RsNick2.Fields(0)
'      End If
'   End If
'End Function

''iSitu 0:+  1:-
'Public Function CompWorkDay(ByVal iAddDay As Integer, strDay As String, Optional iSitu As Integer = 0) As String
' Dim i As Integer
' Dim intI As Integer
' Dim strExc(0) As String
' Dim RsTemp As New ADODB.Recordset
' 'Add by Morgan 2010/11/5
' Dim stDateBoundry As String
'
'On Error GoTo Err
'   Select Case iSitu
'      Case 0
'         'Add by Morgan 2010/11/5
'         '改上下限抓2倍天數,固定3個月有可能不夠
'         'stDateBoundry =  ChangeWDateStringToWString(DateAdd("m", 12, Format(strDay, "####/##/##")))
'         stDateBoundry = CompDate(2, 2 * iAddDay + 30, strDay)
'
'         strExc(0) = "SELECT WD01 FROM WORKDAY WHERE WD01>=" & strDay & _
'            " AND wd01<=" & stDateBoundry & " ORDER BY WD01"
'      Case 1
'         'Add by Morgan 2010/11/5
'         '改上下限抓2倍天數,固定3個月有可能不夠
'         'stDateBoundry =  ChangeWDateStringToWString(DateAdd("m", 12, Format(strDay, "####/##/##")))
'         stDateBoundry = CompDate(2, -1 * 2 * iAddDay - 30, strDay)
'         strExc(0) = "SELECT WD01 FROM WORKDAY WHERE WD01<=" & strDay & " AND wd01>=" & stDateBoundry & " ORDER BY WD01 DESC"
'   End Select
'   intI = 1
'   Set RsTemp = New ADODB.Recordset
'   RsTemp.CursorLocation = adUseClient
'   RsTemp.Open strExc(0), Conn, adOpenStatic, adLockReadOnly
'   'If rsTemp.RecordCount <> 0 Then
'   If Not RsTemp.EOF And Not RsTemp.BOF Then
'        '若有設定工作天數
'        If iAddDay > 0 Then
'            RsTemp.Move iAddDay - 1
'        End If
'      CompWorkDay = RsTemp.Fields(0)
'   Else
'      CompWorkDay = strDay
'   End If
'   Exit Function
'Err:
'   CompWorkDay = ""
'   'MsgBox "無此記錄，請重新輸入 !", vbCritical
'   WLog "CompWorkDay ,無此記錄，請重新輸入 !"
'End Function



''拆字串
'Public Function SystemNumber(ByRef strSystem As String, ByRef Strindex As Integer) As String                'NICK
'Dim Str0001 As String, Str0002 As String, Str0003 As String, Str0004 As String
'Dim Int0001 As Integer, Int0002 As Integer
'If strSystem = "" Or IsNull(strSystem) Then
'Exit Function
'End If
'Int0001 = InStr(1, Trim(strSystem), "-")
'Str0001 = Left(Trim(strSystem), Int0001 - 1)
'Int0002 = InStr(Int0001 + 1, Trim(strSystem), "-")
'Str0002 = Mid(Trim(strSystem), Int0001 + 1, Int0002 - Int0001 - 1)
'Int0001 = InStr(Int0002 + 1, Trim(strSystem), "-")
'Str0003 = Mid(Trim(strSystem), Int0002 + 1, Int0001 - Int0002 - 1)
'Str0004 = Right(Trim(strSystem), Len(Trim(strSystem)) - Int0001)
'
'Select Case Strindex
'Case 1
'      SystemNumber = Str0001
'Case 2
'      SystemNumber = Str0002
'Case 3
'      SystemNumber = Str0003
'Case 4
'      SystemNumber = Str0004
'Case Else
'     MsgBox ("傳入參數錯誤")
'End Select
'End Function

'Public Function GetStaffDepartment(ByVal strStuff As String) As String
'   Dim rsTmp As New ADODB.Recordset
'   Dim strSql As String
'
'   GetStaffDepartment = Empty
'   strSql = "SELECT * FROM Staff " & _
'            "WHERE ST01 = '" & strStuff & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, Conn, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      If IsNull(rsTmp.Fields("ST03")) = False Then
'         GetStaffDepartment = rsTmp.Fields("ST03")
'      End If
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
'End Function

'Public Function GetPrjSalesNM(ByVal Strindex As String) As String          '智權人員,承辦人        NICK
'Dim tmpRSNM As New ADODB.Recordset
'If Strindex <> "" Then
'   strSql = "SELECT ST02 FROM STAFF WHERE ST01='" & Strindex & "' "
'   Set tmpRSNM = New ADODB.Recordset
'   If tmpRSNM.State = 1 Then tmpRSNM.Close
'   tmpRSNM.CursorLocation = adUseClient
'   tmpRSNM.Open strSql, Conn, adOpenStatic, adLockReadOnly
'   If tmpRSNM.RecordCount <> 0 And tmpRSNM.RecordCount > 0 Then
'      If Not IsNull(tmpRSNM.Fields(0)) Then
'          GetPrjSalesNM = tmpRSNM.Fields(0)
'      Else
'          GetPrjSalesNM = ""
'      End If
'   Else
'      GetPrjSalesNM = ""
'   End If
'   Set tmpRSNM = Nothing
'Else
'   GetPrjSalesNM = ""
'End If
'End Function

'Public Function GetCustomerName(ByVal strCustomer As String, Optional ByVal nLanguage As String = 0) As String
'   Dim rsTmp As New ADODB.Recordset
'   Dim strKey As String
'   Dim strSql As String
'
'   ' 檢查取得中文還是英文名稱的範圍
'   If nLanguage < 0 Or nLanguage > 1 Then: nLanguage = 0
'
'   GetCustomerName = Empty
'
'   If Len(strCustomer) < 9 Then: strCustomer = strCustomer & String(9 - Len(strCustomer), "0")
'
'   If Len(strCustomer) > 8 Then
'      strSql = "SELECT * FROM Customer " & _
'               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
'                     "CU02 = '" & Mid(strCustomer, 9, 1) & "'"
'   Else
'      strSql = "SELECT * FROM Customer " & _
'               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
'                     "CU02 = '0' "
'   End If
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, Conn, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      Select Case nLanguage
'         Case 0:
'            If IsNull(rsTmp.Fields("CU04")) = False Then
'               GetCustomerName = rsTmp.Fields("CU04")
'            ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
'               GetCustomerName = rsTmp.Fields("CU05")
'               '92.10.15 add by sonia
'               If IsNull(rsTmp.Fields("CU88")) = False Then
'                  GetCustomerName = GetCustomerName & " " & rsTmp.Fields("CU88")
'               End If
'               If IsNull(rsTmp.Fields("CU89")) = False Then
'                  GetCustomerName = GetCustomerName & " " & rsTmp.Fields("CU89")
'               End If
'               If IsNull(rsTmp.Fields("CU90")) = False Then
'                  GetCustomerName = GetCustomerName & " " & rsTmp.Fields("CU90")
'               End If
'               '92.10.15 end
'            ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
'               GetCustomerName = rsTmp.Fields("CU06")
'            End If
'         Case 1:
'            If IsNull(rsTmp.Fields("CU05")) = False Then
'               GetCustomerName = rsTmp.Fields("CU05")
'               '92.10.15 add by sonia
'               If IsNull(rsTmp.Fields("CU88")) = False Then
'                  GetCustomerName = GetCustomerName & " " & rsTmp.Fields("CU88")
'               End If
'               If IsNull(rsTmp.Fields("CU89")) = False Then
'                  GetCustomerName = GetCustomerName & " " & rsTmp.Fields("CU89")
'               End If
'               If IsNull(rsTmp.Fields("CU90")) = False Then
'                  GetCustomerName = GetCustomerName & " " & rsTmp.Fields("CU90")
'               End If
'               '92.10.15 end
'            ElseIf IsNull(rsTmp.Fields("CU04")) = False Then
'               GetCustomerName = rsTmp.Fields("CU04")
'            ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
'               GetCustomerName = rsTmp.Fields("CU06")
'            End If
'      End Select
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
'End Function

''add by nickc 2006/08/18 檢查收件者，若是已經離職，則改發該部門最小編號者，若是沒有，改發秀玲---秀玲 mail 提的
''Rem by Morgan 2009/2/23 若規則有改時 frm880005 也要同步
'Function ChkMailId(oOldMailID) As String
'ChkMailId = ""
'Dim oStrSQL As String
'Dim MailRS As New ADODB.Recordset
'   '先檢查是否為員工
'   oStrSQL = "select * from staff where st01='" & Replace(UCase(oOldMailID), "@TAIE.COM.TW", "") & "' "
'   Set MailRS = New ADODB.Recordset
'   MailRS.CursorLocation = adUseClient
'   MailRS.Open oStrSQL, Conn, adOpenStatic, adLockReadOnly
'   If MailRS.RecordCount <> 0 Then
'       '檢查在不在職
'       'edit by nickc 2007/09/20 虛建智權人員也發給主管
'       'oStrSQL = "select * from staff where st01='" & Replace(UCase(oOldMailID), "@TAIE.COM.TW", "") & "' and st04='1' "
'       'Modify by Morgan 2010/5/11
'       'oStrSQL = "select * from staff where st01='" & Replace(UCase(oOldMailID), "@TAIE.COM.TW", "") & "' and st04='1' and st01>'63001' "
'       oStrSQL = "select st14 from staff where st01='" & Replace(UCase(oOldMailID), "@TAIE.COM.TW", "") & "' and st04='1' and st01>'63001' "
'       Set MailRS = New ADODB.Recordset
'       MailRS.CursorLocation = adUseClient
'       MailRS.Open oStrSQL, Conn, adOpenStatic, adLockReadOnly
'       If MailRS.RecordCount <> 0 Then
'           'Modify by Morgan 2010/5/11
'           'ChkMailId = oOldMailID
'           If Not IsNull(MailRS.Fields(0)) Then
'               ChkMailId = MailRS.Fields(0)
'           Else
'               ChkMailId = oOldMailID
'           End If
'       Else
'       '已離職
'            '2011/4/13 add by sonia 離職智權人員先發給在職的帶人主管 吳邑君收文CFP-023894已會稿完成
'            oStrSQL = "select s2.st01,s3.st01,s4.st01,s5.st01 from staff s1,staff s2,staff s3,staff s4,staff s5 where s1.st01='" & Replace(UCase(oOldMailID), "@TAIE.COM.TW", "") & "' " & _
'                      "and s1.st52=s2.st01(+) and s2.st04(+)='1' and s1.st53=s3.st01(+) and s3.st04(+)='1' " & _
'                      "and s1.st54=s4.st01(+) and s4.st04(+)='1' and s1.st55=s5.st01(+) and s5.st04(+)='1' "
'            Set MailRS = New ADODB.Recordset
'            MailRS.CursorLocation = adUseClient
'            MailRS.Open oStrSQL, Conn, adOpenStatic, adLockReadOnly
'            If MailRS.RecordCount <> 0 Then
'               If Not IsNull(MailRS(0)) Then
'                  ChkMailId = MailRS(0)
'               ElseIf Not IsNull(MailRS(1)) Then
'                  ChkMailId = MailRS(1)
'               ElseIf Not IsNull(MailRS(2)) Then
'                  ChkMailId = MailRS(2)
'               ElseIf Not IsNull(MailRS(3)) Then
'                  ChkMailId = MailRS(3)
'               End If
'            End If
'            If ChkMailId = "" Then
'            '2011/4/13 end
'
'               'edit by nickc 2007/03/27 改成 st03
'               'oStrSQL = "select min(st01) from staff where st15=(select st15 from staff where st01='" & Replace(UCase(oOldMailID), "@TAIE.COM.TW", "") & "' ) and st04='1' and st01>'63001' "
'               'edit by nickc 2007/08/16 改抓部門檔的主管
'               'oStrSQL = "select min(st01) from staff where st03=(select st03 from staff where st01='" & Replace(UCase(oOldMailID), "@TAIE.COM.TW", "") & "' ) and st04='1' and st01>'63001' "
'               'edit by nickc 2007/09/20 虛建智權人員也發給主管
'               'oStrSQL = "select a0908 from acc090 where a0901=(select st15 from staff where st01='" & Replace(UCase(oOldMailID), "@TAIE.COM.TW", "") & "' ) and st04='1' and st01>'63001' "
'               'Modify by Morgan 2009/2/10 語法錯誤
'               'oStrSQL = "select a0908 from acc090 where a0901=(select st15 from staff where st01='" & Replace(UCase(oOldMailID), "@TAIE.COM.TW", "") & "' ) and st04='1'  "
'               oStrSQL = "select a0908 from acc090 where a0901=(select st15 from staff where st01='" & Replace(UCase(oOldMailID), "@TAIE.COM.TW", "") & "' )"
'               Set MailRS = New ADODB.Recordset
'               MailRS.CursorLocation = adUseClient
'               MailRS.Open oStrSQL, Conn, adOpenStatic, adLockReadOnly
'               If MailRS.EOF And MailRS.BOF Then
'                   '完全沒有
'                   ChkMailId = "83002"
'               Else
'                   ChkMailId = CheckStr(MailRS.Fields(0))
'               End If
'            End If  '2011/4/13 ADD BY SONIA
'       End If
'   Else
'       ChkMailId = oOldMailID
'   End If
'End Function

'Public Function ConvertToBase64(sPathOrString As String, IfFile As Boolean, AddReturns As Boolean) As String
'  Static Enc() As Byte
'  Dim b1() As Byte, b2() As Byte, B76() As Byte
'  Dim i As Long, i2 As Long, i3 As Long, LFil As Long, NumReturns As Long
'  Dim FF2 As Integer
'     On Error Resume Next
'  If (Not Val(Not Enc)) = 0 Then
'    Enc = StrConv("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/", vbFromUnicode)
'  End If
'
'  If (IfFile = True) Then  '-- if converting picture file.
'         FF2 = FreeFile()
'      Open sPathOrString For Binary As #FF2
'       LFil = LOF(FF2)
'         ReDim b1(0 To (LFil - 1)) As Byte
'        For i = 1 To LFil
'           Get #FF2, i, b1(i - 1)
'        Next
'      Close #FF2
'  Else   '-- converting a string.
'     b1() = StrConv(sPathOrString, vbFromUnicode)
'     LFil = UBound(b1) + 1  '-- added as correction 12-2-04
'  End If
'
'
'  ReDim Preserve b1(0 To ((LFil - 1) \ 3) * 3 + 2)
'  ReDim Preserve b2(0 To (UBound(b1) \ 3) * 4 + 3)
'  For i = 0 To UBound(b1) - 1 Step 3
'    b2(i2) = Enc(b1(i) \ 4)
'      i2 = i2 + 1
'    b2(i2) = Enc((b1(i + 1) \ 16) Or (b1(i) And 3) * 16)
'      i2 = i2 + 1
'    b2(i2) = Enc((b1(i + 2) \ 64) Or (b1(i + 1) And 15) * 4)
'      i2 = i2 + 1
'    b2(i2) = Enc(b1(i + 2) And 63)
'      i2 = i2 + 1
'  Next i
'    For i = 1 To i - LFil
'       b2(UBound(b2) - i + 1) = 61
'    Next i
'
'   If (AddReturns = True) And (LFil > 76) Then
'      '-- add returns every 76 characters before converting to string:
'        NumReturns = ((UBound(b2) + 1) \ 76)
'        LFil = (UBound(b2) + (NumReturns * 2)) '--make B76 B2 plus 2 spots for each vbcrlf.
'         ReDim B76(0 To LFil) As Byte
'          i2 = 0
'          i3 = 0
'        For i = 0 To UBound(b2)
'           B76(i2) = b2(i)
'            i2 = i2 + 1
'            i3 = i3 + 1
'           If (i3 = 76) And (i2 < (LFil - 2)) Then   '--extra check. make sure there are still
'              B76(i2) = 13                      '-- 2 spots left for return if at end.
'              B76(i2 + 1) = 10
'              i2 = i2 + 2
'              i3 = 0
'           End If
'        Next
'     ConvertToBase64 = StrConv(B76, vbUnicode)
'   Else
'     ConvertToBase64 = StrConv(b2, vbUnicode)
'   End If
'End Function

''夾背景圖
'Public Function PrepText(ByVal sText As String) As String  '-- edit IMG tags and change = to =3D
'  Dim sToRep As String, sRep As String, sRet As String
'  On Error Resume Next
'   sText = Trim$(sText)
'          '-- set up Content-ID edit for IMG tags.
'   sToRep = "SRC=" & Chr$(34)              '--  SRC="
'   sRep = "SRC=" & Chr$(34) & "cid:"     '--  SRC="cid:
'     sRet = Replace(sText, sToRep, sRep, 1, -1, vbTextCompare)
'          '-- do content-id tags for body background pic if present.
'   sToRep = "BACKGROUND=" & Chr$(34)
'   sRep = sToRep & "cid:"
'     sRet = Replace(sRet, sToRep, sRep, 1, 1, vbTextCompare)
'          '-- replace "=" with "=3D"
'     sRep = "=3D"
'     sRet = Replace(sRet, "=", sRep, 1, -1, vbTextCompare)
'         '-- wrap lines to 70+- characters. Probably not necessary
'         '-- but it is part of the official format standard.
'     sRep = Clip70(sRet)
'
'   PrepText = sRep
'End Function

'Public Function GetPicHeader(ByVal sFil As String, sID As String) As String
'  Dim s As String, sExt As String ', sID As String
'  Dim tmpMailStr As String
'      On Error Resume Next '-- just in case problem with file name.
'    sExt = Right$(sFil, 3)
'      If UCase$(sExt) = "PEG" Then sExt = "jpg"
'        '-- get a GUID to replace file name for Content-ID.
'        '-- very silly, but apparently an official standard practice.
'    'sID = GetGUID()
'  If (Len(sID) > 0) Then   '--replace file string in html IMG tag with new ID.
'     sBodyText = Replace(sBodyText, "cid:" & sFil, "cid:" & sID, 1, -1, vbTextCompare)
'  Else   '-- if generating GUID was not successful then just assign sfil to sID.
'     sID = sFil
'  End If
'    s = cDASH2 & cBoundaryA & vbCrLf                          ' --Boundary_Taie_A
'    s = s & "Content-Type: image/" & sExt & cSEMIC & vbCrLf   'Content-Type: image/gif;
'    'edit by nickc 2006/10/05 加入編碼
'    's = s & vbTab & "Name=" & Chr$(34) & sFil & Chr$(34) & vbCrLf            '  Name = "pic.gif"
'    tmpMailStr = ConvertToBase64(sFil, False, False)
'    If tmpMailStr <> sFil Then
'      s = s & vbTab & "Name=" & Chr$(34) & "=?Big5?B?" & tmpMailStr & "?=" & Chr$(34) & vbCrLf            '  Name = "pic.gif"
'    Else
'      s = s & vbTab & "Name=" & Chr$(34) & sFil & Chr$(34) & vbCrLf            '  Name = "pic.gif"
'    End If
'    s = s & "Content-Transfer-Encoding: base64" & vbCrLf
'    'add by nickc 2006/10/05 加入
'    s = s & "Content-Disposition: attachment" & cSEMIC & vbCrLf
'    tmpMailStr = ConvertToBase64(sFil, False, False)
'    If tmpMailStr <> sFil Then
'      s = s & vbTab & "filename=""=?Big5?B?" & tmpMailStr & "?=""" & vbCrLf
'    Else
'      s = s & vbTab & "filename=""" & sFil & """" & vbCrLf
'    End If
'    s = s & "Content-ID: <" & sID & ">" & vbCrLf & vbCrLf     'Content-ID: <49984ed.9399.2de3. etc.>  '#
'  GetPicHeader = s
'End Function

'Public Function GetAttachmentHeader(ByVal sFil As String) As String
'  Dim s As String, sExt1 As String, sID As String, sCT1 As String
'  Dim Pt1 As Long
'  Dim Boo1 As Boolean
'      On Error Resume Next
'        '-- get MIME type string:
'   Pt1 = InStrRev(sFil, cDOT)
'     If (Pt1 > 0) Then
'       sExt1 = Right$(sFil, (Len(sFil) - (Pt1 - 1)))  '-- get  ".xxx"
'       Boo1 = RegGetString(HKEY_CLASSES_ROOT, sExt1, "Content Type", sCT1)
'     End If
'       If (Len(sCT1) = 0) Then sCT1 = "application/octet-stream"
'
'    s = cDASH2 & cBoundaryA & vbCrLf                          ' --Boundary_Taie_A
'    s = s & "Content-Type: " & sCT1 & cSEMIC & vbCrLf   'Content-Type: image/gif;
'    s = s & vbTab & "Name=" & Chr$(34) & sFil & Chr$(34) & vbCrLf            '  Name = "pic.gif"
'    s = s & "Content-Transfer-Encoding: base64" & vbCrLf
'    s = s & "Content-Disposition: attachment;" & vbCrLf
'    s = s & vbTab & "Filename=" & Chr$(34) & sFil & Chr$(34) & vbCrLf & vbCrLf
'  GetAttachmentHeader = s
'End Function

'Public Function Clip70(ByVal sIn As String) As String
'  Dim s1 As String, sRet As String
'  Dim Pt1 As Long, PtA As Long, LLen As Long
'    On Error Resume Next
'      LLen = Len(sIn)
'        If (LLen < 76) Then
'           Clip70 = sIn
'           Exit Function
'        End If
'      PtA = 1
'   Do
'     s1 = Mid$(sIn, PtA, 70)
'       Pt1 = InStrRev(s1, R_BRACKET)
'        If (Pt1 > 0) Then
'           s1 = Left$(s1, Pt1) & vbCrLf  '-- get line up to last ">" and add vbcrlf.
'           PtA = (PtA + Pt1)
'        Else  '-- no ">" in line.
'          Pt1 = InStrRev(s1, cSPACE)
'            If (Pt1 > 0) Then
'              s1 = Left$(s1, Pt1) & cEQUAL & vbCrLf
'              PtA = (PtA + Pt1)
'            Else
'              s1 = s1 & cEQUAL & vbCrLf
'              PtA = (PtA + 70)
'            End If
'        End If
'       sRet = sRet & s1
'         Select Case (LLen - PtA)
'            Case Is < 0   '-- if at end of text, then quit.
'               Exit Do
'            Case 0 To 75
'               sRet = sRet & Right$(sIn, ((LLen - PtA) + 1))
'               Exit Do
'            Case Else
'               '--
'         End Select
'    Loop
'      Clip70 = sRet
'End Function

'Public Function GetGUID() As String
' Dim ID1 As GUID
' Dim LRet As Long, LLen As Long
' Dim sID As String
'     On Error Resume Next
'   LRet = CoCreateGuid(ID1)
'     If (LRet = 0) Then
'       LLen = 255
'       sID = String$(LLen, 0)
'         LRet = StringFromGUID2(ID1, StrPtr(sID), LLen)
'           If (LRet > 30) Then
'             sID = Left$(sID, (LRet - 2))  '--snip off { and }
'             sID = Right$(sID, (Len(sID) - 1))
'             sID = Replace(sID, cDASH, cDOT)
'             GetGUID = sID
'           End If
'     End If
'End Function
'Public Function RegGetString(ByVal KeyH As Long, sPath As String, sValName As String, sReturn As String) As Boolean
' Dim hKey As Long, LRet As Long, LBuf As Long, LType As Long
' Dim sBuf As String, sRet As String
'  On Error Resume Next
'     LRet = RegOpenKeyEx(ByVal KeyH, sPath, 0&, KEY_QUERY_VALUE, hKey)
'        If (LRet <> 0) Then Exit Function
'     LBuf = 255
'     sBuf = String$(LBuf, 0)
'       LRet = RegQueryValueExString(hKey, sValName, 0&, LType, sBuf, LBuf)
'          If (LRet = 0) And (LBuf > 1) Then
'             sReturn = Left$(sBuf, (LBuf - 1))
'             RegGetString = True
'          End If
'    LRet = RegCloseKey(hKey)
'End Function

'Public Sub PrepareEmail(ByVal sMailText As String, sFils As String, sID As String)
' Dim s1 As String, sFil64 As String, sFilName As String, sHeader As String
' Dim Pt1 As Long
' Dim AFils As Variant
' Dim i As Integer, iPics As Integer
'  On Error Resume Next
'     If (BooPlain = False) Then  '--html
'        sBodyText = PrepText(sMailText)
'     Else
'        sBodyText = sMailText
'     End If
'
'    ReDim APics(0) As String
'    APics(0) = ""
'    iPics = 0
'    sFils = Trim$(sFils)
'     If (Len(sFils) > 0) Then
'        AFils = Split(sFils, cAST)
'          '-- array of picture file full paths.
'          '-- when done this will have pics in APics array, names in apics(0).
'       For i = 0 To UBound(AFils)
'         s1 = Trim$(AFils(i))
'          If (UCFileExists(s1) = True) Then
'             Pt1 = InStrRev(s1, "\")
'             sFilName = Right$(s1, (Len(s1) - Pt1)) '--get file name.
'             '-- for each file attachment, get header for HTML email content
'             '-- or plain text attachment, then encode the file as Base64.
'             '-- Finally, save the whole thing as a single string in the Apics array.
'                If (BooPlain = False) Then  '--html
'                   '-- this will get header for picture file and also get a GUID to use
'                   '-- as Content-ID. The CID will then replace the picture file name
'                   '-- in the body text to match.
'                    sHeader = GetPicHeader(sFilName, sID & "." & Format(i + 1, "000"))
'                Else
'                    sHeader = GetAttachmentHeader(sFilName)
'                End If
'             sFil64 = ConvertToBase64(s1, True, True)
'             sHeader = sHeader & sFil64 & vbCrLf & vbCrLf
'             iPics = iPics + 1
'             ReDim Preserve APics(0 To iPics) As String
'             APics(iPics) = sHeader
'          End If
'       Next
'     End If
'
'End Sub
'Public Function UCFileExists(sFilPath As String) As Boolean
'  Dim i As Integer
'  On Error Resume Next
'  Err.Clear
'   UCFileExists = False
'     i = GetAttr(sFilPath)
'    If (Err = 0) Then
'       If (i And vbDirectory) = 0 Then
'          UCFileExists = True
'       End If
'    End If
'  Err.Clear
'End Function
'Public Function TextBlurb() As String
'  Dim s As String
'     s = "This is an HTML wewbpage email." & vbCrLf
'     s = s & "It may only be viewed as HTML email, but your email program" & vbCrLf
'     s = s & "does not appear to read HTML format." & vbCrLf & vbCrLf
'  TextBlurb = s
'End Function

Public Function QpDecode(inString As String) As String
   Dim myB     As Byte
   Dim myByte1     As Byte, myByte2       As Byte, myByte3 As Byte 'add by nickc 2006/10/05 加第 3 馬
   Dim convStr()     As Byte
   Dim mOutByte     As Byte
   Dim FinishPercent     As Long
   Dim TotalB, k       As Long
   Dim tmpByte     As Byte
     
   convStr = StrConv(inString, vbFromUnicode)
     
   TotalB = UBound(convStr)
   For k = 0 To TotalB
        myB = convStr(k)
        If myB = Asc("=") Then
              k = k + 1
              'add by nickc 2006/10/05 若是最後一碼為=要跳過
              If k <= UBound(convStr) Then
                myByte1 = convStr(k)
                'myByte1 = convStr(k-1)
                If myByte1 = &HA Then
                            '如果是回?,??
                Else
                            '取第二?字?
                      k = k + 1
                      myByte2 = convStr(k)
                      Call DecodeByte(myByte1, myByte2, mOutByte)
                      If mOutByte >= 127 Then
                            If tmpByte <> 0 Then
                                  QpDecode = QpDecode & Chr(Val("&H" & Hex(tmpByte) & Hex(mOutByte)))
                                  tmpByte = 0
                            Else
                                  tmpByte = mOutByte
                            End If
                      Else
                            'QpDecode = QpDecode & Chr(mOutByte)
                            QpDecode = QpDecode & Chr(Format(tmpByte, "#.0") * 256 + Format(mOutByte, "#.0"))
                            tmpByte = 0
                      End If
                End If
             End If
          Else
                  mOutByte = myB
                  QpDecode = QpDecode & Chr(mOutByte)
          End If
   Next
End Function

Private Sub DecodeByte(mInByte1 As Byte, mInByte2 As Byte, mOutByte As Byte)
Dim tbyte1     As Integer, tbyte2       As Integer
   If mInByte1 > Asc("9") Then
           tbyte1 = mInByte1 - Asc("A") + 10
   Else
           tbyte1 = mInByte1 - Asc("0")
   End If
   If mInByte2 > Asc("9") Then
           tbyte2 = mInByte2 - Asc("A") + 10
   Else
           tbyte2 = mInByte2 - Asc("0")
   End If
   mOutByte = tbyte1 * 16 + tbyte2
End Sub

'Private Sub EncodeByte(mInByte As Byte, mOutStr As String)
''此段因為不曉的為何要這樣，先不做 nickc
''  If (mInByte >= 33 And mInByte <= 60) Or (mInByte >= 62 And mInByte <= 126) Then
''          mOutStr = Chr(mInByte)
''  Else
'          If mInByte <= &HF Then
'                  mOutStr = "=0" & Hex(mInByte)
'          Else
'                  mOutStr = "=" & Hex(mInByte)
'          End If
''  End If
'End Sub
''Quoteprint 編碼
'Public Function ConvertToQp(inString As String) As String
'  Dim myB     As Byte
'  Dim convByte()     As Byte
'  Dim mOutStr     As String
'  Dim FinishPercent     As Long
'  Dim TotalB, k       As Long
'
'  convByte = StrConv(inString, vbFromUnicode)
'
'  TotalB = UBound(convByte)
'  For k = 0 To TotalB
'          myB = convByte(k)
'          EncodeByte myB, mOutStr
'          ConvertToQp = ConvertToQp & mOutStr
'  Next
'End Function

''add by nickc 2007/10/01 取得有時效性的承辦天數
''oSys              系統別
''oNation          國家代碼
''oCp10            案件性質
''oStartDayW    起始日(收文日、發文日、文件齊備日、系統日、...)
''oMaxDayW    最大期限(本所期限、法定期限、系統日、...)
''oCFPCp09     收文號判斷專利國內外案皆為同一工程師使用，沒傳入將不檢查國內外
''oDelayDay      傳入P案延遲天數    只會扣 CFP
''若是 oStartDayW  沒有傳入，則將以系統日作判斷
''※※※※注意：有修改此段程式，須一併修改 basQuery 程式
'Public Function Pub_GetHandleDay(ByVal oSys As String, ByVal oNation As String, ByVal oCp10 As String, Optional ByVal oStartDayW As String = "", Optional ByVal oMaxDayW As String = "", Optional ByVal oCFPCp09 As String = "", Optional ByVal oDelayDay As Integer = 0) As String
'Dim rsgnd As New ADODB.Recordset
'Dim rsgnd2 As New ADODB.Recordset
'Dim oMyStdDate As String
'Dim oMyMaxDate As String
'Dim Strgnd As String
'Dim Strgnd2 As String
'Dim MyCp48 As String
'Dim MyBaseDate As String
'oMyStdDate = IIf(oStartDayW = "", strSrvDate(1), oStartDayW)
''oMyMaxDate = IIf(oMaxDayW = "", strSrvDate(1), oMaxDayW)
'oMyMaxDate = oMaxDayW 'Add by Morgan 2011/2/8 改與basQuery相同
'Set rsgnd = New ADODB.Recordset
'If rsgnd.State = 1 Then rsgnd.Close
'Strgnd = "select cf105 as CFDate,1 as oSort from casefee1 where cf101='" & oSys & "' and cf102='" & oNation & "' and cf103='" & oCp10 & "' and cf104=(select max(cf104) from casefee1 where cf101='" & oSys & "' and cf102='" & oNation & "' and cf103='" & oCp10 & "' and cf104<=" & oMyStdDate & ") "
'Strgnd = Strgnd & " union select cf04,2 from casefee where cf01='" & oSys & "' and cf02='" & oNation & "' and cf03='" & oCp10 & "' order by oSort "
'rsgnd.CursorLocation = adUseClient
'rsgnd.Open Strgnd, Conn, adOpenStatic, adLockReadOnly
'If rsgnd.RecordCount <> 0 Then
'    If CheckStr(rsgnd.Fields("CFDate")) <> "" Then
'        MyBaseDate = CheckStr(rsgnd.Fields("CFDate"))
'        If oCFPCp09 <> "" Then   '有國內外案才會傳入
'            Strgnd2 = "select * from (select C1.cp14 as A,C2.Cp14 as B from casemap,caseprogress c1,caseprogress c2 where cm10='0' and c1.cp01='CFP' and c1.cp09='" & oCFPCp09 & "' and c1.cp10 in (" & GetAddStr(CaseMapIn) & ") "
'            Strgnd2 = Strgnd2 & " and c1.cp01=cm01(+) and c1.cp02=cm02(+) and c1.cp03=cm03(+) and c1.cp04=cm04(+) and cm05=c2.cp01(+) and cm06=c2.cp02(+) and cm07=c2.cp03(+) and cm08=c2.cp04(+) and c2.cp10 in (" & GetAddStr(CaseMapOut) & ") ) AA where AA.A=AA.B "
'            Set rsgnd2 = New ADODB.Recordset
'            If rsgnd2.State = 1 Then rsgnd2.Close
'            rsgnd2.CursorLocation = adUseClient
'            rsgnd2.Open Strgnd2, Conn, adOpenStatic, adLockReadOnly
'            If rsgnd2.RecordCount <> 0 Then
'                MyBaseDate = Trim(Val("10") - oDelayDay)
'            End If
'        End If
'        MyCp48 = CompWorkDay(Val(MyBaseDate), oMyStdDate, 0)
'        If oMyMaxDate <> "" Then
'            If MyCp48 > oMyMaxDate Then
'                MyCp48 = oMyMaxDate
'            End If
'        End If
'    End If
'End If
''Add by Morgan 2011/2/8 若有承辦期限且最大期限小於系統日時，承辦期限設定為系統日
'If MyCp48 <> "" And oMaxDayW <> "" And Val(oMaxDayW) < Val(strSrvDate(1)) Then
'   MyCp48 = strSrvDate(1)
'End If
''end 2011/2/8
'Pub_GetHandleDay = MyCp48
'End Function

''****************************
''      陣列字串加單引號
''  IN
''     Strindex    未加單引號之陣列字串
''  OUT
''     GetAddStr   加單引號之陣列字串
''*****************************
'Public Function GetAddStr(Strindex As String) As String
'Dim StrSeekTemp As Variant, StrTempNewStr As String
'Dim i As Integer
'
'StrSeekTemp = Split(Strindex, ",")
'StrTempNewStr = ""
'For i = 0 To UBound(StrSeekTemp)
'   If Len(Trim(StrSeekTemp(i))) <> 0 Then
'      StrTempNewStr = StrTempNewStr & "'" & StrSeekTemp(i) & "'"
'      If i <> UBound(StrSeekTemp) Then
'         StrTempNewStr = StrTempNewStr & ","
'      End If
'   End If
'Next i
'If Right(StrTempNewStr, 1) = "," Then StrTempNewStr = Left(StrTempNewStr, Len(StrTempNewStr) - 1)
'If Left(StrTempNewStr, 1) = "," Then StrTempNewStr = Right(StrTempNewStr, Len(StrTempNewStr) - 1)
'GetAddStr = StrTempNewStr
'End Function

''add by nickc 2007/10/16 取得要發過期逾期信mail
'Public Function Pub_GetSpecMan(oCode As String) As String
'Dim MyTRS As New ADODB.Recordset
'Set MyTRS = New ADODB.Recordset
'If MyTRS.State = 1 Then MyTRS.Close
'Pub_GetSpecMan = ""
'With MyTRS
'    .CursorLocation = adUseClient
'    .Open "select distinct oMan from setSpecMan where ocode='" & oCode & "' ", Conn, adOpenStatic, adLockReadOnly
'    If .RecordCount <> 0 Then
'        Pub_GetSpecMan = Replace(CheckStr(.Fields("oMan")), ",", ";")
'    End If
'    If Right(Pub_GetSpecMan, 1) = ";" Then
'        Pub_GetSpecMan = Mid(Pub_GetSpecMan, 1, Len(Pub_GetSpecMan) - 1)
'    End If
'End With
'MyTRS.Close
'End Function

'Add by Morgan 2008/6/9
'Removed by Morgan 2013/11/20 不再使用
'Public Function GetSpecID(ByVal sCode As String, Optional sName As String) As String
'
'   If Conn.State = adStateClosed Then
'      Conn.ConnectionString = CnStr
'      Conn.Open
'   End If
'
'   Dim rsTemp1 As New ADODB.Recordset
'   Dim stSQL As String
'
'   stSQL = "SELECT oMan FROM SetSpecMan WHERE oCode='" & sCode & "'"
'   rsTemp1.CursorLocation = adUseClient
'   rsTemp1.Open stSQL, Conn, adOpenForwardOnly, adLockReadOnly
'   If rsTemp1.RecordCount > 0 Then
'      GetSpecID = "" & rsTemp1.Fields(0)
'   End If
'   Set rsTemp1 = Nothing
'
'End Function

'Add by Morgan 2008/8/7
'讀取系統代碼
'Public Function CheckSys(ByRef Strindex As Variant, Optional StrIndex2 As Integer = 0) As String
'   Dim AdoRecordSet3 As New ADODB.Recordset, strSql As String
'
'   strSql = "select sk02,sk03,sk04 from systemkind where sk01='" & Strindex & "' "
'   With AdoRecordSet3
'      .CursorLocation = adUseClient
'      .Open strSql, Conn, adOpenStatic, adLockReadOnly
'      If .RecordCount <> 0 Then
'         CheckSys = CheckStr(.Fields(0))
'         If StrIndex2 <> 0 Then
'            CheckSys = CheckSys & CheckStr(.Fields(1))
'         End If
'      Else
'         CheckSys = ""
'      End If
'   End With
'   Set AdoRecordSet3 = Nothing
'
'End Function

'Add by Morgan 2008/8/7
'關閉資料集
Public Sub CloseRst(p_Rst As ADODB.Recordset)
   If p_Rst.State <> adStateClosed Then p_Rst.Close
End Sub

''Add by Morgan 2008/8/7
''讀取地址相關欄位
'Public Function PUB_GetAddrRef(p_CustNo As String, Optional p_CaseNo1 As String, Optional p_CaseNo2 As String, Optional p_CaseNo3 As String, Optional p_CaseNo4 As String, Optional p_CustName As String, Optional p_Contact As String, Optional p_ZipCode As String, Optional p_Address As String) As Boolean
'   Dim stSQL As String, intR As Integer, iSys As Integer, adoRst As New ADODB.Recordset
'   Dim stContactNo As String
'   Dim CU80 As String, CU87 As String, CU64 As String
'
'   p_CustName = ""
'   p_Contact = ""
'   p_ZipCode = ""
'   p_Address = ""
'
'   '有本所案號
'   If p_CaseNo1 <> "" Then
'      iSys = CheckSys(p_CaseNo1)
'      Select Case iSys
'         Case 1 '專利
'            stSQL = "select pa26,pa149 from patent where pa01='" & p_CaseNo1 & "' and pa02='" & p_CaseNo2 & "' and pa03='" & p_CaseNo3 & "' and pa04='" & p_CaseNo4 & "'"
'         Case 2 '商標
'            stSQL = "select tm23,tm123 from trademark where tm01='" & p_CaseNo1 & "' and tm02='" & p_CaseNo2 & "' and tm03='" & p_CaseNo3 & "' and tm04='" & p_CaseNo4 & "'"
'         Case 3 '法務
'            stSQL = "select lc11,lc42 from lawcase where lc01='" & p_CaseNo1 & "' and lc02='" & p_CaseNo2 & "' and lc03='" & p_CaseNo3 & "' and lc04='" & p_CaseNo4 & "'"
'         Case 4 '顧問
'            stSQL = "select hc05,hc23 from hirecase where hc01='" & p_CaseNo1 & "' and hc02='" & p_CaseNo2 & "' and hc03='" & p_CaseNo3 & "' and hc04='" & p_CaseNo4 & "'"
'         Case Else '服務
'            stSQL = "select sp08,sp78 from servicepractice where sp01='" & p_CaseNo1 & "' and sp02='" & p_CaseNo2 & "' and sp03='" & p_CaseNo3 & "' and sp04='" & p_CaseNo4 & "'"
'      End Select
'      CloseRst adoRst
'      With adoRst
'      .CursorLocation = adUseClient
'      .Open stSQL, Conn, adOpenStatic, adLockReadOnly
'      If .RecordCount <> 0 Then
'         p_CustNo = "" & adoRst.Fields(0)
'         stContactNo = "" & adoRst.Fields(1)
'      End If
'      End With
'   End If
'   p_CustNo = Left(p_CustNo & "000", 9)
'   If stContactNo <> "" Then
'      stSQL = "select * from customer,potcustcont where cu01='" & Left(p_CustNo, 8) & "' and cu02='" & Mid(p_CustNo, 9) & "' and pcc01(+)=cu01 and pcc02(+)='" & stContactNo & "'"
'   Else
'      stSQL = "select * from customer,potcustcont where cu01='" & Left(p_CustNo, 8) & "' and cu02='" & Mid(p_CustNo, 9) & "' and pcc01(+)=cu01 and pcc02(+)=cu127"
'   End If
'   CloseRst adoRst
'   With adoRst
'   .CursorLocation = adUseClient
'   .Open stSQL, Conn, adOpenStatic, adLockReadOnly
'   If .RecordCount <> 0 Then
'      CU80 = "" & .Fields("cu80")
'      CU87 = Left("" & .Fields("cu87"), 3)
'      CU64 = "" & .Fields("cu64")
'      '客戶名稱
'      If Not IsNull(.Fields("CU104")) Then
'         p_CustName = .Fields("CU104")
'      ElseIf Not IsNull(.Fields("CU04")) Then
'         p_CustName = .Fields("CU04")
'      Else
'         p_CustName = Trim("" & .Fields("CU05") & " " & .Fields("CU88") & " " & .Fields("CU89") & " " & .Fields("CU90"))
'      End If
'      '有客戶狀態或定稿語文為中文且非台灣、香港、大陸、澳門時不印地址
'      If CU80 <> "" Or (CU64 = "1" And CU87 > "009" And CU87 <> "013" And CU87 <> "020" And CU87 <> "044") Then
'         p_ZipCode = ""
'         p_Address = ""
'         p_Contact = ""
'      '接洽人有聯絡地址
'      ElseIf Not IsNull(.Fields("PCC22")) Then
'         p_Address = Trim(.Fields("PCC22"))
'         p_ZipCode = "" & .Fields("PCC21")
'         p_Contact = "" & .Fields("PCC05")
'      Else
'         '郵遞區號
'         p_ZipCode = "" & .Fields("CU30")
'         '客戶聯絡地址
'         If Not IsNull(.Fields("CU31")) Then
'            p_Address = Trim(.Fields("CU31"))
'         '客戶中文地址
'         Else
'            p_Address = Trim("" & .Fields("CU23"))
'         End If
'         '接洽人
'         If Not IsNull(.Fields("PCC05")) Then
'            p_Contact = .Fields("PCC05")
'         End If
'      End If
'      PUB_GetAddrRef = True
'   End If
'   End With
'End Function

'Copy By Sindy 98/04/02
'Public Function pvGetExt(ByVal sFileName As String) As String
'    pvGetExt = LCase(Right$(sFileName, 3))
'End Function

''Copy By Sindy 98/04/02
''add by nickc 2005/11/23 取圖用
''edit by nickc 2006/05/24
''Public Function pvGetStdPicture(ByVal sFileName As String, bSuccess As Boolean, Optional pic As PictureBox) As StdPicture
'Public Function pvGetStdPicture(ByVal sFileName As String) As StdPicture
'    'add by nickc 2006/05/09
'    Dim tBI      As BITMAP
'    Dim bSuccess
'
'    On Error Resume Next
'
'    If (pvGetExt(sFileName) = "png" Or pvGetExt(sFileName) = "tif") Then
'
'        '-- Use GDI+ loading
'        'Remove by Morgan 2006/8/10 不再使用
'        'Set pvGetStdPicture = mGDIpEx.LoadPictureEx(sFileName)
'        'end 2006/8/10
'
'    'add by nickc 2006/01/05 修正 wmf and emf 秀圖問題
''    ElseIf (pvGetExt(sFileName) = "wmf" Or pvGetExt(sFileName) = "emf") Then
''        Dim tmpPB As PictureBox
''        Set tmpPB = pic
''        Set pvGetStdPicture = mGDIpEx.LoadPictureExWmf(sFileName, tmpPB)
'    Else
'        '-- Use VB LoadPicture
'        Set pvGetStdPicture = LoadPicture(sFileName)
''        If pvGetStdPicture Is Nothing Then
''            Set pvGetStdPicture = mGDIpEx.LoadPictureEx(sFileName)
''        End If
'    End If
'
'    '-- Is there an image ?
'    bSuccess = Not (pvGetStdPicture Is Nothing)
'
'    If (bSuccess = False) Then
'        '-- Nothing loaded
'        Call MsgBox("無法解析的圖檔！", vbExclamation)
'    End If
''edit by nickc 2006/05/24
''    Call GetObject(pvGetStdPicture.handle, Len(tBI), tBI)
''
''    If tBI.bmWidth > 2000 Or tBI.bmHeight > 2000 Then
''        bSuccess = False
''        Call MsgBox("目前尚不提供圖檔太大的格式！" & vbCrLf & "請將圖的寬與高調整到 2000 像素內！", , "圖片太大！")
''    End If
'    On Error GoTo 0
'End Function

'Copy By Sindy 98/04/02
'add by nickc 2006/05/24
'Public Function FileExists(FileName As String) As Boolean
'    If Len(FileName) > 0 Then FileExists = (Len(Dir(FileName, vbNormal Or vbReadOnly Or vbHidden Or vbSystem Or vbArchive)) > 0)
'End Function

''Copy By Sindy 98/04/02
'Public Sub RidFile(FileName As String)
'    If FileExists(FileName) Then
'        SetAttr FileName, vbNormal
'        Kill FileName
'    End If
'End Sub

''Copy By Sindy 98/04/07
''暫停
'Public Sub Sleep(Strindex As Integer)
'Dim TInt1 As Long, TInt2 As Long
'TInt1 = Timer
'TInt2 = TInt1
'Do While TInt2 - TInt1 < Strindex
'    TInt2 = Timer
'Loop
'End Sub

''Copy By Sindy 2009/05/26
'Public Function PUB_CopyImgFile(p_fPA() As String, p_tPA() As String) As Boolean
'   Dim stSQL As String, iTmp As Integer, iFields As Integer
'   Dim ibf() As String
'   Dim bytes() As Byte
'
'   Set cnnConnection = New ADODB.Connection
'   cnnConnection.ConnectionString = CnStr
'   cnnConnection.Open
'
'   stSQL = "select * from imgbytefile where ibf01='" & p_fPA(1) & "' and ibf02='" & p_fPA(2) & "' and ibf03='" & p_fPA(3) & "' and ibf04='" & p_fPA(4) & "' and ibf05='1'"
'   iTmp = 1
'   CheckOC3
'   With AdoRecordSet3
'      .CursorLocation = adUseClient
'      .Open stSQL, cnnConnection, adOpenStatic, adLockOptimistic
'      If .RecordCount > 0 Then
'         iFields = .Fields.Count
'         ReDim ibf(iFields - 1)
'         For iTmp = 0 To iFields - 1
'            If LCase(.Fields(iTmp).Name) = "ibf14" Then
'               ReDim bytes(Val(.Fields("ibf13").Value))
'               bytes() = .Fields(iTmp).GetChunk(Val(.Fields("ibf13")))
'            ElseIf LCase(.Fields(iTmp).Name) = "ibf07" Or _
'               LCase(.Fields(iTmp).Name) = "ibf10" Or _
'               LCase(.Fields(iTmp).Name) = "ibf11" Or _
'               LCase(.Fields(iTmp).Name) = "ibf12" Then
'               ibf(iTmp) = ""
'            ElseIf LCase(.Fields(iTmp).Name) = "ibf08" Then
'               ibf(iTmp) = Val(Format(Date, "yyyymmdd"))
'            ElseIf LCase(.Fields(iTmp).Name) = "ibf09" Then
'               ibf(iTmp) = Val(Format(Time, "HHMM"))
'            Else
'               ibf(iTmp) = "" & .Fields(iTmp)
'            End If
'         Next
'         .AddNew
'
'On Error GoTo CancleUpdate
'
'         For iTmp = 0 To iFields - 1
'            If LCase(.Fields(iTmp).Name) = "ibf14" Then
'               .Fields(iTmp).AppendChunk bytes()
'            ElseIf LCase(.Fields(iTmp).Name) = "ibf01" Then
'               .Fields(iTmp) = p_tPA(1)
'            ElseIf LCase(.Fields(iTmp).Name) = "ibf02" Then
'               .Fields(iTmp) = p_tPA(2)
'            ElseIf LCase(.Fields(iTmp).Name) = "ibf03" Then
'               .Fields(iTmp) = p_tPA(3)
'            ElseIf LCase(.Fields(iTmp).Name) = "ibf04" Then
'               .Fields(iTmp) = p_tPA(4)
'            ElseIf ibf(iTmp) <> "" Then
'               .Fields(iTmp) = ibf(iTmp)
'            End If
'         Next
'         .UPDATE
'         PUB_CopyImgFile = True
'      End If
'   End With
'   Exit Function
'
'CancleUpdate:
'   AdoRecordSet3.CancelUpdate
'
'End Function

''關閉連線
'Public Sub CheckOC3()           'NICK
'If AdoRecordSet3.State <> 0 Then
'   AdoRecordSet3.Close
'End If
'End Sub

'Public Function CompDate(ByVal iSitu As Integer, ByVal iNum As Single, ByVal strTemp As String) As String
' Dim i As Integer, s As Single, strTmp As String
'   If strTemp = "" Then CompDate = "": Exit Function
'   If Len(strTemp) <> 8 Then strTemp = Format(Val(strTemp) + 19110000)
'   Select Case iSitu
'      Case 0 '年
'         strTmp = Format(iNum)
'         'Modify by Morgan 2005/7/22
'         'CompDate = Format(Val(Left(strTemp, 4)) + iNum) & Mid(strTemp, 5, 2) & Right(strTemp, 2)
'         CompDate = Format(DateAdd("yyyy", Int(iNum), DateSerial(Left(strTemp, 4), Mid(strTemp, 5, 2), Right(strTemp, 2))), "YYYYMMDD")
'         i = InStr(strTmp, ".")
'         If i > 0 Then
'            s = (Val(iNum) - Int(iNum)) * 12
'            CompDate = CompDate(1, s, CompDate)
'         End If
'      Case 1 '月
'         CompDate = Format(DateAdd("M", iNum, DateSerial(Left(strTemp, 4), Mid(strTemp, 5, 2), Right(strTemp, 2))), "YYYYMMDD")
'      Case 2 '日
'         CompDate = Format(DateAdd("D", iNum, DateSerial(Left(strTemp, 4), Mid(strTemp, 5, 2), Right(strTemp, 2))), "YYYYMMDD")
'   End Select
'End Function

'Public Function Pub_GetModuleFileName() As String
''宣告變數
'Dim str As String
'
'Pub_GetModuleFileName = ""
'str = String(MAX_PATH, "#")
'GetModuleFileName App.hInstance, str, MAX_PATH
'Pub_GetModuleFileName = Replace(str, "#", "")
'
'End Function

''取得中英混雜字串之長度
'Public Function GetTextLength(ByRef strTemp As String) As Integer
'GetTextLength = LenB(StrConv(strTemp, vbFromUnicode))
'End Function

''Add By Cheng 2003/12/16
''取的新字串長度
''Modify by Morgan 2010/1/19 +bolAddSpace:是否補空白,bolAddLeft:空白補左邊
'Public Function PUB_StrToStr(ByRef Strindex As String, ByRef StrIndex2 As Single, Optional bolAddSpace As Boolean = False, Optional bolAddLeft As Boolean = False) As String
'Dim ii As Integer
'
'PUB_StrToStr = ""
''若傳入的字串非空字串
'If Strindex <> "" Then
'    For ii = 1 To Len(Strindex)
'        If LenB(StrConv(PUB_StrToStr & Mid(Strindex, ii, 1), vbFromUnicode)) <= StrIndex2 Then
'            PUB_StrToStr = PUB_StrToStr & StrConv(MidB(StrConv(Mid(Strindex, ii, 1), vbFromUnicode), 1, 2), vbUnicode)
'        Else
'            Exit For
'        End If
'    Next ii
'    'Add by Morgan 2010/1/19
'    If bolAddSpace Then
'      If bolAddLeft Then
'         PUB_StrToStr = String(StrIndex2 - LenB(StrConv(PUB_StrToStr, vbFromUnicode)), " ") & PUB_StrToStr
'      Else
'         PUB_StrToStr = PUB_StrToStr & String(StrIndex2 - LenB(StrConv(PUB_StrToStr, vbFromUnicode)), " ")
'      End If
'   End If
'End If
'End Function

''Add by Sindy 2011/9/6
''取得人事室出缺勤電子簽核收E-Mail人員
'Public Function GetM21EMailPerson() As String
'Dim rsTmp As New ADODB.Recordset
'
'   strSql = "SELECT OMAN FROM SetSpecMan where OCODE='人事室出缺勤電子簽核' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, Conn, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      If Not IsNull(rsTmp.Fields(0)) Then GetM21EMailPerson = rsTmp.Fields(0)
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
'End Function

'Add by Sindy 2011/9/6
'*************************************
'檢查 員工是否離職
'   strST01     員工編號
'*************************************
'Public Function ChkStaffST04(strST01 As String, Optional bolMsg As Boolean = True) As Boolean
'   Dim s As Integer
'
'   CheckOC3
'   strSql = "select st01 from staff where st01='" & strST01 & "' and st04='1' "
'   With AdoRecordSet3
'      .CursorLocation = adUseClient
'      .Open strSql, Conn, adOpenStatic, adLockReadOnly
'      If .RecordCount <> 0 Then
'         ChkStaffST04 = False
'      Else
'         ChkStaffST04 = True
'         If bolMsg = True Then s = MsgBox("此人員不存在或已離職！！", , "人員錯誤！！")
'      End If
'   End With
'   CheckOC3
'End Function

''Add by Sindy 2011/9/6
''傳回表單確認(每月出缺勤統計確認)的下一處理人員
'Public Function GetNextB1303(strKEY01 As String, strKEY02 As String) As String
'Dim rsTmp As New ADODB.Recordset
'
'   GetNextB1303 = ""
'
'   strSql = "SELECT 1,B0108 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and B0108 is not null and B1304 is null " & _
'            "Union " & _
'            "SELECT 2,B0109 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and B0109 is not null and B1305 is null " & _
'            "Union " & _
'            "SELECT 3,B0110 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and B0110 is not null and B1306 is null " & _
'            "Union " & _
'            "SELECT 4,B0111 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and B0111 is not null and B1307 is null " & _
'            "order by 1 asc"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, Conn, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      GetNextB1303 = rsTmp.Fields(1)
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
'End Function

''Add by Sindy 2011/9/6
''傳回表單確認(每月出缺勤統計確認)的目前處理人員其審核主管為何
'Public Function GetCurrB1303Seqno(strKEY01 As String, strKEY02 As String, strKEY03 As String) As Integer
'Dim rsTmp As New ADODB.Recordset
'
'   GetCurrB1303Seqno = 0
'
'   strSql = "SELECT 1,B0108 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and '" & strKEY03 & "'=B0108(+) " & _
'            "Union " & _
'            "SELECT 2,B0109 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and '" & strKEY03 & "'=B0109(+) " & _
'            "Union " & _
'            "SELECT 3,B0110 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and '" & strKEY03 & "'=B0110(+) " & _
'            "Union " & _
'            "SELECT 4,B0111 FROM ABS013,ABS001 WHERE B1301='" & strKEY01 & "' and B1302='" & strKEY02 & "' and B1302=B0101(+) and '" & strKEY03 & "'=B0111(+) " & _
'            "order by 1 asc"
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, Conn, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      With rsTmp
'         .MoveFirst
'         Do While Not .EOF
'            If Not IsNull(rsTmp.Fields(1)) Then
'               GetCurrB1303Seqno = rsTmp.Fields(0): Exit Do
'            End If
'            .MoveNext
'         Loop
'      End With
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
'End Function

'Removed by Morgan 2014/3/10 整合同名函數到 basQuery
''*************************************************
''  電腦自動給號
''
''*************************************************
'Public Function AutoNo(InputItem As String, InputLength As Integer) As String
'Dim adoaccnum As New ADODB.Recordset
'Dim strItem As String, strYes As String
'
'
''911106 NICK '911106 nick 避免相同連線作做2次 transation
'Dim BolTransOk As Boolean
'BolTransOk = True
'On Error GoTo TransErr
'
'   adoTaie.BeginTrans
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
'         AutoNo = strItem & Mid(ACDate(strSrvDate(1)), 1, 3) & ZeroBeforeNo("2000", InputLength)
'      Else
'         If InputItem = "U" Or InputItem = "V" Then
'            AutoNo = strItem & Mid(ACDate(strSrvDate(1)), 1, 3) & ZeroBeforeNo("0", InputLength)
'         Else
'            'Modify By Sindy 2010/8/17 比對自動編號年度
'            'Modify by Morgan 2011/1/5 +K
'            If InputItem = "A" Or InputItem = "B" Or InputItem = "C" Or _
'               InputItem = "D" Or InputItem = "DP" Or InputItem = "K" Then
'               AutoNo = strItem & CompAutoNumberYear(Val(Mid(ACDate(strSrvDate(1)), 1, 3))) & ZeroBeforeNo("0", InputLength)
'            '2010/8/17 End
'            Else
'               AutoNo = strItem & Val(Mid(ACDate(strSrvDate(1)), 1, 3)) & ZeroBeforeNo("0", InputLength)
'            End If
'         End If
'      End If
'   Else
'      If adoaccnum.Fields("au02").Value <> Val(Mid(strSrvDate(1), 1, 4)) Then
'         If InputItem = "E" Then
'            AutoNo = strItem & Mid(ACDate(strSrvDate(1)), 1, 3) & ZeroBeforeNo("2000", InputLength)
'         Else
'            If InputItem = "U" Or InputItem = "V" Then
'               AutoNo = strItem & Mid(ACDate(strSrvDate(1)), 1, 3) & ZeroBeforeNo("0", InputLength)
'            Else
'               'Modify By Sindy 2010/8/17 比對自動編號年度
'               'Modify by Morgan 2011/1/5 +K
'               If InputItem = "A" Or InputItem = "B" Or InputItem = "C" Or _
'                  InputItem = "D" Or InputItem = "DP" Or InputItem = "K" Then
'                  AutoNo = strItem & CompAutoNumberYear(Val(Mid(ACDate(strSrvDate(1)), 1, 3))) & ZeroBeforeNo("0", InputLength)
'               '2010/8/17 End
'               Else
'                  AutoNo = strItem & Val(Mid(ACDate(strSrvDate(1)), 1, 3)) & ZeroBeforeNo("0", InputLength)
'               End If
'            End If
'         End If
'      Else
'         If InputItem = "U" Or InputItem = "V" Then
'            AutoNo = strItem & Mid(ACDate(strSrvDate(1)), 1, 3) & ZeroBeforeNo(str(adoaccnum.Fields("au03").Value), InputLength)
'         Else
'            'Modify By Sindy 2010/8/17 比對自動編號年度
'            'Modify by Morgan 2011/1/5 +K
'            If InputItem = "A" Or InputItem = "B" Or InputItem = "C" Or _
'               InputItem = "D" Or InputItem = "DP" Or InputItem = "K" Then
'               AutoNo = strItem & CompAutoNumberYear(Val(Mid(ACDate(strSrvDate(1)), 1, 3))) & ZeroBeforeNo(str(adoaccnum.Fields("au03").Value), InputLength)
'            '2010/8/17 End
'            Else
'               AutoNo = strItem & Val(Mid(ACDate(strSrvDate(1)), 1, 3)) & ZeroBeforeNo(str(adoaccnum.Fields("au03").Value), InputLength)
'            End If
'         End If
'      End If
'   End If
'   If Len(InputItem) = 1 Then
'      If InputItem = "U" Or InputItem = "V" Then
'         strYes = SaveAutoNo(InputItem, Mid(AutoNo, 5, InputLength))
'      Else
'         'Modify By Sindy 2010/8/17
'         'Modify by Morgan 2011/1/5 +K
'         If InputItem = "A" Or InputItem = "B" Or InputItem = "C" Or _
'            InputItem = "D" Or InputItem = "DP" Or InputItem = "K" Or _
'            Val(Mid(ACDate(strSrvDate(1)), 1, 3)) <= 99 Then
'            strYes = SaveAutoNo(InputItem, Mid(AutoNo, 4, InputLength))
'         '2010/8/17 End
'         Else
'            strYes = SaveAutoNo(InputItem, Mid(AutoNo, 5, InputLength))
'         End If
'      End If
'   End If
'   adoaccnum.Close
'   If BolTransOk Then
'        adoTaie.CommitTrans
'   End If
''911106 nick 避免相同連線作做2次 transation
'   Exit Function
'TransErr:
'   If Err.Number = -2147168237 Then
'      BolTransOk = False
'      Resume Next
'   End If
'End Function
'
''*************************************************
''  電腦給號存檔
''
''*************************************************
'Public Function SaveAutoNo(InputItem, InputNo As String) As String
'Dim adoaccnum As New ADODB.Recordset
'   adoaccnum.CursorLocation = adUseClient
'   adoaccnum.Open "select * from autonumber where au01 = '" & InputItem & "'", cnnConnection, adOpenDynamic, adLockBatchOptimistic
'   If adoaccnum.RecordCount = 0 Then
'      adoaccnum.AddNew
'      adoaccnum.Fields("au01").Value = InputItem
'   End If
'   adoaccnum.Fields("au02").Value = Mid(strSrvDate(1), 1, 4)
'   adoaccnum.Fields("au03").Value = InputNo
'   adoaccnum.UpdateBatch
'   adoaccnum.Close
'   SaveAutoNo = "Y"
'End Function
'end 2014/3/10

'Add By Sindy 2014/3/26 已積欠總金額
'intKind     : 0.A1k03 1.A1k28
'Modify By Sindy 2014/5/12 +Optional ByRef strPayDate As String
Public Function GetCuFAAccNotPayAmt(ByVal strCustNo As String, ByVal intKind As Integer, _
                     Optional ByRef strName As String, Optional ByRef strIsNotPay As String, _
                     Optional ByRef strPayDATE As String) As Double
   Dim strAccMemo As String 'Add by Amy 2020/09/21
   
   GetCuFAAccNotPayAmt = 0
   If strCustNo = "" Then Exit Function 'Add By Sindy 2020/3/17
   'A1k03.案件代理人
   If intKind = 0 Then
      strExc(0) = "Select sum(nvl(a1k11-nvl(a1k06,0)-nvl(a1k30,0),0)) From acc1k0" & _
                  " Where a1k03='" & strCustNo & "'" & _
                  " and a1k29 is null" & _
                  " and a1k12 is null and a1k25 is null and a1k17 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         GetCuFAAccNotPayAmt = RsTemp.Fields(0)
         'Add By Sindy 2014/5/12
         '最早請款日期
         strExc(0) = "Select a1k02 From acc1k0" & _
                     " Where a1k03='" & strCustNo & "'" & _
                     " and a1k29 is null" & _
                     " and a1k12 is null and a1k25 is null and a1k17 is null" & _
                     " order by a1k02 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            strPayDATE = RsTemp.Fields(0)
         End If
         '2014/5/12 END
      End If
   'A1k28.請款對象
   ElseIf intKind = 1 Then
      strExc(0) = "Select sum(nvl(a1k11-nvl(a1k06,0)-nvl(a1k30,0),0)) From acc1k0" & _
                  " Where a1k28='" & strCustNo & "'" & _
                  " and a1k29 is null" & _
                  " and a1k12 is null and a1k25 is null and a1k17 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         GetCuFAAccNotPayAmt = RsTemp.Fields(0)
         'Add By Sindy 2014/5/12
         '最早請款日期
         strExc(0) = "Select a1k02 From acc1k0" & _
                     " Where a1k28='" & strCustNo & "'" & _
                     " and a1k29 is null" & _
                     " and a1k12 is null and a1k25 is null and a1k17 is null" & _
                     " order by a1k02 asc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsTemp.MoveFirst
            strPayDATE = RsTemp.Fields(0)
         End If
         '2014/5/12 END
      End If
   End If
   If GetCuFAAccNotPayAmt > 0 Then
      'Modify by Amy 2020/09/21 +會計備註 cu79/fa118
      '客戶
      If Left(strCustNo, 1) = "X" Then
         strExc(0) = "Select Decode(CU05,Null, Nvl(CU04, CU06), CU05||' '||CU88||' '||CU89||' '||CU90),CU140,cu79 From customer" & _
                     " Where cu01='" & Left(strCustNo, 8) & "' And cu02='" & Right(strCustNo, 1) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strName = RsTemp.Fields(0)
            strIsNotPay = "" & RsTemp.Fields("CU140")
            strAccMemo = "" & RsTemp.Fields("cu79")
         End If
      '代理人
      Else
         strExc(0) = "Select Decode(FA05,Null, Nvl(FA04, FA06), FA05||' '||FA63||' '||FA64||' '||FA65),FA101,fa118 From Fagent" & _
                     " Where fa01='" & Left(strCustNo, 8) & "' And fa02='" & Right(strCustNo, 1) & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strName = RsTemp.Fields(0)
            strIsNotPay = "" & RsTemp.Fields("FA101")
            strAccMemo = "" & RsTemp.Fields("fa118")
         End If
      End If
      'Modify by Amy 2020/09/21 原:cu140/fa101為N 不寄催款單 改輸1-3
      'If strIsNotPay = "N" Then strIsNotPay = "永久設定不催款"
      If strIsNotPay <> MsgText(601) Then
        Select Case strIsNotPay
            Case "1"
                strIsNotPay = "每月寄對帳單"
            Case "2"
                strIsNotPay = "客戶要求不寄對帳單"
            Case "3"
                strIsNotPay = strAccMemo
        End Select
      End If
      'end 2020/09/21
   End If
End Function

'Add By Sindy 2022/10/11 因為Account有引用 basFlow 會連帶需要引用到 Service1
'但因接洽單電子收文就會呼叫到一些案件系統函數, 所以才建此虛函數
Public Function PUB_AutoRecvCRLMain(strSys As String, strCRL01 As String) As Boolean
End Function
