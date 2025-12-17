VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ClsPrtForm001"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

Dim m_NP01 As String ' 下一程序檔的總收文號
' 本所案號
Dim m_NP02 As String
Dim m_NP03 As String
Dim m_NP04 As String
Dim m_NP05 As String
Dim m_NP07 As String ' 下一程序代號
Dim m_NP10 As String ' 智權人員代號
Dim m_NP10_Name As String ' 智權人員姓名
Dim m_NP10_Dep As String ' 智權人員部門別
Dim m_NP22 As String ' 下一程序檔的序號
Dim m_NP08 As String ' 本所期限
Dim m_NP09 As String ' 法定期限
Dim m_NP23 As String ' 約定期限
Dim m_Nation As String ' 國家代碼
Dim m_Agent As String ' 代理人
'Add By Cheng 2003/06/23
Dim m_TM45 As String '彼所案號
Dim m_TM09 As String ' 商品類別
Dim m_TM12 As String ' 申請案號
Dim m_TM15 As String ' 審定號或專利號數
Dim m_Customer As String ' 申請人代號
Dim m_Customer_CName As String ' 申請人中文名稱
Dim m_Customer_EName As String ' 申請人英文名稱
' 申請人姓名
' 負責人
Dim m_Customer_Main As String
' 接洽人
Dim m_Customer_Keep As String
' 申請人地址
Dim m_Customer_Adrs As String
' 申請人電話
Dim m_Customer_Tel As String
' 申請人傳真
Dim m_Customer_Fax As String
' 申請人國籍
Dim m_Customer_Cny As String
' 申請地址(中)
Dim m_Customer_CAddress As String
' 申請地址(英)
Dim m_Customer_EAddress As String
' 案件名稱
Dim m_TM05 As String
'Add By Cheng 2002/10/25
Dim m_CP01 As String
Dim m_CP02 As String
Dim m_CP03 As String
Dim m_CP04 As String
Dim m_CP10 As String
Dim m_CP12 As String
Dim m_CP12_Name As String
Dim m_CP12_Dep As String
Dim m_CP06 As String
Dim m_CP07 As String
Dim m_CP09 As String
Dim more_Customer As String  '2005/11/2 ADD BY SONIA 多申請人
Dim m_CP05 As String 'Add by Morgan 2006/6/26 收文日
'add by nickc 2008/01/18
Dim m_CuMemo As String
Dim m_SpecCu As String
Dim m_PA08 As String 'Add By Sindy 2009/06/29


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 檢查字串是否為空白
' Input : strData ==> 所要檢查的字串
' Output : 若輸入的字串中沒有非空白的字元存在時, 則傳回True
'          否則傳回 False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function IsEmptyText(ByVal strData As String) As Boolean
   Dim nIndex As Integer
   IsEmptyText = False
   
   If Len(strData) <= 0 Then
      IsEmptyText = True
   Else
      IsEmptyText = True
      For nIndex = 1 To Len(strData)
         If Mid(strData, nIndex, 1) <> " " Then
            IsEmptyText = False
            Exit For
         End If
      Next nIndex
   End If
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 由系統別及案件性質代碼取得案件性質名稱
' Input : strSysKind ==> 本所案號中的系統別
'         strCaseNo ==> 案件性質代碼
'         nLanguage ==> 0 表取得國內案件性質名稱
'                       1 表取得大陸件性質名稱
' Output : 傳回案件性質名稱
' 說明 : 當nCountry不輸入時, 預設值為取得國內的案件性質名稱
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetCaseTypeName(ByVal strSysKind As String, ByVal strCaseNo As String, Optional ByVal nLanguage As Integer = 0) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   ' 檢查取得國內還是大陸名稱的範圍
   If nLanguage < 0 Or nLanguage > 1 Then: nLanguage = 0
   
   GetCaseTypeName = Empty
   strSql = "SELECT * FROM CasePropertyMap " & _
            "WHERE CPM01 = '" & strSysKind & "' AND " & _
                  "CPM02 = '" & strCaseNo & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Select Case nLanguage
         ' 表取得國內的名稱
         Case 0:
            If IsNull(rsTmp.Fields("CPM03")) = False Then
               GetCaseTypeName = rsTmp.Fields("CPM03")
            End If
         ' 表取得大陸的名稱
         Case 1:
            If IsNull(rsTmp.Fields("CPM04")) = False Then
               GetCaseTypeName = rsTmp.Fields("CPM04")
            End If
      End Select
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得國家的名稱
' Input : strNation ==> 國家代碼
'         nLanguage ==> 0 表取得國內名稱
'                       1 表取得國外名稱
' Output : 傳回國家的名稱
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetNationName(ByVal strNation As String, Optional ByVal nLanguage As Integer) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetNationName = Empty
   
   ' 檢查取得國內還是國外名稱的範圍
   If nLanguage < 0 Or nLanguage > 1 Then: nLanguage = 0
   
   strSql = "SELECT * FROM NATION " & _
            "WHERE NA01 = '" & strNation & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Select Case nLanguage
         Case 0:
            If IsNull(rsTmp.Fields("NA03")) = False Then
               GetNationName = rsTmp.Fields("NA03")
            End If
         Case 1:
            If IsNull(rsTmp.Fields("NA04")) = False Then
               GetNationName = rsTmp.Fields("NA04")
            End If
      End Select
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得員工的姓名
' Input : strStuff ==> 員工的代碼
'         bAll ==> 是否包含已離職的員工 (不輸入此參數時預設為False)
'                  True : 表包含所有已離職的員工
'                  False : 只查詢在職的員工 (預設值)
' Output : 傳回員工的姓名
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetStaffName(ByVal strStuff As String, Optional ByVal bAll As Boolean = False) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetStaffName = Empty
   
   strSql = "SELECT * FROM Staff " & _
            "WHERE ST01 = '" & strStuff & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("ST02")) = False Then
         GetStaffName = rsTmp.Fields("ST02")
      End If
      If bAll = False Then
         If IsNull(rsTmp.Fields("ST04")) = False Then
            If rsTmp.Fields("ST04") = "2" Then
               GetStaffName = Empty
            End If
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得國外代理人的姓名
' Input : strAgent ==> 代理人的代碼
' Output : 傳回國外代理人的姓名
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetFAgentName(ByVal strAgent As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim strFA01 As String
   Dim strFA02 As String
   
   ' 設定 KEY 值
   strFA01 = Mid(strAgent, 1, 8)
   If Len(strFA01) < 8 Then: strFA01 = strFA01 & String(8 - Len(strFA01), "0")
   strFA02 = Mid(strAgent, 9, 1)
   If strFA02 = Empty Then: strFA02 = "0"
   
   ' 設定初始值
   GetFAgentName = Empty
   
   ' 組成SQL語法
   strSql = "SELECT * FROM FAGENT " & _
            "WHERE FA01 = '" & strFA01 & "' AND " & _
                  "FA02 = '" & strFA02 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   
   ' 判斷是否有資料
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("FA05")) = False Then
         GetFAgentName = rsTmp.Fields("FA05")
      End If
      If IsEmptyText(GetFAgentName) = True Then
         If IsNull(rsTmp.Fields("FA63")) = False Then
            GetFAgentName = rsTmp.Fields("FA63")
         End If
      End If
      If IsEmptyText(GetFAgentName) = True Then
         If IsNull(rsTmp.Fields("FA64")) = False Then
            GetFAgentName = rsTmp.Fields("FA64")
         End If
      End If
      If IsEmptyText(GetFAgentName) = True Then
         If IsNull(rsTmp.Fields("FA65")) = False Then
            GetFAgentName = rsTmp.Fields("FA65")
         End If
      End If
      If IsEmptyText(GetFAgentName) = True Then
         If IsNull(rsTmp.Fields("FA04")) = False Then
            GetFAgentName = rsTmp.Fields("FA04")
         End If
      End If
      If IsEmptyText(GetFAgentName) = True Then
         If IsNull(rsTmp.Fields("FA06")) = False Then
            GetFAgentName = rsTmp.Fields("FA06")
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 由客戶代碼取得客戶名稱
' Input : strCustomer ==> 客戶代碼
'         nLanguage ==> 0 : 表取得的是中文名稱(無中文取英文, 無英文取日文)
'                       1 : 表取得的是英文名稱(無英文取中文, 無中文取日文)
'                       (預設值為0)
' Output : 傳回客戶的名稱
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetCustomerName(ByVal strCustomer As String, Optional ByVal nLanguage As String = 0) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   ' 檢查取得中文還是英文名稱的範圍
   If nLanguage < 0 Or nLanguage > 1 Then: nLanguage = 0
   
   GetCustomerName = Empty
   
   If Len(strCustomer) < 9 Then: strCustomer = strCustomer & String(9 - Len(strCustomer), "0")
   
   If Len(strCustomer) > 8 Then
      strSql = "SELECT * FROM Customer " & _
               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                     "CU02 = '" & Mid(strCustomer, 9, 1) & "'"
   Else
      strSql = "SELECT * FROM Customer " & _
               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                     "CU02 = '0' "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Select Case nLanguage
         Case 0:
            If IsNull(rsTmp.Fields("CU04")) = False Then
               GetCustomerName = rsTmp.Fields("CU04")
            ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
               GetCustomerName = rsTmp.Fields("CU05")
            ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
               GetCustomerName = rsTmp.Fields("CU06")
            End If
         Case 1:
            If IsNull(rsTmp.Fields("CU05")) = False Then
               GetCustomerName = rsTmp.Fields("CU05")
            ElseIf IsNull(rsTmp.Fields("CU04")) = False Then
               GetCustomerName = rsTmp.Fields("CU04")
            ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
               GetCustomerName = rsTmp.Fields("CU06")
            End If
      End Select
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
'add by nickc 2008/04/23 台灣商標專用
Private Function GetCustomerNameT000(ByVal strCustomer As String, Optional ByVal nLanguage As String = 0) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   ' 檢查取得中文還是英文名稱的範圍
   If nLanguage < 0 Or nLanguage > 1 Then: nLanguage = 0
   
   GetCustomerNameT000 = Empty
   
   If Len(strCustomer) < 9 Then: strCustomer = strCustomer & String(9 - Len(strCustomer), "0")
   
   If Len(strCustomer) > 8 Then
      strSql = "SELECT * FROM Customer " & _
               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                     "CU02 = '" & Mid(strCustomer, 9, 1) & "'"
   Else
      strSql = "SELECT * FROM Customer " & _
               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                     "CU02 = '0' "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Select Case nLanguage
         Case 0:
            If IsNull(rsTmp.Fields("CU104")) = False Then
                GetCustomerNameT000 = rsTmp.Fields("CU104")
            ElseIf IsNull(rsTmp.Fields("CU04")) = False Then
               GetCustomerNameT000 = rsTmp.Fields("CU04")
            ElseIf IsNull(rsTmp.Fields("CU05")) = False Then
               GetCustomerNameT000 = rsTmp.Fields("CU05") & " " & "" & rsTmp.Fields("CU88") & " " & "" & rsTmp.Fields("CU89") & " " & "" & rsTmp.Fields("CU90")
            ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
               GetCustomerNameT000 = rsTmp.Fields("CU06")
            End If
         Case 1:
            If IsNull(rsTmp.Fields("CU05")) = False Then
               GetCustomerNameT000 = rsTmp.Fields("CU05") & " " & "" & rsTmp.Fields("CU88") & " " & "" & rsTmp.Fields("CU89") & " " & "" & rsTmp.Fields("CU90")
            ElseIf IsNull(rsTmp.Fields("CU104")) = False Then
               GetCustomerNameT000 = rsTmp.Fields("CU104")
            ElseIf IsNull(rsTmp.Fields("CU04")) = False Then
               GetCustomerNameT000 = rsTmp.Fields("CU04")
            ElseIf IsNull(rsTmp.Fields("CU06")) = False Then
               GetCustomerNameT000 = rsTmp.Fields("CU06")
            End If
      End Select
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
'add by nickc 2008/04/23 非台灣商標專用
Private Function GetCustomerNameTnot000(ByVal strCustomer As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   
   GetCustomerNameTnot000 = Empty
   
   If Len(strCustomer) < 9 Then: strCustomer = strCustomer & String(9 - Len(strCustomer), "0")
   
   If Len(strCustomer) > 8 Then
      strSql = "SELECT * FROM Customer " & _
               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                     "CU02 = '" & Mid(strCustomer, 9, 1) & "'"
   Else
      strSql = "SELECT * FROM Customer " & _
               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                     "CU02 = '0' "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
        If IsNull(rsTmp.Fields("CU04")) = False Then
           GetCustomerNameTnot000 = rsTmp.Fields("CU04")
        End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得西元的年份
' Input : strDate ==> 輸入的日期
' Output : 傳回西元的年份 YYYY
' Description : 此功能會傳回日期字串中的年, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DBYEAR(ByVal strDate As String) As String
   Dim nIndex As Integer
   Dim strTemp As String
   ' 若日期中有/表示這是日期格式的字串如 YY/MM/DD 或 YYYY/MM/DD
   DBYEAR = "0000"
   If InStr(1, strDate, "/") Then
      nIndex = InStr(1, strDate, "/")
      strTemp = Mid(strDate, 1, nIndex - 1)
      If Len(strTemp) < 4 Then
         strTemp = str(Val(strTemp) + 1911)
      End If
      DBYEAR = strTemp
   Else
      If Len(strDate) > 7 Then
         DBYEAR = Mid(strDate, 1, 4)
      Else
         strTemp = strDate
         DBYEAR = Mid(ChangeTStringToWString(strTemp), 1, 4)
      End If
   End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得西元的月份
' Input : strDate ==> 輸入的日期
' Output : 傳回西元的月份 MM
' Description : 此功能會傳回日期字串中的年, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DBMONTH(ByVal strDate As String) As String
   Dim nBegin As Integer
   Dim nEnd As Integer
   Dim strTemp As String
   DBMONTH = "00"
   ' 若日期中有/表示這是日期格式的字串如 YY/MM/DD 或 YYYY/MM/DD
   If InStr(1, strDate, "/") Then
      nBegin = InStr(1, strDate, "/")
      nEnd = InStr(nBegin + 1, strDate, "/")
      strTemp = Mid(strDate, nBegin + 1, nEnd - nBegin - 1)
      If Len(strTemp) = 1 Then
         DBMONTH = "0" & strTemp
      Else
         DBMONTH = strTemp
      End If
   Else
      If Len(strDate) > 7 Then
         DBMONTH = Mid(strDate, 5, 2)
      Else
         strTemp = strDate
         DBMONTH = Mid(ChangeTStringToWString(strTemp), 5, 2)
      End If
   End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得西元的日數
' Input : strDate ==> 輸入的日期
' Output : 傳回西元的日數 DD
' Description : 此功能會傳回日期字串中的年, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function DBDAY(ByVal strDate As String) As String
   Dim nTemp As Integer
   Dim nIndex As Integer
   Dim strTemp As String
   DBDAY = "00"
   ' 若日期中有/表示這是日期格式的字串如 YY/MM/DD 或 YYYY/MM/DD
   If InStr(1, strDate, "/") Then
      nTemp = InStr(1, strDate, "/")
      nIndex = InStr(nTemp + 1, strDate, "/")
      strTemp = Mid(strDate, nIndex + 1, Len(strDate) - nIndex)
      If Len(strTemp) = 1 Then
         DBDAY = "0" & strTemp
      Else
         DBDAY = strTemp
      End If
   Else
      If Len(strDate) > 7 Then
         DBDAY = Mid(strDate, 7, 2)
      Else
         strTemp = strDate
         DBDAY = Mid(ChangeTStringToWString(strTemp), 7, 2)
      End If
   End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得民國的年份
' Input : strDate ==> 輸入的日期
' Output : 傳回民國的年份 YY
' Description : 此功能會傳回日期字串中的月份, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TAIWANYEAR(ByVal strDate As String) As String
   Dim strTemp As String
   strTemp = DBYEAR(strDate)
   TAIWANYEAR = Val(strTemp) - 1911
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得民國的月份
' Input : strDate ==> 輸入的日期
' Output : 傳回民國的月份 MM
' Description : 此功能會傳回日期字串中的年, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TAIWANMONTH(ByVal strDate As String) As String
   TAIWANMONTH = DBMONTH(strDate)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得民國的日數
' Input : strDate ==> 輸入的日期
' Output : 傳回民國的日數 DD
' Description : 此功能會傳回日期字串中的日數, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function TAIWANDAY(ByVal strDate As String) As String
   TAIWANDAY = DBDAY(strDate)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 初始化
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OnInitial()
   m_NP01 = Empty
   m_NP02 = Empty
   m_NP03 = Empty
   m_NP04 = Empty
   m_NP05 = Empty
   m_NP07 = Empty
   m_NP10 = Empty
   m_NP10_Name = Empty
   m_NP10_Dep = Empty
   m_NP22 = Empty
   m_NP08 = Empty
   m_NP09 = Empty
   m_NP23 = Empty
   m_TM09 = Empty
   m_TM12 = Empty
   m_TM15 = Empty
   m_Customer = Empty
   m_Customer_CName = Empty
   m_Customer_EName = Empty
   m_Customer_Main = Empty
   m_Customer_Keep = Empty
   m_Customer_Adrs = Empty
   m_Customer_Tel = Empty
   m_Customer_Fax = Empty
   m_Customer_Cny = Empty
   m_Customer_CAddress = Empty
   m_Customer_EAddress = Empty
   m_TM05 = Empty
   m_Nation = Empty
   m_Agent = Empty
   'add by nickc 2008/01/18
   m_CuMemo = Empty
   m_SpecCu = Empty
   m_PA08 = Empty 'Add By Sindy 2009/06/29
End Sub

'Add By Cheng 2002/10/25
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 初始化
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OnInitialCP()
   m_CP01 = Empty
   m_CP02 = Empty
   m_CP03 = Empty
   m_CP04 = Empty
   m_CP10 = "601"
   m_CP12 = Empty
   m_CP12_Name = Empty
   m_CP12_Dep = Empty
   m_CP06 = Empty
   m_CP07 = Empty
   m_TM09 = Empty
   m_TM12 = Empty
   m_TM15 = Empty
   m_Customer = Empty
   m_Customer_CName = Empty
   m_Customer_EName = Empty
   m_Customer_Main = Empty
   m_Customer_Keep = Empty
   m_Customer_Adrs = Empty
   m_Customer_Tel = Empty
   m_Customer_Fax = Empty
   m_Customer_Cny = Empty
   m_Customer_CAddress = Empty
   m_Customer_EAddress = Empty
   m_TM05 = Empty
   m_Nation = Empty
   m_Agent = Empty
   'add by nickc 2008/01/18
   m_CuMemo = Empty
   m_SpecCu = Empty
End Sub

'Add By Cheng 2002/10/25
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 初始化
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Sub OnInitialCPLA()
   m_CP01 = Empty
   m_CP02 = Empty
   m_CP03 = Empty
   m_CP04 = Empty
   m_CP10 = "0"
   m_CP12 = Empty
   m_CP12_Name = Empty
   m_CP12_Dep = Empty
   m_CP06 = Empty
   m_CP07 = Empty
   m_TM09 = Empty
   m_TM12 = Empty
   m_TM15 = Empty
   m_Customer = Empty
   m_Customer_CName = Empty
   m_Customer_EName = Empty
   m_Customer_Main = Empty
   m_Customer_Keep = Empty
   m_Customer_Adrs = Empty
   m_Customer_Tel = Empty
   m_Customer_Fax = Empty
   m_Customer_Cny = Empty
   m_Customer_CAddress = Empty
   m_Customer_EAddress = Empty
   m_TM05 = Empty
   m_Nation = Empty
   m_Agent = Empty
   'add by nickc 2008/01/18
   m_CuMemo = Empty
   m_SpecCu = Empty
End Sub

' 取得客戶檔的相關資料
Private Sub QueryCustomer(ByVal strCustomer As String, Optional strCP01 As String, Optional strCP02 As String, Optional strCP03 As String, Optional strCP04 As String)
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   If Len(strCustomer) > 8 Then
      strSql = "SELECT * FROM Customer " & _
               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                     "CU02 = '" & Mid(strCustomer, 9, 1) & "'"
   Else
      strSql = "SELECT * FROM Customer " & _
               "WHERE CU01 = '" & Mid(strCustomer, 1, 8) & "' AND " & _
                     "CU02 = '0' "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 申請人中文名稱
      If IsNull(rsTmp.Fields("CU04")) = False Then
         m_Customer_CName = rsTmp.Fields("CU04")
      End If
      ' 申請人英文名稱
      If IsNull(rsTmp.Fields("CU05")) = False Then
         m_Customer_EName = rsTmp.Fields("CU05")
      End If
      If IsEmptyText(m_Customer_EName) = True Then
         If IsNull(rsTmp.Fields("CU88")) = False Then
            m_Customer_EName = rsTmp.Fields("CU88")
         End If
      End If
      If IsEmptyText(m_Customer_EName) = True Then
         If IsNull(rsTmp.Fields("CU89")) = False Then
            m_Customer_EName = rsTmp.Fields("CU89")
         End If
      End If
      If IsEmptyText(m_Customer_EName) = True Then
         If IsNull(rsTmp.Fields("CU90")) = False Then
            m_Customer_EName = rsTmp.Fields("CU90")
         End If
      End If
      ' 負責人
      If IsNull(rsTmp.Fields("CU07")) = False Then
         m_Customer_Main = rsTmp.Fields("CU07")
      End If
      'Modify by Morgan 2008/8/20
      '' 接洽人
      'If IsNull(rsTmp.Fields("CU08")) = False Then
      '   m_Customer_Keep = rsTmp.Fields("CU08")
      'End If
      '' 聯絡地址
      'If IsNull(rsTmp.Fields("CU31")) = False Then
      '   m_Customer_Adrs = rsTmp.Fields("CU31")
      'End If
      Call PUB_GetAddrRef(, m_NP02, m_NP03, m_NP04, m_NP05, , m_Customer_Keep, , m_Customer_Adrs, , , True)
      'end 2008/8/20
      ' 電話
      If IsNull(rsTmp.Fields("CU16")) = False Then
         m_Customer_Tel = rsTmp.Fields("CU16")
      End If
      ' 傳真
      If IsNull(rsTmp.Fields("CU18")) = False Then
         m_Customer_Fax = rsTmp.Fields("CU18")
      End If
      ' 申請人國籍
      If IsNull(rsTmp.Fields("CU10")) = False Then
         m_Customer_Cny = GetNationName(rsTmp.Fields("CU10"), 0)
      End If
      ' 申請地址(中)
      If IsNull(rsTmp.Fields("CU23")) = False Then
         m_Customer_CAddress = rsTmp.Fields("CU23")
      End If
      ' 申請地址(英)
      If IsNull(rsTmp.Fields("CU24")) = False Then
         m_Customer_EAddress = rsTmp.Fields("CU24")
      End If
      If IsEmptyText(m_Customer_EAddress) = True Then
         If IsNull(rsTmp.Fields("CU25")) = False Then
            m_Customer_EAddress = rsTmp.Fields("CU25")
         End If
      End If
      If IsEmptyText(m_Customer_EAddress) = True Then
         If IsNull(rsTmp.Fields("CU26")) = False Then
            m_Customer_EAddress = rsTmp.Fields("CU26")
         End If
      End If
      If IsEmptyText(m_Customer_EAddress) = True Then
         If IsNull(rsTmp.Fields("CU27")) = False Then
            m_Customer_EAddress = rsTmp.Fields("CU27")
         End If
      End If
      If IsEmptyText(m_Customer_EAddress) = True Then
         If IsNull(rsTmp.Fields("CU28")) = False Then
            m_Customer_EAddress = rsTmp.Fields("CU28")
         End If
      End If
      If IsEmptyText(m_Customer_EAddress) = True Then
         If IsNull(rsTmp.Fields("CU102")) = False Then
            m_Customer_EAddress = rsTmp.Fields("CU102")
         End If
      End If
   End If
   
   'Add By Sindy 2011/1/25 專利商標案的申請人中、英文地址改抓案件基本檔的申請地址
   Select Case strCP01
      '專利
      Case "P", "FCP", "CFP"
         If rsTmp.State <> 0 Then rsTmp.Close
         strSql = "Select * From Patent Where pa01='" & strCP01 & "' And pa02='" & strCP02 & "' And pa03='" & strCP03 & "' And pa04='" & strCP04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            m_Customer_CAddress = Trim("" & rsTmp("pa31").Value)
            m_Customer_EAddress = Trim("" & rsTmp("pa36").Value)
         End If
         
      '商標
      Case "T", "FCT", "CFT", "TF"
         If rsTmp.State <> 0 Then rsTmp.Close
         strSql = "Select * From Trademark Where tm01='" & strCP01 & "' And tm02='" & strCP02 & "' And tm03='" & strCP03 & "' And tm04='" & strCP04 & "' "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            m_Customer_CAddress = Trim("" & rsTmp("tm24").Value)
            m_Customer_EAddress = Trim("" & rsTmp("tm25").Value)
         End If
   End Select
   '2011/1/25 End
   
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取商標基本檔
Private Sub QueryTradeMark()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM TradeMark " & _
            "WHERE TM01 = '" & m_NP02 & "' AND " & _
                  "TM02 = '" & m_NP03 & "' AND " & _
                  "TM03 = '" & m_NP04 & "' AND " & _
                  "TM04 = '" & m_NP05 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 案件名稱
      If IsNull(rsTmp.Fields("TM05")) = False Then
         m_TM05 = rsTmp.Fields("TM05")
      End If
      If IsEmptyText(m_TM05) Then
         If IsNull(rsTmp.Fields("TM06")) = False Then
            m_TM05 = rsTmp.Fields("TM06")
         End If
      End If
      If IsEmptyText(m_TM05) Then
         If IsNull(rsTmp.Fields("TM07")) = False Then
            m_TM05 = rsTmp.Fields("TM07")
         End If
      End If
      ' 商品類別
      If IsNull(rsTmp.Fields("TM09")) = False Then
         m_TM09 = rsTmp.Fields("TM09")
      End If
      ' 國家代碼
      m_Nation = Empty
      If IsNull(rsTmp.Fields("TM10")) = False Then
         m_Nation = rsTmp.Fields("TM10")
      End If
      m_TM12 = Empty
      m_TM15 = Empty
      ' 審定號
      If IsNull(rsTmp.Fields("TM15")) = False Then
         m_TM15 = rsTmp.Fields("TM15")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("TM12")) = False Then
         m_TM12 = rsTmp.Fields("TM12")
      End If
      ' 申請人
      m_Customer = Empty
      If IsNull(rsTmp.Fields("TM23")) = False Then
         m_Customer = rsTmp.Fields("TM23")
      End If
      ' 申請地址
      If IsNull(rsTmp.Fields("TM24")) = False Then
         m_Customer_CAddress = rsTmp.Fields("TM24")
      End If
      ' 代理人
      m_Agent = Empty
      If IsNull(rsTmp.Fields("TM44")) = False Then
         m_Agent = rsTmp.Fields("TM44")
      End If
      'Add By Cheng 2003/06/23
      '彼所案號
      m_TM45 = Empty
      If IsNull(rsTmp.Fields("TM45")) = False Then
         m_TM45 = rsTmp.Fields("TM45")
      End If
      '2009/8/12 ADD BY SONIA
      If IsNull(rsTmp.Fields("TM78")) = False Then
         more_Customer = "Y"
      End If
      '2009/8/2 END
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

' 讀取服務業務基本檔
Private Sub QueryServicePractice()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM ServicePractice " & _
            "WHERE SP01 = '" & m_NP02 & "' AND " & _
                  "SP02 = '" & m_NP03 & "' AND " & _
                  "SP03 = '" & m_NP04 & "' AND " & _
                  "SP04 = '" & m_NP05 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 案件名稱0000000
      If IsNull(rsTmp.Fields("SP05")) = False Then
         m_TM05 = rsTmp.Fields("SP05")
      End If
      If IsEmptyText(m_TM05) Then
         If IsNull(rsTmp.Fields("SP06")) = False Then
            m_TM05 = rsTmp.Fields("SP06")
         End If
      End If
      If IsEmptyText(m_TM05) Then
         If IsNull(rsTmp.Fields("SP07")) = False Then
            m_TM05 = rsTmp.Fields("SP07")
         End If
      End If
      ' 國家代碼
      m_Nation = Empty
      If IsNull(rsTmp.Fields("SP09")) = False Then
         m_Nation = rsTmp.Fields("SP09")
      End If
      m_TM12 = Empty
      m_TM15 = Empty
      ' 申請案號
      If IsNull(rsTmp.Fields("SP11")) = False Then
         m_TM12 = rsTmp.Fields("SP11")
      End If
      ' 申請人
      m_Customer = Empty
      If IsNull(rsTmp.Fields("SP08")) = False Then
         m_Customer = rsTmp.Fields("SP08")
      End If
      '2005/11/2 ADD BY SONIA
      If IsNull(rsTmp.Fields("SP58")) = False Then
         more_Customer = "Y"
      End If
      '2005/11/2 END
      ' 代理人
      m_Agent = Empty
      If IsNull(rsTmp.Fields("SP26")) = False Then
         m_Agent = rsTmp.Fields("SP26")
      End If
      'Add By Cheng 2003/06/23
      '彼所案號
      m_TM45 = Empty
      If IsNull(rsTmp.Fields("SP27")) = False Then
         m_TM45 = rsTmp.Fields("SP27")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub QueryPatent()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & m_NP02 & "' AND " & _
                  "PA02 = '" & m_NP03 & "' AND " & _
                  "PA03 = '" & m_NP04 & "' AND " & _
                  "PA04 = '" & m_NP05 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 案件名稱0000000
      If IsNull(rsTmp.Fields("PA05")) = False Then
         m_TM05 = rsTmp.Fields("PA05")
      End If
      If IsEmptyText(m_TM05) Then
         If IsNull(rsTmp.Fields("PA06")) = False Then
            m_TM05 = rsTmp.Fields("PA06")
         End If
      End If
      If IsEmptyText(m_TM05) Then
         If IsNull(rsTmp.Fields("PA07")) = False Then
            m_TM05 = rsTmp.Fields("PA07")
         End If
      End If
      'Add By Sindy 2009/06/29
      ' 專利種類
      m_PA08 = Empty
      If IsNull(rsTmp.Fields("PA08")) = False Then
         m_PA08 = rsTmp.Fields("PA08")
      End If
      '2009/06/29 End
      ' 國家代碼
      m_Nation = Empty
      If IsNull(rsTmp.Fields("PA09")) = False Then
         m_Nation = rsTmp.Fields("PA09")
      End If
      m_TM12 = Empty
      m_TM15 = Empty
      ' 專利號數
      If IsNull(rsTmp.Fields("PA22")) = False Then
         m_TM15 = rsTmp.Fields("PA22")
      End If
      ' 申請案號
      If IsNull(rsTmp.Fields("PA11")) = False Then
         m_TM12 = rsTmp.Fields("PA11")
      End If
      ' 申請人
      m_Customer = Empty
      If IsNull(rsTmp.Fields("PA26")) = False Then
         m_Customer = rsTmp.Fields("PA26")
      End If
      '2005/11/2 ADD BY SONIA
      If IsNull(rsTmp.Fields("PA27")) = False Then
         more_Customer = "Y"
      End If
      '2005/11/2 END
      ' 代理人
      m_Agent = Empty
      If IsNull(rsTmp.Fields("PA75")) = False Then
         m_Agent = rsTmp.Fields("PA75")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing

End Sub

Private Sub QueryLawCase()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM LAWCASE " & _
            "WHERE LC01 = '" & m_NP02 & "' AND " & _
                  "LC02 = '" & m_NP03 & "' AND " & _
                  "LC03 = '" & m_NP04 & "' AND " & _
                  "LC04 = '" & m_NP05 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 案件名稱0000000
      If IsNull(rsTmp.Fields("LC05")) = False Then
         m_TM05 = rsTmp.Fields("LC05")
      End If
      If IsEmptyText(m_TM05) Then
         If IsNull(rsTmp.Fields("LC06")) = False Then
            m_TM05 = rsTmp.Fields("LC06")
         End If
      End If
      If IsEmptyText(m_TM05) Then
         If IsNull(rsTmp.Fields("LC07")) = False Then
            m_TM05 = rsTmp.Fields("LC07")
         End If
      End If
      ' 國家代碼
      m_Nation = Empty
      If IsNull(rsTmp.Fields("LC15")) = False Then
         m_Nation = rsTmp.Fields("LC15")
      End If
      m_TM12 = Empty
      m_TM15 = Empty
      ' 申請人
      m_Customer = Empty
      If IsNull(rsTmp.Fields("LC11")) = False Then
         m_Customer = rsTmp.Fields("LC11")
      End If
      ' 代理人
      m_Agent = Empty
      If IsNull(rsTmp.Fields("LC22")) = False Then
         m_Agent = rsTmp.Fields("LC22")
      End If
      '2011/2/21 Add By Sindy
      If IsNull(rsTmp.Fields("LC43")) = False Then
         more_Customer = "Y"
      End If
      '2011/2/21 End
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub

Private Sub QueryHireCase()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM HIRECASE " & _
            "WHERE HC01 = '" & m_NP02 & "' AND " & _
                  "HC02 = '" & m_NP03 & "' AND " & _
                  "HC03 = '" & m_NP04 & "' AND " & _
                  "HC04 = '" & m_NP05 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 案件名稱0000000
      If IsNull(rsTmp.Fields("HC06")) = False Then
         m_TM05 = rsTmp.Fields("HC06")
      End If
      ' 國家代碼
      m_Nation = Empty
      m_TM12 = Empty
      m_TM15 = Empty
      ' 申請人
      m_Customer = Empty
      If IsNull(rsTmp.Fields("HC05")) = False Then
         m_Customer = rsTmp.Fields("HC05")
      End If
      ' 代理人
      m_Agent = Empty
      '2011/2/21 Add By Sindy
      If IsNull(rsTmp.Fields("HC24")) = False Then
         more_Customer = "Y"
      End If
      '2011/2/21 End
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Sub
' 讀取下一程序檔
Private Sub QueryCaseProgress1(ByRef p_CP09 As String)
   strSql = "select * from caseprogress where cp09='" & p_CP09 & "'"
   If RsTemp.State = adStateOpen Then RsTemp.Close
   With RsTemp
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         ' 總收文號
         m_NP01 = "" & .Fields("CP43")
         ' 本所案號
         m_NP02 = "" & .Fields("CP01")
         m_NP03 = "" & .Fields("CP02")
         m_NP04 = "" & .Fields("CP03")
         m_NP05 = "" & .Fields("CP04")
         '收文日
         m_CP05 = "" & .Fields("CP05")
         '案件性質
         m_NP07 = "" & .Fields("CP10")
         ' 本所期限
         m_NP08 = "" & .Fields("CP06")
         ' 法定期限
         m_NP09 = "" & .Fields("CP07")
         'Modify by Morgan 2009/12/28 非FMP案才要
         'If m_NP02 = "P" And m_Nation <> 台灣國家代號 Then
         If m_NP02 = "P" And m_Nation <> 台灣國家代號 And Left("" & .Fields("cp12"), 1) <> "F" Then
            m_NP09 = m_NP08
         End If
         ' 智權人員代號
         m_NP10 = "" & .Fields("CP13")
         ' 取得智權人員姓名
         m_NP10_Name = GetStaffName(m_NP10, True)
         '收文業務區
         m_NP10_Dep = "" & .Fields("CP12")
         ' 本所期限
         If m_NP08 <> "" Then
            m_NP08 = ChangeWStringToTDateString(m_NP08)
         End If
         ' 法定期限
         If m_NP09 <> "" Then
            m_NP09 = ChangeWStringToTDateString(m_NP09)
         End If
         '收文日
         If m_CP05 <> "" Then
            m_CP05 = ChangeWStringToTDateString(m_CP05)
         End If
      End If
   End With
End Sub
' 讀取下一程序檔
Private Sub QueryNextProgress()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strSql = "SELECT * FROM NextProgress,caseprogress " & _
            "WHERE NP22 = " & m_NP22 & " and cp09(+)=np01 "
   If IsEmptyText(m_NP02) = False Then
      strSql = strSql & "AND NP02 = '" & m_NP02 & "' "
   End If
   If IsEmptyText(m_NP03) = False Then
      strSql = strSql & "AND NP03 = '" & m_NP03 & "' "
   End If
   If IsEmptyText(m_NP04) = False Then
      strSql = strSql & "AND NP04 = '" & m_NP04 & "' "
   End If
   If IsEmptyText(m_NP05) = False Then
      strSql = strSql & "AND NP05 = '" & m_NP05 & "' "
   End If
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 總收文號
      If IsNull(rsTmp.Fields("NP01")) = False Then
         m_NP01 = rsTmp.Fields("NP01")
      End If
      ' 本所案號
      If IsEmptyText(m_NP02) = True Then
         If IsNull(rsTmp.Fields("NP02")) = False Then
            m_NP02 = rsTmp.Fields("NP02")
         End If
      End If
      If IsEmptyText(m_NP03) = True Then
         If IsNull(rsTmp.Fields("NP03")) = False Then
            m_NP03 = rsTmp.Fields("NP03")
         End If
      End If
      If IsEmptyText(m_NP04) = True Then
         If IsNull(rsTmp.Fields("NP04")) = False Then
            m_NP04 = rsTmp.Fields("NP04")
         End If
      End If
      If IsEmptyText(m_NP05) = True Then
         If IsNull(rsTmp.Fields("NP05")) = False Then
            m_NP05 = rsTmp.Fields("NP05")
         End If
      End If
      ' 案件性質
      'If IsNull(rsTmp.Fields("NP07")) = False Then
      '   m_NP07 = rsTmp.Fields("NP07")
      'End If
      '910628 Sieg
      If bol901 Then
         m_NP07 = "901"
      Else
         If IsNull(rsTmp.Fields("NP07")) = False Then
            m_NP07 = rsTmp.Fields("NP07")
         End If
      End If
      ' 本所期限
      If IsNull(rsTmp.Fields("NP08")) = False Then
         m_NP08 = rsTmp.Fields("NP08")
      End If
      ' 法定期限
      If IsNull(rsTmp.Fields("NP09")) = False Then
         m_NP09 = rsTmp.Fields("NP09")
      End If
      
      ' 約定期限 Add by Morgan 2009/12/7
      If IsNull(rsTmp.Fields("NP23")) = False Then
         m_NP23 = rsTmp.Fields("NP23")
      End If
      
      'MODIFY BY SONIA
      'Modify by Morgan 2009/12/28 FMP除外
      'If m_NP02 = "P" And m_Nation <> 台灣國家代號 Then
      If m_NP02 = "P" And m_Nation <> 台灣國家代號 And Left("" & rsTmp.Fields("CP12"), 1) <> "F" Then
         If IsNull(rsTmp.Fields("NP08")) = False Then
            m_NP09 = rsTmp.Fields("NP08")
         End If
      End If
      
      ' 智權人員代號
      If IsNull(rsTmp.Fields("NP10")) = False Then
         m_NP10 = rsTmp.Fields("NP10")
      End If
   End If
   rsTmp.Close
   'Add by Morgan 2004/12/3 抓最新A類收文的智權人員
   '93.12.29 MODIFY BY SONIA
   'm_NP10 = PUB_GetAKindSalesNo(m_NP02, m_NP03, m_NP04, m_NP05)
   Select Case m_NP02
      Case "FCP"
         m_NP10 = PUB_GetFCPSalesNo(m_NP02, m_NP03, m_NP04, m_NP05)
      Case "FCT"
         'Modified by Lydia 2018/01/10 因為從107/1/1開始FCT承辦人預設為分區主管,所以智權人員改抓NP10
         'm_NP10 = PUB_GetFCTSalesNo(m_NP02, m_NP03, m_NP04, m_NP05)
         If m_NP10 = "" Then
             m_NP10 = PUB_GetFCTSalesNo(m_NP02, m_NP03, m_NP04, m_NP05)
         End If
         'end 2018/01/10
      Case Else
         m_NP10 = PUB_GetAKindSalesNo(m_NP02, m_NP03, m_NP04, m_NP05)
   End Select
   '93.12.29 END
   
   ' 取得智權人員姓名
   'Modify By Cheng 2002/04/15
   '不論是否離職都要帶出智權人員姓名
'   m_NP10_Name = GetStaffName(m_NP10)
   m_NP10_Name = GetStaffName(m_NP10, True)
   ' 取得智權人員部門別
   m_NP10_Dep = GetStaffDepartment(m_NP10)
   ' 本所期限
   m_NP08 = TAIWANYEAR(m_NP08) & "/" & TAIWANMONTH(m_NP08) & "/" & TAIWANDAY(m_NP08)
   ' 法定期限
   m_NP09 = TAIWANYEAR(m_NP09) & "/" & TAIWANMONTH(m_NP09) & "/" & TAIWANDAY(m_NP09)
   ' 約定期限 Add by Morgan 2009/12/7
   If m_NP23 <> "" Then
      m_NP23 = TAIWANYEAR(m_NP23) & "/" & TAIWANMONTH(m_NP23) & "/" & TAIWANDAY(m_NP23)
   End If
   
   Set rsTmp = Nothing
End Sub

' 讀取案件進度檔
Private Sub QueryCaseProgress()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
    'Modify By Cheng 2002/11/26
    '抓A類收文號資料以收文日及收文號由大至小排序
'   strSQL = "SELECT * FROM CaseProgress " & _
'            "WHERE  " & ChgCaseprogress(m_CP01 & m_CP02 & m_CP03 & m_CP04) & " AND CP10='" & m_CP10 & "'  AND CP27 IS NULL AND CP57 IS NULL "
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE  " & ChgCaseprogress(m_CP01 & m_CP02 & m_CP03 & m_CP04) & " AND CP09 < 'B' Order By CP05 DESC, CP09 DESC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 總收文號
'        m_CP09 = "" & rsTmp.Fields("CP09")
      ' 本所案號
      ' 案件性質
      ' 本所期限
'        m_CP06 = "" & rsTmp.Fields("CP06")
      ' 法定期限
'         m_CP07 = "" & rsTmp.Fields("CP07")
      ' 智權人員代號
         m_CP12 = "" & rsTmp.Fields("CP13")
   End If
   If rsTmp.State <> adStateClosed Then rsTmp.Close
   Set rsTmp = Nothing
   
   ' 取得智權人員姓名
   'Modify By Cheng 2002/04/15
   '不論是否離職都要帶出智權人員姓名
   m_CP12_Name = GetStaffName(m_CP12, True)
   ' 取得智權人員部門別
   m_CP12_Dep = GetStaffDepartment(m_CP12)
   ' 本所期限
   m_CP06 = TAIWANYEAR(m_CP06) & "/" & TAIWANMONTH(m_CP06) & "/" & TAIWANDAY(m_CP06)
   ' 法定期限
   m_CP07 = TAIWANYEAR(m_CP07) & "/" & TAIWANMONTH(m_CP07) & "/" & TAIWANDAY(m_CP07)
   
End Sub

'Add By Cheng 2003/03/31
' 讀取案件進度檔
Private Sub QueryCaseProgressLA()
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
    'Modify By Cheng 2002/11/26
    '抓A, B類收文號資料以收文日及收文號由大至小排序
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE  " & ChgCaseprogress(m_CP01 & m_CP02 & m_CP03 & m_CP04) & " AND CP09 < 'C' And CP10='0' Order By CP05 DESC, CP09 DESC "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 總收文號
'        m_CP09 = "" & rsTmp.Fields("CP09")
      ' 本所案號
      ' 案件性質
'        m_CP10 = "" & rsTmp.Fields("CP10")
       ' 本所期限
        m_CP06 = "" & rsTmp.Fields("CP54")
      ' 法定期限
         m_CP07 = "" & rsTmp.Fields("CP54")
      ' 智權人員代號
         m_CP12 = "" & rsTmp.Fields("CP13")
   End If
   If rsTmp.State <> adStateClosed Then rsTmp.Close
   Set rsTmp = Nothing
   
   ' 取得智權人員姓名
   'Modify By Cheng 2002/04/15
   '不論是否離職都要帶出智權人員姓名
   m_CP12_Name = GetStaffName(m_CP12, True)
   ' 取得智權人員部門別
   m_CP12_Dep = GetStaffDepartment(m_CP12)
   ' 本所期限
   m_CP06 = TAIWANYEAR(m_CP06) & "/" & TAIWANMONTH(m_CP06) & "/" & TAIWANDAY(m_CP06)
   ' 法定期限
   m_CP07 = TAIWANYEAR(m_CP07) & "/" & TAIWANMONTH(m_CP07) & "/" & TAIWANDAY(m_CP07)
   
End Sub

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得員工的部門別
' Input : strStuff ==> 員工的代碼
' Output : 傳回員工所屬的部門別
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Private Function GetStaffDepartment(ByVal strStuff As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetStaffDepartment = Empty
   strSql = "SELECT * FROM Staff " & _
            "WHERE ST01 = '" & strStuff & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("ST03")) = False Then
         GetStaffDepartment = rsTmp.Fields("ST03")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

' 讀取資料
Private Sub QueryData(Optional ByRef p_CP09 As String)
   '2005/11/2 ADD BY SONIA
   more_Customer = ""
   'Add by Morgan 2006/12/21
   m_CP05 = ""
   'Modify by Morgan 2006/6/26
   If p_CP09 = "" Then
      ' 讀取下一程序檔
      QueryNextProgress
   Else
      QueryCaseProgress1 p_CP09
   End If
   ' 依系統別
   Select Case m_NP02
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         QueryPatent
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         QueryTradeMark
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         QueryLawCase
      ' 讀取顧問案件基本檔
      Case "LA":
         QueryHireCase
      ' 讀取服務業務基本檔
      Case Else:
         QueryServicePractice
   End Select
   
   ' 取客戶檔的相關資料
   If IsEmptyText(m_Customer) = False Then
      Call QueryCustomer(m_Customer, m_NP02, m_NP03, m_NP04, m_NP05)
   End If
End Sub

'Add By Cheng 2002/10/25
' 讀取資料
Private Sub QueryDataCP()
   ' 讀取案件進度檔
   'Modify by Morgan 2004/7/9
   '智權人員用新規則抓
   'QueryCaseProgress
   ' 智權人員代號
   m_CP12 = PUB_GetAKindSalesNo(m_CP01, m_CP02, m_CP03, m_CP04)
   ' 取得智權人員姓名
   'Modify By Cheng 2002/04/15
   '不論是否離職都要帶出智權人員姓名
   m_CP12_Name = GetStaffName(m_CP12, True)
   ' 取得智權人員部門別
   m_CP12_Dep = GetStaffDepartment(m_CP12)
   ' 本所期限
   m_CP06 = TAIWANYEAR(m_CP06) & "/" & TAIWANMONTH(m_CP06) & "/" & TAIWANDAY(m_CP06)
   ' 法定期限
   m_CP07 = TAIWANYEAR(m_CP07) & "/" & TAIWANMONTH(m_CP07) & "/" & TAIWANDAY(m_CP07)
   'End
   
   
   'Add By Cheng 2002/11/26
    m_CP06 = ""
    m_CP07 = ""
    m_NP02 = m_CP01
    m_NP03 = m_CP02
    m_NP04 = m_CP03
    m_NP05 = m_CP04
   ' 依系統別
   Select Case m_CP01
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         QueryPatent
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         QueryTradeMark
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         QueryLawCase
      ' 讀取顧問案件基本檔
      Case "LA":
         QueryHireCase
      ' 讀取服務業務基本檔
      Case Else:
         QueryServicePractice
   End Select
   
   ' 取客戶檔的相關資料
   If IsEmptyText(m_Customer) = False Then
      Call QueryCustomer(m_Customer, m_CP01, m_CP02, m_CP03, m_CP04)
   End If
End Sub

'Add By Cheng 2003/03/31
' 讀取資料
Private Sub QueryDataCPLA()
   ' 讀取案件進度檔
   QueryCaseProgressLA
    m_NP02 = m_CP01
    m_NP03 = m_CP02
    m_NP04 = m_CP03
    m_NP05 = m_CP04
   ' 依系統別
   Select Case m_CP01
      ' 讀取專利基本檔
      Case "P", "CFP", "FCP":
         QueryPatent
      ' 讀取商標基本檔
      Case "T", "TF", "CFT", "FCT":
         QueryTradeMark
      ' 讀取法務基本檔
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS":
         QueryLawCase
      ' 讀取顧問案件基本檔
      Case "LA":
         QueryHireCase
      ' 讀取服務業務基本檔
      Case Else:
         QueryServicePractice
   End Select
   
   ' 取客戶檔的相關資料
   If IsEmptyText(m_Customer) = False Then
      Call QueryCustomer(m_Customer, m_CP01, m_CP02, m_CP03, m_CP04)
   End If
End Sub

' 列印
'Modify By Cheng 2003/04/16
'Public Sub PrintForm(ByVal strNP22 As String, Optional ByVal strNP02 As String = Empty, Optional ByVal strNP03 As String = Empty, Optional ByVal strNP04 As String = Empty, Optional ByVal strNP05 As String = Empty)
'Modify by Morgan 2005/5/18 加strSpecial=特別處理 1:加印年費減免案件性質
'Modify by Morgan 2009/5/27 +bContinue:是否繼續列印
'Modify by Sindy 2010/5/01 +bNewFunc:使用於結案單
'Modify by Sindy 2010/5/12 +strNP07:案件性質
'Modify By Sindy 2012/6/19 +bolFullPrint:是否完整列印接洽單
'Modify By Sindy 2015/5/1 +bolHadEFile:是否有電子回覆單
'Modify by Amy 2019/06/04 +stFormName:表單名稱
Public Sub PrintForm(ByVal strNP22 As String, Optional ByVal strNP02 As String = Empty, Optional ByVal strNP03 As String = Empty, _
                     Optional ByVal strNP04 As String = Empty, Optional ByVal strNP05 As String = Empty, Optional ByVal strBKindCP09 As String = Empty, _
                     Optional ByVal strSpecial As String = Empty, Optional ByVal bContinue As Boolean, Optional bNewFunc As Boolean = False, _
                     Optional strNP07 As String = "", Optional bolFullPrint As Boolean = True, Optional bolHadEFile As Boolean = False, Optional stFormName As String = "")
   'Add by Amy 2019/06/04
   Dim bolShowMsg As Boolean
   
   'Add by Amy 2019/06/04
   bolShowMsg = False
   If UCase(stFormName) = "FRM210154" Then
        bolShowMsg = True
   End If
   'end 2019/06/04
   OnInitial
   m_NP22 = strNP22
   If IsEmptyText(strNP02) = False And IsEmptyText(strNP03) = False Then
       m_NP02 = strNP02
       m_NP03 = strNP03
       m_NP04 = strNP04
       If IsEmptyText(m_NP04) = True Then: m_NP04 = "0"
       m_NP05 = strNP05
       If IsEmptyText(m_NP05) = True Then: m_NP05 = "00"
   End If
   'Modify by Morgan 2006/6/26
   'QueryData
   QueryData strBKindCP09
   
   'Modify By Cheng 2003/04/16
   '    ExportData
   'Modify By Cheng 2003/06/23
   
   'Add by Sindy 2013/7/16
   If m_NP22 = "000000" Then
       strNP07 = m_NP22
   End If
   '2013/7/16 END
   
   Select Case strNP02
   Case "FCT"
      'Modify By Sindy 2012/6/19 +bolFullPrint
      'Modify By Sindy 2015/5/1 +bolHadeFile
      'Modify by Amy 2019/06/04 +bolShowMsg
      ExportDataFCT strBKindCP09, , bNewFunc, strNP07, bolFullPrint, bolHadEFile, bolShowMsg
   Case Else
      'Modify By Sindy 2015/5/1 +bolHadeFile
      'Modify by Amy 2019/06/04 +bolShowMsg
      ExportData strBKindCP09, strSpecial, , bNewFunc, strNP07, bolHadEFile, bolShowMsg
   End Select
   OnInitial
End Sub

'Add By Cheng 2002/10/25
'若無下一程序檔列印接洽記錄單
Public Sub PrintFormCP(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, StrPA11 As String)
   OnInitialCP
   m_CP01 = strCP01
   m_CP02 = strCP02
   m_CP03 = strCP03
   If IsEmptyText(m_CP03) = True Then: m_CP03 = "0"
   m_CP04 = strCP04
   If IsEmptyText(m_CP04) = True Then: m_CP04 = "00"
   '92.5.8 ADD BY SONIA
   'Modify by Morgan 2010/12/28 申請案號改碼數
   'If Len(StrPA11) > 8 Then
      'If Mid(StrPA11, 3, 1) <> "3" Then
   If Len(StrPA11) > 9 Then
      If Mid(StrPA11, 4, 1) <> "3" Then
         m_CP10 = "602"
      Else
         m_CP10 = "603"
      End If
   End If
   '92.5.8 END
   QueryDataCP
   ExportDataCP
   OnInitialCP
End Sub

'Add By Cheng 2003/03/31
'若無下一程序檔列印接洽記錄單
'Modify by Morgan 2009/5/27 +bContinue:是否繼續列印
Public Sub PrintFormCPLA(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, Optional bContinue As Boolean, Optional bNewFunc As Boolean = False)
   OnInitialCPLA
    m_CP01 = strCP01
    m_CP02 = strCP02
    m_CP03 = strCP03
    If IsEmptyText(m_CP03) = True Then: m_CP03 = "0"
    m_CP04 = strCP04
    If IsEmptyText(m_CP04) = True Then: m_CP04 = "00"
   QueryDataCPLA
   ExportDataCP bContinue, bNewFunc
   OnInitialCPLA
End Sub

'Modify By Cheng 2003/04/16
'Private Sub ExportData()
'Modify by Morgan 2005/5/18 加strSpecial=特別處理 1:加印年費減免案件性質
'Modify by Morgan 2009/5/27 +bContinue:是否繼續列印
'Modify by Sindy 2010/5/01 +bNewFunc:使用於結案單
'Modify by Sindy 2010/5/12 +strNP07:案件性質
'Modify By Sindy 2015/5/1 +bolHadEFile:是否有電子回覆單
'Modify by Amy 2019/06/04 +bolShowMsg:列印完秀訊息
Private Sub ExportData(Optional strBKindCP09 As String = Empty, Optional strSpecial As String = Empty, Optional bContinue As Boolean, _
                       Optional bNewFunc As Boolean = False, Optional strNP07 As String = "", Optional bolHadEFile As Boolean = False, Optional bolShowMsg As Boolean = False)
Dim m_CharWidth As Long
Dim m_CharHeight As Long
Dim strData As String
Dim nCount As Integer
'92.6.8 ADD BY SONIA
Dim m_MORE As String
Dim strSQL1 As String
Dim Rs2 As New ADODB.Recordset
' 報表左方留白的寬度
Dim m_LeftMargin As Integer
' 報表上方留白的高度
Dim m_TopMargin As Integer
Dim nRow As Integer
'減免身分
Dim stCP81 As String
'Add By Sindy 2009/06/29
Dim strYear As String '抓下次繳費年度
Dim m_Nexttimes As String '抓下次繳費次數
Dim strFeeType As String, strYF15 As String
Dim strKey(0 To 5) As String
Dim m_CaseFee(1 To 2) As String
Dim aryCaseFee As Variant
'2009/06/29 End
Dim m_iFixNo As Integer '2010/2/12 ADD BY SONIA 修法次數
Dim strNP15 As String 'Add By Sindy 2017/4/25
   
   If strSrvDate(1) >= 20100501 And bNewFunc = False Then Exit Sub '2010/4/23 ADD BY SONIA 99/5/1起不再列印
    
   m_CharWidth = 120
   m_CharHeight = 280
    'Modify By Cheng 2003/05/29
'   m_LeftMargin = 6
   m_LeftMargin = 4
   m_TopMargin = 4
   If bNewFunc = False Then
      '***********  90.11.14   nickc
      Printer.Orientation = 1
      '*****************************
      Printer.PaperSize = vbPRPSA4
   End If
   Printer.Font.Size = 12
   Printer.FontName = "細明體" 'Add by Morgan 2004/9/10

   nRow = 1
   ' 總收文號
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
    'Modify By Cheng 2003/04/16
'   Printer.Print "總收文號："
   Printer.Print "總收文號：" & IIf(strBKindCP09 <> "", strBKindCP09, "")
   ' 區別
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Select Case m_NP02
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "FCP", "FCT", "FCL", "LIN", "ACS":
         Printer.Print "申請人國籍：" & m_Customer_Cny
      Case Else:
         Printer.Print "區別　　：" & m_NP10_Dep
   End Select
   ' 填表
   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Add By Sindy 2010/4/28
   If bNewFunc = True Then
      Printer.Print "填表日期：" & ChangeTStringToTDateString(strSrvDate(2))
   '2010/4/28 End
   Else
      Printer.Print "填表日期："
   End If
   
   nRow = nRow + 1
   ' 收文日期
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "收文日期：" & m_CP05
   'Add By Sindy 2013/7/16
   '無期限閉卷
   If strNP07 = "000000" Then
      ' 法定期限
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "法定期限："
   Else
   '2013/7/16 END
      ' 法定期限
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "法定期限：" & m_NP09
   End If
   ' 本所人員
   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify By Sindy 2011/7/11
   If bNewFunc = True And m_NP10_Name <> strUserName Then
      Printer.Print "本所人員：" & "★" & strUserName
   '2011/7/11 End
   Else
      Printer.Print "本所人員：" & m_NP10_Name
   End If

   'Modify by Morgan 2007/1/12 本所案號，案件性質向上移一行
   nRow = nRow + 1
   ' 本所案號
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "本所案號：" & m_NP02 & "-" & m_NP03 & "-" & m_NP04 & "-" & m_NP05
   'Add By Sindy 2013/7/16
   '無期限閉卷
   If strNP07 = "000000" Then
      ' 約定期限
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "約定期限："
   Else
   '2013/7/16 END
      ' 約定期限
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "約定期限：" & m_NP23
   End If
   ' 員工編號
   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify By Sindy 2011/7/11
   If bNewFunc = True And m_NP10_Name <> strUserName Then
      Printer.Print "員工編號：" & strUserNum
   '2011/7/11 End
   Else
      Printer.Print "員工編號：" & m_NP10
   End If
   
   '92.6.8 ADD BY SONIA 美國發明領證案檢查有無公開費期限
   m_MORE = ""
   If m_NP02 = "CFP" And m_NP07 = "601" Then
        strSQL1 = "SELECT NP08 FROM NEXTPROGRESS WHERE " & ChgNextProgress(m_NP02 & m_NP03 & m_NP04 & m_NP05) & " And NP06 Is Null And NP07 = 217 And NP09 = " & ChangeTStringToWString(ChangeTDateStringToTString(m_NP09))
        If Rs2.State <> adStateClosed Then Rs2.Close
        Set Rs2 = Nothing
        Rs2.CursorLocation = adUseClient
        Rs2.Open strSQL1, cnnConnection, adOpenStatic, adLockReadOnly
        If Rs2.RecordCount > 0 Then
           m_MORE = "Y"
        End If
   'Add By Sindy 2017/4/25
   ElseIf m_NP02 = "CFT" And m_NP07 = "601" Then
        strSQL1 = "SELECT NP15 FROM NEXTPROGRESS WHERE " & ChgNextProgress(m_NP02 & m_NP03 & m_NP04 & m_NP05) & " And NP01='" & m_NP01 & "'"
        If Rs2.State <> adStateClosed Then Rs2.Close
        Set Rs2 = Nothing
        Rs2.CursorLocation = adUseClient
        Rs2.Open strSQL1, cnnConnection, adOpenStatic, adLockReadOnly
        If Rs2.RecordCount > 0 Then
           strNP15 = "" & Rs2.Fields("NP15")
        End If
   '2017/4/25 END
   End If
   '92.6.8 END

   nRow = nRow + 1
   
   '案件性質
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Add By Sindy 2013/7/16
   '無期限閉卷
   If strNP07 = "000000" Then
      Printer.Print "案件性質：" & "★" & "無期限閉卷"
   '2013/7/16 END
   Else
      'Modify by Morgan 2007/1/12 重新整理各種例外情形
      'Modify By Sindy 2010/5/10
      If bNewFunc = True Then
         'Modify By Sindy 2010/5/12
         If strNP07 <> "" Then
            '台灣
            If m_Nation < "010" Then
               Printer.Print "案件性質：" & strNP07 & " " & GetCaseTypeName(m_NP02, strNP07, 0)
            Else
               Printer.Print "案件性質：" & strNP07 & " " & GetCaseTypeName(m_NP02, strNP07, 1)
            End If
         '2010/5/12 End
         Else
            '台灣
            If m_Nation < "010" Then
               Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0)
            Else
               Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 1)
            End If
         End If
      '2010/5/10 End
      Else
         '台灣
         If m_Nation < "010" Then
            '台灣的行政訴訟(T-403,P-503)要加印準備程序(T-204,P-211),言詞辯論(T-205,P-212)
            If ((m_NP02 = "T" And m_NP07 = "403") Or (m_NP02 = "P" And m_NP07 = "503")) Then
               Printer.Print "案件性質：□ " & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0) & ""
               If (m_NP02 = "T") Then
                  Printer.CurrentX = m_LeftMargin * m_CharWidth
                  Printer.CurrentY = (m_TopMargin + nRow + 1) * m_CharHeight
                  Printer.Print "　　　　　□ 204 " & GetCaseTypeName(m_NP02, "204", 0) & ""
                  Printer.CurrentX = m_LeftMargin * m_CharWidth
                  Printer.CurrentY = (m_TopMargin + nRow + 2) * m_CharHeight
                  Printer.Print "　　　　　□ 205 " & GetCaseTypeName(m_NP02, "205", 0) & ""
               Else
                  Printer.CurrentX = m_LeftMargin * m_CharWidth
                  Printer.CurrentY = (m_TopMargin + nRow + 1) * m_CharHeight
                  Printer.Print "　　　　　□ 211 " & GetCaseTypeName(m_NP02, "211", 0) & ""
                  Printer.CurrentX = m_LeftMargin * m_CharWidth
                  Printer.CurrentY = (m_TopMargin + nRow + 2) * m_CharHeight
                  Printer.Print "　　　　　□ 212 " & GetCaseTypeName(m_NP02, "212", 0) & ""
               End If
            '台灣的第一期註冊費(T-715)要加印全期註冊費(T-717)
            ElseIf (m_NP02 = "T" And m_NP07 = "715") Then
               Printer.Print "案件性質：□ " & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0) & ""
               Printer.CurrentX = m_LeftMargin * m_CharWidth
               Printer.CurrentY = (m_TopMargin + nRow + 1) * m_CharHeight
               Printer.Print "　　　　　□ 717 " & GetCaseTypeName(m_NP02, "717", 0) & ""
            '台灣的領証繳年費(601)要加延緩年費(412)
            ElseIf (m_NP02 = "P" And m_NP07 = "601") Then
               Printer.Print "案件性質：□ " & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0) & ""
               Printer.CurrentX = m_LeftMargin * m_CharWidth
               Printer.CurrentY = (m_TopMargin + nRow + 1) * m_CharHeight
               Printer.Print "　　　　　□ 412 " & GetCaseTypeName(m_NP02, "412", 0) & ""
               Printer.Print "                     ____個月"
            '台灣的申復(P-205)要加印修正(P-204)
            ElseIf (m_NP02 = "P" And m_NP07 = "205") Then
               'Modify by Morgan 2007/2/13 加控制來函的案件性質為1202審查意見通知函的才要
               If GetCP10(m_NP01) = "1202" Then
                  Printer.Print "案件性質：□ 204 " & GetCaseTypeName(m_NP02, "204", 0) & ""
                  Printer.CurrentX = m_LeftMargin * m_CharWidth
                  Printer.CurrentY = (m_TopMargin + nRow + 1) * m_CharHeight
                  Printer.Print "　　　　　□ " & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0) & ""
               Else
                  Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0)
               End If
               'End 2007/2/13
            '減免退費
            ElseIf strSpecial = "1" Then
               Printer.Print "案件性質：" & m_NP07 & GetCaseTypeName(m_NP02, m_NP07, 0) & ",919減免退費"
            '一般
            Else
               Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0)
            End If
         Else
            '2005/12/20 ADD BY SONIA
            If m_NP02 = "CFT" And m_NP07 = "312" Then
               Printer.Print "案件性質：303 " & GetCaseTypeName(m_NP02, "303", 1)
            '2005/12/20 END
            'CFP有公開費期限
            ElseIf m_NP02 = "CFP" And m_MORE = "Y" Then
               Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 1) & ",217 公開費"
            'Add By Sindy 2009/06/12  2009/7/1加限制美國
            ElseIf m_NP02 = "CFP" And GetCP10(m_NP01) = "1006" And m_Nation = "101" Then
               Printer.Print "案件性質：□ " & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 1)
               Printer.CurrentX = m_LeftMargin * m_CharWidth
               Printer.CurrentY = (m_TopMargin + nRow + 1) * m_CharHeight
               Printer.Print "　　　　　□ " & "424" & " " & GetCaseTypeName(m_NP02, "424", 0) & ""
            '一般
            Else
               Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 1)
            End If
         End If
      End If
   End If
   
   'Add By Sindy 2013/7/16
   '無期限閉卷
   If strNP07 = "000000" Then
      ' 本所期限
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "本所期限："
   Else
   '2013/7/16 END
      ' 本所期限
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "本所期限：" & m_NP08
   End If
   ' 申請國家
   If IsEmptyText(m_Nation) = False Then
      Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "申請國家：" & GetNationName(m_Nation, 0)
   End If
   
   nRow = nRow + 1
   'Add By Sindy 2009/06/29
   '下次繳費年度(次數)
   If (m_NP02 = "P" Or m_NP02 = "CFP") And _
      (m_NP07 = "605" Or m_NP07 = "606" Or m_NP07 = "607") Then
      '設定本所案號
      strKey(0) = ""
      strKey(1) = m_NP02
      strKey(2) = m_NP03
      strKey(3) = m_NP04
      strKey(4) = m_NP05
      '取得繳年費的資料
      '2010/2/12 MODIFY BY SONIA 抓修法次數
      'If GetMoneyDate(m_PA08, m_Nation, strKey, m_CaseFee(1), m_CaseFee(2)) = True Then
      If GetMoneyDate(m_PA08, m_Nation, strKey, m_CaseFee(1), m_CaseFee(2), , , m_iFixNo) = True Then
         '取得下次繳費次數/年度
         m_Nexttimes = PUB_Getnexttimes(m_NP02, m_NP03, m_NP04, m_NP05, strYear)
         If m_Nexttimes <> "" Then
            Printer.CurrentX = m_LeftMargin * m_CharWidth
            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
            If m_NP07 = "605" Then '605.年費
               'Printer.Print "繳費年度：第" & strYear & "年至第___年"
               aryCaseFee = Split(m_CaseFee(2), ",")
               If Val(strYear) < Val(aryCaseFee(UBound(aryCaseFee))) Then
                  Printer.Print "年度：第" & strYear & "年至第___年"
               Else
                  '2009/11/2 加註BY SONIA,波蘭最後一年為第9-10年,接洽結案單不可印第9-10年,否則收文會錯
                  Printer.Print "年度：第" & strYear & "年"
               End If
            ElseIf m_NP07 = "606" Or m_NP07 = "607" Then '606.維持費 607.延展費
               '年度說明
               'Memo by Morgan 2022/6/13 結案單已電子化，此處應該不會再被執行到
               strFeeType = PUB_GetNa20Na22Na24(m_Nation, m_PA08)
               'strYF15 = PUB_GetYF15(m_Nation, m_PA08, "Y0000000", strFeeType, CDbl(m_Nexttimes))
               '2010/2/12 MODIFY BY SONIA
               'strYF15 = PUB_GetYF15(m_Nation, m_PA08, "Y0000000", strFeeType, CDbl(strYear))
               strYF15 = PUB_GetYF15(m_Nation, m_PA08, "Y000000" & m_iFixNo, strFeeType, CDbl(strYear))
               'Printer.Print "繳費次數：第" & m_Nexttimes & "次(" & strYF15 & ")"
               Printer.Print "次數：" & strYF15
            End If
         End If
      End If
   End If
   '2009/06/29 End
   
   ' 審定號, 專利號數或申請案號
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   
   'Modify By Cheng 2002/01/04
'   Printer.Print m_TM12
   '91.12.5 MODIFY BY SONIA
   'Printer.Print "申請案號：" & m_TM12
   Select Case m_NP02
      Case "CFT", "FCT", "T", "TF":
         If IsNull(m_TM15) = False Then
            Printer.Print "審定號　：" & m_TM15
         Else
            Printer.Print "申請案號：" & m_TM12
         End If
      Case "CFP":
         If IsNull(m_TM15) = False Then
            Printer.Print "專利號數：" & m_TM15
         Else
            Printer.Print "申請案號：" & m_TM12
         End If
      Case Else
         Printer.Print "申請案號：" & m_TM12
   End Select
   '91.12.5 end
   
   '來函收文號(下一程序檔的收文號)
   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "來函收文號：" & m_NP01

   nRow = nRow + 1
   ' 案件中文名稱
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify By Sindy 2015/5/1
   If bolHadEFile = False Then
      Printer.Print "案件名稱：" & m_TM05
   Else
      Printer.Print "案件名稱：" & StrToStr(m_TM05, 16)
      Printer.CurrentX = 9500 '(m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print " (有電子回覆單)"
   End If
   
   nRow = nRow + 1
   ' 商品類別/符合年費減免
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   If m_NP02 = "P" Then
      stCP81 = PUB_GetCaseDiscStat(m_NP02 & m_NP03 & m_NP04 & m_NP05)
      If stCP81 = "Y" Then
         Printer.Print "符合年費減免：■ 是　□ 否"
      ElseIf stCP81 = "N" Then
         Printer.Print "符合年費減免：□ 是　■ 否"
      Else
         Printer.Print "符合年費減免：□ 是　□ 否"
      End If
   Else
       Printer.Print "商品類別：" & m_TM09
   End If
   
   nRow = nRow + 1
   ' 有FC代理人時
   If IsEmptyText(m_Agent) = False Then
      ' 代理人編號
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "代理人編號：" & Mid(m_Agent, 1, 6) & "-" & Mid(m_Agent, 7, 2) & "-" & Mid(m_Agent, 9, 1)
      ' 代理人 : 新(   ) 舊(   )
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      '92.8.15 MODIFY BY SONIA
      'Printer.Print "代理人：新(   ) 舊(   )"
      Printer.Print "代理人："
      ' 代理人名稱
      'Printer.CurrentX = (m_LeftMargin + 50) * m_CharWidth
      Printer.CurrentX = (m_LeftMargin + 38) * m_CharWidth
      '92.8.15 END
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print GetFAgentName(m_Agent)
      
      ' 下一列
      nRow = nRow + 1
      ' 申請人
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "申請人　：" & m_Customer
      ' 申請人名稱(中)(英)
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      '2005/11/2 MODIFY BY SONIA
      'Printer.Print m_Customer_CName & "  " & m_Customer_EName
      If more_Customer <> "" Then
         Printer.Print m_Customer_CName & " (多申請人) " & m_Customer_EName
      Else
         Printer.Print m_Customer_CName & "  " & m_Customer_EName
      End If
      '2005/11/2 END
      ' 下一列
      nRow = nRow + 1
      ' 申請地址:
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "英文地址：" & m_Customer_EAddress
      
      'Add By Sindy 2017/4/25
      If m_NP02 = "CFT" And m_NP07 = "601" Then
         nRow = nRow + 1
         ' 下一程序備註
         Printer.CurrentX = m_LeftMargin * m_CharWidth
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print "下一程序備註：" & strNP15
      End If
   Else
      ' 客戶編號
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "客戶編號：" & Mid(m_Customer, 1, 6) & "-" & Mid(m_Customer, 7, 2) & "-" & Mid(m_Customer, 9, 1)
      ' 接洽人
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
        'Modify By Cheng 2003/05/29
'      Printer.Print "接洽人：" & m_Customer_Keep
      Printer.Print "接洽人：" & m_Customer_Keep & IIf(m_Customer_Keep <> "", "　", "　　　　　") & "電話：" & m_Customer_Tel & IIf(m_Customer_Tel <> "", "　", "               ") & "傳真：" & m_Customer_Fax
'      ' 電話
'      Printer.CurrentX = (m_LeftMargin + 50) * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "電話：" & m_Customer_Tel
'      ' 傳真
'      Printer.CurrentX = (m_LeftMargin + 70) * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "傳真：" & m_Customer_Fax
      
      ' 下一列
      nRow = nRow + 1
      ' 申請人
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      '2005/11/2 MODIFY BY SONIA
      'Printer.Print "申請人　：" & GetCustomerName(m_Customer, "0")
      If more_Customer <> "" Then
         Printer.Print "申請人　：" & GetCustomerName(m_Customer, "0") & " (多申請人)"
      Else
         Printer.Print "申請人　：" & GetCustomerName(m_Customer, "0")
      End If
      '2005/11/2 END
      ' 負責人
        'Modify By Cheng 2003/05/29
'      Printer.CurrentX = (m_LeftMargin + 50) * m_CharWidth
      Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "負責人：" & m_Customer_Main
   
      nRow = nRow + 1
      ' 聯絡地址
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "聯絡地址：" & m_Customer_Adrs & vbTab
   
      nRow = nRow + 1
      ' 申請地址
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "申請地址：" & m_Customer_CAddress
      
      'Add By Sindy 2017/4/25
      If m_NP02 = "CFT" And m_NP07 = "601" Then
         nRow = nRow + 1
         ' 下一程序備註
         Printer.CurrentX = m_LeftMargin * m_CharWidth
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print "下一程序備註：" & strNP15
      End If
   End If
   'add by nickc 2008/01/18 加印客戶其他資料
   GetOtherCuData m_NP02, m_NP03, m_NP04, m_NP05
'   If IsEmpty(m_CuMemo) = False And CheckStr(m_CuMemo) <> "" And IsEmpty(m_SpecCu) = False And CheckStr(m_SpecCu) <> "" Then
'      Printer.CurrentX = m_LeftMargin * m_CharWidth + 150
'      Printer.CurrentY = (m_TopMargin + 25) * m_CharHeight + 100
'      Printer.Print "　　　　　業務備註：" & m_CuMemo
'      Printer.CurrentX = m_LeftMargin * m_CharWidth + 150
'      Printer.CurrentY = (m_TopMargin + 26) * m_CharHeight + 100
'      Printer.Print "　　　　　此為特殊客戶，特殊原因為：" & m_SpecCu
'   ElseIf IsEmpty(m_CuMemo) = False And CheckStr(m_CuMemo) <> "" Then
'      Printer.CurrentX = m_LeftMargin * m_CharWidth + 150
'      Printer.CurrentY = (m_TopMargin + 26) * m_CharHeight + 100
'      Printer.Print "　　　　　業務備註：" & m_CuMemo
'   ElseIf IsEmpty(m_SpecCu) = False And CheckStr(m_SpecCu) <> "" Then
'      Printer.CurrentX = m_LeftMargin * m_CharWidth + 150
'      Printer.CurrentY = (m_TopMargin + 26) * m_CharHeight + 100
'      Printer.Print "　　　　　此為特殊客戶，特殊原因為：" & m_SpecCu
'   End If
   'Modify By Sindy 2014/7/18
   Dim ii As Integer
   Dim ii_Row As Integer
   Dim strText1 As String
   Dim arrText1, iiiii As Integer
   Dim intStarY As Integer
   intStarY = 11000
   ii_Row = 0
   For ii = 1 To 2
      If ii = 1 Then strText1 = Trim(m_CuMemo)
      If ii = 2 Then strText1 = Trim(m_SpecCu)
      If strText1 > "" Then
         If ii = 1 Then strText1 = "業務備註：" & strText1
         If ii = 2 Then strText1 = "此為特殊客戶，特殊原因為：" & strText1
         arrText1 = Split(strText1, vbCrLf)
         If UBound(arrText1) > 0 Then
            For iiiii = 0 To UBound(arrText1)
               ii_Row = ii_Row + 1
               Printer.CurrentX = 1200
               Printer.CurrentY = intStarY + (300 * ii_Row)
               Printer.Print arrText1(iiiii)
            Next iiiii
         Else
            If Len(strText1) > 41 Then
               Do While Len(strText1) > 41
                  ii_Row = ii_Row + 1
                  Printer.CurrentX = 1200
                  Printer.CurrentY = intStarY + (300 * ii_Row)
                  Printer.Print Mid(strText1, 1, 41)
                  strText1 = Right(strText1, Len(strText1) - 41)
               Loop
               ii_Row = ii_Row + 1
               Printer.CurrentX = 1200
               Printer.CurrentY = intStarY + (300 * ii_Row)
               Printer.Print strText1
            Else
               'Add By Sindy 2017/5/19
               nRow = nRow + 1
               Printer.CurrentX = m_LeftMargin * m_CharWidth
               Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
               '2017/5/19 END
               Printer.Print strText1
            End If
         End If
      End If
   Next ii
   '2014/7/18 END
   If bNewFunc = False Then
      If bContinue Then
         Printer.NewPage
      Else
         Printer.EndDoc
      End If
   End If
End Sub

'Add By Cheng 2003/06/23
'Modify by Morgan 2009/5/27 +bContinue:是否繼續列印
'Modify by Sindy 2010/5/01 +bNewFunc:使用於結案單
'Modify by Sindy 2010/5/12 +strNP07:案件性質
'Modify By Sindy 2012/6/8 改為不套表,直接列印完整接洽單
'Modify By Sindy 2012/6/18 +bolFullPrint
'Modify By Sindy 2015/5/1 +bolHadEFile:是否有電子回覆單
'Modify by Amy 2019/06/04 +bolShowMsg:列印完秀訊息
Private Sub ExportDataFCT(Optional strBKindCP09 As String = Empty, Optional bContinue As Boolean, Optional bNewFunc As Boolean = False, _
                          Optional strNP07 As String = "", Optional bolFullPrint As Boolean = True, Optional bolHadEFile As Boolean = False, Optional bolShowMsg As Boolean = False)
    Dim m_CharWidth As Long
    Dim m_CharHeight As Long
    Dim strData As String
    Dim nCount As Integer
    Dim strSQL1 As String
    Dim Rs2 As New ADODB.Recordset
    ' 報表左方留白的寬度
    Dim m_LeftMargin As Integer
    ' 報表上方留白的高度
    Dim m_TopMargin As Integer
    Dim nRow As Integer
    
    m_CharWidth = 120
    m_CharHeight = 280
    'Modify By Cheng 2003/05/29
'   m_LeftMargin = 6
    m_LeftMargin = 4
    m_TopMargin = 2 '4
    If bNewFunc = False Then
      '***********  90.11.14   nickc
      Printer.Orientation = 1
      '*****************************
      Printer.PaperSize = vbPRPSA4
    End If
    
    Printer.FontName = "細明體" 'Add by Morgan 2004/9/10
   
   'Add By Sindy 2012/6/8
   If bolFullPrint = True Then
      nRow = 1
      Printer.Font.Size = 18
      Printer.FontBold = True
      'Modified by Morgan 2020/3/30 取消公司名稱
      'Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("台一國際專利商標事務所外商案件接洽記錄單") / 2)
      'Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      'Printer.Print "台一國際專利商標事務所外商案件接洽記錄單"
      Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("外商案件接洽記錄單") / 2)
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "外商案件接洽記錄單"
      'end 2020/3/30
      nRow = 4
      Printer.FontBold = False
   Else
      nRow = 2
   End If
   '2012/6/8 End
   
   Printer.Font.Size = 12
   ' 總收文號
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
    'Modify By Cheng 2003/04/16
'   Printer.Print "總收文號："
   Printer.Print "總收文號：" & IIf(strBKindCP09 <> "", strBKindCP09, "")
   ' 區別
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "區別　　：" & m_NP10_Dep
   ' 填表
   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Add By Sindy 2010/4/28
   If bNewFunc = True Then
      Printer.Print "填表日期：" & ChangeTStringToTDateString(strSrvDate(2))
   '2010/4/28 End
   Else
      Printer.Print "填表日期："
   End If
   
   nRow = nRow + 1
   ' 收文日期
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "收文日期："
   'Add By Sindy 2013/7/16
   '無期限閉卷
   If strNP07 = "000000" Then
      ' 法定期限
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "法定期限："
   Else
   '2013/7/16 END
      ' 法定期限
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "法定期限：" & m_NP09
   End If
   ' 本所人員
   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify By Sindy 2011/7/11
   If bNewFunc = True And m_NP10_Name <> strUserName Then
      Printer.Print "本所人員：" & "★" & strUserName
   '2011/7/11 End
   Else
      Printer.Print "本所人員：" & m_NP10_Name
   End If
   nRow = nRow + 1
   
   'add by nickc 2006/09/12
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   If m_NP07 = "403" And m_Nation = "000" Then
       ' 本所案號
       Printer.Print "本所案號：" & m_NP02 & "-" & m_NP03 & "-" & m_NP04 & "-" & m_NP05
   End If
   
   
   ' 約定期限
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "約定期限："
   ' 員工編號
   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify By Sindy 2011/7/11
   If bNewFunc = True And m_NP10_Name <> strUserName Then
      Printer.Print "員工編號：" & strUserNum
   '2011/7/11 End
   Else
      Printer.Print "員工編號：" & m_NP10
   End If
   nRow = nRow + 1
   ' 本所案號
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify By Sindy 2010/5/10
   If bNewFunc = True Then
      Printer.Print "本所案號：" & m_NP02 & "-" & m_NP03 & "-" & m_NP04 & "-" & m_NP05
   '2010/5/10 End
   Else
      'add by nickc 2006/09/12
      If m_NP07 = "403" And m_Nation = "000" Then
          Printer.Print "案件性質：□ " & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0)
      Else
          Printer.Print "本所案號：" & m_NP02 & "-" & m_NP03 & "-" & m_NP04 & "-" & m_NP05
      End If
   End If
   
   'Add By Sindy 2013/7/16
   '無期限閉卷
   If strNP07 = "000000" Then
      ' 本所期限
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "本所期限："
   Else
   '2013/7/16 END
      ' 本所期限
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "本所期限：" & m_NP08
   End If
   ' 申請國家
   If IsEmptyText(m_Nation) = False Then
      Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "申請國家：" & GetNationName(m_Nation, 0)
   End If
   nRow = nRow + 1
   
   'Modify By Sindy 2010/5/10
   If bNewFunc = False Then
      '2005/12/30 ADD BY SONIA
      If m_NP07 = "715" Then
         Printer.CurrentX = m_LeftMargin * m_CharWidth
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print "案件性質：□ " & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0)
      'add by nickc 2006/09/08
      ElseIf m_NP07 = "403" And m_Nation = "000" Then
         Printer.CurrentX = m_LeftMargin * m_CharWidth
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print "　　　　　□ 204 " & GetCaseTypeName(m_NP02, "204", 0) & ""
      End If
      '2005/12/30 END
   End If
   
   ' 審定號, 專利號數或申請案號
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify By Cheng 2002/01/04
'   Printer.Print m_TM12
   If IsNull(m_TM15) = False Then
      Printer.Print "審定號　：" & m_TM15
   Else
      Printer.Print "申請案號：" & m_TM12
   End If
   'Add By Cheng 2002/04/10
   '來函收文號(下一程序檔的收文號)
   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "來函收文號：" & m_NP01
   nRow = nRow + 1
   ' 案件性質
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   
   'Add By Sindy 2013/7/16
   '無期限閉卷
   If strNP07 = "000000" Then
      Printer.Print "案件性質：" & "★" & "無期限閉卷"
   Else
      'Modify By Sindy 2010/5/10
      If bNewFunc = True Then
         'Modify By Sindy 2010/5/12
         If strNP07 <> "" Then
            If m_Nation < "010" Then
               Printer.Print "案件性質：" & strNP07 & " " & GetCaseTypeName(m_NP02, strNP07, 0)
            Else
               Printer.Print "案件性質：" & strNP07 & " " & GetCaseTypeName(m_NP02, strNP07, 1)
            End If
         '2010/5/12 End
         Else
            If m_Nation < "010" Then
               Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0)
            Else
               Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 1)
            End If
         End If
      '2010/5/10 End
      Else
         If m_Nation < "010" Then
            If m_NP07 = "715" Then
               Printer.Print "          □ 717 " & GetCaseTypeName(m_NP02, "717", 0)
            Else
               'add by nickc 2006/09/12
               If m_NP07 = "403" And m_Nation = "000" Then
                  Printer.Print "　　　　　□ 205 " & GetCaseTypeName(m_NP02, "205", 0) & ""
               Else
                  Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0)
               End If
            End If
         Else
               Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 1)
         End If
      End If
   End If
   
   ' 案件中文名稱
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify By Cheng 2002/01/04
'   Printer.Print "主題：" & m_TM05
   'Modify By Sindy 2015/5/1
   If bolHadEFile = False Then
      Printer.Print "案件名稱：" & m_TM05
   Else
      Printer.Print "案件名稱：" & StrToStr(m_TM05, 16)
      Printer.CurrentX = 9500 '(m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print " (有電子回覆單)"
   End If
   
   nRow = nRow + 1
   ' 電子送件
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify by Amy 2019/06/06 不顯示商品類別
   'Printer.Print "商品類別：" & m_TM09
    Printer.Print "　　　　　□電子送件"
    
    'Add by Amy 2019/06/11 商品類別加回
    nRow = nRow + 1
   ' 商品類別
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "商品類別：" & m_TM09
   
  
   nRow = nRow + 1
   ' 有FC代理人時
   If IsEmptyText(m_Agent) = False Then
      ' 代理人編號
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "代理人編號：" & Mid(m_Agent, 1, 6) & "-" & Mid(m_Agent, 7, 2) & "-" & Mid(m_Agent, 9, 1)
      ' 代理人 : 新(   ) 舊(   )
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "代理人：新(   ) 舊(   )"
      ' 代理人名稱
      Printer.CurrentX = (m_LeftMargin + 50) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print GetFAgentName(m_Agent)
      ' 下一列
      nRow = nRow + 1
      ' 代理人英文地址
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "英文地址：" & GetFagentAddress(m_Agent, "1")
      ' 下一列
      nRow = nRow + 1
      '彼所案號
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "彼所案號：" & m_TM45
      ' 下一列
      nRow = nRow + 1
      ' 申請人
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "申請人　：" & Mid(m_Customer, 1, 6) & "-" & Mid(m_Customer, 7, 2) & "-" & Mid(m_Customer, 9, 1)
      ' 申請人名稱(中)(英)
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print m_Customer_CName & "  " & m_Customer_EName
'      ' 下一列
'      nRow = nRow + 1
'      ' 申請地址:
'      Printer.CurrentX = m_LeftMargin * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "英文地址：" & m_Customer_EAddress
   '無代理人時
   Else
      '彼所案號
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "彼所案號：" & m_TM45
      ' 客戶編號
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "客戶編號：" & Mid(m_Customer, 1, 6) & "-" & Mid(m_Customer, 7, 2) & "-" & Mid(m_Customer, 9, 1)
      ' 接洽人
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
        'Modify By Cheng 2003/05/29
'      Printer.Print "接洽人：" & m_Customer_Keep
      Printer.Print "接洽人：" & m_Customer_Keep & IIf(m_Customer_Keep <> "", "　", "　　　　　") & "電話：" & m_Customer_Tel & IIf(m_Customer_Tel <> "", "　", "               ") & "傳真：" & m_Customer_Fax
'      ' 電話
'      Printer.CurrentX = (m_LeftMargin + 50) * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "電話：" & m_Customer_Tel
'      ' 傳真
'      Printer.CurrentX = (m_LeftMargin + 70) * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "傳真：" & m_Customer_Fax
      ' 下一列
      nRow = nRow + 1
      ' 申請人
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "申請人　：" & GetCustomerName(m_Customer, "0")
      ' 負責人
        'Modify By Cheng 2003/05/29
'      Printer.CurrentX = (m_LeftMargin + 50) * m_CharWidth
      Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "負責人：" & m_Customer_Main
      nRow = nRow + 1
      ' 聯絡地址
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "聯絡地址：" & m_Customer_Adrs & vbTab
      nRow = nRow + 1
      ' 申請地址
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "申請地址：" & m_Customer_CAddress
   End If
   
   'Add By Sindy 2012/6/8
   If bolFullPrint = True Then
      Dim LngLineY As Long, LngTextY As Long
      '印框
      '*** 設定值 Star
      nRow = 15 '列數
      intI = 350 '間距
      '*** End
      '[test]
      'PUB_PrintCaseCloseSheet strUserNum
      'Pub_StrUserSt03 = "F11"
      '請款內容
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (500, LngLineY)-(11300, LngLineY)
      Printer.Line (500, LngLineY)-(500, 8050)
      Printer.Line (850, LngLineY)-(850, 8050)
      Printer.Line (11300, LngLineY)-(11300, 8050)
      Printer.CurrentX = 550
      Printer.CurrentY = LngTextY
      Printer.Print "請"
      Printer.Line (1200, LngLineY)-(1200, LngLineY + 350)
      Printer.CurrentX = 1250
      Printer.CurrentY = LngTextY
      Printer.Print "不請款"
      Printer.Line (4000, LngLineY)-(4000, LngLineY + 350)
      Printer.Line (4350, LngLineY)-(4350, LngLineY + 350)
      Printer.CurrentX = 4400
      Printer.CurrentY = LngTextY
      Printer.Print "合併列印請款單"
      'Printer.Line (7650, LngLineY)-(7650, LngLineY + 350)
      Printer.Line (8000, LngLineY)-(8000, LngLineY + 350)
      Printer.CurrentX = 8050
      Printer.CurrentY = LngTextY
      Printer.Print "列印申請人：　　(要列印Y)"
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (850, LngLineY)-(11300, LngLineY)
      Printer.Font.Size = 10
      Printer.CurrentX = 900
      Printer.CurrentY = LngTextY
      Printer.Print "固定列印對象："
      Printer.CurrentX = 3700
      Printer.CurrentY = LngTextY
      Printer.Print "固定請款對象："
      Printer.CurrentX = 6500
      Printer.CurrentY = LngTextY
      Printer.Print "請款對象："
      Printer.CurrentX = 9000
      Printer.CurrentY = LngTextY
      Printer.Print "列印對象："
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (850, LngLineY)-(11300, LngLineY)
      Printer.Font.Size = 12
      Printer.CurrentX = 550
      Printer.CurrentY = LngTextY
      Printer.Print "款"
      Printer.Font.Size = 14
      Printer.Font.Bold = True
      Printer.CurrentX = 900
      Printer.CurrentY = LngTextY
      Printer.Print "總金額："
      Printer.Line (1900, LngLineY + 300)-(3900, LngLineY + 300)
      Printer.CurrentX = 4000
      Printer.CurrentY = LngTextY
      Printer.Print "規費："
      Printer.Line (4800, LngLineY + 300)-(6800, LngLineY + 300)
      Printer.CurrentX = 7000
      Printer.CurrentY = LngTextY
      Printer.Print "點數："
      Printer.Line (7800, LngLineY + 300)-(9800, LngLineY + 300)
      Printer.Font.Size = 12
      Printer.Font.Bold = False
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (850, LngLineY)-(11300, LngLineY)
      Dim ii As Integer
      Printer.CurrentX = 850 + 200
      Printer.CurrentY = LngTextY
      Printer.Print "代　號"
      ii = 850 + 1160
      Printer.Line (ii, LngLineY)-(ii, 8050)
      Printer.CurrentX = ii + 200
      Printer.CurrentY = LngTextY
      Printer.Print "金　額"
      ii = ii + 1160
      Printer.Line (ii, LngLineY)-(ii, 8050)
      Printer.CurrentX = ii + 200
      Printer.CurrentY = LngTextY
      Printer.Print "折　扣"
      ii = ii + 1160
      Printer.Line (ii, LngLineY)-(ii, 8050)
      ii = ii + 25
      Printer.Line (ii, LngLineY)-(ii, 8050)
      Printer.CurrentX = ii + 200
      Printer.CurrentY = LngTextY
      Printer.Print "代　號"
      ii = ii + 1160
      Printer.Line (ii, LngLineY)-(ii, 8050)
      Printer.CurrentX = ii + 200
      Printer.CurrentY = LngTextY
      Printer.Print "金　額"
      ii = ii + 1160
      Printer.Line (ii, LngLineY)-(ii, 8050)
      Printer.CurrentX = ii + 200
      Printer.CurrentY = LngTextY
      Printer.Print "折　扣"
      ii = ii + 1160
      Printer.Line (ii, LngLineY)-(ii, 8050)
      ii = ii + 25
      Printer.Line (ii, LngLineY)-(ii, 8050)
      Printer.CurrentX = ii + 200
      Printer.CurrentY = LngTextY
      Printer.Print "代　號"
      ii = ii + 1160
      Printer.Line (ii, LngLineY)-(ii, 8050)
      Printer.CurrentX = ii + 200
      Printer.CurrentY = LngTextY
      Printer.Print "金　額"
      ii = ii + 1160
      Printer.Line (ii, LngLineY)-(ii, 8050)
      Printer.CurrentX = ii + 200
      Printer.CurrentY = LngTextY
      Printer.Print "折　扣"
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (850, LngLineY)-(11300, LngLineY)
      Printer.CurrentX = 550
      Printer.CurrentY = LngTextY
      Printer.Print "內"
      Printer.CurrentX = 7950
      Printer.CurrentY = LngTextY
      Printer.Print "01FAX"
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (850, LngLineY)-(11300, LngLineY)
      Printer.CurrentX = 7950
      Printer.CurrentY = LngTextY
      Printer.Print "02雜費"
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (850, LngLineY)-(11300, LngLineY)
      nRow = nRow + 1
      Printer.CurrentX = 550
      Printer.CurrentY = LngTextY
      Printer.Print "容"
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (500, LngLineY)-(11300, LngLineY)
      '補正事項
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (500, LngLineY)-(11300, LngLineY)
      Printer.Line (500, LngLineY)-(500, 15050)
      Printer.Line (850, LngLineY)-(850, 15050)
      Printer.Line (11300, LngLineY)-(11300, 15050)
      Printer.CurrentX = 550
      Printer.CurrentY = LngTextY
      Printer.Print "補"
      Printer.Line (1200, LngLineY)-(1200, 10150)
      Printer.CurrentX = 1250
      Printer.CurrentY = LngTextY
      Printer.Print "覆行文"
      Printer.Line (4000, LngLineY)-(4000, LngLineY + 350)
      Printer.Line (4350, LngLineY)-(4350, LngLineY + 350)
      Printer.CurrentX = 4400
      Printer.CurrentY = LngTextY
      Printer.Print "主動行文"
      Printer.Line (7650, LngLineY)-(7650, LngLineY + 350)
      Printer.Line (8000, LngLineY)-(8000, LngLineY + 350)
      Printer.CurrentX = 8050
      Printer.CurrentY = LngTextY
      Printer.Print "電話回覆"
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (850, LngLineY)-(11300, LngLineY)
      Printer.CurrentX = 1250
      Printer.CurrentY = LngTextY
      Printer.Print "覆電話指示－日期：　　年　　月　　日　　審查員：　　　　　先生／小姐　　分機："
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (850, LngLineY)-(11300, LngLineY)
      Printer.CurrentX = 1250
      Printer.CurrentY = LngTextY
      Printer.Print "委任狀"
      Printer.Line (3450, LngLineY)-(3450, 10150)
      Printer.Line (3800, LngLineY)-(3800, 10150)
      Printer.CurrentX = 3850
      Printer.CurrentY = LngTextY
      Printer.Print "移轉契約書"
      Printer.Line (6100, LngLineY)-(6100, 10150)
      Printer.Line (6450, LngLineY)-(6450, 10150)
      Printer.CurrentX = 6500
      Printer.CurrentY = LngTextY
      Printer.Print "更名證明"
      Printer.Line (8750, LngLineY)-(8750, 10150)
      Printer.Line (9100, LngLineY)-(9100, 10150)
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (850, LngLineY)-(11300, LngLineY)
      Printer.CurrentX = 1250
      Printer.CurrentY = LngTextY
      Printer.Print "優先權證明"
      Printer.CurrentX = 3850
      Printer.CurrentY = LngTextY
      Printer.Print "授權契約書"
      Printer.CurrentX = 6500
      Printer.CurrentY = LngTextY
      Printer.Print "再授權契約書"
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (850, LngLineY)-(11300, LngLineY)
      Printer.CurrentX = 1250
      Printer.CurrentY = LngTextY
      Printer.Print "港澳大陸地區證明"
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (850, LngLineY)-(11300, LngLineY)
      Printer.Font.Size = 14
      Printer.Font.Bold = True
      Printer.CurrentX = 900
      Printer.CurrentY = LngTextY
      Printer.Print "工時："
      Printer.Font.Size = 12
      Printer.Font.Bold = False
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      'Printer.Line (850, LngLineY)-(11300, LngLineY)
      Printer.CurrentX = 550
      Printer.CurrentY = LngTextY
      Printer.Print "正"
      nRow = nRow + 6
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      'Printer.Line (850, LngLineY)-(11300, LngLineY)
      Printer.CurrentX = 550
      Printer.CurrentY = LngTextY
      Printer.Print "事"
      nRow = nRow + 6
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      'Printer.Line (850, LngLineY)-(11300, LngLineY)
      Printer.CurrentX = 550
      Printer.CurrentY = LngTextY
      Printer.Print "項"
      nRow = nRow + 1
      LngLineY = (intI * nRow): LngTextY = (intI * nRow) + 50
      Printer.Line (500, LngLineY)-(11300, LngLineY)
      '簽名蓋章處
      Printer.CurrentX = 1500
      Printer.CurrentY = LngTextY
      Printer.Print "經　理："
      Printer.CurrentX = 4500
      Printer.CurrentY = LngTextY
      Printer.Print "主　任："
      Printer.CurrentX = 7500
      Printer.CurrentY = LngTextY
      Printer.Print "會　計："
   End If
   '2012/6/8 End
   
   If bNewFunc = False Then
      If bContinue = True Then
         Printer.NewPage
      Else
         Printer.EndDoc
         If bolShowMsg = True Then
            MsgBox "列印完畢！", vbInformation
         End If
      End If
   End If
End Sub
'Private Sub ExportDataFCT(Optional strBKindCP09 As String = Empty, Optional bContinue As Boolean, Optional bNewFunc As Boolean = False, Optional strNP07 As String = "")
'    Dim m_CharWidth As Long
'    Dim m_CharHeight As Long
'    Dim strData As String
'    Dim nCount As Integer
'    Dim strSQL1 As String
'    Dim Rs2 As New ADODB.Recordset
'    ' 報表左方留白的寬度
'    Dim m_LeftMargin As Integer
'    ' 報表上方留白的高度
'    Dim m_TopMargin As Integer
'    Dim nRow As Integer
'
'    m_CharWidth = 120
'    m_CharHeight = 280
'    'Modify By Cheng 2003/05/29
''   m_LeftMargin = 6
'    m_LeftMargin = 4
'    m_TopMargin = 4
'    If bNewFunc = False Then
'      '***********  90.11.14   nickc
'      Printer.Orientation = 1
'      '*****************************
'      Printer.PaperSize = vbPRPSA4
'    End If
'    Printer.Font.Size = 12
'    Printer.FontName = "細明體" 'Add by Morgan 2004/9/10
'
'   nRow = 1
'   ' 總收文號
'   Printer.CurrentX = m_LeftMargin * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'    'Modify By Cheng 2003/04/16
''   Printer.Print "總收文號："
'   Printer.Print "總收文號：" & IIf(strBKindCP09 <> "", strBKindCP09, "")
'   ' 區別
'   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   Printer.Print "區別　　：" & m_NP10_Dep
'   ' 填表
'   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   'Add By Sindy 2010/4/28
'   If bNewFunc = True Then
'      Printer.Print "填表日期：" & ChangeTStringToTDateString(strSrvDate(2))
'   '2010/4/28 End
'   Else
'      Printer.Print "填表日期："
'   End If
'
'   nRow = nRow + 1
'   ' 收文日期
'   Printer.CurrentX = m_LeftMargin * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   Printer.Print "收文日期："
'   ' 法定期限
'   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   Printer.Print "法定期限：" & m_NP09
'   ' 本所人員
'   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   'Modify By Sindy 2011/7/11
'   If bNewFunc = True And m_NP10_Name <> strUserName Then
'      Printer.Print "本所人員：" & "★" & strUserName
'   '2011/7/11 End
'   Else
'      Printer.Print "本所人員：" & m_NP10_Name
'   End If
'   nRow = nRow + 1
'
'   'add by nickc 2006/09/12
'   Printer.CurrentX = m_LeftMargin * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   If m_NP07 = "403" And m_Nation = "000" Then
'       ' 本所案號
'       Printer.Print "本所案號：" & m_NP02 & "-" & m_NP03 & "-" & m_NP04 & "-" & m_NP05
'   End If
'
'
'   ' 約定期限
'   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   Printer.Print "約定期限："
'   ' 員工編號
'   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   'Modify By Sindy 2011/7/11
'   If bNewFunc = True And m_NP10_Name <> strUserName Then
'      Printer.Print "員工編號：" & strUserNum
'   '2011/7/11 End
'   Else
'      Printer.Print "員工編號：" & m_NP10
'   End If
'   nRow = nRow + 1
'   ' 本所案號
'   Printer.CurrentX = m_LeftMargin * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   'Modify By Sindy 2010/5/10
'   If bNewFunc = True Then
'      Printer.Print "本所案號：" & m_NP02 & "-" & m_NP03 & "-" & m_NP04 & "-" & m_NP05
'   '2010/5/10 End
'   Else
'      'add by nickc 2006/09/12
'      If m_NP07 = "403" And m_Nation = "000" Then
'          Printer.Print "案件性質：□ " & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0)
'      Else
'          Printer.Print "本所案號：" & m_NP02 & "-" & m_NP03 & "-" & m_NP04 & "-" & m_NP05
'      End If
'   End If
'
'   ' 本所期限
'   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   Printer.Print "本所期限：" & m_NP08
'   ' 申請國家
'   If IsEmptyText(m_Nation) = False Then
'      Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "申請國家：" & GetNationName(m_Nation, 0)
'   End If
'   nRow = nRow + 1
'
'   'Modify By Sindy 2010/5/10
'   If bNewFunc = False Then
'      '2005/12/30 ADD BY SONIA
'      If m_NP07 = "715" Then
'         Printer.CurrentX = m_LeftMargin * m_CharWidth
'         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'         Printer.Print "案件性質：□ " & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0)
'      'add by nickc 2006/09/08
'      ElseIf m_NP07 = "403" And m_Nation = "000" Then
'         Printer.CurrentX = m_LeftMargin * m_CharWidth
'         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'         Printer.Print "　　　　　□ 204 " & GetCaseTypeName(m_NP02, "204", 0) & ""
'      End If
'      '2005/12/30 END
'   End If
'
'   ' 審定號, 專利號數或申請案號
'   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   'Modify By Cheng 2002/01/04
''   Printer.Print m_TM12
'   If IsNull(m_TM15) = False Then
'      Printer.Print "審定號　：" & m_TM15
'   Else
'      Printer.Print "申請案號：" & m_TM12
'   End If
'   'Add By Cheng 2002/04/10
'   '來函收文號(下一程序檔的收文號)
'   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   Printer.Print "來函收文號：" & m_NP01
'   nRow = nRow + 1
'   ' 案件性質
'   Printer.CurrentX = m_LeftMargin * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   'Modify By Sindy 2010/5/10
'   If bNewFunc = True Then
'      'Modify By Sindy 2010/5/12
'      If strNP07 <> "" Then
'         If m_Nation < "010" Then
'            Printer.Print "案件性質：" & strNP07 & " " & GetCaseTypeName(m_NP02, strNP07, 0)
'         Else
'            Printer.Print "案件性質：" & strNP07 & " " & GetCaseTypeName(m_NP02, strNP07, 1)
'         End If
'      '2010/5/12 End
'      Else
'         If m_Nation < "010" Then
'            Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0)
'         Else
'            Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 1)
'         End If
'      End If
'   '2010/5/10 End
'   Else
'      If m_Nation < "010" Then
'         If m_NP07 = "715" Then
'            Printer.Print "          □ 717 " & GetCaseTypeName(m_NP02, "717", 0)
'         Else
'            'add by nickc 2006/09/12
'            If m_NP07 = "403" And m_Nation = "000" Then
'               Printer.Print "　　　　　□ 205 " & GetCaseTypeName(m_NP02, "205", 0) & ""
'            Else
'               Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 0)
'            End If
'         End If
'      Else
'            Printer.Print "案件性質：" & m_NP07 & " " & GetCaseTypeName(m_NP02, m_NP07, 1)
'      End If
'   End If
'   ' 案件中文名稱
'   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   'Modify By Cheng 2002/01/04
''   Printer.Print "主題：" & m_TM05
'   Printer.Print "案件名稱：" & m_TM05
'   nRow = nRow + 1
'   ' 商品類別
'   Printer.CurrentX = m_LeftMargin * m_CharWidth
'   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'   Printer.Print "商品類別：" & m_TM09
'   nRow = nRow + 1
'   ' 有FC代理人時
'   If IsEmptyText(m_Agent) = False Then
'      ' 代理人編號
'      Printer.CurrentX = m_LeftMargin * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "代理人編號：" & Mid(m_Agent, 1, 6) & "-" & Mid(m_Agent, 7, 2) & "-" & Mid(m_Agent, 9, 1)
'      ' 代理人 : 新(   ) 舊(   )
'      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "代理人：新(   ) 舊(   )"
'      ' 代理人名稱
'      Printer.CurrentX = (m_LeftMargin + 50) * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print GetFAgentName(m_Agent)
'      ' 下一列
'      nRow = nRow + 1
'      ' 代理人英文地址
'      Printer.CurrentX = m_LeftMargin * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "英文地址：" & GetFagentAddress(m_Agent, "1")
'      ' 下一列
'      nRow = nRow + 1
'      '彼所案號
'      Printer.CurrentX = m_LeftMargin * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "彼所案號：" & m_TM45
'      ' 下一列
'      nRow = nRow + 1
'      ' 申請人
'      Printer.CurrentX = m_LeftMargin * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "申請人　：" & Mid(m_Customer, 1, 6) & "-" & Mid(m_Customer, 7, 2) & "-" & Mid(m_Customer, 9, 1)
'      ' 申請人名稱(中)(英)
'      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print m_Customer_CName & "  " & m_Customer_EName
''      ' 下一列
''      nRow = nRow + 1
''      ' 申請地址:
''      Printer.CurrentX = m_LeftMargin * m_CharWidth
''      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
''      Printer.Print "英文地址：" & m_Customer_EAddress
'   '無代理人時
'   Else
'      '彼所案號
'      Printer.CurrentX = m_LeftMargin * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "彼所案號：" & m_TM45
'      ' 客戶編號
'      Printer.CurrentX = m_LeftMargin * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "客戶編號：" & Mid(m_Customer, 1, 6) & "-" & Mid(m_Customer, 7, 2) & "-" & Mid(m_Customer, 9, 1)
'      ' 接洽人
'      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'        'Modify By Cheng 2003/05/29
''      Printer.Print "接洽人：" & m_Customer_Keep
'      Printer.Print "接洽人：" & m_Customer_Keep & IIf(m_Customer_Keep <> "", "　", "　　　　　") & "電話：" & m_Customer_Tel & IIf(m_Customer_Tel <> "", "　", "               ") & "傳真：" & m_Customer_Fax
''      ' 電話
''      Printer.CurrentX = (m_LeftMargin + 50) * m_CharWidth
''      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
''      Printer.Print "電話：" & m_Customer_Tel
''      ' 傳真
''      Printer.CurrentX = (m_LeftMargin + 70) * m_CharWidth
''      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
''      Printer.Print "傳真：" & m_Customer_Fax
'      ' 下一列
'      nRow = nRow + 1
'      ' 申請人
'      Printer.CurrentX = m_LeftMargin * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "申請人　：" & GetCustomerName(m_Customer, "0")
'      ' 負責人
'        'Modify By Cheng 2003/05/29
''      Printer.CurrentX = (m_LeftMargin + 50) * m_CharWidth
'      Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "負責人：" & m_Customer_Main
'      nRow = nRow + 1
'      ' 聯絡地址
'      Printer.CurrentX = m_LeftMargin * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "聯絡地址：" & m_Customer_Adrs & vbTab
'      nRow = nRow + 1
'      ' 申請地址
'      Printer.CurrentX = m_LeftMargin * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "申請地址：" & m_Customer_CAddress
'   End If
'
'   If bNewFunc = False Then
'      If bContinue = True Then
'         Printer.NewPage
'      Else
'         Printer.EndDoc
'      End If
'   End If
'End Sub

'Add By Cheng 2002/10/25
'Modify by Morgan 2009/5/27 +bContinue:是否繼續列印
'Modify by Sindy 2010/5/01 +bNewFunc:使用於結案單
Private Sub ExportDataCP(Optional ByVal bContinue As Boolean, Optional bNewFunc As Boolean = False)
   Dim m_CharWidth As Long
   Dim m_CharHeight As Long
   Dim strData As String
   Dim nCount As Integer
   ' 報表左方留白的寬度
   Dim m_LeftMargin As Integer
   ' 報表上方留白的高度
   Dim m_TopMargin As Integer
   Dim nRow As Integer
   
   If strSrvDate(1) >= 20100501 And bNewFunc = False Then Exit Sub '2010/4/23 ADD BY SONIA 99/5/1起不再列印
   
   m_CharWidth = 120
   m_CharHeight = 280
   m_LeftMargin = 6
   m_TopMargin = 4
   If bNewFunc = False Then
      '***********  90.11.14   nickc
      Printer.Orientation = 1
      '*****************************
      Printer.PaperSize = vbPRPSA4
   End If
   Printer.Font.Size = 12
   Printer.FontName = "細明體" 'Add by Morgan 2004/9/10

   nRow = 1
   ' 總收文號
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "總收文號："
   ' 區別
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Select Case m_CP01
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "FCP", "FCT", "FCL", "LIN", "ACS":
         Printer.Print "申請人國籍：" & m_Customer_Cny
      Case Else:
         Printer.Print "區別　　：" & m_CP12_Dep
   End Select
   ' 填表
   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Add By Sindy 2010/4/28
   If bNewFunc = True Then
      Printer.Print "填表日期：" & ChangeTStringToTDateString(strSrvDate(2))
   '2010/4/28 End
   Else
      Printer.Print "填表日期："
   End If
   
   nRow = nRow + 1
   ' 收文日期
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "收文日期："
   ' 法定期限
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "法定期限：" & m_CP07
   ' 本所人員
   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify By Sindy 2011/7/11
   If bNewFunc = True And m_CP12_Name <> strUserName Then
      Printer.Print "本所人員：" & "★" & strUserName
   '2011/7/11 End
   Else
      Printer.Print "本所人員：" & m_CP12_Name
   End If

   nRow = nRow + 1
   ' 約定期限
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "約定期限："
   ' 員工編號
   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify By Sindy 2011/7/11
   If bNewFunc = True And m_CP12_Name <> strUserName Then
      Printer.Print "員工編號：" & strUserNum
   '2011/7/11 End
   Else
      Printer.Print "員工編號：" & m_CP12
   End If
   
   nRow = nRow + 1
   ' 本所案號
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "本所案號：" & m_CP01 & "-" & m_CP02 & "-" & m_CP03 & "-" & m_CP04
   ' 本所期限
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "本所期限：" & m_CP07
   ' 申請國家
   If IsEmptyText(m_Nation) = False Then
      Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "申請國家：" & GetNationName(m_Nation, 0)
   End If

   nRow = nRow + 1
   'Add by Morgan 2004/6/30
   '原位置上移一列
   Printer.CurrentX = (m_LeftMargin) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   ' 案件性質
   Printer.Print "案件性質：" & m_CP10 & " " & GetCaseTypeName(m_CP01, m_CP10, 0)
   'End
   
   ' 審定號, 專利號數或申請案號
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   
   'Modify By Cheng 2002/01/04
   '91.12.5 MODIFY BY SONIA
   'Printer.Print "申請案號：" & m_TM12
   Select Case m_NP02
      Case "CFT", "FCT", "T", "TF":
         If IsNull(m_TM15) = False Then
            Printer.Print "審定號　：" & m_TM15
         Else
            Printer.Print "申請案號：" & m_TM12
         End If
      Case "CFP":
         If IsNull(m_TM15) = False Then
            Printer.Print "專利號數：" & m_TM15
         Else
            Printer.Print "申請案號：" & m_TM12
         End If
      Case Else
         Printer.Print "申請案號：" & m_TM12
   End Select
   '91.12.5 end
   
   'Add By Cheng 2002/04/10
   '來函收文號(下一程序檔的收文號)
   Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "來函收文號：" & m_CP09

   nRow = nRow + 1
   
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Add by Morgan 2004/6/30
   If m_Nation = "000" Then
      'Modify by Morgan 2004/8/4
      '所有性質都要放 ■ □
      'If m_NP02 = "P" And (m_NP07 = "601" Or m_NP07 = "605") Then
      If m_NP02 = "P" Then
         Dim stCP81 As String
         stCP81 = PUB_GetCaseDiscStat(m_NP02 & m_NP03 & m_NP04 & m_NP05)
         If stCP81 = "Y" Then
            Printer.Print "符合年費減免：■ 是　□ 否"
         ElseIf stCP81 = "N" Then
            Printer.Print "符合年費減免：□ 是　■ 否"
         Else
            Printer.Print "符合年費減免：□ 是　□ 否"
         End If
      End If
   End If
   'End
   
   ' 案件中文名稱
   Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   
   'Modify By Cheng 2002/01/04
   Printer.Print "案件名稱：" & m_TM05

   nRow = nRow + 1
   ' 商品類別
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "商品類別：" & m_TM09
   
   nRow = nRow + 1
   ' 有FC代理人時
   If IsEmptyText(m_Agent) = False Then
      ' 代理人編號
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "代理人編號：" & Mid(m_Agent, 1, 6) & "-" & Mid(m_Agent, 7, 2) & "-" & Mid(m_Agent, 9, 1)
      ' 代理人 : 新(   ) 舊(   )
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "代理人：新(   ) 舊(   )"
      ' 代理人名稱
      Printer.CurrentX = (m_LeftMargin + 50) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print GetFAgentName(m_Agent)
      
      ' 下一列
      nRow = nRow + 1
      ' 申請人
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "申請人　：" & m_Customer
      ' 申請人名稱(中)(英)
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print m_Customer_CName & "  " & m_Customer_EName
      ' 下一列
      nRow = nRow + 1
      ' 申請地址:
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "英文地址：" & m_Customer_EAddress
      
   Else
      ' 客戶編號
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "客戶編號：" & Mid(m_Customer, 1, 6) & "-" & Mid(m_Customer, 7, 2) & "-" & Mid(m_Customer, 9, 1)
      ' 接洽人
      Printer.CurrentX = (m_LeftMargin + 30) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "接洽人：" & m_Customer_Keep
      Printer.Print "接洽人：" & m_Customer_Keep & IIf(m_Customer_Keep <> "", "　", "　　　　　") & "電話：" & m_Customer_Tel & IIf(m_Customer_Tel <> "", "　", "               ") & "傳真：" & m_Customer_Fax
'      ' 電話
'      Printer.CurrentX = (m_LeftMargin + 50) * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "電話：" & m_Customer_Tel
'      ' 傳真
'      Printer.CurrentX = (m_LeftMargin + 70) * m_CharWidth
'      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'      Printer.Print "傳真：" & m_Customer_Fax
      
      ' 下一列
      nRow = nRow + 1
      ' 申請人
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "申請人　：" & GetCustomerName(m_Customer, "0")
      ' 負責人
      Printer.CurrentX = (m_LeftMargin + 60) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "負責人：" & m_Customer_Main
   
      nRow = nRow + 1
      ' 聯絡地址
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "聯絡地址：" & m_Customer_Adrs & vbTab
   
      nRow = nRow + 1
      ' 申請地址
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "申請地址：" & m_Customer_CAddress
   End If

   'add by nickc 2008/01/18 加印客戶其他資料
   GetOtherCuData m_NP02, m_NP03, m_NP04, m_NP05
   If IsEmpty(m_CuMemo) = False And CheckStr(m_CuMemo) <> "" And IsEmpty(m_SpecCu) = False And CheckStr(m_SpecCu) <> "" Then
      Printer.CurrentX = m_LeftMargin * m_CharWidth + 150
      Printer.CurrentY = (m_TopMargin + 25) * m_CharHeight + 100
      Printer.Print "　　　　　業務備註：" & m_CuMemo
      Printer.CurrentX = m_LeftMargin * m_CharWidth + 150
      Printer.CurrentY = (m_TopMargin + 26) * m_CharHeight + 100
      Printer.Print "　　　　　此為特殊客戶，特殊原因為：" & m_SpecCu
   ElseIf IsEmpty(m_CuMemo) = False And CheckStr(m_CuMemo) <> "" Then
      Printer.CurrentX = m_LeftMargin * m_CharWidth + 150
      Printer.CurrentY = (m_TopMargin + 26) * m_CharHeight + 100
      Printer.Print "　　　　　業務備註：" & m_CuMemo
   ElseIf IsEmpty(m_SpecCu) = False And CheckStr(m_SpecCu) <> "" Then
      Printer.CurrentX = m_LeftMargin * m_CharWidth + 150
      Printer.CurrentY = (m_TopMargin + 26) * m_CharHeight + 100
      Printer.Print "　　　　　此為特殊客戶，特殊原因為：" & m_SpecCu
   End If
   
   If bNewFunc = False Then
      If bContinue Then
         Printer.NewPage
      Else
         Printer.EndDoc
      End If
   End If
End Sub

'Add By Cheng 2002/01/25
' 列印C類接洽記錄單
'Modified by Morgan 2018/10/2 +bolSavePdf:轉pdf
Public Sub PrintCForm(ByVal strCP09 As String, Optional ByVal strMemo As String, Optional ByVal strCP48Desc As String, Optional bolSavePdf As Boolean = False)
'strCP09:總收文號
'列印格式
'收文日期 : Field(0)        承辦人 : Field(1)
'總收文號 : Field(2)        管制人 : Field(3)
'本所期限 : Field(4)        智權人員 : Field(5)
'......
'......
'專利名稱 : Field(10)
'代理人名稱(英文) : Field(11)
'......
'......

Dim m_CharWidth As Long
Dim m_CharHeight As Long
Dim strData As String
Dim nCount As Integer
' 報表左方留白的寬度
Dim m_LeftMargin As Integer
' 報表上方留白的高度
Dim m_TopMargin As Integer
Dim nRow As Integer

Dim strSQL1 As String
Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim jj As Integer
'2010/4/20 ADD BY SONIA
Dim strCustNo As String
Dim varCustNo As Variant
Dim ii As Integer
Dim bolFMP As Boolean 'Add by Morgan 2010/6/18
Dim bolIsSyngentaOneDayCase As Boolean 'Added by Morgan 2015/4/20
'Added by Morgan 2018/10/2
Dim strCOrderPath As String
Dim strPdfName As String
Dim oFileSys As New FileSystemObject
Dim oFile

'設定SQL字串
'Modify by Morgan 2007/4/19 加,CP10,CP43,PA27,PA28,PA29,PA30
'Modify by Morgan 2009/12/1 +P也可印
'2010/4/20 modify by sonia 加cp40,cp41,cp42
'2010/9/27 modify by sonia 加cp24,改變原處分1503加印准駁,
strSQL1 = " Select CP05, S1.ST02, CP09, S2.ST02, CP06, S3.ST02, " & _
         " CP07, CP01||'-'||CP02||'-'||CP03||'-'||CP04 CNo, CP48, decode(cp10,'1503',decode(pa09,'000',CPM03,CPM04)||decode(cp24,'1','－准','2','－駁',cp24),decode(pa09,'000',CPM03,CPM04)) CPM03, " & _
         " PA05, FA05||' '||FA63, FA04, FA06, CU05||' '||CU88, CU04, CU06,PA75,PA26,CP01,CP10,CP43,PA27,PA28,PA29,PA30,CP12,NP23,CP40,CP41,CP42 " & _
         ",CP133,CP134,CP24,CP02,CP03,CP04 From CaseProgress,Patent,Staff S1,Nation,Staff S2,Staff S3,CasePropertyMap,Fagent,Customer,nextprogress" & _
         " Where CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) " & _
         " And CP14=S1.ST01(+) And PA09=NA01(+) And NA16=S2.ST01(+) And CP13=S3.ST01(+) " & _
         " And CP01=CPM01(+) And CP10=CPM02(+) And substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) " & _
         " And substr(PA26,1,8)=CU01(+) And substr(PA26,9,1)=CU02(+) " & _
         " And CP09 = '" & "" & strCP09 & "' AND CP01 in ('FCP','P') and np01(+)=cp09 and np06(+) is null"
'2006/9/11 ADD BY SONIA 加入 FG
'Modify by Morgan 2007/4/19 加,CP10,CP43,SP58,SP59,SP65,SP66
'2010/4/20 modify by sonia 加cp40,cp41,cp42
strSQL1 = strSQL1 & " UNION Select CP05, S1.ST02, CP09, S2.ST02, CP06, S3.ST02, " & _
         " CP07, CP01||'-'||CP02||'-'||CP03||'-'||CP04 CNo, CP48, decode(sp09,'000',CPM03,CPM04) CPM03, " & _
         " SP05, FA05||' '||FA63, FA04, FA06, CU05||' '||CU88, CU04, CU06,SP26,SP08,CP01,CP10,CP43,SP58,SP59,SP65,SP66,CP12,NP23,CP40,CP41,CP42 " & _
         ",CP133,CP134,CP24,CP02,CP03,CP04 From CaseProgress,SERVICEPRACTICE,Staff S1,Nation,Staff S2,Staff S3,CasePropertyMap,Fagent,Customer,nextprogress" & _
         " Where CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) " & _
         " And CP14=S1.ST01(+) And SP09=NA01(+) And NA16=S2.ST01(+) And CP13=S3.ST01(+) " & _
         " And CP01=CPM01(+) And CP10=CPM02(+) And substr(SP26,1,8)=FA01(+) And substr(SP26,9,1)=FA02(+) " & _
         " And substr(SP08,1,8)=CU01(+) And substr(SP08,9,1)=CU02(+) " & _
         " And CP09 = '" & "" & strCP09 & "' AND CP01='FG' and np01(+)=cp09 and np06(+) is null"
If rs1.State <> adStateClosed Then rs1.Close
Set rs1 = Nothing
rs1.CursorLocation = adUseClient
rs1.Open strSQL1, cnnConnection, adOpenStatic, adLockReadOnly
'若有資料則執行列印
If rs1.RecordCount > 0 Then
   'Added by Morgan 2018/10/2
   If bolSavePdf Then
      strCOrderPath = App.path & "\" & strUserNum
      'Added by Lydia 2018/12/17 FCP檔案命名
      If "" & rs1.Fields("cp01") = "FCP" Then
          strPdfName = PUB_FCPCaseNo2FileName(rs1.Fields("cp01"), rs1.Fields("cp02"), rs1.Fields("cp03"), rs1.Fields("cp04")) & "." & rs1.Fields("cp10") & ".ORDER.PDF"
      Else
      'end 2018/12/17
          strPdfName = PUB_CaseNo2FileName(rs1.Fields("cp01"), rs1.Fields("cp02"), rs1.Fields("cp03"), rs1.Fields("cp04")) & "." & rs1.Fields("cp10") & ".ORDER.PDF"
      End If 'Added 2018/12/17
      Load frmPDF
      frmPDF.Show
      frmPDF.StartProcess strCOrderPath, strPdfName
   End If
   'end 2018/10/2

   m_CharWidth = 120
   m_CharHeight = 280
   m_LeftMargin = 6
   m_TopMargin = 4
   Printer.EndDoc 'Add by Morgan 2009/12/31
   '***********  90.11.14   nickc
   Printer.Orientation = 1
   '*****************************
   Printer.PaperSize = vbPRPSA4
   Printer.Font.Size = 12
   Printer.FontName = "細明體"
   'Add By Cheng 2002/02/08
   'For 報表名稱
   Printer.Font.Size = 14

   'Modify by Morgan 2009/12/31 P案紙張不同要上提
   'nRow = 0
   If "" & rs1.Fields("CP01") = "P" Then
      nRow = -1
   Else
      nRow = 0
   End If
   '標題 : C類接洽記錄單
   Printer.CurrentX = (m_LeftMargin + 25) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify by Morgan 2007/4/4
   'Printer.Print "B　類　接　洽　記　錄　單"
   Printer.Print Left(strCP09, 1) & "　類　接　洽　記　錄　單"
   'end 2007/4/4

   'Add By Cheng 2002/02/08
   Printer.Font.Size = 12
   nRow = nRow + 1
   nRow = nRow + 1
   ' 0.收文日期
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify By Cheng 2002/02/08
   '以民國年顯示
'   Printer.Print "收文日期：" & rs1.Fields(0).Value
   Printer.Print "收文日期：" & IIf(IsNull(rs1.Fields(0).Value) Or Len("" & rs1.Fields(0).Value) <= 0, "", _
                              (Left(rs1.Fields(0).Value, 4) - 1911) & "/" & Mid(rs1.Fields(0).Value, 5, 2) & "/" & Right(rs1.Fields(0).Value, 2))
   ' 1.承辦人
   Printer.CurrentX = (m_LeftMargin + 40) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
    'Modify By Cheng 2003/02/20
    '取消判斷是否列印承辦人
'    'Modify By Cheng 2003/02/18
'    'FCP的案件不印承辦人
    'Modified by Morgan 2023/5/24
    'Printer.Print "承辦人　：" & rs1.Fields(1).Value
    PUB_PrintUnicodeText "承辦人　：" & rs1.Fields(1).Value, Printer.CurrentX, Printer.CurrentY, 0
    'end 2023/5/24
'    Printer.Print "承辦人　：" & IIf(rs1("CP01").Value = "FCP", "", rs1.Fields(1).Value)
   nRow = nRow + 1
   ' 2.總收文號
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "總收文號：" & rs1.Fields(2).Value

   'Modify by Morgan 2010/5/5 +控制國外部收文的案件才印
   If Left("" & rs1.Fields("CP12").Value, 1) = "F" Then
      If "" & rs1.Fields("CP01") = "P" Then bolFMP = True 'Add by Morgan 2010/6/18
      ' 3.管制人
      Printer.CurrentX = (m_LeftMargin + 40) * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      'Add By Cheng 2002/02/08
      '若有FC代理人
      'Modified by Lydia 2022/10/21 改成共用模組
'      If Len("" & rs1.Fields(17).Value) > 0 Then
'         strSQL1 = " SELECT ST02 FROM FAGENT,NATION,STAFF WHERE SUBSTR('" & rs1.Fields(17).Value & "',1,8)=FA01(+) AND SUBSTR('" & rs1.Fields(17).Value & "',9,1)=FA02(+) AND FA10=NA01(+) AND NA16=ST01(+)"
'         If Rs2.State <> adStateClosed Then Rs2.Close
'         Set Rs2 = Nothing
'         Rs2.CursorLocation = adUseClient
'         Rs2.Open strSQL1, cnnConnection, adOpenStatic, adLockReadOnly
'         If Rs2.RecordCount > 0 Then
'            Printer.Print "管制人　：" & Rs2.Fields(0).Value
'         End If
'         If Rs2.State <> adStateClosed Then Rs2.Close
'         Set Rs2 = Nothing
'      '若有申請人
'      ElseIf Len("" & rs1.Fields(18).Value) > 0 Then
'         strSQL1 = " SELECT ST02 FROM CUSTOMER,NATION,STAFF WHERE SUBSTR('" & rs1.Fields(18).Value & "',1,8)=CU01(+) AND SUBSTR('" & rs1.Fields(18).Value & "',9,1)=CU02(+) AND CU10=NA01(+) AND NA16=ST01(+)"
'         If Rs2.State <> adStateClosed Then Rs2.Close
'         Set Rs2 = Nothing
'         Rs2.CursorLocation = adUseClient
'         Rs2.Open strSQL1, cnnConnection, adOpenStatic, adLockReadOnly
'         If Rs2.RecordCount > 0 Then
'            Printer.Print "管制人　：" & Rs2.Fields(0).Value
'         End If
'         If Rs2.State <> adStateClosed Then Rs2.Close
'         Set Rs2 = Nothing
'      Else
'         Printer.Print "管制人　："
'      End If
      strSQL1 = PUB_GetFCPHandler(rs1.Fields("cp01"), rs1.Fields("cp02"), rs1.Fields("cp03"), rs1.Fields("cp04"), "" & rs1.Fields("cp10"))
      If strSQL1 <> "" Then
          'Modified by Morgan 2023/5/24
          'Printer.Print "管制人　：" & GetStaffName(strSQL1, True)
          PUB_PrintUnicodeText "管制人　：" & GetStaffName(strSQL1, True), Printer.CurrentX, Printer.CurrentY, 0
          'end 2023/5/24
      Else
          Printer.Print "管制人　："
      End If
      'end 2022/10/21
   End If

   nRow = nRow + 1
   ' 4.本所期限
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "本所期限：" & IIf(IsNull(rs1.Fields(4).Value), "", _
                              (Left(rs1.Fields(4).Value, 4) - 1911) & "/" & Mid(rs1.Fields(4).Value, 5, 2) & "/" & Right(rs1.Fields(4).Value, 2))
   ' 5.智權人員
   Printer.CurrentX = (m_LeftMargin + 40) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modified by Morgan 2023/5/24
   'Printer.Print "智權人員：" & rs1.Fields(5).Value
   PUB_PrintUnicodeText "智權人員：" & rs1.Fields(5).Value, Printer.CurrentX, Printer.CurrentY, 0
   'end 2023/5/24

   nRow = nRow + 1
   ' 6.法定期限
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "法定期限：" & IIf(IsNull(rs1.Fields(6).Value), "", _
                              (Left(rs1.Fields(6).Value, 4) - 1911) & "/" & Mid(rs1.Fields(6).Value, 5, 2) & "/" & Right(rs1.Fields(6).Value, 2))
   ' 7.本所案號
   Printer.CurrentX = (m_LeftMargin + 40) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "本所案號：" & rs1.Fields(7).Value

   nRow = nRow + 1
   ' 8.承辦期限
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modify By Cheng 2002/02/08
   '以民國年顯示
'   Printer.Print "承辦期限：" & rs1.Fields(8).Value

   'Add by Morgan 2010/6/18
   If bolFMP And strCP48Desc = "" Then
      If rs1.Fields("cp133") > 0 And rs1.Fields("cp134") > 0 Then
         PUB_GetFmpCP48 rs1.Fields("cp05"), rs1.Fields("cp06"), rs1.Fields("cp07"), rs1.Fields("cp133"), rs1.Fields("cp134"), , strCP48Desc
      End If
   End If

   Printer.Print "承辦期限：" & IIf(IsNull(rs1.Fields(8).Value) Or Len("" & rs1.Fields(8).Value) <= 0, "", _
                              (Left(rs1.Fields(8).Value, 4) - 1911) & "/" & Mid(rs1.Fields(8).Value, 5, 2) & "/" & Right(rs1.Fields(8).Value, 2)) & IIf(strCP48Desc = "", "", "(" & strCP48Desc & ")")

   ' 9.案件性質
   Printer.CurrentX = (m_LeftMargin + 40) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "案件性質：" & rs1.Fields(9).Value

   nRow = nRow + 1
   'Add by Morgan 2009/12/28 FMP+約定期限
   If Left("" & rs1.Fields("CP12").Value, 1) = "F" Then
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.FontBold = True
      Printer.Print "約定期限：" & ChangeWStringToTDateString("" & rs1.Fields("NP23").Value)
      Printer.FontBold = False
   End If

   nRow = nRow + 1
   ' 10.專利名稱(中文)
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modified by Morgan 2023/5/24
   'Printer.Print "專利名稱(中文)：" & rs1.Fields(10).Value
   PUB_PrintUnicodeText "專利名稱(中文)：" & rs1.Fields(10).Value, Printer.CurrentX, Printer.CurrentY, 0
   'end 2023/5/24

   nRow = nRow + 1
   nRow = nRow + 1
   ' 11.代理人名稱(英文)
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "代理人名稱(英文)：" & rs1.Fields(11).Value

   nRow = nRow + 1
   ' 12.代理人名稱(中文)
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modified by Morgan 2023/5/24
   'Printer.Print "　　　　　(中文)：" & rs1.Fields(12).Value
   PUB_PrintUnicodeText "　　　　　(中文)：" & rs1.Fields(12).Value, Printer.CurrentX, Printer.CurrentY, 0
   'end 2023/5/24
   
   nRow = nRow + 1
   ' 13.代理人名稱(日文)
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   'Modified by Morgan 2023/5/24
   'Printer.Print "　　　　　(日文)：" & rs1.Fields(13).Value
   PUB_PrintUnicodeText "　　　　　(日文)：" & rs1.Fields(13).Value, Printer.CurrentX, Printer.CurrentY, 0
   'end 2023/5/24
   
   ' 下一列
   'Modify by Morgan 2009/12/31 P案紙張可印空間不夠，申請人不印
   If "" & rs1.Fields("CP01") <> "P" Then
      nRow = nRow + 1
      nRow = nRow + 1
      ' 14.申請人名稱(英文)
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "申請人名稱(英文)：" & rs1.Fields(14).Value

      nRow = nRow + 1
      ' 15.申請人名稱(中文)
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      'Modified by Morgan 2023/5/24
      'Printer.Print "　　　　　(中文)：" & rs1.Fields(15).Value
      PUB_PrintUnicodeText "　　　　　(中文)：" & rs1.Fields(15).Value, Printer.CurrentX, Printer.CurrentY, 0
      'end 2023/5/24

      nRow = nRow + 1
      ' 16.申請人名稱(日文)
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      'Modified by Morgan 2023/5/24
      'Printer.Print "　　　　　(日文)：" & rs1.Fields(16).Value
      PUB_PrintUnicodeText "　　　　　(日文)：" & rs1.Fields(16).Value, Printer.CurrentX, Printer.CurrentY, 0
      'end 2023/5/24
   End If

   'Add by Morgan 2007/4/4
   'Modify by Morgan 2009/12/1 +FMP案也要
   If Left("" & rs1.Fields(7).Value, 3) = "FCP" Or Left("" & rs1.Fields(7).Value, 2) = "FG" Or Left("" & rs1.Fields("CP12").Value, 1) = "F" Then
      'Add by Morgan 2007/4/19
      '申請人2~5名稱
      'Modify by Morgan 2009/12/31 P案紙張可印空間不夠，申請人不印
      If "" & rs1.Fields("CP01") <> "P" Then
         For jj = 27 To 30
            If "" & rs1.Fields("PA" & jj) <> "" Then
               strExc(0) = "select  CU05||' '||CU88, CU04, CU06 from customer where cu01||cu02='" & rs1.Fields("PA" & jj) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  nRow = nRow + 1
                  '英
                  Printer.CurrentX = m_LeftMargin * m_CharWidth
                  Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
                  Printer.Print "　　　　　(英文)：" & RsTemp.Fields(0).Value

                  nRow = nRow + 1
                  '中
                  Printer.CurrentX = m_LeftMargin * m_CharWidth
                  Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
                  'Modified by Morgan 2023/5/24
                  'Printer.Print "　　　　　(中文)：" & RsTemp.Fields(1).Value
                  PUB_PrintUnicodeText "　　　　　(中文)：" & RsTemp.Fields(1).Value, Printer.CurrentX, Printer.CurrentY, 0
                  'end 2023/5/24

                  nRow = nRow + 1
                  '日
                  Printer.CurrentX = m_LeftMargin * m_CharWidth
                  Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
                  'Modified by Morgan 2023/5/24
                  'Printer.Print "　　　　　(日文)：" & RsTemp.Fields(2).Value
                  PUB_PrintUnicodeText "　　　　　(日文)：" & RsTemp.Fields(2).Value, Printer.CurrentX, Printer.CurrentY, 0
                  'end 2023/5/24
               End If
            End If
         Next
         'end 2007/4/19
      End If
      nRow = nRow + 2
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "備註："

      'Add by Morgan 2007/4/19 核駁函接洽單上加印備註
      If strMemo = "" Then
         '2010/9/27 modify by sonia 改變原處分-核駁也要印
         'If (rs1("CP10") = "1002" Or rs1("CP10") = "1007") And "" & rs1("CP43") <> "" Then
         If (rs1("CP10") = "1002" Or rs1("CP10") = "1007" Or (rs1("CP10") = "1503" And "" & rs1("CP24") = "2")) And "" & rs1("CP43") <> "" Then
            'Modify by Morgan 2009/12/1 改抓下一程序的案件性質判斷
            'strExc(0) = "SELECT CP10 FROM CASEPROGRESS WHERE CP09='" & rs1("CP43") & "'"
            'intI = 1
            'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            'If intI = 1 Then
            '   Select Case RsTemp.Fields(0)
            '      'Modify by Morgan 2009/10/30
            '      Case "107", "804"
            '         strMemo = "請客戶提供重新簽署的委任書以俾辦理訴願程序"
            '      Case "501"
            '         strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理行政訴訟程序"
            '      Case "503"
            '         If rs1("CP10") = "1002" Then
            '            strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理行政訴訟上訴程序"
            '         Else
            '            strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理抗告程序"
            '         End If
            '      Case "507"
            '         If rs1("CP10") = "1002" Then
            '            strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理行政再審程序"
            '         Else
            '            strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理抗告程序"
            '         End If
            '   End Select
            'End If
            strExc(0) = "SELECT NP07 FROM NEXTPROGRESS WHERE NP01='" & strCP09 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Select Case RsTemp.Fields(0)
                  Case "501" '訴願
                     strMemo = "請客戶提供重新簽署的委任書以俾辦理訴願程序"
                  Case "503" '行政訴訟
                     strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理行政訴訟程序"
                  Case "504" '行政再審
                     strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理行政再審程序"
                  Case "507" '行政訴訟上訴
                     strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理行政訴訟上訴程序"
                  Case "509" '抗告
                     strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理抗告程序"
               End Select
            End If
            'END 2009/12/1
         End If
      End If
      'end 2007/4/19

      'Added by Morgan 2014/2/17
      'Removed by Morgan 2020/9/15 該案已至年費階段並閉卷
      'If rs1("CP10") = "1202" And rs1(7) = "FCP-043077-0-00" Then
      '   If strMemo <> "" Then strMemo = strMemo & vbCrLf
      '   strMemo = strMemo & "工程師通知分割時間限制"
      'End If
      'end 2022/01/21
      'end 2014/2/17

       'Added by Morgan 2015/4/20
      If InStr("Y4830900,Y4830901,Y4830902,Y4830903,Y4830904,Y4830905,Y4830908,Y5132600", Left("" & rs1("pa75"), 8)) > 0 Then
         'Modified by Lydia 2022/01/21 改成共用模組ChkSyngentaNP=>Pub_ChkSyngentaNP
         bolIsSyngentaOneDayCase = Pub_ChkSyngentaNP(strCP09)
         If bolIsSyngentaOneDayCase = True Then
            If strMemo <> "" Then strMemo = strMemo & vbCrLf
            strMemo = strMemo & "收到舉發案訴願或任何法律行動，請於收到的24小時內告知先正達。"
         End If
      End If
      'end 2015/4/20

      If strMemo <> "" Then
         Printer.FontBold = True
         Printer.CurrentX = m_LeftMargin * m_CharWidth + Printer.TextWidth("備註：")
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         'Modified by Morgan 2013/12/11
         'Printer.Print strMemo
         strExc(1) = strMemo
         Do While strExc(1) <> ""
            intI = InStr(strExc(1), vbCrLf)
            If intI > 0 Then
               strExc(0) = Left(strExc(1), intI - 1)
               strExc(1) = Mid(strExc(1), intI + 2)
               Printer.CurrentX = m_LeftMargin * m_CharWidth + Printer.TextWidth("備註：")
               Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
               'Modified by Morgan 2023/5/24
               'Printer.Print strExc(0)
               PUB_PrintUnicodeText strExc(0), Printer.CurrentX, Printer.CurrentY, 0
               'end 2023/5/24
               
               nRow = nRow + 1
            Else
               Printer.CurrentX = m_LeftMargin * m_CharWidth + Printer.TextWidth("備註：")
               Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
               'Modified by Morgan 2023/5/24
               'Printer.Print strExc(1)
               PUB_PrintUnicodeText strExc(1), Printer.CurrentX, Printer.CurrentY, 0
               'end 2023/5/24
               strExc(1) = ""
            End If
         Loop
         'end 2013/12/11
         Printer.FontBold = False

      End If


      'Added by Morgan 2013/8/28
      'Modified by Morgan 2015/4/20 排除先正達承辦期限為2個工作天(24Hr)的來函
      'If (rs1("cp01") = "P" Or rs1("cp01") = "FCP") And (rs1("cp10") = "1202" Or rs1("cp10") = "1227" Or rs1("cp10") = "1002") Then
      'Modified by Lydia 2020/03/06 +被舉發理由1802
      'Modified by Lydia 2021/05/18 +通知補正1201
      If bolIsSyngentaOneDayCase = False And (rs1("cp01") = "P" Or rs1("cp01") = "FCP") And (rs1("cp10") = "1202" Or rs1("cp10") = "1227" Or rs1("cp10") = "1002" Or rs1("cp10") = "1802" Or rs1("cp10") = "1201") Then
      'end 2015/4/20
'Modified by Morgan 2013/9/18 改呼叫共用函數
'         If Left("" & rs1("pa75"), 6) = "Y34210" And Left("" & rs1("pa26"), 6) = "X51446" Then
'            If strMemo <> "" Then nRow = nRow + 1
'            Printer.FontBold = True
'            Printer.CurrentX = m_LeftMargin * m_CharWidth + Printer.TextWidth("備註：")
'            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'            Printer.Print "本程序分二階段報告： 第1 階段寄OA原文及其日譯文+引証案, 第2 階段提供本所意見"
'            Printer.FontBold = False
'         ElseIf Left("" & rs1("pa75"), 6) = "Y47453" And Left("" & rs1("pa26"), 6) = "X55778" Then
'            If strMemo <> "" Then nRow = nRow + 1
'            Printer.FontBold = True
'            Printer.CurrentX = m_LeftMargin * m_CharWidth + Printer.TextWidth("備註：")
'            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'            Printer.Print "本程序分二階段報告：第1階段：收文日1週內提供OA原文"
'            nRow = nRow + 1
'            Printer.CurrentX = m_LeftMargin * m_CharWidth + Printer.TextWidth("備註：本程序分二階段報告：")
'            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'            Printer.Print "第2階段：收文日1個月內提供翻譯文＋本所comment"
'            Printer.FontBold = False
'         'Modified by Morgan 2013/9/6 + Y20065 --邱子瑜
'         ElseIf Left("" & rs1("pa75"), 6) = "Y20065" Then
'            If strMemo <> "" Then nRow = nRow + 1
'            Printer.FontBold = True
'            Printer.CurrentX = m_LeftMargin * m_CharWidth + Printer.TextWidth("備註：")
'            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'            Printer.Print "本程序分二階段報告：第1 階段寄OA原文及其日譯文+檢索報告英譯文 (C類收文15日內),"
'            nRow = nRow + 1
'            Printer.CurrentX = m_LeftMargin * m_CharWidth + Printer.TextWidth("備註：本程序分二階段報告：")
'            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'            Printer.Print "第2 階段提供本所報告及意見 (C類收文35日內)"
'            Printer.FontBold = False
'         End If
         strExc(1) = "": strExc(2) = "": strExc(3) = ""
         'Add by Lydia 2014/12/3 核駁及審查意見通知函備註
         'If PUB_ChkAutoRec(rs1("cp01"), "" & rs1("pa75"), "" & rs1("pa26"), , , , strExc(1), True, "" & rs1("pa27"), "" & rs1("pa28"), "" & rs1("pa29"), "" & rs1("pa30")) = True Then
          strExc(9) = Replace(rs1("cno"), "-", "")
          'Modified by Lydia 2021/11/05 模組修改
          'strExc(2) = "": strExc(3) = "": strExc(4) = "": strExc(5) = ""
          'If Not IsNull(rs1("pa27")) Then strExc(2) = ChangeCustomerL(rs1("pa27"))
          'If Not IsNull(rs1("pa28")) Then strExc(3) = ChangeCustomerL(rs1("pa28"))
          'If Not IsNull(rs1("pa29")) Then strExc(4) = ChangeCustomerL(rs1("pa29"))
          'If Not IsNull(rs1("pa30")) Then strExc(5) = ChangeCustomerL(rs1("pa30"))
          'strExc(1) = PUB_GetIncomMemoNew(strExc(9), rs1("cp01"), "", ChangeCustomerL(rs1("pa75")), ChangeCustomerL(rs1("pa26")), "Y", , _
          '            , strExc(2), strExc(3), strExc(4), strExc(5))
          'strExc(2) = "": strExc(3) = "": strExc(4) = "": strExc(5) = ""
          strExc(1) = PUB_GetIncomMemoNew(strExc(9), rs1.Fields("cp01"), ChangeCustomerL(rs1.Fields("pa75")), ChangeCustomerL(rs1.Fields("pa26")) & "," & ChangeCustomerL(rs1.Fields("pa27")) & "," & ChangeCustomerL(rs1.Fields("pa28")) & "," & ChangeCustomerL(rs1.Fields("pa29")) & "," & ChangeCustomerL(rs1.Fields("pa30")), "Y", "", "" & rs1.Fields("cp10"))
          'end 2021/11/05

          If Len(strExc(1)) > 0 Then
          'end  'Add by Lydia 2014/12/3
            If strExc(1) <> "" Then
               Printer.FontBold = True
               If strMemo <> "" Then nRow = nRow + 1
'Removed by Morgan 2014/3/3
'               intI = InStr(strExc(1), "：")
'               If intI > 0 And InStr(strExc(1), vbCrLf) <> intI + 1 Then
'                  strExc(2) = Left(strExc(1), intI)
'               End If
'end 2014/3/3
               Do While strExc(1) <> ""
                  intI = InStr(strExc(1), vbCrLf)
                  If intI > 0 Then
                     strExc(0) = Left(strExc(1), intI - 1)
                     strExc(1) = Mid(strExc(1), intI + 2)
                     Printer.CurrentX = m_LeftMargin * m_CharWidth + Printer.TextWidth("備註：" & strExc(3))
                     Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
                     'Modified by Morgan 2023/5/24
                     'Printer.Print strExc(0)
                     PUB_PrintUnicodeText strExc(0), Printer.CurrentX, Printer.CurrentY, 0
                     'end 2023/5/24
                     strExc(3) = strExc(2)
                     nRow = nRow + 1
                  Else
                     Printer.CurrentX = m_LeftMargin * m_CharWidth + Printer.TextWidth("備註：" & strExc(3))
                     Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
                     'Modified by Morgan 2023/5/24
                     'Printer.Print strExc(1)
                     PUB_PrintUnicodeText strExc(1), Printer.CurrentX, Printer.CurrentY, 0
                     'end 2023/5/24
                     strExc(1) = ""
                  End If
               Loop
               Printer.FontBold = False
            End If
         End If
'end 2013/9/18

      End If
      'end 2013/8/28

      '2010/4/20 ADD BY SONIA 對造若為本所客戶印出基本資料
      If PUB_ChkCustNameExist("" & rs1.Fields("CP40"), "" & rs1.Fields("CP41"), "" & rs1.Fields("CP42"), False, strCustNo) = True Then
         strMemo = ""
         nRow = nRow + 2
         Printer.CurrentX = m_LeftMargin * m_CharWidth
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print "對造為本所客戶："
         varCustNo = Split(strCustNo, ",")
         For ii = LBound(varCustNo) To UBound(varCustNo)
            '2011/8/19 MODIFY BY SONIA 因varCustNo含(業務區智權人員),故只取前9碼
            'strExc(0) = "select cu01||cu02,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,cu10,cu11 from customer where CU01||CU02='" & varCustNo(ii) & "'"
            strExc(0) = "select cu01||cu02,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,cu10,NVL(cu11,'無') CU11 from customer where CU01||CU02='" & Left(varCustNo(ii), 9) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Select Case RsTemp.Fields("cu10")
                  Case "001", "002", "003", "004", "005", "006", "007", "008", "013", "020", "044"
                     If Not IsNull(RsTemp.Fields("cu04")) Then
                        strMemo = RsTemp.Fields("cu04")
                     ElseIf Not IsNull(RsTemp.Fields("cu05")) Then
                        strMemo = RsTemp.Fields("cu05")
                     ElseIf Not IsNull(RsTemp.Fields("cu06")) Then
                        strMemo = RsTemp.Fields("cu06")
                     End If
                  Case Else
                     If Not IsNull(RsTemp.Fields("cu05")) Then
                        strMemo = RsTemp.Fields("cu05")
                     ElseIf Not IsNull(RsTemp.Fields("cu04")) Then
                        strMemo = RsTemp.Fields("cu04")
                     ElseIf Not IsNull(RsTemp.Fields("cu06")) Then
                        strMemo = RsTemp.Fields("cu06")
                     End If
               End Select
               nRow = nRow + 1
               Printer.CurrentX = m_LeftMargin * m_CharWidth
               Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
               'Modified by Morgan 2023/5/24
               'Printer.Print "　客戶編號：" & varCustNo(ii) & "　名稱：" & strMemo & "　ＩＤ：" & "" & RsTemp.Fields("cu11")
               PUB_PrintUnicodeText "　客戶編號：" & varCustNo(ii) & "　名稱：" & strMemo & "　ＩＤ：" & "" & RsTemp.Fields("cu11"), Printer.CurrentX, Printer.CurrentY, 0
               'end 2023/5/24
            End If
         Next ii
      End If
      '2010/4/20 END
   End If

   Printer.EndDoc

   'Added by Morgan 2018/10/2
   If bolSavePdf Then
      frmPDF.EndtProcess
      Unload frmPDF
      If Dir(strCOrderPath & "\" & strPdfName) <> "" Then
         'Added by Lydia 2018/12/17 FCP案C類接洽單同時列印並且上傳到卷宗區(先列印,再上傳,最後刪除檔案)
On Error GoTo ErrorHand01
         If "" & rs1.Fields("cp01") = "FCP" Then
             'Modified by Lydia 2019/01/04 調整順序,先上傳後可能檔案會比較完整,再列印,最後刪除檔案
             '(FCP-59024)因為Teresa在電子公文來函輸入核駁,在後續的核駁函自動上傳C類接洽單和列印出現"Rdrcef.exe"的錯誤,
             '又重複電子公文來函輸入,產生CA8000016(卷宗區有電子公文)和CA8000017(卷宗區有Order.PDF),後來做人工刪CA8000017並且將公文.PDF歸入CA8000016
             'PUB_PrintPDF strCOrderPath & "\" & strPdfName
             'Set oFile = oFileSys.GetFile(strCOrderPath & "\" & strPdfName)
             'If SaveAttFile_PDF(strCP09, strCOrderPath & "\" & strPdfName, strPdfName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , , True) = True Then
             '      Kill strCOrderPath & "\" & strPdfName
             'End If
             Set oFile = oFileSys.GetFile(strCOrderPath & "\" & strPdfName)
             If SaveAttFile_PDF(strCP09, strCOrderPath & "\" & strPdfName, strPdfName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , , True) = True Then
                   PUB_PrintPDF strCOrderPath & "\" & strPdfName
             End If
             Kill strCOrderPath & "\" & strPdfName
             'end 2019/01/04
         Else
         'end 2018/12/17
             Set oFile = oFileSys.GetFile(strCOrderPath & "\" & strPdfName)
             SaveAttFile_PDF strCP09, strCOrderPath & "\" & strPdfName, strPdfName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , , True
         End If  'Added by Lydia 2018/12/17
      End If
   End If
   'end 2018/10/2

End If
If rs1.State <> adStateClosed Then rs1.Close
Set rs1 = Nothing
'Added by Morgan 2018/10/2
Set oFile = Nothing
Set oFileSys = Nothing

'Added by Lydia 2018/12/17
    Exit Sub

ErrorHand01:
    If Err.Number <> 0 Then
        Resume Next
    End If
'end 2018/12/17
End Sub

'92.1.28 ADD BY SONIA
' 列印CFP案件C類接洽記錄單
'Modified by Morgan 2018/10/2 +bolSavePdf:轉pdf
Public Sub PrintCFForm(ByVal strCP09 As String, Optional strMemo As String, Optional bolSavePdf As Boolean = False)
'strCP09:總收文號
'列印格式
'收文日期 : Field(0)        承辦人   : Field(1)
'總收文號 : Field(2)        申請國家 : Field(19)
'本所期限 : Field(4)        智權人員   : Field(5)
'......
'專利名稱 : Field(10),Field(20),Field(21)
'......
'申請人名稱 : Field(15)
'......

Dim m_CharWidth As Long
Dim m_CharHeight As Long
Dim strData As String
Dim nCount As Integer
' 報表左方留白的寬度
Dim m_LeftMargin As Integer
' 報表上方留白的高度
Dim m_TopMargin As Integer
Dim nRow As Integer

Dim strSQL1 As String
Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset

Dim m_CP64 As String  '2008/11/28 ADD BY SONIA 取得約定期限
'Added by Lydia 2015/12/03
Dim m_IDSmemo As String
Dim tmpArr As Variant
'Added by Morgan 2018/10/2
Dim strCOrderPath As String
Dim strPdfName As String
Dim oFileSys As New FileSystemObject
Dim oFile
'Added by Morgan 2020/12/10
Dim ii As Integer
Dim arrMemo() As String


   'Added by Lydia 2021/08/26 預設頁面最大列數
   Dim m_MaxRow As Integer
   m_MaxRow = 52
   'end 2021/08/26
   
'Modify by Morgan 2005/2/3 加欄位 CP01,CP02,CP03,CP04,CP10
'Modify by Sindy 2011/1/19 +PA47
'Modified by Morgan 2020/12/10 +CP12,PA27,PA28,PA29,PA30
'Modified by Morgan 2021/9/2
strSQL1 = " Select CP05, S1.ST02, CP09, S2.ST02, CP06, S3.ST02, " & _
         " CP07, CP01||'-'||CP02||'-'||CP03||'-'||CP04, CP48, CPM03, " & _
         " PA05, FA05||' '||FA63, FA04, FA06, CU05||' '||CU88, CU04, CU06,PA75,PA26,NA03,PA06,PA07,CP01,CP02,CP03,CP04,CP10,CP64 " & _
         ",PA09,PA47,CP12,PA27,PA28,PA29,PA30 From CaseProgress,Patent,Staff S1,Nation,Staff S2,Staff S3,CasePropertyMap,Fagent,Customer  " & _
         " Where CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) " & _
         " And CP14=S1.ST01(+) And PA09=NA01(+) And NA16=S2.ST01(+) And CP13=S3.ST01(+) " & _
         " And CP01=CPM01(+) And CP10=CPM02(+) And substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) " & _
         " And substr(PA26,1,8)=CU01(+) And substr(PA26,9,1)=CU02(+) " & _
         " And CP09 = '" & "" & strCP09 & "'"
If rs1.State <> adStateClosed Then rs1.Close
Set rs1 = Nothing
rs1.CursorLocation = adUseClient
rs1.Open strSQL1, cnnConnection, adOpenStatic, adLockReadOnly
'若有資料則執行列印
If rs1.RecordCount > 0 Then
   'Added by Morgan 2018/10/2
   If bolSavePdf Then
      strCOrderPath = App.path & "\" & strUserNum
      strPdfName = PUB_CaseNo2FileName(rs1.Fields("cp01"), rs1.Fields("cp02"), rs1.Fields("cp03"), rs1.Fields("cp04")) & "." & rs1.Fields("cp10") & ".ORDER.PDF"
      Load frmPDF
      frmPDF.Show
      frmPDF.StartProcess strCOrderPath, strPdfName
   End If
   'end 2018/10/2

   m_CharWidth = 120
   m_CharHeight = 260
   m_LeftMargin = 6
   m_TopMargin = 4
   '***********  90.11.14   nickc
   Printer.Orientation = 1
   '*****************************
   Printer.PaperSize = vbPRPSA4
   Printer.Font.Size = 12
   Printer.FontName = "細明體"
   
   'For 報表名稱
   Printer.Font.Size = 14
   
   nRow = 0
   '標題 : C類接洽記錄單
   Printer.CurrentX = (m_LeftMargin + 25) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "C　類　接　洽　記　錄　單"
      
   m_CharHeight = 280
   Printer.Font.Size = 12
   
   nRow = nRow + 1
   nRow = nRow + 1
   ' 0.收文日期
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   '以民國年顯示
'   Printer.Print "收文日期：" & rs1.Fields(0).Value
   Printer.Print "收文日期：" & IIf(IsNull(rs1.Fields(0).Value) Or Len("" & rs1.Fields(0).Value) <= 0, "", _
                              (Left(rs1.Fields(0).Value, 4) - 1911) & "/" & Mid(rs1.Fields(0).Value, 5, 2) & "/" & Right(rs1.Fields(0).Value, 2))
   ' 1.承辦人
   Printer.CurrentX = (m_LeftMargin + 40) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "承辦人　：" & rs1.Fields(1).Value
   
   nRow = nRow + 1
   ' 2.總收文號
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "總收文號：" & rs1.Fields(2).Value
   ' 3.申請國家
   Printer.CurrentX = (m_LeftMargin + 40) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "申請國家：" & rs1.Fields(19).Value
   
   nRow = nRow + 1
   ' 4.本所期限
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "本所期限：" & IIf(IsNull(rs1.Fields(4).Value), "", _
                              (Left(rs1.Fields(4).Value, 4) - 1911) & "/" & Mid(rs1.Fields(4).Value, 5, 2) & "/" & Right(rs1.Fields(4).Value, 2))
   ' 5.智權人員
   Printer.CurrentX = (m_LeftMargin + 40) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "智權人員：" & rs1.Fields(5).Value
   
   nRow = nRow + 1
   ' 6.法定期限
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "法定期限：" & IIf(IsNull(rs1.Fields(6).Value), "", _
                              (Left(rs1.Fields(6).Value, 4) - 1911) & "/" & Mid(rs1.Fields(6).Value, 5, 2) & "/" & Right(rs1.Fields(6).Value, 2))
   ' 7.本所案號
   Printer.CurrentX = (m_LeftMargin + 40) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "本所案號：" & rs1.Fields(7).Value
   
   nRow = nRow + 1
   ' 8.承辦期限
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   '以民國年顯示
'   Printer.Print "承辦期限：" & rs1.Fields(8).Value
   Printer.Print "承辦期限：" & IIf(IsNull(rs1.Fields(8).Value) Or Len("" & rs1.Fields(8).Value) <= 0, "", _
                              (Left(rs1.Fields(8).Value, 4) - 1911) & "/" & Mid(rs1.Fields(8).Value, 5, 2) & "/" & Right(rs1.Fields(8).Value, 2))
   
   ' 9.案件性質
   Printer.CurrentX = (m_LeftMargin + 40) * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "案件性質：" & rs1.Fields(9).Value
   
   '2008/11/28 ADD BY SONIA CFP案CP64有 約定期限： 者才印
   'Modified by Morgan 2021/9/2 來函約定期限改不放備註都寫下一程序的約定期限欄位
   'If InStr("" & rs1.Fields("CP64").Value, "約定期限：") > 0 Then 'Add by Morgan 2011/8/8 有輸才抓
   '   m_CP64 = Mid(rs1.Fields("CP64").Value, (InStr(rs1.Fields("CP64").Value, "約定期限：") + 5))
   'End If
   intI = 1
   strSql = "SELECT np23 FROM nextprogress WHERE np01='" & strCP09 & "' and np23 is not null"
   Set Rs2 = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      m_CP64 = Rs2("np23")
   End If
   'end 2021/9/2
   
   nRow = nRow + 1
   ' 約定期限
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "約定期限：" & IIf(IsNull(m_CP64) Or Len("" & m_CP64) <= 0, "", ChangeTStringToTDateString(m_CP64))
   '2008/11/28 END
   'Modify by Sindy 2011/1/19 分所案號
   If Not IsNull(rs1.Fields("PA47").Value) Then
      If rs1.Fields("PA47").Value <> "" Then
         Printer.CurrentX = (m_LeftMargin + 40) * m_CharWidth
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print "分所案號：" & rs1.Fields("PA47").Value
      End If
   End If
   '2011/1/19 End
   
   nRow = nRow + 1
   nRow = nRow + 1
   ' 10.專利名稱(中文)
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "專利　名稱(中文)：" & rs1.Fields(10).Value
   nRow = nRow + 1
   ' 20.專利名稱(英文)
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "　　　　　(英文)：" & rs1.Fields(20).Value
   nRow = nRow + 1
   ' 21.專利名稱(日文)
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "　　　　　(日文)：" & rs1.Fields(21).Value
   
   nRow = nRow + 1
   nRow = nRow + 1
   ' 11.申請人名稱(中文)
   Printer.CurrentX = m_LeftMargin * m_CharWidth
   Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
   Printer.Print "申請人名稱(中文)：" & rs1.Fields(15).Value
   
   'Add by Morgan 2008/5/6 加印備註(Ex. 大個體,費用(點數P))
   If strMemo <> "" Then
      nRow = nRow + 1
      nRow = nRow + 1
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print strMemo
   End If
   'end 2008/5/6
   
   If rs1.Fields(9).Value = "其他來函" Then
      nRow = nRow + 1
      nRow = nRow + 1
      '處理情形 :
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "請工程師勾選：ˍˍˍ１. 退程序歸檔"
      nRow = nRow + 1
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "　　　　　　　ˍˍˍ２. 直接回覆代理人（先退程序做內部收文：回覆代理人）"
      nRow = nRow + 1
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "　　　　　　　ˍˍˍ３. 通知客戶（請將次頁之接洽記錄單一併交予智權人員，以便智權人員收文下一程序）"
      nRow = nRow + 1
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "　　　　　　　ˍˍˍ４. 其他（請說明）"
      
      nRow = nRow + 1
      nRow = nRow + 1
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "工程師意見："
      
   'Add by Morgan 2005/2/3 (通知提供前案,核駁,檢索報告)
   'Modify by Morgan 2006/4/10 加最終核駁(1006),建議性處分書(1220)
   'Mofified by Morgan 2021/2/26 +審查意見書(1202) --林禧佩 Ex:CFP-029978
   ElseIf rs1.Fields("CP10") = "1205" Or rs1.Fields("CP10") = "1002" Or rs1.Fields("CP10") = "1006" Or rs1.Fields("CP10") = "1202" Or rs1.Fields("CP10") = "1209" Or rs1.Fields("CP10") = "1220" Then
      '實審收文情形
      strExc(1) = "" & rs1.Fields("CP01")
      strExc(2) = "" & rs1.Fields("CP02")
      strExc(3) = "" & rs1.Fields("CP03")
      strExc(4) = "" & rs1.Fields("CP04")
      'Modify by Morgan 2010/6/29 新加坡 014 也要判斷[申請檢索報告 421]及[合併檢索及實審 427]
      If rs1.Fields("pa09") = "014" Then
         strExc(0) = Get416State014(strExc)
      Else
      'end 2010/6/29
         strExc(0) = Get416State(strExc)
      End If
      If strExc(0) <> "" Then
         nRow = nRow + 1
         nRow = nRow + 1
         Printer.CurrentX = m_LeftMargin * m_CharWidth
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print strExc(0)
      End If
      'add by sonia 2015/9/1 日本011最終核駁1006要加一句話CFP-025108
      If rs1.Fields("pa09") = "011" And rs1.Fields("CP10") = "1006" Then
         nRow = nRow + 1
         Printer.CurrentX = m_LeftMargin * m_CharWidth
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print "最終核駁發文請一併附委任狀給客戶"
      End If
      'end 2015/9/1
      '多國案資料
      Set AdoRecordSet3 = GetRelCaseData(strExc, intI)
      If intI = 1 Then
         nRow = nRow + 1
         nRow = nRow + 1
         Printer.CurrentX = m_LeftMargin * m_CharWidth
         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
         Printer.Print "其他多國案："
         With AdoRecordSet3
            Do While Not .EOF
               'Modify by Morgan 2010/2/9 美國案准駁改抓最近實質相關程序名稱及收文日
               If IsNull(.Fields("PA57")) And .Fields("PA09") = "101" Then
                  strExc(0) = .Fields(0) & Space(2) & .Fields(1) & Space(2) & .Fields(2) & Space(2) & GetLatestCP(.Fields("pa01"), .Fields("pa02"), .Fields("pa03"), .Fields("pa04"))
                  'Added by Lydia 2015/12/03 CFP美國案之核駁、檢索報告多列出IDS之訊息
                  If rs1.Fields("CP10") = "1002" Or rs1.Fields("CP10") = "1209" Then
                      intI = 1
                      strSql = "SELECT CP27,CP09,CP64 FROM CASEPROGRESS WHERE CP01='" & .Fields("pa01") & "' AND CP02='" & .Fields("pa02") & "' AND CP03='" & .Fields("pa03") & "' AND CP04='" & .Fields("pa04") & "' " & _
                               "AND CP27>0 AND CP10='214' order by 1 "
                      Set Rs2 = ClsLawReadRstMsg(intI, strSql)
                      If intI = 1 Then
                         Rs2.MoveFirst
                         Do While Not Rs2.EOF
                            'IDS-案號
                            'm_IDSmemo = m_IDSmemo & convForm(IIf("" & .Fields("pa03") & .Fields("pa04") <> "000", .Fields("pa01") & .Fields("pa02") & .Fields("pa03") & .Fields("pa04"), .Fields("pa01") & .Fields("pa02")), 14)
                            If Rs2.AbsolutePosition = 1 Then
                               m_IDSmemo = m_IDSmemo & convForm(.Fields("pa01") & "-" & .Fields("pa02") & "-" & .Fields("pa03") & "-" & .Fields("pa04"), 20)
                            Else
                               m_IDSmemo = m_IDSmemo & Space(20)
                            End If
                            'IDS-發文日
                            m_IDSmemo = m_IDSmemo & convForm(ChangeWStringToWDateString(Rs2.Fields("CP27")), 14)
                            'IDS-備註
                            m_IDSmemo = m_IDSmemo & PUB_GetMidString(Rs2.Fields("CP64"), "揭露項目", ";")
                            m_IDSmemo = m_IDSmemo & "|"
                            Rs2.MoveNext
                         Loop
                      End If
                  End If
                  'end 2015/12/03
               Else
                  strExc(0) = .Fields(0) & Space(2) & .Fields(1) & Space(2) & .Fields(2) & Space(2) & .Fields(3)
                  'Added by Morgan 2018/7/16 若無狀態時則顯示最後收文之審查意見通知1202--郭雅娟
                  If IsNull(.Fields(3)) Then
                     intI = 1
                     strSql = "SELECT Decode(PA09,'000',CPM03,CPM04)||'('||sqldatet(cp05)||')' status FROM patent,CASEPROGRESS,casepropertymap" & _
                        " WHERE pa01='" & .Fields("pa01") & "' AND pa02='" & .Fields("pa02") & "' AND pa03='" & .Fields("pa03") & "' AND pa04='" & .Fields("pa04") & "' " & _
                        " and pa57||pa16 is null and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='1202' and cpm01(+)=cp01 and cpm02(+)=cp10 order by cp05 desc"
                     Set Rs2 = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        strExc(0) = strExc(0) & Rs2(0)
                     End If
                  End If
                  'end 2018/7/16
               End If
               Printer.CurrentX = m_LeftMargin * m_CharWidth + Printer.TextWidth("其他多國案：")
               Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
               Printer.Print strExc(0)
               nRow = nRow + 1
               Call PrintCFormPage(m_MaxRow, nRow, m_LeftMargin * m_CharWidth, (m_TopMargin + nRow) * m_CharHeight) 'Added by Lydia 2021/08/26 判斷頁面最大列數，超過就跳頁
               .MoveNext
            Loop
         End With
      End If
   '2005/2/3 end
      'Added by Lydia 2015/12/03 列印美國IDS訊息
      If m_IDSmemo <> "" Then
          nRow = nRow + 1
          nRow = nRow + 1
          Printer.CurrentX = m_LeftMargin * m_CharWidth
          Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
          Printer.Print "IDS訊息："
          tmpArr = Empty
          tmpArr = Split(m_IDSmemo, "|")
          For intI = 0 To UBound(tmpArr)
              If tmpArr(intI) <> "" Then
                 Printer.CurrentX = m_LeftMargin * m_CharWidth + Printer.TextWidth("IDS訊息：")
                 Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
                 Printer.Print tmpArr(intI)
                 nRow = nRow + 1
                 Call PrintCFormPage(m_MaxRow, nRow, m_LeftMargin * m_CharWidth, (m_TopMargin + nRow) * m_CharHeight) 'Added by Lydia 2021/08/26 判斷頁面最大列數，超過就跳頁
              End If
          Next intI
          nRow = nRow - 1
          If nRow = 0 Then nRow = 1 'Added by Lydia 2021/08/26
      End If
      'end 2015/12/03
      
      'Add by Lydia 2014/12/30 EPC提供其他國家檢索報告之控管-列印所主張優先權案件之狀態
      If (rs1.Fields("CP10") = "1209" And rs1.Fields("PA09") = "221") Then
         Dim rsPri As New ADODB.Recordset
         Dim intP As Integer
            strSql = " select * from pridate where pd01='" & strExc(1) & "' and pd02='" & strExc(2) & "' and pd03='" & strExc(3) & "' and pd04='" & strExc(4) & "' "
            intP = 1
            Set rsPri = ClsLawReadRstMsg(intP, strSql)
            If intP = 1 Then
                nRow = nRow + 1
                Call PrintCFormPage(m_MaxRow, nRow, m_LeftMargin * m_CharWidth, (m_TopMargin + nRow) * m_CharHeight) 'Added by Lydia 2021/08/26 判斷頁面最大列數，超過就跳頁
                intP = m_LeftMargin * m_CharWidth + Printer.TextWidth("主張優先權案：")
                With rsPri
                   Do While Not .EOF
                      strExc(0) = PUB_GetEPCPriDLatest(rsPri!pd06, rsPri!pd07, strExc(5), strExc(6), strExc(7), "1")
                    If strExc(0) <> "" Then
                      If .AbsolutePosition = 1 Then
                        Printer.CurrentX = m_LeftMargin * m_CharWidth
                        Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
                        Printer.Print "主張優先權案："
                      End If
                      strExc(1) = strExc(0) & Space(2) & strExc(5) & Space(2) & strExc(6) & Space(2) & strExc(7)
                      Printer.CurrentX = intP
                      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
                      Printer.Print strExc(1)
                      nRow = nRow + 1
                      Call PrintCFormPage(m_MaxRow, nRow, m_LeftMargin * m_CharWidth, (m_TopMargin + nRow) * m_CharHeight) 'Added by Lydia 2021/08/26 判斷頁面最大列數，超過就跳頁
                    End If
                      .MoveNext
                   Loop
                End With
            End If
      End If
   End If
   
   'Added by Morgan 2022/7/1
   If rs1.Fields("CP01") = "CFP" And rs1.Fields("PA26") = "X80027010" Then
      nRow = nRow + 1
      nRow = nRow + 1
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "特殊提醒：「本案申請至答辯2次之費用已確定，本次收到審查意見書，將控管於2次內答辯完成"
      nRow = nRow + 1
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "　　　　　申請歷程，請斟酌案情，製作分析時，提出包括縮減專利範圍建議，若不允許或有困"
      nRow = nRow + 1
      Printer.CurrentX = m_LeftMargin * m_CharWidth
      Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
      Printer.Print "　　　　　難，請事先告知智權同仁與客戶溝通」"
   End If
   'end 2022/7/1
   
   'Added by Morgan 2020/12/10 等請作--TIM
'   If Left("" & rs1("cp12"), 1) = "F" And rs1.Fields("CP10") = "1002" Then
'      If Not IsNull(rs1("pa27")) Then strExc(2) = ChangeCustomerL(rs1("pa27"))
'      If Not IsNull(rs1("pa28")) Then strExc(3) = ChangeCustomerL(rs1("pa28"))
'      If Not IsNull(rs1("pa29")) Then strExc(4) = ChangeCustomerL(rs1("pa29"))
'      If Not IsNull(rs1("pa30")) Then strExc(5) = ChangeCustomerL(rs1("pa30"))
'      strExc(9) = rs1("cp01") & rs1("cp02") & rs1("cp03") & rs1("cp04")
'      strExc(1) = PUB_GetIncomMemoNew(strExc(9), rs1("cp01"), "", ChangeCustomerL(rs1("pa75")), ChangeCustomerL(rs1("pa26")), "Y", , , strExc(2), strExc(3), strExc(4), strExc(5))
'      strExc(2) = "": strExc(3) = "": strExc(4) = "": strExc(5) = ""
'      If strExc(1) <> "" Then
'         nRow = nRow + 2
'         intP = m_LeftMargin * m_CharWidth
'         Printer.CurrentX = intP
'         Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'         Printer.Print "外專備註："
'         arrMemo() = Split(strExc(1), vbCrLf)
'         For intI = LBound(arrMemo) To UBound(arrMemo)
'            nRow = nRow + 1
'            Printer.CurrentX = intP
'            Printer.CurrentY = (m_TopMargin + nRow) * m_CharHeight
'            Printer.Print arrMemo(intI)
'         Next
'      End If
'   End If
   'end 2020/12/10
   
   Printer.EndDoc

   'Added by Morgan 2018/10/2
   If bolSavePdf Then
      frmPDF.EndtProcess
      Unload frmPDF
      If Dir(strCOrderPath & "\" & strPdfName) <> "" Then
         Set oFile = oFileSys.GetFile(strCOrderPath & "\" & strPdfName)
         SaveAttFile_PDF strCP09, strCOrderPath & "\" & strPdfName, strPdfName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , , True
      End If
   End If
   'end 2018/10/2
End If
If rs1.State <> adStateClosed Then rs1.Close
Set rs1 = Nothing
'Added by Morgan 2018/10/2
Set oFile = Nothing
Set oFileSys = Nothing
End Sub

'Added by Lydia 2021/08/26 C類接洽單列印：判斷頁面最大列數，超過就跳頁
Private Sub PrintCFormPage(ByVal pMaxRow As Integer, ByRef pNowRow As Integer, ByVal pPosX As Double, ByVal pPosY As Double)
   
   If pMaxRow = 0 Then Exit Sub
   
   If pMaxRow < pNowRow Then
       Printer.CurrentX = pPosX + Printer.TextWidth(String(20, "　"))
       Printer.CurrentY = pPosY
       Printer.Print "（接下頁）"
       pNowRow = 1
       Printer.NewPage
   End If
   
End Sub

'Add by Morgan 2010/2/9
'抓最近實質相關程序名稱及收文日
Private Function GetLatestCP(stCP01 As String, stCP02 As String, stCP03 As String, stCP04 As String) As String
   Dim stSQL As String, intR As Integer
   Dim adoRst As ADODB.Recordset
   'Modify by Morgan 2010/3/10 +建議性處分書 1220
   'Modify By Sindy 2011/2/15 因用sqldate排序,修改百年蟲問題
'   stSQL = "select cp10,cp27,cpm03,sqldatet(cp05) cp05 from caseprogress c1,casepropertymap" & _
'      " where cp01='" & stCP01 & "' and cp02='" & stCP02 & "' and cp03='" & stCP03 & "' and cp04='" & stCP04 & "'" & _
'      " and (cp10='601' or (cp10 in ('1002','1006','1001','1220') and exists(select * from caseprogress c2" & _
'      " where c2.cp09=c1.cp43 and c2.cp10 in (" & UpdateCaseResultCP10List & "))))" & _
'      " and cpm01(+)=cp01 and cpm02(+)=cp10 order by cp05 desc"
   '2011/11/14 modify by sonia 加1602,1603 CFP-021944
   'Modified by Lydia 2015/10/05 + 1008
   stSQL = "select cp10,cp27,cpm03,cp05 from caseprogress c1,casepropertymap" & _
      " where cp01='" & stCP01 & "' and cp02='" & stCP02 & "' and cp03='" & stCP03 & "' and cp04='" & stCP04 & "'" & _
      " and (cp10='601' or (cp10 in ('1002','1006','1001','1220','1602','1603','1008') and exists(select * from caseprogress c2" & _
      " where c2.cp09=c1.cp43 and c2.cp10 in (" & UpdateCaseResultCP10List & "))))" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 order by cp05 desc"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      If adoRst("cp10") = "601" Then
         If IsNull(adoRst("cp27")) Then
            'Modify By Sindy 2011/2/15 因用sqldate排序,修改百年蟲問題
            'GetLatestCP = "領證中(" & adoRst("cp05") & "收文)"
            GetLatestCP = "領證中(" & ChangeWStringToTDateString(adoRst("cp05")) & "收文)"
         Else
            'Modify By Sindy 2011/2/15 因用sqldate排序,修改百年蟲問題
            'GetLatestCP = "已領證(" & adoRst("cp05") & "收文)"
            GetLatestCP = "已領證(" & ChangeWStringToTDateString(adoRst("cp27")) & "發文)"
         End If
      '2011/11/14 ADD BY SONIA
      ElseIf adoRst("cp10") = "1603" Then
         GetLatestCP = "已發證(" & ChangeWStringToTDateString(adoRst("cp05")) & "收文)"
      '2011/11/14 END
      Else
         'Modify By Sindy 2011/2/15 因用sqldate排序,修改百年蟲問題
         'GetLatestCP = "" & adoRst("cpm03") & "(" & adoRst("cp05") & ")"
         GetLatestCP = "" & adoRst("cpm03") & "(" & ChangeWStringToTDateString(adoRst("cp05")) & ")"
      End If
   End If
   Set adoRst = Nothing
End Function

'Add by Morgan 2005/2/3 取得實審收文情形
Private Function Get416State(ByRef p_Var() As String) As String
   Dim intR As Integer, stSQL As String
   
   'Modify by Morgan 2010/12/7 +未取消收文
   stSQL = "SELECT CP27 FROM CASEPROGRESS WHERE CP01='" & p_Var(1) & "' and CP02='" & p_Var(2) & "' and CP03='" & p_Var(3) & "' and CP04='" & p_Var(4) & "' and CP10='416' and cp57 is null"
   intR = 1
   Set AdoRecordSet3 = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      If IsNull(AdoRecordSet3(0)) Then
         Get416State = "實審已收文未發文"
      Else
         Get416State = "實審已收文已發文"
      End If
   Else
      'Modify by Morgan 2008/5/30 判斷下一程序有實審期限的
      'Get416State = "實審未收文"
      'Modify by Morgan 2010/12/7 +未續辦
      stSQL = "SELECT 1 FROM nextprogress WHERE np02='" & p_Var(1) & "' and np03='" & p_Var(2) & "' and np04='" & p_Var(3) & "' and np05='" & p_Var(4) & "' and np07='416' and np06 is null"
      intR = 1
      Set AdoRecordSet3 = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         Get416State = "實審未收文"
      End If
      'end 2008/5/30
   End If
End Function
'Add by Morgan 2010/6/9 取得新加坡檢索及實審收文情形
Private Function Get416State014(ByRef p_Var() As String) As String
   Dim intR As Integer, stSQL As String
   Dim iFlag As Integer
   
   'Modify by Morgan 2010/12/7 +未取消收文
   stSQL = "SELECT CP10,CP27 FROM CASEPROGRESS WHERE CP01='" & p_Var(1) & "' and CP02='" & p_Var(2) & "' and CP03='" & p_Var(3) & "' and CP04='" & p_Var(4) & "' and CP10 in ('416','421','427') and cp57 is null order by CP10 desc"
   intR = 1
   Set AdoRecordSet3 = ClsLawReadRstMsg(intR, stSQL)
   iFlag = 0
   If intR = 1 Then
      Do While Not AdoRecordSet3.EOF
         If Get416State014 <> "" Then
            Get416State014 = Get416State014 & "、"
         End If
         If AdoRecordSet3(0) = "427" Then
            iFlag = iFlag + 2
            Get416State014 = Get416State014 & "合併檢索及實審已收文"
         ElseIf AdoRecordSet3(0) = "421" Then
            iFlag = iFlag + 1
            Get416State014 = Get416State014 & "申請檢索報告已收文"
         Else
            iFlag = iFlag + 1
            Get416State014 = Get416State014 & "實審已收文"
         End If
         If IsNull(AdoRecordSet3(0)) Then
            Get416State014 = Get416State014 & "未發文"
         Else
            Get416State014 = Get416State014 & "已發文"
         End If
         AdoRecordSet3.MoveNext
      Loop
   End If
   
   If iFlag < 2 Then
      'Modify by Morgan 2010/12/7 +未續辦
      stSQL = "SELECT NP07 FROM nextprogress WHERE np02='" & p_Var(1) & "' and np03='" & p_Var(2) & "' and np04='" & p_Var(3) & "' and np05='" & p_Var(4) & "' and np07 in ('416','421','427') and np06 is null"
      intR = 1
      Set AdoRecordSet3 = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         Do While Not AdoRecordSet3.EOF
            If Get416State014 <> "" Then
               Get416State014 = Get416State014 & "、"
            End If
            If AdoRecordSet3(0) = "427" Then
               Get416State014 = Get416State014 & "合併檢索及實審未收文"
            ElseIf AdoRecordSet3(0) = "421" Then
               Get416State014 = Get416State014 & "申請檢索報告未收文"
            Else
               Get416State014 = Get416State014 & "實審未收文"
            End If
            AdoRecordSet3.MoveNext
         Loop
      End If
   End If
End Function

'Add by Morgan 2005/2/3 讀取多國案資料
Private Function GetRelCaseData(ByRef p_Var() As String, p_iR As Integer) As ADODB.Recordset
   Dim stSQL As String, stVTable As String
   'Modify by Morgan 2008/5/22 改叫公用
'   '其他相關案
'   stVTable = " select CR05 X01,CR06 X02,CR07 X03,CR08 X04 from caserelation" & _
'      " where CR01='" & p_Var(1) & "' AND CR02='" & p_Var(2) & "' AND CR03='" & p_Var(3) & "' AND CR04='" & p_Var(4) & "'" & _
'      " Union" & _
'      " select CR01,CR02,CR03,CR04 from caserelation" & _
'      " where CR05='" & p_Var(1) & "' AND CR06='" & p_Var(2) & "' AND CR07='" & p_Var(3) & "' AND CR08='" & p_Var(4) & "'"
'
'   'Add by Morgan 2007/9/20 加國內外案
'   '國內案
'   stVTable = stVTable & " UNION SELECT CM05,CM06,CM07,CM08 FROM CASEMAP" & _
'      " WHERE CM01='" & p_Var(1) & "' AND CM02='" & p_Var(2) & "' AND CM03='" & p_Var(3) & "' AND CM04='" & p_Var(4) & "'"
'   '國內案的其他國外案
'   stVTable = stVTable & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
'      "(SELECT CM05,CM06,CM07,CM08 FROM CASEMAP" & _
'      " WHERE CM01='" & p_Var(1) & "' AND CM02='" & p_Var(2) & "' AND CM03='" & p_Var(3) & "' AND CM04='" & p_Var(4) & "')"
'   '國外案
'   stVTable = stVTable & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP" & _
'      " WHERE CM05='" & p_Var(1) & "' AND CM06='" & p_Var(2) & "' AND CM07='" & p_Var(3) & "' AND CM08='" & p_Var(4) & "'"
'   '國外案的其他國外案
'   stVTable = stVTable & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
'      "(SELECT CM01,CM02,CM03,CM04 FROM CASEMAP" & _
'      " WHERE CM05='" & p_Var(1) & "' AND CM06='" & p_Var(2) & "' AND CM07='" & p_Var(3) & "' AND CM08='" & p_Var(4) & "')"
   'Modified by Morgan 2021/3/26
   'stVTable = PUB_GetRelCaseVTB(p_Var)
   stVTable = PUB_GetRelCaseVTB2(p_Var(1), p_Var(2), p_Var(3), p_Var(4), True)
   'end 2008/5/22
   
   '判斷是否為多國案:相關案中任一案件為多國案時就算是--秀玲
   'Modify by Morgan 2007/9/20 不限制多國案,國內外也要
   'strSQL = "SELECT 1 FROM (" & _
      " SELECT '" & p_Var(1) & "' X01,'" & p_Var(2) & "' X02,'" & p_Var(3) & "' X03,'" & p_Var(4) & "' X04 FROM DUAL" & _
      " Union All" & stVTable & _
      " ) X WHERE ROWNUM<2" & _
      " AND EXISTS(SELECT * FROM CASEPROGRESS WHERE CP01=X01 AND CP02=X02 AND CP03=X03 AND CP04=X04 AND CP21='Y')"
      
   stSQL = "SELECT distinct RPAD(PA01||'-'||PA02||'-'||PA03||'-'||PA04,15) C01, RPAD(NA03,16) C02, RPAD(PTM03,6) C03, RPAD(DECODE(PA57,'Y','閉卷',DECODE(PA16,'1','核准','2','核駁')),4) C04,PA57,PA09,PA01,PA02,PA03,PA04" & _
      " FROM (" & stVTable & ") X, Patent, Nation, PatentTrademarkMap" & _
      " WHERE NOT (X01='" & p_Var(1) & "' AND X02='" & p_Var(2) & "' AND X03='" & p_Var(3) & "' AND X04='" & p_Var(4) & "')" & _
      " AND PA01=X01 AND PA02=X02 AND PA03=X03 AND PA04=X04" & _
      " AND NA01(+)=PA09" & _
      " AND PTM01(+)=1 AND PTM02(+)=PA08"
   'end 2007/9/20
   p_iR = 1
   Set GetRelCaseData = ClsLawReadRstMsg(p_iR, stSQL)
   
End Function
'Add By Cheng 2003/06/23
'取得代理人地址
Private Function GetFagentAddress(strFA0102 As String, strLanguage As String) As String
'strFA0102 : 代理人代號
'strLanguage : 語文
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

GetFagentAddress = ""
strFA0102 = Left(strFA0102 & "000000000", 9)
StrSQLa = "Select * From Fagent Where FA01='" & Left(strFA0102, 8) & "' And FA02='" & Mid(strFA0102, 9, 1) & "' "
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
    Select Case strLanguage
    Case "0" '中文地址
        GetFagentAddress = Trim("" & rsA("FA17").Value)
    Case "1" '英文地址
        If "" & rsA("FA32").Value <> "" Then
            GetFagentAddress = Trim("" & rsA("FA32").Value & " " & rsA("FA33").Value & " " & rsA("FA34").Value & " " & rsA("FA35").Value & " " & rsA("FA36").Value)
        Else
            'Modify by Morgan 2007/1/23 加地址6 FA70
            GetFagentAddress = Trim("" & rsA("FA18").Value & " " & rsA("FA19").Value & " " & rsA("FA20").Value & " " & rsA("FA21").Value & " " & rsA("FA22").Value & " " & rsA("FA70").Value)
        End If
    Case Else '日文地址
        GetFagentAddress = Trim("" & rsA("FA23").Value)
    End Select
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing

End Function

'edit by nickc 2005/05/30 加一個判斷，要不要秀 msgbox
'Public Sub PrintReturnSheet(ByVal p_stNP01 As String, p_stNP07 As String, ByVal p_stNP09 As String)
'Modify by Morgan 2005/11/9 加參數 p_bolNoPrint=不列印,p_strArt=回傳內容,p_bolCtrlNP07=指定案件性質
'Public Sub PrintReturnSheet(ByVal p_stNP01 As String, p_stNP07 As String, ByVal p_stNP09 As String, Optional b_ShowMsg As Boolean = True)
'Modify by Morgan 2011/5/24 +p_bolPrintLetterHead:是否列印信頭
'Modify by Morgan 2011/6/1 改預設要印信頭
'Modify By Sindy 2012/3/27 +bolSavePic,ofrom
'Modified by Lydia 2020/05/25 +直接存Word檔 bolToWord ; 拿掉bolSavePic,ofrom (沒有使用)
'Public Sub PrintReturnSheet(Optional ByVal p_stNP01 As String, Optional p_stNP07 As String, Optional ByVal p_stNP09 As String, Optional b_ShowMsg As Boolean = True, Optional p_bolNoPrint As Boolean = False, Optional p_strArt As String, Optional p_bolCtrlNP07 As Boolean, Optional p_strCaseNo As String, Optional p_bolPrintLetterHead As Boolean = True, Optional bolSavePic As Boolean = False, Optional ofrom As Form)
'Modified by Lydia 2021/01/11 +定稿呼叫案件回覆單內容加印費用 bolLetterFee
'Modify By Sindy 2022/4/18 傳入呼叫的Form + , Optional pForm As Form = Nothing
Public Sub PrintReturnSheet(Optional ByVal p_stNP01 As String, Optional p_stNP07 As String, Optional ByVal p_stNP09 As String, _
   Optional b_ShowMsg As Boolean = True, Optional p_bolNoPrint As Boolean = False, Optional p_strArt As String, _
   Optional p_bolCtrlNP07 As Boolean, Optional p_strCaseNo As String, Optional p_bolPrintLetterHead As Boolean = True, _
   Optional ByVal bolToWord As Boolean, Optional bolLetterFee As Boolean = False, Optional pForm As Form = Nothing)
Dim strCont As String
Dim iType As Integer '回覆單種類 0:通用 1:領證及繳年費 2:年費 3:第一期註冊費 4:大陸領証 5:大陸年費 6:商標預設
Dim strCPM03s As String
Dim strData() As String
Dim bolReturn As Boolean
Dim strAppNo As String '申請案號
Dim strCaseTitle As String '案件名稱
Dim strCountry As String '申請國家
Dim strApplicant(1 To 5) As String '申請人
Dim i As Integer
Dim arrNP07
Dim intCnt As Integer  '2009/3/4 ADD BY SONIA 商品類別數
Dim strCU10 As String 'Add By Sindy 2009/04/10
Dim strCusCaseNo As String 'Add By Sindy 2009/06/12
'Add By Sindy 2009/06/29
Dim strYear As String '抓下次繳費年度
Dim m_Nexttimes As String '抓下次繳費次數
Dim strFeeType As String, strYF15 As String
Dim strFeeYear As String
Dim strKey(0 To 5) As String
Dim m_CaseFee(1 To 2) As String
Dim aryCaseFee As Variant
'2009/06/29 End
Dim m_iFixNo As Integer '2010/2/12 ADD BY SONIA 修法次數
Dim m_NO As String      '2010/4/1 ADD BY SONIA 分所案號
Dim strCurrTypeNM As String 'Add By Sindy 2010/11/18
Dim OneCase As Double, TwoCase As Double, OneAndTwoCase As Double
Dim strSales As String 'Add By Sindy 2014/6/5
Dim bstNP07 As String 'Added by Lydia 2016/1/12 保留傳入的下一程序案件性質
Dim p_stNP15 As String 'Add by Sindy 2017/4/25 下一程序備註
Dim p_stNP08 As String  'Added by Lydia 2020/05/25 本所期限
Dim p_stNP22 As String 'Added by Lydia 2021/01/11 序號

   bstNP07 = p_stNP07  'Added by Lydia 2016/1/12
   
   ReDim strData(1 To 4)
   'Modify by Morgan 2005/11/15 加是否指定案件性質
   'bolReturn = ReadNPData(p_stNP01, p_stNP09, strCPM03s, strData)
   If p_bolCtrlNP07 = True Then
      'Modified by Lydia 2020/05/25 + p_stNP08
      bolReturn = ReadNPData(p_stNP01, p_stNP09, strCPM03s, strData, p_strCaseNo, p_stNP07, p_stNP15, p_stNP08)
   Else
      'Modified by Lydia 2020/05/25 + p_stNP08
      bolReturn = ReadNPData(p_stNP01, p_stNP09, strCPM03s, strData, p_strCaseNo, , p_stNP15, p_stNP08)
   End If
   
   If bolReturn = False Then
      If p_bolNoPrint = False Then
         MsgBox "無可列印資料！", vbExclamation
      End If
      Exit Sub
   End If
   
   m_NO = "" '2010/4/1 ADD BY SONIA 分所案號
   Select Case strData(1)
      Case "P", "CFP", "FCP"
         ReDim Preserve strData(1 To TF_PA)
         bolReturn = PUB_ReadPatentData(strData, strData(1), strData(2), strData(3), strData(4))
         m_NO = strData(47) '2010/4/1 ADD BY SONIA 分所案號
      Case "T", "CFT", "FCT", "TF"
         ReDim Preserve strData(1 To TF_TM)
         bolReturn = PUB_ReadTradeMarkData(strData, strData(1), strData(2), strData(3), strData(4))
         m_NO = strData(34) '2010/4/1 ADD BY SONIA 分所案號
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS"
         ReDim Preserve strData(1 To TF_LC)
         bolReturn = PUB_ReadLawCaseData(strData, strData(1), strData(2), strData(3), strData(4))
         m_NO = strData(16) '2010/4/1 ADD BY SONIA 分所案號
      Case "LA"
         ReDim Preserve strData(1 To TF_HC)
         bolReturn = PUB_ReadHireCaseData(strData, strData(1), strData(2), strData(3), strData(4))
         m_NO = strData(7) '2010/4/1 ADD BY SONIA 分所案號
      Case Else
         ReDim Preserve strData(1 To tf_SP)
         bolReturn = PUB_ReadServicePracticeData(strData, strData(1), strData(2), strData(3), strData(4))
         m_NO = strData(28) '2010/4/1 ADD BY SONIA 分所案號
   End Select
   
   If bolReturn = False Then
      If p_bolNoPrint = False Then
         MsgBox "無法讀取基本檔資料！", vbExclamation
      End If
      Exit Sub
   End If
   
   iType = 0
   Select Case strData(1)
      Case "P", "CFP", "FCP"
         strCusCaseNo = Trim("" & strData(48)) 'Add By Sindy 2009/06/12
         strAppNo = strData(11)
         strCaseTitle = strData(5) & strData(6) & strData(7)
         strCountry = GetNationName(strData(9))
         If strData(9) = "000" Then
            If p_stNP07 = "601" Then
               iType = 1
            ElseIf p_stNP07 = "605" Then
               iType = 2
            End If
         ElseIf strData(9) = "020" Then
            If p_stNP07 = "601" Then
               iType = 4
            ElseIf p_stNP07 = "605" Then
               iType = 5
            End If
         'Added by Lydia 2014/12/10 EPC告准函-案件回覆單
         ElseIf strData(1) = "CFP" And strData(9) = "221" And p_stNP07 = "601" Then
                 iType = 7
         End If
         'PA26~30
         For i = 1 To 5
            If strData(25 + i) <> "" Then
               strApplicant(i) = GetCustomerName(strData(25 + i))
            End If
         Next
      Case "T", "CFT", "FCT", "TF"
         'add by nickc 2008/04/23 商標的預設
         iType = 6
         strCusCaseNo = Trim("" & strData(35)) 'Add By Sindy 2009/06/12
         strAppNo = strData(12)
         strCaseTitle = strData(5) & strData(6) & strData(7)
         If strData(10) = "000" Then
            'Modify By Sindy 2009/04/23
            'If p_stNP07 = "715" Then
            If p_stNP07 = "715" Or p_stNP07 = "716" Then
               iType = 3
            End If
         End If
         'Add By Sindy 2009/04/10
         '抓取申請人的國籍
         '2011/5/19 modify by sonia 改為有FC代理人先抓FC代理人國籍,無FC代理人才抓申請人國籍
         'strCU10 = GetPrjNationNumber1(strData(23))
         If "" & strData(44) = "" Then
            strCU10 = GetPrjNationNumber1(strData(23))
         Else
            strCU10 = GetPrjNationNumber("" & strData(44))
         End If
         '2011/5/19 END
         '2009/04/10 End
         strCountry = GetNationName(strData(10))
         'TM23
         'edit by nickc 2008/04/23 要區分國家，抓的順序與其他的不同
         'strApplicant(1) = GetCustomerName(strData(23))
         'Modify By Sindy 2016/1/4 mark
'         If strData(10) = "000" Then
         '2016/1/4 END
            strApplicant(1) = GetCustomerNameT000(strData(23))
'         Else
'            strApplicant(1) = GetCustomerNameTnot000(strData(23))
'         End If
         'add by nickc 2007/01/12 加申請人
         For i = 2 To 5
            If strData(76 + i) <> "" Then
               'edit by nickc 2008/04/23 要區分國家，抓的順序與其他的不同
               'Modify By Sindy 2016/1/4 mark
               'strApplicant(i) = GetCustomerName(strData(76 + i))
'                If strData(10) = "000" Then
               '2016/1/4 END
                   strApplicant(i) = GetCustomerNameT000(strData(76 + i))
'                Else
'                   strApplicant(i) = GetCustomerNameTnot000(strData(76 + i))
'                End If
            End If
         Next
      'Modify By Sindy 2009/07/24 增加LIN系統類別
      'modify by sonia 2019/7/29 +ACS系統類別
      Case "L", "CFL", "FCL", "LIN", "ACS"
         strCusCaseNo = Trim("" & strData(17)) 'Add By Sindy 2009/06/12
         strAppNo = ""
         strCaseTitle = strData(5) & strData(6) & strData(7)
         strCountry = GetNationName(strData(15))
         'LC11
         strApplicant(1) = GetCustomerName(strData(11))
      Case "LA"
         strCusCaseNo = "" 'Add By Sindy 2009/06/12
         strAppNo = ""
         strCaseTitle = strData(6)
         strCountry = ""
         'HC05
         strApplicant(1) = GetCustomerName(strData(5))
      Case Else
         strCusCaseNo = Trim("" & strData(29)) 'Add By Sindy 2009/06/12
         strAppNo = strData(11)
         strCaseTitle = strData(5) & strData(6) & strData(7)
         strCountry = GetNationName(strData(9))
         'SP08
         strApplicant(1) = GetCustomerName(strData(8))
         'SP58,SP59
         For i = 2 To 3
            If strData(56 + i) <> "" Then
               strApplicant(i) = GetCustomerName(strData(56 + i))
            End If
         Next
         'SP65,SP66
         For i = 4 To 5
            If strData(61 + i) <> "" Then
               strApplicant(i) = GetCustomerName(strData(61 + i))
            End If
         Next
   End Select
   
   'Add by Morgan 2008/4/18 將案件性質代碼轉中文
   'Added by Lydia 2016/1/12 P,CFP年費移到後面轉中文
   If Not ((strData(1) = "P" Or strData(1) = "CFP") And (InStr(bstNP07, "605") > 0 Or InStr(bstNP07, "606") > 0 Or InStr(bstNP07, "607") > 0)) Then
        arrNP07 = Split(strCPM03s, ",")
        strExc(1) = ""
        For i = LBound(arrNP07) To UBound(arrNP07)
           'edit by nickc 2008/05/21 修改成不是只有專利才可以用
           'If strData(9) = "000" Then
           If (strData(9) = "000" And InStr(1, strData(1), "P") > 0) Or (strData(10) = "000" And InStr(1, strData(1), "T") > 0) Then
              strSql = "select DECODE(CPM03,'（無）',CPM04,CPM03) CPM03 from CasePropertyMap where cpm01='" & strData(1) & "' and cpm02='" & arrNP07(i) & "'"
           Else
              strSql = "select DECODE(CPM04,'（無）',CPM03,CPM04) CPM03 from CasePropertyMap where cpm01='" & strData(1) & "' and cpm02='" & arrNP07(i) & "'"
           End If
           intI = 1
           Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
           If intI = 1 Then
              'Modified by Morgan 2020/11/16
              '歐盟延展+延展(英國)時要加勾選框
              If strData(1) = "CFT" And InStr(strCPM03s, "102") > 0 And InStr(strCPM03s, "110") > 0 Then
                  If arrNP07(i) = "102" Then
                     strExc(1) = strExc(1) & IIf(strExc(1) <> "", ", ", "") & "□歐盟延展"
                  ElseIf arrNP07(i) = "110" Then
                     strExc(1) = strExc(1) & IIf(strExc(1) <> "", ", ", "") & "□英國延展"
                  Else
                     strExc(1) = strExc(1) & IIf(strExc(1) <> "", ", ", "") & "□" & AdoRecordSet3("CPM03")
                  End If
              Else
              'end 2020/11/16
                  strExc(1) = strExc(1) & IIf(strExc(1) <> "", ",", "") & AdoRecordSet3("CPM03")
              End If 'Modified by Morgan 2020/11/16
           End If
        Next
        If strExc(1) <> "" Then strCPM03s = strExc(1)
        'end 2008/4/18
   End If
   'end 2016/1/12
   
   'Modify by Morgan 2006/3/7 調整位置格式--杜副總
   strCont = ""
   'Modify By Sindy 2009/04/22
   'Added by Lydia 2014/12/10 EPC告准函-案件回覆單 iType=7
   If strData(1) = "T" Or iType = 7 Then
      'Modify By Sindy 2019/11/21 +  & vbCrLf
      strCont = strCont & vbCrLf & vbCrLf & vbCrLf
   Else
      strCont = strCont & vbCrLf & vbCrLf & vbCrLf & vbCrLf & vbCrLf
   End If
   '2006/4/4 modify by sonia 依王協理需求不印通知日期,因與工程師之核駁分析日期差好幾天
   'strCont = strCont & "　　　　　　　　　　　　　　案件回覆單　　通知日期：" & (Mid(strSrvDate(1), 1, 4) - 1911) & "年" & Mid(strSrvDate(1), 5, 2) & "月" & Mid(strSrvDate(1), 7, 2) & "日" & vbCrLf
   strCont = strCont & "　　　　　　　　　　　　　　案件回覆單" & vbCrLf
   '2006/4/4 end
   strSales = PUB_GetAKindSalesNo(strData(1), strData(2), strData(3), strData(4)) 'Add By Sindy 2014/6/5 本所服務人員
   If strData(1) = "T" And (iType = 3 Or iType = 6) And strCU10 = "020" Then '大->台
      '不列印本所服務人員
   Else
      'Modified by Morgan 2015/5/19
      'strCont = strCont & "　　　　　　　　　　　　　　　　　　　　　本所服務人員：" & GetStaffName(strSales) & vbCrLf
      strExc(1) = PUB_GetLetterSalesName(strSales)
      If strExc(1) = "" Then
         strExc(1) = GetStaffName(strSales)
      End If
      strCont = strCont & vbCrLf & String(26 - Len(strExc(1)), "　") & "本所服務人員：" & strExc(1) & vbCrLf
   End If
   strCont = strCont & vbCrLf
   
   'Added by Lydia 2020/05/25 CFT延展案的開頭
   If strData(1) = "CFT" And (bstNP07 = "102" Or InStr(bstNP07, "102,") > 0) Then
        strCont = strCont & "申請國家：" & strCountry & vbCrLf
        strCont = strCont & "商標名稱：" & SplitWords(IIf(strCusCaseNo > "", strCaseTitle & "(" & strCusCaseNo & ")", strCaseTitle), 10, 68) & IIf(p_bolNoPrint = True And strData(1) = "T" And (iType = 3 Or iType = 6), GetPicColor(strData(1), strData(2), strData(3), strData(4)), "") & vbCrLf
        strCont = strCont & "註冊號數：" & strData(15) & vbCrLf
        strCont = strCont & "商品類別：第" & strData(9) & "類" & vbCrLf
        strCont = strCont & "本所案號：" & Replace(strData(1) & "-" & strData(2) & "-" & strData(3) & "-" & strData(4), "-0-00", "") & vbCrLf
   Else
   'end 2020/05/25
        'Add By Sindy 2009/04/09
        If strData(1) = "T" And (iType = 3 Or iType = 6) And strCU10 = "020" Then '大->台
           strCont = strCont & "回 覆 者：" & GetPrjName1(strData(44)) & vbCrLf
        End If
        '2009/04/09 End
        'strCont = strCont & "本所案號：" & Replace(strData(1) & "-" & strData(2) & "-" & strData(3) & "-" & strData(4), "-0-00", "") & vbCrLf
        'Modify By Sindy 2009/04/23
        '2010/4/1 modify by sonia 加印分所案號
        If m_NO <> "" Then m_NO = " (分所號：" & m_NO & ")"
        'modify by sonia 2022/9/29 +TM16判斷T-230518
        If strData(1) = "T" And (Trim(strData(15)) <> "" And Not IsNull(strData(15))) And Val(strData(16)) <> "2" Then
           strCont = strCont & "註 冊 號：" & Trim(strData(15)) & " (本所案號：" & Replace(strData(1) & "-" & strData(2) & "-" & strData(3) & "-" & strData(4), "-0-00", "") & ")" & m_NO & vbCrLf
        Else
           strCont = strCont & "申請案號：" & strAppNo & " (本所案號：" & Replace(strData(1) & "-" & strData(2) & "-" & strData(3) & "-" & strData(4), "-0-00", "") & ")" & m_NO & vbCrLf
        End If
        '2009/04/23 End
        
        'Modify By Sindy 2009/06/12
        'strCont = strCont & "案件名稱：" & SplitWords(strCaseTitle, 10, 68) & IIf(p_bolNoPrint = True And strData(1) = "T" And (iType = 3 Or iType = 6), GetPicColor(strData(1), strData(2), strData(3), strData(4)), "") & vbCrLf
        'Modified by Morgan 2025/4/24 中英混合時折行會有問題,調整一列總字數68->67 Ex:P-131714
        strCont = strCont & "案件名稱：" & SplitWords(IIf(strCusCaseNo > "", strCaseTitle & "(" & strCusCaseNo & ")", strCaseTitle), 10, 67) & IIf(p_bolNoPrint = True And strData(1) = "T" And (iType = 3 Or iType = 6), GetPicColor(strData(1), strData(2), strData(3), strData(4)), "") & vbCrLf
        '2009/06/12 End
        
        'Add By Sindy 2009/04/09 增加商品類別及商品名稱
        If strData(1) = "T" And (iType = 3 Or iType = 6) Then
           'Add By Sindy 2009/04/09 抓商品類別數
           intCnt = GetTMKindCnt(strData(1), strData(2), strData(3), strData(4))
           strCont = strCont & "商品類別：第" & strData(9) & "類" & vbCrLf
           If intCnt = 1 Then
              strCont = strCont & "商品名稱：" & GetTMGoods(strData(1), strData(2), strData(3), strData(4)) & vbCrLf
           End If
        End If
        '2009/04/09 End
        strCont = strCont & "申請國家：" & strCountry & vbCrLf
   End If 'Added by Lydia 2020/05/25
   
   strCont = strCont & "申 請 人：" & strApplicant(1) & vbCrLf
   If strApplicant(2) <> "" Then
      strCont = strCont & ("　　　　　") & strApplicant(2) & vbCrLf
   End If
   If strApplicant(3) <> "" Then
      strCont = strCont & ("　　　　　") & strApplicant(3) & vbCrLf
   End If
   If strApplicant(4) <> "" Then
      strCont = strCont & ("　　　　　") & strApplicant(4) & vbCrLf
   End If
   If strApplicant(5) <> "" Then
      strCont = strCont & ("　　　　　") & strApplicant(5) & vbCrLf
   End If
   
   
   'Add By Sindy 2009/06/29
   '下次繳費年度(次數)
   'Modified by Lydia 2016/1/12 年費和延展費合併催函
'   If (strData(1) = "P" Or strData(1) = "CFP") And _
'      (p_stNP07 = "605" Or p_stNP07 = "606" Or p_stNP07 = "607") Then
'      '設定本所案號
'      strKey(0) = ""
'      strKey(1) = strData(1)
'      strKey(2) = strData(2)
'      strKey(3) = strData(3)
'      strKey(4) = strData(4)
'      '取得繳年費的資料
'      '2010/2/12 MODIFY BY SONIA 抓修法次數
'      'If GetMoneyDate(strData(8), strData(9), strKey, m_CaseFee(1), m_CaseFee(2)) = True Then
'      If GetMoneyDate(strData(8), strData(9), strKey, m_CaseFee(1), m_CaseFee(2), , , m_iFixNo) = True Then
'         '取得下次繳費次數/年度
'         m_Nexttimes = PUB_Getnexttimes(strData(1), strData(2), strData(3), strData(4), strYear)
'         If m_Nexttimes <> "" Then
'            If p_stNP07 = "605" Then '605.年費
'               aryCaseFee = Split(m_CaseFee(2), ",")
'               '2012/8/6 MODIFY BY SONIA 年費也改抓YF15
'               'If Val(strYear) < Val(aryCaseFee(UBound(aryCaseFee))) Then
'               '   'strFeeYear = " (第" & strYear & "年至第___年)"
'               '   strFeeYear = " (第" & strYear & "年)"
'               'Else
'               '   strFeeYear = " (第" & strYear & "年)"
'               'End If
'               strFeeType = PUB_GetNa20Na22Na24(strData(9), strData(8))
'               'Added by Morgan 2013/8/23 印尼年費第1次為發證日後1年繳已使用完的年費Ex.CFP-23499 申請日 99/10/20 發證日 101/7/12 第1次繳費期限 102/7/12 須繳 99/10/20-101/10/19 1-2年年費
'               If strData(9) = "017" And strFeeType = "605" And m_Nexttimes = "1" Then
'                  strYF15 = PUB_GetXDesc(strData(1), strData(2), strData(3), strData(4))
'               Else
'               'end 2013/8/23
'                  strYF15 = PUB_GetYF15(strData(9), strData(8), "Y000000" & m_iFixNo, strFeeType, CDbl(strYear))
'               End If 'Added by Morgan 2013/8/23
'
'               strFeeYear = " " & strYF15
'               '2012/8/6 END
'            ElseIf p_stNP07 = "606" Or p_stNP07 = "607" Then '606.維持費 607.延展費
'               '年度說明
'               strFeeType = PUB_GetNa20Na22Na24(strData(9), strData(8))
'               '2010/2/12 MODIFY BY SONIA
'               'strYF15 = PUB_GetYF15(strData(9), strData(8), "Y0000000", strFeeType, CDbl(strYear))
'               strYF15 = PUB_GetYF15(strData(9), strData(8), "Y000000" & m_iFixNo, strFeeType, CDbl(strYear))
'               strFeeYear = " " & strYF15
'            End If
'         End If
'      End If
'   End If
'   '2009/06/29 End
   If (strData(1) = "P" Or strData(1) = "CFP") And (InStr(bstNP07, "605") > 0 Or InStr(bstNP07, "606") > 0 Or InStr(bstNP07, "607") > 0) Then
       'Modified by Lydia 2016/06/07 改成模組
'       '設定本所案號
'       strKey(0) = ""
'       strKey(1) = strData(1)
'       strKey(2) = strData(2)
'       strKey(3) = strData(3)
'       strKey(4) = strData(4)
'       arrNP07 = Empty
'       arrNP07 = Split(bstNP07, ",")
'       strCPM03s = ""
'       For i = LBound(arrNP07) To UBound(arrNP07)
'          '抓案件性質名稱
'          strExc(1) = "": strFeeYear = ""
'          If (strData(9) = "000" And InStr(1, strData(1), "P") > 0) Or (strData(10) = "000" And InStr(1, strData(1), "T") > 0) Then
'             strSql = "select DECODE(CPM03,'（無）',CPM04,CPM03) CPM03 from CasePropertyMap where cpm01='" & strData(1) & "' and cpm02='" & arrNP07(i) & "'"
'          Else
'             strSql = "select DECODE(CPM04,'（無）',CPM03,CPM04) CPM03 from CasePropertyMap where cpm01='" & strData(1) & "' and cpm02='" & arrNP07(i) & "'"
'          End If
'          intI = 1
'          Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
'          If intI = 1 Then
'             strExc(1) = strExc(1) & IIf(strExc(1) <> "", ",", "") & AdoRecordSet3("CPM03")
'          End If
'          'CFP-大馬新型延展費,俄羅斯設計延展費
'          If (strData(1) = "CFP" And strData(9) = "018" And strData(8) = "2" And arrNP07(i) = "607") Or _
'             (strData(1) = "CFP" And strData(9) = "233" And strData(8) = "3" And arrNP07(i) = "607") Then
'               strFeeYear = " " & PUB_GetExpYF607(strData(9), strData(1), strData(2), strData(3), strData(4), strData(10))
'
'          ElseIf (arrNP07(i) = "605" Or arrNP07(i) = "606" Or arrNP07(i) = "607") Then
'              '取得繳年費的資料(抓修法次數)
'              If GetMoneyDate(strData(8), strData(9), strKey, m_CaseFee(1), m_CaseFee(2), , , m_iFixNo) = True Then
'                 '取得下次繳費次數/年度
'                 m_Nexttimes = PUB_Getnexttimes(strData(1), strData(2), strData(3), strData(4), strYear)
'                 If m_Nexttimes <> "" Then
'                    If p_stNP07 = "605" Then '605.年費
'                       aryCaseFee = Split(m_CaseFee(2), ",")
'                       strFeeType = PUB_GetNa20Na22Na24(strData(9), strData(8))
'                       If strData(9) = "017" And strFeeType = "605" And m_Nexttimes = "1" Then
'                          strYF15 = PUB_GetXDesc(strData(1), strData(2), strData(3), strData(4))
'                       Else
'                          strYF15 = PUB_GetYF15(strData(9), strData(8), "Y000000" & m_iFixNo, strFeeType, CDbl(strYear))
'                       End If
'                       strFeeYear = " " & strYF15
'
'                    ElseIf p_stNP07 = "606" Or p_stNP07 = "607" Then '606.維持費 607.延展費
'                       '年度說明
'                       strFeeType = PUB_GetNa20Na22Na24(strData(9), strData(8))
'                       strYF15 = PUB_GetYF15(strData(9), strData(8), "Y000000" & m_iFixNo, strFeeType, CDbl(strYear))
'                       strFeeYear = " " & strYF15
'                    End If
'                 End If
'              End If
'          End If
'          strCPM03s = strCPM03s & strExc(1) & strFeeYear & ","
'       Next
'       If Right(strCPM03s, 1) = "," Then strCPM03s = Mid(strCPM03s, 1, Len(strCPM03s) - 1)
       strCPM03s = PUB_GetCPMnexttimes(strData(1), strData(2), strData(3), strData(4), IIf(InStr(1, strData(1), "T") > 0, strData(10), strData(9)), strData(8), bstNP07, p_stNP07)
   End If
'-------2016/1/12 end

   'Modify by Morgan 2007/1/12 台灣專利的申復要加印修正選項
   If strData(1) = "P" And strData(9) = "000" And p_stNP07 = "205" Then
      'Modify by Morgan 2007/2/13 加控制來函的案件性質為1202審查意見通知函的才要
      If GetCP10(p_stNP01) = "1202" Then
         strCont = strCont & "案件性質：□修正" & vbCrLf
         strCont = strCont & "　　　　　□" & strCPM03s & vbCrLf
      Else
         strCont = strCont & "案件性質：" & strCPM03s & vbCrLf
      End If
      'End 2007/2/13
   'Add By Sindy 2017/4/19 緬甸701領證
   ElseIf strData(1) = "CFT" And strData(10) = "048" And p_stNP07 = "701" Then
      strCont = strCont & "案件性質：□" & strCPM03s & "　□刊登廣告" & vbCrLf
   '2017/4/19 END
   'add by nickc 2008/04/23 加入商標的控制
   ElseIf strData(1) = "T" And strData(10) = "000" And p_stNP07 = "715" Then
      strCont = strCont & "案件性質：註冊費" & vbCrLf
   'Added by Lydia 2014/12/10 EPC告准函-案件回覆單
   ElseIf strData(1) = "CFP" And strData(9) = "221" And p_stNP07 = "601" Then
          'Modified by Morgan 2023/3/6
          'strCont = strCont & "案件性質：" & strCPM03s & "＆指定國註冊登記" & vbCrLf
          strCont = strCont & "案件性質：" & strCPM03s & "，指定國註冊登記，請求UP" & vbCrLf
          'end 2023/3/3
   'Add by Amy 2017/07/20 美國最終核駁下一程序除 "答辨" 外尚可選擇 "請求繼續審查" 或 "訴願"
   ElseIf strData(1) = "CFP" And strData(9) = "101" And p_stNP07 = "107" And GetCP10(p_stNP01) = "1006" Then
        strCont = strCont & "案件性質：□" & strCPM03s & vbCrLf
        strCont = strCont & "　　　　　□請求繼續審查" & vbCrLf
        strCont = strCont & "　　　　　□訴願" & vbCrLf
        
   'Added by Morgan 2020/11/16 歐盟延展費
   ElseIf strData(1) = "CFP" And strData(9) = "239" And p_stNP07 = "607" Then
      If PUB_ChkNPExist(strData(), "613") = True Then
         strCont = strCont & "案件性質：□歐盟" & strCPM03s & vbCrLf
         strCont = strCont & "　　　　　□英國" & strCPM03s & vbCrLf
      Else
         strCont = strCont & "案件性質：" & strCPM03s & vbCrLf
      End If
      
   'Added by Morgan 2023/6/12
   ElseIf strData(1) = "CFP" And strData(9) = "023" And p_stNP07 = "601" Then
      strCont = strCont & "案件性質：□" & strCPM03s & vbCrLf
      strCont = strCont & "　　　　　□申請紙本專利證書" & vbCrLf
   'end 2023/6/12
              
   'add by sonia 2024/5/28 印度040商業使用聲明930
   ElseIf strData(1) = "CFP" And strData(9) = "040" And p_stNP07 = "930" Then
      strCont = strCont & "案件性質：" & strCPM03s & vbCrLf
      strCont = strCont & "　　　　　□有作商業使用" & vbCrLf
      strCont = strCont & "　　　　　□沒有作商業使用，原因：" & vbCrLf
      strCont = strCont & "　　　　　　　□發明專利正在開發/商業試驗中" & vbCrLf
      strCont = strCont & "　　　　　　　□發明專利正在接受監管機構的審查/批准" & vbCrLf
      strCont = strCont & "　　　　　　　□尋求商業授權中" & vbCrLf
      strCont = strCont & "　　　　　　　□其他，具體說明：" & vbCrLf
      strCont = strCont & "是否可供授權該專利" & vbCrLf
      strCont = strCont & "□否" & vbCrLf
      strCont = strCont & "□是：聯絡人信箱：" & vbCrLf
      strCont = strCont & "　　　聯絡人電話：" & vbCrLf
   'end 2024/5/28
              
   Else
      'Modify By Sindy 2009/06/29
      'strCont = strCont & "案件性質：" & strCPM03s & vbCrLf
      'Modified by Lydia 2016/1/11
      'strCont = strCont & "案件性質：" & strCPM03s & strFeeYear & vbCrLf
      strCont = strCont & "案件性質：" & strCPM03s & vbCrLf
      '2009/06/29 End
   End If

   'Added by Lydia 2021/01/11 P、CFP特定案件之回覆單加註費用
   If bolLetterFee = True And strData(1) = "P" And strData(9) = "020" And bstNP07 = "601" Then
       'P(大陸)之領證
       strSql = "select decode(et05,'大陸領證費','1','年費金額','2','3') ord1, et05,et06 from  exceptcondition where et01='05' and et03 in ('03','04') and et05 in ('大陸領證費','年費金額','總費用') and et02 in " & _
                   "(select cp43 from caseprogress where cp09=(select np01 from nextprogress where np02='" & strData(1) & "' and np03='" & strData(2) & "' and np04='" & strData(3) & "' and np05='" & strData(4) & "' and np06 is null and np07='601')) "
       strSql = strSql & "order by 1"
       intI = 1
       Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
       If intI = 1 Then
           strExc(1) = "": strExc(2) = "": strExc(3) = ""
           AdoRecordSet3.MoveFirst
           Do While Not AdoRecordSet3.EOF
               strExc(Val(AdoRecordSet3.Fields("ord1"))) = "" & AdoRecordSet3.Fields("et06")
               AdoRecordSet3.MoveNext
           Loop
           If Trim(strExc(2)) = "" Then '年費未屆
                'modify by sonia 2022/3/28 回覆單以下費用加入新台幣字樣
                strCont = strCont & "費　　用：新台幣" & Format(strExc(3), "###,###,##0") & "元" & vbCrLf
           Else
                 If strExc(1) <> "" Then   '年費將屆：分成兩種選項，第一個選項：費  用：（領證及該年度年費的總額），第二個選項：費  用(含下一年度)：（領證.該年度及下一年度年費的總額）
                      strCont = strCont & "費　　用：新台幣" & Format(strExc(1), "###,###,##0") & "元" & vbCrLf
                 End If
                 If strExc(2) <> "" Then
                      strCont = strCont & "費用(含下一年度)：新台幣" & Format(strExc(3), "###,###,##0") & "元" & vbCrLf
                 End If
           End If
       End If
   End If
   If bolLetterFee = True And strData(1) = "CFP" And InStr("601,605,606,607,", bstNP07 & ",") > 0 Then
       Call PUB_ChkNPExist(strData, bstNP07, 0, p_stNP22, p_stNP01)
       If p_stNP01 <> "" And p_stNP22 <> "" Then
            'CFP之領證、年費、維持費及延展費：帶出「費用：（通知之費用）」
            Call ClsPDGetCaseProperty(strData(1), bstNP07, strExc(0))
            If PUB_CheckCFPtoFee(strData(9), strData(1), p_stNP01, p_stNP22, bstNP07 & " " & strExc(0), p_stNP15, strExc(1), strExc(2), strExc(3)) = True Then
                 If strData(9) = "239" And bstNP07 = "607" Then '歐盟延展案＋英國延展費
                     If PUB_ChkNPExist(strData(), "613") = True Then  '未收文: 歐盟案613延展費(英國)
                          'modify by sonia 2022/3/28 回覆單以下費用加入新台幣字樣
                          strCont = strCont & "費　　用：歐盟　新台幣" & Format(strExc(1), "###,###,##0") & "元、"
                          Call ClsPDGetCaseProperty(strData(1), "613", strExc(0))
                          If PUB_CheckCFPtoFee(strData(9), strData(1), p_stNP01, p_stNP22, "613 " & strExc(0), "", strExc(1), strExc(2), strExc(3)) = True Then
                              strCont = strCont & "英國　新台幣" & Format(strExc(1), "###,###,##0") & "元" & vbCrLf
                          End If
                     Else
                          strCont = strCont & "費　　用：新台幣" & Format(strExc(1), "###,###,##0") & "元" & vbCrLf
                     End If
                 Else
                     strCont = strCont & "費　　用：新台幣" & Format(strExc(1), "###,###,##0") & "元" & vbCrLf
                 End If
            End If
       End If
   End If
   'end 2021/01/11
   
   'Added by Morgan 2024/11/6 大陸補償期年費要報價
   If strData(1) = "P" And bstNP07 = "615" Then
      strExc(1) = PUB_GetCN615Fee(strData())
      strCont = strCont & "費　　用：新台幣" & Format(strExc(1), "###,###,##0") & "元" & vbCrLf
   End If
   'end 2024/11/6
   
   'Added by Lydia 2020/05/25 CFT延展案：+法限+所限+可辦期限
   If strData(1) = "CFT" And (bstNP07 = "102" Or InStr(bstNP07, "102,") > 0) Then
        strExc(1) = TranslateKeyWord(incCNV_CHINESE_MINKO, p_stNP09, Empty, strData(1) & strData(2) & strData(3) & strData(4))
        strCont = strCont & "法定期限：" & strExc(1) & vbCrLf
        strExc(2) = Empty
        strExc(1) = TranslateKeyWord(incCNV_CHINESE_MINKO, p_stNP08, Empty, strData(1) & strData(2) & strData(3) & strData(4))
        strCont = strCont & "本所期限：" & strExc(1) & vbCrLf
        strExc(2) = GetPrjNationName(strData(10), "NA15") '延展時間(月)
        If Val(strExc(2)) > 0 Then
            strExc(1) = CompDate(1, Val(strExc(2)) * -1, p_stNP09) '可辦期限=法限-NA15(月)
            If strData(10) = "011" Then strExc(1) = CompDate(2, 1, strExc(1)) 'Added by Morgan 2023/7/12 日本可辦延展期限起日要+1天(翌日起),定稿要同步修改--May
            strExc(3) = TranslateKeyWord(incCNV_CHINESE_MINKO, strExc(1), Empty, strData(1) & strData(2) & strData(3) & strData(4))
            strCont = strCont & "|#(中字)  (本案於：" & strExc(3) & "起，始能辦理)#|" & vbCrLf
        Else
            strCont = strCont & "|#(中字)  (本案於：      年    月    日起，始能辦理)#|" & vbCrLf
        End If
   'Mark by Lydia 2020/08/05 因回覆單加印本所期限可能會造成客戶誤解回覆日期，反而造成專業部來不及作業，
                                          '約定期限系統可能抓不到(個案改)也不能放。
   'Else  '其他：+所限
   '     strExc(1) = TranslateKeyWord(incCNV_CHINESE_MINKO, p_stNP08, Empty, strData(1) & strData(2) & strData(3) & strData(4))
   '     strCont = strCont & "本所期限：" & strExc(1) & vbCrLf
   'end 2020/08/05
   End If
   'end 2020/05/25
   
   'Modify By Sindy 2017/4/25
   If strData(1) = "CFT" And p_stNP07 = "601" Then
      strCont = strCont & "下一程序備註：" & p_stNP15 & vbCrLf
   Else
   '2017/4/25 END
      strCont = strCont & vbCrLf
   End If
   
   Select Case iType
      Case 0
         strCont = strCont & "□同意辦理，請　貴所代為管制本案後續發展。"
         
         'Modified by Morgan 2013/10/18 本段程式 103/1/1 以後還原
         'strCont = strCont & vbCrLf
         If strData(1) = "CFP" And strData(9) = "101" And p_stNP07 = "601" And Val(strSrvDate(1)) < 20140101 Then
            strSql = "select et06,et05 from nextprogress,exceptcondition where np01='" & p_stNP01 & "' and np07='601' and et01(+)='03' and et02(+)=np01 and et04(+)=np16 and et05 in ('費用總計','新費用總計') order by et05"
            intI = 1
            Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
            If AdoRecordSet3.RecordCount = 2 Then
               strCont = strCont & "(請勾選下列選項)" & vbCrLf
               Do While Not AdoRecordSet3.EOF
                  If AdoRecordSet3("et05") = "費用總計" Then
                     strCont = strCont & "　　□領證費用計為新台幣" & Format(AdoRecordSet3("et06"), "#,##0") & "元整。(2013年12月31日前(含當日))" & vbCrLf
                  Else
                     strCont = strCont & "　　□領證費用計為新台幣" & Format(AdoRecordSet3("et06"), "#,##0") & "元整。(2014年01月01日後(含當日))" & vbCrLf
                  End If
                  AdoRecordSet3.MoveNext
               Loop
            Else
               strCont = strCont & vbCrLf
            End If
         Else
            strCont = strCont & vbCrLf
         End If
         'end 2013/10/18
         
      Case 1 '台灣專利領證及繳年費 601
         strCont = strCont & "□同意辦理領證及繳納第  年至第  年年費，請　貴所代為管制本案後續發展。" & vbCrLf
         strCont = strCont & "　□本人／本公司不符合專利年費減免條件。" & vbCrLf
         strCont = strCont & "　□本人／本公司符合專利年費減免條件。" & vbCrLf
      Case 2 '台灣專利年費 605
         strCont = strCont & "□同意辦理繳納第  年至第  年年費，請　貴所代為管制本案後續發展。" & vbCrLf
         strCont = strCont & "　□本人／本公司不符合專利年費減免條件。" & vbCrLf
         strCont = strCont & "　□本人／本公司符合專利年費減免條件。" & vbCrLf
      Case 3 '商標第一期註冊費 715
         'Add By Sindy 2009/04/10
         If strCU10 > "010" Then '大->台
            '2009/9/10 MODIFY BY SONIA 改抓商品類別數intCnt
            'Dim tmpTM09 As Variant
            'Dim OneCase As Double, TwoCase As Double, OneAndTwoCase As Double
            'tmpTM09 = Split(Trim(strData(9)), ",")
            'OneCase = (UBound(tmpTM09) + 1) * 650
            'TwoCase = (UBound(tmpTM09) + 1) * 775
            'OneAndTwoCase = (UBound(tmpTM09) + 1) * 1025
            'Modify By Sindy 2013/1/22
            If Left(Trim(strData(44)), 6) = "Y51566" Then
               strCurrTypeNM = "新台幣"
               If intCnt > 1 Then '1類以上
                  '服務費2000規費2500*類別數
                  OneAndTwoCase = (intCnt * 2500) + 2000
               Else
                  '服務費1000規費2500
                  OneAndTwoCase = 1000 + 2500
               End If
            Else
               strCurrTypeNM = "人民幣"
               OneCase = intCnt * 650
               TwoCase = intCnt * 775
               OneAndTwoCase = intCnt * 1000
            End If
            '2009/9/10 END
'            'Add By Sindy 2010/11/18
'            If Left(Trim(strData(44)), 6) = "Y51566" Then
'               strCurrTypeNM = "新台幣"
'            Else
'               strCurrTypeNM = "人民幣"
'            End If
            '2013/1/22 End
            'Modify By Sindy 2009/04/23
            If p_stNP07 = "716" Then '第二期註冊費
               strCont = strCont & "□同意繳納第二期註冊費，費用" & strCurrTypeNM & Format(TwoCase, "###,###,##0") & "元。" & vbCrLf
            Else
               'Modify By Sindy 2012/5/29 商標修法
               If Val(strData(13)) >= 20120701 Then
                  'modify by sonia 2022/3/28 回覆單以下費用加入新台幣字樣
                  strCont = strCont & "□同意繳納註冊費，費用新台幣" & strCurrTypeNM & Format(OneAndTwoCase, "###,###,##0") & "元。" & vbCrLf
               Else
               '2012/5/29 End
                  strCont = strCont & "□同意繳納第一期註冊費，費用" & strCurrTypeNM & Format(OneCase, "###,###,##0") & "元。" & vbCrLf
                  strCont = strCont & "□同意繳納全期（含第一期及第二期）註冊費，費用" & strCurrTypeNM & Format(OneAndTwoCase, "###,###,##0") & "元。" & vbCrLf
               End If
            End If
            '2009/04/23 End
         '2009/04/10 End
         Else
            '2008/11/24 ADD BY SONIA 加費用及規費
            'strCont = strCont & "□同意繳納第一期註冊費，並另行給付　貴所服務費。" & vbCrLf
            'strCont = strCont & "□同意同時繳納第一期及第二期註冊費，並另行給付　貴所服務費。" & vbCrLf
            Dim o71713 As Double  '點數為服務費
            Dim o71708 As Double  '規費
            Dim o71613 As Double
            Dim o71608 As Double
            Dim o71513 As Double
            Dim o71508 As Double
            strSql = "select * from casefee where cf01='" & strData(1) & "' and cf02='" & strData(10) & "' and cf03 in ('715','716','717') order by cf03 "
            intI = 1
            Set AdoRecordSet3 = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               AdoRecordSet3.MoveFirst
               Do While Not AdoRecordSet3.EOF
                  '2009/3/4 ADD BY SONIA 抓商品類別數
                  intCnt = GetTMKindCnt(strData(1), strData(2), strData(3), strData(4))
                  '2009/3/4 END
                  '2009/3/4 MODIFY BY SONIA CF08規費要*商品類別數
                  Select Case AdoRecordSet3.Fields("cf03").Value
                     Case "715"
                        o71513 = AdoRecordSet3.Fields("cf13").Value
                        o71508 = AdoRecordSet3.Fields("cf08").Value * intCnt
                     Case "716"
                        o71613 = AdoRecordSet3.Fields("cf13").Value
                        o71608 = AdoRecordSet3.Fields("cf08").Value * intCnt
                     Case "717"
                        o71713 = AdoRecordSet3.Fields("cf13").Value
                        o71708 = AdoRecordSet3.Fields("cf08").Value * intCnt
                     Case Else
                  End Select
                  AdoRecordSet3.MoveNext
               Loop
            End If
            'Modify By Sindy 2009/05/04
            If p_stNP07 = "716" Then '第二期註冊費
               strCont = strCont & "□同意繳納第二期註冊費，新台幣" & Format((o71613 * 1000) + o71608, "###,###,##0") & "元（含規費" & Format(o71608, "###,###,##0") & "元）。" & vbCrLf
            '2009/05/04 End
            Else
               strCont = strCont & "□同意繳納第一期註冊費，新台幣" & Format((o71513 * 1000) + o71508, "###,###,##0") & "元（含規費" & Format(o71508, "###,###,##0") & "元）。" & vbCrLf
               strCont = strCont & "□同意繳納全期（含第一期及第二期）註冊費，新台幣" & Format((o71713 * 1000) + o71708, "###,###,##0") & "元（含規費" & Format(o71708, "###,###,##0") & "元）。" & vbCrLf
            End If
            'strCont = strCont & "　幣" & Format((o71713 * 1000), "###,###,##0") & "元。" & vbCrLf
            '2008/11/24 END
         End If
      Case 4 '大陸專利領證及繳年費 601
         strCont = strCont & "□同意辦理領證及繳納第  年年費，請　貴所代為管制本案後續發展。" & vbCrLf
      Case 5 '大陸專利年費 605
         strCont = strCont & "□同意辦理繳納第  年年費，請　貴所代為管制本案後續發展。" & vbCrLf
      Case 6 '商標的預設值----------------------------------------︾
         'Modify By Sindy 2012/7/20 大->台審定核准
         If strData(1) = "T" And strCU10 = "020" And p_stNP07 = "717" Then
            'Modify By Sindy 2013/1/22
            If Left(Trim(strData(44)), 6) = "Y51566" Then
               'modify by sonia 2018/8/28 改每類人民幣800元,依此類推T-212285
               'If intCnt > 1 Then '1類以上
               '   '服務費2000規費2500*類別數
               '   OneAndTwoCase = 2000 + (intCnt * 2500)
               'Else
               '   '服務費1000規費2500
               '   OneAndTwoCase = 1000 + 2500
               'End If
               'strCont = strCont & "□同意繳納註冊費，新台幣" & Format(OneAndTwoCase, "###,###,##0") & "元。" & vbCrLf
               strCont = strCont & "□同意繳納註冊費，人民幣" & Format((800 * intCnt), "###,###,##0") & "元。" & vbCrLf
               'end 2018/8/28
            Else
               strCont = strCont & "□同意繳納註冊費，人民幣" & Format((1000 * intCnt), "###,###,##0") & "元。" & vbCrLf
            End If
            '2013/1/22 End
         Else
         '2012/7/20 End
            strCont = strCont & "□同意辦理，請　貴所代為處理本案後續發展。" & vbCrLf
         End If
       'Added by Lydia 2014/12/10 EPC告准函-案件回覆單
      Case 7
         'Modified by Morgan 2023/3/6
         'strCont = strCont & "□同意辦理領證，請　貴所代為管制指定國註冊登記期限。" & vbCrLf
         'strCont = strCont & "□同意辦理領證及指定國註冊登記。(勾註此選項者，請告知欲辦理註冊之指定" & vbCrLf
         'strCont = strCont & "　國家)。" & vbCrLf
         'strCont = strCont & "　　指定國：" & vbCrLf & vbCrLf & vbCrLf & vbCrLf
         strCont = strCont & "□同意辦理領證，請　貴所先代為管制指定國註冊登記期限。" & vbCrLf
         strCont = strCont & "□同意辦理領證，請　貴所依規劃表上勾選之國家提供報價。" & vbCrLf
         'end 2023/3/6
   End Select
   
   'Modify By Sindy 2009/04/09
'   If strData(1) = "T" Then
'      strCont = strCont & "□本人／本公司自行處理本案之後續作業，　貴所不需作本案之後續追蹤及通" & vbCrLf
'      strCont = strCont & "　知。" & vbCrLf
'   Else
      strCont = strCont & "□本人／本公司自行處理本案之後續作業，並同意　貴所不需作本案之後續追" & vbCrLf
      strCont = strCont & "　蹤及通知。" & vbCrLf
'   End If
   '2009/04/09 End
   strCont = strCont & "□放棄本案。" & vbCrLf
   strCont = strCont & "□其他，請說明ˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍˍ" & vbCrLf
   'edit by nickc 2006/03/24 加入定稿用代表圖定位器
   'strCont = strCont & vbCrLf
   'Modify By Sindy 2009/04/09
   'strCont = strCont & IIf(p_bolNoPrint = True, IIf(IIf(strData(1) = "P", True, IIf(strData(1) = "CFP", True, False)) And IIf(iType = 2, True, False), "|#代表圖#|", ""), "") & vbCrLf
   If p_bolNoPrint = True And (strData(1) = "P" Or strData(1) = "CFP") And iType = 2 Then
      'Modified by Morgan 2022/10/7 增加跳行調整位置否則會壓到文字 Ex:P-099472
      'strCont = strCont & "|#代表圖#|" & vbCrLf
      strCont = strCont & vbCrLf & "|#代表圖#|" & vbCrLf & vbCrLf
      'end 2022/10/7
   'Modified by Lydia 2020/04/01 T、CFT案件回覆單產生電子檔，一律加上代表圖。
   'ElseIf p_bolNoPrint = True And strData(1) = "T" And (iType = 3 Or iType = 6) Then
   'Modified by Lydia 2020/05/25 直接存/列印Word檔
   'ElseIf p_bolNoPrint = True And (strData(1) = "T" Or strData(1) = "CFT") Then
   ElseIf strData(1) = "T" Or strData(1) = "CFT" Then
      strCont = strCont & vbCrLf 'Add By Sindy 2010/12/9 ex:T-169733
      'Modified by Lydia 2020/05/25 改用自動代表圖
      'strCont = strCont & "|#代表圖#|" & vbCrLf
      strCont = strCont & "|#(自動代表圖)#|" & vbCrLf
   Else
      strCont = strCont & "" & vbCrLf
   End If
   '2009/04/09 End
   strCont = strCont & vbCrLf
   'add by sonia 2024/5/28 印度040商業使用聲明930，取消"往下掉1行"
   If strData(1) = "CFP" And strData(9) = "040" And p_stNP07 = "930" Then
   Else
   'end 2024/5/28
      strCont = strCont & vbCrLf
   End If    'add by sonia 2024/5/28
   'add  by nickc 2005/11/23 往下掉1行
   strCont = strCont & vbCrLf
   strCont = strCont & "　　　　　　　　　　　　　　　回覆者：　　　　　　　　　　　（簽章）" & vbCrLf
   strCont = strCont & vbCrLf
   strCont = strCont & vbCrLf
   strCont = strCont & "　　　　　　　　　　　　　　　回覆日期：　　　年　　　月　　　日" & vbCrLf
   strCont = strCont & vbCrLf
   strCont = strCont & vbCrLf
   'Added by Lydia 2020/05/25 CFT延展案：取消"往下掉2行"
   If strData(1) = "CFT" And (bstNP07 = "102" Or InStr(bstNP07, "102,") > 0) Then
   ElseIf strData(1) = "CFP" And strData(9) = "040" And p_stNP07 = "930" Then   'add by sonia 2024/5/28 印度040商業使用聲明930，取消"往下掉2行"
   Else
   'end 2020/05/25
        'add  by nickc 2005/11/23 往下掉2行
        strCont = strCont & vbCrLf
        strCont = strCont & vbCrLf
        'add  by Sindy 2009/04/10 往下掉2行
        If strData(1) = "T" And (iType = 3 Or iType = 6) Then
           strCont = strCont & vbCrLf
           strCont = strCont & vbCrLf
        End If
   End If 'Added by Lydia 2020/05/25
   
'   If strData(1) = "T" And (iType = 3 Or iType = 6) And strCU10 > "010" Then
'      strCont = strCont & vbCrLf
'   End If
   'Modify By Sindy 2009/04/09
   'If strData(1) = "T" Then
   'Modify By Sindy 2014/6/5 + And (strSales = "67002" Or strSales = "96029" Or strSales = "96030")
   'Modify by Amy 2017/01/10 +判斷"MCTF"開頭 的人員
   If strData(1) = "T" And (iType = 3 Or iType = 6) And strCU10 > "010" And _
      (strSales = "67002" Or strSales = "96029" Or strSales = "96030" Or Left(strSales, 4) = "MCTF") Then '大->台
      strCont = strCont & "※請將此案件回覆單打V填妥後傳真或E-Mail通知本所，以利本所作業，感" & vbCrLf
      strCont = strCont & "　謝您的合作與支持！" & vbCrLf
      strCont = strCont & "　本所傳真：886-2-25011666　E-Mail：tm@taie.com.tw;taie@seed.net.tw"
   Else
      'Modified by Lydia 2018/11/23 加入傳真號碼
      'strCont = strCont & "※請將此案件回覆單打V填妥後傳真通知本所，以利本所作業，感謝您的合作" & vbCrLf
      'strCont = strCont & "　與支持！"
      Select Case PUB_GetST06(strSales)
         Case "2"  '中
                strExc(1) = "04-23227483"
         Case "3"  '南
                strExc(1) = "06-2744030"
         Case "4"  '高
                strExc(1) = "07-2364360"
         Case Else     '剩下的歸北所
                strExc(1) = "02-25011666"
      End Select
      strCont = strCont & "※請將此案件回覆單打V填妥後傳真通知本所(傳真號碼：" & strExc(1) & ")，以利" & vbCrLf
      strCont = strCont & "　本所作業，感謝您的合作與支持！"
      'end 2018/11/23
   End If
   '2009/04/09 End
     
   'Modify by Morgan 2005/11/9 加控制只回傳不列印
   If p_bolNoPrint = True Then
      p_strArt = strCont
   Else
      'edit by nickc 2005/11/23 補本所案號 ，還有 P & CFP 的領正及繳年費才印
      'If DoPrintReturnSheet(strCont) = True Then
      'Modify By Sindy 2012/3/27
      'Modified by Lydia 2020/05/25 用Word直接列印
'      If bolSavePic = True Then
'         If strData(1) = "T" Then
'            'Call frm210126.DoPrintReturnSheet(strCont, strData(1), strData(2), strData(3), strData(4), True, p_bolPrintLetterHead)
'            Call ofrom.DoPrintReturnSheet(strCont, strData(1), strData(2), strData(3), strData(4), True, p_bolPrintLetterHead)
'         Else
'            'Call frm210126.DoPrintReturnSheet(strCont, strData(1), strData(2), strData(3), strData(4), IIf(strData(1) = "P", True, IIf(strData(1) = "CFP", True, False)) And IIf(iType = 2, True, False), p_bolPrintLetterHead)
'            Call ofrom.DoPrintReturnSheet(strCont, strData(1), strData(2), strData(3), strData(4), IIf(strData(1) = "P", True, IIf(strData(1) = "CFP", True, False)) And IIf(iType = 2, True, False), p_bolPrintLetterHead)
'         End If
'      Else
'      '2012/3/27 End
'         'Modify By Sindy 2009/05/01 +T
'         'Modify By Sindy 2012/4/11 +CFT
'         If strData(1) = "T" Or strData(1) = "CFT" Then
'            If DoPrintReturnSheet(strCont, strData(1), strData(2), strData(3), strData(4), True, p_bolPrintLetterHead) = True Then
'               If b_ShowMsg = True Then
'                  MsgBox "列印完畢！", vbInformation
'               End If
'            End If
'         '2009/05/01 End
'         ElseIf DoPrintReturnSheet(strCont, strData(1), strData(2), strData(3), strData(4), IIf(strData(1) = "P", True, IIf(strData(1) = "CFP", True, False)) And IIf(iType = 2, True, False), p_bolPrintLetterHead) = True Then
'            'add by nickc 2005/05/30
'            If b_ShowMsg = True Then
'               MsgBox "列印完畢！", vbInformation
'            End If
'         End If
'      End If
        '引用basLetter裡的變數及函數
        m_strFilePath = PUB_Getdesktop & "\" & strData(1) & strData(2) & strData(3) & strData(4) & "-案件回覆單.doc"
        If bolToWord = False Then
            m_bolSave2File = False
        Else
            m_bolSave2File = True
        End If
        strPrintType = "1"
        m_MySt(1) = strData(1)
        m_MySt(2) = strData(2)
        m_MySt(3) = strData(3)
        m_MySt(4) = strData(4)
        
        'Modify By Sindy 2022/4/22 MCT不需要產生回覆單紙本及電子檔
        'Memo by Lydia 2025/08/07 若取消MCT控制，請一併取消frm210126的控制
        If Mid(PUB_GetAKindSalesNo(strData(1), strData(2), strData(3), strData(4)), 1, 4) = "MCTF" Then
            Exit Sub
        End If
        
        '不寫入定稿資料檔
        If bolToWord = False Then
            'Add By Sindy 2022/3/30 T*回覆單要儲存至卷宗區
            'Modify By Sindy 2022/4/18 + And (TypeName(pForm) = "Nothing" Or UCase(TypeName(pForm)) <> UCase("FRM075007_2"))
            If Left(m_MySt(1), 1) = "T" And (TypeName(pForm) = "Nothing" Or UCase(TypeName(pForm)) <> UCase("FRM075007_2")) Then
               '直接列印, 一併上傳卷宗區
               ExportToMsWord strCont, True, False, True, 0, , , p_stNP01, , , , , , False, True, "BLANK"
            Else
            '2022/3/30 END
               '直接列印
               'Modified by Lydia 2020/12/08 不用更新列印註記; Debug:內商程序常有整批定稿中有幾筆定稿被上不列印註記。
               'ExportToMsWord strCont, True, False, True, 0
               ExportToMsWord strCont, True, False, True, 0, , , , , , , , , False
            End If
        Else
            '桌面存Word檔by frm210126
            ExportToMsWord strCont, True, False, False, 0
        End If
        'end 2020/05/25
   End If
End Sub

'edit by nickc 2005/11/23 補本所案號
'Private Function DoPrintReturnSheet(ByVal p_stContent As String) As Boolean
'Modify by Morgan 2011/5/24 +p_bolPrintLetterHead:是否列印信頭
'Modified by Lydia 2020/05/25 更名保留: DoPrintReturnSheet=>DoPrintReturnSheet_old
Private Function DoPrintReturnSheet(ByVal p_stContent As String, Optional oCP01 As String = "", Optional oCP02 As String = "", Optional oCP03 As String = "", Optional oCP04 As String = "", Optional IsPrintPic As Boolean = False, Optional p_bolPrintLetterHead As Boolean) As Boolean
   Dim iPos As Integer
   Dim iLineHeight As Integer
   Dim iMarginLeft As Integer
   Dim iNowPos As Integer
   Dim iScaleLeft As Integer
   Dim iScaleTop As Integer
   Dim iLineCount As Integer
   Dim tObj As New StdPicture
   Dim stHeadFilePath As String, stFootFilePath As String '圖檔路徑 Add by Morgan 2011/5/24
   Dim bytes() As Byte
   Dim file_num As Integer
   Dim rsPic As New ADODB.Recordset
   Dim pTop As Integer
   Dim pWidth As Integer
   Dim pWidthNew As Integer
   Dim PLeft As Integer
   Dim pHeight As Integer
   Dim pHeightNew As Integer
   Dim IsWmf As Boolean
   Dim bSuccess As Boolean
   
   DoPrintReturnSheet = False
   
   'Add by Morgan 2011/5/24
   If p_bolPrintLetterHead Then
      'Modified by Morgan 2014/8/27
      'PUB_ReadLetterPic oCP01, stHeadFilePath, stFootFilePath
      PUB_ReadLetterPic oCP01, oCP02, oCP03, oCP04, stHeadFilePath, stFootFilePath
   End If
   
   iScaleLeft = Printer.ScaleLeft
   iScaleTop = Printer.ScaleTop
   
On Error GoTo ErrHnd
   
   Printer.Orientation = 1
   Printer.PaperSize = vbPRPSA4
   Printer.Font.Size = 14
   Printer.ScaleLeft = -850
   Printer.ScaleTop = -1600
   
   'Add by Morgan 2011/5/24
   If p_bolPrintLetterHead Then
      If stHeadFilePath <> "" Then
         Set tObj = pvGetStdPicture(stHeadFilePath)
         'Modified by Morgan 2020/4/1 圖片高度改一原比例縮放
         'Printer.PaintPicture tObj, Printer.ScaleLeft, Printer.ScaleTop + 150, Printer.ScaleWidth, Printer.ScaleWidth / 6.2
         pHeight = Trunc(tObj.Height * Printer.ScaleWidth / tObj.Width)
         Printer.PaintPicture tObj, Printer.ScaleLeft, Printer.ScaleTop + 150, Printer.ScaleWidth, pHeight
         'end 2020/4/1
      End If
      If stFootFilePath <> "" Then
         Set tObj = pvGetStdPicture(stFootFilePath)
         'Modified by Morgan 2020/4/1 圖片高度改一原比例縮放
         'Printer.PaintPicture tObj, Printer.ScaleLeft, Printer.ScaleTop + Printer.ScaleHeight - Printer.ScaleWidth / 10 - 100, Printer.ScaleWidth, Printer.ScaleWidth / 10
         pHeight = Trunc(tObj.Height * Printer.ScaleWidth / tObj.Width)
         Printer.PaintPicture tObj, Printer.ScaleLeft, Printer.ScaleTop + Printer.ScaleHeight - pHeight - 200, Printer.ScaleWidth, pHeight
         'end 2020/4/1
      End If
   End If
   
   Printer.FontName = "標楷體"
   iPos = InStr(p_stContent, vbCrLf)
   iNowPos = 0
   iLineCount = 1
   iLineHeight = 360
   Do While iPos > 0
      Printer.CurrentY = iNowPos
      Printer.Print Left(p_stContent, iPos - 1)
      iNowPos = iNowPos + iLineHeight
      p_stContent = Mid(p_stContent, iPos + 1)
      iPos = InStr(p_stContent, vbCrLf)
      iLineCount = iLineCount + 1
   Loop
   Printer.CurrentY = iNowPos
   Printer.Print p_stContent
   'add by nickc 2005/04/15 測試印圖 * start    2005/11/23 正式加入使用
   If Trim(oCP01 & oCP02 & oCP03 & oCP04) <> "" And IsPrintPic = True Then
'        Dim GpInput As mGDIpEx.GdiplusStartupInput
'        GpInput.GdiplusVersion = 1
'        If (mGDIpEx.GdiplusStartup(m_GDIpToken, GpInput) = [OK]) Then
            '算幾行
            Set rsPic = New ADODB.Recordset
            rsPic.CursorLocation = adUseClient
            'Modify By Sindy 2018/10/31 TF馬德里商標圖檔-子案的圖同母案
            '固定都以IBF01=tm01 AND IBF02=substr(tm02,1,5)||'0' AND IBF03='0' AND IBF04='00' 去抓代表圖
            If oCP01 = "TF" Then
               rsPic.Open "select * from ImgByteFile where ibf01='" & oCP01 & "' and ibf02='" & Mid(oCP02, 1, 5) & "0" & "' and ibf03='0' And ibf04='00' and ibf05='1' ", cnnConnection, adOpenStatic, adLockOptimistic
            Else
            '2018/10/31 END
               rsPic.Open "select * from ImgByteFile where ibf01='" & oCP01 & "' and ibf02='" & oCP02 & "' and ibf03='" & oCP03 & "' and ibf04='" & oCP04 & "' and ibf05='1' ", cnnConnection, adOpenStatic, adLockOptimistic
            End If
            If rsPic.RecordCount <> 0 Then
               'edit by nickc 2006/01/13
               If CheckStr(rsPic.Fields("ibf06")) = "1" Or CheckStr(rsPic.Fields("ibf06")) = "2" Then
                  IsWmf = False
               Else
                  IsWmf = True
               End If
               
               'Add By Sindy 2017/8/10
'               If "" & rsPic.Fields("IBF15") <> "" Then
                  If PUB_GetFtpFile(rsPic.Fields("IBF15"), App.path & "\" & oCP01 & oCP02 & oCP03 & oCP04 & "." & IIf(IsWmf, "wmf", "jpg"), UCase("ImgByteFile")) = False Then
                     Exit Function
                  End If
'               Else
'               '2017/8/10 END
'                  ReDim bytes(Val(rsPic.Fields("ibf13").Value))
'                  bytes() = rsPic.Fields("ibf14").GetChunk(Val(rsPic.Fields("ibf13").Value))
'                  file_num = FreeFile
'                  'edit by nickc 2006/03/28
'                  'Open App.Path & "\tmp." & IIf(IsWmf, "wmf", "tif") For Binary Access Write As #file_num
'                  'edit by nickc 2007/11/27 指定檔名
'                  'Open App.Path & "\tmp." & IIf(IsWmf, "wmf", "jpg") For Binary Access Write As #file_num
'                  Open App.path & "\" & oCP01 & oCP02 & oCP03 & oCP04 & "." & IIf(IsWmf, "wmf", "jpg") For Binary Access Write As #file_num
'                     Put #file_num, , bytes()
'                  Close #file_num
'               End If
               
               'edit by nickc 2006/03/28
               'Set tObj = pvGetStdPicture(App.Path & "\tmp." & IIf(IsWmf, "wmf", "tif"))
               'edit by nickc 2006/05/24
               'Set tObj = pvGetStdPicture(App.Path & "\tmp." & IIf(IsWmf, "wmf", "jpg"), bSuccess)
               'edit by nickc 2007/11/27 指定檔名
               'Set tObj = pvGetStdPicture(App.Path & "\tmp." & IIf(IsWmf, "wmf", "jpg"))
               Set tObj = pvGetStdPicture(App.path & "\" & oCP01 & oCP02 & oCP03 & oCP04 & "." & IIf(IsWmf, "wmf", "jpg"))
'               If bSuccess Then
                    pTop = 7780 + ((iLineCount - 35) * iLineHeight)
                    PLeft = 0
                    pWidth = 4200
                    pHeight = 4380
                    If tObj.Height >= tObj.Width Then  '以高為準
                        pWidthNew = tObj.Width / (tObj.Height / pHeight)
                        PLeft = PLeft + ((pWidth - pWidthNew) / 2)
                        pWidth = pWidthNew
                    Else   '已寬為準
                        pHeightNew = tObj.Height / (tObj.Width / pWidth)
                        pTop = pTop + ((pHeight - pHeightNew) / 2)
                        pHeight = pHeightNew
                    End If
                    '因為標準是 35 行 所以總行數 - 35 在乘上行高
                    'Printer.PaintPicture tObj, PLeft, pTop, pWidth, pHeight
                    'Modify By Sindy 2009/05/01
                    'Printer.PaintPicture tObj, 100, pTop, pWidth - 1500, pHeight - 1500
                    'Modify By Sindy 2012/4/11 改乘於比例的方式來調整圖片大小
                    Printer.PaintPicture tObj, 100, pTop, (pWidth / 100 * 85), (pHeight / 100 * 85)
'                End If
'            End If
'            Call mGDIpEx.GdiplusShutdown(m_GDIpToken)
            'Modify by Nick 2005/12/26
            'edit by nickc 2006/03/28
            'If Dir(App.Path & "\tmp." & IIf(IsWmf, "wmf", "tif")) <> "" Then
            '   Kill App.Path & "\tmp." & IIf(IsWmf, "wmf", "tif")
            'edit by nickc 2007/11/27 指定檔名
            'If Dir(App.Path & "\tmp." & IIf(IsWmf, "wmf", "jpg")) <> "" Then
            '   Kill App.Path & "\tmp." & IIf(IsWmf, "wmf", "jpg")
            If Dir(App.path & "\" & oCP01 & oCP02 & oCP03 & oCP04 & "." & IIf(IsWmf, "wmf", "jpg")) <> "" Then
               Kill App.path & "\" & oCP01 & oCP02 & oCP03 & oCP04 & "." & IIf(IsWmf, "wmf", "jpg")
            End If
        End If
   End If
   '*********end
   Printer.EndDoc
   
   DoPrintReturnSheet = True
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   Printer.ScaleLeft = iScaleLeft
   Printer.ScaleTop = iScaleTop
End Function


'Add by Morgan 2005/2/22 讀取下一程序資料
'Add By Sindy 2017/4/25 + Optional ByRef p_stNP15 As String
'Modified by Lydia 2020/05/25 + 本所期限 p_stNP08 as string
Private Function ReadNPData(ByRef p_stNP01 As String, ByRef p_stNP09 As String, ByRef p_stCPM03s As String, ByRef p_stData() As String, _
Optional ByRef p_stCaseNo As String, Optional ByRef p_stNP07 As String, Optional ByRef p_stNP15 As String, Optional ByRef p_stNP08 As String) As Boolean
   
   Dim stCon As String
   
On Error GoTo ErrHnd

   stCon = ""
   If p_stNP09 <> Empty Then
     stCon = " and NP09=" & p_stNP09
   End If
   
   'Add byMorgan 2005/11/15
   '本所案號
   If p_stCaseNo <> Empty Then
      stCon = stCon & " and " & ChgNextProgress(p_stCaseNo)
   End If
   
   '案件性質,多個時用逗號(,)區隔
   If p_stNP07 <> Empty Then
      stCon = stCon & " and instr('" & p_stNP07 & "',NP07)>0"
   End If

   'Add by Morgan 2005/11/16
   If p_stNP01 <> Empty Then
      stCon = stCon & " and NP01='" & p_stNP01 & "'"
   '收文號及本所號不可皆為空白
   ElseIf p_stCaseNo = Empty Then
      Exit Function
   End If
   
    ReadNPData = False
    'Modify by Morgan 2006/5/15 CPM03='（無）'時改抓CPM04
    'Modify By Sindy 2017/4/25 + NP15
    'Modified by Lydia 2020/05/25 +NP08, NP09
    strSql = "select NP02, NP03, NP04, NP05, NP07, DECODE(CPM03,'（無）',CPM04,CPM03) CPM03, NP15,NP08,NP09 " & _
                "from NextProgress,CasePropertyMap " & _
                "Where NP06 is NULL " & stCon & _
                " and CPM01(+)=NP02 and CPM02=NP07 order by np07"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount > 0 Then
         p_stData(1) = .Fields("NP02")
         p_stData(2) = .Fields("NP03")
         p_stData(3) = .Fields("NP04")
         p_stData(4) = .Fields("NP05")
         'Add by Morgan 2005/12/26 需抓實際案件性質以便判斷回覆單種類
         p_stNP07 = .Fields("NP07")
         p_stNP15 = "" & .Fields("NP15") 'Add By Sindy 2017/4/25
         p_stNP08 = "" & .Fields("NP08") 'Added by Lydia 2020/05/25
         Do While Not .EOF
            'Modify by Morgan 2008/4/18 先存代碼,再依國家抓中文
            'p_stCPM03s = p_stCPM03s & IIf(p_stCPM03s <> "", ",", "") & .Fields("CPM03")
            p_stCPM03s = p_stCPM03s & IIf(p_stCPM03s <> "", ",", "") & .Fields("NP07")
            .MoveNext
         Loop
         ReadNPData = True
      End If
   End With
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Add by Morgan 2005/2/22
'iPreLen=前置字串長度
Private Function SplitWords(ByVal stTitle As String, ByVal iPreLen As Integer, Optional iMaxLen As Integer = 66) As String
   Dim iTLen As Integer '字串總長度
   Dim arrWords '單字陣列
   Dim stRtn As String '回傳字串
   Dim iWordLen As Single  '單字長度
   Dim iLineLen As Integer '一列字串長度
   Dim ii As Integer, jj As Integer
   Dim stChar As String
   Dim iRestLen As Integer '列剩餘可印長度
   Dim iWordPos As Integer '列剩餘可印字串位置
   Dim iCharLen As Integer '字元長度

   stTitle = Trim(stTitle)
   iTLen = Len(stTitle)
   
   If iTLen > 0 Then
      '拆字
      arrWords = Split(stTitle, " ")
      iLineLen = iPreLen: stRtn = ""
      
      For ii = LBound(arrWords) To UBound(arrWords)
         arrWords(ii) = Trim(arrWords(ii))
         iRestLen = iMaxLen - iLineLen
         If arrWords(ii) <> "" Then
            iWordLen = 0
            For jj = 1 To Len(arrWords(ii))
               stChar = Mid(arrWords(ii), jj, 1)
               iWordLen = iWordLen + 1
               '全形字算2個
               If Asc(stChar) < 0 Then
                  iWordLen = iWordLen + 1
               'Add by Morgan 2005/11/11
               '英文大寫字元算1.3個
               ElseIf Asc(stChar) >= 65 And Asc(stChar) <= 90 Then
                  iWordLen = iWordLen + 0.3
               '2005/11/11
               End If
            Next jj
            
            '若字串超過列可印長度時斷字
            If iPreLen + iWordLen > iMaxLen Then
               iWordLen = 0: iWordPos = 0
               For jj = 1 To Len(arrWords(ii))
                  stChar = Mid(arrWords(ii), jj, 1)
                  '全形字算2個
                  If Asc(stChar) < 0 Then
                     iCharLen = 2
                  '英文大寫字元算1.5個
                  ElseIf Asc(stChar) >= 65 And Asc(stChar) <= 90 Then
                     iCharLen = 1.5
                  '其餘算1個
                  Else
                     iCharLen = 1
                  End If
                  iWordLen = iWordLen + iCharLen
                  If iLineLen + iWordLen > iMaxLen Then
                     iWordPos = iWordPos + 1
                     If jj - iWordPos > 0 Then
                        stRtn = stRtn & Mid(arrWords(ii), iWordPos, jj - iWordPos)
                     End If
                     iWordPos = jj
                     stRtn = stRtn & vbCrLf & String(iPreLen, " ") & Mid(arrWords(ii), iWordPos, 1)
                     iLineLen = iPreLen + iCharLen
                     iWordLen = 0
                  End If
               Next jj
               arrWords(ii) = Mid(arrWords(ii), iWordPos + 1)
            End If
            
            '若超過每行固定長度則跳行
            If iLineLen + iWordLen > iMaxLen Then
               stRtn = stRtn & vbCrLf & String(iPreLen, " ") & arrWords(ii)
               iLineLen = iPreLen + iWordLen
            Else
               stRtn = stRtn & arrWords(ii)
               iLineLen = iLineLen + iWordLen
            End If
            '若不是最後一個字時加空白
            If iLineLen < iMaxLen Then
               stRtn = stRtn & " "
               iLineLen = iLineLen + 1
            End If
         End If
      Next ii
   End If
   SplitWords = stRtn
End Function
'Add by Morgan 2007/2/13
Private Function GetCP10(p_CP09 As String) As String
   strExc(0) = "select cp10 from caseprogress where cp09='" & p_CP09 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      GetCP10 = "" & RsTemp.Fields(0)
   End If
End Function
'add by nickc 2008/01/18 抓特殊客戶及業務備註
Private Sub GetOtherCuData(p_Cp01 As String, p_Cp02 As String, p_Cp03 As String, p_Cp04 As String)
Dim m_rs As New ADODB.Recordset
Dim m_rs1 As New ADODB.Recordset
Dim m_str As String
Dim m_i As Integer
m_CuMemo = ""
m_SpecCu = ""
'Modify By Sindy 2011/2/21 增加LC43,LC44,LC45,LC46,HC24,HC25,HC26,HC27
m_str = "select pa26,pa27,pa28,pa29,pa30 from patent where pa01='" & p_Cp01 & "' and pa02='" & p_Cp02 & "' and pa03='" & p_Cp03 & "' and pa04='" & p_Cp04 & "' "
m_str = m_str & " union select tm23,tm78,tm79,tm80,tm81 from trademark where tm01='" & p_Cp01 & "' and tm02='" & p_Cp02 & "' and tm03='" & p_Cp03 & "' and tm04='" & p_Cp04 & "' "
m_str = m_str & " union select lc11,lc43,lc44,lc45,lc46 from lawcase where lc01='" & p_Cp01 & "' and lc02='" & p_Cp02 & "' and lc03='" & p_Cp03 & "' and lc04='" & p_Cp04 & "' "
m_str = m_str & " union select hc05,hc24,hc25,hc26,hc27 from hirecase where hc01='" & p_Cp01 & "' and hc02='" & p_Cp02 & "' and hc03='" & p_Cp03 & "' and hc04='" & p_Cp04 & "' "
m_str = m_str & " union select sp08,sp58,sp59,sp65,sp66 from servicepractice where sp01='" & p_Cp01 & "' and sp02='" & p_Cp02 & "' and sp03='" & p_Cp03 & "' and sp04='" & p_Cp04 & "' "
If m_rs.State = 1 Then m_rs.Close
m_rs.CursorLocation = adUseClient
m_rs.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
If Not m_rs.EOF And Not m_rs.BOF Then
    m_rs.MoveFirst
    Do While Not m_rs.EOF
        For m_i = 0 To 4
                m_str = "select cu01,cu02,cu121,nvl(cu04,nvl(cu05||cu88||cu89||cu90,cu06)) cu04,cl04,cu125 from customer, (select * from custspeciallog  where (cl01,cl02,cl03) in (select cl01,cl02,max(cl03) from custspeciallog where cl01='" & CheckStr(m_rs.Fields(m_i)) & "' " & _
                             " and cl02 in (select max(cl02) from custspeciallog where cl01='" & CheckStr(m_rs.Fields(m_i)) & "' ) group by cl01,cl02)) Ncl " & _
                             " Where CU01='" & Mid(CheckStr(m_rs.Fields(m_i)), 1, 8) & "' And CU02='" & Mid(CheckStr(m_rs.Fields(m_i)), 9, 1) & "' and cu01||cu02=Ncl.cl01(+) "
                If m_rs1.State = 1 Then m_rs1.Close
                m_rs1.CursorLocation = adUseClient
                m_rs1.Open m_str, cnnConnection, adOpenStatic, adLockReadOnly
                If Not m_rs1.EOF And Not m_rs1.BOF Then
                    If CheckStr(m_rs1.Fields("cu121")) = "Y" Then
                        If m_SpecCu = "" Then    '秀玲說只要記錄第一筆
                            m_SpecCu = CheckStr(m_rs1.Fields("cl04"))
                        End If
                    End If
                    If CheckStr(m_rs1.Fields("cu125")) <> "" Then
                        If m_CuMemo = "" Then    '記錄第一筆
                            m_CuMemo = CheckStr(m_rs1.Fields("cu125"))
                        End If
                    End If
                End If
        Next m_i
        m_rs.MoveNext
    Loop
End If
End Sub

'Added by Lydia 2019/03/14 列印C類接洽記錄單(Word版)
'Modified by Morgan 2020/9/15 +strPrintMemo(回傳實際列印的C類備註)
Public Sub PrintCFormNew(ByVal strCP09 As String, Optional ByVal strMemo As String, Optional ByVal strCP48Desc As String, Optional bolSavePdf As Boolean = False, Optional ByRef strPrintMemo As String)
'strCP09: 總收文號
'strMemo: 傳入備註
'strCP48Desc: 設定FMP案的承辦期限
'bolSavePdf: 轉pdf

'列印格式
'收文日期 : Field(0)        承辦人 : Field(1)
'總收文號 : Field(2)        管制人 : Field(3)
'本所期限 : Field(4)        智權人員 : Field(5)
'......
'......
'專利名稱 : Field(10)
'代理人名稱(英文) : Field(11)
'......
'......

'Dim m_CharWidth As Long
'Dim m_CharHeight As Long
'Dim strData As String
'Dim nCount As Integer
'' 報表左方留白的寬度
'Dim m_LeftMargin As Integer
'' 報表上方留白的高度
'Dim m_TopMargin As Integer
Dim nRow As Integer '目前列位置
Dim PageRow As Integer '頁面列數
Dim iLHeight As Double '每列固定高度

Dim strSQL1 As String
Dim rs1 As New ADODB.Recordset
Dim Rs2 As New ADODB.Recordset
Dim jj As Integer
Dim strCustNo As String
Dim varCustNo As Variant
Dim ii As Integer
Dim bolFMP As Boolean
Dim bolIsSyngentaOneDayCase As Boolean
Dim strCOrderPath As String
Dim strPdfName As String
Dim oFileSys As New FileSystemObject
Dim oFile
Dim bolOpenWord As Boolean

'設定SQL字串
'Modified by Lydia 2021/08/31 +CP14
strSQL1 = " Select CP05, S1.ST02, CP09, S2.ST02, CP06, S3.ST02, " & _
         " CP07, CP01||'-'||CP02||'-'||CP03||'-'||CP04 CNo, CP48, decode(cp10,'1503',decode(pa09,'000',CPM03,CPM04)||decode(cp24,'1','－准','2','－駁',cp24),decode(pa09,'000',CPM03,CPM04)) CPM03, " & _
         " PA05, FA05||' '||FA63, FA04, FA06, CU05||' '||CU88, CU04, CU06,PA75,PA26,CP01,CP10,CP43,PA27,PA28,PA29,PA30,CP12,NP23,CP40,CP41,CP42 " & _
         ",CP133,CP134,CP24,CP02,CP03,CP04,CP14 From CaseProgress,Patent,Staff S1,Nation,Staff S2,Staff S3,CasePropertyMap,Fagent,Customer,nextprogress" & _
         " Where CP01=PA01(+) And CP02=PA02(+) And CP03=PA03(+) And CP04=PA04(+) " & _
         " And CP14=S1.ST01(+) And PA09=NA01(+) And NA16=S2.ST01(+) And CP13=S3.ST01(+) " & _
         " And CP01=CPM01(+) And CP10=CPM02(+) And substr(PA75,1,8)=FA01(+) And substr(PA75,9,1)=FA02(+) " & _
         " And substr(PA26,1,8)=CU01(+) And substr(PA26,9,1)=CU02(+) " & _
         " And CP09 = '" & "" & strCP09 & "' AND CP01 in ('FCP','P') and np01(+)=cp09 and np06(+) is null"
strSQL1 = strSQL1 & " UNION Select CP05, S1.ST02, CP09, S2.ST02, CP06, S3.ST02, " & _
         " CP07, CP01||'-'||CP02||'-'||CP03||'-'||CP04 CNo, CP48, decode(sp09,'000',CPM03,CPM04) CPM03, " & _
         " SP05, FA05||' '||FA63, FA04, FA06, CU05||' '||CU88, CU04, CU06,SP26,SP08,CP01,CP10,CP43,SP58,SP59,SP65,SP66,CP12,NP23,CP40,CP41,CP42 " & _
         ",CP133,CP134,CP24,CP02,CP03,CP04,CP14 From CaseProgress,SERVICEPRACTICE,Staff S1,Nation,Staff S2,Staff S3,CasePropertyMap,Fagent,Customer,nextprogress" & _
         " Where CP01=SP01(+) And CP02=SP02(+) And CP03=SP03(+) And CP04=SP04(+) " & _
         " And CP14=S1.ST01(+) And SP09=NA01(+) And NA16=S2.ST01(+) And CP13=S3.ST01(+) " & _
         " And CP01=CPM01(+) And CP10=CPM02(+) And substr(SP26,1,8)=FA01(+) And substr(SP26,9,1)=FA02(+) " & _
         " And substr(SP08,1,8)=CU01(+) And substr(SP08,9,1)=CU02(+) " & _
         " And CP09 = '" & "" & strCP09 & "' AND CP01='FG' and np01(+)=cp09 and np06(+) is null"
If rs1.State <> adStateClosed Then rs1.Close
Set rs1 = Nothing
rs1.CursorLocation = adUseClient
rs1.Open strSQL1, cnnConnection, adOpenStatic, adLockReadOnly
'若有資料則執行列印
If rs1.RecordCount > 0 Then
ReOpenWord:
   'Word: A4紙張,直印
   If TypeName(g_WordAp) <> "Application" Then Set g_WordAp = New Word.Application
    
   iLHeight = 14.5
   g_WordAp.Documents.add
   With g_WordAp.Application
        .Visible = True
        .WindowState = wdWindowStateNormal
        .ActiveWindow.ActivePane.View.Zoom.Percentage = 100
        .Selection.Font.Name = "細明體" 'Added by Lydia 2019/03/19
       '邊界
       .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1.2)
       .Selection.PageSetup.RightMargin = .CentimetersToPoints(1.2)
       .Selection.PageSetup.TopMargin = .CentimetersToPoints(1.2)
       .Selection.PageSetup.BottomMargin = .CentimetersToPoints(1.2)
       '新增表格(1*2)
       .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=2
       With .Selection.Tables(1)
          .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
          .Borders(wdBorderRight).LineStyle = wdLineStyleNone
          .Borders(wdBorderTop).LineStyle = wdLineStyleNone
          .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
          .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
          .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
          .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
          .Borders.Shadow = False
       End With
       .Selection.SelectRow
       .Selection.Cells.VerticalAlignment = wdAlignVerticalCenter
       '設定表格高度, 欄位寬度
       .Selection.Cells.SetHeight RowHeight:=iLHeight, HeightRule:=wdRowHeightExactly
       .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(9), RulerStyle:=wdAdjustProportional
'-------------Word輸入---------------------------
        '輸入內容
       .Selection.InsertRows 6
       PageRow = 7
       .Selection.Collapse Direction:=wdCollapseStart
       'Line1. 標題 : C類接洽記錄單
       nRow = nRow + 1
       .Selection.SelectRow
       .Selection.Cells.Merge
       .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
       .Selection.Cells.SetHeight RowHeight:=28, HeightRule:=wdRowHeightExactly
       .Selection.Font.Size = 14
       .Selection.TypeText Text:=Left(strCP09, 1) & "　類　接　洽　記　錄　單"
        '多空一行
       .Selection.MoveRight Unit:=wdCharacter, Count:=5
        'Line3. 收文日期, 承辦人
        nRow = nRow + 2
       .Selection.SelectRow
       .Selection.Font.Size = 12
       .Selection.SelectRow
       .Selection.Collapse Direction:=wdCollapseStart
       .Selection.TypeText Text:="收文日期：" & IIf(IsNull(rs1.Fields(0).Value) Or Len("" & rs1.Fields(0).Value) <= 0, "", _
                                  (Left(rs1.Fields(0).Value, 4) - 1911) & "/" & Mid(rs1.Fields(0).Value, 5, 2) & "/" & Right(rs1.Fields(0).Value, 2))
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       .Selection.TypeText Text:="承辦人　：" & rs1.Fields(1).Value
       'Line4. 總收文號, 管制人
       nRow = nRow + 1
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       .Selection.TypeText Text:="總收文號：" & rs1.Fields(2).Value
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       
       '控制國外部收文的案件才印
       If Left("" & rs1.Fields("CP12").Value, 1) = "F" Then
          If "" & rs1.Fields("CP01") = "P" Then bolFMP = True
          'Modified by Lydia 2022/10/19 改成共用模組
'          '若有FC代理人
'          If Len("" & rs1.Fields(17).Value) > 0 Then
'             strSQL1 = " SELECT ST02 FROM FAGENT,NATION,STAFF WHERE SUBSTR('" & rs1.Fields(17).Value & "',1,8)=FA01(+) AND SUBSTR('" & rs1.Fields(17).Value & "',9,1)=FA02(+) AND FA10=NA01(+) AND NA16=ST01(+)"
'             If Rs2.State <> adStateClosed Then Rs2.Close
'             Set Rs2 = Nothing
'             Rs2.CursorLocation = adUseClient
'             Rs2.Open strSQL1, cnnConnection, adOpenStatic, adLockReadOnly
'             If Rs2.RecordCount > 0 Then
'                 .Selection.TypeText Text:="管制人　：" & Rs2.Fields(0).Value
'             End If
'             If Rs2.State <> adStateClosed Then Rs2.Close
'             Set Rs2 = Nothing
'
'          '若有申請人
'          ElseIf Len("" & rs1.Fields(18).Value) > 0 Then
'             strSQL1 = " SELECT ST02 FROM CUSTOMER,NATION,STAFF WHERE SUBSTR('" & rs1.Fields(18).Value & "',1,8)=CU01(+) AND SUBSTR('" & rs1.Fields(18).Value & "',9,1)=CU02(+) AND CU10=NA01(+) AND NA16=ST01(+)"
'             If Rs2.State <> adStateClosed Then Rs2.Close
'             Set Rs2 = Nothing
'             Rs2.CursorLocation = adUseClient
'             Rs2.Open strSQL1, cnnConnection, adOpenStatic, adLockReadOnly
'             If Rs2.RecordCount > 0 Then
'                 .Selection.TypeText Text:="管制人　：" & Rs2.Fields(0).Value
'             End If
'             If Rs2.State <> adStateClosed Then Rs2.Close
'             Set Rs2 = Nothing
'          Else
'             Printer.Print "管制人　："
'          End If
         strSQL1 = PUB_GetFCPHandler(rs1.Fields("cp01"), rs1.Fields("cp02"), rs1.Fields("cp03"), rs1.Fields("cp04"), "" & rs1.Fields("cp10"))
          If strSQL1 <> "" Then
              .Selection.TypeText Text:="管制人　：" & GetStaffName(strSQL1, True)
          Else
              .Selection.TypeText Text:="管制人　："
          End If
          'end 2022/10/19
       Else
          '非FCP
       End If
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       
       'Line5. 本所期限, 智權人員
       nRow = nRow + 1
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       .Selection.TypeText Text:="本所期限：" & IIf(IsNull(rs1.Fields(4).Value), "", _
                                  (Left(rs1.Fields(4).Value, 4) - 1911) & "/" & Mid(rs1.Fields(4).Value, 5, 2) & "/" & Right(rs1.Fields(4).Value, 2))
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       .Selection.TypeText Text:="智權人員：" & rs1.Fields(5).Value
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       
       'Line6. 法定期限, 本所案號
       nRow = nRow + 1
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       .Selection.TypeText Text:="法定期限：" & IIf(IsNull(rs1.Fields(6).Value), "", _
                                  (Left(rs1.Fields(6).Value, 4) - 1911) & "/" & Mid(rs1.Fields(6).Value, 5, 2) & "/" & Right(rs1.Fields(6).Value, 2))
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       .Selection.TypeText Text:="本所案號：" & rs1.Fields(7).Value
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       
       'Line7. 承辦期限, 案件性質
       nRow = nRow + 1
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       If bolFMP And strCP48Desc = "" Then
          If rs1.Fields("cp133") > 0 And rs1.Fields("cp134") > 0 Then
             PUB_GetFmpCP48 rs1.Fields("cp05"), rs1.Fields("cp06"), rs1.Fields("cp07"), rs1.Fields("cp133"), rs1.Fields("cp134"), , strCP48Desc
          End If
       End If
       .Selection.TypeText Text:="承辦期限：" & IIf(IsNull(rs1.Fields(8).Value) Or Len("" & rs1.Fields(8).Value) <= 0, "", _
                                  (Left(rs1.Fields(8).Value, 4) - 1911) & "/" & Mid(rs1.Fields(8).Value, 5, 2) & "/" & Right(rs1.Fields(8).Value, 2)) & IIf(strCP48Desc = "", "", "(" & strCP48Desc & ")")
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       .Selection.TypeText Text:="案件性質：" & rs1.Fields(9).Value
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
    
       'Line8. FMP+約定期限
       If Left("" & rs1.Fields("CP12").Value, 1) = "F" Then
            nRow = nRow + 1
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            If nRow > PageRow Then
                .Selection.InsertRows
                PageRow = PageRow + 1
            End If
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.Font.Bold = True
            .Selection.TypeText Text:="約定期限：" & ChangeWStringToTDateString("" & rs1.Fields("NP23").Value)
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
       End If
       
       'Line9/Line8. 專利名稱(中文)
       nRow = nRow + 1
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       .Selection.InsertRows
       PageRow = PageRow + 1
       .Selection.Font.Bold = False
       .Selection.TypeText Text:="專利名稱(中文)：" & rs1.Fields(10).Value
       .Selection.SelectRow
       .Selection.Cells.Merge
       .Selection.ParagraphFormat.Alignment = wdAlignParagraphLeft
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       .Selection.InsertRows 5
       PageRow = PageRow + 5
       .Selection.Collapse Direction:=wdCollapseStart
       '專利名稱(中文): 調整列高為自動
       .Selection.MoveUp Unit:=wdLine, Count:=1
       .Selection.SelectRow
       .Selection.Cells.SetHeight RowHeight:=iLHeight + 4, HeightRule:=wdRowHeightAuto
       .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly '固定行距
       .Selection.ParagraphFormat.LineSpacing = 14
       .Selection.MoveDown Unit:=wdLine, Count:=2  '空1行
       
       'Line11/Line 10 代理人名稱(英文)
       nRow = nRow + 2
       .Selection.TypeText Text:="代理人名稱(英文)：" & rs1.Fields(11).Value
       'Line12/Line 11 代理人名稱(中文)
       nRow = nRow + 1
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       .Selection.TypeText Text:="　　　　　(中文)：" & rs1.Fields(12).Value
       'Line13/Line 14 代理人名稱(日文)
       nRow = nRow + 1
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       .Selection.TypeText Text:="　　　　　(日文)：" & rs1.Fields(13).Value
       'Line14/Line 15 申請人1名稱(英文)  , P案紙張可印空間不夠，申請人不印
       If Left("" & rs1.Fields("CP12").Value, 1) = "F" Then
           nRow = nRow + 1
          .Selection.MoveRight Unit:=wdCharacter, Count:=1
          intI = 1
          '申請人(數)
          For jj = 27 To 30
              If "" & rs1.Fields("PA" & jj) <> "" Then
                  intI = intI + 1
              End If
          Next jj
          .Selection.InsertRows (intI * 3)
          PageRow = PageRow + (intI * 3)
          .Selection.Collapse Direction:=wdCollapseStart
          '申請人1
           nRow = nRow + 1
          .Selection.MoveRight Unit:=wdCharacter, Count:=1
          .Selection.TypeText Text:="申請人名稱(英文)：" & rs1.Fields(14).Value
           nRow = nRow + 1
          .Selection.MoveRight Unit:=wdCharacter, Count:=1
          .Selection.TypeText Text:="　　　　　(中文)：" & rs1.Fields(15).Value
           nRow = nRow + 1
          .Selection.MoveRight Unit:=wdCharacter, Count:=1
          .Selection.TypeText Text:="　　　　　(日文)：" & rs1.Fields(16).Value
          
          '申請人2~5
           For jj = 27 To 30
                If "" & rs1.Fields("PA" & jj) <> "" Then
                   strExc(0) = "select  CU05||' '||CU88, CU04, CU06 from customer where cu01||cu02='" & rs1.Fields("PA" & jj) & "'"
                   intI = 1
                   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                   If intI = 1 Then
                      nRow = nRow + 1
                      .Selection.MoveRight Unit:=wdCharacter, Count:=1
                      .Selection.TypeText Text:="　　　　　(英文)：" & RsTemp.Fields(0).Value
                      nRow = nRow + 1
                      .Selection.MoveRight Unit:=wdCharacter, Count:=1
                      .Selection.TypeText Text:="　　　　　(中文)：" & RsTemp.Fields(1).Value
                      nRow = nRow + 1
                      .Selection.MoveRight Unit:=wdCharacter, Count:=1
                      .Selection.TypeText Text:="　　　　　(日文)：" & RsTemp.Fields(2).Value
                   End If
                End If
           Next jj
       End If
       
       'Added by Lydia 2021/11/23 依David的要求，C類來函請於email及接洽單上，加上一句話「※請工程師務必進系統看【各項指示】內容，謝謝。」放置位置在備註的上面並且獨立一行，並非併入備註內容。
       If "" & rs1.Fields("CP01") = "FCP" Or bolFMP = True Then
          nRow = nRow + 1
          .Selection.MoveRight Unit:=wdCharacter, Count:=1
          .Selection.InsertRows 3
          PageRow = PageRow + 3
          .Selection.Collapse Direction:=wdCollapseStart
          .Selection.MoveDown Unit:=wdLine, Count:=1 '空1行
          .Selection.TypeText Text:="※請工程師務必進系統看【各項指示】內容，謝謝。"
       End If
       'end 2021/11/23
       
       '備註
       nRow = nRow + 2
       .Selection.MoveRight Unit:=wdCharacter, Count:=1
       If nRow > PageRow Then
           .Selection.InsertRows (nRow - PageRow)
           PageRow = PageRow + (nRow - PageRow)
           .Selection.Collapse Direction:=wdCollapseStart
       End If
       .Selection.MoveDown Unit:=wdLine, Count:=1 '空1行

          '核駁函接洽單上加印備註(如果沒有傳入備註)
       'Modified by Lydia 2022/01/21 備註改成共用模組basUpdate.PUB_FCPCFormMemo;  因為輸入C類來函通知會一併發email給承辦工程師
'          If strMemo = "" Then
'             '改變原處分-核駁也要印
'             If (rs1("CP10") = "1002" Or rs1("CP10") = "1007" Or (rs1("CP10") = "1503" And "" & rs1("CP24") = "2")) And "" & rs1("CP43") <> "" Then
'                strExc(0) = "SELECT NP07 FROM NEXTPROGRESS WHERE NP01='" & strCP09 & "'"
'                intI = 1
'                Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                If intI = 1 Then
'                   Select Case RsTemp.Fields(0)
'                      Case "501" '訴願
'                         strMemo = "請客戶提供重新簽署的委任書以俾辦理訴願程序"
'                      Case "503" '行政訴訟
'                         strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理行政訴訟程序"
'                      Case "504" '行政再審
'                         strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理行政再審程序"
'                      Case "507" '行政訴訟上訴
'                         strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理行政訴訟上訴程序"
'                      Case "509" '抗告
'                         strMemo = "請客戶提供重新簽署具公、簽證的委任書以俾辦理抗告程序"
'                   End Select
'                End If
'             End If
'          End If
'
'         'Removed by Morgan 2020/9/15 該案已至年費階段並閉卷
'         'If rs1("CP10") = "1202" And rs1(7) = "FCP-043077-0-00" Then
'         '    If strMemo <> "" Then strMemo = strMemo & vbCrLf
'         '    strMemo = strMemo & "工程師通知分割時間限制"
'         'End If
'         'end 2020/9/15
'
'         If InStr("Y4830900,Y4830901,Y4830902,Y4830903,Y4830904,Y4830905,Y4830908,Y5132600", Left("" & rs1("pa75"), 8)) > 0 Then
'             bolIsSyngentaOneDayCase = ChkSyngentaNP(strCP09)
'             If bolIsSyngentaOneDayCase = True Then
'                If strMemo <> "" Then strMemo = strMemo & vbCrLf
'                strMemo = strMemo & "收到舉發案訴願或任何法律行動，請於收到的24小時內告知先正達。"
'             End If
'         End If
'          '核駁及審查意見通知函備註:  排除先正達承辦期限為2個工作天(24Hr)的來函
'          'Modified by Lydia 2021/08/25 國外部凡是C類工程師的來函(排除核准1001、核發1008，另外非核准和一般來函性質1204,1217,1913,1603,1604)，有設核駁及審查意見通知函備註皆要帶備註到接洽單
'          'If bolIsSyngentaOneDayCase = False And (rs1("cp01") = "P" Or rs1("cp01") = "FCP") And (rs1("cp10") = "1202" Or rs1("cp10") = "1227" Or rs1("cp10") = "1002") Then
'          'Modified by Lydia 2021/08/31 +判斷來函承辦人為工程師 And PUB_GetST03("" & rs1.Fields("cp14")); ex.發生了FCP065502通知即將公開1207直接收文告代
'          'Modified by Lydia 2021/08/31 +排除926,和非C類接洽單And Left("" & rs1.Fields(2), 1) = "C"
'          If bolIsSyngentaOneDayCase = False And PUB_GetST03("" & rs1.Fields("cp14")) = "F21" And Left("" & rs1.Fields(2), 1) = "C" And _
'                        ((rs1.Fields("cp01") = "FCP" Or bolFMP = True) And InStr("1001,1008,1204,1217,1913,1603,1604,926", rs1.Fields("cp10")) = 0) Then
'              strExc(1) = "": strExc(2) = "": strExc(3) = ""
'              strExc(9) = Replace(rs1("cno"), "-", "")
'              'Modified by Lydia 2021/11/05 模組修改
'              'strExc(2) = "": strExc(3) = "": strExc(4) = "": strExc(5) = ""
'              'If Not IsNull(rs1("pa27")) Then strExc(2) = ChangeCustomerL(rs1("pa27"))
'              'If Not IsNull(rs1("pa28")) Then strExc(3) = ChangeCustomerL(rs1("pa28"))
'              'If Not IsNull(rs1("pa29")) Then strExc(4) = ChangeCustomerL(rs1("pa29"))
'              'If Not IsNull(rs1("pa30")) Then strExc(5) = ChangeCustomerL(rs1("pa30"))
'              '核駁及審查意見通知函備註(複數備註)
'              'strExc(1) = PUB_GetIncomMemoNew(strExc(9), rs1("cp01"), "", ChangeCustomerL(rs1("pa75")), ChangeCustomerL(rs1("pa26")), "Y", , _
'                          , strExc(2), strExc(3), strExc(4), strExc(5))
'              'strExc(2) = "": strExc(3) = "": strExc(4) = "": strExc(5) = ""
'              strExc(1) = PUB_GetIncomMemoNew(strExc(9), rs1.Fields("cp01"), ChangeCustomerL("" & rs1.Fields("pa75")), ChangeCustomerL("" & rs1.Fields("pa26")) & "," & ChangeCustomerL("" & rs1.Fields("pa27")) & "," & ChangeCustomerL("" & rs1.Fields("pa28")) & "," & ChangeCustomerL("" & rs1.Fields("pa29")) & "," & ChangeCustomerL("" & rs1.Fields("pa30")), "Y", "", "" & rs1.Fields("cp10"))
'              'end 2021/11/05
'
'              If Len(strExc(1)) > 0 Then
'                   strMemo = strMemo & IIf(strMemo <> "", vbCrLf, "") & strExc(1)
'              End If
'          End If
       strSQL1 = PUB_FCPCFormMemo(strCP09, strMemo)
       strMemo = strSQL1
       'end 2022/01/21
      .Selection.SelectRow
      .Selection.Cells.SetHeight RowHeight:=iLHeight, HeightRule:=wdRowHeightAuto '自動調整列高(調整行高以適應行中的最高值)
      .Selection.ParagraphFormat.LineSpacingRule = wdLineSpaceExactly '固定行距
      .Selection.ParagraphFormat.LineSpacing = 14
      .Selection.TypeText Text:="備註：" & Replace(strMemo, vbCrLf, vbCrLf & "　　　")
      strPrintMemo = strMemo 'Added by Morgan 2020/9/15
      
      '對造若為本所客戶印出基本資料
      If PUB_ChkCustNameExist("" & rs1.Fields("CP40"), "" & rs1.Fields("CP41"), "" & rs1.Fields("CP42"), False, strCustNo) = True Then
         strMemo = "": strExc(2) = ""
         nRow = nRow + 2  '預設2行
         .Selection.MoveRight Unit:=wdCharacter, Count:=1
         If nRow > PageRow Then
            .Selection.InsertRows (nRow - PageRow)
            PageRow = PageRow + (nRow - PageRow)
            .Selection.Collapse Direction:=wdCollapseStart
         End If

         nRow = nRow - 1 '回到標題
         .Selection.TypeText Text:="對造為本所客戶："
         .Selection.MoveDown Unit:=wdLine, Count:=1
         For ii = LBound(varCustNo) To UBound(varCustNo)
            strExc(0) = "select cu01||cu02,cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90) cu05,cu06,cu10,NVL(cu11,'無') CU11 from customer where CU01||CU02='" & Left(varCustNo(ii), 9) & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               Select Case RsTemp.Fields("cu10")
                  Case "001", "002", "003", "004", "005", "006", "007", "008", "013", "020", "044"
                     If Not IsNull(RsTemp.Fields("cu04")) Then
                        strExc(2) = RsTemp.Fields("cu04")
                     ElseIf Not IsNull(RsTemp.Fields("cu05")) Then
                        strExc(2) = RsTemp.Fields("cu05")
                     ElseIf Not IsNull(RsTemp.Fields("cu06")) Then
                        strExc(2) = RsTemp.Fields("cu06")
                     End If
                  Case Else
                     If Not IsNull(RsTemp.Fields("cu05")) Then
                        strExc(2) = RsTemp.Fields("cu05")
                     ElseIf Not IsNull(RsTemp.Fields("cu04")) Then
                        strExc(2) = RsTemp.Fields("cu04")
                     ElseIf Not IsNull(RsTemp.Fields("cu06")) Then
                        strExc(2) = RsTemp.Fields("cu06")
                     End If
               End Select
               strMemo = strMemo & IIf(strMemo <> "", vbCrLf, "") & "　客戶編號：" & varCustNo(ii) & "　名稱：" & strExc(2) & "　ＩＤ：" & "" & RsTemp.Fields("cu11")
            End If
         Next ii
         nRow = nRow + 1
         .Selection.TypeText Text:=strMemo
      End If
      
      bolOpenWord = True
   End With
'-------------Word輸入完畢---------------------------
End If
   
On Error GoTo ErrorHand01
   '存成PDF檔
   If bolSavePdf Then
      strCOrderPath = App.path & "\" & strUserNum
      'Added by Morgan 2022/9/20
      If Dir(strCOrderPath, vbDirectory) = "" Then
         MkDir strCOrderPath
      End If
      'end 2022/9/20
      strExc(1) = ""
      'FCP檔案命名
      If "" & rs1.Fields("cp01") = "FCP" Then
          strPdfName = PUB_FCPCaseNo2FileName(rs1.Fields("cp01"), rs1.Fields("cp02"), rs1.Fields("cp03"), rs1.Fields("cp04")) & "." & rs1.Fields("cp10") & ".ORDER.PDF"
      Else
          strPdfName = PUB_CaseNo2FileName(rs1.Fields("cp01"), rs1.Fields("cp02"), rs1.Fields("cp03"), rs1.Fields("cp04")) & "." & rs1.Fields("cp10") & ".ORDER.PDF"
      End If
      'Modified by Lydia 2023/04/27 改模組
      'If PUB_PrintWord2PDF(g_WordAp, strCOrderPath, strPdfName, strExc(1)) = True Then
      If PUB_PrintWord2File(g_WordAp, strCOrderPath, strPdfName, strExc(1)) = True Then
           If Dir(strCOrderPath & "\" & strPdfName) <> "" Then
               'Modified by Lydia 2020/04/06 +FMP案(因應防疫在家上班作業，請將FMP案key來函產生的C類接洽記錄單回存到卷宗區)
               'If "" & rs1.Fields("cp01") = "FCP" Then
               If "" & rs1.Fields("cp01") = "FCP" Or bolFMP = True Then
                    'FCP案C類接洽單同時列印並且上傳到卷宗區(先列印,再上傳,最後刪除檔案)
                    Set oFile = oFileSys.GetFile(strCOrderPath & "\" & strPdfName)
                    If SaveAttFile_PDF(strCP09, strCOrderPath & "\" & strPdfName, strPdfName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , , True) = True Then
                        'PUB_PrintPDF strCOrderPath & "\" & strPdfName 'Removed by Morgan 2023/2/15 有存檔都不必再印紙本--品薇,敏莉
                    End If
                    RidFile strCOrderPath & "\" & strPdfName
               Else  '其他-只上傳
                    Set oFile = oFileSys.GetFile(strCOrderPath & "\" & strPdfName)
                    SaveAttFile_PDF strCP09, strCOrderPath & "\" & strPdfName, strPdfName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , , True
               End If
           End If
      End If
   Else '即時列印
      With g_WordAp
           .PrintOut Background:=False, Copies:=1, Collate:=True
           .ActiveDocument.Close wdDoNotSaveChanges
            If .Documents.Count = 0 Then
               .Quit wdDoNotSaveChanges
            End If
      End With
      Set g_WordAp = Nothing
   End If
   
If rs1.State <> adStateClosed Then rs1.Close
Set rs1 = Nothing
Set oFile = Nothing
Set oFileSys = Nothing

    Exit Sub
    
ErrorHand01:
   Select Case Err.Number
      Case 91, 462:
        '重新開啟
         GoTo ReOpenWord
      Case Else:
         MsgBox "錯誤 : " & Err.Description, vbCritical
         Exit Sub
   End Select

End Sub
