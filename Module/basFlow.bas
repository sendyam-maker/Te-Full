Attribute VB_Name = "basFlow"
Option Explicit

'Flow表單類別
Public Const Flow_結案單 = "1"
Public Const Flow_銷案銷帳單 = "2"
Public Const Flow_接洽單 = "3"
Public Const Flow_指示信 = "4" 'Added by Morgan 2015/11/13

Public Const Flow_不走分案的系統別 = "CFT,S,CFC,ACS"

'*** Flow表單狀態(下方有修改,需確認 GetF0309N 是否也要改) ***
Public Const Flow_主管審核中 = "01"
Public Const Flow_處理中 = "02"
Public Const Flow_已完成 = "03"
Public Const Flow_退回 = "04"
Public Const Flow_判發退回 = "05"
Public Const Flow_歸檔 = "06"
Public Const Flow_重送 = "07"
Public Const Flow_判發重送 = "08"
Public Const Flow_指示信判發中 = "09" 'Added by Morgan 2015/11/13
Public Const Flow_未上傳 = "10" 'Added by Morgan 2015/11/13
Public Const Flow_待寄送 = "11" 'Added by Morgan 2015/11/24
Public Const Flow_待發文 = "12" 'Added by Morgan 2020/1/16
Public Const Flow_待確收 = "13" 'Added by Morgan 2021/3/26
Public Const Flow_待分案 = "14" 'Added by Sindy 2022/8/9
Public Const Flow_智權補件 = "15" 'Added by Sindy 2022/8/9
Public Const Flow_補件完成 = "16" 'Added by Sindy 2022/8/9
Public Const Flow_已分案 = "17" 'Added by Sindy 2022/8/9
Public Const Flow_已收文 = "18" 'Added by Sindy 2022/9/26
Public Const Flow_放棄案源 = "19" 'Added by Sindy 2022/10/3
Public Const Flow_程序補件 = "20" 'Added by Sindy 2022/11/11
'*** End Flow表單狀態(上方有修改,需確認 GetF0309N 是否也要改) ***

'Public Const Flow_補看人員 = "73022" '73022.游登銘 'Removed by Morgan 2024/12/20 沒用了
Public Const Flow_可改簽核人員1 = "A2025,A3023" 'A2025.黃誠安,A3023.陳頌恩
'Modified by Morgan 2015/11/13 +4 指示信
'Modified by Morgan 2018/2/6 +5 帳單
'Modified by Lydia 2019/08/08 銷案銷帳單=>銷案／銷帳單
'Modified by Morgan 2020/1/16 +8 C類未發文
Public Const ShowFlow表單類別中文 = "'1','結案單','2','銷案／銷帳單','3','接洽單','4','指示信','5','帳單','6','客戶函','7','E化函','8','C類未發文'"
'Modified by Morgan 2015/11/13 +09,10,11
'Modified by Morgan 2020/1/16 +12
'Modified by Morgan 2021/3/26 +13
Public Const ShowFlow表單狀態中文 = "'01','主管審核中','02','處理中','03','已完成','04','退回','05','判發退回','06','歸檔','07','重送','08','判發重送','09','指示信判發中','10','未上傳','11','待寄送','12','待發文','13','待確收','14','待分案','15','智權補件','16','補件完成','17','已分案','18','已收文','19','放棄案源','20','程序補件'"
'Add By Sindy 2022/9/22
Public Const ShowFlow特殊簽核人員 = "'A3','櫃檯人員','A4','法務案源','A5','分所專利','A6','北所分案','A7','程序人員'"
'2022/9/22 END
Public Const ShowFlow簽核人員身份 = "'1','簽核主管','2','程序人員','3','審核人員','A0','智權人員','A1','簽核人員','A2','特例主管','A3','櫃檯人員','A4','法務案源','A5','分所專利','A6','北所分案','A7','程序人員'" 'Modify by Amy 2018/08/29 3.判發改審核
Public Const ShowFlow簽核結果 = "'1','同意','2','退回','3','已處理','4','放棄','5','已補件','6','智權補件','7','程序補件'"
'Add by Amy 2018/08/27 結案單寄信內容(操作路徑)
Public Const 結案單外商CF操作路徑 = "請至商標系統的外商->CF資料處理->待處理區，進行結案操作。"
Public Const 結案單補看人員操作路徑 = "請至一般作業->案件表單查詢及簽核->專業部 審核/補看 作業審核。" 'Modify by Amy 2018/08/29 +專業部
'Add by Amy 2025/05/26 FC結案單操作路徑
Public Const 結案單外商FC操作路徑 = "請至外商系統的外商->FC資料處理->待處理區，進行結案操作。"
Public Const 結案單外專FC操作路徑 = "請至外專系統的外專->資料處理->待處理區，進行結案操作。"
Global TF_CRL As Integer
Public m_blnABSActivated As Boolean 'Add By Sindy 2011/9/15
Public m_bolLoanAct As Boolean 'Add by Amy 2017/02/03 判斷是否有圖書借閱記錄需簽核


'Add by Amy 2022/11/15 傳入F0309編號,回傳編號+名稱
Public Function GetF0309N(ByVal stF0309 As String, Optional ByVal bolShowNo As Boolean = True) As String

    GetF0309N = ""
    Select Case stF0309
        Case "01"
            GetF0309N = "主管審核中"
        Case "02"
            GetF0309N = "處理中"
        Case "03"
            GetF0309N = "已完成"
        Case "04"
            GetF0309N = "退回"
        Case "05"
            GetF0309N = "判發退回"
        Case "06"
            GetF0309N = "歸檔"
        Case "07"
            GetF0309N = "重送"
        Case "08"
            GetF0309N = "判發重送"
        Case "09"
            GetF0309N = "指示信判發中"
        Case "10"
            GetF0309N = "未上傳"
        Case "11"
            GetF0309N = "待寄送"
        Case "12"
            GetF0309N = "待發文"
        Case "13"
            GetF0309N = "待確收"
        Case "14"
            GetF0309N = "待分案"
        Case "15"
            GetF0309N = "智權補件"
        Case "16"
            GetF0309N = "補件完成"
        Case "17"
            GetF0309N = "已分案"
        Case "18"
            GetF0309N = "已收文"
        Case "19"
            GetF0309N = "放棄案源"
        Case "20"
            GetF0309N = "程序補件"
    End Select
    If bolShowNo = True Then
        GetF0309N = stF0309 & " " & GetF0309N
    End If
End Function

'案件電子表單簽核-表單編號(電腦自動給號)
Public Function AutoNo_FLOW(InputItem As String, InputLength As Integer) As String
Dim adoaccnum As New ADODB.Recordset
Dim strItem As String, strYes As String
   
   adoTaie.Execute "update autonumber set au03 = au03 where au01 = '" & InputItem & "'"
   If Trim(InputItem) = "CLS" Then
      strItem = ""
   End If
   adoaccnum.CursorLocation = adUseClient
   adoaccnum.Open "select * from autonumber where au01 = '" & InputItem & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoaccnum.RecordCount = 0 Then
      AutoNo_FLOW = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo("0", InputLength)
   Else
      If adoaccnum.Fields("au02").Value <> Val(Mid(strSrvDate(1), 1, 4)) Then
         AutoNo_FLOW = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo("0", InputLength)
      Else
         AutoNo_FLOW = strItem & Mid(strSrvDate(2), 1, 3) & ZeroBeforeNo(str(adoaccnum.Fields("au03").Value), InputLength)
      End If
   End If
   strYes = SaveAutoNo(InputItem, Mid(AutoNo_FLOW, 4, InputLength))
   adoaccnum.Close
   
   Set adoaccnum = Nothing
End Function

'檢查案件結案單是否已存在
'Modify by Amy 2025/05/23 +strCCM17 信件編號
Public Function ChkFlowFormExists(ByVal strFormKind As String, ByVal strNP01 As String, ByVal strNP22 As String, _
                                  ByVal strCP01 As String, ByVal strCP02 As String, ByVal strCP03 As String, ByVal strCP04 As String, _
                                  Optional ByVal strGetColName As String, Optional ByRef strGetColVal As String, _
                                  Optional ByVal strCurrF0301 As String, Optional ByVal strCCM17 As String = "") As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String, strCon As String
Dim i As Integer
Dim strTB As String 'Add by Amy 2025/05/23
   
   ChkFlowFormExists = False
   
   If Val(strNP22) > 0 Then
      strSql = "SELECT NP06,NP24 FROM NextProgress" & _
               " WHERE NP01='" & strNP01 & "' and NP22=" & strNP22 & IIf(strCurrF0301 <> "", " and NP24<>'" & strCurrF0301 & "'", "") & _
               " and NP24 is not null"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
'         ChkFlowFormExists = True
'         If UCase(strGetColName) = UCase("F0301") Then
'            If Len("" & rsTmp.Fields("NP24")) = 8 Then '8碼為電子表單編號
'               strGetColVal = "" & rsTmp.Fields("NP24")
'            End If
'         End If
         If Len("" & rsTmp.Fields("NP24")) = 8 Then '8碼為電子表單編號
            ChkFlowFormExists = True
            If UCase(strGetColName) = UCase("F0301") Then
               strGetColVal = "" & rsTmp.Fields("NP24")
            End If
            rsTmp.Close
            Set rsTmp = Nothing
            Exit Function
         End If
'         rsTmp.Close
'         Set rsTmp = Nothing
'         Exit Function
      End If
      rsTmp.Close
   End If
   
   'Modify By Sindy 2016/1/6 Ex:P-94839 104/11月先做了案件的無期限閉卷,又再105/1要做領證及繳年費的閉卷
'   For i = 1 To 3
'      strSql = "SELECT * FROM FLOW003" & _
'               " WHERE F0302='" & strFormKind & "'"
'      If i = 1 Then
'         If strNP01 <> "" And strNP22 <> "" Then
'            strCon = " and F0303='" & strNP01 & "' and F0304=" & strNP22
'         Else
'            GoTo goStep
'         End If
'      ElseIf i = 2 Then
'         If strNP01 <> "" Then
'            strCon = " and F0303='" & strNP01 & "'"
'         Else
'            GoTo goStep
'         End If
'      Else
'         If strCP01 <> "" Then
'            strCon = " and F0303='" & strCP01 & strCP02 & strCP03 & strCP04 & "'"
'         Else
'            GoTo goStep
'         End If
'      End If
'      If strCurrF0301 <> "" Then
'         strCon = strCon & " and F0301<>'" & strCurrF0301 & "'"
'      End If
'      strSql = strSql & strCon
'      rsTmp.CursorLocation = adUseClient
'      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'      If rsTmp.RecordCount > 0 Then
'         ChkFlowFormExists = True
'
'         If strGetColName <> "" Then
'            strGetColVal = rsTmp.Fields(strGetColName)
'         End If
'
'         rsTmp.Close
'         Set rsTmp = Nothing
'         Exit Function
'      End If
'      rsTmp.Close
'goStep:
'   Next i
   strCon = ""
   'Modify by Amy 2025/04/10 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
   If strSrvDate(1) >= FCP結案單電子化啟用日 Then
      strTB = ",CloseCaseMain"
      If strCCM17 <> "" Then
         strCon = strCon & " And F0301(+)=CCM01 And CCM17='" & strCCM17 & "'"
         If strNP22 <> "" Then
            strTB = strTB & ",NextProgress"
            strCon = strCon & " And CCM02=NP01(+) And CCM03=NP22(+) And length(NP24)=8"
         End If
      ElseIf strNP01 <> "" And strNP22 <> "" Then
         strCon = " and CCM02='" & strNP01 & "' and CCM03=" & strNP22
      ElseIf strNP01 <> "" Then
         strCon = " and CCM02='" & strNP01 & "'"
      ElseIf strCP01 <> "" Then
         strCon = " and CCM02='" & strCP01 & strCP02 & strCP03 & strCP04 & "'"
      End If
      If strCCM17 = "" Then strCon = strCon & " And F0301=CCM01(+) "
   Else
      If strNP01 <> "" And strNP22 <> "" Then
         strCon = " and F0303='" & strNP01 & "' and F0304=" & strNP22
      ElseIf strNP01 <> "" Then
         strCon = " and F0303='" & strNP01 & "'"
      ElseIf strCP01 <> "" Then
         strCon = " and F0303='" & strCP01 & strCP02 & strCP03 & strCP04 & "'"
      End If
   End If
   If strCurrF0301 <> "" Then
      strCon = strCon & " and F0301<>'" & strCurrF0301 & "'"
   End If
   'Modify By Sindy 2015/1/6 + and f0309<>'06'
   strSql = "SELECT * FROM FLOW003" & strTB & _
            " WHERE F0302='" & strFormKind & "' and f0309<>'06'"
   strSql = strSql & strCon
   If strCon <> "" Then
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         If strCCM17 <> "" Then
            '系統收件區多筆案號結案,於最後一筆才更新信件沖銷,若中途未正常結束,可能第一筆結案單已產生,故以信件編號查詢
            strExc(5) = "" & rsTmp.Fields("CCM02") & rsTmp.Fields("CCM03")
            If strExc(5) = strNP01 & strNP22 Then
               '有期限
               ChkFlowFormExists = True
            ElseIf strExc(5) = strCP01 & strCP02 & strCP03 & strCP04 Then
               '無期限
               ChkFlowFormExists = True
            End If
         Else
            ChkFlowFormExists = True
         End If
         If strGetColName <> "" Then
            strGetColVal = rsTmp.Fields(strGetColName)
         End If
         
         rsTmp.Close
         Set rsTmp = Nothing
         Exit Function
      End If
      rsTmp.Close
   End If
      
   Set rsTmp = Nothing
End Function

'檢查該表單編號是屬於那一類案件資料
'NP:下一程序
'CP:案件進度
'MF:主檔
Public Function GetFlowFormReadDBType(ByVal strF0301 As String, Optional ByRef strFormKind As String _
                                      , Optional ByRef strCP01 As String, Optional ByRef strCP02 As String _
                                      , Optional ByRef strCP03 As String, Optional ByRef strCP04 As String) As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
   
   GetFlowFormReadDBType = ""
   
   'Add by Amy 2025/05/08 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
   If strSrvDate(1) >= FCP結案單電子化啟用日 Then
      strSql = "SELECT * FROM FLOW003,CloseCaseMain WHERE F0301='" & strF0301 & "' And F0301=CCM01(+) "
   Else
      strSql = "SELECT * FROM FLOW003 WHERE F0301='" & strF0301 & "'"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      strFormKind = rsTmp.Fields("F0302")
      'Modify by Amy 2025/05/08 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
      If strSrvDate(1) >= FCP結案單電子化啟用日 Then
         If Len(rsTmp.Fields("ccm02")) = 9 And "" & rsTmp.Fields("ccm04") = "" Then
            GetFlowFormReadDBType = "CP"
         ElseIf Len(rsTmp.Fields("ccm02")) = 9 And "" & rsTmp.Fields("ccm04") <> "" Then
            GetFlowFormReadDBType = "NP"
         Else
            GetFlowFormReadDBType = "MF"
            strCP01 = Left(rsTmp.Fields("ccm02"), Len(rsTmp.Fields("ccm02")) - 9)
            strCP02 = Mid(rsTmp.Fields("ccm02"), Len(strCP01) + 1, 6)
            strCP03 = Mid(rsTmp.Fields("ccm02"), Len(strCP01) + 7, 1)
            strCP04 = Right(rsTmp.Fields("ccm02"), 2)
         End If
      Else
         If Len(rsTmp.Fields("F0303")) = 9 And "" & rsTmp.Fields("F0304") = "" Then
            GetFlowFormReadDBType = "CP"
         ElseIf Len(rsTmp.Fields("F0303")) = 9 And "" & rsTmp.Fields("F0304") <> "" Then
            GetFlowFormReadDBType = "NP"
         Else
            GetFlowFormReadDBType = "MF"
            strCP01 = Left(rsTmp.Fields("F0303"), Len(rsTmp.Fields("F0303")) - 9)
            strCP02 = Mid(rsTmp.Fields("F0303"), Len(strCP01) + 1, 6)
            strCP03 = Mid(rsTmp.Fields("F0303"), Len(strCP01) + 7, 1)
            strCP04 = Right(rsTmp.Fields("F0303"), 2)
         End If
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Function

'取得案件表單簽核人員
'Add By Sindy 2020/7/20 + , Optional intGetEmpNum As Integer = 0 : 取幾位簽核人員; 0:全部
'Modify By Sindy 2022/9/19 , Optional ByRef strRecvSet1 As String = "" : 接洽單一般簽核主管
'                          , Optional ByRef strRecvSet2 As String = "" : 接洽單特例簽核主管
Public Function GetFLOW001Person(ByVal StrST01 As String, ByVal strFormKind As String, _
   Optional bolInSelf As Boolean = False, Optional intGetEmpNum As Integer = 0, _
   Optional ByRef strRecvSet1 As String = "", Optional ByRef strRecvSet2 As String = "") As String
   
Dim rsTmp As New ADODB.Recordset
Dim intCnt As Integer 'Add By Sindy 2020/7/20
   
   'Modify By Sindy 2018/8/16 Mark
'   If bolInSelf = False Then
'      '讀取他人簽核人員檔時,均不可過濾掉自己
'      If strST01 <> strUserNum Then
'         bolInSelf = True
'      End If
'   End If
   
   intCnt = 0 'Add By Sindy 2020/7/20
   GetFLOW001Person = ""
   strRecvSet1 = "": strRecvSet2 = "" 'Add By Sindy 2022/9/19
   strSql = "select * from FLOW001 where F0101=" & CNULL(StrST01) & " and F0102=" & CNULL(strFormKind)
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Not IsNull(rsTmp.Fields("F0103")) Then
         If ChkStaffST04(rsTmp.Fields("F0103"), False) = False Then
            'Modify By Sindy 2018/8/16 ex:P-107546
            'If rsTmp.Fields("F0103") <> strUserNum Or bolInSelf = True Then
            If rsTmp.Fields("F0103") <> StrST01 Or bolInSelf = True Then
               GetFLOW001Person = GetFLOW001Person & rsTmp.Fields("F0103") & ","
               'Add By Sindy 2020/7/20
               intCnt = intCnt + 1
               If intGetEmpNum > 0 And intCnt >= intGetEmpNum Then GoTo ExitEnd
               '2020/7/20 END
            End If
         End If
      End If
      If Not IsNull(rsTmp.Fields("F0104")) Then
         If ChkStaffST04(rsTmp.Fields("F0104"), False) = False Then
            'If rsTmp.Fields("F0104") <> strUserNum Or bolInSelf = True Then
            If rsTmp.Fields("F0104") <> StrST01 Or bolInSelf = True Then
               GetFLOW001Person = GetFLOW001Person & rsTmp.Fields("F0104") & ","
               'Add By Sindy 2020/7/20
               intCnt = intCnt + 1
               If intGetEmpNum > 0 And intCnt >= intGetEmpNum Then GoTo ExitEnd
               '2020/7/20 END
            End If
         End If
      End If
      If Not IsNull(rsTmp.Fields("F0105")) Then
         If ChkStaffST04(rsTmp.Fields("F0105"), False) = False Then
            'If rsTmp.Fields("F0105") <> strUserNum Or bolInSelf = True Then
            If rsTmp.Fields("F0105") <> StrST01 Or bolInSelf = True Then
               GetFLOW001Person = GetFLOW001Person & rsTmp.Fields("F0105") & ","
               'Add By Sindy 2020/7/20
               intCnt = intCnt + 1
               If intGetEmpNum > 0 And intCnt >= intGetEmpNum Then GoTo ExitEnd
               '2020/7/20 END
            End If
         End If
      End If
      If strFormKind = Flow_接洽單 Then strRecvSet1 = GetFLOW001Person 'Add By Sindy 2022/9/19
      
      'Modify By Sindy 2022/8/26 改簽核人員4,5,6為接洽單特例簽核主管
      'Add By Sindy 2020/7/20 非接洽單才抓取簽核人員5,6;因接洽單簽核人員5,6為特殊情形
      If Not IsNull(rsTmp.Fields("F0106")) Then
         If ChkStaffST04(rsTmp.Fields("F0106"), False) = False Then
            'If rsTmp.Fields("F0106") <> strUserNum Or bolInSelf = True Then
            If rsTmp.Fields("F0106") <> StrST01 Or bolInSelf = True Then
               If rsTmp.Fields("F0102") <> "3" Then 'Add By Sindy 2020/7/20 + if
                  GetFLOW001Person = GetFLOW001Person & rsTmp.Fields("F0106") & ","
                  'Add By Sindy 2020/7/20
                  intCnt = intCnt + 1
                  If intGetEmpNum > 0 And intCnt >= intGetEmpNum Then GoTo ExitEnd
                  '2020/7/20 END
               'Modify By Sindy 2022/9/19
               Else
                  strRecvSet2 = strRecvSet2 & rsTmp.Fields("F0106") & ","
               End If
            End If
         End If
      End If
      If Not IsNull(rsTmp.Fields("F0107")) Then
         If ChkStaffST04(rsTmp.Fields("F0107"), False) = False Then
            'If rsTmp.Fields("F0107") <> strUserNum Or bolInSelf = True Then
            If rsTmp.Fields("F0107") <> StrST01 Or bolInSelf = True Then
               If rsTmp.Fields("F0102") <> "3" Then 'Add By Sindy 2020/7/20 + if
                  GetFLOW001Person = GetFLOW001Person & rsTmp.Fields("F0107") & ","
                  'Add By Sindy 2020/7/20
                  intCnt = intCnt + 1
                  If intGetEmpNum > 0 And intCnt >= intGetEmpNum Then GoTo ExitEnd
                  '2020/7/20 END
               'Modify By Sindy 2022/9/19
               Else
                  strRecvSet2 = strRecvSet2 & rsTmp.Fields("F0107") & ","
               End If
            End If
         End If
      End If
      If Not IsNull(rsTmp.Fields("F0108")) Then
         If ChkStaffST04(rsTmp.Fields("F0108"), False) = False Then
            'If rsTmp.Fields("F0108") <> strUserNum Or bolInSelf = True Then
            If rsTmp.Fields("F0108") <> StrST01 Or bolInSelf = True Then
               If rsTmp.Fields("F0102") <> "3" Then 'Add By Sindy 2020/7/20 + if
                  GetFLOW001Person = GetFLOW001Person & rsTmp.Fields("F0108") & ","
                  'Add By Sindy 2020/7/20
                  intCnt = intCnt + 1
                  If intGetEmpNum > 0 And intCnt >= intGetEmpNum Then GoTo ExitEnd
                  '2020/7/20 END
               'Modify By Sindy 2022/9/19
               Else
                  strRecvSet2 = strRecvSet2 & rsTmp.Fields("F0108") & ","
               End If
            End If
         End If
      End If
   End If
   
ExitEnd:
   If GetFLOW001Person <> "" Then GetFLOW001Person = Left(GetFLOW001Person, Len(GetFLOW001Person) - 1)
   If strRecvSet1 <> "" Then strRecvSet1 = Left(strRecvSet1, Len(strRecvSet1) - 1) 'Add By Sindy 2022/9/19
   If strRecvSet2 <> "" Then strRecvSet2 = Left(strRecvSet2, Len(strRecvSet2) - 1) 'Add By Sindy 2022/9/19
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'取得Insert案件表單流程備註Sql
'F0403=處理人員
'F0408=簽核身份
'F0409=呈報對象身份
'Modify By Sindy 2022/10/10 + , Optional ByVal strFromF0408 As String, Optional ByVal strToF0409 As String
Public Function GetInsertFLOW004Sql(strF0401 As String, strF0403 As String, strF0404 As String, strF0405 As String, _
   strF0406 As String, strF0407 As String, Optional ByVal strFromF0408 As String, Optional ByVal strToF0409 As String) As String
Dim intSeqno As Integer
Dim rsTmp As New ADODB.Recordset
   
   strSql = "select * from FLOW004 where F0401='" & strF0401 & "' "
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strSql = "select max(F0402) from FLOW004 where F0401='" & strF0401 & "' "
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If Not IsNull(rsTmp.Fields(0).Value) Then intSeqno = rsTmp.Fields(0).Value
      End If
   End If
   'Modify By Sindy 2022/10/10 + ,F0408,F0409
   GetInsertFLOW004Sql = "insert into FLOW004 (F0401,F0402,F0403,F0404,F0405,F0406,F0407,F0408,F0409) " & _
                         "values(" & CNULL(strF0401) & "," & (intSeqno + 1) & "," & _
                         CNULL(strF0403) & "," & strF0404 & "," & strF0405 & "," & _
                         CNULL(strF0406) & "," & CNULL(strF0407) & "," & CNULL(strFromF0408) & "," & CNULL(strToF0409) & ")"
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'取得流程中相關人員(Memo by Amy 2018/06/01 目前結案單用)
'Modify by Amy 2020/03/06 +strCaseNo for CFP業務區劃分
'Modify by Amy 2025/04/16 +strCP10 案件性質,strF0316 結案單業務,IsFMPY53374 FMP寰華案
Public Function GetSignOffEmp(strEmp As String, strCP01 As String, strCP02 As String, strNation As String, Optional ByVal strCaseNo As String = "", _
  Optional ByVal strCP10 As String = "", Optional ByVal strF0316 As String = "", Optional ByVal IsFMPY53374 As Boolean = False) As String
Dim rsA As New ADODB.Recordset
Dim m_CaseNo(1 To 4) As String, ii As Integer, intQ As Integer, arrTxt 'Add  by Amy 2025/04/16
   
   GetSignOffEmp = ""
   'Add by Amy 2025/04/16
   If InStr(strCaseNo, "-") > 0 Then
      arrTxt = Split(strCaseNo, "-")
      intQ = 1
      For ii = LBound(arrTxt) To UBound(arrTxt)
         m_CaseNo(intQ) = arrTxt(ii)
         intQ = intQ + 1
      Next
      If m_CaseNo(3) = "" Then m_CaseNo(3) = "0"
      If m_CaseNo(4) = "" Then m_CaseNo(4) = "00"
   Else
      ChgCaseNo strCaseNo, m_CaseNo
   End If
   'end 2025/04/16
   Select Case UCase(strEmp)
      'Add by Amy 2025/05/08
      Case "CM1" 'FC[結案單]承辦人員主管(ST52)
         GetSignOffEmp = PUB_GetFCPProSup(strF0316) 'FC承辦人員2級主管
      Case "CM2" 'FC[結案單]承辦人員主管(ST53)
         GetSignOffEmp = GetST52SelfList(strF0316, "st53") 'FC承辦人員3級主管
      'end 2025/05/08
      'Add by Amy 2025/09/30
      Case "CM3"
         GetSignOffEmp = GetST52SelfList(strF0316, "st54,st55") 'FC承辦人員4-5級主管
         'Memo by Amy 目前外專之st55=83002
         GetSignOffEmp = Replace("," & GetSignOffEmp, "," & Pub_GetSpecMan("程式管理人員"), "")
         'Add by Amy 2025/10/15
         If GetSignOffEmp <> "" Then
            If Left(GetSignOffEmp, 1) = "," Then GetSignOffEmp = Mid(GetSignOffEmp, 2)
         End If
      'end 2025/09/30
      Case "NP" '程序人員
         'Add by Amy 2025/04/16 +FCP/FG-[外專]程序
         If strCP01 = "FCP" Or strCP01 = "FG" Then
             GetSignOffEmp = PUB_GetFCPHandler(m_CaseNo(1), m_CaseNo(2), m_CaseNo(3), m_CaseNo(4))
         'Modify by Amy 2018/06/22 +CPS走CFP流程
         ElseIf strCP01 = "CFP" Or strCP01 = "CPS" Then
            'Modify by Amy 2020/03/06 +if
            If strSrvDate(1) >= CFP業務區劃分啟用日 Then
                 GetSignOffEmp = PUB_GetCFPHandler(strCaseNo)
            Else
                strExc(0) = "select na73,na74 from nation " & _
                            "where na01='" & strNation & "' "
                intI = 1
                Set rsA = ClsLawReadRstMsg(intI, strExc(0))
                If intI = 1 Then
                   If (strCP02 Mod 2) = 0 Then '雙號
                      GetSignOffEmp = "" & rsA.Fields("na74")
                   Else
                      GetSignOffEmp = "" & rsA.Fields("na73")
                   End If
                End If
            End If
            'end 2020/03/06
          'Modify by Amy 2018/06/20 P/PS走P流程;CFT/S 走CFT流程;T字頭走T流程
         ElseIf strCP01 = "P" Or strCP01 = "PS" Then
            'Added by Morgan 2025/1/14
            If strSrvDate(1) >= P業務區劃分啟用日 Then
                 'Modify by Amy 2025/04/16 +if
                 If strSrvDate(1) >= FCP結案單電子化啟用日 And IsFMPY53374 = True Then
                     'P寰華-[外專]程序結
                     GetSignOffEmp = PUB_GetFCPHandler(m_CaseNo(1), m_CaseNo(2), m_CaseNo(3), m_CaseNo(4))
                 Else
                     'Memo by Amy 2025/04/16  國內P/FC結案的P[非]寰華FMP-[內專]程序結
                     GetSignOffEmp = PUB_GetPHandler(strCaseNo)
                 End If
            Else
            'end 2025/1/14
            
               If strNation = "000" Then '台灣案
                  GetSignOffEmp = Pub_GetSpecMan("PS1")
               Else '非台灣案
                  GetSignOffEmp = Pub_GetSpecMan("PS2")
               End If
               
            End If 'Added by Morgan 2025/1/14
            
         'Modify by Amy 2019/11/27 CFC原GetNA69
         'Modify by Amy 2019/12/05 美國案才抓特殊設定,非美國仍抓GetNA69
         ElseIf strCP01 = "CFC" And strNation = "101" Then
            'Memo by Amy 2021/06/28 此有改需確認 GetCFTSt16Manager 是否也需改
            GetSignOffEmp = Pub_GetSpecMan("CFC案承辦人")
         'Modify by Amy 2018/06/22 +CFC走CFT流程
         'Modify by Amy 2025/06/12 +strNation<>"000"
         ElseIf strCP01 = "CFT" Or (strCP01 = "S" And strNation <> "000") Or strCP01 = "CFC" Then
            'Memo by Amy 2021/06/28 此有改需確認 GetCFTSt16Manager 是否也需改
            'strCP13傳操作人之所別
            Call GetNA69("", strNation, strUserNum, GetSignOffEmp)
         'end 2019/12/05
         'T字頭(ex:TD...)
         ElseIf Left(strCP01, 1) = "T" Then
            GetSignOffEmp = Pub_GetSpecMan("TS1")
         'end 2018/06/20
         'Add by Amy 2025/6/12 FC結案單
         ElseIf strCP01 = "FCT" Or (strCP01 = "S" And strNation = "000") Then
            'Modify by Amy 2025/08/15 FCT爭議案(原:內商結抓系統特殊設定-TS1)全部回到外商結案-114/7/2 秀玲詢問後結果
            'FCT or S 台灣案-[外商]FC案件程序管制人(ST57)
            strExc(9) = "Select st57 From Staff Where st01='" & strF0316 & "' "
            intI = 1
            Set rsA = ClsLawReadRstMsg(intI, strExc(9))
            If intI = 1 Then
               GetSignOffEmp = "" & rsA.Fields("st57")
            End If
         End If
         'end 2025/08/15
   End Select
   'Modify by Amy 2025/10/15 +UCase(strEmp)
   If GetSignOffEmp <> "" And UCase(strEmp) <> "CM3" Then
      GetSignOffEmp = GetSignOffEmp & " " & GetPrjSalesNM(GetSignOffEmp)
   End If
   
   Set rsA = Nothing
End Function

'設定專利處程序組人員下拉選單
Public Sub SetPatentP12Combo(objCbo As Object, strSysID As String, objLbl As Object)
Dim rsTmp As New ADODB.Recordset
Dim i As Integer
   
   objCbo.Clear
   If strSysID = "CFP" Then
      objLbl.Caption = "程序人員："
      strExc(0) = "select na73 from nation " & _
                  "Union " & _
                  "select na74 from nation " & _
                  "order by 1 asc "
      'Add by Amy 2020/03/06
      If strSrvDate(1) >= CFP業務區劃分啟用日 Then
         strExc(0) = "Select Distinct a0916 From Acc090 Order by 1"
      End If
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         rsTmp.MoveFirst
         objCbo.AddItem ""
         Do While Not rsTmp.EOF
            objCbo.AddItem rsTmp.Fields(0) & " " & GetPrjSalesNM(rsTmp.Fields(0))
            rsTmp.MoveNext
         Loop
      End If
      objCbo.ListIndex = 1
      For i = 0 To objCbo.ListCount - 1
         If Trim(Left(objCbo.List(i), 6)) = strUserNum Then
            objCbo.ListIndex = i
            Exit For
         End If
      Next i
      rsTmp.Close
      Set rsTmp = Nothing
   ElseIf strSysID = "P" Then
   
      'Added by Morgan 2025/1/10
      If strSrvDate(1) >= P業務區劃分啟用日 Then
         strExc(0) = "select a0917,st02 from (Select Distinct a0917 From Acc090),staff" & _
            " where a0917 is not null and  st01(+)=a0917 Order by 1"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            rsTmp.MoveFirst
            objCbo.AddItem ""
            Do While Not rsTmp.EOF
               objCbo.AddItem rsTmp.Fields(0) & " " & rsTmp.Fields(1)
               rsTmp.MoveNext
            Loop
         End If
         objCbo.ListIndex = 0
         For i = 0 To objCbo.ListCount - 1
            If Trim(Left(objCbo.List(i), 6)) = strUserNum Then
               objCbo.ListIndex = i
               Exit For
            End If
         Next i
         rsTmp.Close
         Set rsTmp = Nothing
      Else
      'end 2025/1/10
      
         objLbl.Caption = "國家："
         objCbo.AddItem ""
         objCbo.AddItem "1 台灣案"
         objCbo.AddItem "2 非台灣案"
         objCbo.ListIndex = 1
         If Pub_GetSpecMan("PS2") = strUserNum Then '非台灣案
            objCbo.ListIndex = 2
         End If
         
      End If
   End If
End Sub
   
'設定案件表單類別
'Mofieied by Morgan 2015/11/12 是否加指示信選項
'Modified by Morgan 2020/1/16 +pSys
'Modify by Amy 2023/02/13 +stFormN as string
Public Sub Flow_SetF0302Combo(objCbo As Object, Optional bolMore As Boolean = False, Optional pSys As String, Optional ByVal stFormN = "")
   Dim stNotShow As String 'Add by Amy 2023/02/13
   
   '*** Memo 若有增加表單請於下方列示,方便知道修改影響之程式 ***
   Select Case UCase(stFormN)
        Case "FRM040118" '結案單審核作業
            stNotShow = "3"
        Case "FRM210147" '案件目前表單
        Case "FRM210148" '案件簽核作業
        Case "FRM210149" '待處理區
            stNotShow = "3"
   End Select
   
   objCbo.Clear
   'Modify by Amy 2023/02/13 +stNotShow
   If InStr(stNotShow, "0") = 0 Then objCbo.AddItem "0  全部"
   If InStr(stNotShow, "1") = 0 Then objCbo.AddItem "1  結案單"
   'Modified by Lydia 2019/08/08 銷案銷帳單=>銷案／銷帳單
   If InStr(stNotShow, "2") = 0 Then objCbo.AddItem "2  銷案／銷帳單"
   If InStr(stNotShow, "3") = 0 Then objCbo.AddItem "3  接洽單"
   'end 2023/02/13
   If bolMore Then
      'Modify By Sindy 2020/12/28
      If InStr(pSys, "T") = 0 Then
      '2020/12/28 END
         objCbo.AddItem "4  指示信"
      End If
      objCbo.AddItem "5  帳單" 'Added by Morgan 2018/2/6
      'Modify By Sindy 2020/12/28
      If InStr(pSys, "T") = 0 Then
      '2020/12/28 END
         objCbo.AddItem "6  客戶函" 'Added by Morgan 2019/1/17
      End If
      'Added by Morgan 2018/10/25
      'Modified by Morgan 2021/3/29
      'If strSrvDate(1) >= e化客戶啟用日 Then
      'Modified by Morgan 2022/2/14
      'If InStr(pSys, "T") = 0 Then
         objCbo.AddItem "7  E化函"
      'End If
      'end 2018/10/25
      
      'Added by Morgan 2020/1/16
      If pSys = "P" Then
         objCbo.AddItem "8  C類未發文"
      End If
      'end 2020/1/16
   End If
   objCbo.ListIndex = 0
End Sub

'設定操作人員下拉選單及含須代理職務的人員
'Modify By Sindy 2023/1/5 + Optional ByVal bolAddAll As Boolean = False, Optional ByRef strAllEmp As String = ""
'Modify by Amy 2023/02/10 +stFormN
Public Sub SetEmpDutyCombo(objCbo As Object, Optional ByVal bolInTurnover As Boolean = False, _
   Optional ByVal bolAddAll As Boolean = False, Optional stFormN As String = "")
Dim strUser As String, arrData As Variant
Dim i As Integer
   
   '*** Memo 若有增加表單請於下方列示,方便知道修改影響之程式 ***
   Select Case UCase(stFormN)
        Case "FRM210147" '案件目前表單
        Case "FRM210148" '案件簽核作業
   End Select
   
   objCbo.Clear
   'Modify By Sindy 2025/8/5 +And Left(PUB_GetST03(strUserNum), 1) <> "F"
   If bolAddAll = True And Left(PUB_GetST03(strUserNum), 1) <> "F" Then objCbo.AddItem "All   全部" 'Add By Sindy 2023/1/5
   
   'Modify By Sindy 2022/8/22
   'Modify by Amy 2025/07/09 +Wirter ,亮丞/廷瑋 會於櫃台協助收文
   If PUB_GetST03(strUserNum) = "M12" Or PUB_GetST03(strUserNum) = "M13" _
     Or UCase(App.EXEName) = "TEWRITER" Or UCase(App.EXEName) = "WRITER" Then '接待室、打字室
      objCbo.AddItem "A3   櫃檯人員"
   Else
   '2022/8/22 END
      objCbo.AddItem strUserNum & " " & strUserName
   End If
   
   'Add By Sindy 2023/1/3
   If strUserNum = "71011" Then
      objCbo.AddItem "79075 郭雅娟"
   End If
   '2023/1/3 END
   
   '檢查當時是否需要為他人職代
   'Modify By Sindy 2016/10/12
   Call Pub_SetForOthersEmpCombo(strUserNum, objCbo, False, , True) '含特殊職代
   
   'Add By Sindy 2025/8/5
   '國外部承辦組主管的抓法
   If Left(PUB_GetST03(strUserNum), 1) = "F" Then
      '案件簽核作業
      If UCase(stFormN) = "FRM210148" Then
         '本人為主管
         strSql = "select st01,st02 from staff where st03='" & Pub_StrUserSt03 & "' and st04='1' and instr(';'||st52||';'||st53||';'||st54||';'||st55||';',';" & strUserNum & ";')>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            'Add By Sindy 2025/10/27
            If Pub_StrUserSt03 = "F11" Then '外商
               '檢查當時是否需要為他人職代
               Call Pub_SetForOthersEmpCombo(strUserNum, objCbo, False, , True, False)
            Else
            '2025/10/27 END
               strSql = "select distinct st52 from (select st52 from staff where st93='" & PUB_GetST93(strUserNum) & "' and st04='1' and st52 is not null and st52<>'" & strUserNum & "'" & _
                        " union select st53 from staff where st93='" & PUB_GetST93(strUserNum) & "' and st04='1' and st53 is not null and st53<>'" & strUserNum & "'" & _
                        " union select st54 from staff where st93='" & PUB_GetST93(strUserNum) & "' and st04='1' and st54 is not null and st54<>'" & strUserNum & "'" & _
                        ")"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  RsTemp.MoveFirst
                  Do While Not RsTemp.EOF
                     For i = 0 To objCbo.ListCount - 1
                        If InStr(objCbo.List(i), RsTemp(0)) = 1 Then
                           Exit For
                        End If
                     Next i
                     If i = objCbo.ListCount Then
                        objCbo.AddItem RsTemp.Fields("st52") & " " & GetPrjSalesNM(RsTemp.Fields("st52"))
                     End If
                     RsTemp.MoveNext
                  Loop
               End If
            End If
         End If
      '案件目前表單
      ElseIf UCase(stFormN) = "FRM210147" Then
         '帶人主管
         strSql = "select st01,st02 from staff where substr(st01,1,1)<>'F' and st04='1' and instr(';'||st52||';'||st53||';'||st54||';'||st55||';',';" & strUserNum & ";')>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               For i = 0 To objCbo.ListCount - 1
                  If InStr(objCbo.List(i), RsTemp(0)) = 1 Then
                     Exit For
                  End If
               Next i
               If i = objCbo.ListCount Then
                  objCbo.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
               End If
               RsTemp.MoveNext
            Loop
         End If
      End If
   Else
   '2025/8/5 END
      'Add By Sindy 2023/1/4
      If InStr(Pub_GetSpecMan("全所智權部主管"), strUserNum) > 0 Then
         strSql = "select * from staff where substr(st15,1,1)='S' and st04='1' and length(st01)=5 and st01 not in('001-1') order by st15,st01"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            RsTemp.MoveFirst
            Do While Not RsTemp.EOF
               For i = 0 To objCbo.ListCount - 1
                  If InStr(objCbo.List(i), RsTemp(0)) = 1 Then
                     Exit For
                  End If
               Next i
               If i = objCbo.ListCount Then
                  objCbo.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
               End If
               RsTemp.MoveNext
            Loop
         End If
      End If
      '2023/1/4 END
      
      '帶人的權限
      Call Pub_SetSAManageEmpCombo(strUserNum, objCbo, False)
   '   '開放部份智權同仁的資料給彥葶操作
   '   If InStr(Pub_GetSpecMan("A8"), strUserNum) > 0 Then
   '      strTemp = Pub_GetSpecMan("A7")
   '      arrData = Split(strTemp, ";")
   '      For i = 0 To UBound(arrData)
   '         objCbo.AddItem arrData(i) & " " & GetPrjSalesNM(CStr(arrData(i)))
   '      Next
   '   End If
      
      '帶人主管抓虛建編號 ex.86047.高國碩,要帶出20011.中一區
      strSql = "select st01,st02 from staff where st01<'63001' and instr(';'||st52||';'||st53||';'||st54||';'||st55||';',';" & strUserNum & ";')>0"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            For i = 0 To objCbo.ListCount - 1
               If InStr(objCbo.List(i), RsTemp(0)) = 1 Then
                  Exit For
               End If
            Next i
            If i = objCbo.ListCount Then
               objCbo.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
            End If
            RsTemp.MoveNext
         Loop
      End If
      '抓出帶的離職人員
      'Modify By Sindy 2023/1/4 + 離職3個月內的人員
      If bolInTurnover = True Then
         strUser = PUB_GetSalesList(strUserNum, PUB_GetStaffST15(strUserNum, "1"), PUB_GetStaffST15(strUserNum, "1"), PUB_GetST06(strUserNum))
         If strUser <> "" Then
            strSql = "select cp13,st02 from staff,(" & _
                     " select cp13 from caseprogress" & _
                     " where cp13 in(select st01 from staff where st01 in(" & strUser & ") and st04<>'1' and st51>=" & CompDate(1, -3, strSrvDate(1)) & ")" & _
                     " and cp06 is not null and cp27 is null and cp57 is null" & _
                     " Union" & _
                     " select np10 from nextprogress" & _
                     " where np10 in(select st01 from staff where st01 in(" & strUser & ") and st04<>'1' and st51>=" & CompDate(1, -3, strSrvDate(1)) & ")" & _
                     " and np08 is not null and np06 is null) A" & _
                     " where cp13=st01(+)" & _
                     " group by cp13,st02 order by cp13 desc"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               RsTemp.MoveFirst
               Do While Not RsTemp.EOF
                  For i = 0 To objCbo.ListCount - 1
                     If InStr(objCbo.List(i), RsTemp(0)) = 1 Then
                        Exit For
                     End If
                  Next i
                  If i = objCbo.ListCount Then
                     objCbo.AddItem RsTemp.Fields("cp13") & " " & RsTemp.Fields("st02")
                  End If
                  RsTemp.MoveNext
               Loop
            End If
         End If
      End If
   End If

   objCbo.Text = objCbo.List(0)
End Sub

'Add By Sindy 2024/5/7 解析CRL55中案源案號
Public Function GetCRL55toLOS15(strText) As String
Dim arrTmp As Variant, ii As Integer
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String

   GetCRL55toLOS15 = ""
   If Trim(strText) = "" Then Exit Function
   
   arrTmp = Split(strText, ",")
   For ii = LBound(arrTmp) To UBound(arrTmp)
      Str01 = SystemNumber(CStr(arrTmp(ii)), 1)
      Str02 = SystemNumber(CStr(arrTmp(ii)), 2)
      Str03 = SystemNumber(CStr(arrTmp(ii)), 3)
      Str04 = SystemNumber(CStr(arrTmp(ii)), 4)
      If Str01 = "" And Len(arrTmp(ii)) = 8 Then '案源案號
         GetCRL55toLOS15 = arrTmp(ii)
         Exit For
      End If
   Next ii
End Function

'Add By Sindy 2022/8/24
'新增接洽單簽核資料
'bolReSend=True:重送
Public Function PUB_AddConsultRecvFlow(ByVal strCRL01 As String, Optional ByVal bolReSend As Boolean = False) As Boolean
Dim Rs As New ADODB.Recordset
Dim strUpdDate As String, strUpdTime As String
Dim strEmpSet1 As String, strEmpSet2 As String, intCnt As Integer
Dim varTmp As Variant
Dim ii As Integer
Dim strCRL03 As String, strCRL07 As String, strCRL08 As String
Dim strCRL66 As String, strCRL69 As String
Dim strNewCU As String, bolNewEmp As Boolean
Dim strCRL74 As String
Dim strF0308 As String '下一處理人員
Dim strF0309 As String
Dim bolBLawLos17 As Boolean, strCRL55 As String
Dim strLOS15 As String '案源案號 Add By Sindy 2024/5/7
Dim dblCRL146 As Double, strCRL147 As String 'Add By Sindy 2022/11/4
Dim strFlowEmp As String, strCaseType As String
Dim strCRL152 As String '自行送簽核 Add By Sindy 2023/4/7
   
On Error GoTo ErrHand
   
   '讀取接洽單資料
   strSql = "select CONSULTRECORDLIST.*,GetCRCaseNmFee(crl01,'3') as CaseType from CONSULTRECORDLIST where CRL01='" & strCRL01 & "'"
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCRL03 = "" & Rs.Fields("CRL03") '智權人員
      strCRL66 = "" & Rs.Fields("CRL66") '有對造已簽准
      'Modify By Sindy 2024/7/19 因專業部的特例主管都會設定到總經理,
      '但對造都必需事前呈報給總經理同意,方可收文
      '所以在專業部系統面不適用於特例簽核
      If Left(PUB_GetStaffST15(strCRL03, "1"), 1) <> "S" Then strCRL66 = ""
      '2024/7/19 END
      strCRL69 = "" & Rs.Fields("CRL69") '呈主管簽核
      strCRL07 = "" & Rs.Fields("CRL07") '系統別
      strCRL08 = "" & Rs.Fields("CRL08") '案號流水號
      strCRL74 = "" & Rs.Fields("CRL74") '相關案號的類別
      strCRL55 = "" & Rs.Fields("CRL55")
      strLOS15 = GetCRL55toLOS15(strCRL55) '案源案號 Add By Sindy 2024/5/7
      dblCRL146 = Val("" & Rs.Fields("CRL146")) '點數低於底價
      strCaseType = "" & Rs.Fields("CaseType") '案件性質
      strCRL147 = "" & Rs.Fields("CRL147") '價格已核准
      If strCRL147 = "Y" And InStr(strCRL69, "點數低於底價") > 0 Then
         strCRL69 = Trim(Replace(strCRL69, Mid(strCRL69, InStr(strCRL69, "點數低於底價"), InStr(strCRL69, vbCrLf) + 1), ""))
      End If
      strCRL152 = "" & Rs.Fields("CRL152") '自行送簽核 'Add By Sindy 2023/4/7
   End If
   '讀取接洽單申請人資料有新客戶
   strSql = "select * from consultrecapp where CRA01='" & strCRL01 & "' and CRA03='Y'"
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strNewCU = "Y"
   End If
   
   strUpdDate = strSrvDate(1)
   strUpdTime = Right("000000" & ServerTime, 6)
   
   '******************
   '更新表單主檔
   '******************
   strSql = "update FLOW003 set" & _
            " F0307='" & strUserNum & "',F0310='" & strUserNum & "'" & _
            ",F0311=" & strUpdDate & ",F0312=" & strUpdTime & _
            " where F0301='" & strCRL01 & "' and F0302='" & Flow_接洽單 & "'"
   cnnConnection.Execute strSql, intI
   
   '******************
   '新增表單簽核檔
   '******************
   '先清空
   strSql = "delete From FLOW002 where F0201=" & CNULL(strCRL01)
   cnnConnection.Execute strSql
   
   '跨所也是要走特例簽核
   strSql = "select * from consultrecapp where CRA01='" & strCRL01 & "' and CRA27='Y'"
   intI = 1
   Set Rs = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCRL66 = "Y" '特例
   End If
   
   '抓”智權部”的簽核主管
   If Left(PUB_GetStaffST15(strCRL03, "1"), 1) = "S" Then
      'A1.簽核人員
      '用”智權人員”抓簽核主管表單設定
      Call GetFLOW001Person(strCRL03, Flow_接洽單, , , strEmpSet1, strEmpSet2)
      
      'L-888888(案件名稱：特殊狀況發函)電子收文時:
      '智權人員收文時，非區主管人員的接洽單(依L介紹案源流程)須由區主管簽核
      '區主管的接洽單須由杜協理簽核，杜協理的接洽單須由總經理簽核。
      If strCRL07 = "L" And strCRL08 = "888888" Then
         bolNewEmp = True
         If InStr(Pub_GetSpecMan("全所智權部主管"), strCRL03) > 0 Then
            strEmpSet1 = Pub_GetSpecMan("總經理員工編號")
         Else
            strSql = "SELECT ST01,ST15,A0908 " & _
                    "From STAFF, ACC090 WHERE ST01='" & strCRL03 & "' AND ST15=A0901(+) " & _
                    "AND A0908 is not null "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If RsTemp.Fields("A0908") = strCRL03 Then
                  strEmpSet1 = Pub_GetSpecMan("全所智權部主管")
               Else
                  strEmpSet1 = RsTemp.Fields("A0908")
               End If
            Else
               strEmpSet1 = Pub_GetSpecMan("全所智權部主管")
            End If
         End If
      Else
         '簽核人員1非本人者，不可直接收文者,只抓簽核主管1
         strSql = "select * from flow001 where f0101='" & strCRL03 & "' and f0103<>f0101 and f0102='" & Flow_接洽單 & "'"
         intI = 1
         Set Rs = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            bolNewEmp = True
         End If
      End If
      'Add By Sindy 2023/8/22 杜協理提:因扣點數的案件，均須經過我這裡，財務若未見及我的簽核，還是會來問我。
      '                       故若是負數之案件，麻煩還是都到我這裡，謝謝。
      If InStr(strCRL69, "有【負點數】") > 0 Then
         strCRL66 = "Y" '特例
      End If
      '2023/8/22 END
      
      'Modify By Sindy 2025/3/28 杜協理通知:已與林協理進行台中智權部各類簽核的討論，請協助將中所智權人員之各類簽核，
      '          包含人事.案件及帳款簽核的權限，原設定為林協理的部分，請均調整為本人，且各區依北所各區之管理模式進行設定。
      '智權部 "點數低於底價" 超過1000元 為特例簽核
'      '例外:中所智權人員”點數低於底價”不算特例簽核，均需給林柄佑(中所智權部主管)簽核
'      If Left(PUB_GetStaffST15(strCRL03, "1"), 2) = "S2" Then
'         If InStr(strCRL69, "點數低於底價") > 0 Then
'            If InStr(strEmpSet1, Pub_GetSpecMan("中所智權部主管")) = 0 Then
'               strEmpSet1 = strEmpSet1 & "," & Pub_GetSpecMan("中所智權部主管")
'            End If
'         End If
'      Else
      '2025/3/28 END
         If dblCRL146 > 1 And InStr(strCRL69, "點數低於底價") > 0 Then
            strExc(10) = Trim(Replace(strCRL69, Mid(strCRL69, InStr(strCRL69, "點數低於底價"), InStr(strCRL69, vbCrLf) + 1), ""))
            If Len(strExc(10)) = 0 Then
               strCRL69 = ""
               strCRL66 = "Y" '特例
            End If
         End If
'      End If
      
'      '不可直接收文者,只抓簽核主管1
'      If bolNewEmp = True Then
'         intCnt = 0
'         If strEmpSet1 <> "" Then
'            varTmp = Split(strEmpSet1, ",")
'            For ii = 0 To 0 'UBound(varTmp)
'               If varTmp(ii) <> "" Then
'                  intCnt = intCnt + 1
'                  strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A1'," & intCnt & "," & CNULL(CStr(varTmp(ii))) & ")"
'                  cnnConnection.Execute strSql
'               End If
'            Next ii
'         End If
'      End If
'      '有資料需呈報主管簽核
'      If strCRL69 <> "" Then
'         If strEmpSet1 <> "" Then
'            varTmp = Split(strEmpSet1, ",")
'            For ii = 0 To UBound(varTmp)
'               If varTmp(ii) <> "" Then
'                  strSql = "select * from FLOW002 where F0201='" & strCRL01 & "' and F0204='" & varTmp(ii) & "'"
'                  intI = 1
'                  Set Rs = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 0 Then
'                     intCnt = intCnt + 1
'                     strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A1'," & intCnt & "," & CNULL(CStr(varTmp(ii))) & ")"
'                     cnnConnection.Execute strSql
'                  End If
'               End If
'            Next ii
'         Else
'            varTmp = Split(strEmpSet2, ",")
'            For ii = 0 To UBound(varTmp)
'               If varTmp(ii) <> "" Then
'                  intCnt = intCnt + 1
'                  strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A1'," & intCnt & "," & CNULL(CStr(varTmp(ii))) & ")"
'                  cnnConnection.Execute strSql
'               End If
'            Next ii
'         End If
'      End If
'      'A2.特例主管
'      'Add By Sindy 2023/4/7 + strCRL152=Y 自行送簽核
'      If (strCRL66 = "Y" Or (strCRL69 <> "" And intCnt = 0) Or strCRL152 = "Y") And _
'         Not (strCRL07 = "L" And strCRL08 = "888888") Then
'         intCnt = 0
'         If strEmpSet2 <> "" Then
'            varTmp = Split(strEmpSet2, ",")
'            For ii = 0 To UBound(varTmp)
'               If varTmp(ii) <> "" Then
'                  strSql = "select * from FLOW002 where F0201='" & strCRL01 & "' and F0204='" & varTmp(ii) & "'"
'                  intI = 1
'                  Set Rs = ClsLawReadRstMsg(intI, strSql)
'                  If intI = 0 Then
'                     intCnt = intCnt + 1
'                     strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A2'," & intCnt & "," & CNULL(CStr(varTmp(ii))) & ")"
'                     cnnConnection.Execute strSql
'                  End If
'               End If
'            Next ii
'         End If
'      End If
      
   '抓”專業部及其他同仁”的簽核主管
   Else
      'A1.簽核人員
      '專利處程序
      If PUB_GetStaffST15(strUserNum, "1") = "P12" Then
         '只有P案601領證、605年費
         'Modify By Sindy 2023/4/12 + PUB_GetStaffST15(strCRL03, "1") = "P12" ex:P-127887 排除CRL03=顧服組(W2001)
         If strCRL07 = "P" And _
            (InStr(strCaseType, "601") > 0 Or InStr(strCaseType, "605") > 0) And _
            PUB_GetStaffST15(strCRL03, "1") = "P12" Then
            '用”智權人員”抓簽核主管表單設定
            strFlowEmp = strCRL03
         Else
            '用”操作人員”抓簽核主管表單設定
            strFlowEmp = strUserNum
         End If
      Else
         '用”智權人員”抓簽核主管表單設定
         strFlowEmp = strCRL03
      End If
      Call GetFLOW001Person(strFlowEmp, Flow_接洽單, , , strEmpSet1, strEmpSet2)
      '專利處程序於智權人員填寫顧服組時,增加發給總經理簽核
      If strCRL03 = "W2001" And PUB_GetStaffST15(strUserNum, "1") = "P12" Then
         'Modify By Sindy 2024/1/17
         If strEmpSet1 <> "" Then strEmpSet1 = strEmpSet1 & "," & Pub_GetSpecMan("總經理員工編號")
         If strEmpSet2 <> "" Then strEmpSet2 = strEmpSet2 & "," & Pub_GetSpecMan("總經理員工編號")
      End If
      
      '簽核人員1非本人者，不可直接收文者,只抓簽核主管1
      strSql = "select * from flow001 where f0101='" & strFlowEmp & "' and f0103<>f0101 and f0102='" & Flow_接洽單 & "'"
      intI = 1
      Set Rs = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         bolNewEmp = True
      End If
      
'      '不可直接收文需主管簽核者
'      If bolNewEmp = True Then
'         intCnt = 0
'         If strEmpSet1 <> "" Then
'            varTmp = Split(strEmpSet1, ",")
'            For ii = 0 To UBound(varTmp)
'               If varTmp(ii) <> "" Then
'                  intCnt = intCnt + 1
'                  strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A1'," & intCnt & "," & CNULL(CStr(varTmp(ii))) & ")"
'                  cnnConnection.Execute strSql
'               End If
'            Next ii
'         End If
'      End If
'      'A2.特例主管
'      'Add By Sindy 2023/4/7 + strCRL152=Y 自行送簽核
'      If strCRL66 = "Y" Or strCRL69 <> "" Or strCRL152 = "Y" Or _
'         (strCRL07 = "L" And strCRL08 = "888888") Then
'         intCnt = 0
'         If strEmpSet2 <> "" Then
'            varTmp = Split(strEmpSet2, ",")
'            For ii = 0 To UBound(varTmp)
'               If varTmp(ii) <> "" Then
'                  intCnt = intCnt + 1
'                  strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A2'," & intCnt & "," & CNULL(CStr(varTmp(ii))) & ")"
'                  cnnConnection.Execute strSql
'               End If
'            Next ii
'         End If
'      End If
   End If
   
   'Modify By Sindy 2023/11/1 改共用不要寫在上列if裡,而寫2次
   '不可直接收文者,只抓簽核主管1
   If bolNewEmp = True Then
      intCnt = 0
      If strEmpSet1 <> "" Then
         varTmp = Split(strEmpSet1, ",")
         For ii = 0 To 0 'UBound(varTmp)
            If varTmp(ii) <> "" Then
               intCnt = intCnt + 1
               strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A1'," & intCnt & "," & CNULL(CStr(varTmp(ii))) & ")"
               cnnConnection.Execute strSql
            End If
         Next ii
      End If
   End If
   'Add By Sindy 2024/1/17 W3001增加發給總經理簽核
   If strCRL03 = "W3001" Then
      varTmp(0) = Pub_GetSpecMan("總經理員工編號")
      strSql = "select * from FLOW002 where F0201='" & strCRL01 & "' and F0204='" & varTmp(0) & "'"
      intI = 1
      Set Rs = ClsLawReadRstMsg(intI, strSql)
      If intI = 0 Then
         intCnt = intCnt + 1
         strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A1'," & intCnt & "," & CNULL(CStr(varTmp(0))) & ")"
         cnnConnection.Execute strSql
      End If
   End If
   '2024/1/17 END
   '有資料需呈報主管簽核
   If strCRL69 <> "" Then
      If strEmpSet1 <> "" Then
         varTmp = Split(strEmpSet1, ",")
         For ii = 0 To UBound(varTmp)
            If varTmp(ii) <> "" Then
               strSql = "select * from FLOW002 where F0201='" & strCRL01 & "' and F0204='" & varTmp(ii) & "'"
               intI = 1
               Set Rs = ClsLawReadRstMsg(intI, strSql)
               If intI = 0 Then
                  intCnt = intCnt + 1
                  strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A1'," & intCnt & "," & CNULL(CStr(varTmp(ii))) & ")"
                  cnnConnection.Execute strSql
               End If
            End If
         Next ii
      Else
         varTmp = Split(strEmpSet2, ",")
         For ii = 0 To UBound(varTmp)
            If varTmp(ii) <> "" Then
               'Add By Sindy 2023/11/2
               strSql = "select * from FLOW002 where F0201='" & strCRL01 & "' and F0204='" & varTmp(ii) & "'"
               intI = 1
               Set Rs = ClsLawReadRstMsg(intI, strSql)
               If intI = 0 Then
               '2023/11/2 END
                  intCnt = intCnt + 1
                  strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A1'," & intCnt & "," & CNULL(CStr(varTmp(ii))) & ")"
                  cnnConnection.Execute strSql
               End If
            End If
         Next ii
      End If
   End If
   'A2.特例主管
   'Add By Sindy 2023/4/7 + strCRL152=Y 自行送簽核
   If (strCRL66 = "Y" Or (strCRL69 <> "" And intCnt = 0) Or strCRL152 = "Y") And _
      Not (strCRL07 = "L" And strCRL08 = "888888") Then
      intCnt = 0
      If strEmpSet2 <> "" Then
         varTmp = Split(strEmpSet2, ",")
         For ii = 0 To UBound(varTmp)
            If varTmp(ii) <> "" Then
               strSql = "select * from FLOW002 where F0201='" & strCRL01 & "' and F0204='" & varTmp(ii) & "'"
               intI = 1
               Set Rs = ClsLawReadRstMsg(intI, strSql)
               If intI = 0 Then
                  intCnt = intCnt + 1
                  strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A2'," & intCnt & "," & CNULL(CStr(varTmp(ii))) & ")"
                  cnnConnection.Execute strSql
               End If
            End If
         Next ii
      End If
   End If
   '2023/11/1 END
   
   '*****若是代他人填單,簽核檔中若自己也是簽核人員之一時,一併確認掉
   strSql = "update FLOW002 set " & _
            "F0205='" & strUpdDate & "'" & _
            ",F0206='" & strUpdTime & "'" & _
            ",F0207='1'" & _
            " where F0201='" & strCRL01 & "' and F0204='" & strUserNum & "' and F0207 is null"
   cnnConnection.Execute strSql, intI
   '*****END
   
'A3.櫃檯人員
   If strNewCU = "Y" Then '有新客戶
      strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A3',1,'A3')"
      cnnConnection.Execute strSql
   End If
'A4.法務案源(A類)
   If Len(strCRL74) = 2 Then
      If Left(strCRL74, 1) = "A" Then
         strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A4',1,'A4')"
         cnnConnection.Execute strSql
      Else
         'B類的法律接洽單
         strSql = "SELECT los15,los17,los18 FROM LawOfficeSource where (los17='" & strCRL01 & "' or los20='" & strCRL01 & "') and los15='" & strLOS15 & "'"
         intI = 1
         Set Rs = ClsLawReadRstMsg(intI, strSql, , True)
         If intI = 1 Then
            bolBLawLos17 = True
            
            'Add By Sindy 2023/2/8
            strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A4',1,'A4')"
            cnnConnection.Execute strSql
            '2023/2/8 END
         End If
      End If
   End If
   
   'CFT案件,ACS：尚未電子化，於智權人員收文即可
   '法律案源也尚未電子化
   If InStr("," & Flow_不走分案的系統別, "," & strCRL07) = 0 And Left(strCRL74, 1) <> "A" And bolBLawLos17 = False Then
      'A5.分所專利
      If (strCRL07 = "P" Or strCRL07 = "PS" Or strCRL07 = "CFP" Or strCRL07 = "CPS") And _
         PUB_GetST06(strCRL03) <> "1" Then
         strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A5',1,'A5')"
         cnnConnection.Execute strSql
      End If
      'A6.北所分案
      strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A6',1,'A6')"
      cnnConnection.Execute strSql
      'A7.程序人員
      strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A7',1,'A7')"
      cnnConnection.Execute strSql
   End If
   
   '記錄重送訊息
   If bolReSend = True Then
      strSql = GetInsertFLOW004Sql(Trim(strCRL01), strUserNum, strUpdDate, strUpdTime, Flow_重送, "")
      cnnConnection.Execute strSql
   End If
   
   '讀取下一處理人員
   If GetNextProPerson_Flow(strCRL01, strCRL03, strF0308, strF0309) = False Then GoTo ErrHand
   
   PUB_AddConsultRecvFlow = True
   Set Rs = Nothing
   Exit Function
   
ErrHand:
   PUB_AddConsultRecvFlow = False
   Set Rs = Nothing
   
   If Err.Number <> 0 Then
      MsgBox Err.Description & vbCrLf & _
      ";strSql = " & strSql, vbCritical, "PUB_AddConsultRecvFlow"
   End If
End Function

'傳回下一處理人員
'Modified by Morgan 2015/11/3 +bolUnfinished:未完成(目前為非臺灣指示信判發使用)
Public Function GetNextProPerson_Flow(strKEY01 As String, strF0316 As String, ByRef strF0204 As String, _
   Optional ByRef strF0309 As String, Optional bolUnfinished As Boolean = False) As Boolean
   
Dim rsTmp As New ADODB.Recordset
Dim Rs As New ADODB.Recordset
Dim bolRunLoopFindNextP As Boolean
Dim strTemp As Variant, i As Integer
Dim strF0202 As String 'Modify By Sindy 2022/8/10 Integer 改 String
Dim intF0203 As Integer
Dim strCurrF0308 As String, strCRL08 As String, strCurrF0309 As String
Dim m_bolIsRest1Day As Boolean
Dim strF0302 As String, strCRL07 As String, strCRL74 As String 'Add By Sindy 2022/9/19
Dim bolRecv As Boolean, bolHadA6 As Boolean, strCUM01 As String, strCUM02 As String 'Add By Sindy 2022/9/26
Dim strCRL55 As String, bolBLawLos17 As Boolean 'Add By Sindy 2022/10/7
Dim strLOS15 As String '案源案號 Add By Sindy 2024/5/7
Dim strSubject As String, strContent As String 'Add By Sindy 2022/10/7
Dim strCPM35 As String 'Add By Sindy 2022/10/17
Dim strF0307 As String 'Add By Sindy 2022/10/20 紀錄上一處理人員
   
On Error GoTo ErrHand
   
   GetNextProPerson_Flow = True
   
   '讀取案件表單主檔資料
   strSql = "SELECT * FROM FLOW003,CONSULTRECORDLIST WHERE F0301='" & strKEY01 & "'" & _
            " and F0301=CRL01(+)"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCurrF0308 = "" & RsTemp.Fields("F0308") '目前下一處理人員
      strCurrF0309 = "" & RsTemp.Fields("F0309") '目前表單狀態 'Add By Sindy 2021/6/23
      strF0302 = "" & RsTemp.Fields("F0302")
      'Add By Sindy 2022/10/7
      If strF0309 = "" Then '沒有傳入,抓現況
         strF0316 = "" & RsTemp.Fields("F0316")
         strF0204 = "" & RsTemp.Fields("F0308")
         strF0309 = "" & RsTemp.Fields("F0309")
      End If
      '2022/10/7 END
      strCRL07 = "" & RsTemp.Fields("CRL07") '系統別
      strCRL08 = "" & RsTemp.Fields("CRL08")
      strCRL74 = "" & RsTemp.Fields("CRL74") '相關案號的類別
      strCRL55 = "" & RsTemp.Fields("CRL55")
      strLOS15 = GetCRL55toLOS15(strCRL55) '案源案號 Add By Sindy 2024/5/7
   End If
   
   'Add By Sindy 2022/9/26
   If strF0302 = Flow_接洽單 Then
      '接洽記錄單案件性質
      bolRecv = False '未收文
      strSql = "SELECT * FROM ConsultRecCMP WHERE CRC01='" & strKEY01 & "'" & _
               " and CRC08 is null"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 0 Then
         bolRecv = True '已收文
      End If
      '檢查是否有分案簽核
      bolHadA6 = False '無分案簽核
      strSql = "SELECT * FROM FLOW002 WHERE F0201='" & strKEY01 & "'" & _
               " and F0202>='A6'" 'A6.北所分案
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         bolHadA6 = True '有分案簽核
      End If
   End If
   '2022/9/26 END
   
ReCheck:
   bolRunLoopFindNextP = True '預設一進此函數都會Run一次do迴圈
   Do While bolRunLoopFindNextP = True
      bolRunLoopFindNextP = False
      
      '讀取案件表單簽核資料
      strSql = "SELECT * FROM FLOW002 " & _
               "WHERE F0201='" & strKEY01 & "' and F0207 is null order by F0202 asc,F0203 asc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         
         strF0202 = rsTmp.Fields("F0202")
         intF0203 = rsTmp.Fields("F0203")
         strF0204 = rsTmp.Fields("F0204")
         
         'Modify By Sindy 2022/8/9
         If strF0202 = "1" Or strF0202 = "A1" Or strF0202 = "A2" Then '簽核主管
            strF0309 = Flow_主管審核中
         ElseIf strF0202 = "2" Or strF0202 = "A3" Or strF0202 = "A4" Or strF0202 = "A7" Then '行政人員(程序人員)
            strF0309 = Flow_處理中
         'Add By Sindy 2022/8/9
         ElseIf strF0202 = "A5" Or strF0202 = "A6" Then '分案主管
            strF0309 = Flow_待分案
         '2022/8/9 END
         ElseIf strF0202 = "3" Then '補看人員
            'Modified by Morgan 2015/11/3
            'strF0309 = Flow_已完成
            If bolUnfinished Then
               strF0309 = Flow_指示信判發中
            Else
               strF0309 = Flow_已完成
            End If
            'end 2015/11/3
         End If
         
      'Modify By Sindy 2015/8/3
      Else
         strF0204 = ""
         'Add By Sindy 2022/9/22
         If strF0302 = Flow_接洽單 Then
            If bolHadA6 = False Then '無分案簽核
               If bolRecv = True Then
                  strF0204 = strUserNum
                  strF0309 = Flow_已收文
               Else
                  
                  '檢查是否B類的法律接洽單
                  bolBLawLos17 = False
                  If Len(strCRL74) = 2 Then
                     bolBLawLos17 = True 'Add By Sindy 2023/4/14 法律所不能電子收文,維持列印紙本經櫃檯收文
                     If Left(strCRL74, 1) <> "A" Then
                        strSql = "SELECT los15,los17,los18 FROM LawOfficeSource where (los17='" & strKEY01 & "' or los20='" & strKEY01 & "') and los15='" & strLOS15 & "'"
                        intI = 1
                        Set Rs = ClsLawReadRstMsg(intI, strSql, , True)
                        If intI = 1 Then
                           'bolBLawLos17 = True
                           '是否可以收文
                           strSql = "SELECT * FROM FLOW002 WHERE F0201='" & strKEY01 & "'" & _
                                    " and F0202='A4' and F0207='1'" '法律,同意
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                           If intI = 1 Then
                              bolBLawLos17 = False 'B類可以收文
                           End If
                        End If
                     End If
                  End If
               
                  '***接洽單電子收文***
                  If bolBLawLos17 = True Then
                     strF0204 = strUserNum
                     strF0309 = Flow_處理中 'Flow_待分案
                  Else
                     If PUB_AutoRecvCRLMain(strCRL07, strKEY01) = False Then
                        GoTo ErrHand
                     Else
                        bolRecv = True
                        'CFT案件：尚未電子化，於智權人員收文成功後，直接Email通知陳金蓮，以利分案處理。
                        If InStr("," & Flow_不走分案的系統別, "," & strCRL07) > 0 Then
                           '寫入要發通知信的人員
                           'CaseUseMemo:
                           'cum01 = 身份類別
                           'cum02 = 收受者
                           'cum03 = 0
                           'cum04 = 表單狀態
                           'cum05 = 03.電子收文(或簽核)通知信
                           'cum06 = 操作人員
                           'CUM09 = 接洽單編號
                           If strCRL07 = "ACS" Then
                              'strCUM01 = "ACS"
                              'Modify By Sindy 2023/11/7
                              'strCUM02 = "ACS01" 'Modify By Sindy 2023/6/8 Pub_GetSpecMan("ACS分案人員")
                              strSubject = "(接洽單已收文)請進行案件分案"
                              strContent = strSubject & ", 謝謝！"
                              strExc(0) = "select mc01,mc02 from mailcache" & _
                                          " where mc02='" & Pub_GetSpecMan("ACS分案人員") & "'" & _
                                            " and mc07='" & ChgSQL(strSubject) & "' and mc05 is null"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 0 Then
                                 strExc(0) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                             " values ('" & strUserNum & "','" & Pub_GetSpecMan("ACS分案人員") & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss')" & _
                                             ",'" & ChgSQL(strSubject) & "','" & ChgSQL(strContent) & "',null)"
                                 cnnConnection.Execute strExc(0), intI
                              End If
                              '2023/11/7 END
                           Else
                              strCUM01 = "CFT"
                              strCUM02 = Pub_GetSpecMan("D")
                              'Add By Sindy 2025/5/8 + and cum01='" & strCUM01 & "'
                              strExc(0) = "select cum02 from CaseUseMemo" & _
                                          " where cum05='03'" & _
                                            " and cum06=" & CNULL(strUserNum) & _
                                            " and cum02=" & CNULL(strCUM02) & _
                                            " and cum01='" & strCUM01 & "'"
                              intI = 1
                              Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                              If intI = 0 Then
                                 strExc(0) = "insert into CaseUseMemo(cum01,cum02,cum03,cum04,cum05,cum09)" & _
                                             " values('" & strCUM01 & "','" & strCUM02 & "','0','" & Flow_已收文 & "','03','" & strKEY01 & "')"
                                 cnnConnection.Execute strExc(0), intI
                              End If
                           End If
                        End If
                     End If
                     strF0204 = strUserNum
                     strF0309 = Flow_已收文
                  End If
               End If
            Else
               strF0204 = strUserNum
               strF0309 = Flow_已分案
            End If
         End If
         '2022/9/22 END
      '2015/8/3 END
      End If
      If rsTmp.State <> 0 Then rsTmp.Close
   Loop
         
   'Add By Sindy 2023/6/2 簽核主管時,要再確認一下是否有該主管需要一併簽核的資料
   If strF0309 = Flow_主管審核中 Then
      '檢查下一處理人員是否已簽核過,若是,則直接Update其相同的簽核日期資料
      strSql = "SELECT * FROM FLOW002 " & _
               "WHERE F0201='" & strKEY01 & "' and F0204='" & strF0204 & "' and F0207 is not null order by F0202,F0203 asc "
      intI = 1
      Set Rs = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strSql = "update FLOW002 set " & _
                  "F0205=" & Rs.Fields("F0205") & _
                  ",F0206=" & Rs.Fields("F0206") & _
                  ",F0207='" & Rs.Fields("F0207") & "' " & _
                  "WHERE F0201='" & strKEY01 & "' and F0204='" & strF0204 & "' and F0207 is null"
         cnnConnection.Execute strSql
         GoTo ReCheck '重新檢查下一處理人員
      End If
   End If
   '2023/6/2 END
   
   'Add By Sindy 2022/9/19
   If strF0302 = Flow_接洽單 Then
      If strF0204 = "A4" And Left(strCRL74, 1) = "A" Then 'A4.法務案源
         strSql = "SELECT los15 FROM LawOfficeSource where los17='" & strKEY01 & "'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If "" & RsTemp.Fields("los15") <> "" Then
               'A類要EMail 通知法務人員/窗口(BC類在分案)
               PUB_AddMailCache_LOS "1", RsTemp.Fields("los15")
            End If
         End If
      
      '下一處理人員為分案主管時, 未收文就先啟動電子收文
      ElseIf (strF0204 = "A5" Or strF0204 = "A6") And bolRecv = False Then
         '***接洽單電子收文***
         If PUB_AutoRecvCRLMain(strCRL07, strKEY01) = False Then
            GoTo ErrHand
         Else
            bolRecv = True '已收文 ***** 此變數很重要,才不會跑N次迴圈
            
            '專利:
            If strCRL07 = "P" Or strCRL07 = "PS" Or strCRL07 = "CFP" Or strCRL07 = "CPS" Then
               strCPM35 = PUB_GetCPM35(strKEY01, strCRL07)
               '1.先補文件再呈分案主管
               If strCPM35 = "1" Then
                  strF0309 = Flow_程序補件
                  'Modify By Sindy 2023/2/22 ex:P-131020
                  'strF0307 = "A6" '記錄上一處理人員
                  strF0307 = strF0204 '記錄上一處理人員
                  '2023/2/22 END
                  strF0204 = "A7" 'A7.程序人員
                  
               '2.程序承辦不需經由主管分案 3.可能程序或工程師承辦
               ElseIf strCPM35 = "2" Or strCPM35 = "3" Then
                  strF0309 = Flow_處理中
                  strSql = "update FLOW002 set" & _
                           " F0205='" & strSrvDate(1) & "'" & _
                           ",F0206='" & Right("000000" & ServerTime, 6) & "'" & _
                           ",F0207='1',F0204='QPGMR'" & _
                           " where F0201='" & strKEY01 & "' and F0202='A5' and F0207 is null"
                  cnnConnection.Execute strSql
                  strSql = "update FLOW002 set" & _
                           " F0205='" & strSrvDate(1) & "'" & _
                           ",F0206='" & Right("000000" & ServerTime, 6) & "'" & _
                           ",F0207='1',F0204='QPGMR'" & _
                           " where F0201='" & strKEY01 & "' and F0202='A6' and F0207 is null"
                  cnnConnection.Execute strSql
                  strF0307 = "A6" '記錄上一處理人員
                  strF0204 = "A7" 'A7.程序人員
                  
                  '若全部案件性質都有掛承辦人且CP157已有日期,就不用經過程序人員操作分案作業, 系統自動簽核
                  strSql = "select CRC01,CRC03,CRC08,CP14 from ConsultRecCMP,caseprogress" & _
                           " where CRC01='" & strKEY01 & "' and CRC08=cp09(+) and (CP14 is null or CP157 is null or CRC08 is null)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 0 Then
                     '程序人員不需操作分案作業,系統自動簽核
                     strSql = "update FLOW002 set" & _
                              " F0205='" & strSrvDate(1) & "'" & _
                              ",F0206='" & Right("000000" & ServerTime, 6) & "'" & _
                              ",F0207='1',F0204='QPGMR'" & _
                              " where F0201='" & strKEY01 & "' and F0202='A7' and F0207 is null"
                     cnnConnection.Execute strSql
                     strF0307 = "A7" '記錄上一處理人員
                  End If
                  GoTo ReCheck '重新檢查下一處理人員
               End If
               
            '商標:T開頭和FCT
            ElseIf Left(strCRL07, 1) = "T" Or strCRL07 = "FCT" Then
               If strCRL07 = "T" Then
                  'Modify By Sindy 2023/1/7 程序人員還是需要操作程序分案作業,因為 延期要輸入總收文號/放棄專用權要上齊備日
'                  '(原)接洽單自動收文處理程序:T 303延期,201補正,206放棄專用權,211檢送同意書
'                  '有掛承辦人, 就更新北所分案日 ==> 視為程序人員已分案
'                  strSql = "update caseprogress set cp157=" & strSrvDate(1) & _
'                           " where cp09 in(select CRC08 from ConsultRecCMP,caseprogress" & _
'                                          " where CRC01='" & strKEY01 & "' and CRC03 in('303','201','206','211')" & _
'                                          " and CRC08=cp09(+) and cp14 is not null)"
'                  cnnConnection.Execute strSql, intI
               End If
               
               '檢查若全部案件性質都有掛承辦人,不用進分案主管分案,系統自動簽核
               strSql = "select CRC01,CRC03,CRC08,CP14 from ConsultRecCMP,caseprogress" & _
                        " where CRC01='" & strKEY01 & "' and CRC08=cp09(+) and (cp14 is null or CRC08 is null)"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 0 Then '都有掛承辦人
                  '不用進分案主管簽核,系統自動簽核
                  strSql = "update FLOW002 set" & _
                           " F0205='" & strSrvDate(1) & "'" & _
                           ",F0206='" & Right("000000" & ServerTime, 6) & "'" & _
                           ",F0207='1',F0204='QPGMR'" & _
                           " where F0201='" & strKEY01 & "' and F0202='A6' and F0207 is null"
                  cnnConnection.Execute strSql, intI
                  strF0307 = "A6" '記錄上一處理人員
                  
                  '若全部案件性質都有掛承辦人且CP157已有日期,就不用經過程序人員操作分案作業, 系統自動簽核
                  strSql = "select CRC01,CRC03,CRC08,CP14 from ConsultRecCMP,caseprogress" & _
                           " where CRC01='" & strKEY01 & "' and CRC08=cp09(+) and (CP14 is null or CP157 is null or CRC08 is null)"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 0 Then
                     '程序人員不需操作分案作業,系統自動簽核
                     strSql = "update FLOW002 set" & _
                              " F0205='" & strSrvDate(1) & "'" & _
                              ",F0206='" & Right("000000" & ServerTime, 6) & "'" & _
                              ",F0207='1',F0204='QPGMR'" & _
                              " where F0201='" & strKEY01 & "' and F0202='A7' and F0207 is null"
                     cnnConnection.Execute strSql
                     strF0307 = "A7" '記錄上一處理人員
                  End If
                  GoTo ReCheck '重新檢查下一處理人員
               End If
'            Else
'               GoTo ReCheck '重新檢查下一處理人員
            End If
         End If
      End If
   End If
   '2022/9/19 END
   
   'Add By Sindy 2022/10/20
   If strF0307 = "" Then
      If strCurrF0308 <> "" Then
         'Modify By Sindy 2023/2/7
         'A5=分所分案,有可能是北所分案主管直接操作其分案,增加判斷
         If strCurrF0308 = "A5" And PUB_GetST06(strUserNum) = "1" Then
            strF0307 = "A6" 'A6.北所分案
         Else
         '2023/2/7 END
            strF0307 = strCurrF0308 'Modify By Sindy 2022/11/2
         End If
      Else
         strF0307 = strUserNum
      End If
   End If
   '2022/10/20 END
   '有異動下一處理人員時才須更新資料
   'Modify By Sindy 2021/6/23
   'If strF0204 <> "" And strF0204 <> strCurrF0308 Then
   If strF0204 <> "" And strF0309 <> "" And _
      (strF0204 <> strCurrF0308 Or strF0309 <> strCurrF0309) Then
   '2021/6/23 END
      strSql = "update FLOW003 set " & _
                "F0307='" & strF0307 & "'" & _
               ",F0308='" & strF0204 & "'" & _
               ",F0309='" & strF0309 & "'" & _
               " where F0301='" & strKEY01 & "' "
      cnnConnection.Execute strSql
      
      'Add By Sindy 2022/10/4 接洽單的主管審核中,櫃檯人員(收文櫃台 <writer@taie.com.tw>) 通知信
      If strF0302 = Flow_接洽單 Then
'         '中所區主管杜協理說正本給她，副本通知林協理，所有區主管流程上由杜協理控制，無須再切割出台中的區主管，這樣較單純。
'         If strCRL07 = "L" And strCRL08 = "888888" And _
'            strF0309 = Flow_主管審核中 And PUB_GetST06(strF0316) = "2" And InStr(Pub_GetSpecMan("全所智權部主管"), strF0204) > 0 Then
'
'            strSubject = GetPrjSalesNM(strF0316) & "(接洽單L-888888)電子簽核通知"
'            strContent = strSubject & vbCrLf & vbCrLf & "請至案件管理系統的 一般作業->案件表單查詢及簽核 項目中，進行表單簽核處理。"
'            strExc(0) = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
'               " values( '" & strUserNum & "','" & strF0204 & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'               ",'" & strSubject & "','" & strContent & "','" & Pub_GetSpecMan("中所智權部主管") & "')"
'            cnnConnection.Execute strExc(0)
'         Else
         
            If strF0309 = Flow_主管審核中 Or _
               (strF0204 = "A3" And strF0309 = Flow_處理中) Then
               '寫入要發通知信的人員
               'CaseUseMemo:
               'cum01 = 身份類別
               'cum02 = 收受者
               'cum03 = 0
               'cum04 = 表單狀態
               'cum05 = 03.電子收文(或簽核)通知信
               'cum06 = 操作人員
               'CUM09 = 接洽單編號
               strCUM01 = IIf(strF0309 = Flow_主管審核中, "A1", IIf(strF0204 = "A3", "A3", strF0309)) 'Add By Sindy 2025/5/8
               If strF0204 = "A3" Then
                  strCUM02 = "writer"
               Else
                  strCUM02 = strF0204
               End If
               'Add By Sindy 2025/5/8 + and cum01='" & strCUM01 & "'
               strExc(0) = "select cum02 from CaseUseMemo" & _
                           " where cum05='03'" & _
                             " and cum06=" & CNULL(strUserNum) & _
                             " and cum02=" & CNULL(strCUM02) & _
                             " and cum01='" & strCUM01 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  'Modify By Sindy 2025/5/8 改使用 strCUM01 變數
                  strExc(0) = "insert into CaseUseMemo(cum01,cum02,cum03,cum04,cum05,cum09)" & _
                              " values('" & strCUM01 & "'" & _
                              ",'" & strCUM02 & "','0','" & strF0309 & "','03','" & strKEY01 & "')"
                  cnnConnection.Execute strExc(0), intI
               End If
            End If
'         End If
      End If
      '2022/10/4 END
   End If
   
   Set rsTmp = Nothing
   Set Rs = Nothing
   Exit Function
   
ErrHand:
   GetNextProPerson_Flow = False
   Set rsTmp = Nothing
   Set Rs = Nothing
   If Err.Number <> 0 Then
      MsgBox Err.Description & vbCrLf & _
      ";strExc(0) = " & strExc(0) & vbCrLf & _
      ";strSql = " & strSql, vbCritical, "GetNextProPerson_Flow"
   End If
End Function

'取得接洽單案號資料
Public Function PUB_MailGetCRLCaseData(strCRL01 As String) As String
Dim strSalesDeptName As String, strSalesST06 As String, strSalesST15 As String, strCP05 As String
Dim oContext2 As String
Dim strCRA26 As String, strCRA27 As String
   
   oContext2 = ""
   strExc(0) = "select CRL01,CRC08,sqldatet(CP05) CP05" & _
            " from ConsultRecordList,ConsultRecCMP,caseprogress" & _
            " where CRL01='" & strCRL01 & "' and CRL01=CRC01(+) and CRC08=CP09(+)" & _
            " group by CRL01,CRC08,CP05" & _
            " order by CRL01,CRC08,CP05"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strCP05 = " " & RsTemp.Fields("CP05")
   End If
   
   '是否有對造
   strSql = "select * from consultrecapp where CRA01='" & strCRL01 & "' and CRA26='Y'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCRA26 = "Y"
   End If
   '是否有跨所
   strSql = "select * from consultrecapp where CRA01='" & strCRL01 & "' and CRA27='Y'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCRA27 = "Y"
   End If
   
   strExc(0) = "select CRL15,na03 申請國家" & _
            ",CRL07||'-'||decode(CRL08,null,'',CRL08)||decode(CRL09,null,'','-'||CRL09)||decode(CRL10,null,'','-'||CRL10) 案號" & _
            ",CRL17 案件名稱,GetCRCaseNmFee(crl01,'2') 案件性質,sqldatet(CRL12) 本所期限,sqldatet(CRL13) 法定期限" & _
            ",sum(CRC04) 總費用,CRL03,CRL04,CRL69" & _
            " from ConsultRecordList,ConsultRecCMP,nation" & _
            " where CRL01='" & strCRL01 & "' and CRL01=CRC01(+) and CRL15=na01(+)" & _
            " group by CRL01,CRL15,na03,CRL07,CRL08,CRL09,CRL10,CRL17,CRL12,CRL13,CRL03,CRL04,CRL69" & _
            " order by CRL01,CRL15,na03,CRL07,CRL08,CRL09,CRL10,CRL17,CRL12,CRL13,CRL03,CRL04,CRL69"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strSalesST15 = GetST15(RsTemp.Fields("CRL03"), strSalesDeptName, , strSalesST06)
      oContext2 = "本所案號： " + RsTemp.Fields("案號") + vbCrLf _
                + "案件名稱： " + RsTemp.Fields("案件名稱") + vbCrLf _
                + "申請國家：" + RsTemp.Fields("CRL15") + " " + RsTemp.Fields("申請國家") + vbCrLf _
                + "收文日： " + strCP05 + vbCrLf _
                + "案件性質： " + RsTemp.Fields("案件性質") + vbCrLf _
                + "本所期限：" & RsTemp.Fields("本所期限") + vbCrLf _
                + "法定期限：" & RsTemp.Fields("法定期限") + vbCrLf _
                + "智權人員　：" & GetStaffName(RsTemp.Fields("CRL03")) + vbCrLf _
                + "業務區　：" & strSalesDeptName + vbCrLf _
                + "費用　　：" & Format(RsTemp.Fields("總費用"), "##,##0") + vbCrLf _
                + "呈主管簽核：" & RsTemp.Fields("CRL69") & IIf(strCRA26 = "Y", "有對造;", IIf(strCRA27 = "Y", "有跨所;", ""))
   End If
   PUB_MailGetCRLCaseData = oContext2
End Function

'Add By Sindy 2022/10/17 取得案件性質的分案狀況代碼
Public Function PUB_GetCPM35(strCRL01 As String, strCP01 As String) As String
Dim strCRL06 As String, strCRL07 As String, strCP10 As String

   PUB_GetCPM35 = ""
   strSql = "SELECT CPM35,CRL06,CRL07,GetCRCaseNmFee(CRL01,'3') 案件性質代碼 FROM ConsultRecCMP,CasePropertyMap,ConsultRecordList" & _
            " WHERE CRC01='" & strCRL01 & "' and CRC01=CRL01" & _
            " and '" & strCP01 & "'=CPM01(+) and CRC03=CPM02(+)" & _
            " and CPM35='1'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCRL06 = "" & RsTemp.Fields("CRL06")
      strCRL07 = "" & RsTemp.Fields("CRL07")
      strCP10 = "" & RsTemp.Fields("案件性質代碼")
      
      If strCRL07 = "CFP" And (InStr(strCP10, "804") > 0 Or InStr(strCP10, "906") > 0) Then
         '(804)舉發答辯 (906)異同分析,中間接進來才須查詢，原為本所案件則不需補資料或進行查詢
         If strCRL06 = "Y" Then
            '1.先補文件再呈分案主管
            PUB_GetCPM35 = "1"
         Else
            Exit Function
         End If
         
      'P"舊案"的案件性質:(804)舉發答辯 (906)異同分析
      ElseIf strCRL07 = "P" And strCRL06 = "" And (InStr(strCP10, "804") > 0 Or InStr(strCP10, "906") > 0) Then
         '1.先補文件再呈分案主管
         PUB_GetCPM35 = "1"
         
      ElseIf strCRL07 = "P" And strCRL06 = "Y" Then
         '收文時勾選"新案"的案件性質，且未與新申請案同時收文的案件性質,才先由程序人員補資料
         strSql = "SELECT CPM35 FROM ConsultRecCMP,CasePropertyMap" & _
                  " WHERE CRC01='" & strCRL01 & "'" & _
                  " and '" & strCP01 & "'=CPM01(+) and CRC03=CPM02(+)" & _
                  " and instr('" & NewCasePtyList & "', CRC03)>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 0 Then
            '1.先補文件再呈分案主管
            PUB_GetCPM35 = "1"
         Else
            Exit Function
         End If
      End If
      
   Else
      strSql = "SELECT CPM35 FROM ConsultRecCMP,CasePropertyMap" & _
               " WHERE CRC01='" & strCRL01 & "'" & _
               " and '" & strCP01 & "'=CPM01(+) and CRC03=CPM02(+)" & _
               " group by CPM35" & _
               " order by CPM35 desc" 'null排在前面
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         If "" & RsTemp.Fields("CPM35") = "" Then
            '有,無未設定的
         Else
            '2.程序承辦不需經由主管分案
            '3.可能程序或工程師承辦
            PUB_GetCPM35 = RsTemp.Fields("CPM35")
         End If
      End If
   End If
End Function

'E-Mail簽核通知內容
'Modify By Sindy 2022/10/12 + Optional ByVal strF0308 As String = "" : 下一處理人員
'strSendText: 要寄出的內文
Public Function GetEMailContent_Flow(ByVal strF0301 As String, ByRef strSubject As String, _
         Optional ByVal strSendKind As String = "", Optional ByVal strSendText As String = "", _
         Optional ByVal strF0308 As String = "") As String
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim strF0316 As String
Dim strCTB As String, strWhrCP As String, strWhrBas As String, strWhrNp As String 'Add by Amy 2025/05/07
Dim strNation As String 'Add By Sindy 2025/6/4 申請國家
   
   GetEMailContent_Flow = ""
   'Modify By Sindy 2022/9/22 + 接洽單
   'Modify By Sindy 2024/2/21 + 結案單的下一程序資料檔
   'Modify by Amy 2025/06/02 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中,並增加以本所案號串進度(避免「無期限」結案退回會錯)
   If strSrvDate(1) >= FCP結案單電子化啟用日 Then
      strCTB = ",CloseCaseMain"
      strWhrBas = "And F0301=CCM01(+) and length(CCM02)<>9 and CP01(+)=SUBSTR(CCM02, 1, length(CCM02) - 9) and CP02(+)=SUBSTR(CCM02, length(CCM02)- 8, 6) " & _
                              "And CP03(+)=SubStr(CCM02, length(CCM02)- 2,1) And CP04(+)=SubStr(CCM02, length(CCM02)- 1,length(CCM02)) And CP01 is not null "
      strWhrCP = "And F0301=CCM01(+) and length(CCM02)=9 and CCM02=CP09(+) And CP09 is not null "
      strWhrNp = "And F0301=CCM01(+) and CCM02=NP01(+) and CCM03=NP22(+) And NP01 is not null "
   Else
      strWhrBas = "and length(F0303)<>9 and CP01(+)=SUBSTR(F0303, 1, length(F0303) - 9) and CP02(+)=SUBSTR(F0303, length(F0303)- 8, 6) " & _
                              "And CP03(+)=SubStr(F0303, length(F0303)- 2,1) And CP04(+)=SubStr(F0303, length(F0303)- 1,length(F0303)) And CP01 is not null "
      strWhrCP = "And length(F0303)=9 and F0303=CP09(+) And CP09 is not null "
      strWhrNp = "and F0303=NP01(+) and F0304=NP22(+) And NP01 is not null "
   End If
   '結案單
   strSql = "SELECT Distinct FLOW003.*,CP01,CP02,CP03,CP04,2 as sort FROM FLOW003,CaseProgress" & strCTB & " WHERE F0301='" & strF0301 & "' and length(F0301)=8 " & strWhrBas & _
            " Union " & _
            "SELECT FLOW003.*,CP01,CP02,CP03,CP04,2 as sort FROM FLOW003,CaseProgress" & strCTB & " WHERE F0301='" & strF0301 & "' and length(F0301)=8 " & strWhrCP & _
            " Union " & _
            "SELECT FLOW003.*,NP02,NP03,NP04,NP05,1 as sort FROM FLOW003,NextProgress" & strCTB & " WHERE F0301='" & strF0301 & "' and length(F0301)=8 " & strWhrNp
   '接洽單
   strSql = strSql & " Union " & _
            "SELECT FLOW003.*,CRL07,CRL08,CRL09,CRL10,0 as sort FROM FLOW003,CONSULTRECORDLIST WHERE F0301='" & strF0301 & "' and length(F0301)=10 and F0301=CRL01(+) " & _
            " order by sort asc"
   'end 2025/06/02
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      strCP01 = "" & rsTmp.Fields("CP01")
      strCP02 = "" & rsTmp.Fields("CP02")
      strCP03 = "" & rsTmp.Fields("CP03")
      strCP04 = "" & rsTmp.Fields("CP04")
      strNation = GetPrjNation1(strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04) 'Add By Sindy 2025/6/4 申請國家
      If strCP01 = "" Then
         'Modify by Amy 2025/05/23 +FC結案單,將Flow003中屬於結案單資料者拆出
         If strSrvDate(1) >= FCP結案單電子化啟用日 Then
            strCP01 = Left(rsTmp.Fields("CCM02"), Len(rsTmp.Fields("CCM02")) - 9)
            strCP02 = Mid(rsTmp.Fields("CCM02"), Len(strCP01) + 1, 6)
            strCP03 = Mid(rsTmp.Fields("CCM02"), Len(strCP01) + 7, 1)
            strCP04 = Right(rsTmp.Fields("CCM02"), 2)
         Else
            strCP01 = Left(rsTmp.Fields("F0303"), Len(rsTmp.Fields("F0303")) - 9)
            strCP02 = Mid(rsTmp.Fields("F0303"), Len(strCP01) + 1, 6)
            strCP03 = Mid(rsTmp.Fields("F0303"), Len(strCP01) + 7, 1)
            strCP04 = Right(rsTmp.Fields("F0303"), 2)
         End If
         'end 2025/05/23
      End If
      strF0316 = rsTmp.Fields("F0316")
      If strF0308 = "" Then strF0308 = rsTmp.Fields("F0308") '下一處理人員
      If strSendKind = "" Then strSendKind = rsTmp.Fields("F0309")
      
      'Modified by Lydia 2019/08/08 銷案銷帳單=>銷案／銷帳單
      'Modify By Sindy 2022/10/4
      'Modify By Sindy 2025/6/4 整理程式:將接洽單和結案單的程式分開,方便維護
'**************************************************************************
'   接洽單
'**************************************************************************
      If rsTmp.Fields("F0302") = Flow_接洽單 Then
         'Modify By Sindy 2023/11/7 接洽單的智權人員是操作人員時,才掛人名
         If strF0316 = strUserNum Then
            GetEMailContent_Flow = GetPrjSalesNM(strF0316) & "(接洽單)"
         Else
            GetEMailContent_Flow = "(接洽單)"
         End If
         '2023/11/7 END
         Select Case strSendKind
            Case Flow_主管審核中, "A1": 'Modify By Sindy 2025/5/8 + , "A1"
               GetEMailContent_Flow = GetEMailContent_Flow & "電子簽核通知"
            Case Flow_退回:
               GetEMailContent_Flow = GetEMailContent_Flow & "電子簽核【退回】通知"
            Case Flow_已完成:
               GetEMailContent_Flow = GetEMailContent_Flow & "已完成"
            Case Flow_重送:
               GetEMailContent_Flow = GetEMailContent_Flow & "電子簽核【重送】通知"
            Case Else
               'Modify By Sindy 2022/9/22
               'A3.櫃檯人員
               'Modify By Sindy 2025/5/8
               'If strF0308 = "A3" And strSendKind = Flow_處理中 Then
               If strSendKind = "A3" Then
               '2025/5/8 END
                  GetEMailContent_Flow = GetEMailContent_Flow & "【新客戶建檔】通知"
               Else
               '2022/9/22 END
                  strExc(10) = ""
                  If Len(strSendKind) = 2 Then
                     strExc(0) = "select decode('" & strSendKind & "'," & ShowFlow表單狀態中文 & ",'" & strSendKind & "') from dual"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strExc(10) = RsTemp.Fields(0)
                     End If
                  End If
                  GetEMailContent_Flow = GetEMailContent_Flow & "電子簽核【" & IIf(strExc(10) <> "", strExc(10), strSendKind) & "】通知"
               End If
         End Select
         strSubject = GetEMailContent_Flow
         If strSendText <> "" Then GetEMailContent_Flow = strSendText & vbCrLf & vbCrLf & GetEMailContent_Flow 'Add By Sindy 2025/6/4
         
         'Modify By Sindy 2023/2/18 + Or strSendKind = Flow_智權補件
         If strSendKind = Flow_退回 Or strSendKind = Flow_智權補件 Then
            GetEMailContent_Flow = GetEMailContent_Flow & vbCrLf & vbCrLf & "請至案件管理系統的 一般作業->案件表單查詢及簽核->目前表單 項目中，進行表單處理。"
         Else
            'Modify By Sindy 2022/9/22
            'A3.櫃檯人員
            'Modify By Sindy 2025/5/8
            'If strF0308 = "A3" And strSendKind = Flow_處理中 Then
            If strSendKind = "A3" Then
            '2025/5/8 END
               GetEMailContent_Flow = GetEMailContent_Flow & vbCrLf & vbCrLf & "請至收文系統的 收文->接洽記錄單->待建檔區，進行處理。"
            Else
               GetEMailContent_Flow = GetEMailContent_Flow & vbCrLf & vbCrLf & "請至案件管理系統的 一般作業->案件表單查詢及簽核->簽核作業 項目中，進行表單簽核處理。"
            End If
         End If
'**************************************************************************
'   結案單
'**************************************************************************
      Else
      '2022/10/4 END
         GetEMailContent_Flow = GetPrjSalesNM(strF0316) & "(" & _
                           IIf(rsTmp.Fields("F0302") = "1", "結案單", IIf(rsTmp.Fields("F0302") = "2", "銷案／銷帳單", "")) & _
                           strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04 & _
                           ")"
                           
         Select Case strSendKind
            Case Flow_主管審核中:
               GetEMailContent_Flow = GetEMailContent_Flow & "電子簽核通知"
            'Add By Sindy 2025/6/4
            Case Flow_處理中:
               GetEMailContent_Flow = GetEMailContent_Flow & "電子簽核【解除期限】通知"
            Case Flow_已完成, Flow_判發重送:
               GetEMailContent_Flow = GetEMailContent_Flow & "電子簽核【審核/補看】通知"
            '2025/6/4 END
            'Add by Amy 2019/11/27 外商案件判發退回發信通知承辦 +, Flow_判發退回
            Case Flow_退回, Flow_判發退回:
               GetEMailContent_Flow = GetEMailContent_Flow & "電子簽核【退回】通知"
            Case Flow_重送:
               GetEMailContent_Flow = GetEMailContent_Flow & "電子簽核【重送】通知"
            'end 2019/11/27
            Case Else
               strExc(10) = ""
               If Len(strSendKind) = 2 Then
                  strExc(0) = "select decode('" & strSendKind & "'," & ShowFlow表單狀態中文 & ",'" & strSendKind & "') from dual"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strExc(10) = RsTemp.Fields(0)
                  End If
               End If
               GetEMailContent_Flow = GetEMailContent_Flow & "電子簽核【" & IIf(strExc(10) <> "", strExc(10), strSendKind) & "】通知"
         End Select
         strSubject = GetEMailContent_Flow
         If strSendText <> "" Then GetEMailContent_Flow = strSendText & vbCrLf & vbCrLf & GetEMailContent_Flow 'Add By Sindy 2025/6/4
         
         'Add by Amy 2019/11/27
         'Modify By Sindy 2025/6/4 +Flow_處理中
         If strSendKind = Flow_判發退回 Or strSendKind = Flow_處理中 Then '程序人員
            'GetEMailContent_Flow = GetEMailContent_Flow & vbCrLf & vbCrLf & Replace(結案單外商CF操作路徑, "結案", "")
            If strCP01 = "CFT" Or strCP01 = "CFC" Or (strCP01 = "S" And strNation <> "000") Then
               GetEMailContent_Flow = GetEMailContent_Flow & vbCrLf & vbCrLf & "請至國外部商標系統->外商->CF資料處理->待處理區，進行表單處理。"
            ElseIf Left(PUB_GetST03(strF0308), 2) = "F1" Then
               GetEMailContent_Flow = GetEMailContent_Flow & vbCrLf & vbCrLf & "請至國外部商標系統->外商->FC資料處理->待處理區，進行表單處理。"
            ElseIf Left(PUB_GetST03(strF0308), 2) = "F2" Then
               GetEMailContent_Flow = GetEMailContent_Flow & vbCrLf & vbCrLf & "請至國外部專利及承辦人系統->外專->資料處理->待處理區，進行表單處理。"
            End If
         ElseIf strSendKind = Flow_已完成 Or strSendKind = Flow_判發重送 Then '補看人員
            GetEMailContent_Flow = GetEMailContent_Flow & vbCrLf & vbCrLf & "請至案件管理系統的 一般作業->案件表單查詢及簽核->專業部 審核/補看 作業檢核。"
         '2025/6/4 END
         'end 2019/11/27
         ElseIf strSendKind = Flow_退回 Then
            GetEMailContent_Flow = GetEMailContent_Flow & vbCrLf & vbCrLf & "請至案件管理系統的 一般作業->案件表單查詢及簽核->目前表單 項目中，進行表單處理。"
         Else
            GetEMailContent_Flow = GetEMailContent_Flow & vbCrLf & vbCrLf & "請至案件管理系統的 一般作業->案件表單查詢及簽核->簽核作業 項目中，進行表單簽核處理。"
         End If
      End If
      '2025/6/4 END
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Add By Sindy 2022/10/4 整批發通知信: 電子收文(或部分簽核)通知信
Public Function FlowBatchSendMail(strCUM06 As String) As Boolean
Dim strSubject As String, strContent As String, strCon As String
Dim strRestKind As String
Dim strTempCC As String
Dim strRestEmp As String, strNormalEmp As String
Dim ArrStr As Variant, jj As Integer, i As Integer
Dim strCUM01 As String, strCUM02 As String, strCUM04 As String, strCUM09 As String
Dim rsA As New ADODB.Recordset
Dim strCRL03 As String, strRestEmp_A As String
   
   FlowBatchSendMail = False
   
'CaseUseMemo結構:
   'cum01 = 身份類別
   'cum02 = 收受者
   'cum03 = 0
   'cum04 = 表單狀態
   'cum05 = 03.電子收文(或簽核)通知信
   'cum06 = 操作人員
   'CUM09 = 接洽單編號
   'Modify By Sindy 2025/5/7
   For i = 1 To 3 '4
      If i = 1 Then
         strCon = " and cum01 in('A1','A2')"
      ElseIf i = 2 Then
         strCon = " and cum01='A3'"
      ElseIf i = 3 Then
         strCon = " and cum01 in('CFT','ACS')"
      End If
      '2025/5/7 END
      strExc(0) = "select cum01,cum02,cum04,cum05,cum09 from CaseUseMemo" & _
                  " where cum05='03'" & _
                  " and cum06=" & CNULL(strCUM06) & strCon & _
                  " order by cum01"
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         rsA.MoveFirst
         strRestEmp = "": strNormalEmp = "": strRestEmp_A = ""
         Do While Not rsA.EOF
            strCUM01 = "" & rsA.Fields("cum01")
            strCUM02 = "" & rsA.Fields("cum02")
            strCUM04 = "" & rsA.Fields("cum04")
            strCUM09 = "" & rsA.Fields("cum09")
            'Add By Sindy 2023/6/2
            strCRL03 = ""
            If strCUM09 <> "" Then
               strExc(0) = "select CRL01,CRL03 from consultrecordlist" & _
                           " where CRL01='" & strCUM09 & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strCRL03 = RsTemp.Fields("CRL03") '接洽單智權人員
               End If
            End If
            '2023/6/2 END
            
            '因為有休假問題,所以有休假人員各自發信,其他人則一封
            strTempCC = GetCaseDutyAgent(strCUM02, "", False, strRestKind)
            'Add By Sindy 2023/6/2
            '接洽單本身的智權人員就是系統抓出來的職代時,Mail則改抓全發(職代全發)
            If strTempCC = strCRL03 And strCRL03 <> "" And strTempCC <> "" Then
               strRestEmp_A = strRestEmp_A & ";" & strCUM02
            '2023/6/2 END
            ElseIf strTempCC <> "" Then
               strRestEmp = strRestEmp & ";" & strCUM02
            Else
               strNormalEmp = strNormalEmp & ";" & strCUM02
            End If
            rsA.MoveNext
         Loop
         
         strSubject = "": strContent = ""
         'If strCUM01 = "CFT" Then
         If i = 3 Then
            strSubject = "(接洽單已收文)請進行案件分案"
            strContent = strSubject & ", 謝謝！"
         Else
'            If i = 2 Then 'A3
'               strContent = GetEMailContent_Flow(strCUM09, strSubject, strCUM04, , strCUM02)
'            Else
               'Modify By Sindy 2025/5/8 +strCUM01 E-Mail要通知的狀態
               'strContent = GetEMailContent_Flow(strCUM09, strSubject)
               strContent = GetEMailContent_Flow(strCUM09, strSubject, strCUM01)
'            End If
         End If
         '一起通知,減少操作人員等發信的時間
         If strNormalEmp <> "" Then
            strNormalEmp = Mid(strNormalEmp, 2)
            '含特殊職代
            PUB_SendMail strUserNum, strNormalEmp, "", strSubject, strContent, , , , , , , , , , , False, , True
         End If
         '有休假人員各自發信
         If strRestEmp <> "" Then
            strRestEmp = Mid(strRestEmp, 2)
            ArrStr = Split(strRestEmp, ";")
            For jj = 0 To UBound(ArrStr)
               '含特殊職代
               PUB_SendMail strUserNum, ArrStr(jj), "", strSubject, strContent, , , , , , , , , , , False, , True
            Next jj
         End If
         'Add By Sindy 2023/6/2 接洽單本身的智權人員就是系統抓出來的職代時,Mail則改抓全發(職代全發)
         If strRestEmp_A <> "" Then
            strRestEmp_A = Mid(strRestEmp_A, 2)
            ArrStr = Split(strRestEmp_A, ";")
            For jj = 0 To UBound(ArrStr)
               '含特殊職代
               PUB_SendMail strUserNum, ArrStr(jj), "", strSubject, strContent, , , , , , , , , , , False, , True, , , , , , , , , , "A"
            Next jj
         End If
         '2023/6/2 END
         '刪除記錄
         strExc(0) = "delete from CaseUseMemo" & _
                     " where cum05='03'" & _
                    " and cum06=" & CNULL(strCUM06) & strCon
         cnnConnection.Execute strExc(0)
      End If
   Next i
   
   FlowBatchSendMail = True
End Function

'讀取案件表單流程備註檔
'Modify by Amy 2022/10/07 +stWhere
Public Sub SetFlow004TextBox(ByRef txtTempBox As Object, strF0401 As String, Optional ByVal stWhere As String = "")
Dim rsTmp As New ADODB.Recordset
   
   txtTempBox.Text = ""
   'Modify by Amy 2022/10/07 +stWhere
   'Modify By Sindy 2023/12/29
   If strSrvDate(1) >= 新部門啟用日 Then
      strSql = "SELECT sqldateT(F0404),sqltime(substr('000000'||F0405,-6)),decode(F0406,'" & Flow_歸檔 & "','',nvl(A0922,ST02)),decode(F0406," & ShowFlow表單狀態中文 & "),F0407" & _
               " FROM Flow004,Staff,Acc090NEW" & _
               " WHERE F0401='" & strF0401 & "' and F0403=ST01(+) and F0403=A0921(+) " & stWhere & " order by F0402 asc"
   Else
   '2023/12/29 END
      strSql = "SELECT sqldateT(F0404),sqltime(substr('000000'||F0405,-6)),decode(F0406,'" & Flow_歸檔 & "','',nvl(A0902,ST02)),decode(F0406," & ShowFlow表單狀態中文 & "),F0407 FROM Flow004,Staff,Acc090 " & _
               "WHERE F0401='" & strF0401 & "' and F0403=ST01(+) and F0403=A0901(+) " & stWhere & " order by F0402 asc"
   End If
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With rsTmp
         .MoveFirst
         Do While Not .EOF
            If txtTempBox.Text <> "" Then
               txtTempBox.Text = txtTempBox.Text & vbCrLf
               txtTempBox.Text = txtTempBox.Text & "-----------------------------------------------------" & vbCrLf 'Add By Sindy 2022/12/6
            End If
            txtTempBox.Text = txtTempBox.Text & .Fields(0) & " " & _
                                                .Fields(1) & " " & _
                                                IIf(IsNull(.Fields(2)), "", .Fields(2) & " ") & _
                                                IIf(IsNull(.Fields(3)), "", .Fields(3) & " ") & _
                                                IIf(Not IsNull(.Fields(4)) And .Fields(4) > "", IIf(IsNull(.Fields(2)) And IsNull(.Fields(3)), "", "："), "") & _
                                                .Fields(4)
            .MoveNext
         Loop
      End With
   End If
End Sub

'檢查結案單
'Modify By Sindy 2023/6/26 + , Optional ByVal bolIsFMP As Boolean = False : 是否為FMP;True.是
Public Function CheckFlowCloseOk(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, _
   strCP10 As String, Optional ByVal bolIsFMP As Boolean = False) As Boolean
   
   CheckFlowCloseOk = False
   '大陸的標準專利紀錄請求除外
   If strCP01 = "P" And strCP10 = "110" Then
      CheckFlowCloseOk = True
      Exit Function
   End If
   'Modify By Sindy 2020/1/8 Mark
'   If strCP01 = "P" Or strCP01 = "PS" Or strCP01 = "CFP" Or strCP01 = "CPS" Then
   '2020/1/8 END
      'Modify By Sindy 2020/2/5 解除期限FCP-059874;有收文未發文的進度,若CP12為FXX時
      '                         則改以可選擇是否繼續的方式提醒 + order by cp66 desc,cp67 desc
      strExc(0) = "select cp09,cp12 from caseprogress where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & Right("0" & strCP03, 1) & "' and cp04='" & Right("00" & strCP04, 2) & "' and cp27 is null and cp57 is null and substr(cp09,1,1) in('B','C') order by cp66 desc,cp67 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '有B,C類收文未發文
         'Modify By Sindy 2020/2/5
         'If Mid(RsTemp.Fields("cp12"), 1, 1) = "F" Then 'Removed by Morgan 2023/12/4 改都提醒但可繼續
         
            If MsgBox("有B,C類收文未發文，是否確定要解除期限？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
               CheckFlowCloseOk = True
            End If
            
         'Removed by Morgan 2023/12/4 改都提醒但可繼續
         'Else
         ''2020/2/5 END
         '   MsgBox "有B,C類收文未發文不可解除期限！"
         'End If
         'end 2023/12/4
         
      'Add By Sindy 2020/1/8 增加A類的提醒
      Else
         CheckFlowCloseOk = True
         '彈提醒訊息
         strExc(0) = "select cp09 from caseprogress where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & Right("0" & strCP03, 1) & "' and cp04='" & Right("00" & strCP04, 2) & "' and cp27 is null and cp57 is null and substr(cp09,1,1) in('A','D')"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            '有A,D類收文未發文
            MsgBox "尚有A,D類收文未發文資料"
         End If
      '2020/1/8 END
      End If
      
      'Add By Sindy 2023/6/26 操作FMP案解除期限，當程序點選下一程序陳述意見時
      '彈視窗1詢問，Y有 N沒有
      '點選Y ' 彈視窗2「請通知承辦: 須在期限內提出陳述意見，待日後接獲核駁函方可提出分割，若要分割不可解除期限」，確定後回到前畫面。
      '點選N ' 完成不續辦/閉卷。
      If bolIsFMP = True And strCP10 = "205" Then '205.陳述意見
         If MsgBox("聯絡單是否有註記管制分割？" & vbCrLf & vbCrLf & "按否，完成不續辦/閉卷。", vbYesNo + vbDefaultButton1, "FMP案解除期限-陳述意見") = vbYes Then
            If MsgBox("「請通知承辦：須在期限內提出陳述意見，待日後接獲核駁函方可提出分割，若要分割不可解除期限」。" & vbCrLf & vbCrLf & _
               "按否，完成不續辦/閉卷。", vbYesNo + vbDefaultButton1, "FMP案解除期限-陳述意見") = vbYes Then
               CheckFlowCloseOk = False
            Else
               CheckFlowCloseOk = True
            End If
         Else
            CheckFlowCloseOk = True
         End If
         Exit Function
      End If
      '2023/6/26 END
'   Else
'      CheckFlowCloseOk = True
'   End If
End Function

'該案的智權人員
Public Function ShowCurrCP13(ByVal strCP01 As String, ByVal strCP02 As String, ByVal strCP03 As String, ByVal strCP04 As String, _
                             ByVal strNation As String, Optional ByRef strCP12 As String) As String
   
   'Modify By Sindy 2021/6/29 可以改用 PUB_GetAKindSalesNo 統一抓法了
'   If strCP01 = "FCP" Or strCP01 = "FG" Then
'      ShowCurrCP13 = PUB_GetFCPSalesNo(strCP01, strCP02, strCP03, strCP04)
'   ElseIf strCP01 = "FCL" Or strCP01 = "LIN" Then
'      ShowCurrCP13 = PUB_GetFCLSalesNo(strCP01, strCP02, strCP03, strCP04)
'   ElseIf strCP01 = "FCT" Then
'      ShowCurrCP13 = PUB_GetFCTSalesNo(strCP01, strCP02, strCP03, strCP04)
'   ElseIf strCP01 = "S" Then
'     If strNation = "000" Then
'        ShowCurrCP13 = PUB_GetFCTSalesNo(strCP01, strCP02, strCP03, strCP04)
'     Else
'        ShowCurrCP13 = PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04)
'     End If
'   Else
'      ShowCurrCP13 = PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04)
'   End If
   ShowCurrCP13 = PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04)
   '2021/6/29 END
   
   If ShowCurrCP13 <> "" Then
      strCP12 = GetSalesArea(ShowCurrCP13)
   End If
End Function

'Add by Amy 2020/12/10 結案單確認(從frm210149_1 cmdok中搬過來,T延展結案待處理區也用)
'Modify by Amy 2025/01/23 +回傳stNP11
Public Function ChkNotCloseStep(ByRef intStep As Integer, stCP01 As String, stCP02 As String, stCP03 As String, stCP04 As String, stNP01 As String, stNP22 As String, stNP07 As String, _
  Optional ByRef stNP11 As String) As Boolean
    Dim RsQ As New ADODB.Recordset
    Dim strQ As String, intQ As Integer
    
    ChkNotCloseStep = False
    
    '檢查下一程序是否續辦
    intStep = 1
    'Modify by Amy 2025/01/23 +NP11
    strQ = "Select np01,np11 From NextProgress" & _
                " where np02='" & stCP01 & "' and np03='" & stCP02 & "'" & _
                " and np04='" & stCP03 & "' and np05='" & stCP04 & "'" & _
                " and np06 is not null and np01='" & stNP01 & "' and np22=" & stNP22
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 1 Then
        ChkNotCloseStep = True
        stNP11 = "" & RsQ.Fields("np11") 'Add by Amy 2025/01/23
        RsQ.Close
        Exit Function
    End If
    If RsQ.State <> adStateClosed Then RsQ.Close
    
    '檢查收文未發文
    intStep = intStep + 1
    If CheckFlowCloseOk(stCP01, stCP02, stCP03, stCP04, stNP07) = False Then
        ChkNotCloseStep = True
    End If
    Set RsQ = Nothing
End Function

'Add By Sindy 2017/6/19 結案單恢復解除期限
Public Sub PUB_CloseRestoreLimit(m_DelCP09 As String)
Dim rsTmp As New ADODB.Recordset
Dim strDelCP05 As String, strUpdCP140 As String
Dim strUpdNP22 As String, strUpdNP06 As String
Dim strIsClose As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim strCCD08 As String 'Add by Amy 2025/08/08
   
   If Trim(m_DelCP09) = "" Then Exit Sub
   
   strSql = "select cp01,cp02,cp03,cp04,cp05,cp09,cp140,np22,np06 from caseprogress,nextprogress" & _
            " where cp09='" & m_DelCP09 & "'" & _
            " and cp01=np02(+) and cp02=np03(+) and cp03=np04(+) and cp04=np05(+) and cp09=np24(+)"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      strDelCP05 = rsTmp.Fields("cp05")
      strUpdCP140 = "" & rsTmp.Fields("cp140")
      strUpdNP22 = "" & rsTmp.Fields("np22")
      strUpdNP06 = "" & rsTmp.Fields("np06")
      strCP01 = rsTmp.Fields("cp01")
      strCP02 = rsTmp.Fields("cp02")
      strCP03 = rsTmp.Fields("cp03")
      strCP04 = rsTmp.Fields("cp04")
   End If
   rsTmp.Close
   If Len(strUpdCP140) <> 8 Then Exit Sub '非電子結案單已結案的狀況下,離開,不處理
   
   If m_DelCP09 <> "" Then
      strSql = "delete from caseprogress where cp09='" & m_DelCP09 & "'"
      Pub_SeekTbLog strSql '記錄Log
      cnnConnection.Execute strSql
      
      'Add By Sindy 2019/8/1 要先刪.menu才更新收文號為本所案號
      strSql = "DELETE FROM casepaperpdf WHERE cpp01='" & m_DelCP09 & "' AND upper(substr(cpp02,-5))=upper('.menu')"
      Pub_SeekTbLog strSql '記錄Log
      cnnConnection.Execute strSql
      'Add By Sindy 2023/9/13 '先刪除非結案單的附件
      PUB_DelFtpFile2 m_DelCP09, " and (CPP11<>'" & strUpdCP140 & "' or CPP11 is null)"
      strSql = "DELETE FROM casepaperpdf WHERE CPP01='" & m_DelCP09 & "' and (CPP11<>'" & strUpdCP140 & "' or CPP11 is null)"
      Pub_SeekTbLog strSql '記錄Log
      cnnConnection.Execute strSql, intI
      '2023/9/13 END
      '保留結案單的附件
      strSql = "UPDATE casepaperpdf SET cpp01='" & strCP01 & strCP02 & strCP03 & strCP04 & "',cpp10='U' WHERE cpp01='" & m_DelCP09 & "'"
      Pub_SeekTbLog strSql '記錄Log
      cnnConnection.Execute strSql
      '2019/8/1 END
   End If
   If strUpdNP22 <> "" Then
      strSql = "update nextprogress set NP11=null,NP12=null,np24='" & strUpdCP140 & "'"
      If strUpdNP06 <> "Y" Then '已收文了
         strSql = strSql & ",NP06=null"
      End If
      strSql = strSql & " where np22='" & strUpdNP22 & "'" & _
                        " and np02='" & strCP01 & "' and np03='" & strCP02 & "' and np04='" & strCP03 & "' and np05='" & strCP04 & "'"
      Pub_SeekTbLog strSql '記錄Log
      cnnConnection.Execute strSql
   End If
   
   '與閉卷相同的所有取消收文日拿掉
   If Val(strDelCP05) > 0 Then
      strSql = "UPDATE CASEPROGRESS SET CP26=null,CP57=null,CP58=null" & _
               " WHERE CP01='" & strCP01 & "' AND CP02='" & strCP02 & "' AND CP03='" & strCP03 & "' AND CP04='" & strCP04 & "'" & _
               " AND CP57=" & strDelCP05 & " AND CP27 IS NULL"
      Pub_SeekTbLog strSql '記錄Log
      cnnConnection.Execute strSql
   End If
   
   'Modify by Amy 2018/08/07 非P 案結案電子化,加入其他基本檔
   strSql = "select pa57 from patent where pa01='" & strCP01 & "' and pa02='" & strCP02 & "' and pa03='" & strCP03 & "' and pa04='" & strCP04 & "' " & _
   "Union select tm29 from TradeMark where tm01='" & strCP01 & "' and tm02='" & strCP02 & "' and tm03='" & strCP03 & "' and tm04='" & strCP04 & "' " & _
   "Union select lc08 from LawCase where lc01='" & strCP01 & "' and lc02='" & strCP02 & "' and lc03='" & strCP03 & "' and lc04='" & strCP04 & "' " & _
   "Union select hc09 from HireCase where hc01='" & strCP01 & "' and hc02='" & strCP02 & "' and hc03='" & strCP03 & "' and hc04='" & strCP04 & "' " & _
   "Union select sp15 from ServicePractice where sp01='" & strCP01 & "' and sp02='" & strCP02 & "' and sp03='" & strCP03 & "' and sp04='" & strCP04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      strIsClose = "" & rsTmp.Fields("pa57")
   End If
   rsTmp.Close
   '閉卷
   If strIsClose = "Y" Then
      'Modify By Sindy 2019/8/1 不需詢問
'      If MsgBox("要恢復閉卷嗎？", vbExclamation + vbYesNo + vbDefaultButton2, "重要訊息！") = vbNo Then
'         Exit Sub
'      Else
      '2019/8/1 END
         Select Case strCP01
         Case "P", "CFP", "FCP":
            strSql = "UPDATE PATENT SET PA57=null,PA58=null,PA59=null" & _
                        " WHERE PA01='" & strCP01 & "' AND PA02='" & strCP02 & "' AND PA03='" & strCP03 & "' AND PA04='" & strCP04 & "'"
         Case "T", "TF", "CFT", "FCT":
            strSql = "UPDATE TradeMark set tm29=null,tm30=null,tm31=null" & _
                        " where tm01='" & strCP01 & "' and tm02='" & strCP02 & "' and tm03='" & strCP03 & "' and tm04='" & strCP04 & "'"
         Case "L", "CFL", "FCL", "LIN":
            strSql = "UPDATE LawCase set lc09=null,lc10=null,lc11=null" & _
                         " where lc01='" & strCP01 & "' and lc02='" & strCP02 & "' and lc03='" & strCP03 & "' and lc04='" & strCP04 & "'"
         Case "LA":
            strSql = "UPDATE HireCase set hc08=null,hc09=null,hc10=null" & _
                         " where hc01='" & strCP01 & "' and hc02='" & strCP02 & "' and hc03='" & strCP03 & "' and hc04='" & strCP04 & "'"
         Case Else:
            strSql = "UPDATE ServicePractice set sp15=null,sp16=null,sp17=null" & _
                        " where sp01='" & strCP01 & "' and sp02='" & strCP02 & "' and sp03='" & strCP03 & "' and sp04='" & strCP04 & "'"
         End Select
         Pub_SeekTbLog strSql '記錄Log
         cnnConnection.Execute strSql
'      End If
   End If
   'end 2018/08/07
   'Add by Amy 2025/08/08 外專結案單有勾「未付帳款」 及有輸「 管制催款日」將其行事曆刪除
   If strUpdCP140 <> "" Then
      If ChkCCD03(1, "PUB_CloseRestoreLimit", strUpdCP140, , strCCD08) = True Then
         strSql = "Delete Staff_Calendar Where SC20='" & strUpdCP140 & "'"
         Pub_SeekTbLog strSql '記錄Log
         cnnConnection.Execute strSql
         MsgBox "追蹤欠款 管制催款日[" & strCCD08 & "]之行事曆已刪除"
      End If
   End If
   'end 2025/08/08
   Set rsTmp = Nothing
End Sub

'Add by Amy 2021/06/24取得案件表單主檔資料
'strNo:案件表單編號
Public Function GetFlow003Data(strNo As String, Optional strWhere As String = "", Optional ByVal strField As String = "") As String
    Dim RsQ As ADODB.Recordset, intQ As Integer, stSQL As String
    
    GetFlow003Data = ""
    If strField = "" Then strField = "*"
    stSQL = "Select " & strField & " From Flow003 Where F0301='" & strNo & "' " & strWhere
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
    If intQ = 1 Then
        GetFlow003Data = "" & RsQ.Fields(0)
    End If
      
    Set RsQ = Nothing
End Function

'Add By Sindy 2015/1/30
'Modify by Amy 2025/05/29 +strF0316 回傳 智權人員
'Modify by Amy 2025/08/28 +intFCState
'使用的作業:
'　　表單簽核狀況查詢:刪除
'　　案件進度:刪除
Public Sub PUB_CloseFlowDataDel(strF0301 As String, m_CP01 As String, m_CP02 As String, m_CP03 As String, m_CP04 As String, _
                              Optional strCP09 As String = "", Optional ByRef strF0316 As String, Optional ByVal intFCState As Integer)
Dim rsTmp As New ADODB.Recordset
Dim strWhr As String 'Add by Amy 2025/05/13
Dim strTmp As String, strIR01 As String, strIR03 As String, intRec As Integer  'Add by Amy 2025/05/29
Dim strCCD08 As String 'Add by Amy 2025/08/11
   
   'Modify By Sindy 2015/5/11 Mark,不控管編號長度,因若接洽單收文會是10碼,但刪除進度時一樣需要更新NP24=null,刪除卷宗區
   '*******************************************************************************************************************
   '                          注意:案件表單編號和接洽單收文編號不可一樣長度
   '*******************************************************************************************************************
   'If Len(strF0301) <> 8 Then Exit Sub '長度8碼才是案件表單編號
   
   If strF0301 <> "" And m_CP01 <> "" And m_CP02 <> "" And m_CP03 <> "" And m_CP04 <> "" Then
      'Add By Sindy 2015/12/18
      If Len(strF0301) = 8 Then '電子結案單
      '2015/12/18 END
         'Add by Amy 2025/05/29 +FC信件沖銷還原
         'Modify by Amy 2025/07/08 +CFP
         'Modif by Amy 2025/08/28 改intFCState判斷,因外商結案可能使用新與舊結案單
         'If strSrvDate(1) >= FCP結案單電子化啟用日 And (m_CP01 = "FCP" Or m_CP01 = "P" Or m_CP01 = "FG" Or m_CP01 = "CFP") Then
         If intFCState = 1 Or intFCState = 2 Then
            strF0316 = ""
            strWhr = "CCM01='" & strF0301 & "' And F0301(+)=CCM01 And F0302(+)='1' Group by CCM17||';'||F0316"
            strTmp = Pub_GetField("CloseCaseMain,Flow003", strWhr, "CCM17||';'||F0316||';'||count(*)", True)
            intRec = Val(Mid(strTmp, InStrRev(strTmp, ";") + 1))
            strF0316 = Mid(strTmp, InStr(strTmp, ";") + 1, 5)
            '有信件編號,需[還原]信件沖銷
            'Modif by Amy 2025/07/08 拿掉 And intRec = 1,原:有信件編號且只有１筆,需[還原]信件沖銷,改有信件編號就[還原]信件沖銷,即使結多案-Sindy 與 Anny 討論後
            '  主管代為操作結案,會於系統收件區以下拉式選其智權人員操作,故都還原其智權人即可-Sindy
            If strTmp <> "" And Left(strTmp, 1) <> ";" Then
               strIR01 = Mid(strTmp, 1, 8)
               strIR03 = Mid(strTmp, 16, 5)
               strSql = "UPDATE InputRecord SET" & _
                              " ir08=0,ir09=null,ir10=null,ir16=null,ir17=0,ir18=null,ir19=null,ir22=null" & _
                              " WHERE ir01=" & strIR01 & " And ir03='" & strIR03 & "' And ir04='" & strF0316 & "' And ir08>0"
               Pub_SeekTbLog strSql '記錄Log
               cnnConnection.Execute strSql
               
               strSql = "UPDATE ipdeptInput SET" & _
                              " ii27=null,ii16=0" & _
                              " WHERE ii01=" & strIR01 & " And ii03='" & strIR03 & "' And ii16>0"
                  Pub_SeekTbLog strSql '記錄Log
                  cnnConnection.Execute strSql
            End If
            strWhr = ""
         End If
         'Add by Amy 2025/05/15 +FC結案單,將Flow003中屬於結案單資料者拆至結案單主檔中
         If strSrvDate(1) >= FCP結案單電子化啟用日 Then
            '結案單主檔
            strSql = "Delete From CloseCaseMain Where CCM01='" & strF0301 & "'"
            Pub_SeekTbLog strSql '記錄Log
            cnnConnection.Execute strSql
              'Add by Amy 2025/08/11 外專結案單有勾「未付帳款」 及有輸「 管制催款日」將其行事曆刪除
            If ChkCCD03(1, "PUB_CloseRestoreLimit", strF0301, , strCCD08) = True Then
               strSql = "Delete Staff_Calendar Where SC20='" & strF0301 & "'"
               Pub_SeekTbLog strSql '記錄Log
               cnnConnection.Execute strSql
               MsgBox "追蹤欠款 管制催款日[" & strCCD08 & "]之行事曆已刪除"
            End If
            'end 2025/08/11
            '結案單明細
            strSql = "Select * From CloseCaseDetail Where CCD01='" & strF0301 & "' Order by CCD02,CCD03"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               rsTmp.MoveFirst
               Do While Not rsTmp.EOF
                  strSql = "Delete From CloseCaseDetail Where CCD01='" & strF0301 & "' " & _
                                    "And CCD02='" & rsTmp.Fields("CCD02") & "' And CCD03='" & rsTmp.Fields("CCD03") & "' "
                  Pub_SeekTbLog strSql '記錄Log
                  cnnConnection.Execute strSql
                  rsTmp.MoveNext
               Loop
            End If
            rsTmp.Close
         End If
         '流程主檔
         strSql = "DELETE FROM Flow003 WHERE F0301='" & strF0301 & "'"
         Pub_SeekTbLog strSql '記錄Log
         cnnConnection.Execute strSql
      
         '簽核檔
         strSql = "SELECT * FROM Flow002 WHERE F0201='" & strF0301 & "' order by F0202,F0203 asc"
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            With rsTmp
               .MoveFirst
               Do While Not .EOF
                  strSql = "DELETE FROM Flow002 WHERE F0201='" & strF0301 & "' and F0202='" & rsTmp.Fields("F0202") & "' and F0203=" & rsTmp.Fields("F0203")
                  Pub_SeekTbLog strSql '記錄Log
                  cnnConnection.Execute strSql
                  .MoveNext
               Loop
            End With
         End If
         rsTmp.Close
      
         '流程備註檔
         strSql = "SELECT * FROM Flow004 WHERE F0401='" & strF0301 & "' order by F0402 asc "
         rsTmp.CursorLocation = adUseClient
         rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If rsTmp.RecordCount > 0 Then
            With rsTmp
               .MoveFirst
               Do While Not .EOF
                  strSql = "DELETE FROM Flow004 WHERE F0401='" & strF0301 & "' and F0402=" & rsTmp.Fields("F0402")
                  Pub_SeekTbLog strSql '記錄Log
                  cnnConnection.Execute strSql
                  .MoveNext
               Loop
            End With
         End If
         rsTmp.Close
      End If
      
      '無總收文號時,案件表單尚在簽核中程序未處理
      If strCP09 = "" Then
         '更新下一程序
         strSql = "Update NextProgress Set NP24=null WHERE NP24='" & strF0301 & "'"
         Pub_SeekTbLog strSql '記錄Log
         cnnConnection.Execute strSql
         
         'Modify by Amy 2025/05/15 結案單會有單號,避免刪錯,以案號+結案單號刪(因FC結案單要刪.MSG)
         'PUB_DelFtpFile2 m_CP01 & m_CP02 & m_CP03 & m_CP04, " and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))>0 and CPP10='U'"  'Added by Morgan 2015/4/2 檔案改放 FTP,必須在DB資料刪除前執行
         strWhr = " And CPP11='" & strF0301 & "'"
         PUB_DelFtpFile2 m_CP01 & m_CP02 & m_CP03 & m_CP04, strWhr
      
         '刪除卷宗區暫存的回覆單附件
         'strSql = "DELETE FROM casepaperpdf WHERE CPP01='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "' and instr(upper(CPP02),upper('" & EMP_回覆單 & ".pdf'))>0 and CPP10='U'"
         strSql = "DELETE FROM casepaperpdf WHERE CPP01='" & m_CP01 & m_CP02 & m_CP03 & m_CP04 & "'" & strWhr
         'end 2025/05/15
         Pub_SeekTbLog strSql '記錄Log
         cnnConnection.Execute strSql
      Else
         '更新下一程序
         strSql = "Update NextProgress Set NP24=null WHERE NP24='" & strCP09 & "'"
         Pub_SeekTbLog strSql '記錄Log
         cnnConnection.Execute strSql
         
         PUB_DelFtpFile2 strCP09 'Added by Morgan 2015/4/2 檔案改放 FTP,必須在DB資料刪除前執行
         
         '刪除卷宗區資料
         strSql = "DELETE FROM casepaperpdf WHERE CPP01='" & strCP09 & "'"
         Pub_SeekTbLog strSql '記錄Log
         cnnConnection.Execute strSql
      End If
   End If
   
   Set rsTmp = Nothing
End Sub

'Added by Morgan 2018/8/30
'更新結案單目前表單狀態 09指示信判發中 -> 03已完成
Public Sub OrderLetterFlowStatusUpdate(pF0301 As String)
   Dim stSQL As String, intR As Integer
   
   '考慮EPC母案結案時同時會有子案指示信要寄送,需判斷所有
   stSQL = "update flow003 set f0309='03' where f0301='" & pF0301 & "' and f0309='09'" & _
      " and not exists(select * from caseprogress,appform where cp140=f0301 and af01(+)=cp09 and af11=0)"
   cnnConnection.Execute stSQL, intR
End Sub

'Add By Sindy 2022/9/21
'讀取ConsultRecordList接洽記錄單主檔
Public Function ClsPDReadCRLDatabase(ByRef CRL() As String) As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset, i As Integer

On Error GoTo ErrHand
   
   ReDim Preserve CRL(TF_CRL) As String
   
   strSql = "select * from ConsultRecordList where CRL01=" + CNULL(CRL(1))
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsRecordset.RecordCount > 0 Then
      With rsRecordset
         For i = 1 To TF_CRL
            If IsNull(.Fields("CRL" & Format(i, "00")).Value) Then
               CRL(i) = ""
            Else
               CRL(i) = .Fields("CRL" & Format(i, "00")).Value
            End If
         Next
      End With
      
      ClsPDReadCRLDatabase = True
   Else
      ShowMsg "找不到此接洽單之資料"
   End If
   rsRecordset.Close
   Exit Function
ErrHand:
   MsgBox Err.Description
End Function

'Add By Sindy 2022/9/21
'讀取ConsultRecApp接洽記錄單申請人資料
Public Function ClsPDReadCRADatabase(ByRef CRA() As String, intRow As Integer) As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset, i As Integer
Dim tf_CRA As Integer
   
On Error GoTo ErrHand
   
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from ConsultRecApp where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_CRA = AdoRecordSet3.Fields.Count
   CheckOC3
   
   strSql = "select * from ConsultRecApp where CRA01=" + CNULL(CRA(1)) & " and CRA02=" & intRow
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsRecordset.RecordCount > 0 Then
      With rsRecordset
         For i = 1 To tf_CRA
            If IsNull(.Fields("CRA" & Format(i, "00")).Value) Then
               CRA(i) = ""
            Else
               CRA(i) = .Fields("CRA" & Format(i, "00")).Value
            End If
         Next
      End With
      
      ClsPDReadCRADatabase = True
'   Else
'      ShowMsg "找不到此接洽單申請人之資料"
   End If
   rsRecordset.Close
   Exit Function
ErrHand:
   MsgBox Err.Description
End Function
'Added by Lydia 2020/05/20 法律所案源收文：P/T/FCP/FCT分案(限台灣案)，修改案源案件類型
'Modified by Lydia 2020/05/28 +後補案源bolAfter
Public Function PUB_UpdateCP10toPT(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pCP09 As String, ByVal oldCP10 As String, ByVal oldNaId As String, ByVal newCP10 As String, ByVal newNaId As String, _
                       ByVal pCP06 As String, ByVal pCP13 As String, ByVal PCU01 As String, ByVal bolAfter As Boolean) As String
'pCP01~pCP04、pCP09、pCP06、pCP13、pCu01: 分案之本所案號、收文號、本所期限、申請人1
'oldCP10+oldNaId、newCP10+newNaId: 修改前/後的案件性質+申請國家
'Memo by Lydia 2020/05/20
'1,2項：修改案源案件類型=>如果有變化需要對應3,4,5項PUB_UpdateLOS01
'3,4,5項：針對分案PUB_UpdateLOS01
Dim stOldType As String, stNewType As String
Dim strA1 As String, strB1 As String, strLOS15 As String
Dim intR As Integer
Dim rsA As New ADODB.Recordset

    stOldType = PUB_GetLOSkind(pCP01, oldCP10, oldNaId)
    stNewType = PUB_GetLOSkind(pCP01, newCP10, newNaId)
    
    If stOldType = stNewType Then
         'Added by Lydia 2020/05/28 填C類接洽單選不用法律所配合,在分案時選”需要法律所配合”：補案源
         If Left(stNewType, 1) = "C" And bolAfter = True Then
             '自動新增案源
             strLOS15 = AutoNo("LOS", 5, , True)
             strB1 = "insert into LawOfficeSource(LOS01,LOS02,LOS03,LOS04,LOS05,LOS10,LOS11,LOS12,LOS13,LOS15)" & _
                " values (null, '" & Mid(stNewType, 1, Len(stNewType) - 1) & "' ,null " & _
                ",'" & pCP13 & "','" & ChangeCustomerL(PCU01) & "','" & IIf(Left(stNewType, 1) = "B", strA1, "") & "'" & _
                ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss'),'" & strLOS15 & "')"
             cnnConnection.Execute strB1, intR
             strB1 = "update caseprogress set CP162='" & strLOS15 & "' where cp09='" & pCP09 & "' "
             cnnConnection.Execute strB1, intR
             '同一接洽單(案源)以第一筆分案的回答為準
             strB1 = Pub_GetSpecMan("C" & IIf(InStr(pCP01, "P") > 0, "P", "T"))
             strB1 = "update caseprogress set CP162='" & strLOS15 & "' where cp09 in (" & _
                        "select cp09 From caseprogress where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' " & _
                        "and cp05=(select cp05 from caseprogress where cp09='" & pCP09 & "') and instr('" & strB1 & "',cp10) > 0 and cp162 is null) "
             cnnConnection.Execute strB1, intR
         Else
         'end 2020/05/28
             Exit Function
         End If
    Else
         '1.若收文時非BC類但分案改為BC類的案件性質時，自動新增案源。另B類加自動收文TT-999999(736服務費)；
         '案件類型C類增’是否需要法律所配合’，預設Y。=> 畫面會顯示欄位,自動預設Y
         If stOldType = "" And (Left(stNewType, 1) = "B" Or Left(stNewType, 1) = "C") Then
             'B類加自動收文TT-999999(736服務費)
             If Left(stNewType, 1) = "B" Then
                strA1 = AutoNo("B", 6)
                strB1 = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp11,cp12,cp13,CP162)" & _
                   " values('TT','999999','0','00'," & strSrvDate(1) & "," & CNULL(TransDate(pCP06, 2), True) & ",'" & strA1 & "'" & _
                   ",'736','07','" & GetST15(pCP13) & "','" & pCP13 & "',null)"
                cnnConnection.Execute strB1, intR
             Else
                strA1 = pCP09
             End If
             '自動新增案源
             strLOS15 = AutoNo("LOS", 5, , True)
             strB1 = "insert into LawOfficeSource(LOS01,LOS02,LOS03,LOS04,LOS05,LOS10,LOS11,LOS12,LOS13,LOS15)" & _
                " values (null, '" & Mid(stNewType, 1, Len(stNewType) - 1) & "' ,null " & _
                ",'" & pCP13 & "','" & ChangeCustomerL(PCU01) & "','" & IIf(Left(stNewType, 1) = "B", strA1, "") & "'" & _
                ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss'),'" & strLOS15 & "')"
             cnnConnection.Execute strB1, intR
             strB1 = "update caseprogress set CP162='" & strLOS15 & "' where cp09='" & pCP09 & "' "
             cnnConnection.Execute strB1, intR
         End If
         '2.若收文時為BC類但分案改為非BC類案件性質時，將法律所案源檔的放棄日期、放棄人員填入、放棄原因：專業部分案改案件性質，
         '更新案源總收文號(LOS01)為本案之總收文號。另B類要將收據總收文號(TT總收文號抓進度) 上取消收文日及取消收文原因代號99，放棄原因存在進度備註” 放棄原因：專業部分案改案件性質”。
         If stNewType = "" And (Left(stOldType, 1) = "B" Or Left(stOldType, 1) = "C") Then
            '未分案
            strB1 = "SELECT '1' ord1, X.LOS01,X.LOS02, X.LOS10, X.LOS15 FROM LawOfficeSource X WHERE X.LOS15 IN ( " & _
                        "SELECT NVL(CP162,'N') PNO FROM CASEPROGRESS WHERE CP09='" & pCP09 & "' AND CP162 IS NOT NULL) " & _
                        "AND LOS07 IS NULL"
            '已分案
            strB1 = strB1 & " union SELECT '2' ord1, LOS01,LOS02,LOS10,LOS15 FROM LawOfficeSource WHERE LOS01='" & pCP09 & "' AND LOS07 IS NULL "
            strB1 = strB1 & " ORDER BY 1, 2, LOS15 "
            intR = 1
            Set rsA = ClsLawReadRstMsg(intR, strB1)
            If intR = 1 Then
                rsA.MoveFirst
                strB1 = "update LawOfficeSource set los01='" & pCP09 & "', los07=" & strSrvDate(1) & ", los08='" & strUserNum & "', los09='專業部分案改案件性質' where los15='" & rsA.Fields("los15") & "' and los07 is null "
                cnnConnection.Execute strB1, intR
                If "" & rsA.Fields("los10") <> "" And Left(stOldType, 1) = "B" Then
                    strB1 = "update caseprogress set cp57='" & strSrvDate(1) & "', cp58='99',cp64='放棄原因：專業部分案改案件性質;'||cp64 where cp09='" & rsA.Fields("los10") & "' and cp57 is null"
                    cnnConnection.Execute strB1, intR
                End If
            End If
            Set rsA = Nothing
         End If
    
    End If
End Function

'Add By Sindy 2011/8/25 檢查操作人員是否有待處理事項
Public Function ChkIsAbsenceMustPro() As String
   'A.出缺勤表單待簽核
   'Modify By Sindy 2013/1/23
   If strUserNum = "71011" Then '王副總時,請假/出差單要剔除請假日14天以後的待簽核資料
      strSql = "select * from abs010 Where B1003<>B1017 and B1017='" & strUserNum & "' And B1019 Is Null " & _
                                     "And (B1002 in('02') or ((B1002='01' or B1002='03') and B1004<=" & Val(CompDate(2, 14, strSrvDate(1))) & ")) "
   Else
   '2013/1/23 End
      strSql = "select * from abs010 Where B1003<>B1017 and B1017='" & strUserNum & "' And B1019 Is Null "
   End If
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ChkIsAbsenceMustPro = ChkIsAbsenceMustPro & "A,"
   End If
   'B.出缺勤表單待處理
   strSql = "select * from abs010 Where B1003='" & strUserNum & "' and (B1017='" & strUserNum & "' or B1017 is null) And B1019 Is Null "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ChkIsAbsenceMustPro = ChkIsAbsenceMustPro & "B,"
   End If
   'C.出缺勤統計待確認
   strSql = "select * from abs013 Where B1301='04' and B1303='" & strUserNum & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ChkIsAbsenceMustPro = ChkIsAbsenceMustPro & "C,"
   End If
   'D.個人資料明細待確認
   strSql = "select * from abs013 Where B1301='05' and B1303='" & strUserNum & "' "
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ChkIsAbsenceMustPro = ChkIsAbsenceMustPro & "D,"
   End If
   'Add By Sindy 2014/2/18
   'E.打卡異常個人處理待確認
   strSql = "select * from abs014 where b1401='" & strUserNum & "' and b1405 is null and b1411 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ChkIsAbsenceMustPro = ChkIsAbsenceMustPro & "E,"
   End If
   'F.打卡異常主管處理待確認
   strSql = "select * from abs014 where b1408='" & strUserNum & "' and b1409 is null and b1411 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ChkIsAbsenceMustPro = ChkIsAbsenceMustPro & "F,"
   End If
   '2014/2/18 END
   'Add By Sindy 2015/7/2
   'G.案件表單主管待簽核
   strSql = "select * from flow003 Where F0308='" & strUserNum & "' and F0309 in('" & Flow_主管審核中 & "')"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ChkIsAbsenceMustPro = ChkIsAbsenceMustPro & "G,"
   End If
   'H.案件目前表單待處理
   strSql = "select * from flow003 Where F0316='" & strUserNum & "' and F0309 in('" & Flow_退回 & "','" & Flow_智權補件 & "')"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ChkIsAbsenceMustPro = ChkIsAbsenceMustPro & "H,"
   End If
   '2015/7/2 END
   'Add By Sindy 2022/10/17
   'I.外專程序系統收件區
   If PUB_GetST03(strUserNum) = "F22" Then
      '待核准信件:1.輸入 2.不處理 9.回信 5.已處理; 人員待處理信件:未處理和 4.歸卷 9.回信 3.退回
      strSql = "select * from inputrecord Where IR08=0 and IR16 in('1','2','9','5') and ir22='" & strUserNum & "'" & _
               "union select * from inputrecord Where IR08=0 and (IR16 is null or IR16 in('4','9','3')) and IR04='" & strUserNum & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         ChkIsAbsenceMustPro = ChkIsAbsenceMustPro & "I,"
      End If
   End If
   '2022/10/17 END
   'Add By Sindy 2025/11/3
   'J.下班逾30分鐘原因確認
   strSql = "select * from abs015 Where B1501='" & strUserNum & "' and B1504 is null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      ChkIsAbsenceMustPro = ChkIsAbsenceMustPro & "J,"
   End If
   '2025/11/3 END
   
   If ChkIsAbsenceMustPro <> "" Then ChkIsAbsenceMustPro = Left(ChkIsAbsenceMustPro, Len(ChkIsAbsenceMustPro) - 1)
End Function

'Add By Sindy 2011/8/25 各系統須檢查人事出缺勤、簽核表單、系統收件區是否有待辦事件,若有,顯示訊息提醒操作人員
'Modify By Sindy 2022/10/17 + Optional ByRef strData As String : 回傳
Public Function ChkIsAbsenceMustProMsg(Optional ByRef strData As String) As String
Dim strTemp As Variant
Dim i As Integer
'Dim strText As String
'Dim nResponse
   
   strData = ChkIsAbsenceMustPro
   strTemp = Split(strData, ",")
   For i = 0 To UBound(strTemp)
      '出缺勤表單待簽核
      If strTemp(i) = "A" Then ChkIsAbsenceMustProMsg = ChkIsAbsenceMustProMsg & "表單待簽核、"
      '出缺勤表單待處理
      If strTemp(i) = "B" Then ChkIsAbsenceMustProMsg = ChkIsAbsenceMustProMsg & "表單待處理、"
      '出缺勤統計待確認
      If strTemp(i) = "C" Then ChkIsAbsenceMustProMsg = ChkIsAbsenceMustProMsg & "出缺勤統計待確認、"
      '個人資料明細待確認
      If strTemp(i) = "D" Then ChkIsAbsenceMustProMsg = ChkIsAbsenceMustProMsg & "個人資料明細待確認、"
      'Add By Sindy 2015/7/2
      '案件表單主管待簽核
      If strTemp(i) = "G" Then ChkIsAbsenceMustProMsg = ChkIsAbsenceMustProMsg & "案件表單待簽核、"
      '案件目前表單待處理
      If strTemp(i) = "H" Then ChkIsAbsenceMustProMsg = ChkIsAbsenceMustProMsg & "案件表單待處理、"
      '2015/7/2 END
      'Add By Sindy 2022/10/17
      '外專系統收件區
      If strTemp(i) = "I" Then ChkIsAbsenceMustProMsg = ChkIsAbsenceMustProMsg & "系統收件區郵件未處理，請注意。"
      '2022/10/17 END
      'Add By Sindy 2025/11/3
      '下班逾30分鐘原因確認
      If strTemp(i) = "J" Then
         '大部分系統訊息另外彈,所以不用串進來
         If strSrvDate(1) >= 20990101 Then
         If InStr(UCase(App.EXEName), "ACCOUNT") > 0 _
            Or InStr(UCase(App.EXEName), "SALARY") > 0 _
            Or InStr(UCase(App.EXEName), "FINANCE") > 0 Then '要提醒用
            ChkIsAbsenceMustProMsg = ChkIsAbsenceMustProMsg & "下班逾30分鐘原因確認、"
         End If
         End If
      End If
      '2025/11/3 END
   Next i
   If ChkIsAbsenceMustProMsg <> "" Then
      ChkIsAbsenceMustProMsg = Left(ChkIsAbsenceMustProMsg, Len(ChkIsAbsenceMustProMsg) - 1)
   End If
End Function

'Modify By Sindy 2025/11/3 改為共用函數
'Modify by Morgan 2008/11/7 登入系統後要執行的程式集中到這裡做
Public Sub MDIFormStarProc()
Dim strRefData As String, strText As Variant
Dim nResponse
   
   'Added by Morgan 2012/1/5 視窗原為最大化,改指定大小及位置以方便其他應用程式操作
   Static bolFormIsSet As Boolean
   If bolFormIsSet = False Then
      PUB_InitFormPos Forms(0) 'Me Modify By Sindy 2025/11/3
      bolFormIsSet = True
      'Add By Sindy 2025/11/7
      If InStr(UCase(App.EXEName), "PROMOTER") > 0 Then
      '2025/11/7 END
         Call Forms(0).SetPersonMenu 'Added by Lydia 2019/06/27 設定智權部->個人常用區
      End If
   End If
   'end 2012/1/5
      
   'Add By Sindy 2011/9/15 檢查人事出缺勤是否有待辦事件
   '只要執行一次
   If m_blnABSActivated = False Then
      m_blnABSActivated = True
      pub_CallNextABSForm = False
      Call Forms(0).SetTmpForm 'Add By Sindy 2015/7/2
      
      '檢查人事出缺勤、簽核表單、系統收件區是否有待辦事件,若有,顯示訊息提醒操作人員
      strText = ChkIsAbsenceMustProMsg(strRefData)
      If strText <> "" Then
         nResponse = MsgBox("您有" & strText & "，現在是否要進行處理？", vbYesNo + vbCritical + vbQuestion, "電子簽核")
         If nResponse = vbYes Then
            pub_CallNextABSForm = True
            'strText = ChkIsAbsenceMustPro
            If InStr(1, strRefData, "A") > 0 Then
               Tmpfrm180201.Show
            ElseIf InStr(1, strRefData, "B") > 0 Then
               Tmpfrm180101.Show
            ElseIf InStr(1, strRefData, "C") > 0 Then
'               frm160201.intChoose = 1
'               frm160201.Hide
'               Call frm160201.cmdok_Click(0)
''               Unload frm160201
               Tmpfrm180203_1.Show
            ElseIf InStr(1, strRefData, "D") > 0 Then
               Tmpfrm160102.intChoose = 1
               Tmpfrm160102.Hide
               Call Tmpfrm160102.cmdok_Click(0)
'               Unload frm160102
            'Add By Sindy 2015/7/2
            ElseIf InStr(1, strRefData, "G") > 0 Then
               If TypeName(Tmpfrm210148) <> "Nothing" Then
                  Tmpfrm210148.Show
               End If
            ElseIf InStr(1, strRefData, "H") > 0 Then
               If TypeName(Tmpfrm210147) <> "Nothing" Then
                  Tmpfrm210147.Show
               End If
            '2015/7/2 END
            'Add By Sindy 2022/10/17
            ElseIf InStr(1, strRefData, "I") > 0 Then
               If TypeName(Tmpfrm06010616) <> "Nothing" Then
                  Tmpfrm06010616.Show
               End If
            End If
         Else
            Call Forms(0).SysStartCallForm 'Add By Sindy 2011/10/7
         End If
      Else
         Call Forms(0).SysStartCallForm 'Add By Sindy 2011/10/7
      End If
      'Add By Sindy 2025/11/3 +if
      If PUB_GetST03(strUserNum) = "P22" Then '內商程序人員
      '2025/11/3 END
         Call ChkDir1728 'Added by Lydia 2017/06/16 檢查並管理收款寄證資料夾的檔案
      End If
      
      'Add By Sindy 2025/11/3
      If strRefData <> "" And strSrvDate(1) >= 20990101 Then
         If InStr(1, strRefData, "J") > 0 Then
            MsgBox "您尚有未處理的下班逾30分鐘原因未輸入！", , "請確認"
            Tmpfrm160018.Show
         End If
      End If
      '2025/11/3 END
   Else
      If Forms(0).m_ChkIsOpenFrm180203 = False Then 'Add By Sindy 2013/7/8 因此視窗是個強制視窗,若開著時程式往下執行會出現錯誤
         Call Forms(0).SysStartCallForm 'Add By Sindy 2011/11/8
      End If
   End If
   'Add by Amy 2017/02/03 判斷是否有圖書借閱記錄需簽核
   If m_bolLoanAct = False Then
      m_bolLoanAct = True
      If GetLoanRecordApply = True Then
          nResponse = MsgBox("您有圖書借閱資料需簽核，現在是否要進行處理？", vbYesNo + vbCritical + vbQuestion, "圖書借閱簽核")
          If nResponse = vbYes Then
              Tmpfrm010035_2.cmdSearch.Visible = False
              Tmpfrm010035_2.cmdPrePage.Visible = False
              Tmpfrm010035_2.cmdOK(4).Visible = False
              Tmpfrm010035_2.cmdOK(5).Visible = False
              Tmpfrm010035_2.Option1(1).Value = True
              Call Tmpfrm010035_2.cmdSearch_Click
          End If
      End If
      'Added by Lydia 2020/01/15 非外專人員行事曆提醒通知
      'Add By Sindy 2025/11/3 +if
      If Left(PUB_GetST03(strUserNum), 2) <> "F2" Then '非外專人員
      '2025/11/3 END
         If PUB_CheckStaffCalendarDue = True Then
              Tmpfrm060209.m_Role = "F41"
              Tmpfrm060209.Show
         End If
         'end 2020/01/15
      End If
   End If
End Sub

'Added by Lydia 2017/06/16 檢查並管理收款寄證的資料夾
Private Sub ChkDir1728()
Dim strDefDir As String
Dim strAd As String
Dim mF As String, mD As String

    strDefDir = GetMyDocPath & "\收款寄證"
    
    mF = Dir(strDefDir & "\*.pdf", vbNormal)
    Do While mF <> ""
       mD = Format(FileDateTime(strDefDir & "\" & mF), "yyyy/mm/dd")
       '刪除超過30天的請款單 或是上一次登入產生的催款單
       If DateDiff("d", mD, ChangeTStringToTDateString(strSrvDate(1))) > 30 Or InStr(mF, "_DNX") = 0 Then
          If PUB_ChkFileOpening(strDefDir & "\" & mF) = True Then
             MsgBox strDefDir & "\" & mF & vbCrLf & "檔案正在使用中，無法刪除。", vbExclamation
             Exit Do
          End If
          Kill strDefDir & "\" & mF   '刪除檔案
       End If
       mF = Dir()
    Loop
End Sub

'Added by Lydia 2020/05/12 法律所案源收文：取得案源類別
Public Function PUB_GetLOSkind(ByVal iCP01 As String, ByVal iCP10 As String, Optional ByVal iNa01 As String = "000") As String
'iNa01: 申請國家, 不傳入預設為台灣案
Dim strR As String, intR As Integer
Dim rsRD As New ADODB.Recordset

   PUB_GetLOSkind = ""
      
   If InStr(iCP01, "L") > 0 Then
       PUB_GetLOSkind = "A"  'Memo by Lydia 2020/07/22 櫃台收文判斷客戶的cu12非法務，才歸入案源收文
   'Modified by Lydia 2020/06/04 +著作權TC
   ElseIf (iCP01 = "P" Or iCP01 = "FCP" Or iCP01 = "T" Or iCP01 = "FCT" Or iCP01 = "TC") And iCP10 <> "" Then
        If iNa01 <> "000" Then Exit Function '限台灣案 'Move by Lydia 2020/07/22 從If InStr(iCP01, "L") > 0 Then 上面移過來
        strR = "select ocode from setspecman where ocode in ('B1P','B1T','B2P','B2T','CP','CT') and instr(oman," & CNULL(Format(iCP10, "000")) & ")> 0 "
        If iCP01 = "P" Or iCP01 = "FCP" Then
           strR = strR & " and ocode like '%P' "
        Else
           strR = strR & " and ocode like '%T' "
        End If
        intR = 1
        Set rsRD = ClsLawReadRstMsg(intR, strR)
        If intR = 1 Then
             PUB_GetLOSkind = rsRD.Fields("ocode")
        End If
   End If
   Set rsRD = Nothing
End Function

'Added by Lydia 2020/06/04 法律所案源收文：判斷是否為補收文=>案源類別
Public Function PUB_GetLOSplus(ByVal iCP01 As String, ByVal iCP02 As String, ByVal iCP03 As String, ByVal iCP04 As String, ByVal iCP10 As String, Optional ByVal iNa01 As String = "000", Optional ByVal pCode As String) As String
'iNa01: 申請國家, 不傳入預設為台灣案
'pCode: (原)案源類型
Dim strR As String, intR As Integer
Dim strCP10s  As String
Dim rsRD As New ADODB.Recordset
Dim strChuTing As String, strSuYuan As String 'Added by Lydia 2020/06/08 訴願和出庭的案件性質

   
   PUB_GetLOSplus = ""
   If iNa01 <> "000" Then Exit Function '限台灣案
   
   strChuTing = PUB_GetLOSsrcCP10s(IIf(iCP01 = "P" Or iCP01 = "FCP", "P", "T"), strSuYuan)
   
   If (iCP01 = "P" Or iCP01 = "FCP" Or iCP01 = "T" Or iCP01 = "FCT" Or iCP01 = "TC") And iCP10 <> "" And iCP02 <> "" Then
       If iCP03 = "" Then iCP03 = "0"
       If iCP04 = "" Then iCP04 = "00"

       If (iCP01 = "P" Or iCP01 = "FCP") And Right(pCode, 1) <> "P" Then pCode = pCode & "P"
       If (iCP01 = "T" Or iCP01 = "FCT" Or iCP01 = "TC") And Right(pCode, 1) <> "T" Then pCode = pCode & "T"
       
       'C類:若有其他C類已收未發程序則視為同一案源補收文，不算新案源以一般接洽單處理即可
       If Left(pCode, 1) = "C" Then
            strCP10s = Pub_GetSpecMan(pCode)
            'Modified by Lydia 2020/07/20 到收文,分案階段,直接以特殊設定的案件性質判斷; 因為舊案的一般收文會被歸入C類
            'strR = "select * from caseprogress  where " & ChgCaseprogress(iCP01 & iCP02 & iCP03 & iCP04) & _
                    " and cp158=0 and cp159=0 and instr('" & strCP10s & "',cp10)>0"
            'intR = 1
            'Set rsRd = ClsLawReadRstMsg(intR, strR)
            'If intR = 0 Then
            '   PUB_GetLOSplus = "C"
            'End If
            'Modified by Lydia 2020/09/28 取得模組自動將「,」改為「;」 ; ex.P-123453 參加訴訟AA9040295
            'If strCP10s <> "" And InStr(strCP10s & ",", iCP10 & ",") > 0 Then
            If strCP10s <> "" And InStr(strCP10s & ";", iCP10 & ";") > 0 Then
                PUB_GetLOSplus = "C"
            End If
            'end 2020/07/20
       'Mark by Lydia 2020/06/08 先保留
       'Else
       '     If pCode = "P" Then
       '         strCP10s = "211,212"
       '     End If
       '     If pCode = "T" Then
       '         strCP10s = "204,205"
       '     End If
       '     '檢查是否為B2案源補收準備程序/言詞辯論
       '     If strCP10s <> "" And InStr(strCP10s, iCP10) > 0 Then
       '        strR = "select * from caseprogress,lawofficesource where " & ChgCaseprogress(iCP01 & iCP02 & iCP03 & iCP04) & " and cp158=0 and cp159=0 and los15(+)=cp162 and los02='B2'"
       '        intR = 1
       '        Set rsRd = ClsLawReadRstMsg(intR, strR)
       '        '有收文B2類案源
       '        If intR > 0 Then
       '           PUB_GetLOSplus = "B2"
       '        End If
       '     End If
       'end 2020/06/08
       End If
   End If
   'Added by Lydia 2020/06/08 準備程序、言詞辯論、訴願可輸入案源,但不一定是案源
   'Modified by Lydia 2020/11/04 +FCP,FCT
   If (iCP01 = "P" Or iCP01 = "T" Or iCP01 = "TC" Or iCP01 = "FCP" Or iCP01 = "FCT") And iCP10 <> "" Then
     '只提醒,不做存檔檢查
     If (strChuTing <> "" And InStr(strChuTing, iCP10) > 0) Or (strSuYuan <> "" And InStr(strSuYuan, iCP10) > 0) Then
           PUB_GetLOSplus = "D"
     End If
   End If
   'end 2020/06/08
   
   Set rsRD = Nothing
End Function

'Added by Lydia 2020/06/08 法律所案源收文：設定案件性質(準備程序,言詞辯論,訴願)
Public Function PUB_GetLOSsrcCP10s(ByVal pCP01 As String, Optional ByRef pCP10s_2 As String) As String
'PUB_GetLOSsrcCP10s: 1.案件性質: 出庭
'pCP10s_2: 2.案件性質: 訴願
     
     '設定案件性質(準備程序,言詞辯論,訴願)
     'Modified by Lydia 2020/11/04 +FCP
     If pCP01 = "P" Or pCP01 = "FCP" Then
         PUB_GetLOSsrcCP10s = "211,212"
         pCP10s_2 = "501,505"
     Else
         PUB_GetLOSsrcCP10s = "204,205"
         pCP10s_2 = "401,406"
     End If
     
End Function

'Added by Lydia 2020/05/20 法律所案源收文：P/T/FCP/FCT分案(限台灣案)，更新相關資料
Public Sub PUB_UpdateLOS01(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pCP09 As String, ByVal pAppNo As String, ByVal pType As String)
Dim strB1 As String
Dim intB As Integer, intP As Integer
Dim strTempB As Variant
Dim strCUNo As String, strCU01 As String
Dim rsR1 As New ADODB.Recordset
Dim rsB1 As New ADODB.Recordset  'Added by Lydia 2020/09/24
Dim strCRL01 As String, m_F0316 As String, m_F0308 As String, m_F0309 As String 'Add By Sindy 2022/10/7

'pCP01~pCP04、pCP09: 分案之本所案號、收文號
'pAppNo: 傳入申請人編號，用,區隔
'pType: C類”是否需要法律所配合”
    
    '排除法律所案源檔已放棄LOS07
    'Modified by Lydia 2020/09/24 抓法律所案號
    'strB1 = "SELECT CP09,CP16, X.* FROM CASEPROGRESS, LawOfficeSource X WHERE CP09='" & pCP09 & "' AND CP162 IS NOT NULL " & _
                "AND CP162=LOS15(+) AND LOS15 IS NOT NULL AND LOS07 IS NULL"
    strB1 = "SELECT A1.CP09,A1.CP16,B1.CP01 AS LCP01,B1.CP02 AS LCP02,B1.CP03 AS LCP03,B1.CP04 AS LCP04, X.* FROM CASEPROGRESS A1, LawOfficeSource X,CASEPROGRESS B1 " & _
               "WHERE A1.CP09='" & pCP09 & "' AND A1.CP162 IS NOT NULL AND A1.CP162=LOS15(+) AND LOS15 IS NOT NULL AND LOS07 IS NULL AND LOS06=B1.CP09(+) "
    intB = 1
    Set rsR1 = ClsLawReadRstMsg(intB, strB1)
    If intB = 1 Then
        rsR1.MoveFirst
        If "" & rsR1.Fields("LOS01") = "" Then
             strB1 = "update LawOfficeSource set los01='" & pCP09 & "' where los15='" & rsR1.Fields("los15") & "' and los01 is null "
             cnnConnection.Execute strB1, intB
             
             'Added by Morgan 2022/11/10 B2類案源法務案接洽單要設定為P/T案的期限--秀玲
             If rsR1.Fields("los02") = "B2" Then
               strB1 = "update ConsultRecordList set (CRL12,CRL13)=(select cp06,cp07 from caseprogress where cp09='" & pCP09 & "')" & _
                 " where CRL01='" & rsR1.Fields("los17") & "' and CRL13 is null"
               cnnConnection.Execute strB1, intB
             End If
             'end 2022/11/10
        ElseIf Val("" & rsR1.Fields("CP16")) > 0 Then
             strB1 = "update LawOfficeSource set los01='" & pCP09 & "' where los15='" & rsR1.Fields("los15") & "' "
             cnnConnection.Execute strB1, intB
        End If
        If "" & rsR1.Fields("los10") <> "" And Left("" & rsR1.Fields("los02"), 1) = "B" Then
            strB1 = "update caseprogress set cp43='" & pCP09 & "' where cp09='" & rsR1.Fields("los10") & "' "
            cnnConnection.Execute strB1, intB
        End If
        '3.  存檔時以CP162抓法律所案源檔之案源單號，5/26若無案源總收文號(LOS01)或該程序有費用則更新案源總收文號為本案之總收文號。
             '另案件類型B1、B2者，同時再以案源檔之收據總收文號(LOS10)更新TT案的CP43為本案之總收文號，
             '並以本案的申請人編號更新法律案接洽單(LOS17)的申請人編號(CRA05+CRA06)。
        If "" & rsR1.Fields("LOS01") = "" Then '同一接洽單(案源)以第一筆分案的回答為準
            If "" & rsR1.Fields("los17") <> "" And pAppNo <> "" Then
                strTempB = Split(pAppNo, ",")
                For intP = 0 To UBound(strTempB)
                     If Trim(strTempB(intP)) <> "" Then
                         strCUNo = ChangeCustomerL(strTempB(intP))
                         If intP = 0 Then strCU01 = strCUNo
                         strB1 = "update consultrecapp set cra05='" & Mid(strCUNo, 1, 8) & "' , cra06='" & Mid(strCUNo, 9, 1) & "' where cra01='" & rsR1.Fields("los17") & "' and cra02=" & intP + 1 & " "
                         cnnConnection.Execute strB1, intB
                         If intB = 0 Then  '補:接洽記錄單申請人
                             strB1 = "insert consultrecapp(cra01,cra02, cra05,cra06) values('" & rsR1.Fields("los17") & "', " & intP + 1 & ", '" & Mid(strCUNo, 1, 8) & "' ,'" & Mid(strCUNo, 9, 1) & "' )"
                         End If
                         'Mark by Lydia 2021/09/07 判斷接洽單申請人為舊客戶則不更新申請人; 參考frm090801法務舊案客戶和P/T案舊案客戶編號不一致，改成詢問；ex.P-074838用X29213000(個人), L-006229用X29213120(磐石)
                         '-------目前不啟用的原因：因為該案是去年案源未上線前收文，所以理論不應收；修改程式先保留
                         'strB1 = "select cra03 from ConsultRecApp where cra01='" & rsR1.Fields("los17") & "' and cra02=" & intP + 1 & " "
                         'intB = 1
                         'Set rsB1 = ClsLawReadRstMsg(intB, strB1)
                         'If intB = 0 Then
                         '    '補:接洽記錄單申請人
                         '    strB1 = "insert consultrecapp(cra01,cra02, cra05,cra06) values('" & rsR1.Fields("los17") & "', " & intP + 1 & ", '" & Mid(strCUNo, 1, 8) & "' ,'" & Mid(strCUNo, 9, 1) & "' )"
                         '    cnnConnection.Execute strB1, intB
                         'Else
                         '    If "" & rsB1.Fields("cra03") = "Y" Then '接洽單申請人為新客戶
                         '        strB1 = "update consultrecapp set cra05='" & Mid(strCUNo, 1, 8) & "' , cra06='" & Mid(strCUNo, 9, 1) & "' where cra01='" & rsR1.Fields("los17") & "' and cra02=" & intP + 1 & " "
                         '        cnnConnection.Execute strB1, intB
                         '    End If
                         'End If
                         'end 2021/09/07
                     End If
                Next intP
            ElseIf pAppNo <> "" Then
               strCU01 = ChangeCustomerL(Mid(pAppNo, 1, InStr(pAppNo, ",") - 1))
            End If
            
            '4.  案件類型C類增'是否需要法律所配合'，預設Y。不需要者可以以歷程給律師判發，且將法律所案源檔的放棄日期、放棄人員填入、放棄原因：專業部分案不需法律所配合。
            If Left("" & rsR1.Fields("los02"), 1) = "C" And pType <> "Y" And "" & rsR1.Fields("los07") = "" Then
                '放棄人員=操作者
                strB1 = "update LawOfficeSource set los07=" & strSrvDate(1) & ", los08='" & strUserNum & "' , los09='專業部分案不需法律所配合' where los15='" & rsR1.Fields("los15") & "' and los07 is null"
                cnnConnection.Execute strB1, intB
                
            ElseIf Left("" & rsR1.Fields("los02"), 1) = "B" Then
                If "" & rsR1.Fields("LOS01") = "" And "" & rsR1.Fields("LOS06") <> "" Then
                    '5/29 案源已有法律所總收文號者不必通知(A轉B1補收文)。
                    'Added by Lydia 2020/09/24  L-6292(A4類改為B類)通知內專補收文P-125809：補上案件關聯
                    If "" & rsR1.Fields("LCP01") <> "" And "" & rsR1.Fields("LCP02") <> "" Then
                        strB1 = "select cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08 from caserelation1 where cr01='" & rsR1.Fields("LCP01") & "' and cr02='" & rsR1.Fields("LCP02") & "' and cr03='" & rsR1.Fields("LCP03") & "' and cr04='" & rsR1.Fields("LCP04") & "' and cr05='" & pCP01 & "' and cr06='" & pCP02 & "' and cr07='" & pCP03 & "' and cr08='" & pCP04 & "' "
                        intB = 1
                        Set rsB1 = ClsLawReadRstMsg(intB, strB1)
                        If intB = 0 Then
                           strB1 = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(rsR1.Fields("LCP01")) & ", " & CNULL(rsR1.Fields("LCP02")) & ", " & CNULL(rsR1.Fields("LCP03")) & ", " & CNULL(rsR1.Fields("LCP04")) & ", " & CNULL(pCP01) & ", " & CNULL(pCP02) & ", " & CNULL(pCP03) & ", " & CNULL(pCP04) & " ) "
                           cnnConnection.Execute strB1
                        End If
                        strB1 = "select cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08 from caserelation1 where cr01='" & pCP01 & "' and cr02='" & pCP02 & "' and cr03='" & pCP03 & "' and cr04='" & pCP04 & "' and cr05='" & rsR1.Fields("LCP01") & "' and cr06='" & rsR1.Fields("LCP02") & "' and cr07='" & rsR1.Fields("LCP03") & "' and cr08='" & rsR1.Fields("LCP04") & "' "
                        intB = 1
                        Set rsB1 = ClsLawReadRstMsg(intB, strB1)
                        If intB = 0 Then
                           strB1 = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(pCP01) & ", " & CNULL(pCP02) & ", " & CNULL(pCP03) & ", " & CNULL(pCP04) & ", " & CNULL(rsR1.Fields("LCP01")) & ", " & CNULL(rsR1.Fields("LCP02")) & ", " & CNULL(rsR1.Fields("LCP03")) & ", " & CNULL(rsR1.Fields("LCP04")) & " ) "
                           cnnConnection.Execute strB1
                        End If
                    End If
                    'end 2020/09/24
                    
                    'Added by Morgan 2025/4/18 若有補收款還是要通知法律所收文
                    If Not IsNull(rsR1.Fields("los20")) Then
                        pType = "Y"
                    End If
                    'end 2025/4/18
                Else
                    pType = "Y"
                End If
            End If
            
            'Add By Sindy 2022/10/7 取得法律所接洽單編號
            strCRL01 = ""
            If "" & rsR1.Fields("los20") <> "" Then
               strCRL01 = rsR1.Fields("los20")
            ElseIf "" & rsR1.Fields("los17") <> "" Then
               strCRL01 = rsR1.Fields("los17")
            End If
            '2022/10/7 END
            
            '5.  案件類型B1、B2及上述C類選Y者同時E-MAIL給指定法務人員或各所窗口。
            If pType = "Y" Then
'               'Add By Sindy 2022/10/7 要配合: 因此法律接洽單則新增一筆”法務案源”簽核資料
'               If Val("" & rsR1.Fields("los12")) >= 接洽單電子收文啟用日 And strCRL01 <> "" Then
'                  strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(strCRL01) & ",'A4',1,'A4')"
'                  cnnConnection.Execute strSql
'                  '讀取下一處理人員
'                  Call GetNextProPerson_Flow(strCRL01, m_F0316, m_F0308, m_F0309)
'               End If
'               '2022/10/7 END
               
               Call PUB_AddMailCache_LOS("1", rsR1.Fields("LOS15"))
               
            'Add By Sindy 2022/10/7 放棄
            Else
               If Val("" & rsR1.Fields("los12")) >= 接洽單電子收文啟用日 And strCRL01 <> "" Then
                  '表單主檔
                  strSql = "update FLOW003 set " & _
                           "F0307='" & strUserNum & "'" & _
                           ",F0309='" & Flow_放棄案源 & "'" & _
                           " where F0301='" & strCRL01 & "'"
                  cnnConnection.Execute strSql
               End If
               '2022/10/7 END
            End If
        End If
    End If
    Set rsR1 = Nothing
    Set rsB1 = Nothing 'Added by Lydia 2020/09/24
End Sub

'Added by Morgan 2020/5/21 依員工所別取得法務預設窗口
'Modified by Lydia 2020/06/23 +系統別
Public Function PUB_GetLawCaseWindow(pSaleNo As String, pSysNo As String) As String
   Dim stSQL As String, intQ As Integer
   Dim RsQ As ADODB.Recordset
   
   'Added by Lydia 2020/06/23
   If pSysNo = "FCL" Or pSysNo = "LIN" Or pSysNo = "CFL" Then
        stSQL = "select oMan from SetSpecMan where ocode='法務外法窗口' "
   Else
   'end 2020/06/23
        stSQL = "select oMan from SetSpecMan" & _
           " where oCode=(select decode(st06,'2','法務中所窗口','3','法務南所窗口','4','法務高所窗口','法務北所窗口')" & _
           " from staff where st01='" & pSaleNo & "')"
   End If 'Added by Lydia 2020/06/23
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      PUB_GetLawCaseWindow = "" & RsQ.Fields("oMan")
   End If
   Set RsQ = Nothing
End Function

'Added by Lydia 2020/05/22 法律所案源收文：Email通知(mailcache)
'Modified by Lydia 2020/06/16 +iTitle 通知的主旨(部份)
Public Sub PUB_AddMailCache_LOS(ByVal iKind As String, ByVal iPKey As String, Optional ByVal iTitle As String = "")
'iKind : Email類型 1-(智權部)已新增案源、2-法律所分案、3-法律所分案補收配合開庭
'iPKey : 案源單號
Dim strSubject As String, strContent  As String, strTo As String
Dim intJ As Integer, strQ1 As String, strLOS04n As String
Dim rsQuery As New ADODB.Recordset
Dim strCC As String 'Added by Lydia 2020/07/13
'Added by Lydia 2022/05/03
Dim strTmp1 As String, strLOS04new As String
Dim tmpArr As Variant, intK As Integer
Dim rsQD As New ADODB.Recordset
Dim strPCase(0 To 4) As String  'Added by Lydia 2024/04/19

    If iKind = "" Or iPKey = "" Then Exit Sub
    'Modified by Lydia 2024/04/19 +
    strQ1 = "SELECT LOS01 總收文號,SQLDATET(LOS12) 介紹日期,SQLDATET(NVL(C1.CP06,LOS16)) 管制日期 " & _
                ",A0902 業務區,LOS04 介紹人,NVL(CRA07,CRA08) 介紹客戶,CRL57 介紹內容, " & _
                "C1.CP01||DECODE(C1.CP02,'','','-'||C1.CP02||DECODE(C1.CP03||C1.CP04,'000','','-'||C1.CP03||'-'||C1.CP04)) 專業部案號, " & _
                "DECODE(C2.CP01,NULL,CRL07||DECODE(CRL08,'','','-'||CRL08||DECODE(CRL09||CRL10,'000','','-'||CRL09||'-'||CRL10)) ,C2.CP01||DECODE(C2.CP02,'','','-'||C2.CP02||DECODE(C2.CP03||C2.CP04,'000','','-'||C2.CP03||'-'||C2.CP04))) 法律所案號, " & _
                "CRL17 案件名稱,LOS06, LOS03,LOS15,C2.CP14 AS 法律所承辦人,C2.CP29 AS  協辦人員,LOS02,NVL(C2.CP01,CRL07) as CP01T,C1.CP12,C1.CP01 AS PP01,C1.CP02 AS PP02,C1.CP03 AS PP03,C1.CP04 AS PP04 " & _
                "FROM LAWOFFICESOURCE,CASEPROGRESS C1,ACC090,CONSULTRECORDLIST,CONSULTRECAPP,CASEPROGRESS C2 " & _
                "WHERE LOS15='" & iPKey & "' AND C1.CP09(+)=LOS01 AND A0901(+)=NVL(C1.CP12,CRL04) " & _
                "AND CRL01(+)=LOS17 AND CRA01(+)=CRL01 AND C2.CP09(+)=LOS06 "
    strQ1 = strQ1 & "ORDER BY 1, CRA01 "
    intJ = 1
    Set rsQuery = ClsLawReadRstMsg(intJ, strQ1)
    If intJ = 1 Then
         With rsQuery
             .MoveFirst
             strLOS04n = "" & .Fields("介紹人")
             'Modified by Lydia 2022/05/03 法律所A類案源在接洽記錄單填寫時會發副本給介紹人主管，若介紹人為不可自動收文(ST58='N')人員時，請在郵件內文介紹人姓名後面增加"(不可自動收文人員)"
             'If strLOS04n <> "" Then strLOS04n = Replace(PUB_ReadUserData(strLOS04n), ";", "、")
             If strLOS04n <> "" Then
                 tmpArr = Split(strLOS04n, ";")
                 strLOS04n = "": strLOS04new = ""
                 For intK = 0 To UBound(tmpArr)
                     If Trim("" & tmpArr(intK)) <> "" Then
                        strTmp1 = "select st02, st58 from staff where st01='" & tmpArr(intK) & "' "
                        intJ = 1
                        Set rsQD = ClsLawReadRstMsg(intJ, strTmp1)
                         If intJ = 1 Then
                             strLOS04n = strLOS04n & "、" & rsQD.Fields("st02")
                             strLOS04new = strLOS04new & "、" & rsQD.Fields("st02")
                             If iKind = "1" And Left(.Fields("LOS02"), 1) = "A" And "" & rsQD.Fields("st58") = "N" Then
                                 strLOS04new = strLOS04new & "　(不可自動收文人員)"
                             End If
                         End If
                     End If
                 Next
             End If
             If strLOS04n <> "" Then strLOS04n = Mid(strLOS04n, 2)
             If strLOS04new <> "" Then strLOS04new = Mid(strLOS04new, 2)
             'end 2022/05/03
             
             If iKind = "1" Then
                 '已新增案源
                 strTo = "" & .Fields("los03")
                 If strTo = "" Then
                     '用介紹人員第一人抓法務窗口
                     strQ1 = "" & .Fields("介紹人")
                     If InStr(strQ1, ",") > 0 Then strQ1 = Mid(strQ1, 1, InStr(strQ1, ",") - 1)
                     'Modified by Lydia 2020/06/23 +CP01T
                     strTo = PUB_GetLawCaseWindow(strQ1, "" & rsQuery.Fields("CP01T"))
                 End If
                 strSubject = "已新增案源(介紹人:" & strLOS04n & ", 案源單號:" & .Fields("los15") & ")，請執行案源管理作業！"
                 'Modified by Morgan 2020/5/26 A類不用帶(因固定為TT-999999)
                 If Left(.Fields("LOS02"), 1) <> "A" Then
                   strContent = "專業部案號：" & .Fields("專業部案號") & vbCrLf
                 End If
             
             ElseIf iKind = "2" Or iKind = "3" Then
                 '法律所分案
                 strTo = "" & .Fields("介紹人")
                 If InStr(strTo, ",") > 0 Then strTo = Mid(strTo, 1, InStr(strTo, ",") - 1)  'Added by Lydia 2020/06/09 只通知介紹人員第一人
                 
                 If iKind = "2" Then
                      strSubject = .Fields("介紹日期") & "介紹之案源" & .Fields("專業部案號") & "(" & .Fields("總收文號") & ")已為法務收文，案號為" & .Fields("法律所案號")
                 Else
                      '分案和配合開庭通知整合為一封email
                      'Modified by Lydia 2020/06/16 改成另外通知=> frm077005「智財訴訟案需專業部配合通知補收文作業」
                      'strSubject = .Fields("介紹日期") & "介紹之案源(案源單號:" & .Fields("los15") & ")已為法務收文，需智慧所配合開庭，請補收文配合開庭！"
                      strSubject = .Fields("介紹日期") & "介紹之案源(案源單號:" & .Fields("los15") & ")已為法務收文"
                      If iTitle <> "" Then
                          strSubject = strSubject & "，請補收文" & iTitle & "！"
                      Else
                          strSubject = strSubject & "，需智慧所配合開庭，請補收文配合開庭！"
                      End If
                 End If
                 strContent = "法律所案號：" & .Fields("法律所案號") & vbCrLf
                 '法律所承辦人,C2.CP29 AS  協辦人員
                 If "" & .Fields("法律所承辦人") <> "" Then strContent = strContent & "承辦人：　" & GetStaffName(.Fields("法律所承辦人")) & vbCrLf
                 If "" & .Fields("協辦人員") <> "" Then strContent = strContent & "協辦人員：" & GetStaffName(.Fields("協辦人員")) & vbCrLf
             End If
             
             If strTo <> "" Then
                 'Added by Lydia 2020/07/13 副本:介紹人之主管
                 If iKind = "1" Then  'P/T分案
                     strCC = PUB_GetLos04Man("" & .Fields("介紹人"), "1") '副本主管
                     strQ1 = ""
                     If InStr("" & .Fields("介紹內容"), "主管簽核") > 0 Then  '特殊情形主管:應收帳款等任何要主管簽核的部分
                         strSubject = "(需主管簽核)" & strSubject 'Added by Lydia 2020/07/17 有應收帳款管制，在email主旨前方加註
                         strQ1 = PUB_GetLos04Man("" & .Fields("介紹人"), "2")
                         If strQ1 <> "" And (strCC = "" Or (strCC <> "" And InStr(strCC, strQ1) = 0)) Then
                              strCC = strCC & IIf(strCC <> "", ";", "") & strQ1
                         End If
                     End If
                 ElseIf iKind = "2" Then  '法律所分案
                     '國外部介紹案件 , 法律所接案或放棄案源EMAIL給介紹人時同時副本給主管
                     'Modified by Morgan 2020/8/19 不必排除A類
                      'If Left(Trim("" & .Fields("CP12")), 1) = "F" And Left(.Fields("LOS02"), 1) <> "A" Then
                      If Left(Trim("" & .Fields("CP12")), 1) = "F" Then
                      'end 2020/8/19
                           strCC = PUB_GetLos04Man("" & .Fields("介紹人"), "1") '副本主管
                      End If
                      'Added by Lydia 2024/04/19
                      strPCase(0) = "" & .Fields("專業部案號")
                      Call ChgCaseNo(Replace(strPCase(0), "-", ""), strPCase)
                      'end 2024/04/19
                 End If
                 'end 2020/07/13
                 
                 strContent = strContent & "主題：　　" & .Fields("案件名稱") & vbCrLf
                 strContent = strContent & "介紹日期：" & .Fields("介紹日期") & vbCrLf
                 'Modified by Lydia 2022/05/03 加註"(不可自動收文人員)"
                 'strContent = strContent & "介紹人：　" & strLOS04n & vbCrLf
                 strContent = strContent & "介紹人：　" & strLOS04new & vbCrLf
                 If "" & .Fields("管制日期") <> "" Then strContent = strContent & "管制日期：" & .Fields("管制日期") & vbCrLf
                 strContent = strContent & "介紹客戶：" & .Fields("介紹客戶") & vbCrLf
                 strContent = strContent & "介紹內容：" & .Fields("介紹內容")
                 'Modified by Lydia 2020/07/13 +mc09副本:介紹人之主管
                 strQ1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                          " values ('" & strUserNum & "','" & strTo & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss')" & _
                          ",'" & ChgSQL(strSubject) & "','" & ChgSQL(strContent) & "'," & CNULL(strCC) & ")"
                 cnnConnection.Execute strQ1, intJ
             End If
         End With
    End If
    
    'Added by Lydia 2024/04/19 法律所L或CFL、FCL分案時，若為案源案件且案源總收文號LOS01之本所案號為FCP或P案時發EMAIL，(T、FCT不發)
    If iKind = "2" And (strPCase(1) = "P" Or strPCase(1) = "FCP") Then '法律所分案
       'Modified by Lydia 2025/05/08 請判斷P或FCP之案件性質為503行政訴訟,504行政再審,506參加訴訟,507行政訴訟上訴,510行政更審,515參加上訴的才發EMAIL。 + and c1.cp10 in ('503','504','506','507','510','515')
       strQ1 = "select c1.cp09, c1.cp13 as pcp13,s1.st02 as pcp13n,c1.cp14 as pcp14,c2.cp14 as cp14, s2.st02 as cp14n,decode(pa09,'000',cpm03,cpm04) as cpm03 " & _
               "from lawofficesource,caseprogress c1, caseprogress c2,patent,staff s1, staff s2, casepropertymap " & _
               "where los15='" & iPKey & "' and los01=c1.cp09(+) and los06=c2.cp09(+) and c2.cp14=s2.st01(+) " & _
               "and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+) and c1.cp13=s1.st01(+) " & _
               "and c1.cp01=cpm01(+) and c1.cp10=cpm02(+) and c1.cp10 in ('503','504','506','507','510','515') "
       intJ = 1
       Set rsQuery = ClsLawReadRstMsg(intJ, strQ1)
       If intJ = 1 Then
          '電子公文匯入通知P
          strTo = ""
          If strPCase(1) = "P" Then
             strTo = PUB_GetPHandler(strPCase(0)) 'Added by Morgan 2025/1/24 'Memo by Lydia 2025/05/08 舊Code刪除只保留最新Code
             strSubject = strSubject = strPCase(0) & rsQuery.Fields("cpm03") & "之承辦律師為" & rsQuery.Fields("cp14n") & "，請準備打好字之委任狀提供給" & rsQuery.Fields("pcp13n") & "。"
          Else
             strTo = "" & rsQuery.Fields("pcp14")
             strSubject = strPCase(0) & rsQuery.Fields("cpm03") & "之承辦律師為" & rsQuery.Fields("cp14n") & "，請工程師準備委任狀給客戶/代理人。"
          End If
          If strTo <> "" Then
            'Modified by Lydia 2025/05/07 內文null => 如旨 ; ex.Pub_SendMail沒內文會判斷發文號，而P-087216的法務案源L-006220-5-00分案時，P案已發文所以沒有發mail
            strQ1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc13)" & _
                     " values ('" & strUserNum & "','" & strTo & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss')" & _
                     ",'" & ChgSQL(strSubject) & "','如旨','" & rsQuery.Fields("cp09") & "')"
            cnnConnection.Execute strQ1, intJ
          End If
       End If
       
    End If
    'end 2024/04/19
    
    Set rsQuery = Nothing
    Set rsQD = Nothing 'Added by Lydia 2022/05/03
End Sub

'Added by Lydia 2020/07/13 法律所案源收文：案源介紹人之主管
Public Function PUB_GetLos04Man(ByVal pLOS04 As String, Optional ByVal pType As String = "1") As String
'pType: 1-副本主管(簽核人員1~4) , 2-特殊情況主管(簽核人員5~6)
'Modify By Sindy 2022/8/26 改為 1-副本主管(簽核人員1~3) , 2-特殊情況主管(簽核人員4~6)
Dim strA1 As String
Dim intA As Integer
Dim rsAD As New ADODB.Recordset
    
    PUB_GetLos04Man = ""
    If pLOS04 = "" Or pType = "" Then Exit Function
    If InStr(pLOS04, ",") > 0 Then '介紹人第一人
       pLOS04 = Mid(pLOS04, 1, InStr(pLOS04, ",") - 1)
    ElseIf InStr(pLOS04, ";") > 0 Then
       pLOS04 = Mid(pLOS04, 1, InStr(pLOS04, ";") - 1)
    End If
    
    'Modified by Lydia 2025/11/18 +ST03,ST52,ST53,ST54
    strA1 = "SELECT ST01,ST04,A0908,F0103,F0104,F0105,F0106,F0107,F0108,ST03,ST52,ST53,ST54 " & _
                "From STAFF, ACC090, FLOW001 WHERE ST01='" & pLOS04 & "' AND ST15=A0901(+) AND ST01=F0101(+) AND '3'=F0102(+) "
    intA = 1
    Set rsAD = ClsLawReadRstMsg(intA, strA1)
    If intA = 1 Then
        strA1 = ""
        'Added by Lydia 2025/11/18 外專外商人員改抓ST52、ST53、ST54
        If Left("" & rsAD.Fields("st03"), 2) = "F1" Or Left("" & rsAD.Fields("st03"), 2) = "F2" Then
            If "" & rsAD.Fields("st52") <> "" And "" & rsAD.Fields("st52") <> pLOS04 Then
               strA1 = strA1 & IIf(strA1 <> "", ",", "") & rsAD.Fields("st52")
            End If
            If "" & rsAD.Fields("st53") <> "" And "" & rsAD.Fields("st53") <> pLOS04 Then
               strA1 = strA1 & IIf(strA1 <> "", ",", "") & rsAD.Fields("st53")
            End If
            If "" & rsAD.Fields("st54") <> "" And "" & rsAD.Fields("st54") <> pLOS04 Then
               strA1 = strA1 & IIf(strA1 <> "", ",", "") & rsAD.Fields("st54")
            End If
        Else
        'end 2025/11/18
            '副本主管：以介紹人抓FLOW001案件表單簽核人員設定檔之表單類別為3接洽單;
            'Modify By Sindy 2022/8/26 接洽單的簽核主管改為F0103~F0105簽核人員1~3
            'If pType = "1" And "" & rsAD.Fields("F0103") & rsAD.Fields("F0104") & rsAD.Fields("F0105") & rsAD.Fields("F0106") <> "" Then
            If pType = "1" And "" & rsAD.Fields("F0103") & rsAD.Fields("F0104") & rsAD.Fields("F0105") <> "" Then
               '讀得到且F0103簽核人員1為自已則不必發副本；否則發F0103~F0106簽核人員1~4；
               If "" & rsAD.Fields("F0103") <> pLOS04 Then
                   'strA1 = rsAD.Fields("F0103") & "," & rsAD.Fields("F0104") & "," & rsAD.Fields("F0105") & "," & rsAD.Fields("F0106")
                   strA1 = rsAD.Fields("F0103") & "," & rsAD.Fields("F0104") & "," & rsAD.Fields("F0105")
               End If
            '特殊情形主管：以介紹人抓FLOW001案件表單簽核人員設定檔之表單類別為3接洽單：
            'Modify By Sindy 2022/8/26 接洽單的特例簽核主管改為F0106~F0108簽核人員4~6
            'ElseIf pType = "2" And "" & rsAD.Fields("F0107") & rsAD.Fields("F0108") <> "" Then
            ElseIf pType = "2" And "" & rsAD.Fields("F0106") & rsAD.Fields("F0107") & rsAD.Fields("F0108") <> "" Then
               '讀得到且F0107簽核人員5為自已則不必發；否則發F0107~F0108簽核人員5~6；
               'If "" & rsAD.Fields("F0107") <> pLOS04 Then
               If "" & rsAD.Fields("F0106") <> pLOS04 Then
                   'strA1 = rsAD.Fields("F0107") & "," & rsAD.Fields("F0108")
                   strA1 = rsAD.Fields("F0106") & "," & rsAD.Fields("F0107") & "," & rsAD.Fields("F0108")
               End If
            Else
               '若讀不到則副本改發部門主管A0908及秀玲83002；
               strA1 = rsAD.Fields("A0908") & "," & Pub_GetSpecMan("程式管理人員")
            End If
        End If 'Added by Lydia 2025/11/18
        If strA1 <> "" Then
            PUB_GetLos04Man = Replace(Replace(GetAddStr(strA1), "'", ""), ",", ";")
        End If
    End If
    Set rsAD = Nothing
End Function

'Add By Sindy 2022/10/3
'刪除接洽單及簽核流程
Public Function PUB_DelCRLAllData(strCRL01 As String) As Boolean
Dim strCRL74 As String, strCRL55 As String
Dim strLOS15 As String '案源案號 Add By Sindy 2024/5/7
Dim bolHadRecv As Boolean
Dim strLOS10 As String

On Error GoTo CheckingErr
   
   PUB_DelCRLAllData = False
   
   strSql = "select * from CONSULTRECORDLIST where CRL01='" & strCRL01 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strCRL74 = "" & RsTemp.Fields("CRL74") '相關案號的類別:1.查名代號
      strCRL55 = "" & RsTemp.Fields("CRL55") '查名代號
      strLOS15 = GetCRL55toLOS15(strCRL55) '案源案號 Add By Sindy 2024/5/7
   End If
   '是否已收文
   strSql = "select * from CONSULTRECCMP where CRC01='" & strCRL01 & "' AND CRC08 is not null"
   intI = 1
   bolHadRecv = False
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      bolHadRecv = True '已收文
   End If
   
   If bolHadRecv = False Then '未收文才刪
      '查名代號(收文組群)記錄
      If strCRL74 = "1" And strCRL55 <> "" Then
         strSql = "DELETE FROM tmqcasemap WHERE tqc01='" & strCRL55 & "'"
         Pub_SeekTbLog strSql '記錄Log
         cnnConnection.Execute strSql, intI
      End If
      
      '法律案源資料
      If Len(strCRL74) = 2 And strLOS15 <> "" Then
         '補收款接洽單
         strSql = "update LawOfficeSource set LOS20=null where LOS15='" & strLOS15 & "' and LOS20='" & strCRL01 & "'"
         Pub_SeekTbLog strSql '記錄Log
         cnnConnection.Execute strSql, intI
         If intI = 0 Then '其他案源狀況
            'Add By Sindy 2024/9/20 依案源單號檢查是否有已收文,若有,則不可刪除
            strSql = "select cp09,los01,los06,los10,los17,los18,cp01,cp02,cp03,cp04 from lawofficesource,caseprogress" & _
                     " where los15='" & strLOS15 & "' and los06=cp09(+) and cp09 is not null" & _
                     " Union All" & _
                     " select cp09,los01,los06,los10,los17,los18,cp01,cp02,cp03,cp04 from lawofficesource,caseprogress" & _
                     " where los15='" & strLOS15 & "' and los01=cp09(+) and cp09 is not null and los01<>los10"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 0 Then
            '2024/9/20 END
               'Add By Sindy 2024/9/25 檢查案源主檔是否還有其他接洽單存在,若有,則不可刪除
               strSql = "select los15,los17,los18,los20 from lawofficesource,CONSULTRECORDLIST" & _
                        " where los15='" & strLOS15 & "'" & _
                        " and los17=crl01 and los17<>'" & strCRL01 & "' and los17 is not null" & _
                        " union select los15,los17,los18,los20 from lawofficesource,CONSULTRECORDLIST" & _
                        " where los15='" & strLOS15 & "'" & _
                        " and los18=crl01 and los18<>'" & strCRL01 & "' and los18 is not null" & _
                        " union select los15,los17,los18,los20 from lawofficesource,CONSULTRECORDLIST" & _
                        " where los15='" & strLOS15 & "'" & _
                        " and los20=crl01 and los20<>'" & strCRL01 & "' and los20 is not null"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 0 Then
               '2024/9/25 END
                  strSql = "select * from LawOfficeSource where LOS15='" & strLOS15 & "'" & _
                           " and (LOS17='" & strCRL01 & "' or LOS18='" & strCRL01 & "')"
                  intI = 1: strLOS10 = ""
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     strLOS10 = "" & RsTemp.Fields("LOS10") 'TT總收文號
                     If strLOS10 <> "" Then
                        strSql = "DELETE FROM caseprogress where cp09='" & strLOS10 & "'"
                        Pub_SeekTbLog strSql '記錄Log
                        cnnConnection.Execute strSql, intI
                        
                        PUB_DelFtpFile2 strLOS10
                        strSql = "delete from casepaperpdf where cpp01='" & strLOS10 & "'"
                        Pub_SeekTbLog strSql '記錄Log
                        cnnConnection.Execute strSql, intI
                     End If
                  End If
                  
                  strSql = "DELETE FROM LawOfficeSource where LOS15='" & strLOS15 & "' and (LOS17='" & strCRL01 & "' or LOS18='" & strCRL01 & "')"
                  Pub_SeekTbLog strSql '記錄Log
                  cnnConnection.Execute strSql, intI
               End If
            End If
         End If
      End If
   End If
   
   '接洽記錄單主檔
   strSql = "DELETE FROM CONSULTRECORDLIST WHERE CRL01='" & strCRL01 & "'"
   Pub_SeekTbLog strSql '記錄Log
   cnnConnection.Execute strSql, intI
   '案件性質
   strSql = "DELETE FROM ConsultRecCMP WHERE CRC01='" & strCRL01 & "'"
   Pub_SeekTbLog strSql '記錄Log
   cnnConnection.Execute strSql, intI
   '申請人
   strSql = "DELETE FROM consultrecapp WHERE cra01='" & strCRL01 & "'"
   Pub_SeekTbLog strSql '記錄Log
   cnnConnection.Execute strSql, intI
   '發明人
   strSql = "DELETE FROM consultrecinv WHERE cri01='" & strCRL01 & "'"
   Pub_SeekTbLog strSql '記錄Log
   cnnConnection.Execute strSql, intI
   '商標圖檔:
   '刪除電子檔
   PUB_DelFtpFile2 strCRL01, , UCase("ConsultRecImageF") '檔案改放FTP,必須在DB資料刪除前執行
   '刪除DB資料
   strSql = "DELETE FROM consultrecimagef WHERE crif01='" & strCRL01 & "'"
   Pub_SeekTbLog strSql '記錄Log
   cnnConnection.Execute strSql, intI
   
   '附件
   strExc(0) = "select *" & _
               " from casepaperpdf WHERE cpp11='" & strCRL01 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         '刪除電子檔
         PUB_DelFtpFile2 RsTemp.Fields("cpp01"), " and cpp02='" & ChgSQL(RsTemp.Fields("cpp02")) & "'"
         '刪除DB資料
         strSql = "delete from casepaperpdf where cpp01='" & RsTemp.Fields("cpp01") & "' and cpp02='" & ChgSQL(RsTemp.Fields("cpp02")) & "'"
         Pub_SeekTbLog strSql '記錄Log
         cnnConnection.Execute strSql, intI
         RsTemp.MoveNext
      Loop
   End If
   
   '簽核資料
   strSql = "delete from flow003 where f0301='" & strCRL01 & "'"
   Pub_SeekTbLog strSql '記錄Log
   cnnConnection.Execute strSql, intI
   strSql = "delete from flow002 where f0201='" & strCRL01 & "'"
   Pub_SeekTbLog strSql '記錄Log
   cnnConnection.Execute strSql, intI
   strSql = "delete from flow004 where f0401='" & strCRL01 & "'"
   Pub_SeekTbLog strSql '記錄Log
   cnnConnection.Execute strSql, intI
   
   PUB_DelCRLAllData = True
   Exit Function
   
CheckingErr:
   If Err.Description <> "" Then MsgBox (Err.Description), , "PUB_DelCRLAllData"
   PUB_DelCRLAllData = False
End Function

'儲存接洽單附件
Public Function PUB_SaveCRLFile(m_strSaveFiles As String, m_strSaveFiles2 As String, _
   strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, strCRL01 As String, _
   bolNewCase As Boolean, Optional strCaseNA239 As String = "", _
   Optional ByVal bolDelOrgFile As Boolean = False) As Boolean
   
   PUB_SaveCRLFile = False
   
   Screen.MousePointer = vbHourglass
   '新案:附件CPP01=接洽單編號
   If bolNewCase = True Then
      'ex:1110003732.EX.PDF
      '   1110003732.016020772-P0-Q2-C0-R0-L0-E0-B0-S0-H0.pdf
      If m_strSaveFiles <> "" Then
         If PUB_UpdReplyFile(m_strSaveFiles, strCRL01, strCP01, "", "", "", , strCRL01, , "1") = False Then '1.官方文件
            Screen.MousePointer = vbDefault
            Exit Function
         Else
            PUB_SaveCRLFile = True
         End If
      End If
      If m_strSaveFiles2 <> "" Then
         If PUB_UpdReplyFile(m_strSaveFiles2, strCRL01, strCP01, "", "", "", , strCRL01, "case", "2") = False Then '2.內部文件
            Screen.MousePointer = vbDefault
            Exit Function
         Else
            PUB_SaveCRLFile = True
         End If
      End If
      
   '舊案:案件回覆單 ex:T123457.1110003734.Reply.pdf
   Else
      If m_strSaveFiles <> "" Then
         'Mofidied by Morgan 2020/12/21 脫歐英國新案放在歐盟案號+UK
         If strCaseNA239 <> "" Then
            If PUB_UpdReplyFile(m_strSaveFiles, strCaseNA239 & "UK", Left(strCaseNA239, 3), Mid(strCaseNA239, 4, 6), Mid(strCaseNA239, 10, 1), Mid(strCaseNA239, 11, 2), , strCRL01, , "3") = False Then
               Screen.MousePointer = vbDefault
               Exit Function
            Else
               PUB_SaveCRLFile = True
            End If
         Else
         'end 2020/12/21
   
            If PUB_UpdReplyFile(m_strSaveFiles, "", strCP01, strCP02, strCP03, strCP04, , strCRL01, EMP_回覆單, "3") = False Then
               Screen.MousePointer = vbDefault
               Exit Function
            Else
               PUB_SaveCRLFile = True
            End If
         End If
      End If
   End If
   
   If PUB_SaveCRLFile = True Then
      '刪除原始檔
      If bolDelOrgFile = True Then
         If m_strSaveFiles <> "" Then Call PUB_DelPCOrgFile(m_strSaveFiles)
         If m_strSaveFiles2 <> "" Then Call PUB_DelPCOrgFile(m_strSaveFiles2)
      End If
   End If
   Screen.MousePointer = vbDefault
End Function

'Added by Morgan 2021/8/18
'讀取使用者上次的視窗大小及位置
'Move by Lydia 2022/10/07 從frm100101_2_1搬過來
'Modified by Lydia 2022/10/24 是否調整表單大小bolResize
Public Sub PUB_SetPdfForm(pForm As Form, Optional ByVal bolResize As Boolean = True)
   Dim strWinPos As String, lngLeft As Long, lngTop As Long, lngWinPos() As String
   
   strWinPos = GetSetting("TAIE", strUserNum, pForm.Name)
   If strWinPos = "" Then
      MoveFormToCenter pForm
   'Added by Morgan 2022/8/19
   ElseIf strWinPos = "Max" Then
      pForm.WindowState = vbMaximized
   'end 2022/8/19
   Else
      lngWinPos = Split(strWinPos, ",")
      'Modified by Morgan 2022/7/27
      If bolResize = True Then 'Added by Lydia 2022/10/24 是否調整表單大小
         If (lngWinPos(2) < Screen.Width / 2 Or lngWinPos(2) > Screen.Width) Then lngWinPos(2) = Screen.Width / 2
         If (lngWinPos(3) < Screen.Height / 2 Or lngWinPos(3) > Screen.Height - 650) Then lngWinPos(3) = Screen.Height - 650
      End If 'Added by Lydia 2022/10/24
      If lngWinPos(0) < 0 Or lngWinPos(0) > Screen.Width Then lngWinPos(0) = 0
      If lngWinPos(1) < 0 Or lngWinPos(1) > Screen.Height Then lngWinPos(1) = 0
      'end 2022/7/27
      pForm.Move lngWinPos(0), lngWinPos(1), lngWinPos(2), lngWinPos(3)
   End If
End Sub

'Added by Morgan 2021/8/18
'紀錄視窗最後的大小及位置
'Move by Lydia 2022/10/07 從frm100101_2_1搬過來
Public Sub PUB_SavePdfForm(pForm As Form)
   'Added by Morgan 2022/8/19
   If pForm.WindowState = vbMaximized Then
      SaveSetting "TAIE", strUserNum, pForm.Name, "Max"
   Else
   'end 2022/8/19
      SaveSetting "TAIE", strUserNum, pForm.Name, pForm.Left & "," & pForm.Top & "," & pForm.Width & "," & pForm.Height
   End If
End Sub

'Add by Amy 2022/10/12
'確認並設定接洽單上一簽核人員-補件完成用
'stF0201:接洽單編號 /stSingnIdentity:簽核身份: A0:智權人員
'Modify By Sindy 2022/11/2 + Optional ByVal strEmp As String : 員工編號
Public Function SetConultRecPrePerson_Flow002(ByVal stFormN As String, ByVal stF0201 As String, ByVal stSingnIdentity As String, _
Optional ByVal strEmp As String) As Boolean
    Dim RsQ As New ADODB.Recordset, strQ As String, intQ As Integer
    Dim strSql As String, intMaxF0203 As Integer
    
    SetConultRecPrePerson_Flow002 = False
    '判斷簽核身份「無」待簽核資料,新增一筆
    strQ = "Select * From FLOW002 Where F0201='" & stF0201 & "' and F0202='" & stSingnIdentity & "' and F0207 is null "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 0 Then
        strQ = "Select nvl(Max(F0203),0) From FLOW002 Where F0201='" & stF0201 & "' and F0202='" & stSingnIdentity & "' "
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, strQ)
        If intQ = 1 Then
            intMaxF0203 = RsQ.Fields(0)
        End If
        '分案主管(A6)/程序(A7)
        intMaxF0203 = intMaxF0203 + 1
        strSql = "insert into FLOW002 (F0201,F0202,F0203,F0204) values(" & CNULL(stF0201) & ",'" & stSingnIdentity & "'," & intMaxF0203 & ",'" & IIf(strEmp <> "", strEmp, stSingnIdentity) & "')"
        cnnConnection.Execute strSql
    End If
    SetConultRecPrePerson_Flow002 = True
    Set RsQ = Nothing
End Function

'確認簽核身份是否可操作
'Modify by Amy 2022/10/17 +回傳 stF0207
'Modify by Amy 2022/11/15 +回傳訊息 stMsg
'stF0201:接洽單編號  / stSingnIdentity:簽核身份 / IsEConsultRec:是接洽單電子化
Public Function ChkConultRecFlow002(ByVal stFormN As String, ByVal stF0201 As String, ByVal stSingnIdentity As String, ByRef IsEConsultRec As Boolean, Optional ByRef stF0207 As String, Optional ByRef stMsg As String) As Boolean
    Dim RsQ As New ADODB.Recordset, strQ As String, intQ As Integer
    
    ChkConultRecFlow002 = False: IsEConsultRec = False
    stF0207 = "" 'Add by Amy 2022/10/17
    stMsg = "" 'Add by Amy 2022/11/15
    strQ = "Select F0201,F0207,cp09 From Flow002,CaseProgress " & _
              "Where F0201='" & stF0201 & "' And F0202='" & stSingnIdentity & "' And F0201=cp140(+) " & _
              "Order by F0203 Desc "
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, strQ)
    '接洽單電子化
    If intQ = 1 Then
        IsEConsultRec = True
        stF0207 = "" & RsQ.Fields("F0207") 'Add by Amy 2022/10/17
        '程序人員已處理
        If stSingnIdentity = "A7" Then
            ChkConultRecFlow002 = True
        '分案主管簽核結果需同意(F0207=1)
        ElseIf stSingnIdentity = "A6" And "" & RsQ.Fields("F0207") = "1" Then
            ChkConultRecFlow002 = True
        Else
            'Modify by Amy 2022/11/18  原:直接彈訊息
            stMsg = "此案尚未簽核完畢，不可分案！"
        End If
    '舊資料不會有資料
    Else
        ChkConultRecFlow002 = True
    End If
    Set RsQ = Nothing
End Function
'end 2022/10/12

'Add By Sindy 2022/9/7
'接洽單設定關連
Public Function PUB_SetCRLGroup(oGrid As MSHFlexGrid, intCRL01 As Integer, intCRL65 As Integer, ByRef bolUpd As Boolean, _
   intCRL06 As Integer, intCRL55 As Integer, intCRL74 As Integer, intCaseNo As Integer, intCust As Integer, intCaseName As Integer) As Boolean
Dim ii As Integer, intHadSel As Integer, jj As Integer
Dim strCRL65 As String, strCRL55 As String
Dim strCustNm As String, strCaseName As String
Dim bolAnsYes As Boolean, bolCustYes As Boolean, bolCaseNameYes As Boolean
   
   PUB_SetCRLGroup = False: bolUpd = False
   'intHadSel = 0
   For ii = 1 To oGrid.Rows - 1
      If Trim(oGrid.TextMatrix(ii, 0)) = "V" Then
         intHadSel = intHadSel + 1
         
'         '記錄系統別
'         If Trim(oGrid.TextMatrix(jj, intCaseNo)) <> "" Then
'            strSys = SystemNumber(Trim(oGrid.TextMatrix(jj, intCaseNo)), 1)
'         End If
         
         If Trim(oGrid.TextMatrix(ii, intCRL06)) <> "Y" Then
            MsgBox "新案才能設定關連！", vbExclamation, "設定關連"
            Exit Function
         Else
'            If Len(Trim(oGrid.TextMatrix(ii, intCRL74))) = 2 Then
'               MsgBox "(第" & ii & " 筆)案源無須手動設定關連！", vbExclamation, "設定關連"
'               Exit Function
'            ElseIf Trim(oGrid.TextMatrix(ii, intCRL74)) <> "" Then
'               MsgBox "(第" & ii & " 筆)特殊案件無須手動設定關連！", vbExclamation, "設定關連"
'               Exit Function
'            End If
         End If
         
         If intHadSel = 1 Then
            '檢查是否有輸入相同案號
            For jj = ii To oGrid.Rows - 1
               If Trim(oGrid.TextMatrix(jj, 0)) = "V" And Trim(oGrid.TextMatrix(jj, intCRL55)) <> "" Then
                  strCRL55 = Trim(oGrid.TextMatrix(jj, intCRL55))
                  Exit For
               End If
            Next jj
            For jj = ii + 1 To oGrid.Rows - 1
               '有輸入案號就全部人員自行輸入,或全部空白
               If Trim(oGrid.TextMatrix(jj, 0)) = "V" Then
                  If strCRL55 <> "" Then
                     If Trim(oGrid.TextMatrix(jj, intCRL55)) = "" Then
                        If Replace(Trim(oGrid.TextMatrix(jj, intCaseNo)), "-0-00", "") <> strCRL55 Then
                           MsgBox "您點選的資料列(第" & jj & " 筆)請自行輸入相同案號，不可空白！", vbExclamation, "設定關連"
                           Exit Function
                        End If
                     Else
                        If Trim(oGrid.TextMatrix(jj, intCRL55)) <> strCRL55 Then
                           MsgBox "您點選的資料列(第" & jj & " 筆)相同案號與(" & strCRL55 & ")不同，請重新確認！", vbExclamation, "設定關連"
                           Exit Function
                        End If
                     End If
                  End If
               End If
            Next jj
         
            '檢查是否有設定關連
            For jj = ii To oGrid.Rows - 1
               If Trim(oGrid.TextMatrix(jj, 0)) = "V" And Trim(oGrid.TextMatrix(jj, intCRL65)) <> "" Then
                  strCRL65 = Trim(oGrid.TextMatrix(jj, intCRL65))
                  Exit For
               End If
            Next jj
            For jj = ii + 1 To oGrid.Rows - 1
               '只能相同或空白
               If Trim(oGrid.TextMatrix(jj, 0)) = "V" Then
                  If strCRL65 <> "" Then
                     If Trim(oGrid.TextMatrix(jj, intCRL65)) <> strCRL65 Then
                        MsgBox "您點選的資料列(第" & jj & " 筆)已有設定關連，" & vbCrLf & vbCrLf & _
                               "若要重新設定，請先取消關連！", vbExclamation, "設定關連"
                        Exit Function
                     End If
                  End If
               End If
            Next jj
            
            '檢查客戶是否不同,彈提醒
            For jj = ii To oGrid.Rows - 1
               If Trim(oGrid.TextMatrix(jj, 0)) = "V" And Trim(oGrid.TextMatrix(jj, intCust)) <> "" Then
                  strCustNm = Trim(oGrid.TextMatrix(jj, intCust))
                  Exit For
               End If
            Next jj
            For jj = ii + 1 To oGrid.Rows - 1
               '不同彈提醒
               If Trim(oGrid.TextMatrix(jj, 0)) = "V" Then
                  If strCustNm <> "" Then
                     If Trim(oGrid.TextMatrix(jj, intCust)) <> strCustNm And bolCustYes = False Then
                        If MsgBox("客戶不同，確定設定關連嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
                           Exit Function
                        Else
                           bolCustYes = True
                           bolAnsYes = True
                        End If
                     End If
                  End If
               End If
            Next jj
            
            '檢查案件名稱是否不同,彈提醒
            For jj = ii To oGrid.Rows - 1
               If Trim(oGrid.TextMatrix(jj, 0)) = "V" And Trim(oGrid.TextMatrix(jj, intCaseName)) <> "" Then
                  strCaseName = Trim(oGrid.TextMatrix(jj, intCaseName))
                  Exit For
               End If
            Next jj
            For jj = ii + 1 To oGrid.Rows - 1
               '不同彈提醒
               If Trim(oGrid.TextMatrix(jj, 0)) = "V" Then
                  If strCaseName <> "" Then
                     If Trim(oGrid.TextMatrix(jj, intCaseName)) <> strCaseName And bolCaseNameYes = False Then
                        If MsgBox("案件名稱不同，確定設定關連嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
                           Exit Function
                        Else
                           bolCaseNameYes = True
                           bolAnsYes = True
                        End If
                     End If
                  End If
               End If
            Next jj
         End If
      End If
   Next ii
   If intHadSel < 2 Then
      MsgBox "請至少勾選2筆資料列，才能設定關連！", vbExclamation, "設定關連"
      Exit Function
   End If
   If bolAnsYes = False Then
      If MsgBox("確定設定關連嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbNo Then
         Exit Function
      End If
   End If
   For ii = 1 To oGrid.Rows - 1
      If Trim(oGrid.TextMatrix(ii, 0)) = "V" Then
         If strCRL65 = "" Then strCRL65 = Trim(oGrid.TextMatrix(ii, intCRL01))
         strSql = "UPDATE CONSULTRECORDLIST SET CRL65='" & strCRL65 & "' WHERE CRL01='" & Trim(oGrid.TextMatrix(ii, intCRL01)) & "'"
         cnnConnection.Execute strSql, intI
         bolUpd = True
      End If
   Next ii
   
   PUB_SetCRLGroup = True
End Function

'Add By Sindy 2022/9/7
'接洽單取消關連
Public Function PUB_CancelCRLGroup(oGrid As MSHFlexGrid, intCRL01 As Integer, intCRL65 As Integer, ByRef bolUpd As Boolean, _
   intCRL74 As Integer) As Boolean
Dim ii As Integer, intHadSel As Integer
   
   PUB_CancelCRLGroup = False: bolUpd = False
   intHadSel = 0
   For ii = 1 To oGrid.Rows - 1
      If Trim(oGrid.TextMatrix(ii, 0)) = "V" Then
         intHadSel = intHadSel + 1
         If Trim(oGrid.TextMatrix(ii, intCRL65)) = "" Then
            MsgBox "您點選的資料列並未設定關連，無需取消！", vbExclamation, "取消關連"
            Exit Function
         Else
'            If Len(Trim(oGrid.TextMatrix(ii, intCRL74))) = 2 Then
'               MsgBox "(第" & ii & " 筆)案源不能取消關連！", vbExclamation, "設定關連"
'               Exit Function
'            ElseIf Trim(oGrid.TextMatrix(ii, intCRL74)) <> "" Then
'               MsgBox "(第" & ii & " 筆)特殊案件不能取消關連！", vbExclamation, "設定關連"
'               Exit Function
'            End If
         End If
      End If
   Next ii
   If intHadSel = 0 Then
      MsgBox "請至少勾選1筆有設定關連的資料列，才能取消關連！", vbExclamation, "取消關連"
      Exit Function
   End If
   If MsgBox("確定取消關連嗎？", vbYesNo + vbCritical + vbDefaultButton2, "詢問") = vbYes Then
      For ii = 1 To oGrid.Rows - 1
         If Trim(oGrid.TextMatrix(ii, 0)) = "V" Then
            strSql = "UPDATE CONSULTRECORDLIST SET CRL65=null WHERE CRL65='" & Trim(oGrid.TextMatrix(ii, intCRL65)) & "'" & _
                     " and CRL01 in(select CRL01" & _
                     " From ConsultRecordList, flow003" & _
                     " where crl02>=" & 接洽單電子收文啟用日 & _
                     " and CRL01=f0301(+) and f0301 is not null and f0309 is null and crl78='" & strUserNum & "')"
            cnnConnection.Execute strSql, intI
            bolUpd = True
         End If
      Next ii
   End If
   PUB_CancelCRLGroup = True
End Function

'Add By Sindy 2022/12/27 檢查是否為接洽單第一筆案件性質
Public Function Pub_ConIsFirstCRC(strCRL01 As String, strCP09 As String) As Boolean
   Pub_ConIsFirstCRC = False
   strSql = "Select * From consultrecordlist Where CRL01='" & strCRL01 & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      strSql = "Select * From consultrecCMP Where CRC01='" & strCRL01 & "' and CRC08='" & strCP09 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.Fields("CRC02") = 1 Then
            Pub_ConIsFirstCRC = True '是第一筆
         End If
      End If
   Else
      '無接洽單均視為第一筆
      Pub_ConIsFirstCRC = True
   End If
End Function

'Add By Sindy 2024/1/18
'回傳值
'0=無主管機關
'1=有主管機關者
'2=已程序判發回來
'Modify By Sindy 2024/11/11 strEEP01 及 strEEP02 增加 Optional ByVal; 可不傳入
Public Function PUB_ChkhadCF10forEMP_46(strCP01 As String, strPA09 As String, strCP10 As String, _
   Optional ByVal strEEP01 As String = "", Optional ByVal strEEP02 As String = "") As Integer
   
   PUB_ChkhadCF10forEMP_46 = 0 '無主管機關
   'Add By Sindy 2023/11/9 外專要檢查台灣案發文前是否有需要送判程序主管
   If strPA09 = "000" Then
      strExc(0) = "select * from casefee where cf01='" & strCP01 & "' and cf02='" & strPA09 & "' and cf03='" & strCP10 & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         '有主管機關者
         If Trim("" & RsTemp.Fields("cf10")) <> "" Then
            PUB_ChkhadCF10forEMP_46 = 1
            'Add By Sindy 2024/11/11
            If strEEP01 <> "" And strEEP02 <> "" Then
            '2024/11/11 END
               strExc(0) = "select * from empelectronprocess where eep01='" & strEEP01 & "' and eep02=" & strEEP02
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp.Fields("EEP04") = EMP_送件 And InStr("" & RsTemp.Fields("EEP11"), "流程狀態:" & EMP_程序送判) > 0 Then
                     PUB_ChkhadCF10forEMP_46 = 2 '已程序判發回來
                  End If
               End If
            End If
         End If
      End If
   End If
   '2023/11/9 END
End Function

'Add By Sindy 2024/2/21 改成共用函數
'Modify By Sindy 2019/6/24
'Modify By Sindy 2024/10/23 Optional strNotCP10 As String = "": 排除的案件性質
Public Sub frm090801_New_SetComboCase(ii As Integer, strText As String, _
   ByRef Combo1 As Object, strCP01 As String, _
   Optional ByVal Combo1_0_Text As String, Optional ChkPCT As Object, _
   Optional Text1_101 As Object, Optional Text1_102 As Object, Optional Text1_103 As Object, _
   Optional strNotCP10 As String = "")
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strText1 As String 'Added by Morgan 2022/7/14
Dim arrCaseProperty, arrNation

   If ii = 0 Then '申請國家
      '申請國及發明人國籍
      'Modify By Sindy 2023/3/2 排除224.UP
      StrSQLa = "Select NA01, NA03 From Nation Where Length(NA01)=3 And NA01<='9999' And substr(NA02,3,1)='0'" & _
                " And NA01 not in('224')"
      'Add By Sindy 2019/7/9
      If Combo1.ListIndex >= 0 Then
         Combo1.Tag = Combo1.List(Combo1.ListIndex)
      End If
      If Combo1.Tag <> Combo1.Text Then
         StrSQLa = StrSQLa & " and (na01||' '||na03 like '%" & strText & "%')"
         'Added by Morgan 2022/7/14
         strText1 = PUB_ConvFullChar(strText)
         If strText1 <> strText Then
            StrSQLa = StrSQLa & " or na01||' '||na03 like '%" & strText1 & "%'"
         End If
         'end 2022/7/14
      End If
      StrSQLa = StrSQLa & " Order By NA01"
      '2019/7/9 END
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      Combo1.Clear
      While Not rsA.EOF
          Combo1.AddItem "" & rsA.Fields(0).Value & " " & rsA.Fields(1).Value
'          For i = 0 To 9
'             Me.Combo3(i).AddItem "" & rsA.Fields(0).Value & " " & rsA.Fields(1).Value
'          Next i
          rsA.MoveNext
      Wend
      
   '案件性質
   Else
      'Modify By Sindy 2019/7/22
      arrNation = Split(Combo1_0_Text, " ") '申請國家
      If UBound(arrNation) > 0 Then
      '2019/7/22 END
         'Modify By Cheng 2004/03/15
         '若系統類別為L, LA, FCL, CFL則案件性質代碼不要限制三碼及三碼以內
         'edit by nickc 2005/03/02 案件性質台灣抓 cpm03 其餘抓 cpm04
         'edit by nickc 2007/03/23  加入PCT 案件性質
         'Modified by Lydia 2020/11/06  ACS的4碼的10XX且小於103、12XX、19XX不出現。L、CFL、FCL、LIN、LA的4碼90XX不出現
         'StrSQLa = "Select CPM02||' '||" & IIf(ChkPCT.Value = vbChecked And ChkPCT.Enabled = True, "decode(cpm02,'101','發明進入國家','102','新型進入國家' ,", "") & "Decode('" & arrNation(0) & "', '000', CPM03,CPM04)" & IIf(ChkPCT.Value = vbChecked And ChkPCT.Enabled = True, ") ", "") & " as AA,cpm02 From CasePropertyMap Where CPM01='" & Me.Text1(6).Text & "' And Length(CPM02)<Decode(instr(CPM01,'L'), 0, 4, 5) And CPM01 Not In ('T','CFT','FCT') And Decode('" & arrNation(0) & "', '000', CPM03, CPM04)<>'（無）' "
         'Memo by Lydia 2020/11/06 注意修改條件要一併修改Combo1_Validate
         'Modify By Sindy 2024/10/21
         If TypeName(ChkPCT) <> "Nothing" Then
            StrSQLa = "Select CPM02||' '||" & IIf(ChkPCT.Value = vbChecked And ChkPCT.Enabled = True, "decode(cpm02,'101','發明進入國家','102','新型進入國家' ,", "")
            StrSQLa = StrSQLa & "Decode('" & arrNation(0) & "', '000', CPM03,CPM04)" & IIf(ChkPCT.Value = vbChecked And ChkPCT.Enabled = True, ") ", "") & " as AA,cpm02 "
         Else
            StrSQLa = "Select CPM02||' '||Decode('" & arrNation(0) & "', '000', CPM03,CPM04) as AA,cpm02 "
         End If
         '2024/10/21 END
         'Modify By Sindy 2024/12/6 + CPM30 is null And: 接洽單不顯示
         StrSQLa = StrSQLa & "From CasePropertyMap Where CPM30 is null And CPM01='" & strCP01 & "' And CPM01 Not In ('T','CFT','FCT') And Decode('" & arrNation(0) & "', '000', CPM03, CPM04)<>'（無）' "
         'Add By Sindy 2024/10/23 欲排除的案件性質
         If strNotCP10 <> "" Then
            StrSQLa = Trim(StrSQLa) & " and CPM02 not in(" & strNotCP10 & ")"
         End If
         '2024/10/23 END
         If strCP01 = "ACS" Then
              'Modified by Lydia 2020/11/23 改為ACS的4碼的100X、12XX、19XX不出現
              'StrSQLa = StrSQLa & "and not(length(cpm02)=4 and (cpm02 <'103' or cpm02 like '12%' or cpm02 like '19%')) "
              StrSQLa = StrSQLa & "and not(length(cpm02)=4 and (cpm02 like '100%' or cpm02 like '12%' or cpm02 like '19%')) "
         ElseIf InStr(strCP01, "L") > 0 Then '法務
              StrSQLa = StrSQLa & "and not(length(cpm02)=4 and cpm02 > '900') "
         Else  '其他：專利
              StrSQLa = StrSQLa & "and Length(CPM02)<Decode(instr(CPM01,'L'), 0, 4, 5)"
         End If
         'end 2020/11/06
         
         'Modify By Sindy 2016/5/11 取消查名轉換顯示[查名請改系統類別TS(內商)或S(外商)]字樣
         'StrSQLa = StrSQLa & " Union Select CPM02||' '||Decode('" & arrNation(0) & "', '000',cpm03,'020', Decode(CPM02,'001','查名請改系統類別TS(內商)或S(外商)',CPM04), Decode(CPM02,'001','查名請改系統類別TS(內商)或S(外商)',CPM04)),cpm02 From CasePropertyMap Where CPM01='" & Me.Text1(6).Text & "' And Length(CPM02)<4 And CPM01 In ('T','CFT','FCT') And Decode('" & arrNation(0) & "', '000',cpm03,'020', Decode(CPM02,'001','查名請改系統類別TS(內商)或S(外商)',CPM04), Decode(CPM02,'001','查名請改系統類別TS(內商)或S(外商)',CPM04))<>'（無）' "
         'Modify By Sindy 2024/12/6 + CPM30 is null And: 接洽單不顯示
         StrSQLa = StrSQLa & " Union Select CPM02||' '||Decode('" & arrNation(0) & "', '000',cpm03,CPM04) as AA,cpm02 From CasePropertyMap Where CPM30 is null And CPM01='" & strCP01 & "' And Length(CPM02)<4 And CPM01 In ('T','CFT','FCT') And Decode('" & arrNation(0) & "', '000',cpm03,CPM04)<>'（無）' "
         '2016/5/11 END
         'Add By Sindy 2024/10/23 欲排除的案件性質
         If strNotCP10 <> "" Then
            StrSQLa = Trim(StrSQLa) & " and CPM02 not in(" & strNotCP10 & ")"
         End If
         '2024/10/23 END
         'Added by Morgan 2012/12/25
         '20130101 起台灣專利案105聯合申請與125衍生設計申請交換下拉功能列出的順序
         If strSrvDate(1) >= "20130101" And strCP01 = "P" And arrNation(0) = "000" Then
             StrSQLa = "select * from (" & StrSQLa & ") Order By decode(cpm02,'105','125','125','105',cpm02)"
         Else
         'end 2012/12/25
             StrSQLa = StrSQLa & " Order By 1 "
         End If
         
         'Add By Sindy 2019/7/9
         If Combo1.ListIndex >= 0 Then
            Combo1.Tag = Combo1.List(Combo1.ListIndex)
         End If
         If Combo1.Tag <> Combo1.Text Then
            StrSQLa = "select * from (" & StrSQLa & ") where AA like '%" & strText & "%'"
         End If
         '2019/7/9 END
         
         rsA.CursorLocation = adUseClient
   '      'Add By Sindy 2019/6/24
   '      If strText <> "" Then
   '         rsA.Open "select * from (" & StrSQLa & ") where AA like '%" & strText & "%'", cnnConnection, adOpenStatic, adLockReadOnly
   '      Else
   '      '2019/6/24 END
            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   '      End If
         If rsA.RecordCount > 0 Then
      '       For ii = 1 To 4
                 Combo1.Clear
                 
                 'Added by Morgan 2020/5/4
                 If strText = "" Then
                     'Add By Sindy 2024/10/21
                     If TypeName(Text1_101) <> "Nothing" Then
                     '2024/10/21 END
                        Text1_101.Text = "" ' + (ii - 1) * 3
                        Text1_102.Text = "" ' + (ii - 1) * 3
                        Text1_103.Text = "" ' + (ii - 1) * 3
                     End If
                 End If
                 'end 2020/5/4
      '       Next ii
             While Not rsA.EOF
                '2010/1/6 917 超頁、超項費不再使用
                'Modify By Sindy 2010/8/5 閉卷,不續辦,取消收文,緩衝期限不顯示出來
                'Modified by Morgan 2012/12/25 有 "目前無" 字樣不列
                'Modify by Amy 2017/01/20 +法務、顧問案雖不限碼數, 但不可為 9001通知開庭,9002其他來函
                'Modify By Sindy 2023/1/17 排除陳述聲明
                'Modify By Sindy 2024/11/20 CFP要增加案件性質209檢視中說，但僅限於內部收文用，故接洽記錄單要剔除不列在下拉選單，也要檢查不能收文。
'Modify By Sindy 2024/12/6 已增加 SQL CPM30接洽單不顯示; 不需再逐一寫在程式裡了, 因此Mark
'                If "" & rsA.Fields(0).Value <> "917 超頁、超項費" And _
'                   "" & rsA.Fields(0).Value <> "913 閉卷" And _
'                   "" & rsA.Fields(0).Value <> "704 閉卷" And _
'                   "" & rsA.Fields(0).Value <> "907 不續辦" And _
'                   "" & rsA.Fields(0).Value <> "703 不續辦" And _
'                   "" & rsA.Fields(0).Value <> "925 取消收文" And _
'                   "" & rsA.Fields(0).Value <> "718 取消收文" And _
'                   "" & rsA.Fields(0).Value <> "312 緩衝期限" And _
'                   "" & rsA.Fields(0).Value <> "9001 通知開庭" And _
'                   "" & rsA.Fields(0).Value <> "9002 其他來函" And _
'                   InStr("" & rsA.Fields(0).Value, "目前無") = 0 And _
'                   "" & rsA.Fields(0).Value <> "214 陳述聲明" And _
'                   Not (strCP01 = "CFP" And "" & rsA.Fields(0).Value = "209 檢視中說") Then
               If InStr("" & rsA.Fields(0).Value, "目前無") = 0 Then
               '2024/12/6 END
                  'Add By Sindy 2024/2/21
                  strExc(10) = "" & rsA.Fields(0).Value
                  arrCaseProperty = Split(strExc(10), " ")
                  If UBound(arrCaseProperty) > 0 Then
                     If strCP01 = "CFT" Then
                        If arrNation(0) = "018" Or arrNation(0) = "022" Then '馬來西亞、汶萊
                           '馬來西亞、汶萊：不出現「304申請英文證明」
                           If arrCaseProperty(0) = "304" Then '申請英文證明
                              strExc(10) = ""
                           End If
                        Else
                           '非馬來西亞、汶萊：不出現「111英譯證明(CFT用) 」
                           If arrCaseProperty(0) = "111" Then '英譯證明(CFT用)
                              strExc(10) = ""
                           End If
                        End If
                        'Add By Sindy 2024/6/19 CFT增設案件性質，代號112，使用證據(申請)。
                        '   請於接洽記錄單控制僅CFT且為美國案方可收文
                        If arrCaseProperty(0) = "112" And arrNation(0) <> "101" Then
                           strExc(10) = ""
                        End If
                        '2024/6/19 END
                     End If
                     'Combo1.AddItem "" & rsA.Fields(0).Value
                     If strExc(10) <> "" Then
                        Combo1.AddItem strExc(10)
                     End If
                  End If
                  '2024/2/21 END
      '            Me.Combo1(2).AddItem "" & rsA.Fields(0).Value
      '            Me.Combo1(3).AddItem "" & rsA.Fields(0).Value
      '            Me.Combo1(4).AddItem "" & rsA.Fields(0).Value
                End If
                rsA.MoveNext
             Wend
         Else
      '       For ii = 1 To 4
                 Combo1.Clear
      '       Next ii
         End If
      End If
   End If
      
   'Add By Sindy 2019/6/24
   If strText <> "" Then
      If Combo1.ListCount = 1 Then
         Combo1.ListIndex = 0
         Combo1.Tag = Combo1.Text 'Add By Sindy 2019/7/9
      Else
         Combo1 = strText
'         'Add By Sindy 2019/7/9
'         If Combo1.ListIndex >= 0 Then
'            Combo1.Tag = Combo1.List(Combo1.ListIndex)
'         End If
'         '2019/7/9 END
      End If
   End If
   '2019/6/24 END
   
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Sub

'Added by Morgan 2024/6/3
'全E化客戶任一信箱異動時(含聯絡人)，發信通知智權人員
'pCustNo:客戶編號,pType:1=客戶 2=接洽人
Public Function PUB_ECustEmailChangeInform(pCustNo As String, pMsg As String, Optional pType As String = "1") As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stSubject As String, stContent As String, stTO As String, stCC As String
   
   stSQL = "select cu13,cu01||cu02||nvl(cu04,cu05) CName,st02||ac03 OMan" & _
    " from customer,SetSpecMan,staff,allcode where cu01='" & Left(pCustNo, 8) & "' and cu02='" & Mid(pCustNo, 9) & "' and cu176 is not null" & _
    " and ocode='全E化客戶維護人員' and st01(+)=oman and ac01(+)='01' and ac02(+)=st20"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      stSubject = "全E化客戶「" & rsQuery("CName") & "」" & IIf(pType = "2", "接洽人", "") & "連絡信箱異動通知!!"
      stContent = "此客戶已設定為全E化客戶，今因客戶的" & IIf(pType = "2", "接洽人", "") & "連絡信箱有異動，請確認全E化系統之客戶信箱資訊是否須同步進行異動，若須進行，請通知電腦中心" & rsQuery("OMan") & "進行全E化系統之客戶信箱異動。"
      stTO = rsQuery("cu13")
      If stTO <> strUserNum Then
         stCC = strUserNum
      End If
      
      stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
         " values( '" & strUserNum & "','" & stTO & "',to_char(sysdate,'yyyymmdd')" & _
         ",to_char(sysdate,'hh24miss'),'" & ChgSQL(stSubject) & "','" & ChgSQL(stContent) & "','" & stCC & "')"
      cnnConnection.Execute stSQL, intR
      pMsg = stContent
      PUB_ECustEmailChangeInform = True
   End If
   Set rsQuery = Nothing
End Function

'Added by Morgan 2024/6/20
'已設定不索取CF對帳單的代理人清單
Public Sub PUB_NoSoaList(pAutoBatch As Boolean, Optional ByVal pMailTo As String)
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim strSubject As String, stText As String, stTxtFilePath As String
   Dim ff As Integer, lngRecs As Long
   
On Error GoTo ErrHandle
   
   If pMailTo = "" Then pMailTo = Pub_GetSpecMan("外專請款單已收款通知人員")
   
   stSQL = "select FA01||FA02 YNo,trim(FA05||' '||FA04) YName from fagent where fa02='0' and fa133 is not null order by 1"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      strSubject = "已設定不索取CF對帳單的代理人清單"
      lngRecs = 0
         
      If ff > 0 Then Close #ff
      ff = FreeFile
      stTxtFilePath = App.path & "\$" & strSubject & ".txt"
      Open stTxtFilePath For Output As ff
      Print #ff, Space(25) & strSubject
      Print #ff, ""
      Print #ff, "日期：" & Format(strSrvDate(2), "###/##/##")
      Print #ff, ""
      Print #ff, "編號       名稱"
      Print #ff, "---------- ---------------------------------------------------------------"
      
      With rsQuery
      .MoveFirst
      Do While Not .EOF
         lngRecs = lngRecs + 1
         stText = ""
         stText = stText & convForm(.Fields("YNo"), 10)
         stText = stText & " " & .Fields("YName")
         Print #ff, stText
         .MoveNext
      Loop
      End With
      Print #ff, "---------- ---------------------------------------------------------------"
      Print #ff, "共" & lngRecs & "筆"
      Close ff
      ff = 0
         
      PUB_SendMail strUserNum, pMailTo, "", strSubject, "如旨", , stTxtFilePath, , , , , , , , True
   End If
   
ErrHandle:
   If Err.Number <> 0 Then
      If pAutoBatch Then
         WLog strSubject & ":" & Err.Description
      Else
         MsgBox Err.Description, vbCritical
      End If
   End If
   If ff > 0 Then Close #ff: ff = 0
   Set rsQuery = Nothing
End Sub

'Add by Amy 2025/04/18
'設定FC程序組人員下拉選單
Public Sub SetFCCombo(stFormN As String, objCbo As Object, strSysID As String, objLbl As Object)
   Dim RsQ As New ADODB.Recordset, intQ As Integer, stQ As String, i As Integer
   Dim pNoList As String, arrID As Variant
   
   Select Case UCase(stFormN)
      Case "FRM210149" '待處理區
   End Select
   
   If strSysID = "FCT" Then
      '檢查當時是否需要為他人職代
      'Mark by Amy 2025/07/23 改同待送件區-Sindy
      'Call Pub_SetForOthersEmpCombo(strUserNum, , False, pNoList)
      stQ = "And st15='F12' "
   ElseIf strSysID = "FCP" Then
      stQ = "And st15='F22' "
   End If
   
   If stQ <> "" Then
      stQ = "Select * From Staff Where st04='1' And st01<>'" & strUserNum & "' " & stQ & _
                  "Order by st01 asc"
       intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, stQ)
      If intQ = 1 Then
         RsQ.MoveFirst
         objCbo.AddItem strUserNum & " " & strUserName
         Do While RsQ.EOF = False
            objCbo.AddItem RsQ.Fields("st01") & " " & RsQ.Fields("st02")
            RsQ.MoveNext
         Loop
      End If
      Set RsQ = Nothing
   ElseIf Trim(pNoList) <> "" Then
      objCbo.AddItem strUserNum & " " & strUserName
      arrID = Split(pNoList, ";")
      For i = 0 To UBound(arrID)
         If Trim(arrID(i)) <> "" Then
            objCbo.AddItem arrID(i) & " " & GetPrjSalesNM(CStr(arrID(i)))
         End If
      Next i
   End If
  objCbo = strUserNum & " " & strUserName
End Sub


