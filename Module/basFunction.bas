Attribute VB_Name = "basFunction"
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

'Memo by Lydia 2023/03/03 從Service1移到basFunction
'Memo by Lydia 2022/08/30 從frm010005移過來
'Modified by Lydia 2024/02/20 +105: 有關" 香港013專利開放收文集體設計申請105"，請比照" 香港013專利收文設計申請103" --- Anny
Public Const FcpAddTct = "101,102,103,112,125,105" '要輸入追蹤流水號,並且走命名流程
'Modified by Lydia 2022/12/27 +935轉本所案件
'Modified by Lydia 2024/12/13 +605年費,803舉發 ---memo 增加非新案命名之性質，檢查Service1增加FCP/P案號時的系統通知Proc_FCPNewCaseEmail
Public Const AddTrackingNo = "101,102,103,105,110,112,125,307,935,605,803" '要輸入追蹤流水號
Public Const FcpNewCaseEmail = "307分割案,401變更,935案件轉至本所,605年費,803舉發" 'Added by Lydia 2024/12/13 增加FCP/P案號時的系統通知(非新案命名之性質)
'--------Memo by Lydia 2022/08/30 從frm010005移過來

Const 外商分信經理 As String = "80030" '洪琬姿 Add By Sindy 2021/6/17
'Modfied by Lydia 2023/03/10 TF_TCT=119=>120
Public Const TF_TCT As Integer = 120 'Added by Lydia 2023/02/16 外專命名記錄的欄位數
'Modified by Lydia 2025/01/17 +啟動認領日期TCT121,時間122
Public Const TF_TCTnotFS As String = "112,113,114,115,121,122"  'Added by Lydia 2023/02/16 外專命名記錄：排除不修改的欄位;
Public Const TCTforCP14 As String = "203,901,902,422,431"  'Added by Lydia 2023/05/10 外專新案命名一併收文的案件性質，當命名作業的主管確認自動掛承辦人=命名人員並且上已分案
Public Const FCPforEngNum As Integer = 2 'Added by Lydia 2024/02/27 外專工程師英文組數(2024/02/27 外專機械設計組人員異動調整程式：新案認領組別，請取消機械設計組，只留電子電機組及化學組)


'*************************************************
'  流水號之前補零
'
'*************************************************
Public Function ZeroBeforeNo(strInputValue As String, intInputLength As Integer) As String
Dim intCounter As Integer
   For intCounter = 1 To (intInputLength - Len(Trim(str(Val(strInputValue) + 1))))
      ZeroBeforeNo = ZeroBeforeNo & Mid("0", 1, 1)
   Next intCounter
   ZeroBeforeNo = ZeroBeforeNo & (Val(strInputValue) + 1)
End Function

'*************************************************
'  代理人名稱查詢
'
'*************************************************
Public Function FagentQuery(InputNo As String, InputSelect As Integer) As String
Dim adofagent As New ADODB.Recordset
   adofagent.CursorLocation = adUseClient
   adofagent.Open "select * from fagent where fa01 = '" & Mid(InputNo, 1, 8) & "' and fa02 = '" & Mid(InputNo, 9, 1) & "'", adoTaie, adOpenStatic, adLockReadOnly
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
            'edit by nickc 2007/02/08
            'If IsNull(adofagent.Fields("na04").Value) Then
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
'  檢核資料是否存在
'
'*************************************************
Public Function ExistCheck(strTable As String, strField As String, strValue As String, strError As String, Optional bolMsg As Boolean = True) As Boolean
Dim adocheck As New ADODB.Recordset

   adocheck.CursorLocation = adUseClient
   adocheck.Open "select " & strField & " from " & strTable & " where " & strField & " = " & "'" & strValue & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adocheck.RecordCount = 0 Then
      If bolMsg Then
         MsgBox MsgText(28) & strError, , MsgText(5)
      End If
      ExistCheck = False
   Else
      ExistCheck = True
   End If
   adocheck.Close
End Function

'*************************************************
'  計算天數
'
'*************************************************
Public Function CalculateDays(strStartDate As String, strEndDate As String) As Long
   CalculateDays = CDate(Mid(strEndDate, 1, 4) & "/" & Mid(strEndDate, 5, 2) & "/" & Mid(strEndDate, 7, 2)) - CDate(Mid(strStartDate, 1, 4) & "/" & Mid(strStartDate, 5, 2) & "/" & Mid(strStartDate, 7, 2))
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 由系統別, 國家代碼及案件性質取得工作天數
' Input : strSysKind  ==> 系統別
'         strNation   ==> 國家代碼
'         strCaseType ==> 案件性質
' Output : 傳回工作的天數
'          (若資料庫中未定義則傳回空白, 否則傳回工作的天數)
' 此功能會檢查案件(國家)收費表以帶出工作天數
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetWorkDays(ByVal strSysKind As String, ByVal strNation As String, ByVal strCaseType As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetWorkDays = ""
   
   'Modify by Morgan 2007/10/12 改先抓CaseFee1
   'strSQL = strSQL & "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & strSysKind & "' AND " & _
                  "CF02 = '" & strNation & "' AND " & _
                  "CF03 = '" & strCaseType & "' "
   strSql = "select cf105 as CF04,1 as oSort from casefee1" & _
         " where cf101='" & strSysKind & "' and cf102='" & strNation & "' and cf103='" & strCaseType & "'" & _
         " and cf104=(select max(cf104) from casefee1 where cf101='" & strSysKind & "' and cf102='" & strNation & "' and cf103='" & strCaseType & "' and cf104<=" & strSrvDate(1) & ") "
   strSql = strSql & " union SELECT CF04,2 FROM CaseFee " & _
            "WHERE CF01 = '" & strSysKind & "' AND " & _
                  "CF02 = '" & strNation & "' AND " & _
                  "CF03 = '" & strCaseType & "' "
   'end 2007/10/12
   
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CF04")) = False Then
         If IsEmptyText(rsTmp.Fields("CF04")) = False Then
            GetWorkDays = rsTmp.Fields("CF04")
         End If
      End If
   '92.7.4 ADD BY SONIA
   Else
      GetWorkDays = ""
   '92.7.4 END
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Function

'2009/10/15 add by sonia
' 取得FMP專利的規費
' Input : strPA01 ==> 系統類別
'         strCP10 ==> 案件性質
'         strCP07 ==> 法定期限
'         strPA08 ==> 專利種類
'         strPA09 ==> 申請國家
'         strPA16 ==> 目前准駁
Public Function GetFMPOfficialFee(ByVal strPA01 As String, _
                                     ByVal strCP10 As String, _
                                     ByVal strPA09 As String) As String
Dim rsTmp As ADODB.Recordset
Dim strSql As String
Dim strFee As String

   strFee = Empty
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM CASEFEE " & _
            "WHERE CF01 = '" & strPA01 & "' AND " & _
                  "CF02 = '" & strPA09 & "' AND " & _
                  "CF03 = '" & strCP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If Not IsNull(rsTmp.Fields("CF08")) Then
         strFee = rsTmp.Fields("CF08")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   GetFMPOfficialFee = strFee
   
End Function
'2009/10/15 end

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得專利的規費
' Input : strPA01 ==> 系統類別
'         strCP10 ==> 案件性質
'         strCP07 ==> 法定期限
'         strPA08 ==> 專利種類
'         strPA09 ==> 申請國家
'         strPA16 ==> 目前准駁
'911202 nick
'         strPa14 ==> 公告日
'2009/12/30
'         strPa02 ==> 本所案號
'         strPa03 ==> 本所案號
'         strPa04 ==> 本所案號
' 說明 : FCP的申請國家一定是台灣
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modified by Lydia 2017/12/08 是否電子送件 strCP118
Public Function GetPatentOfficialFee(ByVal strPA01 As String, _
                                     ByVal strCP10 As String, _
                                     ByVal strCP07 As String, _
                                     ByVal strPA08 As String, _
                                     ByVal strPA09 As String, _
                                     ByVal strPA16 As String, _
                                     Optional ByVal strPA14 As String = "", _
                                     Optional ByVal strPA02 As String = "", _
                                     Optional ByVal strPA03 As String = "", _
                                     Optional ByVal strPA04 As String = "", _
                                     Optional ByVal strCP118 As String = "") As String
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Dim strFee As String
   
   strFee = Empty
   ' FCP或FG需查CASEFEE檔案
   '92.11.3 MODIFY BY SONIA
   'If strPA01 = "FCP" Then
   If strPA01 = "FCP" Or strPA01 = "FG" Then
   '92.11.3 END
      Set rsTmp = New ADODB.Recordset
      '92.11.3 MODIFY BY SONIA
      'strSQL = "SELECT * FROM CASEFEE " & _
      '         "WHERE CF01 = 'FCP' AND " & _
      '               "CF02 = '000' AND " & _
      '               "CF03 = '" & StrCp10 & "' "
      strSql = "SELECT * FROM CASEFEE " & _
               "WHERE CF01 = '" & strPA01 & "' AND " & _
                     "CF02 = '000' AND " & _
                     "CF03 = '" & strCP10 & "' "
      '92.11.3 END
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         If Not IsNull(rsTmp.Fields("CF08")) Then
            strFee = rsTmp.Fields("CF08")
         End If
      End If
      rsTmp.Close
      Set rsTmp = Nothing
   End If
   
   ' 針對特殊的狀況再做特別的處理
   Select Case strCP10
      ' 追加申請
      Case "104":
         Select Case strPA08
            ' 發明
            Case "1": strFee = "2000"
            ' 新型
            Case "2": strFee = "4500"
         End Select
      ' 再審申請
      Case "107":
         Select Case strPA08
            ' 發明 93.7.1 6000->8000
            'Removed by Morgan 2013/1/10 102.1.1起改 7000 + 超頁超項費(抓 CaseFee)
            'Case "1": strFee = "8000"
            'Added by Morgan 2013/4/21 要檢查只能收7000,若有超頁超項費則須另外收文(CaseFee 改回 3500 )
            Case "1": strFee = "7000"
            
            ' 新型 93.7.1 4500->0
            Case "2": strFee = "0"
            ' 設計
            Case "3": strFee = "3500"
            
         End Select
         'add by sonia 2020/3/2 郭經理說P台灣案一律不檢查
         'Removed by Morgan 2020/3/17 先取消,等確認 P-121731,P-121221後續處理模式後再決定如何管控--秀玲
         'If strPA01 = "P" And strPA09 = "000" Then
         '   strFee = 0
         'End If
         'end 2020/3/17
         'end 2020/3/2
      ' 改請獨立
      Case "306":
         Select Case strPA08
            ' 發明
            Case "1": strFee = "2000"
            ' 新型
            Case "2": strFee = "4500"
            ' 設計
            Case "3": strFee = "3000"
         End Select
      ' 分割
      Case "307":
         Select Case strPA08
            ' 發明 93.7.1 2000->3500
            Case "1": strFee = "3500"
            ' 新型 93.7.1 4500->3000
            Case "2": strFee = "3000"
            ' 設計
            Case "3": strFee = "3000"
         End Select
      ' 變更  MODIFY BY SONIA 91.10.2 改由案件性質控制
      'Case "401":
      '   If IsEmptyText(strPA16) Then
      '      strFee = "300"
      '   Else
      '      strFee = "1000"
      '   End If
      ' 領證及繳年費  MODIFY BY SONIA 91.10.22 改在分案控制
      'Case "601":
         ' 系統日期小於等於法定期限
      '   If Val(DBDATE(SystemDate())) <= Val(DBDATE(strCP07)) Then
      '      strFee = "5000"
      '   Else
      '      If strPA01 = "FCP" Then strFee = strFee * 2
      '   End If
      ' 年費  MODIFY BY SONIA 91.10.22 改在分案控制
      'Case "605":
         ' 系統日期小於等於法定期限
      '   If Val(DBDATE(SystemDate())) <= Val(DBDATE(strCP07)) Then
      '      strFee = "2500"
      '   Else
      '      If strPA01 = "FCP" Then strFee = strFee * 2
      '   End If
      
'Modify by Morgan 2007/1/9 改由案件性質控制
'      ' 讓與
'      Case "701", "708":
'         ' 系統日期小於等於公告日加三個月
'         '911021 NICK 邱小姐說公告日空白不做
'         '911202 nick 修正
'         'If strCP07 <> "" Then
'         '   If Val(DBDATE(SystemDate())) <= Val(DBDATE(AddMonth(strCP07, 3))) Then
'         If strPa14 <> "" Then
'            If Val(DBDATE(SystemDate())) <= Val(DBDATE(AddMonth(strPa14, 3))) Then
'               strFee = "2000"
'            Else
'               '93.7.1 3500->2000
'               strFee = "2000"
'            End If
'         End If
'         '92.8.6 ADD BY SONIA
'         If strPA01 = "P" Then
'            Select Case StrCp10
'               Case "701"
'                  strFee = "2000"
'               Case "708"
'                  '93.7.1 3500->2000
'                  strFee = "2000"
'            End Select
'         End If
'         '92.8.6 END
'end 2007/1/9

'2010/8/25 CANCEL BY SONIA 此為舊法控制
'      ' 合併
'      Case "702":
'         ' 系統日期小於等於公告日加三個月
'         '911021 NICK 邱小姐說公告日空白不做
'         '911202 nick 修正
'         'If strCP07 <> "" Then
'         '   If Val(DBDATE(SystemDate())) <= Val(DBDATE(AddMonth(strCP07, 3))) Then
'         If strPa14 <> "" Then
'            If Val(DBDATE(SystemDate())) <= Val(DBDATE(AddMonth(strPa14, 3))) Then
'               strFee = "2000"
'            Else
'               strFee = "3500"
'            End If
'         End If
'
'      ' 繼承
'      Case "703":
'         ' 系統日期小於等於公告日加三個月
'         '911021 NICK 邱小姐說公告日空白不做
'         '911202 nick 修正
'         'If strCP07 <> "" Then
'         '   If Val(DBDATE(SystemDate())) <= Val(DBDATE(AddMonth(strCP07, 3))) Then
'         If strPa14 <> "" Then
'            If Val(DBDATE(SystemDate())) <= Val(DBDATE(AddMonth(strPa14, 3))) Then
'               strFee = "2000"
'            Else
'               strFee = "3500"
'            End If
'         End If
'2010/8/25 END
         
      ' 異議
      Case "801":
         Select Case strPA08
            ' 發明
            Case "1": strFee = "6000"
            ' 新型
            Case "2": strFee = "4500"
            ' 設計
            Case "3": strFee = "3500"
         End Select
      
'Removed by Morgan 2013/1/17 102新法規費有改,只在收費表設最低價,由接洽單輸入時控制
'      ' 舉發
'      Case "803"
'         Select Case strPA08
'            ' 發明 93.7.1 9000->10000
'            Case "1": strFee = "10000"
'            ' 新型 93.7.1 8500->9000
'            Case "2": strFee = "9000"
'            ' 設計
'            Case "3": strFee = "8000"
'         End Select
'end 2013/1/17

      '2009/12/30 add by sonia 台灣實審990101起調整規費
      ' 實體審查
      Case "416"
         If (strPA09 = "000" Or strPA09 = "") Then
            '先判斷此案是否為台灣99年新規費發明案件
            If Chk99NewCase(strPA01, strPA02, strPA03, strPA04) = False Then
               strFee = "8000"
            Else
               strFee = "7000"
            End If
         End If
      '2009/12/30 end
      'add by sonia 2019/5/8 FCP新型案之更正402規費1000,僅限卷宗性質為申請者,舉發案仍為2000
      Case "402"
         If strSrvDate(1) < 20191101 Then 'Added by Morgan 2019/10/30 108.11.1專利新法更正402規費改都為2000
            Set rsTmp = New ADODB.Recordset
            strSql = "SELECT PA23,PA08 FROM PATENT WHERE PA01='" & strPA01 & "' AND PA02='" & strPA02 & "' AND PA03='" & strPA03 & "' AND PA04='" & strPA04 & "' " & _
                     " AND PA01='FCP' AND PA23='1' AND PA08='2'"
            rsTmp.CursorLocation = adUseClient
            rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If rsTmp.RecordCount > 0 Then
               strFee = "1000"
            End If
            rsTmp.Close
            Set rsTmp = Nothing
         End If 'Added by Morgan 2019/10/30
      'end 2019/5/8
   End Select
   'Added by Lydia 2017/12/08 電子送件的規費-600
   'Modified by Lydia 2017/12/12 +參考P案分割307發文,不可減免 ( And strCP10 <> "307")
   'Modified by Lydia 2108/04/03 限申請案才有電子送件減免
   'If strCP118 = "Y" And Val(strFee) > 600 And strCP10 <> "307" Then
   If (strPA09 = "000" Or strPA09 = "") And strCP118 = "Y" And Val(strFee) > 600 And (InStr("101,102,103,125", strCP10) > 0 Or Left(strCP10, 1) = "3") Then
       strFee = Val(strFee) - 600
   End If
   'end 2017/12/08
   
   GetPatentOfficialFee = strFee
   '2005/6/14 ADD BY SONIA
'edit by nickc 2006/07/11 沒規費阿妙要拿掉
'   If GetPatentOfficialFee = "" Then
'      GetPatentOfficialFee = 0
'   End If
   '2005/6/14 END
End Function


'Move by Lydia 2023/03/22 從basQuery搬過來
Public Function ClsLawGetRelation(ByRef strlc As String, ByRef strRelcp09 As String, ByRef strRelation As String) As Boolean
 Dim RsTemp As New ADODB.Recordset, strQty As String
  
On Error GoTo ErrHand
   'Modified by Lydia 2023/03/22 法律案(L,FCL)不同審級收文期限沖銷：案件不同審級會收文-1、-2…案號，但實際是同一案件。
   'strQty = "select cp09 from caseprogress where " & ChgCaseprogress(strlc) & " and cp09='" + strRelation + "'"
   Dim strCase(1 To 4) As String
   Call ChgCaseNo(strlc, strCase)
   strQty = "select cp09 from caseprogress where cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' "
   If strCase(1) = "L" Or strCase(1) = "FCL" Then
   Else
      strQty = strQty & " and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "' "
   End If
   strQty = strQty & " and cp09='" & strRelation & "' "
   'end 2023/03/22
   
   'add by nickc 2007/04/18
   RsTemp.CursorLocation = adUseClient
   
   RsTemp.Open strQty, cnnConnection
   ClsLawGetRelation = False
   Do While Not RsTemp.EOF
      ClsLawGetRelation = True
      Exit Do
   Loop
   RsTemp.Close
   If ClsLawGetRelation = False Then MsgBox "非此案之其他收文號", vbCritical
   Exit Function
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'Move by Lydia 2023/03/22 從basQuery搬過來
'cp(1),cp(2),cp(3),cp(4) 本所案號,cp(5) 收據號碼,cp(6) 客戶編號,cp(7) 收據抬頭,911118 nick 新增原來申請人,cp(8) 原來申請人
Public Function ClsLawUpdAcc0k0(ByRef cp() As String, Optional bolUpd As Boolean = False, _
   Optional frm040104_9 As Boolean = False) As Boolean
 'Dim RsTemp As New ADODB.Recordset 'Removed by Morgan 2021/2/22
 Dim strQty As String
 
On Error GoTo ErrHand
   ClsLawUpdAcc0k0 = False
   cp(6) = ChangeCustomerL(cp(6))
   
'Removed by Morgan 2021/2/22 不會發生(原控制應為針對舊系統資料),取消--秀玲
'   '911118 nick 加入  若同一收據，但申請人不同，有兩筆以上時，不能修改
'   'StrQty = "SELECT COUNT(DISTINCT A0J02) FROM ACC0J0 WHERE A0J13='" & cp(5) & "'"
'   strQty = "SELECT COUNT(DISTINCT A0J02) FROM ACC0J0 WHERE A0J13='" & cp(5) & "' and a0j11<>'" & cp(8) & "' "
'   'add by nickc 2007/04/18
'   RsTemp.CursorLocation = adUseClient
'
'   RsTemp.Open strQty, cnnConnection
'   If RsTemp.Fields(0) > 1 Then
'      MsgBox "同一收據有其他案號收文資料，不可修改申請人 !", vbInformation
'   Else
'end 2021/2/22

      If bolUpd Then
         '92.3.9 cancel by sonia 都不更新收據抬頭
         'If Not frm040104_9 Then
         '   StrQty = "UPDATE ACC0K0 SET A0K03=" & CNULL(cp(6)) & ",A0K04=" & CNULL(cp(7)) & _
         '      " WHERE A0K01=" & CNULL(cp(5))
         'Else
            'frm040104_9 不更新收據抬頭
            strQty = "UPDATE ACC0K0 SET A0K03=" & CNULL(cp(6)) & _
               " WHERE A0K01=" & CNULL(cp(5))
         'End If
         '92.3.9 end
         cnnConnection.Execute strQty
         
         strQty = "UPDATE ACC0J0 SET A0J11=" & CNULL(cp(6)) & " WHERE A0J13='" & cp(5) & "'"
         cnnConnection.Execute strQty
      End If
      ClsLawUpdAcc0k0 = True
      
'Removed by Morgan 2021/2/22
'   End If
'   RsTemp.Close
'end 2021/2/22

   Exit Function
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'2009/10/15 add by sonia
' 取得FMP的費用
Public Function GetFMPFee(ByVal strPA01 As String, ByVal strCP10 As String, ByVal strPA09 As String) As String
Dim rsTmp As ADODB.Recordset
Dim strSql As String
   
   Set rsTmp = New ADODB.Recordset
   GetFMPFee = Empty
   strSql = "SELECT * FROM CASEFEE " & _
            "WHERE CF01 = '" & strPA01 & "' AND " & _
                  "CF02 = '" & strPA09 & "' AND " & _
                  "CF03 = '" & strCP10 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If Not IsNull(rsTmp.Fields("CF06")) Then
         GetFMPFee = rsTmp.Fields("CF06")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function
'2009/10/15 end

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得FCP的費用
' 說明 : FCP的申請國家一定是台灣
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'92.11.3 MODIFY BY SONIA
'Public Function GetFCPFee(ByVal StrCp10 As String) As String
Public Function GetFCPFee(ByVal strPA01 As String, ByVal strCP10 As String) As String
'92.11.3 END
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   Set rsTmp = New ADODB.Recordset
   GetFCPFee = Empty
   '92.11.3 MODIFY BY SONIA
   'strSQL = "SELECT * FROM CASEFEE " & _
   '         "WHERE CF01 = 'FCP' AND " & _
   '               "CF02 = '000' AND " & _
   '               "CF03 = '" & StrCp10 & "' "
   strSql = "SELECT * FROM CASEFEE " & _
            "WHERE CF01 = '" & strPA01 & "' AND " & _
                  "CF02 = '000' AND " & _
                  "CF03 = '" & strCP10 & "' "
   '92.11.3 END
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If Not IsNull(rsTmp.Fields("CF06")) Then
         GetFCPFee = rsTmp.Fields("CF06")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得商標的規費
' Input : strTM01 ==> 系統類別
'         strCP10 ==> 案件性質
'         strCP07 ==> 法定期限
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetTrademarkOfficialFee(ByVal strTM01 As String, _
                                        ByVal strCP10 As String, _
                                        ByVal strCP07 As String)
   Dim strFee As String
   
   strFee = Empty
   ' 針對特殊的狀況再做特別的處理
   Select Case strCP10
      ' 延展
      Case "102":
         strFee = "4000"
         'Add By Sindy 2011/3/14 若系統日的昨天為非工作天, 則計算出系統日的前一個工作天
         If ChkWorkDay(DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1))))) = False Then
            If Val(CompWorkDay(1, DBDATE(DateAdd("d", -1, ChangeWStringToWDateString(strSrvDate(1)))), 1)) > Val(DBDATE(strCP07)) And strCP07 <> "" Then
               strFee = "8000"
            End If
         '2011/3/14 End
         Else
            If Val(DBDATE(SystemDate())) > Val(DBDATE(strCP07)) And strCP07 <> "" Then
               strFee = "8000"
            End If
         End If
   End Select
   GetTrademarkOfficialFee = strFee
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 由客戶代碼取得客戶名稱
' Input : strCustomer ==> 客戶代碼
' Output : 傳回客戶的國籍
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetCustomerNation(ByVal strCustomer As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   GetCustomerNation = Empty
   
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
      If IsNull(rsTmp.Fields("CU10")) = False Then
         GetCustomerNation = rsTmp.Fields("CU10")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得民國的年份
' Input : strDate ==> 輸入的日期
' Output : 傳回民國的年份 YY
' Description : 此功能會傳回日期字串中的月份, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TAIWANYEAR(ByVal strDate As String) As String
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
Public Function TAIWANMONTH(ByVal strDate As String) As String
   TAIWANMONTH = DBMONTH(strDate)
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得民國的日數
' Input : strDate ==> 輸入的日期
' Output : 傳回民國的日數 DD
' Description : 此功能會傳回日期字串中的日數, 不管輸入的日期
'   是西元日期還是民國日期, 或是字串中有/的字元, 均會自動轉換
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function TAIWANDAY(ByVal strDate As String) As String
   TAIWANDAY = DBDAY(strDate)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得國家的延展年度
' Input : strNation ==> 國家代碼
' Output : 傳回國家的延展年度
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetNationExtentYear(ByVal strNation As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetNationExtentYear = "0"
   
   strSql = "SELECT * FROM NATION " & _
            "WHERE NA01 = '" & strNation & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("NA14")) = False Then
         GetNationExtentYear = rsTmp.Fields("NA14")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 檢查此客戶是否為個人
' Input : strCustomer == 客戶編號
' Output : 若此客戶是個人則傳回True, 否則傳回False
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsCustomerIndividual(ByVal strCustomer As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strKey As String
   Dim strSql As String
   
   'Modify By Sindy 2012/5/24
   'IsCustomerIndividual = True
   IsCustomerIndividual = False
   
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
      If IsNull(rsTmp.Fields("CU15")) = False Then
         'Modify By Sindy 2012/5/24
         'If rsTmp.Fields("CU15") = "1" Then
         If rsTmp.Fields("CU15") = "0" Then
            'Modify By Sindy 2012/5/24
            'IsCustomerIndividual = False
            IsCustomerIndividual = True
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得催審期限的日期
' Input : strSysKind  ==> 系統類別
'         strNation   ==> 申請國家代碼
'         strCaseType ==> 案件性質代碼
'         strDate     ==> 發文日
' Output : 傳回計算出來的催審期限日期(西元日期)
' 說明 : 若資料庫中未定義該系統類別及國別的審查時間, 則會傳回空白
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetUrgeDate(ByVal strSysKind As String, ByVal strNation As String, ByVal strCaseType As String, ByVal strDate As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   GetUrgeDate = Empty
   
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & strSysKind & "' AND " & _
                  "CF02 = '" & strNation & "' AND " & _
                  "CF03 = '" & strCaseType & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields("CF05")) = False Then
        'Modify By Cheng 2003/09/02
'         GetUrgeDate = DBDATE(DateSerial(Val(DBYEAR(strDate)), Val(DBMONTH(strDate)), Val(DBDAY(strDate)) + Val(rsTmp.Fields("CF05"))))
         GetUrgeDate = DBDATE(DateAdd("d", Val(rsTmp.Fields("CF05")), ChangeWStringToWDateString(DBDATE(strDate))))
         'Added by Lydia 2025/11/12 改抓最近工作天
         GetUrgeDate = PUB_GetWorkDay1(GetUrgeDate, True)
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得相關人
' Input : strCP09 ==> 總收文號
' Output : 傳回與此收文號相關的相關人
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetRelatedPerson(ByVal strCP09 As String) As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim bDeal As Boolean
   
   GetRelatedPerson = Empty
   strSql = "SELECT * FROM CaseProgress " & _
            "WHERE CP09 = '" & strCP09 & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      bDeal = False
      If bDeal = False And IsNull(rsTmp.Fields("CP40")) = False Then
         If IsEmptyText(rsTmp.Fields("CP40")) = False Then
            GetRelatedPerson = rsTmp.Fields("CP40")
            bDeal = True
         End If
      End If
      If bDeal = False And IsNull(rsTmp.Fields("CP41")) = False Then
         If IsEmptyText(rsTmp.Fields("CP41")) = False Then
            GetRelatedPerson = rsTmp.Fields("CP41")
            bDeal = True
         End If
      End If
      If bDeal = False And IsNull(rsTmp.Fields("CP42")) = False Then
         If IsEmptyText(rsTmp.Fields("CP42")) = False Then
            GetRelatedPerson = rsTmp.Fields("CP42")
            bDeal = True
         End If
      End If
      If bDeal = False And IsNull(rsTmp.Fields("CP50")) = False Then
         If IsEmptyText(rsTmp.Fields("CP50")) = False Then
            GetRelatedPerson = rsTmp.Fields("CP50")
            bDeal = True
         End If
      End If
      If bDeal = False And IsNull(rsTmp.Fields("CP51")) = False Then
         If IsEmptyText(rsTmp.Fields("CP51")) = False Then
            GetRelatedPerson = rsTmp.Fields("CP51")
            bDeal = True
         End If
      End If
      If bDeal = False And IsNull(rsTmp.Fields("CP52")) = False Then
         If IsEmptyText(rsTmp.Fields("CP52")) = False Then
            GetRelatedPerson = rsTmp.Fields("CP52")
            bDeal = True
         End If
      End If
      If bDeal = False And IsNull(rsTmp.Fields("CP56")) = False Then
         If IsEmptyText(rsTmp.Fields("CP56")) = False Then
            GetRelatedPerson = GetCustomerName(rsTmp.Fields("CP56"))
            bDeal = True
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 檢查來函記錄檔中是否存在該筆無期限的記錄
' Input : strTM01 ==> 本所案號中的系統別
'         strTM02 ==> 本所案號中的流水號
'         strTM03 ==> 本所案號中的追加案號
'         strTM04 ==> 本所案號中的多國多類碼
'         strDate ==> 收件日(來函收文日)
' Output : 傳回該筆無期限的記錄是否存在
'          True  : 存在
'          False : 不存在
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsMailRecNoTermExist(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String, ByVal strDate As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   ' 預設不存在
   IsMailRecNoTermExist = False
   
   strSql = "SELECT * FROM MailRec " & _
            "WHERE MR12 = '" & strTM01 & "' AND " & _
                  "MR13 = '" & strTM02 & "' AND " & _
                  "MR14 = '" & strTM03 & "' AND " & _
                  "MR15 = '" & strTM04 & "' AND " & _
                  "MR02 = " & DBDATE(strDate) & " AND " & _
                  "(MR16 is NULL OR MR16 = 0)"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsMailRecNoTermExist = True
   End If
   rsTmp.Close

   Set rsTmp = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得來函記錄檔的欄位
' Input : strTM01 ==> 本所案號中的系統別
'         strTM02 ==> 本所案號中的流水號
'         strTM03 ==> 本所案號中的追加案號
'         strTM04 ==> 本所案號中的多國多類碼
'         strDate ==> 收件日(來函收文日)
'         strField ==> 欄位名稱
' Output : 傳回該筆記錄的欄位內容
'          (若有存在此筆記錄則傳回其欄位的內容, 否則傳回空白)
' 說明 : 此函式回搜尋來函記錄檔MailRec以找尋符合條件的資料
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetMailRecField(ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String, ByVal strDate As String, ByVal strField As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetMailRecField = Empty
   
   strSql = "SELECT * FROM MailRec " & _
            "WHERE MR12 = '" & strTM01 & "' AND " & _
                  "MR13 = '" & strTM02 & "' AND " & _
                  "MR14 = '" & strTM03 & "' AND " & _
                  "MR15 = '" & strTM04 & "' AND " & _
                  "MR02 = " & DBDATE(strDate) & " "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      If IsNull(rsTmp.Fields(strField)) = False Then
         GetMailRecField = rsTmp.Fields(strField)
      End If
   End If
   rsTmp.Close
   
   Set rsTmp = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 由系統別,國家代碼及案件性質取得下一程序
' Input : strSysKind  ==> 系統別
'         strNation   ==> 國家代碼
'         strCaseType ==> 案件性質
' Output : 傳回案件(國家)收費表的下一救濟程序代號
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modified by Lydia 2020/11/05 +是否自動內部收文下一程序strCF30
Public Function GetNextProgress(ByVal strSysKind As String, ByVal strNation As String, ByVal strCaseType As String, Optional ByRef strCF30 As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   strCF30 = "" 'Added by Lydia 2020/11/05
   GetNextProgress = Empty
   strSql = "SELECT * FROM CaseFee " & _
            "WHERE CF01 = '" & strSysKind & "' AND " & _
                  "CF02 = '" & strNation & "' AND " & _
                  "CF03 = '" & strCaseType & "'"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 下一救濟程序
      If IsNull(rsTmp.Fields("CF15")) = False Then
         GetNextProgress = rsTmp.Fields("CF15")
         strCF30 = "" & rsTmp.Fields("CF30") 'Added by Lydia 2020/11/05
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 檢查條碼公式是否正確
' Input : strData = 條碼的內容
' Output : True = 表示條碼的內容是正確的
'          False = 條碼的內容是不正確的
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsBarcodeCorrect(ByVal strData As String) As Boolean
   Dim nSumEven As Integer
   Dim nSumOdd As Integer
   Dim nSum As Integer
   Dim nIndex As Integer
   Dim nRest As Integer
   Dim nCheck As Integer
   
   IsBarcodeCorrect = False
   
   ' 條碼必須全為數值
   If IsNumeric(strData) = False Then
      GoTo EXITSUB
   End If
   ' 條碼長度必須為13碼
   If Len(strData) <> 13 Then
      GoTo EXITSUB
   End If
      
   ' 取得總數
   nSumEven = 0
   nSumOdd = 0
   For nIndex = 1 To 12
      If nIndex Mod 2 = 0 Then
         nSumEven = nSumEven + Val(Mid(strData, nIndex, 1))
      Else
         nSumOdd = nSumOdd + Val(Mid(strData, nIndex, 1))
      End If
   Next nIndex
   nSum = nSumEven * 3 + nSumOdd
   
   ' 取餘數
   nRest = nSum Mod 10
   
   If nRest = 0 Then
      nCheck = 0
   Else
      nCheck = 10 - nRest
   End If

   ' 比較檢碼
   If CStr(nCheck) <> Mid(strData, 13, 1) Then
      GoTo EXITSUB
   End If

   IsBarcodeCorrect = True
   
EXITSUB:
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 取得自動編號的數字字串
'' Input : strSys == 系統別
'' Output : 傳回該系統別的自動編號
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function GetSysAutoNumber(ByVal strSys As String) As String
'   Dim rsTmp As New ADODB.Recordset
'   Dim strSql As String
'
'   GetSysAutoNumber = "000001"
'   strSql = "SELECT * FROM AutoNumber " & _
'            "WHERE AU01 = '" & strSys & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      rsTmp.MoveFirst
'      If IsNull(rsTmp.Fields("AU03")) = False Then
'         GetSysAutoNumber = String(6 - Len(rsTmp.Fields("AU03")), "0") & rsTmp.Fields("AU03")
'      End If
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
'End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' 檢查字串中的字元是否全為英文字母
' Input : strData ==> 欲檢查的字串
' Output : 若在字串中發現非英文字母的字元時則傳回False
'          否則傳回True
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsAlphabetic(ByVal strData As String) As Boolean
   Dim nIndex As Integer
   Dim nAscii As Integer
   
   IsAlphabetic = False
   For nIndex = 1 To Len(strData)
      nAscii = Asc(Mid(strData, nIndex, 1))
      If nAscii < 65 Then
         GoTo EXITSUB
      End If
      If nAscii > 90 And nAscii < 97 Then
         GoTo EXITSUB
      End If
      If nAscii > 122 Then
         GoTo EXITSUB
      End If
   Next nIndex
   IsAlphabetic = True
EXITSUB:
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得會計科目的名稱
' Input : strData ==> 會計科目的編號
' Output : 傳回會計科目的名稱
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetAccountingTitle(ByVal strData As String) As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   GetAccountingTitle = Empty
   strSql = "SELECT A0102 FROM ACC010 " & _
            "WHERE A0101 = '" & strData & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("A0102")) = False Then
         GetAccountingTitle = rsTmp.Fields("A0102")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''
' 檢查是否超過自動編號
' Input : strAU01 ==> 系統別
'         strAU02 ==> 西元年 (空白表不管)
'         strAU03 ==> 流水號
' Output : 若超過自動編號則傳回 True
'          否則傳回 False
' 說明 : 若不檢查年份時, strAU02 可給空白
'''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsOverAutoNumber(ByVal strAU01 As String, ByVal strAU02 As String, ByVal strAU03 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   
   IsOverAutoNumber = False
   'strSQL = "SELECT * FROM AUTONUMBER " & _
   '         "WHERE AU01 = '" & strAU01 & "' AND " & _
   '               "AU02 = " & strAU02 & " "
   strSql = "SELECT * FROM AUTONUMBER " & _
            "WHERE AU01 = '" & strAU01 & "' "
   If IsEmptyText(strAU02) = False Then
      strSql = strSql & " AND AU02 = '" & strAU02 & "' "
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("AU03")) = False Then
         If Val(strAU03) > Val(rsTmp.Fields("AU03")) Then
            IsOverAutoNumber = True
         End If
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得使用者可使用的系統類別
' Output : 傳回使用者可使用的系統類別(以逗號','間隔)
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function GetUserSystemKind() As String
   Dim strSql As String
   Dim rsTmp As New ADODB.Recordset
   Dim strGroup As String
   Dim strSys As String

   GetUserSystemKind = Empty
   
   strSql = "SELECT ST11 FROM Staff " & _
            "WHERE ST01 = '" & strUserNum & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If IsNull(rsTmp.Fields("ST11")) = False Then
         strGroup = rsTmp.Fields("ST11")
      End If
   End If
   rsTmp.Close
   
   strSys = Empty
   strSql = "SELECT DISTINCT(SG02) FROM STAFF_GROUP " & _
            "WHERE SG01 = '" & strGroup & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      Do While rsTmp.EOF = False
         If IsEmptyText(strSys) = False Then: strSys = strSys & ","
         strSys = strSys & rsTmp.Fields("SG02")
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close

   GetUserSystemKind = strSys
   
   Set rsTmp = Nothing
End Function

'move to basQuery by nick 2004/10/06
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 傳回欲儲存到資料庫中正確的字串值
'' Input : strData ==> 字串內容
'' Output : 此函式會回傳即將存入資料庫SQL語法中的文字資料內容
''          若是有資料的情況, 回傳值自動會加上前後的單引號
''          若沒有資料時, 回傳值會是一個內容為NULL字樣的字串
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function DBNullString(ByVal strData As String) As String
'   If IsEmptyText(strData) = True Then
'      DBNullString = "NULL"
'   Else
'      DBNullString = "'" & ChgSQL(strData) & "'"
'   End If
'End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 傳回欲儲存到資料庫中正確的數值
' Input : strData ==> 數值字串內容
' Output : 此函式會回傳即將存入資料庫SQL語法中的數值資料內容
'          若是有資料的情況, 回傳值即是傳入的文字內容
'          若沒有資料時, 回傳值會是一個內容為NULL字樣的字串
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DBNullNumeric(ByVal strData As String) As String
   If IsEmptyText(strData) = True Then
      DBNullNumeric = "NULL"
   Else
      DBNullNumeric = strData
   End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 傳回欲儲存到資料庫中正確的日期
' Input : strData ==> 數值字串內容
' Output : 此函式會回傳即將存入資料庫SQL語法中的日期資料內容
'          若是有資料的情況, 回傳值即是傳入的西元日期內容
'          若沒有資料時, 回傳值會是一個內容為NULL字樣的字串
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function DBNullDate(ByVal strData As String) As String
   If IsEmptyText(strData) = True Then
      DBNullDate = "NULL"
   Else
      DBNullDate = DBDATE(strData)
   End If
End Function

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 檢查變更事項檔是否存在
' Input : strCE01 ==> 收文號
' Output : True = 該筆收文號有變更事項檔
'          False = 該筆收文號無變更事項檔
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function IsChangeEventExist(ByVal strCE01 As String) As Boolean
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   IsChangeEventExist = False
   strSql = "SELECT * FROM CHANGEEVENT " & _
            "WHERE CE01 = '" & strCE01 & "' "
   Set rsTmp = New ADODB.Recordset
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      IsChangeEventExist = True
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'move to basquery by nickc 2007/02/07
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 檢查案件國家收費表是否存在該筆記錄
'' Input : strCF01 ==> 系統別
''         strCF02 ==> 國家代碼
''         strCF03 ==> 案件性質
'' Output : TRUE = 存在, FALSE = 不存在
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function IsExistCasefee(ByVal strCF01 As String, ByVal strCF02 As String, ByVal strCF03 As String, ByRef strCF23 As String) As Boolean
'   Dim strSQL As String
'   Dim rsTmp As ADODB.Recordset
'
'   IsExistCasefee = False
'   strCF23 = Empty
'   Set rsTmp = New ADODB.Recordset
'   strSQL = "SELECT * FROM CASEFEE " & _
'            "WHERE CF01 = '" & strCF01 & "' AND " & _
'                  "CF02 = '" & strCF02 & "' AND " & _
'                  "CF03 = '" & strCF03 & "' "
'   rsTmp.CursorLocation = adUseClient
'   rsTmp.Open strSQL, cnnConnection, adOpenStatic, adLockReadOnly
'   If rsTmp.RecordCount > 0 Then
'      If Not IsNull(rsTmp.Fields("CF23")) Then
'         If rsTmp.Fields("CF23") <> 0 Then
'            strCF23 = rsTmp.Fields("CF23")
'            IsExistCasefee = True
'         End If
'      End If
'   End If
'   rsTmp.Close
'   Set rsTmp = Nothing
'End Function
'move to basquery by nickc 2007/02/07
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'' 新增一筆收達的下一程序檔
'' Input : strNP01 ==> 總收文號
''         strNP02 ==> 本所案號第一個欄位(系統別)
''         strNP03 ==> 本所案號第二個欄位(流水號)
''         strNP04 ==> 本所案號第三個欄位(追加案號)
''         strNP05 ==> 本所案號第四個欄位(多國類別碼)
''         strDate ==> 本所期限及法定期限的日期
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Public Function InsertNextProgress_997(ByVal strNP01 As String, ByVal strNP02 As String, ByVal strNP03 As String, ByVal strNP04 As String, ByVal strNP05 As String, ByVal strDate As String) As String
'   Dim strSQL As String
'   Dim strNP22 As String
'
'   strNP22 = GetNextProgressNo()
'    'Modify By Cheng 2003/12/08
'    '若本所期限非工作天則抓最近的工作天
''   strSQL = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
''               "VALUES ('" & strNP01 & "','" & strNP02 & "','" & strNP03 & "','" & strNP04 & "','" & strNP05 & "'," & _
''                       "997" & "," & DBDATE(strDate) & "," & DBDATE(strDate) & ",'" & strUserNum & "'," & strNP22 & ") "
'   strSQL = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
'               "VALUES ('" & strNP01 & "','" & strNP02 & "','" & strNP03 & "','" & strNP04 & "','" & strNP05 & "'," & _
'                       "997" & "," & DBDATE(PUB_GetWorkDay1(strDate, True)) & "," & DBDATE(strDate) & ",'" & strUserNum & "'," & strNP22 & ") "
'   cnnConnection.Execute strSQL
'End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 取得專利的規費
' Input : strPA01 ==> 系統類別
'         strPA02 ==> 本所案號第二欄
'         strPA03 ==> 本所案號第三欄
'         strPA04 ==> 本所案號第四欄
'         strCP10 ==> 案件性質
'         strCP07 ==> 法定期限
'
' 說明 : FCP的申請國家一定是台灣
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function ReadPatentOfficialFee(ByVal strPA01 As String, _
                                      ByVal strPA02 As String, _
                                      ByVal strPA03 As String, _
                                      ByVal strPA04 As String, _
                                      ByVal strCP10 As String, _
                                      ByVal strCP07 As String) As String
   Dim rsTmp As ADODB.Recordset
   Dim strSql As String
   ' 規費
   Dim strFee As String
   ' 專利種類
   Dim strPA08 As String
   ' 申請國家
   Dim strPA09 As String
   ' 公告日
   Dim strPA14 As String
   ' 目前准駁
   Dim strPA16 As String
   
   ' 本所案號
   strPA03 = strPA03 & String(1 - Len(strPA03), "0")
   strPA04 = strPA04 & String(2 - Len(strPA04), "0")
   
   ' 讀取專利基本檔
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT * FROM PATENT " & _
            "WHERE PA01 = '" & strPA01 & "' AND " & _
                  "PA02 = '" & strPA02 & "' AND " & _
                  "PA03 = '" & strPA03 & "' AND " & _
                  "PA04 = '" & strPA04 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      ' 專利種類
      If Not IsNull(rsTmp.Fields("PA08")) Then
         strPA08 = rsTmp.Fields("PA08")
      End If
      ' 申請國家
      If Not IsNull(rsTmp.Fields("PA09")) Then
         strPA09 = rsTmp.Fields("PA09")
      End If
      ' 公告日
      If Not IsNull(rsTmp.Fields("PA14")) Then
         strPA14 = rsTmp.Fields("PA14")
      End If
      ' 目前准駁
      If Not IsNull(rsTmp.Fields("PA16")) Then
         strPA16 = rsTmp.Fields("PA16")
      End If
   End If
   rsTmp.Close
   Set rsTmp = Nothing
   
   '2010/8/17 modify by sonia
   'strFee = GetPatentOfficialFee(strPA01, strCP10, strCP07, strPA08, strPA09, strPA16, strPa14)
   strFee = GetPatentOfficialFee(strPA01, strCP10, strCP07, strPA08, strPA09, strPA16, strPA14, strPA02, strPA03, strPA04)
   '2010/8/17 end
   ReadPatentOfficialFee = strFee
End Function

'Add by Morgan 2008/9/25
'讀取會稿加乘適用規則
Public Function PUB_GetCPM05(CPM01 As String, CPM02 As String) As String
   Dim stSQL As String, intR As Integer
   Dim rsTmp As ADODB.Recordset
   stSQL = "select cpm05 from casepropertymap where cpm01='" & CPM01 & "' and cpm02='" & CPM02 & "'"
   intR = 1
   Set rsTmp = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      PUB_GetCPM05 = "" & rsTmp(0)
   End If
   Set rsTmp = Nothing
End Function

'Mark by Amy 2015/09/02 改至basQery 因案件的basFunction也有但有改,此未更新
''Add by Morgan 2008/11/13
''轉員工編號為名稱並加到ListBox
'Public Sub PUB_SetUserList(p_ListBox As ListBox, p_stNums As String)
'   Dim arrID, stSQL As String, intR As Integer, rstTmp As ADODB.Recordset
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
'                  '2012/2/14 MODIFY BY SONIA 員工編號已可非數字需做轉換
'                  p_ListBox.ITEMDATA(0) = PUB_Id2Num(.Fields(0)) '員工編號
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

'Add By Sindy 2012/11/22
' 取得特殊客戶/代理人收文費用
Public Function GetSpecGuestFee(ByVal strXNo As String, ByVal strYNo As String, ByVal strSG03 As String, ByVal strSG04 As String, ByVal strSG05 As String, ByVal strCP05 As String, ByRef strSG07 As String, ByRef strSG08 As String) As Boolean
Dim rsTmp As ADODB.Recordset
Dim strSql As String
Dim i As Integer, Item As Integer
Dim strSG01 As String, strSG02 As String
   
   Set rsTmp = New ADODB.Recordset
   GetSpecGuestFee = False
   Item = 1
   
   If strXNo = "" And strYNo <> "" Then Item = 2
   If strXNo <> "" And strYNo = "" Then Item = 3
   If strXNo <> "" Then strXNo = Left(strXNo & "00000000", 8)
   If strYNo <> "" Then strYNo = Left(strYNo & "00000000", 8)
   For i = Item To 3
      '讀取資料的順序為：X+Y、Y+Y、X+X
      If i = 1 Then
         strSG01 = strXNo: strSG02 = strYNo
      ElseIf i = 2 Then
         strSG01 = strYNo: strSG02 = strYNo
      ElseIf i = 3 Then
         strSG01 = strXNo: strSG02 = strXNo
      End If
      '啟用日期為>=收文日期的最大日期
      strSql = "SELECT * FROM SpecGuestFee " & _
               "WHERE SG01='" & strSG01 & "' AND " & _
                     "SG02='" & strSG02 & "' AND " & _
                     "SG03='" & strSG03 & "' AND " & _
                     "SG04='" & strSG04 & "' AND " & _
                     "SG05='" & strSG05 & "' AND " & _
                     "SG06=(SELECT max(SG06) FROM SpecGuestFee " & _
                            "WHERE SG01='" & strSG01 & "' AND " & _
                                  "SG02='" & strSG02 & "' AND " & _
                                  "SG03='" & strSG03 & "' AND " & _
                                  "SG04='" & strSG04 & "' AND " & _
                                  "SG05='" & strSG05 & "' AND " & _
                                  "SG06>='" & strCP05 & "') "
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If rsTmp.RecordCount > 0 Then
         rsTmp.MoveFirst
         strSG07 = Val(rsTmp.Fields("SG07"))
         strSG08 = Val(rsTmp.Fields("SG08"))
         GetSpecGuestFee = True
         rsTmp.Close
         Exit For
      End If
      rsTmp.Close
   Next i
   Set rsTmp = Nothing
End Function

'Add By Sindy 2014/3/5 FCT延期時,法定日期特殊規則取得”日期”
'ex.FCT-033757
'1.CP43若非C類時再抓其CP43為C類為止
'2.以該C類抓cp07的日,年月為延期月數後的年月
Public Function PUB_FCTGetDelaySpecDay(m_DelayDate As String, m_CP43 As String) As String
Dim Rs As New ADODB.Recordset
Dim strCP43 As String, strDay As String, strTmpDate As String
   
   m_DelayDate = TransDate(m_DelayDate, 1)
   PUB_FCTGetDelaySpecDay = m_DelayDate
   strCP43 = m_CP43
   Do While strCP43 <> ""
      If Left(strCP43, 1) = "C" Then
         strSql = "select cp07 from caseprogress where cp09='" & strCP43 & "'"
         Rs.CursorLocation = adUseClient
         Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If Rs.RecordCount > 0 Then
            If "" & Rs.Fields(0) > 0 Then
               strDay = CStr(Right("" & Rs.Fields(0), 2))
            End If
         End If
         Rs.Close
         Exit Do
      Else
         strSql = "select cp43 from caseprogress where cp09='" & strCP43 & "'"
         Rs.CursorLocation = adUseClient
         Rs.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         If Rs.RecordCount > 0 Then
            If "" & Rs.Fields(0) <> "" Then
               strCP43 = Rs.Fields(0)
            Else
               strCP43 = ""
            End If
         End If
         Rs.Close
      End If
   Loop
   If strDay <> "" Then
      strTmpDate = Left(m_DelayDate, 5) & strDay
      'Modify By Sindy 2017/7/28 + False:不彈訊息
      Do While ChkDate(strTmpDate, False) = False
         strTmpDate = Right("0" & CStr(Val(strTmpDate) - 1), 7)
      Loop
      PUB_FCTGetDelaySpecDay = strTmpDate
   End If
   
   Set Rs = Nothing
End Function

'Add By Sindy 2016/10/28
Public Function PUB_ChkOpenLetterLimit(oAccEmp As String, Optional bolShowMsg As Boolean = True) As Boolean
Dim varTmp As Variant
Dim ii As Integer
Dim rRS As New ADODB.Recordset
   
   'Modify By Sindy 2017/12/26 Mark
   'If Trim(oAccEmp) = "" And Pub_StrUserSt15 <> "M51" Then Exit Function
   If Trim(oAccEmp) = "" Then PUB_ChkOpenLetterLimit = True: Exit Function
   
   PUB_ChkOpenLetterLimit = False
   If Pub_StrUserSt15 = "M51" Or _
      (Pub_StrUserSt15 = "M31" And InStr(UCase(oAccEmp), UCase("account")) > 0) Or _
      (Left(Pub_StrUserSt15, 1) = "P" And InStr(UCase(oAccEmp), UCase("patent")) > 0) Then
      PUB_ChkOpenLetterLimit = True
      Exit Function
   End If
   
   '國外部及專利處信件:
   '雙擊主旨可開信件的部分,檢查操作人員的部門第一碼, 及所有收受者的部門第一碼,
   '必須有一個相同才可以打開信件
   varTmp = Split(oAccEmp, ";")
   For ii = 0 To UBound(varTmp)
      strSql = "select st15 from staff where st01='" & varTmp(ii) & "'"
      intI = 1
      Set rRS = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If "" & rRS.Fields(0) <> "" Then
            If Left(rRS.Fields(0), 1) = Left(Pub_StrUserSt15, 1) Then
               PUB_ChkOpenLetterLimit = True
               Exit Function
            End If
         End If
      End If
   Next ii
   
   If PUB_ChkOpenLetterLimit = False And bolShowMsg = True Then
      MsgBox "無權限開啟此郵件！", vbInformation
   End If
End Function

'Add By Sindy 2016/10/6
Public Sub PUB_ExLetterTransTxt(oFile As Object, Text1 As Object)
   Text1 = oFile.Name
   If Text1 <> oFile.Name Then
      'Add By Sindy 2016/10/4 因為”|”ex. 主旨:Owner of TW-Patent No. 141540 (our ref: DR015-1-TW)  |  TW-Patent No. 97777 (our ref: DR014-1-TW)
      If InStr(Text1, "|") > 0 Then
         oFile.Name = Replace(Text1, "|", "_")
         Text1 = oFile.Name
      End If
      If InStr(Text1, "?") > 0 Then
      '2016/10/4 END
         oFile.Name = Replace(Text1, "?", "_")
         Text1 = oFile.Name
      End If
      'Add By Sindy 2016/10/6
      If InStr(Text1, "o") > 0 Then 'o(德文上面有2點)
         oFile.Name = Replace(Text1, "o", "_")
         Text1 = oFile.Name
      End If
      If InStr(Text1, "a") > 0 Then
         oFile.Name = Replace(Text1, "a", "_") 'a(德文上面有2點)
         Text1 = oFile.Name
      End If
      If InStr(Text1, "￡") > 0 Then
         oFile.Name = Replace(Text1, "￡", "_")
         Text1 = oFile.Name
      End If
      '2016/10/6 END
      'Add By Sindy 2017/1/4
      If InStr(Text1, "&") > 0 Then
         oFile.Name = Replace(Text1, "&", "_")
         Text1 = oFile.Name
      End If
      '2017/1/4 END
      'Add By Sindy 2016/12/12 讀取檔案會失敗
      If InStr(Text1, "〔") > 0 Then
         oFile.Name = Replace(Text1, "〔", "_")
         Text1 = oFile.Name
      End If
      If InStr(Text1, "〕") > 0 Then
         oFile.Name = Replace(Text1, "〕", "_")
         Text1 = oFile.Name
      End If
      '2016/12/12 END
      'Add By Sindy 2016/12/22 讀取檔案會失敗
      If InStr(Text1, "’") > 0 Then
         oFile.Name = Replace(Text1, "’", "_")
         Text1 = oFile.Name
      End If
      If InStr(Text1, "‧") > 0 Then
         oFile.Name = Replace(Text1, "‧", "_")
         Text1 = oFile.Name
      End If
      '2016/12/22 END
   End If
End Sub

'Move by Sindy 2025/2/19 從basConst搬過來
'Move by Lydia 2023/04/27 從basPublic搬過來
'Added by Lydia 2018/07/18 FCT發文自動將下載的FCTxxxx.案件性質.PDF檔,上傳到卷宗區
Public Function Pub_AutoSavePdf_FCT(ByVal iCP01 As String, ByVal iCP02 As String, ByVal iCP03 As String, ByVal iCP04 As String, ByVal iCp09 As String, ByVal iCP10 As String) As Boolean
Dim fs, f
Dim fType As String, fPath As String
Dim strErr As String
Dim strFileName As String, stReName As String
Dim strB01 As String, intB As Integer
Dim rsB1 As New ADODB.Recordset
'Dim bolConn As Boolean
Dim stCP118 As String
Dim strFilePath As String 'Added by Lydia 2023/04/27 FCT_WORKFLOW相對應案號的資料夾路徑
'Add By Sindy 2025/2/17
Dim strErrTitle As String
Dim oFolder As Folder
Dim oFiles As files
Dim oFile
Dim strCP09 As String, strCP10 As String
Dim objOutLook As Object
Dim objMail As Object
Dim strMailDate As String, strMailTime As String, strMailSub As String, strSecName As String
'2025/2/17 END
   
   strErrTitle = "FCT發文自動將下載的PDF檔" 'Add By Sindy 2025/2/17
   If iCP01 <> "FCT" Or iCP02 = "" Or iCp09 = "" Then
       Exit Function
   End If
   
   fType = "DATA"
'      bolConn = True
   Set fs = CreateObject("Scripting.FileSystemObject") 'Add By Sindy 2025/2/14
   
   'Modified by Lydia 2020/03/09 測試抓桌面的相同資料夾以免誤刪真實檔案
   'If Pub_StrUserSt03 = "M51" Then
   If Pub_StrUserSt03 = "M51" Or UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then
JumpChk:
       'Modified by Lydia 2024/07/22 改用變數
       'If MsgBox("外商路徑：\\Typing2\國外部\外商\卷宗匯入區" & vbCrLf & "是否採用？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
       '     fPath = "\\Typing2\國外部\外商\卷宗匯入區"
       If MsgBox("外商路徑：\\" & strTyping2Path & "\國外部\外商\卷宗匯入區" & vbCrLf & "是否採用？", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
            fPath = "\\" & strTyping2Path & "\國外部\外商\卷宗匯入區"
       'end 2024/07/22
       Else
            fPath = InputBox("本機路徑：", , PUB_Getdesktop & "\卷宗匯入區")
            If Dir(fPath, vbDirectory) = "" Then
                MsgBox "路徑不存在！", vbCritical
                GoTo JumpChk
            End If
       End If
   Else
       'Modified by Lydia 2020/07/22 發文之資料夾設定改在發文作業第一畫面，整批也參照該設定
       'fPath = GetSetting("TAIE", "FCP", "FRM040111" & "Dir", "")
       fPath = GetSetting("TAIE", "FCT", "FRM030202_01" & "Dir", "")
       'Modified by Lydia 2024/07/22 改用變數
       'If fPath = "" Then fPath = "\\Typing2\國外部\外商\卷宗匯入區"
       If fPath = "" Then fPath = "\\" & strTyping2Path & "\國外部\外商\卷宗匯入區"
       If Dir(fPath, vbDirectory) = "" Then
             strErr = strErr & vbCrLf & fPath & "，路徑不存在！"
             GoTo JumpExit
       End If
   End If
   
On Error GoTo JumpExit
   'Modified by Lydia 2018/12/18 改成多筆PDF上傳
   'strFileName = Dir(fPath & "\" & iCP01 & "*" & Val(iCP02) & "." & iCP10 & ".pdf")
   'Modified by Lydia 2019/06/04 無法匯入FCT-6586-T案
   'strFileName = Dir(fPath & "\" & iCP01 & "*" & Val(iCP02) & "." & iCP10 & "*.pdf")
   strFileName = Dir(fPath & "\" & iCP01 & "*" & Val(iCP02) & IIf(iCP03 <> "0", "*" & iCP03, "") & IIf(iCP04 <> "00", "*" & iCP04, "") & "." & iCP10 & "*.pdf")
   If strFileName <> "" Then
        '若為電子送件的檔案是否存在
        'Remove by Lydia 2020/07/22 FCT案開放發文後,再自行將匯入區之電子檔整批匯入
        'strB01 = "select cp01,cp02,cp03,cp04,cp09,cp10,cp118,cp121 from caseprogress where cp09=" & CNULL(iCp09)
        'intB = 1
        'Set rsB1 = ClsLawReadRstMsg(intB, strB01)
        'If intB = 1 Then
        '     If "" & rsB1.Fields("cp118") = "Y" And "" & rsB1.Fields("cp121") = "Y" Then
        '         strErr = strErr & vbCrLf & "收文號：" & iCp09 & "，說明書電子檔已上傳！"
        '         GoTo JumpExit
        '     End If
        '     stCP118 = "" & rsB1.Fields("cp118")
        'End If
        'end 2020/07/22
        
        'Added by Lydia 2018/12/18
        Do While strFileName <> ""
             'Set fs = CreateObject("Scripting.FileSystemObject") 'Modify By Sindy 2025/2/14 mark;
             Set f = fs.GetFile(fPath & "\" & strFileName)
             '檔案大小為 0 KB 有誤
             If f.Size = 0 Then
                  strErr = strErr & vbCrLf & fPath & "\" & strFileName & "，" & MsgText(9221)
                  'Modified by Lydia 2018/12/18
                  'GoTo JumpExit
                  GoTo JumpNextDir
             End If
             If PUB_ChkFileOpening(fPath & "\" & strFileName) = True Then
                 strErr = strErr & vbCrLf & fPath & "\" & strFileName & "，檔案正在使用中，請關閉或關閉檔案後間隔1分鐘，方能上傳到卷宗區。"
                 'Modified by Lydia 2018/12/18
                 'GoTo JumpExit
                 GoTo JumpNextDir
             End If
              
             '檢查檔名規則
             If PUB_ChkEmpFlowFNMRule(iCP01 & "-" & iCP02 & "-" & iCP03 & "-" & iCP04, strFileName, "Y", iCP10, , , False, False, strErr) = False Then
                  'Modified by Lydia 2018/12/18
                  'GoTo JumpExit
                  GoTo JumpNextDir
             End If
             '更名
             'Modified by Lydia 2018/12/18
             'If PUB_GetEmpFlowReNameFile(iCP01, iCP02, iCP03, iCP04, iCP10, strFileName, stReName, True, 1, False, strErr, , fType) = False Then
             '     GoTo JumpNextDir
             'End If
             stReName = strFileName
             '案件性質+DATA
             If UCase(Right(stReName, Len("." & iCP10 & ".PDF"))) = "." & iCP10 & ".PDF" Then
                 stReName = Replace(stReName, "." & iCP10 & ".PDF", "." & iCP10 & "." & fType & ".PDF")
                 stReName = Replace(stReName, "." & iCP10 & ".pdf", "." & iCP10 & "." & fType & ".pdf")
             End If
             'Added by Lydia 2020/07/22 本所案號統一為6碼
             'Modified by Lydia 2022/05/13 預設本所案號; ex.FCT047520的AB1018569(DATA,GSN) 案號不足6碼
             'If InStr(stReName, iCP01 & iCP02) = 0 Then
             '    stReName = Replace(stReName, iCP01 & Val(iCP02), iCP01 & iCP02)
             'End If
             'stReName = UCase(stReName) '統一為大寫
             'end 2020/07/22
             If InStr(stReName, "." & iCP10 & ".") = 0 Then
                 stReName = Replace(stReName, "." & iCP10, "." & iCP10 & ".")
             End If
             strB01 = PUB_CaseNo2FileName(iCP01, iCP02, iCP03, iCP04) & Mid(stReName, InStr(stReName, "." & iCP10))
             stReName = UCase(PUB_GetReNameFileSignal(strB01))
             'end 2022/05/13

             'Move by Lydia 2018/12/18 從檢查檔案大小的上面移下來
             '檢查卷宗區檔案是否存在
             'Modified by Lydia 2018/12/18
             'strB01 = "SELECT cpp01,cpp02 FROM casepaperpdf " & _
                           "WHERE cpp01 ='" & iCp09 & "' and instr(upper(cpp02),'" & fType & ".PDF') > 0 and instr(upper(cpp02),'PDF.DEL') = 0 "
             strB01 = "SELECT cpp01,cpp02 FROM casepaperpdf " & _
                           "WHERE cpp01 ='" & iCp09 & "' and instr(upper(cpp02),'" & UCase(stReName) & "') > 0 and instr(upper(cpp02),'PDF.DEL') = 0 "
             intB = 1
             Set rsB1 = ClsLawReadRstMsg(intB, strB01)
             If intB = 1 Then
                  strErr = strErr & vbCrLf & fPath & "\" & rsB1.Fields("cpp02") & "，卷宗區檔案已存在！"
                  'Modified by Lydia 2018/12/18
                  'GoTo JumpExit
                  GoTo JumpNextDir
             End If
             'end 2018/12/18
             
             '上傳到卷宗區
             'cnnConnection.BeginTrans 'Remove by Lydia 2018/12/18
             If SaveAttFile_PDF(iCp09, fPath & "\" & strFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = False Then
                  strErr = strErr & vbCrLf & fPath & "\" & strFileName & "，存檔失敗！" & vbCrLf & Err.Description
'                     bolConn = False
                  'Modified by Lydia 2018/12/18
                  'GoTo JumpExit
                  GoTo JumpNextDir
             End If
             'Remove by Lydia 2018/12/18
             'If stCP118 = "Y" Then
             '     Call UpdateCP121(iCp09, iCP10, fType)
             'End If
             'cnnConnection.CommitTrans
             'end 2018/12/18
             'Added by Lydia 2023/04/27 原本自動上傳卷宗區之提申資料(如申請書、contact…等)同時存至FCT_WORKFLOW。
             If strFilePath = "" Then   'Memo by Lydia 2023/05/03 報告客戶之資料統一存檔FCT_WORKFLOW: 2023/5/5 上線
                strFilePath = Pub_GetEFilePath_All(iCP01, iCP02, iCP03, iCP04)
             End If
             If strFilePath <> "" Then
                 '檢查檔案是否存在
                 If UCase(strFileName) <> UCase(stReName) Then
                    f.Name = stReName
                    strFileName = stReName
                 End If
                 strB01 = Pub_GetEFileName(strFilePath, stReName)
                 If Len(stReName) <> Len(strB01) Then
                    f.Name = strB01
                    strFileName = strB01
                 End If
                 fs.CopyFile fPath & "\" & strFileName, strFilePath & "\" & strFileName
                 Sleep 1000
             End If
             'end 2023/04/27
             fs.DeleteFile fPath & "\" & strFileName, True '刪檔
        
        'Added by Lydia 2018/12/18
JumpNextDir:
             strFileName = Dir()
        Loop
        'Remove by Lydia 2020/07/22 FCT案開放發文後,再自行將匯入區之電子檔整批匯入
        'If stCP118 = "Y" Then '電子送件需更新CP121
        '      Call UpdateCP121(iCp09, iCP10, fType)
        'End If
        'end 2018/12/18
        'end 2020/07/22
        
        Pub_AutoSavePdf_FCT = True
   End If
   
   'Add By Sindy 2025/2/14 FCT未成卷信函: 系統於個案程序同仁發文送件時，
   '   將有本所案號之信函資料夾，自動匯入該案號之最早總收文號之案件性質卷宗區內,並刪除該資料夾
   strErrTitle = "FCT未成卷信函" 'Add By Sindy 2025/2/17
   fPath = Pub_GetSpecMan("FCT未成卷信函存放路徑")
   If Pub_StrUserSt03 = "M51" Or UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then '測試抓桌面的相同資料夾以免誤刪真實檔案
JumpChk2:
      If MsgBox("共用資料夾路徑：" & fPath & vbCrLf & "是否採用？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
           fPath = InputBox("本機路徑：", , PUB_Getdesktop & "\" & Mid(fPath, InStrRev(fPath, "\") + 1))
           If Dir(fPath, vbDirectory) = "" Then
               strErr = strErr & vbCrLf & fPath & "，路徑不存在！"
               GoTo JumpExit
           End If
      End If
   End If
   fPath = fPath & "\" & iCP01 & iCP02 & IIf(iCP03 & iCP04 = "000", "", iCP03 & iCP04)
   If Dir(fPath, vbDirectory) = "" Then
       GoTo JumpExit '離開
   End If
   
   Set oFolder = fs.GetFolder(fPath)
   Set oFiles = oFolder.files
   If oFiles.Count > 0 Then '有檔案
      '該案號之最早總收文號之案件性質
      strB01 = "SELECT cp09,cp10 FROM caseprogress" & _
               " WHERE cp01 ='" & iCP01 & "' and cp02 ='" & iCP02 & "' and cp03 ='" & iCP03 & "' and cp04 ='" & iCP04 & "'" & _
               " and substr(cp09,1,1)='A'" & _
               " order by cp05 asc,cp09 asc"
      intB = 1
      Set rsB1 = ClsLawReadRstMsg(intB, strB01)
      If intB = 1 Then
         strCP09 = rsB1.Fields("cp09")
         strCP10 = rsB1.Fields("cp10")
      End If
      If strCP09 = "" Then
         strErr = strErr & vbCrLf & "抓不到此案號最早的總收文號！"
         GoTo JumpExit
      End If
      
      Set objOutLook = CreateObject("Outlook.Application")
      For Each oFile In oFiles
         strFileName = oFile.Name
         If Dir(fPath & "\" & strFileName) <> "" Then
            Set f = fs.GetFile(fPath & "\" & strFileName)
            '檔案大小為 0 KB 有誤
            If f.Size = 0 Then
               strErr = strErr & vbCrLf & fPath & "\" & strFileName & "，" & MsgText(9221)
               GoTo JumpNextDir2
            End If
            If PUB_ChkFileOpening(fPath & "\" & strFileName) = True Then
               strErr = strErr & vbCrLf & fPath & "\" & strFileName & "，檔案正在使用中，請關閉或關閉檔案後間隔1分鐘，方能上傳到卷宗區。"
               GoTo JumpNextDir2
            End If
            
            '檢查檔名規則
            If UCase(Right(Trim(strFileName), 4)) = ".MSG" Then
               Set objMail = objOutLook.CreateItemFromTemplate(fPath & "\" & strFileName)
               '郵件ReName
               Call PUB_ReadMailText(objMail, , , , , strMailDate, strMailTime, strMailSub)
               '解析主旨
               strSecName = ""
               Call PUB_IPDept_ComparisonCP(strMailSub, "", iCP01, iCP02, iCP03, iCP04, strSecName, "", "")
               'strSecName回傳回來,範例: .paper.msg
               If strSecName <> "" Then
                  If InStr(UCase(strSecName), UCase(Right(Trim(strFileName), 4))) > 0 Then
                     strSecName = Mid(strSecName, 1, Len(strSecName) - 4)
                  End If
                  If Left(strSecName, 1) = "." Then
                     strSecName = Mid(strSecName, 2)
                  End If
               End If
               stReName = PUB_CaseNo2FileName(iCP01, iCP02, iCP03, iCP04) & _
                              "." & strCP10 & IIf(strSecName = "", "", "." & strSecName) & _
                              UCase(Right(Trim(strFileName), 4))
               stReName = PUB_ReMailFileName(strCP09, stReName, strMailDate, strMailTime, strSecName)
            Else
               'If InStr(strFileName, iCP01) = 0 And InStr(strFileName, iCP02) = 0 Then
               strExc(10) = PUB_CaseNo2FileName(iCP01, iCP02, iCP03, iCP04)
               If Left(strFileName, Len(strExc(10))) = strExc(10) Then
                  stReName = strFileName
               Else
                  stReName = PUB_CaseNo2FileName(iCP01, iCP02, iCP03, iCP04) & "." & strCP10 & "." & strFileName
               End If
               If PUB_ChkEmpFlowFNMRule(iCP01 & "-" & iCP02 & "-" & iCP03 & "-" & iCP04, stReName, "Y", strCP10, , , False, False, strErr) = False Then
                  GoTo JumpNextDir2
               End If
               If PUB_GetEmpFlowReNameFile(iCP01, iCP02, iCP03, iCP04, strCP10, ChgSQL(stReName), stReName, True, 1, , , strCP09) = False Then
                  GoTo JumpNextDir2
               End If
            End If
            
            '卷宗區
            If UCase(Right(Trim(stReName), 4)) = ".MSG" _
               Or UCase(Right(Trim(stReName), 4)) = ".PDF" Then
               '檢查卷宗區檔案是否存在
               strB01 = "SELECT cpp01,cpp02 FROM casepaperpdf " & _
                             "WHERE cpp01 ='" & strCP09 & "' and instr(upper(cpp02),'" & UCase(stReName) & "') > 0 and instr(upper(cpp02),'" & UCase(stReName) & ".DEL" & "') = 0 "
               intB = 1
               Set rsB1 = ClsLawReadRstMsg(intB, strB01)
               If intB = 1 Then
                  strErr = strErr & vbCrLf & fPath & "\" & rsB1.Fields("cpp02") & "，卷宗區檔案已存在(" & strCP09 & ")！"
                  GoTo JumpNextDir2
               End If
               
               '上傳到卷宗區
               If SaveAttFile_PDF(strCP09, fPath & "\" & strFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = False Then
                  strErr = strErr & vbCrLf & fPath & "\" & strFileName & "，存檔失敗(" & strCP09 & ")！" & vbCrLf & Err.Description
                  GoTo JumpNextDir2
               Else
                 oFile.Delete True '刪檔
               End If
            
            '原始檔區
            Else
               '檢查原始檔區檔案是否存在
               strB01 = "SELECT cpf01,cpf02 FROM casepaperfile " & _
                             "WHERE cpf01 ='" & strCP09 & "' and instr(upper(cpf02),'" & UCase(stReName) & "') > 0 and instr(upper(cpf02),'" & UCase(stReName) & ".DEL" & "') = 0 "
               intB = 1
               Set rsB1 = ClsLawReadRstMsg(intB, strB01)
               If intB = 1 Then
                  strErr = strErr & vbCrLf & fPath & "\" & rsB1.Fields("cpf02") & "，卷宗區檔案已存在(" & strCP09 & ")！"
                  GoTo JumpNextDir2
               End If
               
               '上傳到原始檔區
               If SaveAttFile_Org(strCP09, fPath & "\" & strFileName, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS")) = False Then
                  strErr = strErr & vbCrLf & fPath & "\" & strFileName & "，存檔失敗(" & strCP09 & ")！" & vbCrLf & Err.Description
                  GoTo JumpNextDir2
               Else
                 oFile.Delete True '刪檔
               End If
            End If
         End If
JumpNextDir2:
      Next
      Set objMail = Nothing
      Set objOutLook = Nothing
      
      Set oFiles = oFolder.files
      strB01 = iCP01 & iCP02 & IIf(iCP03 & iCP04 = "000", "", iCP03 & iCP04) & _
                  GetPrjState4(iCP01 & "-" & iCP02 & "-" & iCP03 & "-" & iCP04, iCP10) & _
                  "發文,FCT未成卷信函歸 (" & GetPrjState4(iCP01 & "-" & iCP02 & "-" & iCP03 & "-" & iCP04, strCP10) & ") "
      If oFiles.Count = 0 Then '無檔案,刪除該資料夾
         If oFolder.SubFolders.Count = 0 Then
            oFolder.Delete True
         End If
         PUB_SendMail strUserNum, PUB_GetFCTSalesNo(iCP01, iCP02, iCP03, iCP04), "", strB01 & ", 已歸卷完成!", "同主旨"
      Else
         PUB_SendMail strUserNum, PUB_GetFCTSalesNo(iCP01, iCP02, iCP03, iCP04), "", strB01 & ", 尚有 " & oFiles.Count & " 個電子檔未歸卷, 請人工處理!", "同主旨" & vbCrLf & vbCrLf & strErr
      End If
   Else
      If oFolder.SubFolders.Count = 0 Then
         oFolder.Delete True
      End If
   End If
   
   Set rsB1 = Nothing
   Exit Function
      
JumpExit:
   
   Set rsB1 = Nothing
   
   If Err.Number <> 0 Then strErr = strErr & vbCrLf & Err.Description
   'If bolConn = False Then cnnConnection.RollbackTrans
   
   If strErr <> "" Then
      'Modify By Sindy 2025/2/17
      'MsgBox "FCT發文自動將下載的PDF檔，上傳到卷宗區作業失敗：" & strErr, vbCritical
      MsgBox strErrTitle & "，上傳到卷宗區作業失敗：" & strErr, vbCritical
      '2025/2/17 END
   End If
End Function

'Add By Sindy 2022/2/8 解析信件內容
'objMail.Recipients(kk).Type: 1.收件者 2.副本
'Modify By Sindy 2024/7/23 + , Optional ByVal bolOnlyReadAddrAnd2 As Boolean = False: True=僅讀副本的Addr
Public Function PUB_ReadMailText(ByVal objMail As Object, Optional ByRef strRecipients_all As String, _
   Optional ByRef strRecipients_1 As String, Optional ByVal bolOnlyReadAddrAnd2 As Boolean = False, _
   Optional ByRef strSender As String, _
   Optional ByRef strMailDate As String, Optional ByRef strMailTime As String, _
   Optional ByRef strMailSub As String) As Boolean
   
Dim kk As Integer
   
   strMailSub = ChgSQL(objMail.Subject)
   strSender = ""
   If objMail.Class = 46 Then '46.olReport
      strSender = "未傳遞的主旨"
      strMailDate = "0"
      strMailTime = ""
   '43.olMail
   Else
      strMailDate = Format(objMail.SentOn, "YYYYMMDD") 'ReceivedTime
      strMailTime = Format(objMail.SentOn, "HHMMSS")
      
      If objMail.SenderEmailType = "EX" Then
         strSender = objMail.SenderName
      Else
         If objMail.SenderName = objMail.senderemailaddress Then
            strSender = objMail.senderemailaddress
         Else
            'Add By Sindy 2024/7/29
            'Modify By Sindy 2025/2/5 + Or objMail.senderemailaddress = ""
            If InStr(UCase(objMail.senderemailaddress), UCase("Recipients/cn=")) > 0 _
               Or InStr(UCase(objMail.senderemailaddress), UCase("Public Folder/CN=")) > 0 _
               Or objMail.senderemailaddress = "" Then
               strSender = objMail.SenderName
            Else
            '2024/7/29 END
               strSender = objMail.SenderName & " [" & objMail.senderemailaddress & "]"
            End If
         End If
      End If
      
'            'Add By Sindy 2020/8/26 出現objMail.Recipients.Count到999,在M51-Win7會出現記憶不足
'            If objMail.Recipients.Count >= 99 Then
'               intRunMax = 99
'            Else
'               intRunMax = objMail.Recipients.Count
'            End If
      'Add By Sindy 2017/10/20
      For kk = objMail.Recipients.Count To 1 Step -1
   '            For kk = intRunMax To 1 Step -1
      '2020/8/26 END
   '               If objMail.Recipients(kk).Type = 1 Then '1.收件者
            strExc(10) = ""
            If InStr(UCase(objMail.Recipients(kk).address), UCase("@taie.com.tw")) > 0 Then
               strExc(10) = objMail.Recipients(kk).address
            
            ElseIf InStr(UCase(objMail.Recipients(kk).Name), UCase("@taie.com.tw")) > 0 Then
               strExc(10) = objMail.Recipients(kk).Name
            
            ElseIf InStr(UCase(objMail.Recipients(kk).Name), UCase("ipdept")) > 0 Or _
                   InStr(UCase(objMail.Recipients(kk).Name), UCase("專利處信箱")) > 0 Or _
                   InStr(UCase(objMail.Recipients(kk).Name), UCase("patent")) > 0 Or _
                   InStr(UCase(objMail.Recipients(kk).Name), UCase("tm")) > 0 Or _
                   InStr(UCase(objMail.Recipients(kk).Name), UCase("account")) > 0 Then
               strExc(10) = objMail.Recipients(kk).Name
            
            Else 'If objMail.Recipients(kk).Name <> objMail.Recipients(kk).address And _
               'InStr(objMail.Recipients(kk).address, "@") = 0 Then
               If InStr(UCase(objMail.Recipients(kk).address), UCase("Recipients/cn=")) > 0 _
                  Or InStr(UCase(objMail.Recipients(kk).address), UCase("Public Folder/CN=")) > 0 Then
                  strExc(10) = "" 'Mid(objMail.Recipients(kk).address, InStr(UCase(objMail.Recipients(kk).address), UCase("Recipients/cn=")) + Len("Recipients/cn="))
               Else
                  strExc(10) = objMail.Recipients(kk).address
               End If
               strExc(10) = Replace(strExc(10), """", "")
               If InStr(strRecipients_all, strExc(10)) = 0 Then
                  'Modify By Sindy 2024/7/23
                  If bolOnlyReadAddrAnd2 = True Then
                     If objMail.Recipients(kk).Type <> 1 Then
                        If strExc(10) <> "" Then
                           strRecipients_all = strRecipients_all & ";" & strExc(10)
                        Else
                           strRecipients_all = strRecipients_all & ";" & objMail.Recipients(kk).Name
                        End If
                     End If
                  Else
                  '2024/7/23 END
                     strRecipients_all = strRecipients_all & ";" & objMail.Recipients(kk).Name & IIf(strExc(10) <> "", "(" & strExc(10) & ")", "")
                  End If
                  If objMail.Recipients(kk).Type = 1 Then
                     strRecipients_1 = strRecipients_1 & ";" & objMail.Recipients(kk).Name & IIf(strExc(10) <> "", "(" & strExc(10) & ")", "")
                  End If
               End If
               strExc(10) = ""
            End If
            
            If strExc(10) <> "" Then
               'Modify By Sindy 2024/7/23
               If bolOnlyReadAddrAnd2 = True Then
                  If objMail.Recipients(kk).Type <> 1 Then
                     strRecipients_all = strRecipients_all & ";" & strExc(10)
                  End If
               Else
               '2024/7/23 END
                  strRecipients_all = strRecipients_all & ";" & strExc(10)
               End If
               If objMail.Recipients(kk).Type = 1 Then
                  strRecipients_1 = strRecipients_1 & ";" & strExc(10)
               End If
            End If
      Next kk
      If strRecipients_all <> "" Then strRecipients_all = Mid(strRecipients_all, 2)
      If strRecipients_1 <> "" Then strRecipients_1 = Mid(strRecipients_1, 2)
      '2017/10/20 END
   End If
   
   PUB_ReadMailText = True
End Function

' 若為郵件修改檔名為日期時間加序號
Public Function PUB_ReMailFileName(strCP09 As String, strFileName As String, _
   strDate As String, strTime As String, Optional ByVal strSecName As String = "") As String
Dim strTempFileName As String
Dim strTempFileName1 As String, strTempFileName2 As String
Dim adoRst As ADODB.Recordset
Dim intRow As Integer
Dim varTemp As Variant
Dim ii As Integer
   
   PUB_ReMailFileName = strFileName
   intRow = 0
   
   varTemp = Split(strFileName, ".")
   For ii = 0 To 1 'UBound(sFile)
      strTempFileName1 = strTempFileName1 & varTemp(ii) & "."
   Next ii
   ii = UBound(varTemp)
   If UCase(varTemp(ii)) <> UCase("msg") Then Exit Function '非郵件,離開
   If strDate = 0 Then
      strDate = strSrvDate(1)
      strTime = ServerTime
   End If
   strTempFileName1 = strTempFileName1 & strDate & Right("000000" & strTime, 6) & "."
   If strSecName = "" Then
      strTempFileName2 = varTemp(ii)
   Else
      strTempFileName2 = varTemp(ii - 1) & "." & varTemp(ii)
   End If
GotoChk:
   If intRow > 0 Then
      strTempFileName = strTempFileName1 & intRow & "." & strTempFileName2
   Else
      strTempFileName = strTempFileName1 & strTempFileName2
   End If
   strSql = "SELECT cpp01 FROM casepaperpdf WHERE cpp01='" & strCP09 & "' and upper(cpp02)=upper('" & ChgSQL(strTempFileName) & "')"
   intI = 1
   Set adoRst = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      intRow = intRow + 1
      GoTo GotoChk
   End If
   PUB_ReMailFileName = strTempFileName
   
   Set adoRst = Nothing
End Function

'Modify By Sindy 2017/3/8 變共用函數
'國外部信件轉入及分類
'strTo:轉寄人員
'strII05:分類
'回傳:是否成功
'Modify By Sindy 2017/5/17 + Optional ByVal strProFileName As String = "" : 指定處理的檔案
'                            strProFileName="N" : 手動拖拉.Msg檔至資料夾再匯入
'Modify By Sindy 2017/6/13 + Optional ByVal strCaseNo As String = "" : 本所案號
Public Function PUB_IPDeptTransMail_New(oForm As Form, Optional ByRef strTo As String, _
   Optional ByRef strErrText As String, Optional ByRef strII05 As String, _
   Optional ByVal strProFileName As String = "", Optional ByRef strCaseNo As String = "") As Boolean
Dim objOutLook As Object
Dim objMail As Object
Dim myForward As Object 'Add By Sindy 2017/6/26
Dim strII03 As String, strII03_2 As String, strII11 As String, strII12 As String, strII13 As String
Dim strII18 As String 'Add By Sindy 2017/8/28
Dim strUpdTime As String
Dim stFtpPath As String
Dim strII06 As String, strII17 As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim oFileSys As New FileSystemObject
Dim oFolder As Folder
Dim oFile As Object
Dim strCP09 As String, strCP10 As String
Dim fs, f
Dim bolSaveEFile As Boolean
Dim lngRonCnt As Long
Dim bolConnect As Boolean
Dim intII03 As Integer 'Add By Sindy 2016/10/4
Dim tmpArr As Variant, j As Integer, strIRUpdTime As String
Dim stReName As String
Dim intCaseKind As Integer 'Add By Sindy 2017/6/23
Dim strCC As String, ArrStr As Variant, strTempCC As String, ArrStrkk As Variant  'Add By Sindy 2017/6/26
Dim jj As Integer, kk As Integer 'Add By Sindy 2017/6/26
Dim strPI03 As String, strPI03_2 As String
Dim strRecipients_1 As String 'Add By Sindy 2017/10/20 收件者
Dim strRecipients_all As String 'Add By Sindy 2019/4/12 全部含副本等收件者
Dim bolForKeyWordDel As Boolean
Dim strII20 As String 'Add By Sindy 2017/11/17
Dim strTableName As String, strUpdWhere As String 'Add By Sindy 2017/11/17
Dim bolTaieStaff  As Boolean 'Add By Sindy 2020/5/28
Dim intSameCnt As Integer
Dim bolPatentDel As Boolean 'Add By Sindy 2022/3/2
Dim strToSys As String, strCP12 As String, bolSendMailErr As Boolean 'Add By Sindy 2022/5/24
Dim strII17AddText As String, strContent As String
Dim strCP13 As String 'Add By Sindy 2023/7/18
Dim strRestKind As String, bolRest As Boolean, strToSpec As String 'Add By Sindy 2024/6/11
   
On Error GoTo ErrHand
   
   bolPatentDel = False
   PUB_IPDeptTransMail_New = False
   strErrText = "": strII18 = "": strII17AddText = ""
   Set oFolder = oFileSys.GetFolder(oForm.txtPathIPDept.Text)
'   cmdTrans.Enabled = False
'   oForm.TxtIPDept.Visible = True
'   oForm.LblCntIPDept.Visible = True
   Set objOutLook = CreateObject("Outlook.Application")
   Set fs = CreateObject("Scripting.FileSystemObject")
   lngRonCnt = 0
   For Each oFile In oFolder.files
      lngRonCnt = lngRonCnt + 1
      oForm.LblCntIPDept.Caption = "已處理件數 / 剩餘件數：" & lngRonCnt & " / " & oFolder.files.Count
      DoEvents
      oForm.TxtIPDept = oFile.Name
      
      If UCase(Right(Trim(oFile.Name), 4)) = UCase(".msg") And _
         (strProFileName = "N" Or UCase(Trim(strProFileName)) = UCase(Trim(oFile.Name))) Then
         Call PUB_ExLetterTransTxt(oFile, oForm.TxtIPDept)
         
         strTo = "" '轉寄人員
         Set objMail = objOutLook.CreateItemFromTemplate(oForm.txtPathIPDept.Text & "\" & oFile.Name)
         DoEvents 'Add By Sindy 2019/12/13
         Screen.MousePointer = vbHourglass
         
         'strII03 = Trim(oFile.Name)
'         strII17 = ChgSQL(objMail.Subject)
'         oForm.TextII17 = objMail.Subject 'Add By Sindy 2021/4/12 Find簡體字
''         oForm.Text2 = strII17 'Add By Sindy 2016/4/21 Re: ML/kc 中?特許出願201510920053.X　貴所整理番?31565－CN　弊所整理番?：P-112987
''         strII17 = ChgSQL(oForm.Text2) '要用文字框存放，因才能把unicode去掉
'         DoEvents
''         If strII17 <> objMail.Subject Then
''            MsgBox "主旨抓的有誤，請洽電腦中心！"
''            GoTo ErrHand
''         End If
'         'If InStr(strII03, "未傳遞的主旨") = 0 And InStr(strII03, "延遲的傳遞") = 0 And Left(strII03, 3) <> "已讀取" Then
         
         'Modify By Sindy 2025/2/17
         strRecipients_1 = "" '收件者
         strRecipients_all = ""
'         If objMail.Class = 46 Then '46.olReport
'            strII11 = "未傳遞的主旨"
'            strII12 = "0"
'            strII13 = ""
'         '43.olMail
'         Else
'            strII11 = PUB_GetMail_ii11(objMail) 'Modify By Sindy 2024/7/30
'            strII12 = Format(objMail.SentOn, "YYYYMMDD") 'ReceivedTime
'            strII13 = Format(objMail.SentOn, "HHMMSS")
'
'            '抓收件者資料
'            Call PUB_ReadMailText_CC(objMail, strRecipients_all, strRecipients_1)
'         End If
         Call PUB_ReadMailText(objMail, strRecipients_all, strRecipients_1, , strII11, strII12, strII13, strII17)
         oForm.TextII17 = strII17 'Add By Sindy 2021/4/12 Find簡體字
         '2025/2/17 END
         
         'Modify By Sindy 2016/4/21 strII17-->Text2
         'strII05 = ToSortOut(strII17, strII11, strII06, strCP01, strCP02, strCP03, strCP04)
         strII05 = PUB_IPDept_ToSortOut(strII17, strII11, strII06, strCP01, strCP02, strCP03, strCP04, strII18)
         
         strUpdTime = Right("000000" & ServerTime, 6)
         strCP09 = "": strCP10 = "": strCP12 = ""
         '個案
         'If strII05 = "1" Then
            If strCP01 <> "" And strCP02 <> "" Then
'               '該案號最大收文日最小Create日期時間的總收文號
'               strExc(0) = "select cp09 from caseprogress" & _
'                           " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
'                           " and cp05=(select max(cp05) from caseprogress" & _
'                           " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "')" & _
'                           " order by cp66 asc,cp67 asc"
               'Modify By Sindy 2017/3/29
               '該案號A,B,C類最大收文日最大總收文號
               'Modify By Sindy 2017/7/18 不剔除D類進度 : and cp09<'D'
               strExc(0) = "select cp09 from caseprogress" & _
                           " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                           " and cp05=(select max(cp05) from caseprogress" & _
                           " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "')" & _
                           " order by SQLDatet2(CP05) DESC, nvl(cp66,cp05) desc, CP67 DESC, CP09 DESC"
                           'Modify By Sindy 2018/6/27 order by cp66 desc,cp67 desc
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strCP09 = RsTemp.Fields("cp09")
                  strCP13 = PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04) '目前智權人員
                  strCP12 = PUB_GetST03(strCP13) '目前智權人員部門
                  strExc(0) = "select cp10 from caseprogress where cp09='" & strCP09 & "'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     strCP10 = RsTemp.Fields("cp10")
                  End If
               End If
            End If
'            '有總收文號:欲收錄到卷宗區所以檔名長度有限制
'            If strCP09 <> "" Then
''               If LenB(strII03) > 74 Then
''                  strII03 = LeftB(strII03, 66) & ".rx.msg" '必須取偶數,不可奇數
''               End If
''               '郵件副檔名要取為.rx.msg
''               If InStr(UCase(strII03), UCase(".rx.msg")) = 0 Then strII03 = Left(strII03, Len(strII03) - 4) & ".rx.msg"
''               'modify by sonia 2016/4/8 1.應檢查IPDeptInput,否則寫入IPDeptInput會違反唯一的限制條件,2.與DAVID討論暫不放卷宗區,以免系統直接放入沒用的信件
''               ''檢查資料庫中是否已有今天相同的檔名存在,若有,檔名再加時間
''               'strExc(0) = "select cpp02 from casepaperpdf" & _
''               '            " where cpp01=" & CNULL(strCP09) & _
''               '            " and upper(cpp02)=upper('" & ChgSQL(strII03) & "')"
''               'intI = 1
''               'Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''               'If intI = 1 Then
''               '   strII03 = Trim(Left(strII03, Len(strII03) - 7)) & "," & strUpdTime & ".rx.msg"
''               '   '加了時間還是有可能重覆,再加當日當時筆數(流水號)
''               '   strExc(0) = "select cpp02 from casepaperpdf" & _
''               '               " where cpp01=" & CNULL(strCP09) & _
''               '               " and upper(cpp02)=upper('" & ChgSQL(strII03) & "')"
''               '   intI = 1
''               '   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''               '   If intI = 1 Then
''               '      strExc(0) = "select count(*) from IPDeptInput" & _
''               '                  " where ii01=" & strSrvDate(1)
''               '      intI = 1
''               '      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''               '      If intI = 1 Then
''               '         strExc(0) = Val(RsTemp.Fields(0)) + 1
''               '      End If
''               '      strII03 = Trim(Left(strII03, Len(strII03) - 7)) & "," & strExc(0) & ".rx.msg"
''               '   End If
''               'End If
''               '檢查資料庫中是否已有今天相同的檔名存在,若有,檔名再加時間
''               strExc(0) = "select ii03 from IPDeptInput" & _
''                           " where ii01=" & strSrvDate(1) & _
''                           " and upper(ii03)=upper('" & ChgSQL(strII03) & "')"
''               intI = 1
''               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''               If intI = 1 Then
''                  strII03 = Trim(Left(strII03, Len(strII03) - 4)) & "," & strUpdTime & ".msg"
''                  '加了時間還是有可能重覆,再加當日當時筆數(流水號)
''                  strExc(0) = "select ii03 from IPDeptInput" & _
''                              " where ii01=" & strSrvDate(1) & _
''                              " and upper(ii03)=upper('" & ChgSQL(strII03) & "')"
''                  intI = 1
''                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''                  If intI = 1 Then
''                     strExc(0) = "select count(*) from IPDeptInput" & _
''                                 " where ii01=" & strSrvDate(1)
''                     intI = 1
''                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
''                     If intI = 1 Then
''                        strExc(0) = Val(RsTemp.Fields(0)) + 1
''                     End If
''                     strII03 = Trim(Left(strII03, Len(strII03) - 4)) & "," & strExc(0) & ".msg"
''                  End If
''               End If
''               'end 2016/4/8
'            Else
'               strII05 = "Z" '其他
'            End If
         'End If

'         '非個案
'         If strII05 <> "1" Then
'            If LenB(strII03) > 74 Then
'               strII03 = LeftB(strII03, 70) & ".msg"
'            End If
'            '檢查資料庫中是否已有今天相同的檔名存在,若有,檔名再加時間
'            strExc(0) = "select ii03 from IPDeptInput" & _
'                        " where ii01=" & strSrvDate(1) & _
'                        " and upper(ii03)=upper('" & ChgSQL(strII03) & "')"
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strII03 = Trim(Left(strII03, Len(strII03) - 4)) & "," & strUpdTime & ".msg"
'               '加了時間還是有可能重覆,再加當日當時筆數(流水號)
'               strExc(0) = "select ii03 from IPDeptInput" & _
'                           " where ii01=" & strSrvDate(1) & _
'                           " and upper(ii03)=upper('" & ChgSQL(strII03) & "')"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  strExc(0) = "select count(*) from IPDeptInput" & _
'                              " where ii01=" & strSrvDate(1)
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     strExc(0) = Val(RsTemp.Fields(0)) + 1
'                  End If
'                  strII03 = Trim(Left(strII03, Len(strII03) - 4)) & "," & strExc(0) & ".msg"
'               End If
'            End If
'         End If
         
         'Add By Sindy 2018/7/5 解析主旨抓出對應的案件性質,副檔名
         Call PUB_IPDept_ComparisonCP(strII17, "", strCP01, strCP02, strCP03, strCP04, "", strCP09, strCP10)
         
         If bolConnect = False Then cnnConnection.BeginTrans: bolConnect = True
         
         '存實體檔案到File Server
         '檢查若為個案必須儲存到卷宗區
         'CANCEL BY SONIA 2016/4/8 與DAVID討論暫不放卷宗區,以免系統直接放入沒用的信件
         'If strCP09 <> "" Then
         '   Set f = fs.GetFile(txtPathIPDept.Text & "\" & oFile.Name)
         '   bolSaveEFile = SaveAttFile_PDF(strCP09, txtPathIPDept.Text & "\" & oFile.Name, strII03, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), True, "F","Y" , , , stFtpPath)
         'Else
         'END 2016/4/8
         '國外部信件區
'            strExc(0) = "select count(*) from IPDeptInput" & _
'                        " where ii01=" & strSrvDate(1)
'            intI = 1
'            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'            'Modify By Sindy 2016/10/4
'            If intI = 1 Then
'               intII03 = Val(RsTemp.Fields(0)) + 1
'            Else
'               intII03 = 1
'            End If
'            strII03 = "F" & Format(intII03, "0000")
'            '2016/10/4 END
            'Modify By Sindy 2019/12/2 自動給號,才能 Keep PKey
            strII03 = AutoNoByDate("F", 4)
            '2019/12/2 END
            strII03_2 = strSrvDate(1) & strUpdTime & "." & strII03 & ".msg"
            bolSaveEFile = PUB_PutFtpFile(oForm.txtPathIPDept.Text & "\" & oFile.Name, strSrvDate(1), strII03_2, stFtpPath, "IPDEPTINPUT")
         'End If  'CANCEL BY SONIA 2016/4/8
         If bolSaveEFile = True Then
            
            'Add By Sindy 2017/3/27 個案存卷宗區
            stReName = "": strCaseNo = ""
            'If strII05 = "1" And strCP09 <> "" Then
            'Modify By Sindy 2022/5/24 FCL案仍直接歸個案
            'If strCP09 <> "" And (Val(strSrvDate(1)) < 外專信件沖銷啟用日 Or Left(strCP12, 2) <> "F2" Or strCP01 = "FCL") Then
            'Modify By Sindy 2023/4/25 + CF案件要從系統中做信件沖銷
            'Modify By Sindy 2023/7/18 FC案要歸卷 + And substr(strCP12, 1, 2) <> "F1"
            If strCP09 <> "" Then
               If (Left(strCP12, 2) <> "F2" Or strCP01 = "FCL") _
                  And Not (Val(strSrvDate(1)) >= 外商CF信件沖銷啟用日 And PUB_IPDept_IsCFMail(strII06) = True And Mid(strCP12, 1, 2) <> "F1") Then
                  Set f = fs.GetFile(oForm.txtPathIPDept.Text & "\" & oFile.Name)
                  
                  'Add By Sindy 2018/12/25 若是由INBOUND INBOUND@taie.com.tw轉來
                  '但寄件者是huanhua <huanhua@sino-elite-ip.com>的郵件，匯入卷宗區時則中文附檔說明
                  '為.msg的"代理人來函(ALTR)"。
                  If InStr(UCase(strII11), UCase("huanhua@sino-elite-ip.com")) > 0 And _
                     strCP01 = "P" And _
                     GetPrjNation1(strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04) <> "000" Then
                     strII03_2 = Replace(strII03_2, ".msg", ".ALTR.msg")
                  Else
                  '2018/12/25 END
                     strII03_2 = Replace(strII03_2, ".msg", ".rx.msg")
                  End If
                  
                  'Modify By Sindy 2019/12/11 本所案號流水號要存足碼
                  'Modify By Sindy 2020/2/19 電子檔名,本所案號使用函數 PUB_CaseNo2FileName
'                     stReName = Trim(strCP01) & Trim(strCP02) & _
'                                 IIf(Val(Trim(strCP03)) = 0 And Val(Trim(strCP04)) = 0, "", "-" & strCP03) & _
'                                 IIf(Val(Trim(strCP04)) = 0, "", "-" & Format(strCP04, "00")) & "." & strCP10 & "." & _
'                                 strII03_2
                  stReName = PUB_CaseNo2FileName(strCP01, strCP02, strCP03, strCP04) & _
                             "." & strCP10 & "." & strII03_2
                  strCaseNo = strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 <> "000", strCP03 & "-" & strCP04, "")
                  
                  If UCase(oForm.Name) = UCase("frmTaOutLook") Then
                     oForm.WLog_Day "欲存: " & strCaseNo & "(" & strCP09 & ")", 國外部收件信箱
                  End If
                  
                  '+ save cpp04
                  bolSaveEFile = SaveAttFile_PDF(strCP09, oForm.txtPathIPDept.Text & "\" & oFile.Name, stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), True, "F", "Y", , , , strII17, strErrText, False)
                  If bolSaveEFile = False Then
                     If UCase(oForm.Name) = UCase("frmTaOutLook") Then
                        oForm.WLog_Day "SaveAttFile_PDF 失敗 : GoTo ErrHand", 國外部收件信箱
                     End If
                     GoTo ErrHand '失敗結束
                  Else
                     If UCase(oForm.Name) = UCase("frmTaOutLook") Then
                        oForm.WLog_Day "SaveAttFile_PDF 成功", 國外部收件信箱
                     End If
                  End If
               End If
            End If
            
'            'Add By Sindy 2017/7/20
'            If UCase(oForm.Name) = UCase("frmTaOutLook") Then
'               If strCaseNo = "" Then '未歸案號
'                  oForm.WLog_Day "找不到對應案件 : " & vbCrLf & strII17 & vbCrLf & _
'                  "==>收到日期:" & IIf(strII12 <> "", Format(ChangeTStringToWDateString(strII12 - 19110000), "YYYY/MM/DD"), "") & " " & strII13 & " 寄件者:" & strII11 & vbCrLf, 國外部收件信箱
'               Else
'                  oForm.WLog_Day strII17 & vbCrLf & _
'                  "==>收到日期:" & IIf(strII12 <> "", Format(ChangeTStringToWDateString(strII12 - 19110000), "YYYY/MM/DD"), "") & " " & strII13 & " 寄件者:" & strII11 & vbCrLf & _
'                  "==>" & strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04 & " : " & strCP09 & "(" & strCP10 & ")==>" & stReName & vbCrLf, 國外部收件信箱
'               End If
'            End If
'            '2017/7/20 END
            
            If bolSaveEFile = True Then
               '存資料到DB
               'MODIFY BY SONIA 2016/4/11 不放卷宗區所以也不存總收文號,否則不同日之信件會開到同一個MSG檔,例:
               'strSql = "insert into IPDeptInput(ii01,ii02,ii03,ii04,ii05,ii06,ii11,ii12,ii13,ii14,ii17,ii18)" & _
                        " values(" & strSrvDate(1) & "," & strUpdTime & _
                        ",'" & ChgSQL(strII03) & "','" & strUserNum & "'" & _
                        ",'" & strII05 & "','" & strII06 & "'" & _
                        "," & CNULL(ChgSQL(strII11)) & "," & CNULL(strII12) & "," & CNULL(strII13) & _
                        ",'" & ChgSQL(stFtpPath) & "','" & strII17 & "','" & strII05 & "')"
               'Modify By Sindy 2016/4/27 寄件者長度太長,截取長度100 ex.MAILER-DAEMON@heramailgw12.hera.idc.justsystem.co.jp[MAILER-DAEMON@heramailgw12.hera.idc.justsystem.co.jp]
               If Len(strII11) > 100 Then
                  strII11 = Mid(strII11, 1, 100)
               End If
               '2016/4/27 END
               'Add By Sindy 2017/8/28
               If strII05 <> "" Then '分類
                  strExc(0) = "select decode('" & strII05 & "'," & Show國外部信件分類 & ",'" & strII05 & "') 分類 from dual"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If strII18 <> "" Then '關鍵字
                        strII18 = Replace(RsTemp.Fields(0) & ";" & strII18, ";;", ";")
                     Else
                        strII18 = RsTemp.Fields(0)
                     End If
                  End If
               End If
               '2017/8/28 END
               'Modify By Sindy 2020/8/26 And Len(strRecipients_all) <= 200 : 收件者太多就不要存值了
               strII18 = strII18 & IIf(strRecipients_all <> "" And Len(strRecipients_all) <= 200, ";收件者:" & strRecipients_all, "") 'Add By Sindy 2017/10/20 + 原收件者
               'Modify By Sindy 2017/3/31 + ii19.總收文號
               strSql = "insert into IPDeptInput(ii01,ii02,ii03,ii04,ii05,ii06,ii11,ii12,ii13,ii14,ii17,ii18,ii19,ii23,ii24,ii25,ii26)" & _
                        " values(" & strSrvDate(1) & "," & strUpdTime & _
                        ",'" & ChgSQL(strII03) & "','" & strUserNum & "'" & _
                        ",'" & strII05 & "','" & strII06 & "'" & _
                        "," & CNULL(ChgSQL(strII11)) & "," & strII12 & "," & CNULL(strII13) & _
                        ",'" & ChgSQL(stFtpPath) & "','" & strII17 & "'," & CNULL(ChgSQL(strII18)) & _
                        "," & CNULL(strCP09) & "," & CNULL(strCP01) & "," & CNULL(strCP02) & "," & CNULL(strCP03) & "," & CNULL(strCP04) & ")"
               cnnConnection.Execute strSql
               
               'Add By Sindy 2022/2/22
               '檢查信件是否同時有寄ipdept及patent信箱
               Dim bolPatentAndIPDept As Boolean, bolExistPatent As Boolean
               Dim strChkSub As String
               bolPatentAndIPDept = False
               bolExistPatent = False
               If InStr(UCase(strRecipients_all), UCase("patent@taie.")) > 0 And _
                  InStr(UCase(strRecipients_all), UCase("ipdept@taie.")) > 0 Then
                  bolPatentAndIPDept = True
                  '檢查專利處是否有此筆郵件
                  strSql = "select pi01,pi03 from patentinput" & _
                           " where pi17 = '" & strII17 & "'" & _
                           " and pi11 = '" & ChgSQL(strII11) & "' and pi12 = " & strII12 & " and pi13 = " & strII13 & _
                           " order by pi01 desc,pi03 desc"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                  If intI = 1 Then
                     bolExistPatent = True '這狀況是不應該發生的
                  End If
               End If
               '2022/2/22 END
               
               'Modify By Sindy 2018/3/6 改為共同函數
               'Modify By Sindy 2019/4/15 + , strRecipients_all, strII06
               bolForKeyWordDel = PUB_OutLookForKeyWordDel("F", strRecipients_1, strII11, strII17, strRecipients_all, strII06)
               '要排除的垃圾郵件:直接刪除
               'admin@taie.com.tw:Admin [admin@taie.com.tw] 出差中，請與職務代理人聯絡
               '@taie.com.tw & 主旨有”自動回覆””目前不在回覆”
               'Modify By Sindy 2017/8/31 + 薛經理:不在辦公室
               'Add By Sindy 2020/5/28 檢查寄件者是否為本所員工
               strExc(0) = "SELECT st01,st02,st04,st69 From staff" & _
                           " where st01>'6' and st01<'F'" & _
                           " AND substr(st01,4,1)<>'9'" & _
                           " AND (st04='1' or (st04='2' and st51>=" & DBDATE(DateAdd("d", -7, Format(strSrvDate(1), "####/##/##"))) & "))" & _
                           " AND st01 NOT IN('60000')" & _
                           " AND substr(st03,1,1)<>'R'" & _
                           " AND (InStr(upper('" & ChgSQL(strII11) & "'), upper(st02)) > 0 Or (InStr(upper('" & ChgSQL(strII11) & "'), upper(ST01)) > 0 and InStr(upper('" & ChgSQL(strII11) & "'), upper('/O=TAIE/OU=DOMAIN')) > 0))"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               bolTaieStaff = False
               If intI = 1 Then
                  bolTaieStaff = True
               End If
               '2020/4/8 END
'               If InStr(UCase(ChgSQL(strII11)), UCase("admin@taie.com.tw")) > 0 Or _
'                  (InStr(UCase(ChgSQL(strII11)), UCase("@taie.com.tw")) > 0 And _
'                   (InStr(strII17, "自動回覆") > 0 Or _
'                   InStr(strII17, "目前不在回覆") > 0 Or _
'                   InStr(strII17, "不在辦公室") > 0)) Or _
'                   bolForKeyWordDel = True Then
               If InStr(UCase(ChgSQL(strII11)), UCase("admin@taie.com.tw")) > 0 Or _
                  (bolTaieStaff = True And _
                    (InStr(strII17, "自動回覆") > 0 Or InStr(strII17, "目前不在回覆") > 0 Or InStr(strII17, "不在辦公室") > 0) _
                  ) Or _
                  bolForKeyWordDel = True Then
               '2020/5/28 END
                  If strII06 = "" Then
                     strSql = "update IPDeptInput set" & _
                              " ii07='Y',ii08=" & strSrvDate(1) & _
                              ",ii09=" & strUpdTime & ",ii10='" & strUserNum & "'" & _
                              ",ii16=" & strSrvDate(1) & ",ii06=null" & _
                              " where ii01=" & strSrvDate(1) & _
                                " and ii02=" & strUpdTime & _
                                " and ii03='" & ChgSQL(strII03) & "'"
                     cnnConnection.Execute strSql
                     'Modify By Sindy 2019/4/15 PUB_OutLookForKeyWordDel函數會回傳strII06變數值
                     'strII06 = "" '不須轉寄
                  Else
                     '剔除某些欲轉寄人員
                     strSql = "update IPDeptInput set" & _
                              " ii06=" & CNULL(strII06) & _
                              " where ii01=" & strSrvDate(1) & _
                                " and ii02=" & strUpdTime & _
                                " and ii03='" & ChgSQL(strII03) & "'"
                     cnnConnection.Execute strSql
                  End If
               End If
               strTo = strII06 '轉寄人員
               
               'Add By Sindy 2017/3/27 有收受者並且有分類者，直接 [轉寄]
               'Modify By Sindy 2022/2/23 + Or bolPatentAndIPDept = True
               If (strTo <> "" And strII05 <> "") Or bolPatentAndIPDept = True Then
                  strIRUpdTime = Right("000000" & ServerTime, 6)
                  
                  '專利處,要進系統
                  'If strII05 = "4" Or InStr(UCase(strTo), UCase("patent")) > 0 Then
                  'Modify By Sindy 2022/2/23 + Or bolPatentAndIPDept = True
                  If InStr(UCase(strTo), UCase("patent")) > 0 Or bolPatentAndIPDept = True Then
                     'Add By Sindy 2016/9/13
                     '檢查收受者若是有專利處(patent)信件也複製一份至專利處收件夾資料
                     '若純寄patent則國外部信件要上刪除日期註記---取消
                     '若有其他單位人員則不需要上註記---取消
                     '*****
                     '專利處收件夾資料
'                     strExc(0) = "select count(*) from PatentInput" & _
'                                 " where PI01=" & strSrvDate(1)
'                     intI = 1
'                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        intII03 = Val(RsTemp.Fields(0)) + 1
'                     Else
'                        intII03 = 1
'                     End If
'                     strPI03 = "P" & Format(intII03, "0000")
                     'Modify By Sindy 2019/12/2 自動給號,才能 Keep PKey
                     strPI03 = AutoNoByDate("P", 4)
                     '2019/12/2 END
                     strPI03_2 = strSrvDate(1) & strIRUpdTime & "." & strPI03 & ".msg"
                     '存實體檔案到PatentInput
                     bolSaveEFile = PUB_PutFtpFile(oForm.txtPathIPDept.Text & "\" & oFile.Name, strSrvDate(1), strPI03_2, stFtpPath, UCase("PatentInput"))
                     If bolSaveEFile = True Then
                        '存資料到專利處收件夾資料
                        strSql = "insert into PatentInput(PI01,PI02,PI03,PI04,PI11,PI12,PI13,PI14,PI17)" & _
                                 " values(" & strSrvDate(1) & "," & strIRUpdTime & _
                                 ",'" & strPI03 & "','" & strUserNum & "'" & _
                                 "," & CNULL(ChgSQL(strII11)) & "," & strII12 & "," & CNULL(strII13) & _
                                 ",'" & ChgSQL(stFtpPath) & "','" & strII17 & "')"
                        cnnConnection.Execute strSql
                        
                        '***** 使用專利處的分信規則 *****
                        Dim strPI11 As String, strPI06 As String, strPI15 As String
                        Dim strPI05 As String
                        strPI05 = PUB_Patent_ToSortOut(oForm, strII17, strPI11, strPI06, strCP01, strCP02, strCP03, strCP04, strPI15)
                        'Modify By Sindy 2022/2/23
                        If bolPatentAndIPDept = False Or bolExistPatent = True Or InStr(UCase(strTo), UCase("patent")) > 0 Then
                        '2022/2/23 END
                           If strPI15 <> "" Then '關鍵字
                              strPI15 = "IPDept;" & strPI15
                           Else
                              strPI15 = "IPDept"
                           End If
                        End If
                        If strPI05 <> "" Then
                           strExc(0) = "select decode('" & strPI05 & "'," & Show專利處信件分類 & ",'" & strPI05 & "') 分類 from dual"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 1 Then
                              'strPI15 = strPI15 & ";" & RsTemp.Fields(0)
                              If strPI15 <> "" Then '關鍵字
                                 strPI15 = Replace(strPI15 & ";" & RsTemp.Fields(0), ";;", ";") 'Replace(RsTemp.Fields(0) & ";" & strPI15, ";;", ";")
                              Else
                                 strPI15 = RsTemp.Fields(0)
                              End If
                           End If
                        End If
                        'Add By Sindy 2022/2/24 + 原收件者
                        If strRecipients_all <> "" And Len(strRecipients_all) <= 200 Then
                           If strPI15 <> "" Then strPI15 = strPI15 & ";"
                           strPI15 = strPI15 & "收件者:" & strRecipients_all
                        End If
                        '2022/2/24 END
                        
                        'Add By Sindy 2022/2/23
                        If bolPatentAndIPDept = True Then '視為Patent的第一封信
                           '不分信直接刪除
                           If PUB_OutLookForKeyWordDel("P", "", strII11, strII17, strRecipients_all, strPI06) = True Then
                              strSql = "update PatentInput set" & _
                                       " pi07='Y',pi08=" & strSrvDate(1) & _
                                       ",pi09=" & strIRUpdTime & ",pi10='" & strUserNum & "'" & _
                                       ",pi16=" & strSrvDate(1) & ",pi06=null" & _
                                       " where pi01=" & strSrvDate(1) & _
                                         " and pi02=" & strIRUpdTime & _
                                         " and pi03='" & strPI03 & "'"
                              cnnConnection.Execute strSql
                           End If
                        Else
                        '2022/2/23 END
                           'Add By Sindy 2019/7/17 增加副本給專利處主管
                           If OL_SendNotifyMailCC("IPDept", "Patent", oForm.txtPathIPDept.Text & "\" & oFile.Name, strII17, strSrvDate(1), strIRUpdTime, strPI03, OL_PatMailCC, strSrvDate(1), strIRUpdTime) = False Then
                              GoTo ErrHand
                           End If
                        End If
                        
                        '記錄本所案號
                        If strCP01 <> "" Then
                           strPI15 = strPI15 & IIf(strPI15 <> "", ";", "") & strCP01 & strCP02 & strCP03 & strCP04
                           
                           '別的信箱轉至PATENT信箱的信件，郭經理尚未操作確認前，在PATENT分信時若已有本所案號者，則將郭經理的記錄上刪除人員QPGMR
                           'Add By Sindy 2020/6/17 副本人員未上核銷,因有本所案號了,系統自動核銷
                           strSql = "update InputRecord set " & _
                                    " ir08=" & strSrvDate(1) & ",ir09=" & strIRUpdTime & ",ir10='QPGMR'" & _
                                    " where ir01=" & strSrvDate(1) & _
                                    " and ir03='" & strPI03 & "'" & _
                                    " and ir04='79075'" & _
                                    " and ir24='Y'" & _
                                    " and ir08=0"
                           cnnConnection.Execute strSql, intI
                           '2020/6/17 END
                        End If
                        strSql = "update PatentInput set PI05=" & CNULL(strPI05) & _
                                 ",PI06=" & CNULL(strPI06) & _
                                 ",PI15=" & CNULL(ChgSQL(strPI15)) & _
                                 ",PI18=" & CNULL(strCP01) & _
                                 ",PI19=" & CNULL(strCP02) & _
                                 ",PI20=" & CNULL(strCP03) & _
                                 ",PI21=" & CNULL(strCP04) & _
                                 " where PI01=" & strSrvDate(1) & _
                                 " and PI02=" & strIRUpdTime & _
                                 " and PI03='" & strPI03 & "'"
                        cnnConnection.Execute strSql
                        '***** END
                        
                        'Modify By Sindy 2022/2/23
                        '更新國外部收件夾資料-記錄專利處信件流水號
                        strSql = "update IPDeptInput set" & _
                                    " ii10='" & strUserNum & "',ii15='" & strPI03 & "'" & _
                                    " where ii01=" & strSrvDate(1) & _
                                      " and ii02=" & strUpdTime & _
                                      " and ii03='" & ChgSQL(strII03) & "'"
                        cnnConnection.Execute strSql
                        'Add By Sindy 2022/8/11 新增郵件轉信讀取記錄(直接沖銷)
                        strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR15,IR08,IR09,IR10)" & _
                                    " values(" & strSrvDate(1) & _
                                             "," & strUpdTime & _
                                             ",'" & ChgSQL(strII03) & "'" & _
                                             ",'patent'," & strSrvDate(1) & "," & _
                                             strUpdTime & ",'" & strUserNum & "','Y'," & _
                                             strSrvDate(1) & "," & strUpdTime & ",'" & strUserNum & "')"
                        cnnConnection.Execute strExc(0)
                        
                        'Modify By Sindy 2022/5/24
'                        If (InStr(UCase(strTo), UCase("patent")) > 0 And Val(strSrvDate(1)) < 外專信件沖銷啟用日) Or _
'                           (UCase(Trim(strTo)) = UCase("patent")) Then
                        If UCase(Trim(strTo)) = UCase("patent") Then
                        '2022/5/24 END
                           '更新國外部收件夾資料-記錄直接上刪除實體檔日期
                           strSql = "update IPDeptInput set" & _
                                    " ii08=" & strSrvDate(1) & ",ii09=" & strIRUpdTime & ",ii10='" & strUserNum & "',ii16=" & strSrvDate(1) & _
                                    " where ii01=" & strSrvDate(1) & _
                                      " and ii02=" & strUpdTime & _
                                      " and ii03='" & ChgSQL(strII03) & "'"
                           cnnConnection.Execute strSql
                        End If
                        
                        'Add By Sindy 2022/2/23 比較2個信箱的狀況(分類狀況)
                        If bolPatentAndIPDept = True Then
                           '專利處=其他並且國外部=個案時,專利處信件刪除
                           If strPI05 = "7" And strII05 = "1" Then
                              bolPatentDel = True
                              strSql = "update PatentInput set" & _
                                       " pi07='Y',pi08=" & strSrvDate(1) & _
                                       ",pi09=" & strIRUpdTime & ",pi10='" & strUserNum & "'" & _
                                       ",pi16=" & strSrvDate(1) & ",pi06=null" & _
                                       " where pi01=" & strSrvDate(1) & _
                                         " and pi02=" & strIRUpdTime & _
                                         " and pi03='" & strPI03 & "'"
                              cnnConnection.Execute strSql
                              '記錄於Patent有關
                              strSql = "update IPDeptInput set" & _
                                       " ii18='Patent;'||ii18" & _
                                       " where ii01=" & strSrvDate(1) & _
                                         " and ii02=" & strUpdTime & _
                                         " and ii03='" & ChgSQL(strII03) & "'"
                              cnnConnection.Execute strSql
                           '專利處有分到案號
                           ElseIf (strCP01 <> "" And strCP02 <> "") And InStr(UCase(strTo), UCase("patent")) > 0 Then
                              '檢查案號是否相同,不同時國外部改成"其他"由人工判斷
                              If InStr(strII18, "(CFP-)") > 0 Or InStr(strII18, "(CPS-)") > 0 Then
                                 If InStr(strII18, "(CFP-)") > 0 Then
                                    strChkSub = Replace(strII17, "CFP--", "CFP-")
                                 End If
                                 If InStr(strII18, "(CPS-)") > 0 Then
                                    strChkSub = Replace(strII17, "CPS--", "CPS-")
                                 End If
                                 '不同時,國外部改成"其他"由人工判斷
                                 If InStr(UCase(strChkSub), strCP01 & "-" & strCP02) = 0 And _
                                    InStr(UCase(strChkSub), strCP01 & " " & strCP02) = 0 And _
                                    InStr(UCase(strChkSub), strCP01 & "-" & Val(strCP02)) = 0 And _
                                    InStr(UCase(strChkSub), strCP01 & " " & Val(strCP02)) = 0 Then
                                    strSql = "update IPDeptInput set" & _
                                             " ii05='Z',ii06=null,ii08=0,ii09=null,ii10=null,ii16=0,ii15=null" & _
                                             " where ii01=" & strSrvDate(1) & _
                                               " and ii02=" & strUpdTime & _
                                               " and ii03='" & ChgSQL(strII03) & "'"
                                    cnnConnection.Execute strSql
                                    '取消於IPDept有關
                                    If Mid(strPI15, 1, 7) = "IPDept;" Then
                                       strPI15 = Mid(strPI15, 8)
                                    ElseIf Mid(strPI15, 1, 6) = "IPDept" Then
                                       strPI15 = Mid(strPI15, 7)
                                    End If
                                    strSql = "update PatentInput set" & _
                                             " pi15=" & CNULL(ChgSQL(strPI15)) & _
                                             " where pi01=" & strSrvDate(1) & _
                                             " and pi02=" & strIRUpdTime & _
                                             " and pi03='" & strPI03 & "'"
                                    cnnConnection.Execute strSql
                                 End If
                              End If
                           End If
                        End If
                        '2022/2/23 END
                     Else
                        GoTo ErrHand
                     End If
                     '要去掉Patent以免下列程式將信寄到Patent
                     If InStr(UCase(strTo), UCase("patent")) > 0 Then 'Modify By Sindy 2022/2/23 +if
                        strTo = Replace(strTo, "patent", "")
                        strTo = Replace(strTo, ";;", ";")
                        If strTo = ";" Then strTo = ""
                        If strTo <> "" Then
                           If Left(strTo, 1) = ";" Then strTo = Mid(strTo, 2)
                           If Right(strTo, 1) = ";" Then strTo = Mid(strTo, 1, Len(strTo) - 1)
                        End If
                     End If
                  End If
                  'Add By Sindy 2022/2/23 暫時收到通知,觀察狀況用
                  If bolPatentAndIPDept = True Then
                     strSql = "select * from PatentInput" & _
                              " where pi01 = " & strSrvDate(1) & _
                                " and pi02 = " & strIRUpdTime & _
                                " and pi03 = '" & strPI03 & "';" & vbCrLf
                     strSql = strSql & "select * from IPDeptInput" & _
                              " where ii01 = " & strSrvDate(1) & _
                                " and ii02 = " & strUpdTime & _
                                " and ii03 = '" & ChgSQL(strII03) & "';"
                     oForm.WLog_Day "【信件同時寄給patent@taie.com.tw和ipdept@taie.com.tw】" & vbCrLf & strII17 & vbCrLf & strSql, 國外部收件信箱
'                     PUB_SendMail strUserNum, "97038", "", _
'                        "【信件同時寄給patent@taie.com.tw和ipdept@taie.com.tw => 暫時收到通知,觀察狀況用】" & strII17, strII17 & vbCrLf & vbCrLf & strSql, , oForm.txtPathIPDept.Text & "\" & oFile.Name, , , , , , , , True, False, , , False, , , False
                  End If
                  '2022/2/23 END
                  
                  '收受者非專利處
                  'If strII05 <> "4" Then
                  'If UCase(strTo) <> UCase("patent") Then
                  strToSys = "": strExc(10) = "": bolSendMailErr = False
                  strToSpec = "" 'Add By Sindy 2024/6/19
                  If strTo <> "" Then
                     'Modify By Sindy 2022/5/24 外專部門(F22,F23)信件收錄進系統收件區,其他維持Outlook轉寄
                     'Modify By Sindy 2022/8/10 David說新知,財務,開拓也不要列入系統收件區
                     'If Val(strSrvDate(1)) >= 外專信件沖銷啟用日 And strII05 <> "6" And strII05 <> "7" And strII05 <> "8" Then
                     If strII05 <> "6" And strII05 <> "7" And strII05 <> "8" Then
                        tmpArr = Split(strTo, ";")
                        For j = 0 To UBound(tmpArr)
                           If tmpArr(j) <> "" Then
                              'Add By Sindy 2023/7/26 若為E-Mail則抓帳號即可
                              If InStr(tmpArr(j), "@") > 0 Then
                                 tmpArr(j) = Mid(tmpArr(j), 1, InStr(tmpArr(j), "@") - 1)
                              End If
                              '2023/7/26 END
                              If PUB_GetST03(CStr(tmpArr(j))) <> "" Then
                                 '外專部門(F22,F23)信件收錄進系統收件區
                                 'Modify By Sindy 2023/4/25 + CF案件要從系統中做信件沖銷
                                 If Trim(PUB_GetST03(CStr(tmpArr(j)))) = "F22" Or _
                                    Trim(PUB_GetST03(CStr(tmpArr(j)))) = "F23" Or _
                                    (Val(strSrvDate(1)) >= 外商CF信件沖銷啟用日 And PUB_IPDept_IsCFMail(CStr(tmpArr(j))) = True) Then
                                    strCaseNo = ""
                                    '收錄進系統收件區的收受者
                                    If strToSys <> "" Then strToSys = strToSys & ";"
                                    strToSys = strToSys & Trim(tmpArr(j))
                                    '記錄要發通知信的收受者
                                    'Modify By Sindy 2025/5/13 先記錄在暫存檔,等時間到才一起寄發通知信
'                                    If InStr(m_strMailTo, Trim(tmpArr(j))) = 0 Then
'                                       If m_strMailTo <> "" Then m_strMailTo = m_strMailTo & ";"
'                                       m_strMailTo = m_strMailTo & Trim(tmpArr(j))
'                                    End If
                                    '寫入要發通知信的人員
                                    'CaseUseMemo:
                                    'cum01 = 'F'身份類別
                                    'cum02 = 收受者
                                    'cum04 = 通知信分類:01-加密 02-個人系統收件區
                                    'cum05 = '04'TaRevOutLook系統自動分信通知信
                                    'cum06 = 操作人員
                                    strExc(0) = "select cum02 from CaseUseMemo" & _
                                                " where cum05='04'" & _
                                                  " and cum02=" & CNULL(Trim(tmpArr(j))) & _
                                                  " and cum04='02'" & _
                                                  " and cum01='F'"
                                    intI = 1
                                    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                                    If intI = 0 Then
                                       strExc(0) = "insert into CaseUseMemo(cum01,cum02,cum03,cum04,cum05)" & _
                                                   " values('F'," & CNULL(Trim(tmpArr(j))) & ",'0','02','04')"
                                       cnnConnection.Execute strExc(0), intI
                                    End If
                                    '2025/5/13 END
                                    '新增郵件轉信讀取記錄
                                    strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR15)" & _
                                                " values(" & strSrvDate(1) & _
                                                         "," & strUpdTime & _
                                                         ",'" & ChgSQL(strII03) & "'" & _
                                                         ",'" & tmpArr(j) & "'," & strSrvDate(1) & "," & _
                                                         strUpdTime & ",'" & strUserNum & "','Y')"
                                    cnnConnection.Execute strExc(0)
                                    'Add By Sindy 2024/6/11 洪琬姿副理提:F11人員出差期間，若有分信至該人員時，加發副本至該人員處，以利訊息接收。
                                    If PUB_GetST03(CStr(tmpArr(j))) = "F11" Then
                                       bolRest = CheckIsPersonRest(CStr(tmpArr(j)), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), strRestKind)
                                       'strRestKind 休假狀況:1.請假3.出差
                                       If bolRest = True And strRestKind = "3" Then
                                          'Outlook要獨立寄發,不轉職代
                                          If strToSpec <> "" Then strToSpec = strToSpec & ";"
                                          strToSpec = strToSpec & Trim(tmpArr(j))
                                       End If
                                    End If
                                    '2024/6/11 END
                                 '非外專部門
                                 Else
                                    '維持Outlook轉寄的收受者
                                    If strExc(10) <> "" Then strExc(10) = strExc(10) & ";"
                                    strExc(10) = strExc(10) & Trim(tmpArr(j))
                                    
                                    'Add By Sindy 2022/8/11
                                    '新增郵件轉信讀取記錄(直接沖銷)
                                    strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR15,IR08,IR09,IR10)" & _
                                                " values(" & strSrvDate(1) & _
                                                         "," & strUpdTime & _
                                                         ",'" & ChgSQL(strII03) & "'" & _
                                                         ",'" & tmpArr(j) & "'," & strSrvDate(1) & "," & _
                                                         strUpdTime & ",'" & strUserNum & "','Y'," & _
                                                         strSrvDate(1) & "," & strUpdTime & ",'" & strUserNum & "')"
                                    cnnConnection.Execute strExc(0)
                                 End If
                              '無部門別
                              Else
                                 '維持Outlook轉寄的收受者
                                 If strExc(10) <> "" Then strExc(10) = strExc(10) & ";"
                                 strExc(10) = strExc(10) & Trim(tmpArr(j))
                                 
                                 'Add By Sindy 2022/8/11
                                 '新增郵件轉信讀取記錄(直接沖銷)
                                 strExc(0) = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR15,IR08,IR09,IR10)" & _
                                             " values(" & strSrvDate(1) & _
                                                      "," & strUpdTime & _
                                                      ",'" & ChgSQL(strII03) & "'" & _
                                                      ",'" & tmpArr(j) & "'," & strSrvDate(1) & "," & _
                                                      strUpdTime & ",'" & strUserNum & "','Y'," & _
                                                      strSrvDate(1) & "," & strUpdTime & ",'" & strUserNum & "')"
                                 cnnConnection.Execute strExc(0)
                              End If
                           End If
                        Next j
                        If strToSys <> "" Then strTo = strExc(10)
                     End If
                     
                     strII17AddText = IIf(strCaseNo <> "", "【" & strCaseNo & " Saved】", "") & IIf(pub_SaveCoRec = True, "【往來記錄 Saved】", "") & IIf(bolPatentDel = True, "【非該單位信件，請轉寄回patent信箱】", "")
                     '使用Outlook轉寄
                     If strTo <> "" Or strToSpec <> "" Then
                        '其他型態信件均用信包信方式寄出
                        strErrText = "objMail.Class(" & objMail.Class & ")以信包信方式寄出" & strII17
                        '以信包信方式寄出 : 2017/7/7 承辦人員在Outlook操作時,才能使用全部回覆,不然客戶Address會出不來
                        '寄件者為ipdept
                        '主旨增加,當個案且有案號時,顯示歸入那一個案號
'                        If UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then '測試資料庫
'                           strTo = Pub_GetSpecMan("電腦中心郵件檢核人員")
'                        End If
                        '********************
                        'Add By Sindy 2017/11/17 SendMail記錄
                        PStr_SendMailKey1 = strSrvDate(1)
                        PStr_SendMailKey2 = strUpdTime
                        PStr_SendMailKey3 = ChgSQL(strII03)
                        strTableName = "IPDEPTINPUT"
                        strUpdWhere = "II01=" & PStr_SendMailKey1 & " and II02=" & PStr_SendMailKey2 & " and II03='" & PStr_SendMailKey3 & "'"
                        strSql = "update " & strTableName & _
                                 " set II20='S',II21=" & strSrvDate(1) & ",II22=" & Right("000000" & ServerTime, 6) & _
                                 " where " & strUpdWhere
                        cnnConnection.Execute strSql
                        '2017/11/17 END
                        '********************
                        
                        'Add By Sindy 2022/8/11
                        strContent = ""
                        If strToSys <> "" Or InStr(UCase(strII06), UCase("patent")) > 0 Then
                           If strToSys <> "" Then strContent = PUB_ReadUserData(strToSys)
                           If InStr(UCase(strII06), UCase("patent")) > 0 Then
                              If strContent <> "" Then strContent = strContent & ","
                              strContent = strContent & "patent"
                           End If
                        End If
                        If strContent <> "" Then strContent = "同時信件已分給下列人員：" & strContent & vbCrLf & vbCrLf
                        '2022/8/11 END
                        'Add By Sindy 2024/6/19
                        'Outlook要獨立寄發,不轉職代
                        If strToSpec <> "" Then
                           strExc(9) = strContent
                           If strTo <> "" Then
                              strExc(9) = strExc(9) & "同時信件已寄給下列人員：" & PUB_ReadUserData(strTo) & vbCrLf & vbCrLf
                           End If
                           If UCase(pub_DbTerminalName) = 正式資料庫電腦名稱 Then
                              PUB_SendMail "inbound", strToSpec, "", _
                                          strII17 & strII17AddText, _
                                          strExc(9) & _
                                          "信件內容參附件" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                                          "同仁收到由 INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 或該封e-mail非貴單位之案件, 請務必以【個人e-mail】及時再轉寄 ipdept@taie.com.tw 以便檢查調整並迅速轉發給適當收信部門, 謝謝!", , oForm.txtPathIPDept.Text & "\" & oFile.Name, , , , , 國外部收件信箱, , , True, False, , , False, strUpdWhere, strTableName, False, , , , , , , "1"
                              DoEvents
                              If bolMailSendOk = False Then
                                 GoTo ErrHand '失敗結束
                              End If
                           End If
                        End If
                        If strTo <> "" Then
                        '2024/6/19 END
                           If UCase(pub_DbTerminalName) = 正式資料庫電腦名稱 Then 'Add By Sindy 2020/2/5
                              'Modify By Sindy 2017/7/10
                              '*注意* 若要寄patent信箱,要傳入完整email:patent@taie.com.tw
                              If strII05 = "6" Then '新知不轉職代
                                 'del strII05 = "1" And
                                 '同仁收到由INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 請及時轉寄或通報國外部顏副理(77015@taie.com.tw)，以便國外部迅速轉發適當收信同仁, 謝謝!
                                 '同仁收到由 INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 或該封e-mail非貴單位之案件, 請務必及時再轉寄 ipdept@taie.com.tw 以便國外部迅速轉發適當收信部門, 謝謝!
                                 PUB_SendMail "inbound", strTo, "", _
                                       strII17 & strII17AddText, _
                                       strContent & _
                                       "信件內容參附件" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                                       "同仁收到由 INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 或該封e-mail非貴單位之案件, 請務必以【個人e-mail】及時再轉寄 ipdept@taie.com.tw 以便檢查調整並迅速轉發給適當收信部門, 謝謝!", , oForm.txtPathIPDept.Text & "\" & oFile.Name, , , , , 國外部收件信箱, , , True, False, , , False, strUpdWhere, strTableName, False, , , , , , , "1"
                              Else
                                 'del strII05 = "1" And
                                 '同仁收到由INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 請及時轉寄或通報國外部顏副理(77015@taie.com.tw)，以便國外部迅速轉發適當收信同仁, 謝謝!
                                 'Modify By Sindy 2020/7/20 外商個案寄發時,不在此處判斷人員休假轉職代問題
                                 '                          ,因在前頭程式要先特別處理了
                                 '同仁收到由 INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 或該封e-mail非貴單位之案件, 請務必及時再轉寄 ipdept@taie.com.tw 以便國外部迅速轉發適當收信部門, 謝謝!
                                 If strII05 = "1" And _
                                    (strCP01 = "FCT" Or strCP01 = "CFT" Or strCP01 = "CFC" Or strCP01 = "S" Or strCP01 = "T" Or strCP01 = "TM") And _
                                    strII06 <> Pub_GetSpecMan("國外部轉信外商群組") Then
                                    
                                    PUB_SendMail "inbound", strTo, "", _
                                          strII17 & strII17AddText, _
                                          strContent & _
                                          "信件內容參附件" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                                          "同仁收到由 INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 或該封e-mail非貴單位之案件, 請務必以【個人e-mail】及時再轉寄 ipdept@taie.com.tw 以便檢查調整並迅速轉發給適當收信部門, 謝謝!", , oForm.txtPathIPDept.Text & "\" & oFile.Name, , , , , 國外部收件信箱, , , True, False, , , False, strUpdWhere, strTableName, False, , , , , , , "1"
                                 Else
                                 '2020/7/20 END
                                    '同仁收到由 INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 或該封e-mail非貴單位之案件, 請務必及時再轉寄 ipdept@taie.com.tw 以便國外部迅速轉發適當收信部門, 謝謝!
                                    PUB_SendMail "inbound", strTo, "", _
                                          strII17 & strII17AddText, _
                                          strContent & _
                                          "信件內容參附件" & vbCrLf & vbCrLf & vbCrLf & vbCrLf & _
                                          "同仁收到由 INBOUND (ipdept) 所寄發之郵件, 若無法暸解為何會收到該封e-mail, 或該封e-mail非貴單位之案件, 請務必以【個人e-mail】及時再轉寄 ipdept@taie.com.tw 以便檢查調整並迅速轉發給適當收信部門, 謝謝!", , oForm.txtPathIPDept.Text & "\" & oFile.Name, , , , , 國外部收件信箱, , , , False, , , False, strUpdWhere, strTableName, False, , , , , , , "1"
                                 End If
                              End If
                              DoEvents
                              If bolMailSendOk = False Then
                                 GoTo ErrHand '失敗結束
                              End If
                              '2017/7/10 END
                           End If
                        End If
                        
                        '********************
                        'Modify By Sindy 2018/10/29
                        strExc(0) = "select ii20 from " & strTableName & _
                                    " where " & strUpdWhere
                        intI = 1
                        Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If "" & RsTemp.Fields(0) = "T" Then
                              strSql = "update " & strTableName & _
                                       " set II20='Y'" & _
                                       " where " & strUpdWhere
                              cnnConnection.Execute strSql
                           Else
                              strSql = "update " & strTableName & _
                                       " set II20='" & PStr_SendMailKey1 & PStr_SendMailKey2 & PStr_SendMailKey3 & "'" & _
                                       " where " & strUpdWhere
                              cnnConnection.Execute strSql
                              If UCase(pub_DbTerminalName) = 正式資料庫電腦名稱 Then bolSendMailErr = True
                           End If
                        End If
                        '2018/10/29 END
                        'Add By Sindy 2017/11/17 SendMail記錄
                        PStr_SendMailKey1 = ""
                        PStr_SendMailKey2 = ""
                        PStr_SendMailKey3 = ""
                        strTableName = ""
                        strUpdWhere = ""
                        '2017/11/17 END
                        '********************
                     End If
                  End If
                  
                  '更新資料
                  'Modify By Sindy 2022/6/20 更新主旨
                  strSql = "update IPDeptInput set ii01=ii01" & _
                           IIf(strII17AddText <> "", ",ii17='" & strII17 & strII17AddText & "'", "") & _
                           " where ii01=" & strSrvDate(1) & _
                             " and ii02=" & strUpdTime & _
                             " and ii03='" & ChgSQL(strII03) & "'"
                  cnnConnection.Execute strSql
                  
                  'Modify By Sindy 2017/3/31
                  If (strToSys <> "" Or strTo <> "") And bolSendMailErr = False Then
                     'Upd 轉寄(分信)記錄
                     strSql = "update IPDeptInput set" & _
                              " ii08=" & strSrvDate(1) & ",ii09=" & strIRUpdTime & ",ii10='" & strUserNum & "'" & _
                              " where ii01=" & strSrvDate(1) & _
                                " and ii02=" & strUpdTime & _
                                " and ii03='" & ChgSQL(strII03) & "'" & _
                                " and ii08=0" '*****
                     cnnConnection.Execute strSql
                  End If
                  
                  'Add By Sindy 2022/5/24 無收錄進系統收件區的信件, 才能直接上刪除實體檔日期
                  If strToSys = "" And strTo <> "" And bolSendMailErr = False Then
                  '2022/5/24 END
                     '因國外部直接用Outlook轉信,所以直接上刪除實體檔日期
                     strSql = "update IPDeptInput set" & _
                              " ii16=" & strSrvDate(1) & _
                              " where ii01=" & strSrvDate(1) & _
                                " and ii02=" & strUpdTime & _
                                " and ii03='" & ChgSQL(strII03) & "'" & _
                                " and ii08>0" '*****
                     cnnConnection.Execute strSql
                  End If
               End If
               '2017/3/27 END
               
               '刪除PC端檔案
               'Kill 刪不掉 "C:\IPdept\【轉知】(1) 經濟部智慧財產局來函，自105年4月1日起提出發明專利加速審查、專利審查高速公路與支援利用專利審查高速公路之專利申請案尚未公開者，不必再申請提早公開；(2) 經濟部智慧財產局來函，公告修正「發明專利加速審查申請書及其申請須知」、「發明專利PPH審查申請書及其申請須知」與「發明專利TW-SUPA審查申請書」.msg"
               'Kill txtPathIPDept.Text & "\" & oFile.Name
               'DoEvents
               Call fs.DeleteFile(oForm.txtPathIPDept.Text & "\" & oFile.Name)
            ElseIf UCase(oForm.Name) = UCase("frmTaOutLook") Then '單筆,失敗結束
               GoTo ErrHand
            End If
         ElseIf UCase(oForm.Name) = UCase("frmTaOutLook") Then '單筆,失敗結束
            GoTo ErrHand
         End If
         
         If bolConnect = True Then cnnConnection.CommitTrans: bolConnect = False
         PUB_IPDeptTransMail_New = True 'Modify By Sindy 2019/12/11
      End If
   Next
   oForm.LblCntIPDept.Caption = "已處理件數 / 剩餘件數：" & lngRonCnt & " / " & oFolder.files.Count '最後再讀一次資料夾的檔案數
   
'   PUB_SaveLastDate Me.Name, strUserNum & "PATH", txtPathIPDept.Text
'   MsgBox "信件轉入完成！" & IIf(oFolder.files.Count > 0, vbCrLf & vbCrLf & "(尚有未轉入的信件，詳情請至資料夾查看)", "")
'
'   GetTodayTotCnt  'add by sonia 2016/4/1 重新計算今日總筆數
'
'   cmdTrans.Enabled = True
'   oForm.TxtIPDept.Visible = False
'   oForm.LblCntIPDept.Visible = False
'   Call QueryData
   
   Screen.MousePointer = vbDefault
   Set f = Nothing
   Set fs = Nothing
   Set oFolder = Nothing
   Set oFile = Nothing
   Set oFileSys = Nothing
   Set objMail = Nothing
   Set objOutLook = Nothing
   Exit Function
   
ErrHand:
   '********************
   'Add By Sindy 2017/11/17 SendMail記錄
   PStr_SendMailKey1 = ""
   PStr_SendMailKey2 = ""
   PStr_SendMailKey3 = ""
   strTableName = ""
   strUpdWhere = ""
   '2017/11/17 END
   '********************
   Screen.MousePointer = vbDefault
   If bolConnect = True Then cnnConnection.RollbackTrans: bolConnect = False
   strErrText = strErrText & "信件轉入失敗！" & vbCrLf & IIf(Err.Number <> 0, "Err.Number:" & Err.Number & ";" & vbCrLf & Err.Description, "")
   If UCase(oForm.Name) = UCase("frmTaOutLook") Then
      oForm.WLog_Day strErrText, 國外部收件信箱
   End If
'   If Err.Number <> 0 Then MsgBox " 信件轉入失敗！" & vbCrLf & Err.Description
'   cmdTrans.Enabled = True
'   Call QueryData
   Set f = Nothing
   Set fs = Nothing
   Set oFolder = Nothing
   Set oFile = Nothing
   Set oFileSys = Nothing
   Set objMail = Nothing
   Set objOutLook = Nothing
End Function

'Add By Sindy 2025/5/13 讀取信件處理人員
Public Function PUB_TaRevMailTo(strMailBox As String) As String
Dim varTmp As Variant
Dim jj As Integer
   
   If strMailBox = "01" Then
      PUB_TaRevMailTo = Pub_GetSpecMan("國外部信件處理人")
   ElseIf strMailBox = "03" Or strMailBox = "04" Then
      'Add By Sindy 2018/2/8 玲玲說分信就她和雅娟經理在處理,休假時不須轉職代,人員休假時不收通知信
      If strMailBox = "03" Then
         PUB_TaRevMailTo = Pub_GetSpecMan("專利處信件處理人")
      ElseIf strMailBox = "04" Then
         PUB_TaRevMailTo = Pub_GetSpecMan("商標處信件處理人")
      End If
      varTmp = Split(PUB_TaRevMailTo, ";")
      PUB_TaRevMailTo = ""
      For jj = 0 To UBound(varTmp)
         '檢查是否休假
         If CheckIsPersonRest(CStr(varTmp(jj)), strSrvDate(1), Format(Left(Right("000000" & ServerTime, 6), 4), "##:##")) = False Then
            If PUB_TaRevMailTo <> "" Then PUB_TaRevMailTo = PUB_TaRevMailTo & ";"
            PUB_TaRevMailTo = PUB_TaRevMailTo & CStr(varTmp(jj))
         End If
      Next jj
      If strMailBox = "03" Then
         If PUB_TaRevMailTo = "" Then PUB_TaRevMailTo = Pub_GetSpecMan("專利處信件處理人")
      ElseIf strMailBox = "04" Then
         If PUB_TaRevMailTo = "" Then PUB_TaRevMailTo = Pub_GetSpecMan("商標處信件處理人")
      End If
      '2018/2/8 END
   End If
   If PUB_TaRevMailTo = "" Then PUB_TaRevMailTo = Pub_GetSpecMan("電腦中心郵件檢核人員")
End Function

'Add By Sindy 2025/5/13 整批發通知信
'Modify By Sindy 2025/5/14 +, bolSendNotic As Boolean: 是否要發通知信
'                          +, Optional bolExitApp As Boolean = False: 是否結束作業而觸發的動作
Public Sub TaRevOutLookBatchSendMail(strMailBox As String, bolSendNotic As Boolean, _
   Optional bolExitApp As Boolean = False)
Dim strF1xEmp As String, strF2xEmp As String 'Add By Sindy 2023/5/23
Dim varTmp As Variant 'Add By Sindy 2023/5/23
Dim jj As Integer
Dim rsA As New ADODB.Recordset
Dim intURGENT As Integer 'Add By Sindy 2019/11/14
Dim strMailTo As String
Dim strMailName As String
Dim strCon As String
   
   Select Case strMailBox
      Case "01"
         strMailName = 國外部收件信箱
      Case "02"
         strMailName = 國外部寄件信箱
      Case "03"
         strMailName = 專利處收件信箱
      Case "04"
         strMailName = 商標處收件信箱
      Case "05"
         strMailName = 法律所寄件信箱
   End Select
   
   'Add By Sindy 2025/5/14
   'bolExitApp = False: 不是結束作業才觸發的動作,才做此通知
   If bolSendNotic = True And bolExitApp = False Then '要發通知信
   '2025/5/14 END
      'Modify By Sindy 2017/12/27 工作天才要通知
      If ChkWorkDay(strSrvDate(1)) = True And _
         (Format(time, "HHMMSS") >= "080000" And Format(time, "HHMMSS") < "183000") Then
         '檢查是否有信件未轉寄
         If strMailBox <> "02" And strMailBox <> "05" Then '排除國外部IPDept寄信郵件
            'If UCase(pub_DbTerminalName) = 正式資料庫電腦名稱 Then '正式資料庫才發信
               strExc(0) = ""
               Select Case strMailBox
                  Case "01" '國外部IPDept收信郵件
                     strExc(0) = "SELECT COUNT(*) FROM ipdeptinput WHERE ii08=0"
                  Case "03" '專利處Patent收信郵件
                     'Modify By Sindy 2018/10/1 雅娟:取消此通知
                     'strExc(0) = "SELECT COUNT(*) FROM patentinput WHERE pi08=0"
                  Case "04" '商標處TM收信郵件
                     strExc(0) = "SELECT COUNT(*) FROM TMinput WHERE Ti08=0"
               End Select
               If strExc(0) <> "" Then
                  intI = 1
                  Set rsA = ClsLawReadRstMsg(intI, strExc(0))
                  If rsA.Fields(0) > 0 Then
                     'Add By Sindy 2019/11/14 主旨裡有 URGENT 字樣者,通知信要加有急件! => IIf(intURGENT > 0, "（有急件！）", "") &
                     intURGENT = 0
                     strExc(0) = ""
                     Select Case strMailBox
                        Case "01" '國外部IPDept收信郵件
                           strExc(0) = "SELECT COUNT(*) FROM ipdeptinput WHERE ii08=0 and instr(upper(ii17),'URGENT')>0"
                        Case "04" '商標處TM收信郵件
                           strExc(0) = "SELECT COUNT(*) FROM TMinput WHERE Ti08=0 and instr(upper(Ti17),'URGENT')>0"
                     End Select
                     If strExc(0) <> "" Then
                        intI = 1
                        Set rsA = ClsLawReadRstMsg(intI, strExc(0))
                        If rsA.Fields(0) > 0 Then
                           intURGENT = rsA.RecordCount
                        End If
                        '2019/11/14 END
                     End If
                     'Modify By Sindy 2019/11/14 + IIf(intURGENT > 0, "（有急件！）", "") &
                     PUB_SendMail strUserNum, PUB_TaRevMailTo(strMailBox), "", IIf(intURGENT > 0, "（有急件！）", "") & "注意：" & strMailName & "尚有未轉寄信件待處理！", "同主旨", , , , , , , , , , IIf(strMailBox = "01", False, True), False, , , False, , , False
                  End If
               End If
            'End If
            
            If strMailBox = "01" Then
               'Modify By Sindy 2018/10/29 信件有遺失,轉寄資訊正常,但確實寄信備份網頁系統找不到信件
               'select ii08,ii09,ii20,ii21,ii22,ii17 from ipdeptinput where ii01='20181025' and ii03 in('F0292','F0304','F0293','F0262');
               '/*
               '      II08       II09 II20                       II21       II22 II17
               '---------- ---------- -------------------- ---------- ---------- --------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
               '  20181025     141308 Y                      20181025     141310 未傳遞的主旨: Mail Delivery Failure
               '  20181026     143250 Y                      20181026     143256 Mail Delivery Failure
               '  20181026     143249 Y                      20181026     143255 IMPORTANT NOTICE
               '  20181026     143249 Y                      20181026     143254 Out of Office Notice
               '*/
               strExc(0) = "select count(*) from ipdeptinput where ii20<>'Y' and ii20 is not null" & _
                           " and ii01>=20181001" & _
                           " order by ii01,ii02"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If RsTemp.Fields(0) > 0 And ChkWorkDay(strSrvDate(1)) = True Then
   '                  PUB_SendMail strUserNum, "97038", "", "【TaRevOutLook】檢查信件是否有遺失(" & RsTemp.Fields(0) & "筆)", strExc(0), , , , , , , , , , , False, , , False, , , False
                  End If
               End If
               '2018/10/29 END
            End If
         End If
      End If
   End If
   
   '要發個人的通知信
   If bolSendNotic = True Then
      '目前只有inbound有個人通知信件
      If strMailBox = "01" Then
         strCon = " and cum01='F'"
      Else
         Set rsA = Nothing
         Exit Sub
      End If
      'Add By Sindy 2022/5/25
      'Modify By Sindy 2025/5/13
      '寄發通知信
      strExc(0) = "select cum02 from CaseUseMemo" & _
                  " where cum05='04'" & strCon
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         Do While Not RsTemp.EOF
            strMailTo = strMailTo & ";" & RsTemp.Fields("cum02")
            RsTemp.MoveNext
         Loop
         strMailTo = Mid(strMailTo, 2)
      '2025/5/13 END
         '區分部門
         strF1xEmp = "": strF2xEmp = ""
         varTmp = Split(strMailTo, ";")
         For jj = 0 To UBound(varTmp)
            If Left(PUB_GetST03(CStr(varTmp(jj))), 2) = "F1" Then '外商
               strF1xEmp = strF1xEmp & ";" & varTmp(jj)
            Else
               strF2xEmp = strF2xEmp & ";" & varTmp(jj)
            End If
         Next jj
         'Call PUB_SendNotifyMail(strMailTo)
         If strF1xEmp <> "" Then
            strF1xEmp = Mid(strF1xEmp, 2)
            Call PUB_SendNotifyMail(strF1xEmp)
         End If
         If strF2xEmp <> "" Then
            strF2xEmp = Mid(strF2xEmp, 2)
            Call PUB_SendNotifyMail(strF2xEmp)
         End If
         '刪除記錄
         strExc(0) = "delete from CaseUseMemo" & _
                     " where cum05='04'" & strCon
         cnnConnection.Execute strExc(0)
      End If
   End If
   
   Set rsA = Nothing
End Sub

'Modify By Sindy 2022/5/24
'寄發通知信
'Modify By Sindy 2022/8/26 +, Optional bolIsSend As Boolean = False: 一定要發通知信=True
Public Sub PUB_SendNotifyMail(strEMailTo As String, Optional bolIsSend As Boolean = False)
Dim ArrStr As Variant
Dim jj As Integer, strContent As String
Dim bolCaseDutyAgentMsg As Boolean, strRestKind As String
Dim strRestEmp As String, strNormalEmp As String
Dim strTempCC As String
Dim strST04is2 As String 'Add By Sindy 2023/5/11
   
   If strEMailTo = "" Then Exit Sub
   
   'Modify By Sindy 2022/8/26
   '工作天才要寄信通知人員處理
   If (ChkWorkDay(strSrvDate(1)) = True And _
      (Format(time, "HHMMSS") >= "080000" And Format(time, "HHMMSS") < "183000")) _
      Or bolIsSend = True Then
   '2022/8/26 END
      strContent = "請至案件管理系統的一般作業\系統收件區，進行看查。" & vbCrLf & vbCrLf & vbCrLf & _
                   "若為職代或主管查詢其他同仁信件時，請在員工編號欄輸入欲查詢同仁之員工編號再按畫面更新按鈕。"
      If strEMailTo <> "" Then
         If Mid(strEMailTo, 1, 1) = ";" Then strEMailTo = Mid(strEMailTo, 2)
         '因為有休假問題,所以有休假人員各自發信,其他人則一封
         ArrStr = Split(strEMailTo, ";")
         For jj = 0 To UBound(ArrStr)
            strTempCC = GetCaseDutyAgent(ArrStr(jj), "", bolCaseDutyAgentMsg, strRestKind)
            If strTempCC <> "" Then
               strRestEmp = strRestEmp & ";" & ArrStr(jj)
            Else
               strNormalEmp = strNormalEmp & ";" & ArrStr(jj)
            End If
            'Add By Sindy 2023/5/11 檢查是否為離職人員
            If ChkStaffST04(CStr(ArrStr(jj)), False) = True Then '已離職
               strST04is2 = strST04is2 & ";" & ArrStr(jj)
            End If
            '2023/5/11 END
         Next jj
         'Add By Sindy 2023/5/11
         If strST04is2 <> "" Then
            If Mid(strST04is2, 1, 1) = ";" Then strST04is2 = Mid(strST04is2, 2)
            PUB_SendMail strUserNum, Pub_GetSpecMan("電腦中心郵件檢核人員"), "", strST04is2 & "此人員已離職，查看其系統收件區的信件狀況！", strContent, , , , , , , , , , True, False, , , False, , , False
         End If
         '2023/5/11 END
         
         '改為一起通知,減少操作人員等發信的時間
         '1.收受者請假不回彈訊息 2.修改信件內容
         If strNormalEmp <> "" Then
            If Mid(strNormalEmp, 1, 1) = ";" Then strNormalEmp = Mid(strNormalEmp, 2)
            PUB_SendMail strUserNum, strNormalEmp, "", "通知已有信件轉入系統收件區", strContent, , , , , , , , , , , False, , , False, , , False, , , , , , , "1"
         End If
      End If
      
      '有休假人員各自發信
      If strRestEmp <> "" Then
         If Mid(strRestEmp, 1, 1) = ";" Then strRestEmp = Mid(strRestEmp, 2)
         ArrStr = Split(strRestEmp, ";")
         For jj = 0 To UBound(ArrStr)
            PUB_SendMail strUserNum, ArrStr(jj), "", "通知已有信件轉入系統收件區", strContent, , , , , , , , , , , False, , , False, , , False, , , , , , , "1"
         Next jj
      End If
   End If
End Sub

'Add By Sindy 2018/5/29
'國際會議郵件歸往來記錄檔
'回傳:是否成功
'strMailType: 0 - Rx.外來郵件 1 - Tx.寄出郵件 "" - 代表執行匯入作業
'StrCR05 : 往來類別 Add By Sindy 2019/3/12
Public Function PUB_IPDeptISDMail(oForm As Form, ByVal strMailType As String, _
   ByVal strISDServerPath As String, ByVal strFilePath As String, ByVal strFileName As String, _
   Optional ByRef intCaseOK As Integer, Optional oListErrBox As ListBox) As Boolean
Dim fs, f
Dim bolConnect As Boolean
Dim rsA As New ADODB.Recordset
Dim strSubject As String, stReName As String
Dim ii As Integer, ArrStr As Variant
Dim strCR01 As String, strCR02 As String, strCR03 As String, StrCR04 As String, strCR06 As String
Dim strCR07 As String, strCR08 As String, strCR09 As String, strCR19 As String, strCF06 As String
Dim strProc As String, intStar As Integer, intEnd As Integer, strTextSubject As String
Dim strData(0 To 3) As String
Dim bolMoveFile As Boolean, bolErr As Boolean, bolErrText As String
Dim objOutLook As Object
Dim objMail As Object
Dim strMeetName As String
Dim strEmp As String, strDirector As String
Dim strChkST69 As String
Dim strMailDate As String, strMailTime As String
Dim strFullFileName As String
Dim strSenderName As String, strSenderBehalfofName As String, strMailKind As String
Dim input_type As String '[ISD.:1 ;其他:2
Dim strSenderNameST01 As String '寄件者,台一同仁 Add By Sindy 2018/10/30
'Add By Sindy 2019/2/22
Dim strPkey2_Type As String '往來類別
Dim strPkey4_No As String '聯絡人/臨時編號
Dim strII03_2 As String '回傳副檔名
Dim strCF02 As String 'Add By Sindy 2019/2/26
Dim strSize As String 'Add By Sindy 2019/2/26
Dim bolSaveFile As Boolean, bolPCC25 As Boolean 'Add By Sindy 2019/3/12
Dim StrCR05 As String
Dim strDataNo1 As String, strDataNo2 As String
Dim strCon As String, strText As String
Dim strCF10 As String, strCF14 As String 'Add By Sindy 2019/12/24 寄件者...
   
   PUB_IPDeptISDMail = False
   
On Error GoTo ErrHand
   
   If Right(Trim(strISDServerPath), 1) <> "\" Then strISDServerPath = Trim(strISDServerPath) & "\"
   If Right(Trim(strFilePath), 1) <> "\" Then strFilePath = Trim(strFilePath) & "\"
   strFullFileName = strFilePath & strFileName
   
   Set objOutLook = CreateObject("Outlook.Application")
   Set objMail = objOutLook.CreateItemFromTemplate(strFullFileName)
   Set fs = CreateObject("Scripting.FileSystemObject")
   Set f = fs.GetFile(strFullFileName)
   
   '內容放主旨
   strSubject = ChgSQL(objMail.Subject)
   oForm.Text1 = strSubject 'Re: ML/kc 中?特許出願201510920053.X　貴所整理番?31565－CN　弊所整理番?：P-112987
'   strSubject = ChgSQL(oForm.Text2) '要用文字框存放，因才能把unicode去掉
   DoEvents
   
   'Add By Sindy 2019/10/15
   If strMailType = "1" Then '寄出郵件
      If InStr(strSubject, "XXXXXXXXX") > 0 Then
         '刪除PC端檔案
         Call fs.DeleteFile(strFullFileName)
         DoEvents
         GoTo ExitFunc
      End If
   End If
   '2019/10/15 END
   
   '解析主旨使用
   strTextSubject = UCase(strSubject)
   'Modify By Sindy 2018/11/30 主旨設定錯誤,為分信順利這此Replace
   '2019 Chinese bamboo calendars to convey our best wishes from Tai E International Patent & Law Office [ISDXXXXXXXXX_ETC] (EY_wc)
   strTextSubject = Replace(strTextSubject, "_ETC] (EY_WC)", ".ETC] (EY/WC)")
   '2018/11/30 END
   strTextSubject = Replace(strTextSubject, "．", ".")
   strTextSubject = Replace(strTextSubject, "..", ".")
   strTextSubject = Replace(strTextSubject, "...", ".")
   strTextSubject = Replace(strTextSubject, "〔", "[")
   strTextSubject = Replace(strTextSubject, "〕", "]")
   strTextSubject = Replace(strTextSubject, " ", "") 'Add By Sindy 2019/2/22 空白去除後再比對
   strTextSubject = Replace(strTextSubject, "　", "") 'Add By Sindy 2019/2/22 空白去除後再比對
   strTextSubject = Replace(strTextSubject, "：", ":")
   '先檢查是否為ISD郵件
   'Modify By Sindy 2019/2/22 + [OURREF:...
   'Modify By Sindy 2021/4/7 + [OURREF:F
   If InStr(strTextSubject, "[ISD.") = 0 And _
      InStr(strTextSubject, "[ISDR") = 0 And _
      InStr(strTextSubject, "[ISDX") = 0 And _
      InStr(strTextSubject, "[ISDY") = 0 And _
      InStr(strTextSubject, "[OURREF:R") = 0 And _
      InStr(strTextSubject, "[OURREF:X") = 0 And _
      InStr(strTextSubject, "[OURREF:Y") = 0 And _
      InStr(strTextSubject, "[OURREF:A") = 0 And _
      InStr(strTextSubject, "[OURREF:B") = 0 And _
      InStr(strTextSubject, "[OURREF:F") = 0 Then
      If strMailType = "" Then
         If TypeName(oListErrBox) <> "Nothing" Then
            oListErrBox.AddItem strFileName & " : 主旨中找不到任何關鍵字", 0
         End If
      End If
      Set objMail = Nothing
      Set objOutLook = Nothing
      Exit Function
   End If
   bolMoveFile = False: bolErr = False
   For ii = 0 To 3 '2
      strData(ii) = ""
   Next ii
   '截取編號資料:
   'Add By Sindy 2019/2/22 + [OURREF:...
   If InStr(strTextSubject, "[OURREF:") > 0 Then
      intStar = InStr(UCase(strTextSubject), UCase("[OURREF:")) + Len(UCase("["))
      intEnd = InStr(intStar, UCase(strTextSubject), UCase("]"))
      input_type = "1"
      'Modify By Sindy 2021/4/7 + [OURREF:F
      If InStr(strTextSubject, "[OURREF:R") > 0 Or _
         InStr(strTextSubject, "[OURREF:X") > 0 Or _
         InStr(strTextSubject, "[OURREF:Y") > 0 Or _
         InStr(strTextSubject, "[OURREF:A") > 0 Or _
         InStr(strTextSubject, "[OURREF:B") > 0 Or _
         InStr(strTextSubject, "[OURREF:F") > 0 Then
         strProc = Trim(Mid(UCase(strTextSubject), intStar, intEnd - intStar))
         If InStr(strProc, ".") = 0 Then '檢查有沒有必需的分隔符號
            If strMailType = "" Then
               If TypeName(oListErrBox) <> "Nothing" Then
                  oListErrBox.AddItem strFileName & " : 關鍵字沒有必需的分隔符號", 0
               End If
            End If
            Set objMail = Nothing
            Set objOutLook = Nothing
            Exit Function
         Else
            strDataNo1 = Mid(strProc, Len("OURREF:Y"), 1)
            strDataNo2 = Mid(strProc, Len("OURREF:Y") + 1, InStr(strProc, ".") - (Len("OURREF:Y") + 1))
         End If
         If (strDataNo1 & strDataNo2) <> "" Then
            '檢查是不是往來類別
            strExc(0) = "select * from allcode" & _
                        " where ac01='11' and ac02='" & (strDataNo1 & strDataNo2) & "'"
            intI = 1
            Set rsA = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               If IsNumeric(Left(strDataNo2, 5)) = False Then '檢查RXY第2碼起是不是正常的數字
                  If strMailType = "" Then
                     If TypeName(oListErrBox) <> "Nothing" Then
                        oListErrBox.AddItem strFileName & " : 關鍵字RXY第2碼起不是正常的數字", 0
                     End If
                  End If
                  Set objMail = Nothing
                  Set objOutLook = Nothing
                  Exit Function
               End If
               '取得資料是不是客戶/代理人/潛在客戶編號
               'Add By Sindy 2021/4/7
               If Left(strDataNo1, 1) = "F" Then
                  strCR03 = "F" & Format(strDataNo2, "0000")
               Else
               '2021/4/7 END
                  strCR03 = Left((strDataNo1 & strDataNo2) & "0000000", 9)
               End If
               Select Case Left(strCR03, 1)
                  Case "R"
                     '國外潛在客戶
                     strExc(0) = "select pcu01,pcu02 from potcustomer" & _
                                 " where pcu01='" & Left(strCR03, 8) & "' and pcu02='0'" '" & Right(strCR03, 1) & "
                  Case "X"
                     '客戶檔
                     strExc(0) = "select cu01,cu02 from customer" & _
                                 " where cu01='" & Left(strCR03, 8) & "' and cu02='0'" '" & Right(strCR03, 1) & "
                  Case "Y"
                     '代理人檔
                     strExc(0) = "select fa01,fa02 from fagent" & _
                                 " where fa01='" & Left(strCR03, 8) & "' and fa02='0'" '" & Right(strCR03, 1) & "
                  'Add By Sindy 2021/4/7
                  Case "F"
                     '客戶平台
                     strExc(0) = "select cw01,'' from custweb" & _
                                 " where cw01='" & Mid(strCR03, 2) & "' and cw03='7'" '媒介平台
               End Select
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  If strMailType = "" Then
                     If TypeName(oListErrBox) <> "Nothing" Then
                        oListErrBox.AddItem strFileName & " : 關鍵字不是往來類別也不是客戶/代理人/潛在客戶編號/客戶媒介平台", 0
                     End If
                  End If
                  Set objMail = Nothing
                  Set objOutLook = Nothing
                  Exit Function
               'Add By Sindy 2019/7/4
               Else
                  'Add By Sindy 2021/4/7
                  If Left(strCR03, 1) <> "F" Then
                  '2021/4/7 END
                     strCR03 = Left(strCR03, 8) & "0"
                  End If
               '2019/7/4 END
               End If
            End If
         Else
            If strMailType = "" Then
               If TypeName(oListErrBox) <> "Nothing" Then
                  oListErrBox.AddItem strFileName & " : 主旨中找不到任何關鍵字", 0
               End If
            End If
            Set objMail = Nothing
            Set objOutLook = Nothing
            Exit Function
         End If
         input_type = "2"
      Else
         If strMailType = "" Then
            If TypeName(oListErrBox) <> "Nothing" Then
               oListErrBox.AddItem strFileName & " : 主旨中找不到任何關鍵字", 0
            End If
         End If
         Set objMail = Nothing
         Set objOutLook = Nothing
         Exit Function
      End If
   End If
   '2019/2/22 END
   '[ISD.2018 INTA.FY01]Nice to meet you at 2018 INTA Annual Conference in Seattle
   If InStr(strTextSubject, "[ISD.") > 0 Then
      intStar = InStr(UCase(strTextSubject), UCase("[ISD.")) + Len(UCase("["))
      intEnd = InStr(intStar, UCase(strTextSubject), UCase("]"))
      input_type = "1" '[ISD.
   End If
   '[ISDY20656.01.2018 INTA]Nice to meet you at 2018 INTA Annual Conference in Seattle
   If intStar = 0 And InStr(strTextSubject, "[ISDR") > 0 Then
      intStar = InStr(UCase(strTextSubject), UCase("[ISDR")) + Len(UCase("["))
      intEnd = InStr(intStar, UCase(strTextSubject), UCase("]"))
      input_type = "2" '其他
   End If
   If intStar = 0 And InStr(strTextSubject, "[ISDX") > 0 Then
      intStar = InStr(UCase(strTextSubject), UCase("[ISDX")) + Len(UCase("["))
      intEnd = InStr(intStar, UCase(strTextSubject), UCase("]"))
      input_type = "2" '其他
   End If
   If intStar = 0 And InStr(strTextSubject, "[ISDY") > 0 Then
      intStar = InStr(UCase(strTextSubject), UCase("[ISDY")) + Len(UCase("["))
      intEnd = InStr(intStar, UCase(strTextSubject), UCase("]"))
      input_type = "2" '其他
   End If
   
   If intStar > 0 And intEnd > 0 And intEnd > intStar Then
      strProc = Trim(Mid(UCase(strTextSubject), intStar, intEnd - intStar))
      ArrStr = Split(strProc, ".")
      For ii = 0 To UBound(ArrStr)
         'If ii > 2 Then Exit For
         If ii > 3 Then Exit For
         'Modify By Sindy 2018/7/20 + Trim
         strData(ii) = Trim(ArrStr(ii))
      Next ii

      '檢查資料正確性
      If InStr(strTextSubject, "[ISD") > 0 And input_type = "2" Then
'         If strData(0) = "" Or strData(1) = "" Or strData(2) = "" Then
'            bolErr = True
'            bolErrText = bolErrText & "規定文字輸入不足,無法辨識;"
'         End If
         If strData(2) = "" Then
            If strData(1) = "" Then
               bolErr = True
               bolErrText = bolErrText & "沒有輸入會議名稱;"
            Else
               strMeetName = strData(1)
            End If
         Else
            strMeetName = strData(2)
         End If
         
         '檢查客戶編號是否存在
         strCR03 = ChangeCustomerL(Mid(strData(0), 4))
         Select Case Left(strCR03, 1)
            Case "R"
               '國外潛在客戶
               strExc(0) = "select pcu01,pcu02 from potcustomer" & _
                           " where pcu01='" & Left(strCR03, 8) & "' and pcu02='0'" '" & Right(strCR03, 1) & "
            Case "X"
               '客戶檔
               strExc(0) = "select cu01,cu02 from customer" & _
                           " where cu01='" & Left(strCR03, 8) & "' and cu02='0'" '" & Right(strCR03, 1) & "
            Case "Y"
               '代理人檔
               strExc(0) = "select fa01,fa02 from fagent" & _
                           " where fa01='" & Left(strCR03, 8) & "' and fa02='0'" '" & Right(strCR03, 1) & "
         End Select
         intI = 1
         Set rsA = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            strCR03 = ""
            bolErr = True
            Select Case Left(strCR03, 1)
               Case "R"
                  bolErrText = bolErrText & "無此潛在客戶資料:" & strCR03 & ";"
               Case "X"
                  bolErrText = bolErrText & "無此客戶資料:" & strCR03 & ";"
               Case "Y"
                  bolErrText = bolErrText & "無此代理人資料:" & strCR03 & ";"
            End Select
         'Add By Sindy 2019/7/4
         Else
            strCR03 = Left(strCR03, 8) & "0"
         '2019/7/4 END
         End If
         '檢查聯絡人是否存在
         If strData(1) <> "" And strCR03 <> "" Then
            If Len(strData(1)) < 3 And IsNumeric(strData(1)) = True Then
               strExc(0) = "select pcc01,pcc02 from potcustcont" & _
                              " where pcc01='" & Left(strCR03, 8) & "' and pcc02='" & Right("00" & strData(1), 2) & "'"
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
                  If strData(2) = "" Then
                     'strCR03 = strCR03 '往來對象
                     StrCR04 = ""
                  Else
                     bolErr = True
                     bolErrText = bolErrText & "無此聯絡人資料:" & Left(strCR03, 8) & "-" & strData(1) & ";"
                  End If
               Else
                  'strCR03 = Left(strCR03, 8) & "0" '往來對象
                  StrCR04 = strData(1) '聯絡人
               End If
            Else
               If strData(2) = "" Then
                  'strCR03 = strCR03 '往來對象
                  StrCR04 = ""
               Else
                  bolErr = True
                  bolErrText = bolErrText & "無此聯絡人資料:" & Left(strCR03, 8) & "-" & strData(1) & ";"
               End If
            End If
         ElseIf strCR03 <> "" Then
            'strCR03 = strCR03 '往來對象
            StrCR04 = ""
         End If
         
      'Add By Sindy 2019/2/22 新規則 [OURREF:... 有客戶編號
      ElseIf InStr(strTextSubject, "[OURREF:") > 0 And input_type = "2" Then
         If strData(0) = "" Or strData(1) = "" Then
            bolErr = True
            bolErrText = bolErrText & "規定文字輸入不足,無法辨識;"
         Else
            'A01.1P 或 A01.2P
            If strData(1) <> "" And strData(2) <> "" Then
               strExc(0) = "select * from allcode" & _
                           " where ac01='11' and ac02='" & strData(1) & "." & strData(2) & "'"
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strData(1) = strData(1) & "." & strData(2)
                  strData(2) = ""
               End If
            End If
            strPkey2_Type = strData(1) '往來類別
            If strData(2) <> "" Then strMeetName = strData(2) '會議名稱
            If strData(3) <> "" Then strPkey4_No = strData(3) '聯絡人/臨時編號
            If strPkey4_No = "" And strMeetName <> "" And Len(strMeetName) <= 2 And IsNumeric(strMeetName) = True Then
               strPkey4_No = Format(strMeetName, "00")
               If strPkey4_No = "00" Then strPkey4_No = ""
               strMeetName = ""
            End If
            'Add By Sindy 2021/4/7
            If Left(strCR03, 1) <> "F" Then
            '2021/4/7 END
               strCR03 = ChangeCustomerL(Mid(strData(0), 8)) '例如OURREF:R
            End If
            '檢查客戶編號是否存在
            Select Case Left(strCR03, 1)
               Case "R"
                  '國外潛在客戶
                  strExc(0) = "select pcu01,pcu02 from potcustomer" & _
                              " where pcu01='" & Left(strCR03, 8) & "' and pcu02='0'" '" & Right(strCR03, 1) & "
               Case "X"
                  '客戶檔
                  strExc(0) = "select cu01,cu02 from customer" & _
                              " where cu01='" & Left(strCR03, 8) & "' and cu02='0'" '" & Right(strCR03, 1) & "
               Case "Y"
                  '代理人檔
                  strExc(0) = "select fa01,fa02 from fagent" & _
                              " where fa01='" & Left(strCR03, 8) & "' and fa02='0'" '" & Right(strCR03, 1) & "
               'Add By Sindy 2021/4/7
               Case "F"
                  '客戶平台
                  strExc(0) = "select cw01,'' from CustWeb" & _
                              " where cw01='" & Mid(strCR03, 2) & "' and cw03='7'" '媒介平台
            End Select
            intI = 1
            Set rsA = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               bolErr = True
               Select Case Left(strCR03, 1)
                  Case "R"
                     bolErrText = bolErrText & "無此潛在客戶資料:" & strCR03 & ";"
                  Case "X"
                     bolErrText = bolErrText & "無此客戶資料:" & strCR03 & ";"
                  Case "Y"
                     bolErrText = bolErrText & "無此代理人資料:" & strCR03 & ";"
                  'Add By Sindy 2021/4/7
                  Case "F"
                     bolErrText = bolErrText & "無此客戶媒介平台資料:" & strCR03 & ";"
               End Select
               strCR03 = ""
            'Add By Sindy 2019/7/4
            Else
               'Add By Sindy 2021/4/7
               If Left(strCR03, 1) <> "F" Then
               '2021/4/7 END
                  strCR03 = Left(strCR03, 8) & "0"
               End If
            '2019/7/4 END
            End If
            '檢查聯絡人是否存在
            If strPkey4_No <> "" And strCR03 <> "" Then
               strExc(0) = "select pcc01,pcc02 from potcustcont" & _
                              " where pcc01='" & Left(strCR03, 8) & "' and pcc02='" & strPkey4_No & "'"
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 0 Then
'                  If strMeetName = "" Then
'                     StrCR04 = ""
'                  Else
                     bolErr = True
                     bolErrText = bolErrText & "無此聯絡人資料:" & Left(strCR03, 8) & "-" & strPkey4_No & ";"
'                  End If
               Else
                  StrCR04 = strPkey4_No '聯絡人
               End If
            ElseIf strCR03 <> "" Then
               StrCR04 = ""
            End If
         End If
      '2019/2/22 END
            
      Else '無客戶編號
         If strData(0) = "" Or strData(1) = "" Or strData(2) = "" Then
            bolErr = True
            bolErrText = bolErrText & "規定文字輸入不足,無法辨識;"
         End If
         'A01.1P 或 A01.2P
         If strData(0) <> "" And strData(1) <> "" Then
            strExc(0) = "select * from allcode" & _
                        " where ac01='11' and ac02='" & strData(0) & "." & strData(1) & "'"
            intI = 1
            Set rsA = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strData(0) = strData(0) & "." & strData(1)
               strData(1) = ""
            End If
         End If
         strPkey2_Type = strData(0) '往來類別
         strMeetName = strData(1)
         '往來對象
         If strData(2) <> "" Then
            strExc(0) = "select pcc01,pcc02 from potcustcont" & _
                           " where Replace(upper(pcc25),' ','')='" & Replace(Trim(strMeetName), " ", "") & "." & UCase(strData(2)) & "'"
            intI = 1
            Set rsA = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               bolPCC25 = True
               strCR03 = rsA.Fields("pcc01") & "0"
               StrCR04 = rsA.Fields("pcc02")
            Else
               bolErr = True
               bolErrText = bolErrText & "無往來對象;"
            End If
            For ii = 1 To Len(strData(2))
               If IsNumeric(Mid(strData(2), ii, 1)) = False Then
                  strChkST69 = strChkST69 & Mid(strData(2), ii, 1)
               'Add By Sindy 2019/1/7
               Else
                  Exit For
               '2019/1/7 END
               End If
            Next ii
            '接洽同仁
            If strChkST69 <> "" Then
               strExc(0) = "select st01,st02,st04,st69 from staff" & _
                           " where st69 is not null and st04='1'" & _
                           " and instr(upper('/" & Trim(strChkST69) & "'),upper(substr(st69,instr(st69,'/'))))>0" & _
                           "order by st04 asc,st01 desc"
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strCR19 = rsA.Fields("st01")
               End If
            End If
         End If
      End If
      
      'Add By Sindy 2019/2/22
      '檢查往來類別是否存在
      If strPkey2_Type <> "" Then
         strExc(0) = "select * from allcode" & _
                     " where ac01='11' and ac02='" & strPkey2_Type & "'"
         intI = 1
         Set rsA = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 0 Then
            bolErr = True
            bolErrText = bolErrText & "無往來類別;"
         Else
            StrCR05 = rsA.Fields("ac02") '往來類別代碼
         End If
      End If
      '2019/2/22 END
      
      If objMail.Class <> 43 Then
         bolErr = True
         bolErrText = bolErrText & "非一般郵件;"
         'Add By Sindy 2019/10/15
         If InStr(strTextSubject, "未傳遞的主旨") > 0 Then
            '刪除PC端檔案
            Call fs.DeleteFile(strFullFileName)
            DoEvents
            GoTo ExitFunc
         End If
         '2019/10/15 END
      Else
         strMailDate = Format(objMail.SentOn, "YYYYMMDD") '信件日期
         strMailTime = Format(objMail.SentOn, "HHMMSS") '信件時間
         '讀取寄件者
         If InStr(objMail.SenderName, "(") = 0 Then
            strSenderName = ChgSQL(objMail.SenderName)
         Else
            strSenderName = ChgSQL(Trim(Mid(objMail.SenderName, 1, InStr(objMail.SenderName, "(") - 1)))
         End If
         strSenderBehalfofName = objMail.sentonbehalfofname
         
         'Modify By Sindy 2025/2/17
         'Add By Sindy 2019/12/24 寄件者
'         strCF10 = PUB_GetMail_ii11(objMail) 'Modify By Sindy 2024/7/30
         Call PUB_ReadMailText(objMail, , , , strCF10)
         '2025/2/17 END
'         If objMail.SenderEmailType = "EX" Then
'            strCF10 = objMail.SenderName
'         Else
'            If objMail.SenderName = objMail.senderemailaddress Then
'               strCF10 = objMail.senderemailaddress
'            Else
'               'Add By Sindy 2024/7/29
'               If InStr(UCase(objMail.senderemailaddress), UCase("Recipients/cn=")) > 0 Then
'                  strCF10 = objMail.SenderName
'               Else
'               '2024/7/29 END
'                  strCF10 = objMail.SenderName & " [" & objMail.senderemailaddress & "]"
'               End If
'            End If
'         End If
         '依主旨解析專業代號寄件人員是誰
         Call BySubjectToStaff(strSubject, strSenderName, strCF14, strDirector, False)
         '2019/12/24 END
         
         If strMailType = "" Then
            If UCase(strSenderBehalfofName) = UCase("ipdept") Then
               strMailKind = "T" '寄出郵件
            End If
'            strExc(0) = "select st01,st02,st04,st69 from staff" & _
'                        " where st02='" & strSenderName & "' and st04='1'"
'            intI = 1
'            Set rsA = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               strSenderNameST01 = rsA.Fields("st01") 'Add By Sindy 2018/10/30
'               strMailKind = "T" '寄出郵件
'            End If
         ElseIf strMailType = "0" Then '0 - Rx.外來郵件
            strMailKind = "R" '外來郵件
         ElseIf strMailType = "1" Then '1 - Tx.寄出郵件
            strMailKind = "T" '寄出郵件
         End If
         'Modify By Sindy 2021/1/6 調整SQL
         strExc(0) = "select st01,st02,st04,st69 from staff" & _
                     " where st02='" & strSenderName & "' AND st04='1' AND substr(st01,1,1)<'F' AND st05 IS NOT NULL" & _
                     " and substr(st01,4,1)<>'9'"
         intI = 1
         Set rsA = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strSenderNameST01 = rsA.Fields("st01") 'Add By Sindy 2018/10/30
            strCF10 = rsA.Fields("st01") '若為員工,存入員工編號
            If strMailKind = "" Then strMailKind = "T" '寄出郵件
         Else
            If strMailKind = "" Then strMailKind = "R" '外來郵件
         End If
      End If
   Else
      bolErr = True
      bolErrText = bolErrText & "解析關鍵字有誤(" & strTextSubject & ");"
   End If
   
   If bolErr = True Then
      If strMailType = "" Then
         If TypeName(oListErrBox) <> "Nothing" Then
            oListErrBox.AddItem strFileName & " : " & bolErrText, 0
         End If
         Set objMail = Nothing
         Set objOutLook = Nothing
         Exit Function '有錯誤,無法匯入
      Else
         bolMoveFile = True '有錯,郵件搬到資料夾,供人員再查核
      End If
   End If
   
   Screen.MousePointer = vbHourglass
   'Set fs = CreateObject("Scripting.FileSystemObject")
   'Set f = fs.GetFile(strFullFileName)
   
   'Modify By Sindy 2019/2/22 解析主旨抓出對應的案件性質,副檔名; 回傳副檔名
   'Modify By Sindy 2025/2/26 0=外來郵件不需要解析副檔名
   If strMailType <> "0" Then
   '2025/2/26 END
      Call PUB_IPDept_ComparisonCP(strTextSubject, "", "", "", "", "", strII03_2, "", "")
   End If
   stReName = strMailDate & strMailTime & "[" & Replace(strProc, ":", " ") & "]" & IIf(strII03_2 <> "", "." & strII03_2 & ".msg", "")
   '2019/2/22 END
   '歸往來記錄檔
   If bolMoveFile = False Or strMailType = "" Then
      If strMailType = "0" Or UCase(InStr(strFileName, -7)) = UCase(".Rx.msg") Or strMailKind = "R" Then  '外來郵件
         'Add By Sindy 2019/2/22
         If strII03_2 = "" Then
         '2019/2/22 END
            stReName = stReName & ".Rx.msg"
         End If
         strCR07 = "外來郵件" '場合
      ElseIf strMailType = "1" Or UCase(InStr(strFileName, -7)) = UCase(".Tx.msg") Or strMailKind = "T" Then '寄出郵件
         'Add By Sindy 2019/2/22
         If strII03_2 = "" Then
         '2019/2/22 END
            stReName = stReName & ".Tx.msg"
         End If
         strCR07 = "寄出郵件"
      Else
         'Add By Sindy 2019/2/22
         If strII03_2 = "" Then
         '2019/2/22 END
            stReName = stReName & ".msg"
         End If
         strCR07 = "Mail"
      End If
      '接洽同仁
      If strCR19 = "" Then
         strExc(0) = "select st01,st02,st04,st69 from staff" & _
                     " where st69 is not null and st04='1'" & _
                     " and instr(upper('" & strSubject & "'),upper(st69))>0" & _
                     " order by length(st69) desc,st04 asc,st01 desc"
         intI = 1
         Set rsA = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strCR19 = rsA.Fields("st01")
         Else
            If strCF14 <> "" Then strCR19 = strCF14
         End If
      End If
      
      strCR02 = strMailDate '往來日期=信件日期
      'Add By Sindy 2018/10/30
      If strCR19 = "" And strSenderNameST01 <> "" Then strCR19 = strSenderNameST01
      '往來類別
      'Add By Sindy 2019/2/22
      If StrCR05 = "" Then
      '2019/2/22 END
         If UCase(strMeetName) = "QUO" Then
            'StrCR05 = "來函詢價/申請文件"
            StrCR05 = "A02"
         ElseIf UCase(strMeetName) = "REC" Then
            'StrCR05 = "互惠"
            StrCR05 = "B29"
         ElseIf UCase(strMeetName) = "INT" Then
            'StrCR05 = "訪談"
            StrCR05 = "B19"
         ElseIf UCase(strMeetName) = "CON" Then
            'StrCR05 = "慰問"
            StrCR05 = "B49"
         ElseIf UCase(strMeetName) = "ETC" Then
            StrCR05 = "其他"
            StrCR05 = "B99"
         Else
            'StrCR05 = "國際會議往來信函"
            StrCR05 = "B11"
            strCR06 = strMeetName '主旨=會議名稱
         End If
         '2018/10/30 END
      'Add By Sindy 2019/2/22
      'Modify By Sindy 2019/11/19 + Or strPkey2_Type = "B61"
      ElseIf strPkey2_Type = "B11" Or strPkey2_Type = "B61" Then
         strCR06 = strMeetName '主旨=會議名稱
      End If
      '2019/2/22 END
      
      If strCR06 = "" Then strCR06 = strSubject '主旨 Add By Sindy 2018/10/30
      strCR08 = strSubject '內容
      'strCR09 = strCR01 & "_" & stReName & " (" & -1 * Int(-1 * f.Size / 1024) & " KB)" '附件檔名
      'Add By Sindy 2019/2/26
      strSize = (-1 * Int(-1 * f.Size / 1024))
      '2019/2/26 END
      
      'Add By Sindy 2019/12/26 檢查寄出郵件是否為整批寄出,若是,寄件者改為QPGMR
      strExc(0) = "select ms01,ms11,ms12 from mailschedule" & _
                  " where upper(ms03)=upper('ipdept@taie.com.tw') AND ms25='Y'" & _
                  " and " & strMailDate & Format(strMailTime, "000000") & ">=ms16||substr(to_char('00'||ms17),-6) and ms16>=" & strMailDate & _
                  " and instr(upper('" & strSubject & "'),upper(ms02))>0"
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         'Modify By Sindy 2020/4/22
'         If Val("" & rsA.Fields("ms11")) > 0 Then
'            If Val(strMailDate & Format(strMailTime, "000000")) <= Val(rsA.Fields("ms11") & Format(rsA.Fields("ms12"), "000000")) Then
'               strCF10 = "QPGMR" '為整批寄出
'            End If
'         Else
'            strCF10 = "QPGMR" '為整批寄出
'         End If
         If strMailKind = "T" Then '寄出郵件
            strCF10 = "QPGMR" '為整批寄出
         End If
         '2020/4/22 END
      End If
      If UCase(strCF10) = "IPDEPT" And strCF14 <> "" Then
         strCF10 = strCF14
      End If
      '2019/12/26 END
      
'      strSql = "insert into ContactRecord(CR01,CR02,CR03,CR04,CR05" & _
'               ",CR06,CR07,CR08,CR09,CR19)" & _
'               " values('" & strCR01 & "'," & strCR02 & "," & CNULL(strCR03) & "," & CNULL(StrCR04) & ",'" & StrCR05 & "'" & _
'               "," & CNULL(strCR06) & ",'" & strCR07 & "','" & ChgSQL(strCR08) & "','" & strCR09 & "'," & CNULL(strCR19) & ")"
      '1.寄出郵件,整批寄信的才要新增往來記錄
      '臨時編號的也要新增往來記錄
      If (strMailType = "1" And UCase(strSenderName) = UCase("ipdept") And _
         Not (StrCR05 <> "" And Len(StrCR05) = 3 And Left(StrCR05, 1) = "A") _
         ) _
         Or bolPCC25 = True Then
RunInsert: '開拓後面程式段沒有找到可以加入電子檔的往來記錄時,就會回至此處新增往來記錄
         cnnConnection.BeginTrans: bolConnect = True
         strCR01 = AutoNo("K", 6) '往來記錄編號
         bolSaveFile = True
         strSql = "insert into ContactRecord(CR01,CR02,CR03,CR04,CR05" & _
                  ",CR06,CR07,CR08,CR19)" & _
                  " values('" & strCR01 & "'," & strCR02 & "," & CNULL(IIf(Left(strCR03, 1) = "F", Mid(strCR03, 2), strCR03)) & "," & CNULL(StrCR04) & ",'" & StrCR05 & "'" & _
                  "," & CNULL(strCR06) & ",'" & strCR07 & "','" & ChgSQL(strCR08) & "'," & CNULL(strCR19) & ")"
         cnnConnection.Execute strSql
      Else
         '若有同往來對象者,更新附件至最近那一筆往來記錄
         If strCR03 <> "" And StrCR05 <> "" Then
            'Modify By Sindy 2021/8/2 Mark:不用理會聯絡人資訊在往來記錄有無的問題, 也不用列入為歸檔的檢查條件
'            If StrCR04 <> "" Then '有聯絡人
'               strCon = strCon & " and cr04='" & StrCR04 & "'"
'               strText = strText & ";聯絡人:" & StrCR04
'            Else
'               strCon = strCon & " and cr04 is null"
'            End If

            'Modify By Sindy 2019/3/27 開拓的會議信函要比對到會議名稱
            'Modify By Sindy 2019/11/19 + Or StrCR05 = "B61"
            If StrCR05 = "國際會議往來信函" Or StrCR05 = "B11" Or StrCR05 = "B61" Then
               strCon = strCon & " and CR06='" & strCR06 & "'"
               strText = strText & ";會議名稱:" & strCR06
            End If
            '2019/3/27 END
            '搜尋資料
            'Left(strCR05, 1) = "B"=開拓單位使用
            'Modify By Sindy 2025/2/26 增加,由業拓同仁所發出的A02往來類別信函，能自動回存。
            If Left(StrCR05, 1) = "B" Or _
               (PUB_GetST93(strCR19) = "B01" And StrCR05 = "A02") Then
               If PUB_GetST93(strCR19) = "B01" And StrCR05 = "A02" Then
                  strCon = strCon & " and st93='B01'"
               End If
               '2025/2/26 END
               '先比較往來類別
               strExc(0) = "select cr01,cr02,cr06 from ContactRecord,staff" & _
                           " where cr03='" & IIf(Left(strCR03, 1) = "F", Mid(strCR03, 2), strCR03) & "'" & _
                           " and cr05='" & StrCR05 & "' and cr19=st01(+)" & strCon & _
                           " order by cr02 desc,cr01 desc"
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  'Modify By Sindy 2023/2/2 B開拓增加新增往來記錄的判斷
                  '比較主旨:
                  '  1. 異：新一筆
                  '  2. 同：再比較日期
                  '    (1)半年以上：新一筆
                  '    (2)半年以內：同一筆
                  '加主旨檢查:(排除字樣 FW: RE:)
                  strCR06 = Trim(Replace(strCR06, "FW:", ""))
                  strCR06 = Trim(Replace(strCR06, "F:", ""))
                  strCR06 = Trim(Replace(strCR06, "RE:", ""))
                  strCR06 = Trim(Replace(strCR06, "R:", ""))
                  strExc(0) = "select cr01,cr02,cr06 from ContactRecord,staff" & _
                              " where cr03='" & IIf(Left(strCR03, 1) = "F", Mid(strCR03, 2), strCR03) & "'" & _
                              " and cr05='" & StrCR05 & "' and upper(cr06)='" & ChgSQL(UCase(strCR06)) & "'" & _
                              " and cr19=st01(+)" & strCon & _
                              " order by cr02 desc,cr01 desc"
                  intI = 1
                  Set rsA = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 0 Then '無,則直接新增
                     GoTo RunInsert
                  Else
                     If rsA.Fields("cr02") > DBDATE(DateAdd("m", -6, Format(strSrvDate(1), "####/##/##"))) Then
                        strCR01 = rsA.Fields("cr01")
                        bolSaveFile = True
                     Else
                        GoTo RunInsert
                     End If
                  End If
                  '2023/2/2 END
               Else
                  'Add By Sindy 2019/3/28 國際會議找不到資料儲存附件,就是新增
                  'Modify By Sindy 2019/11/19 + Or StrCR05 = "B61"
                  If StrCR05 = "國際會議往來信函" Or StrCR05 = "B11" Or StrCR05 = "B61" Then
                     GoTo RunInsert
                  '2019/3/28 END
                  'Add By Sindy 2019/4/29 開拓主旨無輸入聯絡人編號,就歸此代理人且同往來類別最近一道
                  ElseIf StrCR04 = "" Then 'Left(strCR05, 1) = "B" And
                     strExc(0) = "select cr01 from ContactRecord,staff" & _
                                 " where cr03='" & IIf(Left(strCR03, 1) = "F", Mid(strCR03, 2), strCR03) & "'" & _
                                 " and cr05='" & StrCR05 & "' and cr19=st01(+)" & _
                                 " order by cr02 desc,cr01 desc"
                     intI = 1
                     Set rsA = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        strCR01 = rsA.Fields("cr01")
                        bolSaveFile = True
                     Else
'                        'Modify By Sindy 2019/5/30 B類項目,查此編號無該類別的往來記錄時,就直接新增該筆往來記錄
'                        If Left(strCR05, 1) = "B" Then
                           GoTo RunInsert
'                        Else
'                        '2019/5/30 END
'                           bolErr = True
'                           bolErrText = bolErrText & "無解析到相對的往來記錄(對象:" & strCR03 & ";往來類別:" & strCR05 & strText & ")"
'                        End If
                     End If
                  '2019/4/29 END
                  Else
'                     'Modify By Sindy 2019/5/30 B類項目,查此編號無該類別的往來記錄時,就直接新增該筆往來記錄
'                     If Left(strCR05, 1) = "B" Then
                        GoTo RunInsert
'                     Else
'                     '2019/5/30 END
'                        bolErr = True
'                        bolErrText = bolErrText & "無解析到相對的往來記錄(對象:" & strCR03 & ";往來類別:" & strCR05 & strText & ")"
'                     End If
                  End If
               End If
               
            Else '承辦組
               strExc(0) = "select cr01,cr06 from ContactRecord" & _
                           " where cr03='" & IIf(Left(strCR03, 1) = "F", Mid(strCR03, 2), strCR03) & "'" & _
                           " and cr05='" & StrCR05 & "'" & strCon & _
                           " order by cr02 desc,cr01 desc"
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strCR01 = rsA.Fields("cr01")
                  bolSaveFile = True
               Else
                  bolErr = True
                  bolErrText = bolErrText & "無解析到相對的往來記錄(對象:" & strCR03 & ";往來類別:" & StrCR05 & strText & ")"
               End If
            End If
         Else
            bolErr = True
            bolErrText = bolErrText & "無解析到往來對象(" & strCR03 & ")或往來類別(" & StrCR05 & ");"
         End If
      End If
      'Modify By Sindy 2019/3/12
      If bolSaveFile = True And strCR01 <> "" Then
      '2019/3/12 END
         If bolConnect = False Then cnnConnection.BeginTrans: bolConnect = True
         strCF02 = strCR01 & "_" & stReName
         'Modify By Sindy 2019/2/25 "CONTACTRECORD" 改為 "CONTACTFILE"
         If PUB_PutFtpFile(strFullFileName, strCR01, strCR01 & "_" & stReName, strCF06, "CONTACTFILE") = False Then
            bolErr = True
            bolErrText = bolErrText & "存附件失敗" & stReName & ":" & Err.Number & "-" & Err.Description & ";"
            PUB_SendMail strUserNum, Pub_GetSpecMan("電腦中心郵件檢核人員"), "", "存附件失敗:" & stReName, bolErrText, , , , , , , , , , , False, , , False, , , False
            DoEvents
            cnnConnection.RollbackTrans: bolConnect = False
            If strMailType = "" Then
               If TypeName(oListErrBox) <> "Nothing" Then
                  oListErrBox.AddItem strFileName & " : " & bolErrText, 0
               End If
               GoTo ExitFunc '有錯誤,無法匯入
            Else
               '郵件搬到資料夾中
               bolMoveFile = True
               GoTo ErrHand2
            End If
         Else
            'FTP路徑
   '         strSql = "update ContactRecord set CR20='" & strCR20 & "'" & _
   '                  " where CR01='" & strCR01 & "'"
   '         cnnConnection.Execute strSql
            'Add By Sindy 2019/2/26
            'insert into Contactfile(CF01,CF02,CF06,CF07) values('KA8000335','KA8000335_20190314163330[ISDY45803010.ETC].Tx.msg','KA80/KA8000335/KA8000335_20190314163330ISDY45803010.ETC.Tx.msg.001','52')
            'Modify By Sindy 2019/12/24 + ,CF09,CF10,CF11,CF12,CF13,CF14
            'Modify By Sindy 2025/5/21 原CNULL(strCF10)+ChgSQL
            strSql = "insert into Contactfile(CF01,CF02,CF06,CF07,CF09,CF10,CF11,CF12,CF13,CF14)" & _
                     " values(" & CNULL(strCR01) & "," & CNULL(strCF02) & "," & CNULL(strCF06) & "," & CNULL(strSize) & _
                     "," & CNULL(strMailKind) & "," & CNULL(ChgSQL(strCF10)) & "," & strMailDate & "," & strMailTime & "," & CNULL(ChgSQL(strSubject)) & _
                     "," & CNULL(strCF14) & ")"
            cnnConnection.Execute strSql
            '2019/2/26 END
            
            intCaseOK = intCaseOK + 1 '記錄個案筆數
            '正常,記錄Log
            If strMailType <> "" Then
               oForm.WLog_Day strSubject & vbCrLf & _
                        "==>收到日期:" & strMailDate & " " & strMailTime & _
                        "==>電子檔名:" & strCF06 & vbCrLf, UCase("ISD")
            End If
         End If
         If strMailType <> "0" Then '外來郵件一樣後續要分信出去,因此不用刪除電子檔
            Call fs.DeleteFile(strFullFileName)
            DoEvents
         End If
         cnnConnection.CommitTrans
         pub_SaveCoRec = True 'Add By Sindy 2022/6/17 記錄有儲存往來記錄
         bolConnect = False
         GoTo ExitFunc
      End If
   End If
   
ErrHand2:
   'If bolMoveFile = True Then '郵件搬到資料夾中
   'Modify By Sindy 2019/3/12 往來類別是B類的才要存到資料夾中
   If strMailType <> "" Then '匯入區不需處理電子檔問題
      If Left(StrCR05, 1) <> "A" Then
      '2019/3/12 END
         oForm.Text1.Text = strSubject
         oForm.Text1.Text = Replace(oForm.Text1.Text, "/", "")
         oForm.Text1.Text = Replace(oForm.Text1.Text, "\", "")
         oForm.Text1.Text = Replace(oForm.Text1.Text, ":", "")
         oForm.Text1.Text = Replace(oForm.Text1.Text, "*", "")
         oForm.Text1.Text = Replace(oForm.Text1.Text, "?", "")
         oForm.Text1.Text = Replace(oForm.Text1.Text, "<", "")
         oForm.Text1.Text = Replace(oForm.Text1.Text, ">", "")
         oForm.Text1.Text = Replace(oForm.Text1.Text, "|", "")
         '1.郵件搬到資料夾中
         If strMailType = "0" Or UCase(InStr(strFileName, -7)) = UCase(".Rx.msg") Then '外來郵件
            'stReName = stReName & Left(Trim(oForm.Text1.Text), 50) & ".Rx.msg"
            stReName = Trim(Left(oForm.Text1.Text, 60)) & "_" & strFileName '& ".Rx.msg"
            stReName = Replace(stReName, ".msg", ".Rx.msg")
         Else 'If strMailType = "1" Or UCase(InStr(strFileName, -7)) = UCase(".Tx.msg") Then '寄出郵件
            '寄出郵件
            'stReName = stReName & Left(Trim(oForm.Text1.Text), 50) & ".Tx.msg"
            stReName = Trim(Left(oForm.Text1.Text, 60)) & "_" & strFileName '& ".Tx.msg"
            stReName = Replace(stReName, ".msg", ".Tx.msg")
         End If
         
         fs.CopyFile strFullFileName, strISDServerPath & stReName
         DoEvents
         
         '2.檢查檔案複製完成，刪除檔案
         If Dir(strISDServerPath & stReName) <> "" Then
            If bolErr = False Then
               '正常,記錄Log
               oForm.WLog_Day strSubject & vbCrLf & _
                        "==>收到日期:" & strMailDate & " " & strMailTime & _
                        "==>電子檔名:" & strISDServerPath & stReName & vbCrLf, UCase("ISD")
            End If
            If strMailType <> "0" Then  '外來郵件一樣後續要分信出去,因此不用刪除電子檔
               '3.刪除PC端檔案
               Call fs.DeleteFile(strFullFileName)
               DoEvents
            End If
         Else
            bolErr = True
            bolErrText = bolErrText & "儲存電子檔失敗" & stReName & ":" & Err.Number & "-" & Err.Description & ";"
            PUB_SendMail strUserNum, Pub_GetSpecMan("電腦中心郵件檢核人員"), "", "儲存電子檔失敗:" & stReName, bolErrText, , strFullFileName, , , , , , , , , False, , , False, , , False
            DoEvents
         End If
         strEmp = "A4024" '陳增廣
         If strCR19 <> "" Then
            If InStr(strCR19, strEmp) = 0 Then
               strEmp = strCR19 & ";" & strEmp
            End If
         'Modify By Sindy 2025/2/11
         ElseIf strCF14 <> "" Then
            If InStr(strCF14, strEmp) = 0 Then
               strEmp = strCF14 & ";" & strEmp
            End If
         '2025/2/11 END
         End If
         If strDirector <> "" Then strDirector = strDirector & ";"
         strDirector = strDirector & PUB_GetFCPProSup("A4024")
'         'Modify By Sindy 2019/11/5
'         'strDirector = "99098" '楊雯芳
'         strDirector = PUB_GetFCPProSup(strEmp)
'         '2019/11/5 END
      Else
         If strMailType <> "0" Then '外來郵件一樣後續要分信出去,因此不用刪除電子檔
            Call fs.DeleteFile(strFullFileName)
            DoEvents
         End If
         'Modify By Sindy 2019/12/25
         strEmp = strCF14
'         strEmp = "": strDirector = ""
'         '先抓寄出人員
'         'Call BySenderToStaff(strSenderName, strEmp, strDirector)
'         'Modify By Sindy 2019/9/3 依主旨解析寄件人員是誰
'         Call BySubjectToStaff(strSubject, strSenderName, strEmp, strDirector)
         '2019/12/25 END
         '無人暫時就歸開拓
         If strEmp = "" Then
'            strEmp = "A4024" '陳增廣
'            If strCR19 <> "" Then strEmp = strCR19 & ";" & strEmp
'            'Modify By Sindy 2019/11/5
'            'strDirector = "99098" '楊雯芳
'            strDirector = PUB_GetFCPProSup(strEmp)
'            '2019/11/5 END
            strEmp = "QPGMR" 'Modify By Sindy 2025/10/14 改使用QPGMR
            If strCR19 <> "" Then
               If InStr(strCR19, strEmp) = 0 Then
                  strEmp = strCR19 & ";" & strEmp
                  If strDirector <> "" Then strDirector = strDirector & ";"
                  'strDirector = strDirector & PUB_GetFCPProSup("77015") 'Modify By Sindy 2025/10/14 mark
               End If
            Else
               'strDirector = PUB_GetFCPProSup(strEmp) 'Modify By Sindy 2025/10/14 mark
            End If
         End If
      End If
      '4.有錯誤記錄Log
      If bolErr = True Then
         oForm.WLog_Day bolErrText & vbCrLf & strSubject & vbCrLf & _
                     "==>收到日期:" & strMailDate & " " & strMailTime & vbCrLf, UCase("ISD_ERR")
         '記錄通知User_Log
         If strEmp <> "" Then
            strSql = "insert into R100101(R005002,R005004,R005005,R005003,R005007,R005006,R005008,ID)" & _
                     " values('" & strMailDate & "','" & strMailTime & "','系統Log記錄,不可刪除','" & strSenderName & "','" & ChgSQL(strSubject & vbCrLf & bolErrText) & "'," & _
                     "'" & strEmp & "','" & strDirector & "','" & strUserNum & "')"
            cnnConnection.Execute strSql
         End If
      End If
   Else
      If bolErr = True Then
         If TypeName(oListErrBox) <> "Nothing" Then
            oListErrBox.AddItem strFileName & " : " & bolErrText, 0
         End If
         Set objMail = Nothing
         Set objOutLook = Nothing
         Exit Function '有錯誤,無法匯入
      End If
   End If
   
ExitFunc:
   
   'Modify By Sindy 2019/9/3 取消
'   strSql = "insert into R100101(R005002,R005004,R005005,R005003,R005007,R005006,R005008,ID)" & _
'            " values('" & strMailDate & "','" & strMailTime & "','系統Log記錄,不可刪除','" & strSenderName & "','" & ChgSQL(strSubject & IIf(bolErrText <> "", vbCrLf & bolErrText, "") & IIf(strCR01 <> "", vbCrLf & strCR01, "")) & "'," & _
'            "'97038',null,'" & strUserNum & "')"
'   cnnConnection.Execute strSql
   
   PUB_IPDeptISDMail = True
   Screen.MousePointer = vbDefault
   Set objMail = Nothing
   Set objOutLook = Nothing
   Set f = Nothing
   Set fs = Nothing
   Set rsA = Nothing

   Exit Function

ErrHand:
   If bolConnect = True Then cnnConnection.RollbackTrans: bolConnect = False
   'Add By Sindy 2025/5/21
   oForm.WLog_Day bolErrText & " Err.Number:" & Err.Number & " Err.Description:" & Err.Description & vbCrLf & _
                  strSubject & vbCrLf & _
                  "==>收到日期:" & strMailDate & " " & strMailTime & vbCrLf & _
                  "strSql= " & strSql, UCase("ISD_ERR")
   '2025/5/21 END
   Screen.MousePointer = vbDefault
   Set objMail = Nothing
   Set objOutLook = Nothing
   Set f = Nothing
   Set fs = Nothing
   Set rsA = Nothing
End Function

'Add By Sindy 2019/7/17 從其他信箱轉寄到該信箱要通知其單位主管為副本收受者
'Modify By Sindy 2020/8/28 + strFullPath As String, strSubject As String
Public Function OL_SendNotifyMailCC(strFormMailbox As String, strToMailbox As String, _
   strFullPath As String, strSubject As String, _
   strPkey1 As String, strPkey2 As String, strPkey3 As String, strTo As String, _
   strUpdDate As String, strUpdTime As String) As Boolean
   
Dim tmpArr As Variant
Dim j As Integer
Dim strContent As String
Dim bolSendMsg As Boolean 'Add By Sindy 2020/8/28
   
   OL_SendNotifyMailCC = True
   bolSendMsg = False
   tmpArr = Split(strTo, ";")
   '增加副本給主管
   For j = 0 To UBound(tmpArr)
      If tmpArr(j) <> "" Then
         strSql = "insert into inputrecord(IR01,IR02,IR03,IR04,IR11,IR12,IR13,IR24)" & _
                  " values(" & strPkey1 & _
                           "," & strPkey2 & _
                           ",'" & strPkey3 & "'" & _
                           ",'" & tmpArr(j) & "'," & strUpdDate & "," & _
                           strUpdTime & ",'" & strUserNum & "','Y')"
         cnnConnection.Execute strSql
         
         'Add By Sindy 2020/8/28
         If Dir(strFullPath) <> "" Then
            '收受者 非商標處人員 直接寄Outlook(夾帶郵件檔案)
            If UCase(strToMailbox) = "TM" And Left(PUB_GetST03(CStr(tmpArr(j))), 2) <> "P2" Then
               PUB_SendMail strUserNum, tmpArr(j), "", "【" & strFormMailbox & "轉來的信件副本】" & strSubject, vbCrLf & "信件內容參附件！", , strFullPath, , , , , , , , , False, , , False, , , False
               If bolMailSendOk = False Then OL_SendNotifyMailCC = False: Exit Function
               '該收受者上刪除日期時間人員
               strExc(0) = "update InputRecord set " & _
                           " ir08=" & strUpdDate & ",ir09=" & strUpdTime & ",ir10='" & strUserNum & "'" & _
                           " where ir01=" & strPkey1 & _
                             " and ir02=" & strPkey2 & _
                             " and ir03='" & strPkey3 & "'" & _
                             " and ir04='" & tmpArr(j) & "'"
               cnnConnection.Execute strExc(0)
            Else
               bolSendMsg = True
            End If
         Else
            bolSendMsg = True
         End If
         '2020/8/28 END
         
         '尚無通知信,要記錄一筆
         'Modify By Sindy 2020/1/30 為工作天才要寄信通知人員處理
         'Modify By Sindy 2020/8/28 + bolSendMsg = True
         If ChkWorkDay(strSrvDate(1)) = True And _
            (Format(time, "HHMMSS") >= "080000" And Format(time, "HHMMSS") < "183000") And _
            bolSendMsg = True Then
         '2020/1/30 END
            strExc(0) = "select MC01 from mailcache" & _
                        " where MC01=" & CNULL(strUserNum) & _
                        " and MC02=" & CNULL(CStr(tmpArr(j))) & " and mc05 is null"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               strContent = "請至案件管理系統的一般作業\系統收件區，進行看查。" & vbCrLf & vbCrLf & vbCrLf & _
                            "若為職代或主管查詢其他同仁信件時，請在員工編號欄輸入欲查詢同仁之員工編號再按畫面更新按鈕。"
               
               strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08) values(" & _
                        "'" & strUserNum & "','" & tmpArr(j) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                        ",'" & "【" & strFormMailbox & "轉來的信件副本】通知已有信件轉入系統收件區" & "','" & strContent & "')"
               cnnConnection.Execute strSql, intI
            End If
         End If
      End If
   Next j
End Function

'Add By Sindy 2019/9/3 依主旨解析寄件人員是誰
'傳入:主旨
'回傳:員工編號,主管
'Optional bolReplace As Boolean = True => 若找不到專業代號,要不要預設:True.預設 False.不要預設
Public Sub BySubjectToStaff(ByVal strSubject As String, ByVal strII11 As String, _
      ByRef StrST01 As String, ByRef strDirector As String, _
      Optional bolReplace As Boolean = True)
      
Dim strST17 As String 'Add By Sindy 2022/11/30
   
   '依外專承辦英文縮寫分信
   If StrST01 = "" Then
      Call PUB_IPDept_ToSortOutSub_F23(strSubject, strII11, "", "", "1", , StrST01, strDirector)
   End If
   '依外專程序英文縮寫分信
   If StrST01 = "" Then
      'Modify By Sindy 2022/11/30 DY/ => 改抓特殊設定
      'Call PUB_IPDept_ToSortOutSub_F22(strSubject, strII11, "", "", "1", "DY", strST01, strDirector)
      Call GetPrjSalesNM(Pub_GetSpecMan("外專承辦英文組主管"), , strST17) '外專承辦(DY/dy)
      If InStr(strST17, "/") > 0 Then
         strST17 = Left(strST17, InStr(strST17, "/") - 1)
      End If
      Call PUB_IPDept_ToSortOutSub_F22(strSubject, strII11, "", "", "1", UCase(strST17), StrST01, strDirector)
      '2022/11/30 END
      
      'Add By Sindy 2025/9/22 ex:PY/py : (回覆) 返稿-FMP專利新申請案; Your R: PI-250405/1V; Our/Ref: P-136116 [DATA.936] (陳亭妙)
      If StrST01 = "" Then
         '寰華案出去的信件,前頭會掛PY/各程序人員縮寫
         Call PUB_IPDept_ToSortOutSub_F22(strSubject, strII11, "", "", "1", "PY", StrST01, strDirector)
      End If
      '2025/9/22 END
      
      'Add By Sindy 2020/10/6
      'Modify By Sindy 2022/11/30 + And strSrvDate(1) < "20230601"
      If StrST01 = "" And strSrvDate(1) < "20230601" Then
         Call PUB_IPDept_ToSortOutSub_F22(strSubject, strII11, "", "", "1", "WW", StrST01, strDirector)
      End If
      '2020/10/6 END
      'Modify By Sindy 2022/11/30 DY/ => 改抓特殊設定
      If StrST01 = "" Then
         Call GetPrjSalesNM(Pub_GetSpecMan("S"), , strST17) '日文組經理
         Call PUB_IPDept_ToSortOutSub_F22(strSubject, strII11, "", "", "1", UCase(strST17), StrST01, strDirector)
      End If
      '2022/11/30 END
   End If
   '依國外業務拓展英文縮寫分信
   If StrST01 = "" Then
      Call PUB_IPDept_ToSortOutSub_F41(strSubject, strII11, "", "", "1", StrST01, strDirector)
   End If
   '依外專工程師英文縮寫分信
   If StrST01 = "" Then
      Call PUB_IPDept_ToSortOutSub_F21(strSubject, strII11, "", "", "1", StrST01, strDirector)
   End If
   '依外商英文縮寫分信
   If StrST01 = "" Then
      Call PUB_IPDept_ToSortOutSub_F1x(strSubject, strII11, "", "", "1", StrST01, strDirector)
   End If
   
   'Modify By Sindy 2022/7/8
'***************************************
'前面先檢查在職人員, 這裡再統一檢查離職人員
'***************************************
   '依外專承辦英文縮寫分信
   If StrST01 = "" Then
      Call PUB_IPDept_ToSortOutSub_F23(strSubject, strII11, "", "", "2", , StrST01, strDirector)
   End If
   '依外專程序英文縮寫分信
   If StrST01 = "" Then
      Call PUB_IPDept_ToSortOutSub_F22(strSubject, strII11, "", "", "2", "DY", StrST01, strDirector)
      'Add By Sindy 2020/10/6
      If StrST01 = "" Then
         Call PUB_IPDept_ToSortOutSub_F22(strSubject, strII11, "", "", "2", "WW", StrST01, strDirector)
      End If
      '2020/10/6 END
   End If
   '依國外業務拓展英文縮寫分信
   If StrST01 = "" Then
      Call PUB_IPDept_ToSortOutSub_F41(strSubject, strII11, "", "", "2", StrST01, strDirector)
   End If
   '依外專工程師英文縮寫分信
   If StrST01 = "" Then
      Call PUB_IPDept_ToSortOutSub_F21(strSubject, strII11, "", "", "2", StrST01, strDirector)
   End If
   '依外商英文縮寫分信
   If StrST01 = "" Then
      Call PUB_IPDept_ToSortOutSub_F1x(strSubject, strII11, "", "", "2", StrST01, strDirector)
   End If
'***************************************
   
   'Add By Sindy 2021/3/22 增加用寄信人查看是所內那一位員工發的信
   If StrST01 = "" Then
      Call BySenderToStaff(strII11, StrST01, strDirector)
   End If
   '2021/3/22 END
   
   If bolReplace = True Then
      If StrST01 = "" Then StrST01 = "QPGMR" 'Modify By Sindy 2025/10/14 改使用QPGMR
   End If
End Sub

'傳入:寄件者
'回傳:員工編號,主管
'Modify By Sindy 2024/5/15 +, Optional ByVal bolFindAllEmp As Boolean = False: True=查詢全所人員
Public Function BySenderToStaff(ByVal strSendder As String, _
   ByRef StrST01 As String, ByRef strDirector As String, _
   Optional ByVal bolFindAllEmp As Boolean = False) As Boolean
   
Dim RsQ As New ADODB.Recordset
Dim strQ As String
Dim varTemp As Variant, jj As Integer
   
   BySenderToStaff = False
   If Trim(strSendder) = "" Then Exit Function
   
   'Modify By Sindy 2024/5/15 多人時,抓為所內人員的第一筆
   varTemp = Split(strSendder, ";")
   For jj = 0 To UBound(varTemp)
      strSendder = varTemp(jj)
   '2024/5/15 END
      If InStr(strSendder, "張衛民") > 0 Then strSendder = Replace(strSendder, "張衛民", "張�恭�") 'Add By Sindy 2017/10/20
      If InStr(strSendder, "陳思穎") > 0 Then strSendder = Replace(strSendder, "陳思穎", "陳思��") 'Add By Sindy 2017/12/5
      '國外部人員
      'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
      'modify by sonia 2021/1/27 再排除F4104及F4105
      strQ = "select st01,st02,st03,a0908,st16,st52" & _
             " From staff,acc090" & _
             " where st04='1'"
      If bolFindAllEmp = False Then
         strQ = strQ & " and substr(st03,1,1)='F' and st03 not in('F31','F41','F00','F61')"
      End If
      strQ = strQ & _
             " and substr(st01,1,1)>'6' and substr(st01,1,1)<'F'" & _
             " and substr(st01,4,1)<>'9'" & _
             " and st01<>'F4102' and st01<>'F4104' and st01<>'F4105'" & _
             " and st03=a0901(+)" & _
             " and instr('" & ChgSQL(strSendder) & "',st02)>0" & _
             " order by st03,st01"
      RsQ.CursorLocation = adUseClient
      RsQ.Open strQ, cnnConnection, adOpenStatic, adLockReadOnly
      If RsQ.RecordCount <> 0 And RsQ.RecordCount > 0 Then
         RsQ.MoveFirst
   '      Do While Not RsQ.EOF
   '         If InStr(strSendder, RsQ.Fields("st02")) > 0 Then
               BySenderToStaff = True
               StrST01 = RsQ.Fields("st01")
               If "" & RsQ.Fields("st03") = "F23" Then '外專承辦
                  strDirector = "" & RsQ.Fields("a0908")
               ElseIf "" & RsQ.Fields("st03") = "F22" Then '外專程序
                  'Modify By Sindy 2021/3/22
                  'strDirector = "82045;85033;86013;86024" '86024.葉敏莉 4個主任
                  strDirector = "" & RsQ.Fields("st52")
                  '2021/3/22 END
               ElseIf "" & RsQ.Fields("st03") = "F21" Then '外專工程師
                  If "" & RsQ.Fields("st16") = "1" Then '機電
                     strDirector = Pub_GetSpecMan("T")
                  ElseIf "" & RsQ.Fields("st16") = "2" Then '化學
                     strDirector = Pub_GetSpecMan("R")
                  ElseIf "" & RsQ.Fields("st16") = "3" Then '日文
                     strDirector = "" & RsQ.Fields("st52")
                     'Modify By Sindy 2021/3/22 王文安不加發任何人
                     If Not (RsQ.Fields("st01") = "88003") Then
                        'Modify By Sindy 2022/5/19 加發工程師主管
                        strExc(10) = ""
                        strExc(10) = PUB_GetFCPEngSup(RsQ.Fields("st01"), True)
                        'strExc(10) = PUB_GetST70SirEmp(RsQ.Fields("st01"))
                        '2022/5/19 END
                        If InStr(strDirector, strExc(10)) = 0 Then
                           strDirector = strDirector & ";" & strExc(10)
                        End If
                     End If
                     '2021/3/22 END
                  ElseIf "" & RsQ.Fields("st16") = "4" Then '德文
                     strDirector = Pub_GetSpecMan("T1")
   '               Else
   '                  strDirector = RsQ.Fields("a0908")
                  End If
               ElseIf Left("" & RsQ.Fields("st03"), 2) = "F1" Then '外商
                  'Modify By Sindy 2021/6/23 陳經理退休
                  'strDirector = "68005" '68005.陳鳳英
                  '外商改68005為80030洪琬姿及78011葉易雲
                  strDirector = "80030;78011"
                  '2021/6/23 END
               ElseIf "" & RsQ.Fields("a0908") <> "" Then
                  strDirector = RsQ.Fields("a0908")
               End If
   '            Exit Do
   '         End If
   '         RsQ.MoveNext
   '      Loop
            Exit For 'Add By Sindy 2024/5/15
      End If
      RsQ.Close 'Add By Sindy 2024/5/15
   Next jj
   Set RsQ = Nothing 'Modify By Sindy 2024/5/15
End Function

'Add By Sindy 2018/3/6 不分信直接刪除
'strMailKind.信箱代號 = F.國外部 P.專利處
Public Function PUB_OutLookForKeyWordDel(strMailKind As String, strRecipients_1 As String, _
   strSender As String, strSubject As String, _
   Optional strRecipients_all As String, Optional ByRef strSendMan As String) As Boolean
   
Dim varTemp As Variant
Dim jj As Integer
   
   PUB_OutLookForKeyWordDel = False
   If strMailKind = "" Then Exit Function '一定要傳入信箱代號
   
   'Modify By Sindy 2017/10/20
   '2017/10/19 (星期四) 上午 09:26
   'FW: ◎RE: APAA 2017 in Auckland
   '楊雯芳,薛德璟,何金柱:檢查有設定收受者為某些人員時，不分信直接刪除
   If strRecipients_1 <> "" Then
      strSql = "select LK04,LK01,LK12 from ipdeptkeyword" & _
               " where LK12='" & strMailKind & "' and LK02='S' and LK03='3'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         varTemp = Split(RsTemp.Fields("LK04"), ";")
         For jj = 0 To UBound(varTemp)
            If InStr(UCase(strRecipients_1), UCase(varTemp(jj))) > 0 Then
               strSendMan = ""
               'Add By Sindy 2024/5/17 記錄使用次數
               cnnConnection.Execute "update ipdeptkeyword set LK16=LK16+1" & _
                                     " where LK01='" & RsTemp.Fields("LK01") & "' and LK12='" & RsTemp.Fields("LK12") & "'" _
                                     , intI
               '2024/5/17 END
               PUB_OutLookForKeyWordDel = True
               Exit Function
            End If
         Next jj
      End If
   End If
   '2017/10/20 END
   
   'Add By Sind 2018/3/6 智慧局發的收件成功通知信,直接刪除不需分給程序人員
   If InStr(UCase(strSender), UCase("tipo@tiponet.tipo.gov.tw")) > 0 And _
      (InStr(strSubject, "收件成功通知") > 0 Or InStr(strSubject, "電子公文通知") > 0) Then
      strSendMan = ""
      PUB_OutLookForKeyWordDel = True
      Exit Function
   End If
   'Add By Sindy 2019/4/15 婧瑄:如account@taie.COM.tw已經是收件人之一或副本收受者,ipdept分信時請排除,不要再轉發一次
   If strMailKind = "F" Then
      If InStr(UCase(strRecipients_all), UCase("account@taie.com.tw")) > 0 And _
         InStr(UCase(strSendMan), UCase("account")) > 0 Then
         strSendMan = Replace(strSendMan, ",account", "")
         strSendMan = Replace(strSendMan, "account", "")
         PUB_OutLookForKeyWordDel = True
         Exit Function
      End If
   End If
   '2019/4/15 END
End Function

'Add By Sindy 2022/2/9 前提:信件同時有寄ipdept及patent信箱時,才檢查
'strMailKind.信箱代號 = F.國外部 P.專利處
Public Function PUB_ProPatentAndIpDeptMail(strMailKind As String, strRecipients_1 As String, _
   strSender As String, strSubject As String, _
   Optional strRecipients_all As String, Optional ByRef strSendMan As String, _
   Optional ByVal strMailDate As String, Optional ByVal strMailTime As String, _
   Optional ByVal strCurII01 As String, Optional ByVal strCurII03 As String, _
   Optional ByRef strII08 As String, Optional ByRef strII15 As String, _
   Optional ByRef intSameCnt As Integer) As Boolean
Dim varTemp As Variant
Dim jj As Integer
Dim strPi18 As String, strPi19 As String, strPi20 As String, strPi21 As String '本所案號
Dim strChkSub As String
   
   PUB_ProPatentAndIpDeptMail = False
   If strMailKind = "" Then Exit Function '一定要傳入信箱代號
   
   'Add By Sindy 2022/2/9 前提:信件同時有寄ipdept及patent信箱時,才檢查:
   '在匯入信箱時,先檢查是否已有(主旨、寄件者、寄件時間)均相同之同一封信
   If InStr(UCase(strRecipients_all), UCase("patent@taie.")) > 0 And _
      InStr(UCase(strRecipients_all), UCase("ipdept@taie.")) > 0 Then
      strII08 = "": strII15 = "": intSameCnt = 0
      '要進專利處分信系統
      If InStr(UCase(strSendMan), UCase("patent")) > 0 Then
         strSql = "select count(*) from patentinput" & _
                  " where pi17='" & ChgSQL(strSubject) & "'" & _
                  " and pi11='" & ChgSQL(strSender) & "' and pi12=" & strMailDate & " and pi13=" & strMailTime & _
                  " and pi01||pi03<>'" & strCurII01 & strCurII03 & "'" & _
                  " order by pi01 desc,pi03 desc"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            intSameCnt = RsTemp.Fields(0)
            If intSameCnt > 0 Then
               strSql = "select * from patentinput" & _
                        " where pi17='" & ChgSQL(strSubject) & "'" & _
                        " and pi11='" & ChgSQL(strSender) & "' and pi12=" & strMailDate & " and pi13=" & strMailTime & _
                        " and pi01||pi03<>'" & strCurII01 & strCurII03 & "'" & _
                        " order by pi01 desc,pi03 desc"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  strII08 = RsTemp.Fields("pi01")
                  strII15 = RsTemp.Fields("pi03")
                  strPi18 = RsTemp.Fields("pi18")
                  strPi19 = RsTemp.Fields("pi19")
                  strPi20 = RsTemp.Fields("pi20")
                  strPi21 = RsTemp.Fields("pi21")
                  '國外部信箱要進專利處分信系統時
                  If strMailKind = "F" Then
                     strSql = "select * from ipdeptinput" & _
                              " where ii01=" & strCurII01 & _
                              " and ii03='" & strCurII03 & "'"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strSql)
                     If intI = 1 Then
                        '要去掉Patent以免下列程式將信寄到Patent
                        strSendMan = Replace(strSendMan, "patent", "")
                        strSendMan = Replace(strSendMan, ";;", ";")
                        If strSendMan = ";" Then strSendMan = ""
                        If strSendMan <> "" Then
                           If Left(strSendMan, 1) = ";" Then strSendMan = Mid(strSendMan, 2)
                           If Right(strSendMan, 1) = ";" Then strSendMan = Mid(strSendMan, 1, Len(strSendMan) - 1)
                        End If
                        
                        '重覆信件:
                        '檢查案號是否相同,不同時改成"其他"由人工判斷;相同則直接上刪除日期
                        If InStr(RsTemp.Fields("ii18"), "(CFP-)") > 0 Or InStr(RsTemp.Fields("ii18"), "(CPS-)") > 0 Then
                           If InStr(RsTemp.Fields("ii18"), "(CFP-)") > 0 Then
                              strChkSub = Replace(strSubject, "CFP--", "CFP-")
                           End If
                           If InStr(RsTemp.Fields("ii18"), "(CPS-)") > 0 Then
                              strChkSub = Replace(strSubject, "CPS--", "CPS-")
                           End If
                           '案號相同,直接上刪除日期
                           If InStr(strChkSub, strPi18 & "-" & strPi19) > 0 Or _
                              InStr(strChkSub, strPi18 & " " & strPi19) > 0 Then
                              PUB_ProPatentAndIpDeptMail = True
                              Exit Function
                           Else
                              '不同時,改成"其他"由人工判斷
                              strII08 = "": strII15 = ""
                              strSql = "update ipdeptinput set ii05='Z',ii06=null" & _
                                       " where ii01=" & strCurII01 & _
                                       " and ii03='" & strCurII03 & "'"
                              cnnConnection.Execute strSql
                              Exit Function
                           End If
                        '依非案號規則，一樣轉入專利處信件系統分類的註記會顯示F*（*代表直接分給專利處，國外部人員沒經手若分錯要退回ipdept）。
                        Else
                           PUB_ProPatentAndIpDeptMail = True
                           Exit Function
                        End If
                        
                     End If
                  End If
               End If
            End If
         End If
'      '要進國外部分信系統
'      ElseIf InStr(UCase(strSendMan), UCase("ipdept")) > 0 Then
'         strSql = "select count(*) from ipdeptinput" & _
'                  " where ii17='" & ChgSQL(strSubject) & "'" & _
'                  " and ii11='" & ChgSQL(strSender) & "' and ii12=" & strMailDate & " and ii13=" & strMailTime & _
'                  " and ii01||ii03<>'" & strCurII01 & strCurII03 & "'" & _
'                  " order by ii01 desc,ii03 desc"
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'         If intI = 1 Then
'            intSameCnt = RsTemp.Fields(0)
'            If intSameCnt > 0 Then
'               strSql = "select * from ipdeptinput" & _
'                        " where ii17='" & ChgSQL(strSubject) & "'" & _
'                        " and ii11='" & ChgSQL(strSender) & "' and ii12=" & strMailDate & " and ii13=" & strMailTime & _
'                        " and ii01||ii03<>'" & strCurII01 & strCurII03 & "'" & _
'                        " order by ii01 desc,ii03 desc"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'               If intI = 1 Then
'                  strII08 = RsTemp.Fields("ii01")
'                  strII15 = RsTemp.Fields("ii03")
'                  strSendMan = ""
'                  PUB_ProPatentAndIpDeptMail = True
'                  Exit Function
'               End If
'            End If
'         End If
      End If
   End If
   '2022/2/9 END
End Function

'Add By Sindy 2019/4/19
'將匯入商標處的郵件, 再依商標處分類的規則重新分類
Public Function PUB_IPDeptChangeTM(oForm As Form) As Boolean
Dim rsA As ADODB.Recordset
Dim m_strTi11 As String, strTi06 As String, strTi15 As String
Dim strTi05 As String
Dim m_strTi01 As String, m_strTi03 As String, m_strTi15 As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim m_strTi11tmp As String
   
On Error GoTo ErrHand
   
   PUB_IPDeptChangeTM = False
   Screen.MousePointer = vbHourglass
   '檢查是否有待分類的信件
   strExc(0) = "select * from TMInput" & _
               " where TI05 is null"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsA.MoveFirst
      Do While Not rsA.EOF
         m_strTi01 = rsA.Fields("Ti01")
         m_strTi03 = rsA.Fields("Ti03")
         m_strTi15 = "" & rsA.Fields("Ti15") '系統記錄
         m_strTi11 = "" & rsA.Fields("Ti11") '寄件者
         oForm.TextII17 = "" & rsA.Fields("Ti17") '主旨
         '使用商標處的分信規則
         If InStr(m_strTi11, "[") > 0 Then
            m_strTi11tmp = Mid(m_strTi11, InStr(m_strTi11, "[") + 1)
            'Modify By Sindy 2022/5/17
            If InStr(m_strTi11tmp, "]") = 0 Then
               m_strTi11tmp = m_strTi11tmp & "]"
            End If
            '2022/5/17 END
            m_strTi11tmp = Left(m_strTi11tmp, InStr(m_strTi11tmp, "]") - 1)
            strTi05 = PUB_TM_ToSortOut(oForm, oForm.TextII17, m_strTi11tmp, strTi06, strCP01, strCP02, strCP03, strCP04, strTi15)
         Else
            strTi05 = PUB_TM_ToSortOut(oForm, oForm.TextII17, m_strTi11, strTi06, strCP01, strCP02, strCP03, strCP04, strTi15)
         End If
         If strTi15 <> "" Then '關鍵字
            m_strTi15 = m_strTi15 & ";" & strTi15
         End If
         If strTi05 <> "" Then
            strExc(0) = "select decode('" & strTi05 & "'," & Show商標處信件分類 & ",'" & strTi05 & "') 分類 from dual"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_strTi15 = m_strTi15 & ";" & RsTemp.Fields(0)
            End If
         End If
         '記錄本所案號
         If strCP01 <> "" Then
            m_strTi15 = m_strTi15 & ";" & strCP01 & strCP02 & strCP03 & strCP04
         End If
         m_strTi15 = Replace(m_strTi15, ";;", ";")
         strSql = "update TMInput set Ti05=" & CNULL(strTi05) & _
                  ",Ti06=" & CNULL(strTi06) & _
                  ",Ti15=" & CNULL(ChgSQL(m_strTi15)) & _
                  ",Ti18=" & CNULL(strCP01) & _
                  ",Ti19=" & CNULL(strCP02) & _
                  ",Ti20=" & CNULL(strCP03) & _
                  ",Ti21=" & CNULL(strCP04) & _
                  " where Ti01=" & m_strTi01 & _
                  " and Ti03='" & m_strTi03 & "'"
         cnnConnection.Execute strSql
         rsA.MoveNext
      Loop
   End If
   PUB_IPDeptChangeTM = True
   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHand:
   PUB_IPDeptChangeTM = True
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then MsgBox " 從國外部或專利處信箱匯入,欲分類失敗！" & vbCrLf & Err.Description
End Function

'Add By Sindy 2018/1/5
'將匯入專利處的郵件, 再依專利處分類的規則重新分類
Public Function PUB_IPDeptChangePatent(oForm As Form) As Boolean
Dim rsA As ADODB.Recordset
Dim m_strPI11 As String, strPI06 As String, strPI15 As String
Dim strPI05 As String
Dim m_strPI01 As String, m_strPI03 As String, m_strPI15 As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   
On Error GoTo ErrHand
   
   PUB_IPDeptChangePatent = False
   Screen.MousePointer = vbHourglass
   '檢查是否有待分類的信件
   strExc(0) = "select * from PatentInput" & _
               " where PI05 is null"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsA.MoveFirst
      Do While Not rsA.EOF
         m_strPI01 = rsA.Fields("PI01")
         m_strPI03 = rsA.Fields("PI03")
         m_strPI15 = "" & rsA.Fields("PI15") '系統記錄
         m_strPI11 = "" & rsA.Fields("PI11") '寄件者
         oForm.TextII17 = "" & rsA.Fields("PI17") '主旨
         '使用專利處的分信規則
         strPI05 = PUB_Patent_ToSortOut(oForm, oForm.TextII17, m_strPI11, strPI06, strCP01, strCP02, strCP03, strCP04, strPI15)
         If strPI15 <> "" Then '關鍵字
            m_strPI15 = m_strPI15 & ";" & strPI15
         End If
         If strPI05 <> "" Then
            strExc(0) = "select decode('" & strPI05 & "'," & Show專利處信件分類 & ",'" & strPI05 & "') 分類 from dual"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               m_strPI15 = m_strPI15 & ";" & RsTemp.Fields(0)
            End If
         End If
         '記錄本所案號
         If strCP01 <> "" Then
            m_strPI15 = m_strPI15 & ";" & strCP01 & strCP02 & strCP03 & strCP04
         End If
         strSql = "update PatentInput set PI05=" & CNULL(strPI05) & _
                  ",PI06=" & CNULL(strPI06) & _
                  ",PI15=" & CNULL(ChgSQL(m_strPI15)) & _
                  ",PI18=" & CNULL(strCP01) & _
                  ",PI19=" & CNULL(strCP02) & _
                  ",PI20=" & CNULL(strCP03) & _
                  ",PI21=" & CNULL(strCP04) & _
                  " where PI01=" & m_strPI01 & _
                  " and PI03='" & m_strPI03 & "'"
         cnnConnection.Execute strSql
         
         '別的信箱轉至PATENT信箱的信件，郭經理尚未操作確認前，在PATENT分信時若已有本所案號者，則將郭經理的記錄上刪除人員QPGMR
         If strCP01 <> "" Then
            'Add By Sindy 2020/6/9 副本人員未上核銷,因有本所案號了,系統自動核銷
            strSql = "update InputRecord set " & _
                     " ir08=" & strSrvDate(1) & ",ir09=" & Right("000000" & ServerTime, 6) & ",ir10='QPGMR'" & _
                     " where ir01=" & m_strPI01 & _
                     " and ir03='" & m_strPI03 & "'" & _
                     " and ir04='79075'" & _
                     " and ir24='Y'" & _
                     " and ir08=0"
            cnnConnection.Execute strSql, intI
            '2020/6/9 END
         End If
         rsA.MoveNext
      Loop
   End If
   PUB_IPDeptChangePatent = True
   Screen.MousePointer = vbDefault
   Exit Function
   
ErrHand:
   PUB_IPDeptChangePatent = True
   Screen.MousePointer = vbDefault
   If Err.Number <> 0 Then MsgBox " 從國外部或商標處信箱匯入,欲分類失敗！" & vbCrLf & Err.Description
End Function

'Add By Sindy 2017/9/8
Public Function PUB_PatentTransMail(oForm As Form, Optional ByRef strTo As String, _
   Optional ByRef strErrText As String, Optional ByRef strPI05 As String, _
   Optional ByVal strProFileName As String = "", Optional ByRef strCaseNo As String = "") As Boolean
Dim objOutLook As Object
Dim oFileSys As New FileSystemObject
Dim oFolder As Folder
Dim fs, f
Dim strPI03 As String, strPI03_2 As String, strPI11 As String, strPI12 As String, strPI13 As String
Dim strUpdTime As String
Dim stFtpPath As String
Dim strPI06 As String, strPI17 As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
'Modified by Morgan 2017/9/30 不可宣告 File 型態,會與 File 專案名稱相同會導致該專案無法編譯,改為 Object 型態
'Dim oFile As File
Dim oFile As Object
'end 2017/9/30
Dim objMail As Object
Dim bolSaveEFile As Boolean
Dim lngRonCnt As Long
Dim bolConnect As Boolean
Dim intPI03 As Integer
Dim strPI15 As String
Dim strRecipients_1 As String 'Add By Sindy 2022/2/8 收件者
Dim strRecipients_all As String 'Add By Sindy 2022/2/8 全部含副本等收件者
Dim intSameCnt As Integer
   
On Error GoTo ErrHand
   
   PUB_PatentTransMail = False
   strErrText = "": strPI15 = ""
   Set oFolder = oFileSys.GetFolder(oForm.txtPathPatent.Text)
   Set objOutLook = CreateObject("Outlook.Application")
   Set fs = CreateObject("Scripting.FileSystemObject")
   lngRonCnt = 0
   For Each oFile In oFolder.files
      lngRonCnt = lngRonCnt + 1
      oForm.LblCntIPDept.Caption = "已處理件數 / 剩餘件數：" & lngRonCnt & " / " & oFolder.files.Count
      DoEvents
      oForm.TxtIPDept = oFile.Name
      
      If UCase(Right(Trim(oFile.Name), 4)) = UCase(".msg") And _
         (strProFileName = "N" Or UCase(Trim(strProFileName)) = UCase(Trim(oFile.Name))) Then
         Call PUB_ExLetterTransTxt(oFile, oForm.TxtIPDept) '與國外部共用Function
         
         strTo = "" '轉寄人員
         Set objMail = objOutLook.CreateItemFromTemplate(oForm.txtPathPatent.Text & "\" & oFile.Name)
         DoEvents 'Add By Sindy 2019/12/13
         Screen.MousePointer = vbHourglass
         
         'strPI03 = Trim(oFile.Name)
'         strPI17 = ChgSQL(objMail.Subject)
'         oForm.TextII17 = objMail.Subject 'Add By Sindy 2017/11/22 Find簡體字
''         oForm.Text2 = strPI17 'Re: ML/kc 中?特許出願201510920053.X　貴所整理番?31565－CN　弊所整理番?：P-112987
''         strPI17 = ChgSQL(oForm.Text2) '要用文字框存放，因才能把unicode去掉
'         DoEvents
''         If strPI17 <> objMail.Subject Then
''            MsgBox "主旨抓的有誤，請洽電腦中心！"
''            GoTo ErrHand
''         End If

         'Modify By Sindy 2025/2/17
         strRecipients_1 = "" '收件者
         strRecipients_all = ""
'         If objMail.Class = 46 Then '46.olReport
'            strPI11 = "未傳遞的主旨"
'            strPI12 = "0"
'            strPI13 = ""
'         Else
'            strPI11 = PUB_GetMail_ii11(objMail) 'Modify By Sindy 2024/7/30
''            If objMail.SenderEmailType = "EX" Then
''               strPI11 = objMail.SenderName
''            Else
''               If objMail.SenderName = objMail.senderemailaddress Then
''                  strPI11 = objMail.senderemailaddress
''               Else
''                  'Add By Sindy 2024/7/29
''                  If InStr(UCase(objMail.senderemailaddress), UCase("Recipients/cn=")) > 0 Then
''                     strPI11 = objMail.SenderName
''                  Else
''                  '2024/7/29 END
''                     strPI11 = objMail.SenderName & " [" & objMail.senderemailaddress & "]"
''                  End If
''               End If
''            End If
'            strPI12 = Format(objMail.SentOn, "YYYYMMDD")
'            strPI13 = Format(objMail.SentOn, "HHMMSS")
'
'            'Add By Sindy 2022/2/8
'            '抓收件者資料
'            Call PUB_ReadMailText_CC(objMail, strRecipients_all, strRecipients_1)
'            '2022/2/8 END
'         End If
         Call PUB_ReadMailText(objMail, strRecipients_all, strRecipients_1, , strPI11, strPI12, strPI13, strPI17)
         oForm.TextII17 = strPI17 'Add By Sindy 2017/11/22 Find簡體字
         '2025/2/17 END
         strPI05 = PUB_Patent_ToSortOut(oForm, strPI17, strPI11, strPI06, strCP01, strCP02, strCP03, strCP04, strPI15)
         strUpdTime = Right("000000" & ServerTime, 6)
         
         cnnConnection.BeginTrans
         bolConnect = True
         '存實體檔案到File Server
         '專利處信件區
'         strExc(0) = "select count(*) from PatentInput" & _
'                     " where PI01=" & strSrvDate(1)
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            intPI03 = Val(RsTemp.Fields(0)) + 1
'         Else
'            intPI03 = 1
'         End If
'         strPI03 = "P" & Format(intPI03, "0000")
         'Modify By Sindy 2019/12/2 自動給號,才能 Keep PKey
         strPI03 = AutoNoByDate("P", 4)
         '2019/12/2 END
         strPI03_2 = strSrvDate(1) & strUpdTime & "." & strPI03 & ".msg"
         bolSaveEFile = PUB_PutFtpFile(oForm.txtPathPatent.Text & "\" & oFile.Name, strSrvDate(1), strPI03_2, stFtpPath, UCase("PatentInput"))
         If bolSaveEFile = True Then
            '存資料到DB
            If Len(strPI11) > 100 Then
               strPI11 = Mid(strPI11, 1, 100)
            End If
            
            If strPI05 <> "" Then
               strExc(0) = "select decode('" & strPI05 & "'," & Show專利處信件分類 & ",'" & strPI05 & "') 分類 from dual"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If strPI15 <> "" Then '關鍵字
                     strPI15 = Replace(RsTemp.Fields(0) & ";" & strPI15, ";;", ";")
                  Else
                     strPI15 = RsTemp.Fields(0)
                  End If
               End If
            End If
            '記錄本所案號
            If strCP01 <> "" Then
               strPI15 = strPI15 & IIf(strPI15 <> "", ";", "") & strCP01 & strCP02 & strCP03 & strCP04
            End If
            'Modify By Sindy 2022/2/8 And Len(strRecipients_all) <= 200 : 收件者太多就不要存值了
            strPI15 = strPI15 & IIf(strRecipients_all <> "" And Len(strRecipients_all) <= 200, ";收件者:" & strRecipients_all, "") '+ 原收件者
            strSql = "insert into PatentInput(PI01,PI02,PI03,PI04,PI05,PI06,PI11,PI12,PI13,PI14,PI17,PI15,PI18,PI19,PI20,PI21)" & _
                     " values(" & strSrvDate(1) & "," & strUpdTime & _
                     ",'" & strPI03 & "','" & strUserNum & "'" & _
                     ",'" & strPI05 & "','" & strPI06 & "'" & _
                     "," & CNULL(ChgSQL(strPI11)) & "," & strPI12 & "," & CNULL(strPI13) & _
                     ",'" & ChgSQL(stFtpPath) & "','" & strPI17 & "','" & ChgSQL(strPI15) & _
                     "','" & strCP01 & "','" & strCP02 & "','" & strCP03 & "','" & strCP04 & "')"
            cnnConnection.Execute strSql
            
            'Add By Sindy 2018/3/6 不分信直接刪除
            If PUB_OutLookForKeyWordDel("P", "", strPI11, strPI17, strRecipients_all, strPI06) = True Then
               strSql = "update PatentInput set" & _
                           " pi07='Y',pi08=" & strSrvDate(1) & _
                           ",pi09=" & strUpdTime & ",pi10='" & strUserNum & "'" & _
                           ",pi16=" & strSrvDate(1) & ",pi06=null" & _
                           " where pi01=" & strSrvDate(1) & _
                             " and pi02=" & strUpdTime & _
                             " and pi03='" & strPI03 & "'"
               cnnConnection.Execute strSql
               'Modify By Sindy 2019/4/15 PUB_OutLookForKeyWordDel函數會回傳strII06變數值
               'strPI06 = "" '不須轉寄
            End If
            strTo = strPI06 '轉寄人員
'            '有收受者並且有分類者，直接 [轉寄]
'            If strTo <> "" And strPI05 <> "" Then
'            End If
            '2018/3/6 END
            
            '刪除PC端檔案
            'Kill 刪不掉 "C:\IPdept\【轉知】(1) 經濟部智慧財產局來函，自105年4月1日起提出發明專利加速審查、專利審查高速公路與支援利用專利審查高速公路之專利申請案尚未公開者，不必再申請提早公開；(2) 經濟部智慧財產局來函，公告修正「發明專利加速審查申請書及其申請須知」、「發明專利PPH審查申請書及其申請須知」與「發明專利TW-SUPA審查申請書」.msg"
            'Kill txtPathPatent.Text & "\" & oFile.Name
            Call fs.DeleteFile(oForm.txtPathPatent.Text & "\" & oFile.Name)
         ElseIf UCase(oForm.Name) = UCase("frmTaOutLook") Then '單筆,失敗結束
            GoTo ErrHand
         End If
         cnnConnection.CommitTrans
         bolConnect = False
         PUB_PatentTransMail = True 'Modify By Sindy 2019/12/11
      End If
   Next
   oForm.LblCntIPDept.Caption = "已處理件數 / 剩餘件數：" & lngRonCnt & " / " & oFolder.files.Count '最後再讀一次資料夾的檔案數
      
   Screen.MousePointer = vbDefault
   Set f = Nothing
   Set fs = Nothing
   Set oFolder = Nothing
   Set oFile = Nothing
   Set oFileSys = Nothing
   Set objMail = Nothing
   Set objOutLook = Nothing
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault
   If bolConnect = True Then cnnConnection.RollbackTrans
   strErrText = strErrText & "信件轉入失敗！" & vbCrLf & IIf(Err.Number <> 0, "Err.Number:" & Err.Number & ";" & vbCrLf & Err.Description, "")
   
   Set f = Nothing
   Set fs = Nothing
   Set oFolder = Nothing
   Set oFile = Nothing
   Set oFileSys = Nothing
   Set objMail = Nothing
   Set objOutLook = Nothing
End Function

'專利處
'系統分類:
'回傳:分類 及 收受者 及 本所案號
Public Function PUB_Patent_ToSortOut(oForm As Form, strSubject As String, strPI11 As String, _
      ByRef m_Sender As String, ByRef strCP01 As String, ByRef strCP02 As String, _
      ByRef strCP03 As String, ByRef strCP04 As String, Optional ByRef strPI15 As String) As String
Dim strPA150 As String, strCP13 As String
Dim strText As String
Dim rsTmp As New ADODB.Recordset
Dim tmpArr As Variant
Dim YourRefCase As String, OurRefCase As String
Dim strTemp1 As String, strTemp2 As String, strTemp3 As String, StrTemp4 As String
Dim strTemp As String
Dim pa() As String, sp() As String
Dim j As Integer
Dim LongSelLength As Long, intGrd2 As Integer 'Add By Sindy 2017/11/23
Dim strOldSender As String
Dim bolChkOk As Boolean, strWord As String
   
   PUB_Patent_ToSortOut = "": m_Sender = ""
   strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
   YourRefCase = "": OurRefCase = "": strPI15 = ""
   
   '依系統別抓取本所案號
   'If strCP01 = "" Then
   'If strCP01 <> "CFP" And strCP01 <> "CPS" And strCP01 <> "PS" And strCP01 <> "P" Then
   'Modify By Sindy 2018/11/5 排除FCP-
   'FW: TW Patent No. I509237; Your Ref.: FCP-050125; Our Ref.: 2014-OPA-6486/TW_annuity fee
'   strText = strSubject
'   If PUB_IPDeptGetCaseNo(strText, "FCP-", strCP01, strCP02, strCP03, strCP04, strPA150, True) = True Then
'      strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
'   Else
   '2018/11/5 END
      'Modify By Sindy 2021/9/29 + , , False : 正規方式檢查本所案號
      strText = strSubject
      If PUB_IPDeptGetCaseNo(strText, "CFP-", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , False) = False Then
         strText = strSubject
         If PUB_IPDeptGetCaseNo(strText, "P-", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , False) = False Then
            strText = strSubject
            If PUB_IPDeptGetCaseNo(strText, "YOURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , Patent收件匣, , False) = False Then
               strText = strSubject
               If PUB_IPDeptGetCaseNo(strText, "OURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , Patent收件匣, , False) = False Then
                  strText = strSubject
                  If PUB_IPDeptGetCaseNo(strText, "CPS-", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , False) = False Then
                     strText = strSubject
                     If PUB_IPDeptGetCaseNo(strText, "PS-", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , False) = False Then
                        'Modify By Sindy 2018/1/16
                        '關於「用於減少局部脂肪與減少體重的組合物及其醫藥品與應用」美國案(CFP28292)OA通知
                        '15/620521 (TWTAE0741, CFP028690) - OAallowance
                        strText = strSubject
                        If PUB_IPDeptGetCaseNo(strText, "CFP", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , False) = False Then
                           strText = strSubject
                           If PUB_IPDeptGetCaseNo(strText, "CPS", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , False) = False Then
                              strText = strSubject
                              If PUB_IPDeptGetCaseNo(strText, "PS", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , False) = False Then
                                 strText = strSubject
                                 If PUB_IPDeptGetCaseNo(strText, "P", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , False) = False Then
                                 End If
                              End If
                           End If
                        End If
                        '2018/1/16 END
                     End If
                  End If
               End If
            End If
         End If
      End If
      'Add By Sindy 2021/9/29 再用其他方式檢查案號
      If strCP01 = "" Or strCP02 = "" Then
         strText = strSubject
         If PUB_IPDeptGetCaseNo(strText, "CFP-", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , True) = False Then
            strText = strSubject
            If PUB_IPDeptGetCaseNo(strText, "P-", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , True) = False Then
               strText = strSubject
               If PUB_IPDeptGetCaseNo(strText, "YOURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , Patent收件匣, , True) = False Then
                  strText = strSubject
                  If PUB_IPDeptGetCaseNo(strText, "OURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , Patent收件匣, , True) = False Then
                     strText = strSubject
                     If PUB_IPDeptGetCaseNo(strText, "CPS-", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , True) = False Then
                        strText = strSubject
                        If PUB_IPDeptGetCaseNo(strText, "PS-", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , True) = False Then
                           'Modify By Sindy 2018/1/16
                           '關於「用於減少局部脂肪與減少體重的組合物及其醫藥品與應用」美國案(CFP28292)OA通知
                           '15/620521 (TWTAE0741, CFP028690) - OAallowance
                           strText = strSubject
                           If PUB_IPDeptGetCaseNo(strText, "CFP", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , True) = False Then
                              strText = strSubject
                              If PUB_IPDeptGetCaseNo(strText, "CPS", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , True) = False Then
                                 strText = strSubject
                                 If PUB_IPDeptGetCaseNo(strText, "PS", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , True) = False Then
                                    strText = strSubject
                                    If PUB_IPDeptGetCaseNo(strText, "P", strCP01, strCP02, strCP03, strCP04, strPA150, True, Patent收件匣, , True) = False Then
                                    End If
                                 End If
                              End If
                           End If
                           '2018/1/16 END
                        End If
                     End If
                  End If
               End If
            End If
         End If
      End If
'   End If
   'Modify By Sindy 2020/9/23 排除FCP-
   'Modify By Sindy 2020/12/10 排除FCP- ex:FW: Taiwanese patent application 108135657 (Your Ref.: FCP-061998; Maiwald Ref.: B15898TW/HLZ)
   If strCP01 = "FCP" Or InStr(strSubject, "FCP-") > 0 Then
      strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
   End If
   '2020/9/23 END
   
   'Add By Sindy 2025/6/2 Y.比個案優先檢查
   PUB_Patent_ToSortOut = PUB_FindIPdeptKeyWord(" and lk11='Y' and LK12='P'", strSubject, strPI11, m_Sender, strPI15)
   If PUB_Patent_ToSortOut = "" Then
   '2025/6/2 END
      If strCP01 <> "" And strCP02 <> "" And InStr("P,PS,CFP,CPS", strCP01) > 0 Then
         '*****
         If PUB_PatentByChkCP14(strSubject, strPI11, strCP01, strCP02, strCP03, strCP04, PUB_Patent_ToSortOut, m_Sender) = False Then
            'GoTo ChkEnd
         End If
         '*****
      End If
   End If
   
   '********************************************************************************************************
   '關鍵字索引
   '********************************************************************************************************
   If PUB_Patent_ToSortOut = "" Then
      m_Sender = ""
      'Modify By Sindy 2020/3/6 將索引關鍵字改為共用函數
      PUB_Patent_ToSortOut = PUB_FindIPdeptKeyWord(" and lk12='P'", strSubject, strPI11, m_Sender, strPI15)
      '2020/3/6 END
''      'Modify By Sindy 2018/1/10
''      'strSql = "select LK01,LK02,LK03,LK04 from ipdeptkeyword where LK12='P' order by LK13 asc,LK01 asc"
''      strSql = "select lk01,lk02,lk03,lk04,lk13,lk14 from ipdeptkeyword where lk12='P' and lk14 is null" & _
''               " union select ' '||rtrim(ltrim(lk01))||' ' lk01,lk02,lk03,lk04,LK13,lk14 from ipdeptkeyword where lk12='P' and lk14='Y'" & _
''               " order by lk13 asc,lk01 asc"
''      '2018/1/10 END
'      'Modify By Sindy 2018/5/17
'      strSql = "select lk01,lk02,lk03,lk04,lk13,lk14 from ipdeptkeyword" & _
'               " where lk12='P' and lk14 is null and lk03='1'" & _
'               " and InStr('" & UCase(ChgSQL(strSubject)) & "',upper(rtrim(ltrim(lk01)))) > 0" & _
'               " union select ' '||rtrim(ltrim(lk01))||' ' lk01,lk02,lk03,lk04,LK13,lk14 from ipdeptkeyword" & _
'               " where lk12='P' and lk14='Y' and lk03='1'" & _
'               " and InStr('" & UCase(ChgSQL(strSubject)) & "',upper(rtrim(ltrim(lk01)))) > 0" & _
'               " union select lk01,lk02,lk03,lk04,lk13,lk14 from ipdeptkeyword" & _
'               " where lk12='P' and lk14 is null and lk03='2'" & _
'               " and InStr('" & UCase(ChgSQL(strPI11)) & "',upper(rtrim(ltrim(lk01)))) > 0" & _
'               " union select ' '||rtrim(ltrim(lk01))||' ' lk01,lk02,lk03,lk04,LK13,lk14 from ipdeptkeyword" & _
'               " where lk12='P' and lk14='Y' and lk03='2'" & _
'               " and InStr('" & UCase(ChgSQL(strPI11)) & "',upper(rtrim(ltrim(lk01)))) > 0" & _
'               " order by lk13 asc,lk01 asc"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'      If intI = 1 Then
'         With RsTemp
'            RsTemp.MoveFirst
'            Do While RsTemp.EOF = False
''               If RsTemp.Fields("LK03") = "1" Then '主旨
''                  If InStr(UCase(strSubject), UCase(RsTemp.Fields("LK01"))) > 0 Then
''                     strPI15 = RsTemp.Fields("LK01")
''                     PUB_Patent_ToSortOut = RsTemp.Fields("LK02")
''                     If "" & RsTemp.Fields("LK04") <> "" Then '有收受者
''                        tmpArr = Split(RsTemp.Fields("LK04"), ";")
''                        For j = 0 To UBound(tmpArr)
''                           strTemp = Pub_GetSpecMan(CStr(tmpArr(j)))
''                           If strTemp <> "" Then
''                              m_Sender = m_Sender & ";" & strTemp
''                           Else
''                              m_Sender = m_Sender & ";" & tmpArr(j)
''                           End If
''                        Next j
''                     End If
''                  'Add By Sindy 2018/1/10 單字索引
''                  Else
'                  If RsTemp.Fields("LK03") = "1" Then '主旨
'                     strTemp = Trim(UCase(strSubject))
'                  Else
'                     strTemp = Trim(UCase(strPI11))
'                  End If
'                  If "" & RsTemp.Fields("lk14") = "Y" Then
'                     bolChkOk = False
'                     '索引最前面
'                     strWord = Trim(UCase(RsTemp.Fields("LK01"))) & " "
'                     If Left(strTemp, Len(strWord)) = strWord Then
'                        bolChkOk = True
'                     End If
'                     If bolChkOk = False Then
'                        '索引最後面
'                        strWord = " " & Trim(UCase(RsTemp.Fields("LK01")))
'                        If Right(strTemp, Len(strWord)) = strWord Then
'                           bolChkOk = True
'                        End If
'                     End If
'                     'Add By Sindy 2018/6/27
'                     If bolChkOk = False Then
'                        '索引中間
'                        strWord = " " & Trim(UCase(RsTemp.Fields("LK01"))) & " "
'                        If InStr(strTemp, strWord) > 0 Then
'                           bolChkOk = True
'                        End If
'                     End If
'                     '2018/6/27 END
'                     '有索引到
'                     If bolChkOk = True Then
'                        strPI15 = RsTemp.Fields("LK01")
'                        PUB_Patent_ToSortOut = RsTemp.Fields("LK02")
'                        If "" & RsTemp.Fields("LK04") <> "" Then '有收受者
'                           tmpArr = Split(RsTemp.Fields("LK04"), ";")
'                           For j = 0 To UBound(tmpArr)
'                              strTemp = Pub_GetSpecMan(CStr(tmpArr(j)))
'                              If strTemp <> "" Then
'                                 m_Sender = m_Sender & ";" & strTemp
'                              Else
'                                 m_Sender = m_Sender & ";" & tmpArr(j)
'                              End If
'                           Next j
'                           Exit Do
'                        End If
'                     End If
'                  End If
'                  '2018/1/10 END
''               Else '寄件者或網域
''                  If InStr(UCase(strPI11), UCase(RsTemp.Fields("LK01"))) > 0 Then
''                     strPI15 = RsTemp.Fields("LK01")
''                     PUB_Patent_ToSortOut = RsTemp.Fields("LK02")
''                     If "" & RsTemp.Fields("LK04") <> "" Then '有收受者
''                        tmpArr = Split(RsTemp.Fields("LK04"), ";")
''                        For j = 0 To UBound(tmpArr)
''                           strTemp = Pub_GetSpecMan(CStr(tmpArr(j)))
''                           If strTemp <> "" Then
''                              m_Sender = m_Sender & ";" & strTemp
''                           Else
''                              m_Sender = m_Sender & ";" & tmpArr(j)
''                           End If
''                        Next j
''                     End If
''                  End If
''               End If
'               If PUB_Patent_ToSortOut <> "" Then Exit Do
'               RsTemp.MoveNext
'            Loop
'            'Add By Sindy 2018/5/17
'            If PUB_Patent_ToSortOut = "" Then
'               RsTemp.MoveFirst
'               Do While RsTemp.EOF = False
'                  If "" & RsTemp.Fields("lk14") = "" Then
'                     strPI15 = RsTemp.Fields("LK01")
'                     PUB_Patent_ToSortOut = RsTemp.Fields("LK02")
'                     m_Sender = ""
'                     tmpArr = Split("" & RsTemp.Fields("LK04"), ";")
'                     For j = 0 To UBound(tmpArr)
'                        strTemp = Pub_GetSpecMan(CStr(tmpArr(j)))
'                        If strTemp <> "" Then
'                           m_Sender = m_Sender & ";" & strTemp
'                        Else
'                           m_Sender = m_Sender & ";" & tmpArr(j)
'                        End If
'                     Next j
'                     Exit Do
'                  End If
'                  RsTemp.MoveNext
'               Loop
'            End If
'            '2018/5/17 END
'         End With
'      End If
   End If
   
'   'Add By Sindy 2017/11/23
'   If PUB_Patent_ToSortOut = "" Then
'      '比對簡體字檔案
'      oForm.GRD2.Clear: intGrd2 = 0: strOldSender = m_Sender
'      If oForm.GRD2.Rows > 0 Then
'         For j = oForm.GRD2.Rows - 1 To 1 Step -1
'            oForm.GRD2.RemoveItem j
'         Next j
'      End If
'      'oForm.TextBoxP.SetFocus
'      'If oForm.TextBoxP.LineCount > 0 Then
'      If oForm.TextBoxP <> "" Then
'         For j = 1 To Len(oForm.TextBoxP)
'            LongSelLength = InStr(Mid(oForm.TextBoxP, j), vbCrLf)
'            If LongSelLength = 0 Then LongSelLength = Len(oForm.TextBoxP)
'            oForm.TextBox3 = Replace(Mid(oForm.TextBoxP.Text, j, LongSelLength - 1), Chr(10), "")
'            If j > 1 Then '第一筆是標題,跳過
'               If oForm.TextBox3 <> "" Then
'                  '關鍵字簡體#分類#收受者#索引排序
'                  tmpArr = Split(oForm.TextBox3, "#")
'                  If UBound(tmpArr) = 3 Then
'                     strExc(10) = oForm.TextBox3
'                     oForm.TextBox3 = tmpArr(0)
'                     If InStr(oForm.TextII17, oForm.TextBox3) > 0 Then
'                        If PUB_Patent_ToSortOut = "" Then '記錄第一次比對到的值
'                           PUB_Patent_ToSortOut = Trim(tmpArr(1))
'                           strTemp = Pub_GetSpecMan(CStr(Trim(tmpArr(2))))
'                           If strTemp <> "" Then
'                              m_Sender = strOldSender & ";" & strTemp
'                           Else
'                              m_Sender = strOldSender & ";" & Trim(tmpArr(2))
'                           End If
'                        End If
'                     End If
'                     If Val(tmpArr(3)) > 0 Then '有排序欄位值的才記錄下來
'                        If Not (intGrd2 = 0 And oForm.GRD2.Rows = 1) Then
'                           oForm.GRD2.AddItem ""
'                        End If
'                        oForm.GRD2.TextMatrix(intGrd2, 0) = strExc(10)
'                        oForm.GRD2.TextMatrix(intGrd2, 1) = Val(tmpArr(3))
'                        intGrd2 = intGrd2 + 1
'                     End If
'                  End If
'               End If
'            End If
'            j = j + LongSelLength
'         Next j
'         '有排序欄位時,取得優先順序的第一筆資料
'         If intGrd2 > 0 Then
'            oForm.GRD2.col = 1
'            oForm.GRD2.row = 0
'            oForm.GRD2.Sort = 3 '3.數值昇冪 4.數值降冪
'            For j = 0 To intGrd2 - 1
'               tmpArr = Split(oForm.GRD2.TextMatrix(j, 0), "#")
'               '關鍵字簡體#分類#收受者#索引排序
'               If UBound(tmpArr) = 3 Then
'                  oForm.TextBox3 = tmpArr(0)
'                  If InStr(oForm.TextII17, oForm.TextBox3) > 0 Then
'                     PUB_Patent_ToSortOut = Trim(tmpArr(1))
'                     strTemp = Pub_GetSpecMan(CStr(Trim(tmpArr(2))))
'                     If strTemp <> "" Then
'                        m_Sender = strOldSender & ";" & strTemp
'                     Else
'                        m_Sender = strOldSender & ";" & Trim(tmpArr(2))
'                     End If
'                  End If
'               End If
'            Next j
'         End If
'
'      End If
'   End If
'   '2017/11/23 END
   
   'Modify By Sindy 2025/1/8
   If strSrvDate(1) < P業務區劃分啟用日 Then
   '2025/1/8 END
      If PUB_Patent_ToSortOut = "" Then
         If strCP01 = "P" Or strCP01 = "PS" Then 'P其餘歸程序2
            PUB_Patent_ToSortOut = "2"
            m_Sender = Pub_GetSpecMan("專利處轉信非台灣程序2")
         End If
      End If
   End If
   
   '條件檢查完畢:
ChkEnd:
   '無以上條件者分至其他；
   If m_Sender = "" And PUB_Patent_ToSortOut <> "8" Then PUB_Patent_ToSortOut = "" 'Modify By Sindy 2023/7/27
   If PUB_Patent_ToSortOut = "" Then PUB_Patent_ToSortOut = "7" '其他
   
   'Add by Sindy 2019/8/15
   If Trim(strPI11) = "未傳遞的主旨" Then
      PUB_Patent_ToSortOut = "7" '其他
      m_Sender = ""
   End If
   
   '過濾是否有收受者重覆的資料
   If Left(m_Sender, 1) = ";" Then m_Sender = Mid(m_Sender, 2)
   If m_Sender <> "" And InStr(m_Sender, ";") > 0 Then
      strText = m_Sender
      tmpArr = Split(strText, ";")
      m_Sender = ""
      For j = 0 To UBound(tmpArr)
         If tmpArr(j) <> "" Then
            If InStr(m_Sender, tmpArr(j)) = 0 Then
               m_Sender = m_Sender & IIf(m_Sender = "", "", ";") & tmpArr(j)
            End If
         End If
      Next j
   End If
   
   Set rsTmp = Nothing
End Function

'Add By Sindy 2019/3/29
Public Function PUB_TMTransMail(oForm As Form, Optional ByRef strTo As String, _
   Optional ByRef strErrText As String, Optional ByRef strTi05 As String, _
   Optional ByVal strProFileName As String = "", Optional ByRef strCaseNo As String = "") As Boolean
Dim objOutLook As Object
Dim oFileSys As New FileSystemObject
Dim oFolder As Folder
Dim fs, f
Dim strTi03 As String, strTi03_2 As String, strTi11 As String, strTi12 As String, strTi13 As String
Dim strUpdTime As String
Dim stFtpPath As String
Dim strTi06 As String, strTi17 As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim tmpArr As Variant

'不可宣告 File 型態,會與 File 專案名稱相同會導致該專案無法編譯,改為 Object 型態
'Dim oFile As File
Dim oFile As Object

Dim objMail As Object
Dim bolSaveEFile As Boolean
Dim lngRonCnt As Long
Dim bolConnect As Boolean
Dim intTi03 As Integer
Dim strTi15 As String
Dim strRecipients_1 As String 'Add By Sindy 2022/2/8 收件者
Dim strRecipients_all As String 'Add By Sindy 2022/2/8 全部含副本等收件者
   
On Error GoTo ErrHand
   
   PUB_TMTransMail = False
   strErrText = "": strTi15 = ""
   Set oFolder = oFileSys.GetFolder(oForm.txtPathTM.Text)
   Set objOutLook = CreateObject("Outlook.Application")
   Set fs = CreateObject("Scripting.FileSystemObject")
   lngRonCnt = 0
   For Each oFile In oFolder.files
      lngRonCnt = lngRonCnt + 1
      oForm.LblCntIPDept.Caption = "已處理件數 / 剩餘件數：" & lngRonCnt & " / " & oFolder.files.Count
      DoEvents
      oForm.TxtIPDept = oFile.Name
      
      If UCase(Right(Trim(oFile.Name), 4)) = UCase(".msg") And _
         (strProFileName = "N" Or UCase(Trim(strProFileName)) = UCase(Trim(oFile.Name))) Then
         Call PUB_ExLetterTransTxt(oFile, oForm.TxtIPDept) '與國外部共用Function
         
         strTo = "" '轉寄人員
         Set objMail = objOutLook.CreateItemFromTemplate(oForm.txtPathTM.Text & "\" & oFile.Name)
         DoEvents 'Add By Sindy 2019/12/13
         Screen.MousePointer = vbHourglass
         
         'strTi03 = Trim(oFile.Name)
'         strTi17 = ChgSQL(objMail.Subject)
'         oForm.TextII17 = objMail.Subject 'Add By Sindy 2017/11/22 Find簡體字
''         oForm.Text2 = strTi17 'Re: ML/kc 中?特許出願201510920053.X　貴所整理番?31565－CN　弊所整理番?：P-112987
''         strTi17 = ChgSQL(oForm.Text2) '要用文字框存放，因才能把unicode去掉
'         DoEvents
''         If strTi17 <> objMail.Subject Then
''            MsgBox "主旨抓的有誤，請洽電腦中心！"
''            GoTo ErrHand
''         End If
         
         'Modify By Sindy 2025/2/17
         strRecipients_1 = "" '收件者
         strRecipients_all = ""
'         If objMail.Class = 46 Then '46.olReport
'            strTi11 = "未傳遞的主旨"
'            strTi12 = "0"
'            strTi13 = ""
'         Else
'            strTi11 = PUB_GetMail_ii11(objMail) 'Modify By Sindy 2024/7/30
''            If objMail.SenderEmailType = "EX" Then
''               strTi11 = objMail.SenderName
''            Else
''               If objMail.SenderName = objMail.senderemailaddress Then
''                  strTi11 = objMail.senderemailaddress
''               Else
''                  'Add By Sindy 2024/7/29
''                  If InStr(UCase(objMail.senderemailaddress), UCase("Recipients/cn=")) > 0 Then
''                     strTi11 = objMail.SenderName
''                  Else
''                  '2024/7/29 END
''                     strTi11 = objMail.SenderName & " [" & objMail.senderemailaddress & "]"
''                  End If
''               End If
''            End If
'            strTi12 = Format(objMail.SentOn, "YYYYMMDD")
'            strTi13 = Format(objMail.SentOn, "HHMMSS")
'
'            'Add By Sindy 2022/2/8
'            '抓收件者資料
'            Call PUB_ReadMailText_CC(objMail, strRecipients_all, strRecipients_1)
'            '2022/2/8 END
'         End If
         Call PUB_ReadMailText(objMail, strRecipients_all, strRecipients_1, , strTi11, strTi12, strTi13, strTi17)
         oForm.TextII17 = strTi17 'Add By Sindy 2017/11/22 Find簡體字
         '2025/2/17 END
         
'         If objMail.Class = 46 Then '46.olReport
            strTi05 = PUB_TM_ToSortOut(oForm, strTi17, strTi11, strTi06, strCP01, strCP02, strCP03, strCP04, strTi15)
'         Else
'            strTi05 = PUB_TM_ToSortOut(oForm, strTi17, objMail.senderemailaddress, strTi06, strCP01, strCP02, strCP03, strCP04, strTi15)
'         End If
         strUpdTime = Right("000000" & ServerTime, 6)
         
         cnnConnection.BeginTrans
         bolConnect = True
         '存實體檔案到File Server
         '商標處信件區
'         strExc(0) = "select count(*) from TMInput" & _
'                     " where Ti01=" & strSrvDate(1)
'         intI = 1
'         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'         If intI = 1 Then
'            intTi03 = Val(RsTemp.Fields(0)) + 1
'         Else
'            intTi03 = 1
'         End If
'         strTi03 = "T" & Format(intTi03, "0000")
         'Modify By Sindy 2019/12/2 自動給號,才能 Keep PKey
         strTi03 = AutoNoByDate("T", 4)
         '2019/12/2 END
         strTi03_2 = strSrvDate(1) & strUpdTime & "." & strTi03 & ".msg"
         bolSaveEFile = PUB_PutFtpFile(oForm.txtPathTM.Text & "\" & oFile.Name, strSrvDate(1), strTi03_2, stFtpPath, UCase("TMInput"))
         If bolSaveEFile = True Then
            '存資料到DB
            If Len(strTi11) > 100 Then
               strTi11 = Mid(strTi11, 1, 100)
            End If
            
            If strTi05 <> "" Then
               strExc(0) = "select decode('" & strTi05 & "'," & Show商標處信件分類 & ",'" & strTi05 & "') 分類 from dual"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If strTi15 <> "" Then '關鍵字
                     strTi15 = Replace(RsTemp.Fields(0) & ";" & strTi15, ";;", ";")
                  Else
                     strTi15 = RsTemp.Fields(0)
                  End If
               End If
            End If
            '記錄本所案號
            If strCP01 <> "" Then
               strTi15 = strTi15 & IIf(strTi15 <> "", ";", "") & strCP01 & strCP02 & strCP03 & strCP04
            End If
            strTi15 = Replace(strTi15, ";;", ";")
            'Modify By Sindy 2022/2/8 And Len(strRecipients_all) <= 200 : 收件者太多就不要存值了
            strTi15 = strTi15 & IIf(strRecipients_all <> "" And Len(strRecipients_all) <= 200, ";收件者:" & strRecipients_all, "") '+ 原收件者
            strSql = "insert into TMInput(Ti01,Ti02,Ti03,Ti04,Ti05,Ti06,Ti11,Ti12,Ti13,Ti14,Ti17,Ti15,Ti18,Ti19,Ti20,Ti21)" & _
                     " values(" & strSrvDate(1) & "," & strUpdTime & _
                     ",'" & strTi03 & "','" & strUserNum & "'" & _
                     ",'" & strTi05 & "','" & strTi06 & "'" & _
                     "," & CNULL(ChgSQL(strTi11)) & "," & strTi12 & "," & CNULL(strTi13) & _
                     ",'" & ChgSQL(stFtpPath) & "','" & strTi17 & "','" & ChgSQL(strTi15) & _
                     "','" & strCP01 & "','" & strCP02 & "','" & strCP03 & "','" & strCP04 & "')"
            cnnConnection.Execute strSql
            
            '不分信直接刪除
            If PUB_OutLookForKeyWordDel("T", "", strTi11, strTi17) = True Then
               strSql = "update TMInput set" & _
                           " Ti07='Y',Ti08=" & strSrvDate(1) & _
                           ",Ti09=" & strUpdTime & ",Ti10='" & strUserNum & "'" & _
                           ",Ti16=" & strSrvDate(1) & ",Ti06=null" & _
                           " where Ti01=" & strSrvDate(1) & _
                             " and Ti02=" & strUpdTime & _
                             " and Ti03='" & strTi03 & "'"
               cnnConnection.Execute strSql
               strTi06 = "" '不須轉寄
            End If
            strTo = strTi06 '轉寄人員
            '有收受者並且有分類者，直接 [轉寄]
            If strTo <> "" And strTi05 <> "" Then
            End If
            
            '刪除PC端檔案
            'Kill 刪不掉 "C:\IPdept\【轉知】(1) 經濟部智慧財產局來函，自105年4月1日起提出發明專利加速審查、專利審查高速公路與支援利用專利審查高速公路之專利申請案尚未公開者，不必再申請提早公開；(2) 經濟部智慧財產局來函，公告修正「發明專利加速審查申請書及其申請須知」、「發明專利PPH審查申請書及其申請須知」與「發明專利TW-SUPA審查申請書」.msg"
            'Kill txtPathTM.Text & "\" & oFile.Name
            Call fs.DeleteFile(oForm.txtPathTM.Text & "\" & oFile.Name)
         ElseIf UCase(oForm.Name) = UCase("frmTaOutLook") Then '單筆,失敗結束
            GoTo ErrHand
         End If
         cnnConnection.CommitTrans
         bolConnect = False
         PUB_TMTransMail = True 'Modify By Sindy 2019/12/11
      End If
   Next
   oForm.LblCntIPDept.Caption = "已處理件數 / 剩餘件數：" & lngRonCnt & " / " & oFolder.files.Count '最後再讀一次資料夾的檔案數
   
   Screen.MousePointer = vbDefault
   Set f = Nothing
   Set fs = Nothing
   Set oFolder = Nothing
   Set oFile = Nothing
   Set oFileSys = Nothing
   Set objMail = Nothing
   Set objOutLook = Nothing
   Exit Function
   
ErrHand:
   Screen.MousePointer = vbDefault
   'Resume
   If bolConnect = True Then cnnConnection.RollbackTrans
   strErrText = strErrText & "信件轉入失敗！" & vbCrLf & IIf(Err.Number <> 0, "Err.Number:" & Err.Number & ";" & vbCrLf & Err.Description, "")
   
   Set f = Nothing
   Set fs = Nothing
   Set oFolder = Nothing
   Set oFile = Nothing
   Set oFileSys = Nothing
   Set objMail = Nothing
   Set objOutLook = Nothing
End Function

'商標處
'系統分類:1.MCTF 2.大陸案 3.個人 4.非大陸案 5.其他 6.其他信箱匯入
'回傳:分類 及 收受者 及 本所案號
Public Function PUB_TM_ToSortOut(oForm As Form, strSubject As String, strTi11 As String, _
      ByRef m_Sender As String, ByRef strCP01 As String, ByRef strCP02 As String, _
      ByRef strCP03 As String, ByRef strCP04 As String, Optional ByRef strTi15 As String) As String

Dim strPA150 As String, strCP13 As String
Dim strText As String
Dim rsTmp As New ADODB.Recordset
Dim tmpArr As Variant
Dim YourRefCase As String, OurRefCase As String
Dim strTemp As String, strTemp1 As String
Dim j As Integer
Dim LongSelLength As Long, intGrd2 As Integer
Dim strOldSender As String
Dim bolChkOk As Boolean, strWord As String
Dim strKindEmp As String
Dim strMail As String, strDomain As String
Dim strTempSort As String
Dim n As Integer
Dim intPoint As Integer, strProCode As String
Dim strProCodeList(1 To 5) As String 'Add By Sindy 2020/12/16
Dim bolChkCaseNo As Boolean 'Add By Sindy 2021/3/25
   
   PUB_TM_ToSortOut = "": m_Sender = ""
   strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
   YourRefCase = "": OurRefCase = "": strTi15 = ""
   
   '丙規則(關鍵字)優先成立
   If PUB_TM_ToSortOut = "" Then
      PUB_TM_ToSortOut = PUB_FindIPdeptKeyWord(" and lk12='T' and lk11='丙'", strSubject, strTi11, m_Sender, strTi15)
      If PUB_TM_ToSortOut <> "" Then GoTo ChkEnd
   End If
   
'  1.若主旨關鍵字同時有甲及乙規則的信件,則同時以甲組及乙組的成員為收件人
   If PUB_TM_ToSortOut = "" Then
      '檢查是否有專業代號的規則:
      strProCode = "" '取得專業代號
      intPoint = 0 'Add By Sindy 2020/12/16
      If InStr(strSubject, "(") > 0 And InStr(strSubject, ")") > 0 Then
         strProCode = Mid(strSubject, InStr(strSubject, "("))
         strProCode = UCase(Mid(strProCode, 1, InStr(strProCode, ")"))) 'Add By Sindy 2020/9/10 專業代號轉大寫比對
         '()裡面要是單個英文字母(A~Z)或是- ;前後要是()
         'ex:(A) 或 (C-E)
         For n = 2 To Len(strProCode)
            strWord = Mid(strProCode, n, 2)
            'Modify By Sindy 2020/12/16 + Or (Mid(strWord, 2) >= 1 And Mid(strWord, 2) <= 9)
            'ex: Subject: 回复：(L1) RE: 台?商?注?
            If Not ((Asc(Left(strWord, 1)) >= 65 And Asc(Left(strWord, 1)) <= 90) And _
                   (Mid(strWord, 2) = "-" Or _
                    Mid(strWord, 2) = ")" Or _
                    (Mid(strWord, 2) >= 1 And Mid(strWord, 2) <= 9) _
                   )) Then
               strProCode = ""
               Exit For
            End If
            'Modify By Sindy 2020/12/16
            'n = n + 1
            intPoint = intPoint + 1
            If Mid(strWord, 2) = "-" Or Mid(strWord, 2) = ")" Then
               n = n + 1
               strProCodeList(intPoint) = Left(strWord, 1)
            Else
               n = n + 2
               strProCodeList(intPoint) = strWord
            End If
            '2020/12/16 END
         Next n
      End If
      If strProCode <> "" And intPoint > 0 Then '有專業代號
         'Modify By Sindy 2020/12/16
         For n = 1 To intPoint
            '甲規則(專業代號)
            strSql = "select st17,st01,st02,'' as LK12 from staff" & _
                     " where substr(st03,1,2)='P2' and st04='1' and ST17 is not null" & _
                     " and upper(st17)='" & strProCodeList(n) & "'" & _
                     " union select lk01,lk04,st02,LK12 from ipdeptkeyword,staff" & _
                     " where lk12='T' and lk11='甲' and lk04=st01(+)" & _
                     " and upper(rtrim(ltrim(lk01)))='(" & strProCodeList(n) & ")'"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               PUB_TM_ToSortOut = "3" '個人
               If InStr(strTi15, "(甲)") = 0 Then strTi15 = strTi15 & "(甲)"
               rsTmp.MoveFirst
               '取最前面的專業代號 => ex:答复: (E)回覆: (B)Southern Cross(第03類)商標，申請-核准 (我方案號：T226219)
               Do While Not rsTmp.EOF
                  m_Sender = m_Sender & ";" & rsTmp.Fields("st01")
                  'Add By Sindy 2024/5/17 記錄使用次數
                  If "" & rsTmp.Fields("LK12") <> "" Then
                     cnnConnection.Execute "update ipdeptkeyword set LK16=LK16+1" & _
                                           " where LK01='" & rsTmp.Fields("LK01") & "' and LK12='" & rsTmp.Fields("LK12") & "'" _
                                           , intI
                  End If
                  '2024/5/17 END
                  rsTmp.MoveNext
               Loop
            End If
         Next n
'         '甲規則(專業代號)
'         strSql = "select st17,st01,st02 from staff" & _
'                  " where substr(st03,1,2)='P2' and st04='1' and ST17 is not null" & _
'                  " and instr('" & UCase(ChgSQL(strProCode)) & "','('||st17||')') > 0" & _
'                  " union select lk01,lk04,st02 from ipdeptkeyword,staff" & _
'                  " where lk12='T' and lk11='甲' and lk04=st01(+)" & _
'                  " and InStr('" & UCase(ChgSQL(strProCode)) & "',upper(rtrim(ltrim(lk01)))) > 0"
'         intI = 1
'         Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'         If intI = 1 Then
'            PUB_TM_ToSortOut = "3" '個人
'            If InStr(strTi15, "(甲)") = 0 Then strTi15 = strTi15 & "(甲)"
'            rsTmp.MoveFirst
'            '取最前面的專業代號 => ex:答复: (E)回覆: (B)Southern Cross(第03類)商標，申請-核准 (我方案號：T226219)
'            Do While Not rsTmp.EOF
'               strTemp = Replace(rsTmp.Fields("st17"), "(", "")
'               strTemp = Replace(strTemp, ")", "")
'               If m_Sender = "" Then
'                  m_Sender = ";" & rsTmp.Fields("st01")
'                  intPoint = InStr(UCase(ChgSQL(strProCode)), "(" & strTemp & ")")
'               Else
'                  If intPoint > InStr(UCase(ChgSQL(strProCode)), "(" & strTemp & ")") Then
'                     m_Sender = ";" & rsTmp.Fields("st01")
'                     intPoint = InStr(UCase(ChgSQL(strProCode)), "(" & strTemp & ")")
'                  End If
'               End If
'               rsTmp.MoveNext
'            Loop
'         End If
'         '若出現(A-C)的狀況,解析多人
'         If InStr(strProCode, "(") > 0 And InStr(strProCode, ")") > 0 Then
'            strTempSort = Mid(strProCode, InStr(strProCode, "(") + 1)
'            strTempSort = Trim(Mid(strTempSort, 1, InStr(strTempSort, ")") - 1))
'            'strTempSort = Trim(UCase(strTempSort))
'            If Len(strTempSort) > 1 And InStr(strTempSort, "-") > 0 Then
'               '資料只存在 - 及A~Z,才是要解析的資料
'               strKindEmp = ""
'               For n = 1 To Len(strTempSort)
'                  strWord = Mid(strTempSort, n, 1)
'                  'A~Z,-
'                  If Not ((Asc(strWord) >= 65 And Asc(strWord) <= 90) Or _
'                          Asc(strWord) = 45) Then
'                     strKindEmp = "N"
'                     Exit For
'                  End If
'               Next n
'               If strKindEmp = "" Then
'                  For n = 1 To Len(strTempSort)
'                     strWord = Mid(strTempSort, n, 1)
'                     If Asc(strWord) >= 65 And Asc(strWord) <= 90 Then
'                        strSql = "select st17,st01,st02 from staff" & _
'                                 " where substr(st03,1,2)='P2' and st04='1' and ST17 is not null" & _
'                                 " and instr('(" & strWord & ")','('||st17||')') > 0" & _
'                                 " union select lk01,lk04,st02 from ipdeptkeyword,staff" & _
'                                 " where lk12='T' and lk11='甲' and lk04=st01(+)" & _
'                                 " and InStr('(" & strWord & ")',upper(rtrim(ltrim(lk01)))) > 0"
'                        intI = 1
'                        Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'                        If intI = 1 Then
'                           strKindEmp = strKindEmp & ";" & rsTmp.Fields("st01")
'                        End If
'                     End If
'                  Next n
'                  If strKindEmp <> "" Then
'                     m_Sender = m_Sender & strKindEmp
'                     If PUB_TM_ToSortOut = "" Then PUB_TM_ToSortOut = "3" '個人
'                     If InStr(strTi15, "(甲)") = 0 Then strTi15 = strTi15 & "(甲)"
'                  End If
'               End If
'            End If
'         End If
         '2020/12/16 END
         '檢查甲規則是否為MCTF人員,若是,則分類為1.MCTF案
         If m_Sender <> "" Then
            tmpArr = Split(m_Sender, ";")
            If UBound(tmpArr) >= 0 Then
               For j = 0 To UBound(tmpArr)
                  If tmpArr(j) <> "" Then
                     If InStr(Pub_GetSpecMan("MCTF", True), tmpArr(j)) > 0 Then
                        PUB_TM_ToSortOut = "1" 'MCTF
                     End If
                  End If
               Next j
            End If
         End If
      End If
      
      '乙規則(關鍵字):
      strKindEmp = ""
'      If QryIpdeptKeyWord("T", strSubject, strTi11, _
'                         strKindEmp, strTi15, " and lk11='B'") <> "" Then
      'Modify By Sindy 2020/3/6 將索引關鍵字改為共用函數
      If PUB_FindIPdeptKeyWord(" and lk12='T' and lk11='乙'", strSubject, strTi11, strKindEmp, strTemp) <> "" Then
      '2020/3/6 END
         PUB_TM_ToSortOut = "2" '大陸案
         If strKindEmp <> "" Then m_Sender = m_Sender & strKindEmp
         If strTemp <> "" Then strTi15 = IIf(strTi15 <> "", strTi15 & ",", "") & strTemp
      End If
   End If
   
'  2.主旨包含TOOOOO(本所案號)、註冊號、申請號,以案件最後承辦人為收件人
'   If PUB_TM_ToSortOut = "" Then
      '我方案號：T220277~79
      '第5812149號商標續展事宜
      '第5769222，5769267號商標續展事宜
      'V2影像檢測程式 "計算機軟件登記(本所案號:TC010948)"
      '(Our Ref：T-220680)
      ')商標，申請 (我方案號：T220851~52)等2件
      '商??展申?提交函-帕思比 PASPI-?方文?T-154347
      '申請第107058088,107058090,107058094,107058096號等4件商標商品名稱補正事
   '   'Modify By Sindy 2018/11/5 排除FCP-
   '   'FW: TW Patent No. I509237; Your Ref.: FCP-050125; Our Ref.: 2014-OPA-6486/TW_annuity fee
      '檢索是否為個案
      'Modify By Sindy 2021/9/29 + , , False : 正規方式檢查本所案號
      strText = strSubject
      If PUB_IPDeptGetCaseNo(strText, "YOURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , TM收件匣, , False) = False Then
         strText = strSubject
         If PUB_IPDeptGetCaseNo(strText, "OURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , TM收件匣, , False) = False Then
         End If
      End If
      'Add By Sindy 2021/9/29 再用其他方式檢查案號
      If (strCP01 = "" Or strCP02 = "") Then
         strText = strSubject
         If PUB_IPDeptGetCaseNo(strText, "YOURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , TM收件匣, , True) = False Then
            strText = strSubject
            If PUB_IPDeptGetCaseNo(strText, "OURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , TM收件匣, , True) = False Then
            End If
         End If
      End If
      '2021/9/29 END
      If strCP01 = "" And strCP02 = "" Then
         For j = 1 To 2
            If j = 1 Then
               '用註冊號檢索
               strSql = "select tm01,tm02,tm03,tm04,tm12,tm15,tm16,tm57" & _
                        " from trademark" & _
                        " Where Length(tm15)>=6" & _
                        " and tm01 not in('CFT')" & _
                        " and instr('" & ChgSQL(strSubject) & "',tm15)>0" & _
                        " and tm29 is null and tm57 is null" & _
                        " order by length(rtrim(ltrim(tm15))) desc,length(rtrim(ltrim(tm11))) desc,tm01||tm02||tm03||tm04 asc"
            Else
               '用申請案號檢索
               strSql = "select tm01,tm02,tm03,tm04,tm12,tm15,tm16,tm57" & _
                        " from trademark" & _
                        " Where Length(tm12)>=6" & _
                        " and tm01 not in('CFT')" & _
                        " and instr('" & ChgSQL(strSubject) & "',tm12)>0" & _
                        " and tm29 is null and tm57 is null" & _
                        " order by length(rtrim(ltrim(tm12))) desc,length(rtrim(ltrim(tm11))) desc,tm01||tm02||tm03||tm04 asc"
            End If
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strCP01 = rsTmp.Fields("tm01")
               strCP02 = rsTmp.Fields("tm02")
               strCP03 = rsTmp.Fields("tm03")
               strCP04 = rsTmp.Fields("tm04")
               '多案
               If rsTmp.RecordCount > 1 Then
                  'Add By Sindy 2021/3/25 FW: Fw:1572233 Rule 18 bis (1)(a) Australia
                  '嘉雯:主旨中之審定號數:1572233，為TF-000820
                  '     分信規則無法識別,即判斷應分給承辦人: 桂紹楨
                  bolChkCaseNo = False
                  If strCP01 = "TF" Then
                     rsTmp.MoveFirst
                     Do While Not rsTmp.EOF
                        '流水號前5碼相同者為同案
                        If strCP01 = rsTmp.Fields("tm01") And Left(strCP02, 5) = Left(rsTmp.Fields("tm02"), 5) Then
                           bolChkCaseNo = True
                        Else
                           bolChkCaseNo = False
                           Exit Do
                        End If
                        rsTmp.MoveNext
                     Loop
                  End If
                  If bolChkCaseNo = False Then
                  '2021/3/25 END
                     rsTmp.MoveFirst
                     Do While Not rsTmp.EOF
                        '檢查各案件最後承辦人是否為同一人
                        If QryIsTMCaseEmp(rsTmp.Fields("tm01"), rsTmp.Fields("tm02"), rsTmp.Fields("tm03"), rsTmp.Fields("tm04"), PUB_TM_ToSortOut, strTemp, "") = False Then
                           PUB_TM_ToSortOut = ""
                           Exit Do
                        Else
                           If strTemp = "" Then
                              PUB_TM_ToSortOut = ""
                              Exit Do
                           ElseIf strTemp1 <> "" And strTemp1 <> strTemp Then
                              PUB_TM_ToSortOut = ""
                              Exit Do
                           End If
                           strTemp1 = strTemp
                        End If
                        rsTmp.MoveNext
                     Loop
                     If PUB_TM_ToSortOut = "" Then
                        strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
                     End If
                  End If
               End If
            End If
            'Add By Sindy 2021/3/25
            If strCP01 <> "" And strCP02 <> "" Then
               Exit For
            End If
            '2021/3/25 END
         Next j
      End If
      
      'Add by Sindy 2020/8/24 主旨裡面有期限3時,未解析到本所案號再加入T-,T解析案號
      'If (strCP01 = "" Or strCP02 = "") And InStr(strSubject, "期限3") > 0 Then
      If (strCP01 = "" Or strCP02 = "") Then
         'Modify By Sindy 2021/9/29 + , , False : 正規方式檢查本所案號
         strText = strSubject
         If PUB_IPDeptGetCaseNo(strText, "T-", strCP01, strCP02, strCP03, strCP04, strPA150, True, TM收件匣, , False) = False Then
            strText = strSubject
            If PUB_IPDeptGetCaseNo(strText, "T", strCP01, strCP02, strCP03, strCP04, strPA150, True, TM收件匣, , False) = False Then
            End If
         End If
         'Add By Sindy 2021/9/29 再用其他方式檢查案號
         If (strCP01 = "" Or strCP02 = "") Then
            strText = strSubject
            If PUB_IPDeptGetCaseNo(strText, "T-", strCP01, strCP02, strCP03, strCP04, strPA150, , TM收件匣, , True) = False Then
               strText = strSubject
               If PUB_IPDeptGetCaseNo(strText, "T", strCP01, strCP02, strCP03, strCP04, strPA150, , TM收件匣, , True) = False Then
               End If
            End If
         End If
         '2021/9/29 END
      End If
      '2020/8/24 END
      
      If strCP01 <> "" And strCP02 <> "" Then
         '個案
         If QryIsTMCaseEmp(strCP01, strCP02, strCP03, strCP04, PUB_TM_ToSortOut, "", m_Sender) = False Then
            strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
'         Else
'            m_Sender = m_Sender & ";" & strTemp
'            '乙規則的信件,同時加發案件最後承辦人為收件人
'            If InStr(strTi15, "(乙)") > 0 Then
'               '以案件最後承辦人為收件人
'               strSql = "SELECT cp09,cp14,st04,A0908 From caseprogress,staff,acc090" & _
'                        " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
'                        " and cp159='0'" & _
'                        " and cp14=st01(+) and st15=A0901(+) and substr(cp09,1,1)<>'D'" & _
'                        " order by cp05 desc,cp67 desc"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'               If intI = 1 Then
'                  If RsTemp.Fields("st04") = "1" And RsTemp.Fields("cp14") <> "" Then
'                     m_Sender = m_Sender & ";" & RsTemp.Fields("cp14") '承辦人
'                  ElseIf RsTemp.Fields("A0908") <> "" Then
'                     m_Sender = m_Sender & ";" & RsTemp.Fields("A0908") '承辦人主管
'                  End If
'               End If
'            End If
         End If
         
         'Modify By Sindy 2024/9/4 嘉雯:因TF案件的程序採輪流負責,現亦已針對TF案件設定關鍵字分信規則
         '                              取消固定加桂英的機制
'         'Add By Sindy 2021/5/12
'         '嘉雯:TF案,除分給承辦人外,增列79041林桂英主任
'         If strCP01 = "TF" Then
'            m_Sender = m_Sender & ";79041"
'         End If
'         '2021/5/12 END
      End If
      
'   '已分類到,但欲抓取本所案號
'   ElseIf strCP01 = "" And strCP02 = "" Then
'      strText = strSubject
'      If PUB_IPDeptGetCaseNo(strText, "YOURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , TM收件匣) = False Then
'         strText = strSubject
'         If PUB_IPDeptGetCaseNo(strText, "OURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , TM收件匣) = False Then
'         End If
'      End If
'      If strCP01 = "" And strCP02 = "" Then
'         For j = 1 To 2
'            If j = 1 Then
'               '用註冊號檢索
'               strSql = "select tm01,tm02,tm03,tm04,tm12,tm15,tm16,tm57" & _
'                        " from trademark" & _
'                        " Where Length(tm15)>=6" & _
'                        " and tm01 not in('CFT')" & _
'                        " and instr('" & ChgSQL(strSubject) & "',tm15)>0" & _
'                        " order by tm01,tm02,tm03,tm04"
'            Else
'               '用申請案號檢索
'               strSql = "select tm01,tm02,tm03,tm04,tm12,tm15,tm16,tm57" & _
'                        " from trademark" & _
'                        " Where Length(tm12)>=6" & _
'                        " and tm01 not in('CFT')" & _
'                        " and instr('" & ChgSQL(strSubject) & "',tm12)>0" & _
'                        " order by tm01,tm02,tm03,tm04"
'            End If
'            intI = 1
'            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
'            If intI = 1 Then
'               strCP01 = rsTmp.Fields("tm01")
'               strCP02 = rsTmp.Fields("tm02")
'               strCP03 = rsTmp.Fields("tm03")
'               strCP04 = rsTmp.Fields("tm04")
'               '多案
'               If rsTmp.RecordCount > 1 Then
'                  rsTmp.MoveFirst
'                  Do While Not rsTmp.EOF
'                     '檢查各案件最後承辦人是否為同一人
'                     If QryIsTMCaseEmp(rsTmp.Fields("tm01"), rsTmp.Fields("tm02"), rsTmp.Fields("tm03"), rsTmp.Fields("tm04"), strTempSort, strTemp) = False Then
'                        strTempSort = ""
'                        Exit Do
'                     Else
'                        If strTemp = "" Then
'                           strTempSort = ""
'                           Exit Do
'                        ElseIf strTemp1 <> "" And strTemp1 <> strTemp Then
'                           strTempSort = ""
'                           Exit Do
'                        End If
'                        strTemp1 = strTemp
'                     End If
'                     rsTmp.MoveNext
'                  Loop
'                  If strTempSort = "" Then
'                     strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
'                  End If
'               End If
'            End If
'            If strTempSort <> "" Or (strCP01 <> "" And strCP02 <> "") Then Exit For
'         Next j
'      End If
'      If strCP01 <> "" And strCP02 <> "" Then
'         '個案
'         If QryIsTMCaseEmp(strCP01, strCP02, strCP03, strCP04, "", "") = False Then
'            strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
'         Else
'            'Add By Sindy 2020/6/25 乙規則的信件,同時加發案件最後承辦人為收件人
'            If InStr(strTi15, "(乙)") > 0 Then
'               '以案件最後承辦人為收件人
'               strSql = "SELECT cp09,cp14,st04,A0908 From caseprogress,staff,acc090" & _
'                        " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
'                        " and cp159='0'" & _
'                        " and cp14=st01(+) and st15=A0901(+) and substr(cp09,1,1)<>'D'" & _
'                        " order by cp05 desc,cp67 desc"
'               intI = 1
'               Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'               If intI = 1 Then
'                  If RsTemp.Fields("st04") = "1" And RsTemp.Fields("cp14") <> "" Then
'                     m_Sender = m_Sender & ";" & RsTemp.Fields("cp14") '承辦人
'                  ElseIf RsTemp.Fields("A0908") <> "" Then
'                     m_Sender = m_Sender & ";" & RsTemp.Fields("A0908") '承辦人主管
'                  End If
'               End If
'            End If
'            '2020/6/25 END
'         End If
'      End If
'   End If
   
   '********************************************************************************************************
   '其他關鍵字索引
   '********************************************************************************************************
   If PUB_TM_ToSortOut = "" Then
'      PUB_TM_ToSortOut = QryIpdeptKeyWord("T", strSubject, strTi11, _
'                         m_Sender, strTi15, " and (lk11 not in('A','B','C') or lk11 is null)")
      'Modify By Sindy 2020/3/6 將索引關鍵字改為共用函數
      PUB_TM_ToSortOut = PUB_FindIPdeptKeyWord(" and lk12='T' and (lk11 not in('甲','乙','丙') or lk11 is null)", strSubject, strTi11, m_Sender, strTi15)
      '2020/3/6 END
   End If
   
'   '********************************************************************************************************
'   '比對簡體字檔案
'   '********************************************************************************************************
'   Dim bolFind As Boolean
'   If PUB_TM_ToSortOut = "" Or InStr(strTi15, "(甲)") > 0 Then
'      oForm.GRD2.Clear: intGrd2 = 0: strOldSender = m_Sender
'      If oForm.GRD2.Rows > 0 Then
'         For j = oForm.GRD2.Rows - 1 To 1 Step -1
'            oForm.GRD2.RemoveItem j
'         Next j
'      End If
'      bolFind = False
'      If oForm.TextBoxT <> "" Then
'         For j = 1 To Len(oForm.TextBoxT)
'            LongSelLength = InStr(Mid(oForm.TextBoxT, j), vbCrLf)
'            If LongSelLength = 0 Then LongSelLength = Len(oForm.TextBoxT)
'            oForm.TextBox3 = Replace(Mid(oForm.TextBoxT.Text, j, LongSelLength - 1), Chr(10), "")
'            If j > 1 Then '第一筆是標題,跳過
'               If oForm.TextBox3 <> "" Then
'                  '關鍵字簡體#分類#收受者#索引排序
'                  tmpArr = Split(oForm.TextBox3, "#")
'                  If UBound(tmpArr) = 3 Then
'                     strExc(10) = oForm.TextBox3
'                     oForm.TextBox3 = tmpArr(0)
'                     If InStr(oForm.TextII17, oForm.TextBox3) > 0 And bolFind = False Then
'                        If InStr(strTi15, "(甲)") > 0 Then
'                           If InStr(tmpArr(1), "乙") > 0 Then
'                              bolFind = True '比對到資料
'                              PUB_TM_ToSortOut = "2" '大陸案
'                           End If
'                        Else 'If PUB_TM_ToSortOut = "" Then '記錄第一次比對到的值
'                           bolFind = True '比對到資料
'                           PUB_TM_ToSortOut = Left(Trim(tmpArr(1)), 1)
'                        End If
'                        If bolFind = True Then
'                           If InStr(strTi15, "(簡)") = 0 Then
'                              strTi15 = IIf(strTi15 <> "", strTi15 & ",", "") & "(簡)"
'                           End If
'                           strTemp = Pub_GetSpecMan(CStr(Trim(tmpArr(2))))
'                           If strTemp <> "" Then
'                              m_Sender = strOldSender & ";" & strTemp
'                           Else
'                              m_Sender = strOldSender & ";" & Trim(tmpArr(2))
'                           End If
'                        End If
'                     End If
'                     If Val(tmpArr(3)) > 0 Then '有排序欄位值的才記錄下來
'                        If Not (intGrd2 = 0 And oForm.GRD2.Rows = 1) Then
'                           oForm.GRD2.AddItem ""
'                        End If
'                        oForm.GRD2.TextMatrix(intGrd2, 0) = strExc(10)
'                        oForm.GRD2.TextMatrix(intGrd2, 1) = Val(tmpArr(3))
'                        intGrd2 = intGrd2 + 1
'                     End If
'                  End If
'               End If
'            End If
'            j = j + LongSelLength
'         Next j
'         '有排序欄位時,取得優先順序的第一筆資料
'         If intGrd2 > 0 Then
'            oForm.GRD2.col = 1
'            oForm.GRD2.row = 0
'            oForm.GRD2.Sort = 3 '3.數值昇冪 4.數值降冪
'            For j = 0 To intGrd2 - 1
'               tmpArr = Split(oForm.GRD2.TextMatrix(j, 0), "#")
'               '關鍵字簡體#分類#收受者#索引排序
'               If UBound(tmpArr) = 3 Then
'                  oForm.TextBox3 = tmpArr(0)
'                  If InStr(oForm.TextII17, oForm.TextBox3) > 0 Then
'                     If bolFind = False Then
'                        If InStr(strTi15, "(甲)") > 0 Then
'                           If InStr(tmpArr(1), "乙") > 0 Then
'                              bolFind = True '比對到資料
'                              PUB_TM_ToSortOut = "2" '大陸案
'                           End If
'                        Else 'If PUB_TM_ToSortOut = "" Then '記錄第一次比對到的值
'                           bolFind = True '比對到資料
'                           PUB_TM_ToSortOut = Left(Trim(tmpArr(1)), 1)
'                        End If
'                        If bolFind = True Then
'                           If InStr(strTi15, "(簡)") = 0 Then
'                              strTi15 = IIf(strTi15 <> "", strTi15 & ",", "") & "(簡)"
'                           End If
'                           strTemp = Pub_GetSpecMan(CStr(Trim(tmpArr(2))))
'                           If strTemp <> "" Then
'                              m_Sender = strOldSender & ";" & strTemp
'                           Else
'                              m_Sender = strOldSender & ";" & Trim(tmpArr(2))
'                           End If
'                        End If
'                     End If
'                  End If
'               End If
'            Next j
'         End If
'
'      End If
'   End If
   
   '3.若主旨有「申請核准」，同時以甲組及以下網域地址之特定收件人為收件人
'  4.主旨未包含上述關鍵字時,依負責代理人區分收件人
'   If PUB_TM_ToSortOut = "" Or _
'      (InStr(strSubject, "申請核准") > 0 And InStr(strTi15, "(甲)") > 0) Then
   If PUB_TM_ToSortOut = "" Or _
      (InStr(strSubject, "申請核准") > 0 And InStr(strTi15, "(甲)") > 0) Then
      '寄件者:
      If InStr(strTi11, "@") > 0 Then
         'Modify By Sindy 2023/7/21
'         tmpArr = Split(strTi11, "@")
'         strMail = "@" & tmpArr(UBound(tmpArr))
         If InStr(strTi11, " [") > 0 Then
            tmpArr = Split(strTi11, " [")
            strMail = Left(tmpArr(1), Len(tmpArr(1)) - 1)
         Else
            strMail = strTi11
         End If
         tmpArr = Split(strMail, "@")
         strDomain = "@" & tmpArr(UBound(tmpArr))
         '2023/7/21 END
         'E-Mail:
         strSql = "SELECT fa01,fa02,fa120 From fagent" & _
                  " where (" & _
                      "upper(fa16)='" & UCase(ChgSQL(Trim(strMail))) & "'" & _
                  " or upper(fa79)='" & UCase(ChgSQL(Trim(strMail))) & "'" & _
                  " or upper(fa105)='" & UCase(ChgSQL(Trim(strMail))) & "'" & _
                  " or upper(fa80)='" & UCase(ChgSQL(Trim(strMail))) & "'" & _
                  " or upper(fa81)='" & UCase(ChgSQL(Trim(strMail))) & "'" & _
                  " or upper(fa82)='" & UCase(ChgSQL(Trim(strMail))) & "'" & _
                  ")" & _
                  " and fa120 is not null" & _
                  " and fa02='0'" & _
                  " order by fa01 asc"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            'If rsTmp.RecordCount = 1 Then
               PUB_TM_ToSortOut = "3" '個人
               strTi15 = IIf(strTi15 <> "", strTi15 & ",", "") & strMail
               m_Sender = m_Sender & ";" & rsTmp.Fields("fa120")
               If InStr(rsTmp.Fields("fa120"), "MCTF") > 0 Then
                  PUB_TM_ToSortOut = "1" 'MCTF
               End If
            'End If
         Else
            '網域:
            'Add By Sindy 2021/1/11 TM分信規則,請排除@gmail.com和@qq.com之網域判斷
            'Modify By Sindy 2021/5/20 + 排除 @126.com
            'Modify By Sindy 2023/7/21 + 排除 @yahoo.com 及 @yahoo.com.tw
            'Modify By Sindy 2024/7/10 + 排除 @taie.com.tw
            'Modify By Sindy 2024/7/18 + 排除 @msa.hinet.net
            If UCase(strDomain) <> UCase("@gmail.com") And _
               UCase(strDomain) <> UCase("@qq.com") And _
               UCase(strDomain) <> UCase("@126.com") And _
               UCase(strDomain) <> UCase("@yahoo.com.tw") And _
               InStr(UCase(strDomain), UCase("@yahoo.com")) = 0 And _
               UCase(strDomain) <> UCase("@taie.com.tw") And _
               UCase(strDomain) <> UCase("@msa.hinet.net") Then
            '2021/1/11 END
               strSql = "SELECT fa01,fa02,fa120 From fagent" & _
                        " where (" & _
                            "instr(upper(fa16),'" & UCase(ChgSQL(Trim(strDomain))) & "')> 0" & _
                        " or instr(upper(fa79),'" & UCase(ChgSQL(Trim(strDomain))) & "')> 0" & _
                        " or instr(upper(fa105),'" & UCase(ChgSQL(Trim(strDomain))) & "')> 0" & _
                        " or instr(upper(fa80),'" & UCase(ChgSQL(Trim(strDomain))) & "')> 0" & _
                        " or instr(upper(fa81),'" & UCase(ChgSQL(Trim(strDomain))) & "')> 0" & _
                        " or instr(upper(fa82),'" & UCase(ChgSQL(Trim(strDomain))) & "')> 0" & _
                        ")" & _
                        " and fa120 is not null" & _
                        " and fa02='0'" & _
                        " order by fa01 asc"
               intI = 1
               Set rsTmp = ClsLawReadRstMsg(intI, strSql)
               If intI = 1 Then
                  'If rsTmp.RecordCount = 1 Then
                     PUB_TM_ToSortOut = "3" '個人
                     strTi15 = IIf(strTi15 <> "", strTi15 & ",", "") & strDomain
                     m_Sender = m_Sender & ";" & rsTmp.Fields("fa120")
                     If InStr(rsTmp.Fields("fa120"), "MCTF") > 0 Then
                        PUB_TM_ToSortOut = "1" 'MCTF
                     End If
                  'End If
               End If
            End If
         End If
      End If
   End If
   
   '條件檢查完畢:
ChkEnd:
   '無以上條件者分至其他；
'   If PUB_TM_ToSortOut = "甲" Or _
'      PUB_TM_ToSortOut = "乙" Or _
'      PUB_TM_ToSortOut = "丙" Then PUB_TM_ToSortOut = "3" '個人
   If m_Sender = "" Then PUB_TM_ToSortOut = "" 'Add By Sindy 2023/7/21
   If PUB_TM_ToSortOut = "" Then PUB_TM_ToSortOut = "5" '其他
   
   'Add by Sindy 2019/8/15
   If Trim(strTi11) = "未傳遞的主旨" Then
      PUB_TM_ToSortOut = "5" '其他
      'm_Sender = ""
   End If
   
   'Modify By Sindy 2020/8/24 整合在 PUB_TM_ToSortOut_sub 中
   Call PUB_TM_ToSortOut_sub(m_Sender)
   
   Set rsTmp = Nothing
End Function

'Add By Sindy 2020/8/24 解析TM分信收受者
'bolCompEmp : 純比對需要再帶出那些人員
Public Sub PUB_TM_ToSortOut_sub(ByRef m_Sender As String, Optional bolCompEmp As Boolean = False)
Dim strText As String
Dim tmpArr As Variant
Dim j As Integer
Dim strEmp As String, strChkMCTFEmp As String, strMCTFEmp As String
Dim strMCTFMan As String, varMCTFMan As Variant
   
'*****************************************************************************
'檢查需要再帶出那些人員
'*****************************************************************************
   'Modify By Sindy 2021/11/24 收受者為MCTF人員時，檢查是否須增加收信人員
   strSql = "select substr(ocode,1,6) as ocode_id,ocode,oman from setspecman" & _
            " where substr(ocode,1,4)='MCTF' and instr(ocode,'收信人員')>0 and instr('" & m_Sender & "',substr(ocode,1,6))>0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If InStr(m_Sender, RsTemp.Fields("ocode_id")) > 0 Then
            If m_Sender <> "" Then m_Sender = m_Sender & ";"
            m_Sender = m_Sender & RsTemp.Fields("oman")
         End If
         RsTemp.MoveNext
      Loop
   End If
'   'Modify By Sindy 2021/9/1 修改抓特殊人員設定
'   '依寄件者或網域判斷為MCTF04的信件時，同時寄 96003.Monica
'   'If InStr(m_Sender, "MCTF04") > 0 And InStr(m_Sender, "96003") = 0 Then
'   If InStr(m_Sender, "MCTF04") > 0 Then
'      If m_Sender <> "" Then m_Sender = m_Sender & ";"
'      m_Sender = m_Sender & Pub_GetSpecMan("MCTF04收信人員") '";96003"
'   End If
'   '依寄件者或網域判斷為MCTF05的信件時，同時寄 A4009.Cary
'   'If InStr(m_Sender, "MCTF05") > 0 And InStr(m_Sender, "A4009") = 0 Then
'   If InStr(m_Sender, "MCTF05") > 0 Then
'      If m_Sender <> "" Then m_Sender = m_Sender & ";"
'      m_Sender = m_Sender & Pub_GetSpecMan("MCTF05收信人員") '";A4009"
'   End If
   
   '智權人員為96030.巨京商標時, 屬MCTF05業務負責
   If InStr(m_Sender, "96030") > 0 Then
      m_Sender = Replace(m_Sender, "96030", "MCTF05")
   End If
   
'純比對需要再帶出那些人員,離開回傳
If bolCompEmp = True Then Exit Sub
   
   'Add By Sindy 2020/11/4
   If m_Sender <> "" Then
      tmpArr = Split(m_Sender, ";")
      For j = 0 To UBound(tmpArr)
         If tmpArr(j) <> "" Then
            If InStr(OL_TmMail需排除的收受者, tmpArr(j)) > 0 Then
               'Modify By Sindy 2021/11/24
               'm_Sender = ""
               m_Sender = Replace(m_Sender, tmpArr(j), "")
               '2021/11/24 END
               'Exit Function
            End If
         End If
      Next j
   End If
   '2020/11/4 END
   
'*****************************************************************************
'其他
'*****************************************************************************
   'Modify By Sindy 2021/9/1 天雲和筱凌仍互為職代(自110年09月06日(星期一)修正執行):走職代設定檔
'   If strSrvDate(1) < 20210906 Then
'   '2021/9/1 END
'      '收受者有寄 96003.Monica，Monica休假時，則寄Cary
'      If InStr(m_Sender, "96003") > 0 Then
'         If CheckIsPersonRest("96003", strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = True Then
'            m_Sender = m_Sender & ";A4009"
'         End If
'      '收受者有寄 A4009.Cary，Cary休假時，則寄Monica
'      ElseIf InStr(m_Sender, "A4009") > 0 Then
'         If CheckIsPersonRest("A4009", strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)) = True Then
'            m_Sender = m_Sender & ";96003"
'         End If
'      End If
'   End If
   
   strMCTFMan = Pub_GetSpecMan("MCTMember")
   varMCTFMan = Split(strMCTFMan, ";")
   For intI = 0 To UBound(varMCTFMan)
      If Left(varMCTFMan(intI), 4) = "MCTF" Then
         If InStr(m_Sender, varMCTFMan(intI)) > 0 Then
            m_Sender = Replace(m_Sender, varMCTFMan(intI), Pub_GetSpecMan(CStr(varMCTFMan(intI))))
         End If
      End If
   Next intI
'   If InStr(m_Sender, "MCTF01") > 0 Then m_Sender = Replace(m_Sender, "MCTF01", Pub_GetSpecMan("MCTF01"))
'   If InStr(m_Sender, "MCTF02") > 0 Then m_Sender = Replace(m_Sender, "MCTF02", Pub_GetSpecMan("MCTF02"))
'   If InStr(m_Sender, "MCTF03") > 0 Then m_Sender = Replace(m_Sender, "MCTF03", Pub_GetSpecMan("MCTF03"))
'   If InStr(m_Sender, "MCTF04") > 0 Then m_Sender = Replace(m_Sender, "MCTF04", Pub_GetSpecMan("MCTF04"))
'   If InStr(m_Sender, "MCTF05") > 0 Then m_Sender = Replace(m_Sender, "MCTF05", Pub_GetSpecMan("MCTF05"))
   m_Sender = Replace(m_Sender, ",", ";")
   If Left(m_Sender, 1) = ";" Then m_Sender = Mid(m_Sender, 2)
   
   'Add By Sindy 2021/10/1 收受者裡若有MCTF第一位負責人時,
   '                       增加檢查該負責人是否有休假, 若有, 加抓職代
   strMCTFEmp = ""
   strSql = "select ocode,oman from setspecman where substr(ocode,1,4)='MCTF' and length(ocode)=6 and oman is not null"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         If InStr(strMCTFEmp, RsTemp.Fields("oman")) = 0 Then
            strMCTFEmp = strMCTFEmp & ";" & RsTemp.Fields("oman") '列出MCTF負責人
         End If
         RsTemp.MoveNext
      Loop
   End If
'   For j = 1 To 5
'      If j = 1 Then strEmp = Left(Pub_GetSpecMan("MCTF01"), 5)
'      If j = 2 Then strEmp = Left(Pub_GetSpecMan("MCTF02"), 5)
'      If j = 3 Then strEmp = Left(Pub_GetSpecMan("MCTF03"), 5)
'      If j = 4 Then strEmp = Left(Pub_GetSpecMan("MCTF04"), 5)
'      If j = 5 Then strEmp = Left(Pub_GetSpecMan("MCTF05"), 5)
'      If InStr(strMCTFEmp, strEmp) = 0 And strEmp <> "" Then
'         strMCTFEmp = strMCTFEmp & ";" & strEmp '列出MCTF負責人
'      End If
'   Next j
   strChkMCTFEmp = ""
   strMCTFMan = Pub_GetSpecMan("MCTMember")
   varMCTFMan = Split(strMCTFMan, ";")
   For intI = 0 To UBound(varMCTFMan)
      If Left(varMCTFMan(intI), 4) = "MCTF" Then
         strEmp = Left(Pub_GetSpecMan(CStr(varMCTFMan(intI))), 5)
         If InStr(m_Sender, strEmp) > 0 And InStr(strChkMCTFEmp, strEmp) = 0 And strEmp <> "" Then
            strChkMCTFEmp = strChkMCTFEmp & ";" & strEmp '記錄不需要再檢查的人員
            '職代走一般規則(先檢查案件職代再檢查人事職代)
            strText = GetCaseDutyAgent(strEmp, "", False)
            If strText <> "" And InStr(strMCTFEmp, strText) > 0 Then
               If GetCaseDutyAgent(strText, "", False) = "" Then
                  If InStr(m_Sender, strText) = 0 Then
                     m_Sender = m_Sender & ";" & strText
                  End If
                  strChkMCTFEmp = strChkMCTFEmp & ";" & strText '記錄不需要再檢查的人員
               End If
            End If
         End If
      End If
   Next intI
'   For j = 1 To 5
'      If j = 1 Then strEmp = Left(Pub_GetSpecMan("MCTF01"), 5)
'      If j = 2 Then strEmp = Left(Pub_GetSpecMan("MCTF02"), 5)
'      If j = 3 Then strEmp = Left(Pub_GetSpecMan("MCTF03"), 5)
'      If j = 4 Then strEmp = Left(Pub_GetSpecMan("MCTF04"), 5)
'      If j = 5 Then strEmp = Left(Pub_GetSpecMan("MCTF05"), 5)
'      If InStr(m_Sender, strEmp) > 0 And InStr(strChkMCTFEmp, strEmp) = 0 And strEmp <> "" Then
'         strChkMCTFEmp = strChkMCTFEmp & ";" & strEmp '記錄不需要再檢查的人員
'         '職代走一般規則(先檢查案件職代再檢查人事職代)
'         strText = GetCaseDutyAgent(strEmp, "", False)
'         If strText <> "" And InStr(strMCTFEmp, strText) > 0 Then
'            If GetCaseDutyAgent(strText, "", False) = "" Then
'               If InStr(m_Sender, strText) = 0 Then
'                  m_Sender = m_Sender & ";" & strText
'               End If
'               strChkMCTFEmp = strChkMCTFEmp & ";" & strText '記錄不需要再檢查的人員
'            End If
'         End If
'      End If
'   Next j
   '2021/10/1 END
   
   '過濾是否有收受者重覆的資料
   If m_Sender <> "" And InStr(m_Sender, ";") > 0 Then
      strText = m_Sender
      tmpArr = Split(strText, ";")
      m_Sender = ""
      For j = 0 To UBound(tmpArr)
         If tmpArr(j) <> "" Then
            If InStr(m_Sender, tmpArr(j)) = 0 Then
               m_Sender = m_Sender & IIf(m_Sender = "", "", ";") & tmpArr(j)
            End If
         End If
      Next j
   End If
End Sub

'Add By Sindy 2019/4/19
Function QryIsTMCaseEmp(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, _
   ByRef TM_ToSortOut As String, ByRef strEmp As String, ByRef m_Sender As String) As Boolean
Dim tm() As String, sp() As String
Dim salesNo As String, salesArea As String
Dim strCP13 As String
'Dim tmpArr As Variant
'Dim j As Integer
   
   QryIsTMCaseEmp = False:  strEmp = ""
   If strCP01 <> "" And strCP02 <> "" And Left(strCP01, 1) = "T" Then
      ReDim tm(1 To TF_TM) As String
      ReDim sp(1 To tf_SP) As String
      '讀取基本檔資料
      If PUB_ReadTradeMarkData(tm(), strCP01, strCP02, strCP03, strCP04) = False Then
         If PUB_ReadServicePracticeData(sp(), strCP01, strCP02, strCP03, strCP04) = False Then
            Exit Function
         Else
            salesArea = GetCuSales(ChangeCustomerL(sp(8)), salesNo)
            If sp(9) = "020" Then
               TM_ToSortOut = "2" '大陸案
            Else
               TM_ToSortOut = "4" '非大陸案
            End If
         End If
      Else
         salesArea = GetCuSales(ChangeCustomerL(tm(23)), salesNo)
         If tm(10) = "020" Then
            TM_ToSortOut = "2" '大陸案
         Else
            TM_ToSortOut = "4" '非大陸案
         End If
      End If
      
      strCP13 = PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04) '目前智權人員
      
      If Mid(strCP13, 1, 4) = "MCTF" Or _
         salesNo = "96029" Or salesNo = "96030" Then
         TM_ToSortOut = "1" 'MCTF案
      End If
      QryIsTMCaseEmp = True
      
      '讀取案件最後承辦人
      'Modify By Sindy 2020/8/26 承辦人排除掛程序人員
      strSql = "SELECT cp09,cp14,st04,A0908 From caseprogress,staff,acc090" & _
               " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
               " and cp159='0'" & _
               " and cp14=st01(+) and st15=A0901(+) and substr(cp09,1,1)<>'D'" & _
               " and st03<>'P22'" & _
               " order by cp05 desc,cp67 desc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         If RsTemp.Fields("st04") = "1" And RsTemp.Fields("cp14") <> "" Then
            strEmp = RsTemp.Fields("cp14") '承辦人
         'Modify By Sindy 2020/12/3 嘉雯:離職人員不要帶主管,放空 ex:TC-010755 RE: ?以此?准！！！：RE: 回复：(Y)詢問：著作權登記變更( Our Ref:TC010755、56、58、59)
'         ElseIf RsTemp.Fields("A0908") <> "" Then
'            strEmp = RsTemp.Fields("A0908") '承辦人主管
         End If
         '大陸案加發案件最後承辦人
         If TM_ToSortOut = "2" And InStr(m_Sender, strEmp) = 0 Then '大陸案
            m_Sender = m_Sender & ";" & strEmp
         End If
      End If
      
      'MCTF案件,目前收受者為程序人員時,同時加發智權人員
      If TM_ToSortOut = "1" Then 'MCTF案
         strSql = "select st17,st01,st02,st03 from staff where instr('" & m_Sender & "',st01)>0 and st03='P22'"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If InStr(m_Sender, strCP13) = 0 Then
               m_Sender = m_Sender & ";" & strCP13
            End If
         End If
      End If
      
      '若有找到案件,無抓到收受者時:
      If m_Sender = "" Then
         'MCTF歸目前智權人員
         If TM_ToSortOut = "1" And strCP13 <> "" Then
            m_Sender = m_Sender & ";" & strCP13
         ElseIf TM_ToSortOut <> "" And TM_ToSortOut <> "1" And strEmp <> "" Then
            '其他歸承辦人
            m_Sender = m_Sender & ";" & strEmp
         End If
      End If
   End If
End Function

''Add By Sindy 2019/4/19
'Function QryIpdeptKeyWord(ByVal strSys As String, ByVal strSubject As String, ByVal strTi11 As String, _
'   ByRef m_Sender As String, ByRef strTi15 As String, Optional ByVal strConSql As String) As String
'Dim strTemp As String, bolChkOk As Boolean, strWord As Stream
'Dim tmpArr As Variant, j As Integer
'
'   strSql = "select lk01,lk02,lk03,lk04,lk13,lk14 from ipdeptkeyword" & _
'            " where lk12='" & strSys & "' and lk14 is null and lk03='1'" & strConSql & _
'            " and InStr('" & UCase(ChgSQL(strSubject)) & "',upper(rtrim(ltrim(lk01)))) > 0" & _
'            " union select ' '||rtrim(ltrim(lk01))||' ' lk01,lk02,lk03,lk04,LK13,lk14 from ipdeptkeyword" & _
'            " where lk12='" & strSys & "' and lk14='Y' and lk03='1'" & strConSql & _
'            " and InStr('" & UCase(ChgSQL(strSubject)) & "',upper(rtrim(ltrim(lk01)))) > 0" & _
'            " union select lk01,lk02,lk03,lk04,lk13,lk14 from ipdeptkeyword" & _
'            " where lk12='" & strSys & "' and lk14 is null and lk03='2'" & strConSql & _
'            " and InStr('" & UCase(ChgSQL(strTi11)) & "',upper(rtrim(ltrim(lk01)))) > 0" & _
'            " union select ' '||rtrim(ltrim(lk01))||' ' lk01,lk02,lk03,lk04,LK13,lk14 from ipdeptkeyword" & _
'            " where lk12='" & strSys & "' and lk14='Y' and lk03='2'" & strConSql & _
'            " and InStr('" & UCase(ChgSQL(strTi11)) & "',upper(rtrim(ltrim(lk01)))) > 0" & _
'            " order by lk13 asc,lk01 asc"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
'   If intI = 1 Then
'      With RsTemp
'         RsTemp.MoveFirst
'         Do While RsTemp.EOF = False
'               If RsTemp.Fields("LK03") = "1" Then '主旨
'                  strTemp = Trim(UCase(strSubject))
'               Else
'                  strTemp = Trim(UCase(strTi11))
'               End If
'               If "" & RsTemp.Fields("lk14") = "Y" Then
'                  bolChkOk = False
'                  '索引最前面
'                  strWord = Trim(UCase(RsTemp.Fields("LK01"))) & " "
'                  If Left(strTemp, Len(strWord)) = strWord Then
'                     bolChkOk = True
'                  End If
'                  If bolChkOk = False Then
'                     '索引最後面
'                     strWord = " " & Trim(UCase(RsTemp.Fields("LK01")))
'                     If Right(strTemp, Len(strWord)) = strWord Then
'                        bolChkOk = True
'                     End If
'                  End If
'                  'Add By Sindy 2018/6/27
'                  If bolChkOk = False Then
'                     '索引中間
'                     strWord = " " & Trim(UCase(RsTemp.Fields("LK01"))) & " "
'                     If InStr(strTemp, strWord) > 0 Then
'                        bolChkOk = True
'                     End If
'                  End If
'                  '2018/6/27 END
'                  '有索引到
'                  If bolChkOk = True Then
'                     strTi15 = strTi15 & ";" & RsTemp.Fields("LK01")
'                     QryIpdeptKeyWord = RsTemp.Fields("LK02")
'                     If "" & RsTemp.Fields("LK04") <> "" Then '有收受者
'                        tmpArr = Split(RsTemp.Fields("LK04"), ";")
'                        For j = 0 To UBound(tmpArr)
'                           strTemp = Pub_GetSpecMan(CStr(tmpArr(j)))
'                           If strTemp <> "" Then
'                              m_Sender = m_Sender & ";" & strTemp
'                           Else
'                              m_Sender = m_Sender & ";" & tmpArr(j)
'                           End If
'                        Next j
'                        Exit Do
'                     End If
'                  End If
'               End If
'            If QryIpdeptKeyWord <> "" Then Exit Do
'            RsTemp.MoveNext
'         Loop
'         If QryIpdeptKeyWord = "" Then
'            RsTemp.MoveFirst
'            Do While RsTemp.EOF = False
'               If "" & RsTemp.Fields("lk14") = "" Then
'                  strTi15 = strTi15 & ";" & RsTemp.Fields("LK01")
'                  QryIpdeptKeyWord = RsTemp.Fields("LK02")
'                  m_Sender = ""
'                  tmpArr = Split("" & RsTemp.Fields("LK04"), ";")
'                  For j = 0 To UBound(tmpArr)
'                     strTemp = Pub_GetSpecMan(CStr(tmpArr(j)))
'                     If strTemp <> "" Then
'                        m_Sender = m_Sender & ";" & strTemp
'                     Else
'                        m_Sender = m_Sender & ";" & tmpArr(j)
'                     End If
'                  Next j
'                  Exit Do
'               End If
'               RsTemp.MoveNext
'            Loop
'         End If
'      End With
'   End If
'End Function

Public Function PUB_PatentByChkCP14(strSubject As String, strPI11 As String, _
   strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, _
   ByRef ToSortOut As String, ByRef m_Sender As String, _
   Optional ByVal bolShowMsg As Boolean = False) As Boolean
Dim tmpArr As Variant
Dim strTemp As String
Dim pa() As String, sp() As String
   
   ReDim pa(1 To TF_PA) As String
   ReDim sp(1 To tf_SP) As String
   
   ToSortOut = "": m_Sender = ""
   'strText = PUB_ChgEnglishStyle(PUB_ChgNumeralStyle(UCase(strSubject)))
   PUB_PatentByChkCP14 = True
   '讀取基本檔資料
   Select Case strCP01
      Case "P", "CFP":
         If PUB_ReadPatentData(pa(), strCP01, strCP02, strCP03, strCP04) = False Then
            If bolShowMsg = True Then MsgBox "無此案號！", vbCritical
            PUB_PatentByChkCP14 = False
            Exit Function
         End If
      Case Else:
         If PUB_ReadServicePracticeData(sp(), strCP01, strCP02, strCP03, strCP04) = False Then
            If bolShowMsg = True Then MsgBox "無此案號！", vbCritical
            PUB_PatentByChkCP14 = False
            Exit Function
         End If
   End Select
   '依設定之國家區分程序人員
   Select Case strCP01
      Case "CFP", "CPS":
         'Modify by Sindy 2020/3/18
         '109/4/1以後改業務區劃分
         If strSrvDate(1) >= CFP業務區劃分啟用日 Then
            m_Sender = PUB_GetCFPHandler(strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04, ToSortOut)
            If ToSortOut <> "" Then ToSortOut = UCase(ToSortOut) '轉大寫
         Else
         '2020/3/18 END
            'Modify By Sindy 2018/6/21 CFP程序分組變動,調整新規則
            If strSrvDate(1) > 20180622 Then
               strExc(0) = "SELECT na01,na02,na73,na74 FROM nation WHERE na01='" & IIf(pa(9) <> "", pa(9), sp(9)) & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  '美日
                  If RsTemp.Fields("na01") = "101" Or RsTemp.Fields("na01") = "011" Then
                     If strCP02 Mod 2 = 0 Then '雙數
                        ToSortOut = "4"
                        m_Sender = m_Sender & ";" & Pub_GetSpecMan("專利處轉信美日雙號程序")
                     Else
                        '單數
                        ToSortOut = "3"
                        m_Sender = m_Sender & ";" & Pub_GetSpecMan("專利處轉信美日單號程序")
                     End If
                  Else
                     If strCP02 Mod 2 = 0 Then '雙數
                        ToSortOut = "6"
                        m_Sender = m_Sender & ";" & Pub_GetSpecMan("專利處轉信美日以外雙號程序")
                     Else
                        '單數
                        ToSortOut = "5"
                        m_Sender = m_Sender & ";" & Pub_GetSpecMan("專利處轉信美日以外單號程序")
                     End If
                  End If
               End If
            Else
            '2018/6/21 END
               strExc(0) = "SELECT na02,na73,na74 FROM nation WHERE na01='" & IIf(pa(9) <> "", pa(9), sp(9)) & "' and NA02<'C1'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  ToSortOut = "3"
                  m_Sender = m_Sender & ";" & Pub_GetSpecMan("專利處轉信亞洲程序")
               Else
                  strExc(0) = "SELECT na02,na73,na74 FROM nation WHERE na01='" & IIf(pa(9) <> "", pa(9), sp(9)) & "' and NA02='C20'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     ToSortOut = "4"
                     m_Sender = m_Sender & ";" & Pub_GetSpecMan("專利處轉信歐洲程序")
                  Else
                     strExc(0) = "SELECT na02,na73,na74 FROM nation WHERE na01='" & IIf(pa(9) <> "", pa(9), sp(9)) & "' and NA02 in('C10','C30','C40')"
                     intI = 1
                     Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        If strCP02 Mod 2 = 0 Then '雙數
                           ToSortOut = "6"
                           m_Sender = m_Sender & ";" & Pub_GetSpecMan("專利處轉信美洋非洲雙號程序")
                        Else
                           '單數
                           ToSortOut = "5"
                           m_Sender = m_Sender & ";" & Pub_GetSpecMan("專利處轉信美洋非洲單號程序")
                        End If
                     End If
                  End If
               End If
            End If
         End If
   End Select
   
   'Modify By Sindy 2025/1/8
   If m_Sender = "" Then 'Modify By Sindy 2025/2/11 +
      If strSrvDate(1) >= P業務區劃分啟用日 Then
         'P案管制人
         m_Sender = PUB_GetPHandler(strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04, ToSortOut)
         If ToSortOut <> "" Then ToSortOut = UCase(ToSortOut) '轉大寫
      Else
      '2025/1/8 END
         'P,PS其申請國家須為非台灣
         If strCP01 = "P" And pa(9) = "000" Then
            ToSortOut = "7" '其他 'Add By Sindy 2017/12/11 雅娟:P台灣案不要分類,留在其他
            If bolShowMsg = True Then MsgBox "P案須為非台灣！", vbCritical
            PUB_PatentByChkCP14 = False
            Exit Function
         ElseIf strCP01 = "PS" And pa(9) = "000" Then
            ToSortOut = "7" '其他 'Add By Sindy 2017/12/11 雅娟:P台灣案不要分類,留在其他
            If bolShowMsg = True Then MsgBox "P案須為非台灣！", vbCritical
            PUB_PatentByChkCP14 = False
            Exit Function
         End If
      End If
   End If
   'If ToSortOut = "" Then ByCaseNoChkCP14 = False
End Function

'Add by Sindy 2020/3/18
'CFP案管制人-信件分類
'109/4/1以後改業務區劃分
'Public Sub PUB_AddItemCFPHandler(oCombo As ComboBox, oCombo2 As ComboBox)
'Modify By Sindy 2022/3/21 + , Optional intMaxItem As Integer = 4 : 程式設計上CFP固定4個頁籤
'Modify By Sindy 2025/1/9 + , Optional strReadKind As String = "CFP": 抓CFP程序人員 還是 P程序人員
Public Sub PUB_AddItemCFPHandler(oCombo As Object, oCombo2 As Object, _
   Optional intMaxItem As Integer = 4, Optional strReadKind As String = "CFP")
   Dim strSql As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim intItem As Integer
   Dim strType As String
   
   intItem = 0
   'Add By Sindy 2022/4/6
   strType = 1 '預設為1
Type2:
   If strType = 2 Then
      'Modify By Sindy 2025/1/9
      If strReadKind = "P" Then
         strSql = "select ST01 a0917,ST02,ST17 from staff where st04='1' and st05 in('73','75') and st01<'F'" & _
                  " and not exists(select * from acc090 where a0917 IS NOT NULL and a0917=st01)"
      Else
      '2025/1/9 END
         strSql = "select ST01 a0916,ST02,ST17 from staff where st04='1' and st05 in('83','85') and st01<'F'" & _
                  " and not exists(select * from acc090 where a0916 IS NOT NULL and a0916=st01)"
      End If
   Else
   '2022/4/6 END
      'Modify By Sindy 2025/1/9
      If strReadKind = "P" Then
         strSql = "SELECT a0917,ST02,ST17 FROM staff,acc090 WHERE a0917 IS NOT NULL AND a0917=st01(+) AND st04='1'" & _
                  " GROUP BY a0917,ST02,ST17" & _
                  " order by a0917 asc"
      Else
      '2025/1/9 END
         strSql = "SELECT a0916,ST02,ST17 FROM staff,acc090 WHERE a0916 IS NOT NULL AND a0916=st01(+) AND st04='1'" & _
                  " GROUP BY a0916,ST02,ST17" & _
                  " order by a0916 asc"
      End If
   End If
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
      rsQuery.MoveFirst
      Do While Not rsQuery.EOF
         intItem = intItem + 1
         'Modify By Sindy 2022/3/21
         If intItem > intMaxItem Then
            Exit Do
         End If
         '2022/3/21 END
         oCombo.AddItem UCase(rsQuery.Fields("ST17")) & " " & rsQuery.Fields("ST02")
         'Modify By Sindy 2025/1/9 rsQuery.Fields("a0916") => rsQuery.Fields(0)
         oCombo2.AddItem UCase(rsQuery.Fields("ST17")) & " " & rsQuery.Fields(0)
         rsQuery.MoveNext
      Loop
   End If
   'Add By Sindy 2022/4/6 CFP新人暫還沒有要劃分負責的業務區,還是人員手動分信給她,但在分信的頁籤上要出現她的名字
   'Modiofy By Sindy 2022/6/14 只Run一次 + And strType = 1
   If intItem < intMaxItem And strType = 1 Then
      strType = 2
      GoTo Type2
   End If
   '2022/4/6 END
   
   Set rsQuery = Nothing
End Sub

'Modify By Sindy 2017/3/8 變共用函數
'解析本所案號
'Modify By Sindy 2017/12/19 + Optional strMailBox As String = ""
'                             "":沒傳; 傳入收件夾
'Modify By Sindy 2019/3/6 strMailBox="L":代表是卷宗區呼叫此函數
'Modify By Sindy 2021/3/18 回傳:strII18 如何find到案號
'Modify By Sindy 2021/9/29 + , Optional ByVal BolFindOtherKind As Boolean = True : 除了正規的檢查本所案號之外, 還要用其他方式檢查
Public Function PUB_IPDeptGetCaseNo(ByVal strText As String, ByVal strCompText As String, _
   ByRef strCP01 As String, ByRef strCP02 As String, ByRef strCP03 As String, _
   ByRef strCP04 As String, Optional ByRef strPA150 As String, _
   Optional bolIncludeTit As Boolean = False, Optional strMailBox As String = "", _
   Optional ByRef strII18 As String, Optional ByVal BolFindOtherKind As Boolean = True) As Boolean

Dim rsTmp As New ADODB.Recordset
Dim strData As String
Dim intStar As Integer, intIdx As Integer
Dim i As Integer
Dim strSubject As String, j As Integer 'Add By Sindy 2020/9/22
Dim strChkSubject As String 'Add By Sindy 2020/12/8
Dim strChkKeyWord As String, k As Integer 'Add By Sindy 2021/1/29
Dim intLen As Integer 'Add By Sindy 2021/8/12
   
   strSubject = strText 'Add By Sindy 2020/9/22
   
   '***** 因 OURREF 等於 YOURREF 字串的一部份,為了不要二者互相影響
   '程式轉換成 YOURREF ==> YOUREF 比對
   If strCompText = "YOURREF" Then strCompText = "YOUREF"
   
   strText = PUB_ChgEnglishStyle(PUB_ChgNumeralStyle(UCase(strText)))
   
   'Add By Sindy 2020/8/12 為商標處
   'FW: COCO 都可及? 第29? ?方文?：T-221519T；COCO及? 第43? ?方文?：T-221515
   strText = Replace(strText, UCase("?方文?："), "OURREF")
   '2020/8/12 END
   
   strText = Replace(strText, "..", " ")
   strText = Replace(strText, "...", " ")
   strText = Replace(strText, ".", " ")
   strText = Replace(strText, "．", " ")
   strText = Replace(strText, ":", " ")
   strText = Replace(strText, "：", " ")
'   strText = Replace(strText, ",", " ")
'   strText = Replace(strText, "，", " ")
   strText = Replace(strText, "- ", "-") 'Add By Sindy 2021/1/21
   strText = Replace(strText, " -", "-") 'Add By Sindy 2021/1/21
   strText = Replace(strText, "－", "-")
   'strText = Replace(strText, "_", " ") mark:FW: REMINDER Y/Ref:CFP-028551_O/Ref:B18206FR_Decision Certifying Forfeiture
   strText = Replace(strText, "　", " ") '***** 最後才清空白 Add By Sindy 2017/7/18
   'Modify By Sindy 2017/8/28 只留單空白
   Do While InStr(strText, "  ") > 0
      strText = Replace(strText, "  ", " ")
   Loop
   '2017/8/28 END
   'Modify By Sindy 2021/1/20 改寫 ex:Possible New Patent Application in your Country; our Ref: P4640TW00
   strText = Replace(strText, "[", " ") '清空白格,方便解析案號的結束
   strText = Replace(strText, "]", " ")
   strText = Replace(strText, "〔", " ")
   strText = Replace(strText, "〕", " ")
   strText = Replace(strText, "(", " ")
   strText = Replace(strText, ")", " ")
   strText = Replace(strText, "（", " ")
   strText = Replace(strText, "）", " ")
   '2021/1/20 END
   
   'Rreplace時還要考慮到字母排列的問題:如.Your Ref必須比Our Ref先處理,因都有our
   strText = Replace(strText, UCase("Your Reference"), "YOUREF")
   strText = Replace(strText, UCase("Your Ref No"), "YOUREF")
   strText = Replace(strText, UCase("Your Ref"), "YOUREF")
   strText = Replace(strText, UCase("Y Ref"), "YOUREF")
   strText = Replace(strText, UCase("Your File Ref"), "YOUREF")
   strText = Replace(strText, UCase("Your-Ref"), "YOUREF") 'Add By Sindy 2016/4/26
   'strText = Replace(strText, UCase("Y/REF"), "YOUREF") 'Add By Sindy 2016/10/18
   strText = Replace(strText, UCase("YourRef"), "YOUREF") 'Add By Sindy 2016/11/21
   strText = Replace(strText, UCase("Tai E Ref"), "YOUREF")
   strText = Replace(strText, UCase("TaiE refs"), "YOUREF") 'Add By Sindy 2023/2/2 薛經理通知要加入
   strText = Replace(strText, UCase("TaiE ref"), "YOUREF")
   strText = Replace(strText, UCase("Tai E File"), "YOUREF") 'Add By Sindy 2016/10/26
   strText = Replace(strText, UCase("Tai E International Ref"), "YOUREF")
   strText = Replace(strText, UCase("Tai E"), "YOUREF") 'Add By Sindy 2019/9/9
   strText = Replace(strText, UCase("TaiE"), "YOUREF") 'Add By Sindy 2020/9/20 ex:TaiE.FCT-xxxxx
   'Add By Sindy 2020/10/6
   'Y/Ref: FCP-
   'YR. REF.FCP-
   strText = Replace(strText, UCase("Y/Ref"), "YOUREF")
   strText = Replace(strText, UCase("YR REF"), "YOUREF")
   '2020/10/6 END
   
   strText = Replace(strText, UCase("Y/R"), "YOUREF") 'Add By Sindy 2017/10/13 為專利處
   
   strText = Replace(strText, UCase("Our Reference"), "OURREF")
   strText = Replace(strText, UCase("Our Ref No"), "OURREF")
   strText = Replace(strText, UCase("Our Refs"), "OURREF") '多案不可歸個案,分不出來
   strText = Replace(strText, UCase("Our Ref"), "OURREF")
   
   'strText = Replace(strText, UCase("O/REF"), "OURREF") 'Add By Sindy 2016/10/18
   strText = Replace(strText, UCase("OurRef"), "OURREF") 'Add By Sindy 2016/11/21
   
   strText = Replace(strText, UCase("O/R"), "OURREF") 'Add By Sindy 2017/10/13 為專利處
   strText = Replace(strText, UCase("我方案號"), "OURREF") 'Add By Sindy 2017/10/13 為專利處
   strText = Replace(strText, UCase("本所案號"), "OURREF") 'Add By Sindy 2017/10/13 為商標處
   strText = Replace(strText, UCase("貴方案號"), "YOUREF") 'Add By Sindy 2017/10/13 為專利處
   strText = Replace(strText, UCase("貴方卷號"), "YOUREF") 'Add By Sindy 2025/2/21 為內商
   
   strText = Replace(strText, UCase("貴所番?"), "YOUREF")
   strText = Replace(strText, UCase("弊所番?"), "YOUREF")
   strText = Replace(strText, UCase("弊所整理番?"), "YOUREF")
   strText = Replace(strText, UCase("貴所整理番?"), "OURREF")
   strText = Replace(strText, UCase("貴社整理番?"), "OURREF") 'Add By Sindy 2020/9/2
   'Add By Sindy 2024/7/18 嘉雯的困擾
   '  因北京巨京商標部 寄信過來的主旨 T183892續展案提申文件1(本所案號：CMT-240585)
   '  讓系統混淆, 因和本所的案號相近, 易解析錯案號;
   '  為方便使用者不用一直去改案號, 增加此段程式
   strText = Replace(strText, UCase("CMT-"), "ZZ-")
   '2024/7/18 END
   
   'Add By Sindy 2021/8/12 uniocde偵試不到? ex:ELC/jw RE: 台?特許(貴所整理番?：FCP-061414、弊所整理番?：PF20200053-TW)年金?費用見積???伺? [REPLY.]
'   貴所番?
'   弊所番?
'   弊所整理番?
'   貴所整理番?
'   貴社整理番?
   If InStr(strText, "貴所番") > 0 Then
      intLen = InStr(strText, "貴所番")
      strText = Left(strText, intLen - 1) & "YOUREF" & Mid(strText, intLen + 6)
   End If
   If InStr(strText, "弊所番") > 0 Then
      intLen = InStr(strText, "弊所番")
      strText = Left(strText, intLen - 1) & "YOUREF" & Mid(strText, intLen + 6)
   End If
   If InStr(strText, "弊所整理番") > 0 Then
      intLen = InStr(strText, "弊所整理番")
      strText = Left(strText, intLen - 1) & "YOUREF" & Mid(strText, intLen + 6)
   End If
   If InStr(strText, "貴所整理番") > 0 Then
      intLen = InStr(strText, "貴所整理番")
      strText = Left(strText, intLen - 1) & "OURREF" & Mid(strText, intLen + 6)
   End If
   If InStr(strText, "貴社整理番") > 0 Then
      intLen = InStr(strText, "貴社整理番")
      strText = Left(strText, intLen - 1) & "OURREF" & Mid(strText, intLen + 6)
   End If
   '2021/8/12 END
   
'   If InStr(strText, "OURREF") = 0 Then
'      strText = Replace(strText, UCase("O R"), "OURREF")
'   End If
   'strText = Replace(strText, " ", "") '***** 最後才清空白 Add By Sindy 2017/6/14
   
   intStar = 1
   
FindNext:
   PUB_IPDeptGetCaseNo = False
   strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = "": strPA150 = ""
   intIdx = 0
   If InStr(intStar, strText, strCompText) > 0 Then
      intStar = InStr(intStar, strText, strCompText)
      
'      'strData = Replace(Trim(Mid(strText, intStar + Len(strCompText))), " ", "") ex.Taiwanese Pat. Appl. No. 101104838, Your Ref: FCP-045215, Our Ref: 30752/46091A TW
      If bolIncludeTit = True Then '判斷系統別
         'Modify By Sindy 2017/12/29 FW: US Patent Application No. 15/134,372; Your Ref. CFP- 028420, Our Ref. P005925
         'Modify By Sindy 2021/1/21 Mark (bolIncludeTit判斷) ex:RE: URGENT REMINDER: CFP-030029 (0073447.0020 - 20180069.ORI)
         '                          變 CFP-03002900734470020-20180069ORI => 很怪
         strData = Replace(Trim(Mid(strText, intStar)), " ", "\")
         'strData = Replace(Trim(Mid(strText, intStar)), " ", "")
         '2017/12/29 END
      Else
         'strData = Replace(Trim(Mid(strText, intStar + Len(strCompText))), " ", "")
         'Modify By Sindy 2020/11/25 FW RE INDIA NEW PATENT APPLICATION NO202044051106; YOUREF CFP 032032; OURREF 3588-2020/ASB/MS
         'Modify By Sindy 2020/11/26
         '  OURREF P-126239 RE ACK - DY/LY FW NEW NATIONAL PHASE FILING IN CHINA FOR OURREF --0204-700LOG-18CN
         '  P-126239\RE\ACK\-\DY/LY\FW\NEW\NATIONAL\PHASE\FILING\IN\CHINA\FOR\OURREF\--0204-700LOG-18CN
         strData = Replace(Trim(Mid(strText, intStar + Len(strCompText))), " ", "\")
         '  P-126239-RE-ACK---DY/LY-FW-NEW-NATIONAL-PHASE-FILING-IN-CHINA-FOR-OURREF---0204-700LOG-18CN => 會錯,解析不到案號
         'strData = Replace(Trim(Mid(strText, intStar + Len(strCompText))), " ", "-")
      End If

'      For i = 1 To Len(strData)
'         '必須開頭是英文字母A~Z
'         If intIdx = 0 And Not (Asc(Mid(strData, i, 1)) >= 65 And Asc(Mid(strData, i, 1)) <= 90) Then
'            Exit For
'         End If
'         If Asc(Mid(strData, i, 1)) = 32 Then Exit For '空白格離開
'         '-:45
'         '0~9:48~57
'         'A~Z:65~90
'         If Asc(Mid(strData, i, 1)) >= 65 And Asc(Mid(strData, i, 1)) <= 90 _
'            And intIdx <= 1 And Len(strCP01) < 3 Then  '系統別
'            If intIdx = 0 Then intIdx = 1
'            strCP01 = strCP01 & Mid(strData, i, 1)
'         '-:45
'         ElseIf Asc(Mid(strData, i, 1)) = 45 Then
'            intIdx = intIdx + 1
'         '0~9:48~57
'         ElseIf Asc(Mid(strData, i, 1)) >= 48 And Asc(Mid(strData, i, 1)) <= 57 _
'            And intIdx <= 2 And Len(strCP02) < 6 Then '流水號6碼
'            If strCP02 = "" And intIdx = 1 Then intIdx = 2
''            If intIdx = 2 And Len(strCP02) < 6 Then
'               strCP02 = strCP02 & Mid(strData, i, 1)
''            Else
''               Exit For
''            End If
'         '0~9:48~57
'         'A~Z:65~90
'         ElseIf ((Asc(Mid(strData, i, 1)) >= 65 And Asc(Mid(strData, i, 1)) <= 90) Or _
'                 (Asc(Mid(strData, i, 1)) >= 48 And Asc(Mid(strData, i, 1)) <= 57)) _
'            And intIdx <= 3 And Len(strCP03) < 1 Then '第3欄位1碼
'            If strCP01 = "T" Then Exit For 'Add By Sindy 2020/9/10 T,沒有CP03,CP04則不解析
'            If strCP03 = "" And Mid(strData, i - 1, 1) <> "-" Then Exit For 'Add By Sindy 2017/7/21
'            intIdx = 3
'            strCP03 = Mid(strData, i, 1)
'
'         ElseIf Asc(Mid(strData, i, 1)) >= 48 And Asc(Mid(strData, i, 1)) <= 57 _
'            And intIdx <= 4 And Len(strCP04) < 2 Then '第4欄位2碼
'            If strCP04 = "" And Mid(strData, i - 1, 1) <> "-" Then Exit For 'Add By Sindy 2017/7/21
'            If strCP04 = "" And intIdx = 3 Then intIdx = 4
''            If intIdx = 4 And Len(strCP04) < 2 Then
'               strCP04 = strCP04 & Mid(strData, i, 1)
''            Else
''               Exit For
''            End If
'         Else
'            Exit For
'         End If
'      Next i
      'Modify By Sindy 2021/1/20 解析本所案號,改寫 ex:Possible New Patent Application in your Country; our Ref: P4640TW00
      If Trim(strData) <> "" Then
         '必須開頭是英文字母A~Z
         If intIdx = 0 And Asc(Mid(strData, 1, 1)) >= 65 And Asc(Mid(strData, 1, 1)) <= 90 Then
            intIdx = 1
            For i = 1 To Len(strData)
               '-:45
               '0~9:48~57
               'A~Z:65~90
               '系統別3碼
               If intIdx = 1 Then
                  'A~Z
                  If Asc(Mid(strData, i, 1)) >= 65 And Asc(Mid(strData, i, 1)) <= 90 _
                     And Len(strCP01) < 3 Then
                     strCP01 = strCP01 & Mid(strData, i, 1)
                  Else
                     '是否有符合intIdx=2 流水號
                     If (Asc(Mid(strData, i, 1)) >= 48 And Asc(Mid(strData, i, 1)) <= 57) Or _
                        Mid(strData, i, 1) = "-" Then
                        If Mid(strData, i, 1) = "-" Then
                           i = i + 1
                        End If
                        If Mid(strData, i, 1) = "" Then Exit For
                        intIdx = 2
                     Else
                        strCP01 = ""
                        Exit For
                     End If
                  End If
               End If
               '流水號6碼
               If intIdx = 2 Then
                  '0~9
                  If Asc(Mid(strData, i, 1)) >= 48 And Asc(Mid(strData, i, 1)) <= 57 _
                     And Len(strCP02) < 6 Then
                     strCP02 = strCP02 & Mid(strData, i, 1)
                  Else
                     '是否有符合intIdx=3 第3欄位
                     If Mid(strData, i, 1) = "-" Then
                        'Modify By Sindy 2021/1/25
                        'FW: Patent - Registration No. 103519 - Invoice issued - Your ref.: CFP-023682 - Invoice Term: Immediately
                        If InStr(strSubject, strCP01 & "-" & strCP02 & " ") > 0 Or _
                           InStr(strSubject, strCP01 & strCP02 & " ") > 0 Then
                           Exit For
                        End If
                        '2021/1/25 END
                        i = i + 1
                        If Mid(strData, i, 1) = "" Then Exit For
                        intIdx = 3
                     Else
                        If Asc(Mid(strData, i, 1)) = 32 Or Mid(strData, i, 1) = "\" Then
                           Exit For '空白格離開
                        Else
                           'A~Z 或 0~9 : 還有資料代表不是所內案號 ex:Your Ref.: CFP-031511, Our Ref.: P005980
                           If ((Asc(Mid(strData, i, 1)) >= 65 And Asc(Mid(strData, i, 1)) <= 90) Or _
                               (Asc(Mid(strData, i, 1)) >= 48 And Asc(Mid(strData, i, 1)) <= 57)) Then
                              strCP01 = "": strCP02 = ""
                           End If
                           Exit For
                        End If
                     End If
                  End If
               End If
               '第3欄位1碼
               If intIdx = 3 Then
                  'A~Z 或 0~9
                  If ((Asc(Mid(strData, i, 1)) >= 65 And Asc(Mid(strData, i, 1)) <= 90) Or _
                      (Asc(Mid(strData, i, 1)) >= 48 And Asc(Mid(strData, i, 1)) <= 57)) _
                     And Len(strCP03) < 1 Then
                     strCP03 = Mid(strData, i, 1)
                  Else
                     'Modify By Sindy 2023/10/25 mark: ex:CH/bs - New Trademark Application in Taiwan "C'NEWLAB & Device" in Class 03 in the name of Ghang Tai LEE (Your Ref: T23-0133-TW)(TaiE Ref: FCT-051142)[LTR]
                     'If strCP01 = "T" Then Exit For 'Add By Sindy 2020/9/10 T,沒有CP03,CP04則不解析
                     '2023/10/25 END
                     '是否有符合intIdx=4 第4欄位
                     If Mid(strData, i, 1) = "-" Then
                        'Modify By Sindy 2021/1/25
                        'FW: Patent - Registration No. 103519 - Invoice issued - Your ref.: CFP-023682-0 - Invoice Term: Immediately
                        If InStr(strSubject, strCP01 & "-" & strCP02 & "-" & strCP03 & " ") > 0 Or _
                           InStr(strSubject, strCP01 & strCP02 & strCP03 & " ") > 0 Then
                           Exit For
                        End If
                        '2021/1/25 END
                        i = i + 1
                        If Mid(strData, i, 1) = "" Then Exit For
                        intIdx = 4
                     Else
                        If Asc(Mid(strData, i, 1)) = 32 Or Mid(strData, i, 1) = "\" Then
                           Exit For '空白格離開
                        Else
                           'A~Z 或 0~9 : 還有資料代表不是所內案號
                           If ((Asc(Mid(strData, i, 1)) >= 65 And Asc(Mid(strData, i, 1)) <= 90) Or _
                               (Asc(Mid(strData, i, 1)) >= 48 And Asc(Mid(strData, i, 1)) <= 57)) Then
                              strCP01 = "": strCP02 = "": strCP03 = ""
                           End If
                           Exit For
                        End If
                     End If
                  End If
               End If
               '第4欄位2碼
               If intIdx = 4 Then
                  '0~9
                  If Asc(Mid(strData, i, 1)) >= 48 And Asc(Mid(strData, i, 1)) <= 57 _
                     And Len(strCP04) < 2 Then
                     strCP04 = strCP04 & Mid(strData, i, 1)
                     'FW: Y/Ref:CFP-030217-0-12_O/Ref:B31932FR_report
                     If Len(strCP04) = 2 Then Exit For
                  Else
                     'Modify By Sindy 2021/1/25
                     If (InStr(strSubject, strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04 & " ") > 0 Or _
                        InStr(strSubject, strCP01 & strCP02 & strCP03 & strCP04 & " ") > 0) _
                        And Mid(strData, i, 1) = "-" Then
                        Exit For
                     Else
                     '2021/1/25 END
                        If Asc(Mid(strData, i, 1)) = 32 Or Mid(strData, i, 1) = "\" Then
                           Exit For '空白格離開
                        Else
                           'A~Z 或 0~9 : 還有資料代表不是所內案號
                           If ((Asc(Mid(strData, i, 1)) >= 65 And Asc(Mid(strData, i, 1)) <= 90) Or _
                               (Asc(Mid(strData, i, 1)) >= 48 And Asc(Mid(strData, i, 1)) <= 57)) Then
                              strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
                           End If
                           Exit For
                        End If
                     End If
                  End If
               End If
            Next i
         End If
      End If
      intStar = intStar + 1
      'Modify By Sindy 2019/3/6 代表是卷宗區呼叫此函數,任何本所案號都可以歸卷
      If Not (strMailBox = "L") Then
      '2019/3/6 END
         'Modify By Sindy 2017/7/28 國外部分信:排除L
         If strCP01 = "L" Then
            strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
         End If
         '2017/7/28 END
      End If
      If strCP01 <> "" And strCP02 <> "" Then
         strCP02 = Format(strCP02, "000000") '補足6碼
         If strCP03 = "" Then strCP03 = "0"
         If strCP04 = "" Then
            strCP04 = "00"
         'Modify By Sindy 2018/10/4
         Else
            strCP04 = Format(strCP04, "00") '補足2碼
         End If
         '檢查本所案號是否存在
         strExc(0) = "SELECT '' pa150 FROM TRADEMARK WHERE TM01='" & strCP01 & "' AND TM02='" & strCP02 & "' AND TM03='" & strCP03 & "' AND TM04='" & strCP04 & "' "
         strExc(0) = strExc(0) + " union all select pa150 FROM PATENT WHERE PA01='" & strCP01 & "' AND PA02='" & strCP02 & "' AND PA03='" & strCP03 & "' AND PA04='" & strCP04 & "' "
         strExc(0) = strExc(0) + " union all select '' pa150 FROM SERVICEPRACTICE WHERE SP01='" & strCP01 & "' AND SP02='" & strCP02 & "' AND SP03='" & strCP03 & "' AND SP04='" & strCP04 & "' "
         strExc(0) = strExc(0) + " union all select '' pa150 FROM LAWCASE WHERE LC01='" & strCP01 & "' AND LC02='" & strCP02 & "' AND LC03='" & strCP03 & "' AND LC04='" & strCP04 & "' "
         'Modify By Sindy 2017/7/28 國外部分信:排除顧問基本檔
         'strExc(0) = strExc(0) + " union all select '' pa150 FROM HIRECASE WHERE HC01='" & strCP01 & "' AND HC02='" & strCP02 & "' AND HC03='" & strCP03 & "' AND HC04='" & strCP04 & "' "
         'Modify By Sindy 2019/3/6 代表是卷宗區呼叫此函數,任何本所案號都可以歸卷
         If strMailBox = "L" Then
            strExc(0) = strExc(0) + " union all select '' pa150 FROM HIRECASE WHERE HC01='" & strCP01 & "' AND HC02='" & strCP02 & "' AND HC03='" & strCP03 & "' AND HC04='" & strCP04 & "' "
         End If
         '2019/3/6 END
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            strPA150 = "" & RsTemp.Fields("pa150")
            'Add By Sindy 2017/7/24 增加檢查此本所案號是否有進度檔
            strExc(0) = "SELECT cp09 FROM caseprogress WHERE cp01='" & strCP01 & "' AND cp02='" & strCP02 & "' AND cp03='" & strCP03 & "' AND cp04='" & strCP04 & "' "
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               'Modify By Sindy 2017/12/19 Patent只抓CFP,CPS,P,PS
               If Left(strMailBox, 2) = "03" Then 'Patent收件匣 = "03.專利處Patent收信郵件"
                  If strCP01 <> "CFP" And strCP01 <> "CPS" And _
                     strCP01 <> "P" And strCP01 <> "PS" Then
                     strCP01 = ""
                     strCP02 = ""
                     strCP03 = ""
                     strCP04 = ""
                     GoTo RunNext
                  End If
               End If
               strII18 = strCompText 'Add By Sindy 2021/3/18
               PUB_IPDeptGetCaseNo = True '*****
            Else
               strCP01 = ""
               strCP02 = ""
               strCP03 = ""
               strCP04 = ""
            End If
            '2017/7/24 END
         Else
            strCP01 = ""
            strCP02 = ""
            strCP03 = ""
            strCP04 = ""
         End If
      Else
         strCP01 = ""
         strCP02 = ""
         strCP03 = ""
         strCP04 = ""
      End If
      
RunNext:
      'Modify By Sindy 2018/1/12 字串要尋找到最後
      'If (bolIncludeTit = True Or Left(strMailBox, 2) = "03") And _
         strCP01 = "" Then
      If strCP01 = "" Then
         GoTo FindNext
      End If
   End If
   
'*******************************
'商標,申請案號
'*******************************
   'Add By Sindy 2020/11/4
   '  Application No. + 該案申請號
   'Modify By Sindy 2020/11/24 +
   '  App. No. + 該案申請號
   '  App No. + 該案申請號
   '  Appln. No. + 該案申請號
   '  Appln No. + 該案申請號
   'Modify By Sindy 2021/9/29 + And BolFindOtherKind = True : 除了正規的檢查本所案號之外, 還要用其他方式檢查
   If strCP01 = "" And BolFindOtherKind = True Then
'From: no-reply@ngb.co.jp [mailto:no-reply@ngb.co.jp]
'Sent: Thursday, November 05, 2020 2:32 PM
'To: ipdept <ipdept@taie.com.tw>
'Subject: [ipBOX] A new message was uploaded. ID: ipBOX_20201105153223_3328990
      'Modify By Sindy 2022/2/16 商標關鍵字判斷：前面請加 TM 或 Trademark或 Trade Mark
      If InStr(UCase(strText), UCase("TM Application No")) > 0 Or _
         InStr(UCase(strText), UCase("TM App No")) > 0 Or _
         InStr(UCase(strText), UCase("TM Appln No")) > 0 Or _
         InStr(UCase(strText), UCase("Trademark Application No")) > 0 Or _
         InStr(UCase(strText), UCase("Trademark App No")) > 0 Or _
         InStr(UCase(strText), UCase("Trademark Appln No")) > 0 Or _
         InStr(UCase(strText), UCase("Trade Mark Application No")) > 0 Or _
         InStr(UCase(strText), UCase("Trade Mark App No")) > 0 Or _
         InStr(UCase(strText), UCase("Trade Mark Appln No")) > 0 Then
         
         'Add By Sindy 2020/12/8
         strExc(8) = 0: strExc(9) = 0: strExc(10) = 0
         If InStr(UCase(strText), UCase("TM Application No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("TM Application No"))
         ElseIf InStr(UCase(strText), UCase("TM App No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("TM App No"))
         ElseIf InStr(UCase(strText), UCase("TM Appln No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("TM Appln No"))
         End If
         If InStr(UCase(strText), UCase("Trademark Application No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Trademark Application No"))
         ElseIf InStr(UCase(strText), UCase("Trademark App No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Trademark App No"))
         ElseIf InStr(UCase(strText), UCase("Trademark Appln No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Trademark Appln No"))
         End If
         If InStr(UCase(strText), UCase("Trade Mark Application No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Trade Mark Application No"))
         ElseIf InStr(UCase(strText), UCase("Trade Mark App No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Trade Mark App No"))
         ElseIf InStr(UCase(strText), UCase("Trade Mark Appln No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Trade Mark Appln No"))
         End If
         strExc(9) = InStr(UCase(strText), UCase("Your Ref"))
         strExc(10) = InStr(UCase(strText), UCase("Our Ref"))
         If Val(strExc(9)) > 0 And Val(strExc(10)) > 0 Then
            If Val(strExc(9)) > Val(strExc(10)) Then
               strExc(9) = strExc(10)
            End If
         Else
            If Val(strExc(10)) > 0 Then strExc(9) = strExc(10)
         End If
         If Val(strExc(8)) > Val(strExc(9)) Then
            strChkSubject = Mid(strText, strExc(8))
         Else
            strChkSubject = Mid(strText, strExc(8), strExc(9) - strExc(8))
         End If
         '2020/12/8 END
         
         strChkSubject = ReplSymbolToBlank(strChkSubject) 'Add By Sindy 2022/4/18
         '用申請案號檢索
         'RE: URGENT NOTICE AND DEADLINE!!! LY/th Accounting Issue Re: Taiwan Patent Application No. 102135654 & Taiwan Utility Model Patent Application No. 102218423 (Patent No. M501723) Y45130
         '不用判斷閉卷,銷卷 and tm29 is null and tm57 is null / and pa57 is null and pa108 is null
         'Modify By Sindy 2021/3/22 + 申請案號||' '
         '   主旨為 FW: Notice of Laying-open and Initiation of Substantive Examination Chinese Patent Application No. 2019108700211 Y/R:P-122885 O/R:FI-195098-0221
         '   但抓到TM12=870021
         'Modify By Sindy 2022/2/16 =>專利,商標分開抓 ex:RE: LY/th RE: New Taiwan Patent Application Claiming Priority from Singapore Patent Application No. 10202105796S  filed on 01 June 2021; SF Ref:78963TW  [EFILES-SFSG.FID2278534]【CFT-009797 Saved】
         'Modify By Sindy 2022/4/18 tm12 => ' '||tm12||' '
         'Modify By Sindy 2022/10/26 抓未閉卷未銷卷 + and tm30||tm57 is null
         'Modify By Sindy 2023/12/8 mark:" and instr('" & ChgSQL(strChkSubject) & "',' '||tm12||' ')<17"
         strSql = "select tm01,tm02,tm03,tm04,tm11,tm12,tm15,tm16,tm57,instr('" & ChgSQL(strChkSubject) & "',' '||tm12||' ') as m_sort" & _
                  " from trademark" & _
                  " Where ((Length(tm12)>=6 and tm10<>'000') or (Length(tm12)>=9 and tm10='000'))" & _
                  " and tm01 in('CFT','FCT')" & _
                  " and instr('" & ChgSQL(strChkSubject) & "',' '||tm12||' ')>0" & _
                  " and tm30||tm57 is null" & _
                  " order by m_sort asc"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If rsTmp.RecordCount = 1 Then
               strCP01 = rsTmp.Fields("tm01")
               strCP02 = rsTmp.Fields("tm02")
               strCP03 = rsTmp.Fields("tm03")
               strCP04 = rsTmp.Fields("tm04")
               strII18 = "申請案號" 'Add By Sindy 2021/3/18
               PUB_IPDeptGetCaseNo = True '*****
            End If
         'Modify By Sindy 2023/12/8 因上頭的比對是用 strText 此變數在最前頭有做一些空白的處理,
         '                          所以再用原主旨 strSubject 檢查一次
         '(May反應)ex:TMA034588-TW-NF - PING YU CHING (CHINESE) - Cl. 05 - Registration Number 01626974 - [REND] - 2024 - Maintenance Instructions
         Else
            strSql = "select tm01,tm02,tm03,tm04,tm11,tm12,tm15,tm16,tm57,instr('" & ChgSQL(strSubject) & "',' '||tm12||' ') as m_sort" & _
                     " from trademark" & _
                     " Where ((Length(tm12)>=6 and tm10<>'000') or (Length(tm12)>=9 and tm10='000'))" & _
                     " and tm01 in('CFT','FCT')" & _
                     " and instr('" & ChgSQL(strSubject) & "',' '||tm12||' ')>0" & _
                     " and tm30||tm57 is null" & _
                     " order by m_sort asc"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If rsTmp.RecordCount = 1 Then
                  strCP01 = rsTmp.Fields("tm01")
                  strCP02 = rsTmp.Fields("tm02")
                  strCP03 = rsTmp.Fields("tm03")
                  strCP04 = rsTmp.Fields("tm04")
                  strII18 = "申請案號" 'Add By Sindy 2021/3/18
                  PUB_IPDeptGetCaseNo = True '*****
               End If
            End If
         '2023/12/8 END
         End If
      End If
   End If
   '2020/11/4 END
'*******************************
'專利,申請案號
'*******************************
   'Add By Sindy 2020/11/4
   '  Application No. + 該案申請號
   'Modify By Sindy 2020/11/24 +
   '  App. No. + 該案申請號
   '  App No. + 該案申請號
   '  Appln. No. + 該案申請號
   '  Appln No. + 該案申請號
   'Modify By Sindy 2021/9/29 + And BolFindOtherKind = True : 除了正規的檢查本所案號之外, 還要用其他方式檢查
   If strCP01 = "" And BolFindOtherKind = True Then
'From: no-reply@ngb.co.jp [mailto:no-reply@ngb.co.jp]
'Sent: Thursday, November 05, 2020 2:32 PM
'To: ipdept <ipdept@taie.com.tw>
'Subject: [ipBOX] A new message was uploaded. ID: ipBOX_20201105153223_3328990
      If InStr(UCase(strText), UCase("Application No")) > 0 Or _
         InStr(UCase(strText), UCase("App No")) > 0 Or _
         InStr(UCase(strText), UCase("Appln No")) > 0 Then
         
         'Add By Sindy 2020/12/8
         strExc(8) = 0: strExc(9) = 0: strExc(10) = 0
         If InStr(UCase(strText), UCase("Application No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Application No"))
         ElseIf InStr(UCase(strText), UCase("App No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("App No"))
         ElseIf InStr(UCase(strText), UCase("Appln No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Appln No"))
         End If
         strExc(9) = InStr(UCase(strText), UCase("Your Ref"))
         strExc(10) = InStr(UCase(strText), UCase("Our Ref"))
         If Val(strExc(9)) > 0 And Val(strExc(10)) > 0 Then
            If Val(strExc(9)) > Val(strExc(10)) Then
               strExc(9) = strExc(10)
            End If
         Else
            If Val(strExc(10)) > 0 Then strExc(9) = strExc(10)
         End If
         If Val(strExc(8)) > Val(strExc(9)) Then
            strChkSubject = Mid(strText, strExc(8))
         Else
            strChkSubject = Mid(strText, strExc(8), strExc(9) - strExc(8))
         End If
         '2020/12/8 END
         
         strChkSubject = ReplSymbolToBlank(strChkSubject) 'Add By Sindy 2022/4/18
         '用申請案號檢索
         'RE: URGENT NOTICE AND DEADLINE!!! LY/th Accounting Issue Re: Taiwan Patent Application No. 102135654 & Taiwan Utility Model Patent Application No. 102218423 (Patent No. M501723) Y45130
         '不用判斷閉卷,銷卷 and tm29 is null and tm57 is null / and pa57 is null and pa108 is null
         'Modify By Sindy 2021/3/22 + 申請案號||' '
         '   主旨為 FW: Notice of Laying-open and Initiation of Substantive Examination Chinese Patent Application No. 2019108700211 Y/R:P-122885 O/R:FI-195098-0221
         '   但抓到TM12=870021
         'Modify By Sindy 2022/2/16 =>專利,商標分開抓 ex:RE: LY/th RE: New Taiwan Patent Application Claiming Priority from Singapore Patent Application No. 10202105796S  filed on 01 June 2021; SF Ref:78963TW  [EFILES-SFSG.FID2278534]【CFT-009797 Saved】
         'Modify By Sindy 2022/4/18 pa11 => ' '||pa11||' '
         'Modify By Sindy 2022/10/26 抓未閉卷未銷卷 + and pa58||pa108 is null
         'Modify By Sindy 2023/12/8 mark:" and instr('" & ChgSQL(strChkSubject) & "',' '||pa11||' ')<17"
         strSql = "select pa01,pa02,pa03,pa04,pa10,pa11,pa22,pa16,pa108,instr('" & ChgSQL(strChkSubject) & "',' '||pa11||' ') as m_sort" & _
                  " from patent" & _
                  " Where ((Length(pa11)>=6 and pa09<>'000') or (Length(pa11)>=9 and pa09='000'))" & _
                  " and instr('" & ChgSQL(strChkSubject) & "',' '||pa11||' ')>0" & _
                  " and pa58||pa108 is null" & _
                  " order by m_sort asc"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If rsTmp.RecordCount = 1 Then
               strCP01 = rsTmp.Fields("pa01")
               strCP02 = rsTmp.Fields("pa02")
               strCP03 = rsTmp.Fields("pa03")
               strCP04 = rsTmp.Fields("pa04")
               strII18 = "申請案號" 'Add By Sindy 2021/3/18
               PUB_IPDeptGetCaseNo = True '*****
            End If
         End If
      End If
   End If
   '2020/11/4 END
'*******************************
'專利號數
'*******************************
   'Add By Sindy 2020/11/4
   '  Patent No. + 該案專利號
   '  Pat No. + 該案專利號
   '  Registration No. + 該案註冊號
   '  Reg. No. + 該案註冊號
'AP/fc - Taiwan Patent Application No. 109141137 ; Our Ref: BSA595-TW-NP ; Your Ref: FCP-064212 [REP. 1101]
'=>申請案號和專利號要分開抓,因很容易數字會撞上
'LY/th Accounting Issue Re: Taiwan Patent Application No. 102135654 & Taiwan Utility Model Patent Application No. 102218423 (Patent No. M501723) Y45130
'=>專利,商標分開抓
   'Modify By Sindy 2021/9/29 + And BolFindOtherKind = True : 除了正規的檢查本所案號之外, 還要用其他方式檢查
   If strCP01 = "" And BolFindOtherKind = True Then
      If InStr(UCase(strText), UCase("Patent No")) > 0 Or _
         InStr(UCase(strText), UCase("Pat No")) > 0 Then
         'Add By Sindy 2020/12/8
         '[專利處]FW: Renewal of Hong Kong Standard Patent No. 1179579 (Your Ref: P-105045; Our Ref: HKP/2013/66656)
         strExc(8) = 0: strExc(9) = 0: strExc(10) = 0
         If InStr(UCase(strText), UCase("Patent No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Patent No"))
         ElseIf InStr(UCase(strText), UCase("Pat No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Pat No"))
         End If
         strExc(9) = InStr(UCase(strText), UCase("Your Ref"))
         strExc(10) = InStr(UCase(strText), UCase("Our Ref"))
         If Val(strExc(9)) > 0 And Val(strExc(10)) > 0 Then
            If Val(strExc(9)) > Val(strExc(10)) Then
               strExc(9) = strExc(10)
            End If
         Else
            If Val(strExc(10)) > 0 Then strExc(9) = strExc(10)
         End If
         If Val(strExc(8)) > Val(strExc(9)) Then
            strChkSubject = Mid(strText, strExc(8))
         Else
            strChkSubject = Mid(strText, strExc(8), strExc(9) - strExc(8))
         End If
         '2020/12/8 END
         
         strChkSubject = ReplSymbolToBlank(strChkSubject) 'Add By Sindy 2022/4/18
         '用註冊號/專利號數檢索
         '不用判斷閉卷,銷卷 and tm29 is null and tm57 is null / and pa57 is null and pa108 is null
         'Modify By Sindy 2021/3/22 + 專利號數||' '
         'Modify By Sindy 2021/10/6 RE: LY/th - Transfer of Cases in the name of T-DATA SYSTEMS (S) PTE LTD. Taiwan Patent Applications Nos. 100104864 (Patent No. I439862), 099143669 (Patent No. I444887) and 100121200 (Patent No. I509527)
         '                          ==> 專利號數||' ' 取消 ||' '
         '                          加判斷台灣案長度, 及order by instr
         'Modify By Sindy 2022/4/18 pa22 => ' '||pa22||' '
         'Modify By Sindy 2022/10/26 抓未閉卷未銷卷 + and pa58||pa108 is null
         'Modify By Sindy 2023/12/8 mark:" and instr('" & ChgSQL(strChkSubject) & "',' '||pa22||' ')<14"
         strSql = "select pa01,pa02,pa03,pa04,pa10,pa11,pa22,pa16,pa108" & _
                  " from patent" & _
                  " Where ((Length(pa22)>=6 and pa09<>'000') or (Length(pa22)=7 and pa09='000'))" & _
                  " and instr('" & ChgSQL(strChkSubject) & "',' '||pa22||' ')>0" & _
                  " and pa58||pa108 is null" & _
                  " order by instr('" & ChgSQL(strChkSubject) & "',' '||pa22||' ') asc"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If rsTmp.RecordCount = 1 Then
               strCP01 = rsTmp.Fields("pa01")
               strCP02 = rsTmp.Fields("pa02")
               strCP03 = rsTmp.Fields("pa03")
               strCP04 = rsTmp.Fields("pa04")
               strII18 = "專利號數" 'Add By Sindy 2021/3/18
               PUB_IPDeptGetCaseNo = True '*****
            End If
         End If
      End If
   End If
'*******************************
'註冊號
'*******************************
   'Modify By Sindy 2021/9/29 + And BolFindOtherKind = True : 除了正規的檢查本所案號之外, 還要用其他方式檢查
   If strCP01 = "" And BolFindOtherKind = True Then
      '用註冊號/專利號數檢索
      'Modify By Sindy 2023/12/8 +InStr(UCase(strText), UCase("Registration Number")) > 0
      '               (May反應)ex:TMA034588-TW-NF - PING YU CHING (CHINESE) - Cl. 05 - Registration Number 01626974 - [REND] - 2024 - Maintenance Instructions
      If InStr(UCase(strText), UCase("Registration No")) > 0 Or _
         InStr(UCase(strText), UCase("Reg No")) > 0 Or _
         InStr(UCase(strText), UCase("Registration Number")) > 0 Then
         'Add By Sindy 2020/12/8
         strExc(8) = 0: strExc(9) = 0: strExc(10) = 0
         If InStr(UCase(strText), UCase("Registration No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Registration No"))
         ElseIf InStr(UCase(strText), UCase("Reg No")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Reg No"))
         'Add By Sindy 2023/12/8
         ElseIf InStr(UCase(strText), UCase("Registration Number")) > 0 Then
            strExc(8) = InStr(UCase(strText), UCase("Registration Number"))
         '2023/12/8 END
         End If
         strExc(9) = InStr(UCase(strText), UCase("Your Ref"))
         strExc(10) = InStr(UCase(strText), UCase("Our Ref"))
         If Val(strExc(9)) > 0 And Val(strExc(10)) > 0 Then
            If Val(strExc(9)) > Val(strExc(10)) Then
               strExc(9) = strExc(10)
            End If
         Else
            If Val(strExc(10)) > 0 Then strExc(9) = strExc(10)
         End If
         If Val(strExc(8)) > Val(strExc(9)) Then
            strChkSubject = Mid(strText, strExc(8))
         Else
            strChkSubject = Mid(strText, strExc(8), strExc(9) - strExc(8))
         End If
         '2020/12/8 END
         
         'Add By Sindy 2022/4/18
         'RE: (漏信問題) ◎Your ref:  Zacco ref: T133700TW01 Trademark Registration No. 1006596 in Taiwan TENGTOOLS (logo)【CFT-007149 Saved】
         strChkSubject = ReplSymbolToBlank(strChkSubject)
         '2022/4/18 END
         '不用判斷閉卷,銷卷 and tm29 is null and tm57 is null / and pa57 is null and pa108 is null
         'Modify By Sindy 2021/3/22 + 註冊號||' '
         'Modify By Sindy 2022/4/18 tm15 => ' '||tm15||' '
         'Modify By Sindy 2022/10/26 抓未閉卷未銷卷 + and tm30||tm57 is null
         'Modify By Sindy 2023/12/8 mark:" and instr('" & ChgSQL(strChkSubject) & "',' '||tm15||' ')<18"
         strSql = "select tm01,tm02,tm03,tm04,tm11,tm12,tm15,tm16,tm57" & _
                  " from trademark" & _
                  " Where ((Length(tm15)>=6 and tm10<>'000') or (Length(tm15)=8 and tm10='000'))" & _
                  " and tm01 in('CFT','FCT')" & _
                  " and instr('" & ChgSQL(strChkSubject) & "',' '||tm15||' ')>0" & _
                  " and tm30||tm57 is null" & _
                  " order by instr('" & ChgSQL(strChkSubject) & "',' '||tm15||' ') asc"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If rsTmp.RecordCount = 1 Then
               strCP01 = rsTmp.Fields("tm01")
               strCP02 = rsTmp.Fields("tm02")
               strCP03 = rsTmp.Fields("tm03")
               strCP04 = rsTmp.Fields("tm04")
               strII18 = "註冊號" 'Add By Sindy 2021/3/18
               PUB_IPDeptGetCaseNo = True '*****
            End If
         'Modify By Sindy 2023/12/8 因上頭的比對是用 strText 此變數在最前頭有做一些空白的處理,
         '                          所以再用原主旨 strSubject 檢查一次
         '(May反應)ex:TMA034588-TW-NF - PING YU CHING (CHINESE) - Cl. 05 - Registration Number 01626974 - [REND] - 2024 - Maintenance Instructions
         Else
            strSql = "select tm01,tm02,tm03,tm04,tm11,tm12,tm15,tm16,tm57" & _
                     " from trademark" & _
                     " Where ((Length(tm15)>=6 and tm10<>'000') or (Length(tm15)=8 and tm10='000'))" & _
                     " and tm01 in('CFT','FCT')" & _
                     " and instr('" & ChgSQL(strSubject) & "',' '||tm15||' ')>0" & _
                     " and tm30||tm57 is null" & _
                     " order by instr('" & ChgSQL(strSubject) & "',' '||tm15||' ') asc"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               If rsTmp.RecordCount = 1 Then
                  strCP01 = rsTmp.Fields("tm01")
                  strCP02 = rsTmp.Fields("tm02")
                  strCP03 = rsTmp.Fields("tm03")
                  strCP04 = rsTmp.Fields("tm04")
                  strII18 = "註冊號" 'Add By Sindy 2021/3/18
                  PUB_IPDeptGetCaseNo = True '*****
               End If
            End If
         '2023/12/8 END
         End If
      End If
   End If
   '2020/11/4 END
   
'*******************************
'彼所案號
'*******************************
   'Add By Sindy 2020/9/22 Our Ref: 彼所案號 =>從關鍵字Our Ref往後找,直到空格結束
   'Modify By Sindy 2021/9/29 + And BolFindOtherKind = True : 除了正規的檢查本所案號之外, 還要用其他方式檢查
   If strCP01 = "" And BolFindOtherKind = True Then
'      strSubject = UCase(strSubject) ': strText = ""
'      strExc(9) = InStr(strSubject, UCase("Your Ref"))
'      strExc(10) = InStr(strSubject, UCase("Our Ref"))
'      If Val(strExc(9)) > 0 Then
'         If strExc(9) < strExc(10) Then
'            strExc(10) = strExc(9) + Len("Your Ref")
'         Else
'            strExc(10) = 0
'         End If
'      Else
'         strExc(10) = 0
'      End If
'      If Val(strExc(10)) > 0 Then
'         strChkSubject = Mid(strSubject, strExc(10))
'         strChkKeyWord = "Our Ref"
'      End If
      strSubject = Replace(UCase(strSubject), UCase("Your Ref"), UCase("Our Ref"))
      'Modify By Sindy 2022/1/11 個案請加本所彼所案號前置關鍵字 完全一樣之案號
      'Your Reference Number:
      'Our Reference Number:
      strSubject = Replace(UCase(strSubject), UCase("Your Reference Number"), UCase("Our Ref"))
      strSubject = Replace(UCase(strSubject), UCase("Our Reference Number"), UCase("Our Ref"))
      '2022/1/11 END
      'Add By Sindy 2021/1/29 + 加前置關鍵字 Y/R: O/R:
      'InStr(strSubject, UCase("Our Ref")) > 0
      If InStr(UCase(strSubject), UCase("Our Ref")) > 0 Or _
         InStr(UCase(strSubject), UCase("Y/R")) > 0 Or _
         InStr(UCase(strSubject), UCase("O/R")) > 0 Then
         For k = 1 To 3
            strText = ""
            strChkSubject = strSubject
            If k = 1 Then
               strChkKeyWord = UCase("Our Ref")
            ElseIf k = 2 Then
               strChkKeyWord = "Y/R"
            Else
               strChkKeyWord = "O/R"
            End If
            If InStr(UCase(strChkSubject), UCase(strChkKeyWord)) = 0 Then
               strChkSubject = ""
            End If
            If strChkSubject <> "" Then
               For i = InStr(UCase(strChkSubject), UCase(strChkKeyWord)) + Len(strChkKeyWord) To Len(strChkSubject)
                  If strText = "" And _
                     (Mid(strChkSubject, i, 1) = ":" Or Mid(strChkSubject, i, 1) = "." Or _
                      Mid(strChkSubject, i, 1) = " ") Then
                     'It is OK
                  'Modify By Sindy 2022/1/11 ex:strSubject = OUR REF: B1348.70033TW00; OUR REF: FCP-057299 ==> + Or Mid(strChkSubject, i, 1) = ";")
                  ElseIf (Mid(strChkSubject, i, 1) = " " Or Mid(strChkSubject, i, 1) = ";") And strText <> "" Then
                     Exit For
                  Else
                     strText = strText & Mid(strChkSubject, i, 1)
                  End If
               Next i
               '依彼所案號比對本所個案, 若本所有多案彼號相同, 則不自動分信
               If strText <> "" Then
                  strText = ChgSQL(strText) 'Add By Sindy 2020/12/15 ex:OUR REF: CFT-021995[INCOM]舒淨醫材委辦之"YANG'S SUPPORT及圖"英國商申
                  '主檔彼所案號
                  'Modify By Sindy 彼所案號大於3碼的才抓出來比對 此封誤存個案    FW: ◎Our ref.: C 20 0260 B/ fy / Claiming priority from TW-Applications【P-055151 Saved】
                  'Modify By Sindy 2022/10/26 抓未閉卷未銷卷
                  strExc(0) = "SELECT tm01,tm02,tm03,tm04 FROM trademark WHERE TM45='" & strText & "' and TM45 is not null and length(TM45)>3 and tm30||tm57 is null" & _
                              " union " & _
                              "SELECT pa01,pa02,pa03,pa04 FROM patent WHERE PA77='" & strText & "' and PA77 is not null and length(PA77)>3 and pa58||pa108 is null" & _
                              " union " & _
                              "SELECT sp01,sp02,sp03,sp04 FROM servicepractice WHERE SP27='" & strText & "' and SP27 is not null and length(SP27)>3 and sp16||sp61 is null" & _
                              " union " & _
                              "SELECT lc01,lc02,lc03,lc04 FROM lawcase WHERE LC23='" & strText & "' and LC23 is not null and length(LC23)>3 and LC09||LC34 is null"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If RsTemp.RecordCount = 1 Then
                        strCP01 = RsTemp.Fields(0)
                        strCP02 = RsTemp.Fields(1)
                        strCP03 = RsTemp.Fields(2)
                        strCP04 = RsTemp.Fields(3)
                     End If
                  End If
                  '進度彼所案號
                  strExc(0) = "SELECT cp01,cp02,cp03,cp04 FROM CASEPROGRESS WHERE CP45='" & strText & "' and CP45 is not null and length(CP45)>3" & _
                              " group by cp01,cp02,cp03,cp04"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If RsTemp.RecordCount = 1 Then
                        If strCP01 = "" Then
                           strCP01 = RsTemp.Fields(0)
                           strCP02 = RsTemp.Fields(1)
                           strCP03 = RsTemp.Fields(2)
                           strCP04 = RsTemp.Fields(3)
                           'Modify By Sindy 2022/10/26 抓未閉卷未銷卷
                           strExc(0) = "SELECT tm01,tm02,tm03,tm04 FROM trademark WHERE TM01='" & strCP01 & "' and TM02='" & strCP02 & "' and TM03='" & strCP03 & "' and TM04='" & strCP04 & "' and tm30||tm57 is null" & _
                                       " union " & _
                                       "SELECT pa01,pa02,pa03,pa04 FROM patent WHERE PA01='" & strCP01 & "' and PA02='" & strCP02 & "' and PA03='" & strCP03 & "' and PA04='" & strCP04 & "' and pa58||pa108 is null" & _
                                       " union " & _
                                       "SELECT sp01,sp02,sp03,sp04 FROM servicepractice WHERE SP01='" & strCP01 & "' and SP02='" & strCP02 & "' and SP03='" & strCP03 & "' and SP04='" & strCP04 & "' and sp16||sp61 is null" & _
                                       " union " & _
                                       "SELECT lc01,lc02,lc03,lc04 FROM lawcase WHERE LC01='" & strCP01 & "' and LC02='" & strCP02 & "' and LC03='" & strCP03 & "' and LC04='" & strCP04 & "' and LC09||LC34 is null"
                           intI = 1
                           Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                           If intI = 0 Then
                              strCP01 = ""
                              strCP02 = ""
                              strCP03 = ""
                              strCP04 = ""
                           End If
                           '2022/10/26 END
                        ElseIf strCP01 <> RsTemp.Fields(0) And _
                               strCP02 <> RsTemp.Fields(1) And _
                               strCP03 <> RsTemp.Fields(2) And _
                               strCP04 <> RsTemp.Fields(3) Then
                           strCP01 = ""
                           strCP02 = ""
                           strCP03 = ""
                           strCP04 = ""
                        End If
                     Else
                        strCP01 = ""
                        strCP02 = ""
                        strCP03 = ""
                        strCP04 = ""
                     End If
                  End If
                  If strCP01 <> "" And strCP02 <> "" Then
                     strII18 = "彼所案號" 'Add By Sindy 2021/3/18
                     PUB_IPDeptGetCaseNo = True '*****
                     Exit For
                  End If
               End If
            End If
            '2021/1/29 END
         Next k
      End If
   End If
   '2020/9/22 END
   
   Set rsTmp = Nothing
End Function

Public Function ReplSymbolToBlank(ByVal strChkSubject As String) As String
Dim jj As Integer
   
   ReplSymbolToBlank = ""
   For jj = 1 To Len(strChkSubject)
      '全型符號表變空白
      If InStr(WM_全型符號表, Mid(strChkSubject, jj, 1)) > 0 Then
         ReplSymbolToBlank = ReplSymbolToBlank & " "
      '半型符號表變空白
      ElseIf InStr(WM_半型符號表, Mid(strChkSubject, jj, 1)) > 0 Then
         ReplSymbolToBlank = ReplSymbolToBlank & " "
      Else
         ReplSymbolToBlank = ReplSymbolToBlank & Mid(strChkSubject, jj, 1)
      End If
   Next jj
   
   If ReplSymbolToBlank = "" Then
      ReplSymbolToBlank = strChkSubject
   End If
End Function

'Add By Sindy 2020/2/3 檢查此案號是否為國外部辦理的案件
Private Function LocChkCaseNoIsF(ByRef strCP01 As String, ByRef strCP02 As String, _
   ByRef strCP03 As String, ByRef strCP04 As String) As Boolean
   
   LocChkCaseNoIsF = True
   'Add By Sindy 2018/4/9 +if
   'Modify By Sindy 2018/8/29 拿掉,'T','TM' ex:[Our Ref: T2000.0115/AKC/THTL/rmrs] [B&B-SGPMatters.FID344213]【T-200001 Saved】
   If InStr("'FCT','CFT','CFC','S'", "'" & strCP01 & "'") = 0 Then
   '2018/4/9 END
      '歸個案時若該案件進度智權人員都沒有Fxx人員收文時也歸其他
      strExc(0) = "select count(*) from caseprogress" & _
                  " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                  " and substr(cp12,1,1)='F'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Val(RsTemp.Fields(0)) <= 0 Then
            'Modify By Sindy 2016/4/21
            'PUB_IPDept_ToSortOut = "Z" '其他
            'GoTo ChkEnd
            'Modify By Sindy 2017/8/11 檢查是否為國外申請案 CF..
            strExc(0) = "SELECT TM01 FROM TRADEMARK WHERE TM01='" & strCP01 & "' AND TM02='" & strCP02 & "' AND TM03='" & strCP03 & "' AND TM04='" & strCP04 & "' AND TM10='000'"
            strExc(0) = strExc(0) + " union all select PA01 FROM PATENT WHERE PA01='" & strCP01 & "' AND PA02='" & strCP02 & "' AND PA03='" & strCP03 & "' AND PA04='" & strCP04 & "' AND PA09='000'"
            strExc(0) = strExc(0) + " union all select SP01 FROM SERVICEPRACTICE WHERE SP01='" & strCP01 & "' AND SP02='" & strCP02 & "' AND SP03='" & strCP03 & "' AND SP04='" & strCP04 & "' AND SP09='000'"
            strExc(0) = strExc(0) + " union all select LC01 FROM LAWCASE WHERE LC01='" & strCP01 & "' AND LC02='" & strCP02 & "' AND LC03='" & strCP03 & "' AND LC04='" & strCP04 & "' AND LC15='000'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 0 Then
               LocChkCaseNoIsF = False
               strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
               Exit Function
            End If
            '2017/8/11 END
            '2016/4/21 END
            'Add By Sindy 2017/10/25
            '歸個案時若該案件進度承辦人都沒有Fxx人員時也歸其他
            strExc(0) = "select count(*) from caseprogress,staff" & _
                        " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                        " and cp14=st01(+) and substr(st03,1,1)='F'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Val(RsTemp.Fields(0)) <= 0 Then
                  LocChkCaseNoIsF = False
                  strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
                  Exit Function
               End If
            End If
            '2017/10/25 END
         End If
      End If
   End If
End Function

'Modify By Sindy 2017/3/8 變共用函數
'分類
'1.個案 2.外商 3.外專 4.專利處 5.外法 6.新知 7.財務 Z.其他 8.開拓
'回傳:分類 及 收受者
'回傳:strII18 被分到那一組關鍵字
'Modify By Sindy 2022/8/5 + Optional ByVal bolOnlyReadCaseNo As Boolean = False : True=僅為讀取案號
Public Function PUB_IPDept_ToSortOut(strSubject As String, strII11 As String, _
      ByRef m_Sender As String, ByRef strCP01 As String, ByRef strCP02 As String, _
      ByRef strCP03 As String, ByRef strCP04 As String, Optional ByRef strII18 As String, _
      Optional ByVal bolOnlyReadCaseNo As Boolean = False) As String
Dim strPA150 As String
Dim strCP13 As String '目前智權人員
Dim strText As String
Dim rsTmp As New ADODB.Recordset
Dim tmpArr As Variant
Dim YourRefCase As String, OurRefCase As String
Dim strTemp1 As String, strTemp2 As String, strTemp3 As String, StrTemp4 As String
Dim strTemp As String
Dim bolF23EngW As Boolean 'Add By Sindy 2016/5/30
Dim j As Integer
Dim bolAllF22 As Boolean 'Add By Sindy 2017/4/12
Dim strCUNo As String, strCUNoCU10 As String, strFaNo As String, strFANoFA10 As String, strArea As String, strAreaNA55 As String 'Add By Sindy 2017/12/27
Dim bolChkOk As Boolean, strWord As String
Dim strNation As String 'Add By Sindy 2018/3/12
Dim bolFind As Boolean 'Add By Sindy 2018/5/18
Dim strCySender As String 'Add By Sindy 2020/1/16
Dim strTmp As String, strEmpSender As String 'Add By Sindy 2020/3/6
Dim bolChin As Boolean, jj As Integer 'Add By Sindy 2020/3/30
Dim strChkSubject As String 'Add By Sindy 2020/5/14
Dim intChinWord As Integer 'Add By Sindy 2020/5/28 中文字數
Dim strOurII18 As String, strYourII18 As String 'Add By Sindy 2021/3/18
Dim str239TM01 As String, str239TM02 As String, str239TM03 As String, str239TM04 As String 'Add By Sindy 2021/3/18
Dim strEmp As String 'Add By Sindy 2021/6/15
Dim strMail As String, strDomain As String
   
   bolF23EngW = False 'Add By Sindy 2016/5/30
   PUB_IPDept_ToSortOut = "": m_Sender = ""
   strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
   YourRefCase = "": OurRefCase = "": strII18 = ""
   strCySender = ""
   str239TM01 = "": str239TM02 = "": str239TM03 = "": str239TM04 = ""
   '先解析有無本所案號
   strText = strSubject
   
'************************************************************************
   'Modify By Sindy 2018/11/30 主旨設定錯誤,為分信順利在此Replace
   '2019 Chinese bamboo calendars to convey our best wishes from Tai E International Patent & Law Office [ISDXXXXXXXXX_ETC] (EY_wc)
   strText = Replace(strText, "_ETC] (EY_wc)", ".ETC] (EY/wc)")
   '2018/11/30 END
   'Add By Sindy 2024/8/28
   strText = Replace(strText, "　", " ") 'ex: AIPPI　invitation 置換全形空白為單空白,做後續比對
   '多個空白,置換為單空白
   Do While InStr(strText, "  ") > 0
      strText = Replace(strText, "  ", " ")
   Loop
   '2024/8/28 END
'************************************************************************
   
   '直接歸入其他類別:
   'Add By Sindy 2020/4/8 若寄件者為本所員工, 就直接分到其他類別
   '只要是上列條件全都進其他 (個案主旨沒錯  但內文為其他單位  故本所員工轉入, 皆歸其他, 由內文或由主旨新註明單位判斷)
   'Modify By Sindy 2022/11/22 排除主旨裡有(法律所)字樣
   If InStr(strText, "(法律所)") = 0 Then
   '2022/11/22 END
      If UCase(pub_DbTerminalName) = 正式資料庫電腦名稱 Then '正式資料庫
         strExc(0) = "SELECT st01,st02,st04,st69 From staff" & _
                     " where st01>'6' and st01<'F'" & _
                     " AND substr(st01,4,1)<>'9'" & _
                     " AND (st04='1' or (st04='2' and st51>=" & DBDATE(DateAdd("d", -7, Format(strSrvDate(1), "####/##/##"))) & "))" & _
                     " AND st01 NOT IN('60000')" & _
                     " AND substr(st03,1,1)<>'R'" & _
                     " AND (InStr(upper('" & ChgSQL(strII11) & "'), upper(st02)) > 0 Or (InStr(upper('" & ChgSQL(strII11) & "'), upper(ST01)) > 0 and InStr(upper('" & ChgSQL(strII11) & "'), upper('/O=TAIE/OU=DOMAIN')) > 0))"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            PUB_IPDept_ToSortOut = "Z" '其他
            m_Sender = ""
            GoTo ChkEnd
         End If
      End If
   End If
   '2020/4/8 END
   
'Add By Sindy 2019/8/16 未傳遞的主旨略過
If Trim(strII11) <> "未傳遞的主旨" Then
'2019/8/16 END

   'Add By Sindy 2023/7/21 因顧服組將部分客戶的商標案件智權人員移交商標部接手，其客戶來信將寄至 ipdept@taie.com.tw
   '為避免漏信, 增設下列分信規則且優先於信件主旨「本所案號」(個案)及「Initial」分信原則
   '寄件者:
   If InStr(strII11, "@") > 0 Then
      If InStr(strII11, " [") > 0 Then
         tmpArr = Split(strII11, " [")
         strMail = Left(tmpArr(1), Len(tmpArr(1)) - 1)
      Else
         'Modify By Sindy 2025/2/5 ex:"Tamas Gyomber" <no_reply@yesmywine.com>
         If InStr(strII11, " <") > 0 Then
            tmpArr = Split(strII11, " <")
            strMail = Left(tmpArr(1), Len(tmpArr(1)) - 1)
         Else
         '2025/2/5 END
            strMail = strII11
         End If
      End If
      tmpArr = Split(strMail, "@")
      strDomain = "@" & tmpArr(UBound(tmpArr))
      'E-Mail:
      strSql = "SELECT cu01,cu02,cu04,cu20,cu15,cu115,cu116,cu117,cu118 From customer" & _
               " where (" & _
                   "upper(cu20)='" & UCase(ChgSQL(Trim(strMail))) & "'" & _
               " or upper(cu115)='" & UCase(ChgSQL(Trim(strMail))) & "'" & _
               " or upper(cu116)='" & UCase(ChgSQL(Trim(strMail))) & "'" & _
               " or upper(cu117)='" & UCase(ChgSQL(Trim(strMail))) & "'" & _
               " or upper(cu118)='" & UCase(ChgSQL(Trim(strMail))) & "'" & _
               ")" & _
               " and cu13='P2006'" & _
               " and cu02='0'" & _
               " order by cu01 asc"
      intI = 1
      Set rsTmp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         'If rsTmp.RecordCount = 1 Then
            PUB_IPDept_ToSortOut = "2" '外商
            strII18 = "(P2006-" & rsTmp.Fields("cu01") & "-" & strMail & ")"
            m_Sender = Pub_GetSpecMan("IPDept信箱之P2006管理人員")
            GoTo ChkEnd '跳至,條件檢查完畢
         'End If
      Else
         '網域: 排除網域 ＠yahoo.com 及 ＠yahoo.com.tw 及 @gmail.com 及 @qq.com 及 @126.com
         'Modify By Sindy 2024/7/10 + 排除 @taie.com.tw
         'Modify By Sindy 2024/7/18 + 排除 @msa.hinet.net
         If UCase(strDomain) <> UCase("@gmail.com") And _
            UCase(strDomain) <> UCase("@qq.com") And _
            UCase(strDomain) <> UCase("@126.com") And _
            UCase(strDomain) <> UCase("@yahoo.com.tw") And _
            InStr(UCase(strDomain), UCase("@yahoo.com")) = 0 And _
            UCase(strDomain) <> UCase("@taie.com.tw") And _
            UCase(strDomain) <> UCase("@msa.hinet.net") Then
            
            strSql = "SELECT cu01,cu02,cu04,cu20,cu15,cu115,cu116,cu117,cu118 From customer" & _
                     " where (" & _
                         "instr(upper(cu20),'" & UCase(ChgSQL(Trim(strDomain))) & "')> 0" & _
                     " or instr(upper(cu115),'" & UCase(ChgSQL(Trim(strDomain))) & "')> 0" & _
                     " or instr(upper(cu116),'" & UCase(ChgSQL(Trim(strDomain))) & "')> 0" & _
                     " or instr(upper(cu117),'" & UCase(ChgSQL(Trim(strDomain))) & "')> 0" & _
                     " or instr(upper(cu118),'" & UCase(ChgSQL(Trim(strDomain))) & "')> 0" & _
                     ")" & _
                     " and cu13='P2006'" & _
                     " and cu02='0'" & _
                     " order by cu01 asc"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               'If rsTmp.RecordCount = 1 Then
                  PUB_IPDept_ToSortOut = "2" '外商
                  strII18 = "(P2006-" & rsTmp.Fields("cu01") & "-" & strDomain & ")"
                  m_Sender = Pub_GetSpecMan("IPDept信箱之P2006管理人員")
                  GoTo ChkEnd '跳至,條件檢查完畢
               'End If
            End If
         End If
      End If
   End If
   '2023/7/21 END
   
   'Modify By Sindy 2021/3/18 + , , , strII18
   If PUB_IPDeptGetCaseNo(strText, "OURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , , strII18) = False Then
      If PUB_IPDeptGetCaseNo(strText, "YOURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , , strII18) = False Then
      End If
   'Modify By Sindy 2021/9/29 若是用申請案號,專利號,彼所號等抓到資料, 再解析一次案號
   'ex: RE: DY/bc - Taiwan Patent Application No. 106114285; Your Ref: ADVSIL-13-TW / MM; Our Ref: FCP-056692 [PROC.1917]
   ElseIf strII18 <> "OURREF" Then
      strTemp1 = strCP01: strTemp2 = strCP02: strTemp3 = strCP03: StrTemp4 = strCP04: strOurII18 = strII18
      If PUB_IPDeptGetCaseNo(strText, "YOURREF", strCP01, strCP02, strCP03, strCP04, strPA150, , , strII18) = False Then
      End If
      'YOURREF 沒找到 或 找到不是個案, 就用OURREF找到的資料,做後續比對
      If strII18 = "" Or strII18 <> "YOURREF" Then
         strCP01 = strTemp1: strCP02 = strTemp2: strCP03 = strTemp3: strCP04 = StrTemp4: strII18 = strOurII18
      End If
   '2021/9/29 END
   End If
   
'   '再加強尋找 YR OR
'   If strCP01 = "" And strCP02 = "" Then
'      strText = Replace(strText, UCase("YR "), "YOURREF")
'      strText = Replace(strText, UCase("OR "), "OURREF")
'      If PUB_IPDeptGetCaseNo(strText, "YOURREF", strCP01, strCP02, strCP03, strCP04, strPA150) = False Then
'         If PUB_IPDeptGetCaseNo(strText, "OURREF", strCP01, strCP02, strCP03, strCP04, strPA150) = False Then
'         End If
'      End If
'   End If
   'Add By Sindy 2016/3/29
   strTemp1 = "": strTemp2 = "": strTemp3 = "": StrTemp4 = "": strOurII18 = ""
   'Modify By Sindy 2021/3/18 + , , , strOurII18
   'Modify By Sindy 2021/9/29 + , IIf(InStr("申請案號、專利號數、彼所案號", strII18) = 0 And strII18 <> "", False, True):已有抓到本所案號
   If PUB_IPDeptGetCaseNo(strText, "OURREF", strTemp1, strTemp2, strTemp3, StrTemp4, strPA150, , , strOurII18, IIf(InStr("申請案號、專利號數、彼所案號", strII18) = 0 And strII18 <> "", False, True)) = True Then
      OurRefCase = strTemp1 & "-" & strTemp2 & "-" & strTemp3 & "-" & StrTemp4
   End If
   strTemp1 = "": strTemp2 = "": strTemp3 = "": StrTemp4 = "": strYourII18 = ""
   'Modify By Sindy 2021/3/18 + , , , strYourII18
   'Modify By Sindy 2021/9/29 + , IIf(InStr("申請案號、專利號數、彼所案號", strII18) = 0 And strII18 <> "", False, True):已有抓到本所案號
   If PUB_IPDeptGetCaseNo(strText, "YOURREF", strTemp1, strTemp2, strTemp3, StrTemp4, strPA150, , , strYourII18, IIf(InStr("申請案號、專利號數、彼所案號", strII18) = 0 And strII18 <> "", False, True)) = True Then
      YourRefCase = strTemp1 & "-" & strTemp2 & "-" & strTemp3 & "-" & StrTemp4
   End If
   '2016/3/29 END
   
   If strCP01 <> "" And strCP02 <> "" Then
      'Your Ref及Our Ref同時存在時,若有FCP,FCT,CFT,CFP,FG字樣則優先考慮,否則全部歸其他
      If YourRefCase <> "" And OurRefCase <> "" And YourRefCase <> OurRefCase Then
         'Add By Sindy 2020/2/3
         If OurRefCase <> "" Then
            strCP01 = SystemNumber(OurRefCase, 1)
            strCP02 = SystemNumber(OurRefCase, 2)
            strCP03 = SystemNumber(OurRefCase, 3)
            strCP04 = SystemNumber(OurRefCase, 4)
            If LocChkCaseNoIsF(strCP01, strCP02, strCP03, strCP04) = False Then
               OurRefCase = ""
            End If
         End If
         If YourRefCase <> "" Then
            strCP01 = SystemNumber(YourRefCase, 1)
            strCP02 = SystemNumber(YourRefCase, 2)
            strCP03 = SystemNumber(YourRefCase, 3)
            strCP04 = SystemNumber(YourRefCase, 4)
            If LocChkCaseNoIsF(strCP01, strCP02, strCP03, strCP04) = False Then
               YourRefCase = ""
            End If
         End If
         '2020/2/3 END
         If YourRefCase <> "" And OurRefCase <> "" Then
            If SystemNumber(YourRefCase, 1) <> SystemNumber(OurRefCase, 1) Then
               strExc(0) = "'" & SystemNumber(OurRefCase, 1) & "'"
               strExc(1) = "'" & SystemNumber(YourRefCase, 1) & "'"
               'Modify By Sindy 2019/7/3 + ,'P'=> RE: Kind Reminder--DY/ly FW: WC/as/wfc - PRC Patent Application No. 201510131896.1; Your Ref: T2013-001-CN; Our Ref: P-122635 [RMDR.1202]
               If InStr("'FCP','FCT','CFT','CFP','FG'", strExc(0)) > 0 Then
                  strCP01 = SystemNumber(OurRefCase, 1)
                  strCP02 = SystemNumber(OurRefCase, 2)
                  strCP03 = SystemNumber(OurRefCase, 3)
                  strCP04 = SystemNumber(OurRefCase, 4)
                  strII18 = strOurII18 'Add By Sindy 2021/3/18
               ElseIf InStr("'FCP','FCT','CFT','CFP','FG'", strExc(1)) > 0 Then
                  strCP01 = SystemNumber(YourRefCase, 1)
                  strCP02 = SystemNumber(YourRefCase, 2)
                  strCP03 = SystemNumber(YourRefCase, 3)
                  strCP04 = SystemNumber(YourRefCase, 4)
                  strII18 = strYourII18 'Add By Sindy 2021/3/18
               Else
                  If InStr("'P'", strExc(0)) > 0 Then
                     strCP01 = SystemNumber(OurRefCase, 1)
                     strCP02 = SystemNumber(OurRefCase, 2)
                     strCP03 = SystemNumber(OurRefCase, 3)
                     strCP04 = SystemNumber(OurRefCase, 4)
                     strII18 = strOurII18 'Add By Sindy 2021/3/18
                  ElseIf InStr("'P'", strExc(1)) > 0 Then
                     strCP01 = SystemNumber(YourRefCase, 1)
                     strCP02 = SystemNumber(YourRefCase, 2)
                     strCP03 = SystemNumber(YourRefCase, 3)
                     strCP04 = SystemNumber(YourRefCase, 4)
                     strII18 = strYourII18 'Add By Sindy 2021/3/18
                  Else
                     'Modify By Sindy 2016/4/21
                     'PUB_IPDept_ToSortOut = "Z" '其他
                     'GoTo ChkEnd
                     strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
                     GoTo ChkOther
                     '2016/4/21 END
                  End If
               End If
            Else
               'Modify By Sindy 2016/4/21
               'PUB_IPDept_ToSortOut = "Z" '其他
               'GoTo ChkEnd
               strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
               GoTo ChkOther
               '2016/4/21 END
            End If
         'Add By Sindy 2020/2/3
         Else
            If OurRefCase <> "" Then
               strCP01 = SystemNumber(OurRefCase, 1)
               strCP02 = SystemNumber(OurRefCase, 2)
               strCP03 = SystemNumber(OurRefCase, 3)
               strCP04 = SystemNumber(OurRefCase, 4)
               strII18 = strOurII18 'Add By Sindy 2021/3/18
            ElseIf YourRefCase <> "" Then
               strCP01 = SystemNumber(YourRefCase, 1)
               strCP02 = SystemNumber(YourRefCase, 2)
               strCP03 = SystemNumber(YourRefCase, 3)
               strCP04 = SystemNumber(YourRefCase, 4)
               strII18 = strYourII18 'Add By Sindy 2021/3/18
            End If
         End If
         If strCP01 = "" And strCP02 = "" Then
            GoTo ChkOther
         End If
         '2020/2/3 END
      End If
      
      '  4. CFP、CPS：patent@taie.com.tw；
      'Modify By Sindy 2020/12/18 + CFP : 因為有FF案,外專收的CFP,內外專合作處理的案件
      ' CFP移到下列判斷
      'If strCP01 = "CFP" Or strCP01 = "CPS" Then
      If strCP01 = "CPS" Then
      '2020/12/18 END
         PUB_IPDept_ToSortOut = "4" '專利處
         strII18 = "(CPS-)" 'Add By Sindy 2022/2/14
         m_Sender = "patent" '"patent@taie.com.tw"
'      'Add By Sindy 2016/4/21
'      ElseIf strCP01 = "CFT" Then
'         PUB_IPDept_ToSortOut = "2" '外商
'         m_Sender = Pub_GetSpecMan("國外部轉信外商群組")
'      '2016/4/21 END
      Else
         'Add By Sindy 2018/4/9 +if
         'Modify By Sindy 2018/8/29 拿掉,'T','TM' ex:[Our Ref: T2000.0115/AKC/THTL/rmrs] [B&B-SGPMatters.FID344213]【T-200001 Saved】
         If InStr("'FCT','CFT','CFC','S'", "'" & strCP01 & "'") = 0 Then
         '2018/4/9 END
            '歸個案時若該案件進度智權人員都沒有Fxx人員收文時也歸其他
            strExc(0) = "select count(*) from caseprogress" & _
                        " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                        " and substr(cp12,1,1)='F'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               If Val(RsTemp.Fields(0)) <= 0 Then
                  'Modify By Sindy 2016/4/21
                  'PUB_IPDept_ToSortOut = "Z" '其他
                  'GoTo ChkEnd
                  'Modify By Sindy 2017/8/11 檢查是否為國外申請案 CF..
                  strExc(0) = "SELECT TM01 FROM TRADEMARK WHERE TM01='" & strCP01 & "' AND TM02='" & strCP02 & "' AND TM03='" & strCP03 & "' AND TM04='" & strCP04 & "' AND TM10='000'"
                  strExc(0) = strExc(0) + " union all select PA01 FROM PATENT WHERE PA01='" & strCP01 & "' AND PA02='" & strCP02 & "' AND PA03='" & strCP03 & "' AND PA04='" & strCP04 & "' AND PA09='000'"
                  strExc(0) = strExc(0) + " union all select SP01 FROM SERVICEPRACTICE WHERE SP01='" & strCP01 & "' AND SP02='" & strCP02 & "' AND SP03='" & strCP03 & "' AND SP04='" & strCP04 & "' AND SP09='000'"
                  strExc(0) = strExc(0) + " union all select LC01 FROM LAWCASE WHERE LC01='" & strCP01 & "' AND LC02='" & strCP02 & "' AND LC03='" & strCP03 & "' AND LC04='" & strCP04 & "' AND LC15='000'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 0 Then
'                     strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = "" 'Modify By Sindy 2020/12/18 mark
                     GoTo ChkOther
                  End If
                  '2017/8/11 END
                  '2016/4/21 END
                  'Add By Sindy 2017/10/25
                  '歸個案時若該案件進度承辦人都沒有Fxx人員時也歸其他
                  strExc(0) = "select count(*) from caseprogress,staff" & _
                              " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                              " and cp14=st01(+) and substr(st03,1,1)='F'"
                  intI = 1
                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If Val(RsTemp.Fields(0)) <= 0 Then
'                        strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = "" 'Modify By Sindy 2020/12/18 mark
                        GoTo ChkOther
                     End If
                  End If
                  '2017/10/25 END
               End If
            End If
         End If
         
         'Add By Sindy 2017/12/27 讀取案件資料
         strExc(0) = "SELECT TM01,TM23,TM44,TM10 FROM TRADEMARK WHERE TM01='" & strCP01 & "' AND TM02='" & strCP02 & "' AND TM03='" & strCP03 & "' AND TM04='" & strCP04 & "'"
         strExc(0) = strExc(0) + " union all select PA01,PA26,PA75,PA09 FROM PATENT WHERE PA01='" & strCP01 & "' AND PA02='" & strCP02 & "' AND PA03='" & strCP03 & "' AND PA04='" & strCP04 & "'"
         strExc(0) = strExc(0) + " union all select SP01,SP08,SP26,SP09 FROM SERVICEPRACTICE WHERE SP01='" & strCP01 & "' AND SP02='" & strCP02 & "' AND SP03='" & strCP03 & "' AND SP04='" & strCP04 & "'"
         strExc(0) = strExc(0) + " union all select LC01,LC11,LC22,LC15 FROM LAWCASE WHERE LC01='" & strCP01 & "' AND LC02='" & strCP02 & "' AND LC03='" & strCP03 & "' AND LC04='" & strCP04 & "'"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
         strNation = ""
         If intI = 1 Then
            strNation = "" & rsTmp.Fields("TM10") 'Add By Sindy 2018/3/12
            strCUNo = "" & rsTmp.Fields("TM23")
            strCUNoCU10 = GetPrjNationNumber1(strCUNo)
            strFaNo = "" & rsTmp.Fields("TM44")
            strFANoFA10 = GetPrjNationNumber(strFaNo)
            If strFaNo <> "" Then
               strArea = Mid(strFANoFA10, 1, 3)
            Else
               strArea = Mid(strCUNoCU10, 1, 3)
            End If
            strExc(0) = "select na01,na55 from nation where na01='" & strArea & "'"
            intI = 1
            Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
            If intI = 1 Then
               strAreaNA55 = "" & rsTmp.Fields("na55") 'FCT承辦智權人員
            End If
         End If
         
         'Add By Sindy 2021/3/18 CFT英國案若有歐盟案的關聯，改以歐盟案的規則抓收受者，但信件仍歸在英國案。
         'Modify By Sindy 2023/2/18 取消此規則, 直接以英國案抓收受者即可。mail:修改英國商標案件系統分信原則(February 17, 2023 5:07 PM)/(from:陳蒲璇(商標.主任.Alice))
'         If strCP01 = "CFT" And strNation = "201" Then 'CFT英國案
'            strExc(0) = "SELECT tm01,tm02,tm03,tm04,tm10 FROM caserelation1,trademark" & _
'                        " WHERE cr01='" & strCP01 & "' AND cr02='" & strCP02 & "' AND cr03='" & strCP03 & "' AND cr04='" & strCP04 & "'" & _
'                        " AND cr05=tm01(+) AND cr06=tm02(+) AND cr07=tm03(+) AND cr08=tm04(+)" & _
'                        " AND tm10='239'"
'            intI = 1
'            Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'            '歐盟案
'            If intI = 1 Then
'               str239TM01 = "" & rsTmp.Fields("tm01")
'               str239TM02 = "" & rsTmp.Fields("tm02")
'               str239TM03 = "" & rsTmp.Fields("tm03")
'               str239TM04 = "" & rsTmp.Fields("tm04")
'               strNation = "" & rsTmp.Fields("TM10")
'            End If
'         End If
         '2021/3/18 END
         
         'Add By Sindy 2018/4/10
         '目前智權人員
         If strCP01 = "FCP" Or strCP01 = "FG" Then
            strCP13 = PUB_GetFCPSalesNo(strCP01, strCP02, strCP03, strCP04)
         ElseIf strCP01 = "FCL" Or strCP01 = "LIN" Then
            strCP13 = PUB_GetFCLSalesNo(strCP01, strCP02, strCP03, strCP04)
         ElseIf strCP01 = "FCT" Then
            strCP13 = PUB_GetFCTSalesNo(strCP01, strCP02, strCP03, strCP04)
         ElseIf strCP01 = "S" Then
            If strNation = "000" Then
               strCP13 = PUB_GetFCTSalesNo(strCP01, strCP02, strCP03, strCP04)
            Else
               strCP13 = PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04)
            End If
         Else
            'Add By Sindy 2021/3/18
            If str239TM01 <> "" Then
               strCP13 = PUB_GetAKindSalesNo(str239TM01, str239TM02, str239TM03, str239TM04)
            Else
            '2021/3/18 END
               strCP13 = PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04)
            End If
         End If
         '2018/4/10 END
         
         Select Case strCP01
         '甲、有本所案號者：分類為'1'個案，收受者欄依下列條件處理：
         '  1. FCP、FG、P、PS：收受者放如後之a~b，中間以分號區隔：
         '     依外專規則抓
         '       a.案件FCP承辦業務員NA51(若FCP承辦業務員離職則抓其主管ST52)
         '       b.再抓個人主管(即a之ST52)
         '       c.抓該案之最後工程師
         '       d.再抓案件之工程師組別的系統設定管制主管
         'Modify By Sindy 2020/12/18 + CFP : 因為有FF案,外專收的CFP,內外專合作處理的案件
            Case "FCP", "FG", "P", "PS", "CFP"
               'Add By Sindy 2021/3/5 (個案,分給唐韻如) FW: [Lee & Ko IP] Korean Patent Application No. 10-2017-7027533  Your Ref: CFP-032147  Our Ref.: IPCA170958-US
               '                      應該歸入Patent信箱
               If strCP01 = "CFP" And Left(PUB_GetST03(strCP13), 1) <> "F" Then 'CFP案目前智權人員非國外部
                  GoTo ChkOther
               End If
               '2021/3/5 END
               
               PUB_IPDept_ToSortOut = "1" '個案
               '承辦組
               'strCP13 = PUB_GetFCPSalesNo(strCP01, strCP02, strCP03, strCP04)
               'Modify By Sindy 2017/3/29 + st53
               strExc(0) = "select st01,st04,st52,st53 from staff where st01='" & strCP13 & "'"
               intI = 1
               Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If rsTmp.Fields("st04") = "1" Then
                     m_Sender = m_Sender & ";" & strCP13
                  End If
                  If "" & rsTmp.Fields("st52") <> "" Then
                     'Modify By Sindy 2016/11/17 承辦組英文,日文組長的信不用加發二級主管
'                     If InStr(m_Sender, Pub_GetSpecMan("國外部轉信外專承辦英文組長")) = 0 And _
'                        InStr(m_Sender, Pub_GetSpecMan("國外部轉信外專承辦日文組長")) = 0 Then
                     'Modify By Sindy 2017/7/24 承辦組日文組長的信不用加發二級主管
                     If InStr(m_Sender, Pub_GetSpecMan("國外部轉信外專承辦日文組長")) = 0 Then
                        m_Sender = m_Sender & ";" & rsTmp.Fields("st52")
                     '2016/11/17 END
                     End If
                  End If
                  'Modify By Sindy 2017/7/24 Mark 不用加發三級主管
'                  'Add By Sindy 2017/3/29 承辦組英文,日文組長的信不用加發三級主管
'                  If "" & rsTmp.Fields("st53") <> "" Then
'                     If InStr(m_Sender, Pub_GetSpecMan("國外部轉信外專承辦英文組長")) = 0 And _
'                        InStr(m_Sender, Pub_GetSpecMan("國外部轉信外專承辦日文組長")) = 0 Then
'                        m_Sender = m_Sender & ";" & rsTmp.Fields("st53")
'                     End If
'                  End If
'                  '2017/3/29 END
                  'Modify By Sindy 2019/8/2 + and cp14<>'F4102' ex:FCP-052687(不續辦-年費),排除F4102
                  'Modify By Sindy 2022/4/11 + ,st53
                  strExc(0) = "select cp05,cp09,cp14,st02,st15,st52,st04,st16,st53 from caseprogress,staff" & _
                              " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                              " and cp14=st01(+) and cp14<>'F4102'" & _
                              " order by cp05 desc,cp67 desc"
                  intI = 1
                  Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     rsTmp.MoveFirst
                     If "" & rsTmp.Fields("st15") = "F21" Then '是工程師組再加發
                        If rsTmp.Fields("st04") = "1" Then
                           m_Sender = m_Sender & ";" & rsTmp.Fields("cp14")
                        End If
                        'Modify By Sindy 2022/5/19 加發工程師主管
                        'Modify By Sindy 2024/4/29
                        Call PUB_GetF21Manager("" & rsTmp.Fields("st16"), rsTmp.Fields("cp14"), m_Sender)
'                        strExc(10) = ""
'                        strExc(10) = PUB_GetFCPEngSup(rsTmp.Fields("cp14"))
'                        If strExc(10) <> "" Then
'                           m_Sender = m_Sender & ";" & strExc(10)
'                        End If
                        '2024/4/29 END
                        
'                        If "" & rsTmp.Fields("st52") <> "" Then
'                           m_Sender = m_Sender & ";" & rsTmp.Fields("st52")
'                        End If
'                        'Add By Sindy 2019/5/9
'                        If "" & rsTmp.Fields("st16") = "3" Then '外專日文組工程師信件，要加發副理
'                           'Add By Sindy 2022/4/11 + 外專日文組加發三級主管
'                           If "" & rsTmp.Fields("st53") <> "" Then
'                              m_Sender = m_Sender & ";" & rsTmp.Fields("st53")
'                           End If
'                           '2022/4/11 END
'                           strExc(10) = PUB_GetST70SirEmp(rsTmp.Fields("cp14"))
'                           If InStr(m_Sender, strExc(10)) = 0 Then
'                              m_Sender = m_Sender & ";" & strExc(10)
'                           End If
'                        End If
'                        '2019/5/9 END
                        '2022/5/19 END
                     End If
                  End If
      '            strExc(0) = "select cp05,cp09,cp14,st02 from caseprogress,staff" & _
      '                        " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
      '                        " and cp14=st01(+) and st15='F21' and st04='1'" & _
      '                        " order by cp05 desc,cp09 asc"
      '            intI = 1
      '            Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
      '            If intI = 1 Then
      '               rsTmp.MoveFirst
      '               m_Sender = m_Sender & ";" & rsTmp.Fields("cp14")
      '            End If
      '            If strPA150 <> "" Then
      '               m_Sender = m_Sender & ";" & IIf(strPA150 = "1", Pub_GetSpecMan("T"), IIf(strPA150 = "2", Pub_GetSpecMan("R"), IIf(strPA150 = "3", Pub_GetSpecMan("S"), Pub_GetSpecMan("T1"))))
      '            End If
               End If
               
         '  2. FCT、CFT、CFC、S、T、TM：國外部轉信外商群組所有人員；
            Case "FCT", "CFT", "CFC", "S", "T", "TM"
               PUB_IPDept_ToSortOut = "1" '個案
               'Add By Sindy 2018/4/9
               'If strCP01 = "CFT" Or strCP01 = "CFC" Then
               'Modify By Sindy 2018/3/12
               '台灣,大陸
               If strNation = "000" Or strNation = "020" Then '020==>T,020,外商收文
                  'Modify By Sindy 2017/12/27 有案號的FCT案件: 英文組依區別交給陳經理, 洪經理+區主管，日文組則交陳經理, 葉副理
                  If strArea = "011" Then '日本
                     m_Sender = m_Sender & ";" & strAreaNA55 '主管:葉副理May
                     'Add By Sindy 2020/7/20 組員休假,職代走一般規則
                     'Modify By Sindy 2024/1/31 CF組已進系統收件區，休假人員的信不直接寄至職代的系統收件區，由職代自行進入休假人員的系統收件區查看信件。
                     'Modify By Sindy 2024/6/6 CF組上線中含FC英/日主管,她們休假也確定不轉職代
                     If PUB_IPDept_IsCFMail(strAreaNA55) = False Then
                     '2024/1/31 END
                        'Modify By Sindy 2024/3/27 加傳入案號做判斷,抓職代
                        strTemp = GetCaseDutyAgent(strAreaNA55, "", False, , , , strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04)
                     End If
                     If strTemp <> "" Then
                        m_Sender = m_Sender & ";" & strTemp
                     End If
                     '2020/7/20 END
                     
                  Else '非日本,則英文組
                     'Modify By Sindy 2018/3/16
                     strCP13 = PUB_GetFCTSalesNo(strCP01, strCP02, strCP03, strCP04)
                     m_Sender = strCP13
                     'Add By Sindy 2020/7/20 組員休假,職代走一般規則
                     'Modify By Sindy 2024/1/31 CF組已進系統收件區，休假人員的信不直接寄至職代的系統收件區，由職代自行進入休假人員的系統收件區查看信件。
                     'Modify By Sindy 2024/6/6 CF組上線中含FC英/日主管,她們休假也確定不轉職代
                     If PUB_IPDept_IsCFMail(strCP13) = False Then
                     '2024/1/31 END
                        'Modify By Sindy 2024/3/27 加傳入案號做判斷,抓職代
                        strTemp = GetCaseDutyAgent(strCP13, "", False, , , , strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04)
                     End If
                     If strTemp <> "" Then
                        m_Sender = m_Sender & ";" & strTemp
                     End If
                     '2020/7/20 END
                     
                     'Add By Sindy 2020/7/16 (外商英文組案件分信) 案件INBOUND 轉寄email時增列英文承辦區主管為收件人
                     strTemp = GetFLOW001Person(strCP13, "3", , 1) '3.接洽單
                     If strTemp <> "" Then
                        m_Sender = m_Sender & ";" & strTemp
                        'Add By Sindy 2020/7/20 主管休假,指定抓人事職代
                        'Modify By Sindy 2024/1/31 CF組已進系統收件區，休假人員的信不直接寄至職代的系統收件區，由職代自行進入休假人員的系統收件區查看信件。
                        'Modify By Sindy 2024/6/6 CF組上線中含FC英/日主管,她們休假也確定不轉職代
                        If PUB_IPDept_IsCFMail(strTemp) = False Then
                        '2024/1/31 END
                           'Modify By Sindy 2024/3/27 加傳入案號做判斷,抓職代
                           strTemp = GetCaseDutyAgent(strTemp, "", False, , , "1", strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04)
                        End If
                        If strTemp <> "" Then
                           m_Sender = m_Sender & ";" & strTemp
                        End If
                        '2020/7/20 END
                     End If
                     '2020/7/16 END
                     
                     'Modify By Sindy 2021/6/17 加發 2,3,4級主管
                     strExc(0) = "select st01,st52,st53,st54 from staff where st01='" & strCP13 & "'"
                     intI = 1
                     Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
                     If intI = 1 Then
                        If "" & rsTmp.Fields("st52") <> "" Then
                           m_Sender = m_Sender & ";" & rsTmp.Fields("st52")
                           '主管休假,指定抓人事職代
                           'Modify By Sindy 2024/1/31 CF組已進系統收件區，休假人員的信不直接寄至職代的系統收件區，由職代自行進入休假人員的系統收件區查看信件。
                           'Modify By Sindy 2024/6/6 CF組上線中含FC英/日主管,她們休假也確定不轉職代
                           If PUB_IPDept_IsCFMail(rsTmp.Fields("st52")) = False Then
                           '2024/1/31 END
                              'Modify By Sindy 2024/3/27 加傳入案號做判斷,抓職代
                              strTemp = GetCaseDutyAgent(rsTmp.Fields("st52"), "", False, , , "1", strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04)
                           End If
                           If strTemp <> "" Then
                              m_Sender = m_Sender & ";" & strTemp
                           End If
                        End If
                        If "" & rsTmp.Fields("st53") <> "" Then
                           m_Sender = m_Sender & ";" & rsTmp.Fields("st53")
                           '主管休假,指定抓人事職代
                           'Modify By Sindy 2024/1/31 CF組已進系統收件區，休假人員的信不直接寄至職代的系統收件區，由職代自行進入休假人員的系統收件區查看信件。
                           'Modify By Sindy 2024/6/6 CF組上線中含FC英/日主管,她們休假也確定不轉職代
                           If PUB_IPDept_IsCFMail(rsTmp.Fields("st53")) = False Then
                           '2024/1/31 END
                              'Modify By Sindy 2024/3/27 加傳入案號做判斷,抓職代
                              strTemp = GetCaseDutyAgent(rsTmp.Fields("st53"), "", False, , , "1", strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04)
                           End If
                           If strTemp <> "" Then
                              m_Sender = m_Sender & ";" & strTemp
                           End If
                        End If
                        If "" & rsTmp.Fields("st54") <> "" Then
                           m_Sender = m_Sender & ";" & rsTmp.Fields("st54")
                           '主管休假,指定抓人事職代
                           'Modify By Sindy 2024/1/31 CF組已進系統收件區，休假人員的信不直接寄至職代的系統收件區，由職代自行進入休假人員的系統收件區查看信件。
                           'Modify By Sindy 2024/6/6 CF組上線中含FC英/日主管,她們休假也確定不轉職代
                           If PUB_IPDept_IsCFMail(rsTmp.Fields("st54")) = False Then
                           '2024/1/31 END
                              'Modify By Sindy 2024/3/27 加傳入案號做判斷,抓職代
                              strTemp = GetCaseDutyAgent(rsTmp.Fields("st54"), "", False, , , "1", strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04)
                           End If
                           If strTemp <> "" Then
                              m_Sender = m_Sender & ";" & strTemp
                           End If
                        End If
                     End If
                     '2021/6/17 END
                     
                     'Modify By Sindy 2021/6/17 Mark, 外商分信經理=洪經理
'                     strExc(0) = "select st01 from staff where st04='1' and st05='26' and st16='2'"
'                     intI = 1
'                     Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        m_Sender = m_Sender & ";" & rsTmp.Fields("st01") '主管:洪經理
'                        'Add By Sindy 2020/7/20 主管休假,指定抓人事職代
'                        strTemp = GetCaseDutyAgent(rsTmp.Fields("st01"), "", False, , , "1")
'                        If strTemp <> "" Then
'                           m_Sender = m_Sender & ";" & strTemp
'                        End If
'                        '2020/7/20 END
'                     End If
                  End If
'                  m_Sender = m_Sender & ";" & 外商分信經理
'                  'Add By Sindy 2020/7/20 主管休假,指定抓人事職代
'                  strTemp = GetCaseDutyAgent(外商分信經理, "", False, , , "1")
'                  If strTemp <> "" Then
'                     m_Sender = m_Sender & ";" & strTemp
'                  End If
'                  '2020/7/20 END
               
               '其他國家
               Else
                  'Add By Sindy 2021/3/18
                  If str239TM01 <> "" Then
                     Call GetNA69("", strNation, strCP13, strEmp, str239TM01, str239TM02, str239TM03, str239TM04)
                  Else
                  '2021/3/18 END
                     Call GetNA69("", strNation, strCP13, strEmp, strCP01, strCP02, strCP03, strCP04)
                  End If
                  m_Sender = strEmp 'Add By Sindy 2021/6/15
                  'Add By Sindy 2020/7/20 組員休假,職代走一般規則
                  'Modify By Sindy 2024/1/31 CF組已進系統收件區，休假人員的信不直接寄至職代的系統收件區，由職代自行進入休假人員的系統收件區查看信件。
                  'Modify By Sindy 2024/6/6 CF組上線中含FC英/日主管,她們休假也確定不轉職代
                  If PUB_IPDept_IsCFMail(strEmp) = False Then
                  '2024/1/31 END
                     'Modify By Sindy 2024/3/27 加傳入案號做判斷,抓職代
                     strTemp = GetCaseDutyAgent(strEmp, "", False, , , , strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04)
                  End If
                  If strTemp <> "" Then
                     m_Sender = m_Sender & ";" & strTemp
                  End If
                  '2020/7/20 END
                  
                  'Modify By Sindy 2021/6/17 加發 2,3,4級主管
                  strExc(0) = "select st01,st52,st53,st54 from staff where st01='" & strEmp & "'"
                  intI = 1
                  Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If "" & rsTmp.Fields("st52") <> "" Then
                        m_Sender = m_Sender & ";" & rsTmp.Fields("st52")
                        '主管休假,指定抓人事職代
                        'Modify By Sindy 2024/1/31 CF組已進系統收件區，休假人員的信不直接寄至職代的系統收件區，由職代自行進入休假人員的系統收件區查看信件。
                        'Modify By Sindy 2024/6/6 CF組上線中含FC英/日主管,她們休假也確定不轉職代
                        If PUB_IPDept_IsCFMail(rsTmp.Fields("st52")) = False Then
                        '2024/1/31 END
                           'Modify By Sindy 2024/3/27 加傳入案號做判斷,抓職代
                           strTemp = GetCaseDutyAgent(rsTmp.Fields("st52"), "", False, , , "1", strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04)
                        End If
                        If strTemp <> "" Then
                           m_Sender = m_Sender & ";" & strTemp
                        End If
                     End If
                     If "" & rsTmp.Fields("st53") <> "" Then
                        m_Sender = m_Sender & ";" & rsTmp.Fields("st53")
                        '主管休假,指定抓人事職代
                        'Modify By Sindy 2024/1/31 CF組已進系統收件區，休假人員的信不直接寄至職代的系統收件區，由職代自行進入休假人員的系統收件區查看信件。
                        'Modify By Sindy 2024/6/6 CF組上線中含FC英/日主管,她們休假也確定不轉職代
                        If PUB_IPDept_IsCFMail(rsTmp.Fields("st53")) = False Then
                        '2024/1/31 END
                           'Modify By Sindy 2024/3/27 加傳入案號做判斷,抓職代
                           strTemp = GetCaseDutyAgent(rsTmp.Fields("st53"), "", False, , , "1", strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04)
                        End If
                        If strTemp <> "" Then
                           m_Sender = m_Sender & ";" & strTemp
                        End If
                     End If
                     If "" & rsTmp.Fields("st54") <> "" Then
                        m_Sender = m_Sender & ";" & rsTmp.Fields("st54")
                        '主管休假,指定抓人事職代
                        'Modify By Sindy 2024/1/31 CF組已進系統收件區，休假人員的信不直接寄至職代的系統收件區，由職代自行進入休假人員的系統收件區查看信件。
                        'Modify By Sindy 2024/6/6 CF組上線中含FC英/日主管,她們休假也確定不轉職代
                        If PUB_IPDept_IsCFMail(rsTmp.Fields("st54")) = False Then
                        '2024/1/31 END
                           'Modify By Sindy 2024/3/27 加傳入案號做判斷,抓職代
                           strTemp = GetCaseDutyAgent(rsTmp.Fields("st54"), "", False, , , "1", strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04)
                        End If
                        If strTemp <> "" Then
                           m_Sender = m_Sender & ";" & strTemp
                        End If
                     End If
                  End If
                  '2021/6/17 END
                  
'                  '主管
'                  'Modify By Sindy 2021/6/15
'                  'strExc(0) = "select st01,st52 from staff where st01='" & m_Sender & "' and st52 is not null"
'                  strExc(0) = "select st01,st52 from staff where st01='" & strEmp & "' and st52 is not null"
'                  '2021/6/15 END
'                  intI = 1
'                  Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     m_Sender = m_Sender & ";" & rsTmp.Fields("st52")
'                     'Add By Sindy 2020/7/20 主管休假,指定抓人事職代
'                     strTemp = GetCaseDutyAgent(rsTmp.Fields("st52"), "", False, , , "1")
'                     If strTemp <> "" Then
'                        m_Sender = m_Sender & ";" & strTemp
'                     End If
'                     '2020/7/20 END
'                  End If
'                  m_Sender = m_Sender & ";" & 外商分信經理
'                  'Add By Sindy 2020/7/20 主管休假,指定抓人事職代
'                  strTemp = GetCaseDutyAgent(外商分信經理, "", False, , , "1")
'                  If strTemp <> "" Then
'                     m_Sender = m_Sender & ";" & strTemp
'                  End If
'                  '2020/7/20 END
               End If
               '2018/4/9 END
               If m_Sender = "" Then '個案:上列條件未抓到相關人員時,則抓外商群組
                  m_Sender = Pub_GetSpecMan("國外部轉信外商群組")
               End If
               '2018/3/12 END
               
               'Modify By Sindy 2018/3/12 Mark
'               'Modify By Sindy 2017/12/27
'               '有案號的FCT案件: 英文組依區別交給陳經理, 洪經理+區主管，日文組則交陳經理, 葉副理
'               If strSrvDate(1) >= 20180102 And strCP01 = "FCT" Then
'                  m_Sender = Pub_GetSpecMan("V") '陳經理
'                  If strArea = "011" Then '日本
'                     m_Sender = m_Sender & ";" & strAreaNA55 '葉副理
'                  Else '非日本,則英文組
'                     strExc(0) = "select st01 from staff where st04='1' and st05='26' and st16='2'"
'                     intI = 1
'                     Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        m_Sender = m_Sender & ";" & rsTmp.Fields("st01") '洪經理
'                     End If
'                     m_Sender = m_Sender & ";" & strAreaNA55 '區主管
'                  End If
'               Else
'               '2017/12/27 END
'                  m_Sender = Pub_GetSpecMan("國外部轉信外商群組")
'               End If
               
         '  3. FCL、CFL、LIN：抓該案號最大收文日最小收文號之
         '                    承辦人之ST16若為2則設定為國外部轉信外法日文組群組
         '                                   否則設定為國外部轉信外法英文組群組
         '                    再判斷該收文號的業務區CP12，若為F1X部門者則收受者再加國外部轉信外商群組所有人員；
         '                                                若為F2X部門者則收受者再加如FCP案之a~b；
            Case "FCL", "CFL", "LIN"
               PUB_IPDept_ToSortOut = "1" '個案
               'Modify By Sindy 2020/1/16
               '經討論後，請電腦中心協助，對於經由ipdept收、發信件之FCL/LIN案件，請仿照目前國外部FCP分信方式，
               '若系統辨識到主旨含FCL或LIN，請依照系統內建檔資料，把信正本轉給承辦人員/協辦人員/承辦業務；
               '副本：所長/承辦業務直屬主管，
               '以先前FCL-10858為例，當ipdept收到信時，請依系統最近一次程序的承辦人員/協辦人員/承辦業務轉信如下：
               '正本：何律師/Daniel/Jay(外商承辦)
               '副本：所長/Cary(Jay直屬主管)
                              
               'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
               'Modify By Sindy 2020/1/16 +,cp29
               strExc(0) = "select cp05,cp09,cp12,cp14" & _
                           ",s1.st02,s1.st16 s1_ST16,s1.st04 s1_ST04,s1.st52 s1_ST52" & _
                           ",cp29,s2.st02,s2.st04 s2_ST04,s2.st52 s2_ST52" & _
                           " from caseprogress,staff s1,staff s2" & _
                           " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                           " and cp14=s1.st01(+) and cp29=s2.st01(+) and cp14<>'F4102'" & _
                           " order by cp05 desc,cp09 asc"
               intI = 1
               Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  rsTmp.MoveFirst
                  '承辦人員
                  If rsTmp.Fields("s1_ST04") = "1" Then
                     m_Sender = m_Sender & ";" & rsTmp.Fields("cp14")
                  End If
                  If "" & rsTmp.Fields("s1_ST52") <> "" Then
                     strCySender = strCySender & ";" & rsTmp.Fields("s1_ST52")
                  End If
                  '法務協辦人員
                  If Not IsNull(rsTmp.Fields("cp29")) Then
                     If rsTmp.Fields("s2_ST04") = "1" Then
                        m_Sender = m_Sender & ";" & rsTmp.Fields("cp29")
                     End If
                     If "" & rsTmp.Fields("s2_ST52") <> "" Then
                        strCySender = strCySender & ";" & rsTmp.Fields("s2_ST52")
                     End If
                  End If
                  '智權人員
                  strExc(0) = "select st01,st04,st52 from staff where st01='" & strCP13 & "'"
                  intI = 1
                  Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
                  If intI = 1 Then
                     If rsTmp.Fields("st04") = "1" Then
                        m_Sender = m_Sender & ";" & strCP13
                     End If
                     If "" & rsTmp.Fields("st52") <> "" Then
                        strCySender = strCySender & ";" & rsTmp.Fields("st52")
                     End If
                  End If
                  m_Sender = m_Sender & strCySender
               '2020/1/16 END
'                  If "" & rsTmp.Fields("s1_ST16") = "2" Then
'                     m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外法日文組群組")
'                  Else
'                     m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外法英文組群組")
'                  End If
'                  If Mid("" & rsTmp.Fields("cp12"), 1, 2) = "F1" Then
'                     m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外商群組")
'                  ElseIf Mid("" & rsTmp.Fields("cp12"), 1, 2) = "F2" Then
'         '       a.案件FCP承辦業務員NA51(若FCP承辦業務員離職則抓其主管ST52)
'         '       b.再抓個人主管(即a之ST52)
'                     'strCP13 = PUB_GetFCLSalesNo(strCP01, strCP02, strCP03, strCP04)
'                     strExc(0) = "select st01,st04,st52 from staff where st01='" & strCP13 & "'"
'                     intI = 1
'                     Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'                     If intI = 1 Then
'                        If rsTmp.Fields("st04") = "1" Then
'                           m_Sender = m_Sender & ";" & strCP13
'                        End If
'                        If "" & rsTmp.Fields("st52") <> "" Then
'                           m_Sender = m_Sender & ";" & rsTmp.Fields("st52")
'                        End If
'                     End If
'                  End If
               End If
               
               'Modify By Sindy 2020/7/16 FCL,LIN,CFL要檢查是否有案源檔, 加發最大收文時間的介紹人
               strExc(0) = "SELECT cp01,cp02,cp03,cp04,cp09,los06,los04" & _
                           " From caseprogress,lawofficesource" & _
                           " Where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                           " AND cp09=los06(+) AND los06 is not null AND los04 is not null" & _
                           " ORDER BY cp66 desc,cp67 desc"
               intI = 1
               Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  m_Sender = m_Sender & ";" & Replace(rsTmp.Fields("los04"), ",", ";")
               End If
               '2020/7/16 END
         End Select
      End If
   End If
End If 'Add By Sindy 2019/8/16

'*************************
ChkOther:
'*************************
   If bolOnlyReadCaseNo = True Then Exit Function 'Add By Sindy 2022/8/5 僅為讀取案號
   
   If PUB_IPDept_ToSortOut <> "" And m_Sender = "" Then PUB_IPDept_ToSortOut = "" 'Add By Sindy 2018/3/12
   'Add By Sindy 2020/12/18
   If strCP01 = "CFP" And PUB_IPDept_ToSortOut = "" Then
      PUB_IPDept_ToSortOut = "4" '專利處
      strII18 = strII18 & "(CFP-)" 'Add By Sindy 2022/2/14
      m_Sender = "patent" '"patent@taie.com.tw"
   ElseIf PUB_IPDept_ToSortOut = "" Then
      strII18 = "" 'Add By Sindy 2021/3/18
   End If
   '2020/12/18 END
   'Add By Sindy 2023/5/4 ex:自?回复:PY/py : 請辦理申請人更名 （共7件）- FMP專利案; PRC Patent Application No. 201680073087.2; Your Ref: PI-180190/1V; Our Ref: P-12
   If PUB_IPDept_ToSortOut = "" And strCP02 <> "" Then
      strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
   End If
   '2023/5/4 END
   
   'If PUB_IPDept_ToSortOut = "" Then
   '   'Add By Sindy 2016/5/12 直接轉信給靜芳
   ''   @sino-elite-ip.com(寰華)
   ''   @sunyu.com(舜禹)
   ''   @jnkip.com(捷恩凱)
   '   If InStr(UCase(strII11), UCase("@sino-elite-ip.com")) > 0 Or _
   '      InStr(UCase(strII11), UCase("@sunyu.com")) > 0 Or _
   '      InStr(UCase(strII11), UCase("@jnkip.com")) > 0 Then
   '      PUB_IPDept_ToSortOut = "3" '外專
   '      m_Sender = Pub_GetSpecMan("N") 'N:73023
   '   End If
   '   '2016/5/12 END
   'End If
   
   '乙、無本所案號者：檔名依下列字串分類：
   If PUB_IPDept_ToSortOut = "" Then
      'Modify By Sindy 2019/9/3 依外專承辦英文縮寫分信
      Call PUB_IPDept_ToSortOutSub_F23(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "1", bolF23EngW, , , strII18)
   End If
   
'Add By Sindy 2019/8/16 未傳遞的主旨略過
'Modify By Sindy 2020/8/21 FW: CH/jf - KEEN Counterfeits (Our Ref.: FCT-25071 and others)
'                          David反應[未自動分類]:這種主旨可以判斷英文縮寫
'If Trim(strII11) <> "未傳遞的主旨" Then
'2019/8/16 END
   '********************************************************************************************************
   'Add By Sindy 2016/5/25 個案的例外檢查(關鍵字索引)
   'ex.From: huanhua [mailto:huanhua@sino-elite-ip.com]
   '   Subject: ??公布通知? Y/R: P-113774; O/R: PI-160089/1V(江如玉)
   '********************************************************************************************************
   'Modify By Sindy 2016/5/30 再調整位置(若有外專承辦組縮寫就成立)
   'ex.From: daisyzou@sunyu.com [mailto:daisyzou@sunyu.com]
   '   Subject: Re: FW: =急件= <AXIS>新案翻譯 P-1500. P-1501 (共兩件)
   If PUB_IPDept_ToSortOut = "" Or bolF23EngW = False Then '非外專承辦人的英文縮寫
      'Modify By Sindy 2020/3/6 將索引關鍵字改為共用函數
      'Modify By Sindy 2024/8/28 strSubject 調整為 strText
      strTmp = PUB_FindIPdeptKeyWord(" and lk11='Y' and LK12='F'", strText, strII11, strEmpSender, strII18)
      'Modify By Sindy 2020/3/18
      '  strSubject=FW: ◎FCT案     FW: ◎RE:台?商標出願No.108065411　（Our Ref.: F19T12015　Your Ref.: FCP-043896）【FCP-043896 Saved】【FCP-043896 Saved】
      '  strII11=陳毓芳(Elaine.主任)
      'If strTmp <> "" And strEmpSender <> "" Then
      If strTmp <> "" Then
      '2020/3/18 END
         'Add By Sindy 2022/8/10 敏莉說:寰華的來信(寄件者是huanhua <huanhua@sino-elite-ip.com>的郵件)請直接匯到程序管制人員，
         '除非解析不到本所案號，則維持系統設定的人員，由人工轉發
         strExc(10) = ""
         If InStr(UCase(strII11), UCase("huanhua@sino-elite-ip.com")) > 0 And strCP01 <> "" And strCP02 <> "" Then
            strExc(10) = PUB_GetFCPHandler(strCP01, strCP02, strCP03, strCP04) 'FCP程序(管制)人員
         End If
         If strExc(10) <> "" Then
            m_Sender = strExc(10)
         Else
         '2022/8/10 END
            m_Sender = strEmpSender
         End If
         PUB_IPDept_ToSortOut = strTmp
      End If
      '2020/3/6 END
   End If
   '2016/5/25 END
'End If 'Add By Sindy 2019/8/16 + 'Modify By Sindy 2020/8/21 Mark
   
   If PUB_IPDept_ToSortOut = "" Then
      'Modify By Sindy 2019/8/30 依外專程序英文縮寫分信
      Call PUB_IPDept_ToSortOutSub_F22(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "1", "DY", , , strII18)
      'Add By Sindy 2020/10/6
      If PUB_IPDept_ToSortOut = "" Then
         Call PUB_IPDept_ToSortOutSub_F22(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "1", "WW", , , strII18)
      End If
      '2020/10/6 END
   End If
   
   'Add By Sindy 2018/1/4
   If PUB_IPDept_ToSortOut = "" Then
      'Modify By Sindy 2019/8/30 依國外業務拓展英文縮寫分信
      Call PUB_IPDept_ToSortOutSub_F41(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "1", , , strII18)
   End If
   '2018/1/4 END
   
   ''Add By Sindy 2016/4/14
   'If PUB_IPDept_ToSortOut = "" Then
   '   'AC/xx:寰華案件
   '   Call ToSortOutSub_AC(strText, m_Sender, PUB_IPDept_ToSortOut)
   'End If
   ''2016/4/14 END
   
   '***外專工程師:ex.sh  主管+/+個人
   If PUB_IPDept_ToSortOut = "" Then
      'Modify By Sindy 2019/8/30 依外專工程師英文縮寫分信
      Call PUB_IPDept_ToSortOutSub_F21(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "1", , , strII18)
   End If
   
   '*** 外商:主管洪琬姿請假時,英文組承辦人代號AH/會改為FC/
   'Modify By Sindy 2017/12/27
   'If PUB_IPDept_ToSortOut = "" Then
   If PUB_IPDept_ToSortOut = "" Then 'Or (strCP01 = "CFT" Or strCP01 = "CFC" Or strCP01 = "S")
      'Modify By Sindy 2019/8/30 依外商英文縮寫分信
      Call PUB_IPDept_ToSortOutSub_F1x(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "1", , , strII18)
   End If
   
   '*** 法務:  主管+/+個人
   If PUB_IPDept_ToSortOut = "" Then
      'Modify By Sindy 2019/8/30 依外法英文縮寫分信
      Call PUB_IPDept_ToSortOutSub_FLaw(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "1", strII18)
   End If
   
'   '*** 法務:  個人+/
'   If PUB_IPDept_ToSortOut = "" Then
'      'Modify By Sindy 2019/8/30 依外法個人英文縮寫分信
'      Call PUB_IPDept_ToSortOutSub_FLawEmp(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, strII18)
'   End If
   
     
   'Modify By Sindy 2022/7/8
'***************************************
'前面先檢查在職人員, 這裡再統一檢查離職人員
'***************************************
   '*** 依外專承辦英文縮寫分信
   If PUB_IPDept_ToSortOut = "" Then
      Call PUB_IPDept_ToSortOutSub_F23(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "2", , , , strII18)
   End If
   
   '*** 依外專程序英文縮寫分信
   If PUB_IPDept_ToSortOut = "" Then
      Call PUB_IPDept_ToSortOutSub_F22(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "2", "DY", , , strII18)
      If PUB_IPDept_ToSortOut = "" Then
         Call PUB_IPDept_ToSortOutSub_F22(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "2", "WW", , , strII18)
      End If
   End If
   
   '*** 依國外業務拓展英文縮寫分信
   If PUB_IPDept_ToSortOut = "" Then
      Call PUB_IPDept_ToSortOutSub_F41(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "2", , , strII18)
   End If
   
   '*** 依外專工程師英文縮寫分信
   If PUB_IPDept_ToSortOut = "" Then
      Call PUB_IPDept_ToSortOutSub_F21(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "2", , , strII18)
   End If
   
   '*** 依外商英文縮寫分信
   If PUB_IPDept_ToSortOut = "" Then
      Call PUB_IPDept_ToSortOutSub_F1x(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "2", , , strII18)
   End If
   
   '*** 依外法英文縮寫分信
   If PUB_IPDept_ToSortOut = "" Then
      Call PUB_IPDept_ToSortOutSub_FLaw(strText, strII11, m_Sender, PUB_IPDept_ToSortOut, "2", strII18)
   End If
'***************************************
   
'Add By Sindy 2019/8/16 未傳遞的主旨略過
'Modify By Sindy 2024/10/16 + Or InStr(strSubject, "台一關係企業財務信箱變更通知") > 0
If Trim(strII11) <> "未傳遞的主旨" Or InStr(strSubject, "台一關係企業財務信箱變更通知") > 0 Then
'2019/8/16 END
   '********************************************************************************************************
   'Modify By Sindy 2016/5/20 改關鍵字存放Table
   '關鍵字索引
   '********************************************************************************************************
   If PUB_IPDept_ToSortOut = "" Then
      m_Sender = ""
      'Modify By Sindy 2020/3/6 將索引關鍵字改為共用函數
'     Modify By Sindy 2024/5/13
'     113/5/9 David和楊雯芳經理雙方討論好,關鍵字分信規則,調整如下:
'        1. 先歸 編列為非 6.新知 的關鍵字
'        2. 第1點沒找到, 再找 6.新知 的關鍵字
      'PUB_IPDept_ToSortOut = PUB_FindIPdeptKeyWord(" and LK12='F' and LK02 not in('6','8')", strSubject, strII11, m_Sender, strII18)
      'Modify By Sindy 2024/8/28 strSubject 調整為 strText
      PUB_IPDept_ToSortOut = PUB_FindIPdeptKeyWord(" and LK12='F' and LK02 not in('6')", strText, strII11, m_Sender, strII18)
      '2020/3/6 END
   End If
   
   '丙、3. 若無上述二項則以系統類別分類；
   '       a. FCP、FG、P、PS：國外部轉信外專群組
   '       b. FCT、CFT、CFC、S、T、TM：國外部轉信外商群組
   '       c. FCL、CFL、LIN：國外部轉信外法英文組群組；   2016/4/1 改國外部轉信外法群組
   '       c. CFP、CPS：patent@taie.com.tw
   '丁、無系統類別時
   '       Trademark及CTM歸國外部轉信外商群組；
   '       Payment advice、Foreign Payment Notification、Remittance Advice歸財務account@taie.com.tw；
   '       Newsletter、IPO Daily News或網域@aipla.org、@inta.org、@ipo.org歸新知81040(閻副所長);國外資訊分享區(EXTERNAL_NEWS@taie.com.tw)；
   If PUB_IPDept_ToSortOut = "" Then
   '********************************************************************************************************
   '檢查系統別:
   '********************************************************************************************************
      '系統別 3 碼
      'If InStr(strText, " FCP-") > 0 Or Mid(strText, 1, Len("FCP-")) = "FCP-" Then
      If InStr(strText, "FCP-") > 0 Then
         'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "3" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
         PUB_IPDept_ToSortOut = "3" '外專
         strII18 = "(FCP-)" 'Add By Sindy 2022/2/24
         m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外專群組")
   '   If InStr(strText, " FCT-") > 0 Or Mid(strText, 1, Len("FCT-")) = "FCT-" Or _
   '      InStr(strText, " CFT-") > 0 Or Mid(strText, 1, Len("CFT-")) = "CFT-" Or _
   '      InStr(strText, " CFC-") > 0 Or Mid(strText, 1, Len("CFC-")) = "CFC-" Then
      ElseIf InStr(strText, "FCT-") > 0 Or _
         InStr(strText, "CFT-") > 0 Or _
         InStr(strText, "CFC-") > 0 Then
         'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "2" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
         PUB_IPDept_ToSortOut = "2" '外商
         'Add By Sindy 2022/2/24
         If InStr(strText, "FCT-") > 0 Then
            strII18 = "(FCT-)"
         ElseIf InStr(strText, "CFT-") > 0 Then
            strII18 = "(CFT-)"
         Else
            strII18 = "(CFC-)"
         End If
         '2022/2/24 END
         m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外商群組")
   '   If InStr(strText, " FCL-") > 0 Or Mid(strText, 1, Len("FCL-")) = "FCL-" Or _
   '      InStr(strText, " CFL-") > 0 Or Mid(strText, 1, Len("CFL-")) = "CFL-" Or _
   '      InStr(strText, " LIN-") > 0 Or Mid(strText, 1, Len("LIN-")) = "LIN-" Then
      ElseIf InStr(strText, "FCL-") > 0 Or _
         InStr(strText, "CFL-") > 0 Or _
         InStr(strText, "LIN-") > 0 Then
         'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "5" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
         PUB_IPDept_ToSortOut = "5" '外法
         'Add By Sindy 2022/2/24
         If InStr(strText, "FCL-") > 0 Then
            strII18 = "(FCL-)"
         ElseIf InStr(strText, "CFL-") > 0 Then
            strII18 = "(CFL-)"
         Else
            strII18 = "(LIN-)"
         End If
         '2022/2/24 END
         'modify by sonia 2016/4/1 改國外部轉信外法群組
         'm_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外法英文組群組") & ";99021"
         m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外法群組") & ";" & Pub_GetSpecMan("國外部轉信外專承辦日文組長") '99021
   '   If InStr(strText, " CFP-") > 0 Or Mid(strText, 1, Len("CFP-")) = "CFP-" Or _
   '      InStr(strText, " CPS-") > 0 Or Mid(strText, 1, Len("CPS-")) = "CPS-" Then
      ElseIf InStr(strText, "CFP-") > 0 Or _
         InStr(strText, "CPS-") > 0 Then
         'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "4" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
         PUB_IPDept_ToSortOut = "4" '專利處
         'Add By Sindy 2022/2/14
         If InStr(strText, "CFP-") > 0 Then
            strII18 = "(CFP-)"
         Else
            strII18 = "(CPS-)"
         End If
         '2022/2/14 END
         m_Sender = m_Sender & ";patent" 'patent@taie.com.tw"
      '系統別 2 碼
      ElseIf InStr(strText, " FG-") > 0 Or Mid(strText, 1, Len("FG-")) = "FG-" Or _
         InStr(strText, " PS-") > 0 Or Mid(strText, 1, Len("PS-")) = "PS-" Then
   '   ElseIf InStr(strText, "FG-") > 0 Or _
   '      InStr(strText, "PS-") > 0 Then
         'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "3" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
         PUB_IPDept_ToSortOut = "3" '外專
         'Add By Sindy 2022/2/24
         If InStr(strText, " FG-") > 0 Or Mid(strText, 1, Len("FG-")) = "FG-" Then
            strII18 = "(FG-)"
         Else
            strII18 = "(PS-)"
         End If
         '2022/2/24 END
         m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外專群組")
      ElseIf InStr(strText, " TM-") > 0 Or Mid(strText, 1, Len("TM-")) = "TM-" Then
      'ElseIf InStr(strText, "TM-") > 0 Then
         'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "2" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
         PUB_IPDept_ToSortOut = "2" '外商
         strII18 = "(TM-)" 'Add By Sindy 2022/2/24
         m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外商群組")
      '系統別 1 碼
      ElseIf InStr(strText, " P-") > 0 Or Mid(strText, 1, Len("P-")) = "P-" Then
   '   ElseIf InStr(strText, "P-") > 0 Then
         'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "3" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
         PUB_IPDept_ToSortOut = "3" '外專
         strII18 = "(P-)" 'Add By Sindy 2022/2/24
         m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外專群組")
      'Modify By Sindy 2016/5/9 ex.PCT-application, national phase in China, our Ref 101308 WO-CN ==>分錯給外商
      ElseIf InStr(strText, " S-") > 0 Or Mid(strText, 1, Len("S-")) = "S-" Or _
         InStr(strText, " T-") > 0 Or Mid(strText, 1, Len("T-")) = "T-" Then
   '   ElseIf InStr(strText, "S-") > 0 Or _
   '      InStr(strText, "T-") > 0 Then
      '2016/5/9 END
         'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "2" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
         PUB_IPDept_ToSortOut = "2" '外商
         'Add By Sindy 2022/2/24
         If InStr(strText, " S-") > 0 Or Mid(strText, 1, Len("S-")) = "S-" Then
            strII18 = "(S-)"
         Else
            strII18 = "(T-)"
         End If
         '2022/2/24 END
         m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外商群組")
      End If
   End If
   
   '********************************************************************************************************
   'Modify By Sindy 2017/4/14 + 副所長:最後再分新知,開拓信件
'   Modify By Sindy 2024/5/13
'   113/5/9 David和楊雯芳經理雙方討論好,關鍵字分信規則,調整如下:
'      1. 先歸 編列為非 6.新知 的關鍵字
'      2. 第1點沒找到, 再找 6.新知 的關鍵字
   '********************************************************************************************************
   If PUB_IPDept_ToSortOut = "" Then
      m_Sender = ""
      'Modify By Sindy 2020/3/6 將索引關鍵字改為共用函數
      'PUB_IPDept_ToSortOut = PUB_FindIPdeptKeyWord(" and LK12='F' and LK02 in('6','8')", strSubject, strII11, m_Sender, strII18)
      'Modify By Sindy 2024/8/28 strSubject 調整為 strText
      PUB_IPDept_ToSortOut = PUB_FindIPdeptKeyWord(" and LK12='F' and LK02 in('6')", strText, strII11, m_Sender, strII18)
      '2020/3/6 END
   End If
   
   'Add By Sindy 2020/3/30 Lawoffice要整併到Ipdept
   '有中文字無含unicode時,則發給 國內信件管理人員/Z.其他 去做分信處理
   If PUB_IPDept_ToSortOut = "" Then
      'Add By Sindy 2020/11/4 FW: ◎REGISTRATION FOR PROTECTION OF INTELLECTUAL PROPERTY RIGHTS IN TAIWAN [寄件者：北所(九樓)#410]
      '主旨截取到”[寄件者”之前即可
      strText = strSubject
      If InStr(strText, "[寄件者") > 0 Then
         strText = Mid(strText, 1, InStr(strText, "[寄件者") - 1)
      End If
      '2020/11/4 END
      'Add By Sindy 2022/4/11 自動回覆: Check out our new INTAEvents listing
      If Left(strText, 5) = "自動回覆:" Then
         strText = Trim(Mid(strText, 6))
      End If
      '2022/4/11 END
      
      'Modify By Sindy 2020/5/14 先過濾主旨文字
      'vbNarrow:轉換字串中寬的(雙位元)字元為窄的(單位元)字元；適用於遠東地區。
      '＜Password＞Out of Service of Our Facsimile System ==>
      '<Password>Out of Service of Our Facsimile System : 全形符號轉半形符號
      strChkSubject = ""
      For jj = 1 To Len(strText)
         strChkSubject = strChkSubject & Chr(Asc(StrConv(Mid(strText, jj, 1), vbNarrow)))
      Next jj
      '2020/5/14 END
      
      bolChin = False
      intChinWord = 0 'Add By Sindy 2020/5/28 中文字數
      '檢查字串是否有中文或全形字
      For jj = 1 To Len(strChkSubject)
         'Modify By Sindy 2022/4/6 排除全型符號表
         'Modify By Sindy 2022/4/25 + 排除16進位符號表(一般文字:半形及全形字元)
         If InStr(WM_全型符號表, Mid(strChkSubject, jj, 1)) = 0 And _
            Not (Hex(AscW(Mid(strChkSubject, jj, 1))) >= "FE30" And Hex(AscW(Mid(strChkSubject, jj, 1))) <= "FFED") Then
         '2022/4/6 END
            If Asc(Mid(strChkSubject, jj, 1)) <= 0 Then
               'Add By Sindy 2020/5/28 中文字數超過2時,才算是有中文字的信件,因為有些特殊符號是雙位元
               'ex:FW: Transfer of Matters for Jelmar, LLC – Taiwan
               intChinWord = intChinWord + 1
               If intChinWord > 2 Then
               '2020/5/28 END
                  bolChin = True
                  Exit For
               End If
            End If
         End If
      Next jj
      '有中文字且無含unicode
      If bolChin = True And InStr(strChkSubject, "?") = 0 Then
         m_Sender = Pub_GetSpecMan("國內信件管理人員")
         PUB_IPDept_ToSortOut = "Z" '其他
      End If
   End If
   '2020/3/30 END
   
End If 'Add By Sindy 2019/8/16
   
'*************************
ChkEnd: '條件檢查完畢
'*************************
   'Call PUB_IPDept_ToSortOutSub_AC(strText, m_Sender, PUB_IPDept_ToSortOut) 'Add By Sindy 2016/4/15 AC/xx:寰華案件
   '戊、無以上條件者分至其他；
   If m_Sender = "" Then PUB_IPDept_ToSortOut = "" 'Add By Sindy 2023/7/21
   If PUB_IPDept_ToSortOut = "" Then PUB_IPDept_ToSortOut = "Z" '其他
   'If PUB_IPDept_ToSortOut = "Z" Then m_Sender = "" 'Modify By Sindy 2016/5/5 Mark
   'Add By Sindy 2017/7/28 4.專利處 6.新知 8.開拓 Z.其他 : 無歸卷宗區的狀況
   If PUB_IPDept_ToSortOut = "4" Or _
      PUB_IPDept_ToSortOut = "6" Or _
      PUB_IPDept_ToSortOut = "8" Or _
      PUB_IPDept_ToSortOut = "Z" Then
      strCP01 = "": strCP02 = "": strCP03 = "": strCP04 = ""
   End If
   '2017/7/28 END
   
   'Add by Sindy 2019/8/15
   'Modify By Sindy 2021/2/4 strII18 <> "" => PUB_IPDept_ToSortOut = ""
   If Trim(strII11) = "未傳遞的主旨" And _
      (PUB_IPDept_ToSortOut = "" Or PUB_IPDept_ToSortOut = "1") Then
      PUB_IPDept_ToSortOut = "Z" '其他
      m_Sender = ""
   End If
   
   'Modify By Sindy 2022/8/22 Mark
'   'Add By Sindy 2017/10/6 8.開拓
'   '目前系統有關APAA分信原則為：
'   '主旨中含有APAA　則轉寄給　國外部轉信開拓群組(即陳增廣主任/楊雯芳副理/何主秘/閻副所長); Amanda
'   '改成將 ==>
'   If PUB_IPDept_ToSortOut = "8" And InStr(UCase(strII18), UCase("apaa")) > 0 Then
'      '主旨中含有APAA及???　轉寄給　國外部轉信開拓群組(即陳增廣主任/楊雯芳副理/何主秘/閻副所長); William
'      'Modify By Sindy 2017/10/16 + 或 主旨中含有APAA且寄信者網域為.jp
'      m_Sender = ""
'      If InStr(strSubject, "???") > 0 Or _
'         Right(Replace(Trim(strII11), "]", ""), 3) = ".jp" Then
'         tmpArr = Split("國外部轉信開拓群組;88003", ";")
'      '主旨中含有APAA且無???　轉寄給　國外部轉信開拓群組(即陳增廣主任/楊雯芳副理/何主秘/閻副所長); Amanda
'      'Modify By Sindy 2017/10/16 + 或 主旨中含有APAA且寄信者網域為非.jp
'      Else
'         'Modify By Sindy 2020/4/10 Widen說要移除 88003
'         'tmpArr = Split("國外部轉信開拓群組;80030", ";")
'         tmpArr = Split("國外部轉信開拓群組", ";")
'      End If
'      For j = 0 To UBound(tmpArr)
'         strTemp = Pub_GetSpecMan(CStr(tmpArr(j)))
'         If strTemp <> "" Then
'            m_Sender = m_Sender & ";" & strTemp
'         Else
'            m_Sender = m_Sender & ";" & tmpArr(j)
'         End If
'      Next j
'   End If
'   '2017/10/6 END
   
   '過濾是否有收受者重覆的資料
   If Left(m_Sender, 1) = ";" Then m_Sender = Mid(m_Sender, 2)
   If m_Sender <> "" And InStr(m_Sender, ";") > 0 Then
      strText = m_Sender
      tmpArr = Split(strText, ";")
      m_Sender = ""
      For j = 0 To UBound(tmpArr)
         If tmpArr(j) <> "" Then
            If InStr(m_Sender, tmpArr(j)) = 0 Then
               m_Sender = m_Sender & IIf(m_Sender = "", "", ";") & tmpArr(j)
            End If
         End If
      Next j
   End If
   
   'Add By Sindy 2019/3/14
   '國外部商標二區案件，來自美洲、非洲、法國、瑞典代理人或客戶相關emails，原系統僅設定傳吳國安主任，
   '自即日起，請調整設定，增加傳給陳昇鴻專員。
   '收受者有A3022吳國安就加發A6005陳昇鴻
   If InStr(m_Sender, "A3022") > 0 And InStr(m_Sender, "A6005") = 0 Then
      m_Sender = m_Sender & ";A6005"
   End If
   '2019/3/14 END
   
   'Add By Sindy 2019/2/22 特殊狀況,外專承辦組人員要加發David.77015
   'A5023.洪培堯的信件要加發David
   'Modify By Sindy 2019/3/20 Mark.取消
   'If strSrvDate(1) >= 20190224 And InStr(m_Sender, "77015") = 0 And InStr(m_Sender, "A5023") > 0 Then
   'Modify By Sindy 2021/12/2 原 Ryan(A6010) 及 Franny(A8013) 之轉寄主管增設 David
   'Modify By Sindy 2021/1/3 Mark
'   If InStr(m_Sender, "77015") = 0 And _
'      (InStr(m_Sender, "A6010") > 0 Or InStr(m_Sender, "A8013") > 0) Then
'      m_Sender = m_Sender & ";77015"
'   End If
   '2019/2/22 END
   
'   'Add By Sindy 2017/4/6 外專:承辦組英文,日文組長的信不用加發David,目前(國外部轉信外專承辦英文組長)是David
'   If PUB_IPDept_ToSortOut = "3" Then '外專
'      If InStr(m_Sender, Pub_GetSpecMan("國外部轉信外專承辦英文組長")) = 0 And _
'         InStr(m_Sender, Pub_GetSpecMan("國外部轉信外專承辦日文組長")) = 0 Then '收受者沒有外專承辦英,日文組長
'         'Add By Sindy 2017/4/7 單純的外專程序組信件,不須再加David
'         bolAllF22 = False
'         tmpArr = Split(m_Sender, ";")
'         For j = 0 To UBound(tmpArr)
'            If tmpArr(j) <> "" Then
'               strExc(0) = "select st03" & _
'                           " from staff" & _
'                           " where st01='" & tmpArr(j) & "' and st04='1'"
'               intI = 1
'               Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'               If intI = 1 Then
'                  If rsTmp.Fields("st03") = "F22" Then
'                     bolAllF22 = True
'                  Else
'                     bolAllF22 = False
'                     Exit For
'                  End If
'               End If
'            End If
'         Next j
'         If bolAllF22 = False Then
'            m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外專承辦英文組長")
'         End If
'      End If
'   End If
'   '2017/4/6 END
   
   Set rsTmp = Nothing
End Function

'Add By Sindy 2024/4/29
Public Sub PUB_GetF21Manager(ByVal strST16 As String, ByVal StrST01 As String, _
   ByRef m_Sender As String)
   
   'Modify By Sindy 2022/5/19 加發工程師主管
   strExc(10) = ""
   If strST16 = "3" Then '日文
      strExc(10) = PUB_GetFCPEngSup(StrST01)
      If strExc(10) <> "" Then
         m_Sender = m_Sender & ";" & strExc(10)
      End If
   'Modify By Sindy 2024/4/29 專利國外部電子、化學及機械組
   Else
      '核稿主管與判發主管加入e-mail通知。
      strExc(10) = Left(PUB_GetFCPEngSup(StrST01, , , True), 5)
      If strExc(10) <> "" Then
         m_Sender = m_Sender & ";" & strExc(10)
      End If
      '當承辦工程師與最高主管組別不同時，不通知最高主管。
      strExc(10) = PUB_GetFCPEngSup(StrST01)
      If PUB_GetStaffST16(strExc(10)) = strST16 Then
         If InStr(m_Sender, strExc(10)) = 0 Then
            m_Sender = m_Sender & ";" & strExc(10)
         End If
      End If
   '2024/4/29 END
   End If
End Sub

'Add By Sindy 2020/3/6
'尋找ipdeptkeyword的信件關鍵字
Public Function PUB_FindIPdeptKeyWord(ByVal strConSql As String, ByVal strSubject As String, _
   ByVal strII11 As String, ByRef m_Sender As String, ByRef strII18 As String) As String
Dim rsTmp As New ADODB.Recordset
Dim strTemp As String
Dim bolChkOk As Boolean
Dim strWord As String
Dim tmpArr As Variant
Dim j As Integer

'      'Modify By Sindy 2018/1/10
'      'strSql = "select LK01,LK02,LK03,LK04 from ipdeptkeyword where LK12='F' and LK02 in('6','8') order by LK13 asc,LK01 asc"
'      strSql = "select lk01,lk02,lk03,lk04,lk13,lk14 from ipdeptkeyword where LK12='F' and LK02 in('6','8') and lk14 is null" & _
'               " union select ' '||rtrim(ltrim(lk01))||' ' lk01,lk02,lk03,lk04,LK13,lk14 from ipdeptkeyword where LK12='F' and LK02 in('6','8') and lk14='Y'" & _
'               " order by lk13 asc,lk01 asc"
'      '2018/1/10 END
   'Modify By Sindy 2018/5/17
   strSql = "select lk01,lk02,lk03,lk04,to_number(lk13) lk13,lk14,LK11,LK12 from ipdeptkeyword" & _
            " where lk14 is null and lk03='1'" & strConSql & _
            " and InStr('" & UCase(strSubject) & "',upper(rtrim(ltrim(lk01)))) > 0" & _
            " union select ' '||rtrim(ltrim(lk01))||' ' lk01,lk02,lk03,lk04,to_number(lk13) lk13,lk14,LK11,LK12 from ipdeptkeyword" & _
            " where lk14='Y' and lk03='1'" & strConSql & _
            " and InStr('" & UCase(strSubject) & "',upper(rtrim(ltrim(lk01)))) > 0" & _
            " union select lk01,lk02,lk03,lk04,to_number(lk13) lk13,lk14,LK11,LK12 from ipdeptkeyword" & _
            " where lk14 is null and lk03='2'" & strConSql & _
            " and InStr('" & UCase(ChgSQL(strII11)) & "',upper(rtrim(ltrim(lk01)))) > 0" & _
            " union select ' '||rtrim(ltrim(lk01))||' ' lk01,lk02,lk03,lk04,to_number(lk13) lk13,lk14,LK11,LK12 from ipdeptkeyword" & _
            " where lk14='Y' and lk03='2'" & strConSql & _
            " and InStr('" & UCase(ChgSQL(strII11)) & "',upper(rtrim(ltrim(lk01)))) > 0" & _
            " order by lk13 asc,lk01 asc"
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      With rsTmp
         rsTmp.MoveFirst
         Do While rsTmp.EOF = False
'               If rsTmp.Fields("LK03") = "1" Then '主旨
'                  If InStr(UCase(strSubject), UCase(rsTmp.Fields("LK01"))) > 0 Then
'                     strII18 = rsTmp.Fields("LK01") 'Add By Sindy 2017/8/28
'                     PUB_FindIPdeptKeyWord = rsTmp.Fields("LK02")
'                     tmpArr = Split("" & rsTmp.Fields("LK04"), ";")
'                     For j = 0 To UBound(tmpArr)
'                        strTemp = Pub_GetSpecMan(CStr(tmpArr(j)))
'                        If strTemp <> "" Then
'                           m_Sender = m_Sender & ";" & strTemp
'                        Else
'                           m_Sender = m_Sender & ";" & tmpArr(j)
'                        End If
'                     Next j
'                  'Add By Sindy 2018/1/10 單字索引
'                  Else
               If rsTmp.Fields("LK03") = "1" Then '主旨
                  strTemp = Trim(UCase(strSubject))
               Else
                  strTemp = Trim(UCase(strII11))
               End If
               If "" & rsTmp.Fields("lk14") = "Y" Then
                  bolChkOk = False
                  '索引最前面
                  strWord = Trim(UCase(rsTmp.Fields("LK01"))) & " "
                  If Left(strTemp, Len(strWord)) = strWord Then
                     bolChkOk = True
                  End If
                  If bolChkOk = False Then
                     '索引最後面
                     strWord = " " & Trim(UCase(rsTmp.Fields("LK01")))
                     If Right(strTemp, Len(strWord)) = strWord Then
                        bolChkOk = True
                     End If
                  End If
                  'Add By Sindy 2018/6/27
                  If bolChkOk = False Then
                     '索引中間
                     strWord = " " & Trim(UCase(rsTmp.Fields("LK01"))) & " "
                     If InStr(strTemp, strWord) > 0 Then
                        bolChkOk = True
                     End If
                  End If
                  'Add By Sindy 2020/12/3 Calendar
                  If bolChkOk = False Then
                     'From: donjoweber@aol.com [mailto:donjoweber@aol.com]
                     'Sent: Thursday, December 03, 2020 5:02 AM
                     'To: ipdept <ipdept@taie.com.tw>
                     'Subject:Calendar
                     '就單字
                     strWord = Trim(UCase(rsTmp.Fields("LK01")))
                     If strTemp = strWord Then
                        bolChkOk = True
                     End If
                  End If
'                     '有索引到
'                     If bolChkOk = True Then
'                        strII18 = rsTmp.Fields("LK01") 'Add By Sindy 2017/8/28
'                        PUB_FindIPdeptKeyWord = rsTmp.Fields("LK02")
'                        tmpArr = Split("" & rsTmp.Fields("LK04"), ";")
'                        For j = 0 To UBound(tmpArr)
'                           strTemp = Pub_GetSpecMan(CStr(tmpArr(j)))
'                           If strTemp <> "" Then
'                              m_Sender = m_Sender & ";" & strTemp
'                           Else
'                              m_Sender = m_Sender & ";" & tmpArr(j)
'                           End If
'                        Next j
'                        Exit Do
'                     End If
               'Add By Sindy 2020/3/6 優先搜尋到的關鍵字也有可能不是單字索引
               Else
                  bolChkOk = True
               End If
               '有索引到
               If bolChkOk = True Then
                  strII18 = strII18 & ";" & rsTmp.Fields("LK01") 'Add By Sindy 2017/8/28
                  'Add By Sindy 2020/6/25
                  If "" & rsTmp.Fields("lk12") = "T" Then
                     If "" & rsTmp.Fields("lk11") <> "" Then
                        strII18 = strII18 & "(" & rsTmp.Fields("lk11") & ")"
                     End If
                  End If
                  '2020/6/25 END
                  PUB_FindIPdeptKeyWord = rsTmp.Fields("LK02")
                  tmpArr = Split("" & rsTmp.Fields("LK04"), ";")
                  For j = 0 To UBound(tmpArr)
                     strTemp = Pub_GetSpecMan(CStr(tmpArr(j)))
                     If strTemp <> "" Then
                        m_Sender = m_Sender & ";" & strTemp
                     Else
                        m_Sender = m_Sender & ";" & tmpArr(j)
                     End If
                  Next j
                  'Add By Sindy 2024/5/17 記錄使用次數
                  cnnConnection.Execute "update ipdeptkeyword set LK16=LK16+1" & _
                                        " where LK01='" & ChgSQL(rsTmp.Fields("LK01")) & "' and LK12='" & rsTmp.Fields("LK12") & "'" _
                                        , intI
                  '2024/5/17 END
                  Exit Do
               End If
               '2020/3/6 END
               '2018/1/10 END
'               Else '寄件者或網域
'                  If InStr(UCase(strII11), UCase(rsTmp.Fields("LK01"))) > 0 Then
'                     strII18 = rsTmp.Fields("LK01") 'Add By Sindy 2017/8/28
'                     PUB_FindIPdeptKeyWord = rsTmp.Fields("LK02")
'                     tmpArr = Split("" & rsTmp.Fields("LK04"), ";")
'                     For j = 0 To UBound(tmpArr)
'                        strTemp = Pub_GetSpecMan(CStr(tmpArr(j)))
'                        If strTemp <> "" Then
'                           m_Sender = m_Sender & ";" & strTemp
'                        Else
'                           m_Sender = m_Sender & ";" & tmpArr(j)
'                        End If
'                     Next j
'                  End If
'               End If
            If PUB_FindIPdeptKeyWord <> "" Then Exit Do
            rsTmp.MoveNext
         Loop
         'Add By Sindy 2018/5/17
         If PUB_FindIPdeptKeyWord = "" Then
            rsTmp.MoveFirst
            Do While rsTmp.EOF = False
               If "" & rsTmp.Fields("lk14") = "" Then
                  strII18 = strII18 & ";" & rsTmp.Fields("LK01")
                  'Add By Sindy 2020/6/25
                  If "" & rsTmp.Fields("lk12") = "T" Then
                     If "" & rsTmp.Fields("lk11") <> "" Then
                        strII18 = strII18 & "(" & rsTmp.Fields("lk11") & ")"
                     End If
                  End If
                  '2020/6/25 END
                  PUB_FindIPdeptKeyWord = rsTmp.Fields("LK02")
                  m_Sender = ""
                  tmpArr = Split("" & rsTmp.Fields("LK04"), ";")
                  For j = 0 To UBound(tmpArr)
                     strTemp = Pub_GetSpecMan(CStr(tmpArr(j)))
                     If strTemp <> "" Then
                        m_Sender = m_Sender & ";" & strTemp
                     Else
                        m_Sender = m_Sender & ";" & tmpArr(j)
                     End If
                  Next j
                  'Add By Sindy 2024/5/17 記錄使用次數
                  cnnConnection.Execute "update ipdeptkeyword set LK16=LK16+1" & _
                                        " where LK01='" & rsTmp.Fields("LK01") & "' and LK12='" & rsTmp.Fields("LK12") & "'" _
                                        , intI
                  '2024/5/17 END
                  Exit Do
               End If
               rsTmp.MoveNext
            Loop
         End If
         '2018/5/17 END
      End With
   End If
   
   Set rsTmp = Nothing
End Function

'Add By Sindy 2019/9/3 依外法英文縮寫分信
'Modify By Sindy 2021/1/19 + ,Optional ByRef strII18 As String
'Modify By Sindy 2022/7/8 + , ByVal strST04 As String: 1.在職 2.離職
Private Sub PUB_IPDept_ToSortOutSub_FLaw(ByVal strText As String, ByVal strII11 As String, ByRef m_Sender As String, _
                                 ByRef m_ToSortOut As String, ByVal strST04 As String, Optional ByRef strII18 As String)
Dim rsTmp As New ADODB.Recordset
Dim strCon As String 'Add By Sindy 2022/7/8
   
   'Modify By Sindy 2022/7/8
   If strST04 = "2" Then
      'Modify By Sindy 2020/11/18 抓離職2個月內的人員
      strCon = " and st04='2'" & _
               " and st51 is not null and st51>=" & CompDate(1, -2, strSrvDate(1))
   Else
      strCon = " and st04='1'"
   End If
   '2022/7/8 END
   
'***法務:  主管+/+個人
   'ST01     ST02         ST0 ST17        ST52   ST1
   '-------- ------------ --- ----------- ------ ---
   '98003    林美宏       F31 HG/nl       73029  1
   '98020    江郁仁       F31 HG/YJ       73029  2
   '選取了 2 筆資料列.
   'Modify By Sindy 2020/3/31 ipdept分信,F31處再加LXX部門
   'and st03 in('F31','L01','L02') => and (st03='F31' or substr(st03,1,1)='L')
   strExc(0) = "select st01,st02,st03,st17,st52,st16" & _
               " from staff" & _
               " where (st03='F31' or substr(st03,1,1)='L') and substr(st01,1,1)<>'F'" & _
               " and st17 is not null and instr(st17,'/')>0" & strCon
   'Modify By Sindy 2022/7/8
   If strST04 = "2" Then
      strExc(0) = strExc(0) & " order by st51 desc"
   End If
   '2022/7/8 END
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         'Modify By Sindy 2022/7/8 英文縮寫加空白格後, 再搜尋主旨; 亦也會有 WC/as/wfc 這樣的縮寫, 故也增加/判斷
         '1. CH/YCP+空白格 2. WC/as/wfc
         'Modify By Sindy 2022/7/12 英文縮寫加) ex: 3. (EY/ey)
         'Modify By Sindy 2024/12/31 取消:InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & "/") > 0 Or
         If InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & " ") > 0 Or _
            (InStr(UCase(strText), "/" & UCase(rsTmp.Fields("st17"))) > 0 And InStr(UCase(rsTmp.Fields("st17")), "/") > 0) Or _
            InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & ")") > 0 Then
            
            strII18 = strII18 & ";" & UCase(rsTmp.Fields("st17")) 'Add By Sindy 2021/1/19
            'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "5" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
            m_ToSortOut = "5" '外法
            If "" & rsTmp.Fields("st16") = "2" Then
               m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外法日文組群組")
            Else
               m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外法英文組群組")
            End If
            Exit Do
         End If
         rsTmp.MoveNext
      Loop
   End If
   'rsTmp.Close
'   'Add By Sindy 2017/10/24 再檢查離職人員
'   'Modify By Sindy 2020/11/18 抓離職2個月內的人員
'   If m_ToSortOut = "" Then
'      strExc(0) = "select st01,st02,st03,st17,st52,st16" & _
'                  " from staff" & _
'                  " where st04='2' and st03 in('F31','L01','L02') and substr(st01,1,1)<>'F'" & _
'                  " and st17 is not null and instr(st17,'/')>0 and st51 is not null" & _
'                  " and st51>=" & CompDate(1, -2, strSrvDate(1)) & _
'                  " order by st51 desc"
'      intI = 1
'      Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         rsTmp.MoveFirst
'         Do While Not rsTmp.EOF
'            If InStr(UCase(strText), UCase(rsTmp.Fields("st17"))) > 0 Then
'               strII18 = strII18 & ";" & UCase(rsTmp.Fields("st17")) 'Add By Sindy 2021/1/19
'               'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "5" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
'               m_ToSortOut = "5" '外法
'               If "" & rsTmp.Fields("st16") = "2" Then
'                  m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外法日文組群組")
'               Else
'                  m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外法英文組群組")
'               End If
'               Exit Do
'            End If
'            rsTmp.MoveNext
'         Loop
'      End If
'   End If

   Set rsTmp = Nothing
End Sub

'Add By Sindy 2019/9/3 依外商英文縮寫分信
'Modify By Sindy 2021/1/19 + ,Optional ByRef strII18 As String
'Modify By Sindy 2022/7/8 + , ByVal strST04 As String: 1.在職 2.離職
Private Sub PUB_IPDept_ToSortOutSub_F1x(ByVal strText As String, ByVal strII11 As String, ByRef m_Sender As String, _
                                 ByRef m_ToSortOut As String, ByVal strST04 As String, _
                                 Optional ByRef m_Emp As String = "", Optional ByRef strDirector As String = "", _
                                 Optional ByRef strII18 As String)
Dim rsTmp As New ADODB.Recordset
'Dim bolChkFindCFT As Boolean 'Add By Sindy 2018/3/9
Dim strCon As String 'Add By Sindy 2022/7/8
   
'***外商:主管洪琬姿請假時,英文組承辦人代號AH/會改為FC/
   
   'Modify By Sindy 2022/7/8
   If strST04 = "2" Then
      'Modify By Sindy 2020/11/18 抓離職2個月內的人員
      strCon = " and s1.st04='2'" & _
               " and s1.st51 is not null and s1.st51>=" & CompDate(1, -2, strSrvDate(1))
   Else
      strCon = " and s1.st04='1'"
   End If
   '2022/7/8 END
   
'   bolChkFindCFT = False 'Add By Sindy 2018/3/9
'2017/12/27 END
'      strExc(0) = "select st01,st02,st03,st17,st52" & _
'                  " from staff" & _
'                  " where st04='1' and st03 in('F10','F11') and substr(st01,1,1)<>'F'" & _
'                  " and instr(st17,'/')>0" & _
'                  " union " & _
'                  "select st01,st02,st03,replace(st17,'AH/','FC/') st17,st52" & _
'                  " from staff" & _
'                  " where st04='1' and st03='F11' and substr(st01,1,1)<>'F'" & _
'                  " and instr(st17,'/')>0 and substr(st17,1,3)='AH/'"
'      'Modify By Sindy 2017/6/30 葉易雲:請幫忙我的改為如下二個：MY/my 及 AH/ah
'      strExc(0) = strExc(0) & " union " & _
'                  "select st01,st02,st03,upper(substr(st17,instr(st17,'/')+1))||substr(st17,instr(st17,'/')) st17,st52" & _
'                  " from staff" & _
'                  " where st04='1' and st03 in('F10','F11') and substr(st01,1,1)<>'F'" & _
'                  " and instr(st17,'/')>0 and st05 in('26','28')"
   'Modify By Sindy 2018/3/8 外商改設個人英文縮寫(ll),沒有設一整組(FC/ll)
   strExc(0) = "select s1.st01 st01,s1.st02,s1.st03,decode(instr(s1.st17,'/'),0,upper(substr(s2.st17,instr(s2.st17,'/')+1))||'/'||s1.st17,s1.st17) st17,s1.st52,s1.st16 st16,s1.st51 st51" & _
               " from staff s1,staff s2" & _
               " where s1.st03 in('F10','F11') and substr(s1.st01,1,1)<>'F'" & _
               " and s1.st17 is not null and s1.st52 is not null" & _
               " and s1.st52=s2.st01(+) and s2.st17 is not null" & strCon
   strExc(0) = strExc(0) & " union " & _
               "select s1.st01 st01,s1.st02,s1.st03,decode(instr(s1.st17,'/'),0,upper(substr(s2.st17,instr(s2.st17,'/')+1))||'/'||s1.st17,s1.st17) st17,s1.st52,s1.st16 st16,s1.st51 st51" & _
               " from staff s1,staff s2" & _
               " where s1.st03 in('F10','F11') and substr(s1.st01,1,1)<>'F'" & _
               " and s1.st17 is not null and s1.st53 is not null" & _
               " and s1.st53=s2.st01(+) and s2.st17 is not null" & strCon
   strExc(0) = strExc(0) & " union " & _
               "select s1.st01 st01,s1.st02,s1.st03,decode(instr(s1.st17,'/'),0,upper(substr(s2.st17,instr(s2.st17,'/')+1))||'/'||s1.st17,s1.st17) st17,s1.st52,s1.st16 st16,s1.st51 st51" & _
               " from staff s1,staff s2" & _
               " where s1.st03 in('F10','F11') and substr(s1.st01,1,1)<>'F'" & _
               " and s1.st17 is not null and s1.st54 is not null" & _
               " and s1.st54=s2.st01(+) and s2.st17 is not null" & strCon
   strExc(0) = strExc(0) & " union " & _
               "select s1.st01 st01,s1.st02,s1.st03,decode(instr(s1.st17,'/'),0,upper(substr(s2.st17,instr(s2.st17,'/')+1))||'/'||s1.st17,s1.st17) st17,s1.st52,s1.st16 st16,s1.st51 st51" & _
               " from staff s1,staff s2" & _
               " where s1.st03 in('F10','F11') and substr(s1.st01,1,1)<>'F'" & _
               " and s1.st17 is not null and s1.st55 is not null" & _
               " and s1.st55=s2.st01(+) and s2.st17 is not null" & strCon
   'Modify By Sindy 2017/6/30 葉易雲:請幫忙我的改為如下二個：MY/my 及 AH/ah
   strExc(0) = strExc(0) & " union " & _
               "select st01,st02,st03,decode(instr(st17,'/'),0,upper(substr(st17,instr(st17,'/')+1))||'/'||st17,st17) st17,st52,st16,st51" & _
               " from staff" & _
               " where st03 in('F10','F11') and substr(st01,1,1)<>'F'" & _
               " and st17 is not null and st05 in('26','28')"
            If strST04 = "2" Then
               strExc(0) = strExc(0) & " and st04='2'" & _
                                       " and st51 is not null and st51>=" & CompDate(1, -2, strSrvDate(1))
            Else
               strExc(0) = strExc(0) & " and st04='1'"
            End If
   'Add By Sindy 2020/9/10 + 信函Initial資料檔
   'Modify By Sindy 2021/11/3 Mark: and (st16='2' or st16 is null)
   strExc(0) = strExc(0) & " union " & _
               "select st01,st02,st03,ID03 st17,ID02 st52,st16,st51" & _
               " from InitialData,staff" & _
               " where ID01=st01 and st03 in('F10','F11') and substr(st01,1,1)<>'F'"
            If strST04 = "2" Then
               strExc(0) = strExc(0) & " and st04='2'" & _
                                       " and st51 is not null and st51>=" & CompDate(1, -2, strSrvDate(1))
            Else
               strExc(0) = strExc(0) & " and st04='1'"
            End If
   'Modify By Sindy 2022/7/8
   If strST04 = "2" Then
      strExc(0) = strExc(0) & " order by st51 desc"
   End If
   '2022/7/8 END
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         'Modify By Sindy 2022/7/8 英文縮寫加空白格後, 再搜尋主旨; 亦也會有 WC/as/wfc 這樣的縮寫, 故也增加/判斷
         '1. CH/YCP+空白格 2. WC/as/wfc
         'Modify By Sindy 2022/7/12 英文縮寫加) ex: 3. (EY/ey)
         If InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & " ") > 0 Or _
            (InStr(UCase(strText), "/" & UCase(rsTmp.Fields("st17"))) > 0 And InStr(UCase(rsTmp.Fields("st17")), "/") > 0) Or _
            InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & ")") > 0 Then
            
            'Modify By Sindy 2017/12/27
            'Modify By Sindy 2018/3/12
'               If strSrvDate(1) >= 20180102 And (strCP01 = "CFT" Or strCP01 = "CFC" Or strCP01 = "S") Then
'                  m_Sender = Pub_GetSpecMan("V") & ";" & rsTmp.Fields("st52") & ";" & rsTmp.Fields("st01") '陳經理+ST52
'               Else
'               '2017/12/27 END
'                  m_Sender = Pub_GetSpecMan("國外部轉信外商群組")
'               End If
            If "" & rsTmp.Fields("st16") = "4" Then '4.日文組
               m_Sender = Trim(Pub_GetSpecMan("外商日文組通知主管")) '"78011" '改直接寫固定寄給May
               'Modify By Sindy 2024/8/27 檢查是否已離職
               If ChkStaffST04(m_Sender, False) = True Or m_Sender = "" Then
                  m_Sender = ""
                  GoTo ReadNext
               End If
               '2024/8/27 END
               
'               'Modify By Sindy 2018/3/26
'               If Pub_GetSpecMan("V") = rsTmp.Fields("st52") Then
'                  m_Sender = Pub_GetSpecMan("V") & ";" & rsTmp.Fields("st01") '陳經理+ST01
'               Else
'               '2018/3/26 END
'                  'Modify By Sindy 2021/5/25 日文組改直接寫固定寄給May+陳經理
'                  'm_Sender = Pub_GetSpecMan("V") & ";" & rsTmp.Fields("st52") '陳經理+ST52(May)
'                  m_Sender = "78011;" & Pub_GetSpecMan("V") '改直接寫固定寄給May+陳經理
'               End If
            Else
               'Modify By Sindy 2024/8/27 排除離職人員，直接分給該縮寫組合之主管
               If strST04 = "1" Then m_Sender = rsTmp.Fields("st01") '個人
               
               'Modify By Sindy 2021/6/17 + 加發 2,3,4級主管
               strExc(0) = "select st01,st52,st53,st54 from staff where st01='" & rsTmp.Fields("st01") & "'"
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  If "" & RsTemp.Fields("st52") <> "" Then
                     'Modify By Sindy 2024/8/27 檢查是否已離職
                     If ChkStaffST04(RsTemp.Fields("st52"), False) = False Then
                        m_Sender = m_Sender & ";" & RsTemp.Fields("st52")
                     End If
                  End If
                  If "" & RsTemp.Fields("st53") <> "" Then
                     'Modify By Sindy 2024/8/27 檢查是否已離職
                     If ChkStaffST04(RsTemp.Fields("st53"), False) = False Then
                        m_Sender = m_Sender & ";" & RsTemp.Fields("st53")
                     End If
                  End If
                  If "" & RsTemp.Fields("st54") <> "" Then
                     'Modify By Sindy 2024/8/27 檢查是否已離職
                     If ChkStaffST04(RsTemp.Fields("st54"), False) = False Then
                        m_Sender = m_Sender & ";" & RsTemp.Fields("st54")
                     End If
                  End If
               End If
               '2021/6/17 END
               'm_Sender = 外商分信經理 & ";" & rsTmp.Fields("st52") & ";" & rsTmp.Fields("st01") '經理+ST52+個人
            End If
            '2018/3/12 END
'               m_Sender = m_Sender & ";" & rsTmp.Fields("st01")
'               If "" & rsTmp.Fields("st52") <> "" And "" & rsTmp.Fields("st52") <> rsTmp.Fields("st01") Then
'                  m_Sender = m_Sender & ";" & rsTmp.Fields("st52")
'               End If
            'Modify By Sindy 2024/8/27 皆已離職，需人工分信
            If m_Sender = "" Then
               GoTo ReadNext
            End If
            '2024/8/27 END
            
            strII18 = strII18 & ";" & UCase(rsTmp.Fields("st17")) & _
                           IIf(strST04 = "2", "(" & rsTmp.Fields("st01") & "已離職)", "") 'Add By Sindy 2021/1/19
            m_Emp = rsTmp.Fields("st01") 'Add By Sindy 2019/9/3
            strDirector = "" & rsTmp.Fields("ST52") 'Add By Sindy 2019/9/3
            
            'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "2" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
'            bolChkFindCFT = True 'Add By Sindy 2018/3/9
            
            m_ToSortOut = "2" '外商
            
            Exit Do
         End If
ReadNext:
         rsTmp.MoveNext
      Loop
   End If
   'rsTmp.Close
'   'Add By Sindy 2017/10/24 再檢查離職人員
'   'Modify By Sindy 2017/12/27
'   'If PUB_IPDept_ToSortOut = "" Then
'   'Modify By Sindy 2018/3/9
'   'If PUB_IPDept_ToSortOut = "" Or (strCP01 = "CFT" Or strCP01 = "CFC" Or strCP01 = "S") Then
''      If PUB_IPDept_ToSortOut = "" Or _
''         ((strCP01 = "CFT" Or strCP01 = "CFC" Or strCP01 = "S") And bolChkFindCFT = False) Then
'   If m_ToSortOut = "" Then
'   '2018/3/9 END
'   '2017/12/27 END
''         strExc(0) = "select st01,st02,st03,st17,st52,st51" & _
''                     " from staff" & _
''                     " where st04='2' and st03 in('F10','F11') and substr(st01,1,1)<>'F'" & _
''                     " and instr(st17,'/')>0 and st51 is not null" & _
''                     " union " & _
''                     "select st01,st02,st03,replace(st17,'AH/','FC/') st17,st52,st51" & _
''                     " from staff" & _
''                     " where st04='2' and st03='F11' and substr(st01,1,1)<>'F'" & _
''                     " and instr(st17,'/')>0 and substr(st17,1,3)='AH/' and st51 is not null"
''         '葉易雲:請幫忙我的改為如下二個：MY/my 及 AH/ah
''         strExc(0) = strExc(0) & " union " & _
''                     "select st01,st02,st03,upper(substr(st17,instr(st17,'/')+1))||substr(st17,instr(st17,'/')) st17,st52,st51" & _
''                     " from staff" & _
''                     " where st04='2' and st03 in('F10','F11') and substr(st01,1,1)<>'F'" & _
''                     " and instr(st17,'/')>0 and st05 in('26','28') and st51 is not null" & _
''                     " order by st51 desc"
'      'Modify By Sindy 2018/3/8 外商改設個人英文縮寫(ll),沒有設一整組(FC/ll)
'      'Modify By Sindy 2020/11/18 抓離職2個月內的人員
'      strExc(0) = "select s1.st01 st01,s1.st02,s1.st03,decode(instr(s1.st17,'/'),0,upper(substr(s2.st17,instr(s2.st17,'/')+1))||'/'||s1.st17,s1.st17) st17,s1.st52,s1.st51,s1.st16 st16" & _
'                  " from staff s1,staff s2" & _
'                  " where s1.st04='2' and s1.st03 in('F10','F11') and substr(s1.st01,1,1)<>'F'" & _
'                  " and s1.st17 is not null and s1.st52 is not null" & _
'                  " and s1.st52=s2.st01(+) and s2.st17 is not null and s1.st51 is not null" & _
'                  " and s1.st51>=" & CompDate(1, -2, strSrvDate(1))
'      strExc(0) = strExc(0) & " union " & _
'                  "select s1.st01 st01,s1.st02,s1.st03,decode(instr(s1.st17,'/'),0,upper(substr(s2.st17,instr(s2.st17,'/')+1))||'/'||s1.st17,s1.st17) st17,s1.st52,s1.st51,s1.st16 st16" & _
'                  " from staff s1,staff s2" & _
'                  " where s1.st04='2' and s1.st03 in('F10','F11') and substr(s1.st01,1,1)<>'F'" & _
'                  " and s1.st17 is not null and s1.st53 is not null" & _
'                  " and s1.st53=s2.st01(+) and s2.st17 is not null and s1.st51 is not null" & _
'                  " and s1.st51>=" & CompDate(1, -2, strSrvDate(1))
'      strExc(0) = strExc(0) & " union " & _
'                  "select s1.st01 st01,s1.st02,s1.st03,decode(instr(s1.st17,'/'),0,upper(substr(s2.st17,instr(s2.st17,'/')+1))||'/'||s1.st17,s1.st17) st17,s1.st52,s1.st51,s1.st16 st16" & _
'                  " from staff s1,staff s2" & _
'                  " where s1.st04='2' and s1.st03 in('F10','F11') and substr(s1.st01,1,1)<>'F'" & _
'                  " and s1.st17 is not null and s1.st54 is not null" & _
'                  " and s1.st54=s2.st01(+) and s2.st17 is not null and s1.st51 is not null" & _
'                  " and s1.st51>=" & CompDate(1, -2, strSrvDate(1))
'      strExc(0) = strExc(0) & " union " & _
'                  "select s1.st01 st01,s1.st02,s1.st03,decode(instr(s1.st17,'/'),0,upper(substr(s2.st17,instr(s2.st17,'/')+1))||'/'||s1.st17,s1.st17) st17,s1.st52,s1.st51,s1.st16 st16" & _
'                  " from staff s1,staff s2" & _
'                  " where s1.st04='2' and s1.st03 in('F10','F11') and substr(s1.st01,1,1)<>'F'" & _
'                  " and s1.st17 is not null and s1.st55 is not null" & _
'                  " and s1.st55=s2.st01(+) and s2.st17 is not null and s1.st51 is not null" & _
'                  " and s1.st51>=" & CompDate(1, -2, strSrvDate(1))
'      '葉易雲:幫忙改為如下二個：MY/my 及 AH/ah
'      strExc(0) = strExc(0) & " union " & _
'                  "select st01,st02,st03,decode(instr(st17,'/'),0,upper(substr(st17,instr(st17,'/')+1))||'/'||st17,st17) st17,st52,st51,st16" & _
'                  " from staff" & _
'                  " where st04='2' and st03 in('F10','F11') and substr(st01,1,1)<>'F'" & _
'                  " and st17 is not null and st05 in('26','28') and st51 is not null" & _
'                  " and st51>=" & CompDate(1, -2, strSrvDate(1))
'      'Add By Sindy 2020/9/10 + 信函Initial資料檔
'      'Modify By Sindy 2021/11/3 Mark: and (st16='2' or st16 is null)
'      strExc(0) = strExc(0) & " union " & _
'                  "select st01,st02,st03,ID03 st17,ID02 st52,st51,st16" & _
'                  " from InitialData,staff" & _
'                  " where ID01=st01 and st04='2' and st03 in('F10','F11') and substr(st01,1,1)<>'F'" & _
'                  " and st51 is not null" & _
'                  " and st51>=" & CompDate(1, -2, strSrvDate(1))
'      strExc(0) = strExc(0) & " order by st51 desc"
'      intI = 1
'      Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         rsTmp.MoveFirst
'         Do While Not rsTmp.EOF
'            If InStr(UCase(strText), UCase(rsTmp.Fields("st17"))) > 0 Then
'               strII18 = strII18 & ";" & UCase(rsTmp.Fields("st17")) 'Add By Sindy 2021/1/19
'               bolChkFindCFT = True 'Add By Sindy 2018/3/9
'
'               m_Emp = rsTmp.Fields("st01") 'Add By Sindy 2019/9/3
'               strDirector = "" & rsTmp.Fields("ST52") 'Add By Sindy 2019/9/3
'
'               m_ToSortOut = "2" '外商
'               'Modify By Sindy 2017/12/27
'               'Modify By Sindy 2018/3/12
''                  If strSrvDate(1) >= 20180102 And (strCP01 = "CFT" Or strCP01 = "CFC" Or strCP01 = "S") Then
'               If "" & rsTmp.Fields("st16") = "4" Then '4.日文組
'                  m_Sender = "78011" '改直接寫固定寄給May
''                  'Modify By Sindy 2018/3/26
''                  If Pub_GetSpecMan("V") = rsTmp.Fields("st52") Then
''                     m_Sender = Pub_GetSpecMan("V") & ";" & rsTmp.Fields("st01") '經理+ST01
''                  Else
''                  '2018/3/26 END
''                     'Modify By Sindy 2021/5/25 日文組改直接寫固定寄給May+陳經理
''                     'm_Sender = Pub_GetSpecMan("V") & ";" & rsTmp.Fields("st52") '陳經理+ST52(May)
''                     m_Sender = "78011;" & Pub_GetSpecMan("V") '改直接寫固定寄給May+陳經理
''                  End If
'               Else
'                  m_Sender = rsTmp.Fields("st01") '個人
'                  'Modify By Sindy 2021/6/17 + 加發 2,3,4級主管
'                  strExc(0) = "select st01,st52,st53,st54 from staff where st01='" & rsTmp.Fields("st01") & "'"
'                  intI = 1
'                  Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'                  If intI = 1 Then
'                     If "" & RsTemp.Fields("st52") <> "" Then
'                        m_Sender = m_Sender & ";" & RsTemp.Fields("st52")
'                     End If
'                     If "" & RsTemp.Fields("st53") <> "" Then
'                        m_Sender = m_Sender & ";" & RsTemp.Fields("st53")
'                     End If
'                     If "" & RsTemp.Fields("st54") <> "" Then
'                        m_Sender = m_Sender & ";" & RsTemp.Fields("st54")
'                     End If
'                  End If
'                  '2021/6/17 END
'                  'm_Sender = 外商分信經理 & ";" & rsTmp.Fields("st52") & ";" & rsTmp.Fields("st01") '經理+ST52+個人
'               End If
'               '2018/3/12 END
''                  Else
''                  '2017/12/27 END
''                     m_Sender = Pub_GetSpecMan("國外部轉信外商群組")
''                  End If
'               Exit Do
'            End If
'            rsTmp.MoveNext
'         Loop
'      End If
'   End If
   
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2019/9/3 依外專工程師英文縮寫分信
'Modify By Sindy 2021/1/19 + ,Optional ByRef strII18 As String
'Modify By Sindy 2022/7/8 + , ByVal strST04 As String: 1.在職 2.離職
Private Sub PUB_IPDept_ToSortOutSub_F21(ByVal strText As String, ByVal strII11 As String, ByRef m_Sender As String, _
                                 ByRef m_ToSortOut As String, ByVal strST04 As String, _
                                 Optional ByRef m_Emp As String = "", Optional ByRef strDirector As String = "", _
                                 Optional ByRef strII18 As String)
Dim rsTmp As New ADODB.Recordset
'Add By Sindy 2022/4/11
Dim intRunTimes As Integer '檢查次數:1.在職 2.離職
Dim strCon As String
'2022/4/11 END
   
   'Modify By Sindy 2022/7/8
   If strST04 = "2" Then
      'Modify By Sindy 2020/11/18 抓離職2個月內的人員
      strCon = " and s1.st04='2'" & _
               " and s1.st51 is not null and s1.st51>=" & CompDate(1, -2, strSrvDate(1))
   Else
      strCon = " and s1.st04='1'"
   End If
   '2022/7/8 END
   
'   intRunTimes = 1 '1.在職 Add By Sindy 2022/4/11
'   strCon = "s1.st04='1'"
'
'ReStarChk:
'***外專工程師:ex.sh  主管+/+個人
   '   strExc(0) = "select s1.st01 st01,s1.st02,s1.st03,s2.st17||'/'||s1.st17 st17,oman st52,s1.st16" & _
   '               " from staff s1,setspecman,staff s2" & _
   '               " where s1.st04='1' and s1.st03='F21' and substr(s1.st01,1,1)<>'F'" & _
   '               " and s1.st17 is not null" & _
   '               " and decode(s1.st16,'1','T','2','R','3','S','4','T1',s1.st16)=OCODE(+)" & _
   '               " and oman=s2.st01(+)"
   'Modify By Sindy 2016/12/30 + AL/dl: Reminder: RE: Request for cost estimate for a design patent search (O/R: M6X2016634_TW)
   'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
   'Modified by Sindy 2020/10/14 因日文組所要加抓第三,四級主管
   'modify by sonia 2021/1/27 再排除F4104及F4105
   'Modify By Sindy 2025/6/26 增加檢查層級主管有值才抓 ex: and s1.st52 is not null
   strExc(0) = "select s1.st01 st01,s1.st02,s1.st03,upper(s2.st17)||'/'||s1.st17 st17,oman st52,s1.st16 st16,s1.st17 st01_17,s1.st51 st51" & _
               " from staff s1,setspecman,staff s2" & _
               " where s1.st03='F21' and substr(s1.st01,1,1)<>'F' and s1.st01<>'F4102' and s1.st01<>'F4104' and s1.st01<>'F4105'" & _
               " and s1.st17 is not null" & _
               " and decode(s1.st16,'1','T','2','R','3','S','4','T1',s1.st16)=OCODE(+)" & _
               " and oman=s2.st01(+)" & strCon & _
               " Union " & _
               "select s1.st01 st01,s1.st02,s1.st03,upper(nvl(s2.st17,s1.st17))||'/'||s1.st17 st17,s1.st52,s1.st16 st16,s1.st17 st01_17,s1.st51 st51" & _
               " from staff s1,staff s2" & _
               " where s1.st03='F21' and substr(s1.st01,1,1)<>'F' and s1.st01<>'F4102' and s1.st01<>'F4104' and s1.st01<>'F4105'" & _
               " and s1.st17 is not null" & _
               " and s1.st52=s2.st01(+) and s1.st52 is not null" & strCon & _
               " Union " & _
               "select s1.st01 st01,s1.st02,s1.st03,upper(nvl(s2.st17,s1.st17))||'/'||s1.st17 st17,s1.st53,s1.st16 st16,s1.st17 st01_17,s1.st51 st51" & _
               " from staff s1,staff s2" & _
               " where s1.st03='F21' and substr(s1.st01,1,1)<>'F' and s1.st01<>'F4102' and s1.st01<>'F4104' and s1.st01<>'F4105'" & _
               " and s1.st17 is not null" & _
               " and s1.st53=s2.st01(+) and s1.st53 is not null" & strCon & _
               " Union " & _
               "select s1.st01 st01,s1.st02,s1.st03,upper(nvl(s2.st17,s1.st17))||'/'||s1.st17 st17,s1.st54,s1.st16 st16,s1.st17 st01_17,s1.st51 st51" & _
               " from staff s1,staff s2" & _
               " where s1.st03='F21' and substr(s1.st01,1,1)<>'F' and s1.st01<>'F4102' and s1.st01<>'F4104' and s1.st01<>'F4105'" & _
               " and s1.st17 is not null" & _
               " and s1.st54=s2.st01(+) and s1.st54 is not null" & strCon
   'Modify By Sindy 2022/7/8
   If strST04 = "2" Then
      strExc(0) = strExc(0) & " order by st51 desc,st01 asc"
   Else
      strExc(0) = strExc(0) & " order by st01 asc"
   End If
   '2022/7/8 END
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         'Modify By Sindy 2022/7/8 英文縮寫加空白格後, 再搜尋主旨; 亦也會有 WC/as/wfc 這樣的縮寫, 故也增加/判斷
         '1. CH/YCP+空白格 2. WC/as/wfc
         'Modify By Sindy 2022/7/12 英文縮寫加) ex: 3. (EY/ey)
         If InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & " ") > 0 Or _
            (InStr(UCase(strText), "/" & UCase(rsTmp.Fields("st17"))) > 0 And InStr(UCase(rsTmp.Fields("st17")), "/") > 0) Or _
            InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & ")") > 0 Then
            
            strII18 = strII18 & ";" & UCase(rsTmp.Fields("st17")) & _
                           IIf(strST04 = "2", "(" & rsTmp.Fields("st01") & "已離職)", "") 'Add By Sindy 2021/1/19
            m_Emp = rsTmp.Fields("st01") 'Add By Sindy 2019/9/3
            'Add By Sindy 2019/9/3
            If "" & rsTmp.Fields("st16") = "1" Then '機電
               strDirector = Pub_GetSpecMan("T")
            ElseIf "" & rsTmp.Fields("st16") = "2" Then '化學
               strDirector = Pub_GetSpecMan("R")
            ElseIf "" & rsTmp.Fields("st16") = "3" Then '日文
               strDirector = Pub_GetSpecMan("S")
            ElseIf "" & rsTmp.Fields("st16") = "4" Then '德文
               strDirector = Pub_GetSpecMan("T1")
            End If
            
            'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "3" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
            m_ToSortOut = "3" '外專
            'Modify By Sindy 2024/8/27 排除離職人員，直接分給該縮寫組合之主管
            If strST04 = "1" Then m_Sender = m_Sender & ";" & rsTmp.Fields("st01")
            
'            If "" & rsTmp.Fields("st52") <> "" And "" & rsTmp.Fields("st52") <> rsTmp.Fields("st01") Then
'               m_Sender = m_Sender & ";" & rsTmp.Fields("st52")
'            End If

            'Modify By Sindy 2019/6/6 王文安不加發任何人
            If Not (rsTmp.Fields("st01") = "88003") Then
            '2019/6/6 END
               
               'Modify By Sindy 2022/5/19 加發工程師主管
               'Modify By Sindy 2024/4/29
               Call PUB_GetF21Manager("" & rsTmp.Fields("st16"), rsTmp.Fields("st01"), m_Sender)
'               strExc(10) = ""
'               strExc(10) = PUB_GetFCPEngSup(rsTmp.Fields("st01"))
'               If strExc(10) <> "" Then
'                  m_Sender = m_Sender & ";" & strExc(10)
'               End If
               '2024/4/29 END
               
               If rsTmp.Fields("st16") = "3" Then '日文
'                  'Add By Sindy 2019/5/9
'                  If "" & rsTmp.Fields("st16") = "3" Then '外專日文組工程師信件，要加發副理
'                     strExc(10) = PUB_GetST70SirEmp(rsTmp.Fields("st01"))
'                     If InStr(m_Sender, strExc(10)) = 0 Then
'                        m_Sender = m_Sender & ";" & strExc(10)
'                     End If
'                  End If
'                  '2019/5/9 END
                  '2022/5/19 END
                  
                  '日文加發承辦組Elaine
                  m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外專承辦日文組長") '99021
               Else
                  '非日文加發David,Elisa
                  'Modify By Sindy 2017/7/24 英文組長要加發Anny,Widen
                  m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外專承辦英文組長") '77015;A4011;A4024
               End If
            End If
            
            'Modify By Sindy 2022/3/23
            '潘子微(外專.主任.Anny):目前工程師的信函除了工程師自行E出外，其餘都由程序E出，像此信函寄出的sender為(程序wfc)-- HC/tt/wfc
            '故退信通知應包含寄出的sender(程序),以便sender即時處理
            If Trim(strII11) = "未傳遞的主旨" Then
               Call PUB_IPDept_ToSortOutSub_F22(strText, strII11, m_Sender, m_ToSortOut, strST04, UCase(rsTmp.Fields("st01_17")), , , strII18, False)
               If m_Sender <> "" Then Exit Do
               Call PUB_IPDept_ToSortOutSub_F22(strText, strII11, m_Sender, m_ToSortOut, strST04, UCase(rsTmp.Fields("st17")), , , strII18, False)
               If m_Sender <> "" Then Exit Do
            End If
            '2022/3/23 END
            Exit Do
         Else
            'Modify By Sindy 2019/8/30 未傳遞的主旨要再增加外專程序檢查
            If Trim(strII11) = "未傳遞的主旨" Then
               Call PUB_IPDept_ToSortOutSub_F22(strText, strII11, m_Sender, m_ToSortOut, strST04, UCase(rsTmp.Fields("st01_17")), , , strII18, False)
               If m_Sender <> "" Then Exit Do
               Call PUB_IPDept_ToSortOutSub_F22(strText, strII11, m_Sender, m_ToSortOut, strST04, UCase(rsTmp.Fields("st17")), , , strII18, False)
               If m_Sender <> "" Then Exit Do
            End If
            '2019/8/30 END
         End If
         rsTmp.MoveNext
      Loop
   End If
   
'   'Add By Sindy 2017/10/24 再檢查離職人員
'   If m_ToSortOut = "" And intRunTimes = 1 Then
'      intRunTimes = 2 '2.離職
'      strCon = "s1.st04='2'"
'      GoTo ReStarChk
'   End If
   
   Set rsTmp = Nothing
End Sub

'Add By Sindy 2019/9/3 依外專承辦英文縮寫分信
'Modify By Sindy 2021/1/19 + ,Optional ByRef strII18 As String
'Modify By Sindy 2022/7/8 + , ByVal strST04 As String: 1.在職 2.離職
Private Sub PUB_IPDept_ToSortOutSub_F23(ByVal strText As String, ByVal strII11 As String, ByRef m_Sender As String, _
                                 ByRef m_ToSortOut As String, ByVal strST04 As String, Optional ByRef m_bolF23EngW As Boolean, _
                                 Optional ByRef m_Emp As String = "", Optional ByRef strDirector As String = "", _
                                 Optional ByRef strII18 As String)
Dim rsTmp As New ADODB.Recordset
   
'    將檔名以大寫格式與專業代號 -國外ST17之大寫比較
'    1. 先讀取F字頭部門有ST17者，再讀取其主管之ST17，依撰寫信函之郵件寄送規則；
'       若信件檔名含此字串者，則收受者更新為該員工及其主管(分承辦主管或工程師主管)
'    2. 若無個人/主管之ST17代號，則抓個人之ST17代號判斷；
   '***外專承辦:ex.DY/elc
'   ST01     ST02         ST03 ST17                 ST52
'   -------- ------------ ---- -------------------- ------
'   77015    顏裕洋       F23  DY/dy
'   99021    陳毓芳       F23  DY/elc               77015
'   A1021    吳彩菱       F23  ELC/jw               99021
'   A1032    邱子瑜       F23  ELC/kc               99021
'   A4010    郭怡瑩       F23  ELC/mk               99021
'   A4011    潘子微       F23  DY/ap                77015
'   A4024    陳增廣       F23  DY/wc                77015
'   A5006    陳佩貞       F23  AP/lc                A4011
'   A5006    陳佩貞       F23  DY/lc                A4011
'   A5023    洪培堯       F23  DY/th                A4024
'   A5023    洪培堯       F23  WC/th                A4024
'   A6007    羅暐曄       F23  DY/jl                A4024
'   A6007    羅暐曄       F23  WC/jl                A4024
'   A6010    劉興杰       F23  AP/rl                A4011
'   A6010    劉興杰       F23  DY/rl                A4011
   '選取了 15 筆資料列.
   strExc(0) = "select st01,st02,st03,st17,st52" & _
               " from staff" & _
               " where st03='F23' and substr(st01,1,1)<>'F'" & _
               " and st17 is not null"
   'Modify By Sindy 2022/7/8
   If strST04 = "2" Then
      'Modify By Sindy 2020/11/18 抓離職2個月內的人員
      strExc(0) = strExc(0) & " and st04='2'" & _
                              " and st51 is not null and st51>=" & CompDate(1, -2, strSrvDate(1))
   Else
      strExc(0) = strExc(0) & " and st04='1'"
   End If
   '2022/7/8 END
   strExc(0) = strExc(0) & " union " & _
               "select s1.st01,s1.st02,s1.st03,upper(substr(s2.st17,instr(s2.st17,'/')+1))||substr(s1.st17,instr(s1.st17,'/')),s1.st52" & _
               " from staff s1,staff s2" & _
               " where s1.st03='F23' and substr(s1.st01,1,1)<>'F'" & _
               " and s2.st03='F23' and substr(s2.st01,1,1)<>'F'" & _
               " and s1.st17 is not null and s1.st52=s2.st01 and s2.st17 is not null" & _
               " and s1.st17<>upper(substr(s2.st17,instr(s2.st17,'/')+1))||substr(s1.st17,instr(s1.st17,'/'))"
   'Modify By Sindy 2022/7/8
   If strST04 = "2" Then
      'Modify By Sindy 2020/11/18 抓離職2個月內的人員
      strExc(0) = strExc(0) & " and s1.st04='2' and s2.st04='2' " & _
                              " and s1.st51 is not null and s1.st51>=" & CompDate(1, -2, strSrvDate(1)) & _
                              " and s2.st51 is not null and s2.st51>=" & CompDate(1, -2, strSrvDate(1))
   Else
      strExc(0) = strExc(0) & " and s1.st04='1' and s2.st04='1' "
   End If
   '2022/7/8 END
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         'Modify By Sindy 2022/7/8 英文縮寫加空白格後, 再搜尋主旨; 亦也會有 WC/as/wfc 這樣的縮寫, 故也增加/判斷
         '1. CH/YCP+空白格 2. WC/as/wfc
         'Modify By Sindy 2022/7/12 英文縮寫加) ex: 3. (EY/ey)
         If InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & " ") > 0 Or _
            (InStr(UCase(strText), "/" & UCase(rsTmp.Fields("st17"))) > 0 And InStr(UCase(rsTmp.Fields("st17")), "/") > 0) Or _
            InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & ")") > 0 Then
            
            'Modify By Sindy 2024/8/27 皆已離職，需人工分信
            If strST04 = "2" _
               And (ChkStaffST04("" & rsTmp.Fields("st52"), False) = True Or "" & rsTmp.Fields("st52") = "") Then
               GoTo ReadNext
            End If
            '2024/8/27 END
            
            strII18 = strII18 & ";" & UCase(rsTmp.Fields("st17")) & _
                           IIf(strST04 = "2", "(" & rsTmp.Fields("st01") & "已離職)", "") 'Add By Sindy 2021/1/19
            m_Emp = rsTmp.Fields("st01") 'Add By Sindy 2019/9/3
            strDirector = "" & rsTmp.Fields("ST52") 'Add By Sindy 2019/9/3
            'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "3" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
            m_ToSortOut = "3" '外專
            
            'Modify By Sindy 2024/8/27 排除離職人員，直接分給該縮寫組合之主管
            If strST04 = "1" Then m_Sender = m_Sender & ";" & rsTmp.Fields("st01")
            
            If "" & rsTmp.Fields("st52") <> "" Then
               m_Sender = m_Sender & ";" & rsTmp.Fields("st52")
            End If
            m_bolF23EngW = True 'Add By Sindy 2016/5/30
            Exit Do
         End If
ReadNext:
         rsTmp.MoveNext
      Loop
   End If
   'rsTmp.Close
'   'Add By Sindy 2017/10/24 再檢查離職人員
'   'Modify By Sindy 2020/11/18 抓離職2個月內的人員
'   If m_ToSortOut = "" Then
'      strExc(0) = "select st01,st02,st03,st17,st52" & _
'                  " from staff" & _
'                  " where st04='2' and st03='F23' and substr(st01,1,1)<>'F'" & _
'                  " and st17 is not null and st51 is not null" & _
'                  " and st51>=" & CompDate(1, -2, strSrvDate(1)) & _
'                  " order by st51 desc"
'      intI = 1
'      Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         rsTmp.MoveFirst
'         Do While Not rsTmp.EOF
'            If InStr(UCase(strText), UCase(rsTmp.Fields("st17"))) > 0 Then
'               strII18 = strII18 & ";" & UCase(rsTmp.Fields("st17")) 'Add By Sindy 2021/1/19
'               m_Emp = rsTmp.Fields("st01") 'Add By Sindy 2019/9/3
'               strDirector = "" & rsTmp.Fields("ST52") 'Add By Sindy 2019/9/3
'               m_ToSortOut = "3" '外專
'               m_Sender = m_Sender & ";" & rsTmp.Fields("st01")
'               If "" & rsTmp.Fields("st52") <> "" Then
'                  m_Sender = m_Sender & ";" & rsTmp.Fields("st52")
'               End If
'               m_bolF23EngW = True 'Add By Sindy 2016/5/30
'               Exit Do
'            End If
'            rsTmp.MoveNext
'         Loop
'      End If
'   End If
'   '2017/10/24 END

   Set rsTmp = Nothing
End Sub

'Add By Sindy 2019/8/30 依外專程序英文縮寫分信
'Modify By Sindy 2021/1/19 + ,Optional ByRef strII18 As String
'Modify By Sindy 2022/7/8 + , ByVal strST04 As String: 1.在職 2.離職
'Modify By Sindy 2022/7/8 + , Optional ByVal bolAddF23 As Boolean = True: 要加F23承辦群組; 已抓到外專工程師英文縮寫時,呼叫此函數時,不用再加承辦群組
Private Sub PUB_IPDept_ToSortOutSub_F22(ByVal strText As String, ByVal strII11 As String, ByRef m_Sender As String, _
                                 ByRef m_ToSortOut As String, ByVal strST04 As String, ByVal strQuyText As String, _
                                 Optional ByRef m_Emp As String = "", Optional ByRef strDirector As String = "", _
                                 Optional ByRef strII18 As String, Optional ByVal bolAddF23 As Boolean = True)
Dim rsTmp As New ADODB.Recordset
Dim strCon As String 'Add By Sindy 2022/7/8
   
   'Modify By Sindy 2022/7/8
   If strST04 = "2" Then
      'Modify By Sindy 2020/11/18 抓離職2個月內的人員
      strCon = " and st04='2'" & _
               " and st51 is not null and st51>=" & CompDate(1, -2, strSrvDate(1))
   Else
      strCon = " and st04='1'"
   End If
   '2022/7/8 END
   
   If Trim(strQuyText) = "" Then
      'Modify By Sindy 2022/11/30 DY/ => 改抓特殊設定
      'strQuyText = "DY"
      Call GetPrjSalesNM(Pub_GetSpecMan("外專承辦英文組主管"), , strQuyText) '外專承辦(DY/dy)
      If InStr(strQuyText, "/") > 0 Then
         strQuyText = UCase(Left(strQuyText, InStr(strQuyText, "/") - 1))
      End If
      '2022/11/30 END
   End If
   
'***外專程序:ex.DY/sh
   ''DY/'||st17 st17 =>'" & strQuyText & "/'||st17 st17
   strExc(0) = "select st01,st02,st03,'" & strQuyText & "/'||st17 st17,st52" & _
               " from staff" & _
               " where st03='F22' and substr(st01,1,1)<>'F'" & _
               " and st17 is not null" & strCon
   'Modify By Sindy 2022/7/8
   If strST04 = "2" Then
      strExc(0) = strExc(0) & " order by st51 desc"
   End If
   '2022/7/8 END
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         'Modify By Sindy 2022/7/8 英文縮寫加空白格後, 再搜尋主旨; 亦也會有 WC/as/wfc 這樣的縮寫, 故也增加/判斷
         '1. CH/YCP+空白格 2. WC/as/wfc
         'Modify By Sindy 2022/7/12 英文縮寫加) ex: 3. (EY/ey)
         If InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & " ") > 0 Or _
            (InStr(UCase(strText), "/" & UCase(rsTmp.Fields("st17"))) > 0 And InStr(UCase(rsTmp.Fields("st17")), "/") > 0) Or _
            InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & ")") > 0 Then
            
            'Modify By Sindy 2024/8/27 皆已離職，需人工分信
            If strST04 = "2" _
               And (ChkStaffST04("" & rsTmp.Fields("st52"), False) = True Or "" & rsTmp.Fields("st52") = "") Then
               GoTo ReadNext
            End If
            '2024/8/27 END
            
            strII18 = strII18 & ";" & UCase(rsTmp.Fields("st17")) & _
                           IIf(strST04 = "2", "(" & rsTmp.Fields("st01") & "已離職)", "") 'Add By Sindy 2021/1/19
            m_Emp = rsTmp.Fields("st01") 'Add By Sindy 2019/9/3
            strDirector = "" & rsTmp.Fields("ST52") 'Add By Sindy 2019/9/3
            'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "3" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
            'm_ToSortOut = "3" '外專
            If m_ToSortOut = "" Then m_ToSortOut = "3" '外專
            
            'Modify By Sindy 2024/8/27 排除離職人員，直接分給該縮寫組合之主管
            If strST04 = "1" Then m_Sender = m_Sender & ";" & rsTmp.Fields("st01")
            
            If "" & rsTmp.Fields("st52") <> "" Then
               m_Sender = m_Sender & ";" & rsTmp.Fields("st52")
            End If
            'Add By Sindy 2019/8/30 未傳遞的主旨略過
            If Trim(strII11) <> "未傳遞的主旨" Then
            '2019/8/30 END
               If bolAddF23 = True Then
                  m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外專群組") 'Modify By Sindy 2016/4/14 加發承辦組主管(77015;96022;99021)
               End If
            End If
            Exit Do
         End If
ReadNext:
         rsTmp.MoveNext
      Loop
   End If
   'rsTmp.Close
'   'Add By Sindy 2017/10/24 再檢查離職人員
'   'Modify By Sindy 2020/11/18 抓離職2個月內的人員
'   If m_ToSortOut = "" Then
'      strExc(0) = "select st01,st02,st03,'" & strQuyText & "/'||st17 st17,st52" & _
'                  " from staff" & _
'                  " where st04='2' and st03='F22' and substr(st01,1,1)<>'F'" & _
'                  " and st17 is not null and st51 is not null" & _
'                  " and st51>=" & CompDate(1, -2, strSrvDate(1)) & _
'                  " order by st51 desc"
'      intI = 1
'      Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         rsTmp.MoveFirst
'         Do While Not rsTmp.EOF
'            If InStr(UCase(strText), UCase(rsTmp.Fields("st17"))) > 0 Then
'               strII18 = strII18 & ";" & UCase(rsTmp.Fields("st17")) 'Add By Sindy 2021/1/19
'               m_Emp = rsTmp.Fields("st01") 'Add By Sindy 2019/9/3
'               strDirector = "" & rsTmp.Fields("ST52") 'Add By Sindy 2019/9/3
'               'If PUB_IPDept_ToSortOut <> "" And PUB_IPDept_ToSortOut <> "3" Then PUB_IPDept_ToSortOut = "Z": GoTo ChkEnd '符合一個條件以上,歸其他
'               m_ToSortOut = "3" '外專
'               m_Sender = m_Sender & ";" & rsTmp.Fields("st01")
'               If "" & rsTmp.Fields("st52") <> "" Then
'                  m_Sender = m_Sender & ";" & rsTmp.Fields("st52")
'               End If
'               'Add By Sindy 2019/8/30 未傳遞的主旨略過
'               If Trim(strII11) <> "未傳遞的主旨" Then
'               '2019/8/30 END
'                  m_Sender = m_Sender & ";" & Pub_GetSpecMan("國外部轉信外專群組") 'Modify By Sindy 2016/4/14 加發承辦組主管(77015;96022;99021)
'               End If
'               Exit Do
'            End If
'            rsTmp.MoveNext
'         Loop
'      End If
'   End If

   Set rsTmp = Nothing
End Sub

'Add By Sindy 2019/9/3 依國外業務拓展英文縮寫分信
'Modify By Sindy 2021/1/19 + ,Optional ByRef strII18 As String
'Modify By Sindy 2022/7/8 + , ByVal strST04 As String: 1.在職 2.離職
Private Sub PUB_IPDept_ToSortOutSub_F41(ByVal strText As String, ByVal strII11 As String, ByRef m_Sender As String, _
                                 ByRef m_ToSortOut As String, ByVal strST04 As String, _
                                 Optional ByRef m_Emp As String = "", Optional ByRef strDirector As String = "", _
                                 Optional ByRef strII18 As String)
Dim rsTmp As New ADODB.Recordset
Dim strCon As String 'Add By Sindy 2022/7/8
   
   'Modify By Sindy 2022/7/8
   If strST04 = "2" Then
      'Modify By Sindy 2020/11/18 抓離職2個月內的人員
      strCon = " and st04='2'" & _
               " and st51 is not null and st51>=" & CompDate(1, -2, strSrvDate(1))
   Else
      strCon = " and st04='1'"
   End If
   '2022/7/8 END
   
'***國外部開拓:ex.EY/wc
'   ST01     ST02         ST03 ST17       ST52
'   -------- ------------ ---- ---------- ------
'   99033    楊雯芳       F41  EY/ey
'   A4024    陳增廣       F41  EY/wc      99033
   strExc(0) = "select st01,st02,st03,st17,st52" & _
               " from staff" & _
               " where st03='F41' and substr(st01,1,1)<>'F'" & _
               " and st17 is not null" & strCon
   'Modify By Sindy 2022/7/8
   If strST04 = "2" Then
      strExc(0) = strExc(0) & " order by st51 desc"
   End If
   '2022/7/8 END
   intI = 1
   Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         'Modify By Sindy 2022/7/8 英文縮寫加空白格後, 再搜尋主旨; 亦也會有 WC/as/wfc 這樣的縮寫, 故也增加/判斷
         '1. CH/YCP+空白格 2. WC/as/wfc
         'Modify By Sindy 2022/7/12 英文縮寫加) ex: 3. (EY/ey)
         If InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & " ") > 0 Or _
            (InStr(UCase(strText), "/" & UCase(rsTmp.Fields("st17"))) > 0 And InStr(UCase(rsTmp.Fields("st17")), "/") > 0) Or _
            InStr(UCase(strText), UCase(rsTmp.Fields("st17")) & ")") > 0 Then
            
            strII18 = strII18 & ";" & UCase(rsTmp.Fields("st17")) 'Add By Sindy 2021/1/19
            m_Emp = rsTmp.Fields("st01") 'Add By Sindy 2019/9/3
            strDirector = "" & rsTmp.Fields("ST52") 'Add By Sindy 2019/9/3
            m_ToSortOut = "8" '開拓
            m_Sender = "" 'Pub_GetSpecMan("國外部轉信開拓群組")
            Exit Do
         End If
         rsTmp.MoveNext
      Loop
      'Modify By Sindy 2018/1/5 分給開拓全組人員
      If m_ToSortOut = "8" And m_Sender = "" Then
         strExc(0) = "select st01,st02,st03,st17,st52" & _
                     " from staff" & _
                     " where st04='1' and st03='F41' and substr(st01,1,1)<>'F'" & _
                     " and st17 is not null"
         intI = 1
         Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            rsTmp.MoveFirst
            Do While Not rsTmp.EOF
               m_Sender = m_Sender & ";" & rsTmp.Fields("st01")
               rsTmp.MoveNext
            Loop
         End If
      End If
      '2018/1/5 END
   End If
   
'   '再檢查離職人員
'   'Modify By Sindy 2020/11/18 抓離職2個月內的人員
'   If m_ToSortOut = "" Then
'      strExc(0) = "select st01,st02,st03,st17,st52" & _
'                  " from staff" & _
'                  " where st04='1' and st03='F41' and substr(st01,1,1)<>'F'" & _
'                  " and st17 is not null and st51 is not null" & _
'                  " and st51>=" & CompDate(1, -2, strSrvDate(1)) & _
'                  " order by st51 desc"
'      intI = 1
'      Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'         rsTmp.MoveFirst
'         Do While Not rsTmp.EOF
'            If InStr(UCase(strText), UCase(rsTmp.Fields("st17"))) > 0 Then
'               strII18 = strII18 & ";" & UCase(rsTmp.Fields("st17")) 'Add By Sindy 2021/1/19
'               m_Emp = rsTmp.Fields("st01") 'Add By Sindy 2019/9/3
'               strDirector = "" & rsTmp.Fields("ST52") 'Add By Sindy 2019/9/3
'               m_ToSortOut = "8" '開拓
'               m_Sender = "" 'Pub_GetSpecMan("國外部轉信開拓群組")
'               Exit Do
'            End If
'            rsTmp.MoveNext
'         Loop
'         'Modify By Sindy 2018/1/5 分給開拓全組人員
'         If m_ToSortOut = "8" And m_Sender = "" Then
'            strExc(0) = "select st01,st02,st03,st17,st52" & _
'                        " from staff" & _
'                        " where st04='1' and st03='F41' and substr(st01,1,1)<>'F'" & _
'                        " and st17 is not null"
'            intI = 1
'            Set rsTmp = ClsLawReadRstMsg(intI, strExc(0))
'            If intI = 1 Then
'               rsTmp.MoveFirst
'               Do While Not rsTmp.EOF
'                  m_Sender = m_Sender & ";" & rsTmp.Fields("st01")
'                  rsTmp.MoveNext
'               Loop
'            End If
'         End If
'         '2018/1/5 END
'      End If
'   End If

   Set rsTmp = Nothing
End Sub

'Add By Sindy 2018/7/5 解析主旨抓出對應的案件性質
Public Sub PUB_IPDept_ComparisonCP(ByVal strTextSubject As String, ByVal strFileName As String, _
   ByVal strCP01 As String, ByVal strCP02 As String, ByVal strCP03 As String, ByVal strCP04 As String, _
   ByRef strII03_2 As String, _
   ByRef strCP09 As String, ByRef strCP10 As String)
Dim intStar As Integer, intEnd As Integer
Dim strProc As String
Dim rsA As New ADODB.Recordset
Dim RsQ As New ADODB.Recordset
Dim bolExistsPROC As Boolean 'Add By Sindy 2018/12/27
   
   '解析主旨使用
   strTextSubject = strTextSubject
   strTextSubject = Replace(strTextSubject, "．", ".")
   strTextSubject = Replace(strTextSubject, "..", ".")
   strTextSubject = Replace(strTextSubject, "...", ".")
   'Add By Sindy 2018/5/16 歸入正確的案件性質
   If strII03_2 = "" Then
      If InStr(UCase(strTextSubject), UCase("[PROC.")) > 0 Then
         intStar = InStr(UCase(strTextSubject), UCase("[PROC.")) + Len(UCase("[PROC."))
         intEnd = InStr(intStar, UCase(strTextSubject), UCase("]"))
         If intEnd > 0 And intEnd > intStar Then
            strProc = Trim(Mid(UCase(strTextSubject), intStar, intEnd - intStar))
            'Modify By Sindy 2019/2/22 + And strCP01 <> "" And strCP02 <> ""
            If IsNumeric(strProc) = True And strCP01 <> "" And strCP02 <> "" Then
               'Modify By Sindy 2018/7/2 and cp01=cpm01(+) ==> and 'FCP'=cpm01(+)
               strExc(0) = "select cp09,cp10,cpm26 from caseprogress,casepropertymap" & _
                           " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                           " and cp10='" & strProc & "'" & _
                           " and 'FCP'=cpm01(+) and cp10=cpm02(+)" & _
                           " order by nvl(cp66,cp05) desc,cp67 desc"
               intI = 1
               Set rsA = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  strCP09 = rsA.Fields("cp09")
                  strCP10 = rsA.Fields("cp10")
                  bolExistsPROC = True 'Add By Sindy 2018/12/27
                  If strFileName <> "" Then
                     If "" & rsA.Fields("cpm26") <> "" Then
                        strII03_2 = Replace(strFileName, ".msg", "." & rsA.Fields("cpm26") & ".msg")
                     'Modify By Sindy 2018/12/27 Mark
'                     Else
'                        strII03_2 = Replace(strFileName, ".msg", ".tx.msg")
                     End If
                  End If
               End If
            End If
         End If
      End If
      'Modify By Sindy 2018/12/27 + And bolExistsPROC = False
      'Modify By Sindy 2019/2/22 + or instr(upper('" & ChgSQL(strTextSubject) & "'),upper('['||EFC02||']'))>0)
      If strII03_2 = "" And bolExistsPROC = False Then
         strExc(0) = "select EFC01,EFC02 from efilecaption where EFC06='Y'" & _
                     " and (instr(upper('" & ChgSQL(strTextSubject) & "'),upper('['||EFC02||'.'))>0 or instr(upper('" & ChgSQL(strTextSubject) & "'),upper('['||EFC02||']'))>0)" & _
                     " and efc01 in('ALL','" & strCP01 & "') and efc02<>'PROC'"
         intI = 1
         Set RsQ = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            RsQ.MoveFirst
            Do While Not RsQ.EOF
               'Add By Sindy 2019/2/23
               If InStr(UCase(strTextSubject), UCase("[" & RsQ.Fields("EFC02") & "]")) = 0 Then
               '2019/2/23 END
                  intStar = InStr(UCase(strTextSubject), UCase("[" & RsQ.Fields("EFC02") & ".")) + Len(UCase("[" & RsQ.Fields("EFC02") & "."))
                  intEnd = InStr(intStar, UCase(strTextSubject), UCase("]"))
                  If intEnd > 0 And intEnd > intStar Then
                     strProc = Trim(Mid(UCase(strTextSubject), intStar, intEnd - intStar))
                     'Modify By Sindy 2019/2/22 + And strCP01 <> "" And strCP02 <> ""
                     If IsNumeric(strProc) = True And strCP01 <> "" And strCP02 <> "" Then
                        strExc(0) = "select cp09,cp10 from caseprogress" & _
                                    " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                                    " and cp10='" & strProc & "'" & _
                                    " order by nvl(cp66,cp05) desc,cp67 desc"
                        intI = 1
                        Set rsA = ClsLawReadRstMsg(intI, strExc(0))
                        If intI = 1 Then
                           If strFileName <> "" Then
                              strII03_2 = Replace(strFileName, ".msg", "." & RsQ.Fields("EFC02") & ".msg")
                           End If
                           strCP09 = rsA.Fields("cp09")
                           strCP10 = rsA.Fields("cp10")
                           Exit Do
                        End If
                     End If
                  End If
               End If
               RsQ.MoveNext
            Loop
            'Modify By Sindy 2018/6/29 David:有可能在寄信時,還沒有進度資料
            If strII03_2 = "" Then
               RsQ.MoveFirst
               If strFileName <> "" Then
                  strII03_2 = Replace(strFileName, ".msg", "." & RsQ.Fields("EFC02") & ".msg")
               'Add By Sindy 2019/2/22
               Else
                  strII03_2 = RsQ.Fields("EFC02")
                  '2019/2/22 END
               End If
            End If
            '2018/6/29 END
         End If
      End If
   End If
   'Modify By Sindy 2018/10/5 Mark
   'If UCase(strRecipients_1) = UCase("backup") Then '收件者為backup;代表信件純為歸卷宗區
   '2018/10/5 END
   If strII03_2 = "" Then
      If InStr(UCase(strTextSubject), UCase("[紙本寄出]")) > 0 Then '紙本寄出
         If strFileName <> "" Then
            strII03_2 = Replace(strFileName, ".msg", ".paper.msg")
         End If
      ElseIf InStr(UCase(strTextSubject), UCase("[平台下載]")) > 0 Then '平台下載
         If strFileName <> "" Then
            strII03_2 = Replace(strFileName, ".msg", ".dnl.msg")
         End If
      ElseIf InStr(UCase(strTextSubject), UCase("[平台上傳]")) > 0 Then '平台上傳
         If strFileName <> "" Then
            strII03_2 = Replace(strFileName, ".msg", ".upl.msg")
         End If
      'Add By Sindy 2018/12/27
      Else
         If strFileName <> "" Then
            strII03_2 = Replace(strFileName, ".msg", ".tx.msg") '寄出郵件
         End If
      '2018/12/27 END
      End If
   End If
   
   Set rsA = Nothing
   Set RsQ = Nothing
End Sub

'Add By Sindy 2017/5/10
'Modify By Sindy 2017/10/30 + ByVal strType As String
'strType:1.申請人名稱
'        2.代理人
Public Function PUB_FilterBulletinSpecWord(ByVal strType As String, ByVal strName As String, ByVal strNationNm As String) As String
Dim strCompText As String
Dim ii As Integer
   
   If strType = "1" Then '申請人名稱
      strName = Replace(strName, "‧", "．")
      strName = Replace(strName, "˙", "．")
      'Modify By Sindy 2018/10/17 Mark,不去掉字樣,保留原申請人名稱
'      If InStr(strName, "商．") > 0 Then
'         strName = Mid(strName, InStr(strName, "商．") + 2)
'      ElseIf InStr(strName, "區．") > 0 Then
'         strName = Mid(strName, InStr(strName, "區．") + 2)
'      ElseIf InStr(strName, "籍．") > 0 Then
'         strName = Mid(strName, InStr(strName, "籍．") + 2)
'      Else
'         For ii = 1 To 13
'            'Modify By Sindy 2018/10/12 申請人名稱,公司時有可能真的有台灣字樣在公司名稱前
'            'ex:台灣積體電路製造股份有限公司 (106135670 / I637464)
'            strCompText = ""
'            'If ii = 1 Then strCompText = "台灣"
'            '2018/10/12 END
'            If ii = 2 Then strCompText = "美商"
'            If ii = 3 Then strCompText = "英商"
'            If ii = 4 Then strCompText = "日商"
'            If ii = 5 Then strCompText = "法商"
'            If ii = 6 Then strCompText = "德商"
'            If ii = 7 Then strCompText = "德籍"
'            If ii = 8 Then strCompText = "韓籍"
'            If ii = 9 Then strCompText = "韓商"
'            If ii = 10 Then strCompText = "南韓商"
'            If ii = 11 Then strCompText = "開曼群島商"
'            If ii = 12 Then strCompText = "塞席爾商"
'            If ii = 13 Then strCompText = "大陸地區"
'            If strCompText <> "" Then 'Add By Sindy 2018/10/12 + if判斷
'               If Left(strName, Len(strCompText)) = strCompText Then
'                  strName = Replace(strName, strCompText, "")
'                  GoTo LoadComp
'               End If
'            End If
'         Next ii
'         If strNationNm <> "" Then
'            If Left(strName, Len(Trim(strNationNm) & "商")) = Trim(strNationNm) & "商" Then
'               strName = Replace(strName, strNationNm & "商", "")
'               GoTo LoadComp
'            End If
'            If Left(strName, Len(Trim(strNationNm) & "籍")) = Trim(strNationNm) & "籍" Then
'               strName = Replace(strName, strNationNm & "籍", "")
'               GoTo LoadComp
'            End If
'         End If
'      End If
   End If
   
   'Add By Sindy 2017/5/4
LoadComp:
   
   'Modify By Sindy 2022/4/13
   'PUB_FilterBulletinSpecWord = strName
   Forms(0).TextComp.Text = strName '為顯示出Unicode的?
   PUB_FilterBulletinSpecWord = Forms(0).TextComp.Text
   'Modify By Sindy 2022/4/13 在造字還沒全部更新時, 還是先比對此檔案
   '但因前頭程式文字檔已支援Unicode的字碼, 此if判斷不出來, 先Mark
   If InStr(PUB_FilterBulletinSpecWord, "?") > 0 Then '有特殊字時,進行比對
   '2022/4/13 END
      strSql = "select BS03 from BulletinSpecWord WHERE BS01='" & strType & "' and BS02='" & PUB_FilterBulletinSpecWord & "'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         PUB_FilterBulletinSpecWord = RsTemp.Fields("BS03")
      End If
   Else
      PUB_FilterBulletinSpecWord = strName
   End If
   '2017/5/4 END
End Function

'從frm06010301_1各式申請書-電子送件-新案 抽出來變共用函數 Modify By Sindy 2017/9/25
'Modify By Sindy 2018/12/24 + Optional ByVal bolShowPage As Boolean = True
'Modify By Sindy 2019/2/1 + , Optional ByVal p_CP01 As String, Optional ByVal p_ET01 As String, Optional ByVal p_ET03 As String
'Modified by Lydia 2020/09/25 增加分節處理頁碼 + Optional ByVal bolSectionPages As Boolean
Public Function PUB_MakeDoc(ByVal p_Text As String, ByVal p_Name As String, _
   Optional ByVal bolShowPage As Boolean = True, Optional ByVal p_Cp01 As String, _
   Optional ByVal p_ET01 As String, Optional ByVal p_ET03 As String, Optional ByVal bolSectionPages As Boolean) As Boolean
   
   Dim b2Time As Boolean
   
   p_Text = PUB_Big5toUnicode(p_Text)
   
On Error GoTo ErrHnd
   
   If TypeName(g_WordAp) <> "Application" Then
      Set g_WordAp = New Word.Application
   End If
   
   With g_WordAp
      .Visible = True
      .Documents.add

      .Selection.PageSetup.LeftMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.RightMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.TopMargin = .CentimetersToPoints(2)
      .Selection.PageSetup.BottomMargin = .CentimetersToPoints(2)
      'Add By Sindy 2019/2/1
      If p_Cp01 = "FCP" And p_ET01 = "01" And InStr("31,32,33,34,35", p_ET03) > 0 Then
         .Selection.ParagraphFormat.DisableLineHeightGrid = False
         .Selection.Font.Name = "標楷體"
         .Selection.PageSetup.Orientation = wdOrientPortrait
         .Selection.Orientation = wdTextOrientationHorizontal
         .Selection.Font.Size = 14
      Else
      '2019/2/1 END
         .Selection.ParagraphFormat.DisableLineHeightGrid = True
         'Add By Sindy 2018/4/30
         .Selection.Font.Name = "新細明體"
         .Selection.Font.Name = "Times New Roman"
         '2018/4/30 END
      End If
      
      'Add By Sindy 2018/12/24
      If bolShowPage = True Then
      '2018/12/24 END
         'Added by Morgan 2017/8/9
         .ActiveDocument.Repaginate
         If .ActiveWindow.View.SplitSpecial = wdPaneNone Then
            .ActiveWindow.ActivePane.View.Type = wdPageView
         Else
            .ActiveWindow.View.Type = wdPageView
         End If
         .ActiveWindow.ActivePane.View.SeekView = wdSeekCurrentPageFooter
         .Selection.TypeParagraph
         'Added by Lydia 2020/09/25 增加分節處理: 因為外商承辦不能將申請書和基本資料表分開轉檔，所以只能合併為一個檔案經過轉檔分別產生2個PDF；頁尾的「第x頁，共x頁」，以Word的章節內容計算頁數。
         If bolSectionPages = True Then
            '要注意Word內文的換頁符號Chr(12)，視情況替換為分節符號 "|#(分節)#|"
            '設頁碼格式不接續前一章節
            .Selection.HeaderFooter.PageNumbers.RestartNumberingAtSection = True
            .Selection.HeaderFooter.PageNumbers.StartingNumber = 1
            '------------------------------------
            .Selection.TypeText Text:="第"
            .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldPage
            .Selection.TypeText Text:="頁，共"
            .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldSectionPages, Text:="SECTIONPAGES ", PreserveFormatting:=True
            .Selection.TypeText Text:="頁"
         Else
         'end 2020/09/25
            .Selection.TypeText Text:="第"
            .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldPage
            .Selection.TypeText Text:="頁，共"
            .Selection.Fields.add Range:=.Selection.Range, Type:=wdFieldEmpty, Text:="NUMPAGES ", PreserveFormatting:=True
            .Selection.TypeText Text:="頁"
         End If 'Added by Lydia 2020/09/25
         .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
         .ActiveWindow.ActivePane.View.SeekView = wdSeekMainDocument
         'end 2017/8/9
      End If
      
      .ActiveWindow.Selection.TypeText p_Text
      
      ChgWordFormat g_WordAp, p_Text 'Added by Morgan 2018/1/12
      
      If p_Name <> "" Then 'Added by Lydia 2019/07/10 判斷沒有檔案名稱，不用存檔(內商電子送件)
            'Added by Morgan 2017/7/6 先存doc格式,可能需要手動轉pdf
            .ActiveDocument.SaveAs p_Name & ".doc", FileFormat:=0
            'end 2017/7/6
            
            'Removed by Morgan 2017/8/9 目前 Html2Pdf 有問題,暫取消
            'Word 2007 以上
            'If Val(.Version) >= 12 Then
            '   .ActiveDocument.SaveAs p_Name & ".htm", FileFormat:=8
            'Else
            '   .ActiveDocument.SaveAs p_Name & ".htm", FileFormat:=100
            'End If
            '.ActiveDocument.Save
            'end 2017/8/9
            
            .ActiveDocument.Close wdDoNotSaveChanges
      End If 'end 2019/07/10
   End With
   
   If p_Name <> "" Then 'Added by Lydia 2019/07/10 判斷沒有檔案名稱，不用存檔(內商電子送件)
      g_WordAp.Quit wdDoNotSaveChanges
      Set g_WordAp = Nothing
   End If
   
   PUB_MakeDoc = True
   Exit Function
   
ErrHnd:

   Select Case Err.Number
      Case 91:
         g_WordAp.Documents.add
         Resume Next
      Case 462:
         Set g_WordAp = New Word.Application
         g_WordAp.Documents.add
         Resume Next
      Case Else:
         MsgBox "錯誤 : " & Err.Description, vbCritical
         Exit Function
   End Select
End Function

'從frm06010301_1各式申請書-電子送件-新案 抽出來變共用函數 Modify By Sindy 2017/9/25
'Modify By Sindy 2018/4/19 + 增加判讀是否有變更地址是否要抓個案資料 : Optional bolChageAddr As Boolean = False
'Modify By Sindy 2018/10/19 + , Optional strApplNum As String : 欲讀取的申請人資料
'Modify By Sindy 2018/10/19 + , Optional strRepresentative As String : 欲讀取的代表人資料
'Modify By Sindy 2018/12/5 + , Optional bolShowEng As Boolean = False : 顯示英文資料
'Modify By Sindy 2019/1/22 + , Optional ByRef strNA81Appl As String = "" : 回傳有外商國名者
Public Function StartLetterPA_EData(ByVal ET01 As String, ByVal ET03 As String, ByVal strReceiveNo As String, _
   pa() As String, cp() As String, Optional bolShowInvEmp As Boolean = True, _
   Optional bolChageAddr As Boolean = False, Optional strApplNum As String, _
   Optional strRepresentative As String, Optional bolShowEng As Boolean = False, Optional ByRef strNA81Appl As String = "") As Boolean
   
   Dim strTxt(110) As String, strTmp As String
   Dim ii As Integer, jj As Integer
   Dim strInventor As String 'Add By Sindy 2014/11/14
   
   ii = 0
   EndLetter ET01, strReceiveNo, ET03, strUserNum
   
   'Modify By Sindy 2017/11/15
   'Modify By Sindy 2018/12/5 + bolShowEng
   'Modify By Sindy 2019/1/22 + strNA81Appl
   Call PUB_GetApplPA_EData(ET01, ET03, strReceiveNo, pa(), IIf(bolChageAddr = True, False, True), strApplNum, strRepresentative, bolShowEng, strNA81Appl)
   
   '預設出名代理人
   Dim lstNameAgent As ListBox
   If cp(110) = "" Then PUB_SetOurAgent lstNameAgent, pa(), cp(110), cp(10)
   
   'Modify By Sindy 2020/4/8 申請書:出名代理人
   Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 2, cp(110), ET01, strReceiveNo, ET03, ii, strTxt())
   
   'Add By Sindy 2017/10/24 電子送件的補正,基本資料表不顯示發明人資料
   If bolShowInvEmp = True Then
      ii = ii + 1
      'Modify By Sindy 2022/4/29
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明人資料','" & strInventor & "')"
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','發明人要印','♀')"
      '2022/4/29 END
   End If
   '2014/11/14 END
   
   'Add By Sindy 2025/7/7
   strTmp = "專利申請人"
   strSql = "select pa15 from patent" & _
            " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      If Trim("" & RsTemp.Fields("pa15")) <> "" Then
         strTmp = "專利權人"
      End If
   End If
   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','身分類別','" & strTmp & "')"
   '2025/7/7 END
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      StartLetterPA_EData = True
   End If
End Function

'Add By Sindy 2017/11/15
'各式申請書-電子送件:申請人
'Modify By Sindy 2018/4/19 + 增加判讀是否要抓個案資料 : Optional bolReadPA As Boolean = True
'Modify By Sindy 2018/10/19 + , Optional strApplNum As String : 欲讀取的申請人資料
'Modify By Sindy 2018/10/19 + , Optional strRepresentative As String : 欲讀取的代表人資料
'Modify By Sindy 2018/12/5 + , Optional bolShowEng As Boolean = False : 顯示英文資料
'Modify By Sindy 2019/1/22 + , Optional ByRef strNA81Appl As String = "" : 回傳有外商國名者
Public Function PUB_GetApplPA_EData(ByVal ET01 As String, ByVal ET03 As String, ByVal strReceiveNo As String, _
   pa() As String, Optional bolReadPA As Boolean = True, Optional strApplNum As String, _
   Optional strRepresentative As String, Optional bolShowEng As Boolean = False, _
   Optional ByRef strNA81Appl As String = "") As Boolean
   
   Dim strTxt(110) As String, strTmp As String, strTmp2 As String
   Dim ii As Integer, jj As Integer
   Dim strChaName As String, strEngName As String
   Dim kk As Integer, k_star As Integer, k_end As Integer, intRow As Integer
   Dim strApplEmp(1 To 5) As String, varTemp As Variant 'Add By Sindy 2018/10/19
   Dim strRepresentativeEmp(1 To 30) As String 'Add By Sindy 2018/10/19
   Dim idx As Integer 'Add By Sindy 2018/10/22
   Dim strCP10 As String 'Add By Sindy 2019/12/13
   
   'Add By Sindy 2019/12/13
   strExc(0) = "SELECT cp09,cp10 from caseprogress where cp09='" & strReceiveNo & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      strCP10 = RsTemp.Fields("cp10")
   End If
   '2019/12/13 END
   
   strNA81Appl = "" 'Add By Sindy 2019/1/22
   ii = 0
   'Add By Sindy 2018/10/19 有變更申請人時,已傳入的資料讀取資料
   For jj = 1 To 5
      strApplEmp(jj) = ""
   Next jj
   If Trim(strApplNum) <> "" Then
      varTemp = Split(strApplNum, "@")
      For jj = 1 To UBound(varTemp)
         strApplEmp(jj) = Trim(varTemp(jj - 1))
      Next jj
   Else
      For jj = 1 To 5
         If pa(25 + jj) <> "" Then
            strApplEmp(jj) = pa(25 + jj)
         End If
      Next jj
   End If
   '變更代表人
   For jj = 1 To 30
      strRepresentativeEmp(jj) = ""
   Next jj
   If Trim(strRepresentative) <> "" Then
      varTemp = Split(strRepresentative, "@")
      For jj = 1 To UBound(varTemp)
         strRepresentativeEmp(jj) = Trim(varTemp(jj - 1))
      Next jj
   End If
   '2018/10/19 END
   '申請人
   For jj = 1 To 5
      'Modify By Sindy 2018/10/19
      'If pa(25 + jj) <> "" Then
      If strApplEmp(jj) <> "" Then
      '2018/10/19 END
         '申請人
         'Modify By Sindy 2019/4/12 + ,CU07,CU103
         strExc(0) = " SELECT C.*,N1.NA72 X1,N2.NA72 X2,CU07,CU103" & _
            " FROM CUSTOMER C,NATION N1,NATION N2 WHERE CU01='" & Left(ChangeCustomerL(strApplEmp(jj)), 8) & "'" & _
            " and cu02='" & Mid(ChangeCustomerL(strApplEmp(jj)), 9) & "' AND N1.NA01(+)=CU10 AND N2.NA01(+)=CU87"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-國籍','" & RsTemp("X1") & "')"
            
            If RsTemp("CU15") = "0" Then
               strTmp = "自然人"
            Else
               strTmp = "法人公司機關學校"
            End If
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-身分種類','" & strTmp & "')"
            
            If RsTemp("CU10") < "010" Then
               'Add By Sindy 2018/5/2
               If RsTemp("CU15") = "0" And "" & RsTemp("CU11") = "" Then '個人無ID時也要顯示標題
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-ID','♀')"
               Else
               '2018/5/2 END
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-ID','" & RsTemp("CU11") & "')"
               End If
            End If
            
            If RsTemp("CU15") = "0" Then
               strTmp = "申請人" & jj & "-中文姓名"
            Else
               strTmp = "申請人" & jj & "-中文名稱"
            End If
            'Add By Sindy 2017/11/15 修法:106/12/01開始中文名稱要加外商國名
            If Val(strSrvDate(2)) >= 1061201 And RsTemp("CU15") = "1" Then '1.公司
               'Modify By Sindy 2020/11/11 敏莉說申請人名稱她們有可能輸入"後補",不要加外商國名
               strTmp2 = ChgSQL("" & RsTemp("CU04"))
               If strTmp2 <> "後補" Then
                  strTmp2 = GetPrjNationName("" & RsTemp("CU10"), "NA81", pa(1)) & strTmp2
               End If
               '2020/11/11 END
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & strTmp2 & "')"
               'Add By Sindy 2019/1/22
               If GetPrjNationName("" & RsTemp("CU10"), "NA81", pa(1)) <> "" Then
                  strNA81Appl = strNA81Appl & "、" & strTmp2
               End If
               '2019/1/22 END
            Else
            '2017/11/15 END
               'Add By Sindy 2018/4/16 柏翰提個人的姓和名中間要有,號
               If RsTemp("CU15") = "0" Then '自然人
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & PUB_ConvertNameFormat(ChgSQL("" & RsTemp("CU04"))) & "')"
               Else
               '2018/4/16 END
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL("" & RsTemp("CU04")) & "')"
               End If
            End If
            
            'Add By Sindy 2018/12/5 是否顯示英文資料
            'Modify By Sindy 2020/4/10 排除專利處領證及繳年費,年費固定不顯示英文
            If (bolShowEng = True Or RsTemp("CU10") > "010") And _
                Not (Left(Pub_StrUserSt03, 1) = "P" And (strCP10 = 領證及繳年費 Or strCP10 = 年費)) Then
            '2018/12/5 END
               If RsTemp("CU15") = "0" Then
                  strTmp = "申請人" & jj & "-英文姓名"
               Else
                  strTmp = "申請人" & jj & "-英文名稱"
               End If
'               If bolShowEng = True And ChgSQL(RTrim(Trim("" & RsTemp("CU05")) & " " & Trim("" & RsTemp("CU88")) & " " & Trim("" & RsTemp("CU89")) & " " & Trim("" & RsTemp("CU90")))) = "" Then
'                  ii = ii + 1
'                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','♀')"
'               Else
               'Modify By Sindy 2019/6/3 無欄位值帶♀
               If Trim(Trim("" & RsTemp("CU05")) & " " & Trim("" & RsTemp("CU88")) & " " & Trim("" & RsTemp("CU89")) & " " & Trim("" & RsTemp("CU90"))) = "" Then
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','♀')"
               Else
               '2019/6/3 END
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL(RTrim(Trim("" & RsTemp("CU05")) & " " & Trim("" & RsTemp("CU88")) & " " & Trim("" & RsTemp("CU89")) & " " & Trim("" & RsTemp("CU90")))) & "')"
               End If
            End If
            
            '目前抓客戶基本檔資料,等基本檔加欄位後需改抓
            'Modify By Sindy 2019/5/20
            If Left(Pub_StrUserSt03, 1) = "F" Then '外專抓取地址國籍
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-居住國','" & RsTemp("X2") & "')"
            Else
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-居住國','" & RsTemp("X1") & "')"
            End If
            
            'Add By Sindy 2018/4/19 抓個案地址
            If bolReadPA = True Then
               'Add By Sindy 2019/2/19
               If RsTemp("CU10") < "010" And Trim(PUB_ChgNumeralStyle(ChgSQL("" & RsTemp("CU23")))) = Trim(PUB_ChgNumeralStyle(ChgSQL(pa(30 + jj)))) Then
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-郵遞區號','" & PUB_ChgNumeralStyle("" & RsTemp("CU112")) & "')"
               'Add By Sindy 2019/8/14
               ElseIf RsTemp("CU10") < "010" Then
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-郵遞區號','♀')"
               '2019/8/14 END
               End If
               '2019/2/19 END
               
               'Add By Sindy 2019/12/13 '年費申請人是否出名為"N"時,顯示內容不一樣
               If (strCP10 = 年費 And pa(143) = "N") Then
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-中文地址','" & RsTemp("X2") & "')"
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-英文地址','" & RsTemp("X2") & "')"
               Else
               '2019/12/13 END
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-中文地址','" & PUB_ChgNumeralStyle(ChgSQL(pa(30 + jj))) & "')"
                  'Add By Sindy 2018/12/5 是否顯示英文資料
                  'Modify By Sindy 2020/4/10 排除專利處領證及繳年費,年費固定不顯示英文
                  If (bolShowEng = True Or RsTemp("CU10") > "010") And _
                      Not (Left(Pub_StrUserSt03, 1) = "P" And (strCP10 = 領證及繳年費 Or strCP10 = 年費)) Then
                  '2018/12/5 END
   '                  If bolShowEng = True And ChgSQL(pa(35 + jj)) = "" Then
   '                     ii = ii + 1
   '                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-英文地址','♀')"
   '                  Else
                     'Modify By Sindy 2019/6/3 無欄位值帶♀
                     If Trim(ChgSQL(pa(35 + jj))) = "" Then
                        ii = ii + 1
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-英文地址','♀')"
                     Else
                     '2019/6/3 END
                        ii = ii + 1
                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-英文地址','" & ChgSQL(pa(35 + jj)) & "')"
                     End If
                  End If
               End If
            Else
            '2018/4/19 END
               If RsTemp("CU10") < "010" Then
                  ii = ii + 1
                  'Modify By Sindy 2020/4/10 ex:P-124163
                  'PUB_ChgNumeralStyle("" & RsTemp("CU112")) => IIf("" & RsTemp("CU112") = "", "♀", PUB_ChgNumeralStyle("" & RsTemp("CU112")))
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-郵遞區號','" & IIf("" & RsTemp("CU112") = "", "♀", PUB_ChgNumeralStyle("" & RsTemp("CU112"))) & "')"
               End If
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-中文地址','" & PUB_ChgNumeralStyle(ChgSQL("" & RsTemp("CU23"))) & "')"
               'Add By Sindy 2018/12/5 是否顯示英文資料
               'Modify By Sindy 2020/4/10 排除專利處領證及繳年費,年費固定不顯示英文
               If (bolShowEng = True Or RsTemp("CU10") > "010") And _
                  Not (Left(Pub_StrUserSt03, 1) = "P" And (strCP10 = 領證及繳年費 Or strCP10 = 年費)) Then
               '2018/12/5 END
'                  If bolShowEng = True And ChgSQL(RTrim(Trim("" & RsTemp("CU24")) & " " & Trim("" & RsTemp("CU25")) & " " & Trim("" & RsTemp("CU26")) & " " & Trim("" & RsTemp("CU27")) & " " & Trim("" & RsTemp("CU28")))) = "" Then
'                     ii = ii + 1
'                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-英文地址','♀')"
'                  Else
                  'Modify By Sindy 2019/6/3 無欄位值帶♀
                  If Trim(Trim("" & RsTemp("CU24")) & " " & Trim("" & RsTemp("CU25")) & " " & Trim("" & RsTemp("CU26")) & " " & Trim("" & RsTemp("CU27")) & " " & Trim("" & RsTemp("CU28")) & " " & Trim("" & RsTemp("CU102"))) = "" Then
                     ii = ii + 1
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-英文地址','♀')"
                  Else
                  '2019/6/3 END
                     ii = ii + 1
'                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-英文地址','" & ChgSQL(RTrim(Trim("" & RsTemp("CU24")) & " " & Trim("" & RsTemp("CU25")) & " " & Trim("" & RsTemp("CU26")) & " " & Trim("" & RsTemp("CU27")) & " " & Trim("" & RsTemp("CU28")))) & "')"
                     'Modified by Morgan 2019/9/9 補CU102 Ex: FCP060229
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-英文地址','" & ChgSQL(RTrim(Trim("" & RsTemp("CU24")) & " " & Trim("" & RsTemp("CU25")) & " " & Trim("" & RsTemp("CU26")) & " " & Trim("" & RsTemp("CU27")) & " " & Trim("" & RsTemp("CU28")) & " " & Trim("" & RsTemp("CU102")))) & "')"

                  End If
               End If
            End If
            
            'Modify By Sindy 2018/1/16 依申請人帶出代表人資料
            'Add By Sindy 2018/4/19 抓個案地址
            If "" & RsTemp("CU15") <> "0" Then '非自然人才要帶出代表人資料
               'Add By Sindy 2019/12/13 '年費申請人是否出名為"N"時,顯示內容不一樣
               If (strCP10 = 年費 And pa(143) = "N") Then
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-代表人中文姓名','與卷存一致')"
                  '不顯示申請人1-代表人英文姓名
               Else
               '2019/12/13 END
               
                  strChaName = "": strEngName = ""
                  'Modify By Sindy 2018/10/19 變更代表人(1~30)
                  If Trim(strRepresentative) <> "" Then
                     intRow = 0
                     For kk = 1 To 2
                        intRow = intRow + 1
                        If kk = 1 Then idx = jj + (jj - 1) * 5 '1,7,13,19,25
                        If kk = 2 Then idx = idx + 3 '4,10,16,22,28
                        '代表人中文姓名-->非自然人時為必要欄位
                        strTmp = strRepresentativeEmp(idx)
                        If strTmp <> "" Then
                           If Len(strTmp) = 3 Then strTmp = PUB_ConvertNameFormat(strTmp)
                           strChaName = strChaName & " " & intRow & "." & strTmp
                        Else
                           '只有一個代表人時不要有1.
                           If strChaName <> "" Then
                              strChaName = Replace(strChaName, "1.", "")
                           End If
                        End If
                        '代表人英文姓名-->非必要欄位
                        strTmp = strRepresentativeEmp(idx + 1)
                        If strTmp <> "" Then
                           strEngName = strEngName & " " & intRow & "." & strTmp
                        Else
                           '只有一個代表人時不要有1.
                           If strEngName <> "" Then
                              strEngName = Replace(strEngName, "1.", "")
                           End If
                        End If
                     Next kk
                     '2018/10/19 END
                     
                  'Modify By Sindy 2020/11/5 和敏莉確認,代表人只有二種狀況,一種是抓變更畫面上的,另一種是抓個案代表人
                  ElseIf bolReadPA = True Or Left(Pub_StrUserSt03, 1) = "F" Then '抓個案資料
                     If jj < 3 Then
                        k_star = 1: k_end = 2
                     ElseIf jj = 3 Then
                        k_star = 3: k_end = 4
                     ElseIf jj = 4 Then
                        k_star = 5: k_end = 6
                     ElseIf jj = 5 Then
                        k_star = 7: k_end = 8
                     End If
                     intRow = 0
                     For kk = k_star To k_end
                        intRow = intRow + 1
                        '代表人中文姓名-->非自然人時為必要欄位
                        If jj = 1 Then
                           strTmp = pa(79 + 3 * (kk - 1))
                        Else
                           strTmp = pa(109 + 3 * (kk - 1))
                        End If
                        If strTmp <> "" Then
                           strChaName = strChaName & " " & intRow & "." & strTmp
                        Else
                           'Modify By Sindy 2018/1/17 只有一個代表人時不要有1.
                           If strChaName <> "" Then
                              strChaName = Replace(strChaName, "1.", "")
                           End If
                           '2018/1/17 END
                        End If
                        '代表人英文姓名-->非必要欄位
                        If jj = 1 Then
                           strTmp = pa(80 + 3 * (kk - 1))
                        Else
                           strTmp = pa(110 + 3 * (kk - 1))
                        End If
                        If strTmp <> "" Then
                           strEngName = strEngName & " " & intRow & "." & strTmp
                        Else
                           'Modify By Sindy 2018/1/17 只有一個代表人時不要有1.
                           If strEngName <> "" Then
                              strEngName = Replace(strEngName, "1.", "")
                           End If
                           '2018/1/17 END
                        End If
                     Next kk
                     
                  Else '抓客戶檔資料
                     If Left(Pub_StrUserSt03, 1) = "F" Then '外專要帶後補2個字
                        intRow = 0
                        For kk = 1 To 6
                           intRow = intRow + 1
                           '代表人中文姓名-->非自然人時為必要欄位
                           strTmp = "" & RsTemp("CU" & CStr(39 + 3 * (kk - 1)))
                           If strTmp <> "" Then
                              If Len(strTmp) = 3 Then strTmp = PUB_ConvertNameFormat(strTmp)
                              strChaName = strChaName & " " & intRow & "." & strTmp
                           Else
                              'Modify By Sindy 2018/1/17 只有一個代表人時不要有1.
                              If strChaName <> "" Then
                                 strChaName = Replace(strChaName, "1.", "")
                              End If
                              '2018/1/17 END
                           End If
                           '代表人英文姓名-->非必要欄位
                           strTmp = "" & RsTemp("CU" & CStr(40 + 3 * (kk - 1)))
                           If strTmp <> "" Then
                              strEngName = strEngName & " " & intRow & "." & strTmp
                           Else
                              'Modify By Sindy 2018/1/17 只有一個代表人時不要有1.
                              If strEngName <> "" Then
                                 strEngName = Replace(strEngName, "1.", "")
                              End If
                              '2018/1/17 END
                           End If
                        Next kk
                     Else
                        'Add By Sindy 2019/4/12
                        '公司負責人
                        If "" & RsTemp.Fields("cu07") <> "" Then
                           strChaName = Trim("" & RsTemp.Fields("cu07"))
                           If Len(strChaName) = 3 Then strChaName = PUB_ConvertNameFormat(strChaName)
                        End If
                        '公司英文負責人
                        strEngName = Trim("" & RsTemp.Fields("cu103"))
                        '2019/4/12 END
                     End If
                  End If
   '            If RsTemp("CU15") <> "0" Then
   '               If jj < 3 Then
   '                  strTmp = pa(79 + 3 * (jj - 1))
   '               Else
   '                  strTmp = pa(109 + 3 * (jj - 1))
   '               End If
   '               If strTmp = "" Then
   '                  strTmp = "後補"
   '               End If
   '               '代表人中文姓名-->非自然人時為必要欄位
   '               ii = ii + 1
   '               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-代表人中文姓名','" & ChgSQL(strTmp) & "')"
   '               '代表人英文姓名-->非必要欄位
   '               If jj < 3 Then
   '                  strTmp = pa(80 + 3 * (jj - 1))
   '               Else
   '                  strTmp = pa(110 + 3 * (jj - 1))
   '               End If
   '               If strTmp <> "" Then
   '                  ii = ii + 1
   '                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-代表人英文姓名','" & ChgSQL(strTmp) & "')"
   '               End If
   '            End If
                  '代表人中文姓名
                  If Trim(strChaName) = "" Then
                     If Left(Pub_StrUserSt03, 1) = "F" Then '外專要帶後補2個字
                        strChaName = "後補"
                     Else
                        strChaName = "♀"
                     End If
                  Else
                     strChaName = Trim(strChaName)
                  End If
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-代表人中文姓名','" & ChgSQL(strChaName) & "')"
                  'Add By Sindy 2018/12/5 是否顯示英文資料
   '               If bolShowEng = True Then
                  '2018/12/5 END
                     '代表人英文姓名
                     'Add By Sindy 2019/8/26
                     'Modify By Sindy 2019/8/29 敏莉說中文有後補,英文無資料才要帶後補
                     If Trim(strEngName) = "" And Left(Pub_StrUserSt03, 1) = "F" And strChaName = "後補" Then '外專要帶後補2個字
                        strEngName = "後補"
                     End If
                     '2019/8/26 END
                     'Modify By Sindy 2019/6/3 是否顯示英文資料
                     'Modify By Sindy 2020/4/10 排除專利處領證及繳年費,年費固定不顯示英文
                     If (strEngName <> "" Or bolShowEng = True) And _
                        Not (Left(Pub_StrUserSt03, 1) = "P" And (strCP10 = 領證及繳年費 Or strCP10 = 年費)) Then
                        'Modify By Sindy 2019/6/3 無欄位值帶♀
                        If Trim(strEngName) = "" Then
                           strEngName = "♀"
                        Else
                        '2019/6/3 END
                           strEngName = Trim(strEngName)
                        End If
   '                     If bolShowEng = True And ChgSQL(strEngName) = "" Then
   '                        ii = ii + 1
   '                        strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '                           " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-代表人英文姓名','♀')"
   '                     Else
                           ii = ii + 1
                           strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                              " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-代表人英文姓名','" & ChgSQL(strEngName) & "')"
   '                     End If
                     End If
   '               End If
               End If
            End If
            '2018/1/16 END
         End If
      End If
   Next jj
   
   'Add By Sindy 2019/1/22
   If strNA81Appl <> "" Then
      strNA81Appl = Mid(strNA81Appl, 2)
   End If
   '2019/1/22 END
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      PUB_GetApplPA_EData = True
   End If
End Function

'Add By Sindy 2017/12/11 智慧局機關來函清單加註
'Modify By Sindy 2018/7/17 改傳strCaseNo本所案號
Public Function PUB_ReadIPOListMemo(ByVal strCaseNo As String, ByVal strYid As String, ByVal strXid As String, _
   ByVal strApproveMemo As String, ByVal strlimitDate As String) As String
Dim rsA As New ADODB.Recordset
Dim strCon As String
Dim bolChkilm04 As Boolean
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
   
   PUB_ReadIPOListMemo = ""
   
   strXid = Left(strXid & "00000000", 8)
   strYid = Left(strYid & "00000000", 8)
   'Add By Sindy 2018/7/17 解析本所案號
   'Add By Sindy 2019/6/11
   Dim tmpArr
   tmpArr = Split(strCaseNo, "-")
   '2019/6/11 END
   If InStr(strCaseNo, "-") > 0 Then
      'Add By Sindy 2019/6/11
      If UBound(tmpArr) = 1 Then
      '2019/6/11 END
         strCaseNo = strCaseNo & "-0-00"
      End If
      strCP01 = SystemNumber(strCaseNo, 1)
      strCP02 = SystemNumber(strCaseNo, 2)
      strCP03 = SystemNumber(strCaseNo, 3)
      strCP04 = SystemNumber(strCaseNo, 4)
      strExc(0) = "select pa26,pa27,pa28,pa29,pa30" & _
                  " from patent where pa01='" & strCP01 & "' and pa02='" & strCP02 & "'" & _
                  " and pa03='" & strCP03 & "' and pa04='" & strCP04 & "'"
      intI = 1
      Set rsA = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Not IsNull(rsA.Fields("pa26")) Then
            strXid = Left(rsA.Fields("pa26"), 8)
         End If
         If Not IsNull(rsA.Fields("pa27")) Then
            strXid = strXid & ";" & Left(rsA.Fields("pa27"), 8)
         End If
         If Not IsNull(rsA.Fields("pa28")) Then
            strXid = strXid & ";" & Left(rsA.Fields("pa28"), 8)
         End If
         If Not IsNull(rsA.Fields("pa29")) Then
            strXid = strXid & ";" & Left(rsA.Fields("pa29"), 8)
         End If
         If Not IsNull(rsA.Fields("pa30")) Then
            strXid = strXid & ";" & Left(rsA.Fields("pa30"), 8)
         End If
      End If
   Else
      strCP01 = strCaseNo
   End If
   '2018/7/17 END
   If strlimitDate <> "" Then
      strlimitDate = Replace(strlimitDate, "/", "")
   End If
   
   strCon = ""
   If strCP01 <> "" Then
      strCon = strCon & " and instr(','||ilm03,'," & strCP01 & "')>0"
   End If
   'Modify By Sindy 2018/7/17 and ilm02='" & strXid & "' ==> and instr('" & strXid & "',ilm02)>0
   'Modify By Sindy 2022/12/27 + and ilm02<>'0'
   strExc(0) = "select ilm01,ilm02,ilm03,ilm04,ilm05,ilm06,1 sort" & _
               " from IPOListMemo where ilm01='" & strYid & "' and instr('" & strXid & "',ilm02)>0 and ilm02<>'0'" & strCon & _
               " union select ilm01,ilm02,ilm03,ilm04,ilm05,ilm06,2 sort" & _
               " from IPOListMemo where ilm01='" & strYid & "' and ilm01<>'0'" & strCon & _
               " union select ilm01,ilm02,ilm03,ilm04,ilm05,ilm06,3 sort" & _
               " from IPOListMemo where instr('" & strXid & "',ilm02)>0 and ilm02<>'0'" & strCon & _
               " order by sort asc"
   intI = 1
   Set rsA = ClsLawReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      rsA.MoveFirst
      Do While Not rsA.EOF
         bolChkilm04 = False
         If rsA.Fields("sort") <> "1" Then
            If rsA.Fields("sort") = "2" Then '比對到代理人
               '檢查申請人
               If rsA.Fields("ilm02") <> "0" Then
                  'Modify By Sindy 2018/7/17
                  'If strXid <> rsA.Fields("ilm02") Then
                  If InStr(strXid, rsA.Fields("ilm02")) = 0 Then
                  '2018/7/17 END
                     GoTo ReadNext
                  End If
               End If
            ElseIf rsA.Fields("sort") = "3" Then '比對到申請人
               '檢查代理人
               If rsA.Fields("ilm01") <> "0" Then
                  If strYid <> rsA.Fields("ilm01") Then
                     GoTo ReadNext
                  End If
               End If
            End If
         End If
         '是否含核准:N.不含核准 / Null.不分
         If "" & rsA.Fields("ilm04") = "N" Then
            'Modify By Sindy 2017/12/22
            'If strApproveMemo <> "核准" Then
            If InStr(strApproveMemo, "核准") = 0 Then 'N.不含
            '2017/12/22 END
               bolChkilm04 = True
            End If
         ElseIf "" & rsA.Fields("ilm04") = "Y" Then
            'Modify By Sindy 2017/12/22
            'If strApproveMemo = "核准" Then
            If InStr(strApproveMemo, "核准") > 0 Then 'Y.核准
            '2017/12/22 END
               bolChkilm04 = True
            End If
         '不考慮此條件
         Else
            bolChkilm04 = True
         End If
         If bolChkilm04 = True Then
            '有無期限:Y.有期限
            If "" & rsA.Fields("ilm05") = "Y" Then
               If Val(strlimitDate) > 0 Then '有期限
                  PUB_ReadIPOListMemo = rsA.Fields("ilm06")
                  Exit Do
               End If
            '不考慮此條件
            Else
               PUB_ReadIPOListMemo = rsA.Fields("ilm06")
               Exit Do
            End If
         End If
ReadNext:
         rsA.MoveNext
      Loop
   End If
   rsA.Close
   Set rsA = Nothing
End Function

'Add by Morgan 2005/7/13
'設定出名代理人清單
'2010/5/7 MODIFY BY SONIA 加傳CP10以判斷是否為新案案件性質
'Modify By Sindy 2018/4/12 ByVal p_CP110 As String ==> ByRef p_CP110 As String
'Modified by Lydia 2018/11/15 改成Form 2.0可用
'Public Sub PUB_SetOurAgent(ByRef p_ListBox As ListBox, ByRef p_CP() As String, Optional ByRef p_CP110 As String, Optional ByVal p_CP10 As String)
Public Sub PUB_SetOurAgent(ByRef p_Listbox As Control, ByRef p_CP() As String, Optional ByRef p_CP110 As String, Optional ByVal p_CP10 As String, Optional ByVal bForm2 As Boolean = False)
   Dim iNum As Integer '已勾選項目數
   Dim tmpOA As String 'Added by Lydia 2018/11/15
   Dim pStart As String 'Added by Lydia 2019/08/02  第一個列出的出名代理人
   Dim intQ As Integer 'Added by Lydia 2019/08/02
   
On Error GoTo flgErr
   
   With adoRecordset
      '若未設定時抓最近發文的AB類收文
      If p_CP110 = "" Then
         '2010/5/7 ADD BY SONIA 新案預設OA04='Y'者
         'Modified by Lydia 2023/08/08 預設CFT申請英文證明書
         'If p_CP10 <> "" And ((p_CP(1) = "FCP" Or p_CP(1) = "P") And (InStr(NewCasePtyList & ",803", p_CP10) > 0)) Or ((p_CP(1) = "FCT" Or p_CP(1) = "T") And p_CP10 = "101") Then
         If p_CP10 <> "" And (((p_CP(1) = "FCP" Or p_CP(1) = "P") And InStr(NewCasePtyList & ",803", p_CP10) > 0) Or ((p_CP(1) = "FCT" Or p_CP(1) = "T") And p_CP10 = "101") Or (p_CP(1) = "CFT" And p_CP10 = "304")) Then
'            '2011/11/30 add by sonia 電子電機組案件預設林特助
'            If p_CP(1) = "FCP" And p_CP(150) = "1" Then
'               p_CP110 = "94007"
'            Else
'            '2011/11/30 end
               strSql = "select OA02,OA03 from ouragent where oa01='" & p_CP(1) & "' and OA04='Y' order by OA03 DESC, OA01 desc"
               CheckOC
               .CursorLocation = adUseClient
               .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If .RecordCount > 0 Then
                  Do While Not .EOF
                     p_CP110 = p_CP110 & .Fields(0) & ","
                     .MoveNext
                  Loop
               End If
'            End If

'Removed by Morgan 2012/6/15 101.6.1 改與其他組相同預設，不必預設
'            '2011/12/2 Modify by Sindy 電子電機組案件預設林特助 - 程式段往下移至此,因內商Run到(p_CP(150)程式段)陣列索引超出範圍
'            If p_CP(1) = "FCP" Then
'               If p_CP(150) = "1" Then
'                  p_CP110 = "94007"
'               End If
'            End If
'            '2011/12/2 End
'end 2012/6/15

         'FCP 下列性質不預設
         ElseIf p_CP(1) = "FCP" And ((p_CP10 = "901" Or p_CP10 = "902" Or p_CP10 = "903" Or p_CP10 = "904" Or p_CP10 = "906" Or p_CP10 = "912")) Then
         '2011/5/10 ADD BY SONIA FCT延展(102)，變更(301)，移轉(501)，授權(502)，再授權(504)，設定質權(506)設定出名代理人為"閻啟泰，林景郁"
         'Modified by Lydia 2019/08/02 嘉雯表示: 因為電子送件大部份都是林+閻, 所以改成預設
         'ElseIf p_CP(1) = "FCT" And ((p_CP10 = "102" Or p_CP10 = "301" Or p_CP10 = "501" Or p_CP10 = "502" Or p_CP10 = "504" Or p_CP10 = "506")) Then
         'Modified by Lydia 2019/08/15 桂英: 增加下列案件性質 202申請意見書,601異議,602異議答辯,603評定,604評定答辯,605廢止,606廢止答辯,623部分廢止,624部分廢止答辯,627部分異議,628部分異議答辯,629部分評定,630部分評定答辯
         'ElseIf (p_CP(1) = "FCT" Or p_CP(1) = "T") And ((p_CP10 = "102" Or p_CP10 = "301" Or p_CP10 = "501" Or p_CP10 = "502" Or p_CP10 = "504" Or p_CP10 = "506")) Then
         ElseIf (p_CP(1) = "FCT" Or p_CP(1) = "T") And InStr("102,301,501,502,504,506,202,601,602,603,604,605,606,623,624,627,628,629,630", p_CP10) > 0 Then
            'modify by sonia 2016/6/4 改二人順序
            p_CP110 = "94007,81040"
         '2011/5/10 END
         Else
         '2010/5/7 END
            'Modify by Morgan 2005/8/26 改抓有出名代理人的進度
            'strSQL = "select cp110 from caseprogress where cp01='" & p_CP(1) & "' and cp02='" & p_CP(2) & "' and cp03='" & p_CP(3) & "' and cp04='" & p_CP(4) & "' and cp27 is not null and cp09<'C' order by cp27 desc"
            'edit by nickc 2006/02/24 阿蓮說 FCT 2006/02/07 以後發文的才抓預設的
            'strSQL = "select cp110 from caseprogress where cp01='" & p_CP(1) & "' and cp02='" & p_CP(2) & "' and cp03='" & p_CP(3) & "' and cp04='" & p_CP(4) & "' and cp27 is not null and cp09<'C' and cp110 is not null order by cp27 desc"
            'edit by nickc 2008/01/24 阿蓮說 FCT 預設出名代理人時，不要抓回代的
            'strSQL = "select cp110 from caseprogress where cp01='" & p_CP(1) & "' and cp02='" & p_CP(2) & "' and cp03='" & p_CP(3) & "' and cp04='" & p_CP(4) & "' and cp27 is not null " & IIf(p_CP(1) = "FCT", " and cp27>=20060207 ", "") & " and cp09<'C' and cp110 is not null order by cp27 desc"
            'Modified by Morgan 2020/3/20 +cp10
            strSql = "select cp110,cp10 from caseprogress where cp01='" & p_CP(1) & "' and cp02='" & p_CP(2) & "' and cp03='" & p_CP(3) & "' and cp04='" & p_CP(4) & "' and cp27 is not null " & IIf(p_CP(1) = "FCT", " and cp27>=20060207 ", "") & " and cp09<'C' and cp01||cp10 not in ('FCT720') and cp110 is not null order by cp27 desc"
            CheckOC
            .CursorLocation = adUseClient
            .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
            If .RecordCount > 0 Then
               p_CP110 = "" & .Fields(0)
               'Added by Morgan 2020/3/20
               '專利非年費的發文代理人優先
               If (p_CP(1) = "P" Or p_CP(1) = "FCP") And .Fields("cp10") = "605" Then
                  .MoveNext
                  Do While Not .EOF
                     If .Fields("cp10") <> "605" Then
                        p_CP110 = "" & .Fields(0)
                        Exit Do
                     End If
                     .MoveNext
                  Loop
               End If
               'end 2020/3/20
            End If
            
            'Added by Morgan 2020/3/20
            '專利年費取消桂所長出名
            If (p_CP(1) = "P" Or p_CP(1) = "FCP") And p_CP10 = "605" Then
               If InStr(p_CP110, "76012") > 0 Then
                  p_CP110 = Replace(p_CP110, "76012", "")
                  If Left(p_CP110, 1) = "," Then p_CP110 = Mid(p_CP110, 2)
                  If Right(p_CP110, 1) = "," Then p_CP110 = Mid(p_CP110, 1, Len(p_CP110) - 1)
               End If
            End If
            'end 2020/3/20
            
         End If
      End If
      'Add by Morgan 2005/12/26 加排序欄位
      'strSQL = "select st01,st02 from ouragent,staff where oa01='" & p_CP(1) & "' and st01=oa02 order by 1 desc"
      'Modify by Morgan 2008/5/28 只抓在職的
      'Modify By Sindy 2018/8/1 +
      If TypeName(p_Listbox) <> "Nothing" Then
      '2018/8/1 END
         strSql = "select st01,st02,OA03 from ouragent,staff where oa01='" & p_CP(1) & "' and st01=oa02 and st04='1'  order by 3 DESC, 1 desc"
         CheckOC
         .CursorLocation = adUseClient
         .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
         iNum = 0
         If .RecordCount > 0 Then
            p_Listbox.Clear
            If bForm2 = True Then p_Listbox.Tag = "" 'Added by Lydia 2018/11/15 原本放在ItemData,改放在Tag
           
            Do While Not .EOF
               'Modify by Morgan 2006/2/24 有勾選的排前面
               If InStr(p_CP110, "" & .Fields(0)) > 0 Then
                  If pStart = "" Then pStart = "" & .Fields(0) 'Added by Lydia 2019/08/02 第一個列出的出名代理人
                  p_Listbox.AddItem "" & .Fields(1), 0
                  'modify by sonia 2016/10/5 員工編號已可非數字需做轉換
                  'p_ListBox.ItemData(0) = .Fields(0) '員工編號
                  'Added by Lydia 2018/11/15 原本放在ItemData,改放在Tag
                  If bForm2 = True Then
                      tmpOA = .Fields(0) & IIf(tmpOA <> "", ",", "") & tmpOA
                  Else
                  'end 2018/11/15
                       p_Listbox.ItemData(0) = PUB_Id2Num(.Fields(0)) '員工編號
                  End If 'end 2018/11/15
                  p_Listbox.Selected(0) = True
                  iNum = iNum + 1
               Else
                  p_Listbox.AddItem "" & .Fields(1), iNum
                  'modify by sonia 2016/10/5 員工編號已可非數字需做轉換
                  'p_ListBox.ItemData(iNum) = .Fields(0) '員工編號
                  'Added by Lydia 2018/11/15 原本放在ItemData,改放在Tag
                  If bForm2 = True Then
                      'Modified by Lydia 2018/12/13
                      'tmpOA = tmpOA & IIf(tmpOA <> "", ",", "") & .Fields(0)
                      'Added by Lydia 2019/08/02 排在第一個列出的出名代理人的後面
                      If pStart <> "" Then
                          intQ = InStr(tmpOA, pStart & ",")
                          If intQ = 0 Then
                              tmpOA = tmpOA & "," & .Fields(0)
                          Else
                              tmpOA = Mid(tmpOA, 1, intQ + Len(pStart)) & .Fields(0) & "," & Mid(tmpOA, intQ + Len(pStart & ","))
                          End If
                      Else
                      'end 2019/08/02
                          tmpOA = .Fields(0) & IIf(tmpOA <> "", ",", "") & tmpOA
                      End If  'end 2019/08/02
                  Else
                  'end 2018/11/15
                      p_Listbox.ItemData(iNum) = PUB_Id2Num(.Fields(0)) '員工編號
                  End If 'end 2018/11/15
               End If
               .MoveNext
            Loop
         End If
      End If
      'Added by Lydia 2018/11/15
      If tmpOA <> "" Then
          p_Listbox.Tag = tmpOA
          p_Listbox.ListIndex = 0
      End If
      'end 2018/11/15
   End With
   
flgErr:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Sub

'Add By Morgan 2018/5/21
Public Function ReplaceMadeWord(pText As String, Optional pReplace As String = "　") As String
   Dim ii As Integer, strBig5 As String
   Dim strNew As String, strChar As String

   strNew = ""
   For ii = 1 To Len(pText)
      strChar = Mid(pText, ii, 1)
      strBig5 = Hex(Asc(strChar))
      '造字區1:ＦＡ４０－ＦＥＦＥ
      If strBig5 >= "FA40" And strBig5 <= "FEFE" Then
         strChar = pReplace
      '造字區2:８Ｅ４０－Ａ０ＦＥ
      ElseIf strBig5 >= "8E40" And strBig5 <= "A0FE" Then
         strChar = pReplace
      '造字區3:８１４０－８ＤＦＥ
      ElseIf strBig5 >= "8140" And strBig5 <= "8DFE" Then
         strChar = pReplace
      '造字區4:Ｃ６Ａ１－Ｃ８ＦＥ
      ElseIf strBig5 >= "C6A1" And strBig5 <= "C8FE" Then
         strChar = pReplace
      End If
      strNew = strNew & strChar
   Next
   ReplaceMadeWord = strNew
End Function

'Modify By Sindy 2018/11/19 從工程師系統複製過來,改為程序操作
'Add By Sindy 2018/4/10
'申請書
Public Sub Pub_P_NewCaseStartLetter2(ByVal ET01 As String, ByVal ET03 As String, ByVal strCP09 As String, _
   pa() As String, cp() As String, bolCM10 As Boolean, Optional bolHadPOA As Boolean = False, _
   Optional ByRef m_bolShowEng As Boolean = False, Optional bolHadPOAeFile As Boolean = False)
   
   Dim strTxt(110) As String, strTmp As String
   Dim ii As Integer, jj As Integer
   Dim strInventor As String
   Dim strTemp As String
   Dim strCaseNo As String
   
   ii = 0
   EndLetter ET01, strCP09, ET03, strUserNum

   ii = ii + 1
   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
      " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','本所案號','" & pa(1) & Val(pa(2)) & IIf(pa(3) <> "0" Or pa(4) <> "00", "-" & pa(3), IIf(pa(4) <> "00", "-" & pa(4), "")) & "')"

   If pa(8) = "3" Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','設計種類','" & PUB_GetCaseAttributeName(pa(158), pa(8)) & "')"
   End If
   
   '123.主張優惠期
   If PUB_ChkCPExist(cp, "123") Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','有主張優惠期','♀')"
   End If
   '106.主張國際優先權
   '121.主張國內優先權
   If PUB_ChkCPExist(cp, "106") Or PUB_ChkCPExist(cp, "121") Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','有主張優先權','♀')"
      'Add By Sindy 2019/5/17
      'POA資料夾中若有相同案號的PRI檔案(權先權證明文件)一併帶入資料夾
      If PUB_ChkCPExist(cp, "106") Then '106.主張國際優先權,才有電子檔
         strCaseNo = Trim(pa(1)) & Val(Trim(pa(2))) & _
                    IIf(Val(Trim(pa(3))) = 0 And Val(Trim(pa(4))) = 0, "", "-" & pa(3)) & _
                    IIf(Val(Trim(pa(4))) = 0, "", "-" & Format(pa(4), "00"))
         'Modify By Sindy 2022/10/25 改用常變數 str_P_OrderPath
         'If Dir("\\pat1\Order_SCAN\POA\" & strCaseNo & ".PRI.pdf") = "" Then
         If Dir(str_P_OrderPath & "\POA\" & strCaseNo & ".PRI.pdf") = "" Then
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','尚缺優先權文件','♀')"
         End If
      End If
      '2019/5/17 END
   End If
   '946.摘要英譯
   If PUB_ChkCPExist(cp, "946") Then
      ii = ii + 1
      m_bolShowEng = True
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','有摘要英譯','♀')"
   End If
   If pa(26) <> "" Then
      If GetPrjNationNumber1(pa(26)) > "010" Then m_bolShowEng = True
   End If
   If pa(27) <> "" Then
      If GetPrjNationNumber1(pa(27)) > "010" Then m_bolShowEng = True
   End If
   If pa(28) <> "" Then
      If GetPrjNationNumber1(pa(28)) > "010" Then m_bolShowEng = True
   End If
   If pa(29) <> "" Then
      If GetPrjNationNumber1(pa(29)) > "010" Then m_bolShowEng = True
   End If
   If pa(30) <> "" Then
      If GetPrjNationNumber1(pa(30)) > "010" Then m_bolShowEng = True
   End If
   '顯示英文
   If m_bolShowEng = True Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','顯示英文','♀')"
   End If
   '一案兩請
   If bolCM10 = True And _
      (cp(10) = "101" Or cp(10) = "102") Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','有一案兩請','♀')"
   End If
   '有委任書
'   If bolHadPOA = True Then
'      ii = ii + 1
'      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'         " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','有委任書','♀')"
      'Add By Sindy 2019/5/17
      If bolHadPOAeFile = False Then
         ii = ii + 1
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','尚缺委任書','♀')"
      End If
      '2019/5/17 END
'   End If
   
   '申請人
   'Modify By Sindy 2018/12/5 + m_bolShowEng
   Call PUB_GetApplPA_EData(ET01, ET03, strCP09, pa(), False, , , m_bolShowEng)
   
   '預設出名代理人
   Dim lstNameAgent As ListBox
   If cp(110) = "" Then PUB_SetOurAgent lstNameAgent, pa(), cp(110), cp(10)
   '出名代理人
'   strExc(0) = "select oa05,ST26,st02 from ouragent,staff where oa01='" & pa(1) & "' and instr('" & cp(110) & "',oa02)>0 and st01(+)=oa02 order by OA03"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   If intI = 1 Then
'      With RsTemp
'      jj = 1
'      Do While Not .EOF
'         ii = ii + 1
'         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'            " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & .Fields("st02")) & "')"
'         jj = jj + 1
'         .MoveNext
'      Loop
'      End With
'   End If
   'Modify By Sindy 2020/4/8 申請書:出名代理人
   Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 1, cp(110), ET01, strCP09, ET03, ii, strTxt())
   
'   '讀取發明人資料
'   If pa(8) = "1" Then
'      strExc(1) = "發明人"
'   ElseIf pa(8) = "2" Then
'      strExc(1) = "新型創作人"
'   Else
'      strExc(1) = "設計人"
'   End If
'   strInventor = ""
'   strExc(0) = " SELECT IN03,IN04,IN05,IN11,NA72" & _
'               " FROM PatentInventor,INVENTOR,NATION" & _
'               " WHERE pi01=" + CNULL(pa(1)) + " and pi02=" + CNULL(pa(2)) + " and pi03=" + CNULL(pa(3)) + " and pi04=" + CNULL(pa(4)) & _
'               " AND IN01=substr(pi06,1,8) AND IN02=substr(pi06,9,2)" & _
'               " AND NA01(+)=IN11" & _
'               " order by pi05 asc"
'   intI = 1
'   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'   '發明人TAG後面加序號,取消內縮
'   If intI = 1 Then
'      RsTemp.MoveFirst
'      Do While Not RsTemp.EOF
'         'Modify By Sindy 2018/4/16 Mark:不要空行
'         'If strInventor <> "" Then strInventor = strInventor & vbCrLf
'         'Modify By Sindy 2018/10/25 增加英文名稱格式化 PUB_FCPIN05Format_EName
'         'Modify By Sindy 2019/5/17 中文名稱3個字的也要加逗號
'         If strInventor <> "" Then strInventor = strInventor & vbCrLf & vbCrLf 'Add By Sindy 2019/5/30
'         strInventor = strInventor & "【" & strExc(1) & intI & "】" & _
'                                     vbCrLf & "　　【國籍】　　　　　　　　　" & RsTemp("NA72") & _
'                                     vbCrLf & "　　【中文姓名】　　　　　　　" & IIf("" & RsTemp("IN11") = "000" Or Len(ChgSQL("" & RsTemp("IN04"))) = 3, PUB_ConvertNameFormat(ChgSQL("" & RsTemp("IN04"))), ChgSQL("" & RsTemp("IN04"))) & _
'                                     IIf("" & RsTemp("IN11") = "000" And m_bolShowEng = False, "", vbCrLf & "　　【英文姓名】　　　　　　　" & ChgSQL(PUB_FCPIN05Format_EName("" & RsTemp("IN05"), "" & RsTemp("NA72"))))
'         RsTemp.MoveNext
'         intI = intI + 1
'      Loop
'   Else
'      strInventor = "【" & strExc(1) & "1】" & _
'                    vbCrLf & "　　【國籍】　　　　　　　　　" & _
'                    vbCrLf & "　　【中文姓名】　　　　　　　"
'      '專利處
'      If m_bolShowEng = True Then
'         strInventor = strInventor & vbCrLf & "　　【英文姓名】　　　　　　　"
'      End If
'   End If
'   ii = ii + 1
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','發明人資料','" & strInventor & "')"
   'Add By Sindy 2022/4/29 目前這函數使用在新案基本資料表, 要顯示發明人資料
   'If bolShowInvEmp = True Then
      ii = ii + 1
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','發明人要印','♀')"
   'End If
   '2022/4/29 END
   
   'Add By Sindy 2018/12/5 有主張優先權
   If PUB_ChkCPExist(cp, "106") Or PUB_ChkCPExist(cp, "121") Then
   '2018/12/5 END
      '優先權資料
      strExc(0) = "SELECT sqldatew(pd05) pd05,na72,pd06,pd07,decode(pd08,'1','發明','2','新型','3','設計',pd08) pd08,pd09" & _
         " FROM pridate,nation where pd01='" & pa(1) & "' and pd02='" & pa(2) & "' and pd03='" & pa(3) & "' and pd04='" & pa(4) & "'" & _
         " and na01(+)=pd07"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         jj = 0
         Do While Not RsTemp.EOF
            jj = jj + 1
            If jj > 10 Then
               MsgBox "優先權資料超過 10 筆，請自行維護！"
               Exit Do
            End If
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-日','" & RsTemp("pd05") & "')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-國','" & RsTemp("na72") & "')"
            ii = ii + 1
            strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-號','" & RsTemp("pd06") & "')"
            '輸入優先權國家代碼時,代表是以電子交換檢送
            If RsTemp("pd07") = "" & RsTemp("pd09") Then
               '電子交換
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-碼','交換')"
            ElseIf Not IsNull(RsTemp("pd09")) Then
               '非電子交換
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-種類','" & IIf("" & RsTemp("pd08") = "", "♀", ChgSQL("" & RsTemp("pd08"))) & "')"
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','優先權" & jj & "-碼','" & ChgSQL(RsTemp("pd09")) & "')"
            End If
            RsTemp.MoveNext
         Loop
      End If
   End If
   
'   ii = ii + 1
'   strTemp = ""
'   If GetPrjPeople1(GetPrjPeopleNum1(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))) <> "" Then
'      strTemp = GetPrjPeople1(GetPrjPeopleNum1(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)))
'   End If
'   If GetPrjPeople1(GetPrjPeopleNum2(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))) <> "" Then
'      strTemp = strTemp & "、" & GetPrjPeople1(GetPrjPeopleNum2(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)))
'   End If
'   If GetPrjPeople1(GetPrjPeopleNum3(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))) <> "" Then
'      strTemp = strTemp & "、" & GetPrjPeople1(GetPrjPeopleNum3(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)))
'   End If
'   If GetPrjPeople1(GetPrjPeopleNum4(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))) <> "" Then
'      strTemp = strTemp & "、" & GetPrjPeople1(GetPrjPeopleNum4(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)))
'   End If
'   If GetPrjPeople1(GetPrjPeopleNum5(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4))) <> "" Then
'      strTemp = strTemp & "、" & GetPrjPeople1(GetPrjPeopleNum5(pa(1) & "-" & pa(2) & "-" & pa(3) & "-" & pa(4)))
'   End If
'   strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
'      " VALUES ('" & ET01 & "','" & strCP09 & "','" & ET03 & "','" & strUserNum & "','收據抬頭','" & strTemp & "')"
   'Modify By Sindy 2020/4/9 收據抬頭(3)
   Call PUB_ReadPToAppBaseData(pa(1), pa(2), pa(3), pa(4), 3, , ET01, strCP09, ET03, ii, strTxt())
   
   If Not ClsLawExecSQL(ii, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   End If
End Sub

'Add By Sindy 2018/12/17 更新審定號作業
Public Function frm030603_Process(strTMBM07 As String) As Long
Dim strSql As String
Dim strTemp As String
Dim rsTmp As New ADODB.Recordset
Dim nAffect As Long
Dim nCount As Long
Dim nTotal As Long
'Add By Cheng 2003/05/16
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
'Add By Cheng 2003/05/16
'刪除商標已公告缺公告日暫存檔資料
StrSQLa = "Delete From R030603 Where ID='" & strUserNum & "' "
cnnConnection.Execute StrSQLa
 '911107 nick transation
On Error GoTo CheckingErr
cnnConnection.BeginTrans
    
   nAffect = 0
   frm030603_Process = 0
    'Add By Cheng 2003/05/16
    '申請國家為台灣者
    'Modify By Cheng 2003/06/24
'    strSQLA = "Select TM01, TM02, TM03, TM04, TM12, TM15,'" & strUserNum & "' From Trademark " & _
'                    "WHERE TM12 IN (SELECT TMBM04 AS TM12 FROM TMBULLETIN " & _
'                    "WHERE TMBM07 = '" & textTMBM07 & "') AND " & _
'                    "(TM01 = 'T' OR TM01 = 'FCT') And TM14 Is Null And TM10 < '010' "
    StrSQLa = "Select TM01, TM02, TM03, TM04, TM12, TMBM01,'" & strUserNum & "' From Trademark, TMBULLETIN " & _
                    "WHERE TM12=TMBM04 And " & _
                    " TMBM07 = '" & strTMBM07 & "' AND " & _
                    "(TM01 = 'T' OR TM01 = 'FCT') And TM14 Is Null And TM10 < '010' "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            StrSQLa = "Insert Into R030603 Values ('" & rsA.Fields(0).Value & "','" & rsA.Fields(1).Value & "','" & rsA.Fields(2).Value & "','" & rsA.Fields(3).Value & "','" & rsA.Fields(4).Value & "','" & rsA.Fields(5).Value & "','" & rsA.Fields(6).Value & "' )"
            cnnConnection.Execute StrSQLa
            rsA.MoveNext
        Wend
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
   ' 以公報卷期找商標公報檔的公報卷期, 在所找到的記錄中以申請案號找尋商標基本檔的申請案號欄且申請國家必須為台灣, 系統別必須為T或FCT, 若有找到則更新商標基本檔的審定號為商標公報檔的審定號
   strSql = "UPDATE TRADEMARK SET TM15 = (SELECT TMBM01 FROM TMBULLETIN " & _
                                         "WHERE TMBM07 = '" & strTMBM07 & "' AND " & _
                                               "TMBM04 = TM12) " & _
            "WHERE TM12 IN (SELECT TMBM04 AS TM12 FROM TMBULLETIN " & _
                           "WHERE TMBM07 = '" & strTMBM07 & "') AND " & _
                  "(TM01 = 'T' OR TM01 = 'FCT') And TM10 < '010' "
   cnnConnection.Execute strSql, nAffect
   frm030603_Process = nAffect
   
   'add by sonia 2022/10/5 卷宗性質為申請且沒有准駁案件要更新為准T-209148(105054824)
   strSql = "UPDATE TRADEMARK SET TM16=DECODE(TM16,NULL,'1',TM16) " & _
            "WHERE TM12 IN (SELECT TMBM04 AS TM12 FROM TMBULLETIN " & _
                           "WHERE TMBM07 = '" & strTMBM07 & "') AND " & _
                  "(TM01 = 'T' OR TM01 = 'FCT') And TM10 < '010' AND TM28='1'"
   cnnConnection.Execute strSql
   'end 2022/10/5
   
   nTotal = 0
   nCount = 0
   ' 以公報卷期找商標公報檔的公報卷期, 在所找到的記錄中以申請案號找尋商標基本檔的正商標號數且申請國家必須為台灣, 系統別必須為T或FCT, 若有找到則更新商標基本檔的正商標號數為商標公報檔的審定號
   strSql = "SELECT TMBM04 FROM TMBULLETIN " & _
            "WHERE TMBM07 = '" & strTMBM07 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   strTemp = Empty
   If rsTmp.RecordCount > 0 Then
      rsTmp.MoveFirst
      Dim nIndex As Integer
      nIndex = 0
      Do While rsTmp.EOF = False
         If IsNull(rsTmp.Fields("TMBM04")) = False Then
            If IsEmptyText(rsTmp.Fields("TMBM04")) = False Then
               If IsEmptyText(strTemp) = False Then: strTemp = strTemp & ","
               strTemp = strTemp & "'" & rsTmp.Fields("TMBM04") & "'"
               nIndex = nIndex + 1
               ' 90.10.09 modify by louis (串列數目不可超過254)
               If nIndex > 250 Then
                  strSql = "UPDATE TRADEMARK SET TM27 = (SELECT TMBM01 FROM TMBULLETIN " & _
                                            "WHERE TMBM07 = '" & strTMBM07 & "' AND " & _
                                                  "TMBM04 = TM27) " & _
                           "WHERE TM27 IN (" & strTemp & ") AND " & _
                                 "(TM01 = 'T' OR TM01 = 'FCT') And TM10 < '010' "
                  cnnConnection.Execute strSql, nCount
                  
                  nTotal = nTotal + nCount
                  strTemp = Empty
                  nIndex = 0
               End If
            End If
         End If
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   
   If IsEmptyText(strTemp) = False Then
      strSql = "UPDATE TRADEMARK SET TM27 = (SELECT TMBM01 FROM TMBULLETIN " & _
                                            "WHERE TMBM07 = '" & strTMBM07 & "' AND " & _
                                                  "TMBM04 = TM27) " & _
               "WHERE TM27 IN (" & strTemp & ") AND " & _
                     "(TM01 = 'T' OR TM01 = 'FCT') And TM10 < '010' "
      cnnConnection.Execute strSql, nCount
      nTotal = nTotal + nCount
      If frm030603_Process = 0 Then
         'frm030603_Process = nAffect
         frm030603_Process = nTotal
      End If
   End If
   Set rsTmp = Nothing
                           
   'strSQL = "UPDATE TRADEMARK SET TM27 = (SELECT TMBM01 FROM TMBULLETIN " & _
   '                                      "WHERE TMBM07 = '" & textTMBM07 & "' AND " & _
   '                                            "TMBM04 = TM27) " & _
   '         "WHERE TM27 IN (SELECT TMBM04 AS TM27 FROM TMBULLETIN " & _
   '                        "WHERE TMBM07 = '" & textTMBM07 & "') AND " & _
   '               "(TM01 = 'T' OR TM01 = 'FCT') AND " & _
   '               "TM10 < '010' "
   'cnnConnection.Execute strSQL
   
 '911107 nick transation
  cnnConnection.CommitTrans
     Exit Function
CheckingErr:
    MsgBox (Err.Description)
     cnnConnection.RollbackTrans
End Function

'Add By Sindy 2019/4/18
Public Sub PUB_SetListScroll(oForm As Form, oList As ListBox)
   Dim ii As Integer
   Dim lWnow As Long, lWmax As Long
   
   lWmax = 0
   For ii = 0 To oList.ListCount - 1
      lWnow = oForm.TextWidth(oList.List(ii) & " ")
      If lWnow > lWmax Then
         lWmax = lWnow
      End If
   Next
  
   If oForm.ScaleMode = vbTwips Then lWmax = lWmax / Screen.TwipsPerPixelX  ' if twips change to pixels
   SendMessageByNum oList.hWnd, LB_SETHORIZONTALEXTENT, lWmax, 0
End Sub

'Add By Sindy 2021/4/22 抓取信件主檔資料
Public Function PUB_GetInputData(strII01 As String, strII03 As String, strCol As String) As String
Dim rsTmp As New ADODB.Recordset
   
   PUB_GetInputData = ""
   If strII01 = "" Or strII03 = "" Then Exit Function 'Add By Sindy 2021/5/6
   
   '專利處信箱
   If Len(strII03) = 5 And Left(strII03, 1) = "P" Then
      If strCol = "主旨" Then
         strCol = "pi17"
      End If
      strSql = "select pi01,pi03," & strCol & " from patentinput" & _
               " where pi01=" & strII01 & " and pi03='" & ChgSQL(strII03) & "'"
   '商標處信箱
   ElseIf Len(strII03) = 5 And Left(strII03, 1) = "T" Then
      If strCol = "主旨" Then
         strCol = "ti17"
      End If
      strSql = "select Ti01,Ti03," & strCol & " from TMinput" & _
               " where Ti01=" & strII01 & " and Ti03='" & ChgSQL(strII03) & "'"
   '國外部信箱
   Else
      If strCol = "主旨" Then
         strCol = "ii17"
      End If
      strSql = "select ii01,ii03," & strCol & " from ipdeptinput" & _
               " where ii01=" & strII01 & " and ii03='" & ChgSQL(strII03) & "'"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      PUB_GetInputData = "" & rsTmp.Fields(strCol)
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Add By Sindy 2019/6/4
'從其他信箱匯入,且轉寄的收受者為何
Public Function PUB_GetMailInputData(strII01 As String, strII03 As String) As String
Dim rsTmp As New ADODB.Recordset
   
   PUB_GetMailInputData = ""
   '專利處信箱
   If Len(strII03) = 5 And Left(strII03, 1) = "P" Then
      '檢查有無國外部收受者
      strSql = "select pi15 as note,ii06 as emp,ii18 as remark from patentinput,ipdeptinput" & _
               " where pi01=" & strII01 & " and pi03='" & ChgSQL(strII03) & "'" & _
               " and pi01=ii08(+) and instr(ii15,pi03)>0"
      '檢查有無商標處收受者
      strSql = strSql & " union " & _
               "select pi15 as note,ti06 as emp,ti15 as remark from patentinput,TMinput" & _
               " where pi01=" & strII01 & " and pi03='" & ChgSQL(strII03) & "'" & _
               " and pi01=Ti08(+) and instr(Ti22,pi03)>0"
   '商標處信箱
   ElseIf Len(strII03) = 5 And Left(strII03, 1) = "T" Then
      '檢查有無國外部收受者
      strSql = "select ti15 as note,ii06 as emp,ii18 as remark from TMinput,ipdeptinput" & _
               " where Ti01=" & strII01 & " and Ti03='" & ChgSQL(strII03) & "'" & _
               " and Ti01=ii08(+) and instr(ii15,Ti03)>0"
      '檢查有無專利處收受者
      strSql = strSql & " union " & _
               "select ti15 as note,pi06 as emp,pi15 as remark from TMinput,patentinput" & _
               " where Ti01=" & strII01 & " and Ti03='" & ChgSQL(strII03) & "'" & _
               " and Ti01=Pi08(+) and instr(Pi22,Ti03)>0"
   '國外部信箱
   Else
      '檢查有無商標處收受者
      strSql = "select ii18 as note,ti06 as emp,ti15 as remark from ipdeptinput,TMinput" & _
               " where ii01=" & strII01 & " and ii03='" & ChgSQL(strII03) & "'" & _
               " and ii01=Ti08(+) and instr(Ti22,ii03)>0"
      '檢查有無專利處收受者
      strSql = strSql & " union " & _
               "select ii18 as note,pi06 as emp,pi15 as remark from ipdeptinput,patentinput" & _
               " where ii01=" & strII01 & " and ii03='" & ChgSQL(strII03) & "'" & _
               " and ii01=Pi08(+) and instr(Pi22,ii03)>0"
   End If
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      If "" & rsTmp.Fields("note") <> "" Then
         If Left(UCase(Trim(rsTmp.Fields("note"))), 2) = "TM" Then
            PUB_GetMailInputData = "(TM)=>"
         ElseIf Left(UCase(Trim(rsTmp.Fields("note"))), 6) = "PATENT" Then
            PUB_GetMailInputData = "(PATENT)=>"
         ElseIf Left(UCase(Trim(rsTmp.Fields("note"))), 6) = "IPDEPT" Then
            PUB_GetMailInputData = "(IPDEPT)=>"
         End If
      End If
      PUB_GetMailInputData = PUB_GetMailInputData & PUB_ReadUserData("" & rsTmp.Fields("emp"))
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Add By Sindy 2019/6/20 整理收受者資訊
Public Function PUB_IR04DataMakeUp(strPI06 As String) As String
Dim strTmp As String
Dim tmpArr As Variant
Dim j As Integer
   
   strPI06 = Trim(strPI06)
   PUB_IR04DataMakeUp = strPI06
   If UCase(strPI06) = "PATENT" Or UCase(strPI06) = "IPDEPT" Or UCase(strPI06) = "TM" Then
      PUB_IR04DataMakeUp = LCase(strPI06) '小寫
   End If
   If InStr(UCase(strPI06), "@") > 0 Or _
      InStr(UCase(strPI06), ";") > 0 Then
      strTmp = ""
      tmpArr = Split(strPI06, ";")
      For j = 0 To UBound(tmpArr)
         If tmpArr(j) <> "" Then
            If InStr(tmpArr(j), "@") > 0 Then
               tmpArr(j) = Mid(tmpArr(j), 1, InStr(tmpArr(j), "@") - 1)
               tmpArr(j) = LCase(tmpArr(j)) '小寫
            End If
            strTmp = strTmp & ";" & tmpArr(j)
         End If
      Next j
      If strTmp <> "" Then strTmp = Mid(strTmp, 2)
      PUB_IR04DataMakeUp = strTmp
   End If
End Function

'Add By Sindy 2020/4/7
'申請書:出名代理人 0,1,2
'申請書:收據抬頭 3
Public Function PUB_ReadPToAppBaseData(pCP01 As String, pCP02 As String, pCP03 As String, pCP04 As String, _
   pIntShowKind As Integer, Optional pCP110 As String, _
   Optional pET01 As String, Optional pReceiveNo As String, Optional pET03 As String, _
   Optional ByRef pRow As Integer, Optional ByRef strTxt As Variant, Optional strTitName As String = "") As String
   
Dim srtQ As String, intQ As Integer
Dim RsQ As ADODB.Recordset
Dim varTmp As Variant
Dim jj As Integer
Dim strTemp As String
   
   Select Case pIntShowKind
      Case 0, 1, 2 '出名代理人
         If pCP110 <> "" Then
'            varTmp = Split(pCP110, ",")
'            For jj = 0 To UBound(varTmp)
'               srtQ = "select st01,st02,st26,oa08 from ouragent,staff" & _
'                      " where st01='" & varTmp(jj) & "' and oa01='" & pCP01 & "' and st01=oa02" & _
'                      " order by OA03"
'               intQ = 1
'               Set RsQ = ClsLawReadRstMsg(intQ, srtQ)
'               If intQ = 1 Then
            srtQ = "select st01,st02,st26,oa08 from ouragent,staff" & _
                   " where oa01='" & pCP01 & "' and instr('" & pCP110 & "',oa02)>0 and st01(+)=oa02" & _
                   " order by OA03"
            intQ = 1
            Set RsQ = ClsLawReadRstMsg(intQ, srtQ)
            If intQ = 1 Then
               RsQ.MoveFirst
               jj = 0
               Do While Not RsQ.EOF
                  jj = jj + 1
                  Select Case pIntShowKind
                     Case 0
                        '專利處
                        If Left(Pub_StrUserSt03, 2) = "P1" Then
                           'ex:桂齊恆、林景郁
                           PUB_ReadPToAppBaseData = PUB_ReadPToAppBaseData & "、" & RsQ.Fields("st02")
                        Else
                           'ex:桂,齊恆、林,景郁
                           PUB_ReadPToAppBaseData = PUB_ReadPToAppBaseData & "、" & PUB_ConvertNameFormat("" & RsQ.Fields("st02"))
                        End If
                     Case 1
                        pRow = pRow + 1
                        strTxt(pRow) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           " VALUES ('" & pET01 & "','" & pReceiveNo & "','" & pET03 & "','" & strUserNum & "','" & strTitName & "代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & RsQ.Fields("st02")) & "')"
                     Case 2
                        pRow = pRow + 1
                        strTxt(pRow) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           " VALUES ('" & pET01 & "','" & pReceiveNo & "','" & pET03 & "','" & strUserNum & "','" & strTitName & "代理人" & jj & "-中文姓名','" & PUB_ConvertNameFormat("" & RsQ.Fields("st02")) & "')"
                        pRow = pRow + 1
                        strTxt(pRow) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           " VALUES ('" & pET01 & "','" & pReceiveNo & "','" & pET03 & "','" & strUserNum & "','" & strTitName & "代理人" & jj & "-證書字號','" & RsQ.Fields("oa08") & "')"
                        pRow = pRow + 1
                        strTxt(pRow) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                           " VALUES ('" & pET01 & "','" & pReceiveNo & "','" & pET03 & "','" & strUserNum & "','" & strTitName & "代理人" & jj & "-ID','" & RsQ.Fields("ST26") & "')"
                  End Select
                  RsQ.MoveNext
               Loop
               If PUB_ReadPToAppBaseData <> "" Then PUB_ReadPToAppBaseData = Mid(PUB_ReadPToAppBaseData, 2)
            End If
         End If
      Case 3 '收據抬頭
         strTemp = ""
         If GetPrjPeople1(GetPrjPeopleNum1(pCP01 & "-" & pCP02 & "-" & pCP03 & "-" & pCP04)) <> "" Then
            strTemp = GetPrjPeople1(GetPrjPeopleNum1(pCP01 & "-" & pCP02 & "-" & pCP03 & "-" & pCP04), , True)
         End If
         If GetPrjPeople1(GetPrjPeopleNum2(pCP01 & "-" & pCP02 & "-" & pCP03 & "-" & pCP04)) <> "" Then
            strTemp = strTemp & "、" & GetPrjPeople1(GetPrjPeopleNum2(pCP01 & "-" & pCP02 & "-" & pCP03 & "-" & pCP04), , True)
         End If
         If GetPrjPeople1(GetPrjPeopleNum3(pCP01 & "-" & pCP02 & "-" & pCP03 & "-" & pCP04)) <> "" Then
            strTemp = strTemp & "、" & GetPrjPeople1(GetPrjPeopleNum3(pCP01 & "-" & pCP02 & "-" & pCP03 & "-" & pCP04), , True)
         End If
         If GetPrjPeople1(GetPrjPeopleNum4(pCP01 & "-" & pCP02 & "-" & pCP03 & "-" & pCP04)) <> "" Then
            strTemp = strTemp & "、" & GetPrjPeople1(GetPrjPeopleNum4(pCP01 & "-" & pCP02 & "-" & pCP03 & "-" & pCP04), , True)
         End If
         If GetPrjPeople1(GetPrjPeopleNum5(pCP01 & "-" & pCP02 & "-" & pCP03 & "-" & pCP04)) <> "" Then
            strTemp = strTemp & "、" & GetPrjPeople1(GetPrjPeopleNum5(pCP01 & "-" & pCP02 & "-" & pCP03 & "-" & pCP04), , True)
         End If
         pRow = pRow + 1
         strTxt(pRow) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & pET01 & "','" & pReceiveNo & "','" & pET03 & "','" & strUserNum & "','收據抬頭','" & strTemp & "')"
   End Select
   
   Set RsQ = Nothing
End Function

'Modify By Sindy 2020/12/17 特殊的代理/申請人
'回傳需顯示固定字樣的EMail內容
'  S1:根據陶氏化學的《代表指南》，我們僅將掃描的專利證書副本發送給您。
'Modify By Sindy 2021/3/9
'  S2:原始專利證書將在不久的將來通過快遞發送給ASM America，Inc.的Michelle Sympson女士。
Public Function frm060317_1_SpecNO(strPA26 As String, strPA75 As String, _
   Optional ByRef m_iCopys As Integer, _
   Optional ByRef strSpecNO As Boolean, _
   Optional ByRef bolSpecFax As Boolean) As String
   
   frm060317_1_SpecNO = ""
   strPA26 = ChangeCustomerL(strPA26)
   strPA75 = ChangeCustomerL(strPA75)
   'ADD BY SONIA 2014/4/28 除特定客戶/代理人外不管是否E化都要印名條
   strSpecNO = False
   bolSpecFax = False 'Add By Sindy 2016/2/24
   'MODIFY BY SONIA 2014/5/9 取消Y52218再加入X47833,X47833020,X17901010
   'If "" & rsTmp.Fields("PA75") = "Y52218000" Or "" & rsTmp.Fields("PA75") = "Y20085000" Or "" & rsTmp.Fields("PA26") = "X34291000" Or "" & rsTmp.Fields("PA26") = "X21382010" Then
   'Modify By Sindy 2015/11/18
   m_iCopys = 2 '全部統一印2份
   '特殊客戶要產生電子檔,此處修改名單則basPublic的PUB_PrintFCPEmpBill之證書函07也要修改
   Select Case strPA75
      'Modify By Sindy 2016/9/26 + Y54047000
      'Modify By Sindy 2017/8/30 + Y54747000
      'Modify By Sindy 2019/4/17 + Y52418000
      'Modify By Sindy 2019/9/6 + Y48309000 Y48309010 Y48309030 Y48309040 Y48309050 Y48309080 Y51326000
      ', "Y51982000", "Y27856B30"
      'Modify By Sindy 2020/7/3 + Y55240000 DuPont
      'Modify By Sindy 2020/11/2 + Y22327000 MKS Instruments, Inc
      'Modify By Sindy 2021/1/5 + 請協助設定 Y21775 ROBERT BOSCH GMBH 寄證書定稿只出一份
      Case "Y48292030", "Y52075000", "Y54047000", "Y54747000", "Y52418000", _
           "Y48309000", "Y48309010", "Y48309030", "Y48309040", "Y48309050", _
           "Y48309080", "Y51326000", "Y55240000", "Y22327000", "Y21775000"
         'Modify By Sindy 2020/7/3 Y55240000但排除申請人
         'Y55240 +X48049001 (DOW CORNING TORAY CO., LTD.)
         'Y55240 +X48049000 (DuPont Toray Specialty Materials Kabushiki Kaisha)
         If strPA75 = "Y55240000" And Left(strPA26, 8) = "X4804900" Then
         Else
         '2020/7/3 END
            strSpecNO = True
            m_iCopys = 1
         End If
      'Modify By Sindy 2016/1/20 +Y34310000
      'Modify By Sindy 2016/1/25 +Y20064000
      'Modify By Sindy 2016/2/3  +Y34271000, Y51817040
      'Modify By Sindy 2016/2/24 +Y47168000
      'Modify By Sindy 2016/3/3  +Y20078000
      'Modify By Sindy 2016/3/17 +Y20876010
      'Modify By Sindy 2016/3/23 +Y54391000
      'Modify By Sindy 2016/3/29 +Y20624000
      'Modify By Sindy 2016/3/29 +Y52989000
      'Modify By Sindy 2016/5/5  +Y47032000
      'Modify By Sindy 2016/6/15 +Y51333010
      'Modify By Sindy 2016/6/15 +Y51982000
      'Modify By Sindy 2016/7/13 +Y47453000
      'Modify By Sindy 2016/9/21 +Y52643000
      'Modify By Sindy 2017/10/6 +Y48651000
      'Modified by Lydia 2018/09/27 +Y34210010
      Case "Y27766000", "Y51774000", "Y45204000", "Y34210000", "Y34210010", "Y34210020", "Y34210030", _
           "Y45149000", "Y51742000", "Y34310000", "Y20064000", "Y34271000", "Y51817040", _
           "Y47168000", "Y20078000", "Y20876010", "Y54391000", "Y20624000", "Y52989000", _
           "Y47032000", "Y51333010", "Y51982000", "Y47453000", "Y52643000", "Y48651000"
         strSpecNO = True
         m_iCopys = 2
         'Add By Sindy 2016/2/24
         If strPA75 = "Y47168000" Then
            bolSpecFax = True
         End If
         '2016/2/24 END
      'Add By Sindy 2020/12/17 Tim:檢查9碼足
      Case "Y22457000", "Y22457010", "Y22457020", "Y48842000", "Y48048000", _
           "Y48645000", "Y49562000", "Y52322000", "Y52322B10", "Y55020000"
         frm060317_1_SpecNO = "S1"
         strSpecNO = True
         m_iCopys = 1
         '2020/12/17 END
   End Select
   'Add By Sindy 2020/12/17 Tim:檢查9碼足
   Select Case strPA26
      Case "X22457000", "X27727000", "X48049001", "X48049C10", "X48049C11", _
           "X49346000", "X49346001", "X60507000", "X60507001", "X60507010", _
           "X67402000", "X67402010", "X67402020", "X69605000", "X70197000", _
           "X70749000", "X70406000", "X70406001", "X71137000", "X71927000", _
           "X72756000", "X80705000", "X80705C10"
         frm060317_1_SpecNO = "S1"
         strSpecNO = True
         m_iCopys = 1
   End Select
   '2020/12/17 END
   'modify by sonia 2015/11/30 只抓前8碼
   'Select Case "" & rsTmp.Fields("PA26")
   '   Case "X34291000", "X21382010", "X47833000", "X47833020", "X17901010"
   Select Case Left(strPA26, 8)
      Case "X3429100"
   'end 2015/11/30
         strSpecNO = True
         m_iCopys = 1
   End Select
   '2015/11/18 END
   'Add By Sindy 2016/9/2
   If strPA75 = "Y34232000" And Left(strPA26, 8) = "X4863700" Then
      strSpecNO = True
      m_iCopys = 2
   End If
   '2016/9/2 END
   'Add By Sindy 2021/3/9
   'Y33801 (SNELL & WILMER (PHOENIX OFFICE)) + X68646 (ASM IP Holding B.V.)
   'Y33801B2 (Snell & Wilmer L.L.P. (Los Angeles Office))+ X68646 (ASM IP Holding B.V.)
   'Y33801 (SNELL & WILMER (PHOENIX OFFICE)) + X47178 (ASM AMERICA, INC.)
   'Y33801B2 (Snell & Wilmer L.L.P. (Los Angeles Office))+X47178 (ASM AMERICA, INC.)
   If (strPA75 = "Y33801000" And Left(strPA26, 8) = "X6864600") Or _
      (strPA75 = "Y33801B20" And Left(strPA26, 8) = "X6864600") Or _
      (strPA75 = "Y33801000" And Left(strPA26, 8) = "X4717800") Or _
      (strPA75 = "Y33801B20" And Left(strPA26, 8) = "X4717800") Then
      frm060317_1_SpecNO = "S2"
      strSpecNO = True
      m_iCopys = 2
   End If
   '2021/3/9 END
End Function

'Add By Sindy 2021/10/14 ACS顧服組
'沒有核完的控管2個工作天(不含當天)發mail提醒，進承辦人工作進度時也要提醒
'沒有判發的控管2個工作天(不含當天)發mail提醒，進承辦人工作進度時也要提醒
'判發後隔天沒有發文的要發mail提醒，進承辦人工作進度時也要提醒
Public Function PUB_ChkEmpElePro(strEmp As String, strType As String, _
   Optional ByRef rsTmp As ADODB.Recordset) As Boolean
   
Dim m_EMPCon As String, strDept As String
'Dim rsTmp As New ADODB.Recordset
Dim strConSql As String
   
   'A:有已逾時的待核判案件
   'B:有之前尚未歸檔欲發文的案件
   PUB_ChkEmpElePro = False
   
   If strType = "A" Then
      m_EMPCon = "'" & EMP_送英核 & "','" & EMP_送核 & "','" & EMP_送判 & "','" & EMP_轉回 & "'"
      strConSql = " And e1.EEP09='Y' And e1.EEP06<=" & CompWorkDay(3, strSrvDate(1), 1)
      strDept = "W20" 'TEST: strDept = "P11"
      
   Else 'strType = "B"
      m_EMPCon = "'" & EMP_判發 & "'"
      strConSql = " And e1.EEP13='Y' And e1.EEP06<" & strSrvDate(1)
      strDept = "W20" 'TEST: strDept = "P22": m_EMPCon = "'" & EMP_送件 & "'"
   End If
   
   strSql = "Select ' ' as V,EP01 as 目次,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,PA05||PA06||PA07 as 案件名稱," & _
            "NA03 as 國家,Decode(PA09,'000',PTM03,PTM04) as 種類,Decode(PA09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.EEP06 a,e1.EEP07 b,'' as ChkDate,cpm02,EEP15,e1.EEP05 c" & _
            " From EmpElectronProcess e1,CaseProgress,EngineerProgress,Patent,staff s1,staff s2,staff s3,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
            " Where" & IIf(strEmp <> "", " e1.EEP05='" & strEmp & "' And", "") & " e1.EEP04 in(" & m_EMPCon & ")" & strConSql & _
            " AND e1.EEP01=CP09(+)" & _
            " And e1.EEP01=EP02(+)" & _
            " And CP01=PA01 And CP02=PA02 And CP03=PA03 And CP04=PA04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+) And e1.EEP05=s3.ST01(+) And s3.ST15='" & strDept & "'" & _
            " And PA09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '1'=PTM01(+) AND PA08=PTM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " and cp158=0 and cp159=0"
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,TM05||TM06||TM07 as 案件名稱," & _
            "NA03 as 國家,Decode(TM10,'000',PTM03,PTM04) as 種類,Decode(TM10,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.EEP06 a,e1.EEP07 b,decode(e1.EEP09,'Y',decode(sign(e1.EEP07-160000),1,to_char(to_date(e1.EEP06,'YYYYMMDD')+1,'YYYYMMDD'),e1.EEP06),'') AS ChkDate,cpm02,EEP15,e1.EEP05 c" & _
            " From EmpElectronProcess e1,CaseProgress,EngineerProgress,Trademark,staff s1,staff s2,staff s3,nation,CasePropertyMap,PatentTradeMarkMap,allcode" & _
            " Where " & IIf(strEmp <> "", " e1.EEP05='" & strEmp & "' And", "") & " e1.EEP04 in(" & m_EMPCon & ")" & strConSql & _
            " AND e1.EEP01=CP09(+)" & _
            " And e1.EEP01=EP02(+)" & _
            " And CP01=TM01 And CP02=TM02 And CP03=TM03 And CP04=TM04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+) And e1.EEP05=s3.ST01(+) And s3.ST15='" & strDept & "'" & _
            " And TM10=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And '2'=PTM01(+) AND TM08=PTM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " and cp158=0 and cp159=0"
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,SP05||SP06||SP07 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode(SP09,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.EEP06 a,e1.EEP07 b,decode(e1.EEP09,'Y',decode(sign(e1.EEP07-160000),1,to_char(to_date(e1.EEP06,'YYYYMMDD')+1,'YYYYMMDD'),e1.EEP06),'') AS ChkDate,cpm02,EEP15,e1.EEP05 c" & _
            " From EmpElectronProcess e1,CaseProgress,EngineerProgress,servicepractice,staff s1,staff s2,staff s3,nation,CasePropertyMap,allcode" & _
            " Where " & IIf(strEmp <> "", " e1.EEP05='" & strEmp & "' And", "") & " e1.EEP04 in(" & m_EMPCon & ")" & strConSql & _
            " AND e1.EEP01=CP09(+)" & _
            " And e1.EEP01=EP02(+)" & _
            " And CP01=SP01 And CP02=SP02 And CP03=SP03 And CP04=SP04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+) And e1.EEP05=s3.ST01(+) And s3.ST15='" & strDept & "'" & _
            " And SP09=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " and cp158=0 and cp159=0"
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,LC05||LC06||LC07 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode(LC15,'000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.EEP06 a,e1.EEP07 b,decode(e1.EEP09,'Y',decode(sign(e1.EEP07-160000),1,to_char(to_date(e1.EEP06,'YYYYMMDD')+1,'YYYYMMDD'),e1.EEP06),'') AS ChkDate,cpm02,EEP15,e1.EEP05 c" & _
            " From EmpElectronProcess e1,CaseProgress,EngineerProgress,Lawcase,staff s1,staff s2,staff s3,nation,CasePropertyMap,allcode" & _
            " Where " & IIf(strEmp <> "", " e1.EEP05='" & strEmp & "' And", "") & " e1.EEP04 in(" & m_EMPCon & ")" & strConSql & _
            " AND e1.EEP01=CP09(+)" & _
            " And e1.EEP01=EP02(+)" & _
            " And CP01=LC01 And CP02=LC02 And CP03=LC03 And CP04=LC04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+) And e1.EEP05=s3.ST01(+) And s3.ST15='" & strDept & "'" & _
            " And LC15=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " and cp158=0 and cp159=0"
   strSql = strSql & " union " & _
            "Select ' ' as V,EP01 as 目次,SqlDateT(e1.EEP06)||' '||sqltime(e1.EEP07) as 流程日期,CP01||'-'||CP02||'-'||CP03||'-'||CP04 as 本所案號,HC06 as 案件名稱," & _
            "NA03 as 國家,'' as 種類,Decode('000','000',CPM03,CPM04) as 案件性質," & _
            "SqlDateT (cp06) as 本所期限, s1.ST02 as 承辦人, SqlDateT(cp48) as 承辦期限, s2.ST02 as 智權人員, ac03 as 目前流程狀態," & _
            "e1.EEP01 as 總收文號,' ' as 不顯示,e1.EEP02,e1.EEP06 a,e1.EEP07 b,decode(e1.EEP09,'Y',decode(sign(e1.EEP07-160000),1,to_char(to_date(e1.EEP06,'YYYYMMDD')+1,'YYYYMMDD'),e1.EEP06),'') AS ChkDate,cpm02,EEP15,e1.EEP05 c" & _
            " From EmpElectronProcess e1,CaseProgress,EngineerProgress,Hirecase,staff s1,staff s2,staff s3,nation,CasePropertyMap,allcode" & _
            " Where " & IIf(strEmp <> "", " e1.EEP05='" & strEmp & "' And", "") & " e1.EEP04 in(" & m_EMPCon & ")" & strConSql & _
            " AND e1.EEP01=CP09(+)" & _
            " And e1.EEP01=EP02(+)" & _
            " And CP01=HC01 And CP02=HC02 And CP03=HC03 And CP04=HC04" & _
            " And CP14=s1.ST01(+) And CP13=s2.ST01(+) And e1.EEP05=s3.ST01(+) And s3.ST15='" & strDept & "'" & _
            " And '000'=NA01(+)" & _
            " And CP01=CPM01(+) And CP10=CPM02(+)" & _
            " And ac01='09' And e1.EEP04=ac02(+)" & _
            " and cp158=0 and cp159=0"
   strSql = strSql & " order by c asc,a desc,b desc"
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      PUB_ChkEmpElePro = True
   Else
      rsTmp.Close
      Set rsTmp = Nothing
   End If
End Function

'Addde by Lydia 2019/12/19 外專人員執行整批更換FC代理人作業，發案件清單給承辦和程序人員。
'Modified by Lydia 2022/02/24  從frm110104_2搬來: + pType, pFilePath, pWorkDate, pContent
Public Function PUB_ChgPA75List(ByVal pCaseList As String, ByVal pType As String, ByVal pFilePath As String, Optional ByVal pWorkDate As String, Optional ByVal pContent As String) As Boolean
'pType: 0-更換代理人(錯誤會彈訊息), 1-每日批次
'pFilePath: 傳入檔案路徑
'pWorkDate: 0-更換代理人(執行日期), 1-每日批次(前一工作天)
'pContent: 傳入Email內文
Dim intR As Integer
Dim rsRD As New ADODB.Recordset
Dim strR1 As String
Dim strTitle1 As String, strCon1 As String
Dim arrTitle As Variant, arrTmp As Variant
Dim strGrp As String
Dim iRow As Integer, maxCol As Integer
Dim strTemp As String
Dim strFileName As String, strFilePath As String
Dim strToCC As String
Dim xlsReport
Dim wksReport
'Added by Lydia 2019/12/30
Dim mESeqNo As String '暫存檔的序號
Dim rsAD As New ADODB.Recordset
Dim intPS As Integer  '特殊備註的數量
Dim intA As Integer

On Error GoTo ErrHandle
    
    PUB_ChgPA75List = True
    pCaseList = Replace(pCaseList, "N", "")
    If InStr(pCaseList, "'") = 0 Then pCaseList = GetAddStr(pCaseList)
    
    'Added by Lydia 2019/12/30 先將特殊備註設定丟到暫存檔
    strR1 = "Select '10' as ord1,'下一程序固定備註' as ctitle,Nm03,Nm02,'' as cmemo From Npmemo Where Nm03 In (" & pCaseList & ") And Nm02 Is Not Null"
    strR1 = strR1 & " Union All Select '20' as ord1,'核准函輸入備註' as ctitle, Am03,Am02,'' as cmemo From Approvalmemo2 Where Am03 In (" & pCaseList & ") And Am02 Is Not Null"
    strR1 = strR1 & " union all Select '30' as ord1,'核駁及審查意見通知函' as ctitle,Im03,Im02,'' as cmemo From Incommemo Where Im03 In (" & pCaseList & ") And Im02 Is Not Null"
    strR1 = strR1 & " union all Select '40' as ord1,'請款函預設備註' as ctitle,Dnps03, Dnps02 ,'' as cmemo From Debitnoteps Where Dnps03 In (" & pCaseList & ") And Dnps02 Is Not Null"
    strR1 = strR1 & " union all Select '50' as ord1,'承辦單設定'||'-'||feb06 as ctitle,feb03, feb02,feb06||'-'||decode(feb06,'01','告准','02','公告公報','03','寄證書','04','繳年費通知','05','實審請款','06','年證費請款','(補)')||':' as cmemo  From FcpEMPbill Where feb03 In (" & pCaseList & ") And feb02 Is Not Null"
    strR1 = strR1 & " union all Select '60' as ord1,'通知告准加註' as ctitle,aps03, aps02,'' as cmemo From Approvalps Where aps03 In (" & pCaseList & ") And aps02 Is Not Null"
    strR1 = strR1 & " Order by 1, 2"
    intR = 1
    intPS = 6
    Set rsRD = ClsLawReadRstMsg(intR, strR1)
    If intR = 1 Then
        'Modified by Lydia 2022/02/24 Me.Name => PUB_ChgPA75List
        Set rsAD = PUB_CreateRecordset(rsRD, , , , "PUB_ChgPA75List", mESeqNo)
    End If
    'end 2019/12/30

    strCon1 = "PA48,PA49,PA50,PA151,PA152," & _
                   "PA88,PA89,PA90,PA146," & _
                   "PA71,PA70,PA156," & _
                   "PA78,PA133,PA134," & _
                   "PA76,PA135,PA159," & _
                   "PA105,PA106,PA107," & _
                   "PA142,PA155"
    'Modified by Lydia 2019/12/30 +特殊備註維護PS01~PS06(下一程序固定備註、核准函輸入備註、核駁及審查意見通知函、請款函預設備註、承辦單設定、通知告准加註)
                                                 '+案件備註PA91:如有＊符號，請帶出"有備註"
    'strTitle1 = "本所案號," & _
                    "客戶案件案號,全部折扣,申請/翻譯折扣,領證折扣,年費折扣," & _
                    "固定請款對象,不續辦但准通知,信函是否印TITLE,C類收文是否請款," & _
                    "領證自動代繳,年費自動代繳,年費特殊管制," & _
                    "D/N是否列印申請人,D/N固定列印對象,年費D/N列印對象," & _
                    "年費代理人,年費聯絡人,CLIENT_MATTER_ID," & _
                    "年費請款對象,年費彼所案號,年費單筆不跑," & _
                    "是否以Email通知,Email同時寄紙本,案件備註" & _
                    "下一程序固定備註,核准函輸入備註,核駁及審查意見通知函,請款函預設備註,通知告准加註"
    'Modified by Lydia 2020/12/08 案件備註後面+個案各項指示
    'Modify by Amy 2025/08/06 不續辦但准通知 改為 後續准駁簡單報告
    strTitle1 = "本所案號," & _
                    "客戶案件案號,全部折扣,申請/翻譯折扣,領證折扣,年費折扣," & _
                    "固定請款對象,後續准駁簡單報告,信函是否印TITLE,C類收文是否請款," & _
                    "領證自動代繳,年費自動代繳,年費特殊管制," & _
                    "D/N是否列印申請人,D/N固定列印對象,年費D/N列印對象," & _
                    "年費代理人,年費聯絡人,CLIENT_MATTER_ID," & _
                    "年費請款對象,年費彼所案號,年費單筆不跑," & _
                    "是否以Email通知,Email同時寄紙本,案件備註,個案各項指示," & _
                    "下一程序固定備註,核准函輸入備註,核駁及審查意見通知函,請款函預設備註,承辦單設定備註,通知告准加註"
'本所案號,
'客戶案件案號,全部折扣,申請/翻譯折扣,領證折扣,年費折扣,
'固定請款對象,不續辦但准通知,信函是否印TITLE,C類收文是否請款,
'領證自動代繳,年費自動代繳,年費特殊管制,
'D/N是否列印申請人,D/N固定列印對象,'年費D/N列印對象,
'年費代理人,年費聯絡人,CLIENT_MATTER_ID,
'年費請款對象,年費彼所案號,年費單筆不跑,
'是否以Email通知,Email同時寄紙本,案件備註,個案各項指示
    'Modified by Lydia 2022/02/24 改成傳入路徑
    'strFileName = "更換FC代理人案件清單"
    'strR1 = Dir(App.path & "\*" & strFileName & ".*")
    'If strR1 <> "" Then
    '   Kill App.path & "\" & strR1
    'End If
    'strFilePath = App.path & "\" & strSrvDate(1) & "_" & strFileName '不指定為.xls或.xlsx
    If pType = "1" Then
        strFileName = "申請人異動請檢查個案設定之清單"
    Else
        strFileName = "代理人異動請檢查個案設定之清單"  'Memo by Lydia 2022/02/24 Anny: 一併更名
    End If
    strTemp = pFilePath
    If pFilePath <> MsgText(601) Then
        strTemp = Dir(pFilePath, vbDirectory)
    End If
    If strTemp = MsgText(601) Then
        If Dir(App.path & "\" & strUserNum, vbDirectory) = MsgText(601) Then
            MkDir App.path & "\" & strUserNum
        End If
        strTemp = App.path & "\" & strUserNum
        pFilePath = strTemp
    End If
    If pWorkDate = "" Then pWorkDate = strSrvDate(1)
    strFilePath = pFilePath & "\" & pWorkDate & "_" & strFileName   '不指定為.xls或.xlsx
    If pContent = "" Then pContent = vbCrLf & vbCrLf & "***詳情請見附件***" & vbCrLf & vbCrLf '傳入Email內文
    'end 2022/02/24
    
    'Added by Lydia 2020/12/08 增加"個案各項指示"
    strR1 = "IC04" '預設外專; 檢查是否已完成確認(依使用者部門)
    If pType = "0" Then 'Added by Lydia 2022/02/24 更換代理人作業
        Select Case Left(Pub_StrUserSt03, 2)
             Case "F2":  strR1 = "IC04"
             Case "P1":  strR1 = "IC06"
             Case "F1":  strR1 = "IC08"
             Case "P2":  strR1 = "IC10"
        End Select
    End If 'Added by Lydia 2022/02/24
    strTemp = "SELECT ITS02,SUM(DECODE(ITS04,NULL,0,1)) CNT,SUM(DECODE(" & strR1 & ",NULL,0,1)) CON " & _
                    "From INSTRUCTIONS, INSTCONFIRM WHERE ITS02 IN (" & pCaseList & ") AND ITS02=IC02(+) GROUP BY ITS02 "
    'end 2020/12/08
    
    '以承辦提供的欄位為準，凡個案欄位有設定值則提供清單excel檔，excel顯示本所案號+指定欄位；
    'excel做為附件直接email給承辦+程序，不用等候承辦回應直接更換FC代理人
    
    'Modified by Lydia 2019/12/30 +特殊備註維護(下一程序固定備註、核准函輸入備註、核駁及審查意見通知函、請款函預設備註、通知告准加註)
                                                 '+案件備註PA91:如有＊符號，請帶出"有備註"
    'strR1 = "SELECT PA01||'-'||PA02||DECODE(PA03||PA04,'000','','-'||PA03||'-'||PA04) AS CASENO, " & strCon1 & _
               ",FA10,NA51,NA16 FROM PATENT,FAGENT,NATION" & _
               " WHERE PA01||PA02||PA03||PA04 IN (" & pCaseList & ")" & _
               " AND " & Replace(strCon1, ",", "||") & " IS NOT NULL" & _
               " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=NA01(+)"
    'strR1 = strR1 & " ORDER BY NA51, 1 " '排序：代理人國籍、本所案號
    If mESeqNo = "" Then
        'Modified by Lydia 2020/12/08 增加"個案各項指示" CNT=各項指示筆數,CON=完成確認
        'strR1 = "SELECT 0 AS V02, PA01||'-'||PA02||DECODE(PA03||PA04,'000','','-'||PA03||'-'||PA04) AS CASENO, " & strCon1 & _
                   ",PA91,FA10,NA51,NA16,PA01,PA02,PA03,PA04 FROM PATENT,FAGENT,NATION" & _
                   " WHERE PA01||PA02||PA03||PA04 IN (" & pCaseList & ")" & _
                   " AND (" & Replace(strCon1, ",", "||") & " IS NOT NULL  OR INSTR(PA91,'＊') > 0 )" & _
                   " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=NA01(+)"
        'Modified by 2024/05/31 NA51改成Case => decode(pa75," & Pub_GetSpecFCP & ",na51)
        strR1 = "SELECT 0 AS V02, PA01||'-'||PA02||DECODE(PA03||PA04,'000','','-'||PA03||'-'||PA04) AS CASENO, " & strCon1 & _
                   ",PA91,' ' AS ITS06,FA10,decode(pa75," & Pub_GetSpecFCP & ",na51) as NA51,NA16,PA01,PA02,PA03,PA04,CNT,CON FROM PATENT,FAGENT,NATION,(" & strTemp & ") " & _
                   " WHERE PA01||PA02||PA03||PA04 IN (" & pCaseList & ") AND PA01||PA02||PA03||PA04=ITS02(+) " & _
                   " AND (" & Replace(strCon1, ",", "||") & " IS NOT NULL  OR INSTR(PA91,'＊') > 0 OR CNT > 0 OR CON > 0)" & _
                   " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=NA01(+)"
    Else
        'Modified by Lydia 2020/12/08 增加"個案各項指示" CNT=各項指示筆數,CON=完成確認
        'strR1 = "SELECT NVL(V02,0) V02 , PA01||'-'||PA02||DECODE(PA03||PA04,'000','','-'||PA03||'-'||PA04) AS CASENO, " & strCon1 & _
                   ",PA91,FA10,NA51,NA16,PA01,PA02,PA03,PA04 FROM PATENT,FAGENT,NATION" & _
                   ",(SELECT R003 AS V01,COUNT(R004) AS V02 FROM RDATAFACTORY WHERE FORMNAME='" & Me.Name & "' AND ID='" & strUserNum & "' AND SEQNO='" & mESeqNo & "' GROUP BY R003 ) VB01 " & _
                   " WHERE PA01||PA02||PA03||PA04 IN (" & pCaseList & ")" & _
                   " AND (" & Replace(strCon1, ",", "||") & " IS NOT NULL  OR INSTR(PA91,'＊') > 0 OR V02>0 )" & _
                   " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=NA01(+)" & _
                   " AND PA01||PA02||PA03||PA04=V01(+)"
        'Modified by Lydia 2022/02/24 Me.Name => PUB_ChgPA75List
        'Modified by 2024/05/31 NA51改成Case => decode(pa75," & Pub_GetSpecFCP & ",na51)
        strR1 = "SELECT NVL(V02,0) V02 , PA01||'-'||PA02||DECODE(PA03||PA04,'000','','-'||PA03||'-'||PA04) AS CASENO, " & strCon1 & _
                   ",PA91,' ' AS ITS06,FA10,decode(pa75," & Pub_GetSpecFCP & ",na51) as NA51,NA16,PA01,PA02,PA03,PA04,CNT,CON FROM PATENT,FAGENT,NATION,(" & strTemp & ") " & _
                   ",(SELECT R003 AS V01,COUNT(R004) AS V02 FROM RDATAFACTORY WHERE FORMNAME='" & "PUB_ChgPA75List" & "' AND ID='" & strUserNum & "' AND SEQNO='" & mESeqNo & "' GROUP BY R003 ) VB01 " & _
                   " WHERE PA01||PA02||PA03||PA04 IN (" & pCaseList & ") AND PA01||PA02||PA03||PA04=ITS02(+) " & _
                   " AND (" & Replace(strCon1, ",", "||") & " IS NOT NULL  OR INSTR(PA91,'＊') > 0 OR V02>0 OR CNT > 0 OR CON > 0)" & _
                   " AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND FA10=NA01(+)" & _
                   " AND PA01||PA02||PA03||PA04=V01(+)"
    End If
    strR1 = strR1 & " ORDER BY NA51, PA01,PA02 " '排序：代理人國籍、本所案號
    'end 2019/12/30
    intR = 1
    Set rsRD = ClsLawReadRstMsg(intR, strR1)
    If intR = 1 Then
        arrTitle = Split(strTitle1, ",")
        maxCol = UBound(arrTitle) + 1
        ReDim arrTmp(1 To maxCol)
        rsRD.MoveFirst
        Do While Not rsRD.EOF
            If strGrp <> "" & rsRD.Fields("NA51") Then
                If strGrp <> "" Then
                    xlsReport.Workbooks(1).SaveAs strFilePath
                    xlsReport.Workbooks.Close
                    'Modified by Lydia 2022/02/24
                    'PUB_SendMail strUserNum, strGrp, "", strSrvDate(1) & "_" & strFileName, "請參考附件。", , strFilePath, True, , , strToCC
                    strR1 = Dir(strFilePath & ".*")
                    'Modified by Lydia 2022/03/09 debug pFilePath & "\" & strTemp=>pFilePath & "\" & strR1
                    PUB_SendMail strUserNum, strGrp, "", pWorkDate & "_" & strFileName, pContent, , pFilePath & "\" & strR1, False, , , strToCC, , , , , IIf(pType = "1", False, True), , , IIf(pType = "1", False, True), , , IIf(pType = "1", False, True)
                    'end 2022/02/24
                    Set wksReport = Nothing
                    Set xlsReport = Nothing
                End If
                strR1 = Dir(strFilePath & ".*")
                If strR1 <> "" Then
                   'Modified by Lydia 2022/02/24
                   'Kill App.path & "\" & strR1
                   Kill pFilePath & "\" & strR1
                End If

                Set xlsReport = CreateObject("Excel.Application")
                xlsReport.SheetsInNewWorkbook = 1
                xlsReport.Workbooks.add
                xlsReport.Visible = False
                
                iRow = 1
                Set wksReport = xlsReport.Worksheets(1)
                wksReport.Cells.NumberFormatLocal = "@"
                '抬頭
                wksReport.Range(Pub_NumberToSystem26(1) & iRow & ":" & Pub_NumberToSystem26(maxCol) & iRow).Value = arrTitle
                For intR = 1 To maxCol
                    wksReport.Range(Pub_NumberToSystem26(intR) & ":" & Pub_NumberToSystem26(intR)).ColumnWidth = 12 '欄寬
                Next intR
                iRow = iRow + 1
                wksReport.Range("B2").Select
                xlsReport.ActiveWindow.FreezePanes = True '凍結窗格
                wksReport.Range("A1").Select
                strGrp = "" & rsRD.Fields("NA51")
                strToCC = "" & rsRD.Fields("NA16")
            End If
            '案件資料
            For intR = 1 To maxCol
                 'Modified by Lydia 2019/12/30 分別處理
                 'arrTmp(intR) = "" & rsRd.Fields(intR) -1
                 arrTmp(intR) = "" & rsRD.Fields(intR)
                 If intR >= maxCol - intPS - 1 Then
                    strR1 = arrTmp(intR) 'Added by Lydia 2020/12/08
                    arrTmp(intR) = ""
                    If intR = maxCol - intPS - 1 Then
                        'Modified by Lydia 2020/12/08
                        'If arrTmp(intR) <> "" And InStr(arrTmp(intR), "＊") > 0 Then
                        If strR1 <> "" And InStr(strR1, "＊") > 0 Then
                            arrTmp(intR) = "有備註" '案件備註: 如有＊符號，請帶出"有備註"
                        End If
                    'Added by Lydia 2020/12/08 個案各項指示
                    ElseIf intR = maxCol - intPS Then
                        If Val("" & rsRD.Fields("cnt")) > 0 Then
                            arrTmp(intR) = "個案指示"
                        End If
                        If Val("" & rsRD.Fields("con")) > 0 And arrTmp(intR - 1) <> "" And InStr(arrTmp(intR - 1), "有備註") > 0 Then
                            arrTmp(intR - 1) = "完成確認"
                        End If
                    'end 2020/12/08
                    End If
                 End If
            Next intR
            If Val("" & rsRD.Fields("V02")) > 0 Then
                For intR = 1 To intPS
                     'Modified by Lydia 2022/02/24 Me.Name => PUB_ChgPA75List
                     strTemp = "SELECT " & IIf(intR = 5, "R005||", "") & "R004 as cmemo FROM RDATAFACTORY WHERE FORMNAME='" & "PUB_ChgPA75List" & "' AND ID='" & strUserNum & "' AND SEQNO='" & mESeqNo & "' " & _
                                     " AND R001 LIKE '" & intR & "%' AND R003='" & rsRD.Fields("PA01") & rsRD.Fields("PA02") & rsRD.Fields("PA03") & rsRD.Fields("PA04") & "' ORDER BY R002"
                     intA = 1
                     Set rsAD = ClsLawReadRstMsg(intA, strTemp)
                     If intA = 1 Then
                        strTemp = rsAD.GetString(adClipString, , , vbCrLf)
                        arrTmp(maxCol - intPS + intR) = Mid(strTemp, 1, Len(strTemp) - 2)
                     Else
                        arrTmp(maxCol - intPS + intR) = ""
                     End If
                Next intR
            End If
            'end 2019/12/30
            wksReport.Range(Pub_NumberToSystem26(1) & iRow & ":" & Pub_NumberToSystem26(maxCol) & iRow).Value = arrTmp
            
            If strToCC = "" Then
                strToCC = "" & rsRD.Fields("NA16")
            ElseIf InStr(strToCC, "" & rsRD.Fields("NA16")) = 0 Then
                strToCC = strToCC & ";" & rsRD.Fields("NA16")
            End If
            iRow = iRow + 1
            rsRD.MoveNext
        Loop
        xlsReport.Workbooks(1).SaveAs strFilePath
        xlsReport.Workbooks.Close
        xlsReport.Quit
        'Mark by Lydia 2022/02/24 改成傳入Email內文
        'strR1 = "智權人員：" & frm110104_1.txtCaseField(4) & " " & frm110104_1.lblSName & vbCrLf & _
                    "變更案件條件：代理人：" & frm110104_1.txtCaseField(1) & " " & frm110104_1.lblAgent & vbCrLf & _
                    "　　　　　　　申請人：" & frm110104_1.txtCaseField(2) & " " & frm110104_1.lblCustomer & vbCrLf & _
                    "新代理人：" & frm110104_1.txtCaseField(3) & " " & frm110104_1.NewAgent & vbCrLf & _
                    "　　　　　　　" & IIf(frm110104_1.Check1.Value = 1, "■", "□") & "含閉卷或銷卷案件　　　　　　" & IIf(frm110104_1.Check4.Value = 1, "■", "□") & "清除案件聯絡人資料" & vbCrLf & _
                    "　　　　　　　" & IIf(frm110104_1.Check2.Value = 1, "■", "□") & "彼所案號清除　　　　　　　　" & IIf(frm110104_1.Check3.Value = 1, "■", "□") & "案件聯絡人同時更改"

        strTemp = Dir(strFilePath & ".*") '取得存檔後的檔名，因為存檔時未指定副檔名
        'Modified by Lydia 2022/02/24
        'PUB_SendMail strUserNum, strGrp, "", strSrvDate(1) & "_" & strFileName, strR1, , App.path & "\" & strTemp, False, , , strToCC
        PUB_SendMail strUserNum, strGrp, "", pWorkDate & "_" & strFileName, pContent, , pFilePath & "\" & strTemp, False, , , strToCC, , , , , IIf(pType = "1", False, True), , , IIf(pType = "1", False, True), , , IIf(pType = "1", False, True)
    End If
    
    Set wksReport = Nothing
    Set xlsReport = Nothing
    Exit Function
    
ErrHandle:
    PUB_ChgPA75List = False
    If Err.Number <> 0 Then
        'Added by Lydia 2022/02/24
        If pType = "1" Then   '每日批次
            WLog Err.Description
        Else
        'end 2022/02/24
            MsgBox Err.Description
        End If
    End If
End Function

'Added by Lydia 2022/03/22 產生用於檢查造字更換的Excel
Public Function PUB_BatchDay113Excel(ByVal pType As String, ByVal pPath As String) As Boolean
'pType: 1=整批檢查(frm001_1)目前匯入資料, 2=直接更新(更新完後自動發送) 3=每日批次(更新完後自動發送)
Dim strQ As String, intQ As Integer
Dim rsQD As New ADODB.Recordset
Dim xlsDATA113
Dim wksData113
Dim intCounter As Integer
Dim strFileName As String
Dim strToUser As String
Dim strWd1
Dim strWd2
   
    If pType = "1" Then
        strQ = "select * from editeudclog where eel01 in (9998,9999) order by 1,2 "
        strToUser = strUserNum
    ElseIf pType = "2" Then
        strQ = "select * from editeudclog where eel01=" & strSrvDate(1) & " and eel09 like '直接更新:%' order by 1,2 "
        strToUser = strUserNum
    ElseIf pType = "3" Then
        strQ = "select * from editeudclog where eel01=" & strSrvDate(1) & " and eel09 like '每日批次:%' order by 1,2 "
    End If
    intQ = 1
    Set rsQD = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 0 Then
         If pType = "1" Then
             MsgBox "無資料可檢查！！"
         ElseIf pType = "2" Then
             MsgBox "無資料可更新！！"
         End If
         PUB_BatchDay113Excel = False
    Else
       If pPath = "" Then
           Call PUB_KillTempFile("*檢查_造字更換表.*")
           pPath = App.path & "\"
       Else
           Call PUB_KillTempFile(pPath & "*檢查_造字更換表.*")
           pPath = App.path & pPath
       End If
       strFileName = pPath & strSrvDate(1) & "檢查_造字更換表"
       rsQD.MoveFirst
       intCounter = 1
       If pType = "3" Then strToUser = "" & rsQD.Fields("EEL03") '每日批次=通知匯入的人員
       Set xlsDATA113 = CreateObject("Excel.Application")
       xlsDATA113.SheetsInNewWorkbook = 1
       xlsDATA113.Workbooks.add
       Set wksData113 = xlsDATA113.Worksheets(1)
       wksData113.Activate
       xlsDATA113.Visible = True
       If Val(xlsDATA113.Version) < 12 Then
           xlsDATA113.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=-4143
       Else
           xlsDATA113.Workbooks(1).SaveAs FileName:=strFileName, FileFormat:=56
       End If
       wksData113.Range("A:A").ColumnWidth = 9
       wksData113.Range("A:A").NumberFormatLocal = "@"
       wksData113.Range("A" & intCounter).Value = "更改前"
       wksData113.Range("B:B").ColumnWidth = 9
       wksData113.Range("B:B").NumberFormatLocal = "@"
       wksData113.Range("B" & intCounter).Value = "內碼"
       wksData113.Range("C:C").ColumnWidth = 9
       wksData113.Range("C:C").NumberFormatLocal = "@"
       wksData113.Range("C" & intCounter).Value = "更改後"
       wksData113.Range("D:D").ColumnWidth = 9
       wksData113.Range("D:D").NumberFormatLocal = "@"
       wksData113.Range("D" & intCounter).Value = "內碼"
       wksData113.Range("E:E").ColumnWidth = 20
       wksData113.Range("E:E").NumberFormatLocal = "@"
       wksData113.Range("E" & intCounter).Value = "處理備註"
       intCounter = intCounter + 1
       Do While Not rsQD.EOF
            strWd1 = Trim(PUB_StringFilter("" & rsQD.Fields("EEL04")))
            strWd2 = Trim(PUB_StringFilter("" & rsQD.Fields("EEL06")))
            wksData113.Range("A" & intCounter).Value = strWd1
            wksData113.Range("B" & intCounter).Value = "" & rsQD.Fields("EEL05")
            wksData113.Range("C" & intCounter).Value = strWd2
            wksData113.Range("D" & intCounter).Value = "" & rsQD.Fields("EEL07")
            If pType = "1" Then
                wksData113.Range("E" & intCounter).Value = IIf("" & rsQD.Fields("EEL09") <> "", "不會更新資料=" & rsQD.Fields("EEL09"), "")
            Else
                wksData113.Range("E" & intCounter).Value = "" & rsQD.Fields("EEL09")
            End If
            intCounter = intCounter + 1
            rsQD.MoveNext
       Loop
       xlsDATA113.Workbooks(1).Save
       xlsDATA113.Quit

       strExc(1) = Dir(strFileName & ".xls")
       If strExc(1) = "" Then strExc(1) = Dir(strFileName & ".xlsx")
       If strExc(1) <> "" Then
           If pType = "1" Then
               strExc(2) = strSrvDate(1) & "檢查_造字更換表"
           ElseIf pType = "2" Then
               strExc(2) = strSrvDate(1) & "直接更新_造字更換表"
           ElseIf pType = "3" Then
               strExc(2) = strSrvDate(1) & "每日批次_造字更換表"
           End If
           PUB_SendMail strUserNum, strToUser, "", strExc(2), "請參考附件；", , pPath & strExc(1), False, , , , , , , , IIf(pType = "3", False, True), , , IIf(pType = "3", False, True), , , IIf(pType = "3", False, True)
       End If
    End If
     Set wksData113 = Nothing
     Set xlsDATA113 = Nothing
     PUB_BatchDay113Excel = True
         
     Exit Function
     
ErrHand01:
     If Err.Number <> 0 Then
        If pType <> "3" Then
            MsgBox "檢查_造字更換表失敗：" & vbCrLf & Err.Description
        Else
            WLog Err.Description
        End If
     End If
End Function

'Added by Lydia 2022/03/22 整批直接更新造字
Public Function Pub_BatchDay113Proc(ByVal pType As String) As Boolean
'pType: 2=直接更新(frm001_1), 3=每日批次StrMenu113
Dim strA1 As String, intA As Integer, intQ As Integer
Dim strShow As String, strUpd As String
Dim tmpShow As Variant, tmpUpd As Variant
Dim rsAD As New ADODB.Recordset
Dim strBeginTime As String
Dim lngCnt As Long, lngTot As Long
Dim strConList As String
Dim strTBname As String
Dim strWd1 As String, strWd2 As String
Dim strWhere As String, strSet As String
Dim bolTrans As Boolean

   Pub_BatchDay113Proc = False
'-----檢查造字欄位表是否有重複欄位
   strA1 = "select * from EditOraInEudc order by 1 "
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strA1)
   If intA = 1 Then
       rsAD.MoveFirst
       Do While Not rsAD.EOF
           strShow = "": strUpd = ""
           tmpShow = Empty: tmpUpd = Empty
           If "" & rsAD.Fields("showlist") <> "" Then
               tmpShow = Split("" & rsAD.Fields("showlist"), ",")
               For intA = 0 To UBound(tmpShow)
                   If Trim(tmpShow(intA)) <> "" Then
                      If strShow = "" Then
                          strShow = "," & tmpShow(intA)
                      Else
                          If InStr(strShow, tmpShow(intA)) = 0 Then
                              strShow = strShow & "," & tmpShow(intA)
                          End If
                      End If
                   End If
               Next intA
               strShow = Mid(strShow, 2)
           End If
           If "" & rsAD.Fields("updlist") <> "" Then
               tmpUpd = Split("" & rsAD.Fields("updlist"), ",")
               For intA = 0 To UBound(tmpUpd)
                   If Trim(tmpUpd(intA)) <> "" Then
                      If strUpd = "" Then
                          strUpd = "," & tmpUpd(intA)
                      Else
                          If InStr(strUpd, tmpUpd(intA)) = 0 Then
                              strUpd = strUpd & "," & tmpUpd(intA)
                          End If
                      End If
                   End If
               Next intA
               strUpd = Mid(strUpd, 2)
           End If
           If strShow <> "" & rsAD.Fields("showlist") Or strUpd <> rsAD.Fields("updlist") Then
               cnnConnection.Execute strA1
           End If
           rsAD.MoveNext
       Loop
   End If
   
On Error GoTo ErrorHandle
'整批直接更新造字
    strA1 = "select * from editeudclog where eel01=9999 order by 1,2 "
    intA = 1
    Set rsAD = ClsLawReadRstMsg(intA, strA1)
    If intA = 1 Then
         rsAD.MoveFirst
         Do While Not rsAD.EOF
            If strBeginTime = Format(ServerTime, "000000") Then
                Sleep 1000
            End If
            strBeginTime = Format(ServerTime, "000000")
            bolTrans = False
            lngCnt = 0: lngTot = 0
            strWd1 = "" & rsAD.Fields("EEL04") '更改前
            strWd2 = "" & rsAD.Fields("EEL06") '更改後
            cnnConnection.BeginTrans
            If "" & rsAD.Fields("eel09") <> "" Then  '不更新的記錄
                strSql = "Update editeudclog set eel01=" & strSrvDate(1) & ", eel02=" & CNULL(strBeginTime, True) & ", eel08=" & CNULL(Format(ServerTime, "000000"), True) & ", eel09='" & IIf(pType = "2", "直接更新:", "每日批次:") & "不更新;'||eel09 " & _
                            "where eel01=" & rsAD.Fields("eel01") & "  and eel02=" & rsAD.Fields("eel02")
                cnnConnection.Execute strSql
                If bolTrans = False Then bolTrans = True
            Else
                strA1 = "select * from EditOraInEudc order by sno asc "
                intI = 1
                Set RsTemp = ClsLawReadRstMsg(intI, strA1)
                If intI = 1 Then
                    RsTemp.MoveFirst
                    Do While Not RsTemp.EOF
                        If "" & RsTemp.Fields("tbname") <> "" And "" & RsTemp.Fields("updlist") <> "" Then
                            strWhere = "": strSet = ""
                            strTBname = RsTemp.Fields("tbname")
                            strA1 = "" & RsTemp.Fields("updlist")
                            tmpUpd = Empty
                            tmpUpd = Split(strA1, ",")
                            For intQ = 0 To UBound(tmpUpd)
                                '更改前的文字
                                strSet = strSet & ", " & Trim(tmpUpd(intQ)) & "=Replace(" & Trim(tmpUpd(intQ)) & ",'" & strWd1 & "' "
                                strWhere = strWhere & "or instr(" & tmpUpd(intQ) & ",'" & strWd1 & "')>0 "
                                '更改後的文字
                                strSet = strSet & ",'" & strWd2 & "') "
                            Next intQ
                            If strSet <> "" Then
                               strConList = "Update " & RsTemp.Fields("tbname") & " Set " & Mid(strSet, 2) & " where (" & Mid(strWhere, 4) & ") "
                               cnnConnection.Execute strConList, lngCnt
                               lngTot = lngTot + lngCnt
                               If lngTot > 0 And bolTrans = False Then bolTrans = True
                            End If
                        End If
                        RsTemp.MoveNext
                    Loop
                End If
                strSql = "Update editeudclog set eel01=" & strSrvDate(1) & ", eel02=" & CNULL(strBeginTime, True) & ", eel08=" & CNULL(Format(ServerTime, "000000"), True) & ", eel09='" & IIf(pType = "2", "直接更新:", "每日批次:") & "更新" & lngTot & "筆;'||eel09 " & _
                            "where eel01=" & rsAD.Fields("eel01") & "  and eel02=" & rsAD.Fields("eel02")
                cnnConnection.Execute strSql
                If bolTrans = False Then bolTrans = True
            End If
            cnnConnection.CommitTrans
            rsAD.MoveNext
         Loop 'Do While Not rsAD.EOF
         If pType = "2" Then MsgBox "更新完成!!", vbInformation
    End If
    Pub_BatchDay113Proc = True
    Exit Function
    
ErrorHandle:
   If Err.Number <> 0 Then
       If bolTrans = True Then
           cnnConnection.RollbackTrans
       End If
       If pType = "2" Then
           MsgBox "整批直接更新造字失敗：" & strWd1 & "->" & strWd2 & ";" & Err.Description
       ElseIf pType = "3" Then
          WLog Err.Description
       End If
   End If
   Set rsAD = Nothing
End Function

'Add By Sindy 2022/5/25 設定屬智權人員作業的下拉選單(共用模組)
'Modified by Lydia 2022/05/27 傳入人員編號pUserNo;因為frm090801接洽單有業助幫智權人員代填的需求
'Modify by Amy 2023/02/10 +stFormN 表單
Public Sub PUB_SetCombo1Sales(ByRef oCombo As Object, Optional ByVal pUserNo As String, Optional stFromN As String = "")
   Dim ii As Integer, jj As Integer
   Dim stDef As String 'Add by Amy 2019/11/05 預設
   Dim strQ As String, strMCTF As String 'Add by Amy 2020/02/19
   Dim strTemp As String, arrData As Variant
   
   '*** Memo 若有增加表單請於下方列示,方便知道修改影響之程式 ***
   Select Case UCase(strFormName)
        Case "FRM090127" '查覆區
        Case "FRM090202_3" '專利／商標會稿
        Case "FRM210144", "FRM210145" '寄發文件,寄件查詢
   End Select
   
   If pUserNo = "" Then pUserNo = strUserNum 'Added by Lydia 2022/05/27
   
   oCombo.Clear
   'Modified by Lydia 2022/05/27
   'oCombo.AddItem strUserNum & " " & strUserName
   oCombo.AddItem pUserNo & " " & GetStaffName(pUserNo, True)
   
   '檢查當時是否需要為他人職代
   'Mofified by Lydia 2022/05/27 strUserNum=> pUserNo
   Call Pub_SetForOthersEmpCombo(pUserNo, oCombo, False)
   
   'Add by Amy 2023/02/10 +stFormN
   'Modify By Sindy 2014/9/4 帶人的權限
   'Mofified by Lydia 2022/05/27 strUserNum=> pUserNo
   Call Pub_SetSAManageEmpCombo(pUserNo, oCombo, False, , True, , stFromN)
   '2014/9/4 END
   'Modified by Lydia 2020/06/08 +增加特殊權限"AREA"
   'Mofified by Lydia 2022/05/27 strUserNum=> pUserNo
   Call Pub_SetSAManageEmpCombo(pUserNo, oCombo, False, , , "AREA", stFromN)
   'end 2023/02/10
   
   'Added by Morgan 2014/5/15
   '專利處智權同仁代處理人
   'Modify by Amy 2015/03/13 +特殊設定(總經理業務工作代理人員)
   'Mofified by Lydia 2022/05/27 strUserNum=> pUserNo
   If InStr(Pub_GetSpecMan("A8"), pUserNo) > 0 Or InStr(Pub_GetSpecMan("總經理業務工作代理人員"), pUserNo) > 0 Then
      If InStr(Pub_GetSpecMan("A8"), pUserNo) > 0 Then
        strSql = "select st01,st02 from setSpecMan,staff where ocode='A7' and instr(';'||replace(oMan,',',';')||';',';'||st01||';')>0"
      Else
        strSql = "select st01,st02 from setSpecMan,staff where ocode='總經理員工編號' and instr(';'||replace(oMan,',',';')||';',';'||st01||';')>0"
      End If
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            For ii = 0 To oCombo.ListCount - 1
               If InStr(oCombo.List(ii), RsTemp(0)) = 1 Then
                  Exit For
               End If
            Next
            If ii = oCombo.ListCount Then
               oCombo.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
            End If
            
            RsTemp.MoveNext
         Loop
      End If
   End If
   'end 2014/5/15
   
   'Add By Sindy 2014/5/29 開放部份智權同仁的資料給彥葶操作
   'If Pub_GetSpecMan("A8") = strUserNum Then
   'Mofified by Lydia 2022/05/27 strUserNum=> pUserNo
   If InStr(Pub_GetSpecMan("A8"), pUserNo) > 0 Then
      strTemp = Pub_GetSpecMan("A7")
      arrData = Split(strTemp, ";")
      For jj = 0 To UBound(arrData)
         For ii = 0 To oCombo.ListCount - 1
            If InStr(oCombo.List(ii), arrData(jj)) = 1 Then
               Exit For
            End If
         Next ii
         If ii = oCombo.ListCount Then
            oCombo.AddItem arrData(jj) & " " & GetPrjSalesNM(CStr(arrData(jj)))
         End If
     Next jj
   End If
   '2014/5/29 END
      
   'Add By Sindy 2022/5/25 顧服組人員可以操作 W2001
   'Mofified by Lydia 2022/05/27 strUserNum=> pUserNo
   If Left(PUB_GetST03(pUserNo), 1) = "W" Then
      oCombo.AddItem "W2001" & " " & GetPrjSalesNM("W2001")
   End If
   '2022/5/25 END
   
   '帶人主管抓虛建編號
   'Mofified by Lydia 2022/05/27 strUserNum=> pUserNo
   strSql = "select st01,st02 from staff where st01<'63001' and instr(';'||st52||';'||st53||';'||st54||';'||st55||';',';" & pUserNo & ";')>0"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      RsTemp.MoveFirst
      Do While Not RsTemp.EOF
         For ii = 0 To oCombo.ListCount - 1
            If InStr(oCombo.List(ii), RsTemp(0)) = 1 Then
               Exit For
            End If
         Next
         If ii = oCombo.ListCount Then
            oCombo.AddItem RsTemp.Fields("st01") & " " & RsTemp.Fields("st02")
         End If
         RsTemp.MoveNext
      Loop
   End If
   
   oCombo.ListIndex = 0
   'Modify By Sindy 2020/11/20 有預設值就帶
   'If Pub_StrUserSt03 = "P22" Then oCombo = stDef 'Add by Amy 2019/11/05 內商程序人預設 P2002
   If stDef <> "" Then oCombo = stDef 'Add by Amy 2019/11/05 內商程序人預設 P2002
End Sub

'Move by Lydia 2022/09/06 從basQuery搬過來
'cp(1),cp(2),cp(3),cp(4) 本所案號
'cp(5) 案件性質代號
'cp(6) 案件性質名稱
'cp(7) 收文日(民國年)
'cp(8) 總收文號
'911118 nick 新增原來申請人
'cp(9)  原來申請人
Public Function ClsLawChkSameCase(ByRef cp() As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
 Dim strID As String, strName As String
 Dim strTmp As String
 Dim i As Integer
 Dim strQty As String
 
On Error GoTo ErrHand
   ClsLawChkSameCase = True
   strQty = "SELECT COUNT(*) FROM CASEPROGRESS WHERE CP01='" & cp(1) & "' AND " & _
      "CP02='" & cp(2) & "' AND CP03='" & cp(3) & "' AND CP04='" & cp(4) & "' AND " & _
      "CP05='" & TransDate(cp(7), 2) & "' AND CP10='" & cp(5) & "'"
   'add by nickc 2007/04/18
   RsTemp.CursorLocation = adUseClient
   
   RsTemp.Open strQty, cnnConnection
   
   If RsTemp.Fields(0) > 0 Then
      MsgBox "該案件當日相同案件性質有一筆以上，請自行確認 !" & vbCrLf & _
         "本所案號 : " & cp(1) & cp(2) & cp(3) & cp(4) & vbCrLf & _
         "案件性質 : " & cp(6) & vbCrLf & _
         "收文日 : " & cp(7) & vbCrLf, vbInformation
   End If
   RsTemp.Close
   
   'Added by Morgan 2014/3/6
   '有收據/請款單才要繼續
   strQty = "SELECT CP60 FROM CASEPROGRESS WHERE CP09='" & cp(8) & "' and cp60 is not null"
   RsTemp.CursorLocation = adUseClient
   RsTemp.Open strQty, cnnConnection, adOpenForwardOnly, adLockReadOnly
   If RsTemp.EOF And RsTemp.BOF Then
      RsTemp.Close
      Exit Function
   End If
   RsTemp.Close
   'end 2014/3/6
   
   strQty = "select sk02 from systemkind where sk01='" & cp(1) & "'"
   'add by nickc 2007/04/18
   RsTemp.CursorLocation = adUseClient
   
   RsTemp.Open strQty, cnnConnection
   i = RsTemp.Fields(0)
   RsTemp.Close
   
   Select Case i
      Case 1
         strQty = "SELECT PA26 FROM PATENT WHERE PA01='" & cp(1) & "' AND " & _
            "PA02='" & cp(2) & "' AND PA03='" & cp(3) & "' AND PA04='" & cp(4) & "'"
      Case 2
         strQty = "SELECT TM23 FROM TRADEMARK WHERE TM01='" & cp(1) & "' AND " & _
            "TM02='" & cp(2) & "' AND TM03='" & cp(3) & "' AND TM04='" & cp(4) & "'"
      Case 3
         strQty = "SELECT LC11 FROM LAWCASE WHERE LC01='" & cp(1) & "' AND " & _
            "LC02='" & cp(2) & "' AND LC03='" & cp(3) & "' AND LC04='" & cp(4) & "'"
      Case 4
         strQty = "SELECT HC05 FROM HIRECASE WHERE HC01='" & cp(1) & "' AND " & _
            "HC02='" & cp(2) & "' AND HC03='" & cp(3) & "' AND HC04='" & cp(4) & "'"
      Case Else
         strQty = "SELECT SP08 FROM SERVICEPRACTICE WHERE SP01='" & cp(1) & "' AND " & _
            "SP02='" & cp(2) & "' AND SP03='" & cp(3) & "' AND SP04='" & cp(4) & "'"
   End Select
      
   strID = ""
   'add by nickc 2007/04/18
   RsTemp.CursorLocation = adUseClient
   
   RsTemp.Open strQty, cnnConnection
   If Not RsTemp.EOF And Not RsTemp.BOF Then
      If Not IsNull(RsTemp.Fields(0)) Then
         strID = RsTemp.Fields(0)
      End If
      RsTemp.Close
   Else
      RsTemp.Close
      
      strQty = "SELECT CP01,CP02,CP03,CP04 FROM CASEPROGRESS WHERE CP09='" & cp(8) & "'"
      'add by nickc 2007/04/18
      RsTemp.CursorLocation = adUseClient
      
      RsTemp.Open strQty, cnnConnection
      
      Select Case i
         Case 1
            strQty = "SELECT PA26 FROM PATENT WHERE PA01='" & RsTemp.Fields(0) & "' AND " & _
               "PA02='" & RsTemp.Fields(1) & "' AND PA03='" & RsTemp.Fields(2) & "' AND PA04='" & RsTemp.Fields(3) & "'"
         Case 2
            strQty = "SELECT TM23 FROM TRADEMARK WHERE TM01='" & RsTemp.Fields(0) & "' AND " & _
               "TM02='" & RsTemp.Fields(1) & "' AND TM03='" & RsTemp.Fields(2) & "' AND TM04='" & RsTemp.Fields(3) & "'"
         Case 3
            strQty = "SELECT LC11 FROM LAWCASE WHERE LC01='" & RsTemp.Fields(0) & "' AND " & _
               "LC02='" & RsTemp.Fields(1) & "' AND LC03='" & RsTemp.Fields(2) & "' AND LC04='" & RsTemp.Fields(3) & "'"
         Case 4
            strQty = "SELECT HC05 FROM HIRECASE WHERE HC01='" & RsTemp.Fields(0) & "' AND " & _
               "HC02='" & RsTemp.Fields(1) & "' AND HC03='" & RsTemp.Fields(2) & "' AND HC04='" & RsTemp.Fields(3) & "'"
         Case Else
            strQty = "SELECT SP08 FROM SERVICEPRACTICE WHERE SP01='" & RsTemp.Fields(0) & "' AND " & _
               "SP02='" & RsTemp.Fields(1) & "' AND SP03='" & RsTemp.Fields(2) & "' AND SP04='" & RsTemp.Fields(3) & "'"
      End Select
      
      RsTemp.Close
      
      strID = ""
      'add by nickc 2007/04/18
      RsTemp.CursorLocation = adUseClient
      
      RsTemp.Open strQty, cnnConnection
      If Not RsTemp.EOF And Not RsTemp.BOF Then
         If Not IsNull(RsTemp.Fields(0)) Then
            strID = RsTemp.Fields(0)
         End If
         RsTemp.Close
      End If
   End If
   
'Modified by Morgan 2014/3/4 改判斷申請人必須相同,收據客戶編號不用更新
If strID <> ChangeCustomerL(cp(9)) Then
   MsgBox "申請人不同，不可轉本所案號 ! !", vbInformation
   ClsLawChkSameCase = False
End If

'   Dim strTmp1 As String
''edit by nickc 2008/01/18 因為下面已經有做了，所以秀玲說不用
''   If ClsLawGetCusCAJnam(strID, strName, strTmp, strTmp1) Then
''
''   End If
'
'   strTmp1 = ""
'   strQty = "SELECT CP60 FROM CASEPROGRESS WHERE CP09='" & cp(8) & "'"
'   'add by nickc 2007/04/18
'   RsTemp.CursorLocation = adUseClient
'
'   RsTemp.Open strQty, cnnConnection
'   If Not RsTemp.EOF And Not RsTemp.BOF Then
'      If Not IsNull(RsTemp.Fields(0)) Then
'
'         strTmp1 = RsTemp.Fields(0)
'         RsTemp.Close
'
'         '911118 nick 加入  若同一收據，但申請人不同，有兩筆以上時，不能修改
'         'StrQty = "SELECT COUNT(DISTINCT A0J02) FROM ACC0J0 WHERE A0J13='" & strTmp1 & "'"
'         strQty = "SELECT COUNT(DISTINCT A0J02) FROM ACC0J0 WHERE A0J13='" & strTmp1 & "' and a0j11<>'" & cp(9) & "' "
'         'add by nickc 2007/04/18
'         RsTemp.CursorLocation = adUseClient
'
'         RsTemp.Open strQty, cnnConnection
'         If RsTemp.Fields(0) > 1 Then
'            MsgBox "同一收據有其他案號收文資料，不可轉本所案號 !", vbInformation
'            ClsLawChkSameCase = False
'         Else
'
'            'Removed by Morgan 2012/8/31 只要更新該收文號,改在存檔時用 PUB_UpdateAccData
'            'strQty = "UPDATE ACC0J0 SET A0J02='" & cp(1) & cp(2) & cp(3) & cp(4) & "',A0J11=" & CNULL(strID) & " WHERE A0J13=" & CNULL(strTmp1)
'            'cnnConnection.Execute strQty
'            'end 2012/8/31
'
'            '92.3.8 MODIFY BY SONIA 不改收據抬頭
'            'StrQty = "UPDATE ACC0K0 SET A0K03=" & CNULL(strID) & ",A0K04=" & CNULL(strName) & " WHERE A0K01=" & CNULL(strTmp1)
'            strQty = "UPDATE ACC0K0 SET A0K03=" & CNULL(strID) & " WHERE A0K01=" & CNULL(strTmp1)
'            '92.3.8 END
'            cnnConnection.Execute strQty
'         End If
'         RsTemp.Close
'      End If
'   End If
'end 2014/3/4
   
   Exit Function
ErrHand:
   ClsLawChkSameCase = False
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'Move by Lydia 2022/09/06 從basQuery搬過來
Public Function ClsLawChkMRec(mr02 As String, strlc As String, mr16 As String, mr17 As String) As Boolean
 Dim RsTemp As New ADODB.Recordset
 Dim strQty As String
On Error GoTo ErrHand
   strQty = "select mr16,mr17 from mailrec where " & ChgMailRec(strlc) & " AND mr02=" + CNULL(mr02)
   ClsLawChkMRec = False
   mr16 = ""
   mr17 = ""
   'add by nickc 2007/04/18
   RsTemp.CursorLocation = adUseClient
   
   RsTemp.Open strQty, cnnConnection
   Do While Not RsTemp.EOF
      If IsNull(RsTemp.Fields(0)) = False Then mr16 = RsTemp.Fields(0)
      If IsNull(RsTemp.Fields(1)) = False Then mr17 = RsTemp.Fields(1)
      ClsLawChkMRec = True
      Exit Do
   Loop
   RsTemp.Close
   Exit Function
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'Move by Lydia 2022/09/06 從basQuery搬過來
Public Function ClsLawGetCaseFee(ByRef strKind As String, ByRef strRelNation As String, ByRef strProperty As String, ByRef dblDay As Double) As Boolean
 Dim RsTemp As New ADODB.Recordset, strQty As String
   dblDay = 0
   strQty = "select cf04 from casefee where cf01='" + strKind + "' and cf02='" + strRelNation + "' and cf03='" + strProperty + "'"
   ClsLawGetCaseFee = False
   'add by nickc 2007/04/18
   RsTemp.CursorLocation = adUseClient
   
   RsTemp.Open strQty, cnnConnection
   Do While Not RsTemp.EOF
      If Not IsNull(RsTemp.Fields(0)) Then dblDay = RsTemp.Fields(0)
      ClsLawGetCaseFee = True
      Exit Do
   Loop
   RsTemp.Close
   Exit Function
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'Move by Lydia 2022/09/06 從basQuery搬過來
'傳進西元年，傳出西元年
Public Function ClsLawGetCaseFeeDelay(ByVal CF01 As String, ByVal CF02 As String, ByVal CF03 As String, ByRef CFOther() As String, Optional p_CaseNo As String, Optional p_iDays As Integer, Optional p_iMonths As Integer) As Boolean
 Dim RsTemp As New ADODB.Recordset
 Dim stSQL As String, intR As Integer
 Dim stCF01 As String
 Dim iDays As Integer 'Added by Morgan 2019/7/11
 
On Error GoTo ErrHand
   ClsLawGetCaseFeeDelay = True
   stCF01 = CF01
   'Add by Morgan 2008/1/7 FCP的申復延期要判斷若申請人均為本國人抓國內案設定
   'Modify by Morgan 2008/5/30 +204,206也要
   If CF01 = "FCP" And CF02 = "000" And (CF03 = "205" Or CF03 = "204" Or CF03 = "206") And p_CaseNo <> "" Then
      If PUB_ExistForeigner(p_CaseNo) = False Then
         stCF01 = "P"
      End If
   End If
   'end 2008/1/7
   
   CFOther(1) = CFOther(0) '法定
   CFOther(2) = CFOther(0) '本所
   CFOther(3) = CFOther(0) '約定 Add By Sindy 2021/5/7
   stSQL = "SELECT CF22,CF25,CF27 FROM CASEFEE WHERE CF01='" & stCF01 & "' AND CF02='" & CF02 & "' AND CF03='" & CF03 & "'"
   'add by nickc 2007/04/18
   RsTemp.CursorLocation = adUseClient
   
   RsTemp.Open stSQL, cnnConnection
   Do While Not RsTemp.EOF
      If IsNull(RsTemp.Fields(0)) Or RsTemp.Fields(0) = 0 Then
         If Not IsNull(RsTemp.Fields(1)) Then
            p_iDays = RsTemp(1) * 30 'Add by Morgan 2008/5/30
            p_iMonths = RsTemp(1) 'Add by Morgan 2008/9/5
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
                     iDays = 4 'Added by Morgan 2019/7/11
                  Else
                     CFOther(2) = CompDate(2, -2, CFOther(1))
                     iDays = 2 'Added by Morgan 2019/7/11
                  End If
            End Select
            
         End If
      Else
         If Not IsNull(RsTemp.Fields(0)) Then
            p_iDays = RsTemp(0) 'Add by Morgan 2008/5/30
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
                     iDays = 4 'Added by Morgan 2019/7/11
                  Else
                     CFOther(2) = CompDate(2, -2, CFOther(1))
                     iDays = 2 'Added by Morgan 2019/7/11
                  End If
            End Select
         End If
      End If
      
      'Added by Morgan 2014/10/28
      'Modified by Morgan 2014/11/20 外專改回舊規則
      If CF02 = 台灣國家代號 And strSrvDate(1) >= 台灣案所限新規則啟用日 And CF01 <> "FCP" And CF01 <> "FG" Then
         CFOther(2) = PUB_GetOurDeadline(CFOther(1))
         
      'Added by Morgan 2019/7/11 外專台灣案所限以改工作天計算
      ElseIf strSrvDate(1) >= 外專台灣案所限新規則啟用日 And (CF01 = "FCP" Or CF01 = "FG") And iDays > 0 Then
         'Add By Sindy 2021/5/7 +, , CFOther(3)
         CFOther(2) = PUB_GetFCPOurDeadline(CFOther(1), iDays, , CFOther(3))
      
      'end 2019/7/11
      End If
      'end 2014/10/28
      
      ClsLawGetCaseFeeDelay = True
      Exit Do
   Loop
   RsTemp.Close
   Exit Function
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbCritical
   Resume
End Function

'Move by Lydia 2022/09/06 從basQuery搬過來
'add by nick 2004/09/27
'傳入客戶編號
'回傳公司負責人英文名稱
'Mark by Lydia 2024/07/03 改成傳入變數
'Public Sub GetCu103ByCustomer(oForm As Form, ByVal oCu As String)
'CheckOC3
'oCu = oCu & "00000000"
'strSql = "SELECT * FROM Customer " & _
'         "WHERE CU01 = '" & Mid(oCu, 1, 8) & "' AND " & _
'               "CU02 = '" & Mid(oCu, 9, 1) & "'"
'   AdoRecordSet3.CursorLocation = adUseClient
'   AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
'   If AdoRecordSet3.RecordCount > 0 Then
'      AdoRecordSet3.MoveFirst
'      oForm.m_CU103 = CheckStr(AdoRecordSet3.Fields("CU103").Value)
'      oForm.m_CU05 = CheckStr(AdoRecordSet3.Fields("CU05").Value)
'      oForm.m_CU88 = CheckStr(AdoRecordSet3.Fields("CU88").Value)
'      oForm.m_CU89 = CheckStr(AdoRecordSet3.Fields("CU89").Value)
'      oForm.m_CU90 = CheckStr(AdoRecordSet3.Fields("CU90").Value)
'      'add by nickc 2006/01/20
'      oForm.m_CU112 = CheckStr(AdoRecordSet3.Fields("CU112").Value)
'    'add by nickc 2007/08/10 多申請人時，才不會衝突
'      'Add By Sindy 2012/2/8
'      oForm.m_CU39 = CheckStr(AdoRecordSet3.Fields("CU39").Value)
'      oForm.m_CU40 = CheckStr(AdoRecordSet3.Fields("CU40").Value)
'      oForm.m_CU41 = CheckStr(AdoRecordSet3.Fields("CU41").Value)
'      '2012/2/8 End
'      'Add By Sindy 2012/10/31
'      oForm.m_CU10 = CheckStr(AdoRecordSet3.Fields("CU10").Value)
'      '2012/10/31 End
'    Else
'        oForm.m_CU103 = ""
'        oForm.m_CU05 = ""
'        oForm.m_CU88 = ""
'        oForm.m_CU89 = ""
'        oForm.m_CU90 = ""
'        oForm.m_CU112 = ""
'        'Add By Sindy 2012/2/8
'        oForm.m_CU39 = ""
'        oForm.m_CU40 = ""
'        oForm.m_CU41 = ""
'        '2012/2/8 End
'        'Add By Sindy 2012/10/31
'        oForm.m_CU10 = ""
'        '2012/10/31 End
'    End If
'CheckOC3
'End Sub

'Added by Lydia 2024/07/03 取得內商發文->申請人及公司負責人輸入
Public Sub Pub_GetDataFrm020102(ByVal oCu As String, ByRef pCU103 As String, ByRef pCU05 As String, ByRef pCU88 As String, ByRef pCU89 As String, ByRef pCU90 As String _
            , ByRef pCU112 As String, ByRef pCU39 As String, ByRef pCU40 As String, ByRef pCU41 As String, ByRef pCU10 As String)
Dim intQ As Integer, strQuery As String
Dim rsQuery As New ADODB.Recordset
   
   oCu = ChangeCustomerL(oCu)
   pCU112 = ""
   pCU103 = ""
   pCU05 = ""
   pCU88 = ""
   pCU89 = ""
   pCU90 = ""
   pCU39 = ""
   pCU40 = ""
   pCU41 = ""
   pCU10 = ""
   strQuery = "SELECT * FROM Customer WHERE CU01='" & Mid(oCu, 1, 8) & "' AND CU02='" & Mid(oCu, 9, 1) & "' "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      pCU112 = "" & rsQuery.Fields("CU112")
      pCU103 = "" & rsQuery.Fields("CU103")
      pCU05 = "" & rsQuery.Fields("CU05")
      pCU88 = "" & rsQuery.Fields("CU88")
      pCU89 = "" & rsQuery.Fields("CU89")
      pCU90 = "" & rsQuery.Fields("CU90")
      pCU39 = "" & rsQuery.Fields("CU39")
      pCU40 = "" & rsQuery.Fields("CU40")
      pCU41 = "" & rsQuery.Fields("CU41")
      pCU10 = "" & rsQuery.Fields("CU10")
   End If
   Set rsQuery = Nothing
End Sub

'Added by Morgan 2022/12/7
'檢查客戶選擇的證書/註冊證形式是否與最後收文設定相同
'pCustNo=客戶編號,pCP01=系統別, pType=證書/註冊證形式, pShowMsg=是否彈訊息
Public Function PUB_ChkCustCertType(ByVal pCustNo As String, ByVal pCP01 As String, ByVal pType As String, Optional ByVal pShowMsg As Boolean = False) As Boolean
   Dim strQ As String, intQ As Integer
   Dim rstQ As ADODB.Recordset
   Dim strTypName As String, strTypName2 As String, strCertName As String, strCaseType As String
   
   If pCustNo = "" Or pCP01 = "" Or (pType <> "1" And pType <> "2") Then
      MsgBox "參數錯誤！", vbCritical
      Exit Function
   End If
   
   PUB_ChkCustCertType = True
   
   pCustNo = ChangeCustomerL(pCustNo)
   strQ = ""
   If pCP01 = "P" Then
      strCertName = "證書"
      strCaseType = "台灣專利案"
      strQ = "select pa178 TYP,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNO" & _
         " from customer,patent,caseprogress" & _
         " where cu01='" & Mid(pCustNo, 1, 8) & "' and pa26(+)=cu01||cu02 and pa09='000' and pa178 is not null" & _
         " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04 and cp10='601'" & _
         " order by cp05 desc,cp09 desc"
   ElseIf pCP01 = "T" Then
      strCertName = "註冊證"
      strCaseType = "台灣商標案"
      strQ = "select tm136 TYP,cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) CNO" & _
         " from customer,trademark,caseprogress" & _
         " where cu01='" & Mid(pCustNo, 1, 8) & "' and tm23(+)=cu01||cu02 and tm10='000' and tm136 is not null" & _
         " and cp01(+)=tm01 and cp02(+)=tm02 and cp03(+)=tm03 and cp04(+)=tm04 and cp10='717'" & _
         " order by cp05 desc,cp09 desc"
   End If
   If strQ <> "" Then
      intQ = 1
      Set rstQ = ClsLawReadRstMsg(intQ, strQ)
      If intQ = 1 Then
         If rstQ(0) <> pType Then
            PUB_ChkCustCertType = False
            If pShowMsg Then
               If pType = "1" Then
                  strTypName = "電子"
                  strTypName2 = "紙本"
               Else
                  strTypName = "紙本"
                  strTypName2 = "電子"
               End If
               MsgBox "本次" & strCaseType & "選擇的" & strCertName & "形式為【" & strTypName & "】與該客戶前次收文的選擇不同，請留意！" & vbCrLf & vbCrLf & "前次收文案號：" & rstQ(1) & "【" & strTypName2 & "】", vbExclamation
            End If
         End If
      End If
   End If
End Function

'Added by Lydia 2020/10/07 法律所案源收文：用於案源之接洽人取得在職員工編號和介紹人第一人
'Move by Lydia 2022/09/08 從basPublic搬過來
'Move by Lydia 2023/02/14 從Service1搬過來
Public Function PUB_GetNowStaff(ByVal pStLst As String, Optional ByRef oFstNo As String) As String
'oFstNo : 介紹人第一人
Dim intQ As Integer, rsQ1 As New ADODB.Recordset
Dim strQ As String, strMidNo As String
Dim tmpArr1 As Variant, intA As Integer
    oFstNo = ""
    PUB_GetNowStaff = ""
    
    If pStLst <> "" Then
        tmpArr1 = Split(pStLst, ",")
        For intA = 0 To UBound(tmpArr1)
            If Trim("" & tmpArr1(intA)) <> "" Then
                'Modified by Lydia 2022/05/09 介紹人若離職，則改抓其ST15之A0908
                'strQ = "select st01 from staff where st04='1' and st01=" & CNULL(Trim("" & tmpArr1(intA)))
                strQ = "select st01,st04,a0908 from staff,acc090 where st01=" & CNULL(Trim("" & tmpArr1(intA))) & " and st15=a0901(+) "
                intQ = 1
                Set rsQ1 = ClsLawReadRstMsg(intQ, strQ)
                If intQ = 1 Then
                    If "" & rsQ1.Fields("st04") = "1" Then  'Added by Lydia 2022/05/09 判斷是否離職
                        If oFstNo = "" Then oFstNo = Trim("" & tmpArr1(intA))
                        strMidNo = strMidNo & "," & Trim("" & tmpArr1(intA))
                    'Added by Lydia 2022/05/09 介紹人若離職，則改抓其ST15之A0908
                    Else
                        If oFstNo = "" Then oFstNo = "" & rsQ1.Fields("a0908")
                        strMidNo = strMidNo & "," & rsQ1.Fields("a0908")
                    End If
                    'end 2022/05/09
                End If
            End If
        Next intA
        If strMidNo <> "" Then
            PUB_GetNowStaff = Mid(strMidNo, 2)
        End If
    End If
End Function

Public Function PUB_GetTCTmail(ByVal bolCache As Boolean, ByVal iSta As Integer, ByVal iCase01 As String, ByVal iCase02 As String, ByVal iCase03 As String, ByVal iCase04 As String, ByVal iNo As String, Optional ByVal iMang As String, _
                   Optional ByVal iNto As String = "", Optional ByVal iChgGrp As String = "", Optional ByVal iSubStart As String = "", Optional ByVal inContext As String = "", Optional ByVal inCC As String = "") As Boolean
'bolCache 是否存mailcache
'iSta     0：櫃台收新案=新案立卷(Added by Lydia 2023/02/17)
           '1：分案通知
           '2：改組別通知
           '3：刪除(收文)通知
           '4：命名完成通知email(工程師主管確認)
           'Memo by Lydia 2021/04/6 原本「4:通知上傳檔案(.msg) 5.通知上傳檔案(外文原文本)」檢查後確定不再使用
'end 2023/02/17
'iCase01~04 本所案號
'iNo    　收文號
'iMang　　 分案-工程師主管: 只有櫃台收新案會特別傳入B
'iNto     通知命名人員
Dim rsR1 As New ADODB.Recordset
Dim rsB1 As New ADODB.Recordset
Dim intR As Integer
Dim Str01 As String, Str02 As String
Dim strSub As String, strTo  As String
Dim strCont As String
Dim stCC As String 'Added by Lydia 2018/03/07 程序和承辦主管改在副本
Dim strTempA As String, strNA16 As String 'Added by Lydia 2021/04/16
Dim intCnt As Integer 'Added by Lydia 2021/10/19
Dim m_CP13 As String 'Added by Lydia 2022/08/19
Dim m_TCT01 As String 'Added by Lydia 2023/02/21
Dim strCP142 As String, strCP164 As String 'Added by Lydia 2023/4/21

On Error GoTo ErrHandle
    
    'Modified by Lydia 2019/10/03 +提申日
    'Str01 = "SELECT 1 ord1,TCT01,TCT02,TCT03,TCT04,TCT07,TCT10,CP06,CP07 FROM TransCaseTitle,CASEPROGRESS WHERE TCT01='" & iNo & "' AND TCT01=CP09(+) "
    'Str01 = Str01 & "union all SELECT 2 ord1,TCT01,TCT02,TCT03,TCT04,TCT07,TCT10,CP06,CP07 FROM CASEPROGRESS,TransCaseTitle " & _
            "WHERE CP01='" & iCase01 & "' AND CP02='" & iCase02 & "' AND CP03='" & iCase03 & "' AND CP04='" & iCase04 & "' " & _
            "AND CP10 IN (" & NewCasePtyList & ") AND CP158=0 AND CP159=0 AND CP09=TCT01(+) "
    'Modified by Lydia 2021/04/16 +CP158,PA150,TCT20,TCT117
    'Modified by Lydia 2021/04/29 +CP10
    'Modifeid by Lydia 2021/05/05 +CP01~CP04
    'Modified by Lydia 2022/08/19 +CP13
    'Modified Lydia 2022/12/07 抓代理人名稱,申請人名稱
    'Str01 = "SELECT 1 ord1,TCT01,TCT02,TCT03,TCT04,TCT07,TCT10,CP06,CP07,PA10,CP158,PA150,TCT20,TCT117,CP01,CP02,CP03,CP04,CP10,CP13 " & _
                "FROM TransCaseTitle,CASEPROGRESS,PATENT WHERE TCT01='" & iNo & "' AND TCT01=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) "
    'Str01 = Str01 & "union all SELECT 2 ord1,TCT01,TCT02,TCT03,TCT04,TCT07,TCT10,CP06,CP07,PA10,CP158,PA150,TCT20,TCT117,CP01,CP02,CP03,CP04,CP10,CP13 " & _
                "FROM CASEPROGRESS,TransCaseTitle,PATENT " & _
                "WHERE CP01='" & iCase01 & "' AND CP02='" & iCase02 & "' AND CP03='" & iCase03 & "' AND CP04='" & iCase04 & "' " & _
                "AND CP10 IN (" & NewCasePtyList & ") AND CP158=0 AND CP159=0 AND CP09=TCT01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) "
    'Modified by Lydia 2023/04/21 +客戶指定送件日期CP142,CP164
    'Modified by Lydia 2024/11/22 +trackingcasename
    Str01 = "SELECT 1 ord1,TCT01,TCT02,TCT03,TCT04,TCT07,TCT10,CP06,CP07,PA10,CP158,PA150,TCT20,TCT117,CP01,CP02,CP03,CP04," & _
                "CP10,CP13,PA75,NVL(FA05,NVL(FA04,FA06)) PA75N,PA26,NVL(CU05,NVL(CU04,CU06)) PA26N,CP142,CP164,tcn13 " & _
                "FROM TransCaseTitle,CASEPROGRESS,PATENT,FAGENT,CUSTOMER,trackingcasename WHERE TCT01='" & iNo & "' AND TCT01=CP09(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
                "AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and tct01=tcn05(+) "
    Str01 = Str01 & "union all SELECT 2 ord1,TCT01,TCT02,TCT03,TCT04,TCT07,TCT10,CP06,CP07,PA10,CP158,PA150,TCT20,TCT117,CP01,CP02,CP03,CP04," & _
                "CP10,CP13,PA75,NVL(FA05,NVL(FA04,FA06)) PA75N,PA26,NVL(CU05,NVL(CU04,CU06)) PA26N,CP142,CP164,tcn13 " & _
                "FROM CASEPROGRESS,TransCaseTitle,PATENT,FAGENT,CUSTOMER,trackingcasename " & _
                "WHERE CP01='" & iCase01 & "' AND CP02='" & iCase02 & "' AND CP03='" & iCase03 & "' AND CP04='" & iCase04 & "' " & _
                "AND CP10 IN (" & NewCasePtyList & ") AND CP158=0 AND CP159=0 AND CP09=TCT01(+) AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) " & _
                "AND SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and tct01=tcn05(+) "
    Str01 = Str01 & " ORDER BY ord1 ASC,TCT01 DESC "
    intR = 1
    Set rsR1 = ClsLawReadRstMsg(intR, Str01)
    If intR = 0 Then
       Exit Function
    Else
      rsR1.MoveFirst
       m_CP13 = "" & rsR1.Fields("CP13") 'Added by Lydia 2022/08/19
       m_TCT01 = "" & rsR1.Fields("TCT01") 'Added by Lydia 2023/02/21
       
       '郵件主旨:(急件！) FCP-00000新案命名(完成期限：本所期限，譯畢期限：急件才顯示)
       'Modified by Lydia 2021/04/16 新案無卷命名email設定：依階段有不同主旨
       'strSub = iCase01 & "-" & iCase02 & IIf(iCase03 & iCase04 <> "000", "-" & iCase03 & "-" & iCase04, "") & "新案命名"
       strSub = "新案命名"
       'Modified by Lydia 2021/05/04 因為P案需要有卷才能作業，所以排除P案
       'Modified by Lydia 2021/05/13 配合防疫措施，英文組使用新案無卷命名email設定
       'If "" & rsR1.Fields("pa150") = "3" And Val("" & rsR1.Fields("cp158")) = 0 And "" & rsR1.Fields("cp01") <> "P" Then
       'Modified by Lydia 2021/05/20 區分櫃台新案立卷
       'If Val("" & rsR1.Fields("cp158")) = 0 And "" & rsR1.Fields("cp01") <> "P" Then
       'Modified by Lydia 2021/10/19 FMP新案皆已無卷命名流程
       'If Val("" & rsR1.Fields("cp158")) = 0 And "" & rsR1.Fields("cp01") <> "P" And iMang <> "B" Then
       If Val("" & rsR1.Fields("cp158")) = 0 And iMang <> "B" Then
           If iSta = 1 Then
               strSub = "進行無卷命名流程"
           ElseIf iSta = 4 Then
               strSub = "命名完成"
           End If
           strCont = strCont & iCase01 & "-" & iCase02 & IIf(iCase03 & iCase04 <> "000", "-" & iCase03 & "-" & iCase04, "") & "【" & strSub & "】" & vbCrLf
           'Added by Lydai 2023/04/21 命名完成通知信增加「指定送件日期」=客戶指定送件日期
           If iSta = "4" And "" & rsR1.Fields("CP142") <> "" Then
               Select Case "" & rsR1.Fields("CP164")
                   Case "1": Str02 = " 當天"
                   Case "2": Str02 = " 之前"
                   Case "3": Str02 = " 之後"
                   Case Else: Str02 = ""
               End Select
               strCont = strCont & "指定送件日期：" & ChangeWStringToTDateString("" & rsR1.Fields("cp142")) & Str02 & vbCrLf
           End If
           'end 2023/04/21
           strCont = strCont & "提申本所期限：" & ChangeWStringToTDateString("" & rsR1.Fields("cp06")) & vbCrLf
           strCont = strCont & "提申法定期限：" & ChangeWStringToTDateString("" & rsR1.Fields("cp07")) & vbCrLf
       End If
       strSub = iCase01 & "-" & iCase02 & IIf(iCase03 & iCase04 <> "000", "-" & iCase03 & "-" & iCase04, "") & strSub
       'end 2021/04/16
       If iSta = 3 Then
          strSub = "已刪除, " & strSub
       'Added by Lydia 2017/12/04 通知上傳檔案(.msg)
       'Modified by Lydia 2017/12/13 +外文原文本 (5)
       'Remove by Lydia 2021/04/16
       'ElseIf iSta = 4 Or iSta = 5 Then
       '   strSub = strSub & "，請到卷宗區上傳" & IIf(iSta = 4, FcpTcnFKey01, FcpTcnFKey02) & "檔"
       '   strCont = "請到卷宗區的收文號" & rsR1.Fields("TCT01") & "上傳檔案 !"
       ''end 2017/12/04
       'end 2021/04/16
       Else
           strSub = IIf(Val("" & rsR1.Fields("TCT02")) > 0, "急件！", "") & strSub
       End If
       
       'Modified by Lydia 2018/12/17 去掉完成期限(所限)
       'If "" & rsR1.Fields("CP06") <> "" Or InStr(strSub, "急件") > 0 Then
       If InStr(strSub, "急件") > 0 Then
          strSub = strSub & "("
          'Remove by Lydia 2018/12/17 將完成期限刪除以免工程師誤判命名deadline(ex.FCP-060061工程師堅持要到所限才完成)
         'If "" & rsR1.Fields("CP06") <> "" Then
          '   strSub = strSub & "完成期限：" & ChangeTStringToTDateString(TransDate(rsR1.Fields("CP06"), 1)) & IIf(InStr(strSub, "急件") > 0, "，", "")
          'End If
          'end 2018/12/17
          If InStr(strSub, "急件") > 0 Then
             strSub = strSub & "譯畢期限：" & ChangeTStringToTDateString(TransDate(rsR1.Fields("TCT02"), 1)) & " " & Format(rsR1.Fields("TCT03"), "00:00")
          End If
          strSub = strSub & ")"
       
       'Added by Lydia 2019/10/03 (立案或分組,提申前)法定期限當天或前一天,Email通知主旨+急件!
       ElseIf iSta < 3 And Val("" & rsR1.Fields("PA10")) = 0 And Val("" & rsR1.Fields("CP07")) > 0 Then
            If strSrvDate(1) >= CompWorkDay(2, "" & rsR1.Fields("CP07"), 1) Then
                strSub = "急件！" & strSub
            End If
       'end 2019/10/03
       End If
       
       'Added by Lydia 2022/12/07 櫃台的新案收文(101-103)通知函，麻煩請增列該案代理人及申請人名稱 (若申請人未設則空白)，以利辨識新案立卷 , (尤其是急件) ---David ,Sharon(同意)
       'Modified by Lydia 2023/02/17區分「櫃台收新案=新案立卷」 iSta=1=>0
       If iSta = 0 And iMang = "B" Then
          'Added by Lydia 2023/02/21 外專新案認領：現有新案立卷Email通知EMAIL承辦及程序時，加註是否已進入認領階段。事後不必再通知。
          Str02 = ""
          If strSrvDate(1) >= 外專新案認領啟用日 Then
              Str01 = "select tcn16, tcn20, tcn23 from TrackingCaseName where tcn05=" & CNULL(m_TCT01)
              intR = 1
              Set rsB1 = ClsLawReadRstMsg(intR, Str01)
              If intR = 1 Then
                  If "" & rsB1.Fields("tcn16") = "Y" Then
                      strSub = strSub & "【暫不認領】"
                  ElseIf "" & rsB1.Fields("tcn20") <> "" Then
                      strSub = strSub & "【" & PUB_GetFCPGrpName("" & rsB1.Fields("tcn20")) & "】"
                  Else
                      strSub = strSub & "【進入" & IIf(Val("" & rsB1.Fields("tcn23")) = 0, "急件", "") & "認領階段】"
                  End If
              End If
          End If
          'end 2023/02/21
          strCont = "代理人：" & IIf("" & rsR1.Fields("PA75") <> "", rsR1.Fields("PA75") & " " & rsR1.Fields("PA75N"), "（空白）") & vbCrLf & _
                        "申請人：" & IIf("" & rsR1.Fields("PA26") <> "", rsR1.Fields("PA26") & " " & rsR1.Fields("PA26N"), "（空白）") & vbCrLf & strCont
       End If
       'end 2022/12/07
       strSub = iSubStart & strSub 'Added by Lydia 2017/12/15
       'Added by Lydia 2024/11/22 已收參考本的重新命名通知Email，增加註記----Bobbie
       If "" & rsR1.Fields("tcn13") = "3" Then
          strSub = strSub & "〔非英說案: 已收參考本〕"
       End If
       'end 2024/11/22
       
       If strCont = "" Then strCont = "同主旨"
       'Modified by Lydia 2023/06/14 email=Y/X編號+急件翻譯說明; ex.FCP-069710
       'If inContext <> "" Then strCont = inContext 'Added by Lydia 2018/01/03
       If inContext <> "" Then strCont = strCont & vbCrLf & vbCrLf & inContext
       
       strTo = ""
       stCC = "" 'Added by Lydia 2018/03/07
       '分案通知, 刪除通知,通知上傳檔案
       'Modified by Lydia 2021/04/16
       'If iSta = 1 Or iSta = 3 Or iSta = 4 Or iSta = 5 Then
       'Modified by Lydia 2023/02/17區分「櫃台收新案=新案立卷」+ 0
       If iSta = 0 Or iSta = 1 Or iSta = 3 Then
          '指定通知人員
          'Modified by Lydia 2022/08/08 將David加入(英文組)所有FCP, P櫃台新案立卷 (101-103)通知收件人之一
          'If iNto <> "" Then
          If iNto <> "" And Left(iNto, 1) <> "+" Then
            strTo = iNto & ";"
          Else
            'Modified by Lydia 2023/02/17區分「櫃台收新案=新案立卷」+ 0
            If iSta = 0 Or iSta = 1 Or iSta = 3 Then
                '分案-工程師主管
                If iSta <> 0 Then 'Added by Lydia 2023/02/17區分「櫃台收新案=新案立卷」
                   If iMang <> "" And iMang <> "B" Then '退程序只通知FCP程序管制人和主管、FCP承辦管制人和主管
                      strTo = iMang & ";"
                   ElseIf "" & rsR1.Fields("TCT04") <> "" Then
                      strTo = rsR1.Fields("TCT04") & ";"
                   End If
                End If 'Added by Lydia 2023/02/13
                
                'Adde by Lydia 2018/01/09 退程序主旨改成"新案立卷"以資區別(By 敏莉)
                If iMang = "B" Then strSub = Replace(strSub, "新案命名", "新案立卷")
                
                'FCP程序管制人和主管
                Str01 = PUB_GetFCPHandler(iCase01, iCase02, iCase03, iCase04)
                If Str01 <> "" Then
                   strTo = strTo & Str01 & ";"
                   'Modified by Lydia 2022/05/23 改用模組
                   'Str01 = "select nvl(st52,'N') from staff where st01='" & Str01 & "' "
                   'Set rsB1 = ClsLawReadRstMsg(intR, Str01)
                   'If intR = 1 Then
                   '   'Modified by Lydia 2018/03/07 改副本
                   '   'strTo = strTo & IIf("" & rsB1.Fields(0) <> "", rsB1.Fields(0) & ";", "")
                   '   If "" & rsB1.Fields(0) <> "N" Then stCC = stCC & rsB1.Fields(0) & ";"
                   'End If
                   Str02 = PUB_GetFCPProSup(Str01)
                   stCC = stCC & Str02 & ";"
                   'end 2022/05/23
                End If
            End If
            'FCP承辦管制人和主管
            Str01 = PUB_GetFCPSalesNo(iCase01, iCase02, iCase03, iCase04)
            If Str01 <> "" Then
               strTo = strTo & Str01 & ";"
               'Modified by Lydia 2022/05/23 改用模組
               'Str01 = "select nvl(st52,'N') from staff where st01='" & Str01 & "' "
               'Set rsB1 = ClsLawReadRstMsg(intR, Str01)
               'If intR = 1 Then
               '    'Modified by Lydia 2018/03/07 改副本
               '   'strTo = strTo & IIf("" & rsB1.Fields(0) <> "", rsB1.Fields(0) & ";", "")
               '   If "" & rsB1.Fields(0) <> "N" Then stCC = stCC & rsB1.Fields(0) & ";"
               'End If
               Str02 = PUB_GetFCPProSup(Str01)
               stCC = stCC & Str02 & ";"
               'end 2022/05/23
            End If
            
            If iSta = 0 Then 'Added by Lydia 2023/02/17區分「櫃台收新案=新案立卷」
               'Memo by Lydia 2022/12/27 新案立卷通知有變更，需檢查Service1「增加FCP/P案號時的系統通知 (請電腦中心比照附件新案立卷PUB_GetTCTmail通知承辦及相關人員)」是否要一併修改
               'Added by Lydia 2022/08/08 將David加入(英文組)所有FCP, P櫃台新案立卷 (101-103)通知收件人之一
               If iNto <> "" And Left(iNto, 1) = "+" And InStr(strTo, Mid(iNto, 2)) = 0 Then
                   strTo = strTo & Mid(iNto, 2) & ";"
               End If
               'end 2022/08/08
               'Added by Lydia 2022/08/19 將收文之智權人員列為收件人
               'Modified by Lydia 2023/02/17區分「櫃台收新案=新案立卷
               'If iSta = 1 And InStr(strSub, "新案立卷") > 0 And InStr(strTo & ";" & stCC, m_CP13) = 0 And m_CP13 <> "" Then
               If InStr(strTo & ";" & stCC, m_CP13) = 0 And m_CP13 <> "" Then
                    strTo = strTo & m_CP13 & ";"
               End If
               'end 2022/08/19
            End If 'Added by Lydia 2023/02/17
            
            'Added by Lydia 2018/03/05 退程序通知人員 'memo by Lydia 2019/08/01 查無資料誰要求額外加上通知退程序
            'Remove by Lydia 2019/08/01 實務上改工程師組別交給各區程序,已不需額外通知
            'If iMang = "B" Then
            '   Str01 = Pub_GetSpecMan("FCP退程序通知")
            '   If Str01 <> "" Then
            '       strTo = strTo & Str01 & ";"
            '   End If
            'End If
            'end 2018/03/05
            
          End If
       'Added by Lydia 2017/11/28 改組通知
       ElseIf iSta = 2 Then
           If iNto = "" Then
              strTo = "" & rsR1.Fields("TCT04") & ";"
           Else
              strTo = iNto & ";"
           End If
           If InStr(iChgGrp, "B") > 0 Then '退程序
              'FCP程序管制人和主管
              Str01 = PUB_GetFCPHandler(iCase01, iCase02, iCase03, iCase04)
              If Str01 <> "" Then
                 strTo = strTo & Str01 & ";"
                 'Modified by Lydia 2022/05/23 改用模組
                 'Str01 = "select nvl(st52,'N') from staff where st01='" & Str01 & "' "
                 'Set rsB1 = ClsLawReadRstMsg(intR, Str01)
                 'If intR = 1 Then
                 '   'Modified by Lydia 2018/03/07 改副本
                 '   'strTo = strTo & IIf("" & rsB1.Fields(0) <> "", rsB1.Fields(0) & ";", "")
                 '   'Modified by Lydia 2018/09/27 +判斷非N
                 '   If "" & rsB1.Fields(0) <> "N" Then stCC = stCC & rsB1.Fields(0) & ";"
                 'End If
                 Str02 = PUB_GetFCPProSup(Str01)
                 stCC = stCC & Str02 & ";"
                 'end 2022/05/23
              End If
              'Added by Lydia 2018/03/05 退程序通知人員 'memo by Lydia 2019/08/01 查無資料誰要求額外加上通知退程序
              'Remove by Lydia 2019/08/01 實務上改工程師組別交給各區程序,已不需額外通知
              'Str01 = Pub_GetSpecMan("FCP退程序通知")
              'If Str01 <> "" Then
              '    strTo = strTo & Str01 & ";"
              'End If
              ''end 2018/03/05
           End If
           strCont = "原分案組別:" & Left(iChgGrp, 1) & vbCrLf & "新分案組別:" & Right(iChgGrp, 1)
           strCont = Replace(strCont, "1", PUB_GetFCPGrpName("1"))
           strCont = Replace(strCont, "2", PUB_GetFCPGrpName("2"))
           strCont = Replace(strCont, "3", PUB_GetFCPGrpName("3"))
           strCont = Replace(strCont, "4", PUB_GetFCPGrpName("4"))
           strCont = Replace(strCont, "B", "退程序")
       'end 2017/11/28
       'Added by Lydia 2021/04/16
       'Modified by Lydia 2021/05/04 因為P案需要有卷才能作業，所以排除P案
       'Memo by Lydia 2021/10/19 後面有FMP無卷命名作業
       ElseIf iSta = 4 And Val("" & rsR1.Fields("cp158")) = 0 And "" & rsR1.Fields("cp01") <> "P" Then
           '命名完成時：工程師主管最後按確認時，系統自動發email通知如下：To: 程序，命名工程師, cc: 程序主管,工程師主管 (副理)
           'If "" & rsR1.Fields("pa150") = "3" Then 'Remove by Lydia 2021/05/13 配合防疫措施，英文組使用新案無卷命名email設定
               '程序
               strNA16 = PUB_GetFCPHandler(iCase01, iCase02, iCase03, iCase04)
               If strNA16 <> "" Then
                   strTo = strTo & strNA16 & ";"
                   'Mark by Lydia 2022/06/10 副本不寄 程序主管
                   'Str02 = PUB_GetFCPProSup(strNA16)
                   'If Str02 <> "" Then stCC = stCC & Str02 & ";"
                   'end 2022/06/10
               End If
               '工程師
               Str01 = "" & rsR1.Fields("tct10")
               If Str01 <> "" Then
                   strTo = strTo & Str01 & ";"
                   'Modified by Lydia 2022/12/21 命名作業完成通知排除Owen
                   'Str02 = PUB_GetFCPEngSup(Str01, True)
                   Str02 = PUB_GetFCPEngSup(Str01, True, True)
                   Str02 = PUB_GetStateForMan(Str02) 'Added by Lydia 2022/10/12 特殊情況之指定職代
                   If Str02 <> "" Then stCC = stCC & Str02 & ";"
               End If
               '內文
               strCont = strCont & vbCrLf
                'Added by Lydia 2021/04/29 若為103設計案 'Move by Lydia 2021/05/03 從下方移上來
                If "" & rsR1.Fields("CP10") = "103" Then
                   strCont = strCont & "本案為設計案，程序人員請將卷退工程師，進行製作外文提申本。" & vbCrLf
                End If
                'end 2021/04/29
               Str01 = "": Str02 = ""
               strTempA = "union select 'x1' as ord1, cp09,cp10,cpm03,cp48 from caseprogress, casepropertymap, staff " & _
                                 "where cp01='" & iCase01 & "' and cp02='" & iCase02 & "' and cp03='" & iCase03 & "' and cp04='" & iCase04 & "' and cp01=cpm01(+) and cp10=cpm02(+) and cp158=0 and cp159=0 and cp65=st01(+) "
               '命名有勾選提申前告代/主動修正
               If "" & rsR1.Fields("tct20") = "2" Or "" & rsR1.Fields("tct20") = "3" Then
                   Str01 = Str01 & Replace(strTempA & " and cp10='901' and substr(cp09,1,1) = 'B' and (st03='F21' or st03='M51') and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' ", "x1", "01")
               End If
               If "" & rsR1.Fields("tct117") = "2" Then
                   Str01 = Str01 & Replace(strTempA & " and cp10='203' and substr(cp09,1,1) = 'B' and (st03='F21' or st03='M51') and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' ", "x1", "02")
               End If
              '有收文A類 告代 / 主動修正 / 回代
               Str01 = Str01 & Replace(strTempA & " and cp10 in ('901','902','203') and substr(cp09,1,1) = 'A' ", "x1", "10")
               Str01 = Mid(Str01, 6) & "order by ord1,cp10 "
               intR = 1
               Set rsB1 = ClsLawReadRstMsg(intR, Str01)
               If intR = 0 Then
                   If "" & rsR1.Fields("CP10") <> "103" Then  'Added by Lydia 2021/05/19 排除103設計案
                         strCont = strCont & "無提申前告代/主動修正：程序人員可進行提申作業" & vbCrLf
                   End If 'Added by Lydia 2021/05/19
               Else
                   Str01 = "": Str02 = ""
                   rsB1.MoveFirst
                   Do While Not rsB1.EOF
                       If Val("" & rsB1.Fields("ord1")) < 10 Then '命名有勾選提申前告代/主動修正
                            If "" & rsB1.Fields("cp10") = "901" Then
                                 Str01 = Str01 & "/提申前告代" & IIf("" & rsB1.Fields("cp48") <> "", "(" & ChangeWStringToTDateString("" & rsB1.Fields("cp48")) & ")", "")
                            ElseIf "" & rsB1.Fields("cp10") = "203" Then
                                 Str01 = Str01 & "/提申前主動修正" & IIf("" & rsB1.Fields("cp48") <> "", "(" & ChangeWStringToTDateString("" & rsB1.Fields("cp48")) & ")", "")
                            End If
                       ElseIf Val("" & rsB1.Fields("ord1")) >= 10 Then ''有收文A類 告代 / 主動修正 / 回代
                            If "" & rsB1.Fields("cp10") = "901" Then
                                 Str02 = Str02 & "/告代" & IIf("" & rsB1.Fields("cp48") <> "", "(" & ChangeWStringToTDateString("" & rsB1.Fields("cp48")) & ")", "")
                            ElseIf "" & rsB1.Fields("cp10") = "902" Then
                                 Str02 = Str02 & "/回代" & IIf("" & rsB1.Fields("cp48") <> "", "(" & ChangeWStringToTDateString("" & rsB1.Fields("cp48")) & ")", "")
                            ElseIf "" & rsB1.Fields("cp10") = "203" Then
                                 Str02 = Str02 & "/主動修正" & IIf("" & rsB1.Fields("cp48") <> "", "(" & ChangeWStringToTDateString("" & rsB1.Fields("cp48")) & ")", "")
                            End If
                       End If
                       rsB1.MoveNext
                   Loop
                   'Memo by Lydia 2021/05/03 移到上方
                   If Str01 <> "" Then
                      strCont = strCont & "有" & Mid(Str01, 2) & "：程序人員將卷退工程師，請工程師進行作業。" & vbCrLf
                   End If
                   If Str02 <> "" Then
                      strCont = strCont & "有收文" & Mid(Str02, 2) & "：請工程師回覆 " & GetStaffName(strNA16) & " 為提申前 or 提申後作業，再進行相關之流程。" & vbCrLf
                   End If
                   'Modified by Lydia 2021/04/29 排除103設計案
                   'If Str01 & Str02 = "" Then
                   If Str01 & Str02 = "" And "" & rsR1.Fields("CP10") <> "103" Then
                      strCont = strCont & "無提申前告代/主動修正：程序人員可進行提申作業" & vbCrLf
                   End If
               End If
          'End If　'Remove by Lydia 2021/05/13 配合防疫措施，英文組使用新案無卷命名email設定
       'end 2021/04/16
       'Added by Lydia 2021/10/19 FMP無卷命名作業
       ElseIf iSta = 4 And Val("" & rsR1.Fields("cp158")) = 0 And "" & rsR1.Fields("cp01") = "P" Then
           '命名完成時：工程師主管最後按確認時，系統自動發email通知如下：To: 程序，命名工程師, cc: 程序主管,工程師主管 (副理)
            '程序
            strNA16 = PUB_GetFCPHandler(iCase01, iCase02, iCase03, iCase04)
            If strNA16 <> "" Then
                strTo = strTo & strNA16 & ";"
                'Mark by Lydia 2022/06/10 副本不寄 程序主管
                'Str02 = PUB_GetFCPProSup(strNA16)
                'If Str02 <> "" Then stCC = stCC & Str02 & ";"
                'end 2022/06/10
            End If
            '工程師
            Str01 = "" & rsR1.Fields("tct10")
            If Str01 <> "" Then
                strTo = strTo & Str01 & ";"
                'Modified by Lydia 2022/12/21 命名作業完成通知排除Owen
                'Str02 = PUB_GetFCPEngSup(Str01, True)
                Str02 = PUB_GetFCPEngSup(Str01, True, True)
                Str02 = PUB_GetStateForMan(Str02) 'Added by Lydia 2022/10/12 特殊情況之指定職代
                If Str02 <> "" Then stCC = stCC & Str02 & ";"
            End If
            '內文
            strCont = strCont & vbCrLf
            Str01 = "": Str02 = "": intCnt = 0
            '以下依案件狀況增加內文項目:
            '1. 收文【201新案翻譯】
            strTempA = "select cp09, cp10||cpm04 as cp10n,sqldatet(tf26) tf26t from caseprogress, casepropertymap, transfee " & _
                              "where cp01='" & iCase01 & "' and cp02='" & iCase02 & "' and cp03='" & iCase03 & "' and cp04='" & iCase04 & "' and cp10='201' and cp159=0 and cp158=0 and cp01=cpm01(+) and cp10=cpm02(+) and cp09=tf01(+) "
            intR = 1
            Set rsB1 = ClsLawReadRstMsg(intR, strTempA)
            If intR = 1 Then
                 intCnt = intCnt + 1
                 strCont = strCont & intCnt & ". 收文【" & rsB1.Fields("cp10n") & "】" & vbCrLf
                 If "" & rsB1.Fields("tf26t") <> "" Then
                      strCont = strCont & "   翻譯交稿期限：" & rsB1.Fields("tf26t") & vbCrLf
                 End If
                 strCont = strCont & "   程序人員請通知Sharon進行分案作業，待交稿後通知工程師進行作業。" & vbCrLf
                 strCont = strCont & vbCrLf
            End If
            '2. 收文【209檢視中說】
            strTempA = "select cp09, cp10||cpm04 as cp10n from caseprogress, casepropertymap " & _
                              "where cp01='" & iCase01 & "' and cp02='" & iCase02 & "' and cp03='" & iCase03 & "' and cp04='" & iCase04 & "' and cp10='209' and cp159=0 and cp158=0 and cp01=cpm01(+) and cp10=cpm02(+)  "
            intR = 1
            Set rsB1 = ClsLawReadRstMsg(intR, strTempA)
            If intR = 1 Then
                 intCnt = intCnt + 1
                 strCont = strCont & intCnt & ". 收文【" & rsB1.Fields("cp10n") & "】" & vbCrLf
                 '與FCPxxxxx相關(帶出相關FCP案-分案之多國案卷號)
                 strTempA = "select cm05,cm06,cm07,cm08 from casemap where cm01='" & iCase01 & "' and cm02='" & iCase02 & "' and cm03='" & iCase03 & "' and cm04='" & iCase04 & "' and cm10='0' and cm05='FCP' " & _
                                   "union all select cm01,cm02,cm03,cm04 from casemap where cm05='" & iCase01 & "' and cm06='" & iCase02 & "' and cm07='" & iCase03 & "' and cm08='" & iCase04 & "' and cm10='0' and cm01='FCP' "
                               
                 strTempA = "select cp01||'-'||cp02||'-'||cp03||'-'||cp04 as caseno ,sqldatet(tf26) as tf26t,sqldatet(ep09) as ep09t " & _
                                   "From caseprogress, transfee, engineerprogress where (cp01,cp02,cp03,cp04) in (" & strTempA & _
                                   ") and cp10='201' and cp159=0 and cp09=tf01(+) and cp09=ep02(+) "
                 intR = 1
                 Set rsB1 = ClsLawReadRstMsg(intR, strTempA)
                 If intR = 1 Then
                      'Mark by Lydia 2024/01/31 備註寫錯
                      'If "" & rsB1.Fields("caseno") <> "" Then
                      '    strCont = strCont & "   與" & rsB1.Fields("caseno") & "相關" & vbCrLf
                      'End If
                      'end 2024/01/31
                      If "" & rsB1.Fields("ep09t") <> "" Then
                           '請工程師進行檢視中說相關作業。
                           strCont = strCont & "   請工程師進行檢視中說相關作業。" & vbCrLf
                      Else
                           If "" & rsB1.Fields("tf26t") <> "" Then
                                strCont = strCont & "   翻譯交稿期限：" & rsB1.Fields("tf26t") & vbCrLf
                           End If
                           'Modified by Lydia 2024/01/31 備註寫錯; ex.(1/23) P-132937命名完成
                           strCont = strCont & "   待交稿後通知工程師進行檢視中說作業。" & vbCrLf
                      End If
                 End If
                 strCont = strCont & vbCrLf
            End If
            
            '3. 收文【942檢視PCT公開本與FCP相異】
            strTempA = "select cp09, cp10||cpm04 as cp10n from caseprogress, casepropertymap " & _
                              "where cp01='" & iCase01 & "' and cp02='" & iCase02 & "' and cp03='" & iCase03 & "' and cp04='" & iCase04 & "' and cp10='942' and cp159=0 and cp158=0 and cp01=cpm01(+) and cp10=cpm02(+)  "
            intR = 1
            Set rsB1 = ClsLawReadRstMsg(intR, strTempA)
            If intR = 1 Then
                 intCnt = intCnt + 1
                 strCont = strCont & intCnt & ". 收文【" & rsB1.Fields("cp10n") & "】" & vbCrLf
                 '與FCPxxxxx相關(帶出相關FCP案-分案之多國案卷號)
                 strTempA = "select cm05,cm06,cm07,cm08 from casemap where cm01='" & iCase01 & "' and cm02='" & iCase02 & "' and cm03='" & iCase03 & "' and cm04='" & iCase04 & "' and cm10='" & iCase03 & "' and cm05='FCP' " & _
                                   "union all select cm01,cm02,cm03,cm04 from casemap where cm05='" & iCase01 & "' and cm06='" & iCase02 & "' and cm07='" & iCase03 & "' and cm08='" & iCase04 & "' and cm10='" & iCase03 & "' and cm01='FCP' "
                 intR = 1
                 Set rsB1 = ClsLawReadRstMsg(intR, strTempA)
                 If intR = 1 Then
                      strCont = strCont & "   與" & rsB1.Fields("cm05") & "-" & rsB1.Fields("cm06") & "-" & rsB1.Fields("cm07") & "-" & rsB1.Fields("cm08") & "相關" & vbCrLf
                      strCont = strCont & "   請工程師進行檢視PCT公開本與FCP相異相關作業。" & vbCrLf
                 End If
                 strCont = strCont & vbCrLf
            End If
            '4. 有收文【901告代】【902回代】【924會稿】【203主動補正】【228呈國際階段修正】
            strTempA = "select cp10||cpm04 as cp10n from caseprogress, casepropertymap " & _
                              "where cp01='" & iCase01 & "' and cp02='" & iCase02 & "' and cp03='" & iCase03 & "' and cp04='" & iCase04 & "' " & _
                              "and cp10 in ('901','902','924','203','228') and cp159=0 and cp158=0 and cp01=cpm01(+) and cp10=cpm02(+) order by cp10"
            intR = 1
            Set rsB1 = ClsLawReadRstMsg(intR, strTempA)
            If intR = 1 Then
                 strTempA = rsB1.GetString(adClipString, , , ",")
                 If Right(strTempA, 1) = "," Then strTempA = Mid(strTempA, 1, Len(strTempA) - 1)
                 intCnt = intCnt + 1
                 strCont = strCont & intCnt & ". 收文【" & Replace(strTempA, ",", "】、【") & "】" & vbCrLf
                 strCont = strCont & "   請工程師進行相關作業，若有需待翻譯交稿，待交稿後再行處理。" & vbCrLf
            End If
            'Added by Lydia 2022/12/13  收文【210撰稿】
            strTempA = "select cp10||cpm04 as cp10n from caseprogress, casepropertymap " & _
                              "where cp01='" & iCase01 & "' and cp02='" & iCase02 & "' and cp03='" & iCase03 & "' and cp04='" & iCase04 & "' " & _
                              "and cp10 ='210'  and cp159=0 and cp158=0 and cp01=cpm01(+) and cp10=cpm02(+) order by cp10"
            intR = 1
            Set rsB1 = ClsLawReadRstMsg(intR, strTempA)
            If intR = 1 Then
                 strCont = strCont & "   請工程師進行撰稿相關作業。" & vbCrLf
            End If
            'end 2022/12/13
       'end 2021/10/19
       Else
       End If
       
       'Added by Lydia 2018/04/17 傳入副本收受者
       'Modified by Lydia 2020/04/27 排除已在收件者
       If inCC <> "" And InStr(strTo, inCC) = 0 Then
           stCC = stCC & inCC & ";"
      End If
       'Added by Lydia 2018/03/07
          If strTo = "" Then
               Exit Function
          'Modified by Lydia 2022/08/19 加判斷
          'Else
          ElseIf Right(strTo, 1) = ";" Then
               strTo = Mid(strTo, 1, Len(strTo) - 1)
          End If
          'Modified by Lydia 2022/08/19 加判斷
          'If stCC <> "" Then
          If stCC <> "" And Right(stCC, 1) = ";" Then
              stCC = Mid(stCC, 1, Len(stCC) - 1)
          End If
       'end 2018/03/07
       
       '發email
       If bolCache = True Then
          'Modified by Lydia 2018/03/07 +副本
          'strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
             " values( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
             ",to_char(sysdate,'hh24miss'),'" & strSub & "','" & strCont & "')"
          'Modified by Lydia 2018/03/27 +Replace
          'Modified by Lydia 2022/12/16 chgsql
          strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
             " values( '" & strUserNum & "','" & Replace(strTo, ",", ";") & "',to_char(sysdate,'yyyymmdd')" & _
             ",to_char(sysdate,'hh24miss'),'" & strSub & "','" & ChgSQL(strCont) & "'," & CNULL(Replace(stCC, ",", ";")) & ")"
          cnnConnection.Execute strSql
       Else
          'Modified by Lydia 2018/03/07 +副本
          'PUB_SendMail strUserNum, strTo, "", strSub, vbCrLf & strCont
          PUB_SendMail strUserNum, strTo, "", strSub, vbCrLf & strCont, , , , , , stCC
       End If
       PUB_GetTCTmail = True
       
    End If
    Set rsR1 = Nothing
    Set rsB1 = Nothing
    Exit Function
    
ErrHandle:
    If Err.Number <> 0 Then
       MsgBox Err.Description
    End If
End Function

'Added by Lydia 2017/12/04 FCP案件命名電子化:讀取資料
'Modified by Lydia 2018/04/18  +新申請案:是否電子送件(rCP118)、發文日(rCP27), 新案翻譯:所限(tCP06)、發文日(tCP27)
'Move by Lydia 2023/02/15 從basUpdate搬過來
Public Function PUB_GetTCTread(ByRef fm As Form, ByRef rCase() As String, ByRef rStrKind As String, _
                                                 ByRef rCP118 As String, Optional ByRef rCP27 As String, Optional ByRef tCP06 As String, Optional ByRef tCP27 As String) As Boolean
Dim rsR1 As New ADODB.Recordset
Dim rsR2 As New ADODB.Recordset
Dim intR As Integer
Dim strTmp1 As String


On Error GoTo ErrHandle
    
    'Added by Lydia 2023/03/01 排除顯示的表單
    Dim strJumpList As String
    strJumpList = "frm090902_1,frm090908_1"
    'end 2023/03/01
    
    PUB_GetTCTread = False
    'Modified by Lydia 2018/04/18 9-申請日,10-公告日(PA14),11-目前准/駁(PA16)
    'Modified by Lydia 2020/02/17 +12-名稱有特殊字(PA174)
    strTmp1 = "select pa150,DECODE(pa150,'1','" & PUB_GetFCPGrpName("1") & "','2','" & PUB_GetFCPGrpName("2") & "','3','" & PUB_GetFCPGrpName("3") & "','4','" & PUB_GetFCPGrpName("4") & "',pa150) grpname " & _
                ",pa08,pa09,pa26,pa75,cu10,n1.na03 cna03,fa10,n2.na03 fna03 " & _
                ", pa05,pa06,pa07,pa158,nvl(fa104,'N') fa104,nvl(cu174,'N') cu174,nvl(pa49,nvl(pa50,0)) pa_dc,nvl(fa25,nvl(fa26,0)) fa_dc,nvl(cu36,nvl(cu37,0)) cu_dc " & _
                ",pa10,pa14,pa16,pa174 from patent,fagent,customer,nation n1,nation n2 " & _
                "where pa01='" & rCase(1) & "' and pa02='" & rCase(2) & "'  and pa03='" & rCase(3) & "' and pa04='" & rCase(4) & "' " & _
                "and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and cu10=n1.na01(+) " & _
                "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and fa10=n2.na01(+) "
    intR = 1
    Set rsR1 = ClsLawReadRstMsg(intR, strTmp1)
    If intR = 1 Then
       '專利種類
       rCase(5) = "" & rsR1.Fields("pa08")
       '申請國家
       rCase(6) = "" & rsR1.Fields("pa09")
       fm.lblData(2).Caption = "" & PUB_GetPatentKindName(rCase(5), rCase(6))

       '分案組別
       rCase(7) = "" & rsR1.Fields("pa150")
       fm.lblData(11).Caption = "" & rsR1.Fields("grpname")
       
       '專利種類-設計:顯示設計案屬性=>pa158
       'Modified by Lydia 2020/07/20 只有台灣案有設計案屬性
       'If fm.Name <> "frm090902_1" And rCase(5) <> "3" Then
       'Modified by Lydia 2023/03/01 fm.Name <> "frm090902_1"  => InStr(strJumpList, fm.Name) = 0
       If InStr(strJumpList, fm.Name) = 0 And Not (rCase(5) = "3" And rCase(6) = "000") Then
          fm.Frame1.Visible = False
          For intR = 35 To 39
             fm.Chk2(intR).Visible = False
          Next
          fm.txtData(45).Visible = False
       End If
       
       rCase(8) = "" & rsR1.Fields("pa158")
       '預設代入基本檔案件屬性
       'Modified by Lydia 2023/03/01 fm.Name <> "frm090902_1"  => InStr(strJumpList, fm.Name) = 0
       If InStr(strJumpList, fm.Name) = 0 And Val(rCase(8)) >= 1 And Val(rCase(8)) <= 4 Then
          fm.opt1(Val(rCase(8)) - 1).Value = 1
       End If
       'Added by Lydia 2018/04/18
       rCase(9) = "" & rsR1.Fields("pa10") '申請日
       rCase(10) = "" & rsR1.Fields("pa14") '公告日
       rCase(11) = "" & rsR1.Fields("pa16") '目前准/駁
       rCase(12) = "" & rsR1.Fields("pa174") 'Added by Lydia 2020/02/17 名稱有特殊字
       
       '預設代入基本檔名稱
       'Modified by Lydia 2023/03/01 fm.Name <> "frm090902_1"  => InStr(strJumpList, fm.Name) = 0
       If InStr(strJumpList, fm.Name) = 0 Then
          fm.txtData(3).Text = "" & rsR1.Fields("pa05")
          fm.txtData(4).Text = "" & rsR1.Fields("pa06")
          fm.txtData(5).Text = "" & rsR1.Fields("pa07")
          fm.txtData(5).Locked = True '日文不可輸入
       End If
       
       If Trim("" & rsR1.Fields("fna03") <> "") Then
           fm.Label4(5).Caption = "代理人國籍："
           fm.lblData(10).Caption = "" & rsR1.Fields("fna03")
       Else
           fm.Label4(5).Caption = "申請人國籍："
           fm.lblData(10).Caption = "" & rsR1.Fields("cna03")
       End If
       
       'FCP是否電子送件,外專發文時做檢查用
       strExc(1) = ""
       If "" & rsR1.Fields("fa104") = "Y" Then
          strExc(1) = "代理人"
       End If
       If "" & rsR1.Fields("cu174") = "Y" Then
          strExc(1) = strExc(1) & IIf(strExc(1) <> "", "、", "") & "客戶"
       End If
       fm.lblData(16).Caption = IIf(strExc(1) <> "", "電子送件", "")
       
       'Added by Lydia 2017/12/08 抓固定報價之XY清單
       'Modified by Lydia 2018/04/01 代號'FCPtcn'=>'FCPtct'
       'Modified by Lydia 2018/08/27 改成模組
       'strTmp1 = "select AAL04 from AddressA4List where AAL01='FCPtct' order by AAL03"
       'intR = 1
       'Set rsR2 = ClsLawReadRstMsg(intR, strTmp1)
       'If intR = 1 Then
       '   strTmp1 = rsR2.GetString(adClipString, , , ",")
       'End If
       strTmp1 = Pub_GetPa62Flag(rCase(1) & rCase(2) & rCase(3) & rCase(4))
       
       'Added by Lydia 2018/09/26 有相似度只能上班翻譯
       strExc(1) = "select tf01,tf19,tf20 from caseprogress,transfee " & _
                         "where cp01='" & rCase(1) & "' and cp02='" & rCase(2) & "' and cp03='" & rCase(3) & "' and cp04='" & rCase(4) & "' and cp10 in (" & FcpTctPtys & ") and cp159=0 and cp09=tf01(+) "
       intR = 1
       Set rsR2 = ClsLawReadRstMsg(intR, strExc(1))
       If intR = 1 Then
          If Val("" & rsR2.Fields("tf19")) > 0 Then strTmp1 = "Y"
       End If
       'end 2018/09/26
       
       '若此案有折扣(Y編號、X編號、個案有設專利翻譯折扣、專利全部折扣者或外專提供的固定報價之客戶清單)，則工程師只能點選B上班翻，無法點選A下班翻。
       'Modified by Lydia 2018/08/27 固定報價判斷改成模組
       'If Val("" & rsR1.Fields("pa_dc")) + Val("" & rsR1.Fields("fa_dc")) + Val("" & rsR1.Fields("cu_dc")) > 0 _
            Or (InStr(strTmp1, "" & rsR1.Fields("pa75")) > 0 And "" & rsR1.Fields("pa75") <> "") _
            Or (InStr(strTmp1, "" & rsR1.Fields("pa26")) > 0 And "" & rsR1.Fields("pa26") <> "") Then
       If strTmp1 = "Y" Or Val("" & rsR1.Fields("pa_dc")) + Val("" & rsR1.Fields("fa_dc")) + Val("" & rsR1.Fields("cu_dc")) > 0 Then
          rStrKind = "B上班翻譯"
       Else
          rStrKind = "A下班翻譯或B上班翻譯"
       End If
       
       'Added by Lydia 2018/10/22 顯示相關案(台灣大陸案件提示)
       strTmp1 = " select pa01,pa02,pa03,pa04,pa09 from casemap,patent where cm10='0' and cm01='" & rCase(1) & "' and cm02='" & rCase(2) & "' and cm03='" & rCase(3) & "' and cm04='" & rCase(4) & "'" & _
                      " and cm05=pa01(+) and cm06=pa02(+) and cm07=pa03(+) and cm08=pa04(+) and pa09=" & CNULL(IIf(rCase(6) = "000", "020", "000"))
       strTmp1 = strTmp1 & " union select pa01,pa02,pa03,pa04,pa09 from casemap,patent where cm10='0' and cm05='" & rCase(1) & "' and cm06='" & rCase(2) & "' and cm07='" & rCase(3) & "' and cm08='" & rCase(4) & "'" & _
                      " and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and pa09=" & CNULL(IIf(rCase(6) = "000", "020", "000"))
       strTmp1 = strTmp1 & " order by 1,2 desc "
       intR = 1
       Set rsR2 = ClsLawReadRstMsg(intR, strTmp1)
       If intR = 1 Then
            fm.lblCMboth.Caption = "相關案號：" & rsR2.Fields("pa01") & "-" & rsR2.Fields("pa02") & "-" & rsR2.Fields("pa03") & "-" & rsR2.Fields("pa04")
            fm.lblCMboth.Tag = "" & rsR2.Fields("pa01") & rsR2.Fields("pa02") & rsR2.Fields("pa03") & rsR2.Fields("pa04")
       End If
       'end 2018/10/22

       'Modified by Lydia 2018/03/01 +實審416
       'Modified by Lydia 2018/04/17 +回代902,主動修正203, 限制A類收文
       'Modified by Lydia 2018/04/18 +CP118,CP27
       'Modified by Lydia 2018/05/10 +FMP案414,938,939,106,228
       strTmp1 = "select 2 as ord1 ,sqldatet(cp05) cp05,sqldatet(cp06) cp06,sqldatet(cp07)cp07,cp09,cp10,cpm03,cp13,st02,cp118,CP27 " & _
                   "from caseprogress,staff,casepropertymap " & _
                   "where cp01='" & rCase(1) & "' and cp02='" & rCase(2) & "'  and cp03='" & rCase(3) & "' and cp04='" & rCase(4) & "' " & _
                   "and cp10 in (" & GetAddStr(FcpTctPtys & ",416,902,203,414,938,939,106,228") & ") and substr(cp09,1,1)='A' and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) "
       strTmp1 = strTmp1 & "union all select 1 as ord1 ,sqldatet(cp05) cp05,sqldatet(cp06) cp06,sqldatet(cp07)cp07,cp09,cp10,cpm03,cp13,st02,cp118,CP27 " & _
                   "from caseprogress,staff,casepropertymap " & _
                   "where cp01='" & rCase(1) & "' and cp02='" & rCase(2) & "'  and cp03='" & rCase(3) & "' and cp04='" & rCase(4) & "' " & _
                   "and cp10 in (" & GetAddStr(NewCasePtyList) & ") and cp13=st01(+) and cp01=cpm01(+) and cp10=cpm02(+) "
       strTmp1 = strTmp1 & "order by ord1,cp09 "
       intR = 1
       Set rsR2 = ClsLawReadRstMsg(intR, strTmp1)
       If intR = 1 Then
          With rsR2
              intR = 1: strExc(1) = ""   'Added by Lydia 2018/05/10
              .MoveFirst
              Do While Not .EOF
                 If Val("" & .Fields("ord1")) = 2 Then  '中說進度
                    '總收文號：因為同時顯示2道，所以只抓6碼
                    'Modified by Lydia 2018/04/17 第一道收文號顯示9碼,後面顯示3碼
                    'fm.lblData(4).Caption = fm.lblData(4).Caption & "," & IIf(.AbsolutePosition = 4, vbCrLf, "") & Right("" & .Fields("cp09"), 6)
                    'Modified by Lydia 2018/05/10 +FMP案
                    'fm.lblData(4).Caption = fm.lblData(4).Caption & "," & IIf(.AbsolutePosition = 5, vbCrLf, "") & Right("" & .Fields("cp09"), 3)
                    strExc(1) = strExc(1) & "," & Right("" & .Fields("cp09"), 3)
                    
                    If InStr(FcpTctPtys, "" & .Fields("cp10")) > 0 Then

                        '中說類型
                        If "" & .Fields("cp10") = "242" Then
                           '製作中說210＆外文提申本242 是一起產生
                           fm.lblData(3).Caption = fm.lblData(3).Caption & "＆外文提申本"
                        Else
                           fm.lblData(3).Caption = "" & .Fields("cpm03")
                           'Added by Lydia 2018/04/18 中說-所限,發文日
                           tCP06 = TransDate(Replace("" & .Fields("cp06"), "/", ""), 2)
                           tCP27 = TransDate(Replace("" & .Fields("cp27"), "/", ""), 2)
                        End If
                    End If
                 Else
                    '所限
                    'Modified by Lydia 2018/05/22
                    'fm.lblData(0).Caption = "" & .Fields("cp06")
                    fm.lblData(1).Caption = "" & .Fields("cp06")
                    '法限
                    'Modified by Lydia 2018/05/22
                    'fm.lblData(1).Caption = "" & .Fields("cp07")
                    fm.lblData(0).Caption = "" & .Fields("cp07")
                    '收文號
                    'Modified by Lydia 2018/04/17 第一道收文號顯示9碼
                    'fm.lblData(4).Caption = Right("" & .Fields("cp09"), 6)
                    fm.lblData(4).Caption = "" & .Fields("cp09")
                    '收文日
                    fm.lblData(5).Caption = "" & .Fields("cp05")
                    '本所案號
                    fm.lblData(6).Caption = rCase(1) & "-" & rCase(2) & "-" & rCase(3) & "-" & rCase(4)
                    '智權人員
                    'fm.lblData(8).Caption = "" & .Fields("cp13") 'Remove by Lydia 2018/03/06 改成命名人員
                    fm.lblData(9).Caption = "" & .Fields("st02")
                    '案件性質
                    fm.lblData(15).Caption = "" & .Fields("cpm03")
                    fm.lblData(15).Tag = "" & .Fields("cp10")
                    'Added by Lydia 2018/04/18
                    rCP118 = "" & .Fields("cp118") '新案是否電子送件
                    rCP27 = "" & .Fields("cp27") '新案發文日
                 End If
                 intR = intR + 1  'Added by Lydia 2018/05/10
                 .MoveNext
              Loop
              fm.lblData(4).Caption = fm.lblData(4) & "~" & Right(strExc(1), 3) 'Added by Lydia 2018/05/10 只顯示最後收文號3碼
          End With
          PUB_GetTCTread = True
       End If
    End If
    
ErrHandle:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
Set rsR1 = Nothing
Set rsR2 = Nothing
End Function

'Added by Lydia 2017/12/27 預設外專案件共用資料夾路徑
'Modified by Lydia 2018/05/09 +pA01 系統別,判斷是否為FMP
'Remove by Lydia 2021/12/06 (109/4/6)已將\\Typing2的"English_Vers"和"專利案件"的案件資料夾，全部搬到原始檔區
'Public Function Pub_GetFCPcaseFilePath(ByVal pNo As String, Optional ByVal bolParent As Boolean = False, Optional ByVal pa01 As String = "") As String
''pNo 本所案號流水號 (6碼)
''bolParent 取得案件資料夾的上一層
'    Pub_GetFCPcaseFilePath = ""
'    If pNo = "" Then Exit Function
'    'Added by Lydia 2018/05/09 判斷是否為FMP
'    If pa01 = "P" Then
'        '用6碼卷號
'        Pub_GetFCPcaseFilePath = FCP命名追蹤收文區 & "\" & Val(pNo)
'    Else
'     'end 2018/05/09
'        Pub_GetFCPcaseFilePath = FCP命名追蹤收文區 & "\" & Mid(Val(pNo), 1, 3) & IIf(bolParent = True, "", "\" & Val(pNo))
'    End If 'end 2018/05/09
'End Function
'end 2017/12/27
'end 2021/12/06

'Added by Lydia 2018/02/23 Typing2_FTP上傳檔案
'Move by Lydia 2023/02/15 從basUpdate搬過來
Public Function Pub_FtpPutTyping2(ByVal pFromPath As String, ByVal pToPath As String) As Boolean
Dim stFtpIP As String
Dim stFtpPath As String
Dim stDir As String
Dim stFName As String

    Pub_FtpPutTyping2 = False
    stFtpIP = Pub_GetSpecMan("FTP_TYPING2")
    stFtpPath = Replace(pToPath, "\", "/")

    If stFtpIP = "" Then Exit Function
    
    stDir = "//" & Mid(stFtpPath, InStr(3, stFtpPath, "/") + 1)
    stDir = Mid(stDir, 1, InStrRev(stDir, "/") - 1)
    stFName = Mid(stFtpPath, InStrRev(stFtpPath, "/") + 1)
    If PUB_FtpPutFile(pFromPath, stDir & "/" & stFName, , , stFtpIP) = True Then
        Pub_FtpPutTyping2 = True
    End If
End Function

'Added by Lydia 2018/02/23 Typing2_FTP刪除檔案
'Move by Lydia 2023/02/15 從basUpdate搬過來
Public Function Pub_FtpDelTyping2(ByVal pKind As String, ByVal pToDir As String, Optional ByVal pFName As String = "") As Boolean
Dim stFtpIP As String
Dim stFtpPath As String
Dim stDir As String

    Pub_FtpDelTyping2 = False
    stFtpIP = Pub_GetSpecMan("FTP_TYPING2")
    stFtpPath = Replace(pToDir, "\", "/")

    If stFtpIP = "" Then Exit Function
    
    stDir = "//" & Mid(stFtpPath, InStr(3, stFtpPath, "/") + 1)
    If UCase(pKind) = "TRACKING_NO" And pFName = "" Then
         '命名追蹤刪除->直接刪除FTP資料夾
         If PUB_FtpDelFile(stDir, , , , stFtpIP) = True Then
               Pub_FtpDelTyping2 = True
         End If
    Else
         If PUB_FtpDelFile(stDir, pFName, , , stFtpIP) = True Then
               Pub_FtpDelTyping2 = True
         End If
    End If

End Function

'Added by Lydia 2018/02/27 Typing2_FTP 搬移檔案
'Move by Lydia 2023/02/15 從basUpdate搬過來
'Mark by Lydia 2023/10/12 改用原始檔區存放
'Public Function PUB_FtpRenTyping2(pOldDir As String, pNewDir As String, Optional pChgName As String = "", Optional pFtpIp As String = "", Optional pCreate As Boolean = False, _
'                                                    Optional pMsgList As String = "", Optional pErrMsg As String, Optional pRaiseErr As Boolean = True) As Boolean
''pChgName 是否更名
''pCreate    是否新建目的地目錄
''pMsgList  *.MSG檔的名稱
'
'   Dim hConnection As Long
'   Dim pData As WIN32_FIND_DATA
'   Dim hReturn As Long, hFind As Long, LRet  As Long, stFileName As String
'   Dim nConnection As Long
'   Dim nReturn As Long, nFind As Long
'   Dim nData As WIN32_FIND_DATA
'   Dim strNewFName As String
'   'Added by Lydia 2020/02/13
'   Dim cMsg As Integer, strTmp As String, strTmp2 As String
'   Dim strOldPath As String
'
'    hConnection = PUB_GetFtpConnect(pErrMsg, , , pFtpIp)
'   'Added by Lydia 2020/02/13 English_Vers檔案：判斷啟用日
'   If strSrvDate(1) >= XY特殊權限啟用日by檔案 Then
'       If Dir(pNewDir, vbDirectory) = "" Then
'           MkDir pNewDir
'       End If
'       strOldPath = FCP命名追蹤暫存 & Mid(pNewDir, InStrRev(pNewDir, "\"))
'   Else
'   'end 2020/02/13
'       If PUB_SetFtpDirectory(hConnection, pNewDir, pErrMsg, pRaiseErr, pCreate) = False Then GoTo OutPort
'
'    '若資料夾已存在,將資料夾的檔案名稱後+"_old"
'       If pCreate = False Then
'            nConnection = PUB_GetFtpConnect(pErrMsg, , , pFtpIp)
'            If FtpSetCurrentDirectory(nConnection, pNewDir) = 1 Then
'               nData.cFileName = String(MAX_PATH, 0)
'               nFind = FtpFindFirstFile(nConnection, "*.*", nData, 0, 0)
'               If nFind <> 0 Then
'                  Do
'                     stFileName = Left(nData.cFileName, InStr(1, nData.cFileName, String(1, 0), vbBinaryCompare) - 1)
'                     If stFileName <> "." And stFileName <> ".." Then
'                         If InStrRev(stFileName, ".") > 0 And UCase(stFileName) <> "THUMBS.DB" Then
'                              strNewFName = Mid(stFileName, 1, InStrRev(stFileName, ".") - 1) & "_old" & Mid(stFileName, InStrRev(stFileName, "."))
'                         Else
'                              strNewFName = stFileName
'                         End If
'                         If FtpRenameFile(nConnection, stFileName, pNewDir & "/" & strNewFName) <> 1 Then
'                            pErrMsg = "檔案 " & stFileName & " 移動失敗！"
'                            GoTo OutPort
'                         End If
'                     End If
'                     LRet = InternetFindNextFile(nFind, nData)
'                  Loop While LRet <> 0
'                  InternetCloseHandle nFind
'                  nFind = 0
'               End If
'            Else
'               pErrMsg = "FTP目錄 " & pNewDir & " 切換失敗！"
'               GoTo OutPort
'            End If
'       End If
'   End If
''------------------------------------
''將來源資料夾的所有檔案搬移到目的資料夾後,刪除來源資料夾
'   If FtpSetCurrentDirectory(hConnection, pOldDir) = 1 Then
'      pData.cFileName = String(MAX_PATH, 0)
'      hFind = FtpFindFirstFile(hConnection, "*.*", pData, 0, 0)
'      If hFind <> 0 Then
'         Do
'            stFileName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
'            strNewFName = "" 'Added by Lydia 2020/02/13
'            If stFileName <> "." And stFileName <> ".." Then
'               If pChgName = "" Then
'                  strNewFName = stFileName
'               Else
'                    If Right(UCase(stFileName), Len(FcpTcnFKey02)) = FcpTcnFKey02 Then   '外文原文本
'                          'Modified by Lydia 2018/03/14 拼錯
'                          'strNewFName = pChgName & ".ForignSpec" & FcpTcnFKey02
'                          'Modified by Lydia 2018/03/16 會議後通知,拿掉ForeignSpec
'                          'strNewFName = pChgName & ".ForeignSpec" & FcpTcnFKey02
'                          strNewFName = pChgName & FcpTcnFKey02
'                    'Added by Lydia 2020/02/13 郵件：*.msg檔，統一為系統日期+時間
'                    ElseIf Right(UCase(stFileName), Len(FcpTcnFKey01)) = UCase(FcpTcnFKey01) And strSrvDate(1) >= XY特殊權限啟用日by檔案 Then
'                          cMsg = cMsg + 1
'                          If strTmp = "" Then strTmp = Format(ServerTime, "000000") '時間
'                          '避免重覆檔名
'                          If strTmp2 = strTmp Then
'                              strTmp = Format(Val(strTmp) + 1, "000000")
'                          End If
'                          strNewFName = pChgName & "." & strSrvDate(1) & strTmp & ".rx" & FcpTcnFKey01
'                          strTmp2 = strTmp
'                    'end 2020/02/13
'                    ElseIf UCase(stFileName) <> "THUMBS.DB" Then  '跳過Thumbs.db會造成無法刪來源資料夾
'                         strNewFName = pChgName & "." & stFileName
'                    Else
'                         strNewFName = stFileName '隱藏的系統檔
'                    End If
'               End If
'
'                'Added by Lydia 2020/02/13 English_Vers檔案：判斷啟用日
'                If strSrvDate(1) >= XY特殊權限啟用日by檔案 Then
'                    If stFileName <> "" And stFileName <> "." And stFileName <> ".." And UCase(stFileName) <> "THUMBS.DB" Then
'                       '改成先下載到本機端暫存區
'                        FileCopy strOldPath & "\" & stFileName, pNewDir & "\" & strNewFName
'                        pMsgList = pMsgList & IIf(pMsgList <> "", "&", "") & strNewFName '檔案全部記錄
'                    End If
'                    If stFileName <> "." And stFileName <> ".." Then
'                       If FtpDeleteFile(hConnection, stFileName) <> 1 Then '刪除FTP端的檔案
'                          'MsgBox stDir & "/" & stFileName & " 檔案刪除失敗！", vbCritical
'                          pErrMsg = "檔案 " & stFileName & " 檔案刪除失敗！"
'                          GoTo OutPort
'                       End If
'                    End If
'                Else
'                'end 2020/02/13
'                    If FtpRenameFile(hConnection, stFileName, pNewDir & "/" & strNewFName) <> 1 Then
'                       pErrMsg = "檔案 " & stFileName & " 移動失敗！"
'                       GoTo OutPort
'                    End If
'                    '記錄*.MSG檔的名稱
'                    If Right(UCase(strNewFName), Len(FcpTcnFKey01)) = UCase(FcpTcnFKey01) Then
'                         'Modified by Lydia 2018/03/06 改用&
'                         'pMsgList = pMsgList & IIf(pMsgList <> "", ";", "") & strNewFName
'                         pMsgList = pMsgList & IIf(pMsgList <> "", "&", "") & strNewFName
'                    End If
'                End If 'end 2020/02/13
'            End If
'            LRet = InternetFindNextFile(hFind, pData)
'         Loop While LRet <> 0
'         InternetCloseHandle hFind
'         hFind = 0
'
'         If FtpRemoveDirectory(hConnection, pOldDir) <> 1 Then
'            pErrMsg = pOldDir & "舊目錄刪除失敗！"
'            GoTo OutPort
'         End If
'      End If
'   Else
'      pErrMsg = "FTP目錄 " & pOldDir & " 切換失敗！"
'      GoTo OutPort
'   End If
'   PUB_FtpRenTyping2 = True
'
'OutPort:
'   If Err.Number <> 0 Then pErrMsg = Err.Number & vbCrLf & Err.Description
'   If hConnection <> 0 Then InternetCloseHandle hConnection
'   If hFind <> 0 Then InternetCloseHandle hFind
'
'   If PUB_FtpRenTyping2 = False And pRaiseErr = True Then
'      Err.Raise 999, , pErrMsg
'   End If
'
'End Function
'end 2023/10/12

'Added by Lydia 2018/03/01 FTP目錄是否存在
'Modified by Lydia 2018/03/08 +pMode 後續處理, pKey 檔案類型, pList 回傳結果
'Move by Lydia 2023/02/15 從basUpdate搬過來
Public Function PUB_ChkFtpDirectory(pFtpIp As String, pFtpPath As String, _
                    Optional ByVal pMode As String, Optional ByVal pKey As String, Optional ByRef pList As String) As Boolean
Dim pErrMsg As String, pRaiseErr As Boolean
Dim arrDir() As String, ii As Integer
Dim hDir As Long
Dim pConnection As Long
 'Added by Lydia 2018/03/08
Dim midList As String
Dim pData As WIN32_FIND_DATA
Dim hFind As Long, LRet  As Long, stFileName As String
'end 2018/03/08

On Error GoTo OutPort

   PUB_ChkFtpDirectory = False
   pList = "" 'Added by Lydia 2018/03/08
   
   pConnection = PUB_GetFtpConnect(pErrMsg, , , pFtpIp)
   
   If pFtpIp = "" Or pFtpPath = "" Or pConnection = 0 Then Exit Function
   
   arrDir = Split(pFtpPath, "/")
   For ii = LBound(arrDir) To UBound(arrDir)
      If arrDir(ii) <> "" Then
         PUB_ChkFtpDirectory = True
         If FtpSetCurrentDirectory(pConnection, arrDir(ii)) <> 1 Then
            PUB_ChkFtpDirectory = False
            'Modified by Lydia 2018/03/12 切斷連線
            'Exit Function
            GoTo OutPort
         End If
      End If
   Next
   
   'Added by Lydia 2018/03/08 針對FTP資料夾內的檔案做處理
   If pMode <> "" And pKey <> "" Then
        pData.cFileName = String(MAX_PATH, 0)
        hFind = FtpFindFirstFile(pConnection, "*.*", pData, 0, 0)
        If hFind <> 0 Then
           Do
              stFileName = Left(pData.cFileName, InStr(1, pData.cFileName, String(1, 0), vbBinaryCompare) - 1)
              If stFileName <> "." And stFileName <> ".." Then
                    If pMode = "R" Then
                        If InStrRev(stFileName, ".") > 0 And UCase(stFileName) <> "THUMBS.DB" Then
                            If pKey <> "*.*" Then
                                 'Modified by Lydia 2018/11/19 傳入本所案號.xxxx
                                 'If Right(UCase(stFileName), Len(pKey)) = UCase(pKey) Then
                                 If Left(pKey, 1) <> "." And Right(pKey, 1) = "." Then
                                        If Left(UCase(stFileName), Len(pKey)) = UCase(pKey) Then
                                            midList = midList & stFileName & "&"
                                        End If
                                 ElseIf Left(pKey, 1) = "." And Right(pKey, 1) = "." Then '傳入副檔名(.XXX.)
                                        If InStr(UCase(stFileName), UCase(pKey)) > 0 Then
                                            midList = midList & stFileName & "&"
                                        End If
                                 ElseIf Right(UCase(stFileName), Len(pKey)) = UCase(pKey) Then '傳入副檔名(.XXX)
                                 'end 2018/11/19
                                        midList = midList & stFileName & "&"
                                 End If
                            Else
                                 midList = midList & stFileName & "&"
                            End If
                        End If
                    ElseIf pMode = "D" Then
                        If InStrRev(stFileName, ".") > 0 Then
                            If pKey <> "*.*" Then
                                'Modified by Lydia 2018/11/19 傳入本所案號.xxxx
                                'If Right(UCase(stFileName), Len(pKey)) = UCase(pKey) Then
                                If Left(pKey, 1) <> "." And Right(pKey, 1) = "." Then
                                    If Left(UCase(stFileName), Len(pKey)) = UCase(pKey) Then
                                        If FtpDeleteFile(pConnection, stFileName) <> 1 Then
                                            midList = midList & stFileName & "檔案刪除失敗！" & "&"
                                            GoTo OutPort
                                         Else
                                            midList = midList & stFileName & "&"
                                         End If
                                    End If
                                ElseIf Left(pKey, 1) = "." And Right(pKey, 1) = "." Then '傳入副檔名(.XXX.)
                                    If InStr(UCase(stFileName), UCase(pKey)) > 0 Then
                                        If FtpDeleteFile(pConnection, stFileName) <> 1 Then
                                            midList = midList & stFileName & "檔案刪除失敗！" & "&"
                                            GoTo OutPort
                                         Else
                                            midList = midList & stFileName & "&"
                                         End If
                                    End If
                                ElseIf Right(UCase(stFileName), Len(pKey)) = UCase(pKey) Then '傳入副檔名(.XXX)
                                'end 2018/11/19
                                    If FtpDeleteFile(pConnection, stFileName) <> 1 Then
                                        midList = midList & stFileName & "檔案刪除失敗！" & "&"
                                        GoTo OutPort
                                     Else
                                        midList = midList & stFileName & "&"
                                     End If
                                End If
                            Else
                                If FtpDeleteFile(pConnection, stFileName) <> 1 Then
                                   midList = midList & stFileName & "檔案刪除失敗！" & "&"
                                   GoTo OutPort
                                Else
                                   midList = midList & stFileName & "&"
                                End If
                            End If
                        End If
                    End If
              End If
              LRet = InternetFindNextFile(hFind, pData)
           Loop While LRet <> 0
           
           '刪除FTP資料夾
           If pMode = "D" And pKey = "*.*" Then
               If FtpRemoveDirectory(pConnection, pFtpPath) <> 1 Then
                  midList = midList & pFtpPath & "目錄刪除失敗！" & "&"
                  GoTo OutPort
               Else
                  midList = midList & pFtpPath & "目錄刪除成功" & "&"
               End If
           End If
           
           InternetCloseHandle hFind
           hFind = 0
        End If
   End If
   'end 2018/03/08
   
OutPort:
   If Err.Number <> 0 Then midList = midList & Err.Description & "&"
   If pMode <> "" And pKey <> "" Then pList = midList 'Added by Lydia 2018/03/08
   If pConnection <> 0 Then InternetCloseHandle (pConnection)
   If hFind <> 0 Then InternetCloseHandle hFind

End Function

'Added by Lydia 2023/02/15 外專新案認領：更新狀態
'Modified by Lydia 2023/05/05 +bolErrMsg 是否彈錯誤訊息
Public Function PUB_UpdateTCNstate(ByVal pStatus As String, ByVal pCaseNo As String, Optional ByVal bolErrMsg As Boolean = True) As Boolean
'pCaseNo: 本所案號
'pStatus: 1-櫃台收文, 2-工程師主管認領
Dim intA As Integer, strA1 As String, strA2 As String, strA3 As String
Dim strTo As String, strCC As String
Dim strBCase(1 To 4) As String
Dim rsAD As New ADODB.Recordset
Dim strUpd  As String, m_CP09 As String, m_CP10 As String, m_CP157 As String
Dim m_Grp As String, m_GrpMan As String '工程師組別／主管
Dim m_CP27 As String, m_CP05 As String, m_CP06 As String, m_CP14 As String
Dim m_PA09 As String, m_PA150 As String, m_PA10 As String
Dim strEDate As String, strETime As String
Dim strSpecSub As String
Dim strCont As String 'Added by Lydia 2023/05/23 固定Email內文
Dim rsRD As New ADODB.Recordset  'Added by Lydia 2023/06/14

'*****************2023/06/14整理
'可以更新認領狀態的：櫃台收文(新案立卷)同時處理、新案發文時觸發PUB_UpdateTCNstate、非英說案件在客戶提供文件frm060120確收文件(TCN13)/取消暫不認領(TCN16)
'TCN22 最後認領期限-時間, TCN23 認領期限狀態
'0=急件認領期限0.5h
'1=主管期限2h
'2=通知職代認領+1h
'3=協調認領期限2h(2組以上認領之協調)
'4=非英說認領期限2h(確收文件TCN13=3,4並且判斷是否進入第二認領階段TCN23=4)
'
'認領期限狀態的分類原則如下:
'1.暫不認領TCN16=N =>TCN21=99999999, 英文組和日文組都有暫不認領
'2.103設計申請+125衍生設計會自動分案給機械設計組(TCN20預設4)=>PA150,TCT04(主管),TCN20='4'並且發email; 相似舊案指定組別做同樣處理
'3.日文組案件TCN19=Null=>PA150,TCT04(主管),TCN20='3'並且發email
'4.提申急件預設組別CU154=> PA150,TCT04(主管),TCN20=CU154並且發email
'  4.1 提申急件（指定期限或法限為當日，並且於下午兩點後才收文之新案）。
'  4.2 取消暫不認領TCN16會發生「當日以前的收文而提申期限為當日者」
'  4.3 (5/3會議) 命名作業之(名稱)譯畢期限列入判斷：收文2小時內(包含午休)設為急件。

'5.非上述發認領通知email =>TCN21=系統日,TCN22=PUB_CompWorkTime(系統時間,分), TCN23=0/1,預設逾期通知次數TCN25=0
'   P.S. TCN20=null, TCN21~TCN23控制期限；休息時段沒有設定變數from Sindy
'
'其他
'6.FMP案可以省略輸入工程師組別，因為和frm040101_1共用模組PUB_SavePtoUpd2,PUB_SavePtoUpd4
'   2/15 Phoebe: FMP案仍送樓上補基本資料,可以省略輸入工程師組別
'7.發文: 新案發文時觸發PUB_UpdateTCNstate，急件翻譯名稱在提申後重新進入認領
'  非英說案件在客戶提供文件frm060120確收文件(TCN13)觸發PUB_UpdateTCNstate/取消暫不認領(TCN16)
'
'批次自動:
'  1.固定每5分鐘檢查案件認領期限(在5分鐘內)，若為”主管認領狀態”TCN23=1逾期，發職代Email並且改狀態為”職代認領狀態”TCN23=2。
'  2.認領逾期通知TCN25: 第一次在逾通知期限3小時通知，第二次在逾建檔日期１天通知
'　3.每個工作日下午２點(14:00)若有前日未核判之新案(非提申急件)，由系統寄email提醒通知國外部最高主管進行核判。
'*****************

   PUB_UpdateTCNstate = True
   Call ChgCaseNo(pCaseNo, strBCase)
   
   'Modified by Lydia 2023/05/23 帶出代理人PA75N和申請人名稱PA26N
   strA1 = "select s1.st16,cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp14,cp66,cp67,cp157,cp158,pa09,pa10,pa150,tct04,tct02,tct03,cu154," & _
              "PA75,NVL(FA05,NVL(FA04,FA06)) PA75N,PA26,NVL(CU05,NVL(CU04,CU06)) PA26N, " & _
              "W.* from caseprogress,trackingcasename W,transcasetitle,patent,customer,staff s1,fagent " & _
              "where cp01='" & strBCase(1) & "' and cp02='" & strBCase(2) & "' and cp03='" & strBCase(3) & "' and cp04='" & strBCase(4) & "' " & _
               "and cp31='Y' and cp159=0 and cp09=tcn05(+) and tcn05 is not null " & _
               "and cp09=tct01(+) and tct01 is not null and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
               "and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and tcn03=s1.st01(+) and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) "
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strA1)
   If intA = 1 Then
      m_CP05 = "" & rsAD.Fields("cp05")
      m_CP06 = "" & rsAD.Fields("cp06")
      m_CP09 = "" & rsAD.Fields("cp09")
      m_CP10 = "" & rsAD.Fields("cp10")
      m_CP14 = "" & rsAD.Fields("cp14")
      m_CP27 = "" & rsAD.Fields("cp158")
      m_CP157 = "" & rsAD.Fields("cp157") '北所分案日
      m_PA09 = "" & rsAD.Fields("pa09")
      m_PA10 = "" & rsAD.Fields("pa10")
      m_PA150 = "" & rsAD.Fields("pa150")
      'Added by Lydia 2023/05/23 固定Email內文
      If "" & rsAD.Fields("tct02") & rsAD.Fields("tct03") <> "" Then
           strCont = "譯畢期限：" & ChangeWStringToTDateString(rsAD.Fields("tct02")) & " " & Format(rsAD.Fields("tct03"), "00:00") & vbCrLf & vbCrLf
      End If
      strCont = strCont & "代理人：" & IIf("" & rsAD.Fields("PA75") <> "", rsAD.Fields("PA75") & " " & rsAD.Fields("PA75N"), "（空白）") & vbCrLf & _
                    "申請人：" & IIf("" & rsAD.Fields("PA26") <> "", rsAD.Fields("PA26") & " " & rsAD.Fields("PA26N"), "（空白）") & vbCrLf
      '2023/05/23
      'Modified by Lydia 2023/05/26 Email主旨開頭改成模組
      strSpecSub = PUB_GetTCNmTitle(rsAD.Fields("cp01"), rsAD.Fields("cp02"), rsAD.Fields("cp03"), rsAD.Fields("cp04"), m_PA10, "" & rsAD.Fields("tcn13"), "SPEC")
      
      '排除程序人員在新案建檔設定
      If "" & rsAD.Fields("tcn23") = "9" Then
          GoTo JumpToExec
      End If
      '(共用)工程師主管認領
      If pStatus = "2" And "" & rsAD.Fields("tcn20") <> "" Then
         m_Grp = "" & rsAD.Fields("tcn20")
         GoTo JumpToExec
      End If

      strUpd = ""
      m_Grp = "": m_GrpMan = ""
      '1.暫不認領TCN16=N =>TCN21=99999999, 英文組和日文組都有暫不認領
      If "" & rsAD.Fields("tcn16") = "Y" Then
         strUpd = "Update trackingcasename Set TCN21=99999999 where tcn01='" & rsAD.Fields("tcn01") & "' "
         cnnConnection.Execute strUpd
         GoTo JumpToExec
      End If
      '3.日文組案件TCN19=Null=>PA150,TCT04(主管),TCN20='3'並且發email
      If "" & rsAD.Fields("st16") = "2" And "" & rsAD.Fields("tcn19") <> "Y" Then
          m_Grp = "3"
          'Modified by Lydia 2023/06/14 +, tcn25=0
          strUpd = "Update trackingcasename Set tcn20=" & CNULL(m_Grp) & ", tcn25=0 where tcn01='" & rsAD.Fields("tcn01") & "' "
          cnnConnection.Execute strUpd
          GoTo JumpToExec
      End If
      'Move by Lydia 2023/08/28 從上面移下來：日文組案件不需要英文組認領時，一律都給日文組
      '2.設計申請103會自動分案給機械設計組(TCN20預設4)=>PA150,TCT04(主管),TCN20='4'並且發email
      '＋相似舊案指定組別
      'Modified by Lydia 2024/02/20 +105: 有關" 香港013專利開放收文集體設計申請105"，請比照" 香港013專利收文設計申請103"
      If InStr("103,125,105", "" & rsAD.Fields("cp10")) > 0 Or "" & rsAD.Fields("tcn18") <> "" Then
         'Memo by Lydia 2024/02/27 外專機械設計組人員異動調整程式：新案認領組別，請取消機械設計組，案件歸電子組; 由Wilson代機械組主管T1
         m_Grp = IIf("" & rsAD.Fields("tcn18") <> "", "" & rsAD.Fields("tcn18"), "4")
         
         'Modified by Lydia 2023/06/14 +, tcn25=0
         strUpd = "Update trackingcasename Set tcn20=" & CNULL(m_Grp) & ", tcn25=0 where tcn01='" & rsAD.Fields("tcn01") & "' "
         cnnConnection.Execute strUpd
         GoTo JumpToExec
      End If
      '4.提申急件預設組別CU154=> PA150,TCT04(主管),TCN20=CU154並且發email
         '4.1 提申急件（指定期限或法限為當日，並且於下午兩點後才收文之新案）。
         '4.2 取消暫不認領TCN16會發生「當日以前的收文而提申期限為當日者」
         '4.3 (5/3會議) 命名作業之(名稱)譯畢期限列入判斷：收文2小時內(包含午休)設為急件。
      strA1 = ""
      If Val(m_PA10) = 0 Then
         If strSrvDate(1) >= "" & rsAD.Fields("cp66") Then '取消暫不認領TCN16，所以用系統日判斷是否急件
             strA2 = strSrvDate(1): strA3 = Left(Format(ServerTime, "000000"), 4)
         Else
             strA2 = "" & rsAD.Fields("cp66"): strA3 = "" & rsAD.Fields("cp67")
         End If
         '判斷(名稱)譯畢期限
         If "" & rsAD.Fields("tcn20") = "" And "" & rsAD.Fields("tct02") = strA2 And "" & rsAD.Fields("tct03") <> "" Then
             strA1 = Abs(DateDiff("n", Format(rsAD.Fields("tct03") & "00", "00:00:00"), Format(strA3 & "00", "00:00:00")))
         End If
      End If
      '所限/法限當天或過期
      If Val(m_PA10) = 0 And "" & rsAD.Fields("tcn20") = "" And ((((m_CP06 <> "" And strA2 >= m_CP06) Or ("" & rsAD.Fields("cp07") <> "" And strA2 >= "" & rsAD.Fields("cp07"))) And strA3 >= "1400") _
          Or (Val(strA1) > 0 And Val(strA1) < 120)) Then
          If "" & rsAD.Fields("cu154") <> "" Then
              m_Grp = "" & rsAD.Fields("cu154")
              m_GrpMan = Pub_GetFCPGrpMan(m_Grp)
              m_GrpMan = PUB_GetStateForMan(m_GrpMan) '特殊情況之指定職代
              'Modified by Lydia 2023/06/14 +, tcn25=0
              strUpd = "Update trackingcasename Set tcn20=" & CNULL(m_Grp) & ", tcn25=0 where tcn01='" & rsAD.Fields("tcn01") & "' "
              cnnConnection.Execute strUpd
              strA1 = Replace(strSpecSub, "SPEC", "急件") & "，請協助翻譯發明名稱以先進行提申，謝謝！"
              'Modified by Lydia 2023/05/23 固定Email內文strCont
              strUpd = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                           " values( '" & strUserNum & "','" & m_GrpMan & "',to_char(sysdate,'yyyymmdd')" & _
                           ",to_char(sysdate,'hh24miss'),'" & strA1 & "','" & ChgSQL(strCont) & "')"
              cnnConnection.Execute strUpd
              GoTo JumpToExec
          Else
              Call PUB_CompWorkTime(ServerTime, 30, strETime, strSrvDate(1), strEDate)
              If strEDate <> "" And strETime <> "" Then
                 strUpd = "Update trackingcasename Set tcn23='0', tcn21=" & CNULL(strEDate) & ", tcn22=" & CNULL(Left(strETime, 4)) & " where tcn01='" & rsAD.Fields("tcn01") & "' "
                 cnnConnection.Execute strUpd
                 '急件認領通知Email主旨：主旨：新案急件FCP0*****/P******，請協助確認組別，謝謝！；內文：無。和非急件認領通知Email只加註急件，若半小時內未有組別認領，由系統自動發EMAIL通知最高主管和職代。整列資料顯示為紅色。
                 strTo = PUB_GetEngGrpMan(strCC)
                 If strTo <> "" Then
                    strA1 = Replace(strSpecSub, "SPEC", "急件") & "，請協助確認組別，謝謝！"
                    'Modified by Lydia 2023/05/23 固定Email內文strCont
                    strUpd = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                             " values( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
                             ",to_char(sysdate,'hh24miss'),'" & strA1 & "','" & ChgSQL(strCont) & "','" & strCC & "')"
                    cnnConnection.Execute strUpd
                 End If
                 GoTo JumpToExec
              End If
          End If
      End If
      '5.非上述發認領通知email =>TCN21=系統日,TCN22=PUB_CompWorkTime(系統時間,分), TCN23=0/1,預設逾期通知次數TCN25=0
      '   P.S. TCN20=null, TCN21~TCN22控制期限；休息時段沒有設定變數from Sindy
      'Modified by Lydia 2023/06/14 +非英說案件=4
      'If Val("" & rsAD.Fields("tcn23")) = 0 Or "" & rsAD.Fields("tcn23") = "1" Or "" & rsAD.Fields("tcn23") = "2" Then
      '   strA3 = Val("" & rsAD.Fields("tcn23")) + 1
      If Val("" & rsAD.Fields("tcn23")) = 0 Or "" & rsAD.Fields("tcn23") = "1" Or "" & rsAD.Fields("tcn23") = "2" Or "" & rsAD.Fields("tcn23") = "4" Then
         If "" & rsAD.Fields("tcn23") <> "4" Then
            strA3 = Val("" & rsAD.Fields("tcn23")) + 1
         Else
            strA3 = Val("" & rsAD.Fields("tcn23"))
         End If
      'end 2023/06/14
         Select Case strA3
             Case "1" '1=主管期限2h
                Call PUB_CompWorkTime(ServerTime, 120, strETime, strSrvDate(1), strEDate)
             Case "2" '2=通知職代認領+1h
                Call PUB_CompWorkTime("" & rsAD.Fields("tcn22"), 60, strETime, "" & rsAD.Fields("tcn21"), strEDate)
             Case "3", "4" '3=協調認領期限2h、4=非英說案件協調認領2h(2023/06/14)
                Call PUB_CompWorkTime(ServerTime, 120, strETime, strSrvDate(1), strEDate)
         End Select
         If strEDate <> "" And strETime <> "" Then
            'Modified by Lydia 2023/06/14 +, tcn25=0
            strUpd = "Update trackingcasename Set tcn23='" & strA3 & "', tcn21=" & CNULL(strEDate) & ", tcn22=" & CNULL(Left(strETime, 4)) & ", tcn25=0 where tcn01='" & rsAD.Fields("tcn01") & "' "
            cnnConnection.Execute strUpd
            '認領通知Email主旨：新案FCP0*****/P******，請協助確認組別，謝謝！；內文：無。
            'P.S.現有新案立卷Email通知EMAIL承辦及程序時(PUB_GetTCTmail)，加註是否已進入認領階段。事後不必再通知。
            If strA3 = "1" Or strA3 = "2" Then
               strTo = PUB_GetEngGrpMan(strCC)
               If strTo <> "" Then
                  strA1 = Replace(strSpecSub, "SPEC", "") & "，請協助確認組別，謝謝！"
                  'Modified by Lydia 2023/05/23 固定Email內文strCont
                  strUpd = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                           " values( '" & strUserNum & "','" & IIf(strA3 = "1", strTo, strCC) & "',to_char(sysdate,'yyyymmdd')" & _
                           ",to_char(sysdate,'hh24miss'),'" & strA1 & "','" & ChgSQL(strCont) & "')"
                  cnnConnection.Execute strUpd
               End If
            'Added by Lydia 2023/06/14 非英說案件：第一階段認領=Y
            ElseIf strA3 = "4" Then
               '有經過認領: 在PUB_UpdateReTCN已先新增記錄
               strA1 = "select st16 from transfeeassign,staff where tfa01='" & rsAD.Fields("tcn05") & "' and tfa04=st01(+) and tfa09='4' order by st16 "
               intA = 1
               strTo = ""
               Set rsRD = ClsLawReadRstMsg(intA, strA1)
               If intA = 1 Then
                  strA1 = rsRD.GetString(adClipString, , , ",")
                  strTo = PUB_GetEngGrpMan(strCC, strA1) '第一階段認領=Y的組別
               Else '沒有經過認領: 急件有預設組別CU154
                  strTo = PUB_GetEngGrpMan(strCC)  '全英文組
               End If
               If strTo <> "" Then
                  strA1 = Replace(strSpecSub, "SPEC", "") & "，請協調以確認組別，謝謝！"
                  strUpd = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                               " values( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
                               ",to_char(sysdate,'hh24miss'),'" & strA1 & "','" & ChgSQL(strCont) & "')"
                  cnnConnection.Execute strUpd
               End If
            'end 2023/06/14
            End If
            GoTo JumpToExec
         End If
      End If
   End If
   
JumpToExec:
   If m_Grp <> "" Then
       If m_GrpMan = "" Then m_GrpMan = Pub_GetFCPGrpMan(m_Grp)
       m_GrpMan = PUB_GetStateForMan(m_GrpMan) '特殊情況之指定職代
       '更新基本檔
       strUpd = "Update Patent Set PA150='" & m_Grp & "' Where PA01='" & strBCase(1) & "' and PA02='" & strBCase(2) & "' and PA03='" & strBCase(3) & "' and PA04='" & strBCase(4) & "' "
       Pub_SeekTbLog strUpd, , , , "外專新案認領(PUB_UpdateTCNstate)"
       cnnConnection.Execute strUpd
       'FMP的處理: 參考frm040101_1
       If strBCase(1) = "P" Then
           'Mark by Lydia 2023/06/06 debug: FMP案只預設中說性質(工程師承辦)PUB_SavePtoUpd4，新申請案和其他性質仍交由分案人員進行分案作業; ex.P-131648
           'strUpd = "Update CaseProgress Set CP14=" & CNULL(m_GrpMan) & IIf(m_CP157 = "", ", CP157=" & strSrvDate(1), "") & " Where CP09=" & CNULL(m_CP09)
           'cnnConnection.Execute strUpd
           'end 2023/06/06
           '比照PUB_GetPcm10，香港案=4，澳門案=5
           Call PUB_SavePtoUpd2(True, strBCase, m_CP14, m_CP09, m_CP10, m_CP05, m_CP06, m_CP27, IIf(m_PA09 = "013", "4", IIf(m_PA09 = "044", "5", "")))
           Call PUB_SavePtoUpd4(strBCase, m_PA09, m_Grp, m_PA150, "Y", m_CP09, m_CP10, m_GrpMan, m_CP14, , , pStatus)
       Else
           Call ChangeTCTGrp(pStatus, strBCase, m_CP09, m_Grp, m_PA150)
       End If

   End If
   Set rsAD = Nothing
   Set rsRD = Nothing 'Added by Lydia 2023/06/14
   
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
       PUB_UpdateTCNstate = False
       cnnConnection.RollbackTrans
       If bolErrMsg = True Then
          MsgBox Err.Description & vbCrLf & strUpd, vbCritical, "外專新案認領：更新狀態"
       End If
   End If

End Function

'Added by Lydia 2023/02/16 傳入時間+日期，增加n分鐘，取得符合工時的新時間
Public Sub PUB_CompWorkTime(ByVal pSTime As String, ByVal pAddNS As Integer, ByRef pETime As String, Optional ByVal pSDate As String, Optional ByRef pEDate As String)
'pEDate: 格式yyyymmdd
'pETime: 格式hhnnss
Dim tmpDate As String
Dim diffAdd As Integer
Dim strTime As String

    If pSDate = "" Then pSDate = strSrvDate(1)
    pSTime = Replace(pSTime, ":", "")
    If Len(pSTime) > 4 Then
       pSTime = Left(Format(pSTime, "000000"), 4)
    Else
       pSTime = Format(pSTime, "0000")
    End If
    tmpDate = Format(DateAdd("n", pAddNS, ChangeTStringToTDateString(pSDate) & " " & Format(pSTime & "00", "00:00:00")), "yyyymmdd hhnnss")
    'Modified by Lydia 2023/05/31 含午休時段的判斷; ex.FCP-069520在中午12點取消暫不認領
    'If pSTime < "1200" And pAddNS > 0 Then
    If ((pSTime >= "1200" And pSTime <= "1330") Or pSTime < "1200") And pAddNS > 0 Then
       '跨午休
       If Left(tmpDate, 8) = pSDate And Mid(tmpDate, 10) > "120000" Then   '跨午休
           tmpDate = Format(DateAdd("n", pAddNS + 90, ChangeTStringToTDateString(pSDate) & " " & Format(pSTime & "00", "00:00:00")), "yyyymmdd hhnnss")
       End If
    End If
    If pSTime >= "1200" And pAddNS > 0 Then
       '跨日
       If Left(tmpDate, 8) > pSDate Or Mid(tmpDate, 10) > "180000" Then
           diffAdd = DateDiff("n", "18:00:00", Format(Mid(tmpDate, 10), "00:00:00"))
           tmpDate = Format(DateAdd("n", diffAdd, ChangeTStringToTDateString(CompWorkDay(2, pSDate)) & " 09:00:00"), "yyyymmdd hhnnss")
       End If
    End If
    pEDate = Left(tmpDate, 8)
    pETime = Mid(tmpDate, 10)
End Sub

'Added by Lydia 2023/02/23 取得外專英文組工程師主管和職代
'Modified by Lydia 2024/02/27 外專機械設計組人員異動調整程式：新案認領組別，請取消機械設計組，只留電子電機組及化學組
'Public Function PUB_GetEngGrpMan(Optional ByRef pABS0102 As String, Optional ByVal pGrpList As String = "1,2,4") As String
Public Function PUB_GetEngGrpMan(Optional ByRef pABS0102 As String, Optional ByVal pGrpList As String = "1,2") As String
'PUB_GetEngGrpMan:工程師組別之主管
'pABS0102: 主管的第一職代和第二職代
Dim tmpArr As Variant
Dim intB As Integer
Dim strB00 As String, strB01 As String, strB02 As String, strB03 As String
   
   PUB_GetEngGrpMan = ""
   pABS0102 = ""
   If pGrpList <> "" Then
      tmpArr = Split(pGrpList, ",")
      For intB = 0 To UBound(tmpArr)
         If Trim(tmpArr(intB)) <> "" Then
            strB00 = Pub_GetFCPGrpMan(tmpArr(intB))
            If strB00 <> "" Then
                Call GetABS001_CaseSys(strB00, strB01, strB02, strB03, "")
                PUB_GetEngGrpMan = PUB_GetEngGrpMan & ";" & strB00
                pABS0102 = pABS0102 & IIf(strB01 <> "", ";" & strB01, "") & IIf(strB02 <> "", ";" & strB02, "") & IIf(strB03 <> "", ";" & strB03, "")
            End If
         End If
      Next intB
      If PUB_GetEngGrpMan <> "" Then PUB_GetEngGrpMan = Mid(PUB_GetEngGrpMan, 2)
      If pABS0102 <> "" Then pABS0102 = Mid(pABS0102, 2)
   End If
End Function

'Move by Lydia 2023/02/21 從basPublic 搬過來;
'Modified by Lydia 2023/02/21 tCP14 As Object=> tCP14 As String
'Added by Lydia 2016/07/07 內專分案(工程師主管分案)-寰華案相關和核稿人
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Sub PUB_SavePtoUpd2(ByVal bFMP As Boolean, ByRef mpa() As String, ByRef tCP14 As String, ByVal mCP09 As String, ByVal mCP10 As String, ByVal mCP05 As String, ByVal mCP06 As String, ByVal mCP27 As String, ByVal InCM10 As String, Optional ByVal iEP04 As String)
Dim rsR1 As New ADODB.Recordset
Dim strR1 As String, strR2 As String
Dim inR As Integer

    If bFMP Then
       '除有大陸案關聯之香港或澳門案外,新申請案期限為收文日+1個月(沒期限才更新)
       If mCP27 = "" And mCP06 = "" And InStr(NewCasePtyList, mCP10) > 0 And Not (InCM10 = "4" Or InCM10 = "5") Then
          strR1 = PUB_GetDeadLine(DBDATE(mCP05), "", 2)
          'FMP 案非PCT且未主張優先權則不掛法限(所限=收文日+1個月)
          strR1 = PUB_GetWorkDay1(strR1, True)
          strR2 = "update caseprogress set cp06=" & Val(strR1) & _
             " where cp09='" & mCP09 & "' and cp06 is null"
          inR = 0
          cnnConnection.Execute strR2, inR
       End If

        '外專收文翻譯案件，若承辦人為外專工程師時核稿人設為自己
        If mCP10 = "201" Or mCP10 = "209" Or mCP10 = "235" Or mCP10 = "210" Or mCP10 = "942" Then
           '更新承辦期限、核稿期限
           '抓有所限的新案(若無法限時也會設所限)
           strR1 = "select cp05,cp07 from caseprogress where cp01='" & mpa(1) & "' and cp02='" & mpa(2) & "' and cp03='" & mpa(3) & "' and cp04='" & mpa(4) & "' and cp10 in ('101','102','103') and cp06>0"
           inR = 1
           Set rsR1 = ClsLawReadRstMsg(inR, strR1)
           If inR = 1 Then
              '承辦期限
              strR1 = PUB_GetDeadLine(rsR1.Fields("cp05"), "" & rsR1.Fields("cp07"), 3)
              strR2 = "update caseprogress set cp48=" & CNULL(strR1, True) & " where cp09='" & mCP09 & "'"
              cnnConnection.Execute strR2, inR
              
              '201 要預設核稿期限
              If mCP10 = "201" Then
                 '核稿期限
                 strR1 = PUB_GetDeadLine(rsR1.Fields("cp05"), "" & rsR1.Fields("cp07"), 4)
                 strR2 = "update engineerprogress set ep08=" & CNULL(strR1, True) & " where ep02='" & mCP09 & "'"
                 cnnConnection.Execute strR2, inR
              End If
           End If
           
           '201 要預設核稿人 =FCP
           'Remvoe by Lydia 2021/06/30 已不符合現況; 取消內專分案預設新案翻譯之核稿人, ex.P-126779的核稿人F5370
           'If tCP14 <> "" And mCP10 = "201" Then
           '     '改抓ST15='F21'的
           '     'Modified by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
           '     'modify by sonia 2021/1/26 再排除F4104及F4105
           '     strR1 = "select 1 from staff_idmap,staff where '" & tCP14 & "' in (sim01,sim02)" & _
           '        " and st01(+)=sim01 and ST15='F21' and st04='1' and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' "
           '     inR = 1
           '     Set rsR1 = ClsLawReadRstMsg(inR, strR1)
           '     If inR = 1 Then
           '        strR2 = "update engineerprogress set ep04='" & tCP14 & "' where ep02='" & mCP09 & "'"
           '        cnnConnection.Execute strR2, inR
           '     End If
           'End If
           'end 2021/06/30
        End If
        '審查意見或核駁修改承辦人時一併修改相關收文號之告代承辦人
        If tCP14 <> "" And (mCP10 = "1202" Or mCP10 = "1002") Then
           strR2 = "update caseprogress set  cp14='" & tCP14 & "' where cp43='" & mCP09 & "' and cp10='901' and cp27 is null"
           cnnConnection.Execute strR2, inR
        End If
    End If
    '工程師主管分案可設定核稿人
    If iEP04 <> "" And mCP10 <> "" And InStr(FCPHaveEP04, mCP10) > 0 Then
        strR2 = "update engineerprogress set ep04='" & iEP04 & "' where ep02='" & mCP09 & "'"
        cnnConnection.Execute strR2, inR
    End If
End Sub

'Added by Lydia 2023/02/21 內專分案作業frm040101_1：與外專新案認領：更新狀態PUB_UpdateTCNstate共用
Public Sub PUB_SavePtoUpd4(ByRef pCase() As String, ByVal pPA09 As String, ByVal nEngGrp As String, ByVal pEngGrp As String, _
              ByVal pCP31 As String, ByVal pCP09 As String, ByVal pCP10 As String, ByVal nCP14 As String, ByVal pCP14 As String, _
              Optional ByVal pPA05 As String, Optional ByVal pPA06 As String, Optional ByVal pStatus As String = "1")
Dim intB As Integer
Dim strB1 As String, strB2 As String, strB3 As String, strB4 As String
Dim rsBD As New ADODB.Recordset

   'Added by Morgan 2012/3/19 輸入工程師組別同時設定該案未分案之新案翻譯,製作中說,檢視中說,901告知代理人的承辦人為該組管制人
   If pEngGrp <> nEngGrp And nEngGrp <> "" Then
      'Modifed by Morgan 2012/7/4 +942,203
      'Modified by Morgan 2013/11/6 +235核對中說格式
      strB1 = "select cp09,cp10 from caseprogress where cp01='" & pCase(1) & "' and cp02='" & pCase(2) & "' and cp03='" & pCase(3) & "' and cp04='" & pCase(4) & "' and cp14 is null and cp57 is null and cp10 in ('201','209','235','210','901','942','203')"
      intB = 1
      Set rsBD = ClsLawReadRstMsg(intB, strB1)
      If intB = 1 Then
         With rsBD
         Do While Not .EOF
            'Modify by Amy 2015/04/07 自動上cp157 (北所分案日)
            'Modified by Morgan 2023/9/27 +判斷非寰華案不上北所分案日--玲玲
            'strB1 = "update caseprogress set cp14=(select oMan from SetSpecMan where OCODE=decode('" & nEngGrp & "','1','T','2','R','3','S','4','T1')),cp157=" & strSrvDate(1) & _
               " where cp09='" & .Fields("cp09") & "'"
            strB1 = "update caseprogress a set cp14=(select oMan from SetSpecMan where OCODE=decode('" & nEngGrp & "','1','T','2','R','3','S','4','T1'))" & _
               ",cp157=(select decode(count(*),0,a.cp157," & strSrvDate(1) & ") from caseprogress b where cp01=a.cp01 and cp02=a.cp02 and cp03=a.cp03 and cp04=a.cp04 and cp44='Y53374000' and cp31='Y')" & _
               " where cp09='" & .Fields("cp09") & "'"
            'end 2023/9/27
            cnnConnection.Execute strB1, intB
            
            'Modified by Morgan 2013/11/6 +235核對中說格式
            'Modified by Morgan 2015/3/11 只有 201 要預設核稿人及期限  --靜芳
            'If .Fields("cp10") = "201" Or .Fields("cp10") = "209" Or .Fields("cp10") = "235" Or .Fields("cp10") = "210" Or .Fields("cp10") = "942" Then
            If .Fields("cp10") = "201" Then
            'end 2015/3/11
               strB1 = "update engineerprogress set ep04=(select cp14 from caseprogress where cp09=ep02) where ep02='" & .Fields("cp09") & "'"
               cnnConnection.Execute strB1, intB
            End If
            .MoveNext
         Loop
         End With
      End If
      Set rsBD = Nothing
   End If
   'end 2012/3/19
   
   'Added by Lydia 2018/05/22 FMP案-命名作業分工程師組別,通知相關人員
   'Memo by Lydia 2023/02/21 為了與FCP共用，另外抽成模組ChangeTCTGrp
   Call ChangeTCTGrp(pStatus, pCase, pCP09, nEngGrp, pEngGrp, pPA05, pPA06)
   
   'Added by Lydia 2021/05/31 內專程序分案後，系統自動發email通知
   'Modified by Lydia 2021/11/17 排除香港案(確定要排除母案為寰華案衍生的香港案通知，且包含所有香港案收文之分案 by Phoebe)
   If pCP14 = "" And pCP14 <> nCP14 And pPA09 <> "013" Then
      'Modified by Lydia 2022/06/30 分案(307)改通知工程師
      'If pCP31 = "Y" Then '新案:通知程序, CC:程序主管(二級主管)
      If pCP31 = "Y" And pCP10 <> "307" Then
         strB1 = PUB_GetFCPHandler(pCase(1), pCase(2), pCase(3), pCase(4))
         strB2 = PUB_GetFCPProSup(strB1)
      Else  '中間程序:通知輸入的承辦人員
         strB1 = Trim(nCP14)
         strB3 = PUB_GetST03(strB1)
         'CC: 日文組工程師為副理+協理; 英文組工程師為副理; 其餘部門為人員主管(二級主管)
         If strB3 = "F21" Then
             strB2 = PUB_GetFCPEngSup(strB1, True)
             If nEngGrp = "3" Then
                 strB4 = Pub_GetSpecMan("S")
                 If InStr(strB1 & ";" & strB2, strB4) = 0 Then
                     strB2 = strB2 & IIf(strB2 <> "", ";", "") & strB4
                 End If
             End If
         Else
             'FMP案只通知工程師，寰華案全部都通知
             If PUB_GetFMP2toP(pCase(1), pCase(2), pCase(3), pCase(4)) = True Then
                 strB1 = ""
             Else
                 strB2 = PUB_GetFCPProSup(strB1)
             End If
         End If
      End If
      If strB1 <> "" Then
          'Added by Lydia 2022/06/30 分案(307)改通知工程師,並CC程序
          If pCP31 = "Y" And pCP10 = "307" Then
               strB3 = PUB_GetFCPHandler(pCase(1), pCase(2), pCase(3), pCase(4))
               strB2 = strB2 & ";" & strB3
          End If
          'end 2022/06/30
          intB = ClsPDGetCaseProperty(pCase(1), pCP10, strB4, IIf(pPA09 <> "000", True, False))
          strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
             " values( '" & strUserNum & "','" & strB1 & "',to_char(sysdate,'yyyymmdd')" & _
             ",to_char(sysdate,'hh24miss'),'" & pCase(1) & pCase(2) & IIf(pCase(3) <> "0", pCase(3), "") & IIf(pCase(4) <> "00", pCase(4), "") & "【" & strB4 & "已分案】請處理後續！" & "','同主旨','" & strB2 & "' ) "
          cnnConnection.Execute strSql
      End If
   End If
   'end 2021/05/31
End Sub

'Added by Lydia 2023/02/21 命名作業:分案／變更工程師組別,通知相關人員
Private Sub ChangeTCTGrp(ByVal pType As String, ByRef pCase() As String, ByVal pCP09 As String, ByVal nGRP As String, Optional ByVal pGrp As String, Optional ByVal pPA05 As String, Optional pPA06 As String)
'pType：2=工程師主管認領, 3=由原組別重新進行命名作業(2023/07/27 〔非英說案〕在第1次認領階段中沒有2組以上認領，等到提供”英說翻譯or簡繁體中說”雖然不需要【進入認領階段】，但仍需要原組別檢視案件名稱，所以由原組別重新進行命名作業。)
Dim strQ1 As String, strQ2 As String, strQ3 As String
Dim strUpd As String, strCP27 As String
Dim intQ As Integer
Dim rsRD As New ADODB.Recordset
'Added by Lydia 2023/05/04
Dim strB1 As String, intB As Integer, strCC As String
Dim rsBD As New ADODB.Recordset
Dim m_Context As String 'Added by Lydia 2024/04/23

   'Added by Lydia 2018/05/22 FMP案-命名作業分工程師組別,通知相關人員
   'Modifiedby Lydia 2021/04/08 判斷有改組別; ex.P-126940進來改AB0010843的承辦人,清空命名記錄
   'Modified by Lydia 2023/07/27 + 〔非英說案〕由原組別重新進行命名作業 +Or pType = "3"
   If (pGrp <> nGRP Or pType = "2" Or pType = "3") And pCP09 <> "" Then
      'Modified by Lydia 2023/05/04 +PA10,CP44
      'Modified by Lydia 2023/06/14 +TCN13
      'strQ1 = "Select cp158,pa05,pa06,PA10,CP44 from caseprogress, patent where cp09='" & pCP09 & "' and cp159=0 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
      'Modified by Lydia 2023/06/21 +CP47
      'Modified by Lydia 2024/04/23
      'strQ1 = "Select cp158,pa05,pa06,PA10,CP44,tcn13,CP47 from caseprogress, patent,trackingcasename,transcasetitle where cp09='" & pCP09 & "' and cp159=0 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp09=tcn05(+)"
      strQ1 = "Select cp158,pa05,pa06,PA10,CP44,tcn13,CP47,tct07,tct10 from caseprogress, patent,trackingcasename,transcasetitle where cp09='" & pCP09 & "' and cp159=0 and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp09=tcn05(+) and cp09=tct01(+)"
      intQ = 1
      Set rsRD = ClsLawReadRstMsg(intQ, strQ1)
      If intQ = 1 Then
         strCP27 = "" & rsRD.Fields("cp158")
         If pPA05 = "" Then
             pPA05 = "" & rsRD.Fields("pa05")
         End If
         If pPA06 = "" Then
             pPA06 = "" & rsRD.Fields("pa06")
         End If
         strQ1 = Pub_GetFCPGrpMan(nGRP)
         If strQ1 = "" Then strQ1 = "B"
         strQ1 = PUB_GetStateForMan(strQ1) '各組主管=>特殊情況之指定職代
         'Added by Lydia 2024/04/23
         m_Context = ""
         If Trim("" & rsRD.Fields("TCT07")) <> "" Then
            m_Context = m_Context & "原分案主任：" & rsRD.Fields("TCT07") & GetStaffName("" & rsRD.Fields("TCT07"), True) & vbCrLf
         End If
         If Trim("" & rsRD.Fields("TCT10")) <> "" Then
            m_Context = m_Context & "原命名人員：" & rsRD.Fields("TCT10") & GetStaffName("" & rsRD.Fields("TCT10"), True) & vbCrLf
         End If
         If m_Context <> "" Then m_Context = m_Context & vbCrLf
         'end 2024/04/23
         
         'Added by Lydia 2023/05/04 副本通知
         'P案完成認領，通知分案人員
         'Modified by Lydia 2023/05/29 在櫃台收文時直接給工程師組; ex.P-131591有指定相似案的組別TCN17+TCN18
         'If pCase(1) = "P" And "" & rsRd.Fields("pa10") = "" And pType = "2" Then
         'Modified by Lydia 2023/06/21 因為PCT案提申日也記錄在PA10，改用新案之代理人提申日來判斷未提申; ex.P-131748
         'If pCase(1) = "P" And "" & rsRd.Fields("pa10") = "" Then
         If pCase(1) = "P" And "" & rsRD.Fields("CP47") = "" Then
             If PUB_GetFMP2toP(pCase(1), pCase(2), pCase(3), pCase(4)) = True Then
                 strCC = strCC & IIf(strCC <> "", ";", "") & Pub_GetSpecMan("C")
             Else
                 strCC = strCC & IIf(strCC <> "", ";", "") & Pub_GetSpecMan("FMP非寰華分案人員")
             End If
         End If
         '由職代認領時,副本給職代
         If pType = "2" And InStr(strQ1, strUserNum) = 0 Then
             strB1 = "select tfa04,tfa09 from transfeeassign,staff where tfa01='" & pCP09 & "' and tfa04=st01(+) and st16='" & nGRP & "' and tfa05='Y' order by tfa09 desc "
             intB = 1
             Set rsBD = ClsLawReadRstMsg(intB, strB1)
             If intB = 1 Then
                If InStr(strQ1, "" & rsBD.Fields("tfa04")) = 0 Then
                   strCC = strCC & IIf(strCC <> "", ";", "") & rsBD.Fields("tfa04")
                End If
             End If
         End If
         'end 2023/05/04
                     
         strUpd = ""
         'Modified by Lydia 2023/06/14 判斷有改組別 And pGrp <> nGRP
         'Modified by Lydia 2023/07/27 + 〔非英說案〕由原組別重新進行命名作業
         'If pGrp <> "" And pGrp <> nGRP Then
         If (pGrp <> "" And pGrp <> nGRP) Or pType = "3" Then
           For intQ = 5 To TF_TCT
             If InStr(TF_TCTnotFS, Format(intQ, "000")) = 0 Then
                 Select Case intQ
                    Case 16  '案件名稱(中文)
                       If pPA05 <> "" Then
                           strUpd = strUpd & ", TCT16=" & CNULL(ChgSQL(pPA05))
                       End If
                    Case 17  '案件名稱(英文)
                       If pPA06 <> "" Then
                           strUpd = strUpd & ", TCT17=" & CNULL(ChgSQL(pPA06))
                       End If
                    Case 19 '是否收文主動修正
                       If PUB_ChkBCPisExist(pCase, "203") Then
                           strUpd = strUpd & ", TCT19='Y', TCT117=" & CNULL(IIf(Val(strCP27) > 0, "1", "2"))
                       Else
                           strUpd = strUpd & ", TCT19=NULL "
                       End If
                    Case 20 '是否收文告代
                       If PUB_ChkBCPisExist(pCase, "901") Then
                           strUpd = strUpd & ", TCT20=" & CNULL(IIf(Val(strCP27) > 0, "1", "2"))
                       Else
                           strUpd = strUpd & ", TCT20=NULL "
                       End If
                    Case 23, 24, 117
                       '---排除相似案號、相似度
                    Case Else
                       strUpd = strUpd & ", TCT" & IIf(intQ < 100, Format(intQ, "00"), Format(intQ, "000")) & "=NULL "
                 End Select
             End If
           Next intQ
         End If
         strUpd = strUpd & ", TCT112='" & strUserNum & "', TCT113=" & strSrvDate(1) & ", TCT114=" & Mid(Format(ServerTime, "000000"), 1, 4)
         strUpd = "Update TransCaseTitle Set TCT04=" & CNULL(strQ1) & strUpd & " Where TCT01='" & pCP09 & "' "
         cnnConnection.Execute strUpd, intQ
         If intQ > 0 Then
            '更改分案組別 , 通知雙方
            'Modified by Lydia 2023/06/14 若仍維持原命名的組別一樣不須重新命名
            'If pGrp <> "" Then
            If pType = "2" And pGrp = nGRP Then
               'Modified by Lydia 2023/05/26 Email主旨開頭改成模組
               strUpd = PUB_GetTCNmTitle(pCase(1), pCase(2), pCase(3), pCase(4), "" & rsRD.Fields("pa10"), "" & rsRD.Fields("tcn13"), "")
               strUpd = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                            " values( '" & strUserNum & "','" & strQ1 & "',to_char(sysdate,'yyyymmdd')" & _
                            ",to_char(sysdate,'hh24miss')," & CNULL(strUpd & "已完成認領，不需重新命名") & ",'同主旨','" & strCC & "')"
               cnnConnection.Execute strUpd
            'Modified by Lydia 2023/07/27 + 〔非英說案〕由原組別重新進行命名作業 + Or pType = "3"
            ElseIf pGrp <> "" Or pType = "3" Then
            'end 2023/06/14
                  '重新分案 , 清除卷宗區記錄, 直到新組別主管確認再次產生
                  strUpd = "delete from casepaperpdf where cpp01='" & pCP09 & "' and instr(cpp02,'" & FCP命名記錄 & "') > 0 "
                  cnnConnection.Execute strUpd, intQ
                   '命名作業的主管確認自動掛承辦人=命名人員並且上已分案,所以改工程師組別時一併清空,直到下次主管確認一併更新
                  'Modified by Lydia 2023/05/10 參考frm060102: 保留FCP案的告代901和主動修正203；因FCP新案急件重新認領，修改進度檔若有提申後告代、主動修正再發一次mail通知舊和新承辦人之事。
                  'strUpd = "Update caseprogress set cp14=null, cp122=null where " & ChgCaseprogress(pCase(1) & pCase(2) & pCase(3) & pCase(4)) & _
                           " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' and cp10 in ('902','203') "
                  strUpd = "Update CaseProgress set cp14=null, cp122=null where " & ChgCaseprogress(pCase(1) & pCase(2) & pCase(3) & pCase(4)) & _
                         " and cp158=0 and cp159=0 and substr(cp09,1,1)='A' and cp10 in (" & GetAddStr(IIf(pCase(1) = "FCP", Replace(TCTforCP14, "203,901,", ""), TCTforCP14)) & ") "
                  cnnConnection.Execute strUpd, intQ
                  
                  If pCase(1) <> "FCP" Then  'Added by Lydia 2023/05/10 保留FCP案
                     '清空-工程師收告代901和主動修正203
                     If pGrp <> nGRP Then
                         strUpd = ", cp64='" & ChangeWStringToWDateString(strSrvDate(1)) & "修改工程師組別：" & PUB_GetFCPGrpName(pGrp) & "->" & PUB_GetFCPGrpName(nGRP) & ";'||cp64 "
                     Else
                         strUpd = ""
                     End If
                     strUpd = "Update caseprogress set cp14=null, cp122=null " & strUpd & " where cp09 in (select cp09 from caseprogress,staff where " & ChgCaseprogress(pCase(1) & pCase(2) & pCase(3) & pCase(4)) & _
                                 " and cp158=0 and cp159=0 and substr(cp09,1,1)='B' and cp10 in ('203','901') and cp65=st01(+) and (st03='F21' or st03='M51') )  "
                     cnnConnection.Execute strUpd, intQ
                  End If 'Added by Lydia 2023/05/10
                  
                  strQ2 = pGrp & "-" & nGRP
                  strQ3 = Pub_GetFCPGrpMan(pGrp)
                  If strQ3 = "" Then strQ1 = "B"
                  strQ3 = PUB_GetStateForMan(strQ3) '特殊情況之指定職代
                  '改組別：分別通知兩個組
                  If pGrp <> nGRP Then
                     If PUB_GetTCTmail(True, 2, pCase(1), pCase(2), pCase(3), pCase(4), pCP09, "", strQ1, strQ2, , , strQ3) Then
                     End If
                     Sleep 1000 'Added by Lydi1 2023/05/10 避免mailcache出錯
                  End If
                  '最後：分案通知
                  If strQ1 <> "B" And strQ1 <> "" Then
                     'Modified by Lydia 2023/05/04 +指定副本人員
                     'Modified by Lydia 2024/04/23 +內文帶出原命名人員m_Context
                     'Modified by Lydia 2024/11/22 +
                     If PUB_GetTCTmail(True, 1, pCase(1), pCase(2), pCase(3), pCase(4), pCP09, "", strQ1, , , m_Context, strCC) Then
                     'end 2024/04/23
                     End If
                  End If
            Else
                  'Modified by Lydia 2023/05/04 +指定副本人員
                  'Modified by Lydia 2024/04/23 +內文帶出原命名人員m_Context
                  If PUB_GetTCTmail(True, 1, pCase(1), pCase(2), pCase(3), pCase(4), pCP09, "", strQ1, , , m_Context, strCC) Then
                  End If
            End If
         End If 'If intQ > 0 The
      End If
      Set rsRD = Nothing
      Set rsBD = Nothing 'Added by Lydia 2023/05/04
   End If
End Sub

'Added by Lydia 2023/03/02 外專新案認領：寄email通知國外部最高主管進行核判
Public Function PUB_GetTCNEmail(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pType As String, Optional ByVal pFilePath As String) As Boolean
Dim strQ1 As String, intQ As Integer
Dim strAttFList As String, strFileName As String
Dim strSpecSub As String
Dim rsQuery As New ADODB.Recordset

   'Email主旨：新案FCP0*****/P******請協助確認組別，謝謝！，並且將已存在原始檔區ORI.PDF當做附件。
   '2023/4/21 薛經理要改Email: 新案FCP0*****/P******認領有衝突，請惠予指派組別
   If pCP01 <> "" And pCP02 <> "" Then
       strQ1 = "SELECT CPF01,CPF02,CPF13 FROM CASEPROGRESS A,CASEPAPERFILE B " & _
                  "WHERE CP01='" & pCP01 & "' AND CP02='" & pCP02 & "' AND CP03='" & pCP03 & "' AND CP04='" & pCP04 & "' AND SUBSTR(CP09,1,1) = 'D' AND CP159=0 AND CP10='" & cntEnglish_Vers & "' " & _
                  "AND CP09=CPF01(+) AND NVL(CPF10,'N') <> 'D' AND UPPER(CPF02) LIKE '%.ORI.PDF' " & _
                  "ORDER BY CPF06 DESC, CPF07 DESC "
       intQ = 1
       Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
       If intQ = 1 Then
           If pFilePath = "" Then pFilePath = App.path
           If Right(pFilePath, 1) <> "\" Then pFilePath = pFilePath & "\"
           rsQuery.MoveFirst
           Do While Not rsQuery.EOF
              If "" & rsQuery.Fields("CPF01") <> "" And "" & rsQuery.Fields("CPF02") <> "" And "" & rsQuery.Fields("CPF13") <> "" Then
                  strFileName = pFilePath & "$$" & rsQuery.Fields("CPF02")   '下載檔案名稱+路徑
                  If PUB_GetFtpFile("" & rsQuery.Fields("CPF13"), strFileName, "CASEPAPERFILE") = True Then
                      strAttFList = strAttFList & strFileName & "*"
                  End If
              End If
              rsQuery.MoveNext
           Loop
       End If
       
       'Modified by Lydia 20223/05/26
       'strQ1 = "select pa10 from patent where pa01='" & pCP01 & "' and pa02='" & pCP02 & "' and pa03='" & pCP03 & "' and pa04='" & pCP04 & "' "
       strQ1 = "select pa10,tcn13 from patent, caseprogress, trackingcasename where pa01='" & pCP01 & "' and pa02='" & pCP02 & "' and pa03='" & pCP03 & "' and pa04='" & pCP04 & "' " & _
                   "and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and cp31='Y' and cp09=tcn05(+) "
       intQ = 1
       Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
       If intQ = 1 Then
          'Modified by Lydia 2023/05/26 Email主旨開頭改成模組
          strSpecSub = PUB_GetTCNmTitle(pCP01, pCP02, pCP03, pCP04, "" & rsQuery.Fields("pa10"), "" & rsQuery.Fields("tcn13"), "SPEC")
       End If
       Select Case pType
           Case "0": strSpecSub = Replace(strSpecSub, "SPEC", "(急件逾期)")
           Case "1": strSpecSub = Replace(strSpecSub, "SPEC", "(認領逾期)")
           Case Else: strSpecSub = Replace(strSpecSub, "SPEC", "(認領有衝突)")
       End Select
      strQ1 = Pub_GetSpecMan("外專新案命名核判主管")
      If strQ1 <> "" Then
         'Added by Lydia 2024/01/29
         If pType = "0" Or pType = "1" Then
            PUB_SendMail strUserNum, strQ1, "", strSpecSub & "，請惠予指派組別", "請由〔台一系統->工程師->外專新案認領區-主管核判〕處，進行指派。" & IIf(strAttFList <> "", vbCrLf & "***請參考附件***", "") & vbCrLf & vbCrLf & _
                "P.S.電臘中心人員可以調整認領期限並且通知未認領的主管，語法參考\\LINUX\PolyCOM\TaieNew\Script\外專新案認領.txt", , strAttFList, , , , , "QPGMR"
         Else
         'end 2024/01/29
            PUB_SendMail strUserNum, strQ1, "", strSpecSub & "，請惠予指派組別", "請由〔台一系統->工程師->外專新案認領區-主管核判〕處，進行指派。" & IIf(strAttFList <> "", vbCrLf & "***請參考附件***", ""), , strAttFList, , , , , "QPGMR"
         End If 'Added by Lydia 2024/01/29
         PUB_GetTCNEmail = True
      End If
      Set rsQuery = Nothing
   End If
End Function

'Added by Lydia 2023/03/02 外專新案認領：急件(新案)發文時，重新進入認領流程
'Modified by Lydia 2023/06/14 +重新認領bolRe
Public Function PUB_UpdateReTCN(ByRef tmpPA() As String, ByRef tmpCP() As String, Optional bolRe As Boolean = False) As Boolean
Dim intQ As Integer, strQuery As String, intB As Integer
Dim rsQuery As New ADODB.Recordset
Dim strA1 As String, strB1 As String, tmpArr As Variant
Dim strTCN23 As String
Dim strTCN13 As String 'Added by Lydia 2023/09/11
Dim strTCN24 As String 'Added by Lydia 2024/02/01
   PUB_UpdateReTCN = True
   '重新認領狀況1=新案發文(尚未輸入PA10), 狀況2=客戶提供文件bolRe
   
   '排除日文案PA150=3、103設計申請+125衍生設計、TCN17=TCN18=空白(相似案組別)、TCN23=0=空白
   'Modified by Lydia 2023/06/14 +重新認領bolRe
   'If (tmpPA(1) = "FCP" Or tmpPA(1) = "P") And tmpCP(31) = "Y" And InStr(FcpAddTct, tmpCP(10)) > 0 And Val(tmpPA(10)) = 0 Then
   If (tmpPA(1) = "FCP" Or tmpPA(1) = "P") And tmpCP(31) = "Y" And InStr(FcpAddTct, tmpCP(10)) > 0 And (Val(tmpPA(10)) = 0 Or bolRe = True) Then
      'Modified by Lydia 2024/02/20 +105: 有關" 香港013專利開放收文集體設計申請105"，請比照" 香港013專利收文設計申請103"
      If tmpPA(150) <> "3" And InStr("103,125,105", tmpCP(10)) = 0 Then
          'Modified by Lydia 2023/06/14 +tcn13外文本的對應英/中說
          'Modified by Lydia 2024/02/01 +TCN24 主管核判
          strQuery = "select tcn20,tcn17,tcn18,tcn23,nvl(tcn13,0) tcn13,tcn24 from TrackingCaseName Where TCN05='" & tmpCP(9) & "' "
          intQ = 1
          Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
          If intQ = 1 Then
             strTCN23 = "" & rsQuery.Fields("TCN23")
             strTCN13 = "" & rsQuery.Fields("TCN13") 'Added by Lydia 2023/09/11
             strTCN24 = "" & rsQuery.Fields("TCN24") 'Added by Lydia 2024/02/01
             If "" & rsQuery.Fields("tcn20") = tmpPA(150) And "" & rsQuery.Fields("TCN17") & rsQuery.Fields("TCN18") = "" And Val(strTCN23) = 0 Then
                 'Modified by Lydia 2023/06/14 (Email:112/5/15新案急件組別認領
                 'strQuery = "Update TransCaseTitle Set TCT04=null,TCT05=null,TCT06=null Where TCT01='" & tmpCP(9) & "' "
                 'cnnConnection.Execute strQuery
                 'If PUB_UpdateTCNstate("1", tmpPA(1) & tmpPA(2) & tmpPA(3) & tmpPA(4)) = False Then
                 '    PUB_UpdateReTCN = False
                 'End If
                 '******
                  '為了不延誤提申，仍維持由第一組認領時先去命名、提申；
                  '提申後判斷完成急件認領主管少於3人，則重新進入認領時仍保留第一次認領的狀態，即已放棄或已認領的資料仍會保留下來；等待未處理的主管確認後，如果沒有第二組認領則案件維持原命名的組別也不須重新命名；如果有第二組認領則進入協商階段。
                  '待最後結果若仍維持原命名的組別一樣不須重新命名；除非是不同組別認領到才需要重新走命名流程。
                  If InStr("1,2", "" & rsQuery.Fields("tcn13")) > 0 Then
                     '未確收之非英說案件：到客戶提供文件確收才判斷是否進入協調認領；
                  ElseIf InStr("3,4", "" & rsQuery.Fields("tcn13")) > 0 Then
                     '已確收之非英說案件：與非英說案件的協調認領階段判斷一致
                     GoTo JumpToRe
                  Else
                     strQuery = "select nvl(sum(decode(tfa05,null,0,1)),0) tot1,nvl(sum(decode(tfa05,'Y',1,0)),0) tot2 " & _
                                     "from transfeeassign where tfa01='" & tmpCP(9) & "' and tfa09='0' "
                     intQ = 1
                     Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
                     If intQ = 1 Then
                        '急件認領主管少於3人 或 無人認領(急件有預設組別CU154)
                        'Modified by Lydia 2024/02/27 改成常數
                        'If Val("" & rsQuery.Fields("tot1")) < 3 Or Val("" & rsQuery.Fields("tot2")) = 0 Then
                        If Val("" & rsQuery.Fields("tot1")) < FCPforEngNum Or Val("" & rsQuery.Fields("tot2")) = 0 Then
                           strA1 = "Update TransCaseTitle Set TCT04=null Where TCT01='" & tmpCP(9) & "' "
                           cnnConnection.Execute strA1
                           'Added by Lydia 2024/02/01 清空主管核判
                           If strTCN24 <> "" Then
                              strA1 = "Update TrackingCaseName Set TCN24=null Where TCN05='" & tmpCP(9) & "' "
                              cnnConnection.Execute strA1
                           End If
                           'end 2024/02/01
                           
                           If PUB_UpdateTCNstate("1", tmpPA(1) & tmpPA(2) & tmpPA(3) & tmpPA(4)) = True Then
                               If Val("" & rsQuery.Fields("tot1")) > 0 Then
                                  strA1 = "insert into transfeeassign (tfa01,tfa02,tfa03,tfa04,tfa05,tfa09) " & _
                                                  "select tfa01,to_char(sysdate,'yyyymmdd') tfa02,to_char(sysdate, 'HH24MISS') tfa03,tfa04,tfa05,'1' tfa09 " & _
                                                  "from transfeeassign where tfa01='" & tmpCP(9) & "' and tfa09='0' "
                                  cnnConnection.Execute strA1
                               End If
                           Else
                               PUB_UpdateReTCN = False
                           End If
                        End If
                     End If
                  End If
                 'end 2023/06/14
             'Added by Lydia 2023/06/14 客戶提供文件：確定已收文件／無文件
             ElseIf bolRe = True And InStr("3,4,", "" & rsQuery.Fields("tcn13")) > 0 Then
JumpToRe:
                  If "" & rsQuery.Fields("tcn20") = "" Then  '還在認領中
                     '客戶提供文件有發email
                  ElseIf "" & rsQuery.Fields("TCN17") & rsQuery.Fields("TCN18") <> "" Then
                     '排除有相似舊案指定組別
                  Else
                     '非英說案件：  設定為「有」or 「待確定」第一次認領階段需要三組主管都有輸入，若有兩組以上認領，則先給第一組進行命名，等待客戶提供文件後進入協調認領階段(TCN13=3,4)。
                     strQuery = "select tfa09,nvl(sum(decode(tfa05,null,0,1)),0) tot1,nvl(sum(decode(tfa05,'Y',1,0)),0) tot2 " & _
                                     "from transfeeassign where tfa01='" & tmpCP(9) & "' and tfa09 < 4 group by tfa09  order by tfa09 desc "
                     intQ = 1
                     Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
                     If intQ = 0 Then
                        strA1 = "0" '沒有經過認領: 急件有預設組別CU154
                        intQ = 0
                     Else
                        rsQuery.MoveFirst
                        strA1 = ""
                         '急件認領主管少於3人 或 無人認領(急件有預設組別CU154)
                        'Modified by Lydia 2024/02/27 改成常數
                        'If "" & rsQuery.Fields("tfa09") = "0" And (Val("" & rsQuery.Fields("tot1")) < 3 Or Val("" & rsQuery.Fields("tot2")) = 0) Then
                        If "" & rsQuery.Fields("tfa09") = "0" And (Val("" & rsQuery.Fields("tot1")) < FCPforEngNum Or Val("" & rsQuery.Fields("tot2")) = 0) Then
                             strA1 = rsQuery.Fields("tfa09")
                        '第一次認領階段若有兩組以上認領，等待客戶提供文件後進入第二認領階段
                        ElseIf Val("" & rsQuery.Fields("tfa09")) < 4 Then
                             If Val("" & rsQuery.Fields("tot2")) > 1 Then
                                strA1 = rsQuery.Fields("tfa09")
                             'Added by Lydia 2023/07/27 雖然不需要【進入認領階段】，但仍需要原組別檢視案件名稱，所以由原組別重新進行命名作業。
                             'Modified by Lydia 2023/09/11 確定無英說參考本(TCN13=4)，不用重新命名
                             ElseIf strTCN13 <> "4" Then
                                Call ChangeTCTGrp("3", tmpPA, tmpCP(9), tmpPA(150), tmpPA(150))
                             'end 2023/07/27
                             End If
                        End If
                        intQ = Val("" & rsQuery.Fields("tot2"))
                     End If
                     
                     If strA1 <> "" Then '進入第二認領階段
                        '判斷是否有認領過: 急件 或 無人認領(急件有預設組別CU154)=>全英文組
                        strB1 = PUB_GetEngGrpMan()
                        If strB1 <> "" Then
                           tmpArr = Split(strB1, ";")
                           For intQ = 0 To UBound(tmpArr)
                              If Trim(tmpArr(intQ)) <> "" Then
                                strQuery = "select a.* from transfeeassign a, staff b where a.tfa01='" & tmpCP(9) & "' and a.tfa09='" & strA1 & "' and a.tfa04=b.st01(+) and b.st16=" & CNULL(PUB_GetStaffST16("" & tmpArr(intQ)))
                                intB = 1
                                Set rsQuery = ClsLawReadRstMsg(intB, strQuery)
                                strB1 = ""
                                If intB = 1 Then '有認領=Y
                                   If "" & rsQuery.Fields("tfa05") = "Y" Then
                                      strB1 = "insert into transfeeassign (tfa01,tfa02,tfa03,tfa04,tfa05,tfa09) " & _
                                              "values ('" & tmpCP(9) & "', to_char(sysdate,'yyyymmdd'), to_char(sysdate, 'HH24MISS'), '" & rsQuery.Fields("tfa04") & "', '', '4') "
                                   End If
                                Else '沒有認領過
                                   strB1 = "insert into transfeeassign (tfa01,tfa02,tfa03,tfa04,tfa05,tfa09) " & _
                                           "values ('" & tmpCP(9) & "', to_char(sysdate,'yyyymmdd'), to_char(sysdate, 'HH24MISS'), '" & tmpArr(intQ) & "', '', '4') "
                                End If
                                If strB1 <> "" Then cnnConnection.Execute strB1
                              End If
                           Next intQ
                        End If
                        cnnConnection.Execute "Update TrackingCaseName Set TCN23='4' Where TCN05='" & tmpCP(9) & "' "
                        cnnConnection.Execute "Update TransCaseTitle Set TCT04=null Where TCT01='" & tmpCP(9) & "' "
                         
                        If PUB_UpdateTCNstate("1", tmpPA(1) & tmpPA(2) & tmpPA(3) & tmpPA(4)) = True Then
                        Else
                           cnnConnection.Execute "delete from transfeeassign where tfa01='" & tmpCP(9) & "' and tfa09='4' "
                           PUB_UpdateReTCN = False
                        End If
                     Else
                        '不用進入第二認領階段
                     End If
                  End If
             'end 2023/06/14
             End If
          End If
      End If 'If tmpPA(150) <> "3" And InStr("103,125", tmpCP(10)) = 0 Then
   End If
   
   Set rsQuery = Nothing
End Function

'Move by Lydia 2023/03/22 從basUpdate搬過來
'Memo by Lydia 2023/03/22 刪除PUB_GetApprMemo(單數備註)
'Added by Lydia 2022/07/29  核准函輸入備註檔(複數筆備註)
'pSelType: 查詢類型, pCaseNo:本所案號, pProperty:案件性質, pFCAgent:代理人編號, pApplicant:申請人編號1~5
Public Function PUB_GetApprMemo2(ByVal pSelType As String, ByVal pCaseNo As String, ByVal pProperty As String, ByVal pFCAgent As String, ByVal pApplicant As String) As String
   Dim stSQL As String, iR As Integer, mSQL As String
   Dim rsQuery As ADODB.Recordset
   Dim ixR As Integer, strMemo As String, midRep As String
   Dim BolExit As Boolean '是否繼續讀取Recordset
   Dim iPA26 As Variant, strCU As String
   Dim strLevel As String '記錄符合的階段
   Dim bolThisJump As Boolean  '是否跳過這次讀取
    
   '４.初審=核准管制分割期限
   If pSelType = "4" Then
       mSQL = " and AM07 ='4' "
   '一般核准用
   ElseIf pSelType = "1" Then
      mSQL = " and AM07 in ('1','3') "
   '核對已准用
   ElseIf pSelType = "2" Then
      mSQL = " and AM07 in ('2','3') "
   End If
   pProperty = Replace(pProperty, ",", "','")
   iPA26 = Split(pApplicant, ",")
   
   For ixR = 0 To UBound(iPA26)
      '順序 １.本所案號 ２.代理人+申請人 ３.代理人 ４.申請人
      If Trim(iPA26(ixR)) <> "" Then
         strCU = ChangeCustomerL(iPA26(ixR))
         'Modified by Lydia 2023/03/22 從先處理Y編號，改成先處理Y+X編號 ---Sharon
         'stSQL = "select AM02,0 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM03='" & pCaseNo & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,2 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04='" & Left(pFCAgent, 8) & "' and AM05='" & Left(strCU, 8) & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,4 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04='" & Left(pFCAgent, 8) & "' and AM05='" & Left(strCU, 6) & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,6 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04='" & Left(pFCAgent, 8) & "' and AM05 is null and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,8 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04='" & Left(pFCAgent, 6) & "' and AM05='" & Left(strCU, 8) & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,10 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04='" & Left(pFCAgent, 6) & "' and AM05='" & Left(strCU, 6) & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,12 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04='" & Left(pFCAgent, 6) & "' and AM05 is null and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,14 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04 is null and AM05='" & Left(strCU, 8) & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,16 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04 is null and AM05='" & Left(strCU, 6) & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " order by od1,od2"
         stSQL = "select AM02,0 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM03='" & pCaseNo & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,1 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04='" & Left(pFCAgent, 8) & "' and AM05='" & Left(strCU, 8) & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,2 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04='" & Left(pFCAgent, 8) & "' and AM05='" & Left(strCU, 6) & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,3 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04='" & Left(pFCAgent, 6) & "' and AM05='" & Left(strCU, 8) & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,4 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04='" & Left(pFCAgent, 6) & "' and AM05='" & Left(strCU, 6) & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,5 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04='" & Left(pFCAgent, 8) & "' and AM05 is null and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,6 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04='" & Left(pFCAgent, 6) & "' and AM05 is null and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,7 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04 is null and AM05='" & Left(strCU, 8) & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " union select AM02,8 Od1, AM07 Od2, AM06, am03 as K01, am04 as K02, am05 as K03 from ApprovalMemo2 where AM04 is null and AM05='" & Left(strCU, 6) & "' and AM06 in ('" & pProperty & "')" & mSQL & _
            " order by od1,od2"
         BolExit = False
         iR = 1
         Set rsQuery = ClsLawReadRstMsg(iR, stSQL)
         If iR = 1 Then
            rsQuery.MoveFirst
            If Len(strMemo) = 0 Then '第一次符合條件
               strMemo = strMemo & rsQuery(0) '傳回-備註
               If pSelType = "4" Then '４.初審=核准管制分割期限，只抓第一筆符合
                   BolExit = True
                   Exit For
               Else
                   '不同申請人-> 記上一次條件的title
                   If Not IsNull(rsQuery.Fields("k01")) Then
                      midRep = "(" & rsQuery.Fields("k01") & ")" & vbCrLf
                      strLevel = "1.個案:" & "(" & rsQuery.Fields("k01") & ")"
                   ElseIf Not IsNull(rsQuery.Fields("k02")) And Not IsNull(rsQuery.Fields("k03")) Then
                         midRep = "(" & rsQuery.Fields("k02") & " + " & rsQuery.Fields("k03") & ")" & vbCrLf
                         strLevel = "2.Y+X:" & "(" & rsQuery.Fields("k02") & " + " & rsQuery.Fields("k03") & ")"
                   Else
                        '３－只有代理人符合
                        If Not IsNull(rsQuery.Fields("k02")) And IsNull(rsQuery.Fields("k03")) Then
                            If ixR = 0 Then
                               midRep = "(" & rsQuery.Fields("k02") & ")" & vbCrLf
                               strLevel = "3.Y:" & "(" & rsQuery.Fields("k02") & ")"
                            Else
                               midRep = ""
                            End If
                        End If
                        '４－只有申請人符合
                        If IsNull(rsQuery.Fields("k02")) And Not IsNull(rsQuery.Fields("k03")) Then
                            midRep = "(" & rsQuery.Fields("k03") & ")" & vbCrLf
                            strLevel = "3.X:" & "(" & rsQuery.Fields("k03") & ")"
                        End If
                   End If
               End If '不同申請人-> 記上一次條件的title
            End If '第一次符合條件
            '第一次迴圈的個案在內部判斷
            If ixR = 0 Or (ixR > 0 And rsQuery.RecordCount > 0) Or Len(strMemo) = 0 Then
               Do While Not rsQuery.EOF
                  '新規則順位修改成三階段，依序符合階段就不再抓資料
                  '(1)有個案：個案+ (Y+X1~X5) => 個案+Y or X  (2)Y+X  (3)Y or X
                  '=>整合來看只要有Y+X之後就不單獨讀取Y or X
                  bolThisJump = False
                  '第一次迴圈有判斷個案，之後要跳過個案
                  If InStr(strLevel, "1.個案") > 0 And ixR > 0 And Not IsNull(rsQuery.Fields("k01")) Then
                    bolThisJump = True
                  End If
                  '個案+ (Y+X1~X5)
                  If Not IsNull(rsQuery.Fields("k02")) And Not IsNull(rsQuery.Fields("k03")) Then
                     If InStr(strLevel, ";Y+X:(" & Left(rsQuery.Fields("k02"), 8) & " + " & Left(rsQuery.Fields("k03"), 8)) > 0 Or (InStr(strLevel, ";Y+X:(" & Left(rsQuery.Fields("k02"), 6)) > 0 And InStr(strLevel, " + " & Left(rsQuery.Fields("k03"), 6)) > 0) Then
                        bolThisJump = True
                     End If
                     strLevel = strLevel & ";Y+X:(" & rsQuery.Fields("k02") & " + " & rsQuery.Fields("k03") & ")" & ",Jump=" & IIf(bolThisJump = True, "T", "F")
                  '個案+ Y or X
                  ElseIf Not IsNull(rsQuery.Fields("k02")) Or Not IsNull(rsQuery.Fields("k03")) Then
                     If InStr(strLevel & ",", "Y+X") > 0 Then
                        bolThisJump = True
                        BolExit = True 'Y+X1之後就不單獨讀取Y or X
                     End If
                     '曾經抓到Y or X編號，就不用讀取；1.先抓到8碼XY,後抓到6碼，2.因為設定C類來函性質可以同一Y or X有一筆以上的設定
                     If Not IsNull(rsQuery.Fields("k02")) And InStr(strLevel, ";Y:(" & rsQuery.Fields("k02")) > 0 Then      'Y編號
                         bolThisJump = True
                     End If
                     If Not IsNull(rsQuery.Fields("k03")) And InStr(strLevel, ";X:(" & rsQuery.Fields("k03")) > 0 Then   'X編號
                         bolThisJump = True
                     End If
                     strLevel = strLevel & IIf(Not IsNull(rsQuery.Fields("k02")), ";Y:(" & rsQuery.Fields("k02") & ")", _
                                    ";X:(" & rsQuery.Fields("k03") & ")") & ",Jump=" & IIf(bolThisJump = True, "T", "F")
                  End If
                  'Debug.Print strLevel
                  
                  If bolThisJump = False Then
                    If Left(strMemo, 10) <> Left(rsQuery(0), 10) Then
                      If Not IsNull(rsQuery.Fields("k02")) And IsNull(rsQuery.Fields("k03")) Then
                         If ixR = 0 Then strMemo = strMemo & "PS-SPACES" & rsQuery(0)
                      Else
                         strMemo = strMemo & "PS-SPACES" & rsQuery(0)
                      End If
                      '不同申請人-> 記上一次條件的title
                      If Not (Left(strMemo, 2) = "(X" Or Left(strMemo, 2) = "(Y" Or Left(strMemo, 2) = "(F" Or Left(strMemo, 2) = "(P") Then
                         strMemo = midRep & strMemo
                      End If
                    Else
                       ' strMemo = "PSSPACES-" & strMemo
                    End If
                    If Not IsNull(rsQuery.Fields("k01")) Then
                      strMemo = Replace(strMemo, "PSSPACES-", "(" & rsQuery.Fields("k01") & ")" & vbCrLf)
                      strMemo = Replace(strMemo, "PS-SPACES", "" & vbCrLf & "(" & rsQuery.Fields("k01") & ")" & vbCrLf)
                      midRep = "(" & rsQuery.Fields("k01") & ")" & vbCrLf
                    ElseIf Not IsNull(rsQuery.Fields("k02")) And Not IsNull(rsQuery.Fields("k03")) Then
                       strMemo = Replace(strMemo, "PSSPACES-", "(" & rsQuery.Fields("k02") & " + " & rsQuery.Fields("k03") & ")" & vbCrLf)
                       strMemo = Replace(strMemo, "PS-SPACES", "" & vbCrLf & "(" & rsQuery.Fields("k02") & " + " & rsQuery.Fields("k03") & ")" & vbCrLf)
                       midRep = "(" & rsQuery.Fields("k02") & " + " & rsQuery.Fields("k03") & ")" & vbCrLf
                    Else
                       '３－只有代理人符合 (代理人只取第一次的值)
                        If Not IsNull(rsQuery.Fields("k02")) And IsNull(rsQuery.Fields("k03")) And ixR = 0 Then
                            strMemo = Replace(strMemo, "PSSPACES-", "(" & rsQuery.Fields("k02") & ")" & vbCrLf)
                            strMemo = Replace(strMemo, "PS-SPACES", "" & vbCrLf & "(" & rsQuery.Fields("k02") & ")" & vbCrLf)
                            midRep = "(" & rsQuery.Fields("k02") & ")" & vbCrLf
                        End If
                        '４－只有申請人符合
                        If IsNull(rsQuery.Fields("k02")) And Not IsNull(rsQuery.Fields("k03")) Then
                            strMemo = Replace(strMemo, "PSSPACES-", "(" & rsQuery.Fields("k03") & ")" & vbCrLf)
                            strMemo = Replace(strMemo, "PS-SPACES", "" & vbCrLf & "(" & rsQuery.Fields("k03") & ")" & vbCrLf)
                            midRep = "(" & rsQuery.Fields("k03") & ")" & vbCrLf
                        End If
                    End If
        
                    If BolExit = True Then Exit Do
                  End If 'bolThisJump
                  rsQuery.MoveNext
               Loop
            End If  'If ixR = 0 Or (ixR > 0 And rsQuery.RecordCount > 0) Or Len(strMemo) = 0 Then
         End If
      End If   'If Trim(iPA26(ixR)) <> "" Then
   Next ixR
   
   PUB_GetApprMemo2 = strMemo
   Set rsQuery = Nothing
End Function

'Move by Lydia 2023/03/22 從basUpdate搬過來
'Memo by Lydia 2023/03/22 刪除PUB_GetNpMemo(單數備註)
'Added by Lydia 2022/08/01 下一程序固定備註(複數筆備註)
'pSelType: 查詢類型(1-查詢複數, 2-更新單筆), pCaseNo:本所案號,pProperty:案件性質,pFCAgent:代理人編號,pApplicant:申請人編號1~5
Public Function PUB_GetNpMemo2(ByVal pSelType As String, ByVal pCaseNo As String, ByVal pProperty As String, ByVal pFCAgent As String, ByVal pApplicant As String) As String
   Dim stSQL As String, iR As Integer
   Dim rsQuery As ADODB.Recordset
   Dim ixR As Integer, strMemo As String, midRep As String
   Dim BolExit As Boolean '是否繼續讀取Recordset
   Dim iPA26 As Variant, strCU As String
   Dim strLevel As String '記錄符合的階段
   Dim bolThisJump As Boolean  '是否跳過這次讀取
 
   iPA26 = Split(pApplicant, ",")

   For ixR = 0 To UBound(iPA26)
      '順序 １.本所案號 ２.代理人+申請人 ３.代理人 ４.申請人
      If Trim(iPA26(ixR)) <> "" Then
         strCU = ChangeCustomerL(iPA26(ixR))
         'Modified by Lydia 2023/03/22 從先處理Y編號，改成先處理Y+X編號 ---Sharon
         'stSQL = "select NM02,0 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from npmemo where nm03='" & pCaseNo & "' and nm06='" & pProperty & "'" & _
            " union select NM02,2 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 8) & "' and nm05='" & Left(iPA26(ixR), 8) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,4 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 8) & "' and nm05='" & Left(iPA26(ixR), 6) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,6 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 8) & "' and nm05 is null and nm06='" & pProperty & "'" & _
            " union select NM02,8 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 6) & "' and nm05='" & Left(iPA26(ixR), 8) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,10 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 6) & "' and nm05='" & Left(iPA26(ixR), 6) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,12 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 6) & "' and nm05 is null and nm06='" & pProperty & "'" & _
            " union select NM02,14 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04 is null and nm05='" & Left(iPA26(ixR), 8) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,16 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04 is null and nm05='" & Left(iPA26(ixR), 6) & "' and nm06='" & pProperty & "'" & _
            " order by od1"
         'Modified by Morgan 2023/7/19
         'stSQL = "select NM02,0 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from npmemo where nm03='" & pCaseNo & "' and nm06='" & pProperty & "'" & _
            " union select NM02,1 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 8) & "' and nm05='" & Left(iPA26(ixR), 8) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,2 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 8) & "' and nm05='" & Left(iPA26(ixR), 6) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,3 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 6) & "' and nm05='" & Left(iPA26(ixR), 8) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,4 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 6) & "' and nm05='" & Left(iPA26(ixR), 6) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,5 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 8) & "' and nm05 is null and nm06='" & pProperty & "'" & _
            " union select NM02,6 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 6) & "' and nm05 is null and nm06='" & pProperty & "'" & _
            " union select NM02,7 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04 is null and nm05='" & Left(iPA26(ixR), 8) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,8 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04 is null and nm05='" & Left(iPA26(ixR), 6) & "' and nm06='" & pProperty & "'" & _
            " order by od1"
         stSQL = "select NM02,0 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from npmemo where nm03='" & pCaseNo & "' and nm06='" & pProperty & "'" & _
            " union select NM02,1 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 8) & "' and nm05='" & Left(strCU, 8) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,2 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 8) & "' and nm05='" & Left(strCU, 6) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,3 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 6) & "' and nm05='" & Left(strCU, 8) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,4 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 6) & "' and nm05='" & Left(strCU, 6) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,5 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 8) & "' and nm05 is null and nm06='" & pProperty & "'" & _
            " union select NM02,6 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04='" & Left(pFCAgent, 6) & "' and nm05 is null and nm06='" & pProperty & "'" & _
            " union select NM02,7 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04 is null and nm05='" & Left(strCU, 8) & "' and nm06='" & pProperty & "'" & _
            " union select NM02,8 Od1, NM06, nm03 as K01, nm04 as K02, nm05 as K03 from  npmemo where nm04 is null and nm05='" & Left(strCU, 6) & "' and nm06='" & pProperty & "'" & _
            " order by od1"
         BolExit = False
         iR = 1
         Set rsQuery = ClsLawReadRstMsg(iR, stSQL)
         If iR = 1 Then
            rsQuery.MoveFirst
            If Len(strMemo) = 0 Then '第一次符合條件
               strMemo = strMemo & rsQuery(0) '傳回-備註
               If pSelType = "2" Then '更新單筆，只抓第一筆符合
                  BolExit = True
                  Exit For
               Else
                   '不同申請人-> 記上一次條件的title
                   If Not IsNull(rsQuery.Fields("k01")) Then
                      midRep = "(" & rsQuery.Fields("k01") & ")"
                      strLevel = "1.個案:" & "(" & rsQuery.Fields("k01") & ")"
                   ElseIf Not IsNull(rsQuery.Fields("k02")) And Not IsNull(rsQuery.Fields("k03")) Then
                         midRep = "(" & rsQuery.Fields("k02") & " + " & rsQuery.Fields("k03") & ")"
                         strLevel = "2.Y+X:" & "(" & rsQuery.Fields("k02") & " + " & rsQuery.Fields("k03") & ")"
                   Else
                        '３－只有代理人符合
                        If Not IsNull(rsQuery.Fields("k02")) And IsNull(rsQuery.Fields("k03")) Then
                            If ixR = 0 Then
                               midRep = "(" & rsQuery.Fields("k02") & ")"
                               strLevel = "3.Y:" & "(" & rsQuery.Fields("k02") & ")"
                            Else
                               midRep = ""
                            End If
                        End If
                        '４－只有申請人符合
                        If IsNull(rsQuery.Fields("k02")) And Not IsNull(rsQuery.Fields("k03")) Then
                            midRep = "(" & rsQuery.Fields("k03") & ")"
                            strLevel = "3.X:" & "(" & rsQuery.Fields("k03") & ")"
                        End If
                   End If
               End If
            End If '第一次符合條件
            '第一次迴圈的個案在內部判斷
            If ixR = 0 Or (ixR > 0 And rsQuery.RecordCount > 0) Or Len(strMemo) = 0 Then
               Do While Not rsQuery.EOF
                  '新規則順位修改成三階段，依序符合階段就不再抓資料
                  '(1)有個案：個案+ (Y+X1~X5) => 個案+Y or X  (2)Y+X  (3)Y or X
                  '=>整合來看只要有Y+X之後就不單獨讀取Y or X
                  bolThisJump = False
                  '第一次迴圈有判斷個案，之後要跳過個案
                  If InStr(strLevel, "1.個案") > 0 And ixR > 0 And Not IsNull(rsQuery.Fields("k01")) Then
                    bolThisJump = True
                  End If
                  '個案+ (Y+X1~X5)
                  If Not IsNull(rsQuery.Fields("k02")) And Not IsNull(rsQuery.Fields("k03")) Then
                     If InStr(strLevel, ";Y+X:(" & Left(rsQuery.Fields("k02"), 8) & " + " & Left(rsQuery.Fields("k03"), 8)) > 0 Or (InStr(strLevel, ";Y+X:(" & Left(rsQuery.Fields("k02"), 6)) > 0 And InStr(strLevel, " + " & Left(rsQuery.Fields("k03"), 6)) > 0) Then
                        bolThisJump = True
                     End If
                     strLevel = strLevel & ";Y+X:(" & rsQuery.Fields("k02") & " + " & rsQuery.Fields("k03") & ")" & ",Jump=" & IIf(bolThisJump = True, "T", "F")
                  '個案+ Y or X
                  ElseIf Not IsNull(rsQuery.Fields("k02")) Or Not IsNull(rsQuery.Fields("k03")) Then
                     If InStr(strLevel & ",", "Y+X") > 0 Then
                        bolThisJump = True
                        BolExit = True 'Y+X1之後就不單獨讀取Y or X
                     End If
                     '曾經抓到Y or X編號，就不用讀取；1.先抓到8碼XY,後抓到6碼，2.因為設定C類來函性質可以同一Y or X有一筆以上的設定
                     If Not IsNull(rsQuery.Fields("k02")) And InStr(strLevel, ";Y:(" & rsQuery.Fields("k02")) > 0 Then      'Y編號
                         bolThisJump = True
                     End If
                     If Not IsNull(rsQuery.Fields("k03")) And InStr(strLevel, ";X:(" & rsQuery.Fields("k03")) > 0 Then   'X編號
                         bolThisJump = True
                     End If
                     strLevel = strLevel & IIf(Not IsNull(rsQuery.Fields("k02")), ";Y:(" & rsQuery.Fields("k02") & ")", _
                                    ";X:(" & rsQuery.Fields("k03") & ")") & ",Jump=" & IIf(bolThisJump = True, "T", "F")
                  End If
                  'Debug.Print strLevel

                  If bolThisJump = False Then
                    If Left(strMemo, 10) <> Left(rsQuery(0), 10) Then
                      If strMemo <> "" Then strMemo = strMemo & ";"
                      If Not IsNull(rsQuery.Fields("k02")) And IsNull(rsQuery.Fields("k03")) Then
                         If ixR = 0 Then strMemo = strMemo & "PS-SPACES" & rsQuery(0)
                      Else
                         strMemo = strMemo & "PS-SPACES" & rsQuery(0)
                      End If
                      '不同申請人-> 記上一次條件的title
                      If Not (Left(strMemo, 2) = "(X" Or Left(strMemo, 2) = "(Y" Or Left(strMemo, 2) = "(F" Or Left(strMemo, 2) = "(P") Then
                         strMemo = midRep & strMemo
                      End If
                    Else
                       ' strMemo = "PSSPACES-" & strMemo
                    End If
                    If Not IsNull(rsQuery.Fields("k01")) Then
                      strMemo = Replace(strMemo, "PSSPACES-", "(" & rsQuery.Fields("k01") & ")")
                      strMemo = Replace(strMemo, "PS-SPACES", "(" & rsQuery.Fields("k01") & ")")
                      midRep = "(" & rsQuery.Fields("k01") & ")"
                    ElseIf Not IsNull(rsQuery.Fields("k02")) And Not IsNull(rsQuery.Fields("k03")) Then
                       strMemo = Replace(strMemo, "PSSPACES-", "(" & rsQuery.Fields("k02") & " + " & rsQuery.Fields("k03") & ")")
                       strMemo = Replace(strMemo, "PS-SPACES", "(" & rsQuery.Fields("k02") & " + " & rsQuery.Fields("k03") & ")")
                       midRep = "(" & rsQuery.Fields("k02") & " + " & rsQuery.Fields("k03") & ")"
                    Else
                       '３－只有代理人符合 (代理人只取第一次的值)
                        If Not IsNull(rsQuery.Fields("k02")) And IsNull(rsQuery.Fields("k03")) And ixR = 0 Then
                            strMemo = Replace(strMemo, "PSSPACES-", "(" & rsQuery.Fields("k02") & ")")
                            strMemo = Replace(strMemo, "PS-SPACES", "(" & rsQuery.Fields("k02") & ")")
                            midRep = "(" & rsQuery.Fields("k02") & ")"
                        End If
                        '４－只有申請人符合
                        If IsNull(rsQuery.Fields("k02")) And Not IsNull(rsQuery.Fields("k03")) Then
                            strMemo = Replace(strMemo, "PSSPACES-", "(" & rsQuery.Fields("k03") & ")")
                            strMemo = Replace(strMemo, "PS-SPACES", "(" & rsQuery.Fields("k03") & ")")
                            midRep = "(" & rsQuery.Fields("k03") & ")"
                        End If
                    End If

                    If BolExit = True Then Exit Do
                  End If 'bolThisJump
                  rsQuery.MoveNext
               Loop
            End If  'If ixR = 0 Or (ixR > 0 And rsQuery.RecordCount > 0) Or Len(strMemo) = 0 Then
         End If
      End If   'If Trim(iPA26(ixR)) <> "" Then
   Next ixR
   
   PUB_GetNpMemo2 = strMemo
   Set rsQuery = Nothing
End Function

'Move by Lydia 2023/03/22 從basUpdate搬過來
'Added by Lydia 2015/02/06 請款函預設備註
'dbCaseNo:本所案號,dbPty:案件性質,dbFA:代理人編號,dbCu:申請人編號
Public Function PUB_GetDebitNotePS(dbCaseNo As String, dbPty As String, dbFA As String, dbCu As String) As String
   Dim stSQL As String, iR As Integer
   Dim rsQuery As ADODB.Recordset
   'Added by Lydia 2023/08/01
   Dim iCall As Integer, iRound As Integer, tmpArr As Variant
   '判斷有幾個申請人
   tmpArr = Split(dbCu, ",")
   For iR = 0 To UBound(tmpArr)
       If Trim(tmpArr(iR)) <> "" Then
          tmpArr(iR) = ChangeCustomerL(tmpArr(iR))
          iCall = iCall + 1
       End If
   Next iR
   For iRound = 1 To iCall
   'end 2023/08/01
      '順序 1.本所案號 2.代理人+申請人 3.代理人 4.申請人
      '例外判斷 --"FCP039517000", "FCP039142000", "FCP045373000", "FCP035771000", "FCP051024000",605 年費 除外
      'Modified by Lydia 2023/08/01
      'If InStr("X5403200,X5988900", Left(IIf(Len(dbCu) = 6, dbCu & "00", dbCu), 8)) > 0 Then
      If InStr("X5403200,X5988900", Left(tmpArr(iRound - 1), 8)) > 0 Then
         If InStr("FCP039517000,FCP039142000,FCP045373000,FCP035771000,FCP051024000", dbCaseNo) > 0 Or dbPty = "605" Then
            Exit Function
         End If
      End If
      'Modified by Lydia 2023/03/22 從先處理Y編號，改成先處理Y+X編號 ---Sharon
      'stSQL = "select DNPS02,0 Od1 from DebitNotePS where DNPS03='" & dbCaseNo & "' " & _
         " union select DNPS02,1 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 8) & "' and DNPS05='" & Left(dbCu, 8) & "' " & _
         " union select DNPS02,2 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 8) & "' and DNPS05='" & Left(dbCu, 6) & "' " & _
         " union select DNPS02,3 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 8) & "' and DNPS05 is null" & _
         " union select DNPS02,4 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 6) & "' and DNPS05='" & Left(dbCu, 8) & "' " & _
         " union select DNPS02,5 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 6) & "' and DNPS05='" & Left(dbCu, 6) & "' " & _
         " union select DNPS02,6 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 6) & "' and DNPS05 is null" & _
         " union select DNPS02,7 Od1 from DebitNotePS where DNPS04 is null and DNPS05='" & Left(dbCu, 8) & "' " & _
         " union select DNPS02,8 Od1 from DebitNotePS where DNPS04 is null and DNPS05='" & Left(dbCu, 6) & "' " & _
         " order by od1"
      'Modified by Lydia 2023/08/01 +增加案件性質,順便拿掉XY編號6碼的判斷
      'stSQL = "select DNPS02,0 Od1 from DebitNotePS where DNPS03='" & dbCaseNo & "' " & _
         " union select DNPS02,1 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 8) & "' and DNPS05='" & Left(dbCu, 8) & "' " & _
         " union select DNPS02,2 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 8) & "' and DNPS05='" & Left(dbCu, 6) & "' " & _
         " union select DNPS02,3 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 6) & "' and DNPS05='" & Left(dbCu, 8) & "' " & _
         " union select DNPS02,4 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 6) & "' and DNPS05='" & Left(dbCu, 6) & "' " & _
         " union select DNPS02,5 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 8) & "' and DNPS05 is null" & _
         " union select DNPS02,6 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 6) & "' and DNPS05 is null" & _
         " union select DNPS02,7 Od1 from DebitNotePS where DNPS04 is null and DNPS05='" & Left(dbCu, 8) & "' " & _
         " union select DNPS02,8 Od1 from DebitNotePS where DNPS04 is null and DNPS05='" & Left(dbCu, 6) & "' " & _
         " order by od1"
      stSQL = "select DNPS02,0 Od1 from DebitNotePS where DNPS03='" & dbCaseNo & "' and DNPS12='" & dbPty & "'" & _
         " union select DNPS02,1 Od1 from DebitNotePS where DNPS03='" & dbCaseNo & "' and DNPS12 is null" & _
         " union select DNPS02,2 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 8) & "' and DNPS05='" & Left(tmpArr(iRound - 1), 8) & "' and DNPS12='" & dbPty & "'" & _
         " union select DNPS02,3 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 8) & "' and DNPS05='" & Left(tmpArr(iRound - 1), 8) & "' and DNPS12 is null" & _
         " union select DNPS02,4 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 8) & "' and DNPS05 is null and DNPS12='" & dbPty & "'" & _
         " union select DNPS02,5 Od1 from DebitNotePS where DNPS04='" & Left(dbFA, 8) & "' and DNPS05 is null and DNPS12 is null" & _
         " union select DNPS02,6 Od1 from DebitNotePS where DNPS04 is null and DNPS05='" & Left(tmpArr(iRound - 1), 8) & "' and DNPS12='" & dbPty & "'" & _
         " union select DNPS02,7 Od1 from DebitNotePS where DNPS04 is null and DNPS05='" & Left(tmpArr(iRound - 1), 8) & "' and DNPS12 is null" & _
         " order by od1"
      iR = 1
      Set rsQuery = ClsLawReadRstMsg(iR, stSQL)
      If iR = 1 Then
         PUB_GetDebitNotePS = "" & rsQuery(0)
         GoTo EXITSUB 'Added by Lydia 2023/08/01
      End If
   Next iRound  'Added by Lydia 2023/08/01 For iRound = 1 To iCall
   
EXITSUB: 'Added by Lydia 2023/08/01
   Set rsQuery = Nothing
   
End Function

'Move by Lydia 2023/03/22 從basQuery搬過來; 並且整合frm06010602_3.GetApprovalPS, +pKind=2通知工程師Email
'Modified by Morgan 2022/10/18 從frm060316_1移來
'Modified by Lydia 2022/10/05 增加"定稿語文"pStLang
'Modified by Lydia 2020/12/30 pJnClaims 日文定稿請求項
'Added by Lydia 2019/03/11 通知告准加註(ApprvoalPS)
Public Function PUB_GetApprovalPS(ByVal pKind As String, ByVal dbCaseNo As String, ByVal dbFA As String, ByVal dbCu As String, Optional ByRef pMemo As String = "", Optional ByRef pJnClaims As String = "", Optional ByVal pStLang As String = "2") As Boolean
'dbCaseNo:本所案號,dbFA:代理人編號,dbCu:申請人編號
'pKind=2通知工程師Email
Dim stSQL As String, iR As Integer
Dim stCon As String
Dim rsQuery As ADODB.Recordset
'逐筆判斷Y代理人+X申請人1~5;若有一筆以上,只使用第一筆符合
Dim m_Memo As String
Dim iCall As Integer, iRound As Integer
Dim tmpArr As Variant
Dim m_Claims As String 'Added by Lydia 2020/12/30
  
   '判斷有幾個申請人
   tmpArr = Split(dbCu, ",")
   For iR = 0 To UBound(tmpArr)
       If Trim(tmpArr(iR)) <> "" Then
           iCall = iCall + 1
       End If
   Next iR
   
   For iRound = 1 To iCall
        '順序 1.本所案號 2.代理人+申請人 3.代理人 4.申請人
        'Modified by Lydia 2020/12/30 + APS12 日文定稿請求項
        'Modified by Lydia 2022/10/05 +APS15 日文定稿加註
        'Moeified by Lydia 2023/03/22 +APS13,APS14通知工程師Email主旨,內文
        'Modified by Lydia 2023/03/22 從先處理Y編號，改成先處理Y+X編號 ---Sharon
        'stSQL = "select 0 Od1, APS02, APS12, APS15 from ApprovalPS where APS03='" & dbCaseNo & "' " & stCon & _
           " union select 1 Od1, APS02, APS12, APS15 from ApprovalPS where APS04='" & Left(dbFA, 8) & "' and APS05='" & Left(tmpArr(iRound - 1), 8) & "' " & stCon & _
           " union select 2 Od1, APS02, APS12, APS15 from ApprovalPS where APS04='" & Left(dbFA, 8) & "' and APS05='" & Left(tmpArr(iRound - 1), 6) & "' " & stCon & _
           " union select 3 Od1, APS02, APS12, APS15 from ApprovalPS where APS04='" & Left(dbFA, 8) & "' and APS05 is null" & stCon & _
           " union select 4 Od1, APS02, APS12, APS15 from ApprovalPS where APS04='" & Left(dbFA, 6) & "' and APS05='" & Left(tmpArr(iRound - 1), 8) & "' " & stCon & _
           " union select 5 Od1, APS02, APS12, APS15 from ApprovalPS where APS04='" & Left(dbFA, 6) & "' and APS05='" & Left(tmpArr(iRound - 1), 6) & "' " & stCon & _
           " union select 6 Od1, APS02, APS12, APS15 from ApprovalPS where APS04='" & Left(dbFA, 6) & "' and APS05 is null" & stCon & _
           " union select 7 Od1, APS02, APS12, APS15 from ApprovalPS where APS04 is null and APS05='" & Left(tmpArr(iRound - 1), 8) & "' " & stCon & _
           " union select 8 Od1, APS02, APS12, APS15 from ApprovalPS where APS04 is null and APS05='" & Left(tmpArr(iRound - 1), 6) & "' " & stCon & _
           " order by Od1, APS02"
        stSQL = "select 0 Od1, APS02, APS12, APS15, APS13, APS14  from ApprovalPS where APS03='" & dbCaseNo & "' " & stCon & _
           " union select 1 Od1, APS02, APS12, APS15, APS13, APS14  from ApprovalPS where APS04='" & Left(dbFA, 8) & "' and APS05='" & Left(tmpArr(iRound - 1), 8) & "' " & stCon & _
           " union select 2 Od1, APS02, APS12, APS15, APS13, APS14  from ApprovalPS where APS04='" & Left(dbFA, 8) & "' and APS05='" & Left(tmpArr(iRound - 1), 6) & "' " & stCon & _
           " union select 3 Od1, APS02, APS12, APS15, APS13, APS14  from ApprovalPS where APS04='" & Left(dbFA, 6) & "' and APS05='" & Left(tmpArr(iRound - 1), 8) & "' " & stCon & _
           " union select 4 Od1, APS02, APS12, APS15, APS13, APS14  from ApprovalPS where APS04='" & Left(dbFA, 6) & "' and APS05='" & Left(tmpArr(iRound - 1), 6) & "' " & stCon & _
           " union select 5 Od1, APS02, APS12, APS15, APS13, APS14  from ApprovalPS where APS04='" & Left(dbFA, 8) & "' and APS05 is null" & stCon & _
           " union select 6 Od1, APS02, APS12, APS15, APS13, APS14  from ApprovalPS where APS04='" & Left(dbFA, 6) & "' and APS05 is null" & stCon & _
           " union select 7 Od1, APS02, APS12, APS15, APS13, APS14  from ApprovalPS where APS04 is null and APS05='" & Left(tmpArr(iRound - 1), 8) & "' " & stCon & _
           " union select 8 Od1, APS02, APS12, APS15, APS13, APS14  from ApprovalPS where APS04 is null and APS05='" & Left(tmpArr(iRound - 1), 6) & "' " & stCon & _
           " order by Od1, APS02"
            iR = 1
            Set rsQuery = ClsLawReadRstMsg(iR, stSQL)
            If iR = 1 Then
               'Modified by Lydia 2021/03/09 重新整理,逐筆判斷只使用第一筆符合; FCP-58141曾經另外設工程師Email通知並且優先權更大
               rsQuery.MoveFirst
               Do While Not rsQuery.EOF
                  'Added by Lydia 2023/03/22
                  If pKind = "2" Then
                    If "" & rsQuery.Fields("APS13") <> "" And rsQuery.Fields("APS14") <> "" Then
                         m_Memo = "" & rsQuery.Fields("APS13") '通知工程師Email主旨
                         m_Claims = "" & rsQuery.Fields("APS14") '通知工程師Email內文
                         GoTo JumpToEnd
                    End If
                  Else
                  'end 2023/02/22
                    'Modified by Lydia 2022/10/05
                    'If "" & rsQuery.Fields("APS02") & rsQuery.Fields("APS12") <> "" Then
                    '     m_Memo = "" & rsQuery.Fields("APS02")
                    If (pStLang <> "3" And "" & rsQuery.Fields("APS02") <> "") Or _
                        (pStLang = "3" And "" & rsQuery.Fields("APS15") & rsQuery.Fields("APS12") <> "") Then
                         If pStLang <> "3" Then
                            m_Memo = "" & rsQuery.Fields("APS02")
                         Else
                            m_Memo = "" & rsQuery.Fields("APS15")
                         End If
                    'end 2022/10/05
                         m_Claims = "" & rsQuery.Fields("APS12")
                         GoTo JumpToEnd
                    End If
                  End If 'Added by Ldia 2023/03/22
                  rsQuery.MoveNext
               Loop
               'end 2021/03/09
            End If
   Next iRound
   
JumpToEnd:
   pMemo = m_Memo
   pJnClaims = m_Claims 'Added by Lydia 2020/12/30
   'Added by Lydia 2021/02/02 改判斷是否有備註
   If pMemo <> "" Or pJnClaims <> "" Then
       PUB_GetApprovalPS = True
   End If
   'end 2021/02/02
   Set rsQuery = Nothing
End Function

'Added by Lydia 2023/05/26 外專新案認領：Email主旨開頭
Public Function PUB_GetTCNmTitle(ByVal pPA01 As String, ByVal pPA02 As String, ByVal pPA03 As String, ByVal pPA04 As String, ByVal pPA10 As String, ByVal pTCN13 As String, ByVal pSpec As String) As String
Dim strMid As String
   
   PUB_GetTCNmTitle = ""
   If Val(pPA10) > 0 Then
      strMid = strMid & "(已提申)"
   End If
   Select Case pTCN13
      'Modified by Lydia 2024/10/22 debug
      'Case "1", "2"
      '   strMid = strMid & "〔非英說案〕"
      Case "1"
         strMid = strMid & "〔非英說案：有參考本〕"
      Case "2"
         strMid = strMid & "〔非英說案：待確定〕"
      'end 2024/10/22
      Case "3"
         strMid = strMid & "〔非英說案: 已收參考本〕"
      Case "4"
         strMid = strMid & "〔非英說案: 確定無參考本〕"
   End Select
   strMid = strMid & "新案" & pSpec & pPA01 & "-" & pPA02 & IIf(pPA03 & pPA04 <> "000", "-" & pPA03 & "-" & pPA04, "")
   
   PUB_GetTCNmTitle = strMid
End Function

'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：發Email通知並且自動產生FG案之資訊變更964收文
Public Function PUB_GetFCPlinkMC(ByVal pType As String, ByVal pStartDate As String, ByRef p_PA() As String, ByVal pCP09 As String, ByVal pCP10 As String, Optional ByVal pNewCP10 As String = "1001", Optional ByVal pNowCP12 As String, Optional ByVal pNowCP13 As String, Optional ByVal pNowCP14 As String) As Boolean
'若FCP案之基本檔「專利連結通知PA177」為Y時,若藥證號數欄位有資料,同時進行尋找欄位中有相同藥證資料的FG案,進行以下自動收文「資訊變更964」,帶入藥證號至進度備註
'p_PA():專利基本檔
'pNowCP12, pNowCP13, pNowCP14: 目前收文的業務區,智權,承辦
'pStartDate: FCP案來函日期

Dim strR1 As String, intR As Integer
Dim rsRD As New ADODB.Recordset
Dim strBNo As String, strTmp1 As String, strTmp2 As String, strCon1 As String, strCon2 As String
Dim strCP05 As String, strCP06 As String, strCP48 As String
Dim strTo As String, strCC As String, strSubject As String, strCont As String
Dim strFG(1 To 4) As String
Dim bolAddFG As Boolean
Dim stPA89Memo  As String 'Add by Amy 2025/08/05

On Error GoTo ErrHandle

   PUB_GetFCPlinkMC = False
   strCP06 = ""
   bolAddFG = True 'pType 1~6 => 有相同藥證資料的FG案,進行自動收文「資訊變更964」,帶入藥證號至進度備註
   pCP10 = Replace(pCP10, "'", "")
   pNewCP10 = Replace(pNewCP10, "'", "")
   
   '核駁:若有舉發成立確定輸入來函時，發一封Email給承辦工程師
   If pType = "1" Then
      '舉發成立確定==>1.舉發答辯804駁
                     '2.舉發答辯804訴願501駁(經過收文->舉發答辯->核駁->訴願->核駁(現在來函m_NewReceiveNo)的流程)
                     '3.行政訴訟503駁
                     '4.行政訴訟上訴507駁
      Select Case pCP10
         Case "804"  '舉發成立確定==>1.舉發答辯804駁
            Call ClsPDGetCaseProperty(p_PA(1), pCP10, strCon1)
            strCon1 = strCon1 & "駁"
            strCon2 = "若為" & strCon1 & "，請優先報告並告知客戶若未提起訴願，專利連結系統中資訊變更或刪除之期限為AAAA；"
         Case "503"  '舉發成立確定==>3.行政訴訟503駁
            Call ClsPDGetCaseProperty(p_PA(1), pCP10, strCon1)
            strCon1 = strCon1 & "駁"
            strCon2 = "若為" & strCon1 & "，請優先報告並告知客戶若未提起上訴，專利連結系統中資訊變更或刪除之期限為AAAA；"
         Case "807"  '舉發成立確定==>4.行政訴訟上訴507駁
            Call ClsPDGetCaseProperty(p_PA(1), pCP10, strCon1)
            strCon1 = strCon1 & "駁"
            strCon2 = "若為" & strCon1 & "，請優先報告並告知客戶專利連結系統中資訊變更或刪除之期限為AAAA；"
         Case Else   '舉發成立確定==>2.舉發答辯804訴願501駁(經過收文->舉發答辯->核駁->訴願->核駁(現在來函)的流程)
            strR1 = "select c1.cp01||'-'||c1.cp02 as caseno,c1.cp10 as c1cp10,c1.cp43 as c1cp43,c2.cp10 as c2cp10, c2.cp43 as c2cp43" & _
                        " ,c3.cp10 as c3cp10,c3.cp43 as c3cp43 " & _
                        " from caseprogress c1,caseprogress c2,caseprogress c3" & _
                        " where c1.cp09='" & pCP09 & "' and c1.cp43=c2.cp09(+) and c2.cp43=c3.cp09(+)"
            intR = 1
            Set rsRD = ClsLawReadRstMsg(intR, strR1)
            If intI = 1 Then
               If "" & rsRD.Fields("c1cp10") = "501" And "" & rsRD.Fields("c2cp10") = "1002" And "" & rsRD.Fields("c3cp10") = "804" Then
                  strCon1 = "舉發答辯訴願駁"
                  strCon2 = "若為" & strCon1 & "，請優先報告並告知客戶若未提起行政訴訟，專利連結系統中資訊變更或刪除之期限為AAAA；"
               End If
            End If
      End Select
   End If
   '核准：專利權延長415、更正402發Email通知工程師，並且自動設行事曆管控兩天=>行事曆在frm06010602_3產生
   'Modified by Lydia 2024/01/04 程序解除行事曆後再收文資訊變更並直接顯示正確的法限
   'Mark by Lydia 2024/01/04
   'If pType = "2" Then
   '   Call ClsPDGetCaseProperty(p_PA(1), pCP10, strCon1)
   '   strCon1 = "本案" & IIf(strCon1 = "更正", "請求項更正", strCon1) & "已於" & ChangeWStringToTDateString(DBDATE(pStartDate)) & "核准，"
   '   strCon2 = "請待公告日確定後優先報告並告知客戶專利連結系統中資訊變更或刪除之期限為公告後之次日起45天；"
   'End If
   'end 2024/01/04
   '解除行事曆:當程序確認公報刊載日期後解除行事曆自動收文「通知資訊變更961」,發一封Email給承辦工程師
   If pType = "2A" Then
      'bolAddFG = False 'Mark by Lydia 2024/01/04 解除行事曆自動收文「通知資訊變更961」和FG案之「資訊變更964」
      Call ClsPDGetCaseProperty(p_PA(1), pCP10, strCon1)
      strCon1 = "本案" & IIf(strCon1 = "更正", "請求項更正", strCon1) & "預定" & ChangeWStringToTDateString(DBDATE(pStartDate)) & "公告，"
      strCon2 = "請於預定公告日後確認有公告之事實後告知客戶專利連結系統中資訊變更或刪除之期限為AAAA；"
   End If
   '收文特定案件性質, 自動收文「通知資訊變更961」,發一封Email給承辦工程師
   If pType = "3" Then
      If InStr("讓與701、變更401、授權704、專利權部分拋棄439、合併702、繼承703、終止授權705、專利權讓與708、專屬授權709、放棄專利權429、申請權拋棄440", pCP10) > 0 Then
         Call ClsPDGetCaseProperty(p_PA(1), pCP10, strCon1)
         strCon2 = "本案專利權/專利資訊異動事項請告知客戶專利連結系統中資訊變更或刪除之期限為" & strCon1 & "後之次日起45天，請一併向客戶確認事實發生日以計算法定期限。"
      End If
      '若又收文年費，「通知資訊變更」尚未發文時，提醒承辦工程師已收文年費繳費，不需通知資訊變更。
      If pCP10 = "601" Or pCP10 = "605" Then
         If PUB_ChkCPExist(p_PA, "961", 1, strR1) = True Then
           strSubject = p_PA(1) & "-" & p_PA(2) & IIf(p_PA(3) & p_PA(4) <> "000", "-" & p_PA(3) & "-" & p_PA(4), "") & "已收文年費繳費，不需通知資訊變更。"
           strTo = PUB_GetFCPPromoterNo(strR1, "961")
           strCC = PUB_GetFCPEngSup(strTo)
           strSql = "insert into MailCache(MC01,MC02,MC03,MC04,MC07,MC08,MC09)" & _
                " values( '" & strUserNum & "','" & strTo & "'" & _
                ",to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & ChgSQL(strSubject) & "','同主旨','" & strCC & "')"
           cnnConnection.Execute strSql
         End If
      End If
   End If
   '輸入通知年費逾繳1605自動收文「通知資訊變更961」,發一封Email給承辦工程師----frm040331
   If pType = "4" Then
      strR1 = "select np09 from nextprogress WHERE (NP06 is null or NP06='N') and NP07 in ('605','601') and np02='" & p_PA(1) & "'and np03='" & p_PA(2) & "' and np04='" & p_PA(3) & "' and np05='" & p_PA(4) & "' order by np09 desc "
      intR = 1
      Set rsRD = ClsLawReadRstMsg(intR, strR1)
      If intR = 1 Then
         strCon1 = "逾繳函"
         strCon2 = "本案收到年費逾繳函，補繳期限為" & ChangeWStringToTDateString("" & rsRD.Fields("np09")) & "，" & _
                   "請確認程序是否有寄發逾繳提醒，並告知客戶應於專利權消滅之次日起45日內刪除專利資訊，" & _
                   "若已逾45日期限，則無須告知客戶，可請程序銷收文。"
      End If
   End If
   '專利權期滿:由系統每日批次自動檢查當天專利期滿日,自動收文「通知資訊變更961」,發一封Email給承辦工程師
   If pType = "5" Then
      strCon1 = "本案專利權已於" & ChangeWStringToTDateString(DBDATE(pStartDate)) & "消滅，"
      strCon2 = "請優先報告並告知客戶專利連結系統中資訊變更或刪除之期限為AAAA；"
   End If
   '客戶指示閉卷：輸入閉卷913自動收文「通知資訊變更961」,發一封Email給承辦工程師
   If pType = "6" Then
      Call ClsPDGetCaseProperty(p_PA(1), pCP10, strCon1)
      strCon2 = "本案收文" & strCon1 & "，若" & strCon1 & "原因將導致專利權消滅，請告知客戶專利連結系統中資訊變更或刪除之期限為專利權消滅之次日起45天內，" & _
               "若專利登錄是由本所辦理，請客戶同意辦理資訊變更。"
   End If
   
   If strCon1 <> "" Then
      If pStartDate = "" Then pStartDate = strSrvDate(1)
      If pNowCP13 = "" Then pNowCP13 = PUB_GetFCPSalesNo(p_PA(1), p_PA(2), p_PA(3), p_PA(4))
      If pNowCP12 = "" Then pNowCP12 = GetSalesArea(pNowCP13)
      '承辦人預設為最新進度之工程師(來函性質空白以新案核准1001來預設)
      'Modified by Lydia 2023/12/19 判斷承辦人非外專工程師,預設為最新進度之工程師
      If pNowCP14 = "" Or PUB_GetST03(pNowCP14) <> "F21" Then
         pNowCP14 = PUB_GetFCPPromoterNo(pCP09, pNewCP10)
      End If
      
      '期限:官方來函日後之次日起45天
      strTmp2 = CompDate(2, 45, pStartDate)
      If strTmp2 <= strSrvDate(1) Then
         strTmp2 = strSrvDate(1)
         strTmp1 = strSrvDate(1)
      Else
         strTmp1 = PUB_GetFCPOurDeadline(strTmp2) '所限
         If strTmp1 <= strSrvDate(1) Then
            strTmp1 = strSrvDate(1)
         End If
      End If
      strCP05 = strTmp1
      strCP06 = strTmp2
      '承辦期限:所限-5工作天;比照PUB_GetFCPsetCP48Limit
      strCP48 = CompWorkDay(6, strTmp2, 1)
      If strCP48 < strSrvDate(1) Then strCP48 = strSrvDate(1)
      
      If bolAddFG = True Then
         '有相同藥證資料的FG案,進行自動收文「資訊變更964」,帶入藥證號至進度備註
         strR1 = "select mcm02 as CP01,mcm03 as CP02,mcm04 as CP03,mcm05 as CP04,listagg(mc02,',') within group (order by mcm02,mcm03,mcm04,mcm05) as mlist" & _
                 " From medicinecodemap, medicinecode where mcm02<>'" & p_PA(1) & "' and mcm01 in (select mcm01 from medicinecodemap where mcm02='" & p_PA(1) & "' and mcm03='" & p_PA(2) & "' and mcm04='" & p_PA(3) & "' and mcm05='" & p_PA(4) & "')" & _
                 " and mcm01=mc01 group by mcm02,mcm03,mcm04,mcm05"
         intR = 1
         Set rsRD = ClsLawReadRstMsg(intR, strR1)
         If intR = 1 Then
            rsRD.MoveFirst
            Do While Not rsRD.EOF
               strFG(1) = "" & rsRD.Fields("cp01"): strFG(2) = "" & rsRD.Fields("cp02"): strFG(3) = "" & rsRD.Fields("cp03"): strFG(4) = "" & rsRD.Fields("cp04")
               '若無收文「資訊登錄963」,卻收到客戶要做「資訊變更964」指示,假收文「資訊登錄963」
               If PUB_ChkCPExist(strFG, "963") = False Then
                  strBNo = AutoNo("B", 6)
                  strR1 = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP12,CP13,CP14,CP20,CP26,CP27,CP32) VALUES ('" & strFG(1) & "','" & strFG(2) & "','" & strFG(3) & "','" & strFG(4) & "'," & _
                        " '19221111','" & strBNo & "','963','" & pNowCP12 & "','" & pNowCP13 & "','" & pNowCP14 & "','N','N',19221111,'N')"
                  cnnConnection.Execute strR1
               End If
               strBNo = AutoNo("B", 6)
               'Modified by Lydia 2024/01/04 解除行事曆: 收文日改系統日pStartDate =>IIf(pType = "2A", strSrvDate(1), pStartDate)
               strR1 = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP48,CP64) VALUES ('" & strFG(1) & "','" & strFG(2) & "','" & strFG(3) & "','" & strFG(4) & "'," & _
                        IIf(pType = "2A", strSrvDate(1), pStartDate) & "," & strCP05 & "," & strCP06 & ",'" & strBNo & "','964','" & pNowCP12 & "','" & pNowCP13 & "','" & pNowCP14 & "'," & strCP48 & ",'" & ChgSQL("藥證號：" & rsRD.Fields("mlist") & "(" & "專利連結案：" & p_PA(1) & "-" & p_PA(2) & "-" & p_PA(3) & "-" & p_PA(4)) & ");')"
               cnnConnection.Execute strR1
               rsRD.MoveNext
            Loop
         End If
      End If
      If pType = "2A" Then '解除行事曆:當程序確認公報刊載日期後解除行事曆自動收文「通知資訊變更961」
         strR1 = "select cp09,cp64 from caseprogress where cp09='" & pCP09 & "' " '行事曆傳入來函收文號
         intR = 1
         Set rsRD = ClsLawReadRstMsg(intR, strR1)
         If intR = 1 Then
            strBNo = AutoNo("B", 6)
            strR1 = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP48,CP43,CP64) VALUES ('" & p_PA(1) & "','" & p_PA(2) & "','" & p_PA(3) & "','" & p_PA(4) & "'," & _
                     strSrvDate(1) & "," & strCP05 & "," & strCP06 & ",'" & strBNo & "','961','" & pNowCP12 & "','" & pNowCP13 & "','" & pNowCP14 & "'," & strCP48 & ",'" & pCP09 & "','" & ChgSQL("" & rsRD.Fields("cp64")) & "')"
            cnnConnection.Execute strR1
         End If
      End If
      If pType = "3" Or pType = "4" Then  ''3:收文特定案件性質 , 4:輸入通知年費逾繳1605
         strBNo = AutoNo("B", 6)  'pCP09傳入相關收文號
         strR1 = "INSERT INTO CASEPROGRESS (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP48,CP43) VALUES ('" & p_PA(1) & "','" & p_PA(2) & "','" & p_PA(3) & "','" & p_PA(4) & "'," & _
                  strSrvDate(1) & "," & strCP05 & "," & strCP06 & ",'" & strBNo & "','961','" & pNowCP12 & "','" & pNowCP13 & "','" & pNowCP14 & "'," & strCP48 & ",'" & pCP09 & "')"
         cnnConnection.Execute strR1
      End If
      
      'Add by Amy 2025/08/05 後續准駁簡單報告=Y,輸C類來函[主旨]最前面加【請簡單報告】-Winfrey
      '1.若有舉發成立確定輸入來函時 /4.輸入通知年費逾繳1605 --發Email給承辦工程師
      If (pType = "1" Or pType = "4") And p_PA(89) = "Y" Then stPA89Memo = "【請簡單報告】"
      
      strTo = pNowCP14
      
      'CC: 工程師主管、程序管制人員、程序主管
      strTmp1 = PUB_GetFCPEngSup(pNowCP14)
      If strTmp1 <> "" Then strCC = strCC & ";" & strTmp1
      strTmp1 = PUB_GetFCPHandler(p_PA(1), p_PA(2), p_PA(3), p_PA(4))
      strCC = strCC & ";" & strTmp1
      strTmp1 = PUB_GetFCPProSup(strTmp1)
      strCC = strCC & ";" & strTmp1
      strCC = Mid(strCC, 2)
   
      '主旨
      'Modify by Amy 2025/08/05 +stPA89Memo
      strSubject = stPA89Memo & p_PA(1) & "-" & p_PA(2) & IIf(p_PA(3) & p_PA(4) <> "000", "-" & p_PA(3) & "-" & p_PA(4), "") & "請優先報告並告知客戶專利連結系統中資訊變更或刪除之期限"
      If pType = "2" Then '核准輸入(Memo by Amy 2025/08/05 frm06010602_3 已Mark)
         strSubject = strSubject & "事宜"
      ElseIf pType = "4" Then '4:輸入通知年費逾繳1605
         strSubject = stPA89Memo & p_PA(1) & "-" & p_PA(2) & IIf(p_PA(3) & p_PA(4) <> "000", "-" & p_PA(3) & "-" & p_PA(4), "") & "收到" & strCon1 & ", 請告知客戶專利連結系統中資訊變更或刪除之期限為" & ChangeWStringToTDateString(strCP06)
      ElseIf pType = "6" Then  '6:閉卷913
         strSubject = p_PA(1) & "-" & p_PA(2) & IIf(p_PA(3) & p_PA(4) <> "000", "-" & p_PA(3) & "-" & p_PA(4), "") & "客戶指示" & strCon1 & "或放棄專利權, 請告知客戶專利連結系統中資訊變更或刪除之期限為" & ChangeWStringToTDateString(strCP06)
      Else
         strSubject = stPA89Memo & strSubject & "為" & ChangeWStringToTDateString(strCP06)
      End If
      'end 2025/08/05
      '內文
      strCont = p_PA(1) & "-" & p_PA(2) & IIf(p_PA(3) & p_PA(4) <> "000", "-" & p_PA(3) & "-" & p_PA(4), "") & "需通知客戶專利連結系統中專利資訊變更或刪除之期限，"
      If pType = "2" Or pType = "2A" Then   '核准輸入,行事曆解除期限
         strCont = strCont & strCon1 & Replace(strCon2, "AAAA", ChangeWStringToTDateString(strCP06))
      ElseIf pType = "3" Or pType = "4" Or pType = "5" Or pType = "6" Then  ''3:收文特定案件性質 , 4:輸入通知年費逾繳1605 , 5:專利權期滿, 6:閉卷913
         strCont = strCont & Replace(strCon2, "AAAA", ChangeWStringToTDateString(strCP06))
      Else
         strCont = strCont & "本案已於" & ChangeWStringToTDateString(DBDATE(pStartDate)) & "收到" & strCon1 & "，" & _
                  Replace(strCon2, "AAAA", ChangeWStringToTDateString(strCP06))
      End If
      
      strSql = "insert into MailCache(MC01,MC02,MC03,MC04,MC07,MC08,MC09)" & _
            " values( '" & strUserNum & "','" & strTo & "'" & _
            ",to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & ChgSQL(strSubject) & "','" & ChgSQL(strCont) & "','" & strCC & "')"
      cnnConnection.Execute strSql
   End If

   Set rsRD = Nothing
   
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical, "FCP專利連結案自動產生FG案之資訊變更收文"
   End If
End Function

'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：行事曆解除期限檢查
Public Function PUB_ChkFCPlinkSC(ByVal pSC01 As String, ByVal pSC02 As String) As String
Dim intB As Integer, strB1 As String
Dim rsBD As New ADODB.Recordset
Dim strTempInput As String
Dim strKind As String 'Added by Lydia 2023/08/25

   PUB_ChkFCPlinkSC = ""
      
   'Modified by Lydia 2023/08/25
   'strB1 = "select cp09,cpm03||GetRelateCasePropertyName(cp09,'1') as cpm,cp64 from staff_calendar,caseprogress,casepropertymap" & _
           " where sc01='" & pSC01 & "' and sc02='" & pSC02 & "' and sc20 is not null and sc20=cp09(+) and cp01=cpm01(+) and cp10=cpm02(+)"
   strB1 = "select c1.cp09,cpm03||GetRelateCasePropertyName(c1.cp09,'1') as cpm,c1.cp64,c2.cp10 as scp10 from staff_calendar,caseprogress c1,casepropertymap,caseprogress c2" & _
           " where sc01='" & pSC01 & "' and sc02='" & pSC02 & "' and sc20 is not null and sc20=c1.cp09(+) and c1.cp01=cpm01(+) and c1.cp10=cpm02(+) and c1.cp43=c2.cp09(+)"
   intB = 1
   Set rsBD = ClsLawReadRstMsg(intB, strB1)
   If intB = 1 Then
      'Modified by Lydia 2023/08/25 更正用"勘誤日期"／專利權延長用"公告日期"
      'If InStr("" & rsBD.Fields("cp64") & ",", "勘誤日期") = 0 Then
      '   MsgBox "【" & rsBD.Fields("cpm") & "】收文號:" & rsBD.Fields("cp09") & vbCrLf & "尚未輸入勘誤日期，不可解除期限！", vbExclamation + vbOKOnly, "FCP專利連結案管制"
      If "" & rsBD.Fields("scp10") = "415" Then
         strKind = "公告日期"
      Else
         strKind = "勘誤日期"
      End If
      If InStr("" & rsBD.Fields("cp64") & ",", strKind) = 0 Then
         MsgBox "【" & rsBD.Fields("cpm") & "】收文號:" & rsBD.Fields("cp09") & vbCrLf & "尚未輸入" & strKind & "，不可解除期限！", vbExclamation + vbOKOnly, "FCP專利連結案管制"
      'end 2023/08/25
      Else
JumpToRe:
         'Modified by Lydia 2023/08/25 改用變數strKind
         strTempInput = InputBox("【" & rsBD.Fields("cpm") & "】收文號:" & rsBD.Fields("cp09") & vbCrLf & "進度備註:" & rsBD.Fields("cp64") & vbCrLf & vbCrLf & "請在下方輸入" & strKind & "，例如:" & strSrvDate(2), "輸入" & strKind, strTempInput)
         If strTempInput = "" Then
            MsgBox "尚未輸入" & strKind & "，不可解除期限！"
         Else
            If Len(strTempInput) <> 7 Then
               GoTo JumpToRe
            Else
               If ChkDate(strTempInput) = False Then
                  GoTo JumpToRe
               End If
               If strTempInput < strSrvDate(2) Then
                  If MsgBox(strKind & "小於系統日，是否繼續解除期限？", vbInformation + vbYesNo + vbDefaultButton2, "FCP專利連結案管制") = vbNo Then
                     GoTo JumpToRe
                  End If
               End If
            End If
            If Len(strTempInput) = 7 Then
               PUB_ChkFCPlinkSC = strTempInput
            End If
         End If
      End If
   End If
   Set rsBD = Nothing
   
End Function

'Added by Lydia 2023/07/28 外專-FCP專利連結案管制：若又收文年費，「通知資訊變更」尚未發文時，提醒承辦工程師已收文年費繳費，不需通知資訊變更。年費發文時，自動取消收文「通知資訊變更」
Public Sub PUB_ChkFCPlinkYearFee(ByRef p_PA() As String)
Dim strB1 As String, strEx01 As String
Dim strTo As String, strCC As String

   If PUB_ChkCPExist(p_PA, "961", 1, strB1) = False Then
      strEx01 = "Update Caseprogress Set cp57=" & strSrvDate(1) & ", cp58='99', cp64=sqldatet('" & strSrvDate(1) & "')||'取消收文原因：年費發文;'||cp64 Where cp09 = '" & strB1 & "' "
      cnnConnection.Execute strEx01
      strTo = PUB_GetFCPPromoterNo(strB1, "961")
      If strTo <> "" Then
         strCC = PUB_GetFCPEngSup(strCC)
         strEx01 = "insert into MailCache(MC01,MC02,MC03,MC04,MC07,MC08,MC09)" & _
             " values( '" & strUserNum & "','" & strTo & "'" & _
             ",to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & ChgSQL(p_PA(1) & "-" & p_PA(2) & IIf(p_PA(3) & p_PA(4) <> "000", "-" & p_PA(3) & "-" & p_PA(4), "") & "取消收文" & strB1) & "','同主旨','" & strCC & "')"
         cnnConnection.Execute strEx01
      End If
   End If
   
End Sub

'Added by Morgan 2023/10/4
'Modified by Morgan 2024/2/2 +pSubject
'Modified by Morgan 2024/5/24 +pContent,pNoEncrypt
'Modified by Morgan 2024/6/18 +pCCInBox: CC給所內郵件收件者
'EMail薪資單/年終獎金/翻譯費明細
Public Function PUB_SalarySendMail(pSubject As String, pTo As String, pFile As String, ByRef pMsg As String, Optional pToMe As Boolean = False, Optional ByVal pContent As String, Optional pNoEncrypt As Boolean = False, Optional pCCInBox As Boolean = False) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim strAtt As String, strTo As String, strContent As String, strCC As String
   Dim bolOK As Boolean
   
   If pContent <> "" Then
      strContent = pContent
   Else
      strContent = "如旨，見附件。"
   End If
   
   
   'Added by Morgan 2024/9/5
   If pNoEncrypt = False Then
      strContent = strContent & vbCrLf & vbCrLf & "※檔案開啟需輸入密碼，該密碼是""登入薪資系統密碼""。"
   End If
   'end  2024/9/5
   
   pMsg = ""
   stSQL = "select st18,nvl(DECODEPWD(sp04),st26) pwd,st04,st14 from staff,staff_pwd where st01='" & pTo & "' and sp01(+)=st01"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      
      If rsQuery("st04") = "2" Then pMsg = "已離職 "
      
      If pToMe Then
         strTo = strUserNum
         strCC = ""
      Else
         strTo = "" & rsQuery("st18")
         If pCCInBox Then strCC = "" & rsQuery("st14") 'Added by Morgan 2024/6/18
      End If
      
      If strTo = "" Then
         pMsg = pMsg & "未設定外部信箱"
      Else
         bolOK = False
         If pNoEncrypt Then
            strAtt = pFile
            bolOK = True
         Else
            If PUB_AddSalaryPWD(pFile, rsQuery("pwd"), strAtt) = True Then
               bolOK = True
               Kill pFile
            Else
               pMsg = pMsg & "加密失敗"
            End If
         End If
         
         If bolOK Then
            'Modified by Morgan 2024/2/2
            'PUB_SendMail strUserNum, strTo, "", Val(Text1) & "年" & Val(Text2) & "月份薪資單", "如旨，見附件。", , strAtt
            'Modified by Morgan 2025/8/5
            'PUB_SendMail strUserNum, strTo, "", pSubject, strContent, , strAtt, , , , strCC
            '薪資對外發出通知信，不管操作人員，一律以account為寄件者，不用操作人名義
            PUB_SendMail strUserNum, strTo, "", pSubject, strContent, , strAtt, , , , strCC, strAccMailBox, "財務(Account)"
            'end 2024/2/2
            If bolMailSendOk = True Then
               PUB_SalarySendMail = True
            Else
               pMsg = pMsg & "寄信失敗"
            End If
         End If
      End If
   Else
      pMsg = pMsg & "信箱讀取失敗"
   End If
   Set rsQuery = Nothing
End Function

'Added by Morgan 2013/7/30
Public Function PUB_AddSalaryPWD(pInFile As String, pPWD As String, ByRef pOutFile As String) As Boolean
   Dim program_name As String, program_path As String, strCmd As String, ii As Integer
   Dim process_id As Long
   Dim process_handle As Long

   program_name = "pdftk.exe"
       
On Error GoTo ShellError
   
   pOutFile = Left(pInFile, Len(pInFile) - 4) & "S.pdf"
   '加密
   'pdftk.exe "T240275.1701.PDF" output unsecured.pdf user_pw "12345"
   strCmd = "pdftk.exe """ & pInFile & """ output """ & pOutFile & """ user_pw """ & pPWD & """"
   process_id = Shell(strCmd, vbHide)
   process_handle = OpenProcess(PROCESS_TERMINATE, 0, process_id)
   If process_handle <> 0 Then
      For ii = 1 To 10
         If PUB_CheckIsRunning(pub_PdftkName) = True Then
            Sleep 1000
         Else
            Exit For
         End If
      Next
      If ii >= 10 Then
         TerminateProcess process_handle, 0&
         CloseHandle process_handle
         GoTo ExitPoint
      Else
         CloseHandle process_handle
      End If
   Else
      GoTo ExitPoint
   End If
      
   PUB_AddSalaryPWD = True
   Exit Function
   
ShellError:
   MsgBox " " & _
     program_name & vbCrLf & _
     Err.Description, vbOKOnly Or vbExclamation, _
     "Error"
     
ExitPoint:
End Function

'Move by Lydia 2022/09/02 從basQuery搬過來
'Add by Amy 2022/08/19 將共同查詢的名稱查詢,寫成共用Function (傳出來的是SQL語法)
'stChkWay:>0-模糊比對/=1-字首比對 / IsDevelop:包含 投資法務開拓資料 / IsContrast:包含 對造資料
'stCont1~5:對造用 / stNA01:國籍 /stTKSpec:特取文字(以;區隔) : ex.(股)有限公司;有限公司; 一些特取字 Replace再查
'IsNoAdvertise:是查「不得宣傳」'Modify by Amy2023/07/04
Public Function GetSearchNameSql(ByVal stFormN As String, ByVal stFindTxt As String, ByVal stChkWay As String, ByVal IsDevelop As Boolean, ByVal IsContrast As Boolean _
                            , Optional ByVal stCont1 As String, Optional ByVal stCont2 As String, Optional ByVal stCont3 As String, Optional ByVal stCont4 As String, Optional ByVal stCont5 As String _
                            , Optional ByVal stNA01 As String = "", Optional ByVal stTKSpec As String = "", Optional ByVal IsNoAdvertise As Boolean = False) As String
    Dim ii As Integer, jj As Integer
    'Modify by Amy 2023/12/11 +風險檢查對象
    Dim stTB(9) As String, stVTB(9) As String, stContact(2) As String, stF(9) As String, stRepDB As String, stWhere As String, stTemp As String
    Dim stTP(9) As String, stMTp(9) As String
    Dim arrTxt
    Dim stSTB As String 'Add by Amy 2023/06/26
    
    'Memo 共同查詢 用程式有改 [frm100102_1 申請人/frm100114_1 代理人/frm140407 國外潛在客戶/frm210130 國內潛在客戶] 都要測
    'Modify by Amy 2023/12/11 +風險檢查對象
'*** 欄位 ***
    'Memo 'X' as sField->後面會取代成 1-中文/ 2-英文 /3-日文 欄位
    
    'Modify by Amy 2023/12/28 +風險檢查資料維護 (frm12040163)
    If UCase(stFormN) = UCase("frm140419") Or UCase(stFormN) = UCase("frm12040163") Then
      'Modify by Amy 2024/05/03 拆開
        If UCase(stFormN) = UCase("frm140419") Then
        '*** 潛在案量客戶比對 ***
            '客戶檔
            stF(0) = "cu01||cu02 as FNo,cu20 as MailAddr,cu13 as SalesNo,cu12 as SalesAreaNo,cu04 as FName,'X' as sField"
            '代理人檔
            stF(1) = "fa01||fa02 as FNo,fa16 as MailAddr,'' as SalesNo,'' as SalesAreaNo,fa04 as FName,'X' as sField"
            '國外潛在客戶 檔
            stF(2) = "pcu01||pcu02 as FNo,pcu18 as MailAddr,pcu38 as SalesNo,'' as SalesAreaNo,pcu08 as FName,'X' as sField"
            '國內潛在客戶 檔
            stF(3) = "poc01||poc02 as FNo,poc09 as MailAddr,poc13 as SalesNo,'' as SalesAreaNo,poc03 as FName,'X' as sField"
            '不得代理案件之客戶或代理人
            stF(4) = "nt01 as FNo,'' as MailAddr,'' as SalesNo,'' as SalesAreaNo,nt02 as FName,'X' as sField"
            '聯絡人
            stF(5) = "pcc01||'0-'||pcc02 as FNo,pcc08 as MailAddr,'' as SalesNo,'' as SalesAreaNo,PCC05 as FName,'X' as sField"
            '開拓客戶
            stF(6) = "ecd02||'-'||LPAD(ecd01,6,'0') as FNo,ecd13 as MailAddr,'' as SalesNo,'' as SalesAreaNo,NVL(ecd03,'')||NVL(ecd04,'') as FName,9 as sField"
            '國內開拓函特定公司不列印者(共同查詢不需查)
            stF(7) = "'' as FNo,'' as MailAddr,'' as SalesNo,'' as SalesAreaNo,tbnp01 as FName,9 as sField"
            '對造
            stF(8) = "'對造' as FNo,'' as MailAddr,cp13,cp12,R021002||' '||R021001 as FName,9 as sField"
            '風險檢查對象
            stF(9) = "RCL01 as FNo,'' as MailAddr,'' as SalesNo,'' as SalesAreaNo,RCL02 as FName,'X' as sField"
        Else
        '*** 風險檢查資料維護 ***
            '客戶檔
            stF(0) = "cu01||cu02 as FNo,cu11 as mID,cu13 as SalesNo,cu12 as SalesAreaNo,cu04 as FName,'X' as sField,'客戶' as State"
            '代理人檔 (業務設 開發者->建立者)
            stF(1) = "fa01||fa02 as FNo,'' as mID,Nvl(fa94,fa46) as SalesNo,'' as SalesAreaNo,fa04 as FName,'X' as sField,'代理人' as State"
            '國外潛在客戶 檔
            stF(2) = "pcu01||pcu02 as FNo,'' as mID,pcu38 as SalesNo,'' as SalesAreaNo,pcu08 as FName,'X' as sField,'國外潛客' as State"
            '國內潛在客戶 檔
            stF(3) = "poc01||poc02 as FNo,'' as mID,poc13 as SalesNo,'' as SalesAreaNo,poc03 as FName,'X' as sField,'國內潛客' as State"
            '不得代理案件之客戶或代理人
            stF(4) = "nt01 as FNo,'' as mID,'' as SalesNo,'' as SalesAreaNo,nt02 as FName,'X' as sField,'不得代理' as State"
            '聯絡人
            stF(5) = "pcc01||'0-'||pcc02 as FNo,'' as mID,'' as SalesNo,'' as SalesAreaNo,PCC05 as FName,'X' as sField,'聯絡人' as State"
            '開拓客戶(以前秘書寄資料留記錄用,不需查-秀玲)
            stF(6) = ""
            '國內開拓函特定公司不列印者(同 共同查詢,不需查)
            stF(7) = ""
            '對造
            strExc(2) = ",Replace(Replace(Replace(Replace(Replace(Replace(R021002,'@@CP40',''),'@@CP41',''),'@@CP42',''),'@@CP50',''),'@@CP51',''),'@@CP52','')  as FName "
            stF(8) = ",Decode(Substr(r021002,InStr(r021002,'@@')),'@@CP40',1,'@@CP41',2,'@@CP42',3,'@@CP50',1,'@@CP51',2,'@@CP52',3,0) AS sField"
            stF(8) = "R021001 as FNo,'' as mID,cp13,cp12" & strExc(2) & stF(8) & ",'對造' as State"
            '風險檢查對象(風險檢查資料維護已檢查)
            stF(9) = ""
        End If
        'end 2024/05/03
        'Add by Amy 2023/07/04 +不得宣傳(！！！有加欄位最後語法Group by 也要加！！！)
        If IsNoAdvertise = True Then
            '客戶檔
            stF(0) = stF(0) & ",cu10 as Na"
            '代理人檔
            stF(1) = stF(1) & ",fa10 as Na"
            '國外潛在客戶 檔
            stF(2) = stF(2) & ",pcu09 as Na"
            '國內潛在客戶 檔
            stF(3) = stF(3) & ",poc04 as Na"
            '不得代理案件之客戶或代理人
            stF(4) = stF(4) & ",nt08 as Na"
            '聯絡人
            stF(5) = stF(5) & ",Nation as Na"
            '開拓客戶
            stF(6) = stF(6) & ",ecd10 as Na"
            '國內開拓函特定公司不列印者(共同查詢不需查)
            stF(7) = stF(7) & ",'' as Na"
            '對造
            stF(8) = stF(8) & ",'' as Na"
            '風險檢查對象
            stF(9) = stF(9) & ",'' as Na"
        End If
        'Add by Amy 2024/02/16 +狀態
        If UCase(stFormN) = UCase("frm140419") Then
            '客戶檔
            stF(0) = stF(0) & ",CU80 as Status"
            '代理人檔
            stF(1) = stF(1) & ",FA69 as Status"
            '國外潛在客戶 檔
            stF(2) = stF(2) & ",PCU39 as Status"
            '國內潛在客戶 檔
            stF(3) = stF(3) & ",POC14 as Status"
            '不得代理案件之客戶或代理人
            stF(4) = stF(4) & ",Decode(NT21,null,'不得代理','') as Status"
            '聯絡人
            stF(5) = stF(5) & ",'' as Status"
            '開拓客戶
            stF(6) = stF(6) & ",'投法開拓'||Decode(ECD15,null,null,'-'||ECD15) as Status"
            '國內開拓函特定公司不列印者(共同查詢不需查)
            stF(7) = stF(7) & ",'' as Status"
            '對造
            stF(8) = stF(8) & ",Decode(R021004,'1','對造','其他相關人') as Status"
            '風險檢查對象
            stF(9) = stF(9) & ",Decode(RCL24,null,'風險警示','') as Status"
        End If
    Else
    '** 共同查詢 用 **
        '客戶檔
        stF(0) = "' ' AS V,CU01||CU02||Decode(CU02,'0','','＊')||Decode(CU111,'Y','$','')||Decode(CU121,'Y','●','') AS 編號,'中：'||CU04 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,CU79 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & _
                    ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明,'??'||Decode('??',null,'',CU01||CU02) AS OrgN"
        '代理人檔
        stF(1) = "' ' AS V,FA01||FA02||Decode(FA02,'0','','＊')||Decode(FA77,'Y','$','') AS 編號,'中：'||FA04 AS 名稱,NA03 AS 國籍,' ' AS 智權人員,Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態, FA29 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & _
                    ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明,'??'||Decode('??',null,'',FA01||FA02) AS OrgN"
        '國外潛在客戶 檔
        stF(2) = "' ' AS V,PCU01||PCU02||Decode(PCU02,'0','','＊') AS 編號,'中：'||PCU08 AS 名稱,NA03 AS 國籍,PCU38 AS 智權人員,PCU39 AS 狀態,PCU40 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & _
                    ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明,'??'||Decode('??',null,'',PCU01||PCU02) AS OrgN"
        '國內潛在客戶 檔
        stF(3) = "' ' AS V,POC01||POC02||Decode(POC02,'0','','＊') AS 編號,'中：'||POC03 AS 名稱,NA03 AS 國籍,POC13 AS 智權人員,POC14 AS 狀態,POC15 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & _
                    ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明,'??'||Decode('??',null,'',POC01||POC02) AS OrgN"
        '不得代理案件之客戶或代理人
        stF(4) = "' ' AS V,NT01||Decode(NT21,null,'⊕','') AS 編號,'中：'||NT02 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(NT21,null,'不得代理','') AS 狀態, Decode(NT21,null,'','撤銷日期：'||sqldatet(NT21)||'；')||NT20 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & _
                    ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明,'??'||Decode('??',null,'',NT01) AS OrgN"
        '聯絡人
        stF(5) = "' ' AS V,PCC01||'0-'||PCC02 AS 編號,'中：'||PCC05 AS 名稱,NA03 AS 國籍,'%%' AS 智權人員,'@@' AS 狀態,PCC13 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & _
                     ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明,PCC05||Decode(PCC05,null,'',PCC01||'0-'||PCC02) AS OrgN"
        '開拓客戶
        stF(6) = "' ' AS V,ECD02||'-'||LPAD(ECD01,6,'0') AS 編號,Nvl(ECD03,'')||Nvl(ECD04,'') AS 名稱,NA03 AS 國籍,' ' AS 智權人員,'投法開拓'||Decode(ECD15,null,null,'-'||ECD15) AS 狀態,ECD16 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & _
                     ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明,Nvl(ECD03,'')||Nvl(ECD04,'')||Decode(Nvl(ECD03,'')||Nvl(ECD04,''),null,'',ECD02||'-'||LPAD(ECD01,6,'0')) AS OrgN"
        '客戶端平台帳號資料
        stF(7) = "' ' AS V,'平台'||CW01 AS 編號,CW12 AS 名稱,'平台' AS 國籍,' ' AS 智權人員,Nvl(CW19,'') AS 狀態,'' AS 備註,' ' AS 申請國家,'' AS 總收文號,CW03 AS 案件性質,CW01 AS 收文日" & _
                     ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明,CW12||Decode(CW12,null,'','平台'||CW01) AS OrgN"
        '對造
        stF(8) = "' ' AS V,R021001 AS 編號,R021002 AS 名稱,'' AS 國籍,' ' AS 智權人員,Decode(R021004,'1','對造','其他相關人') AS 狀態,'' AS 備註,'' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & _
                     ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明,R021002||Decode(R021002,null,'',R021001) AS OrgN"
        '風險檢查對象
        stF(9) = "' ' AS V,RCL01 AS 編號,'中：'||RCL02 AS 名稱,NA03 AS 國籍,ST02 AS 智權人員,Decode(RCL24,null,'風險警示','') AS 狀態, Decode(RCL24,null,'','撤銷日期：'||sqldatet(RCL24)||'；')||RCL23 AS 備註,' ' AS 申請國家,'' AS 總收文號,'' AS 案件性質,'' AS 收文日" & _
                    ",'' AS 關聯編號,'' AS 關聯名稱,'' AS 關聯關係,'' AS 關聯說明,'??'||Decode('??',null,'',RCL01) AS OrgN"
    '** End 共同查詢 用 **
    End If
    
    '國外潛在客戶
    If UCase(stFormN) = UCase("frm140407") Then
        '多「類別」欄
        stF(0) = Replace(UCase(stF(0)), "ST02 AS 智權人員,", "ST02 AS 智權人員,' ' AS 類別,")
        stF(1) = Replace(UCase(stF(1)), "' ' AS 智權人員,", "' ' AS 智權人員,' ' AS 類別,")
        stF(2) = Replace(UCase(stF(2)), "PCU38 AS 智權人員,", "PCU38 AS 智權人員,DECODE(PCU11," & 國外潛在客戶類別 & ",PCU11) AS 類別,")
        stF(3) = Replace(UCase(stF(3)), "POC13 AS 智權人員,", "POC13 AS 智權人員,' ' AS 類別,")
        stF(4) = Replace(UCase(stF(4)), "ST02 AS 智權人員,", "ST02 AS 智權人員,' ' AS 類別,")
        'stF(5)=聯絡人下方處理
        stF(6) = Replace(UCase(stF(6)), "' ' AS 智權人員,", "' ' AS 智權人員,ECD02 AS 類別,")
        stF(7) = Replace(UCase(stF(7)), "' ' AS 智權人員,", "' ' AS 智權人員,' ' AS 類別,")
        stF(8) = Replace(UCase(stF(8)), "' ' AS 智權人員,", "' ' AS 智權人員,' ' AS 類別,")
        '風險檢查對象
        stF(9) = Replace(UCase(stF(9)), "ST02 AS 智權人員,", "' ' AS 智權人員,' ' AS 類別,")
    End If
'*** End 欄位 ***
    
'*** 子查詢  ***
    'Modify by Amy 2023/01/07 取代改共用函數
    'Modify by Amy 2023/06/26 改抓ReplaceSign DB函數
    'stRepDB = Pub_ReplaceSign(True, "$$")
    'stFindTxt = Pub_ReplaceSign(False, stFindTxt)
    stRepDB = "$$"
    stFindTxt = UCase(stFindTxt)
    
    '特取文字
    If stTKSpec <> MsgText(601) Then
        arrTxt = Split(stTKSpec, ";")
        For ii = LBound(arrTxt) To UBound(arrTxt)
            stRepDB = "Replace(" & stRepDB & ",'" & arrTxt(ii) & "','')"
            stFindTxt = Replace(stFindTxt, arrTxt(ii), "")
        Next ii
    End If
    
    'Modify by Amy 2023/12/28 +風險檢查資料維護 (frm12040163)
    If UCase(stFormN) = UCase("frm140419") Or UCase(stFormN) = UCase("frm12040163") Then
         '潛在案量客戶名稱比對 (因要取代特取字-stTKSpec,故使用原寫法)
         'Modify by Amy 2023/01/22 +ChgSQL 避免單引號錯誤
         stFindTxt = Pub_GetField("Dual", "1=1", "ReplaceSign(TO_MULTI_BYTE(Upper('" & ChgSQL(stFindTxt) & "')))")
    Else
          '改用 ReplaceSign DB函數,避免抓太慢,先用下列語Repalce符號
         stSTB = ",(Select ReplaceSign(TO_MULTI_BYTE(Upper('" & ChgSQL(stFindTxt) & "'))) kw From Dual) x "
    End If
    'end 2023/06/26
    'Add by Amy 2023/06/08 傳入要查之字串先將數字、英文變全型 ex:DB是ＬＧ(全型)用LG(半型)會查不到
    'Mark by Amy 2023/06/13 ex:Y55074 J-star 會查不到,因DB 使用的TO_MULTI_BYTE和PUB_ChangeZIPToSir轉出來的「-」ASCII不同會查不到,TO_MULTI_BYTE()轉出的字貼到記事本會變?
    'stFindTxt = PUB_ChangeZIPToSir(stFindTxt)
    
    '國籍
    If stNA01 <> MsgText(601) Then
        stTemp = "And SubStr(Nation,1,3)='" & stNA01 & "' "
        If stNA01 = "000" Then stTemp = "And SubStr(Nation,1,2)='00' " '判斷台灣案; ex:001,002
    End If
    
    'Modify by Amy 2023/06/08 +TO_MULTI_BYTE 數字、英文變全型再抓, ex:DB是ＬＧ(全型)用LG(半型)會查不到
    'Modify by Amy 2023/06/13 ex:Y55074 J-star 會查不到,使用PUB_ChangeZIPToSir轉-和TO_MULTI_BYTE轉的不一致
    'Modify by Amy 2023/06/26 改抓新欄位
    'Modify by Amy 2023/12/28 +風險檢查資料維護 (frm12040163)
    '潛在案量客戶名稱比對 / 風險檢查資料維護
    If UCase(stFormN) = UCase("frm140419") Or UCase(stFormN) = UCase("frm12040163") Then
        '客戶檔
        If stTemp <> MsgText(601) Then stWhere = Replace(UCase(stTemp), "NATION", "CU10")
        If UCase(stFormN) = UCase("frm12040163") Then
            stVTB(0) = "And " & Replace(stRepDB, "$$", "CU192") & "='" & ChgSQL(stFindTxt) & "' " & stWhere
        Else
            stVTB(0) = "And Instr(" & Replace(stRepDB, "$$", "CU192") & ",'" & ChgSQL(stFindTxt) & "')" & stChkWay & " " & stWhere
        End If
        '代理人檔
        If stTemp <> MsgText(601) Then stWhere = Replace(UCase(stTemp), "NATION", "FA10")
        If UCase(stFormN) = UCase("frm12040163") Then
            stVTB(1) = "And " & Replace(stRepDB, "$$", "FA130") & "='" & ChgSQL(stFindTxt) & "' " & stWhere
        Else
            stVTB(1) = "And Instr(" & Replace(stRepDB, "$$", "FA130") & ",'" & ChgSQL(stFindTxt) & "')" & stChkWay & " " & stWhere
        End If
        '國外潛在客戶檔
        If stTemp <> MsgText(601) Then stWhere = Replace(UCase(stTemp), "NATION", "PCU09")
        If UCase(stFormN) = UCase("frm12040163") Then
            stVTB(2) = "And " & Replace(stRepDB, "$$", "PCU52") & "='" & ChgSQL(stFindTxt) & "' " & stWhere
        Else
            stVTB(2) = "And Instr(" & Replace(stRepDB, "$$", "PCU52") & ",'" & ChgSQL(stFindTxt) & "')" & stChkWay & " " & stWhere
        End If
        '國內潛在客戶檔
        If stTemp <> MsgText(601) Then stWhere = Replace(UCase(stTemp), "NATION", "POC04")
        If UCase(stFormN) = UCase("frm12040163") Then
            stVTB(3) = "And " & Replace(stRepDB, "$$", "POC29") & "='" & ChgSQL(stFindTxt) & "' " & stWhere
        Else
            stVTB(3) = "And Instr(" & Replace(stRepDB, "$$", "POC29") & ",'" & ChgSQL(stFindTxt) & "')" & stChkWay & " " & stWhere
        End If
        '不得代理案件之客戶或代理人
        If stTemp <> MsgText(601) Then stWhere = Replace(UCase(stTemp), "NATION", "NT08")
        If UCase(stFormN) = UCase("frm12040163") Then
            stVTB(4) = "And " & Replace(stRepDB, "$$", "NT32") & "='" & ChgSQL(stFindTxt) & "' " & stWhere
        Else
            stVTB(4) = "And Instr(" & Replace(stRepDB, "$$", "NT32") & ",'" & ChgSQL(stFindTxt) & "')" & stChkWay & " " & stWhere
        End If
        '聯絡人
        If UCase(stFormN) = UCase("frm12040163") Then
            stVTB(5) = "And " & Replace(stRepDB, "$$", "PCC29") & "='" & ChgSQL(stFindTxt) & "' "
        Else
            stVTB(5) = "And Instr(" & Replace(stRepDB, "$$", "PCC29") & ",'" & ChgSQL(stFindTxt) & "')" & stChkWay & " "
        End If
        '開拓客戶
        If UCase(stFormN) = UCase("frm12040163") Then
            stVTB(6) = "And (" & Replace(stRepDB, "$$", "ECD17") & "='" & ChgSQL(stFindTxt) & "' " & _
                                  " Or " & Replace(stRepDB, "$$", "ECD18") & "='" & ChgSQL(stFindTxt) & "' ) "
        Else
            stVTB(6) = "And (Instr(" & Replace(stRepDB, "$$", "ECD17") & ",'" & ChgSQL(stFindTxt) & "')" & stChkWay & _
                                  " Or Instr(" & Replace(stRepDB, "$$", "ECD18") & ",'" & ChgSQL(stFindTxt) & "')" & stChkWay & ") "
        End If
        '國內開拓函特定公司不列印者(共同查詢不需查)
        If UCase(stFormN) = UCase("frm12040163") Then
            stVTB(7) = ""
        Else
            stVTB(7) = "And InStr(" & Replace(stRepDB, "$$", "ReplaceSign(TO_MULTI_BYTE(UPPER(TBNP01)))") & ",'" & ChgSQL(stFindTxt) & "')" & stChkWay & " "
        End If
        '對造-stF(8)下方
        '風險檢查對象
        If stTemp <> MsgText(601) Then stWhere = Replace(UCase(stTemp), "NATION", "RCL08")
        If UCase(stFormN) = UCase("frm12040163") Then
            '避免重覆檢查,故於此表單自行檢查
        Else
            stVTB(9) = "And Instr(" & Replace(stRepDB, "$$", "RCL33") & ",'" & ChgSQL(stFindTxt) & "')" & stChkWay & " " & stWhere
        End If
    Else
    '** 共同查詢 用 **
        '客戶檔
        If stTemp <> MsgText(601) Then stWhere = Replace(UCase(stTemp), "NATION", "CU10")
        stVTB(0) = "Select Distinct cu01 As A1 From Customer" & stSTB & " Where Instr(" & Replace(stRepDB, "$$", "CU192") & "(+),kw)" & stChkWay & " And CU01 is not null " & stWhere
        '代理人檔
        If stTemp <> MsgText(601) Then stWhere = Replace(UCase(stTemp), "NATION", "FA10")
        stVTB(1) = "Select Distinct fa01 As A1 From Fagent" & stSTB & " Where Instr(" & Replace(stRepDB, "$$", "FA130") & "(+),kw)" & stChkWay & " And FA01 is not null " & stWhere
        '國外潛在客戶檔
        If stTemp <> MsgText(601) Then stWhere = Replace(UCase(stTemp), "NATION", "PCU09")
        stVTB(2) = "Select Distinct pcu01 As A1 From PotCustomer" & stSTB & " Where Instr(" & Replace(stRepDB, "$$", "PCU52") & "(+),kw)" & stChkWay & " And PCU01 is not null " & stWhere
        '國內潛在客戶檔
        If stTemp <> MsgText(601) Then stWhere = Replace(UCase(stTemp), "NATION", "POC04")
        stVTB(3) = "Select Distinct poc01 As A1 From PotCustomer1" & stSTB & " Where Instr(" & Replace(stRepDB, "$$", "POC29") & "(+),kw)" & stChkWay & " And POC01 is not null " & stWhere
        '不得代理案件之客戶或代理人
        If stTemp <> MsgText(601) Then stWhere = Replace(UCase(stTemp), "NATION", "NT08")
        stVTB(4) = "Select Distinct nt01 As A1 From NotAgent" & stSTB & " Where Instr(" & Replace(stRepDB, "$$", "NT32") & "(+),kw)" & stChkWay & " And NT01 is not null " & stWhere
        '聯絡人
        stVTB(5) = "Select * From PotCustCont" & stSTB & " Where Instr(" & Replace(stRepDB, "$$", "PCC29") & "(+),kw)" & stChkWay & " And PCC01 is not null "
        '開拓客戶
        'Memo by Amy 2023/06/26 外部結合的運算子 (+) 不可以用在運算元 OR 或 IN中,故拆成兩句
        stVTB(6) = "Select Distinct Nvl(ecd01,'')||Nvl(ecd02,'') as A1 From ExPandCusDetail" & stSTB & " Where Instr(" & Replace(stRepDB, "$$", "ECD17") & "(+),kw)" & stChkWay & " And ECD01 is not null "
        stVTB(6) = stVTB(6) & "Union "
        stVTB(6) = stVTB(6) & "Select Distinct Nvl(ecd01,'')||Nvl(ecd02,'') as A1 From ExPandCusDetail" & stSTB & " Where Instr(" & Replace(stRepDB, "$$", "ECD18") & "(+),kw)" & stChkWay & " And ECD01 is not null "
        '客戶端平台帳號資料
        stVTB(7) = "Select Distinct cw01 as A1 From CustWeb" & stSTB & " Where Instr(" & Replace(stRepDB, "$$", "CW20") & "(+),kw)" & stChkWay & " And CW01 is not null "
        '對造-stF(8)下方
        '風險檢查對象
        If stTemp <> MsgText(601) Then stWhere = Replace(UCase(stTemp), "NATION", "RCL08")
        stVTB(9) = "Select Distinct RCL01 As A1 From RiskCheckList" & stSTB & " Where Instr(" & Replace(stRepDB, "$$", "RCL33") & "(+),kw)" & stChkWay & " And RCL01 is not null " & stWhere
    '**End 共同查詢 用 **
    End If
    'end 2023/06/26
    'end 2023/06/13
    'end 2023/06/08
'*** End 子查詢 ***

'*** 客戶/代理人/國外(內)潛在客戶/不得代理案件之客戶或代理人/風險檢查對象 ***
    For ii = 0 To 2
        '顯示欄位
        Select Case ii
             Case 0 '中文
                stMTp(0) = Replace(UCase(stF(0)), "'??'", "CU04") '客戶
                stMTp(1) = Replace(UCase(stF(1)), "'??'", "FA04") '代理人
                stMTp(2) = Replace(UCase(stF(2)), "'??'", "PCU08") '國外潛在客戶
                stMTp(3) = Replace(UCase(stF(3)), "'??'", "POC03") '國內潛在客戶
                stMTp(4) = Replace(UCase(stF(4)), "'??'", "NT02") '不得代理案件之客戶或代理人
                stMTp(9) = Replace(UCase(stF(9)), "'??'", "RCL33") '風險檢查對象
            Case 1 '英文
                stMTp(0) = Replace(Replace(Replace(UCase(stF(0)), "CU04", "CU05||' '||CU88||' '||CU89||' '||CU90"), "中：", "英："), "'??'", "CU05||CU88||CU89||CU90")
                stMTp(1) = Replace(Replace(Replace(UCase(stF(1)), "FA04", "FA05||' '||FA63||' '||FA64||' '||FA65"), "中：", "英："), "'??'", "FA05||FA63||FA64||FA65")
                stMTp(2) = Replace(Replace(Replace(UCase(stF(2)), "PCU08", "PCU03||' '||PCU04||' '||PCU05||' '||PCU06"), "中：", "英："), "'??'", "PCU03||PCU04||PCU05||PCU06")
                stMTp(3) = Replace(Replace(Replace(UCase(stF(3)), "POC03", "POC23||' '||POC24||' '||POC25||' '||POC26"), "中：", "英："), "'??'", "POC23||POC24||POC25||POC26")
                stMTp(4) = Replace(Replace(Replace(UCase(stF(4)), "NT02", "NT03||' '||NT04||' '||NT05||' '||NT06"), "中：", "英："), "'??'", "NT03||NT04||NT05||NT06")
                stMTp(9) = Replace(Replace(Replace(UCase(stF(9)), "RCL02", "RCL03||' '||RCL04||' '||RCL05||' '||RCL06"), "中：", "英："), "'??'", "RCL03||RCL04||RCL05||RCL06") '風險檢查對象
            Case 2 '日文
                stMTp(0) = Replace(Replace(Replace(UCase(stF(0)), "CU04", "CU06"), "中：", "日："), "'??'", "CU06")
                stMTp(1) = Replace(Replace(Replace(UCase(stF(1)), "FA04", "FA06"), "中：", "日："), "'??'", "FA06")
                stMTp(2) = Replace(Replace(Replace(UCase(stF(2)), "PCU08", "PCU07"), "中：", "日："), "'??'", "PCU07")
                stMTp(3) = Replace(Replace(Replace(UCase(stF(3)), "POC03", "POC27"), "中：", "日："), "'??'", "POC27")
                stMTp(4) = Replace(Replace(Replace(UCase(stF(4)), "NT02", "NT07"), "中：", "日："), "'??'", "NT07")
                stMTp(9) = Replace(Replace(Replace(UCase(stF(9)), "RCL02", "RCL07"), "中：", "日："), "'??'", "RCL07") '風險檢查對象
        End Select
        'Modify by Amy 2023/12/28 +(frm12040163)
        '潛在案量客戶名稱比對 / 風險檢查資料維護
        If UCase(stFormN) = UCase("frm140419") Or UCase(stFormN) = UCase("frm12040163") Then
            stMTp(0) = Replace(UCase(stMTp(0)), "'X'", ii + 1) '客戶
            stMTp(1) = Replace(UCase(stMTp(1)), "'X'", ii + 1) '代理人
            stMTp(2) = Replace(UCase(stMTp(2)), "'X'", ii + 1) '國外潛在客戶
            stMTp(3) = Replace(UCase(stMTp(3)), "'X'", ii + 1) '國內潛在客戶
            stMTp(4) = Replace(UCase(stMTp(4)), "'X'", ii + 1) '不得代理案件之客戶或代理人
            If UCase(stFormN) = UCase("frm140419") Then
               stMTp(9) = Replace(UCase(stMTp(9)), "'X'", ii + 1) '風險檢查對象
            End If
        End If
        '查詢欄位
        Select Case ii
            Case 0 '中文
                stTP(0) = stVTB(0) '客戶
                stTP(1) = stVTB(1) '代理人
                stTP(2) = stVTB(2) '國外潛在客戶
                stTP(3) = stVTB(3) '國內潛在客戶
                stTP(4) = stVTB(4) '不得代理案件之客戶或代理人
                stTP(9) = stVTB(9) '風險檢查對象
            Case 1 '英文
                stTP(0) = Replace(UCase(stVTB(0)), "CU192", "CU193")
                stTP(1) = Replace(UCase(stVTB(1)), "FA130", "FA131")
                stTP(2) = Replace(UCase(stVTB(2)), "PCU52", "PCU53")
                stTP(3) = Replace(UCase(stVTB(3)), "POC29", "POC30")
                stTP(4) = Replace(UCase(stVTB(4)), "NT32", "NT33")
                stTP(9) = Replace(UCase(stVTB(9)), "RCL33", "RCL34") '風險檢查對象
            Case 2 '日文
                stTP(0) = Replace(UCase(stVTB(0)), "CU192", "CU194")
                stTP(1) = Replace(UCase(stVTB(1)), "FA130", "FA132")
                stTP(2) = Replace(UCase(stVTB(2)), "PCU52", "PCU54")
                stTP(3) = Replace(UCase(stVTB(3)), "POC29", "POC31")
                stTP(4) = Replace(UCase(stVTB(4)), "NT32", "NT34")
                stTP(9) = Replace(UCase(stVTB(9)), "RCL33", "RCL35") '風險檢查對象
        End Select
        'Modify by Amy 2023/12/28 +frm12040163
        '潛在案量客戶名稱比對 / 風險檢查資料維護
        If UCase(stFormN) = UCase("frm140419") Or UCase(stFormN) = UCase("frm12040163") Then
            stTB(0) = stTB(0) & " Union All Select " & stMTp(0) & " From Customer Where 1=1 " & stTP(0)
            stTB(1) = stTB(1) & " Union All Select " & stMTp(1) & " From Fagent Where 1=1 " & stTP(1)
            stTB(2) = stTB(2) & " Union All Select " & stMTp(2) & " From PotCustomer Where 1=1 " & stTP(2)
            stTB(3) = stTB(3) & " Union All Select " & stMTp(3) & " From PotCustomer1 Where 1=1 " & stTP(3)
            stTB(4) = stTB(4) & " Union All Select " & stMTp(4) & " From NotAgent Where 1=1 " & stTP(4)
            If UCase(stFormN) = UCase("frm140419") Then
               stTB(9) = stTB(9) & " Union All Select " & stMTp(9) & " From RiskCheckList Where 1=1 " & stTP(9) '風險檢查對象
            End If
        Else
        '**共同查詢 用 **
            stTB(0) = stTB(0) & " Union All Select " & stMTp(0) & " From Customer,Nation,Staff,(" & stTP(0) & ") A Where cu10=na01(+) And cu01=A.A1 And cu13=st01(+) "
            stTB(1) = stTB(1) & " Union All Select " & stMTp(1) & " From Fagent,Nation,(" & stTP(1) & ") A Where fa10=na01(+) And fa01=A.A1 "
            stTB(2) = stTB(2) & " Union All Select " & stMTp(2) & " From PotCustomer,Nation,Staff,(" & stTP(2) & ") A Where pcu09=na01(+) And pcu01=A.A1 And SubStr(LTrim(pcu38),1,5)=st01(+) "
            stTB(3) = stTB(3) & " Union All Select " & stMTp(3) & " From PotCustomer1,Nation,Staff,(" & stTP(3) & ") A Where poc04=na01(+) And poc01=A.A1 And poc13=st01(+) "
            stTB(4) = stTB(4) & " Union All Select " & stMTp(4) & " From NotAgent,Nation,Staff,(" & stTP(4) & ") A Where nt08=na01(+) And nt01=A.A1 And nt18=st01(+) "
            'Modify by Amy 2024/01/31
            stTB(9) = stTB(9) & " Union All Select " & stMTp(9) & " From RiskCheckList,Nation,Staff,(" & stTP(9) & ") A Where RCL08=na01(+) And RCL01=A.A1 And RCL22=st01(+) " '風險檢查對象
        '**End 共同查詢 用 **
        End If
    Next ii
'*** End 客戶/代理人/國外(內)潛在客戶/不得代理案件之客戶或代理人/風險檢查對象 ***
'*** 聯絡人 ***
    For ii = 0 To 2
        Select Case ii
            Case 0 '中文
                stTP(0) = stVTB(5)
                stMTp(0) = stF(5)
            Case 1 '英文
                stTP(0) = Replace(UCase(stVTB(5)), "PCC29", "PCC27")
                stMTp(0) = Replace(Replace(UCase(stF(5)), "PCC05", "PCC03"), "中：", "英：")
            Case 2 '日文
                stTP(0) = Replace(UCase(stVTB(5)), "PCC29", "PCC28")
                stMTp(0) = Replace(Replace(UCase(stF(5)), "PCC05", "PCC04"), "中：", "日：")
        End Select
         'Modify by Amy 2023/12/28 +frm12040163
        '潛在案量客戶名稱比對 / 風險檢查資料維護
        If UCase(stFormN) = UCase("frm140419") Or UCase(stFormN) = UCase("frm12040163") Then
            stTP(0) = "Select * From PotCustCont Where 1=1 " & stTP(0)
        End If
        'Add by Amy 2023/12/27 風險檢查資料維護
        If UCase(stFormN) = UCase("frm12040163") Then stTP(1) = Replace(UCase(stMTp(0)), "'X'", ii + 1)
        
        For jj = 0 To 3
            Select Case jj
                Case 0 '客戶
                    'Modify by Amy 2023/12/28 +frm12040163
                    '潛在案量客戶名稱比對 / 風險檢查資料維護
                    If UCase(stFormN) = UCase("frm140419") Or UCase(stFormN) = UCase("frm12040163") Then
                        'Modify by Amy 2023/12/27 潛在案量客戶名稱比對 才replace
                        If UCase(stFormN) = UCase("frm140419") Then stTP(1) = Replace(UCase(stMTp(0)), "'X'", ii)
                        'Add by Amy 2023/07/04 +不得宣傳,抓國籍欄位
                        If IsNoAdvertise = True Then stTP(1) = Replace(UCase(stTP(1)), "NATION", "Cu10")
                    Else
                    '**共同查詢 用 **
                        '國外潛在客戶
                        If UCase(stFormN) = UCase("frm140407") Then
                            '多「類別」欄
                            stTP(1) = Replace(UCase(stMTp(0)), "'%%' AS 智權人員,", "ST02 AS 智權人員,' ' AS 類別,")
                        Else
                            stTP(1) = Replace(UCase(stMTp(0)), "'%%' AS 智權人員,", "ST02 AS 智權人員,")
                        End If
                        stTP(1) = Replace(UCase(stTP(1)), "'@@' AS 狀態,", UCase("Decode(CU142,null,CU80,GetDizhang(CU142,'Y')) AS 狀態,"))
                    '** End 共同查詢 用 **
                    End If
                    stTP(2) = "Staff,Customer"
                    stTP(3) = "And cu10=na01(+) And pcc01=cu01(+) And cu02='0' And cu13=st01(+) "
                Case 1 '代理人
                     'Modify by Amy 2023/12/28 +frm12040163
                    '潛在案量客戶名稱比對 /風險檢查資料維護
                    If UCase(stFormN) = UCase("frm140419") Or UCase(stFormN) = UCase("frm12040163") Then
                        'Modify by Amy 2023/12/27 潛在案量客戶名稱比對 才replace
                        If UCase(stFormN) = UCase("frm140419") Then stTP(1) = Replace(UCase(stMTp(0)), "'X'", ii)
                        'Add by Amy 2023/07/04 +不得宣傳,抓國籍欄位
                        If IsNoAdvertise = True Then stTP(1) = Replace(UCase(stTP(1)), "NATION", "Fa10")
                    Else
                    '**共同查詢 用 **
                        '國外潛在客戶
                        If UCase(stFormN) = UCase("frm140407") Then
                            '多「類別」欄
                            stTP(1) = Replace(UCase(stMTp(0)), "'%%' AS 智權人員,", "' ' AS 智權人員,' ' AS 類別,")
                        Else
                            stTP(1) = Replace(UCase(stMTp(0)), "'%%' AS 智權人員,", "' ' AS 智權人員,")
                        End If
                        stTP(1) = Replace(UCase(stTP(1)), "'@@' AS 狀態,", UCase("Decode(FA103,null,FA69,GetDizhang(FA103,'Y')) AS 狀態,"))
                    '** End 共同查詢 用 **
                    End If
                    stTP(2) = "Fagent"
                    stTP(3) = "And fa10=na01(+) And pcc01=fa01(+) And fa02='0' "
                Case 2 '國外潛在客戶
                     'Modify by Amy 2023/12/28 +frm12040163
                    '潛在案量客戶名稱比對 / 風險檢查資料維護
                    If UCase(stFormN) = UCase("frm140419") Or UCase(stFormN) = UCase("frm12040163") Then
                        'Modify by Amy 2023/12/27 潛在案量客戶名稱比對 才replace
                        If UCase(stFormN) = UCase("frm140419") Then stTP(1) = Replace(UCase(stMTp(0)), "'X'", ii)
                        'Add by Amy 2023/07/04 +不得宣傳,抓國籍欄位
                        If IsNoAdvertise = True Then stTP(1) = Replace(UCase(stTP(1)), "NATION", "Pcu09")
                    Else
                    '**共同查詢 用 **
                        '國外潛在客戶
                        If UCase(stFormN) = UCase("frm140407") Then
                            '多「類別」欄
                            stTP(1) = Replace(UCase(stMTp(0)), "'%%' AS 智權人員,", "PCU38 AS 智權人員,' ' AS 類別,")
                        Else
                            stTP(1) = Replace(stMTp(0), "'%%' AS 智權人員,", "PCU38 AS 智權人員,")
                        End If
                        stTP(1) = Replace(UCase(stTP(1)), "'@@' AS 狀態,", "PCU39 AS 狀態,")
                    '** End 共同查詢 用 **
                    End If
                    stTP(2) = "Staff,PotCustomer"
                    stTP(3) = "And pcu09=na01(+) And pcc01=pcu01(+) And pcu02='0' And SubStr(LTrim(pcu38),1,5)=st01(+) "
                Case 3 '國內潛在客戶
                    'Modify by Amy 2023/12/28 +frm12040163
                    '潛在案量客戶名稱比對 / 風險檢查資料維護
                    If UCase(stFormN) = UCase("frm140419") Or UCase(stFormN) = UCase("frm12040163") Then
                        'Modify by Amy 2023/12/27 潛在案量客戶名稱比對 才replace
                        If UCase(stFormN) = UCase("frm140419") Then stTP(1) = Replace(UCase(stMTp(0)), "'X'", ii)
                        'Add by Amy 2023/07/04 +不得宣傳,抓國籍欄位
                        If IsNoAdvertise = True Then stTP(1) = Replace(UCase(stTP(1)), "NATION", "Poc04")
                    Else
                    '**共同查詢 用 **
                        '國外潛在客戶
                        If UCase(stFormN) = UCase("frm140407") Then
                            '多「類別」欄
                            stTP(1) = Replace(UCase(stMTp(0)), "'%%' AS 智權人員,", "POC13 AS 智權人員,' ' AS 類別,")
                        Else
                            stTP(1) = Replace(UCase(stMTp(0)), "'%%' AS 智權人員,", "POC13 AS 智權人員,")
                        End If
                        stTP(1) = Replace(UCase(stTP(1)), "'@@' AS 狀態,", "POC14 AS 狀態,")
                    '** End 共同查詢 用 **
                    End If
                    stTP(2) = "Staff,PotCustomer1"
                    stTP(3) = "And poc04=na01(+) And pcc01=poc01(+) And poc02='0' and poc13=st01(+) "
            End Select
            stContact(ii) = stContact(ii) & " Union All Select " & stTP(1) & " From Nation," & stTP(2) & ",(" & stTP(0) & ") A Where 1=1 " & stTP(3)
        Next jj
        stTB(5) = stTB(5) & stContact(ii)
    Next ii
'*** End 聯絡人 ***
'*** 開拓客戶 ***
    If IsDevelop = True Then
        'Modify by Amy 2023/06/26 外部結合的運算子 (+) 不可以用在運算元 OR 或 IN中,故拆開
        stTP(0) = UCase(stVTB(6))
        stMTp(0) = UCase(stF(6))
        'Modify by Amy 2023/12/28 +frm12040163 (2024/05/03 不需查)
        '潛在案量客戶名稱比對 /風險檢查資料維護
        If UCase(stFormN) = UCase("frm140419") Or UCase(stFormN) = UCase("frm12040163") Then
            stTB(6) = stTB(6) & " Union All Select " & stMTp(0) & " From ExPandCusdetail,ExPandCusAttr Where 1=1 " & stTP(0)
            
            stTP(0) = Replace(Replace(stTP(0), "ECD17", "ECD19"), "ECD18", "ECD20")
            stMTp(0) = Replace(stMTp(0), "NVL(ECD03,'')||NVL(ECD04,'')", "NVL(ECD11,'')||NVL(ECD12,'')")
            stTB(6) = stTB(6) & " Union All Select " & stMTp(0) & " From ExPandCusdetail,ExPandCusAttr Where 1=1 " & stTP(0)
        '共同查詢
        Else
            stTB(6) = stTB(6) & " Union All Select " & stMTp(0) & " From ExPandCusdetail,ExPandCusAttr,Nation,(" & stTP(0) & ") A Where ecd10=na01(+) And ecd02=eca01(+) And Nvl(ecd01,'')||Nvl(ecd02,'')=A.A1 "
            
            stTP(0) = Replace(Replace(stTP(0), "ECD17", "ECD19"), "ECD18", "ECD20")
            stMTp(0) = Replace(stMTp(0), "NVL(ECD03,'')||NVL(ECD04,'')", "NVL(ECD11,'')||NVL(ECD12,'')")
            stTB(6) = stTB(6) & " Union All Select " & stMTp(0) & " From ExPandCusdetail,ExPandCusAttr,Nation,(" & stTP(0) & ") A Where ecd10=na01(+) And ecd02=eca01(+) And Nvl(ecd01,'')||Nvl(ecd02,'')=A.A1 "
        End If
        'end 2023/06/26
    End If
'*** End 開拓客戶 ***
'Modify by Amy 2023/12/28 +風險檢查資料維護 (frm12040163)
If UCase(stFormN) = UCase("frm12040163") Then
   '風險檢查資料維護名稱查詢,同 共同查詢 資料,故stTB(7)抓 國內開拓函特定公司不列印者,不需查
'潛在案量客戶名稱比對
ElseIf UCase(stFormN) = UCase("frm140419") Then
    '國內開拓函特定公司不列印者
    stTB(7) = " Union All Select " & stF(7) & " From TMBulletinnp Where 1=1 " & stVTB(7)
'共同查詢
Else
    '客戶端平台帳號資料
    'Modify by Amy 2023/06/13 ex:Y55074 J-star 會查不到,使用PUB_ChangeZIPToSir轉-和TO_MULTI_BYTE轉的不一致
    'Modify by Amy 2023/06/08 +TO_MULTI_BYTE 數字、英文變全型再抓, ex:DB是ＬＧ(全型)用LG(半型)會查不到
    'Modify by Amy 2023/06/26 改抓新欄位
    stTB(7) = " Union All Select " & stF(7) & " From CustWeb,(" & stVTB(7) & ") A Where Nvl(cw01,'')=A.A1 "
End If
'end 2023/06/26

'*** 對造 ***
    If IsContrast = True Then
        'Modify by Amy 2023/12/28 +風險檢查資料維護 (frm12040163)
        If UCase(stFormN) = UCase("frm12040163") Then
            'Modify by Amy 2024/01/23 使用>進度資料會以☆區隔
            Call Pub_ProcR100102_1(strUserNum & "@" & stFormN, stCont1, stCont2, stCont3, stCont4, stCont5, stFindTxt, ">0", True, True)
        '潛在案量客戶名稱比對
        ElseIf UCase(stFormN) = UCase("frm140419") Then
            Call Pub_ProcR100102_1(strUserNum & "@" & stFormN, stCont1, stCont2, stCont3, stCont4, stCont5, stFindTxt, stChkWay, True)
        Else
        '**共同查詢 用 **
            Call Pub_ProcR100102_1(strUserNum & "@" & stFormN, stCont1, stCont2, stCont3, stCont4, stCont5, stFindTxt, stChkWay, , True)
            '名稱前顯示查到的欄位為中/英/日
            stMTp(0) = "Decode(InStr(R021002,'@@@'),0,'##','***：'||SubStr(R021002,1,InStr(r021002,'@@@')-1))"
            stMTp(1) = "@@@"
            stMTp(4) = "***"
            stMTp(3) = stMTp(0)
            For ii = 0 To 2
                Select Case ii
                    Case 0
                        strExc(1) = "中"
                    Case 1
                        strExc(1) = "英"
                    Case 2
                        strExc(1) = "日"
                End Select
                stMTp(2) = "@@CP4" & ii
                stMTp(3) = Replace(UCase(stMTp(3)), UCase(stMTp(1)), UCase(stMTp(2))) '將@@@取代為欄名
                stMTp(3) = Replace(UCase(stMTp(3)), UCase(stMTp(4)), strExc(1)) '將***取代為中/英/日
                stMTp(3) = Replace(UCase(stMTp(3)), "'##'", UCase(stMTp(0))) '將'##'取代為Decode(...)
            Next ii
            For ii = 0 To 2
                Select Case ii
                    Case 0
                        strExc(1) = "中"
                    Case 1
                        strExc(1) = "英"
                    Case 2
                        strExc(1) = "日"
                End Select
                stMTp(2) = "@@CP5" & ii
                stMTp(3) = Replace(UCase(stMTp(3)), UCase(stMTp(1)), UCase(stMTp(2)))
                stMTp(3) = Replace(UCase(stMTp(3)), UCase(stMTp(4)), strExc(1))
                If ii = 2 Then
                    stMTp(3) = Replace(UCase(stMTp(3)), "'##'", "R021002")
                Else
                    stMTp(3) = Replace(UCase(stMTp(3)), "'##'", UCase(stMTp(0)))
                End If
            Next ii
        '** End 共同查詢 用 **
        End If
        'Modify by Amy 2024/01/17 +風險檢查資料維護 (frm12040163)
        If UCase(stFormN) = UCase("frm12040163") Then
            stTB(8) = " Union All Select Distinct " & stF(8) & " From R100102_1,CaseProgress Where ID='" & strUserNum & "@" & stFormN & "' And R021004<3 And R021006=cp09(+) "
        '潛在案量客戶名稱比對 / 風險檢查資料維護
        ElseIf UCase(stFormN) = UCase("frm140419") Then
            stTB(8) = " Union All Select Distinct " & stF(8) & " From R100102_1,CaseProgress Where ID='" & strUserNum & "@" & stFormN & "' And R021004<3 And R021006=cp09(+) "
        '共同查詢
        Else
            stF(8) = Replace(UCase(stF(8)), "R021002 AS 名稱", UCase(stMTp(3)) & " AS 名稱 ")
            stF(8) = Replace(UCase(stF(8)), UCase("R021002 AS OrgN"), UCase("replace(replace(replace(replace(replace(replace(R021002,'@@CP40',''),'@@CP41',''),'@@CP42',''),'@@CP50',''),'@@CP51',''),'@@CP52','') AS OrgN"))
            stTB(8) = " Union All Select Distinct " & stF(8) & " From R100102_1 Where ID='" & strUserNum & "@" & stFormN & "' And R021004<3 "
        End If
    End If
'*** End 對造 ***
    'Memo by Amy 使用Union All因查「金杜」應出現2筆(因對造 中/日 欄都有金杜)-2020/09/04
    For ii = LBound(stTB) To UBound(stTB)
        GetSearchNameSql = GetSearchNameSql & stTB(ii)
    Next ii
    If GetSearchNameSql <> MsgText(601) Then
        GetSearchNameSql = Mid(GetSearchNameSql, 12)
        'Add by Amy 2023/07/04 +不得宣傳
        'Modify by Amy 2024/02/16 +狀態
        If IsNoAdvertise = True Then
            GetSearchNameSql = "Select O.*,Min(cr02) as CDate  From(" & GetSearchNameSql & ") O,ContactRecord " & _
                                                      "Where Instr(cr05,'A14')>0 And Cr03=FNo(+) And Fno is not null " & _
                                                      "Group by FNo,MailAddr,SalesNo,SalesAreaNo,FName,sField,Na,Status "
        End If
    End If
End Function

'Added by Lydia 2024/04/15 顧問聘任LA取得聘任期間、服務次數,服務時數
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function Pub_GetLAforTimes(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, _
           Optional ByRef pDate01 As String, Optional ByRef pDate02 As String, Optional ByRef pCP15 As String, Optional ByRef pHours As String, Optional ByRef pTimes As String) As Boolean
Dim rsBD As New ADODB.Recordset
Dim intB As Integer
Dim strB1 As String

   Pub_GetLAforTimes = False
   If pCP01 <> "LA" And Len(pCP02) <> 6 Then Exit Function
   
   pDate01 = "": pDate02 = "": pHours = "": pTimes = "": pCP15 = ""
   strB1 = "select c1.cp09,c1.cp15,substr(sqldatet(c1.cp53),1,9) cp53t,substr(sqldatet(c1.cp54),1,9) cp54t,sum(x02) x02,sum(x03) x03 from caseprogress c1" & _
                 ",(select cp09 as x00,cp05 as x01, 1 as x02,nvl(cp113,0) as x03 from caseprogress where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & Left(pCP03 & "0", 1) & "' and cp04='" & Left(pCP04 & "00", 2) & "' and substr(cp09,1,1)='A' and nvl(cp18,0)=0) x1 " & _
                 "where c1.cp09 in ( " & _
                 "select substr(max(cp05||cp09),9,9) mno from caseprogress where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & Left(pCP03 & "0", 1) & "' and cp04='" & Left(pCP04 & "00", 2) & "' and cp10='0' and cp158=0 and cp159=0 " & _
                  ") and x01>=c1.cp53 and x01<=c1.cp54 group by c1.cp09,c1.cp15,substr(sqldatet(c1.cp53),1,9) ,substr(sqldatet(c1.cp54),1,9) "
   intB = 1
   Set rsBD = ClsLawReadRstMsg(intB, strB1)
   If intB = 1 Then
      If "" & rsBD.Fields("cp53t") <> "" And "" & rsBD.Fields("cp54t") <> "" Then
         pDate01 = "" & rsBD.Fields("cp53t")
         pDate02 = "" & rsBD.Fields("cp54t")
         If Val("" & rsBD.Fields("cp15")) > 0 Then pCP15 = Val("" & rsBD.Fields("cp15"))
         If Val("" & rsBD.Fields("x02")) > 0 Then pTimes = Val("" & rsBD.Fields("x02"))
         If Val("" & rsBD.Fields("x03")) > 0 Then pHours = Val("" & rsBD.Fields("x03"))
         Pub_GetLAforTimes = True
      End If
   End If
   Set rsBD = Nothing
End Function

'Added by Morgan 2016/7/26 原寫在P的核准、核駁及一般來函
'台灣、大陸及澳門之發明及設計案件,若有同時辦美國案(美國案未領證未閉卷且需為發明案),則於核准、核駁及審查意見通知書定稿中加入一段美國提IDS之提醒字眼
'Modified by Morgan 2016/7/26 不要排除美專-1,-2...之案件--郭雅娟
'Modified by Morgan 2016/9/8 加判斷有發明申請發文--郭雅娟
'Added by Morgan 2018/1/4 +有主張或被主張優先權也要--郭雅娟
'Modified by Morgan 2018/9/7 加判斷有113,114,122,307發文也算--郭雅娟
'Modified by Morgan 2020/12/25 美國領證發文4週後的才不要 --郭雅娟 109/8/21 請作
'Modified by Morgan 2021/2/25 考慮有多個美國案
'Modified by Morgan 2021/3/25 +CIP,CPA,CA,分割案(沒建多國關聯者)--郭雅娟
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function PUB_GetUSCaseNo(pPA01 As String, pPA02 As String, pPA03 As String, pPA04 As String) As String
   Dim stVTB As String, stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stCaseNoList As String

'   '國外案及國內案的其他國外案(相同案)
'   stSQL = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo" & _
'      " from casemap c1,patent where (cm05,cm06,cm07,cm08) in" & _
'      " (select '" & pPA01 & "','" & pPA02 & "','" & pPA03 & "','" & pPA04 & "' from dual" & _
'      " UNION select c2.cm05,c2.cm06,c2.cm07,c2.cm08 from casemap c2" & _
'      " where c2.cm01='" & pPA01 & "' and c2.cm02='" & pPA02 & "' and c2.cm03='" & pPA03 & "' and c2.cm04='" & pPA04 & "')" & _
'      " and pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04" & _
'      " and pa09='101' and pa08='1' and pa57 is null" & _
'      " and exists(select * from caseprogress where  CP01 = PA01 And cp02 = pa02" & _
'      " And cp03 = pa03 And cp04 = pa04 and cp10 in ('101','113','114','122','307') and cp27>0 and cp159=0)" & _
'      " and not exists(select * from caseprogress Where CP01 = PA01 And cp02 = pa02" & _
'      " And cp03 = pa03 And cp04 = pa04 and cp10='601' and cp27>0 and cp27<to_char(sysdate-28,'yyyymmdd') and cp159=0)"
'
'   '+本案
'   stVTB = "select cm01,cm02,cm03,cm04 from casemap c1 where (cm05,cm06,cm07,cm08) in" & _
'      " (select '" & pPA01 & "','" & pPA02 & "','" & pPA03 & "','" & pPA04 & "' from dual" & _
'      " UNION select c2.cm05,c2.cm06,c2.cm07,c2.cm08 from casemap c2" & _
'      " where c2.cm01='" & pPA01 & "' and c2.cm02='" & pPA02 & "' and c2.cm03='" & pPA03 & "' and c2.cm04='" & pPA04 & "')" & _
'      " UNION select '" & pPA01 & "','" & pPA02 & "','" & pPA03 & "','" & pPA04 & "' from dual"
'
'   '被美國案主張優先權
'   stSQL = stSQL & " union" & _
'      " select p2.pa01||'-'||p2.pa02||decode(p2.pa03||p2.pa04,'000','','-'||p2.pa03||'-'||p2.pa04) CNo" & _
'      " from (" & stVTB & "), patent p1, pridate, patent p2" & _
'      " where p1.pa01(+)=cm01 and p1.pa02(+)=cm02 and p1.pa03(+)=cm03 and p1.pa04(+)=cm04" & _
'      " and pd06(+)=p1.pa11 and pd01='CFP'" & _
'      " and p2.pa01(+)=pd01 and p2.pa02(+)=pd02 and p2.pa03(+)=pd03 and p2.pa04(+)=pd04" & _
'      " and p2.pa09='101' and p2.pa08='1' and p2.pa57 is null" & _
'      " and exists(select * from caseprogress where  CP01 = pd01 And cp02 = pd02 And cp03 = pd03" & _
'      " And cp04 = pd04 and cp10 in ('101','113','114','122','307') and cp27>0 and cp159=0)" & _
'      " and not exists(select * from caseprogress" & _
'      " Where CP01 = pd01 And cp02 = pd02 And cp03 = pd03 And cp04 = pd04 and cp10='601'" & _
'      " and cp27>0 and cp27<to_char(sysdate-28,'yyyymmdd') and cp159=0)"
'
'   '主張美國案優先權
'   stSQL = stSQL & " union" & _
'      " select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo" & _
'      " from (" & stVTB & "), pridate, patent p1" & _
'      " where pd01(+)=cm01 and pd02(+)=cm02 and pd03(+)=cm03 and pd04(+)=cm04 and pd07='101'" & _
'      " and pa11(+)=pd06 and pa09='101' and pa08='1' and pa57 is null" & _
'      " and exists(select * from caseprogress where  CP01 = PA01 And cp02 = pa02" & _
'      " And cp03 = pa03 And cp04 = pa04 and cp10 in ('101','113','114','122','307') and cp27>0 and cp159=0)" & _
'      " and not exists(select * from caseprogress Where CP01 = PA01 And cp02 = pa02" & _
'      " And cp03 = pa03 And cp04 = pa04 and cp10='601' and cp27>0 and cp27<to_char(sysdate-28,'yyyymmdd') and cp159=0)"
  
   stVTB = PUB_GetRelCaseVTB2(pPA01, pPA02, pPA03, pPA04, True)
   
   'Modified by Morgan 2021/4/14 要排除本案 Ex:CFP-31173核駁
   'Modified by Morgan 2023/4/28 美國案改領證提申日+4週後才不管制(原為發文日+4週)
   'Modified by Morgan 2023/7/18 改美國案要管制到發證日-1天--郭
   'stSQL = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo" & _
      " from (" & stVTB & ") x,patent" & _
      " where pa01(+)=x01 and pa02(+)=x02 and pa03(+)=x03 and pa04(+)=x04 and pa08='1' and pa09='101' and pa57 is null" & _
      " and exists(select * from caseprogress where  CP01 = PA01 And cp02 = pa02" & _
      " And cp03 = pa03 And cp04 = pa04 and cp10 in ('101','113','114','122','307') and cp27>0 and cp159=0)" & _
      " and not exists(select * from caseprogress Where CP01 = PA01 And cp02 = pa02" & _
      " And cp03 = pa03 And cp04 = pa04 and cp10='601' and cp47>0 and cp47<to_char(sysdate-28,'yyyymmdd') and cp159=0)" & _
      " and not (pa01='" & pPA01 & "' and pa02='" & pPA02 & "' and pa03='" & pPA03 & "' and pa04='" & pPA04 & "')"
   'Modified by Morgan 2023/11/20 改美國案未發文也要管制(拿掉 cp27>0 條件)
   stSQL = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CNo" & _
      " from (" & stVTB & ") x,patent" & _
      " where pa01(+)=x01 and pa02(+)=x02 and pa03(+)=x03 and pa04(+)=x04 and pa08='1' and pa09='101' and pa57 is null" & _
      " and (pa21 is null or pa21>" & strSrvDate(1) & ") and exists(select * from caseprogress where  CP01 = PA01 And cp02 = pa02" & _
      " And cp03 = pa03 And cp04 = pa04 and cp10 in ('101','113','114','122','307') and cp159=0)" & _
      " and not (pa01='" & pPA01 & "' and pa02='" & pPA02 & "' and pa03='" & pPA03 & "' and pa04='" & pPA04 & "')"
   stSQL = stSQL & " order by 1 asc"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      stCaseNoList = .Fields("CNo")
      .MoveNext
      Do While Not .EOF
         stCaseNoList = stCaseNoList & "、" & .Fields("CNo")
         .MoveNext
      Loop
      End With
   End If
   PUB_GetUSCaseNo = stCaseNoList
End Function

'Added by Morgan 2020/12/18
'是否為要更新基本檔准駁的性質
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function PUB_ChkIsRltPty(pCP01 As String, pCP10 As String, Optional pPA09 As String) As Boolean
   
   If pCP01 = "CFP" Then
      'Modified by Morgan 2025/8/7 +120
      If (pCP10 >= "101" And pCP10 <= "105") Or pCP10 = "107" Or pCP10 = "113" Or pCP10 = "114" Or pCP10 = "120" Or pCP10 = "122" Or pCP10 = "126" Or (pCP10 >= "301" And pCP10 <= "307") Or pCP10 = "424" Or pCP10 = "438" Or pCP10 = "501" Or pCP10 = "805" Then
         PUB_ChkIsRltPty = True
      End If
   ElseIf pCP01 = "P" Then
      '大陸復審例外
      If ((pCP10 >= "101" And pCP10 <= "105") Or (pPA09 <> "020" And pCP10 = "107") Or pCP10 = "125" Or (pCP10 >= "301" And pCP10 <= "308") Or pCP10 = "503" Or pCP10 = "504" Or pCP10 = "802" Or pCP10 = "804") Then
         PUB_ChkIsRltPty = True
      End If
   End If
End Function

'Added by Morgan 2023/12/11
'檢查CFP來函性質是否需管制IDS
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function PUB_CheckIDSOA(pPA01 As String, pPA09 As String, pCP10 As String, pRefCP10 As String) As Boolean
   If pPA01 = "CFP" And (pCP10 = "1209" Or pCP10 = "1002" Or pCP10 = "1006" Or pCP10 = "1220" Or pCP10 = "1202" Or pCP10 = "1815" Or pCP10 = "1801" Or pCP10 = "1802") Then
      If pCP10 = "1002" Then
         If PUB_ChkIsRltPty("CFP", pRefCP10, pPA09) = False Then
            If InStr("801,802,803,804", pRefCP10) = 0 Then 'Added by Morgan 2023/12/26 +801異議、802異議答辯 、803舉發、804舉發答辯 --郭
               Exit Function
            End If
         End If
      End If
      PUB_CheckIDSOA = True
   End If
End Function

'Added by Morgan 2020/12/11
'來函輸入管制美國案IDS期限
'pCDate:官方來文日
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Sub PUB_SetUsIDS(pPA01 As String, pPA02 As String, pPA03 As String, pPA04 As String, pCP09 As String, pCDate As String, Optional pPA09 As String, Optional pCP10 As String, Optional pRefCP10 As String, Optional pAddIDS As Boolean = False)
   Dim stCaseNo As String, stUSNo As String, arrNo() As String, pa(4) As String
   Dim stSQL As String, intQ As Integer, rsQuery As ADODB.Recordset
   Dim ii As Integer, stNP08 As String, stNP09 As String, stNP10 As String, stNP15 As String
   'Added by Morgan 2021/2/25
   Dim arrUSNO() As String, jj As Integer
   Dim strSub As String, strContent As String
   Dim stIDSNP09 As String, stIDSNP08 As String, stIDSTo As String, stIDSCC As String 'Added by Morgan 2023/4/28
   
   'P 由來函畫面決定
   If pPA01 = "P" And pAddIDS = True Then
      stUSNo = PUB_GetUSCaseNo(pPA01, pPA02, pPA03, pPA04)
   'CFP 檢索報告1209(PCT除外)、核駁1002、最終核駁1006、建議性處分書 1220
   'Modified by Morgan 2021/2/25 +審查意見通知函1202
   'Modified by Morgan 2022/7/14 +第三方意見1815--郭
   'Modified by Morgan 2022/9/26 PCT檢索報告1209改不排除,被異議 1801,被舉發 1802--郭
   'Modified by Morgan 2023/12/11 CFP判斷是否OA改用函數判斷(因核准也要用)，另增加直接核准時由User決定是否管制IDS的判斷(pAddIDS)
   'ElseIf pPA01 = "CFP" And (pCP10 = "1209" Or pCP10 = "1002" Or pCP10 = "1006" Or pCP10 = "1220" Or pCP10 = "1202" Or pCP10 = "1815" Or pCP10 = "1801" Or pCP10 = "1802") Then
   '   '核駁1002(先用會更新基本檔准駁的性質,有例外再說)
   '   If pCP10 = "1002" Then
   '      If PUB_ChkIsRltPty(pPA01, pRefCP10, pPA09) = False Then
   '         Exit Sub
   '      End If
   '   End If
   ElseIf pPA01 = "CFP" Then
      If pAddIDS = False Then
         If PUB_CheckIDSOA(pPA01, pPA09, pCP10, pRefCP10) = False Then
            Exit Sub
         End If
      End If
   'end 2023/12/12
      'Modified by Morgan 2021/3/25
      'stUSNo = PUB_CFPGetUSCaseNo(pPA01, pPA02, pPA03, pPA04)
      stUSNo = PUB_GetUSCaseNo(pPA01, pPA02, pPA03, pPA04)
   End If
   
   If stUSNo <> "" Then
      'Modified by Morgan 2025/2/19 備註改放官方發文日(原為收文日)--郭
      stSQL = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||' '||na03||'/'||ptm03||'/'||decode(pa09,'000',cpm03,cpm04)||'('||sqldatet(" & DBDATE(pCDate) & ")||')'" & _
         " from caseprogress,patent,nation,patenttrademarkmap,casepropertymap" & _
         " where cp09='" & pCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and na01(+)=pa09" & _
         " and ptm01(+)='1' and ptm02(+)=pa08 and cpm01(+)=cp01 and cpm02(+)=cp10"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         stNP15 = rsQuery(0)
      End If
      
      stCaseNo = pPA01 & "-" & pPA02 & IIf(pPA03 & pPA04 = "000", "", "-" & pPA03 & "-" & pPA04)
      stNP09 = CompDate(1, 3, pCDate) '法限=官方來文日+3個月
      stNP08 = PUB_GetWorkDay1(CompDate(2, -14, stNP09), True) '所限=法限-2週
      If stNP08 < strSrvDate(1) Then stNP08 = strSrvDate(1)
         
      'Modified by Morgan 2021/2/25 考慮有多個美國案
      'arrNo = Split(stUSNo, "-")
      arrUSNO = Split(stUSNo, "、")
      For jj = LBound(arrUSNO) To UBound(arrUSNO)
         arrNo = Split(arrUSNO(jj), "-")
      'end 2021/2/25
         Erase pa() 'Added by Morgan 2021/11/19
         intQ = 1
         For ii = LBound(arrNo) To UBound(arrNo)
            pa(intQ) = arrNo(ii)
            intQ = intQ + 1
         Next
         If pa(3) = "" Then pa(3) = "0"
         If pa(4) = "" Then pa(4) = "00"
         
         'Added by Morgan 2023/4/28
         '檢查美國案是否有4週內已提申的領證,若有則以領證提申日+4週作為最新IDS的法定期限,法限提前1週作為本限,同時發出以下MAIL提醒智權同仁及最新一道程序之工程師:
         'Modified by Morgan 2023/7/18
         '改檢查美國案是否已有發證日,若有則以發證日-1天(抓工作日)作為最新IDS的法定期限,法限提前2個工作天為本限
         'stSQL = "select cp47,pa05,cu04 from caseprogress,patent,customer" & _
            " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
            " and cp10='601' and cp47>=to_char(sysdate-28,'yyyymmdd') and cp159=0" & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
            " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
         stSQL = "select pa05,pa21,cu04 from patent,customer" & _
            " where pa01='" & pa(1) & "' and pa02='" & pa(2) & "' and pa03='" & pa(3) & "' and pa04='" & pa(4) & "'" & _
            " and pa21>" & strSrvDate(1) & _
            " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            'Modified by Morgan 2023/7/18
            'stIDSNP09 = CompDate(2, 28, rsQuery("cp47")) '法限=領證提申日+4週
            'stIDSNP08 = PUB_GetWorkDay1(CompDate(2, -7, stIDSNP09), True) '所限=法限-1週
            stIDSNP09 = PUB_GetWorkDay1(CompDate(2, -1, rsQuery("pa21")), True) '法限=發證日-1天(再抓工作日)
            stIDSNP08 = CompWorkDay(2, CompDate(2, -1, stIDSNP09), 1) '所限=法限提前2個工作天
            'end 2023/7/18
            If stIDSNP08 < strSrvDate(1) Then stIDSNP08 = strSrvDate(1)
            
            strSub = (PUB_DBYEAR(stIDSNP09) - 1911) & "年" & PUB_DBMONTH(stIDSNP09) & "月" & PUB_DBDAY(stIDSNP09) & "日"
            strSub = arrUSNO(jj) & " 美國案已提出領證,惟相關案 " & stCaseNo & " 有新的引證前案發出,若要提IDS則必須盡快於" & strSub & "前提交,請盡速與本案的工程師連繫。"
            strContent = "本所案號：" & arrUSNO(jj)
               strContent = strContent & vbCrLf & "案件名稱：" & rsQuery("pa05")
               strContent = strContent & vbCrLf & "案件性質：IDS"
               strContent = strContent & vbCrLf & "申請人 ：" & rsQuery("cu04")
               strContent = strContent & vbCrLf & "本所期限：" & ChangeWStringToTDateString(stIDSNP08)
               strContent = strContent & vbCrLf & "法定期限：" & ChangeWStringToTDateString(stIDSNP09)
               strContent = strContent & vbCrLf & "他國官方來函：" & stNP15
            
            stIDSTo = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
            
            stSQL = "Select CP14,ST04 FROM CASEPROGRESS,STAFF WHERE " & ChgCaseprogress(pa(1) & pa(2) & pa(3) & pa(4)) & " AND CP14=ST01(+) AND ST03<>'P12' ORDER BY CP05 DESC, CP09 DESC "
            intQ = 1
            Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
            If intQ = 1 Then
               If rsQuery("st04") = "1" Then
                  stIDSCC = rsQuery("cp14")
               Else
                  stIDSCC = "99050"
               End If
            End If
            stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               " values('" & strUserNum & "','" & stIDSTo & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'" & strSub & "','" & strContent & "')"
            cnnConnection.Execute stSQL, intQ
         Else
            stIDSNP09 = stNP09
            stIDSNP08 = stNP08
         End If
         'end 2023/4/28
         
         'Added by Morgan 2021/2/25
         '若有已收文未發文IDS則更新期限及備註並通知承辦人
         stSQL = "select cp09,cp14,cp07,pa05,cu04 from caseprogress,patent,customer" & _
            " where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
            " and cp10='214' and cp27||cp57 is null" & _
            " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
            " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
         If intQ = 1 Then
            stSQL = "update caseprogress set cp64='" & stNP15 & ";'||cp64"
            '更新期限(期限較早者)
            If IsNull(rsQuery("cp07")) Or Val("" & rsQuery("cp07")) > stIDSNP09 Then
               stSQL = stSQL & ",cp06=" & stIDSNP08 & ",cp07=" & stIDSNP09
            End If
            '更新備註
            stSQL = stSQL & " where cp09='" & rsQuery("cp09") & "'"
            cnnConnection.Execute stSQL, intQ
            
            '通知承辦人
            If Not IsNull(rsQuery("cp14")) Then
               strSub = arrUSNO(jj) & " IDS 已有他國官方來函 " & stCaseNo & ",請儘速辦理"
               strContent = "本所案號：" & arrUSNO(jj)
               strContent = strContent & vbCrLf & "案件名稱：" & rsQuery("pa05")
               strContent = strContent & vbCrLf & "案件性質：IDS"
               strContent = strContent & vbCrLf & "申請人 ：" & rsQuery("cu04")
               strContent = strContent & vbCrLf & "本所期限：" & ChangeWStringToTDateString(stIDSNP08)
               strContent = strContent & vbCrLf & "法定期限：" & ChangeWStringToTDateString(stIDSNP09)
               strContent = strContent & vbCrLf & "他國官方來函：" & stNP15
               
               stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                  " values('" & strUserNum & "','" & rsQuery("cp14") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                  ",'" & strSub & "','" & strContent & "')"
               cnnConnection.Execute stSQL, intQ
            End If
         Else
         'end 2021/2/25
         
            stNP10 = PUB_GetAKindSalesNo(pa(1), pa(2), pa(3), pa(4))
            
            stSQL = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP15,NP22)" & _
               " select '" & pCP09 & "','" & pa(1) & "','" & pa(2) & "','" & pa(3) & "','" & pa(4) & "',214," & stIDSNP08 & "," & stIDSNP09 & _
               ",'" & stNP10 & "','" & stNP15 & ";',NP22" & _
               " from (select nvl(max(np22),0)+1 NP22 from nextprogress) n"
            cnnConnection.Execute stSQL, intQ
            
         End If 'Added by Morgan 2021/2/25
      Next 'Added by Morgan 2021/2/25
   End If
      
End Sub

'Added by Lydia 2024/05/30 勘誤公報控管：新增自動掛相關總收文號
Public Function Pub_GetProcCRC(ByVal pType As String, ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, Optional ByVal pCP10 As String) As String
'pType : 1-輸入證書號, 2-分案作業
Dim strR1 As String, intR As Integer
Dim rsRD As New ADODB.Recordset
Dim strCP10 As String
   
   Pub_GetProcCRC = ""
   If pType = "1" Then
      strCP10 = "402"
   Else
      strCP10 = pCP10
   End If
   If strCP10 = "" Then Exit Function
   
   strR1 = "select cp09,cp43, substr(mno,9,9) as mno from caseprogress, " & _
           "(select cp01 as m01,cp02 as m02,cp03 as m03,cp04 as m04,max(cp05||cp09) mno " & _
           "from caseprogress where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp10='1228' and cp159=0 group by cp01,cp02,cp03,cp04) vtb1 " & _
           "where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp10='" & strCP10 & "' " & _
           "and cp159=0 and cp158=0 and cp01=m01(+) and cp02=m02(+) and cp03=m03(+) and cp04=m04(+) "
   intR = 1
   Set rsRD = ClsLawReadRstMsg(intR, strR1)
   If intR = 1 Then
      If "" & rsRD.Fields("mno") <> "" And "" & rsRD.Fields("CP43") = "" Then
         If pType = "1" Then  '1-輸入證書號
            Pub_GetProcCRC = "Update caseprogress Set CP43='" & rsRD.Fields("mno") & "' where cp09='" & rsRD.Fields("cp09") & "' "
         Else  '分案作業
            '若分案「更正402」的時候，判斷此案已有公告號，在進分案作業畫面中，彈提醒: 此案已公告，一併掛相關收文號為公告公報
            If strCP10 = "402" Then
                MsgBox "此案已公告，一併掛相關收文號為公告公報。", vbInformation
                Pub_GetProcCRC = "" & rsRD.Fields("mno")
            '分案「變更401、更改403」的時候，判斷此案已有公告號，在進分案作業畫面中，詢問: 此案已公告，是否一併掛相關收文號為公告公報，可選是跟否 (P.S因為後面有可能收"變更")
            Else
                intR = MsgBox("此案已公告，是否一併掛相關收文號為公告公報？", vbInformation + vbYesNo + vbDefaultButton1)
                If intR = 6 Then 'Yes=是
                   Pub_GetProcCRC = "" & rsRD.Fields("mno")
                End If
            End If
         End If
      End If
   End If
   
   Set rsRD = Nothing
End Function

'add by nick 2004/11/29 傳入 CP01020304 檢查  若非台灣，在新增CP 資料時， Cp44,cp45 存該號最大之 A 或 B 類之 Cp44,Cp45
'Move by Lydia 2024/06/11 從basQuery搬過來
Public Sub Pub_UpdateFromMaxCP27(oCP01 As String, oCP02 As String, oCP03 As String, oCP04 As String)
   Dim tmpSQL As String
   Dim intQ As Integer, stCP09 As String 'Added by Morgan 2024/10/9

   'Added by Morgan 2024/10/9
   'P、CFP案申請、年費、IDS可能有不同代理人,要依相關收文號的性質來控制
   If oCP01 = "CFP" Or oCP01 = "P" Then
      tmpSQL = "select a.cp09,b.cp10,b.cp44,b.cp45 from caseprogress a,patent,caseprogress b where a.cp01='" & oCP01 & "' and a.cp02='" & oCP02 & "' and a.cp03='" & oCP03 & "' and a.cp04='" & oCP04 & "'" & _
         " and a.cp09<'D' AND a.CP66=" & strSrvDate(1) & " AND a.CP65='" & strUserNum & "' and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04 and pa09>'000'" & _
         " and b.cp09(+)=a.cp43 order by a.cp67 desc"
      intQ = 1
      Set AdoRecordSet3 = ClsLawReadRstMsg(intQ, tmpSQL)
      If intQ = 1 Then
         stCP09 = ""
         With AdoRecordSet3
         '年費或IDS直接抓相關收文號的代理人
         If .Fields("cp10") = "214" Or .Fields("cp10") = "605" Or .Fields("cp10") = "606" Or .Fields("cp10") = "607" Then
            tmpSQL = "update caseprogress set cp44='" & .Fields("cp44") & "',cp45='" & .Fields("cp45") & "' where cp09='" & .Fields("cp09") & "'"
            cnnConnection.Execute tmpSQL, intQ
         Else
            stCP09 = .Fields("cp09")
         End If
         End With
         
         '其他則抓年費及IDS以外的最新代理人
         If stCP09 <> "" Then
            'Modified by Morgan 2024/11/11 +排除907不續辦,936回覆委任代理人,957詢問代理人 Ex:CFP-033363
            'Modified by Morgan 2025/3/24 +排除相關收文號的案件性質是214,605,606,607 Ex:CFP-033407 催提申-IDS
            tmpSQL = "select cp44,cp45,cp09,cp10,cp27 from caseprogress b where cp09=(select substr(max(to_char(cp27)||cp09),9) from caseprogress a" & _
               " where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp09>='A' and cp09<'C' and cp27>0 and cp44 is not null and cp10 not in ('214','605','606','607','907','936','957')" & _
               " and not exists(select * from caseprogress c where cp09=a.cp43 and cp10 in ('214','605','606','607')))"
            intQ = 1
            Set AdoRecordSet3 = ClsLawReadRstMsg(intQ, tmpSQL)
            If intQ = 1 Then
               With AdoRecordSet3
               tmpSQL = "update caseprogress set cp44='" & .Fields("cp44") & "',cp45='" & .Fields("cp45") & "' where cp09='" & stCP09 & "'"
               cnnConnection.Execute tmpSQL, intQ
               End With
            End If
         End If
      End If
      CheckOC3
   Else
   'end 2024/10/9
   
      'Modify By Sindy 2012/2/22 CFT改預設最近一次A類收文之代理人及加servicepractice
      'Modify By Sindy 2013/4/12 CFT,CFC,S案件之CP44預設規則統一:抓A,B類收文且發文日最大者,但剔除711文件簽證及304申請英文證明
      Select Case oCP01
         Case "CFC", "CFT", "S"
            'modify by sonia 2017/3/14 CFT案再剔除Y99999999(網路查名查詢）CFT-018361,但S案不剔除
            tmpSQL = " select * from caseprogress where cp09 in ("
            'modify by sonia 2018/6/8 取消and cp57 is null 條件, P-120097新案發文後取消收文
            tmpSQL = tmpSQL & "select substr(max(to_char(cp27)||cp09),9,9) from caseprogress ,trademark  where cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and '000'<tm10  and cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp09<'C' and cp27 is not null and cp44 is not null AND CP10 NOT IN ('711','304') AND CP44<>'Y99999999' "
            tmpSQL = tmpSQL & " union select substr(max(to_char(cp27)||cp09),9,9) from caseprogress ,servicepractice where cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and '000'<sp09  and cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp09<'C' and cp27 is not null and cp44 is not null AND CP10 NOT IN ('711','304') "
            tmpSQL = tmpSQL & ") "
      '2012/2/22 End
         
         Case Else
            tmpSQL = " select * from caseprogress where cp09 in ("
            '2008/6/12 modify by sonia 不檢查cp45
            'tmpSQL = tmpSQL & " select substr(max(to_char(cp27)||cp09),9,9) from caseprogress ,patent        where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and '000'<pa09  and cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "'  and cp09>='A' and cp09<'C' and cp27 is not null and cp57 is null and cp44 is not null and cp45 is not null "
            'tmpSQL = tmpSQL & " union select substr(max(to_char(cp27)||cp09),9,9) from caseprogress ,trademark  where cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and '000'<tm10  and cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "'  and cp09>='A' and cp09<'C' and cp27 is not null and cp57 is null and cp44 is not null and cp45 is not null "
            'tmpSQL = tmpSQL & " union select substr(max(to_char(cp27)||cp09),9,9) from caseprogress ,lawcase     where cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and '000'<lc15  and cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "'  and cp09>='A' and cp09<'C' and cp27 is not null and cp57 is null and cp44 is not null and cp45 is not null "
            'Modify By Sindy 2012/2/22 加servicepractice
            'modify by sonia 2018/6/8 取消and cp57 is null 條件, P-120097新案發文後取消收文,已被取消就都抓不到
            'Modified by Morgan 2024/3/22 +排除CFP的IDS(214)或相關收文號為IDS的進度,因IDS會固定給Y20825000且不變更原案件的代理人--郭
            'Modified by Morgan 2024/10/9 P,CFP規則特別改上面單獨控制
            tmpSQL = tmpSQL & " select substr(max(to_char(cp27)||cp09),9,9) from caseprogress ,patent        where cp01=pa01 and cp02=pa02 and cp03=pa03 and cp04=pa04 and '000'<pa09  and cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp09>='A' and cp09<'C' and cp27 is not null and cp44 is not null "
            tmpSQL = tmpSQL & " union select substr(max(to_char(cp27)||cp09),9,9) from caseprogress ,trademark  where cp01=tm01 and cp02=tm02 and cp03=tm03 and cp04=tm04 and '000'<tm10  and cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp09>='A' and cp09<'C' and cp27 is not null and cp44 is not null "
            tmpSQL = tmpSQL & " union select substr(max(to_char(cp27)||cp09),9,9) from caseprogress ,lawcase     where cp01=lc01 and cp02=lc02 and cp03=lc03 and cp04=lc04 and '000'<lc15  and cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp09>='A' and cp09<'C' and cp27 is not null and cp44 is not null "
            tmpSQL = tmpSQL & " union select substr(max(to_char(cp27)||cp09),9,9) from caseprogress ,servicepractice where cp01=sp01 and cp02=sp02 and cp03=sp03 and cp04=sp04 and '000'<sp09  and cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp09>='A' and cp09<'C' and cp27 is not null and cp44 is not null "
            tmpSQL = tmpSQL & ") "
      End Select
      
      CheckOC3
      With AdoRecordSet3
          .CursorLocation = adUseClient
          .Open tmpSQL, cnnConnection, adOpenStatic, adLockReadOnly
          If .RecordCount <> 0 Then
              'modify by sonia 2017/2/22 加入排除D類收文號CFP-027614
              'modify by sonia 2018/6/8 max(cp09)改為substr(max(to_char(cp67,'0000')||cp09),6),抓資料條件再加AND CP66=" & strSrvDate(1) & " AND CP65='" & strUserNum & "',否則當日有B,C類時會只抓到C類所以要加當日該操作人員最大時間資料
              cnnConnection.Execute " update caseprogress set cp44='" & CheckStr(.Fields("cp44").Value) & "',cp45='" & CheckStr(.Fields("cp45").Value) & "' where " & _
                                    " cp09 in (select substr(max(to_char(cp67,'0000')||cp09),6) from caseprogress where cp01='" & oCP01 & "' and cp02='" & oCP02 & "' and cp03='" & oCP03 & "' and cp04='" & oCP04 & "' and cp09<'D' AND CP66=" & strSrvDate(1) & " AND CP65='" & strUserNum & "')  "
          End If
      End With
      CheckOC3
   End If
End Sub

'Added by Lydia 2024/06/14 對申請人1~5的重複輸入檢查
Public Function Pub_ChkAppList(ByRef pErrNum As String, ByVal pAppListNo As String) As Boolean
Dim tmpArr1 As Variant
Dim intP As Integer, intN As Integer
Dim strChkList As String

   Pub_ChkAppList = False
   If pAppListNo <> "" Then
      pErrNum = ""
      tmpArr1 = Split(pAppListNo, ",")
      For intP = 0 To UBound(tmpArr1)
         If Trim(tmpArr1(intP)) = "" Then
            If intN = 0 Then
              intN = intP + 1
            End If
         Else
            If InStr(strChkList & ",", ChangeCustomerL(Trim(tmpArr1(intP)))) = 0 Then
               If intN > 0 Then
                  MsgBox "請從申請人" & intN & "開始輸入!", vbCritical + vbOKOnly, MsgText(9001)
                  pErrNum = intP + 1
                  Exit Function
               Else
                  strChkList = strChkList & ChangeCustomerL(Trim(tmpArr1(intP))) & ","
               End If
            Else
               MsgBox "申請人不可重複!", vbCritical + vbOKOnly, MsgText(9001)
               pErrNum = intP + 1
               Exit Function
            End If
         End If
      Next intP
   End If
   Pub_ChkAppList = True
End Function

'Added by Lydia 2024/06/24 內商-【查名單(TradeMarkQurey)】、【查名單-網中(TMQAPPForm)】：檢查圖形路徑是否存在、重覆
Public Function Pub_ChkTMR3isExist(ByVal pRNo As String, ByVal bolMsg As Boolean, Optional ByRef pNewNo As String, Optional ByRef pCName As String) As Boolean
Dim intR As Integer, intX As Integer, strR1 As String
Dim rsRD As New ADODB.Recordset
Dim tmpNo As String, strErrNo As String, strChkNo As String, strDualNo As String
Dim tmpArr1 As Variant, tmpList As String

   Pub_ChkTMR3isExist = False
   pCName = ""
   If Trim(pRNo) = "" Then Exit Function
   
   tmpArr1 = Split(pRNo, ",")
   For intX = 0 To UBound(tmpArr1)
      If Trim(tmpArr1(intX)) <> "" Then
         tmpNo = Trim(tmpArr1(intX))
         If InStr(tmpNo, "-") = 0 Then
            If Len(tmpNo) <> 5 Then
               strErrNo = strErrNo & "," & tmpNo
               tmpNo = ""
            End If
         Else
            If Len(tmpNo) <> 7 Or Mid(tmpNo, 3, 1) <> "-" And Mid(tmpNo, 5, 1) <> "-" Then
               strErrNo = strErrNo & "," & tmpNo
               tmpNo = ""
            End If
         End If
         If tmpNo <> "" Then
            tmpNo = Replace(tmpNo, "-", "")
            strR1 = "SELECT * FROM TMQAPPR3 WHERE TMR301='" & Mid(tmpNo, 1, 2) & "' AND TMR302='" & Mid(tmpNo, 3, 1) & "' AND TMR303='" & Mid(tmpNo, 4) & "' "
            intR = 1
            Set rsRD = ClsLawReadRstMsg(intR, strR1)
            If intR = 1 Then
               tmpNo = Mid(tmpNo, 1, 2) & "-" & Mid(tmpNo, 3, 1) & "-" & Mid(tmpNo, 4)
               If InStr(tmpList & ",", "," & tmpArr1(intX) & ",") > 0 Then
                  strDualNo = strDualNo & "," & tmpArr1(intX)
               Else
                  strChkNo = strChkNo & "," & tmpNo
                  pCName = pCName & "," & rsRD.Fields("TMR304")
                  tmpList = tmpList & "," & Trim(tmpArr1(intX))
               End If
            Else
               strErrNo = strErrNo & "," & Trim(tmpArr1(intX))
            End If
         End If
      End If
   Next intX
   
   If strErrNo & strDualNo <> "" Then
      If strErrNo <> "" Then
         strErrNo = Mid(strErrNo, 2)
         If bolMsg = True Then
            MsgBox "請輸入正確並且存在的代號，例如：01-A-00 或 01A00" & vbCrLf & vbCrLf & "錯誤代號：" & strErrNo, vbCritical, "圖形路徑輸入檢查"
         End If
      End If
      If strDualNo <> "" Then
         strDualNo = Mid(strDualNo, 2)
         If bolMsg = True Then
            MsgBox "重覆輸入，請查明再輸!" & vbCrLf & vbCrLf & "錯誤代號：" & strDualNo, vbCritical, "圖形路徑輸入檢查"
         End If
      End If
   Else
      pNewNo = Mid(strChkNo, 2)
      pCName = Mid(pCName, 2)
      Pub_ChkTMR3isExist = True
   End If
     
   Set rsRD = Nothing
End Function

'Added by Lydia 2024/06/24 內商-【查名單(TradeMarkQurey)】、【查名單-網中(TMQAPPForm)】：檢查委查類別/組群是否存在、重覆
Public Function Pub_ChkTMQCisExist(ByVal pFrmName As String, ByRef pGrpTxt As String, ByVal pType As String, ByVal pStatus As String, Optional ByRef pCName As String, Optional ByVal pOrgList As String, Optional ByVal bolMsg As Boolean = True) As Boolean
'pType: 1-類別, 2-組群
'pStatus: W-文字/文字+圖形, P-圖形
'pOrgList: 已輸入的編號
Dim StrArray As Variant
Dim intA As Integer
Dim strGrp As String, strNameList As String, tmpList As String
Dim strQ1 As String, intQ As Integer, rsQD As New ADODB.Recordset

   StrArray = ""
   pCName = ""
   Pub_ChkTMQCisExist = False
   If Len(pGrpTxt) <> 0 Or Len(pOrgList) <> 0 Then
      If Len(pOrgList) <> 0 And pGrpTxt <> pOrgList Then  '已輸入的編號+(檢查)新增的編號
         StrArray = Split(pOrgList & "," & pGrpTxt, ",")
      Else
         StrArray = Split(pGrpTxt, ",")
      End If
      strGrp = "-"
      For intA = 0 To UBound(StrArray)
         If pType = "1" Then  '類別
            If StrArray(intA) <> "" And (Len(StrArray(intA)) <> 2 Or IsNumeric(StrArray(intA)) = False) Then
               If bolMsg = True Then MsgBox "委查類別格式輸入錯誤!!!", vbCritical, "輸入檢查"
               Exit Function
            End If
            strQ1 = "select tmqc01,nvl(tmqc06,'(空白)') tmqc06 from tmqclass where length(tmqc01)=2 and tmqc01=" & CNULL(Mid("" & StrArray(intA), 1, 2))
            intQ = 1
            Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
            If intQ = 0 Then
                If bolMsg = True Then MsgBox "類別 " & Mid("" & StrArray(intA), 1, 2) & " 查無資料!", vbCritical, "輸入檢查"
                Exit Function
            Else
                strNameList = strNameList & "," & rsQD.Fields("tmqc06")
            End If
         '-----組群
         Else
            If StrArray(intA) <> "" Or IsNumeric(StrArray(intA)) = False Then
               If IsNumeric(StrArray(intA)) = False Then
                  If bolMsg = True Then MsgBox "委查組群請輸入數字!!!", vbCritical, "輸入檢查"
                  Exit Function
               End If
               If Mid(StrArray(intA), 1, 4) <> "3519" And Len(StrArray(intA)) <> 4 Then
                  If bolMsg = True Then MsgBox "非3519組群請輸入4碼!!!", vbCritical, "輸入檢查"
                  Exit Function
               End If
               If Mid(StrArray(intA), 1, 4) = "3519" And Len(StrArray(intA)) <> 6 Then
                  If bolMsg = True Then MsgBox "3519組群請輸入6碼!!!", vbCritical, "輸入檢查"
                  Exit Function
               End If
               '檢查不可存在於組群刪除資料檔
               strQ1 = "Select * From ClassDelete Where CD01='" & StrArray(intA) & "' "
               intQ = 1
               Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
               If intQ = 1 Then
                  If bolMsg = True Then MsgBox StrArray(intA) & "為已刪除的組群，輸入錯誤!!!", vbExclamation + vbOKOnly, "輸入檢查"
                  Exit Function
               End If
            End If

            If pStatus = "W" And Mid(StrArray(intA), 1, 4) <> "3519" Then '文字只檢查類別(2碼)
                strQ1 = "select tmqc01,nvl(tmqc06,'(空白)') tmqc06 from tmqclass where length(tmqc01)=2 and tmqc01=" & CNULL(Mid("" & StrArray(intA), 1, 2))
            Else
                strQ1 = "select tmqc01,nvl(tmqc06,'(空白)') tmqc06 from tmqclass where tmqc01=" & CNULL("" & StrArray(intA))
            End If
            intQ = 1
            Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
            If intQ = 0 Then
                If bolMsg = True Then MsgBox "組群 " & StrArray(intA) & " 查無資料!", vbCritical, "輸入檢查"
                Exit Function
            Else
                strNameList = strNameList & "," & rsQD.Fields("tmqc06")
            End If
            If strGrp = "-" Then
               strGrp = Mid(StrArray(intA), 1, 2)
            End If
            '(原查名單frm090126)文字限同類; 查名單-網中>>可以跨類
            If pStatus = "W" And pFrmName = "frm090126" And strGrp <> Mid(StrArray(intA), 1, 2) Then
               If bolMsg = True Then MsgBox "文字檢索的委查組群必須同一類，請查明再輸!", vbCritical, "輸入檢查"
               Exit Function
            End If
         End If '-----組群
         If InStr(tmpList & ",", "," & StrArray(intA) & ",") > 0 Then
            If bolMsg = True Then MsgBox "委查" & IIf(pType = "1", "類別", "組群") & "重覆輸入：" & StrArray(intA) & "，請查明再輸!", vbCritical, "輸入檢查"
            Exit Function
         Else
            tmpList = tmpList & "," & StrArray(intA)
         End If
      Next intA
   End If
   
   If strNameList <> "" Then
      If Len(pOrgList) <> 0 And pGrpTxt <> pOrgList Then
         pCName = Mid(strNameList, InStrRev(strNameList, ",") + 1)
      Else
         pCName = Mid(strNameList, 2)
      End If
      Pub_ChkTMQCisExist = True
   End If
   Set rsQD = Nothing
   
End Function

'Added by Morgan 2024/11/20
'檢查一案兩請是否發明案已公告且有放棄新型(pa60='Y' or 新型案有429放棄專利權發文)
Public Function PUB_ChkDualCase(pa() As String) As Boolean
   Dim stSQL As String, intR As Integer
   Dim rsQuery As ADODB.Recordset
   
   stSQL = "select cm01,cm02,cm03,cm04" & _
      " from (select cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08 from casemap where cm05='" & pa(1) & "' and cm06='" & pa(2) & "' and cm07='" & pa(3) & "' and cm08='" & pa(4) & "' and cm10='3'" & _
      " union select cm05,cm06,cm07,cm08,cm01,cm02,cm03,cm04 from casemap where cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "' and cm10='3'" & _
      ") X,patent where pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04 and pa08='1' and pa14>0" & _
      " and (pa60='Y' or exists(select * from caseprogress where cp01=cm05 and cp02=cm06 and cp03=cm07 and cp04=cm08 and cp10='429' and cp27>0))"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      PUB_ChkDualCase = True
   End If
   
   Set rsQuery = Nothing
End Function

'檢查是否有處理狀態
'Modify By Sindy 2025/1/8 + Optional ByVal ChkType As Integer = 0
'ChkType=0: 檢查是否有處理狀態
'        1: 檢查是否要清空處理狀態
Public Function PUB_CheckIRStatus(pIR01 As String, pIR02 As String, pIR03 As String, pIR04 As String, _
   Optional ByRef pIR16 As String, Optional ByVal ChkType As Integer = 0) As Boolean
   
   Screen.MousePointer = vbHourglass 'Add By Sindy 2025/6/19
   'Add By Sindy 2025/1/8
   If ChkType = 1 Then
      strExc(0) = "select ir16,decode(ir16," & 信件處理狀態 & ",ir16) as ir16Nm" & _
                  " from inputRecord,staff where ir01=" & pIR01 & " and ir03='" & pIR03 & "'" & _
                  " and ir04=st01(+) and st04='1'" & _
                  " and ((ir16||ir24 is not null and st03='" & PUB_GetST03(pIR04) & "') or (ir15='Y' and ir04='" & pIR04 & "'))"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 0 Then
         PUB_CheckIRStatus = True
      Else
         PUB_CheckIRStatus = False
      End If
   Else
   '2025/1/8 END
      'Modify By Sindy 2017/12/26
      'strExc(0) = "select ir16 from inputRecord where ir01=" & pIR01 & " and ir02=" & pIR02 & " and ir03='" & pIR03 & "' and ir04='" & pIR04 & "'"
      'Modified by Morgan 2023/4/12 +8.退回2
      strExc(0) = "select ir16,decode(ir16," & 信件處理狀態 & ",ir16) as ir16Nm" & _
                  " from inputRecord where ir01=" & pIR01 & " and ir03='" & pIR03 & "' and ir04='" & pIR04 & "'" & _
                  " and ir16 not in('4','3','8') and ir16 is not null" '歸卷/退回不算沖銷
      intI = 1
      pIR16 = ""
      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
      If intI = 1 Then
         If Trim("" & RsTemp("ir16")) <> "" Then '改判斷有無處理狀態, 原本判斷=1.輸入
            pIR16 = RsTemp("ir16Nm")
            PUB_CheckIRStatus = True
         Else
            PUB_CheckIRStatus = False
         End If
      'Add By Sindy 2025/5/15
      Else
         PUB_CheckIRStatus = False
         '他人轉寄沖銷掉了
         strExc(0) = "select ir08,ir10" & _
                     " from inputRecord where ir01=" & pIR01 & " and ir03='" & pIR03 & "' and ir04='" & pIR04 & "'" & _
                     " and ir16 is null and ir08>0"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
         If intI = 1 Then
            pIR16 = GetPrjSalesNM("" & RsTemp.Fields("ir10")) & "已轉寄"
            PUB_CheckIRStatus = True
         End If
      '2025/5/15 END
      End If
   End If
   Screen.MousePointer = vbDefault 'Add By Sindy 2025/6/19
End Function

'Added by Morgan 2025/3/13 台灣案自請撤回不可分案/發文管控
'1.要撤回的相關收文號為申請程序(101,102,103,125)且距申請日／最早優先權日超過15個月。
'2.要撤回的相關收文號案件性質為(101,102,103,125,107,803,421)且已有結果。
'3.要撤回的相關收文號案件性質為(301,302,307)。
'pCP09:自請撤回相關總收文號,pMail:EMail承辦人及智權
Public Function PUB_ChkTW413(pCP09 As String, Optional pMail As Boolean = False) As Boolean
   Dim stSQL As String, intQ As Integer, stMsg As String, stTemp As String
   Dim rsQuery As ADODB.Recordset
   Dim bRtn As Boolean
   
   bRtn = True
   
   stSQL = "select min(pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04)) CaseNo" & _
      ",min(a.cp24) Rlt,nvl(min(pd05),min(pa10)) DDate,min(a.cp10) Pty,min(cpm03) PtyC,min(b.cp14) Eng,min(b.cp13) Sal,min(b.cp09) RecNo" & _
      " from caseprogress a,caseprogress b,patent,pridate,casepropertymap" & _
      " where a.cp09='" & pCP09 & "' and b.cp43(+)=a.cp09 and b.cp10(+)='413' and b.cp158(+)=0 and b.cp159(+)=0" & _
      " and pa01(+)=a.cp01 and pa02(+)=a.cp02 and pa03(+)=a.cp03 and pa04(+)=a.cp04 and pa09='000'" & _
      " and pd01(+)=pa01 and pd02(+)=pa02 and pd03(+)=pa03 and pd04(+)=pa04" & _
      " and cpm01(+)=a.cp01 and cpm02(+)=a.cp10"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      
      '審定來函EMail通知要檢查有未發文的自請撤回
      If pMail Then
         If IsNull(.Fields("RecNo")) Then
            Exit Function
         ElseIf Not IsNull(.Fields("Rlt")) Then
            stMsg = .Fields("CaseNo") & "已收到審定，自請撤回已不可辦理！"
         
            stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               " values( '" & strUserNum & "','" & .Fields("Eng") & ";" & .Fields("Sal") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'" & stMsg & "','如旨')"
            cnnConnection.Execute stSQL, intQ
            
            bRtn = False
         End If
      Else
         '1
         If InStr("301,302,307,421", .Fields("Pty")) > 0 Then
            stMsg = .Fields("PtyC") & "(" & pCP09 & ")不可辦理自請撤回！"
         '2
         ElseIf Not IsNull(.Fields("Rlt")) Then
            stMsg = .Fields("PtyC") & "(" & pCP09 & ")已有結果，自請撤回已不可辦理！"
         '1
         'Removed by Morgan 2025/4/21 取消--韻丞
         'ElseIf InStr("101,102,103,125", .Fields("Pty")) > 0 And .Fields("DDate") > 0 Then
         '   stTemp = CompDate(1, 15, .Fields("DDate"))
         '   If Val(strSrvDate(1)) > Val(stTemp) Then
         '      stMsg = "已距申請日／最早優先權日超過15個月，" & .Fields("PtyC") & "(" & pCP09 & ")不可辦理自請撤回！"
         '   End If
         End If
         If stMsg <> "" Then
            MsgBox stMsg, vbCritical, .Fields("CaseNo") & "自請撤回相關收文號檢查"
            bRtn = False
         End If
      End If
      End With
   End If
   PUB_ChkTW413 = bRtn
End Function

'Added by Lydia 2025/03/25 查名單-網中：取得查名單最大流水號+1碼
Public Function Pub_GetAutoTMA01() As String
Dim strB1 As String, intB As Integer
Dim rsBD As New ADODB.Recordset
   
   Pub_GetAutoTMA01 = ""
   '編碼：開頭H+民國年3碼+流水號5碼
   strB1 = "SELECT 'H'||(SUBSTR(TO_CHAR(SYSDATE,'YYYYMMDD'),1,4)-1911)||LPAD(NVL(SUBSTR(MAX(TMA01),5,5),0)+1,5,'0') AS ANO " & _
           "From TMQAPPFORM WHERE TMA01 LIKE 'H'||(SUBSTR(TO_CHAR(SYSDATE,'YYYYMMDD'),1,4)-1911)||'%' "
   intB = 1
   Set rsBD = ClsLawReadRstMsg(intB, strB1)
   If intB = 1 Then
      Pub_GetAutoTMA01 = "" & rsBD.Fields("ANO")
      '保留流水號
      strB1 = "Insert Into TMQAPPFORM (TMA01,TMA03) values ('" & Pub_GetAutoTMA01 & "','AAAAAA') "
      cnnConnection.Execute strB1
   End If
   
End Function

'Added by Morgan 2025/9/25
'P案工程師(非FMP)
Public Function PUB_GetPPromoter(pCaseNo As String) As String
   Dim stSQL As String, intQ As Integer, ii As Integer
   Dim rsQuery As ADODB.Recordset
   Dim arrTxt() As String
   Dim pa(4) As String
   
   If InStr(pCaseNo, "-") > 0 Then
      arrTxt = Split(pCaseNo, "-")
      intQ = 1
      For ii = LBound(arrTxt) To UBound(arrTxt)
         pa(intQ) = arrTxt(ii)
         intQ = intQ + 1
      Next
      If pa(3) = "" Then pa(3) = "0"
      If pa(4) = "" Then pa(4) = "00"
   Else
      ChgCaseNo pCaseNo, pa
   End If
   
   '先抓本案號最後的承辦工程師
   stSQL = "select cp14,st04 from caseprogress,staff where cp01='" & pa(1) & "' and cp02='" & pa(2) & "' and cp03='" & pa(3) & "' and cp04='" & pa(4) & "'" & _
      " and cp159=0 and st01(+)=cp14 and st03 in ('P10','P11') and substr(st01,1,1)<>'F' order by nvl(cp27,cp05) desc"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      If rsQuery("st04") = "1" Then
         PUB_GetPPromoter = rsQuery("cp14")
      Else
         PUB_GetPPromoter = "99050"
      End If
   '沒有再抓國內案最後的承辦工程師
   Else
      stSQL = "select cp14,st04 from caseprogress,staff where (cp01,cp02,cp03,cp04) in (select cm05,cm06,cm07,cm08" & _
         " from casemap where cm01='" & pa(1) & "' and cm02='" & pa(2) & "' and cm03='" & pa(3) & "' and cm04='" & pa(4) & "' and cm10='0')" & _
         " and cp159=0 and st01(+)=cp14 and st03 in ('P10','P11') and substr(st01,1,1)<>'F' order by nvl(cp27,cp05) desc"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         If rsQuery("st04") = "1" Then
            PUB_GetPPromoter = rsQuery("cp14")
         Else
            PUB_GetPPromoter = "99050"
         End If
      Else
         PUB_GetPPromoter = "99050"
      End If
   End If
   Set rsQuery = Nothing
End Function

