Attribute VB_Name = "bas088"
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo by Morgan2010/12/28 申請案號欄已修改
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit


'指定國家年費
'Modified by Morgan 2020/8/7 +strFAgentNo
Public Sub ModifyMoneyCountry(ByRef strCountry As String, strMoneyCountry As String, ByRef strMoney As String, Optional ByRef strFagentNo As String)
frm880009.strCountry = strCountry
frm880009.strMoneyCountry = strMoneyCountry
frm880009.strMoney = strMoney
frm880009.strFagentNo = strFagentNo
frm880009.Show vbModal
strCountry = frm880009.strCountry
strMoneyCountry = frm880009.strMoneyCountry
strMoney = frm880009.strMoney
strFagentNo = frm880009.strFagentNo
End Sub

'Add by Morgan 2009/3/19
'同案件同時段多筆發文設定,參數皆為字串不同文以逗號區隔
'strCP09:正在發文的收文號,strCP09s:要更新的收文號,strCP123s:是否經發文室-主管機關,strCP27 發文日,bolDefer 是否為B類延期
'modify by sonia 2014/6/23 +strCP84s 發文規費
Public Function ModifyDispatch(ByVal strCP09 As String, ByRef strCP09s As String, ByRef strCP123s As String, ByRef strCP84s As String, Optional strCP27 As String, Optional bolBDefer As Boolean) As Boolean
   Dim lngCursor As Long
   If pub_strUserOffice <> "1" Then
      strCP09s = ""
      ModifyDispatch = True
      Exit Function
   End If
   frm880015.strCP09 = strCP09
   frm880015.strCP27 = strCP27
   frm880015.strCP84 = strCP84s      'add by sonia 2014/6/23
   frm880015.bolBDefer = bolBDefer
   If frm880015.CheckShowList Then
      lngCursor = Screen.MousePointer
      Screen.MousePointer = vbDefault
      frm880015.Show vbModal
      Screen.MousePointer = lngCursor
   End If
   strCP09s = frm880015.strCP09s
   strCP123s = frm880015.strCP123s
   ModifyDispatch = frm880015.bolOK
   
   Unload frm880015
   Set frm880015 = Nothing
End Function

'Add by Sindy 2009/4/24
'發文項目選擇,參數皆為字串不同主管機關以逗號區隔
'strCP09:正在發文的收文號
'strCP09s:要更新的收文號
'strCP123s:是否經發文室-主管機關
'strCP130s:要更新的主管機關
'strCP27:發文日
'bolIsDefer:是否延展發文 Added by Morgan 2011/11/3
'bolIsEApp:是否電子送件 Added by Morgan 2016/4/29
'bolIsCaseNum:是否算發文室件數 Add by Sindy 2018/8/3
Public Function ModifyDispatchCp130(ByVal strCP09 As String, ByRef strCP09s As String, _
   ByRef strCP123s As String, ByRef strCP130s As String, Optional strCP27 As String, _
   Optional bolIsDefer As Boolean, Optional bolIsEApp As Boolean, Optional ByRef bolIsCaseNum As Boolean) As Boolean
   
   'Removed by Morgan 2020/4/6 分所不再送件,但會執行發文作業仍由北所送件
   'If pub_strUserOffice <> "1" Then
   '   strCP09s = ""
   '   ModifyDispatchCp130 = True
   '   Exit Function
   'End If
   'end 2020/4/6
   
   frm880016.strCP09 = strCP09
   frm880016.strCP27 = strCP27
   frm880016.bolIsDefer = bolIsDefer
   frm880016.bolIsEApp = bolIsEApp 'Added by Morgan 2016/4/29
   frm880016.bolIsCaseNum = bolIsCaseNum 'Add by Sindy 2018/8/3
   If frm880016.CheckShowList Then
      frm880016.Show vbModal
   End If
   strCP09s = frm880016.strCP09s
   strCP123s = frm880016.strCP123s
   strCP130s = frm880016.strCP130s
   bolIsCaseNum = frm880016.bolIsCaseNum 'Add by Sindy 2018/8/3
   ModifyDispatchCp130 = frm880016.bolOK
   
   Unload frm880016
   Set frm880016 = Nothing
End Function


'Move from basQuery by Morgan 2007/5/22

''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 刪除資料記錄
' Input : nType == 刪除檔案的種類
'                  0 : 表刪除的是基本檔, 此時 strKey1 放的是本所案號
'                  1 : 表刪除的是案件進度檔, 此時 strKey1 放的是總收文號, 此時 strKey2 放的是本所案號
'                  2 : 表刪除的是下一程序檔, 此時 strKey1 放的是下一程序檔的序號, 此時 strKey2 放的是本所案號
'         strKey1 == 輸入的Key
'                  當刪除的是基本檔時時, strKey1 放的是本所案號
'                  當刪除的是案件進度檔時, strKey1 放的是總收文號
'                  當刪除的是下一程序檔時, strKey1 放的是下一程序檔的序號
'         strKey2 == 當刪除的案件進度檔或下一程序檔時, strKey2 放的是本所案號
'         strDD23 == 失誤人員的預設值(可不給)
'         strDD24 == 刪除備註的預設致(可不給)
'         bShowForm == 是否要顯示畫面讓使用者輸入失誤人員及刪除備註
' Output : 傳回處理的結果
'         0 : 表處理成功
'         -1 : 表輸入的參數不完整或不正確
'         -2 : 原始檔案的該筆記錄不存在
'         -3 : 使用者按下取消鍵
'         此外, 程式會將失誤人員放入strDD23, 刪除備註放入strDD24中傳回去
''''''''''''''''''''''''''''''''''''''''''''''''''''''
Public Function OnDataDeleteRecord(ByVal nType As Integer, ByVal strKey1 As String, Optional ByVal StrKey2 = Empty, Optional ByRef strDD23 As String = Empty, Optional ByRef strDD24 As String = Empty, Optional ByVal bShowForm As Boolean = True) As Integer
   Dim strTM As String
   Dim strTM01 As String
   Dim strTM02 As String
   Dim strTM03 As String
   Dim strTM04 As String
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   Dim nPos As Integer
   Dim strTemp As String
   OnDataDeleteRecord = 0

   strTM = Empty
   strTM01 = Empty
   strTM02 = Empty
   strTM03 = Empty
   strTM04 = Empty
   
   ' 檢查刪除的檔案種類是否正確
   If nType < 0 Or nType > 2 Then
      OnDataDeleteRecord = -1
      GoTo EXITSUB
   End If
   ' 檢查是否有輸入 Key
   If IsEmptyText(strKey1) = True Then
      OnDataDeleteRecord = -1
      GoTo EXITSUB
   End If

   ' 此段程式碼在檢查所輸入的參數是否正確
   Select Case nType
      ' 刪除基本檔
      Case 0:
         strTM = strKey1
      ' 刪除案件進度檔
      Case 1:
         ' 檢查總收文號的值是否正確
         Select Case Mid(strKey1, 1, 1)
            'Modified by Lydia 2016/12/26 + D類收文
            Case "A", "B", "C", "D":
            Case Else:
               OnDataDeleteRecord = -1
               GoTo EXITSUB
         End Select
         ' 本所案號
         strTM = StrKey2
      ' 刪除下一程序檔
      Case 2:
         ' 檢查下一程序的序號值是否正確
         If IsNumeric(strKey1) = False Then
            OnDataDeleteRecord = -1
            GoTo EXITSUB
         End If
         ' 本所案號
         strTM = StrKey2
   End Select
   
   ' 90.07.23 modify by louis (由於錯誤的本所案號會導致無法刪除, 故改變檢查的方式)
   For nPos = 1 To Len(strTM)
      If IsNumeric(Mid(strTM, nPos, 1)) = True Then
         Exit For
      End If
   Next nPos
   If nPos > 1 Then
      strTM01 = Mid(strTM, 1, nPos - 1)
      strTemp = Right(strTM, Len(strTM) - (nPos - 1))
      strTM04 = Right(strTemp, 2)
      strTemp = Left(strTemp, Len(strTemp) - 2)
      strTM03 = Right(strTemp, 1)
      strTemp = Left(strTemp, Len(strTemp) - 1)
      strTM02 = strTemp
   Else
      OnDataDeleteRecord = -1
      GoTo EXITSUB
   End If
   
   Select Case nType
      Case 1:
         ' 設定
         frm880010.SetData 4, strKey1, True
         ' 設定本所案號
         If IsEmptyText(strTM01) = False Then
            frm880010.SetData 0, strTM01, False
            frm880010.SetData 1, strTM02, False
            frm880010.SetData 2, strTM03, False
            frm880010.SetData 3, strTM04, False
         End If
      Case 2:
         frm880010.SetData 5, strKey1, True
         ' 設定本所案號
         If IsEmptyText(strTM01) = False Then
            frm880010.SetData 0, strTM01, False
            frm880010.SetData 1, strTM02, False
            frm880010.SetData 2, strTM03, False
            frm880010.SetData 3, strTM04, False
         End If
      Case Else:
         frm880010.SetData 0, strTM01, True
         frm880010.SetData 1, strTM02, False
         frm880010.SetData 2, strTM03, False
         frm880010.SetData 3, strTM04, False
   End Select
   
   If IsEmptyText(strDD23) = False Then
      frm880010.SetData 6, strDD23, False
   End If
   If IsEmptyText(strDD24) = False Then
      frm880010.SetData 7, strDD24, False
   End If

   ' 讀取檔案失敗
   If frm880010.QueryData() = False Then
      OnDataDeleteRecord = -2
   End If
   
   If bShowForm = True Then
      ' 顯示畫面讓使用者輸入
      frm880010.Show vbModal
      ' 檢查使用者是按下OK還是Cancel
      If frm880010.IsOK = False Then
         OnDataDeleteRecord = -3
      Else
         OnDataDeleteRecord = 0
      End If
   Else
      frm880010.OnSaveData
      OnDataDeleteRecord = 0
   End If
   strDD23 = frm880010.GetData(0)
   strDD24 = frm880010.GetData(1)
   
   Unload frm880010
EXITSUB:
End Function

'Remove by Morgan 2009/7/24 (搬到專利系統的frm050102_1)
''intComeFrom=0為要秀出EMail    =1時直接進入發文
'Public Function Where020102ToGo(Optional intComeFrom As Integer = 0)
'   If intComeFrom = 0 Then
'      '911118 nick 12:22  邱小姐說改成只檢查 cp79 有值就秀 mail，cp16 不檢查
'      '***** start
'      'If frm020102_1.grdDataList.TextMatrix(frm020102_1.grdDataList.Row, 6) = "" Or frm020102_1.grdDataList.TextMatrix(frm020102_1.grdDataList.Row, 6) = "0" Then
'      If frm020102_1.grdDataList.TextMatrix(frm020102_1.grdDataList.row, 6) <> "" And frm020102_1.grdDataList.TextMatrix(frm020102_1.grdDataList.row, 6) <> "0" Then
'         '911029 nick 邱小姐說 增加檢查 cp16 <>0 因為有費用才要收款
'         'Dim nick911029rs As New ADODB.Recordset
'         'Dim nickstrsql As String
'         'nickstrsql = "select cp16 from caseprogress where cp09='" & frm020102_1.grdDataList.TextMatrix(frm020102_1.grdDataList.Row, 5) & "' "
'         'Set nick911029rs = New ADODB.Recordset
'         'nick911029rs.CursorLocation = adUseClient
'         'nick911029rs.Open nickstrsql, cnnConnection, adOpenStatic, adLockReadOnly
'         'If nick911029rs.RecordCount <> 0 Then
'         '   If Val(CheckStr(nick911029rs.Fields(0).Value)) > 0 Then
'      'Else
'                frm020102_K.Show
'                Exit Function
'          '  End If
'         'End If
'      '***** end
'      End If
'   End If
'   Select Case intPCaseKind
'      Case 專利
'         Select Case intPWhere
'            Case 國內
'
'            Case 國外_CF
'               If frm020102_1.grdDataList.TextMatrix(frm020102_1.grdDataList.row, 9) = 專利 Then
'                  Select Case frm020102_1.grdDataList.TextMatrix(frm020102_1.grdDataList.row, 7)
'                     Case 延期
'                        'Add By Cheng 2002/06/20
'                        '延期記錄資料來源為下一程序檔
'                        frm050102_2.m_str_DL05 = "2"
'
'                        frm050102_2.intWhereComeFrom = 2
'                        frm050102_2.Show
'                     'Add by Morgan 2006/8/14 加122CA申請
'                     Case 發明申請, 新型申請, 設計申請, 聯合申請, CIP申請, CPA申請, 再發行, 美國暫時申請, 分割, "122"
'                        frm050102_3.Show
'                     Case 變更
'                        frm050102_4.Show
'                     'Modify by Morgan 2007/7/27 加"繼承"
'                     Case 實體審查, 答辯, 修正, 主動修正, 提供前案資料, 選取, 讓與, "214", "427", 繼承
'                        frm050102_5.Show
'                     Case 補文件
'                        frm050102_6.Show
'                     Case 申請優先權證明
'                        frm050102_7.Show
'                     Case 領證及繳年費
'                        frm050102_8.Show
'                     Case 年費, 維持費, 延展費
'                        frm050102_9.Show
'                     Case 授權
'                        frm050102_a.Show
'                     Case Else
'                        frm050102_6.Show
'                  End Select
'               Else
'                  'Add By Cheng 2002/07/30
'                  '93.11.10 MODIFY by sonia 美國暫時申請 改收CFP
'                  'Select Case frm020102_1.grdDataList.TextMatrix(frm020102_1.grdDataList.Row, 7)
'                  'Case 美國暫時申請
'                  '   frm050102_3.Show
'                  'Case Else
'                  '   frm050102_6.Show
'                  'End Select
'                  frm050102_6.Show
'                  '93.11.10 END
'               End If
'            Case 國外_FC
'
'         End Select
'      Case 商標
'         Select Case intPWhere
''            Case 國內
''               If frm020102_1.grdDataList.TextMatrix(frm020102_1.grdDataList.Row, 9) = 商標 Then
''                  Select Case frm020102_1.grdDataList.TextMatrix(frm020102_1.grdDataList.Row, 7)
''                     Case "01"
''                        frm020102_2.Show
''                     Case "02"
''                        frm020102_3.Show
''                     Case "03"
''                        frm020102_4.Show
''                     Case 申請
''
''                     Case 移轉
''
''                     Case 異議
''                     Case 評定
''                     Case 廢止
''                  End Select
''               Else
''
''               End If
''            Case 國外_CF
'
'   End Select
'End Select
'End Function
'指定國家領證
'Modified by Morgan 2020/8/14 +strType
Public Sub ModifyLicenceCountry(ByRef strCountry As String, strLicenceCountry As String, Optional ByVal strPA10 As String, Optional strCP10 As String)
Dim strTemp As String
frm880008.strCountry = strCountry
frm880008.strLicenceCountry = strLicenceCountry
frm880008.strPA10 = strPA10
'Added by Morgan 2020/8/14
'Modified by Morgan 2023/3/7 改傳入案件性質
'If strType = "1" Then
'   frm880008.Caption = "指定國註冊費"
'ElseIf strType = "2" Then
'   frm880008.Caption = "年費"
'End If
frm880008.strCP10 = strCP10
If strCP10 <> "" Then
   Call ClsPDGetCaseProperty("CFP", strCP10, strTemp)
   frm880008.Caption = strTemp
End If
'end 2020/8/14
frm880008.Show vbModal
strCountry = frm880008.strCountry
strLicenceCountry = frm880008.strLicenceCountry

End Sub

'變更
'Modify by Morgan 2009/7/24 +strCP09
Public Sub ModifyChange(ByRef strChange As String, ByRef bolIsChange As Boolean, ByRef intFunction As Integer, ByRef strCP09 As String)
frm880007.strCP09 = strCP09
frm880007.strChange = strChange
frm880007.bolIsChange = bolIsChange
frm880007.intFunction = intFunction
frm880007.Show vbModal
strChange = frm880007.strChange
bolIsChange = frm880007.bolIsChange

End Sub
'補件期限
Public Sub ModifyAddDeadline(ByRef strAddDeadline1 As String, ByRef strAddDeadline2 As String, ByRef strAddDeadline3 As String)
frm880003.strAddDeadline1 = strAddDeadline1
frm880003.strAddDeadline2 = strAddDeadline2
frm880003.strAddDeadline3 = strAddDeadline3
frm880003.Show vbModal
strAddDeadline1 = frm880003.strAddDeadline1
strAddDeadline2 = frm880003.strAddDeadline2
strAddDeadline3 = frm880003.strAddDeadline3

End Sub
'Add by Morgan 2009/11/11
'補件期限(非台灣案)
Public Sub ModifyAddDeadline1(ByRef strCP09 As String, ByVal strCP06 As String, Optional ByVal strCP07 As String, Optional ByRef siSaveFlag As Single, Optional ByVal bolNoAdd As Boolean = False, Optional ByRef strUnSaveData As String, Optional ByRef bolAddFix As Boolean = False, Optional ByRef strUnSaveData2 As String)
   With frm880017
   .m_CP43 = strCP09
   .m_CP06 = strCP06
   .m_CP07 = strCP07
   .m_siSaveFlag = siSaveFlag
   .m_stUnSaveData = strUnSaveData
   .m_bolNoAdd = bolNoAdd
   .m_bolAddFix = bolAddFix 'Added by Morgan 2014/3/5
   .m_stUnSaveData2 = strUnSaveData2 'Added by Morgan 2014/3/5
   frm880017.Show vbModal
   siSaveFlag = .m_siSaveFlag
   strUnSaveData = .m_stUnSaveData
   strUnSaveData2 = .m_stUnSaveData2 'Added by Morgan 2014/3/5
   End With
   Set frm880017 = Nothing
End Sub

'發明人
'Modified by Morgan 2020/2/19 +strCallForm
Public Sub ModifyInventor(ByRef strPetition As String, ByRef strInventorNo As String, Optional pCallForm As Form)

Set frm880006.fmCallForm = pCallForm 'Added by Morgan 2020/2/19
frm880006.strPetition = strPetition
frm880006.strInventorNo = strInventorNo
frm880006.Show vbModal
strInventorNo = frm880006.strInventorNo
Set frm880006 = Nothing 'Added by Morgan 2020/2/19

End Sub

'Remove by Morgan 2009/7/24 不再使用
''設定回發文之第一個畫面時應做何處理 是繼續發文或清空
'Public Sub SetPubishTextGo(ByRef strSystem As String, ByRef strCode1 As String, ByRef strCode2 As String, ByRef strCode3 As String, ByRef bolGoOn As Boolean)
'If bolGoOn = False Then
'   frm020102_1.optChoose(0).Value = True
'   frm020102_1.txtSystem = ""
'   frm020102_1.txtCode(0) = ""
'   frm020102_1.txtCode(1) = ""
'   frm020102_1.txtCode(2) = ""
'   frm020102_1.txtTFCode(0) = ""
'   frm020102_1.txtTFCode(1) = ""
'   frm020102_1.txtTFCode(2) = ""
'   frm020102_1.txtTFCode(3) = ""
'   frm020102_1.txtReceiveCode = ""
'ElseIf frm020102_1.optChoose(0).Value Then
'   'TF為馬德里案，另外判斷
'   frm020102_1.optChoose(1).Value = True
'   frm020102_1.txtSystem = strSystem
'   If strSystem = 馬德里案 Then
'      frm020102_1.txtTFCode(0) = Left(strCode1, 5)
'      frm020102_1.txtTFCode(1) = IIf(Right(strCode1, 1) = "0", "", Right(strCode1, 1))
'      frm020102_1.txtTFCode(2) = IIf(strCode2 = "0", "", strCode2)
'      frm020102_1.txtTFCode(3) = IIf(strCode3 = "00", "", strCode3)
'   Else
'      frm020102_1.txtCode(0) = strCode1
'      frm020102_1.txtCode(1) = IIf(strCode2 = "0", "", strCode2)
'      frm020102_1.txtCode(2) = IIf(strCode3 = "00", "", strCode3)
'   End If
'End If
'End Sub

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 當符合條件時, 顯示商標基本資料維護的畫面或專利基本資料維護的畫面
'
' Input : strCP09 ==> 收文號
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Modify By Sindy 2010/10/15 增加
'NotRun : 不檢查
'strFrmName : 傳遞執行的作業
'Modify By Sindy 2024/5/3 +, Optional ByRef strRefText As String = "": 回傳文字
Public Sub ShowMaintainForm(ByVal strCP09 As String, Optional ByRef NotRun As String, _
   Optional ByRef strFrmName As String, Optional ByRef m_form As Form, _
   Optional ByRef strRefText As String = "")
   
   Dim strSql As String
   Dim rsTmp As ADODB.Recordset
   Dim bShow As Boolean
   Dim strCP01 As String
   Dim strCP02 As String
   Dim strCP03 As String
   Dim strCP04 As String
   'add by nickc 2006/12/05
   Dim frm As Form
   Dim frmIsUse As Boolean
   'Add By Cheng 2002/01/11
   Dim strCP10 As String '案件性質
   Dim ii As Integer '回圈序號
   
   Set rsTmp = New ADODB.Recordset
   strSql = "SELECT CP01,CP02,CP03,CP04,CP10 FROM CASEPROGRESS " & _
            "WHERE CP09 = '" & strCP09 & "' "
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsTmp.RecordCount > 0 Then
      bShow = False
      ' 錯誤
      If IsNull(rsTmp.Fields("CP01")) Then
         rsTmp.Close
         Set rsTmp = Nothing
         Exit Sub
      End If
      strCP01 = rsTmp.Fields("CP01")
      If Not IsNull(rsTmp.Fields("CP02")) Then
         strCP02 = rsTmp.Fields("CP02")
      End If
      If Not IsNull(rsTmp.Fields("CP03")) Then
         strCP03 = rsTmp.Fields("CP03")
      End If
      If Not IsNull(rsTmp.Fields("CP04")) Then
         strCP04 = rsTmp.Fields("CP04")
      End If
      
      'Add By Cheng 2002/01/11
      strCP10 = "" & rsTmp.Fields("CP10")
      
      Select Case strCP01
         ' 讀取商標基本檔
         Case "T", "TF", "CFT", "FCT":
            If Not IsNull(strCP10) Then
               If strCP10 = "101" Then
                  rsTmp.Close
                  Set rsTmp = Nothing
                  Exit Sub
               End If
               '2011/9/7 ADD BY SONIA 跨類107,主張優先權108 若A類申請程序尚未發文也不必顯示基本檔(外商需求,內商一併做)
               '2016/3/22 modify by sonia +714超項費,711文件公簽證
'modify by sonia 2016/8/24 不限制案件性質,只要有A類申請案未發文都不檢查
'               If strCP10 = "107" Or strCP10 = "108" Or strCP10 = "714" Or strCP10 = "711" Then    '2016/8/24 CANCEL BY SONIA
               'modify by sonia 2017/6/8 +分割308
                  strSql = "SELECT CP27 FROM CASEPROGRESS WHERE " & _
                                 "CP01 = '" & strCP01 & "' AND CP02 = '" & strCP02 & "' AND " & _
                                 "CP03 = '" & strCP03 & "' AND CP04 = '" & strCP04 & "' AND " & _
                                 "(CP10='101' OR CP10='308') AND CP09<'B'"
                  rsTmp.Close
                  rsTmp.CursorLocation = adUseClient
                  rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsTmp.RecordCount > 0 Then
                     If "" & rsTmp.Fields(0) = "" Then
                        rsTmp.Close
                        Set rsTmp = Nothing
                        Exit Sub
                     'ADD BY SONIA 2016/3/29 申請101已發文再判斷是不是系統日
                     ElseIf rsTmp.Fields(0) = strSrvDate(1) Then
                        rsTmp.Close
                        Set rsTmp = Nothing
                        Exit Sub
                     'END 2016/3/29
                     End If
                  End If
'               End If   '2016/8/24 CANCEL BY SONIA
               '2011/9/7 END
               'Modify by Amy 2024/05/15 +TM11
               strSql = "SELECT TM28,TM12,TM22,TM11 FROM TRADEMARK " & _
                        "WHERE TM01 = '" & strCP01 & "' AND " & _
                              "TM02 = '" & strCP02 & "' AND " & _
                              "TM03 = '" & strCP03 & "' AND " & _
                              "TM04 = '" & strCP04 & "' "
               rsTmp.Close
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount > 0 Then
                  If NotRun <> "Y" Then 'Add By Sindy 2010/10/15 增加NotRun判斷是否要執行
                     If Not IsNull(rsTmp.Fields("TM28")) Then
   '                     ' 卷宗性質不是申請
                        ' 卷宗性質是申請
                        If rsTmp.Fields("TM28") = "1" Then
                           'Modify by Amy 2024/05/15 電子收文上線後申請案號會有值,故改抓申請日
'                           ' 申請案號是空的
'                           If IsNull(rsTmp.Fields("TM12")) Then
'                              bShow = True
'                           ElseIf IsEmptyText(rsTmp.Fields("TM12")) Then
'                              bShow = True
'                           End If
                           ' 申請日是空的
                           If IsNull(rsTmp.Fields("TM11")) Then
                              bShow = True
                           End If
                           'end 2024/05/15
                        End If
                     End If
                  End If
                  'Add By Sindy 2010/10/15
                  If bShow = False And strFrmName = "分案" Then
                     If strCP10 = "102" Then
                        rsTmp.Close
                        Set rsTmp = Nothing
                        Exit Sub
                     ElseIf Not IsNull(rsTmp.Fields("TM22")) Then
                        '2011/5/31 MODIFY BY SONIA FCT-18127的移轉案
                        'If Val(rsTmp.Fields("TM22")) <> 0 And (Val(rsTmp.Fields("TM22")) <= strSrvDate(1)) Then
                        If Val(rsTmp.Fields("TM22")) <> 0 And (Val(rsTmp.Fields("TM22")) < strSrvDate(1)) Then
                           strRefText = "此案專用期已過，請確認" 'Add By Sindy 2024/5/3
                           MsgBox strRefText & "！" '"此案專用期已過，請確認！"
                           bShow = True
                        '2011/5/31 Add BY SONIA FCT-18127的移轉案
                        ElseIf Val(rsTmp.Fields("TM22")) <> 0 And (Val(rsTmp.Fields("TM22")) = strSrvDate(1)) Then
                           strRefText = "此案專用期今日到期，請確認" 'Add By Sindy 2024/5/3
                           MsgBox strRefText & "！"
                           bShow = True
                        '2011/5/31 end
                        End If
                     End If
                  End If
                  '2010/10/15 End
               End If
               rsTmp.Close
               If bShow Then
                  ' 設定滑鼠游標為等待狀態
                  'add by nickc 2006/12/05 檢查維護畫面若已經開啟，將警告
'                  frmIsUse = False
                  For Each frm In Forms
                      If frm.Name = "frm020501" Then
                          'MsgBox "基本資料維護畫面已經開啟，將不為您代資料，請另行補資料！" & vbCrLf & vbCrLf & "注意：目前基本資料維護畫面將不會是此筆資料！", , "嚴重錯誤！"
'                          frmIsUse = True
                          
                          'Modify by Morgan 2010/8/5 一律關閉原視窗,否則若為修改狀態又選擇取消時會存錯資料
                          'frm020501.Form_KeyDown vbKeyF10, 0
                          'frm020501.Form_KeyDown vbKeyF4, 0
                          If frm020501.m_EditMode = 1 Or frm020501.m_EditMode = 2 Then
                              frm020501.SetFocus
                              MsgBox "基本資料維護畫面即將更新，未完成之作業將要取消，請另行操作！"
                          End If
                          Unload frm020501
                          'end 2010/8/5
                          Exit For
                      End If
                  Next
'                  If frmIsUse = False Then
                  
                    Screen.MousePointer = vbHourglass
                    If strCP01 = "T" Or strCP01 = "TF" Then
                       frm020501.SetSystem 0
                    Else
                       frm020501.SetSystem 1
                    End If
                    ' 顯示基本資料維護的畫面
                    'Modify By Sindy 2012/6/1 +m_form
                    frm020501.SetCurrKey strCP01, strCP02, strCP03, strCP04, m_form
                    frm020501.ShowCurrRecord strCP01, strCP02, strCP03, strCP04 'Add By Sindy 2016/11/28
                    frm020501.Show
                    ' 設定滑鼠游標為預設
                    Screen.MousePointer = vbDefault
'                  End If
               End If
            End If
         ' 讀取專利基本檔
         Case "P", "CFP", "FCP":
            If Not IsNull(strCP10) Then
               'Modified by Morgan 2012/10/8 +125
               If (CInt(strCP10) <= 105) Or (CInt(strCP10) >= 108 And CInt(strCP10) <= 115 And CInt(strCP10) <> 111) _
               Or (CInt(strCP10) = 116) Or (CInt(strCP10) = 125) Or (CInt(strCP10) = 201) _
               Or (CInt(strCP10) = 203) Or (CInt(strCP10) = 209) _
               Or (CInt(strCP10) = 210) Or (CInt(strCP10) = 901) Then
                  rsTmp.Close
                  Set rsTmp = Nothing
                  Exit Sub
               End If
               strSql = "SELECT PA23, PA11 FROM PATENT " & _
                        "WHERE PA01 = '" & strCP01 & "' AND " & _
                              "PA02 = '" & strCP02 & "' AND " & _
                              "PA03 = '" & strCP03 & "' AND " & _
                              "PA04 = '" & strCP04 & "' "
               rsTmp.Close
               rsTmp.CursorLocation = adUseClient
               rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
               If rsTmp.RecordCount > 0 Then
                  'If Not IsNull(rsTmp.Fields("PA23")) Then
                     ' 卷宗性質是申請
                     If rsTmp.Fields("PA23") = "1" Then
                        ' 申請案號是空的
                        If IsNull(rsTmp.Fields("PA11")) Then
                           bShow = True
                        ElseIf IsEmptyText(rsTmp.Fields("PA11")) Then
                           bShow = True
                        End If
                     End If
                  'End If
               End If
               rsTmp.Close
               
               'Add By Cheng 2002/01/11
               '內專分案, 若案件性質為"專利權讓與"時, 檢查專利基本檔的申請日, 申請案號, 公告日, 公告號, 發證日,
               '專利號數, 專用期間, 繳年費記錄, 若有缺資料之欄位, 顯示專利基本維護畫面
               If bShow = False And strCP01 = "P" And strCP10 = 專利權讓與 Then
                  strSql = "SELECT PA10,PA11,PA14,PA15,PA21,PA22,PA24,PA25,PA72 FROM PATENT " & _
                           "WHERE PA01 = '" & "" & strCP01 & "' AND " & _
                                 "PA02 = '" & "" & strCP02 & "' AND " & _
                                 "PA03 = '" & "" & strCP03 & "' AND " & _
                                 "PA04 = '" & "" & strCP04 & "'"
                  If rsTmp.State <> adStateClosed Then rsTmp.Close
                  rsTmp.CursorLocation = adUseClient
                  rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
                  If rsTmp.RecordCount > 0 Then
                     For ii = 0 To rsTmp.Fields.Count - 1
                        If IsNull(rsTmp.Fields(ii).Value) Then
                           bShow = True
                           Exit For
                        Else
                           If rsTmp.Fields(ii).Value = "" Then
                              bShow = True
                              Exit For
                           End If
                        End If
                     Next ii
                  End If
                  If rsTmp.State <> adStateClosed Then rsTmp.Close
                  Set rsTmp = Nothing
               End If
               
               If bShow Then
                  'add by nickc 2006/12/05 檢查維護畫面若已經開啟，將警告
'                  frmIsUse = False
                  For Each frm In Forms
                      If frm.Name = "frm050701" Then
                          'MsgBox "基本資料維護畫面已經開啟，將不為您代資料，請另行補資料！" & vbCrLf & vbCrLf & "注意：目前基本資料維護畫面將不會是此筆資料！", , "嚴重錯誤！"
'                          frmIsUse = True

                          'Modify by Morgan 2010/8/5 一律關閉原視窗,否則若為修改狀態又選擇取消時會存錯資料
                          'frm050701.Form_KeyDown vbKeyF10, 0
                          'frm050701.Form_KeyDown vbKeyF4, 0
                          If frm050701.ActionEdit = 0 Or frm050701.ActionEdit = 1 Then
                              frm050701.SetFocus
                              MsgBox "基本資料維護畫面即將更新，未完成之作業將要取消，請另行操作！"
                          End If
                          Unload frm050701
                          'end 2010/8/5
                          Exit For
                      End If
                  Next
'                  If frmIsUse = False Then
                        ' 設定滑鼠游標為等待狀態
                        Screen.MousePointer = vbHourglass
                        ' 顯示基本資料維護的畫面
                        strSysKind = strCP01
                          Load frm050701
                          frm050701.Show: DoEvents
                          'Modified by Morgan 2012/6/18 +m_form
                          frm050701.SetCurrKey strCP01, strCP02, strCP03, strCP04, m_form: DoEvents
                          
                          frm050701.Text1(1).Text = strCP01
                          frm050701.Text1(2).Text = strCP02
                          frm050701.Text1(3).Text = Left(strCP03 & "0", 1)
                          frm050701.Text1(4).Text = Left(strCP04 & "00", 2)
                          '按下確定
                          frm050701.Form_KeyDown vbKeyF9, 0
                          '按下修改
                          frm050701.Form_KeyDown vbKeyF3, 0
                        ' 設定滑鼠游標為預設
                        Screen.MousePointer = vbDefault
'                 End If
               End If
            End If
         Case Else:
            rsTmp.Close
      End Select
   End If
   Set rsTmp = Nothing
End Sub

'Modify by Morgan 2004/3/16
'主管機關來函畫面進入畫面與接洽紀錄分開
Public Sub Where01ToGo(ByVal intLeaveKind As Integer, Optional stFormName As String = "frm010001")

    Dim oTmp As Form
    
    If stFormName = "frm010001_1" Then
        Set oTmp = frm010001_1
    Else
        Set oTmp = frm010001
    End If

    If intLeaveKind = 1 Then
       Select Case oTmp.intModifyKind
                 Case 0
                            'Modifie by Lydia 2023/05/16 櫃台(陳金妙)反應有時新增收文存檔後，無法回到前一畫面，即使再按選單呼叫也不會出現，必須重登系統；
                                '在檢查”視窗”有看到”接洽單－新增”的清單
                            'If oTmp.intReceiveKind = 0 Then
                            If TypeName(oTmp) = "frm010001" Then '一律開啟
                               oTmp.Show
                            Else
                               'edit by nickc 2007/02/07 不用 dll 了
                               'Set obj001 = Nothing
                            End If
                 Case 1, 2
                            oTmp.Show
       End Select
    Else
       'edit by nickc 2007/02/07 不用 dll 了
       'Set obj001 = Nothing
       Unload oTmp
    End If
    
End Sub

'再確認輸入
'Modify by Amy 2015/01/28 +strCaption 表單Caption顯示文字/ strLabel 顯示文字內容
'Modified by Morgan 2021/12/4 oText 改為 Object 以相容 Form2.0
Public Function CheckReKey(ByRef txtTemp As Object, Optional strCaption As String = "", Optional strLabel As String = "") As Boolean

If txtTemp.Tag = "" Then
   CheckReKey = True
ElseIf txtTemp <> txtTemp.Tag Then
   bolTheSame = False 'Added by Lydia 2017/10/17
     
   Set frm880004.txtTemp = txtTemp
    'Add by Amy 2015/01/28
   If Trim(strCaption) <> "" Then frm880004.Caption = strCaption & frm880004.Caption
   If Trim(strLabel) <> "" Then frm880004.Label1.Caption = strLabel
   frm880004.Show vbModal
   'Modified by Lydia 2017/10/17 改成共用變數
   'If frm880004.bolTheSame Then CheckReKey = True
   If bolTheSame Then CheckReKey = True
Else
   CheckReKey = True
End If
End Function

Public Sub ModifyAssignCountry(ByRef strCountry As String, Optional ByVal strPA10 As String)
frm880001.strCountry = strCountry
frm880001.strPA10 = strPA10
frm880001.Show vbModal
strCountry = frm880001.strCountry
End Sub

'Move from basStart by Morgan 2004/3/12
Public Sub Where1103ComeFrom(Optional frmTemp As Form, Optional strCode1 As String, Optional strCode2 As String, Optional strCode3 As String, Optional strCode4 As String)
    Static frmLastTemp As Form
    
    If frmTemp Is Nothing Then
       frmLastTemp.Show
    Else
       frm1103_2.intWhereComeFrom = 2
       frm1103_2.lblSystem = strCode1
       If strCode1 = 馬德里案 Then
          'edit by nickc 2006/04/20 原先的 bug 吧
          'frm1103_2.lblTFCode(0) = Left(strCode3, 5)
          'frm1103_2.lblTFCode(1) = IIf(Right(strCode3, 1) = "0", "", Right(strCode3, 1))
          frm1103_2.lblTFCode(0) = Left(strCode2, 5)
          frm1103_2.lblTFCode(1) = IIf(Right(strCode2, 1) = "0", "", Right(strCode2, 1))
          frm1103_2.lblTFCode(2) = IIf(strCode3 = "0", "", strCode3)
          frm1103_2.lblTFCode(3) = IIf(strCode4 = "00", "", strCode4)
       Else
          frm1103_2.lblCode(0) = strCode2
          frm1103_2.lblCode(1) = IIf(strCode3 = "0", "", strCode3)
          frm1103_2.lblCode(2) = IIf(strCode4 = "00", "", strCode4)
       End If
       frm1103_2.Show
       Set frmLastTemp = frmTemp
       frmTemp.Hide
    End If
End Sub

'優先權
'Modify by Morgan 2005/11/4 加p_bolDblCheck=重新輸入檢查,p_strCaseNo=本所案號,p_bolAppCheck=申請人是否一致檢查
'Modify by Morgan 2007/4/24 加strPriority4
'Modify by Amy 2014/03/21 + strPriority5
'Modify by Sindy 2017/9/29 + strPriority6
'Modify by Amy 2023/01/05 +m_PrevForm
Public Sub ModifyPriority(ByRef strPriority1 As String, ByRef strPriority2 As String, _
   ByRef strPriority3 As String, Optional ByVal p_stPA08 As String = "", _
   Optional ByRef p_bolDblCheck As Boolean = False, Optional ByVal p_strCaseNo As String, _
   Optional ByVal p_stPA09 As String = "", Optional ByRef p_bolAppCheck As Boolean = False, _
   Optional ByRef strPriority4 As String, Optional ByRef strPriority5 As String, _
   Optional ByRef strPriority6 As String, Optional ByVal m_PrevForm As Form = Nothing)

'Add By Cheng 2002/06/05
frm880002.m_blnAddNew = False
'Add by Morgan 2004/6/7
frm880002.m_stPA08 = p_stPA08
'Add by Morgan 2005/11/4
frm880002.m_stPA09 = p_stPA09
frm880002.m_bolDblCheck = p_bolDblCheck
frm880002.m_strCaseNo = p_strCaseNo
frm880002.m_bolAppCheck = p_bolAppCheck
'Add By Sindy 2017/10/17 商品類別非必要傳入欄位值,如專利系統
Dim varPriorityTemp2 As Variant, i As Integer
If strPriority2 <> "" And strPriority6 = "" Then
   varPriorityTemp2 = Split(strPriority2, "，")
   For i = 1 To UBound(varPriorityTemp2)
      strPriority6 = strPriority6 & "，"
   Next i
End If
'2017/10/17 END
frm880002.strPriority1 = strPriority1
frm880002.strPriority2 = strPriority2
frm880002.strPriority3 = strPriority3
frm880002.strPriority4 = strPriority4
frm880002.strPD09 = strPriority5
frm880002.strPriority6 = strPriority6 'Add By Sindy 2017/9/29
'Modify by Amy 2023/01/05 +stFormN
If TypeName(m_PrevForm) = "Nothing" Then
    frm880002.Show vbModal
    strPriority1 = frm880002.strPriority1
    strPriority2 = frm880002.strPriority2
    strPriority3 = frm880002.strPriority3
    strPriority4 = frm880002.strPriority4
    strPriority5 = frm880002.strPD09
    strPriority6 = frm880002.strPriority6 'Add By Sindy 2017/9/29
    p_bolDblCheck = frm880002.m_bolDblCheck
    'Add by Amy 2023/01/05 bug發現資料沒清(先開有優先權資料,關掉再開無優先權資料)
    frm880002.m_bolDblCheck = False
    frm880002.strPriority1 = ""
    frm880002.strPriority2 = ""
    frm880002.strPriority3 = ""
    frm880002.strPriority4 = ""
    frm880002.strPD09 = ""
    frm880002.strPriority6 = ""
    'Add by Amy 2014/04/08 解決先做專利基本檔修改優先權資料
    '再改商標基本檔優先權資料不應該彈專利種類必輸
    Set frm880002 = Nothing
Else
    Set frm880002.frmParent = m_PrevForm
    frm880002.Show
End If

End Sub

'Add by Morgan 2007/7/3
Public Sub PUB_BatchPrint(p_LD05 As String)
   
   'Dim iDefaultPrinter As Integer 'Remove by Morgan 2010/2/3
   
   Screen.MousePointer = vbDefault
   
   Load frm880011
   'Modify by Morgan 2010/2/3
   'iDefaultPrinter = frm880011.GetPrinterIndex
   pub_OsPrinter = PUB_GetOsDefaultPrinter
   'end 2010/2/3
   frm880011.Show 1
   
   Screen.MousePointer = vbHourglass
   PrinterLetterDemand p_LD05
   
   '還原控制台&Word預設印表機
   'Modify by Morgan 2010/2/3
   'If iDefaultPrinter <> -1 Then
   '   If Printers(iDefaultPrinter).DeviceName <> Printer.DeviceName Then
   '      Printer.TrackDefault = True
   '      CreateObject("WScript.Network").SetDefaultPrinter Printers(iDefaultPrinter).DeviceName
   '      PUB_SetWordActivePrinter
   '   End If
   'End If
   If pub_OsPrinter <> "" Then
      CreateObject("WScript.Network").SetDefaultPrinter pub_OsPrinter
   End If
   'end 2010/2/3
End Sub

'Modify By Sindy 2014/5/27 Mark統一使用basQuery中的函數
''*************************************************
''  電腦自動給號(傳票號碼)
''
''*************************************************
'Public Function AccAutoNo(InputItem As String, InputLength As Integer) As String
'Dim adoaccnum As New ADODB.Recordset
'Dim strItem As String, strYes As String
'   If Len(InputItem) > 1 Then
'      strItem = Mid(InputItem, 2, 1)
'   Else
'      strItem = InputItem
'   End If
'   adoaccnum.CursorLocation = adUseClient
'   adoaccnum.Open "select * from acc1r0 where a1r01 = '" & InputItem & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If InputItem <> "X" Then
'      If adoaccnum.RecordCount = 0 Then
'         AccAutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & Mid(ACDate(ServerDate), 4, 2) & ZeroBeforeNo("0", InputLength)
'      Else
'         If adoaccnum.Fields("a1r03").Value <> Val(Mid(ACDate(ServerDate), 4, 2)) Then
'            AccAutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & Mid(ACDate(ServerDate), 4, 2) & ZeroBeforeNo("0", InputLength)
'         Else
'            AccAutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & Mid(ACDate(ServerDate), 4, 2) & ZeroBeforeNo(str(adoaccnum.Fields("a1r04").Value), InputLength)
'         End If
'      End If
'   Else
'      If adoaccnum.RecordCount = 0 Then
'         AccAutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & ZeroBeforeNo("0", InputLength)
'      Else
'         If adoaccnum.Fields("a1r02").Value <> Val(Mid(ServerDate, 1, 4)) Then
'            AccAutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & ZeroBeforeNo("0", InputLength)
'         Else
'            AccAutoNo = strItem & Mid(ACDate(ServerDate), 1, 3) & ZeroBeforeNo(str(adoaccnum.Fields("a1r04").Value), InputLength)
'         End If
'      End If
'   End If
'   adoaccnum.Close
'End Function
''*************************************************
''  電腦給號存檔(傳票號碼)
''
''*************************************************
'Public Function AccSaveAutoNo(InputItem As String, InputNo As String) As String
'Dim adoaccnum As New ADODB.Recordset
'
'   adoaccnum.CursorLocation = adUseClient
'   adoaccnum.Open "select * from acc1r0 where a1r01 = '" & InputItem & "'", adoTaie, adOpenStatic, adLockReadOnly
'   If adoaccnum.RecordCount = 0 Then
'      adoTaie.Execute "insert into acc1r0 (a1r01, a1r02, a1r03, a1r04) values ('" & InputItem & "', '" & Mid(ServerDate, 1, 4) & "', '" & Mid(ServerDate, 5, 2) & "', '" & InputNo & "')"
'   Else
'      adoTaie.Execute "UPDATE ACC1R0 SET A1R01 = '" & InputItem & "', A1R02 = '" & Mid(ServerDate, 1, 4) & "', A1R03 = '" & Mid(ServerDate, 5, 2) & "', A1R04 = '" & InputNo & "' WHERE A1R01 = '" & InputItem & "'"
'   End If
'   AccSaveAutoNo = MsgText(602)
'   adoaccnum.Close
'End Function
''*************************************************
''  刪除資料檢核
''
''*************************************************
'Public Function DeleteCheck(strSql As String) As String
'Dim adoDeleteCheck As New ADODB.Recordset
'
'   adoDeleteCheck.CursorLocation = adUseClient
'   adoDeleteCheck.Open strSql, adoTaie, adOpenStatic, adLockReadOnly
'   If adoDeleteCheck.RecordCount = 0 Then
'      MsgBox MsgText(28), , MsgText(5)
'      DeleteCheck = MsgText(603)
'   Else
'      DeleteCheck = MsgText(602)
'   End If
'   adoDeleteCheck.Close
'End Function


'Modify By Cheng 2002/12/05
'多加傳入參數聯絡單備註, 申請國家, 申請人
'Public Function PrintEmail(ByRef intCaseKind As Integer, ByRef intWhere As Integer, ByRef strReceiveCode As String, ByRef strUserName As String) As Boolean
Public Function ClsPPPrintEmail(ByRef intCaseKind As Integer, ByRef intWhere As Integer, ByRef strReceiveCode As String, ByRef strUserName As String, _
    Optional ByRef strNote As String = "", Optional ByRef strNation As String = "", Optional ByRef strCustomer As String = "") As Boolean
 On Error GoTo ErrHand

   DataEnvPublic.Commands(1).CommandText = ClsPPGetSQL(intCaseKind, intWhere, strReceiveCode)
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
    PUB_SetOsPrtAsApp 'Add by Morgan 2010/2/23
   datrptPublic.PrintReport
   PUB_RestoreOsPrt 'Add by Morgan 2010/2/23
   Unload datrptPublic
   DataEnvPublic.rscmdContact.Close
   ClsPPPrintEmail = True
   Exit Function
ErrHand:
      'edit by nickc 2007/02/02
   'ErrorLog
   MsgBox Err.Description
End Function

'Remove by Morgan 2011/10/3 因已達檔案上限要移除不再使用的
''Public Function PrtFCPMail(ByVal CP09 As String, ByVal strUserName As String) As Boolean
'Public Function ClsPPPrtFCPMail(ByVal CP09 As String, ByVal strUserName As String, Optional ByVal bShow As Boolean = True) As Boolean
' Dim strTmp As String
''Add By Cheng 2002/12/05
' On Error GoTo ErrHand
'
'   DataEnvPublic.Commands(2).CommandText = "select st03 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
'      "pa07 a06,cu05||cu88 a07,FA05||FA63 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
'      "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN WHERE cp09='" & CP09 & "' AND " & _
'      "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
'      "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
'      "PA11=TPB01(+)"
'   DataEnvPublic.cmd060302
'   'datrpt060302.Orientation = rptOrientPortrait
'   datrpt060302.Sections(1).Controls("rptlblUserName").Caption = strUserName
'    'Add By Cheng 2002/12/20
'    Screen.MousePointer = vbDefault
'   ' 90.08.16 modify by louis
'   If bShow = True Then
'      datrpt060302.Show vbModal
'   End If
'   PUB_SetOsPrtAsApp 'Add by Morgan 2010/2/23
'   datrpt060302.PrintReport
'   PUB_RestoreOsPrt 'Add by Morgan 2010/2/23
'   Unload datrpt060302
'   DataEnvPublic.rscmdContact.Close
'   ClsPPPrtFCPMail = True
'   Exit Function
'ErrHand:
'      'edit by nickc 2007/02/02
'   'ErrorLog
'   MsgBox Err.Description
'End Function


'Remove by Morgan 2011/10/3 因已達檔案上限要移除不再使用的
''Add By Cheng 2003/01/09
'Public Function ClsPPPrtPMail(ByVal CP09 As String, ByVal strUserName As String, Optional ByVal bShow As Boolean = True) As Boolean
' Dim strTmp As String
' On Error GoTo ErrHand
'
'   DataEnvPublic.Commands(2).CommandText = "select st03 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
'      "pa07 a06,cu04 a07,FA04 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
'      "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN WHERE cp09='" & CP09 & "' AND " & _
'      "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
'      "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
'      "PA11=TPB01(+)"
'   DataEnvPublic.cmd060302
'   'datrpt060302.Orientation = rptOrientPortrait
'   datrpt060302.Sections(1).Controls("rptlblUserName").Caption = strUserName
'    'Add By Cheng 2002/12/20
'    Screen.MousePointer = vbDefault
'   ' 90.08.16 modify by louis
'   If bShow = True Then
'      datrpt060302.Show vbModal
'   End If
'   PUB_SetOsPrtAsApp 'Add by Morgan 2010/2/23
'   datrpt060302.PrintReport
'   PUB_RestoreOsPrt 'Add by Morgan 2010/2/23
'   Unload datrpt060302
'   DataEnvPublic.rscmdContact.Close
'   ClsPPPrtPMail = True
'   Exit Function
'ErrHand:
'      'edit by nickc 2007/02/02
'   'ErrorLog
'   MsgBox Err.Description
'End Function

'Removed by Morgan 2020/4/15
''Add By Cheng 2002/06/26
'Public Function ClsPPPrtMail(ByVal CP09 As String, ByVal strUserName As String, Optional ByVal bShow As Boolean = True, Optional strNote As String, Optional strDept As String, Optional strDate As String) As Boolean
' Dim strTmp As String
''Add By Cheng 2002/12/05
'On Error GoTo ErrHand
'
'   DataEnvPublic.Commands(2).CommandText = "select A0902 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
'      "pa07 a06,cu04 a07,FA05||FA63 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
'      "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN,ACC090 WHERE ST03=A0901 AND cp09='" & CP09 & "' AND " & _
'      "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
'      "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
'      "PA11=TPB01(+)"
'   DataEnvPublic.cmd060302
'   datrptPublicA.Sections(1).Controls("rptlblUserName").Caption = strUserName
'   datrptPublicA.Sections(1).Controls("Label15").Caption = strNote
'   datrptPublicA.Sections(1).Controls("Label14").Caption = strDept
'   datrptPublicA.Sections(1).Controls("rptlblDate").Caption = strDate
'    'Add By Cheng 2002/12/20
'    Screen.MousePointer = vbDefault
'   If bShow = True Then
'      datrptPublicA.Show vbModal
'   Else
'      PUB_SetOsPrtAsApp 'Add by Morgan 2010/2/23
'      datrptPublicA.PrintReport
'      PUB_RestoreOsPrt 'Add by Morgan 2010/2/23
'   End If
'   Unload datrptPublicA
'   DataEnvPublic.rscmd060302.Close
'   ClsPPPrtMail = True
'   Exit Function
'ErrHand:
'      'edit by nickc 2007/02/02
'   'ErrorLog
'   MsgBox Err.Description
'End Function

'Remove by Morgan 2011/10/3 因已達檔案上限要移除不再使用的
''Add By Cheng 2002/10/14
'Public Function ClsPPPrtMail_1(ByVal strCPNO01 As String, ByVal strCPNO02 As String, ByVal strCPNO03 As String, ByVal strCPNO04 As String, _
'      ByVal strUserName As String, Optional ByVal bShow As Boolean = True, Optional strNote As String, Optional strDept As String, Optional strDate As String, Optional strSaleZone As String, Optional strSaleName As String, _
'      Optional strCPName1 As String, Optional strCPName2 As String, Optional strCPName3 As String) As Boolean
'   Dim strTmp As String
'   Dim StrSQLa As String
'
''Add By Cheng 2002/12/05
'On Error GoTo ErrHand
'
'    'Modify By Cheng 2002/12/05
''   strSQLA = "select CU04 FROM PATENT,CUSTOMER " & _
''      "WHERE PA01='" & strCPNO01 & "' AND PA02='" & strCPNO02 & "' AND PA03='" & strCPNO03 & "' AND PA04='" & strCPNO04 & "' " & _
''      "AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) " & _
''      " UNION select CU04 FROM TRADEMARK,CUSTOMER " & _
''      "WHERE TM01='" & strCPNO01 & "' AND TM02='" & strCPNO02 & "' AND TM03='" & strCPNO03 & "' AND TM04='" & strCPNO04 & "' " & _
''      "AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) " & _
''      " UNION select CU04 FROM LAWCASE,CUSTOMER " & _
''      "WHERE LC01='" & strCPNO01 & "' AND LC02='" & strCPNO02 & "' AND LC03='" & strCPNO03 & "' AND LC04='" & strCPNO04 & "' " & _
''      "AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) " & _
''      " UNION select CU04 FROM HIRECASE,CUSTOMER " & _
''      "WHERE HC01='" & strCPNO01 & "' AND HC02='" & strCPNO02 & "' AND HC03='" & strCPNO03 & "' AND HC04='" & strCPNO04 & "' " & _
''      "AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) " & _
''      " UNION select CU04 FROM SERVICEPRACTICE,CUSTOMER " & _
''      "WHERE SP01='" & strCPNO01 & "' AND SP02='" & strCPNO02 & "' AND SP03='" & strCPNO03 & "' AND SP04='" & strCPNO04 & "' " & _
''      "AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) "
'   StrSQLa = "select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03 FROM PATENT,CUSTOMER,NATION " & _
'      "WHERE PA01='" & strCPNO01 & "' AND PA02='" & strCPNO02 & "' AND PA03='" & strCPNO03 & "' AND PA04='" & strCPNO04 & "' " & _
'      "AND SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) AND PA09=NA01(+) " & _
'      " UNION select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03 FROM TRADEMARK,CUSTOMER,NATION " & _
'      "WHERE TM01='" & strCPNO01 & "' AND TM02='" & strCPNO02 & "' AND TM03='" & strCPNO03 & "' AND TM04='" & strCPNO04 & "' " & _
'      "AND SUBSTR(TM23,1,8)=CU01(+) AND SUBSTR(TM23,9,1)=CU02(+) AND TM10=NA01(+) " & _
'      " UNION select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03 FROM LAWCASE,CUSTOMER,NATION " & _
'      "WHERE LC01='" & strCPNO01 & "' AND LC02='" & strCPNO02 & "' AND LC03='" & strCPNO03 & "' AND LC04='" & strCPNO04 & "' " & _
'      "AND SUBSTR(LC11,1,8)=CU01(+) AND SUBSTR(LC11,9,1)=CU02(+) AND LC15=NA01(+) " & _
'      " UNION select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03 FROM HIRECASE,CUSTOMER,NATION " & _
'      "WHERE HC01='" & strCPNO01 & "' AND HC02='" & strCPNO02 & "' AND HC03='" & strCPNO03 & "' AND HC04='" & strCPNO04 & "' " & _
'      "AND SUBSTR(HC05,1,8)=CU01(+) AND SUBSTR(HC05,9,1)=CU02(+) AND '000' = NA01(+) " & _
'      " UNION select nvl(CU04,decode(CU05,NULL,CU06,CU05||' '||CU88||' '||CU89||' '||CU90)) as CU04,nvl(NA03,NA04) as NA03 FROM SERVICEPRACTICE,CUSTOMER,NATION " & _
'      "WHERE SP01='" & strCPNO01 & "' AND SP02='" & strCPNO02 & "' AND SP03='" & strCPNO03 & "' AND SP04='" & strCPNO04 & "' " & _
'      "AND SUBSTR(SP08,1,8)=CU01(+) AND SUBSTR(SP08,9,1)=CU02(+) AND SP09=NA03(+) "
'
'   DataEnvPublic.Commands("cmd1106").CommandText = StrSQLa
'   DataEnvPublic.cmd1106
'   datrptPublicB.Sections(1).Controls("lblA01").Caption = strSaleZone
'   datrptPublicB.Sections(1).Controls("lblA02").Caption = strSaleName
'   datrptPublicB.Sections(1).Controls("lblA03").Caption = strCPNO01 & "-" & strCPNO02 & "-" & strCPNO03 & "-" & strCPNO04
'   datrptPublicB.Sections(1).Controls("lblA04").Caption = strCPName1
'   datrptPublicB.Sections(1).Controls("lblA05").Caption = strCPName2
'   datrptPublicB.Sections(1).Controls("lblA06").Caption = strCPName3
'   datrptPublicB.Sections(1).Controls("rptlblUserName").Caption = strUserName
'   datrptPublicB.Sections(1).Controls("Label15").Caption = strNote
'   datrptPublicB.Sections(1).Controls("Label14").Caption = strDept
'   datrptPublicB.Sections(1).Controls("rptlblDate").Caption = strDate
'    'Add By Cheng 2002/12/20
'    Screen.MousePointer = vbDefault
'   If bShow = True Then
'      datrptPublicB.Show vbModal
'   Else
'      PUB_SetOsPrtAsApp 'Add by Morgan 2010/2/23
'      datrptPublicB.PrintReport
'      PUB_SetOsPrtAsApp 'Add by Morgan 2010/2/23
'   End If
'   Unload datrptPublicB
'   DataEnvPublic.rscmd1106.Close
'   ClsPPPrtMail_1 = True
'   Exit Function
'ErrHand:
'      'edit by nickc 2007/02/02
'   'ErrorLog
'   MsgBox Err.Description
'End Function

'Remove by Morgan 2011/10/3 因已達檔案上限要移除不再使用的
''Add By Cheng 2002/06/26
''Modify By Cheng 2002/12/20
''Public Function PrtMail_2(ByVal CP09 As String, ByVal strUserName As String, Optional ByVal bShow As Boolean = True, Optional strNote As String, Optional strDept As String, Optional strDate As String) As Boolean
'Public Function ClsPPPrtMail_2(ByVal CP09 As String, ByVal strUserName As String, Optional ByVal bShow As Boolean = True, Optional strNote As String, Optional strDept As String, Optional strDate As String, Optional strDate1 As String) As Boolean
' Dim strTmp As String
''Add By Cheng 2002/12/05
'On Error GoTo ErrHand
'
'   DataEnvPublic.Commands(2).CommandText = "select A0902 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
'      "pa07 a06,cu04 a07,FA05||FA63 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
'      "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN,ACC090 WHERE ST03=A0901 AND cp09='" & CP09 & "' AND " & _
'      "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
'      "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
'      "PA11=TPB01(+)"
'   DataEnvPublic.cmd060302
'   datrptPublicC.Sections(1).Controls("rptlblUserName").Caption = strUserName
'   datrptPublicC.Sections(1).Controls("Label15").Caption = strNote
'   datrptPublicC.Sections(1).Controls("Label14").Caption = strDept
'   datrptPublicC.Sections(1).Controls("rptlblDate").Caption = strDate
'   datrptPublicC.Sections(1).Controls("rptlblDate1").Caption = strDate1
'    'Add By Cheng 2002/12/18
'   datrptPublicC.Sections(1).Controls("rptlblUserName_1").Caption = strUserName
'   datrptPublicC.Sections(1).Controls("Label15_1").Caption = strNote
'   datrptPublicC.Sections(1).Controls("Label14_1").Caption = strDept
'   datrptPublicC.Sections(1).Controls("rptlblDate_1").Caption = strDate
'   datrptPublicC.Sections(1).Controls("rptlblDate1_1").Caption = strDate1
'    'Add By Cheng 2002/12/20
'    Screen.MousePointer = vbDefault
'   If bShow = True Then
'      datrptPublicC.Show vbModal
'   Else
'      PUB_SetOsPrtAsApp 'Add by Morgan 2010/2/23
'      datrptPublicC.PrintReport
'      PUB_RestoreOsPrt 'Add by Morgan 2010/2/23
'   End If
'   Unload datrptPublicC
'   DataEnvPublic.rscmd060302.Close
'   ClsPPPrtMail_2 = True
'   Exit Function
'ErrHand:
'      'edit by nickc 2007/02/02
'   'ErrorLog
'   MsgBox Err.Description
'End Function

'Remove by Morgan 2011/10/3 因已達檔案上限要移除不再使用的
''Add By Cheng 2002/12/20
'Public Function ClsPPPrtMail_3(ByVal CP09 As String, ByVal strUserName As String, Optional ByVal bShow As Boolean = True, Optional strNote As String, Optional strDept As String, Optional strDate As String, Optional strDate1 As String) As Boolean
' Dim strTmp As String
'On Error GoTo ErrHand
'
'   DataEnvPublic.Commands(2).CommandText = "select A0902 a01,st02 a02," & ChgCaseprogress("", 1) & " a03,pa05 a04,pa06 a05," & _
'      "pa07 a06,cu04 a07,FA05||FA63 a08," & SQLDate("PA14", True) & " a09,TPB08 a10 " & _
'      "FROM CASEPROGRESS,PATENT,CASEPROPERTYMAP,STAFF,CUSTOMER,FAGENT,TPBULLETIN,ACC090 WHERE ST03=A0901 AND cp09='" & CP09 & "' AND " & _
'      "CP01=PA01 and CP02=PA02 and CP03=PA03 and CP04=PA04 and cp01=cpm01(+) and cp10=cpm02(+) and cp13=st01(+) and " & _
'      "SUBSTR(PA26,1,8)=CU01(+) AND SUBSTR(PA26,9,1)=CU02(+) and SUBSTR(PA75,1,8)=FA01(+) AND SUBSTR(PA75,9,1)=FA02(+) AND " & _
'      "PA11=TPB01(+)"
'   DataEnvPublic.cmd060302
'   datrptPublicD.Sections(1).Controls("rptlblUserName").Caption = strUserName
'   datrptPublicD.Sections(1).Controls("Label15").Caption = strNote
'   datrptPublicD.Sections(1).Controls("Label14").Caption = strDept
'   datrptPublicD.Sections(1).Controls("rptlblDate").Caption = strDate
'   datrptPublicD.Sections(1).Controls("rptlblDate1").Caption = strDate1
'    'Add By Cheng 2002/12/20
'    Screen.MousePointer = vbDefault
'   If bShow = True Then
'      datrptPublicD.Show vbModal
'   Else
'      PUB_SetOsPrtAsApp 'Add by Morgan 2010/2/23
'      datrptPublicD.PrintReport
'      PUB_RestoreOsPrt 'Add by Morgan 2010/2/23
'   End If
'   Unload datrptPublicD
'   DataEnvPublic.rscmd060302.Close
'   ClsPPPrtMail_3 = True
'   Exit Function
'ErrHand:
'      'edit by nickc 2007/02/02
'   'ErrorLog
'   MsgBox Err.Description
'End Function

'Added by Morgan 2022/3/28
'檢查EMail是否需客戶回覆
Public Function PUB_ChkIsRegMail(pLP01 As String) As Boolean
   Dim strQ As String
   Dim intQ As Integer
   Dim rstQ As ADODB.Recordset
   
   strQ = "select lp01 from letterprogress where lp01='" & pLP01 & "' and lp52='Y'"
   'Added by Morgan 2023/4/10
   '排除已閉卷的專利年費逾期通知
   strQ = strQ & " and not exists(select * from caseprogress,patent where cp09=lp01 and cp10='1605'" & _
      " and cp01 in ('P','CFP') and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa57 is not null)"
   'end 2023/4/10
   intQ = 1
   Set rstQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      PUB_ChkIsRegMail = True
   End If
   Set rstQ = Nothing
End Function

'Added by Morgan 2022/3/8
'下一程序
Public Function PUB_GetRefNPName(pCP09 As String) As String
   Dim strQ As String, strVTB As String
   Dim intQ As Integer
   Dim rstQ As ADODB.Recordset
   Dim arrKey(4) As String, strNextYear As String, strNextYearDesc As String
   
   If Left(pCP09, 1) = "D" Then
      strVTB = "select cp01,cp02,cp03,cp04,np07,nvl(pa09,tm10) pa09" & _
         " from caseprogress,nextprogress,patent,trademark" & _
         " where cp09='" & pCP09 & "' and np01(+)=cp43 and np22(+)=cp30 and np01 is not null and np06 is null" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04"
         
   ElseIf Left(pCP09, 1) = "C" Then
      strVTB = "select cp01,cp02,cp03,cp04,substr(min(np09||np07),9) np07,max(nvl(pa09,tm10)) pa09" & _
      " from caseprogress,nextprogress,patent,trademark" & _
      " where cp09='" & pCP09 & "' and np01(+)=cp09" & _
      " and np02(+)=cp01 and np03(+)=cp02 and np04(+)=cp03 and np05(+)=cp04 and length(np07)=3 and np06 is null" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " group by cp01,cp02,cp03,cp04,np07"
      
   End If
   If strVTB <> "" Then
      strQ = "select decode(pa09,'000',cpm03,cpm04),cp01,cp02,cp03,cp04,np07 from (" & strVTB & "),casepropertymap" & _
         " where cpm01(+)=cp01 and cpm02(+)=np07"
      intQ = 1
      Set rstQ = ClsLawReadRstMsg(intQ, strQ)
      If intQ = 1 Then
         PUB_GetRefNPName = "" & rstQ(0)
         If Right(rstQ("cp01"), 1) = "P" And InStr("605,606,607", rstQ("np07")) > 0 Then
            arrKey(1) = rstQ("cp01")
            arrKey(2) = rstQ("cp02")
            arrKey(3) = rstQ("cp03")
            arrKey(4) = rstQ("cp04")
            strNextYear = PUB_GetNextYear(arrKey, strNextYearDesc)
            If strNextYearDesc <> "" Then
               PUB_GetRefNPName = PUB_GetRefNPName & "[" & strNextYearDesc & "]"
            End If
         End If
         
      End If
   End If
End Function

'Added by Morgan 2022/4/12
'顧服組特定客戶專利案有固定承辦工程師，給客戶的信函要雙署名且密件副本給工程師
'Modified by Morgan 2024/8/5 +pIsFlow 是否歷程寄信
Public Function PUB_ChkW2001XPCase(ByVal pCP09 As String, ByRef pBCC As String, ByRef pSignature As String, ByVal pIsFlow As Boolean) As Boolean
   Dim strQ As String, strVTB As String
   Dim intQ As Integer
   Dim rstQ As ADODB.Recordset
   'Added by Morgan 2024/8/5
   Dim strEngPhone As String
   Dim strSign1 As String
   Dim strSign2  As String
   
   '雙署名有Unicode要抓員工檔
   'Modified by Morgan 2023/4/28 智權人員署名也改抓Table
   'strQ = "select cem03,cem04,st02 from caseprogress,patent,customer,custengmap,staff" & _
         " where cp09='" & pCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and cem01(+)=cu01 and (cem02='*' or cem02=pa158) and st01='A5024'" & _
         " order by decode(cem02,'*',2,1) asc"
   'Modified by Morgan 2024/8/5
   'strQ = "select cem03,cem04,cem05,cem06 from caseprogress,patent,customer,custengmap" & _
         " where cp09='" & pCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and cem01(+)=cu01 and (cem02='*' or cem02=pa158)" & _
         " order by decode(cem02,'*',2,1) asc"
   strQ = "select cem03,cem04,cem05,cem06,cu13,ed01,ed03,st02,ac03 from caseprogress,patent,customer,custengmap,ExtensionData,staff,allcode" & _
         " where cp09='" & pCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9) and cem01(+)=cu01 and (cem02='*' or cem02=pa158)" & _
         " and ed02(+)=cem03 and st01(+)=cem03 and ac02(+)=st20 and ac01(+)='01' order by decode(cem02,'*',2,1) asc"
   intQ = 1
   Set rstQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      PUB_ChkW2001XPCase = True
      pBCC = "" & rstQ("cem03")
      '因經理的名字有Unicode造成文字框顯示會對不齊(應該是2.0的問題),前面故意多加1個半形空白才會對齊,工程師名字前面也配合加1個半形空白
      '但若不加也無妨,收到的信還是對齊的
      'Modified by Morgan 2023/4/28
      'pSignature = "創新業務部　經理　 " & rstQ("st02") & "(#281)　專利代理人" & vbCrLf & "專利國內部　" & rstQ("cem04") & "　敬上"
      If Not IsNull(rstQ("cem06")) Then
         pBCC = pBCC & ";" & rstQ("cem06")
      End If
      
      
      'Added by Morgan 2024/8/5
      '智權改30015後落款改為歷程會稿工程師要帶電話分機,寄發文件則不必
      If rstQ("cu13") = "30015" Then
         strSign1 = rstQ("cem05")
         strSign2 = rstQ("cem04")
         If pIsFlow Then
            '北
            If rstQ("ed03") = "1" Then
               strEngPhone = "02-25061023"
            '中
            ElseIf rstQ("ed03") = "2" Then
               strEngPhone = "04-23270288"
            '南
            ElseIf rstQ("ed03") = "3" Then
               strEngPhone = "06-2743866"
            '高
            ElseIf rstQ("ed03") = "4" Then
               strEngPhone = "07-2363602"
            End If
            strSign2 = strSign2 & "(" & strEngPhone & "#" & rstQ("ed01") & ")"
         End If
         
         If GetTextLength(strSign1) > GetTextLength(strSign2) Then
            strSign2 = strSign2 & String(Trunc((GetTextLength(strSign1) - GetTextLength(strSign2)) / 2), "　")
         End If
         pSignature = strSign1 & vbCrLf & strSign2 & "　敬上"
      Else
      'end 2024/8/5
         pSignature = rstQ("cem05") & vbCrLf & "專利國內部　" & rstQ("cem04") & "　敬上"
      End If
      'end 2023/4/28
   End If
   Set rstQ = Nothing
End Function

'Added by Morgan 2014/12/18
'pbolDone:Email是否有寄發
'Modified by Morgan 2016/3/2 + pSubject:主旨, pOurDeadLine:本所期限, pOffDeadLine : 法定期限
'Modify By Sindy 2018/8/30 + , Optional pSaveMailBackup As Boolean = False : 是否存寄件備份
'                            , Optional ByRef strUpdDate As String, Optional ByRef strUpdTime As String
'                            , Optional pbolQueryMailData As Boolean = False : 查詢寄件備份
'                            , Optional pstrSeqno As String : 歷程序號
'                            , Optional pbolEMPFlow As Boolean = False : 是否為歷程作業
'                            , Optional pbolAutoTransmit As Boolean = False : 自動啟動轉寄功能
'                            , Optional pNote As String = "" : 備註
'                            , Optional pForm As Form : 呼叫E-Mail的Form
'Modify By Sindy 2018/9/27 + , Optional pManyCaseNum As Integer = 1 : 傳入案件數量
'Modify By Sindy 2018/10/22 + , Optional pRetrunRecvs As String = "" : 傳入多件收文文號
'Modified by Morgan 2018/10/30 , Optional pECustNo As String : 傳入ｅ化客戶編號(可能是申請人、副本收受人、移轉人...)
'Modify by Amy 2020/01/02 +strLP01:傳信函編號
'Modified by Morgan 2021/12/2 +pSalesNo
'Modified by Morgan 2022/2/18 +pIsRegMail:是否掛號直寄信
'Modified by Sindy 2024/10/16 +bolReadLP42:是否要讀"定稿合併收文號"的案件資料
Public Sub PUB_ShowMailForm(pCP09 As String, pFiles As String, pProperty As String, _
   Optional pbolDone As Boolean, Optional pSubject As String, Optional pOurDeadLine As String, _
   Optional pOffDeadLine As String, Optional pSaveMailBackup As Boolean = False, _
   Optional ByRef pstrUpdDate As String, Optional ByRef pstrUpdTime As String, _
   Optional pbolQueryMailData As Boolean = False, Optional pstrSeqno As String, _
   Optional pbolEMPFlow As Boolean = False, Optional pbolAutoTransmit As Boolean = False, _
   Optional pNote As String = "", Optional pForm As Form, Optional pManyCaseNum As Integer = 1, _
   Optional pRetrunRecvs As String = "", Optional pECustNo As String, Optional ByVal strLP01 As String = "", _
   Optional pSalesNo As String, Optional ByVal pIsRegMail As Boolean = False, Optional ByVal bolReadLP42 As Boolean = False)

Dim stContent As String
Dim stSQL As String, intR As Integer
Dim rsQuery As ADODB.Recordset
Dim rsQuery2 As ADODB.Recordset 'Added by Morgan 2025/1/10 原用RsTemp改為rsQuery2
'Add By Sindy 2018/8/30
Dim varTemp As Variant
Dim ii As Integer, lngInt As Long
Dim strFilePath As String, strFileName As String, strNewFileName As String
Dim fs As Object
Dim strCaseNumber As String
'2018/8/30 END
Dim strCopy As String 'Modify By Sindy 2019/9/5
Dim strSubEx As String   'Added by Morgan 2022/2/17主旨最後一段
Dim strRedText As String 'Added by Morgan 2022/2/17 要提醒的內文(紅色)
Dim strNPName As String 'Added by Morgan 2022/3/8 下一程序案件性質
Dim strBCC As String 'Added by Morgan 2022/4/12 密件副本
Dim bolW2001XPCase As Boolean, strEngNo As String, strSignature As String 'Added by Morgan 2022/4/12 是否顧服組特定客戶專利案件,承辦工程師,雙署名
Dim stCustNo As String, stContactNo As String 'Added by Morgan 2025/1/10
   
   'Add By Sindy 2024/10/16 檢查是否有多個案號
   If pRetrunRecvs = "" And bolReadLP42 = True Then
      strSql = "Select lp01 From letterprogress Where lp42='" & pCP09 & "' and lp01<>'" & pCP09 & "' order by lp01 asc"
      intI = 1
      Set rsQuery2 = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         pManyCaseNum = rsQuery2.RecordCount + 1
         pRetrunRecvs = pCP09
         rsQuery2.MoveFirst
         Do While Not rsQuery2.EOF
            pRetrunRecvs = pRetrunRecvs & "," & rsQuery2.Fields("lp01")
            rsQuery2.MoveNext
         Loop
         '必須為多個案號
         strSql = "Select cp01,cp02,cp03,cp04 From caseprogress Where cp09 in('" & Replace(pRetrunRecvs, ",", "','") & "') group by cp01,cp02,cp03,cp04"
         intI = 1
         Set rsQuery2 = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If rsQuery2.RecordCount = 1 Then
               pRetrunRecvs = ""
            End If
         End If
      End If
   End If
   If pRetrunRecvs = "" Then pManyCaseNum = 1
   '2024/10/16 END
   
   'Added by Morgan 2022/4/12
   'Modified by Morgan 2024/8/5 +傳入pbolEMPFlow
   bolW2001XPCase = PUB_ChkW2001XPCase(pCP09, strEngNo, strSignature, pbolEMPFlow)
   If bolW2001XPCase Then
      strBCC = strEngNo
   End If
   'end 2022/4/12
   
   stContent = ""
   'Added by Morgan 2022/2/17
   'Modified by Morgan 2022/2/18 +掛號直寄才要
   'Modified by Morgan 2022/3/28 不必再限定全E化客戶
   'If pECustNo <> "" And pIsRegMail = True Then
   If pbolEMPFlow = False Then
      If pIsRegMail = False Then
         pIsRegMail = PUB_ChkIsRegMail(pCP09)
      End If
   End If
   If pIsRegMail = True Then
   'end 2022/3/28
      strSubEx = "，請於收到後務必回覆已收到此通知信函"
      strRedText = "，請務必回覆已確實收到此通知信函，謝謝"
      'Added by Morgan 2022/3/8
      strNPName = PUB_GetRefNPName(pCP09)
      If strNPName <> "" Then
         stContent = "敬啟者：" & vbCrLf & vbCrLf & _
            "本案下一程序為" & strNPName & "，謹通知如附件並靜候指示。" & vbCrLf & vbCrLf
      End If
      'end 2022/3/8
   Else
      strSubEx = ",供參考"
      strRedText = ""
   End If
   'end 2022/2/17
   
   'Modified by Morgan 2016/3/2 +專利以外也可用
   'stSQL = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",pa48 客戶案件案號,cpm03 案件性質,cu04 申請人,pa11 申請案號,pa05 案件名稱,NA03 申請國家,DECODE(PA01,'CFP',PTM03,DECODE(PA09,'000',PTM03,PTM04)) 專利種類" & _
      ",pa26,pa27,pa75,pa149,cp01,cp02,cp03,cp04,st02,st06,ed01" & _
      " From caseprogress, patent,nation,PATENTTRADEMARKMAP, CUSTOMER, casepropertymap,staff,ExtensionData" & _
      " where cp09='" & pCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and na01(+)=pa09 AND PTM01(+)='1' AND PTM02(+)=PA08 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and st01(+)=cp13 and ed02(+)=cp13"
   'Modified by Morgan 2018/4/12 +pa77
   'Modify by Sindy 2018/9/13 + st15
   'Modify by Sindy 2018/9/20 + cp10
   'Modified by Morgan 2018/10/30 +pa10/tm11
   'Modified by Lydia 2019/09/02 +申請人N1
   'Modified by Morgan 2021/3/18 +pa09,pa106
   'Modified by Morgan 2024/1/19 申請人無中文時改抓英文 Ex:P-129951
   'Modified by Morgan 2024/5/15 +CU13
   stSQL = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",pa48 客戶案件案號,nvl(cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90)) 申請人,pa11 申請案號,nvl(pa05,nvl(pa06,pa07)) 案件名稱,NA03 申請國家,DECODE(PA01,'CFP',PTM03,DECODE(PA09,'000',PTM03,PTM04)) 專利種類" & _
      ",pa26,pa27,pa75,pa149,cp01,cp02,cp03,cp04,st02,st06,ed01,'1' 案件種類,pa77,st15,cp10,'' 類別" & _
      ",pa10,pa26 as 申請人N1,pa09,pa106,CU13 From caseprogress, patent,nation,PATENTTRADEMARKMAP, CUSTOMER,staff,ExtensionData" & _
      " where cp09='" & pCP09 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & _
      " and na01(+)=pa09 AND PTM01(+)='1' AND PTM02(+)=PA08 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " and st01(+)='" & strUserNum & "' and ed02(+)=st01"
   stSQL = stSQL & " union all " & _
      "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",tm35 客戶案件案號,nvl(cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90)) 申請人,tm12 申請案號,tm05 案件名稱,NA03 申請國家,DeCODE(TM10,'000',PTM03,PTM04) 專利種類" & _
      ",tm23,tm78,tm44,tm123,cp01,cp02,cp03,cp04,st02,st06,ed01,'2' 案件種類,tm45,st15,cp10,tm09 類別" & _
      ",tm11,tm23 as 申請人N1,tm10,'' pa106,CU13 From caseprogress, trademark,nation,PATENTTRADEMARKMAP, CUSTOMER,staff,ExtensionData" & _
      " where cp09='" & pCP09 & "' and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & _
      " and na01(+)=tm10 AND PTM01(+)='2' AND PTM02(+)=tm08 and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9)" & _
      " and st01(+)='" & strUserNum & "' and ed02(+)=st01"
   stSQL = stSQL & " union all " & _
      "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",lc17 客戶案件案號,nvl(cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90)) 申請人,'' 申請案號,lc05 案件名稱,NA03 申請國家,'' 專利種類" & _
      ",lc11,'',lc22,lc42,cp01,cp02,cp03,cp04,st02,st06,ed01,'3' 案件種類,lc23,st15,cp10,'' 類別" & _
      ",null pa10,lc11 as 申請人N1,lc15,'' pa106,CU13 From caseprogress, lawcase,nation, CUSTOMER,staff,ExtensionData" & _
      " where cp09='" & pCP09 & "' and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04 and lc01 is not null" & _
      " and na01(+)=lc15 and cu01(+)=substr(lc11,1,8) and cu02(+)=substr(lc11,9)" & _
      " and st01(+)='" & strUserNum & "' and ed02(+)=st01"
   stSQL = stSQL & " union all " & _
      "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",'' 客戶案件案號,nvl(cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90)) 申請人,'' 申請案號,hc06 案件名稱,'' 申請國家,'' 專利種類" & _
      ",hc05,'','',hc23,cp01,cp02,cp03,cp04,st02,st06,ed01,'4' 案件種類,'',st15,cp10,'' 類別" & _
      ",null pa10,hc05 as 申請人N1,'' pa09,'' pa106,CU13 From caseprogress, hirecase, CUSTOMER,staff,ExtensionData" & _
      " where cp09='" & pCP09 & "' and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04 and hc01 is not null" & _
      " and cu01(+)=substr(hc05,1,8) and cu02(+)=substr(hc05,9)" & _
      " and st01(+)='" & strUserNum & "' and ed02(+)=st01"
   stSQL = stSQL & " union all " & _
      "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",sp29 客戶案件案號,nvl(cu04,rtrim(cu05||' '||cu88||' '||cu89||' '||cu90)) 申請人,sp11 申請案號,sp05 案件名稱,NA03 申請國家,'' 專利種類" & _
      ",sp08,sp58,sp26,sp78,cp01,cp02,cp03,cp04,st02,st06,ed01,'5' 案件種類,sp27,st15,cp10,'' 類別" & _
      ",null pa10,sp08 as 申請人N1,sp09,'' pa106,CU13 From caseprogress, servicepractice,nation, CUSTOMER,staff,ExtensionData" & _
      " where cp09='" & pCP09 & "' and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & _
      " and na01(+)=sp09 and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9)" & _
      " and st01(+)='" & strUserNum & "' and ed02(+)=st01"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With rsQuery
      
      Screen.MousePointer = vbHourglass
      'Add By Sindy 2018/8/30
      If pSaveMailBackup = True Then
         'E-Mail呼叫 frm880019:要將寄信的內容及寄信的成功時間儲存在資料庫中，便於事後查詢。
         frm880019.m_bolSaveMail = True
         frm880019.m_CP01 = .Fields("cp01")
         frm880019.m_CP02 = .Fields("cp02")
         frm880019.m_CP03 = .Fields("cp03")
         frm880019.m_CP04 = .Fields("cp04")
         frm880019.m_CP09 = pCP09
         frm880019.m_CP10 = .Fields("cp10") 'Add By Sindy 2018/9/20
         frm880019.SetParent pForm
         frm880019.m_LP01 = strLP01 'Add by Amy 2020/01/02
      End If
     '2018/8/30 END
     
      frm880019.m_CU13 = "" & .Fields("CU13") 'Added by Morgan 2024/5/15
      frm880019.m_CustCaseNo = "" & .Fields("客戶案件案號") 'Added by Morgan 2023/7/19
      
      '主旨
      'Modified by Morgan 2022/2/17
      If pSubject <> "" Then
         frm880019.txtSubject = pSubject
      'Add By Sindy 2018/9/27
      ElseIf pManyCaseNum > 1 Then '多案件
         'Modify By Sindy 2018/10/15 調整主旨內容
         'frm880019.txtSubject = "呈送案件(" & pProperty & ")" & pManyCaseNum & "件電子檔" & IIf(pbolEMPFlow = True, ",請會稿", ",供參考")
         frm880019.txtSubject = .Fields("本所案號") & " 號" & "「" & .Fields("案件名稱") & "」" & "(" & pProperty & ")" & "電子檔" & IIf(pbolEMPFlow = True, ",請會稿", strSubEx)
         If .Fields("案件種類") = "2" Then
            frm880019.txtSubject = "「" & .Fields("申請人") & "」商標(" & pProperty & ")電子檔(共" & pManyCaseNum & "件)" & IIf(pbolEMPFlow = True, ",請會稿", strSubEx)
         End If
      
      'Added by Morgan 2021/3/18 寶齡富錦 Y55435 案件--唐韻如
      ElseIf .Fields("pa75") = "Y55435000" Then
         frm880019.txtSubject = "LY/th - " & PUB_GetNationEngNameForLet(.Fields("pa09")) & " Patent Application No. " & .Fields("申請案號") & "; Your Ref: " & .Fields("pa77") & "; Our Ref: " & .Fields("本所案號") & "--" & .Fields("pa106")
      'end2021/3/18
      
      'Added by Morgan 2018/4/12 有代理人及彼所案號 --玲玲
      ElseIf Not IsNull(.Fields("pa75")) And Not IsNull(.Fields("pa77")) Then
         frm880019.txtSubject = "呈送 貴方案號：" & .Fields("pa77") & ",我方案號：" & .Fields("本所案號") & IIf(IsNull(.Fields("客戶案件案號")), "", "(" & .Fields("客戶案件案號") & ")") & " 案(" & pProperty & ")電子檔" & IIf(pbolEMPFlow = True, ",請會稿", strSubEx)
      'end 2018/4/12
      Else
         'Modify By Sindy 2018/9/5
         'frm880019.txtSubject = "呈送 " & .Fields("本所案號") & IIf(IsNull(.Fields("客戶案件案號")), "", "(" & .Fields("客戶案件案號") & ")") & " 案(" & pProperty & ")電子檔" & IIf(pbolEMPFlow = True, ",請會稿", ",供參考")
         'Modify By Sindy 2018/10/12 調整主旨內容
         'frm880019.txtSubject = "呈送 " & IIf(IsNull(.Fields("客戶案件案號")), "", "[客戶案號:" & .Fields("客戶案件案號") & "]") & .Fields("本所案號") & " 案(" & pProperty & ")電子檔" & IIf(pbolEMPFlow = True, ",請會稿", ",供參考")
         '[客戶案號:lll] P-120821 號[案件名稱](實用新型申請)電子檔,請會稿
         '[客戶案號:lll] T-217403 號[案件名稱](商標申請)(第30類)電子檔,請會稿
         'Added by Lydia 2019/09/02 和碩案件之電子郵件主旨請依客戶要求帶入下列字樣：[專利申請] 客戶案件案號(本所案號)
         If .Fields("案件種類") = "1" And Left("" & .Fields("申請人N1"), 6) = "X70017" Then
            frm880019.txtSubject = "[專利申請]" & IIf(IsNull(.Fields("客戶案件案號")), "", "[客戶案號:" & .Fields("客戶案件案號") & "]") & .Fields("本所案號") & " 號" & "「" & .Fields("案件名稱") & "」" & "(" & pProperty & ")" & "電子檔" & IIf(pbolEMPFlow = True, ",請會稿", strSubEx)
            
         ElseIf .Fields("案件種類") = "2" Then
            'Modified by Morgan 2022/3/28
            'frm880019.txtSubject = IIf(IsNull(.Fields("客戶案件案號")), "", "[客戶案號:" & .Fields("客戶案件案號") & "]") & .Fields("本所案號") & " 號" & "「" & .Fields("案件名稱") & "」" & "(" & pProperty & ")" & "(第" & .Fields("類別") & "類)" & "電子檔" & IIf(pbolEMPFlow = True, ",請會稿", strSubEx)
            '有客戶案件案號: 客戶案件案號+案件性質+本所案號+案件名稱+類別
            If Not IsNull(.Fields("客戶案件案號")) Then
               frm880019.txtSubject = "[客戶案號:" & .Fields("客戶案件案號") & "](" & pProperty & ")" & .Fields("本所案號") & " 號「" & .Fields("案件名稱") & "」(第" & .Fields("類別") & "類)" & "電子檔" & IIf(pbolEMPFlow = True, ",請會稿", strSubEx)
            '無客戶案件案號: 本所案號+案件性質+案件名稱+類別
            Else
               frm880019.txtSubject = .Fields("本所案號") & " 號(" & pProperty & ")「" & .Fields("案件名稱") & "」(第" & .Fields("類別") & "類)" & "電子檔" & IIf(pbolEMPFlow = True, ",請會稿", strSubEx)
            End If
            'end 2022/3/28
         Else
            'Modified by Morgan 2022/3/28
            'frm880019.txtSubject = IIf(IsNull(.Fields("客戶案件案號")), "", "[客戶案號:" & .Fields("客戶案件案號") & "]") & .Fields("本所案號") & " 號" & "「" & .Fields("案件名稱") & "」" & "(" & pProperty & ")" & "電子檔" & IIf(pbolEMPFlow = True, ",請會稿", strSubEx)
            '有客戶案件案號: 客戶案件案號+案件性質+本所案號+案件名稱
            If Not IsNull(.Fields("客戶案件案號")) Then
               frm880019.txtSubject = "[客戶案號:" & .Fields("客戶案件案號") & "](" & pProperty & ")" & .Fields("本所案號") & " 號「" & .Fields("案件名稱") & "」電子檔" & IIf(pbolEMPFlow = True, ",請會稿", strSubEx)
            '無客戶案件案號: 本所案號+案件性質+案件名稱
            Else
               frm880019.txtSubject = .Fields("本所案號") & " 號(" & pProperty & ")「" & .Fields("案件名稱") & "」電子檔" & IIf(pbolEMPFlow = True, ",請會稿", strSubEx)
            End If
         End If
      End If
      
      'Add By Sindy 2018/9/27
      If pManyCaseNum = 1 Then '單筆案件
      '2018/9/27 END
         If Not IsNull(.Fields("客戶案件案號")) Then
            stContent = stContent & "客戶案件案號：" & .Fields("客戶案件案號") & vbCrLf
         End If
         stContent = stContent & "本所案號：" & .Fields("本所案號") & vbCrLf
         
         If .Fields("案件種類") = "3" Then
            stContent = stContent & "相關國家：" & .Fields("申請國家") & vbCrLf
         ElseIf .Fields("案件種類") <> "4" Then
            stContent = stContent & "申請國家：" & .Fields("申請國家") & vbCrLf
         End If
         
         If .Fields("案件種類") = "1" Then
            stContent = stContent & "專利種類：" & .Fields("專利種類") & vbCrLf
         ElseIf .Fields("案件種類") = "2" Then
            'Modify By Sindy 2018/10/12 + 商品類別
            stContent = stContent & "商標種類：" & .Fields("專利種類") & "(第" & .Fields("類別") & "類)" & vbCrLf
         End If
         
         stContent = stContent & "案件性質：" & pProperty & vbCrLf
         If pOurDeadLine <> "" Then
            stContent = stContent & "本所期限：" & pOurDeadLine & vbCrLf
         End If
         
         If pOffDeadLine <> "" Then
            stContent = stContent & "法定期限：" & pOffDeadLine & vbCrLf
         End If
         
         If .Fields("案件種類") = "3" Or .Fields("案件種類") = "4" Then
            stContent = stContent & "當事人　：" & .Fields("申請人") & vbCrLf
         Else
            stContent = stContent & "申請人　：" & .Fields("申請人") & vbCrLf
         End If
         
         If Not IsNull(.Fields("pa27")) Then
            stContent = stContent & "　　　　（多人申請）" & vbCrLf
         End If
         
         'Modify By Sindy 2018/10/12 申請案號有資料時,才出現此列
         If (.Fields("案件種類") = "1" Or .Fields("案件種類") = "2" Or .Fields("案件種類") = "5") And _
            "" & .Fields("申請案號") <> "" Then
            stContent = stContent & "申請案號：" & .Fields("申請案號") & vbCrLf
         End If
         
         'Added by Morgan 2018/10/30 配合ｅ化客戶要求，"通知申請案號"之電子郵件內容，增加"申請日期"--文雄
         If .Fields("cp10") = "1101" Then
            If Not IsNull(.Fields("pa10")) Then
               stContent = stContent & "申請日期：" & ChangeWStringToTDateString(.Fields("pa10")) & vbCrLf
            End If
         End If
         'end 2018/10/29
         
         stContent = stContent & "案件名稱：" & .Fields("案件名稱") & vbCrLf
         
         If pFiles <> "" Then
            'Modify By Sindy 2018/8/30 檢查附件檔名是否有客戶案件案號,若沒,要加註上去
            If Not IsNull(.Fields("客戶案件案號")) Then
               strExc(10) = ""
               Set fs = CreateObject("Scripting.FileSystemObject")
               varTemp = Split(pFiles, ";")
               For ii = 0 To UBound(varTemp)
                  If varTemp(ii) <> "" Then 'Add By Sindy 2018/9/17 +if
                     lngInt = InStrRev(varTemp(ii), "\")
                     strFilePath = Left(varTemp(ii), lngInt)
                     strFileName = Right(varTemp(ii), Len(varTemp(ii)) - lngInt)
                     'Modify By Sindy 2019/4/29 ex:CFP-30337 exciplex;5,7spiroTB-010-ET;PNFL-001-001-ET
                     strCaseNumber = PUB_FilterEFileSymbol(.Fields("客戶案件案號"))
                     If InStr(UCase(strFileName), UCase(strCaseNumber)) = 0 Then
                        strNewFileName = strFilePath & strCaseNumber & "." & strFileName
                        fs.CopyFile varTemp(ii), strNewFileName
                        strExc(10) = strExc(10) & ";" & strNewFileName
                     Else
                        strExc(10) = strExc(10) & ";" & varTemp(ii)
                     End If
                  End If
               Next ii
               pFiles = Mid(strExc(10), 2)
               Set fs = Nothing
            End If
            '2018/8/30 END
         End If
         
      'Modify By Sindy 2018/9/27
      Else '多筆案件
         stContent = PUB_GetMailManyCaseData(pRetrunRecvs)
      '2018/9/27 END
      End If
      
      If pFiles <> "" Then
         stContent = stContent & vbCrLf & "附檔為本案之相關電子檔文件" & strRedText & "。" & vbCrLf
      End If
      If pNote <> "" Then stContent = stContent & vbCrLf & pNote 'Add By Sindy 2018/9/13
      stContent = stContent & vbCrLf & vbCrLf
      
   'Add By Sindy 2021/5/3 轉寄不要加簽名檔
   If pbolAutoTransmit = False Then
   '2021/5/3 END
      
      'Added by Morgan 2022/4/12
      If bolW2001XPCase Then
         stContent = stContent & strSignature & vbCrLf & vbCrLf
      'Added by Morgan 2024/8/5
      ElseIf .Fields("cu13") = "30015" Then
         stContent = stContent & "智權部主任　鄭鈺華(06-2743866#66)　敬上" & vbCrLf & vbCrLf
      'end 2024/8/5
      End If
      'end 2022/4/12
   
      stContent = stContent & PUB_GetCompName(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04")) & vbCrLf '公司
      
      'Added by Morgan 2022/4/15 顧服組特定客戶專利案件為雙署名,固定帶北所及patent信箱
      If bolW2001XPCase Then
         stContent = stContent & "電　話：(02)25061023" & vbCrLf
         stContent = stContent & "傳　真：(02)25011666" & vbCrLf
         'Added by Morgan 2024/8/5
         If .Fields("cu13") = "30015" Then
            stContent = stContent & "E-MAIL：inno.ip@taie.com.tw" & vbCrLf
         Else
         'end 2024/8/5
            stContent = stContent & "E-MAIL：patent@taie.com.tw" & vbCrLf
         End If
      
      'Added by Morgan 2024/8/5
      ElseIf .Fields("cu13") = "30015" Then
         stContent = stContent & "電　話：(02)25061023" & vbCrLf
         stContent = stContent & "傳　真：(02)25011666" & vbCrLf
         stContent = stContent & "E-MAIL：inno.ip@taie.com.tw" & vbCrLf
      'end 2024/8/5
      
      'Added by Morgan 2025/5/15 P1004非雙署名客戶帶固定落款
      ElseIf .Fields("cu13") = "P1004" Then
         stContent = stContent & "郭雅娟" & vbCrLf
         stContent = stContent & "電　話：(02)25061023#350" & vbCrLf
         stContent = stContent & "傳　真：(02)25011666" & vbCrLf
         stContent = stContent & "E-MAIL：patent@taie.com.tw" & vbCrLf
      'end 2024/5/15
      
      Else
      'end 2022/4/15
      
         'Modified by Morgan 2018/4/12 有代理人時不要帶智權人員 --玲玲
         If IsNull(.Fields("pa75")) Then
            stContent = stContent & .Fields("st02") & vbCrLf '智權人員
         End If
         'end 2018/4/12
         Select Case "" & .Fields("st06")
         Case "1" '北
            stContent = stContent & "電　話：(02)25061023" & IIf(IsNull(.Fields("ed01")), "", "#" & .Fields("ed01")) & vbCrLf
            stContent = stContent & "傳　真：(02)25011666" & vbCrLf
         Case "2" '中
            stContent = stContent & "電　話：(04)23270288" & IIf(IsNull(.Fields("ed01")), "", "#" & .Fields("ed01")) & vbCrLf
            stContent = stContent & "傳　真：(04)23227483" & vbCrLf
         Case "3" '南
            stContent = stContent & "電　話：(06)2743866" & IIf(IsNull(.Fields("ed01")), "", "#" & .Fields("ed01")) & vbCrLf
            stContent = stContent & "傳　真：(06)2744030" & vbCrLf
         Case "4" '高
            stContent = stContent & "電　話：(07)2363602" & IIf(IsNull(.Fields("ed01")), "", "#" & .Fields("ed01")) & vbCrLf
            stContent = stContent & "傳　真：(07)2364360" & vbCrLf
         End Select
         'Add By Sindy 2018/9/5 E-Mail
         'stContent = stContent & "E-MAIL：" & strUserNum & "@taie.com.tw" & vbCrLf
         'Modify By Sindy 2018/9/19
         If Left(Pub_StrUserSt15, 2) = "P1" Then '專利處
            stContent = stContent & "E-MAIL：patent@taie.com.tw" & vbCrLf
         ElseIf Left(Pub_StrUserSt15, 1) = "S" Then '智權部
            stContent = stContent & "E-MAIL：" & strUserNum & "@taie.com.tw" & vbCrLf
         End If
         '2018/9/19 END
      
      End If 'Added by Morgan 2022/4/15
      
      stContent = stContent & "URL:https://www.taie.com.tw" & vbCrLf
      stContent = stContent & "*************保密警語******************** " & vbCrLf
      stContent = stContent & "本信件僅授權於指定之收信人取閱之用，信件中可能含有機密性資訊。" & vbCrLf
      stContent = stContent & "如果您並非被指定之收信人，任何未經授權而擅自使用此信件所含之機密資訊的行為是被嚴格禁止的。" & vbCrLf
      stContent = stContent & "如果您在任何未經授權的情形之下收到本信件，煩請您立即告知原發信人並將此信件回傳至以上地址。" & vbCrLf
      stContent = stContent & "謝謝您的合作。" & vbCrLf
      
   End If
   
      '本文
      frm880019.txtContent = stContent
      '附件
      frm880019.SetAttach pFiles
      
      'Added by Morgan 2021/12/2
      If pSalesNo <> "" Then
         frm880019.txtSubject = "【全E化會稿】" & frm880019.txtSubject
         frm880019.txtContent = "本案為全E化客戶案件，請確認後回覆以便後續EMail通知客戶！" & vbCrLf & vbCrLf & stContent
         frm880019.txtReceiver = pSalesNo & " (" & GetStaffName(pSalesNo, True) & "); "
      'end 2021/12/2
      'Added by Morgan 2021/3/18 寶齡富錦 Y55435 案件--唐韻如
      ElseIf .Fields("pa75") = "Y55435000" Then
         frm880019.SetPBFEmail .Fields("pa26")
      'end 2021/3/18
      
      'Added by Morgan 2018/10/30
      '若有傳入ｅ化客戶編號時收件人要抓該編號的指定信箱
      ElseIf pECustNo <> "" Then
         frm880019.m_RedText = strRedText
         frm880019.SetECustEmail pECustNo
      Else
         frm880019.m_RedText = strRedText 'Added by Morgan 2022/6/10 非全E化也要
      'end 2018/10/30
      
         'Modified by Morgan 2018/4/12 有代理人時副本不要寄給自己 --玲玲
         'Modify By Sindy 2018/9/4 + IIf(IsNull(.Fields("pa75")), False, True) ==> IIf(pSaveMailBackup = True, True, IIf(IsNull(.Fields("pa75")), False, True))
         'Modify By Sindy 2018/9/13 S單位才要副本=自己 + IIf(IsNull(.Fields("pa75")), False, True) ==> IIf(Left(.Fields("st15"), 1) = "S" And pSaveMailBackup = True, False, IIf(IsNull(.Fields("pa75")), False, True))
         'Modify By Sindy 2019/9/5 + 客服組專利會稿工程師執行客戶會稿，同時副本通知客服組成員。
         If UCase(TypeName(pForm)) = UCase("frm090202_2") Then
            If InStr(Pub_GetSpecMan("WSpecial"), pForm.m_FlowUserNum) > 0 Or _
               InStr(Pub_GetSpecMan("客服組專利會稿工程師"), strUserNum) > 0 Then '創新業務部可個人收文成員
               strCopy = pForm.m_FlowUserNum
            End If
         End If
         'Modify By Sindy 2019/9/5 + strCopy
         'Modified by Morgan 2022/4/12 +strBCC
         'Modified by Morgan 2024/5/15 P1004的客戶也不要副本給自己(客戶轉移的過渡期會由智權人員操作)
         'Modified by Morgan 2025/1/10 若為副本信函時收件人要抓LP33,LP34
         'frm880019.SetEmail "" & .Fields("pa26"), "" & .Fields("pa149"), "" & .Fields("pa75"), , IIf(pSaveMailBackup = True, IIf(Left(.Fields("st15"), 1) = "S" And .Fields("cu13") <> "P1004", False, True), IIf(IsNull(.Fields("pa75")), False, True)), strCopy, strBCC
         stCustNo = "" & .Fields("pa26")
         stContactNo = "" & .Fields("pa149")
         stSQL = "select lp33,lp34 from caseprogress,letterprogress where cp09='" & pCP09 & "' and cp10='990' and lp01(+)=cp09 and lp33 is not null"
         intR = 1
         Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
         If intR = 1 Then
            stCustNo = rsQuery("lp33")
            stContactNo = "" & rsQuery("lp34")
         End If
         frm880019.SetEmail stCustNo, stContactNo, "" & .Fields("pa75"), , IIf(pSaveMailBackup = True, IIf(Left(.Fields("st15"), 1) = "S" And .Fields("cu13") <> "P1004", False, True), IIf(IsNull(.Fields("pa75")), False, True)), strCopy, strBCC
         'end 2025/1/10
         
      End If 'Added by Morgan 2018/10/30
      
      'Add By Sindy 2018/8/31 查詢寄件備份
      If pbolQueryMailData = True Then
         frm880019.Hide
         frm880019.Caption = "寄件備份"
         frm880019.cmdExit.Caption = "結束"
'         frm880019.cmdSend.Caption = "轉寄"
'         frm880019.cmdSend.Visible = True
         'Modify By Sindy 2018/10/30
         If pbolAutoTransmit = True Then
            frm880019.cmdSend.Caption = "轉寄"
            frm880019.txtContent.Tag = frm880019.txtContent.Text 'Add By Sindy 2021/5/3
         End If
         '2018/10/30 END
         frm880019.cmdSend.Visible = False
         frm880019.cmdAttach.Visible = False
         frm880019.cmdReceiver(0).Visible = False
         frm880019.cmdReceiver(1).Visible = False
         frm880019.cmdReceiver(2).Visible = False
         frm880019.txtReceiver.Locked = True
         frm880019.txtBCC.Locked = True
         frm880019.txtCopy.Locked = True
         frm880019.txtSubject.Locked = True
         frm880019.txtAttachment.Locked = True
         frm880019.txtContent.Locked = True
         frm880019.FramePrint.Visible = True
         frm880019.m_CP09 = pCP09 '總收文號
         If pstrSeqno <> "" Then
            frm880019.m_SMB11 = pstrSeqno '歷程序號
         Else
            frm880019.m_SMB02 = pstrUpdDate '寄件日期
            frm880019.m_SMB03 = pstrUpdTime '寄件時間
         End If
         'frm880019.SetParent oForm
         If frm880019.QueryData = False Then
            Unload frm880019
            Set rsQuery = Nothing
            Exit Sub
         End If
         If pbolAutoTransmit = True Then '轉寄功能
            'frm880019.cmdSend.Caption = "轉寄" 'Modify By Sindy 2018/10/30 Mark
            frm880019.cmdSend.Visible = True
            frm880019.cmdSend_Click
         End If
      
      Else
         'Added by Morgan 2023/9/15 開放文雄回條測試
         'Modified by Morgan 2023/10/5 10/11全面開放
         If strUserNum = "A4023" Or Pub_StrUserSt03 = "M51" Or strSrvDate(1) >= "20231011" Then
            frm880019.chkReceipt.Visible = True
         End If
         'end 2023/9/15
      End If
      
      Screen.MousePointer = vbDefault
      
      frm880019.Show vbModal
      'Added by Morgan 2015/6/17
      pbolDone = frm880019.m_bolDone
      pstrUpdDate = frm880019.m_SMB02 'Add by Sindy 2018/8/30
      pstrUpdTime = frm880019.m_SMB03 'Add by Sindy 2018/8/30
      Unload frm880019
      'end 2015/6/17
      End With
   End If
   
   Set rsQuery = Nothing
End Sub

'Add By Sindy 2018/10/22 讀取案件資訊
'Add By Sindy 2020/7/8 + , Optional ByRef strCaseNoData As String
'Add By Sindy 2020/10/13 + , Optional ByRef strCP14 As String : 回傳承辦人
'                          , Optional ByVal bolTBase As Boolean = False : 商標處給代理人的案件資訊
'Add By Sindy 2020/12/1 + Optional ByRef strCPMNm As String = "" : 回傳多案的案件性質
Public Function PUB_GetMailManyCaseData(pCP09 As String, Optional ByRef strCaseNoData As String, _
   Optional ByRef strCP14 As String, Optional ByVal bolTBase As Boolean = False, _
   Optional ByRef strCPMNm As String = "")
Dim stContent As String
Dim stSQL As String, intR As Integer
Dim rsQuery As ADODB.Recordset
Dim bolShowCPMNm As Boolean 'Add By Sindy 2020/12/1
   
   Screen.MousePointer = vbHourglass
   
   PUB_GetMailManyCaseData = "": stContent = "": strCaseNoData = "": strCPMNm = ""
   
   stSQL = "select /*+index(caseprogress primary_key)*/ cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",pa48 客戶案件案號,cu04 申請人,pa11 申請案號,pa05 案件名稱,NA03 申請國家,DECODE(PA01,'CFP',PTM03,DECODE(PA09,'000',PTM03,PTM04)) 專利種類,nvl(DECODE(Pa09,'000',cpm03,cpm04),cp10) AS 案件性質" & _
      ",pa26,pa27,pa75,pa149,cp01,cp02,cp03,cp04,'1' 案件種類,pa77,cp10,'' 類別,SQLDATET(cp06) as 本所期限,SQLDATET(cp07) as 法定期限,CP14,pa22 tm15" & _
      " From caseprogress, patent, nation, PATENTTRADEMARKMAP, CUSTOMER, casepropertymap" & _
      " where cp09 in('" & Replace(pCP09, ",", "','") & "') and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa01 is not null" & _
      " and na01(+)=pa09 AND PTM01(+)='1' AND PTM02(+)=PA08 and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)" & _
      " and cp01=cpm01(+) and cp10=cpm02(+)"
   stSQL = stSQL & " union " & _
      "select /*+index(caseprogress primary_key)*/ cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",tm35 客戶案件案號,cu04 申請人,tm12 申請案號,tm05 案件名稱,NA03 申請國家,DeCODE(TM10,'000',PTM03,PTM04) 專利種類,nvl(DECODE(tm10,'000',cpm03,cpm04),cp10) AS 案件性質" & _
      ",tm23,tm78,tm44,tm123,cp01,cp02,cp03,cp04,'2' 案件種類,tm45,cp10,tm09 類別,SQLDATET(cp06) as 本所期限,SQLDATET(cp07) as 法定期限,CP14,tm15" & _
      " From caseprogress, trademark,nation, PATENTTRADEMARKMAP, CUSTOMER, casepropertymap" & _
      " where cp09 in('" & Replace(pCP09, ",", "','") & "') and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04 and tm01 is not null" & _
      " and na01(+)=tm10 AND PTM01(+)='2' AND PTM02(+)=tm08 and cu01(+)=substr(tm23,1,8) and cu02(+)=substr(tm23,9)" & _
      " and cp01=cpm01(+) and cp10=cpm02(+)"
   stSQL = stSQL & " union " & _
      "select /*+index(caseprogress primary_key)*/ cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",lc17 客戶案件案號,cu04 申請人,'' 申請案號,lc05 案件名稱,NA03 申請國家,'' 專利種類,nvl(DECODE(lc15,'000',cpm03,cpm04),cp10) AS 案件性質" & _
      ",lc11,'',lc22,lc42,cp01,cp02,cp03,cp04,'3' 案件種類,lc23,cp10,'' 類別,SQLDATET(cp06) as 本所期限,SQLDATET(cp07) as 法定期限,CP14,'' tm15" & _
      " From caseprogress, lawcase, nation, CUSTOMER, casepropertymap" & _
      " where cp09 in('" & Replace(pCP09, ",", "','") & "') and lc01(+)=cp01 and lc02(+)=cp02 and lc03(+)=cp03 and lc04(+)=cp04 and lc01 is not null" & _
      " and na01(+)=lc15 and cu01(+)=substr(lc11,1,8) and cu02(+)=substr(lc11,9)" & _
      " and cp01=cpm01(+) and cp10=cpm02(+)"
   stSQL = stSQL & " union " & _
      "select /*+index(caseprogress primary_key)*/ cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",'' 客戶案件案號,cu04 申請人,'' 申請案號,hc06 案件名稱,'' 申請國家,'' 專利種類,nvl(cpm03,cp10) AS 案件性質" & _
      ",hc05,'','',hc23,cp01,cp02,cp03,cp04,'4' 案件種類,'',cp10,'' 類別,SQLDATET(cp06) as 本所期限,SQLDATET(cp07) as 法定期限,CP14,'' tm15" & _
      " From caseprogress, hirecase, CUSTOMER, casepropertymap" & _
      " where cp09 in('" & Replace(pCP09, ",", "','") & "') and hc01(+)=cp01 and hc02(+)=cp02 and hc03(+)=cp03 and hc04(+)=cp04 and hc01 is not null" & _
      " and cu01(+)=substr(hc05,1,8) and cu02(+)=substr(hc05,9)" & _
      " and cp01=cpm01(+) and cp10=cpm02(+)"
   stSQL = stSQL & " union " & _
      "select /*+index(caseprogress primary_key)*/ cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04) 本所案號" & _
      ",sp29 客戶案件案號,cu04 申請人,sp11 申請案號,sp05 案件名稱,NA03 申請國家,'' 專利種類,nvl(DECODE(SP09,'000',cpm03,cpm04),cp10) AS 案件性質" & _
      ",sp08,sp58,sp26,sp78,cp01,cp02,cp03,cp04,'5' 案件種類,sp27,cp10,'' 類別,SQLDATET(cp06) as 本所期限,SQLDATET(cp07) as 法定期限,CP14,sp14 tm15" & _
      " From caseprogress, servicepractice, nation, CUSTOMER, casepropertymap" & _
      " where cp09 in('" & Replace(pCP09, ",", "','") & "') and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04 and sp01 is not null" & _
      " and na01(+)=sp09 and cu01(+)=substr(sp08,1,8) and cu02(+)=substr(sp08,9)" & _
      " and cp01=cpm01(+) and cp10=cpm02(+)"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      'Add By Sindy 2020/12/1
      '先檢查是否有案件性質不同,若有,要加註案件性質名稱
      bolShowCPMNm = False
      rsQuery.MoveFirst
      strCPMNm = rsQuery.Fields("案件性質") '第一筆
      Do While Not rsQuery.EOF
         If strCPMNm <> "" & rsQuery.Fields("案件性質") Then
            bolShowCPMNm = True
            If InStr(strCPMNm, "" & rsQuery.Fields("案件性質")) = 0 Then
               strCPMNm = strCPMNm & "、" & rsQuery.Fields("案件性質")
            End If
            'Exit Do
         End If
         rsQuery.MoveNext
      Loop
      '2020/12/1 END
      
      With rsQuery
      rsQuery.MoveFirst
      Do While Not .EOF
         strCP14 = "" & .Fields("cp14") 'Add By Sindy 2020/10/13
         strCaseNoData = strCaseNoData & "," & .Fields("本所案號") 'Add By Sindy 2020/7/8
         
         'Add By Sindy 2020/10/14 商標處給代理人的案件資訊
         If bolTBase = True Then
            If Not IsNull(.Fields("pa77")) Then
               stContent = stContent & "貴方卷號：" & .Fields("pa77") & vbCrLf
            End If
            stContent = stContent & "我方案號：" & .Fields("本所案號") & vbCrLf
            stContent = stContent & "申 請 人：" & .Fields("申請人") & vbCrLf
            stContent = stContent & "商 標：" & .Fields("案件名稱") & vbCrLf
            If Not IsNull(.Fields("類別")) Then
               stContent = stContent & "類 別：" & .Fields("類別") & vbCrLf
            End If
            If "" & .Fields("tm15") <> "" Then
               stContent = stContent & "註冊號數：" & "" & .Fields("tm15") & vbCrLf
            Else
               stContent = stContent & "申請案號：" & .Fields("申請案號") & vbCrLf
            End If
            If bolShowCPMNm = True Then
               stContent = stContent & "案件性質：" & .Fields("案件性質") & vbCrLf
            End If
            stContent = stContent & vbCrLf
         Else
         '2020/10/14 END
            
            'Add By Sindy 2020/7/21
            If Not IsNull(.Fields("pa77")) Then
               stContent = stContent & "貴方卷號：" & .Fields("pa77") & vbCrLf
            End If
            '2020/7/21 END
            If Not IsNull(.Fields("客戶案件案號")) Then
               stContent = stContent & "客戶案件案號：" & .Fields("客戶案件案號") & vbCrLf
            End If
            stContent = stContent & "本所案號：" & .Fields("本所案號") & vbCrLf
            
            If .Fields("案件種類") = "3" Then
               stContent = stContent & "相關國家：" & .Fields("申請國家") & vbCrLf
            ElseIf .Fields("案件種類") <> "4" Then
               stContent = stContent & "申請國家：" & .Fields("申請國家") & vbCrLf
            End If
            
            If .Fields("案件種類") = "1" Then
               stContent = stContent & "專利種類：" & .Fields("專利種類") & vbCrLf
            ElseIf .Fields("案件種類") = "2" Then
               'Modify By Sindy 2018/10/12 + 商品類別
               stContent = stContent & "商標種類：" & .Fields("專利種類") & "(第" & .Fields("類別") & "類)" & vbCrLf
            End If
            
            stContent = stContent & "案件性質：" & .Fields("案件性質") & vbCrLf 'pProperty
            'If pOurDeadLine <> "" Then
            If "" & .Fields("本所期限") <> "" Then
               stContent = stContent & "本所期限：" & .Fields("本所期限") & vbCrLf 'pOurDeadLine
            End If
            
            'If pOffDeadLine <> "" Then
            If "" & .Fields("法定期限") <> "" Then
               stContent = stContent & "法定期限：" & .Fields("法定期限") & vbCrLf 'pOffDeadLine
            End If
            
            If .Fields("案件種類") = "3" Or .Fields("案件種類") = "4" Then
               stContent = stContent & "當事人　：" & .Fields("申請人") & vbCrLf
            Else
               stContent = stContent & "申請人　：" & .Fields("申請人") & vbCrLf
            End If
            
            If Not IsNull(.Fields("pa27")) Then
               stContent = stContent & "　　　　（多人申請）" & vbCrLf
            End If
            
            'Modify By Sindy 2018/10/12 申請案號有資料時,才出現此列
            If (.Fields("案件種類") = "1" Or .Fields("案件種類") = "2" Or .Fields("案件種類") = "5") And _
               "" & .Fields("申請案號") <> "" Then
               stContent = stContent & "申請案號：" & .Fields("申請案號") & vbCrLf
            End If
            
            stContent = stContent & "案件名稱：" & .Fields("案件名稱") & vbCrLf & vbCrLf
         End If
         
         rsQuery.MoveNext
      Loop
      If strCaseNoData <> "" Then strCaseNoData = Mid(strCaseNoData, 2) 'Add By Sindy 2020/7/8
      End With
   End If
   
   PUB_GetMailManyCaseData = stContent
   Set rsQuery = Nothing
   
   Screen.MousePointer = vbDefault
End Function

'Added by Morgan 2014/12/18
Public Function PUB_GetCompName(pPA01 As String, pPA02 As String, pPA03 As String, pPA04 As String)
   Dim stCompNo As String
   stCompNo = PUB_GetReceiptComp(pPA01, pPA02, pPA03, pPA04)
   '有設定
   '專利商標
   If stCompNo = "T" Then
      PUB_GetCompName = CompNameQuery("1")
   '智權公司
   ElseIf stCompNo = "J" Then
      PUB_GetCompName = CompNameQuery("J")
   '未設定(專利法律)
   Else
      PUB_GetCompName = CompNameQuery("2")
   End If
End Function

'Added by Morgan 2021/3/30
'E化函確收
Public Function PUB_RecpConfirm(ByVal pLP01 As String, ByRef pCallForm As Form) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stLP39 As String, stLP40 As String
   
   stSQL = "select lp39,lp40,st02,lp46,lp47,sqldatet(lp47)||' '||sqltime6(lp48) dt,lp49 from letterprogress,staff where lp01='" & pLP01 & "' and lp39>0 and st01(+)=lp46"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      If rsQuery("lp47") > 0 Then
         MsgBox "信函已確收！" & vbCrLf & vbCrLf & "確收人員：" & rsQuery("st02") & vbCrLf & "確收時間：" & rsQuery("dt") & vbCrLf & "確收備註：" & rsQuery("lp49"), vbCritical
         PUB_RecpConfirm = True
         Exit Function
      End If
      stLP39 = rsQuery("lp39")
      stLP40 = rsQuery("lp40")
      '顯示寄件備份
      With frm880019
      .Hide
      .Caption = "寄件備份"
      .cmdExit.Caption = "結束"
      .cmdNoSend.Caption = "確收"
      .cmdNoSend.Visible = True
      .cmdSend.Visible = False
      .cmdAttach.Visible = False
      .cmdReceiver(0).Visible = False
      .cmdReceiver(1).Visible = False
      .cmdReceiver(2).Visible = False
      .txtReceiver.Locked = True
      .txtCopy.Locked = True
      .txtBCC.Locked = True
      .txtSubject.Locked = True
      .txtAttachment.Locked = True
      .txtContent.Locked = True
      .FramePrint.Visible = False
      .m_CP09 = pLP01 '總收文號
      .m_SMB02 = stLP39 '寄件日期
      .m_SMB03 = stLP40 '寄件時間
      .SetParent pCallForm
      If .QueryData = True Then
         .Show vbModal
      End If
      PUB_RecpConfirm = .m_bolDone
      End With
      Unload frm880019
   End If
End Function

'Added by Morgan 2021/5/7
'檢查附件是否開啟中
Public Function PUB_ChkAttFile(pAttFiles As String) As Boolean
   Dim arrFile() As String
   Dim ii As Integer
   Dim bolHadShowMsg As Boolean
   
   If pAttFiles <> "" Then
      arrFile = Split(pAttFiles, ";")
      For ii = LBound(arrFile) To UBound(arrFile) - 1
         If PUB_ChkFileOpening(CStr(arrFile(ii)), bolHadShowMsg) = True Then
            If bolHadShowMsg = False Then
               MsgBox arrFile(ii) & vbCrLf & "檔案正在使用中，請關閉才可執行送出！", vbExclamation
            End If
            Exit Function
         End If
      Next
   End If
   PUB_ChkAttFile = True
End Function

'Added by Lydia 2024/03/25 取得TIPS請款金額
Public Function Pub_GetCP144Val(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pKind As String, ByVal pVAL01 As String) As String
Dim strA1 As String, intA As Integer
Dim strCTitle As String
Dim rsA1 As New ADODB.Recordset

   strCTitle = "TIPS請款金額："  'CP144內容=>TIPS請款金額：XXXXXX;
   Pub_GetCP144Val = ""
   
   Select Case pKind
      Case "0" 'CP144內容=>TIPS請款金額：XXXXXX;
         Pub_GetCP144Val = strCTitle & pVAL01 & ";"
      Case "1" '其他階段請款金額
         strA1 = " select c2.cp156,c2.cp144 from caseprogress c1, caseprogress c2" & _
                 " where c1.cp01='" & pCP01 & "' and c1.cp02='" & pCP02 & "' and c1.cp03='" & pCP03 & "' and c1.cp04='" & pCP04 & "' and c1.cp159=0 and c1.cp10 in (" & ACSforTIPSstep & ")" & _
                 " and c1.cp09=c2.cp43(+) and c2.cp159=0 and nvl(c2.cp156,0)>0 and instr(c2.cp144,'" & strCTitle & "') > 0"
         If Len(pVAL01) > 9 And Mid(pVAL01, 1, 1) < "D" Then
            strA1 = strA1 & " and c2.cp09 not in (" & GetAddStr(pVAL01) & ") "
         End If
         intA = 1
         Set rsA1 = ClsLawReadRstMsg(intA, strA1)
         If intA = 1 Then
            rsA1.MoveFirst
            Do While Not rsA1.EOF
               If InStr("" & rsA1.Fields("cp144"), strCTitle) > 0 Then
                  strA1 = Mid("" & rsA1.Fields("cp144"), InStr("" & rsA1.Fields("cp144"), strCTitle) + Len(strCTitle))
                  If InStr(strA1, ";") > 0 Then
                     strA1 = Mid(strA1, 1, InStr(strA1, ";") - 1)
                  End If
                  Pub_GetCP144Val = Val(Pub_GetCP144Val) + Val(strA1)
               End If
               rsA1.MoveNext
            Loop
         End If
      Case "2" 'CP144>>取得金額
         If InStr(pVAL01, strCTitle) > 0 Then
            strA1 = Mid(pVAL01, InStr(pVAL01, strCTitle) + Len(strCTitle))
            Pub_GetCP144Val = Val(Mid(strA1, 1, InStr(strA1, ";") - 1))
         Else
            Pub_GetCP144Val = pVAL01
         End If
   End Select
   Set rsA1 = Nothing
End Function


