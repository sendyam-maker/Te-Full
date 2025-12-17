Attribute VB_Name = "Service1"
'Memo By Sindy 2012/12/5 ´¼Åv¤H­ûÄæ¤w­×§ï
'Memo By Sindy 2011/2/15 SQLDate¤wÀË¬d
'Memo By Sindy 2010/8/4 ¤é´ÁÄæ¤w­×§ï
Option Explicit

'Modify By Sindy 2023/1/31
'¦¬¤å¦³´Á­­ªº®×¥ó¡]P¡^¡A­Y¦P®É¦¬¡]938¡^¡]939¡^¡]¶W­¶¡A¶W¶µ¶O¡^®É¡A¦¹¨â¹D¤£±¾´Á­­¡C¡]­ì¬°¤H¤u§R°£´Á­­¡^
Public Const P®×¤£±¾´Á­­ªº®×¥ó©Ê½è = "938,939"
'¦¬¤å¦³´Á­­ªº®×¥ó¡]T¡^¡A­Y¦P®É¦¬¡]501¡^¡]²¾Âà¡^®É¡A¦¹¹D¤£±¾´Á­­¡C¡]­ì¬°¤H¤u§R°£´Á­­¡^
Public Const T®×¤£±¾´Á­­ªº®×¥ó©Ê½è = "501"

'Added by Lydia 2022/09/05
Dim intJ As Integer, intK As Integer   '¨ú¥N¦@¥ÎÅÜ¼Æ intI ©Î i ¤§Ãþ
Dim mStrSql As String  '¨ú¥N¦@¥ÎÅÜ¼ÆstrSql
Dim strTmp1(0 To 10) As String  '¨ú¥N¦@¥ÎÅÜ¼ÆstrExc(0 to 10)


'³B²z¸ê®Æ°Ï¶¡°ÝÃD ex. Time,Country....
Public Function Process_txtRange(TxtRange1 As Control, TxtRange2 As Control, strMsg As String, fldname As String) As String
Dim strTemp As String

   If Trim$(TxtRange1.Text) = "" And Trim$(TxtRange2.Text) = "" Then
      strTemp = ""
   Else
      If Trim$(TxtRange1.Text) = "" Then
         strTemp = " and " + fldname + " < " And "'" And Trim$(TxtRange2.Text)
      Else
         If Trim$(TxtRange2.Text) Then
            strTemp = " and " + fldname + " > " And "'" And Trim$(TxtRange1.Text) + "' and "
         Else
            If Val(TxtRange1.Text) < Val(TxtRange2.Text) Then
               Call ShowMsg(strMsg)
               Exit Function
            End If
            strTemp = " and " + fldname + " between '" + Trim$(TxtRange1.Text) + _
                      "' and '" + Trim$(TxtRange2.Text) + "'"
         End If
      End If
   End If
   Process_txtRange = strTemp
End Function

Public Sub Clear_AllTxtAry(TxtCtrl As Variant, TxtNum1 As Integer, TxtNum2 As Integer)
Dim i As Integer

    For i = TxtNum1 To TxtNum2
       TxtCtrl(i).Text = ""
    Next
End Sub

Public Sub ShowDetail(TxtCtrl As Variant, TxtNum1 As Integer, TxtNum2 As Integer, Rss As ADODB.Recordset)
Dim i As Integer

    For i = TxtNum1 To TxtNum2
       If i <> 1 Then
          TxtCtrl(i).Text = IIf(IsNull(Rss(i)), "", Rss(i))
       Else
          TxtCtrl(i).Text = IIf(Rss(i) = 0, "", Rss(i))
       End If
    Next
End Sub

Public Sub OnOff_Button(TlBarCtrl As Control, ButtonValue As Boolean)
Dim i As Integer
   For i = 1 To 4
      TlBarCtrl.Buttons.Item(i).Enabled = ButtonValue
   Next
   For i = 5 To 9
      TlBarCtrl.Buttons.Item(i).Enabled = ButtonValue
   Next
   If ButtonValue = True Then
      TlBarCtrl.Buttons.Item(11).Enabled = False
      TlBarCtrl.Buttons.Item(12).Enabled = False
   Else
      TlBarCtrl.Buttons.Item(11).Enabled = True
      TlBarCtrl.Buttons.Item(12).Enabled = True
   End If
   TlBarCtrl.Buttons.Item(14).Enabled = ButtonValue
End Sub
Public Sub MessageShow(strMsg As String, Optional strLabel As String)
   MsgBox strLabel & strMsg, , MsgText(9001)
End Sub
'*************************************************
' ¨Ì¦¬¤å¸¹¨ú¦^¥»©Ò®×¸¹
'
'*************************************************
Public Function GetCaseNo(strDocNo As String) As String
Dim Rs As New ADODB.Recordset
   strSql = "select * from caseprogress where CP09 = '" & strDocNo & "'"
   Set Rs = ClsPDReadRst(strSql)
   If Rs.RecordCount <> 0 Then
      GetCaseNo = Rs(0) & Rs(1) & Rs(2) & Rs(3)
   Else
      MessageShow "", MsgText(307)
      GetCaseNo = ""
   End If
   Rs.Close
End Function
'±N¦r¦êSize=6 Âà´«¦¨ size=8 or 9(+"00")
Public Sub StrProcessingA(ByRef strNum As String, Optional Size As String = "8", Optional strType As String = "O")
    If strType = "I" Then
        Select Case Len(strNum)
        Case 6
            Select Case Size
            Case "8"
                strNum = strNum + "00"
            Case "9"
                strNum = strNum + "000"
            End Select
        End Select
    Else
        Select Case Len(strNum)
        Case 8
            If Mid(strNum, 7, 2) = "00" Then
                strNum = Mid(strNum, 1, 6)
            End If
        Case 9
            If Mid(strNum, 7, 3) = "000" Then
                strNum = Mid(strNum, 1, 6)
            End If
        End Select
    End If
End Sub
'¨Ì°ê¤º¥~§O¨ÓÀË¬d¸ê®Æ
Public Function ChkDateByWhere(strDate As String, Optional intWhere As Integer = 0) As Boolean
    Select Case intWhere
    Case 1
        ChkDateByWhere = CheckIsDate(strDate)
    Case Else
        ChkDateByWhere = CheckIsTaiwanDate(strDate)
    End Select
End Function
Public Function GetNowDateByWhere(intWhere As Integer) As String
    Select Case intWhere
    Case 0
            GetNowDateByWhere = GetTaiwanTodayDate
    Case 1
            GetNowDateByWhere = GetTodayDate
    End Select
End Function
Public Function AdjustCustomer(strTemp As String) As String
        If Right(strTemp, 2) = "00" And Len(strTemp) = 8 Then
        AdjustCustomer = Mid$(strTemp, 1, 6)
        End If
End Function
'move to basquery by nickc 2007/02/07
''¨ú±o®É¶¡
'Public Function GetTime() As String
'Dim strTime As String, strmin As String
'strTime = time
'If Mid(strTime, 1, 2) = "PM" And Mid(strTime, 4, 2) <> "00" Then
'   strmin = str(Val(Mid(strTime, 4, 2)) + 12)
'   GetTime = strmin + Mid(strTime, 7, 2)
'Else
'   GetTime = Mid(strTime, 4, 2) + Mid(strTime, 7, 2)
'End If
'End Function

Public Function ChkTime(ByRef strTime As String) As Boolean
If Val(Left(strTime, 2)) > 24 Or Val(Right(strTime, 2)) > 60 Then
    DataErrorMessage 1, "®É¶¡"
    ChkTime = False
Else
   ChkTime = True
End If
End Function
'ÀË¬d¤é´Á½d³ò
Public Function ChkDateRange(Compo1 As Control, Compo2 As Control) As Boolean
    If Val(Compo1) > Val(Compo2) Then
            ShowMsg MsgText(9016)
            ChkDateRange = False
    Else
            ChkDateRange = True
    End If
End Function
'¶Ç¦^ªøªº«È¤á¥N¸¹
Public Function SPChangeCustomerL(ByRef strTemp As String, Optional AcType As Integer = 0) As String
Select Case AcType
Case 0
    If strTemp <> "" Then SPChangeCustomerL = strTemp + String(9 - Len(strTemp), "0")
Case 1
    If strTemp <> "" Then SPChangeCustomerL = strTemp + String(8 - Len(strTemp), "0")
End Select
End Function
'¶Ç¦^µuªº«È¤á¥N¸¹
Public Function SPChangeCustomerS(ByRef strTemp As String, Optional AcType As Integer = 0) As String
Select Case AcType
Case 0
    If strTemp <> "" Then SPChangeCustomerS = IIf(Right(strTemp, 3) = "000", Mid(strTemp, 1, 6), IIf(Right(strTemp, 1) = "0", Mid(strTemp, 1, 8), strTemp))
Case 1
    If strTemp <> "" Then SPChangeCustomerS = IIf(Right(strTemp, 2) = "00", Mid(strTemp, 1, 6), strTemp)
End Select
End Function
Public Function GetCustomer1(ByRef strAgent As String, ByRef strAgentName As String) As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset, i As Integer, strTemp
   strAgent = ChangeCustomerL(strAgent)
   strSql = "select cu01||cu02,nvl(fa05,nvl(fa04,fa06)),nvl(fa15,nvl(fa16||' '||fa17||' '||fa18||' '||fa19||' '||fa20,fa21)),fa01||fa02 from customer,fagent where fa01=cu03 and " + ChgCustomer(strAgent) + " order by fa01||fa02"
   Set rsRecordset = ClsPDReadRst(strSql)
   If rsRecordset.EOF Then
    ShowMsg MsgText(9051)
    rsRecordset.Close
    GetCustomer1 = False
    Exit Function
   Else
    strAgent = ChangeCustomerS(strAgent)
    strAgentName = IIf(IsNull(rsRecordset(1)), "", rsRecordset(1))
    rsRecordset.Close
    GetCustomer1 = True
    Exit Function
   End If
End Function
Public Function ChkList(strCon1 As String, strCon2 As String) As Boolean
Dim Rss As ADODB.Recordset
    strSql = "select * from UserMenu where UM01=" + CNULL(strCon1) + " and UM02=" + CNULL(strCon2) + " and UM03=" + CNULL(strGroup)
    Set Rss = ClsPDReadRst(strSql)
    If Rss.EOF Then
        Rss.Close
        ChkList = False
        Exit Function
    Else
        Rss.Close
        ChkList = True
    End If
End Function

'¨ú±o¥N²z¤H¦WºÙ
Public Function GetAgent(ByRef strAgent As String, ByRef strAgentName As String) As Boolean
Dim strSql As String, rsRecordset As New ADODB.Recordset, i As Integer, strTemp

On Error GoTo ErrHand
strAgent = ChangeCustomerL(strAgent)
strSql = "select fa01||fa02,nvl(fa05||fa63||fa64||fa65,nvl(fa04,fa06)) from fagent where " + ChgFagent(strAgent) + " order by fa01||fa02"
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSql, cnnConnection
If rsRecordset.RecordCount > 0 Then
   strAgentName = rsRecordset(1)
   GetAgent = True
Else
   GetAgent = False
End If
rsRecordset.Close
Exit Function
ErrHand:
    GetAgent = False
End Function

'Added by Lydia 2022/09/05 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G·s¼W°Ó¼Ð°ò¥»ÀÉ(±qfrm010004.InsertTrademarkDatabase©â¥X¨Ó)
'Modify By Sindy 2023/5/31 + , Optional ByRef RetVal As String
Private Function InsertTrademarkDB(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intChoose As Integer, _
                ByRef mTM() As String, ByRef mCP() As String, ByVal mCU30 As String, ByVal mSaveControl As String, Optional ByRef IsSaveData As Boolean, _
                Optional ByVal pType As String, Optional ByVal pCaseNo As String, Optional ByRef RetVal As String) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
'mSaveControl: »ô³Æ¤éºÞ¨î
Dim strAutoNumber As String
Dim np13 As String, np14 As String, bolRt As Boolean
Dim np14ForCP41 As String, np14ForCP42 As String 'Add By Sindy 2025/1/24
Dim bolError As Boolean, intW As Integer
Dim adoquery As New ADODB.Recordset
Dim strBKindCP09 As String 'BÃþ¦¬¤å¸¹
Dim strApply As Variant, strAllApp As String 'Add by Amy 2017/03/09
Dim strCP08 As String, strCP43 As String 'Add by Sindy 2022/9/28
Dim m_CaseNaTmp() As String  '¯S®íºÞ¨î¤§ÃöÁp®×
'ªk«ß©Ò®×·½¦¬¤å
Dim m_LOS02 As String '®×·½®×¥óÃþ«¬
Dim m_LOS15 As String '®×·½³æ¸¹
Dim rsQD As New ADODB.Recordset
'Add By Sindy 2023/5/31
Dim m_bMRecvBatch As Boolean '«H¥ó¨R¾P¦h®×¦¬¤å
Dim m_bolRecvOK As Boolean '¬O§_¦¬§¹¤å
Dim m_strMCR11 As String '¦h®×¦¬¤å®É,²Ä¤@µ§ªºÁ`¦¬¤å¸¹
Dim m_strIR01 As String, m_strIR02 As String, m_strIR03 As String, m_strIR04 As String '«H¥ó¨R¾PPK
'2023/5/31 END

   If IsSaveData = True Then
       Exit Function
   End If
   IsSaveData = True
   
'*********¯S®íºÞ¨îªºÅÜ¼Æ*************
   If pType = "CFT­^°ê²æ¼Ú®×" And pCaseNo <> "" Then
       ReDim m_CaseNaTmp(1 To TF_TM)
       Call ChgCaseNo(pCaseNo, m_CaseNaTmp)
   'Modify By Sindy 2025/8/18 µo¥Í¤F®×·½+«H¥ó¨R¾P ex:FCP-057445/FCL-011034
   ' mark,§ï¦b¤U¦C¥t¥~¼gif
   'Add By Sindy 2023/5/31
'   ElseIf InStr(pType, "«H¥ó¨R¾P") > 0 And pCaseNo <> "" Then
'       '¦]¬°«H¥ó¨R¾P¬O±q¨t²Î¦¬¥ó°Ï¡A©Ò¥H¤£·|¸ò®×·½¦¬¤å(frm090801)­«Å|
'       m_CaseNaTmp = Split(pCaseNo, ",")
'       m_strIR01 = m_CaseNaTmp(0)
'       m_strIR02 = m_CaseNaTmp(1)
'       m_strIR03 = m_CaseNaTmp(2)
'       m_strIR04 = m_CaseNaTmp(3)
'       ReDim m_CaseNaTmp(1 To 4)  '¹w³]°}¦CÁ×§Kµ{¦¡¥X¿ù
'       If InStr(pType, "¦h®×¦¬¤å") > 0 Then m_bMRecvBatch = True
   Else
       ReDim m_CaseNaTmp(1 To 4) '¹w³]°}¦CÁ×§Kµ{¦¡¥X¿ù
       'Modify By Sindy 2025/8/18
       'If pType = "LOS®×·½¦¬¤å" And pCaseNo <> "" Then
       If InStr(pType, "LOS®×·½¦¬¤å") > 0 And pCaseNo <> "" Then
       '2025/8/18 END
           m_LOS02 = Mid(pCaseNo, 1, InStr(pCaseNo, ",") - 1) '®×·½®×¥óÃþ«¬
           m_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1, 8) '®×·½³æ¸¹ 'Modify By Sindy 2025/8/18 +, 8)
       ElseIf pType = "CFT½q¨l­«·s¥Ó½Ð®×" And pCaseNo <> "" Then
           Call ChgCaseNo(pCaseNo, m_CaseNaTmp)
       End If
   End If
   'Modify By Sindy 2025/8/18
   If InStr(pType, "«H¥ó¨R¾P") > 0 And pCaseNo <> "" Then
      m_CaseNaTmp = Split(pCaseNo, "-")
      If InStr(pCaseNo, "-") > 0 Then
         strExc(10) = m_CaseNaTmp(1)
      Else
         strExc(10) = m_CaseNaTmp(0)
      End If
      m_CaseNaTmp = Split(strExc(10), ",")
   '2025/8/18 END
      m_strIR01 = m_CaseNaTmp(0)
      m_strIR02 = m_CaseNaTmp(1)
      m_strIR03 = m_CaseNaTmp(2)
      m_strIR04 = m_CaseNaTmp(3)
      'ReDim m_CaseNaTmp(1 To 4)  '¹w³]°}¦CÁ×§Kµ{¦¡¥X¿ù
      If InStr(pType, "¦h®×¦¬¤å") > 0 Then m_bMRecvBatch = True
      If InStr(pType, "LOS®×·½¦¬¤å") > 0 Then pType = "LOS®×·½¦¬¤å" 'Add By Sindy 2025/8/18
   End If
'***********************************

   On Error GoTo ErrHand
   '¶Ç¤J0¬°­«½Æ¤§¥»©Ò®×¸¹(·s¼WÂÂ®×)¡A1¬°¥¿½T¤§¥»©Ò®×¸¹(·s¼W·s®×)
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.BeginTrans
   End If
   If intSaveMode = 1 Then
      Cls001SetTMFileProperty mCP(10), mTM(28)
      If mTM(2) = "" Or mTM(2) = "0" Then
         If ClsPDGetAutoNumber(mTM(1), strAutoNumber, True, False) Then
            mTM(2) = strAutoNumber
            'Added by Lydia 2025/08/21 T¥xÆW·s®×¦¬¤å707 ½Õ¬d®É¡A½Ð³]©w¨÷©v©Ê½è¬°:4(¼o¤î)
            If mTM(1) = "T" And mTM(10) = "000" And mCP(10) = "707" Then
               mTM(28) = "4"
            End If
            'end 2025/08/21
         Else
            bolError = True
         End If
      End If
      If bolError = False Then

         If ClsPDGetSystemKind(mTM(1), , , intW) Then
            mTM(53) = IIf(intW = 2, 2, 1)
            mCP(2) = mTM(2)
            'Modify by Morgan 2008/8/5 +TM123
            'Add By Sindy 2012/7/19 ­Y¥Ó½Ð¤H©Î¥N²z¤H¬°¿ÕµØ¤½¥qªÌ¡A®×¥ó³Æµù­YµL"¤£¾P¨÷"¦r¼Ë,«h­n¥[¤J
            If (mTM(23) <> "" And InStr(strTmNovartisCust, Left(mTM(23), 6)) > 0) Or _
               (mTM(78) <> "" And InStr(strTmNovartisCust, Left(mTM(78), 6)) > 0) Or _
               (mTM(79) <> "" And InStr(strTmNovartisCust, Left(mTM(79), 6)) > 0) Or _
               (mTM(80) <> "" And InStr(strTmNovartisCust, Left(mTM(80), 6)) > 0) Or _
               (mTM(81) <> "" And InStr(strTmNovartisCust, Left(mTM(81), 6)) > 0) Or _
               (mTM(44) <> "" And InStr(strTmNovartisCust, Left(mTM(44), 6)) > 0) Then
               mTM(58) = ChangeTStringToTDateString(strSrvDate(2)) & "¤£¾P¨÷"
            End If
            '2012/7/19 end
            'ADD BY SONIA 2015/11/24 ­Y¥Ó½Ð¬°®Èª°°ê»Ú¤Î³¡¤ÀÃö«Y¥ø·~ªÌ¡A®×¥ó³Æµù­YµL"¤£¾P¨÷"¦r¼Ë,«h­n¥[¤J
            'mTM(58) = "" 'Memo by Lydia 2022/08/11 ­ìCode¦³°ÝÃD
            If (mTM(23) <> "" And InStr(strTmTRAVEL_FOXCust, Left(mTM(23), 8)) > 0) Or _
               (mTM(78) <> "" And InStr(strTmTRAVEL_FOXCust, Left(mTM(78), 8)) > 0) Or _
               (mTM(79) <> "" And InStr(strTmTRAVEL_FOXCust, Left(mTM(79), 8)) > 0) Or _
               (mTM(80) <> "" And InStr(strTmTRAVEL_FOXCust, Left(mTM(80), 8)) > 0) Or _
               (mTM(81) <> "" And InStr(strTmTRAVEL_FOXCust, Left(mTM(81), 8)) > 0) Then
               mTM(58) = ChangeTStringToTDateString(strSrvDate(2)) & "¤£¾P¨÷"
            End If
            'END 2015/11/24
            'Modify by Sindy 2012/7/19 +TM58
            'MODIFY BY SONIA 2014/8/12 +TM16(©µ®i·s®×­n¦stm16,§_«h­«ÂÐ¦¬¤åÀË¬d¤£¥X¨Ó)
            mTM(130) = GetReceiptCmp(Left(mTM(23), 8), Mid(mTM(23), 9, 1), mTM(1), mTM(10)) 'Add by Amy 2018/10/11 +¦¬¾Ú¤½¥q§Otm130
            'Added by Lydia 2020/11/19 CFT­^°ê²æ¼Ú®×ºÞ¨î¡G·s¼W­^°ê®×®É¦P®É§â¼Ú·ù®×¬ÛÃöÄæ¦ì±a¹L¨Ó(°Ñ¦ÒPUB_SaveCountry)
            'Modified by Lydia 2021/03/05 §PÂ_«D¥Ó½Ð®×¡FCFT¼Ú·ù©|¥¼µù¥U®×Âà´«­^°ê¥Ó½Ð®×¦¬¤å±±ºÞ
            If mTM(1) = "CFT" And pType = "CFT­^°ê²æ¼Ú®×" And m_CaseNaTmp(1) <> "" And m_CaseNaTmp(2) <> "" And mCP(10) <> "101" Then
                If PUB_ReadTradeMarkData(m_CaseNaTmp(), m_CaseNaTmp(1), m_CaseNaTmp(2), m_CaseNaTmp(3), m_CaseNaTmp(4)) Then
                   strTmp1(0) = "": strTmp1(1) = ""
                   'Added by Lydia 2020/12/09 ±M¥Î´Á¶¡(¤î)=¼Ú·ù®×¤§ªk­­; Á×§K¼Ú·ù®×¥ý¦¬©µ®i,§ó·s°Ó¼Ð°ò¥»ÀÉªº±M¥Î´Á¶¡(¤î)
                   strTmp1(9) = ""
                   strTmp1(2) = "select np09 from nextprogress where np02='" & m_CaseNaTmp(1) & "' and np03='" & m_CaseNaTmp(2) & "' and np04='" & m_CaseNaTmp(3) & "' and np05='" & m_CaseNaTmp(4) & "' and np07=" & CNULL(IIf(mCP(10) = "102", "110", mCP(10))) & " and np06 is null "
                   intK = 1
                   Set rsQD = ClsLawReadRstMsg(intK, strTmp1(2))
                   If intK = 1 Then
                       strTmp1(9) = "" & rsQD.Fields("np09")
                   End If
                   'end 2020/12/09
                   For intJ = 5 To TF_TM
                       Select Case intJ
                          Case 59, 60, 61, 62, 63, 64, 57, 73, 74, 75 'Create + Update, ¾P¨÷¤é
                          Case 10 '¥Ó½Ð°ê®a
                              strTmp1(0) = strTmp1(0) & "TM" & Format(intJ, "00") & "," 'Insert
                              strTmp1(1) = strTmp1(1) & " '201' as TM10, " 'Select
                          Case 12, 15 '¥Ó½Ð¸¹¡B¼f©w¸¹¡GUK009+¼Ú·ù¸¹¼Æ«á8½X(®³±¼²Ä1½X0)
                              strTmp1(0) = strTmp1(0) & "TM" & Format(intJ, "00") & ","
                              strTmp1(1) = strTmp1(1) & " 'UK009'||substr(tm15,2,8) AS TM" & Format(intJ, "00") & ","
                          'Added by Lydia 2020/12/09 ±M¥Î´Á¶¡(¤î)=¼Ú·ù®×¤§ªk­­; Á×§K¼Ú·ù®×¥ý¦¬©µ®i,§ó·s°Ó¼Ð°ò¥»ÀÉªº±M¥Î´Á¶¡(¤î)
                          Case 22    '±M¥Î´Á¶¡(¤î¤é)
                              strTmp1(0) = strTmp1(0) & "TM" & Format(intJ, "00") & ","
                              strTmp1(1) = strTmp1(1) & " " & IIf(strTmp1(9) <> "", CNULL(strTmp1(9)), "TM22") & " as TM22, "
                          'end 2020/12/09
                          Case 58  '®×¥ó³Æµù: ¥[µù¼Ú·ù®×®×¸¹
                              strTmp1(0) = strTmp1(0) & "TM" & Format(intJ, "00") & ","
                              strTmp1(1) = strTmp1(1) & CNULL("¼Ú·ù®×®×¸¹¡G" & m_CaseNaTmp(1) & m_CaseNaTmp(2) & m_CaseNaTmp(3) & m_CaseNaTmp(4) & ";") & "||TM" & Format(intJ, "00") & " AS TM" & Format(intJ, "00") & ","
                          Case Else
                              strTmp1(0) = strTmp1(0) & "TM" & Format(intJ, "00") & ","
                              strTmp1(1) = strTmp1(1) & "TM" & Format(intJ, "00") & ","
                       End Select
                   Next
                   strTmp1(0) = Left(strTmp1(0), Len(strTmp1(0)) - 1)
                   strTmp1(1) = Left(strTmp1(1), Len(strTmp1(1)) - 1)
                   mStrSql = "INSERT INTO TRADEMARK (TM01,TM02,TM03,TM04," & strTmp1(0) & ") " & _
                               "SELECT '" & mTM(1) & "' as TM01,'" & mTM(2) & "' as TM02,'" & mTM(3) & "' as TM03,'" & mTM(4) & "' as TM04, " & strTmp1(1) & _
                               " FROM TRADEMARK WHERE TM01='" & m_CaseNaTmp(1) & "' and TM02='" & m_CaseNaTmp(2) & "' and TM03='" & m_CaseNaTmp(3) & "' and TM04='" & m_CaseNaTmp(4) & "' "
                   cnnConnection.Execute mStrSql
                   'Added by Lydia 2021/01/08 ½Æ»s«ü©w°Ó«~©ÎªA°È¦WºÙ
                   mStrSql = "INSERT INTO TMGOODS(TG01,TG02,TG03,TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17) " & _
                               "SELECT '" & mTM(1) & "' as TG01, '" & mTM(2) & "' as TG02, '" & mTM(3) & "' as  TG03, '" & mTM(4) & "' as  TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17 " & _
                               "FROM TMGOODS WHERE TG01='" & m_CaseNaTmp(1) & "'  AND TG02='" & m_CaseNaTmp(2) & "' AND TG03='" & m_CaseNaTmp(3) & "'  AND TG04='" & m_CaseNaTmp(4) & "'  AND TG18 IS NULL "
                   cnnConnection.Execute mStrSql
                   'end 2021/01/08
                End If
                
            Else
            'end 2020/11/19
                'Modify by Amy 2018/10/11 +¦¬¾Ú¤½¥q§Otm130
                'Modified by Lydia 2022/08/11 +(¨Ö¤J) tm34, tm35
                'Modify By Sindy 2022/12/7 + tm136
                'Modify By Sindy 2022/12/7 + tm72
                'Modify By Sindy 2024/2/2 + tm12 ¥Ó½Ð®×¸¹
                'Modify By Sindy 2024/7/9 IIf(mTM(15) = "", "NULL", "1") => IIf(mTM(15) <> "" and mTM(28) = "1", "1", "NULL")
                mStrSql = "insert into trademark (tm01, tm02, tm03, tm04, tm05, tm08, tm09, tm10," + _
                  "tm23,tm28,tm34,tm44,tm17,TM15,tm78,tm79,tm80,tm81,tm32,tm45,tm123,tm58,tm16,tm130,tm35,tm136,tm72,tm12) " & _
                  "values (" + CNULL(mTM(1)) + "," + CNULL(mTM(2)) + "," + CNULL(mTM(3)) + "," + CNULL(mTM(4)) + "," + CNULL(ChgSQL(mTM(5))) + "," + _
                  CNULL(mTM(8)) + "," + CNULL(mTM(9)) + "," + CNULL(mTM(10)) + "," + CNULL(mTM(23)) + "," + CNULL(mTM(28)) + "," + CNULL(mTM(34)) + "," + CNULL(mTM(44)) + ", ''" + "," + CNULL(ChgSQL(mTM(15))) + "," + CNULL(mTM(78)) + "," + _
                  CNULL(mTM(79)) + "," + CNULL(mTM(80)) + "," + CNULL(mTM(81)) + "," + CNULL(mTM(32)) + "," + CNULL(ChgSQL(mTM(45))) + "," + CNULL(mTM(123)) + _
                  "," + CNULL(mTM(58)) + "," + IIf(mTM(15) <> "" And mTM(28) = "1", "1", "NULL") + "," + CNULL(mTM(130)) + "," + CNULL(mTM(35)) + "," + CNULL(mTM(136)) + "," + CNULL(mTM(72)) + "," + CNULL(ChgSQL(mTM(12))) + ")"
                'end 2018/10/11
                 cnnConnection.Execute mStrSql
                'Memo by Lydia 2020/11/19 CFT­^°ê²æ¼Ú®×ºÞ¨î¡G·s¼W­^°ê®×®É¦P®É§â¼Ú·ù®×¬ÛÃöÄæ¦ì±a¹L¨Ó¡A©Ò¥H¤£­nÅÜ§ó¸ê®Æ
                 mStrSql = "update trademark set tm24=(select nvl(cu112,'')||cu23 from customer where cu01=" + CNULL(Mid(mTM(23), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(23), 9, 1)) + _
                    "),tm25=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(mTM(23), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(23), 9, 1)) + _
                    "),tm26=(select cu29 from customer where cu01=" + CNULL(Mid(mTM(23), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(23), 9, 1)) + ") " + _
                    "where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
                 cnnConnection.Execute mStrSql
                'add by nickc 2006/11/30
                mStrSql = "update trademark set tm82=(select cu23 from customer where cu01=" + CNULL(Mid(mTM(78), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(78), 9, 1)) + _
                   "),tm86=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(mTM(78), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(78), 9, 1)) + _
                   "),tm90=(select cu29 from customer where cu01=" + CNULL(Mid(mTM(78), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(78), 9, 1)) + ") " + _
                   "where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
                cnnConnection.Execute mStrSql
                mStrSql = "update trademark set tm83=(select cu23 from customer where cu01=" + CNULL(Mid(mTM(79), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(79), 9, 1)) + _
                   "),tm87=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(mTM(79), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(79), 9, 1)) + _
                   "),tm91=(select cu29 from customer where cu01=" + CNULL(Mid(mTM(79), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(79), 9, 1)) + ") " + _
                   "where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
                cnnConnection.Execute mStrSql
                mStrSql = "update trademark set tm84=(select cu23 from customer where cu01=" + CNULL(Mid(mTM(80), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(80), 9, 1)) + _
                   "),tm88=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(mTM(80), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(80), 9, 1)) + _
                   "),tm92=(select cu29 from customer where cu01=" + CNULL(Mid(mTM(80), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(80), 9, 1)) + ") " + _
                   "where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
                cnnConnection.Execute mStrSql
                mStrSql = "update trademark set tm85=(select cu23 from customer where cu01=" + CNULL(Mid(mTM(81), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(81), 9, 1)) + _
                   "),tm89=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(mTM(81), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(81), 9, 1)) + _
                   "),tm93=(select cu29 from customer where cu01=" + CNULL(Mid(mTM(81), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(81), 9, 1)) + ") " + _
                   "where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
                cnnConnection.Execute mStrSql
            End If 'Added by Lydia 2020/11/19
             
            'Added by Lydia 2021/03/05 CFT¼Ú·ù©|¥¼µù¥U®×Âà´«­^°ê¥Ó½Ð®×¦¬¤å±±ºÞ¡G¤@¨Ö±N¼Ú·ùÃöÁp®×¤§¡u°Ó¼Ð¹Ï¼Ë¡v¡B¡u°Ó«~/ªA°ÈÃþ§O¤Î¦WºÙ¡v¡B¡u¥Ó½Ð¤é¡v¡B¡uÀu¥ýÅv¸ê®Æ¡v±a¤J·s®×¸¹
            If mTM(1) = "CFT" And pType = "CFT­^°ê²æ¼Ú®×" And m_CaseNaTmp(1) <> "" And m_CaseNaTmp(2) <> "" And mCP(10) = "101" Then
                mStrSql = "update trademark set (tm09,tm11) = (select tm09,tm11 from trademark where TM01='" & m_CaseNaTmp(1) & "' and TM02='" & m_CaseNaTmp(2) & "' and TM03='" & m_CaseNaTmp(3) & "' and TM04='" & m_CaseNaTmp(4) & "' " & _
                             ") where TM01='" & mTM(1) & "' and TM02='" & mTM(2) & "' and TM03='" & mTM(3) & "' and TM04='" & mTM(4) & "' "
                cnnConnection.Execute mStrSql
                '½Æ»s«ü©w°Ó«~©ÎªA°È¦WºÙ
                mStrSql = "INSERT INTO TMGOODS(TG01,TG02,TG03,TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17) " & _
                            "SELECT '" & mTM(1) & "' as TG01, '" & mTM(2) & "' as TG02, '" & mTM(3) & "' as  TG03, '" & mTM(4) & "' as  TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17 " & _
                            "FROM TMGOODS WHERE TG01='" & m_CaseNaTmp(1) & "'  AND TG02='" & m_CaseNaTmp(2) & "' AND TG03='" & m_CaseNaTmp(3) & "'  AND TG04='" & m_CaseNaTmp(4) & "'  AND TG18 IS NULL "
                cnnConnection.Execute mStrSql
            End If
            'end 2021/03/05
            mCP(31) = "Y"
         Else
            bolError = True
         End If
      End If
      'Add by Amy 2017/01/03 MCTF±±ºÞ T¦rÀY·s®×¥B¦³¿éFC¥N²z¤H¥B¦¬¤å·~°È°Ï¬° P2¦rÀY¥B¥Ó½Ð°ê®a¬O¥xÆW®É,¨ÌFC¥N²z¤H¤§ºÞ±±´¼Åv¤H­û§ó·s«È¤áÀÉ¤§CU12,CU13
      'Memo by Amy 2017/03/22 ­×§ïUpdMCTF_Cu13 ®³±¼¥Ó½Ð°ê®a¬O¥xÆWªº§PÂ_
      If Len(Trim(mTM(44))) > 0 Then
        For intJ = 0 To 4
            If intJ = 0 Then
                If Len(Trim(mTM(23))) = 0 Then Exit For
                strAllApp = strAllApp & "," & mTM(23)
            ElseIf Len(Trim(mTM(78 + intJ))) = 0 Then
               Exit For
            Else
                strAllApp = strAllApp & "," & mTM(78 + intJ)
            End If
        Next intJ
        If strAllApp <> MsgText(601) Then
            strApply = Split(Mid(strAllApp, 2), ",")
            If UpdMCTF_Cu13(mTM(1), ChangeCustomerL(mTM(44)), mTM(10), strApply, PUB_GetST03(mCP(13))) = False Then
                  GoTo ErrHand
            End If
        End If
      End If
      'end 2017/03/09
      
      'Add By Sindy 2024/3/15 ¹q¤l¦¬¤å¤À³Î®×
      If UCase(pFormName) = UCase("frm090801_New") Then
         If mCP(10) = "308" Then
            '¨ú±o¥À¸¹
Dim str308TM01 As String
Dim str308TM02 As String
Dim str308TM03 As String
Dim str308TM04 As String
            mStrSql = "select TM01,TM02,TM03,TM04,TM08,TM09 from ConsultRecordList,trademark" & _
               " where CRL01='" & mCP(140) & "' and CRL07 is not null" & _
                 " and CRL07=TM01 and CRL08=TM02 and CRL09=TM03 and CRL10=TM04"
            intJ = 1
            Set rsQD = ClsLawReadRstMsg(intJ, mStrSql)
            If intJ = 1 Then
               str308TM01 = rsQD.Fields("TM01")
               str308TM02 = rsQD.Fields("TM02")
               str308TM03 = rsQD.Fields("TM03")
               str308TM04 = rsQD.Fields("TM04")
               '¥DÀÉ
               mStrSql = "update trademark set" & _
                         " TM08='" & rsQD.Fields("TM08") & "',TM09='" & rsQD.Fields("TM09") & "'" & _
                         " where TM01='" & mTM(1) & "' and TM02='" & mTM(2) & "'" & _
                           " and TM03='" & mTM(3) & "' and TM04='" & mTM(4) & "'"
               cnnConnection.Execute mStrSql, intJ
               '¥Nªí¹Ï
               mStrSql = "select * from ImgByteFile" & _
                        " where IBF01='" & str308TM01 & "'" & _
                          " and IBF02='" & str308TM02 & "'" & _
                          " and IBF03='" & str308TM03 & "'" & _
                          " and IBF04='" & str308TM04 & "'"
               intJ = 1
               Set rsQD = ClsLawReadRstMsg(intJ, mStrSql)
               If intJ = 1 Then
                  rsQD.MoveFirst
                  Do While Not rsQD.EOF
                     strTmp1(9) = ""
                     If GetImgByteFile_Case(str308TM01, str308TM02, str308TM03, str308TM04, strTmp1(9), rsQD.Fields("IBF05"), strTmp1(5), strTmp1(6)) = True Then
                         Call SaveImgByteFile(strTmp1(9), mTM(1), mTM(2), mTM(3), mTM(4), strTmp1(5), strTmp1(6))
                     End If
                     rsQD.MoveNext
                  Loop
               End If
               If mTM(10) = "020" Then
                  '°Ó«~ÀÉ
                  mStrSql = "insert into TMGoods(TG01,TG02,TG03,TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17)" & _
                            " select '" & mTM(1) & "','" & mTM(2) & "','" & mTM(3) & "','" & mTM(4) & "',TG05,TG06,TG07,TG08,TG15,TG16,TG17" & _
                            " from TMGoods" & _
                            " where TG01='" & str308TM01 & "' and TG02='" & str308TM02 & "'" & _
                              " and TG03='" & str308TM03 & "' and TG04='" & str308TM04 & "'"
                  cnnConnection.Execute mStrSql, intJ
               End If
            End If
         End If
         '2024/3/15 END
      End If
   End If
   If bolError = False Then
      If ClsPDGetAutoNumber(Left(mCP(9), 1), strAutoNumber, True, True) Then
         If mCP(56) <> "" Then
            mCP(55) = mTM(23)
            'Add by Morgan 2006/11/22
            'Åý»P¤H2-5,¨üÅý¤H2-5
            mCP(93) = mTM(78)
            mCP(94) = mTM(79)
            mCP(95) = mTM(80)
            mCP(96) = mTM(81)
            'end 2006/6/23
         End If
         mCP(9) = mCP(9) + strAutoNumber
         'Modify By Sindy 2025/1/24 + np14ForCP41,np14ForCP42
         bolRt = Cls001GetNextProgressData(mTM(1), mTM(2), mTM(3), mTM(4), mCP(10), np13, np14, np14ForCP41, np14ForCP42)

         'Modify By Sindy 2012/11/06 +CP150 (¨Ö¤J) ¦³¡¹¡¹ªºÀ³¦¬±b´ÚÃ±®Ö±±ºÞ
         'Modify By Sindy 2022/9/28 +,cp140
'         If mTM(28) <> "1" And mCP(31) = "Y" Then
'            If bolRt Then
'               mstrsql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp12,cp13,cp14," + _
'                   "cp16,cp17,cp18,cp19,cp31,cp32,cp40,cp55,cp56,cp33,cp34,cp37, CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150,cp140) values (" + CNULL(mTM(1)) + "," + CNULL(mTM(2)) + "," + CNULL(mTM(3)) + "," + CNULL(mTM(4)) + "," + CNULL(mCP(5)) + "," + _
'                   CNULL(mCP(6)) + "," + CNULL(mCP(7)) + "," + CNULL(np13) + "," + CNULL(mCP(9)) + "," + CNULL(mCP(10)) + "," + CNULL(mCP(11)) + "," + CNULL(mCP(12)) + "," + CNULL(mCP(13)) + "," + CNULL(mCP(14)) + "," + CNULL(mCP(16)) + "," + _
'                   CNULL(mCP(17)) + "," + CNULL(mCP(18)) + "," + CNULL(mCP(19)) + "," + CNULL(mCP(31)) + "," + CNULL(mCP(32)) + "," + CNULL(ChgSQL(np14)) + "," + CNULL(mCP(55)) + "," + CNULL(mCP(56)) + ", " & CNULL(mCP(33)) + ", " + CNULL(mCP(34)) + _
'                   "," + CNULL(ChgSQL(mTM(5))) + "," + CNULL(ChgSQL(mCP(64))) + "," + CNULL(mCP(89)) + "," + CNULL(mCP(90)) + "," + CNULL(mCP(91)) + "," + CNULL(mCP(92)) + "," + CNULL(mCP(93)) + "," + CNULL(mCP(94)) + "," + CNULL(mCP(95)) + "," + CNULL(mCP(96)) + "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + ")"
'            Else
'               mstrsql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp12,cp13,cp14," + _
'                   "cp16,cp17,cp18,cp19,cp31,cp32,cp55,cp56,cp33,cp34,cp37, CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150,cp140) values (" + CNULL(mTM(1)) + "," + CNULL(mTM(2)) + "," + CNULL(mTM(3)) + "," + CNULL(mTM(4)) + "," + CNULL(mCP(5)) + "," + _
'                   CNULL(mCP(6)) + "," + CNULL(mCP(7)) + "," + CNULL(mCP(9)) + "," + CNULL(mCP(10)) + "," + CNULL(mCP(11)) + "," + CNULL(mCP(12)) + "," + CNULL(mCP(13)) + "," + CNULL(mCP(14)) + "," + CNULL(mCP(16)) + "," + _
'                   CNULL(mCP(17)) + "," + CNULL(mCP(18)) + "," + CNULL(mCP(19)) + "," + CNULL(mCP(31)) + "," + CNULL(mCP(32)) + "," + CNULL(mCP(55)) + "," + CNULL(mCP(56)) + ", " & CNULL(mCP(33)) + ", " + CNULL(mCP(34)) + _
'                   "," + CNULL(ChgSQL(mTM(5))) + "," + CNULL(ChgSQL(mCP(64))) + "," + CNULL(mCP(89)) + "," + CNULL(mCP(90)) + "," + CNULL(mCP(91)) + "," + CNULL(mCP(92)) + "," + CNULL(mCP(93)) + "," + CNULL(mCP(94)) + "," + CNULL(mCP(95)) + "," + CNULL(mCP(96)) + "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + ")"
'            End If
'         Else
'            If bolRt Then
'               mstrsql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp12,cp13,cp14," + _
'                   "cp16,cp17,cp18,cp19,cp31,cp32,cp40,cp55,cp56,cp33,cp34,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150,cp140) values (" + CNULL(mTM(1)) + "," + CNULL(mTM(2)) + "," + CNULL(mTM(3)) + "," + CNULL(mTM(4)) + "," + CNULL(mCP(5)) + "," + _
'                   CNULL(mCP(6)) + "," + CNULL(mCP(7)) + "," + CNULL(np13) + "," + CNULL(mCP(9)) + "," + CNULL(mCP(10)) + "," + CNULL(mCP(11)) + "," + CNULL(mCP(12)) + "," + CNULL(mCP(13)) + "," + CNULL(mCP(14)) + "," + CNULL(mCP(16)) + "," + _
'                   CNULL(mCP(17)) + "," + CNULL(mCP(18)) + "," + CNULL(mCP(19)) + "," + CNULL(mCP(31)) + "," + CNULL(mCP(32)) + "," + CNULL(ChgSQL(np14)) + "," + CNULL(mCP(55)) + "," + CNULL(mCP(56)) + ", " & CNULL(mCP(33)) + ", " + CNULL(mCP(34)) + _
'                   "," + CNULL(ChgSQL(mCP(64))) + "," + CNULL(mCP(89)) + "," + CNULL(mCP(90)) + "," + CNULL(mCP(91)) + "," + CNULL(mCP(92)) + "," + CNULL(mCP(93)) + "," + CNULL(mCP(94)) + "," + CNULL(mCP(95)) + "," + CNULL(mCP(96)) + "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + ")"
'            Else
'               mstrsql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp12,cp13,cp14," + _
'                   "cp16,cp17,cp18,cp19,cp31,cp32,cp55,cp56,cp33,cp34,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150,cp140) values (" + CNULL(mTM(1)) + "," + CNULL(mTM(2)) + "," + CNULL(mTM(3)) + "," + CNULL(mTM(4)) + "," + CNULL(mCP(5)) + "," + _
'                   CNULL(mCP(6)) + "," + CNULL(mCP(7)) + "," + CNULL(mCP(9)) + "," + CNULL(mCP(10)) + "," + CNULL(mCP(11)) + "," + CNULL(mCP(12)) + "," + CNULL(mCP(13)) + "," + CNULL(mCP(14)) + "," + CNULL(mCP(16)) + "," + _
'                   CNULL(mCP(17)) + "," + CNULL(mCP(18)) + "," + CNULL(mCP(19)) + "," + CNULL(mCP(31)) + "," + CNULL(mCP(32)) + "," + CNULL(mCP(55)) + "," + CNULL(mCP(56)) + ", " & CNULL(mCP(33)) + ", " + CNULL(mCP(34)) + _
'                   "," + CNULL(ChgSQL(mCP(64))) + "," + CNULL(mCP(89)) + "," + CNULL(mCP(90)) + "," + CNULL(mCP(91)) + "," + CNULL(mCP(92)) + "," + CNULL(mCP(93)) + "," + CNULL(mCP(94)) + "," + CNULL(mCP(95)) + "," + CNULL(mCP(96)) + "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + ")"
'            End If
'         End If
         'Modify By Sindy 2023/1/12 +,cp141,cp142
         'Modify By Sindy 2023/4/18 +,cp151
         'Modify By Sindy 2023/12/12 +,cp164
         mStrSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp12,cp13,cp14," + _
             "cp16,cp17,cp18,cp19,cp31,cp32,cp55,cp56,cp33,cp34,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150," + _
             "cp140,cp141,cp142,cp151,cp164) values (" + CNULL(mTM(1)) + "," + CNULL(mTM(2)) + "," + CNULL(mTM(3)) + "," + CNULL(mTM(4)) + "," + CNULL(mCP(5)) + "," + _
             CNULL(mCP(6)) + "," + CNULL(mCP(7)) + "," + CNULL(mCP(9)) + "," + CNULL(mCP(10)) + "," + CNULL(mCP(11)) + "," + CNULL(mCP(12)) + "," + CNULL(mCP(13)) + "," + CNULL(mCP(14)) + "," + CNULL(mCP(16)) + "," + _
             CNULL(mCP(17)) + "," + CNULL(mCP(18)) + "," + CNULL(mCP(19)) + "," + CNULL(mCP(31)) + "," + CNULL(mCP(32)) + "," + CNULL(mCP(55)) + "," + CNULL(mCP(56)) + ", " & CNULL(mCP(33)) + ", " + CNULL(mCP(34)) + _
             "," + CNULL(ChgSQL(mCP(64))) + "," + CNULL(mCP(89)) + "," + CNULL(mCP(90)) + "," + CNULL(mCP(91)) + "," + CNULL(mCP(92)) + "," + CNULL(mCP(93)) + "," + CNULL(mCP(94)) + "," + CNULL(mCP(95)) + "," + CNULL(mCP(96)) + "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + _
             "," + CNULL(mCP(141)) + "," + CNULL(mCP(142)) + "," + CNULL(mCP(151)) + "," + CNULL(mCP(164)) + ")"
         cnnConnection.Execute mStrSql
         If bolRt Then
            'Modify By Sindy 2025/1/24 + np14ForCP41,np14ForCP42
            mStrSql = "update caseprogress set CP08=" + CNULL(np13) + _
                      ",cp40=" + CNULL(ChgSQL(np14)) + _
                      ",cp41=" + CNULL(ChgSQL(np14ForCP41)) + _
                      ",cp42=" + CNULL(ChgSQL(np14ForCP42))
            mStrSql = mStrSql + " where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
         End If
         If mTM(28) <> "1" And mCP(31) = "Y" Then
            mStrSql = "update caseprogress set cp37=" + CNULL(ChgSQL(mTM(5))) + _
                      " where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
         End If
         '2023/1/12 END
         
         'Add By Sindy 2022/12/7 ¼W¥[§ó·s¥xÆWÃÒ®Ñ§Î¦¡
         Dim strUpdTMCol As String
         If mTM(136) <> "" Then
            strUpdTMCol = strUpdTMCol & ",tm136=" + CNULL(mTM(136))
         End If
         'Add By Sindy 2024/4/11 MCT¦³©¼©Ò®×¸¹®É,­n§ó·s¨äÄæ¦ì
         If mTM(45) <> "" And mTM(10) = "000" And intSaveMode = 0 Then
            strUpdTMCol = strUpdTMCol & ",tm45=" + CNULL(mTM(45))
         End If
         '2024/4/11 END
         If strUpdTMCol <> "" Then
            strUpdTMCol = Mid(strUpdTMCol, 2)
            mStrSql = "update trademark set " & strUpdTMCol
            mStrSql = mStrSql & " where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
            cnnConnection.Execute mStrSql
         End If
         '2022/12/7 END
         
         'Add By Cheng 2004/03/16
         '­Y¬°CFTªº°Ó¥Ó®×
         If mTM(1) = "CFT" And mCP(10) = "101" Then
             '­Y¥Ó½Ð°ê®a¬°•´§Q¨È(032), ¬ì«Â¯S(028), ¥ì®Ô(025), ¤ÚªL(041), ¦B®q(218), ¤¦³Á(216), §ÆÃ¾(212), ¤g¦Õ¨ä(235), ®¿«Â(215), ·ç¨å(214), À³²£¥ÍBÃþ"¥Ó½Ð­^¤åÃÒ©ú"(304)
             '2009/1/10 modify by sonia ¨ú®ø¥ì®Ô(025)¡A¤g¦Õ¨ä(235)¡A®¿«Â(215)
             'If mtm(10) = "032" Or mtm(10) = "028" Or mtm(10) = "025" Or mtm(10) = "041" Or mtm(10) = "218" Or mtm(10) = "216" Or mtm(10) = "212" Or mtm(10) = "235" Or mtm(10) = "215" Or mtm(10) = "214" Then
             '2010/5/10 modify by sonia ¼W¥[058¥§ªyº¸
             '2010/6/22 modify by sonia ¨ú®ø216¤¦³Á
             '2010/9/7  modify by sonia ¨ú®ø028¬ì«Â¯S
             '2016/9/22 modify by sonai ¨ú®ø218¦B®q
             'modify by sonia 2017/4/10 ¨ú®ø041¤ÚªL
             'modify by sonia 2018/9/18 ¨ú®ø212§ÆÃ¾
             If mTM(10) = "032" Or mTM(10) = "214" Or mTM(10) = "058" Then
                 strBKindCP09 = AutoNo("B", 6)
                 mStrSql = "Insert Into Caseprogress (CP01,CP02,CP03,CP04,CP05,CP09,CP10,CP11,CP13,CP14,CP20,CP32) " & _
                                 " Values (" + CNULL(mTM(1)) + "," + CNULL(mTM(2)) + "," + CNULL(mTM(3)) + "," + CNULL(mTM(4)) + "," + CNULL(mCP(5)) + "," + _
                                 CNULL(strBKindCP09) + "," + CNULL("304") + "," + CNULL(mCP(11)) + "," + CNULL(mCP(13)) + "," + CNULL(mCP(14)) + "," + CNULL("N") + "," + CNULL("N") + ")"
                 cnnConnection.Execute mStrSql
                 mStrSql = "Update Caseprogress Set CP12=(Select ST15 From Staff Where ST01=" + CNULL(mCP(13)) + ") Where CP09=" + CNULL(strBKindCP09)
                 cnnConnection.Execute mStrSql
             End If
         End If
         
         'Modify By Sindy 2013/8/23 +cp118(¹q¤l°e¥ó) 'Memo by Lydia 2022/08/11 ¨Ìµe­±§ó·s
         'Modified by Lydia 2022/09/20 ¦]¬°P®×¦³Trigger·|¦Û°Ê³]©w¹q¤l°e¥óCP118 = Y, ©Ò¥H§ï¦¨¨â­Ó§PÂ_
         'mstrsql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(mCP(13)) + ") " & ",cp118=" & CNULL(mCP(118)) & " where cp09=" + CNULL(mCP(9))
         mStrSql = IIf(mCP(118) = "YY", ", CP118='Y' ", IIf(mCP(118) = "YN", ", CP118=null ", ""))
         mStrSql = "update caseprogress set cp12=(select st15 from staff where st01=" & CNULL(mCP(13)) & ") " & mStrSql & " where cp09=" & CNULL(mCP(9))
         'end 2022/09/20
         cnnConnection.Execute mStrSql
           
         'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G¥xÆW®×B1¡BB2¤ÎC¦¬¤å®É¡A¼W¥["®×·½³æ¸¹"Äæ¦ì¤@©w­n¿é¤J¡A¨Ã±N®×·½³æ¸¹§ó·s¦Ü¸Óµ§¦¬¤åªºCP162¡C
         If intModifyKind = 0 And mTM(10) = "000" And (mTM(1) = "FCT" Or mTM(1) = "T" Or mTM(1) = "TC") And m_LOS02 <> "" And m_LOS15 <> "" Then
              If Left(m_LOS02, 1) = "B" Or Left(m_LOS02, 1) = "C" Then
                  mStrSql = "update caseprogress set CP162='" & m_LOS15 & "' where cp09='" & mCP(9) & "' "
                  cnnConnection.Execute mStrSql
              End If
         End If
         'end 2020/05/20
      
         '­Y¬°±µ¬¢°O¿ý³æ(Âd¥x¦¬¤å)
         'Modify by Morgan 2007/10/26 ¶O¥Î¥i§ï®É¤~°µ¡A§_«h¤w¦¬´Ú¸ê®Æ·|³QÁÙ­ì
         If intChoose = 0 And mCP(60) = "" Then 'mCP(60) = "" => txtTrademark(14).Enabled = True
             '¥¼¦¬ª÷ÃB = ¶O¥Î
             mStrSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(mCP(9))
             cnnConnection.Execute mStrSql
         End If
         'Add By Cheng 2002/01/15
         '­Y¬°¤º³¡¦¬¤å, ¥u­n¨t²ÎÃþ§O¬°T¶}ÀY©ÎFCTªº®×¥ó
         '¥B®×¥ó©Ê½è¬°201¸É¥¿, 203­×¥¿, 302§ó¥¿, 305¶Ê¼f, 306¦Û½ÐºM¦^, 307¦Û½ÐºM¾P, 614¥¼¸É²z¥Ñ, 615¥¼µªÅG, 706¨ä¥L
         '§ì¨t²Î¤é´Á§ó·s¨ä®×¥ó¶i«×ÀÉªºµo¤å¤é(CP27)
         If intChoose = 1 Then
            If Left(mTM(1), 1) = "T" Or mTM(1) = "FCT" Then
               If mCP(10) = "201" Or mCP(10) = "203" Or mCP(10) = "302" Or mCP(10) = "305" Or _
                  mCP(10) = "306" Or mCP(10) = "307" Or mCP(10) = "614" Or mCP(10) = "615" Or _
                  mCP(10) = "706" Then
                  mStrSql = "update caseprogress set cp27= '" & strSrvDate(1) & "' where cp09=" + CNULL(mCP(9))
                  cnnConnection.Execute mStrSql
               End If
            End If
         End If
         
         'Added by Lydia 2022/11/29 «D¤º³¡¦¬¤å¨Ã¥B¦³¶O¥Î¡A¥ý²Î¤@³]©wCP20=Null ;
         If intChoose = 0 And Val(mCP(16)) > 0 Then
             mStrSql = "update caseprogress set cp20=null where cp09=" + CNULL(mCP(9))
             cnnConnection.Execute mStrSql
         End If
         'end 2022/11/29
         
         '92.5.8 ADD BY SONIA
         If mTM(1) = "FCT" Then
            If Val(mCP(16)) = 0 Then
               mStrSql = "update caseprogress set cp20='N',CP32='N' where cp09=" + CNULL(mCP(9))
               cnnConnection.Execute mStrSql
            End If
         End If
         '92.5.8 END
         'Add By Cheng 2002/05/10
         '­Y¬°¤º³¡¦¬¤å§@·~®É, ®×¥ó¶i«×ÀÉªº¬O§_¦V«È¤á¦¬´Ú³]©w¬°"N"
         If intChoose = 1 Then
            mStrSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
         End If
         'Added by Lydia 2020/11/19 CFT­^°ê²æ¼Ú®×ºÞ¨î
         'Modified by Lydia 2021/03/05 §PÂ_«D¥Ó½Ð®×
         If mTM(1) = "CFT" And mCP(31) = "Y" And pType = "CFT­^°ê²æ¼Ú®×" And m_CaseNaTmp(1) <> "" And m_CaseNaTmp(2) <> "" And mCP(10) <> "101" Then
             strTmp1(0) = "select cp09,cp30 from caseprogress where cp01='" & m_CaseNaTmp(1) & "' and cp02='" & m_CaseNaTmp(2) & "' and cp03='" & m_CaseNaTmp(3) & "' and cp04='" & m_CaseNaTmp(4) & "' " & _
                              "and substr(cp09,1,1) ='C' and cp10='1730' and cp159=0 order by cp05 desc "
             intK = 1
             Set rsQD = ClsLawReadRstMsg(intK, strTmp1(0))
             If intK = 1 Then
                'A. ¼Ú·ù®×­Y¦³¡u³qª¾­^°ê¦Aµù¥U¡vªºCÃþ¨Ó¨ç1730¤§CP30¦s¦Ü·s­^°ê®×¤§¼f©w¸¹TM15
                If "" & rsQD.Fields("CP30") <> "" Then
                    mStrSql = "update TRADEMARK set TM15='" & ChgSQL(rsQD.Fields("cp30")) & "' where TM01='" & mTM(1) & "' and TM02='" & mTM(2) & "' and TM03='" & mTM(3) & "' and TM04='" & mTM(4) & "' "
                    cnnConnection.Execute mStrSql
                End If
                'B. ¼Ú·ù®×­Y¦³¡u³qª¾­^°ê¦Aµù¥U¡vªºCÃþ¨Ó¨ç¤]Âà¦Ü·s­^°ê®×¸¹
                If "" & rsQD.Fields("cp09") <> "" Then
                     mStrSql = "update caseprogress set cp01='" & mTM(1) & "', cp02='" & mTM(2) & "', cp03='" & mTM(3) & "', cp04='" & mTM(4) & "' where cp09='" & rsQD.Fields("cp09") & "' "
                     cnnConnection.Execute mStrSql
                End If
             End If
             'Added by Lydia 2020/12/01
             If mCP(10) = "710" Then
                  '©e¥ô¥N²z¤H¤WÄò¿ì¡F¤U¤@µ{§Ç³Æµù¥[µù¡u­^°ê®×®×¸¹¡v
                  mStrSql = "update nextprogress set np06='Y', np24='" & mCP(9) & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "­^°ê®×®×¸¹¡G" & mTM(1) & mTM(2) & mTM(3) & mTM(4) & ";'||np15 " & _
                              "where np02='" & m_CaseNaTmp(1) & "' and np03='" & m_CaseNaTmp(2) & "' and np04='" & m_CaseNaTmp(3) & "' and np05='" & m_CaseNaTmp(4) & "' and np07='710' and np06 is null "
                  cnnConnection.Execute mStrSql
                  'E. ­Y¦¬¤å¡u©e¥ô¥N²z¤H(CFT.710)¡v®É¼Ú·ù®×¤U¤@µ{§Ç¤§/¡u©µ®i(­^°ê)110¡v´Á­­Âà¦Ü·s®×¸¹¨Ã§ï®×¥ó©Ê½è¬°¡u©µ®i102¡v¡F¤U¤@µ{§Ç³Æµù¥[µù¡u¼Ú·ù®×®×¸¹¡v
                  'Modified by Lydia 2020/12/16 ±NNP01§ï¬°­^°ê®×¦¬¤å¸¹; §_«h¤À®×§@·~·|¿ù»~(¥»©Ò®×¸¹¤£¦P)
                  mStrSql = "update nextprogress set np01='" & mCP(9) & "', np02='" & mTM(1) & "', np03='" & mTM(2) & "', np04='" & mTM(3) & "', np05='" & mTM(4) & "', np07='102', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "¼Ú·ù®×®×¸¹¡G" & m_CaseNaTmp(1) & m_CaseNaTmp(2) & m_CaseNaTmp(3) & m_CaseNaTmp(4) & ";'||np15 " & _
                              "where np02='" & m_CaseNaTmp(1) & "' and np03='" & m_CaseNaTmp(2) & "' and np04='" & m_CaseNaTmp(3) & "' and np05='" & m_CaseNaTmp(4) & "' and np07='110' and np06 is null "
                  cnnConnection.Execute mStrSql
             Else
             'end 2020/12/01
                  'C. ¼Ú·ù®×¤U¤@µ{§Ç¤§¡u©µ®i(­^°ê)¡v(CFT.110)´Á­­¤WÄò¿ìNP06¡A¤U¤@³æ¾Ú½s¸¹NP24°O¿ý·s­^°ê®×¤§Á`¦¬¤å¸¹¡F¤U¤@µ{§Ç³Æµù¥[µù¡u­^°ê®×®×¸¹¡v
                  mStrSql = "update nextprogress set np06='Y', np24='" & mCP(9) & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "­^°ê®×®×¸¹¡G" & mTM(1) & mTM(2) & mTM(3) & mTM(4) & ";'||np15 where np02='" & m_CaseNaTmp(1) & "' and np03='" & m_CaseNaTmp(2) & "' and np04='" & m_CaseNaTmp(3) & "' and np05='" & m_CaseNaTmp(4) & "' and np07='110' and np06 is null "
                  cnnConnection.Execute mStrSql
                  'Added by Lydia 2020/12/04 ±N¼Ú·ù®×¡u©e¥ô¥N²z¤H¡v´Á­­Âà¦Ü·s­^°ê®×¡F¤U¤@µ{§Ç³Æµù¥[µù¡u¼Ú·ù®×®×¸¹¡v
                  'Modified by Lydia 2020/12/16 ±NNP01§ï¬°­^°ê®×¦¬¤å¸¹; §_«h¤À®×§@·~·|¿ù»~(¥»©Ò®×¸¹¤£¦P)
                  mStrSql = "update nextprogress set np01='" & mCP(9) & "', np02='" & mTM(1) & "', np03='" & mTM(2) & "', np04='" & mTM(3) & "', np05='" & mTM(4) & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "¼Ú·ù®×®×¸¹¡G" & m_CaseNaTmp(1) & m_CaseNaTmp(2) & m_CaseNaTmp(3) & m_CaseNaTmp(4) & ";'||np15 " & _
                              "where np02='" & m_CaseNaTmp(1) & "' and np03='" & m_CaseNaTmp(2) & "' and np04='" & m_CaseNaTmp(3) & "' and np05='" & m_CaseNaTmp(4) & "' and np07='710' and np06 is null "
                  cnnConnection.Execute mStrSql
                  'end 2020/12/04
             End If 'Added by Lydia 2020/12/01
             'D. «Ø¥ß¼Ú·ù®×¤Î­^°ê®×¤§ÃöÁp(¬ÛÃö¨÷¸¹)
             mStrSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(mTM(1)) & ", " & CNULL(mTM(2)) & ", " & CNULL(mTM(3)) & ", " & CNULL(mTM(4)) & ", " & CNULL(m_CaseNaTmp(1)) & ", " & CNULL(m_CaseNaTmp(2)) & ", " & CNULL(m_CaseNaTmp(3)) & ", " & CNULL(m_CaseNaTmp(4)) & " ) "
             cnnConnection.Execute mStrSql
             mStrSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_CaseNaTmp(1)) & ", " & CNULL(m_CaseNaTmp(2)) & ", " & CNULL(m_CaseNaTmp(3)) & ", " & CNULL(m_CaseNaTmp(4)) & ", " & CNULL(mTM(1)) & ", " & CNULL(mTM(2)) & ", " & CNULL(mTM(3)) & ", " & CNULL(mTM(4)) & " ) "
             cnnConnection.Execute mStrSql
             'Added by Lydia 2020/12/04 ¼Ú·ù®×®×¥ó³Æµù¥[µù¡u­^°ê®×®×¸¹¡v¡F·s­^°ê®×¤§·s®×¦¬¤åªº¶i«×³Æµù¥[µù¡u¼Ú·ù®×®×¸¹¡v
             mStrSql = "Update Trademark set TM58=" & CNULL("­^°ê®×®×¸¹¡G" & mTM(1) & mTM(2) & mTM(3) & mTM(4) & ";") & "||TM58 where tm01='" & m_CaseNaTmp(1) & "' and tm02='" & m_CaseNaTmp(2) & "' and tm03='" & m_CaseNaTmp(3) & "' and tm04='" & m_CaseNaTmp(4) & "' "
             cnnConnection.Execute mStrSql
             mStrSql = "Update CaseProgress set CP64=" & CNULL("¼Ú·ù®×®×¸¹¡G" & m_CaseNaTmp(1) & m_CaseNaTmp(2) & m_CaseNaTmp(3) & m_CaseNaTmp(4) & ";") & "||CP64 where CP09='" & mCP(9) & "' "
             cnnConnection.Execute mStrSql
             'end 2020/12/04
             'Added by Lydia 2021/01/11 ½Æ»sÀu¥ýÅv¸ê®Æ
             mStrSql = "insert into pridate (pd01,pd02,pd03,pd04,pd05,pd06,pd07,pd08,pd09,pd10) " & _
                          "select " & CNULL(mTM(1)) & ", " & CNULL(mTM(2)) & ", " & CNULL(mTM(3)) & ", " & CNULL(mTM(4)) & ", pd05,pd06,pd07,pd08,pd09,pd10 from pridate where pd01='" & m_CaseNaTmp(1) & "' and pd02='" & m_CaseNaTmp(2) & "' and pd03='" & m_CaseNaTmp(3) & "' and pd04='" & m_CaseNaTmp(4) & "' "
             cnnConnection.Execute mStrSql
             'end 2021/01/11
             'Added by Lydia 2021/04/15 CFT­^°ê²æ¼Ú©e¥ô¥N²z¤§«áÄò³B²z¡G¦¬¤å­^°ê©µ®i¤Î©e¥ô¥N²z¤H·s®×¡A¦P®É±N¥N²z¤H¦s¤JCP44¡C
             strTmp1(0) = "select np01,np15 from nextprogress where np07='710' and np15 like '%²æ¼Ú­^°ê®×¥N²z¤H¡G%' " & _
                              "and ((np02='" & mTM(1) & "' and np03='" & mTM(2) & "' and np04='" & mTM(3) & "' and np05='" & mTM(4) & "') or (np02='" & m_CaseNaTmp(1) & "' and np03='" & m_CaseNaTmp(2) & "'  and np04='" & m_CaseNaTmp(3) & "'  and np05='" & m_CaseNaTmp(4) & "')) "
             intK = 1
             Set rsQD = ClsLawReadRstMsg(intK, strTmp1(0))
             If intK = 1 Then
                 strTmp1(1) = Mid("" & rsQD.Fields("np15"), InStr(rsQD.Fields("np15"), "²æ¼Ú­^°ê®×¥N²z¤H¡G") + 9, 9)
                 If Left(strTmp1(1), 1) = "Y" Then
                     mStrSql = "Update CaseProgress set cp44=" & CNULL(strTmp1(1)) & " where cp09=" & CNULL(mCP(9))
                     cnnConnection.Execute mStrSql
                 End If
             End If
             'end 2021/04/15
             PUB_EUtoUK mTM(1), mTM(2), mTM(3), mTM(4), m_CaseNaTmp(1), m_CaseNaTmp(2), m_CaseNaTmp(3), m_CaseNaTmp(4), mCP(9), mCP(10) 'Added by Morgan 2020/12/21 ¦^ÂÐ³æÂk¨÷
         End If
         'end 2020/11/19
        
         'Added by Lydia 2021/03/05 CFT¼Ú·ù©|¥¼µù¥U®×Âà´«­^°ê¥Ó½Ð®×¦¬¤å±±ºÞ¡G¤@¨Ö±N¼Ú·ùÃöÁp®×¤§¡u°Ó¼Ð¹Ï¼Ë¡v¡B¡u°Ó«~/ªA°ÈÃþ§O¤Î¦WºÙ¡v¡B¡u¥Ó½Ð¤é¡v¡B¡uÀu¥ýÅv¸ê®Æ¡v±a¤J·s®×¸¹
         If mTM(1) = "CFT" And mCP(31) = "Y" And pType = "CFT­^°ê²æ¼Ú®×" And m_CaseNaTmp(1) <> "" And m_CaseNaTmp(2) <> "" And mCP(10) = "101" Then
             '«Ø¥ß¼Ú·ù®×¤Î­^°ê®×¤§ÃöÁp(¬ÛÃö¨÷¸¹)
             mStrSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(mTM(1)) & ", " & CNULL(mTM(2)) & ", " & CNULL(mTM(3)) & ", " & CNULL(mTM(4)) & ", " & CNULL(m_CaseNaTmp(1)) & ", " & CNULL(m_CaseNaTmp(2)) & ", " & CNULL(m_CaseNaTmp(3)) & ", " & CNULL(m_CaseNaTmp(4)) & " ) "
             cnnConnection.Execute mStrSql
             mStrSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_CaseNaTmp(1)) & ", " & CNULL(m_CaseNaTmp(2)) & ", " & CNULL(m_CaseNaTmp(3)) & ", " & CNULL(m_CaseNaTmp(4)) & ", " & CNULL(mTM(1)) & ", " & CNULL(mTM(2)) & ", " & CNULL(mTM(3)) & ", " & CNULL(mTM(4)) & " ) "
             cnnConnection.Execute mStrSql
             '½Æ»sÀu¥ýÅv¸ê®Æ
             mStrSql = "insert into pridate (pd01,pd02,pd03,pd04,pd05,pd06,pd07,pd08,pd09,pd10) " & _
                          "select " & CNULL(mTM(1)) & ", " & CNULL(mTM(2)) & ", " & CNULL(mTM(3)) & ", " & CNULL(mTM(4)) & ", pd05,pd06,pd07,pd08,pd09,pd10 from pridate where pd01='" & m_CaseNaTmp(1) & "' and pd02='" & m_CaseNaTmp(2) & "' and pd03='" & m_CaseNaTmp(3) & "' and pd04='" & m_CaseNaTmp(4) & "' "
             cnnConnection.Execute mStrSql
         End If
         'end 2021/03/05
        
         'Added by Lydia 2021/04/15 CFT­^°ê²æ¼Ú©e¥ô¥N²z¤§«áÄò³B²z¡G¦¬¤å­^°ê©µ®i¤Î©e¥ô¥N²z¤H·s®×¡A¦P®É±N¥N²z¤H¦s¤JCP44¡C
                                                 '¦P¤@¤Ñ±µ¬¢³æ¤§«á¦¬¤åªº³B²z
         If mTM(1) = "CFT" And mCP(31) <> "Y" And mTM(10) = "201" And (mCP(10) = "710" Or mCP(10) = "102") And mTM(58) <> "" And InStr(mTM(58), "¼Ú·ù®×®×¸¹¡G") > 0 Then
             strTmp1(0) = "select cp44 from caseprogress where cp01='" & mTM(1) & "' and cp02='" & mTM(2) & "' and cp03='" & mTM(3) & "' and cp04='" & mTM(4) & "' and cp05=" & mCP(5) & " and cp10 in ('102','710') and cp159=0 And cp31='Y' "
             intK = 1
             Set rsQD = ClsLawReadRstMsg(intK, strTmp1(0))
             If intK = 1 Then
                 If "" & rsQD.Fields("cp44") <> "" Then
                     mStrSql = "Update CaseProgress set cp44=" & CNULL(rsQD.Fields("cp44")) & " where cp09=" & CNULL(mCP(9))
                     cnnConnection.Execute mStrSql
                 End If
             End If
         End If
         'end 2021/04/15
        
         'Added by Lydia 2020/12/15 CFT½q¨l­«·s¥Ó½Ð®×¡G«Ø¥ßÃöÁp
         If mTM(1) = "CFT" And mTM(10) = "048" And pType = "CFT½q¨l­«·s¥Ó½Ð®×" And mCP(10) = "101" And m_CaseNaTmp(1) <> "" And m_CaseNaTmp(2) <> "" Then
             mStrSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(mTM(1)) & ", " & CNULL(mTM(2)) & ", " & CNULL(mTM(3)) & ", " & CNULL(mTM(4)) & ", " & CNULL(m_CaseNaTmp(1)) & ", " & CNULL(m_CaseNaTmp(2)) & ", " & CNULL(m_CaseNaTmp(3)) & ", " & CNULL(m_CaseNaTmp(4)) & " ) "
             cnnConnection.Execute mStrSql
             mStrSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_CaseNaTmp(1)) & ", " & CNULL(m_CaseNaTmp(2)) & ", " & CNULL(m_CaseNaTmp(3)) & ", " & CNULL(m_CaseNaTmp(4)) & ", " & CNULL(mTM(1)) & ", " & CNULL(mTM(2)) & ", " & CNULL(mTM(3)) & ", " & CNULL(mTM(4)) & " ) "
             cnnConnection.Execute mStrSql
             'Added by Lydia 2021/02/01 ½Æ»s°Ó«~/ªA°ÈÃþ§O¤Î¦WºÙ¡BÀu¥ýÅv¸ê®Æ
             mStrSql = "INSERT INTO TMGOODS(TG01,TG02,TG03,TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17) " & _
                         "SELECT '" & mTM(1) & "' as TG01, '" & mTM(2) & "' as TG02, '" & mTM(3) & "' as  TG03, '" & mTM(4) & "' as  TG04,TG05,TG06,TG07,TG08,TG15,TG16,TG17 " & _
                         "FROM TMGOODS WHERE TG01='" & m_CaseNaTmp(1) & "'  AND TG02='" & m_CaseNaTmp(2) & "' AND TG03='" & m_CaseNaTmp(3) & "'  AND TG04='" & m_CaseNaTmp(4) & "'  AND TG18 IS NULL "
             cnnConnection.Execute mStrSql
             mStrSql = "insert into pridate (pd01,pd02,pd03,pd04,pd05,pd06,pd07,pd08,pd09,pd10) " & _
                          "select " & CNULL(mTM(1)) & ", " & CNULL(mTM(2)) & ", " & CNULL(mTM(3)) & ", " & CNULL(mTM(4)) & ", pd05,pd06,pd07,pd08,pd09,pd10 from pridate where pd01='" & m_CaseNaTmp(1) & "' and pd02='" & m_CaseNaTmp(2) & "' and pd03='" & m_CaseNaTmp(3) & "' and pd04='" & m_CaseNaTmp(4) & "' "
             cnnConnection.Execute mStrSql
             'end 2021/02/01
         End If
         'end 2020/12/15
         
         'Modify By Sindy 2023/11/3 mark:¦¬¤å¤w¤£»Ý°õ¦æ¦¹¬qµ{¦¡,¦]±µ¬¢³æ¤w·|¦^¼g¶l»¼°Ï¸¹
'         mstrsql = "update customer set cu30=" + CNULL(mCU30) + " where cu01=" + CNULL(Mid(mTM(23), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(23), 9, 1))
'         cnnConnection.Execute mstrsql
         '2023/11/3 END
         
         adoquery.CursorLocation = adUseClient
         'add by nickc 2007/10/24 ¤º°Ó¥þ´Áµù¥U¶O(717)®É¡A§ì²Ä¤@´Áµù¥U¶O(715)
         'edit by nickc 2007/10/25 ¥[¤J¥~°Ó
         If (mTM(1) = "T" Or mTM(1) = "FCT") And (mCP(10) = "717" Or mCP(10) = "715") Then
               adoquery.Open "select np01 from nextprogress where np02 = '" & mTM(1) & "' and np03 = '" & mTM(2) & "' and np04 = '" & mTM(3) & "' and np05 = '" & mTM(4) & "' and np06 is null and np07 in('715','717') ", cnnConnection, adOpenStatic, adLockReadOnly
         Else
            'Add By Sindy 2012/7/4 °¨¼w¨½¨Ï¥Î«Å»}
            If mTM(1) = "TF" And mCP(10) = "105" Then
               adoquery.Open "select np01 from nextprogress where np02 = '" & mTM(1) & "' and substr(np03,1,5) = '" & Left(mTM(2), 5) & "' and np06 is null and np07 = '" & mCP(10) & "' and np08=" & IIf(mCP(6) = "", 0, mCP(6)), cnnConnection, adOpenStatic, adLockReadOnly
            Else
            '2012/7/4 End
               adoquery.Open "select np01 from nextprogress where np02 = '" & mTM(1) & "' and np03 = '" & mTM(2) & "' and np04 = '" & mTM(3) & "' and np05 = '" & mTM(4) & "' and np06 is null and np07 = '" & mCP(10) & "'", cnnConnection, adOpenStatic, adLockReadOnly
            End If
         End If
         If adoquery.RecordCount > 0 Then
            'Modify By Sindy 2012/7/4 °¨¼w¨½¨Ï¥Î«Å»}¦³¥i¯à¦hµ§
            If (adoquery.RecordCount = 1 Or (mTM(1) = "TF" And mCP(10) = "105" And adoquery.RecordCount >= 1)) Then
            '2012/7/4 End
               If IsNull(adoquery.Fields(0).Value) = False Then
                  '2011/6/16 add by sonia ²§Ä³µªÅG¡Bµû©wµªÅG¡B¼o¤îµªÅG­n¤@¨Ã§ó·s¹ï³y¸ê®Æ
                  If (mCP(10) = "602" Or mCP(10) = "604" Or mCP(10) = "606") Then
                     cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp80) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42,b.cp80 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & mCP(9) & "'", intK
                  Else
                  '2011/6/16 end
                     cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & mCP(9) & "'"
                  End If  '2011/6/16 add by sonia
               End If
               'add by nick 2004/09/08
               If mCP(10) <> "305" Then
                   'add by nickc 2007/10/24 ¤º°Ó¥þ´Áµù¥U¶O(717)®É¡A§ì²Ä¤@´Áµù¥U¶O(715)
                   'edit by nickc 2007/10/25 ¥[¤J¥~°Ó
                   If (mTM(1) = "T" Or mTM(1) = "FCT") And (mCP(10) = "717" Or mCP(10) = "715") Then
                       mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np02=" + CNULL(mTM(1)) + " and np03=" + _
                          CNULL(mTM(2)) + " and np04=" + CNULL(mTM(3)) + " and np05=" + CNULL(mTM(4)) + _
                          " and np07 in('715','717') and np06 is null"
                       cnnConnection.Execute mStrSql
                   Else
                     'Add By Sindy 2012/7/4 °¨¼w¨½¨Ï¥Î«Å»}
                     If mTM(1) = "TF" And mCP(10) = "105" Then
                        mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np02=" + CNULL(mTM(1)) + " and substr(np03,1,5)=" + _
                           CNULL(Left(mTM(2), 5)) + _
                           " and np07=" + CNULL(mCP(10)) + " and np06 is null and np08=" & IIf(mCP(6) = "", 0, mCP(6)) + _
                           " and np02||np03||np04||np05 in(select tm01||tm02||tm03||tm04 from trademark where tm01=np02 and tm02=np03 and tm03=np04 and tm04=np05 and tm29 is null)"
                        cnnConnection.Execute mStrSql
                     Else
                     '2012/7/4 End
                        mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np02=" + CNULL(mTM(1)) + " and np03=" + _
                           CNULL(mTM(2)) + " and np04=" + CNULL(mTM(3)) + " and np05=" + CNULL(mTM(4)) + _
                           " and np07=" + CNULL(mCP(10)) + " and np06 is null"
                        cnnConnection.Execute mStrSql
                     End If
                   End If
               End If
            End If
         Else
            adoquery.Close
            adoquery.CursorLocation = adUseClient
            'Add By Sindy 2025/3/18 ¼W¥[¨R¾P±ø¥ó: ªk©w´Á­­¦³­È®É,¶·ÀË¬d¤@¼Ëªº´Á­­¤~¨R¾P
            strExc(10) = ""
            If Val(mCP(7)) > 0 Then
               strExc(10) = " and np09=" & mCP(7)
            End If
            'add by nickc 2007/10/24 ¤º°Ó¥þ´Áµù¥U¶O(717)®É¡A§ì²Ä¤@´Áµù¥U¶O(715)
            'edit by nickc 2007/10/25 ¥[¤J¥~°Ó
             If (mTM(1) = "T" Or mTM(1) = "FCT") And (mCP(10) = "717" Or mCP(10) = "715") Then
               adoquery.Open "select np01 from nextprogress where np02 = '" & mTM(1) & "' and np03 = '" & mTM(2) & "' and np04 = '" & mTM(3) & "' and np05 = '" & mTM(4) & "' and np06 <>'Y' and np07 in('715','717')" & strExc(10), cnnConnection, adOpenStatic, adLockReadOnly
            Else
               'Add By Sindy 2012/7/4 °¨¼w¨½¨Ï¥Î«Å»}
               If mTM(1) = "TF" And mCP(10) = "105" Then
                  adoquery.Open "select np01 from nextprogress where np02 = '" & mTM(1) & "' and substr(np03,1,5) = '" & Left(mTM(2), 5) & "' and np06 <>'Y' and np07 = '" & mCP(10) & "' and np08=" & IIf(mCP(6) = "", 0, mCP(6)) & strExc(10), cnnConnection, adOpenStatic, adLockReadOnly
               Else
               '2012/7/4 End
                  adoquery.Open "select np01 from nextprogress where np02 = '" & mTM(1) & "' and np03 = '" & mTM(2) & "' and np04 = '" & mTM(3) & "' and np05 = '" & mTM(4) & "' and np06 <>'Y' and np07 = '" & mCP(10) & "'" & strExc(10), cnnConnection, adOpenStatic, adLockReadOnly
               End If
            End If
            If adoquery.RecordCount > 0 Then
               'Modify By Sindy 2012/7/4 °¨¼w¨½¨Ï¥Î«Å»}¦³¥i¯à¦hµ§
               If (adoquery.RecordCount = 1 Or (mTM(1) = "TF" And mCP(10) = "105" And adoquery.RecordCount >= 1)) Then
               '2012/7/4 End
                  If IsNull(adoquery.Fields(0).Value) = False Then
                     '2011/6/16 add by sonia ²§Ä³µªÅG¡Bµû©wµªÅG¡B¼o¤îµªÅG­n¤@¨Ã§ó·s¹ï³y¸ê®Æ
                     If (mCP(10) = "602" Or mCP(10) = "604" Or mCP(10) = "606") Then
                        cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp80) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42,b.cp80 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & mCP(9) & "'", intK
                     Else
                     '2011/6/16 end
                        cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & mCP(9) & "'"
                     End If  '2011/6/16 add by sonia
                  End If
                  'add by nick 2004/09/08
                  If mCP(10) <> "305" Then
                       'add by nickc 2007/10/24 ¤º°Ó¥þ´Áµù¥U¶O(717)®É¡A§ì²Ä¤@´Áµù¥U¶O(715)
                       'edit by nickc 2007/10/25 ¥[¤J¥~°Ó
                       If (mTM(1) = "T" Or mTM(1) = "FCT") And (mCP(10) = "717" Or mCP(10) = "715") Then
                           mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np02=" + CNULL(mTM(1)) + " and np03=" + _
                              CNULL(mTM(2)) + " and np04=" + CNULL(mTM(3)) + " and np05=" + CNULL(mTM(4)) + _
                              " and np07 in('715','717') and np06 <> 'Y'" & strExc(10)
                           cnnConnection.Execute mStrSql
                       Else
                           'Add By Sindy 2012/7/4 °¨¼w¨½¨Ï¥Î«Å»}
                           If mTM(1) = "TF" And mCP(10) = "105" Then
                              mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np02=" + CNULL(mTM(1)) + " and substr(np03,1,5)=" + _
                                 CNULL(Left(mTM(2), 5)) + _
                                 " and np07=" + CNULL(mCP(10)) + " and np06 <> 'Y' and np08=" & IIf(mCP(6) = "", 0, mCP(6)) + strExc(10) + _
                                 " and np02||np03||np04||np05 in(select tm01||tm02||tm03||tm04 from trademark where tm01=np02 and tm02=np03 and tm03=np04 and tm04=np05 and tm29 is null)"
                              cnnConnection.Execute mStrSql
                           Else
                           '2012/7/4 End
                              mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np02=" + CNULL(mTM(1)) + " and np03=" + _
                                 CNULL(mTM(2)) + " and np04=" + CNULL(mTM(3)) + " and np05=" + CNULL(mTM(4)) + _
                                 " and np07=" + CNULL(mCP(10)) + " and np06 <> 'Y'" & strExc(10)
                              cnnConnection.Execute mStrSql
                           End If
                       End If
                   End If
               End If
            End If
            '2025/3/18 END
         End If
         adoquery.Close
         '92.2.19 END
         
         'Modify By Sindy 2022/9/28 ¹q¤l¦¬¤å:±q±µ¬¢³æMove¦¹³B§ó·s
         'Add By Sindy 2015/9/15
         '¸É¥¿,ÀË°e¦P·N®Ñ,©ñ±ó±M¥ÎÅv
         If mCP(1) = "T" And _
            (mCP(10) = "201" Or mCP(10) = "211" Or mCP(10) = "206") Then
            
            'Add By Sindy 2025/2/26 ¼W¥[¨R¾P±ø¥ó: ªk©w´Á­­¦³­È®É,¶·ÀË¬d¤@¼Ëªº´Á­­¤~¨R¾P
            strExc(10) = ""
            If Val(mCP(7)) > 0 Then
               strExc(10) = " and np09=" & mCP(7)
            End If
            '¥ýÅª ¥Ó½Ð·N¨£®Ñ¥BNP06 is null ¦Û°Ê¨R¾P¤U¤@µ{§Ç
            adoquery.CursorLocation = adUseClient
            adoquery.Open "select * from nextprogress where " & ChgNextProgress(mCP(1) & mCP(2) & mCP(3) & mCP(4)) & _
                          " and np06 is null and np07='202'" & strExc(10) & " order by np08 asc", cnnConnection, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount > 0 Then
               adoquery.MoveFirst
               strCP08 = "" & adoquery.Fields("np13")
               strCP43 = "" & adoquery.Fields("np01")
               cnnConnection.Execute "update caseprogress set cp43 = '" & strCP43 & "',cp08='" & strCP08 & "' where cp09 = '" & mCP(9) & "'"
               'Modify By Sindy 2019/10/29 ¦Û°Ê¦¬¤åNP24­n§ó·s¬°Á`¦¬¤å¸¹
               cnnConnection.Execute "update nextprogress set np06='Y',np24='" & mCP(9) & "' where " & ChgNextProgress(mCP(1) & mCP(2) & mCP(3) & mCP(4)) & _
                                     " and np06 is null and np07='202'"
            Else
               adoquery.Close
               '¦AÅª ¥Ó½Ð·N¨£®Ñ¥Bnp06='N' ¦Û°Ê¨R¾P¤U¤@µ{§Ç
               adoquery.Open "select * from nextprogress where " & ChgNextProgress(mCP(1) & mCP(2) & mCP(3) & mCP(4)) & _
                             " and np06='N' and np07='202'" & strExc(10) & " order by np08 desc", cnnConnection, adOpenStatic, adLockReadOnly
               If adoquery.RecordCount > 0 Then
                  adoquery.MoveFirst
                  strCP08 = "" & adoquery.Fields("np13")
                  strCP43 = "" & adoquery.Fields("np01")
                  cnnConnection.Execute "update caseprogress set cp43 = '" & strCP43 & "',cp08='" & strCP08 & "' where cp09 = '" & mCP(9) & "'"
                  'Modify By Sindy 2019/10/29 ¦Û°Ê¦¬¤åNP24­n§ó·s¬°Á`¦¬¤å¸¹
                  cnnConnection.Execute "update nextprogress set np06='Y',np24='" & mCP(9) & "' where " & ChgNextProgress(mCP(1) & mCP(2) & mCP(3) & mCP(4)) & _
                                        " and np06='N' and np07='202'"
               End If
            End If
            '2025/2/26 END
            adoquery.Close
         End If
         '2015/9/15 END
         
         ' Add by Sindy 98/03/02
         '¦¬¤å®É­Y¸Ó®×¸¹¤U¤@µ{§Ç¤´¦³F4103¥B¬O§_Äò¿ì¬°NULLªÌ,
         '§ó·s¤U¤@µ{§ÇF4103¬°¦¬¤å´¼Åv¤H­û
         If mTM(1) = "FCT" Then
            mStrSql = "update nextprogress set np10='" & mCP(13) & "' " & _
               "where np02=" + CNULL(mTM(1)) + " and np03=" + _
               CNULL(mTM(2)) + " and np04=" + CNULL(mTM(3)) + " and np05=" + CNULL(mTM(4)) + _
               " and np10='F4103' and np06 is null"
            cnnConnection.Execute mStrSql
         End If
         ' 98/03/02 End
                  
         'Added by Morgan 2021/6/22
         'T»PFCT¦@¦P±±ºÞ®×¥ó³qª¾
         'Modified by Morgan 2023/5/4 §ï¦bPUB_2SysCaseInform¤º§ì¨t²Î¯S®í³]©w
         'If (mTM(1) = "T" And (mTM(2) = "211948" Or mTM(2) = "211949") And mTM(3) = "0" And mTM(4) = "00") Or (mTM(1) = "FCT" And (mTM(2) = "047561" Or mTM(2) = "047562") And mTM(3) = "0" And mTM(4) = "00") Then
         If mTM(1) = "T" Or mTM(1) = "FCT" Then
         'end 2023/5/4
            PUB_2SysCaseInform mTM(1), mTM(2), mTM(3), mTM(4), mCP(9), 2
         End If
         'end 2021/6/22
         
         'Add By Sindy 2024/7/12 ÀË¬d¬O§_¦³¨R¤U¤@µ{§Çªº´Á­­
         '                       ­Y¦³,¦¹µ§¦¬¤å´Á­­§ó·s¬°¤U¤@µ{§Çªº´Á­­
         If UCase(pFormName) = UCase("frm090801_New") Then
            strTmp1(0) = "select np01,np08,np09 from nextprogress where np24 = '" & mCP(9) & "'"
            intJ = 1
            Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
            If intJ = 1 Then
               mStrSql = "update caseprogress set CP06=" & "" & rsQD.Fields("np08") & _
                                                ",CP07=" & "" & rsQD.Fields("np09")
               mStrSql = mStrSql & " where cp09=" & CNULL(mCP(9))
               cnnConnection.Execute mStrSql, intJ
            End If
         End If
         '2024/7/12 END
         
         'Add By Sindy 2024/1/4 ¹q¤l¦¬¤å¤£¶·°õ¦æ¦¹¨ç¼Æ
         If UCase(pFormName) <> UCase("frm090801_New") Then
         '2024/1/4 END
            If Cls001SetCaseProgressFee(mTM(1), mTM(10), mCP(10), mCP(9)) = False Then bolError = True
         End If
      Else
         bolError = True
      End If
   End If

   If bolError = False Then
      'Remove by Lydia 2018/08/22 (À³¦¬±b´ÚºÞ±±)¨ú®ø¹w©w¦¬´Ú¤é,§ï¦¨¥I´Ú¶g´Á
'      Dim rtCnt As Integer
'      'Modify by Morgan 2010/12/9
'      'If txtTrademark(34) <> "" Then
'      '    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " ", rtCnt
'      If txtTrademark(34) <> "" And txtTrademark(34) <> txtTrademark(34).Tag Then
'          cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
'      'end 2010/12/9
'          If rtCnt = 0 Then
'              cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " from dual "
'          End If
'      End If
      'end 2018/08/22
      
      If mSaveControl <> "" Then
          'Modified by Lydia 2022/09/29 ¶Ç¤J¨t²Î§O,°ê®a,®×¥ó©Ê½è=>mTM(1), mTM(10), mCP(10)
          Call PUB_SaveByControl(mCP(9), mSaveControl, mTM(1), mTM(10), mCP(10))
      End If
   End If
   
   'Add by Sindy 2023/5/31
   'Modify By Sindy 2025/8/18
   'If InStr(pType, "«H¥ó¨R¾P") > 0 And m_strIR01 <> "" Then
   If m_strIR01 <> "" Then
   '2025/8/18 END
      m_bolRecvOK = True
      m_strMCR11 = ""
      If m_bMRecvBatch = True Then '¦h®×¦¬¤å
         '§ó·sÁ`¦¬¤å¸¹
         mStrSql = "update multiCaseRecv set mcr11='" & mCP(9) & "'" & _
                  " where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                  " and mcr02='" & mTM(1) & "' and mcr03='" & mTM(2) & "' and mcr04='" & mTM(3) & "' and mcr05='" & mTM(4) & "'" & _
                  " and mcr06='" & mCP(10) & "'"
                  cnnConnection.Execute mStrSql
                  
         'Modify By Sindy 2022/8/26
         '¤U¸ü«H¥óÀÉ,¤W¶Ç¨÷©v°Ï
         Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, mCP(9))
                  
         'ÀË¬d¦h®×¦¬¤åª¬ªp
         strTmp1(0) = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                     " and mcr02||mcr03||mcr04||mcr05<>'" & mTM(1) & mTM(2) & mTM(3) & mTM(4) & "'" & _
                     " and mcr11 is null"
         intJ = 1
         Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
         If intJ = 1 Then
            m_bolRecvOK = False '©|¦³¥¼¦¬¤å

            'Modify By Sindy 2022/8/26 ¦¹³BMark,µ{¦¡©¹¤W²¾
'            '¤U¸ü«H¥óÀÉ,¤W¶Ç¨÷©v°Ï
'            Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, mCP(9))
         Else
            m_bolRecvOK = True '¬O§_¦¬§¹¤å
            '§ì²Ä¤@µ§ªºÁ`¦¬¤å¸¹
            strTmp1(0) = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                        " and mcr02||mcr03||mcr04||mcr05=mcr07||mcr08||mcr09||mcr10 and mcr11 is not null"
            intJ = 1
            Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
            If intJ = 1 Then
               m_strMCR11 = rsQD.Fields("mcr11")
               RetVal = RetVal & IIf(RetVal <> "", ",", "") & "MCR11:" & m_strMCR11
            Else
               MsgBox "¦h®×¦¬¤å¡AµLÅª¨ú¨ì²Ä¤@µ§®×¥óªºÁ`¦¬¤å¸¹¡A½Ð¬¢¹q¸£¤¤¤ß!!", vbExclamation '¦¹ª¬ªpÀ³¤£·|µo¥Í, ¥H¨¾¥~¤@
               GoTo ErrHand
            End If
         End If
      End If
      If m_bolRecvOK = True Then '¬O§_¦¬§¹¤å=>¥þ³¡¦¬§¹¤å
         RetVal = RetVal & IIf(RetVal <> "", ",", "") & "m_bolRecvOK = True"
         '¦h®×¦¬¤åªºÁ`¦¬¤å¸¹­n¶Ç¤J²Ä¤@µ§Á`¦¬¤å¸¹
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, _
               IIf(m_strMCR11 <> "", "¦h®×¦¬¤å", "frm010001"), _
               IIf(m_strMCR11 <> "", m_strMCR11, mCP(9))
      End If
   End If
   '2023/5/31 END
   
   If bolError Then
      'Add By Sindy 2022/9/27
      If UCase(pFormName) <> UCase("frm090801_New") Then
      '2022/9/27 END
         cnnConnection.RollbackTrans
      End If
      ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
      IsSaveData = False
   Else
      'Add By Sindy 2022/9/27
      If UCase(pFormName) <> UCase("frm090801_New") Then
      '2022/9/27 END
         cnnConnection.CommitTrans
      End If
      InsertTrademarkDB = True
      '²¾¨ì¥~¼hÅÜ§ó
      ' If mTM(1) = "TF" Then
      '    txtTFCode(0) = Mid(tm02, 1, 5)
      '    txtTFCode(1) = Mid(tm02, 6, 1)
      ' Else
      '    mHC(2) = tm02
      ' End If
   End If
   
   Set rsQD = Nothing
   Set adoquery = Nothing
   
   Exit Function
ErrHand:
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.RollbackTrans
   End If
   'ShowMsg MsgText(9004)
   ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
   IsSaveData = False
   
   Set rsQD = Nothing
   Set adoquery = Nothing
End Function

'Added by Lydia 2022/09/05 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G­×§ï°Ó¼Ð°ò¥»ÀÉ(±qfrm010004.UpdateTrademarkDatabase©â¥X¨Ó)
Private Function UpdateTrademarkDB(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intChoose As Integer, _
                ByRef mTM() As String, ByRef mCP() As String, ByVal mCU30 As String, ByVal mSaveControl As String, Optional ByRef IsSaveData As Boolean) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'mSaveControl: »ô³Æ¤éºÞ¨î
Dim adoquery As New ADODB.Recordset

   If IsSaveData = True Then
       Exit Function
   End If
   IsSaveData = True
   
   On Error GoTo ErrHand
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.BeginTrans
   End If
   'Modify By Sindy 2022/12/7  + ", tm136=" + CNULL(mTM(136))
   mStrSql = "update trademark set tm05=" + CNULL(ChgSQL(mTM(5))) + _
      ",tm08=" + CNULL(mTM(8)) + ",tm09=" + CNULL(mTM(9)) + ",tm10=" + _
      CNULL(mTM(10)) + ",tm23=" + CNULL(mTM(23)) + ",tm44=" + CNULL(mTM(44)) + ",tm78=" + CNULL(mTM(78)) + ",tm79=" + CNULL(mTM(79)) + ",tm80=" + CNULL(mTM(80)) + ",tm81=" + CNULL(mTM(81)) + _
      ",tm32=" + CNULL(mTM(32)) + ",tm45=" + CNULL(ChgSQL(mTM(45))) + ", tm123=" + CNULL(mTM(123)) + ", tm136=" + CNULL(mTM(136))
   mStrSql = mStrSql & " where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
   cnnConnection.Execute mStrSql
   
   'Add By Sindy 2012/7/19 ­Y¥Ó½Ð¤H©Î¥N²z¤H¬°¿ÕµØ¤½¥qªÌ¡A®×¥ó³Æµù­YµL"¤£¾P¨÷"¦r¼Ë,«h­n¥[¤J
   If (mTM(23) <> "" And InStr(strTmNovartisCust, Left(mTM(23), 6)) > 0) Or _
      (mTM(78) <> "" And InStr(strTmNovartisCust, Left(mTM(78), 6)) > 0) Or _
      (mTM(79) <> "" And InStr(strTmNovartisCust, Left(mTM(79), 6)) > 0) Or _
      (mTM(80) <> "" And InStr(strTmNovartisCust, Left(mTM(80), 6)) > 0) Or _
      (mTM(81) <> "" And InStr(strTmNovartisCust, Left(mTM(81), 6)) > 0) Or _
      (mTM(44) <> "" And InStr(strTmNovartisCust, Left(mTM(44), 6)) > 0) Then
      mStrSql = "update trademark" & _
               " set tm58=decode(tm58,null,'" & ChangeTStringToTDateString(strSrvDate(2)) & "¤£¾P¨÷','" & ChangeTStringToTDateString(strSrvDate(2)) & "¤£¾P¨÷,'||tm58)" & _
               " Where tm01='" & mTM(1) & "' and tm02='" & mTM(2) & "' and tm03='" & mTM(3) & "' and tm04='" & mTM(4) & "'" & _
               " and (instr(tm58,'¤£¾P¨÷')=0 or tm58 is null)"
      cnnConnection.Execute mStrSql
   End If
   '2012/7/19 end
   'ADD BY SONIA 2015/11/24 ­Y¥Ó½Ð¬°®Èª°°ê»Ú¤Î³¡¤ÀÃö«Y¥ø·~ªÌ¡A®×¥ó³Æµù­YµL"¤£¾P¨÷"¦r¼Ë,«h­n¥[¤J
   If (mTM(23) <> "" And InStr(strTmTRAVEL_FOXCust, Left(mTM(23), 8)) > 0) Or _
      (mTM(78) <> "" And InStr(strTmTRAVEL_FOXCust, Left(mTM(78), 8)) > 0) Or _
      (mTM(79) <> "" And InStr(strTmTRAVEL_FOXCust, Left(mTM(79), 8)) > 0) Or _
      (mTM(80) <> "" And InStr(strTmTRAVEL_FOXCust, Left(mTM(80), 8)) > 0) Or _
      (mTM(81) <> "" And InStr(strTmTRAVEL_FOXCust, Left(mTM(81), 8)) > 0) Then
      mStrSql = "update trademark" & _
               " set tm58=decode(tm58,null,'" & ChangeTStringToTDateString(strSrvDate(2)) & "¤£¾P¨÷','" & ChangeTStringToTDateString(strSrvDate(2)) & "¤£¾P¨÷,'||tm58)" & _
               " Where tm01='" & mTM(1) & "' and tm02='" & mTM(2) & "' and tm03='" & mTM(3) & "' and tm04='" & mTM(4) & "'" & _
               " and (instr(tm58,'¤£¾P¨÷')=0 or tm58 is null)"
      cnnConnection.Execute mStrSql
   End If
   ''END 2015/11/24
   
   mStrSql = "update trademark set tm24=(select cu23 from customer where cu01=" + CNULL(Mid(mTM(23), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(23), 9, 1)) + _
      "),tm25=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(mTM(23), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(23), 9, 1)) + _
      "),tm26=(select cu29 from customer where cu01=" + CNULL(Mid(mTM(23), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(23), 9, 1)) + ") where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
   cnnConnection.Execute mStrSql
   'add by nickc 2006/11/30
   mStrSql = "update trademark set tm82=(select cu23 from customer where cu01=" + CNULL(Mid(mTM(78), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(78), 9, 1)) + _
      "),tm86=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(mTM(78), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(78), 9, 1)) + _
      "),tm90=(select cu29 from customer where cu01=" + CNULL(Mid(mTM(78), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(78), 9, 1)) + ") where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
   cnnConnection.Execute mStrSql
   mStrSql = "update trademark set tm83=(select cu23 from customer where cu01=" + CNULL(Mid(mTM(79), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(79), 9, 1)) + _
      "),tm87=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(mTM(79), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(79), 9, 1)) + _
      "),tm91=(select cu29 from customer where cu01=" + CNULL(Mid(mTM(79), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(79), 9, 1)) + ") where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
   cnnConnection.Execute mStrSql
   mStrSql = "update trademark set tm84=(select cu23 from customer where cu01=" + CNULL(Mid(mTM(80), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(80), 9, 1)) + _
      "),tm88=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(mTM(80), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(80), 9, 1)) + _
      "),tm92=(select cu29 from customer where cu01=" + CNULL(Mid(mTM(80), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(80), 9, 1)) + ") where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
   cnnConnection.Execute mStrSql
   mStrSql = "update trademark set tm85=(select cu23 from customer where cu01=" + CNULL(Mid(mTM(81), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(81), 9, 1)) + _
      "),tm89=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(mTM(81), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(81), 9, 1)) + _
      "),tm93=(select cu29 from customer where cu01=" + CNULL(Mid(mTM(81), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(81), 9, 1)) + ") where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
   cnnConnection.Execute mStrSql
   
   mStrSql = "Update Trademark Set TM34='" & ChgSQL(mTM(34)) & "', TM35='" & ChgSQL(mTM(35)) & "' Where tm01=" + CNULL(mTM(1)) + " and tm02=" + CNULL(mTM(2)) + " and tm03=" + CNULL(mTM(3)) + " and tm04=" + CNULL(mTM(4))
   cnnConnection.Execute mStrSql

   'Modify By Sindy 2012/11/06 +CP150  ¦³¡¹¡¹ªºÀ³¦¬±b´ÚÃ±®Ö±±ºÞ
   mStrSql = "update caseprogress set cp05=" + CNULL(mCP(5)) + ",cp06=" + CNULL(mCP(6)) + ",cp07=" + CNULL(mCP(7)) + ",cp10=" + CNULL(mCP(10)) + _
            ",cp11=" + CNULL(mCP(11)) + ",cp13=" + CNULL(mCP(13)) + ",cp14=" + CNULL(mCP(14)) + ",cp16=" + CNULL(mCP(16)) + ",cp17=" + CNULL(mCP(17)) + _
            ",cp18=" + CNULL(mCP(18)) + ",cp19=" + CNULL(mCP(19)) + ",cp32=" + CNULL(mCP(32)) + ",cp56=" + CNULL(mCP(56)) + ",cp33=" & CNULL(mCP(33)) & ",cp34=" & CNULL(mCP(34)) & ",CP64=" + CNULL(ChgSQL(mCP(64))) + _
            ",cp89=" + CNULL(mCP(89)) + ",cp90=" + CNULL(mCP(90)) + ",cp91=" + CNULL(mCP(91)) + ",cp92=" + CNULL(mCP(92)) + ",cp150=" & CNULL(mCP(150)) & " where cp09='" + mCP(9) + "'"
   cnnConnection.Execute mStrSql
   '2009/10/19 End
   
   'Modify By Sindy 2013/8/23 +cp118(¹q¤l°e¥ó)
   'mStrSQL = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(mcp(13)) + ") where cp09=" + CNULL(mcp(9))
   'Modified by Lydia 2022/09/20 ¦]¬°P®×¦³Trigger·|¦Û°Ê³]©w¹q¤l°e¥óCP118 = Y, ©Ò¥H§ï¦¨¨â­Ó§PÂ_
   'mStrSQL = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(mCP(13)) + ") " & ",cp118=" & CNULL(mCP(118)) & " where cp09=" + CNULL(mCP(9))
   mStrSql = IIf(mCP(118) = "YY", ", CP118='Y' ", IIf(mCP(118) = "YN", ", CP118=null ", ""))
   mStrSql = "update caseprogress set cp12=(select st15 from staff where st01=" & CNULL(mCP(13)) & ") " & mStrSql & " where cp09=" & CNULL(mCP(9))
   'end 2022/09/20
   cnnConnection.Execute mStrSql
           
   '­Y¬°±µ¬¢°O¿ý³æ(Âd¥x¦¬¤å)
   'Modify by Morgan 2007/10/26 ¶O¥Î¥i§ï®É¤~°µ¡A§_«h¤w¦¬´Ú¸ê®Æ·|³QÁÙ­ì
   If intChoose = 0 And mCP(60) = "" Then 'mCP(60) = "" => txtTrademark(14).Enabled = True
       '¥¼¦¬ª÷ÃB = ¶O¥Î
       mStrSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(mCP(9))
       cnnConnection.Execute mStrSql
   End If
           
   'Add By Cheng 2002/01/15
   '­Y¬°¤º³¡¦¬¤å, ¥u­n¨t²ÎÃþ§O¬°T¶}ÀY©ÎFCTªº®×¥ó
   '¥B®×¥ó©Ê½è¬°201¸É¥¿, 203­×¥¿, 302§ó¥¿, 305¶Ê¼f, 306¦Û½ÐºM¦^, 307¦Û½ÐºM¾P, 614¥¼¸É²z¥Ñ, 615¥¼µªÅG, 706¨ä¥L
   '§ì¨t²Î¤é´Á§ó·s¨ä®×¥ó¶i«×ÀÉªºµo¤å¤é(CP27)
   If intChoose = 1 Then
      If Left(mTM(1), 1) = "T" Or mTM(1) = "FCT" Then
         If mCP(10) = "201" Or mCP(10) = "203" Or mCP(10) = "302" Or mCP(10) = "305" Or _
            mCP(10) = "306" Or mCP(10) = "307" Or mCP(10) = "614" Or mCP(10) = "615" Or _
            mCP(10) = "706" Then
            mStrSql = "update caseprogress set cp27= '" & strSrvDate(1) & "' where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
         End If
      End If
   End If
   'Added by Lydia 2022/11/29 «D¤º³¡¦¬¤å¨Ã¥B¦³¶O¥Î¡A¥ý²Î¤@³]©wCP20=Null ;
   If intChoose = 0 And Val(mCP(16)) > 0 Then
       mStrSql = "update caseprogress set cp20=null where cp09=" + CNULL(mCP(9))
       cnnConnection.Execute mStrSql
   End If
   'end 2022/11/29
   
   '92.5.8 ADD BY SONIA
   If mTM(1) = "FCT" Then
      If Val(mCP(16)) = 0 Then
         mStrSql = "update caseprogress set cp20='N',CP32='N' where cp09=" + CNULL(mCP(9))
         cnnConnection.Execute mStrSql
      End If
   End If
   '92.5.8 END
   'Add By Cheng 2002/05/10
   '­Y¬°¤º³¡¦¬¤å§@·~®É, ®×¥ó¶i«×ÀÉªº¬O§_¦V«È¤á¦¬´Ú³]©w¬°"N"
   If intChoose = 1 Then
      mStrSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(mCP(9))
      cnnConnection.Execute mStrSql
   End If
   
   'Modify By Sindy 2023/11/3 mark:¦¬¤å¤w¤£»Ý°õ¦æ¦¹¬qµ{¦¡,¦]±µ¬¢³æ¤w·|¦^¼g¶l»¼°Ï¸¹
'   mStrSql = "update customer set cu30=" + CNULL(mCU30) + " where cu01=" + CNULL(Mid(mTM(23), 1, 8)) + " and cu02=" + CNULL(Mid(mTM(23), 9, 1))
'   cnnConnection.Execute mStrSql
   '2023/11/3 END
   
   UpdateTrademarkDB = True
   adoquery.CursorLocation = adUseClient
   'add by nickc 2007/10/24 ¤º°Ó¥þ´Áµù¥U¶O(717)®É¡A§ì²Ä¤@´Áµù¥U¶O(715)
   'edit by nickc 2007/10/25 ¥[¤J¥~°Ó
   If (mTM(1) = "T" Or mTM(1) = "FCT") And (mCP(10) = "717" Or mCP(10) = "715") Then
       adoquery.Open "select np01 from nextprogress where np02 = '" & mTM(1) & "' and np03 = '" & mTM(2) & "' and np04 = '" & mTM(3) & "' and np05 = '" & mTM(4) & "' and np06 is null and np07 in('715','717') ", cnnConnection, adOpenStatic, adLockReadOnly
   Else
       adoquery.Open "select np01 from nextprogress where np02 = '" & mTM(1) & "' and np03 = '" & mTM(2) & "' and np04 = '" & mTM(3) & "' and np05 = '" & mTM(4) & "' and np06 is null and np07 = '" & mCP(10) & "'", cnnConnection, adOpenStatic, adLockReadOnly
   End If
   'Modify By Cheng 2002/05/10
   '­Y¦b¤U¤@µ{§ÇÀÉ¥u§ì¨ì¤@µ§¸ê®Æ®É, ¤~­n§ì¤U¤@µ{§ÇÀÉªºÁ`¦¬¤å¸¹§ó·s®×¥ó¶i«×ÀÉªº¬ÛÃöÁ`¦¬¤å¸¹
   If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
      If IsNull(adoquery.Fields(0).Value) = False Then
         '2011/6/16 add by sonia ²§Ä³µªÅG¡Bµû©wµªÅG¡B¼o¤îµªÅG­n¤@¨Ã§ó·s¹ï³y¸ê®Æ
         If (mCP(10) = "602" Or mCP(10) = "604" Or mCP(10) = "606") Then
            cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp80) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42,b.cp80 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & mCP(9) & "'", intJ
         Else
         '2011/6/16 end
            cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & mCP(9) & "'"
         End If  '2011/6/16 add by sonia
      End If
   End If
   adoquery.Close
   
   ' Add by Sindy 98/03/02
   '¦¬¤å®É­YÅý®×¸¹¤U¤@µ{§Ç¤´¦³F4103¥B¬O§_Äò¿ì¬°NULLªÌ,
   '§ó·s¤U¤@µ{§ÇF4103¬°¦¬¤å´¼Åv¤H­û
   If mTM(1) = "FCT" Then
      mStrSql = "update nextprogress set np10='" & mCP(13) & "' " & _
         "where np02=" + CNULL(mTM(1)) + " and np03=" + _
         CNULL(mTM(2)) + " and np04=" + CNULL(mTM(3)) + " and np05=" + CNULL(mTM(4)) + _
         " and np10='F4103' and np06 is null"
      cnnConnection.Execute mStrSql
   End If
   ' 98/03/02 End
   
   'add by nickc 2008/05/02 Àx¦s¹w©w¦¬´Ú¤é
   'Remove by Lydia 2018/08/22 (À³¦¬±b´ÚºÞ±±)¨ú®ø¹w©w¦¬´Ú¤é,§ï¦¨¥I´Ú¶g´Á
'   Dim rtCnt As Integer
'   'Modify by Morgan 2010/12/9
'   'If txtTrademark(34) <> "" Then
'   '    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " ", rtCnt
'   If txtTrademark(34) <> "" And txtTrademark(34) <> txtTrademark(34).Tag Then
'       cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
'   'end 2010/12/9
'       If rtCnt = 0 Then
'           cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtTrademark(34)) & " from dual "
'       End If
'   End If
   'end 2018/08/22
   
   If mSaveControl <> "" Then
       'Modified by Lydia 2022/09/29 ¶Ç¤J¨t²Î§O,°ê®a,®×¥ó©Ê½è=>mTM(1), mTM(10), mCP(10)
       Call PUB_SaveByControl(mCP(9), mSaveControl, mTM(1), mTM(10), mCP(10))
   End If
   
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.CommitTrans
   End If
   Set adoquery = Nothing
   Exit Function
   
ErrHand:
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.RollbackTrans
   End If
   ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
   IsSaveData = False
   Set adoquery = Nothing
End Function

'Added by Lydia 2022/09/05 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G°Ó¼Ð¦¬¤å(±qfrm010004.SaveDatabase©â¥X¨Ó)
'Modify By Sindy 2023/5/31 + , Optional ByRef RetVal As String
'Modify By Sindy 2024/11/21 + , Optional ByVal m_intCRC As Integer = 0: ¦Û°Ê¦¬¤åªº®×¥ó©Ê½è¶¶§Ç
Public Function PUB_SaveFrm010004(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intChoose As Integer, _
                ByRef mTM() As String, ByRef mCP() As String, ByVal mCU30 As String, ByVal mSaveControl As String, Optional ByRef IsSaveData As Boolean, _
                Optional ByVal pType As String, Optional ByVal pCaseNo As String, Optional ByRef RetVal As String, _
                Optional ByVal m_intCRC As Integer = 0) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
'mSaveControl: »ô³Æ¤éºÞ¨î
Dim adoquery As New ADODB.Recordset
Dim oMailCount As String
Dim m_SalesST15 As String, m_SalesST06 As String
Dim m_CP10Name As String '¦¬¤å¤§®×¥ó©Ê½è¦WºÙ
Dim m_Na01Name As String '¥Ó½Ð°ê®a¦WºÙ
Dim m_SalesDeptName As String
Dim m_CaseNaTmp() As String  '¯S®íºÞ¨î¤§ÃöÁp®×
'ªk«ß©Ò®×·½¦¬¤å
Dim m_LOS02 As String '®×·½®×¥óÃþ«¬
Dim m_LOS15 As String '®×·½³æ¸¹
'Add By Sindy 2022/9/27
Dim bolIsOverDt As Boolean, strDivisionalEmp As String
Dim strRDate As String, strRTime As String, strRestKind As String
Dim bolCP14Rest As Boolean
'2022/9/27 END
   
'*********¯S®íºÞ¨îªºÅÜ¼Æ*************
   If pType = "CFT­^°ê²æ¼Ú®×" And pCaseNo <> "" Then
       ReDim m_CaseNaTmp(1 To TF_TM)
       Call ChgCaseNo(pCaseNo, m_CaseNaTmp)
   Else
       ReDim m_CaseNaTmp(1 To 4) '¹w³]°}¦CÁ×§Kµ{¦¡¥X¿ù
       'Modify By Sindy 2025/8/18
       'If pType = "LOS®×·½¦¬¤å" And pCaseNo <> "" Then
       If InStr(pType, "LOS®×·½¦¬¤å") > 0 And pCaseNo <> "" Then
       '2025/8/18 END
           m_LOS02 = Mid(pCaseNo, 1, InStr(pCaseNo, ",") - 1) '®×·½®×¥óÃþ«¬
           m_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1, 8) '®×·½³æ¸¹ 'Modify By Sindy 2025/8/18 +, 8)
       ElseIf pType = "CFT½q¨l­«·s¥Ó½Ð®×" And pCaseNo <> "" Then
           Call ChgCaseNo(pCaseNo, m_CaseNaTmp)
       End If
   End If
'***********************************
   If mTM(1) = °¨¼w¨½®× Then
       intJ = ClsPDGetCaseProperty(mTM(1), mCP(10), m_CP10Name)
   Else
       intJ = ClsPDGetCaseProperty(mTM(1), mCP(10), m_CP10Name, IIf(mTM(10) <> "000", True, False))
   End If
   m_Na01Name = PUB_GetNationName(mTM(10))
   m_SalesST15 = GetST15(mCP(13), m_SalesDeptName, , m_SalesST06)
   'Added by Lydia 2023/05/11 ¦]¬°PUB_ReadCaseData·|¦^¶Ç6½X«È¤á½s¸¹,©Ò¥H¥ý²Î¤@«È¤á½s¸¹
   mTM(23) = ChangeCustomerL(mTM(23))
   mTM(78) = ChangeCustomerL(mTM(78))
   mTM(79) = ChangeCustomerL(mTM(79))
   mTM(80) = ChangeCustomerL(mTM(80))
   mTM(81) = ChangeCustomerL(mTM(81))
   'end 2023/05/11
   
   RetVal = "" '¦^¶Ç­È Add By Sindy 2023/5/31
   
   If intModifyKind = 0 Then
       'Modify By Sindy 2023/5/31 + RetVal
       PUB_SaveFrm010004 = InsertTrademarkDB(pFormName, intSaveMode, intModifyKind, intChoose, mTM, mCP, mCU30, mSaveControl, IsSaveData, pType, pCaseNo, RetVal)
   Else
       PUB_SaveFrm010004 = UpdateTrademarkDB(pFormName, intSaveMode, intModifyKind, intChoose, mTM, mCP, mCU30, mSaveControl, IsSaveData)
   End If
   If PUB_SaveFrm010004 = False Then Exit Function 'Add By Sindy 2022/9/28 ¦sÀÉ¥¢±Ñ,«áÄò¤£ÀË¬d
   'add by nickc 2007/11/09 ´ú¸Õ¸Ñ¨Mmail µo¤£¨ìªº®É­Ô·|¦s¨âµ§ªº¿ù»~
   On Error GoTo 0 'Âk¹s
   On Error GoTo ErrHand 'Add By Sindy 2022/9/28
   'Add By Sindy 2022/12/29 ­«ÅªCP,¦]«eÀYUpdate¨ç¼Æµ{¦¡¦³¥i¯àª½±µ¦sDB,¨S¦³§ó·scp³¯¦C­È
   strTmp1(0) = mCP(9)
   Erase mCP
   ReDim Preserve mCP(TF_CP) As String
   mCP(9) = strTmp1(0)
   'Modified by Lydia 2023/05/11 + false
   Call PUB_ReadCaseProgressDatabase(mCP(), 1, False)
   '2022/12/29 END
   
   'Add By Sindy 2020/2/7
   '¥~°Ó¦¬¤å®É, ®×¥ó¦³ FC¥N²z¤H, ¥B¥N²z¤H°êÄy¬°¤é¥»®É,
   '¤£½×¬O°Ó¼Ð®×¥ó©ÎªA°È·~°È®×¥ó, ¦sÀÉ®É©w½Z»y¤åÄæ­Y¬°ªÅ­È®É,
   '¤@«ß§ó·s¬°3.¤é¤å
   If mTM(44) <> "" And (mTM(1) = "FCT" Or mTM(1) = "T" Or mTM(1) = "S") Then
      mStrSql = "SELECT fa01,fa02,fa10 FROM fagent" & _
               " WHERE fa01=" & CNULL(Left(mTM(44), 8)) & _
               " and fa02=" & CNULL(Mid(mTM(44), 9, 1))
      adoquery.CursorLocation = adUseClient
      adoquery.Open mStrSql, cnnConnection, adOpenStatic, adLockReadOnly
      If adoquery.RecordCount > 0 Then
         If Left("" & adoquery.Fields("fa10"), 3) = "011" Then '¤é¥»
            mStrSql = "UPDATE TradeMark SET TM53='3'" & _
                     " WHERE TM01 = '" & mTM(1) & "' AND TM02 = '" & mTM(2) & "' AND TM03 = '" & mTM(3) & "' AND TM04 = '" & mTM(4) & "' AND TM53 is null"
            cnnConnection.Execute mStrSql
         End If
      End If
      adoquery.Close
   End If
   '2020/2/7 END
   
   'Added by Lydia 2021/01/08 CFT­^°ê²æ¼Ú®×¡G¤@¨Ö½Æ»s°Ó¼Ð¹Ï¡B«ü©w°Ó«~©ÎªA°È¦WºÙ(InsertTrademarkDatabase)
   'Memo by Lydia 2021/03/05 §tCFT¼Ú·ù©|¥¼µù¥U®×Âà´«­^°ê¥Ó½Ð®×¦¬¤å±±ºÞ
   If mTM(1) = "CFT" And mCP(31) = "Y" And pType = "CFT­^°ê²æ¼Ú®×" And m_CaseNaTmp(1) <> "" And m_CaseNaTmp(2) <> "" Then
      strTmp1(9) = ""
      If GetImgByteFile_Case(m_CaseNaTmp(1), m_CaseNaTmp(2), m_CaseNaTmp(3), m_CaseNaTmp(4), strTmp1(9), 0, strTmp1(5), strTmp1(6)) = True Then
          Call SaveImgByteFile(strTmp1(9), mTM(1), mTM(2), mTM(3), mTM(4), strTmp1(5), strTmp1(6))
      End If
   End If
   'end 2021/01/08
   
   'Added by Lydia 2021/02/01 CFT½q¨l­«·s¥Ó½Ð®×¡G½q¨l°Ó¼Ð­«·s¥Ó½Ð¦¬¤å®É¡A¤@¨Ö±NÂÂ®×¤§¡u°Ó¼Ð¹Ï¼Ë¡v¡B¡u°Ó«~/ªA°ÈÃþ§O¤Î¦WºÙ¡v¡B¡uÀu¥ýÅv¸ê®Æ¡v±a¤J·s®×¸¹
   If mTM(1) = "CFT" And mTM(10) = "048" And mCP(31) = "Y" And pType = "CFT½q¨l­«·s¥Ó½Ð®×" And mCP(10) = "101" And m_CaseNaTmp(1) <> "" And m_CaseNaTmp(2) <> "" Then
      strTmp1(9) = ""
      If GetImgByteFile_Case(m_CaseNaTmp(1), m_CaseNaTmp(2), m_CaseNaTmp(3), m_CaseNaTmp(4), strTmp1(9), 0, strTmp1(5), strTmp1(6)) = True Then
          Call SaveImgByteFile(strTmp1(9), mTM(1), mTM(2), mTM(3), mTM(4), strTmp1(5), strTmp1(6))
      End If
   End If
   'end 2021/02/01
   
   '¬d¦W³æ¹ïÀ³¦sÀÉ
   If pType = "T¬d¦W³æ" And pCaseNo <> "" Then
      strTmp1(1) = Mid(pCaseNo, 1, InStr(pCaseNo, "|") - 1)
      strTmp1(2) = Mid(pCaseNo, InStr(pCaseNo, "|") + 1)
      'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
      'Modified by Lydia 2024/03/14 +Fasle
      'If PUB_TMQtoCP("", mCP(9), strTmp1(2), strTmp1(1), , , IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)) = False Then
      If PUB_TMQtoCP(False, "", mCP(9), strTmp1(2), strTmp1(1), , , IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)) = False Then
      End If
   End If
      
'add by nickc 2005/09/05
If intModifyKind = 0 Then
   Dim oContext As String, strCaseNo As String
   Dim strTemp As String
   Dim m_strState As String
   
   'Add By Sindy 2021/2/1 ¤£±o¥N²zªº«áÄòÂÂ®×¦¬¤å±±ºÞ¡A³qª¾¦¬¤å¤H­û¡]CP13¡^
   If mTM(44) <> "" Then
     If GetAgentAndState(mTM(44), strTmp1(1), , , , mTM(1), m_strState, IIf(intSaveMode = 0, True, False)) Then
       If InStr(m_strState, "¤£±o¥N²z") > 0 Then
          oContext = oContext & vbCrLf + "¥N²z¤H¡G " + mTM(44) + " " + strTmp1(1) + vbCrLf
          strTemp = strTemp & "," & mTM(44)
       End If
     End If
   End If
   If mTM(23) <> "" Then
     'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , mTM(2), mTM(3), mTM(4)
     If GetCustomerAndState(mTM(23), strTmp1(1), , , , mTM(1), m_strState, IIf(intSaveMode = 0, True, False), , mTM(2), mTM(3), mTM(4)) Then
       If InStr(m_strState, "¤£±o¥N²z") > 0 Then
          oContext = oContext & vbCrLf + "¥Ó½Ð¤H1¡G " + mTM(23) + " " + strTmp1(1) + vbCrLf
          strTemp = strTemp & "," & mTM(23)
       End If
     End If
   End If
   If mTM(78) <> "" Then
     'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , mTM(2), mTM(3), mTM(4)
     If GetCustomerAndState(mTM(78), strTmp1(1), , , , mTM(1), m_strState, IIf(intSaveMode = 0, True, False), , mTM(2), mTM(3), mTM(4)) Then
       If InStr(m_strState, "¤£±o¥N²z") > 0 Then
          oContext = oContext & vbCrLf + "¥Ó½Ð¤H2¡G " + mTM(78) + " " + strTmp1(1) + vbCrLf
          strTemp = strTemp & "," & mTM(78)
       End If
     End If
   End If
   If mTM(79) <> "" Then
     'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , mTM(2), mTM(3), mTM(4)
     If GetCustomerAndState(mTM(79), strTmp1(1), , , , mTM(1), m_strState, IIf(intSaveMode = 0, True, False), , mTM(2), mTM(3), mTM(4)) Then
       If InStr(m_strState, "¤£±o¥N²z") > 0 Then
          oContext = oContext & vbCrLf + "¥Ó½Ð¤H3¡G " + mTM(79) + " " + strTmp1(1) + vbCrLf
          strTemp = strTemp & "," & mTM(79)
       End If
     End If
   End If
   If mTM(80) <> "" Then
     'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , mTM(2), mTM(3), mTM(4)
     If GetCustomerAndState(mTM(80), strTmp1(1), , , , mTM(1), m_strState, IIf(intSaveMode = 0, True, False), , mTM(2), mTM(3), mTM(4)) Then
       If InStr(m_strState, "¤£±o¥N²z") > 0 Then
          oContext = oContext & vbCrLf + "¥Ó½Ð¤H4¡G " + mTM(80) + " " + strTmp1(1) + vbCrLf
          strTemp = strTemp & "," & mTM(80)
       End If
     End If
   End If
   If mTM(81) <> "" Then
     'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , mTM(2), mTM(3), mTM(4)
     If GetCustomerAndState(mTM(81), strTmp1(1), , , , mTM(1), m_strState, IIf(intSaveMode = 0, True, False), , mTM(2), mTM(3), mTM(4)) Then
       If InStr(m_strState, "¤£±o¥N²z") > 0 Then
          oContext = oContext & vbCrLf + "¥Ó½Ð¤H5¡G " + mTM(81) + " " + strTmp1(1) + vbCrLf
          strTemp = strTemp & "," & mTM(81)
       End If
     End If
   End If
   If mCP(56) <> "" Then
     'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , mTM(2), mTM(3), mTM(4)
     'Modify By Sindy 2025/3/26 ¶Ç¤J®×¥ó©Ê½è , mCP(10)
     If GetCustomerAndState(mCP(56), strTmp1(1), , , , mTM(1), m_strState, IIf(intSaveMode = 0, True, False), , mTM(2), mTM(3), mTM(4), mCP(10)) Then
       If InStr(m_strState, "¤£±o¥N²z") > 0 Then
          oContext = oContext & vbCrLf + "²¾Âà¥Ó½Ð¤H1¡G " + mCP(56) + " " + strTmp1(1) + vbCrLf
          strTemp = strTemp & "," & mCP(56)
       End If
     End If
   End If
   If mCP(89) <> "" Then
     'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , mTM(2), mTM(3), mTM(4)
     'Modify By Sindy 2025/3/26 ¶Ç¤J®×¥ó©Ê½è , mCP(10)
     If GetCustomerAndState(mCP(89), strTmp1(1), , , , mTM(1), m_strState, IIf(intSaveMode = 0, True, False), , mTM(2), mTM(3), mTM(4), mCP(10)) Then
       If InStr(m_strState, "¤£±o¥N²z") > 0 Then
          oContext = oContext & vbCrLf + "²¾Âà¥Ó½Ð¤H2¡G " + mCP(89) + " " + strTmp1(1) + vbCrLf
          strTemp = strTemp & "," & mCP(89)
       End If
     End If
   End If
   If mCP(90) <> "" Then
     'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , mTM(2), mTM(3), mTM(4)
     'Modify By Sindy 2025/3/26 ¶Ç¤J®×¥ó©Ê½è , mCP(10)
     If GetCustomerAndState(mCP(90), strTmp1(1), , , , mTM(1), m_strState, IIf(intSaveMode = 0, True, False), , mTM(2), mTM(3), mTM(4), mCP(10)) Then
       If InStr(m_strState, "¤£±o¥N²z") > 0 Then
          oContext = oContext & vbCrLf + "²¾Âà¥Ó½Ð¤H3¡G " + mCP(90) + " " + strTmp1(1) + vbCrLf
          strTemp = strTemp & "," & mCP(90)
       End If
     End If
   End If
   If mCP(91) <> "" Then
     'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , mTM(2), mTM(3), mTM(4)
     'Modify By Sindy 2025/3/26 ¶Ç¤J®×¥ó©Ê½è , mCP(10)
     If GetCustomerAndState(mCP(91), strTmp1(1), , , , mTM(1), m_strState, IIf(intSaveMode = 0, True, False), , mTM(2), mTM(3), mTM(4), mCP(10)) Then
       If InStr(m_strState, "¤£±o¥N²z") > 0 Then
          oContext = oContext & vbCrLf + "²¾Âà¥Ó½Ð¤H4¡G " + mCP(91) + " " + strTmp1(1) + vbCrLf
          strTemp = strTemp & "," & mCP(91)
       End If
     End If
   End If
   If mCP(92) <> "" Then
     'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , mTM(2), mTM(3), mTM(4)
     'Modify By Sindy 2025/3/26 ¶Ç¤J®×¥ó©Ê½è , mCP(10)
     If GetCustomerAndState(mCP(91), strTmp1(1), , , , mTM(1), m_strState, IIf(intSaveMode = 0, True, False), , mTM(2), mTM(3), mTM(4), mCP(10)) Then
       If InStr(m_strState, "¤£±o¥N²z") > 0 Then
          oContext = oContext & vbCrLf + "²¾Âà¥Ó½Ð¤H5¡G " + mCP(92) + " " + strTmp1(1) + vbCrLf
          strTemp = strTemp & "," & mCP(92)
       End If
     End If
   End If
   If oContext <> "" Then
      strTemp = Mid(strTemp, 2)
      If mTM(1) = °¨¼w¨½®× Then
         strCaseNo = mTM(1) + "-" + mTM(2) + "-" + mTM(3) + "-" + mTM(4)
         oContext = "¥»©Ò®×¸¹¡G " + strCaseNo + vbCrLf + _
                    "®×¥ó¦WºÙ¡G " + mTM(5) + vbCrLf + _
                    "¥Ó½Ð°ê®a¡G " + mTM(10) + " " + m_Na01Name + vbCrLf + _
                    "¦¬¤å¤é¡G " + ChangeTStringToTDateString(TransDate(mCP(5), 1)) + vbCrLf + _
                    "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf + vbCrLf + _
                    "¡i¤£±o¥N²z¡j" + vbCrLf + _
                    oContext
      Else
         strCaseNo = IIf("-" + mTM(3) + "-" + mTM(4) = "-0-00", mTM(1) + "-" + mTM(2), mTM(1) + "-" + mTM(2) + "-" + mTM(3) + "-" + mTM(4))
         oContext = "¥»©Ò®×¸¹¡G " + strCaseNo + vbCrLf + _
                    "®×¥ó¦WºÙ¡G " + mTM(5) + vbCrLf + _
                    "¥Ó½Ð°ê®a¡G " + mTM(10) + " " + m_Na01Name + vbCrLf + _
                    "¦¬¤å¤é¡G " + ChangeTStringToTDateString(TransDate(mCP(5), 1)) + vbCrLf + _
                    "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf + vbCrLf + _
                    "¡i¤£±o¥N²z¡j" + vbCrLf + _
                    oContext
      End If
      oMailCount = mCP(13) & ";" & PUB_GetFCPProSup(mCP(13), True)
'      PUB_SendMail strUserNum, oMailCount, "", strCaseNo & _
'         " ¤w½T»{Äò¦æ¦¬¤å¡A½Ðª`·N¸Ó" & strTemp & "½s¸¹¤w³]¬°¤£±o¥N²z¡C", oContext
      'Modify By Sindy 2022/9/29
      'Modify By Sindy 2023/3/27 +,mc13
      mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
         " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
         ",'" & strCaseNo & _
         " ¤w½T»{Äò¦æ¦¬¤å¡A½Ðª`·N¸Ó" & strTemp & "½s¸¹¤w³]¬°¤£±o¥N²z¡C(¤å¸¹:" & mCP(9) & ")','" & oContext & "',null,'" & mCP(9) & "')"
      cnnConnection.Execute mStrSql
      '2022/9/29 END
   End If
   '2021/2/1 END
  
   'add by nickc 2007/05/16 ¥[¤J­Y¬O¥»©Ò´Á­­¤p©óµ¥©ó·í¤Ñ¡A­nµomail  ³qª¾
   Dim oContext2 As String
   oContext = "": oContext2 = ""
   If mTM(1) = °¨¼w¨½®× Then
     oContext = "¥»©Ò®×¸¹¡G " + mTM(1) + "-" + mTM(2) + "-" + mTM(3) + "-" + mTM(4) + vbCrLf + "®×¥ó¦WºÙ¡G " + mTM(5) + vbCrLf + "¦¬¤å¤é¡G " + ChangeTStringToTDateString(TransDate(mCP(5), 1)) + vbCrLf + "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf
     'add by nickc 2007/05/16 ¥[¤J­Y¬O¥»©Ò´Á­­¤p©óµ¥©ó·í¤Ñ¡A­nµomail  ³qª¾
     oContext2 = "¥»©Ò®×¸¹¡G " + mTM(1) + "-" + mTM(2) + "-" + mTM(3) + "-" + mTM(4) + vbCrLf + "®×¥ó¦WºÙ¡G " + mTM(5) + vbCrLf + "¥Ó½Ð°ê®a¡G" + mTM(10) + " " + m_Na01Name + vbCrLf + "¦¬¤å¤é¡G " + ChangeTStringToTDateString(TransDate(mCP(5), 1)) + vbCrLf + "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf
   Else
     oContext = "¥»©Ò®×¸¹¡G " + mTM(1) + "-" + mTM(2) + "-" + mTM(3) + "-" + mTM(4) + vbCrLf + "®×¥ó¦WºÙ¡G " + mTM(5) + vbCrLf + "¦¬¤å¤é¡G " + ChangeTStringToTDateString(TransDate(mCP(5), 1)) + vbCrLf + "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf
     'add by nickc 2007/05/16 ¥[¤J­Y¬O¥»©Ò´Á­­¤p©óµ¥©ó·í¤Ñ¡A­nµomail  ³qª¾
     oContext2 = "¥»©Ò®×¸¹¡G " + mTM(1) + "-" + mTM(2) + "-" + mTM(3) + "-" + mTM(4) + vbCrLf + "®×¥ó¦WºÙ¡G " + mTM(5) + vbCrLf + "¥Ó½Ð°ê®a¡G" + mTM(10) + " " + m_Na01Name + vbCrLf + "¦¬¤å¤é¡G " + ChangeTStringToTDateString(TransDate(mCP(5), 1)) + vbCrLf + "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf
   End If
   
   'Modify By Sindy 2024/11/6 §ï¦¨¦@¥Î¨ç¼Æ: ¦¬¤å®É,ÀË¬d¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¬O§_¦³»~
   '§ï¼g­ì¥Ñ¬O¦]¬°¥Ó½Ð¤H1~5 ³v¤@ÀË¬d,¦³»~§¡­nµo mail
   'edit by nickc 2007/08/21 ­Y¥Ó½Ð¤H¥þªÅ¥Õ¡A¤£µo
   If Not (mTM(23) = "" And mTM(78) = "" And mTM(79) = "" And mTM(80) = "" And mTM(81) = "") Then
      'Modify By Sindy 2024/11/21 ¦Û°Ê¦¬¤åªº®×¥ó©Ê½è¶¶§Ç=1 ©Î¯È¥»¦¬¤å¥¼«ü©w
      If m_intCRC = 1 Or m_intCRC = 0 Then
      '2024/11/21 END
         Call RecvChkApplCust("¥Ó½Ð¤H1", mTM(23), mCP(13), Trim(mTM(44)), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
         Call RecvChkApplCust("¥Ó½Ð¤H2", mTM(78), mCP(13), Trim(mTM(44)), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
         Call RecvChkApplCust("¥Ó½Ð¤H3", mTM(79), mCP(13), Trim(mTM(44)), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
         Call RecvChkApplCust("¥Ó½Ð¤H4", mTM(80), mCP(13), Trim(mTM(44)), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
         Call RecvChkApplCust("¥Ó½Ð¤H5", mTM(81), mCP(13), Trim(mTM(44)), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
      End If
   End If
   '2024/11/6 END
'Modify By Sindy 2024/11/6 mark
'Dim oStrCuSales1 As String
'Dim oStrCuSales2 As String
'Dim oStrCuSales3 As String
'Dim oStrCuSales4 As String
'Dim oStrCuSales5 As String
''Dim oContext As String
''add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'Dim IsMail As Boolean
'
'   IsMail = True
'   oStrCuSales1 = ""
'   oStrCuSales2 = ""
'   oStrCuSales3 = ""
'   oStrCuSales4 = ""
'   oStrCuSales5 = ""
'
'   oMailCount = ""
'
'   'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'   'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'   'Modify By Sindy 2023/2/2 +, , oStrCuSales1 : ¦^¶Ç­ì´¼Åv¤H­û
'   If ChkSameCuArea(mTM(23), mCP(13), , , , , Trim(mTM(44)), , oStrCuSales1) = False And mCP(13) <> "" And mTM(23) <> "" Then
'      'Add By Sindy 2009/10/19
'      'Modified by Lydia 2019/02/14
'      'Modify By Sindy 2023/2/2
'      'If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mTM(23), oStrCuSales1)), 1) = "F" Then
'      If Left(m_SalesST15, 1) = "F" And Left(GetSalesArea(oStrCuSales1), 1) = "F" Then
'      '2023/2/2 END
'         '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'      Else
'         oMailCount = oMailCount & oStrCuSales1 & ";"
'         'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'         If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales1), 1) = "S" And _
'            InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'            oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'         End If
'         '2023/11/7 END
'         oContext = oContext & vbCrLf + "¥Ó½Ð¤H1¡G " + GetCustomerName(mTM(23)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales1)
'      End If
'   'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'   Else
'        If mCP(13) <> "" And mTM(23) <> "" Then
'            IsMail = False
'        End If
'   End If
'   'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'   If m_SalesST06 <> "" And mTM(23) <> "" And mCP(13) <> "" Then
'       'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'       'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'       If PUB_ChkOldCustomer(True, mTM(23), mCP(13), m_SalesST15, m_SalesST06, _
'               IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'           IsMail = False
'       End If
'   End If
'
'   'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'   'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'   'Modify By Sindy 2023/2/2 +, , oStrCuSales2 : ¦^¶Ç­ì´¼Åv¤H­û
'   If ChkSameCuArea(mTM(78), mCP(13), , , , , Trim(mTM(44)), , oStrCuSales2) = False And mCP(13) <> "" And mTM(78) <> "" Then
'      'Add By Sindy 2009/10/19
'      'Modified by Lydia 2019/02/14
'      'Modify By Sindy 2023/2/2
'      'If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mTM(78), oStrCuSales2)), 1) = "F" Then
'      If Left(m_SalesST15, 1) = "F" And Left(GetSalesArea(oStrCuSales2), 1) = "F" Then
'      '2023/2/2 END
'         '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'      Else
'         oMailCount = oMailCount & oStrCuSales2 & ";"
'         'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'         If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales2), 1) = "S" And _
'            InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'            oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'         End If
'         '2023/11/7 END
'         oContext = oContext & vbCrLf + "¥Ó½Ð¤H2¡G " + GetCustomerName(mTM(78)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales2)
'      End If
'   'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'   Else
'        If mCP(13) <> "" And mTM(78) <> "" Then
'            IsMail = False
'        End If
'   End If
'   'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'   If m_SalesST06 <> "" And mTM(78) <> "" And mCP(13) <> "" Then
'       'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'       'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'       If PUB_ChkOldCustomer(True, mTM(78), mCP(13), m_SalesST15, m_SalesST06, _
'               IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'           IsMail = False
'       End If
'   End If
'
'   'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'   'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'   'Modify By Sindy 2023/2/2 +, , oStrCuSales3 : ¦^¶Ç­ì´¼Åv¤H­û
'   If ChkSameCuArea(mTM(79), mCP(13), , , , , Trim(mTM(44)), , oStrCuSales3) = False And mCP(13) <> "" And mTM(79) <> "" Then
'      'Add By Sindy 2009/10/19
'      'Modified by Lydia 2019/02/14
'      'Modify By Sindy 2023/2/2
'      'If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mTM(79), oStrCuSales3)), 1) = "F" Then
'      If Left(m_SalesST15, 1) = "F" And Left(GetSalesArea(oStrCuSales3), 1) = "F" Then
'      '2023/2/2 END
'         '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'      Else
'         oMailCount = oMailCount & oStrCuSales3 & ";"
'         'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'         If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales3), 1) = "S" And _
'            InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'            oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'         End If
'         '2023/11/7 END
'         oContext = oContext & vbCrLf + "¥Ó½Ð¤H3¡G " + GetCustomerName(mTM(79)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales3)
'      End If
'   'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'   Else
'        If mCP(13) <> "" And mTM(79) <> "" Then
'            IsMail = False
'        End If
'   End If
'   'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'   If m_SalesST06 <> "" And mTM(79) <> "" And mCP(13) <> "" Then
'       'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'       'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'       If PUB_ChkOldCustomer(True, mTM(79), mCP(13), m_SalesST15, m_SalesST06, _
'               IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'           IsMail = False
'       End If
'   End If
'
'   'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'   'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'   'Modify By Sindy 2023/2/2 +, , oStrCuSales4 : ¦^¶Ç­ì´¼Åv¤H­û
'   If ChkSameCuArea(mTM(80), mCP(13), , , , , Trim(mTM(44)), , oStrCuSales4) = False And mCP(13) <> "" And mTM(80) <> "" Then
'      'Add By Sindy 2009/10/19
'      'Modified by Lydia 2019/02/14
'      'Modify By Sindy 2023/2/2
'      'If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mTM(80), oStrCuSales4)), 1) = "F" Then
'      If Left(m_SalesST15, 1) = "F" And Left(GetSalesArea(oStrCuSales4), 1) = "F" Then
'      '2023/2/2 END
'         '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'      Else
'         oMailCount = oMailCount & oStrCuSales4 & ";"
'         'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'         If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales4), 1) = "S" And _
'            InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'            oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'         End If
'         '2023/11/7 END
'         oContext = oContext & vbCrLf + "¥Ó½Ð¤H4¡G " + GetCustomerName(mTM(80)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales4)
'      End If
'   'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'   Else
'        If mCP(13) <> "" And mTM(80) <> "" Then
'            IsMail = False
'        End If
'   End If
'   'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'   If m_SalesST06 <> "" And mTM(80) <> "" And mCP(13) <> "" Then
'       'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'       'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'       If PUB_ChkOldCustomer(True, mTM(80), mCP(13), m_SalesST15, m_SalesST06, _
'               IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'           IsMail = False
'       End If
'   End If
'
'   'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'   'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'   'Modify By Sindy 2023/2/2 +, , oStrCuSales5 : ¦^¶Ç­ì´¼Åv¤H­û
'   If ChkSameCuArea(mTM(81), mCP(13), , , , , Trim(mTM(44)), , oStrCuSales5) = False And mCP(13) <> "" And mTM(81) <> "" Then
'      'Add By Sindy 2009/10/19
'      'Modified by Lydia 2019/02/14
'      'If Left(Trim(GetST15(txtTrademark(12).Text)), 1) = "F" And Left(Trim(GetCuSales(mTM(81), oStrCuSales5)), 1) = "F" Then
'      'Modify By Sindy 2023/2/2
'      'If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mTM(81), oStrCuSales5)), 1) = "F" Then
'      If Left(m_SalesST15, 1) = "F" And Left(GetSalesArea(oStrCuSales5), 1) = "F" Then
'      '2023/2/2 END
'         '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'      Else
'         oMailCount = oMailCount & oStrCuSales5 & ";"
'         'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'         If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales5), 1) = "S" And _
'            InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'            oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'         End If
'         '2023/11/7 END
'         oContext = oContext & vbCrLf + "¥Ó½Ð¤H5¡G " + GetCustomerName(mTM(81)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales5)
'      End If
'   'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'   Else
'        If mCP(13) <> "" And mTM(81) <> "" Then
'            IsMail = False
'        End If
'   End If
'   'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'   If m_SalesST06 <> "" And mTM(81) <> "" And mCP(13) <> "" Then
'       'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'       'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'       If PUB_ChkOldCustomer(True, mTM(81), mCP(13), m_SalesST15, m_SalesST06, _
'               IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'           IsMail = False
'       End If
'   End If
'
'   'edit by nickc 2007/08/21 ­Y¥Ó½Ð¤H¥þªÅ¥Õ¡A¤£µo
'   If IsMail = False Or (mTM(23) = "" And mTM(78) = "" And mTM(79) = "" And mTM(80) = "" And mTM(81) = "") Then
'        oMailCount = ""
'   End If
'
'   If UCase(Mid(mTM(1), 1, 1)) <> "F" And oMailCount <> "" Then
'      'Modify By Sindy 2010/11/26 ¥Ó½Ð¤H1~5¬° X65299 ©Î X03072 ªº©Ò¦³Ãö«Y¥ø·~³£¤£ÀË¬d·~°È°Ï
'      If Left(mTM(23), 6) <> "X65299" And Left(mTM(23), 6) <> "X03072" And _
'         Left(mTM(78), 6) <> "X65299" And Left(mTM(78), 6) <> "X03072" And _
'         Left(mTM(79), 6) <> "X65299" And Left(mTM(79), 6) <> "X03072" And _
'         Left(mTM(80), 6) <> "X65299" And Left(mTM(80), 6) <> "X03072" And _
'         Left(mTM(81), 6) <> "X65299" And Left(mTM(81), 6) <> "X03072" Then
'         'Modify By Sindy 2022/9/27
'         If UCase(pFormName) <> UCase("frm090801_New") Then
'         '2022/9/27 END
'            MsgBox "¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¤£¦P·~°È°Ï¡A·Ç³Æµo mail ¡I", , "ª`·N¡I"
'         End If
'         'Modify By Sindy 2022/9/29 §ï§ì Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'         oMailCount = oMailCount & mCP(13) & ";" & Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'         oContext = oContext & vbCrLf + "¦¬¤å´¼Åv¤H­û¡G " + GetStaffName(mCP(13)) + vbCrLf + vbCrLf + "´¼Åv¤H­û(°Ï)¤£¦P¡I"
''         PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å³qª¾--¦¹®×¦¬¤å«D­ì´¼Åv¤H­û(°Ï)¡I", oContext
'         'Modify By Sindy 2022/9/29
'         'Modified by Lydia 2022/12/23 +chgsql
'         'Modify By Sindy 2023/3/27 +,mc13
'         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
'            " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'            ",'" & "®×¥ó¦¬¤å³qª¾--¦¹®×¦¬¤å«D­ì´¼Åv¤H­û(°Ï)¡I(¤å¸¹:" & mCP(9) & ")','" & ChgSQL(oContext) & "',null,'" & mCP(9) & "')"
'         cnnConnection.Execute mStrSql
'         '2022/9/29 END
'      End If
'   End If
'2024/11/6 mark END
   
   'add by nickc 2007/05/16 ¥[¤J­Y¬O¥»©Ò´Á­­¤p©óµ¥©ó·í¤Ñ¡A­nµomail  ³qª¾
   oMailCount = ""
   If Mid(mTM(1), 1, 1) = "T" Then   'T,TF®×
        oMailCount = Pub_GetSpecMan("E")
   ElseIf Right(mTM(1), 1) = "T" Then
        'Modified by Lydia 2021/07/30 °Ó¼Ð¤Î°Ó¼ÐªA°È·~°È¦¬¤å-¦]¥~°Ó³¯¸g²z°h¥ð¦Ó­×§ïµ{¦¡±±¨î
        If mTM(1) = "CFT" Then
           '¥H¥»©Ò®×¸¹©I¥sGetCFTSt16Manager§ì¥DºÞ
            oMailCount = GetCFTSt16Manager(mTM(1), mTM(2), mTM(3), mTM(4))
        Else
            '¥ý¥H¥»©Ò®×¸¹©I¥sPUB_GetFCTSalesNo§ì¥X­t³dªº¤H¡A¦A§ì¸Ó­û°£ST55¤§¥~ªº³Ì°ª¥DºÞNVL(NVL(ST54,ST53),ST52)
            strTmp1(1) = PUB_GetFCTSalesNo(mTM(1), mTM(2), mTM(3), mTM(4))
            If strTmp1(1) = "" Then
                oMailCount = Pub_GetSpecMan("D")
            Else
                oMailCount = PUB_GetSTManLimit(strTmp1(1), "4") '¼Ò²Õ¤Æ
                'Added by Lydia 2022/01/28 µoµ¹¨t²Î¯S®í³]©w¡uD¡v¤§¤H­û©M¥DºÞ
                strTmp1(2) = Pub_GetSpecMan("D")
                oMailCount = oMailCount & ";" & strTmp1(2)
                'end 2022/01/28
            End If
        End If
        'end 2021/07/30
        'add by nickc 2007/06/23 ¥[¤JFCT ª§Ä³®×³qª¾¤º°Ó°Óª§  84027;69008     ®×¥ó©Ê½è 202 °£¥~¡AÁÙ¬O°e¥~°Ó¡Aªü½¬»¡¦Û¤v§PÂ_¡A­Y¬°¤º°Ó®×¥ó¡A¥L·|¦AÂà¹L¨Ó
        If mTM(1) = "FCT" Then    'add by nickc 2007/08/01 ¤£§PÂ_ªº¸Ü  CFT ¤]·|¶i¤J
            Dim tmp960623 As New ADODB.Recordset
            Set tmp960623 = New ADODB.Recordset
            If tmp960623.State = 1 Then tmp960623.Close
            tmp960623.CursorLocation = adUseClient
            tmp960623.Open "select * from staff_group where sg01='C1' and sg02='FCT' and sg03='" & mCP(10) & "' and sg03<>'202'  ", cnnConnection, adOpenStatic, adLockReadOnly
            If tmp960623.RecordCount <> 0 Then
                 If mCP(6) <> "" And mCP(7) <> "" Then
                     'Modified by Lydia 2021/07/30 ¥[µo³Ì°ª¥DºÞ
                     strTmp1(1) = Pub_GetSpecMan("F")
                     If InStr(strTmp1(1), oMailCount) = 0 Then
                         oMailCount = strTmp1(1) & IIf(oMailCount = "", "", ";" & oMailCount)
                     Else
                         oMailCount = strTmp1(1)
                     End If
                     'end 2021/07/30
                 End If
            End If
            tmp960623.Close
            Set tmp960623 = Nothing
        End If
   ElseIf mTM(1) = "S" Then
        oMailCount = Pub_GetSpecMan("D")
   'Modified by Lydia 2021/07/30 debug: CFT¦b«e­±
   'ElseIf mTM(1) = "CFC" Or mTM(1) = "CFT" Then
   ElseIf mTM(1) = "CFC" Then
        oMailCount = Pub_GetSpecMan("L")
   End If
   strDivisionalEmp = oMailCount 'Add By Sindy 2022/9/27 °O¿ý¤À®×¤H­û,«áÄò·|¥Î¨ì
   'Add By Sindy 2023/1/11 ­Y¤w¦³¤À©Ó¿ì¤H,¦P®É¤@¨Ö³qª¾
   If mCP(14) <> "" Then
      If oMailCount <> "" Then oMailCount = oMailCount & ";"
      oMailCount = oMailCount & mCP(14)
   End If
   '2023/1/11 END
   
   'µoMail
   If mCP(6) < strSrvDate(1) And mCP(6) <> "" And Trim(oMailCount) <> "" Then
      bolIsOverDt = True 'Add By Sindy 2022/9/27
      '2007/8/13 MODIFY BY SONIA ¥[´¼Åv¤H­û
      'Modify By Sindy 2010/12/16 ¥[·~°È°Ï,¶O¥Î
      'Modify By Sindy 2013/4/11 +³W¶O,ÂI¼Æ
'      PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¤w¹O¥»©Ò´Á­­¡A½Ð¾¨³t¿ì²z¡I", oContext2 & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(mCP(6)) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & vbCrLf & "³W¶O¡@¡@¡G" & Format(mCP(17), "##,##0") & vbCrLf & "ÂI¼Æ¡@¡@¡G" & mCP(18)
      'Modify By Sindy 2022/9/29
      'Modified by Lydia 2022/12/23 +chgsql
      'Modify By Sindy 2023/3/27 +,mc13
      mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
         " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
         ",'" & "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¤w¹O¥»©Ò´Á­­¡A½Ð¾¨³t¿ì²z¡I(¤å¸¹:" & mCP(9) & ")','" & ChgSQL(oContext2) & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(mCP(6)) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & vbCrLf & "³W¶O¡@¡@¡G" & Format(mCP(17), "##,##0") & vbCrLf & "ÂI¼Æ¡@¡@¡G" & mCP(18) & "',null,'" & mCP(9) & "')"
      cnnConnection.Execute mStrSql
      '2022/9/29 END
   End If
   If mCP(6) = strSrvDate(1) And mCP(6) <> "" And Trim(oMailCount) <> "" Then
      bolIsOverDt = True 'Add By Sindy 2022/9/27
      '2007/8/13 MODIFY BY SONIA ¥[´¼Åv¤H­û
      'Modify By Sindy 2010/12/16 ¥[·~°È°Ï,¶O¥Î
      'Modify By Sindy 2013/4/11 +³W¶O,ÂI¼Æ
'      PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¤w©¡¥»©Ò´Á­­¡A½Ð¾¨³t¿ì²z¡I", oContext2 & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(mCP(6)) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & vbCrLf & "³W¶O¡@¡@¡G" & Format(mCP(17), "##,##0") & vbCrLf & "ÂI¼Æ¡@¡@¡G" & mCP(18)
      'Modify By Sindy 2022/9/29
      'Modified by Lydia 2022/12/23 +chgsql
      'Modify By Sindy 2023/3/27 +,mc13
      mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
         " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
         ",'" & "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¤w©¡¥»©Ò´Á­­¡A½Ð¾¨³t¿ì²z¡I(¤å¸¹:" & mCP(9) & ")','" & ChgSQL(oContext2) & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(mCP(6)) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & vbCrLf & "³W¶O¡@¡@¡G" & Format(mCP(17), "##,##0") & vbCrLf & "ÂI¼Æ¡@¡@¡G" & mCP(18) & "',null,'" & mCP(9) & "')"
      cnnConnection.Execute mStrSql
      '2022/9/29 END
   End If
   'add by nickc 2007/10/16 °²¤é«e¦¬¤å¡A¥B´Á­­¬°°²¤é
   'edit by nickc 2008/03/06 ¨q¬Â»¡§PÂ_¥¼¨ì´Áªº´N¦n
   'If txtTrademark(11).Text <> "" Then
   If mCP(6) > strSrvDate(1) Then
      If (mTM(1) = "T" Or mTM(1) = "FCT") And ChkMyWeek(mCP(6)) = True And Trim(oMailCount) <> "" Then
         bolIsOverDt = True 'Add By Sindy 2022/9/27
         'Modify By Sindy 2010/12/16 ¥[·~°È°Ï,¶O¥Î
         'Modify By Sindy 2013/4/11 +³W¶O,ÂI¼Æ
'         PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×§Y±N©¡¥»©Ò´Á­­¡A¥B¥»©Ò´Á­­¬°°²¤é¡A½Ð¾¨³t¿ì²z¡I", oContext2 & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(mCP(6)) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & vbCrLf & "³W¶O¡@¡@¡G" & Format(mCP(17), "##,##0") & vbCrLf & "ÂI¼Æ¡@¡@¡G" & mCP(18)
         'Modify By Sindy 2022/9/29
         'Modified by Lydia 2022/12/23 +chgsql
         'Modify By Sindy 2023/3/27 +,mc13
         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×§Y±N©¡¥»©Ò´Á­­¡A¥B¥»©Ò´Á­­¬°°²¤é¡A½Ð¾¨³t¿ì²z¡I(¤å¸¹:" & mCP(9) & ")','" & ChgSQL(oContext2) & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(mCP(6)) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & vbCrLf & "³W¶O¡@¡@¡G" & Format(mCP(17), "##,##0") & vbCrLf & "ÂI¼Æ¡@¡@¡G" & mCP(18) & "',null,'" & mCP(9) & "')"
         cnnConnection.Execute mStrSql
         '2022/9/29 END
      End If
      'add by nickc 2008/01/24 ­Y¬O¤À©Ò¦¬¤å¡A´Á­­¬°¤u§@¤Ñ¥B¬°¹j¤Ñ¤]­n³qª¾
      If (mTM(1) = "T" Or mTM(1) = "FCT") And pub_strUserOffice > "1" And Val(CompWorkDay(2, strSrvDate(1), 0)) = Val(mCP(6)) And Trim(oMailCount) <> "" Then
         bolIsOverDt = True 'Add By Sindy 2022/9/27
         'Modify By Sindy 2010/12/16 ¥[·~°È°Ï,¶O¥Î
         'Modify By Sindy 2013/4/11 +³W¶O,ÂI¼Æ
'         PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¬°¤À©Ò®×¥ó¥B±N©¡¥»©Ò´Á­­¡A¥»©Ò´Á­­¬°¤U¤@¤u§@¤é¡A½Ð¾¨³t¿ì²z¡I", oContext2 & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(mCP(6)) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & vbCrLf & "³W¶O¡@¡@¡G" & Format(mCP(17), "##,##0") & vbCrLf & "ÂI¼Æ¡@¡@¡G" & mCP(18)
         'Modify By Sindy 2022/9/29
         'Modified by Lydia 2022/12/23 +chgsql
         'Modify By Sindy 2023/3/27 +,mc13
         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¬°¤À©Ò®×¥ó¥B±N©¡¥»©Ò´Á­­¡A¥»©Ò´Á­­¬°¤U¤@¤u§@¤é¡A½Ð¾¨³t¿ì²z¡I(¤å¸¹:" & mCP(9) & ")','" & ChgSQL(oContext2) & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(mCP(6)) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & vbCrLf & "³W¶O¡@¡@¡G" & Format(mCP(17), "##,##0") & vbCrLf & "ÂI¼Æ¡@¡@¡G" & mCP(18) & "',null,'" & mCP(9) & "')"
         cnnConnection.Execute mStrSql
         '2022/9/29 END
      End If
   End If
   
   'Add By Sindy 2022/9/27
   If UCase(pFormName) = UCase("frm090801_New") Then
      If mCP(14) <> "" Then 'Txx¦³¹w³]©Ó¿ì¤H And Mid(mTM(1), 1, 1) = "T"
         '¥Ø«e¨t²Î¤é´Á®É¶¡
         strRDate = strSrvDate(1)
         strRTime = Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)
         '­Y¹L¤U¯Z®É¶¡,ÀË¬d¤é´Á§ï¬°¤U¤@­Ó¤u§@¤Ñ,®É¶¡¹w³]¬°07:30
         'Modify By Sindy 2025/4/2 ¬°17:00«á, §ï¬°§PÂ_¹j¤é¬O§_¤H­û¥ð°²
         'If Val(Format(strRTime, "hhmm")) > 1800 Then
         If Val(Format(strRTime, "hhmm")) > 1700 Then
            strRDate = CompWorkDay(2, strSrvDate(1), 0)
            strRTime = "07:30" '¤W¤È¬O§_¦³¥ð°²
         End If
         bolCP14Rest = CheckIsPersonRest(mCP(14), strRDate, strRTime, strRestKind)
         If bolCP14Rest = True Then
            'Modify By Sindy 2025/4/2 ¤H­û¥ð°²,¤£ºÞ¬O§_¹O´Á³£¨ú®ø¹w¤À©Ó¿ì¤H,§ï¥Ñ¥DºÞ¤À®×
'            If bolIsOverDt = True Then
'               '·í¤é(¤w¹O)(±N©¡)¥»©Ò´Á­­®×¥ó, ©ÒÄÝ©Ó¿ì¤H½Ð°², ±Ä¤H¤u¤À®×
               'Add By Sindy 2025/4/2 ¤è«K¥DºÞª¾¹D³o¬O¥i¹w¤À©Ó¿ì¤H,¦ý¤H­û¥ð°²§ï¥DºÞ¤À®×
               mStrSql = "update ConsultRecCMP set CRC09='" & mCP(14) & "'" & _
                         " where CRC01='" & mCP(140) & "'" & _
                         " and CRC03='" & mCP(10) & "' and CRC08 is null and CRC09 is null"
               cnnConnection.Execute mStrSql, intI
               '2025/4/2 END
               
               mCP(14) = ""
               mStrSql = "UPDATE caseprogress SET cp14=null WHERE cp09 = '" & mCP(9) & "'"
               cnnConnection.Execute mStrSql, intI
'            Else
''               '¨¾¦hµ§µoMail; ex:T¤À³Î®×
''               strtmp1(0) = "select * from FLOW002 where F0201='" & mCP(140) & "' and F0202='A6' and F0207 is null"
''               intI = 1
''               Set RsTemp = ClsLawReadRstMsg(intI, strtmp1(0))
''               If intI = 1 Then
'                  '©Ó¿ì¤H½Ð°²:µomail³qª¾ "®×¥ó¦¬¤å³qª¾,¥»©Ò´Á­­¬°xxx/xx/xx¦]©Ó¿ì¤H½Ð°²,½Ð°Æ¥»¦¬¨üªÌ¥N¬°¿ì²z¡I"¥Ø«e°Æ¥»¥[±¾³qª¾¹Å¶²©M©Ó¼z
'                  'Modified by Lydia 2022/12/23 +chgsql
'                  'Modify By Sindy 2023/3/27 +,mc13
'                  mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
'                     " values( '" & strUserNum & "','" & mCP(14) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'                     ",'" & "®×¥ó¦¬¤å³qª¾,¥»©Ò´Á­­¬°" & ChangeWStringToTDateString(mCP(6)) & "¦]©Ó¿ì¤H½Ð°²,½Ð°Æ¥»¦¬¨üªÌ¥N¬°¿ì²z¡I(¤å¸¹:" & mCP(9) & ")','" & ChgSQL(oContext2) & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(mCP(6)) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & vbCrLf & "³W¶O¡@¡@¡G" & Format(mCP(17), "##,##0") & vbCrLf & "ÂI¼Æ¡@¡@¡G" & mCP(18) & "','84027;86048','" & mCP(9) & "')"
'                  cnnConnection.Execute mStrSql
''               End If
'            End If
         End If

         '½T©w¤w¤À©Ó¿ì¤H
         If mCP(14) <> "" Then
            '­pºâ©Ó¿ì´Á­­
            Call PUB_CountUpdTxCP48(mCP(9), mCP(10), mCP(143), mCP(5), mCP(6), mCP(7), mCP(13), mCP(122), mTM(1), mTM(10), mCP(48))
         End If
      End If
      
      If ERecvSaveProgress(mCP, mTM, strDivisionalEmp, oContext2) = False Then
         GoTo ErrHand
      End If
   End If
   
End If
   
   Set adoquery = Nothing
   Exit Function
   
ErrHand:
   PUB_SaveFrm010004 = False 'Add By Sindy 2022/10/25
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical, "PUB_SaveFrm010004"
   End If
   Set adoquery = Nothing
End Function

'Add By Sindy 2022/9/28 ¹q¤l¦¬¤å«á­n³B²zªºµ{¦¡
Private Function ERecvSaveProgress(ByRef modCP() As String, ByRef modBase() As String, _
   strDivisionalEmp As String, oContext2 As String) As Boolean
   
Dim strUpdTime As String
Dim m_SalesDeptName As String
Dim intKind As Integer
Dim strCRL() As String
Dim strFolder As String, pSavePath As String
Dim adoquery As New ADODB.Recordset
Dim strContext As String
   
   ERecvSaveProgress = False
   strUpdTime = Right("000000" & ServerTime, 6)
   Call GetST15(modCP(13), m_SalesDeptName)
   
   '±µ¬¢³æ¥DÀÉ
   ReDim Preserve strCRL(TF_CRL) As String
   strCRL(1) = modCP(140)
   Call ClsPDReadCRLDatabase(strCRL)
   
   Call ClsPDGetSystemKind(modCP(1), intKind)
   
   '·s®×®É¡A±µ¬¢³æªº¦¬¾Ú¤½¥q­n¼g¤J®×¥ó°ò¥»ÀÉªº¯S®í¥X¦W¤½¥qÄæ(J©ÎL)
   If strCRL(6) = "Y" And strCRL(49) <> "" Then
      mStrSql = ""
      If intKind = ±M§Q Then
         mStrSql = "update patent set pa161='" & strCRL(49) & "' where pa01='" & modCP(1) & "' and pa02='" & modCP(2) & "' and pa03='" & modCP(3) & "' and pa04='" & modCP(4) & "'"
      ElseIf intKind = °Ó¼Ð Then
         mStrSql = "update trademark set tm130='" & strCRL(49) & "' where tm01='" & modCP(1) & "' and tm02='" & modCP(2) & "' and tm03='" & modCP(3) & "' and tm04='" & modCP(4) & "'"
      ElseIf intKind = ªk°È Then
         mStrSql = "update lawcase set lc48='" & strCRL(49) & "' where LC01='" & modCP(1) & "' and LC02='" & modCP(2) & "' and LC03='" & modCP(3) & "' and LC04='" & modCP(4) & "'"
      ElseIf intKind = ÅU°Ý Then
         mStrSql = "" 'µL
      Else 'ªA°È
         mStrSql = "update servicepractice set sp85='" & strCRL(49) & "' where sp01='" & modCP(1) & "' and sp02='" & modCP(2) & "' and sp03='" & modCP(3) & "' and sp04='" & modCP(4) & "'"
      End If
      If mStrSql <> "" Then cnnConnection.Execute mStrSql, intJ
   End If
   
'Removed by Morgan 2025/8/4 ¥Ø«e¦¬¤å¦³±±¨î³¬¨÷®×¥ó³£ÁÙ¬O»Ý­n¤H¤u¤À®×¡A¬G¦¹³B¤£·|³Q°õ¦æ¡A¥ý¨ú®ø¥H§KÀË¬dµ{¦¡®É·|»~§P
'   'ÀË¬d®×¥ó¬O§_¤w³¬¨÷:
'   'Modify By Sindy 2023/1/6
'   'ÂÂ®×:P®×»âÃÒ¤ÎÃº¦~¶O,¦~¶O
'   'Morgan»¡¤º±Mªº¤º±M»âÃÒ¦~¶O¾ã§åµo¤åfrm040104_i¡A¤w¸g¤£¦Ò¼{CP157¤ÎCP140¡A©Ò¥H»âÃÒ¦~¶Oªº¯S®íª¬ªp¥u­n¤£¹w³]©Ó¿ì¤H§Y¥i¡C
'   'Modify By Sindy 2023/1/6 ¨q¬Â:³¬¨÷®×¦¬¤å¼u°T®§¡u½Ðª`·N¡I¦¹®×¤w³¬¨÷¬O§_½T©w­n¦¬¤å¡H¡v§Y¥i
'   '¡A¨ú®ø¡u( ­Y½T©w­n¦¬¤å·|¨ú®ø³¬¨÷!! )¡C¡vªº¤å¦r¡A¨Ã¶}©ñ¤´¥i¦¬¤å¦ý¤£¥i¨ú®ø³¬¨÷¡F¥u¦³P®×ªº¦~¶O¥i¥H¨ú®ø³¬¨÷¡C
'   '¤]¤£¥²µoMAILµ¹±M·~³¡¡A¸g²z»¡²{¦b«Ü§Ö´N·|¤À®×¡C
'   If strCRL(6) = "" And modCP(1) = "P" And (modCP(10) = "601" Or modCP(10) = "605") And modCP(14) <> "" Then
'      strTmp1(0) = "select pa57,'P' sType from patent where pa01='" & modCP(1) & "' and pa02='" & modCP(2) & "' and pa03='" & modCP(3) & "' and pa04='" & modCP(4) & "' and pa57 is not null" & _
'                  " union select tm29,'T' sType from trademark where tm01='" & modCP(1) & "' and tm02='" & modCP(2) & "' and tm03='" & modCP(3) & "' and tm04='" & modCP(4) & "' and tm29 is not null" & _
'                  " union select sp15,'S' sType from servicepractice where sp01='" & modCP(1) & "' and sp02='" & modCP(2) & "' and sp03='" & modCP(3) & "' and sp04='" & modCP(4) & "' and sp15 is not null" & _
'                  " union select lc08,'L' sType from lawcase where lc01='" & modCP(1) & "' and lc02='" & modCP(2) & "' and lc03='" & modCP(3) & "' and lc04='" & modCP(4) & "' and lc08 is not null" & _
'                  " union select hc09,'H' sType from hirecase where hc01='" & modCP(1) & "' and hc02='" & modCP(2) & "' and hc03='" & modCP(3) & "' and hc04='" & modCP(4) & "' and hc09 is not null"
'      intJ = 1
'      Set adoquery = ClsLawReadRstMsg(intJ, strTmp1(0))
'      If intJ = 1 Then
'         '«ì´_¦¨¥¼³¬¨÷
'         If adoquery.Fields("sType") = "P" Then
'            mStrSql = "update patent set pa57=null,pa58=null,pa59=null where pa01='" & modCP(1) & "' and pa02='" & modCP(2) & "' and pa03='" & modCP(3) & "' and pa04='" & modCP(4) & "'"
'         ElseIf adoquery.Fields("sType") = "T" Then
'            mStrSql = "update trademark set tm29=null,tm30=null,tm31=null where tm01='" & modCP(1) & "' and tm02='" & modCP(2) & "' and tm03='" & modCP(3) & "' and tm04='" & modCP(4) & "'"
'         ElseIf adoquery.Fields("sType") = "S" Then
'            mStrSql = "update servicepractice set sp15=null,sp16=null,sp17=null where sp01='" & modCP(1) & "' and sp02='" & modCP(2) & "' and sp03='" & modCP(3) & "' and sp04='" & modCP(4) & "'"
'         ElseIf adoquery.Fields("sType") = "L" Then
'            mStrSql = "update lawcase set LC08=null,LC09=null,LC10=null where LC01='" & modCP(1) & "' and LC02='" & modCP(2) & "' and LC03='" & modCP(3) & "' and LC04='" & modCP(4) & "'"
'         Else 'H
'            mStrSql = "update hirecase set hc09=null,hc10=null,hc11=null where hc01='" & modCP(1) & "' and hc02='" & modCP(2) & "' and hc03='" & modCP(3) & "' and hc04='" & modCP(4) & "'"
'         End If
'         Pub_SeekTbLog mStrSql 'Add By Sindy 2023/1/6
'         cnnConnection.Execute mStrSql, intJ
'         'µoMail³qª¾¤À®×¤H­û
'         If strDivisionalEmp = "" Then strDivisionalEmp = Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'         'Modified by Lydia 2022/12/23 +chgsql
'         'Modify By Sindy 2023/3/27 +,mc13
'         'Removed by Morgan 2025/8/4 ¤W­±¦³³Æµù¤£¥²EMail¥B¹ê»Ú¤W¤]¨S¦³·s¼W
'         'mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
'            " values( '" & strUserNum & "','" & strDivisionalEmp & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'            ",'" & "½Ðª`·N¡I¦¹®× " & modCP(1) + "-" + modCP(2) + "-" + modCP(3) + "-" + modCP(4) & "¡A¨t²Î¤w¨ú®ø³¬¨÷¡I(¤å¸¹:" & modCP(9) & ")','" & ChgSQL(oContext2) & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(modCP(6)) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(modCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(modCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(modCP(16), "##,##0") & vbCrLf & "³W¶O¡@¡@¡G" & Format(modCP(17), "##,##0") & vbCrLf & "ÂI¼Æ¡@¡@¡G" & modCP(18) & "',null,'" & modCP(9) & "')"
'      End If
'   End If
'end 2025/8/4
   
   'Add By Sindy 2023/3/30
   '±µ¬¢³æ½Ð¥[±±¨î¡GACSÂÂ®×¦¬¤å¡A­Y¸Ó®×¸¹´¿¦³ªþ¥óCÃþ¤À¼í«¬ºA®×¥ó©Ê½èªº¦¬¤å(Pub_GetSpecMan("ACS-C"))¡A
   '¦ý²{¦b¦¬¤å«DCÃþ¥B¦³ÂI¼Æ®É¡A½Ð¥[µoEMAILµ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v
   If modCP(1) = "ACS" And modCP(31) = "" Then
      strTmp1(0) = "select cp09 from caseprogress where cp01='" & modCP(1) & "' and cp02='" & modCP(2) & "'" & _
                  " and cp03='" & modCP(3) & "' and cp04='" & modCP(4) & "'" & _
                  " and cp159=0 and instr('" & Pub_GetSpecMan("ACS-C") & "',cp10)>0" & _
                  " and (cp140<>'" & modCP(140) & "' or cp140 is null)"
      intJ = 1
      Set adoquery = ClsLawReadRstMsg(intJ, strTmp1(0))
      If intJ = 1 Then
         strTmp1(0) = "select cp09 from caseprogress where cp140='" & modCP(140) & "'" & _
                     " and instr('" & Pub_GetSpecMan("ACS-C") & "',cp10)>0"
         intJ = 1
         Set adoquery = ClsLawReadRstMsg(intJ, strTmp1(0))
         If intJ = 0 Then '²{¦b¦¬¤å«DCÃþ
            strTmp1(0) = "select cp09 from caseprogress where cp140='" & modCP(140) & "'" & _
                        " and cp18>0"
            intJ = 1
            Set adoquery = ClsLawReadRstMsg(intJ, strTmp1(0))
            If intJ = 1 Then '¦³ÂI¼Æ
               strContext = ChgSQL(oContext2) & "´¼Åv¤H­û¡@¡G" & GetStaffName(modCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf
               '¥Ó½Ð¤H
               If GetCustomerAndState(modBase(11), strTmp1(1), , , , modBase(1), , , , modBase(2), modBase(3), modBase(4)) Then
                 strContext = strContext & "¥Ó½Ð¤H¡G " + modBase(11) + " " + strTmp1(1) + vbCrLf
               End If
               strContext = strContext & "¶O¥Î¡@¡@¡G" & Format(modCP(16), "##,##0") & vbCrLf & "³W¶O¡@¡@¡G" & Format(modCP(17), "##,##0") & vbCrLf & "ÂI¼Æ¡@¡@¡G" & modCP(18)
               mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
                  " values( '" & strUserNum & "','" & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                  ",'" & "ACS-¥»©Ò®×¸¹¤§TIPS®×¥ó«áÄò¦³ªA°È¶O¤§¦¬¤å³qª¾¡I(¤å¸¹:" & modCP(9) & ")','" & strContext & "',null,'" & modCP(9) & "')"
               cnnConnection.Execute mStrSql
            End If
         End If
      End If
   End If
   '2023/3/30 END
   
   '¦Û°Ê¦¬¤å¤W½u®É¡ACP150§ï¬°¦s©ñ¬O§_¦³¯S¨ÒÃ±®Öªºµù°O
   'Modify By Sindy 2023/5/15 §ïÀË¬dflow002¬O§_¦³¯S¨ÒÃ±®Ö
   'strtmp1(0) = "select cra01 from consultrecapp where cra01='" & modCP(140) & "' and (cra26='Y' or cra27='Y')"
   strTmp1(0) = "select f0201 from flow002 where f0201='" & modCP(140) & "' and f0202='A2'"
   intJ = 1
   Set adoquery = ClsLawReadRstMsg(intJ, strTmp1(0))
   If intJ = 1 Then
      mStrSql = "update caseprogress set cp150='Y' where cp09='" & modCP(9) & "'"
      cnnConnection.Execute mStrSql, intJ
   End If
   '¬O§_«æ¥ó
   mStrSql = "update caseprogress set cp122='" & strCRL(90) & "' where cp09='" & modCP(9) & "' and cp122 is null"
   cnnConnection.Execute mStrSql, intJ
   
'   'ÂÂ®×,§ó·sÃÒ®Ñ§Î¦¡
'   If strCRL(6) = "" And strCRL(59) <> "" Then 'ÂÂ®×
'      If strCRL(7) = "P" Then
'         mstrSql = "update patent set pa178='" & strCRL(59) & "' where pa01='" & modCP(1) & "' and pa02='" & modCP(2) & "' and pa03='" & modCP(3) & "' and pa04='" & modCP(4) & "'"
'         cnnConnection.Execute mstrSql, intj
'      Else
'         mstrSql = "update trademark set tm136='" & strCRL(59) & "' where tm01='" & modCP(1) & "' and tm02='" & modCP(2) & "' and tm03='" & modCP(3) & "' and tm04='" & modCP(4) & "'"
'         cnnConnection.Execute mstrSql, intj
'      End If
'   End If
   
   '·s¼W¤@µ§±µ¬¢°O¿ý³æOrder¦Ü¨÷©v°Ï
   mStrSql = "insert into casepaperpdf(cpp01,cpp02,cpp03,CPP05,CPP06,CPP07,cpp08,cpp09,cpp10)" & _
            " values('" & modCP(9) & "','" & PUB_CaseNo2FileName(modCP(1), modCP(2), modCP(3), modCP(4)) & _
                    "." & modCP(10) & "." & EMP_±µ¬¢³æ & ".menu',0,'" & strUserNum & "'," & _
                    strSrvDate(1) & "," & strUpdTime & "," & _
                    strSrvDate(1) & "," & strUpdTime & ",'Y')"
   cnnConnection.Execute mStrSql, intJ
   
   ERecvSaveProgress = True
   Set adoquery = Nothing
End Function

'Add By Sindy 2022/9/29 ­pºâ°Ó¼ÐTxªº©Ó¿ì´Á­­; §ï¬°¦@¥Î¨ç¼Æ
'Optional ByRef strDate As String : ¦^¶Ç©Ó¿ì´Á­­
Public Function PUB_CountUpdTxCP48(m_CP09 As String, strCP10 As String, strCP143 As String, _
   strCP05 As String, strCP06 As String, strCP07 As String, strCP13 As String, strCP122 As String, _
   m_TM01 As String, m_TM10 As String, Optional ByRef strDate As String) As Boolean
   
Dim strEP06 As String
Dim strCP149 As String
Dim strCP142 As String
Dim strCP164 As String
Dim strTmp As String, strTmp1 As String, strExSql As String, intA As Integer, rsAD As New ADODB.Recordset 'Added by Lydia 2024/05/15

   PUB_CountUpdTxCP48 = False
   
   '¶Ç¤Jªº¸ê®Æ²Î¤@Âà¦è¤¸¦~
   strCP143 = DBDATE(strCP143)
   strCP05 = DBDATE(strCP05)
   strCP06 = DBDATE(strCP06)
   strCP07 = DBDATE(strCP07)
   
   '¤å¥ó»ô³Æ¤é
   strTmp = "select ep06 from engineerprogress where ep02='" & m_CP09 & "'"
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strTmp)
   If intA = 1 Then
      strEP06 = "" & rsAD.Fields("ep06")
   End If
   
'******************************************************************************
   ' ­pºâ©Ó¿ì´Á­­
'******************************************************************************
   'Add By Sindy 2022/4/27 ¨ú±o¤À®×¤é
   'Modify By Sindy 2024/1/15 +,cp142,cp164
   strTmp = "select cp09,cp149,cp48,cp142,cp164 from caseprogress WHERE CP09 = '" & m_CP09 & "' "
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strTmp)
   If intA = 1 Then
      strDate = "" & rsAD.Fields("cp48")
      strCP149 = "" & rsAD.Fields("cp149")
      If strSrvDate(1) >= ±µ¬¢³æ¹q¤l¦¬¤å±Ò¥Î¤é And Val(strCP149) > 0 Then
         strCP149 = CompWorkDay(1, CompDate(2, 1, strCP149), 0) '¤£§t·í¤é,¥[1¤Ñ; ¹Å¶²¦Ò¼{¨ì­Yªñ¤U¯Z®É¶¡¦¬¤å,¹ï©Ó¿ì¤H¤£¤½¥­
      End If
      'Add By Sindy 2024/1/15
      strCP142 = "" & rsAD.Fields("cp142")
      strCP164 = "" & rsAD.Fields("cp164")
      '2024/1/15 END
   End If
   '2022/4/27 END
   'Modify by Morgan 2003/12/05
'      strDay = GetWorkDays(m_TM01, m_TM10, strCP10)
'      If IsEmptyText(strDay) = False Then
'         ' 90.07.03 ©Ó¿ì´Á­­¥H¤u§@¤Ñ­pºâ
'         'strDate = DBDATE(DateSerial(Val(DBYEAR(strCP05)), Val(DBMONTH(strCP05)), Val(DBDAY(strCP05)) + Val(strDay)))
'         strDate = DBDATE(CompWorkDay(Val(strDay), DBDATE(strCP05), 0))
'
'         strexsql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
'                  "WHERE CP09 = '" & m_CP09 & "' "
'         cnnConnection.Execute strexsql
'      End If
   'edit by nick 2004/12/08
   'If m_CP10 = "102" Then
   If strCP10 = "102" Then '©µ®i
      Dim tmpDate As Date
      'MODIFY BY SONIA 2013/6/10 °¨¼w¨½Äò®i§ïªk©w´Á­­«e¤T­Ó¤ë TF-000110
      'tmpDate = DateAdd("M", -6, ChangeTStringToWDateString(strCP07))
      If m_TM01 = "TF" Then
         'Modified by Lydia 2019/04/12
         tmpDate = DateAdd("M", -3, ChangeWStringToWDateString(strCP07))
      Else
         'Modify By Sindy 2014/5/8
         If m_TM10 = "020" Then '¦]T¤j³°­×ªk¡A©Ó¿ì´Á­­§ï¬°ªk©w´Á­­´î¤@¦~ ex.T-1731721
            tmpDate = DateAdd("M", -12, ChangeWStringToWDateString(strCP07))
         Else
         '2014/5/8 END
            tmpDate = DateAdd("M", -6, ChangeWStringToWDateString(strCP07))
         End If
      End If
      '2013/6/10 END
      '¦¬¤å¤é¤j©óªk©w´Á­­¡G¦¬¤å¤é´Á+3¤Ñ
      If (Val(strCP05) > Val(strCP07)) Then
         'Modified by Lydia 2019/04/12 §ï¦¨+3¤u§@¤Ñ(¤£§t·í¤Ñ)
         'strDate = Format(DateAdd("d", 3, ChangeTStringToWDateString(strCP05)), "YYYYMMDD")
         strDate = CompWorkDay(4, DBDATE(strCP05))
      '¦¬¤å¤é¤p©ó©Ó¿ì´Á­­¡G©Ó¿ì´Á­­¡]ªk©w´Á­­´î¤@¦~©Î¥b¦~©Î3­Ó¤ë¡^+3¤Ñ
      'modify by sonia 2022/10/5 T-184822 tmpDate¬O¦³/ªº¤é´Á¡AVal(tmpDate)¥u·|§ì¨ì/¤§«eªº¦~«×4½X¡A©Ò¥HVal(strCP05)¬O8½X¤£¥i¯à¤p©ó4½X
      'ElseIf Val(strCP05) < Val(tmpDate) Then 'Format(tmpDate, "YYYYMMDD") - 19110000
      ElseIf Val(strCP05) < Val(Format(tmpDate, "YYYYMMDD")) Then
         'Modified by Lydia 2019/04/12 §ï¦¨+3¤u§@¤Ñ(¤£§t·í¤Ñ)
         'strDate = Format(tmpDate + 3, "YYYYMMDD")
         strDate = CompWorkDay(4, Format(tmpDate, "YYYYMMDD"))
      '§_«h¡A¦¬¤å¤é´Á+3¤Ñ¡A­Y¤j©óªk©w´Á­­«h©Ó¿ì´Á­­=ªk©w´Á­­
      Else
         'Modified by Lydia 2019/04/12 §ï¦¨+3¤u§@¤Ñ(¤£§t·í¤Ñ)
         'strDate = Format(DateAdd("d", 3, ChangeTStringToWDateString(strCP05)), "YYYYMMDD")
         strDate = CompWorkDay(4, DBDATE(strCP05))
         'Modify By Sindy 2022/10/6
         'If strDate > ChangeTStringToWString(strCP07) Then
         If strDate > strCP07 Then
            'Modify By Sindy 2022/10/6
            'strDate = ChangeTStringToWString(strCP07)
            strDate = strCP07
         End If
      End If
      strExSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
               "WHERE CP09 = '" & m_CP09 & "' "
      cnnConnection.Execute strExSql
   Else
''''edit by nickc 2007/10/11 §ï§ì¦³®É®Ä©Êªº
''''         strDay = GetWorkDays(m_TM01, m_TM10, strCP10)
''''         If IsEmptyText(strDay) = False Then
''''            strDate = DBDATE(CompWorkDay(Val(strDay), DBDATE(strCP05), 0))
      'Add By Sindy 2012/5/8
      'Modified by Lydia 2018/12/10 §PÂ_°Óª§®×
      'If Frame21.Visible = True Then
      'Modified by Lydia 2022/07/15 ­­¨îT,FCT®× => And (m_TM01 = "T" Or m_TM01 = "FCT")
      'Modify By Sindy 2025/7/30 FCT®×727¤ÀªR¤£ÄÝ©óª§Ä³®× + FCT_NotTMdebate
      If (m_TM01 = "T" Or m_TM01 = "FCT") And InStr(TMdebate, strCP10) > 0 _
         And Not (m_TM01 = "FCT" And InStr(FCT_NotTMdebate, strCP10) > 0) Then
         
         '©Ó¿ì¤HÄæ¥ÑµL¡Ð¡Ö¦³®É¡A¥B¤w¿é¤J¸ê®Æ»ô³Æ®É«h­pºâ©Ó¿ì´Á­­
         'Modify By Sindy 2022/4/27 + And Val(strCP149) > 0
         If Val(strEP06) > 0 And Val(strCP149) > 0 Then
            'Modify By Sindy 2022/4/27
            'µL»ô³Æ¤é®É,¤£­pºâ©Ó¿ì´Á­­, ­pºâ©Ó¿ì´Á­­¥H¤À®×¤é¬°°_ºâ¤é
            'strDate = PUB_TMdebateCountCP48(textCP06, textCP122, m_EP06DT, m_CP09, textCP13)
            strDate = PUB_TMdebateCountCP48(strCP06, strCP122, strCP149, m_CP09, strCP13)
            '2022/4/27 END
            strExSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
                     "WHERE CP09 = '" & m_CP09 & "' "
            cnnConnection.Execute strExSql
         End If
      'Added by Lydia 2018/12/10 «Dª§Ä³®×¦¬¤å¤é¦bT®×¦¬¤å»ô³Æ±Ò¥Î¤é¤§«á
      'Memo by Lydia 2019/04/11 «Dª§Ä³®×(AÃþ)T®×¦¬¤å»ô³Æ±Æ°£ªº®×¥ó©Ê½è¬Ò¤£¥ÎºÞ¨î»ô³Æ¤é(¹w³]¤å¥ó»ô³Æ=Y)
      'Modified by Lydia 2022/07/15  T¤j³°®×¤§»ô³Æ¤éºÞ±±; TC®×¤§¤å¥ó»ô³Æ¤éºÞ±±;
      'ElseIf Frame21.Visible = True And InStr(TMdebate, strCP10) = 0 And DBDATE(strCP05) >= T®×¦¬¤å»ô³Æ±Ò¥Î¤é Then
      ElseIf (m_TM01 = "TC" Or (m_TM01 = "T" And InStr(TMdebate, strCP10) = 0 And DBDATE(strCP05) >= T®×¦¬¤å»ô³Æ±Ò¥Î¤é)) Then
         
         'Added by Lydia 2019/01/30 ©Ó¿ì¤HÄæ¥ÑµL¡Ð¡Ö¦³®É¡A¥B¤w¿é¤J¤å¥ó©M¬d¦W»ô³Æ®É«h­pºâ©Ó¿ì´Á­­
         'Modified by Lydia 2022/07/15 T¤j³°®×¤§»ô³Æ¤éºÞ±±; TC®×¤§¤å¥ó»ô³Æ¤éºÞ±±;
         'If (textCP14.Tag = "" And textCP14 <> "") And ((strCP10 = ¥Ó½Ð And textEP06 = "Y" And textCP143 = "Y") _
                              Or (strCP10 <> ¥Ó½Ð And textEP06 = "Y")) Then
         If ((m_TM01 = "T" And strCP10 = ¥Ó½Ð And Val(strEP06) > 0 And Val(strCP143) > 0) _
            Or (m_TM01 = "T" And strCP10 <> ¥Ó½Ð And Val(strEP06) > 0) Or m_TM01 = "TC") Then
            
            'Modified by Lydia 2019/04/11 ©Ó¿ì´Á­­¥H»ô³Æ¤é+®×¥ó©Ê½è©Ò³]¤u§@¤Ñ¼Æ
            'strDate = PUB_TMdebateCountCP48(textCP06, textCP122, m_EP06DT, m_CP09, textCP13)
            strTmp1 = ""
            If Val(strEP06) > 0 Then
               strTmp1 = strEP06
              If Val(strCP143) > 0 Then
                  If Val(strTmp1) < Val(strCP143) Then
                       strTmp1 = strCP143
                  End If
               End If
            End If
            'Modify By Sindy 2022/4/27
            'µL»ô³Æ¤é®É,¤£­pºâ©Ó¿ì´Á­­
            '¦Ó­pºâ©Ó¿ì´Á­­¥H¤À®×¤é¬°°_ºâ¤é
            If Val(strTmp1) = 0 Or Val(strCP149) = 0 Then 'µL»ô³Æ¤é©ÎµL¤À®×¤é
               'strtmp1 = DBDATE(strCP05) 'µL»ô³Æ¤é,´N¥Î¦¬¤å¤é
               strDate = ""
               strExSql = "UPDATE CaseProgress SET CP48 = null " & _
                        "WHERE CP09 = '" & m_CP09 & "' "
               cnnConnection.Execute strExSql
            Else
               'strDate = Pub_GetHandleDay(m_TM01, m_TM10, strCP10, strtmp1, DBDATE(textCP06), textCP09)
               'Memo by Lydia 2022/07/15 TC®×¤§¤å¥ó»ô³Æ¤éºÞ±±: ¦Û¤å¥ó»ô³Æ¤é°_ºâ¡A¤­­Ó¤u§@¤Ñ¡F»P¨q¬Â°Q½×¨M©wª½±µ­×§ïCaseFee¡A¦³³]©wªº©Ê½è3¤Ñ§ï¬°»OÆW®×5¤Ñ/¤j³°®×6¤Ñ
                                                              '»OÆW®×ªº5­Ó¤u§@¤Ñ§t·í¤Ñ(by ¹Å¶²)¡F¤j³°®×ªº5­Ó¤u§@¤Ñ¤£§t·í¤Ñ(by ©Ó¼z)¡C
               strDate = Pub_GetHandleDay(m_TM01, m_TM10, strCP10, strCP149, DBDATE(strCP06), m_CP09)
               If strDate <> "" Then
               'end 2019/04/11
                    strExSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
                             "WHERE CP09 = '" & m_CP09 & "' "
                    cnnConnection.Execute strExSql
               End If
            End If
            '2022/4/27 END
         Else
            'Added by Lydia 2019/04/12 T¥xÆW®×(¥]§t¤¤¶¡µ{§Ç)¦¬¤å»ô³Æ¤éºÞ¨î, ±Æ°£¯S©w®×¥ó©Ê½è¤£¥Î¿é¤J¤å¥ó»ô³Æ; ¦Û°Ê­pºâ©Ó¿ì´Á­­
            'Memo by Lydia 2022/07/15 T¤j³°®×¤£­n¡uT®×¦¬¤å»ô³Æ±Æ°£¡v=> And PA09 = "000"
            If m_TM01 = "T" And m_TM10 = "000" And Val(strEP06) = 0 And InStr(T®×¦¬¤å»ô³Æ±Æ°£, strCP10) > 0 Then
               strDate = Pub_GetHandleDay(m_TM01, m_TM10, strCP10, strCP149)
               strExSql = "UPDATE CaseProgress SET CP48 = " & CNULL(strDate) & " WHERE CP09 = '" & m_CP09 & "' "
               cnnConnection.Execute strExSql, intA
            End If
            'end 2019/04/12
         End If
      'end 2018/12/10
      Else
      '2012/5/8 End
         'Added by Lydia 2019/04/11 Àx¦s¤å¥ó»ô³Æ¤é (ex.«Dª§Ä³®×ªº¤º³¡¦¬¤å(BÃþ)©MT®×¦¬¤å»ô³Æ±Æ°£ªº®×¥ó©Ê½è)
         'Modify By Sindy 2022/4/27
         'µL»ô³Æ¤é®É,¤£­pºâ©Ó¿ì´Á­­, ­pºâ©Ó¿ì´Á­­¥H¤À®×¤é¬°°_ºâ¤é
         If Val(strEP06) > 0 And Val(strCP149) > 0 Then
            'strDate = Pub_GetHandleDay(m_TM01, m_TM10, strCP10, DBDATE(strCP05), DBDATE(textCP06), textCP09)
            strDate = Pub_GetHandleDay(m_TM01, m_TM10, strCP10, strCP149, DBDATE(strCP06), m_CP09)
            If IsEmptyText(strDate) = False Then
               strExSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
                        "WHERE CP09 = '" & m_CP09 & "' "
               cnnConnection.Execute strExSql
            End If
         End If
      End If
      
      'Add By Sindy 2024/1/15
      '«eÀY¦³ºâ¥X©Ó¿ì´Á­­ªÌ,­Y¦³«ü©w¤é´ÁªºT¡BTC®×¥ó,¨Ì«ü©w°e¥ó¤é­«·s­pºâ©Ó¿ì´Á­­
      If strSrvDate(1) >= «ü©w¤é´Á±Ò¥Î¤é Then
         If Val(strDate) > 0 And Val(strCP142) > 0 And Val(strCP164) > 0 Then
            If (m_TM01 = "T" Or m_TM01 = "TC") And InStr(TMdebate, strCP10) = 0 Then '±Æ°£ª§Ä³®×
               '«ü©w¤é´Á·í¤é°e¥ó¡G©Ó¿ì´Á­­¬°«ü©w¤é´Á«e¤T¤é
               If strCP164 = "1" Then
                  strDate = CompWorkDay(4, DBDATE(strCP142), 1) '(¤£§t·í¤Ñ)
               '«ü©w¤é´Á¤§«e°e¥ó¡G©Ó¿ì´Á­­¬°«ü©w¤é´Á«e¤T¤é
               ElseIf strCP164 = "2" Then
                  strDate = CompWorkDay(4, DBDATE(strCP142), 1)
               '«ü©w¤é´Á¤§«á°e¥ó¡G©Ó¿ì´Á­­¬°«ü©w¤é´Á«á¤C¤é
               ElseIf strCP164 = "3" Then
                  strDate = CompWorkDay(8, DBDATE(strCP142))
               End If
               'Add By Sindy 2025/1/3
               '°Ó¥Ó®×¥ó¦³«ü©w°e¥ó¤é®É:
               '1.¦p©Ó¿ì´Á­­¤p©ó¤À®×¤é®É¡A½Ð¥H¤À®×¤é+3¤é¬°©Ó¿ì´Á­­
               '2.©Ó¤W¡A­Y+3¤é«á¡A©Ó¿ì´Á­­¤j©ó«ü©w°e¥ó¤é®É¡A«h©Ó¿ì´Á­­§Y¬°«ü©w°e¥ó¤é
               If strDate < strCP149 Then
                  strDate = CompWorkDay(4, strCP149)
                  If strDate > strCP142 Then
                     strDate = strCP142
                  End If
               End If
               '2025/1/3 END
               strExSql = "UPDATE CaseProgress SET CP48 = " & strDate & " " & _
                        "WHERE CP09 = '" & m_CP09 & "' "
               cnnConnection.Execute strExSql
            End If
         End If
      End If
      '2024/1/15 END
   End If
   'End 2003/12/05
   
   'Add By Sindy 2025/3/14 ©Ó¿ì´Á­­¤£¯à¤j©ó¥»©Ò´Á­­,­Y¤j©ó,«h§ï¬°©Ó¿ì´Á­­=¥»©Ò´Á­­
   strExSql = "UPDATE CaseProgress SET CP48=cp06" & _
              " WHERE CP09 = '" & m_CP09 & "' and cp48>0 and cp06>0 and cp48>cp06"
   cnnConnection.Execute strExSql, intI
   '2025/3/14 END
   
   PUB_CountUpdTxCP48 = True
   Set rsAD = Nothing 'Added by Lydia 2024/05/15
End Function

'add by nickc 2007/10/16 ÀË¬d¬P´Á¥|¤­¦¬ªº¤å¡A´Á­­¬O§_¬°°²¤é
Private Function ChkMyWeek(oDate As String) As Boolean
   ChkMyWeek = False
   If Weekday(ChangeWStringToWDateString(strSrvDate(1))) = 5 Or Weekday(ChangeWStringToWDateString(strSrvDate(1))) = 6 Then
       If ChkWorkDay(DBDATE(oDate)) = False Then
           If GetWorkDay(DBDATE(oDate), strSrvDate(1)) <= 2 Then
               ChkMyWeek = True
           End If
       End If
   End If
End Function

'Added by Lydia 2022/08/22 §ó·s°Ó¼Ð®×ªº»ô³Æ¤é; ±qfrm010004.SaveFrame21­×§ï
'Modified by Lydia 2022/09/29 ¶Ç¤J¨t²Î§O,°ê®a,®×¥ó©Ê½è=>ByVal strCP01 As String, ByVal strNa01 As String, ByVal strCP10 As String
Public Sub PUB_SaveByControl(ByVal strCP09 As String, ByVal ProcList As String, ByVal strCP01 As String, ByVal strNA01 As String, ByVal strCP10 As String)
Dim strTmpA As String, intJ As Integer
Dim tmpArr1 As Variant, tmpArr2 As Variant
   
   If strCP09 = "" Or ProcList = "" Then Exit Sub
   tmpArr1 = Split(ProcList, ",")
   
   For intJ = 0 To UBound(tmpArr1)
       If Trim(tmpArr1(intJ)) <> "" Then
          tmpArr2 = Split(tmpArr1(intJ), "|")
          If Trim(tmpArr2(0)) = "EP06" Then
             '¸ê®Æ¬O§_»ô³Æ
             'Added by Lydia 2022/09/29 T¥xÆW®×(¥]§t¤¤¶¡µ{§Ç)¦¬¤å»ô³Æ¤éºÞ¨î, ±Æ°£¯S©w®×¥ó©Ê½è¤£¥Î¿é¤J¤å¥ó»ô³Æ; ¹w³]¤å¥ó»ô³Æ=Y(°Ñ¦Òfrm090801)
             If strCP01 = "T" And strNA01 = "000" And InStr(T®×¦¬¤å»ô³Æ±Æ°£, strCP10) > 0 Then
                 strTmpA = "update engineerprogress set ep06=" & strSrvDate(1) & ",ep36=" & strSrvDate(1) & " where ep02='" & strCP09 & "'"
                 cnnConnection.Execute strTmpA
             Else
             'end 2022/09/29
                 'Memo by Lydia 2018/12/10 T¥xÆW°Ó¥Ó®×=¤å¥ó¬O§_»ô³Æ
                 If Trim(tmpArr2(1)) = "Y" Then
                     strTmpA = "update engineerprogress set ep06=" & strSrvDate(1) & ",ep36=" & strSrvDate(1) & " where ep02='" & strCP09 & "'"
                     cnnConnection.Execute strTmpA
                 ElseIf Trim(tmpArr2(1)) = "N" Then
                     strTmpA = "update engineerprogress set ep06=0,ep36=0 where ep02='" & strCP09 & "'"
                     cnnConnection.Execute strTmpA
                 ElseIf Trim(tmpArr2(1)) = "" Then
                     strTmpA = "update engineerprogress set ep06=null,ep36=null where ep02='" & strCP09 & "'"
                     cnnConnection.Execute strTmpA
                 End If
                 '¥¼»ô³Æ=>¤w»ô³Æ
                 If (tmpArr2(2) = "" Or tmpArr2(2) = "N") And Trim(tmpArr2(1)) = "Y" Then
                     strTmpA = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07)" & _
                              " values('" & strCP09 & "','1','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & "," & strSrvDate(1) & ",'¦¬¤å')"
                     cnnConnection.Execute strTmpA
                 '¨ú®ø»ô³Æ
                 ElseIf tmpArr2(2) = "Y" And (Trim(tmpArr2(1)) = "N" Or Trim(tmpArr2(1)) = "") Then
                     strTmpA = "insert into tmctldate(tcd01,tcd02,tcd03,tcd04,tcd05,tcd06,tcd07)" & _
                              " values('" & strCP09 & "','1','" & strUserNum & "'," & strSrvDate(1) & "," & Right("000000" & ServerTime, 6) & ",null,'¦¬¤å¨ú®ø»ô³Æ')"
                     cnnConnection.Execute strTmpA
                 End If
             End If 'Added by Lydia 2022/09/29
          End If 'If tmparr2(0) = "EP06" Then
          If tmpArr2(0) = "EP34" Then
             '¬O§_·|½Z
              strTmpA = "update engineerprogress set ep34=" & CNULL(Trim(tmpArr2(1))) & " where ep02='" & strCP09 & "'"
              cnnConnection.Execute strTmpA
          End If 'If tmparr2(0) = "EP34" Then
   
          If tmpArr2(0) = "CP122" Then
             '¬O§_«æ¥ó
              strTmpA = "update caseprogress set cp122=" & CNULL(Trim(tmpArr2(1))) & " where cp09='" & strCP09 & "' "
              cnnConnection.Execute strTmpA
          End If 'If tmparr2(0) = "CP122" The
   
          If tmpArr2(0) = "CP143" Then
             '¬d¦W¬O§_»ô³Æ
              strTmpA = "update caseprogress set cp143=" & CNULL(Trim(tmpArr2(1))) & " where cp09='" & strCP09 & "' "
              cnnConnection.Execute strTmpA
          End If 'If tmparr2(0) = "CP143" Then
       End If
   Next intJ
   
End Sub

'Added by Lydia 2022/09/05 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G·s¼W±M§Q°ò¥»ÀÉ(±qfrm010005.InsertPatentDatabase©â¥X¨Ó)
Private Function InsertPatentDB(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intChoose As Integer, _
                ByRef m_Pa() As String, ByRef mCP() As String, ByVal mCU30 As String, ByVal mInventorNo As String, ByVal mChkVal As String, Optional ByRef IsSaveData As Boolean, _
                Optional ByVal pType As String, Optional ByVal pCaseNo As String, Optional ByRef RetVal As String, Optional ByVal mTCTVal As String, Optional ByRef mTCTList As String) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
'reTurnVal : ¦^¶Ç­È
'mChkVal¡G¶Ç¤J¨ä¥L¾Þ§@µ²ªG
'mTctVal ¡G¶Ç¤Jµe­±¦³Ãö©R¦W§@·~ªº¸ê®Æ
'mTctList ¡G¦^¶Ç©R¦W§@·~¤@¨Ö²£¥Í¤§¦¬¤å¸¹
Dim strAutoNumber As String
Dim np13 As String, np14 As String, bolRt As Boolean
Dim np14ForCP41 As String, np14ForCP42 As String 'Add By Sindy 2025/1/24
Dim bolError As Boolean, intW As Integer
Dim strCustomer(4) As String
Dim varInventorNo As Variant
Dim strInventor(100) As String
Dim m_strCPM34 As String
Dim bolNoAutoCP14 As Boolean
Dim adoquery As New ADODB.Recordset
Dim m_CaseNaTmp() As String  '¯S®íºÞ¨î¤§ÃöÁp®×
'ªk«ß©Ò®×·½¦¬¤å
Dim m_LOS02 As String '®×·½®×¥óÃþ«¬
Dim m_LOS15 As String '®×·½³æ¸¹
Dim rsQD As New ADODB.Recordset
Dim m_bMRecvBatch As Boolean '«H¥ó¨R¾P¦h®×¦¬¤å
Dim m_bolRecvOK As Boolean 'Add By Sindy 2022/7/8 ¬O§_¦¬§¹¤å
Dim m_strMCR11 As String 'Add By Sindy 2022/7/8 ¦h®×¦¬¤å®É,²Ä¤@µ§ªºÁ`¦¬¤å¸¹
Dim m_strIR01 As String, m_strIR02 As String, m_strIR03 As String, m_strIR04 As String '«H¥ó¨R¾PPK
Dim tmpArr As Variant
'Added by Morgan 2024/12/25 »âÃÒ/¦~¶O¦¬¤å¬O§_¦Û°Ê¤º³¡¦¬¤å414¥Ó½Ð´_Åv
Dim bolAutoFCP414 As Boolean, strFCP414BCP06 As String, strFCP414BCP07 As String, strFCP414BCP09 As String, strFCP414BCP16 As String, strFCP414BCP17 As String, strFCP414BCP18 As String, strFCP414BCP48 As String
Dim strChkDate As String, strAlert As String
'end 2024/12/25
   
   If IsSaveData = True Then
      Exit Function
   End If
   IsSaveData = True
   
'*********¯S®íºÞ¨îªºÅÜ¼Æ*************
   If pType = "CFP­^°ê²æ¼Ú®×" And pCaseNo <> "" Then
       ReDim m_CaseNaTmp(1 To TF_PA)
       Call ChgCaseNo(pCaseNo, m_CaseNaTmp)
   'Modify By Sindy 2025/8/18 µo¥Í¤F®×·½+«H¥ó¨R¾P ex:FCP-057445/FCL-011034
   ' mark,§ï¦b¤U¦C¥t¥~¼gif
   'Modify By Sindy 2023/5/31
   'ElseIf InStr(pType, "¥~±M«H¥ó¨R¾P") > 0 And pCaseNo <> "" Then
'   ElseIf InStr(pType, "«H¥ó¨R¾P") > 0 And pCaseNo <> "" Then
'   '2023/5/31 END
'       '¦]¬°¥~±M«H¥ó¨R¾P¬O±q¨t²Î¦¬¥ó°Ï¡A©Ò¥H¤£·|¸ò®×·½¦¬¤å(frm090801)­«Å|
'       m_CaseNaTmp = Split(pCaseNo, ",")
'       m_strIR01 = m_CaseNaTmp(0)
'       m_strIR02 = m_CaseNaTmp(1)
'       m_strIR03 = m_CaseNaTmp(2)
'       m_strIR04 = m_CaseNaTmp(3)
'       ReDim m_CaseNaTmp(1 To 4)  '¹w³]°}¦CÁ×§Kµ{¦¡¥X¿ù
'       If InStr(pType, "¦h®×¦¬¤å") > 0 Then m_bMRecvBatch = True
   Else
       ReDim m_CaseNaTmp(1 To 4) '¹w³]°}¦CÁ×§Kµ{¦¡¥X¿ù
       'Modify By Sindy 2025/8/18
       'If pType = "LOS®×·½¦¬¤å" And pCaseNo <> "" Then
       If InStr(pType, "LOS®×·½¦¬¤å") > 0 And pCaseNo <> "" Then
       '2025/8/18 END
           m_LOS02 = Mid(pCaseNo, 1, InStr(pCaseNo, ",") - 1) '®×·½®×¥óÃþ«¬
           'm_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1) '®×·½³æ¸¹
           m_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1, 8) '®×·½³æ¸¹ 'Modify By Sindy 2025/8/18 +, 8)
       End If
   End If
   'Modify By Sindy 2025/8/18
   If InStr(pType, "«H¥ó¨R¾P") > 0 And pCaseNo <> "" Then
      m_CaseNaTmp = Split(pCaseNo, "-")
      If InStr(pCaseNo, "-") > 0 Then
         strExc(10) = m_CaseNaTmp(1)
      Else
         strExc(10) = m_CaseNaTmp(0)
      End If
      m_CaseNaTmp = Split(strExc(10), ",")
   '2025/8/18 END
      m_strIR01 = m_CaseNaTmp(0)
      m_strIR02 = m_CaseNaTmp(1)
      m_strIR03 = m_CaseNaTmp(2)
      m_strIR04 = m_CaseNaTmp(3)
      If InStr(pType, "¦h®×¦¬¤å") > 0 Then m_bMRecvBatch = True
      'ReDim m_CaseNaTmp(1 To 4)  '¹w³]°}¦CÁ×§Kµ{¦¡¥X¿ù
      If InStr(pType, "LOS®×·½¦¬¤å") > 0 Then pType = "LOS®×·½¦¬¤å" 'Add By Sindy 2025/8/18
   End If
   '2025/8/18 END
'***********************************
   strTmp1(0) = "select cpm34 from casepropertymap where cpm01='" & mCP(1) & "' and cpm02='" & mCP(10) & "' "
   Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
   If intJ = 1 Then
       m_strCPM34 = "" & rsQD.Fields("cpm34")
   End If
   
On Error GoTo ErrHand
   '¶Ç¤J0¬°­«½Æ¤§¥»©Ò®×¸¹(·s¼WÂÂ®×)¡A1¬°¥¿½T¤§¥»©Ò®×¸¹(·s¼W·s®×)
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.BeginTrans
   End If
   If intSaveMode = 1 Then
      'Modified by Lydia 2022/08/11 ®³±¼PA46
      Cls001SetPAFileProperty mCP(10), m_Pa(23)
      If m_Pa(2) = "" Then
         If ClsPDGetAutoNumber(m_Pa(1), strAutoNumber, True, False) Then
            m_Pa(2) = strAutoNumber
         Else
            bolError = True
         End If
      End If
      If bolError = False Then
         If ClsPDGetSystemKind(m_Pa(1), , , intW) Then
            m_Pa(85) = IIf(intW = 2, 2, 1)
            mCP(2) = m_Pa(2)
            'Modify by Toni   2008/8/26   ¬°µo©ú¤H
            varInventorNo = Split(mInventorNo, ",")
            For intJ = 0 To UBound(varInventorNo)
               strInventor(intJ) = varInventorNo(intJ)
            Next
            For intJ = intJ + 1 To 99
               strInventor(intJ) = ""
            Next
            'Add By Sindy 2014/11/6 §ó·s±M§Qµo©ú¤HÀÉ
            For intJ = 0 To 99
               If strInventor(intJ) <> "" Then
                  mStrSql = "INSERT into patentInventor(pi01,pi02,pi03,pi04,pi05,pi06) VALUES(" & _
                           CNULL(m_Pa(1)) & "," & CNULL(m_Pa(2)) & "," & CNULL(m_Pa(3)) & "," & CNULL(m_Pa(4)) & "," & intJ + 1 & ",'" & strInventor(intJ) & "')"
                  Pub_SeekTbLog mStrSql 'Add By Sindy 2017/8/23
                  cnnConnection.Execute mStrSql
               Else
                  Exit For
               End If
            Next intJ
            '2014/11/6 END
            
            m_Pa(161) = GetReceiptCmp(Left(m_Pa(26), 8), Mid(m_Pa(26), 9, 1), m_Pa(1), m_Pa(9)) 'Added by Amy 2018/10/11 ++¦¬¾Ú¤½¥q§Opa161
            'Added by Lydia 2020/11/19 CFP­^°ê²æ¼Ú®×ºÞ¨î¡G·s¼W­^°ê®×®É¦P®É§â¼Ú·ù®×¬ÛÃöÄæ¦ì±a¹L¨Ó(°Ñ¦ÒPUB_SaveCountry)
            If m_Pa(1) = "CFP" And pType = "CFP­^°ê²æ¼Ú®×" And m_CaseNaTmp(1) <> "" And m_CaseNaTmp(2) <> "" Then
                If PUB_ReadPatentData(m_CaseNaTmp(), m_CaseNaTmp(1), m_CaseNaTmp(2), m_CaseNaTmp(3), m_CaseNaTmp(4)) Then
                   strTmp1(0) = "": strTmp1(1) = ""
                   For intJ = 5 To TF_PA
                       Select Case intJ
                          Case 92, 93, 94, 95, 96, 97, 108, 136, 137, 138 'Create + Update, ¾P¨÷¤é
                          Case 9  '¥Ó½Ð°ê®a
                              strTmp1(0) = strTmp1(0) & "PA" & Format(intJ, "00") & "," 'Insert
                              strTmp1(1) = strTmp1(1) & " '201' as PA09, " 'Select
                          Case 22 '±M§Q¸¹¼Æ: ±M§Q¸¹¥H¼Ú·ù³]­p±M§Q¸¹(®³±¼-²Å¸¹)«e¥[¤W9
                              strTmp1(0) = strTmp1(0) & "PA" & Format(intJ, "00") & ","
                              strTmp1(1) = strTmp1(1) & " '9'||REPLACE(PA22,'-','') AS PA" & Format(intJ, "00") & ", "
                          Case 91  '®×¥ó³Æµù: ¥[µù¼Ú·ù®×®×¸¹
                              strTmp1(0) = strTmp1(0) & "PA" & Format(intJ, "00") & ","
                              strTmp1(1) = strTmp1(1) & CNULL("¼Ú·ù®×®×¸¹¡G" & m_CaseNaTmp(1) & m_CaseNaTmp(2) & m_CaseNaTmp(3) & m_CaseNaTmp(4) & ";") & "||PA" & Format(intJ, "00") & " AS PA" & Format(intJ, "00") & ","
                          Case Else
                              strTmp1(0) = strTmp1(0) & "PA" & Format(intJ, "00") & ","
                              strTmp1(1) = strTmp1(1) & "PA" & Format(intJ, "00") & ","
                       End Select
                   Next
                   strTmp1(0) = Left(strTmp1(0), Len(strTmp1(0)) - 1)
                   strTmp1(1) = Left(strTmp1(1), Len(strTmp1(1)) - 1)
                   mStrSql = "INSERT INTO PATENT (PA01,PA02,PA03,PA04," & strTmp1(0) & ") " & _
                               "SELECT '" & m_Pa(1) & "' as PA01,'" & m_Pa(2) & "' as PA02,'" & m_Pa(3) & "' as PA03,'" & m_Pa(4) & "' as pa04, " & strTmp1(1) & _
                               " FROM PATENT WHERE pa01='" & m_CaseNaTmp(1) & "' and pa02='" & m_CaseNaTmp(2) & "' and pa03='" & m_CaseNaTmp(3) & "' and pa04='" & m_CaseNaTmp(4) & "' "
                   cnnConnection.Execute mStrSql
                   
                   If mCP(7) <> "" Then PUB_UpdUkPayYr mCP(7), m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4) 'Added by Morgan 2020/12/8 §ó·s­^°ê®×Ãº¶O¬ö¿ý
                End If
            Else
            'end 2020/11/19
                'Modify by Amy 2018/10/11 +¦¬¾Ú¤½¥q§Opa161
                'Modified by Lydia 2022/08/17 (¨Ö¤J) +PA47¤À©Ò®×¸¹,PA48«È¤á®×¥ó®×¸¹
                'Modify By Sindy 2022/12/7 +PA178
                'Modify By Sindy 2023/3/30 +PA179
                mStrSql = "insert into patent (pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa08,pa09,pa23,pa26," & _
                   "pa27,pa28,pa29,pa30,pa46,pa47,pa48,pa75,pa17,pa77,pa149,pa51,pa52,pa53,pa54,pa55,pa56,pa158,pa150,pa161,pa176,pa178,pa179) " & _
                   "values (" & CNULL(m_Pa(1)) & "," & CNULL(m_Pa(2)) & "," & CNULL(m_Pa(3)) & "," & CNULL(m_Pa(4)) & "," & CNULL(ChgSQL(m_Pa(5))) & "," & _
                   CNULL(Replace(m_Pa(6), "'", "''")) & "," & CNULL(ChgSQL(m_Pa(7))) & "," & CNULL(m_Pa(8)) & "," & CNULL(m_Pa(9)) & "," & CNULL(m_Pa(23)) & "," & CNULL(m_Pa(26)) & "," & CNULL(m_Pa(27)) & "," & _
                   CNULL(m_Pa(28)) & "," & CNULL(m_Pa(29)) & "," & CNULL(m_Pa(30)) & "," & CNULL(m_Pa(46)) & "," & CNULL(m_Pa(47)) & "," & CNULL(m_Pa(48)) & "," & CNULL(m_Pa(75)) & ", ''," & CNULL(m_Pa(77)) & "," & CNULL(m_Pa(149)) & "," & _
                   CNULL(ChgSQL(m_Pa(51))) & "," & CNULL(ChgSQL(m_Pa(52))) & "," & CNULL(ChgSQL(m_Pa(53))) & "," & CNULL(ChgSQL(m_Pa(54))) & "," & CNULL(ChgSQL(m_Pa(55))) & "," & CNULL(ChgSQL(m_Pa(56))) & "," & CNULL((m_Pa(158))) _
                   & "," & CNULL(m_Pa(150)) & "," & CNULL(ChgSQL(m_Pa(161))) & "," & CNULL(m_Pa(176)) & "," & CNULL(m_Pa(178)) & "," & CNULL(m_Pa(179)) & ")"
                cnnConnection.Execute mStrSql
                '2014/11/6 END
                strCustomer(0) = m_Pa(26)
                strCustomer(1) = m_Pa(27)
                strCustomer(2) = m_Pa(28)
                strCustomer(3) = m_Pa(29)
                strCustomer(4) = m_Pa(30)
                'Memo by Lydia 2020/11/19 CFP­^°ê²æ¼Ú®×ºÞ¨î¡G·s¼W­^°ê®×®É¦P®É§â¼Ú·ù®×¬ÛÃöÄæ¦ì±a¹L¨Ó¡A©Ò¥H¤£­nÅÜ§ó¸ê®Æ
                For intJ = 0 To 4
                       mStrSql = "update patent set pa" + Format(31 + intJ) + "=(select cu23 from customer where cu01=" + CNULL(Mid(strCustomer(intJ), 1, 8)) + " and cu02=" + CNULL(Mid(strCustomer(intJ), 9, 1)) + _
                          "),pa" + Format(36 + intJ) + "=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(strCustomer(intJ), 1, 8)) + " and cu02=" + CNULL(Mid(strCustomer(intJ), 9, 1)) + _
                          "),pa" + Format(41 + intJ) + "=(select cu29 from customer where cu01=" + CNULL(Mid(strCustomer(intJ), 1, 8)) + " and cu02=" + CNULL(Mid(strCustomer(intJ), 9, 1)) + ") where pa01=" + CNULL(m_Pa(1)) + " and pa02=" + CNULL(m_Pa(2)) + " and pa03=" + CNULL(m_Pa(3)) + " and pa04=" + CNULL(m_Pa(4))
                       cnnConnection.Execute mStrSql
                Next
            End If 'Added by Lydia 2020/11/19
           
            mCP(31) = "Y"
         Else
            bolError = True
         End If
      Else
         bolError = True
      End If
   
   'Modify By Sindy 2023/5/10 ¦³µo©ú¤H¸ê®Æ¥u·s¼W ex:P-130984(µo©ú¥Ó½Ð«D±¾·s®×,«e­±¦³¦¬±M§Q½Õ¬d)
   ElseIf mInventorNo <> "" Then
      '¬°µo©ú¤H
      varInventorNo = Split(mInventorNo, ",")
      For intJ = 0 To UBound(varInventorNo)
         strInventor(intJ) = varInventorNo(intJ)
      Next
      For intJ = intJ + 1 To 99
         strInventor(intJ) = ""
      Next
      For intJ = 0 To 99
         If strInventor(intJ) <> "" Then
            strTmp1(0) = "select * from patentInventor where pi01='" & mCP(1) & "' and pi02='" & mCP(2) & "' and pi03='" & mCP(3) & "' and pi04='" & mCP(4) & "' and pi06='" & strInventor(intJ) & "'"
            intK = 1
            Set rsQD = ClsLawReadRstMsg(intK, strTmp1(0))
            If intK = 0 Then
               strTmp1(0) = "select max(pi05) from patentInventor where pi01='" & mCP(1) & "' and pi02='" & mCP(2) & "' and pi03='" & mCP(3) & "' and pi04='" & mCP(4) & "'"
               intK = 1
               Set rsQD = ClsLawReadRstMsg(intK, strTmp1(0))
               If intK = 1 Then
                  intK = Val("" & rsQD.Fields(0)) 'Modify By Sindy 2023/6/7
                  mStrSql = "INSERT into patentInventor(pi01,pi02,pi03,pi04,pi05,pi06) VALUES(" & _
                           CNULL(mCP(1)) & "," & CNULL(mCP(2)) & "," & CNULL(mCP(3)) & "," & CNULL(mCP(4)) & "," & intK + 1 & ",'" & strInventor(intJ) & "')"
                  Pub_SeekTbLog mStrSql 'Add By Sindy 2017/8/23
                  cnnConnection.Execute mStrSql
               End If
            End If
         Else
            Exit For
         End If
      Next intJ
      '2023/5/10 END
   End If
   
   If bolError = False Then
      'Modify By Cheng 2002/01/09
      'Modified by Lydia 2018/05/09 §ï¦¨¼Ò²Õ
      'Modified by Lydia 2025/06/25
      'Pub_SetPAIsCase m_Pa(1), mCP(10), mCP(26)
      If PUB_GetCPMbyCP10(m_Pa(1), mCP(10), "cpm05") = "N" Then
          mCP(26) = "N"
      End If
      'end 2018/05/09
      mCP(20) = ""

      'Add By Sindy 2022/3/31
      If InStr(mChkVal, "m_bolFMP") > 0 Then
         mCP(48) = Pub_GetHandleDay("FCP", m_Pa(9), mCP(10), , mCP(6))
      '2022/3/31 END
      ElseIf m_Pa(1) = "FCP" Or m_Pa(1) = "FG" Then
         'Add by Morgan 2008/8/28 «D¨Ò¥~ªº®×¥ó©Ê½è­n¹w³]©Ó¿ì´Á­­
         'Modify by Morgan 2008/10/23
         'Ciba Y45697ªº¦~¶O©Ó¿ì´Á­­±¾15­Ó¤u§@¤Ñ
         If m_Pa(75) = "Y45697000" And mCP(10) = "605" Then
            mCP(48) = CompWorkDay(15, strSrvDate(1))
            If Val(mCP(6)) > 0 And Val(mCP(48)) > Val(mCP(6)) Then
               mCP(48) = mCP(6)
            End If
         'end 2008/10/23
         
         'Added by Morgan 2012/7/13
         '¥[³t¼f¬d­n§PÂ_¤w¿é¤J³qª¾¹ê¼f¤é¤~±¾©Ó¿ì´Á­­
         'Modified by Morgan 2024/11/18 +477¦A¼f¬d¥[³t¼f¬d¨Ã§ï¥Î±M¥Î¼Ò²Õ§PÂ_
         'ElseIf mCP(10) = "422" Then
         '   If PUB_ChkCPExist(m_Pa(), "1204") Then
         ElseIf mCP(10) = "422" Or mCP(10) = "447" Then
            If PUB_Chk1204(m_Pa) Then
         'end 2024/11/13
               mCP(48) = Pub_GetHandleDay("FCP", "000", mCP(10), , mCP(6))
            End If
         'end 2012/7/13
         
         'Add By Sindy 2021/6/24 968¦^´_»¡©ú®Ñ®Õ¾\
         ElseIf mCP(10) = "968" Then
            mCP(48) = Pub_GetHandleDay("FCP", "000", mCP(10), , mCP(6), , , m_Pa(1) & "-" & m_Pa(2) & "-" & m_Pa(3) & "-" & m_Pa(4))
            
         ElseIf InStr(SkipCasePtyList, mCP(10)) = 0 Then
            mCP(48) = Pub_GetHandleDay("FCP", "000", mCP(10), , mCP(6))
            'Y54732000 & X30299000²Õ¦X¤§¦^¥N©Ó¿ì´Á­­¤U­±¥t¦³§ó·s
         End If
         
         'Add By Sindy 2021/4/29 ¤£¬O¥DºÞ¾÷Ãö´Á­­
         If m_strCPM34 = "N" And strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            '(2)¦¬¤å®ÉµL³]¥»©Ò´Á­­¡A¥H©Ó¿ì´Á­­¡Ï5­Ó¤u§@¤Ñ¬°¥»©Ò´Á­­
            If Val(mCP(6)) = 0 Then
               mCP(6) = PUB_GetFCPOurDeadline(DBDATE(mCP(48)), , , , "N")
            '(1)¦¬¤å®É¦³³]¥»©Ò´Á­­¡A¦Û°Ê³Æµù:¥»©Ò´Á­­¬°yyy/mm/dd(¥»©Ò´Á­­)
            Else
               mCP(64) = "¥»©Ò´Á­­¬°" & ChangeWStringToTDateString(mCP(6)) & ";" & mCP(64)
            End If
         End If
      End If
      
      '2010/11/26 add by sonia Pªº942¹w³]¤£½Ð´Ú
      'Modified by Morgan 2019/8/16 +°ê¥~³¡¦¬¤å±ø¥ó(¥Ø«e®×¥ó©Ê½è¹ï·Óªí³]©w¥u¦³¥~±M¥Î)
      If Left(mCP(12), 1) = "F" And (m_Pa(1) = "FCP" Or m_Pa(1) = "FG" Or m_Pa(1) = "P") Then
         mCP(20) = PUB_GetCP20(m_Pa(1), mCP(10))
      End If
      '2010/11/26 END
      'Modified by Lydia 2024/05/28 §ï¦¨¼Ò²Õ
      ''Added by Lydia 2022/05/03  FCP-062174¼f©w«e¤£¦¬¶O±±¨î:¸É¤W¬O§_¦V«È¤á¦¬´Ú=N
      'If m_Pa(16) = "" And InStr("FCP062174000", m_Pa(1) & m_Pa(2) & m_Pa(3) & m_Pa(4)) > 0 Then
      '    mCP(20) = "N"
      'End If
      ''end 2022/05/03
      ''Added by Lydia 2022/05/03 FCP-067004®Ö­ã«e¤£¦¬¶O±±¨î¡G¥Ó½Ð¦Ü®Ö­ã(¼È¤£¥]§t»âÃÒ)¤£¦¬¥ô¦ó¦¬¶O (¥]§t³W¶O¤ÎªA°È¶O¡B­Y«È¤á´£AEP¤]¤£¦¬¶O)
      'If m_Pa(16) <> "1" And InStr("FCP067004000", m_Pa(1) & m_Pa(2) & m_Pa(3) & m_Pa(4)) > 0 Then
      '    mCP(20) = "N"
      'End If
      ''end 2022/05/03
      If PUB_GetCP20forSpec(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), m_Pa(16)) = "N" Then
          mCP(20) = "N"
      End If
      'end 2024/05/28
      
      '92.11.3 END
      If ClsPDGetAutoNumber(Left(mCP(9), 1), strAutoNumber, True, True) Then
         If mCP(56) <> "" Then
            mCP(55) = m_Pa(26)
            'Add by Morgan 2006/6/23
            'Åý»P¤H2-5,¨üÅý¤H2-5
            mCP(93) = m_Pa(27)
            mCP(94) = m_Pa(28)
            mCP(95) = m_Pa(29)
            mCP(96) = m_Pa(30)
            'end 2006/6/23
         End If
         mCP(9) = mCP(9) + strAutoNumber
         'Modify By Sindy 2025/1/24 + np14ForCP41,np14ForCP42
         bolRt = Cls001GetNextProgressData(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), mCP(10), np13, np14, np14ForCP41, np14ForCP42)

         'Modify By Sindy 2012/11/06 +CP150
'         If m_Pa(23) <> "1" And mCP(31) = "Y" Then
'            'Modify by Morgan 2006/6/23 ¥[cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96
'            If bolRt Then
'               'Add By Sindy 2012/11/06 ¦³¡¹¡¹ªºÀ³¦¬±b´ÚÃ±®Ö±±ºÞ (¨Ö¤J)
'               'Modify By Sindy 2022/9/28 +,cp140
'               mStrSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp13,cp14," & _
'                 "cp16,cp17,cp18,cp19,cp20,cp26,cp31,cp32,cp33,cp34,cp37,cp38,cp39,cp40,CP48,cp55,cp56,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150,cp140) values (" & CNULL(m_Pa(1)) & "," & CNULL(m_Pa(2)) & "," & CNULL(m_Pa(3)) & "," & CNULL(m_Pa(4)) & "," & CNULL(mCP(5)) & "," & _
'                 CNULL(mCP(6)) & "," & CNULL(mCP(7)) & "," & CNULL(np13) & "," & CNULL(mCP(9)) & "," & CNULL(mCP(10)) & "," & CNULL(mCP(11)) & "," & CNULL(mCP(13)) & "," & CNULL(mCP(14)) & "," & CNULL(mCP(16)) & "," & _
'                 CNULL(mCP(17)) & "," & CNULL(mCP(18)) & "," & CNULL(mCP(19)) & "," & CNULL(mCP(20)) & "," & CNULL(mCP(26)) & "," & CNULL(mCP(31)) & "," & CNULL(mCP(32)) & ", " & CNULL(mCP(33)) & ", " & CNULL(mCP(34)) & ", " & CNULL(ChgSQL(m_Pa(5))) & ", " & CNULL(ChgSQL(m_Pa(6))) & ", " & CNULL(ChgSQL(m_Pa(7))) & "," & CNULL(ChgSQL(np14)) & "," & CNULL(mCP(48), True) & "," & CNULL(mCP(55)) & "," & CNULL(mCP(56)) & "," & CNULL(ChgSQL(mCP(64))) & "," & _
'                 CNULL(mCP(89)) & "," & CNULL(mCP(90)) & "," & CNULL(mCP(91)) & "," & CNULL(mCP(92)) & "," & CNULL(mCP(93)) & "," & CNULL(mCP(94)) & "," & CNULL(mCP(95)) & "," & CNULL(mCP(96)) & "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + ")"
'            Else
'               mStrSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp13,cp14," & _
'                 "cp16,cp17,cp18,cp19,cp20,cp26,cp31,cp32,cp33,cp34,cp37,cp38,cp39,CP48,cp55,cp56,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150,cp140) values (" & CNULL(m_Pa(1)) & "," & CNULL(m_Pa(2)) & "," & CNULL(m_Pa(3)) & "," & CNULL(m_Pa(4)) & "," & CNULL(mCP(5)) & "," & _
'                 CNULL(mCP(6)) & "," & CNULL(mCP(7)) & "," & CNULL(mCP(9)) & "," & CNULL(mCP(10)) & "," & CNULL(mCP(11)) & "," & CNULL(mCP(13)) & "," & CNULL(mCP(14)) & "," & CNULL(mCP(16)) & "," & _
'                 CNULL(mCP(17)) & "," & CNULL(mCP(18)) & "," & CNULL(mCP(19)) & "," & CNULL(mCP(20)) & "," & CNULL(mCP(26)) & "," & CNULL(mCP(31)) & "," & CNULL(mCP(32)) & ", " & CNULL(mCP(33)) & ", " & CNULL(mCP(34)) & ", " & CNULL(ChgSQL(m_Pa(5))) & ", " & CNULL(ChgSQL(m_Pa(6))) & ", " & CNULL(ChgSQL(m_Pa(7))) & "," & CNULL(mCP(48), True) & "," & CNULL(mCP(55)) & "," & CNULL(mCP(56)) & "," & CNULL(ChgSQL(mCP(64))) & "," & _
'                 CNULL(mCP(89)) & "," & CNULL(mCP(90)) & "," & CNULL(mCP(91)) & "," & CNULL(mCP(92)) & "," & CNULL(mCP(93)) & "," & CNULL(mCP(94)) & "," & CNULL(mCP(95)) & "," & CNULL(mCP(96)) & "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + ")"
'            End If
'         Else
'            If bolRt Then
'               mStrSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp13,cp14," & _
'                 "cp16,cp17,cp18,cp19,cp20,cp26,cp31,cp32,cp33,cp34,cp40,CP48,cp55,cp56,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150,cp140) values (" & CNULL(m_Pa(1)) & "," & CNULL(m_Pa(2)) & "," & CNULL(m_Pa(3)) & "," & CNULL(m_Pa(4)) & "," & CNULL(mCP(5)) & "," & _
'                 CNULL(mCP(6)) & "," & CNULL(mCP(7)) & "," & CNULL(np13) & "," & CNULL(mCP(9)) & "," & CNULL(mCP(10)) & "," & CNULL(mCP(11)) & "," & CNULL(mCP(13)) & "," & CNULL(mCP(14)) & "," & CNULL(mCP(16)) & "," & _
'                 CNULL(mCP(17)) & "," & CNULL(mCP(18)) & "," & CNULL(mCP(19)) & "," & CNULL(mCP(20)) & "," & CNULL(mCP(26)) & "," & CNULL(mCP(31)) & "," & CNULL(mCP(32)) & ", " & CNULL(mCP(33)) & ", " & CNULL(mCP(34)) & "," & CNULL(ChgSQL(np14)) & "," & CNULL(mCP(48), True) & "," & CNULL(mCP(55)) & "," & CNULL(mCP(56)) & "," & CNULL(ChgSQL(mCP(64))) & "," & _
'                 CNULL(mCP(89)) & "," & CNULL(mCP(90)) & "," & CNULL(mCP(91)) & "," & CNULL(mCP(92)) & "," & CNULL(mCP(93)) & "," & CNULL(mCP(94)) & "," & CNULL(mCP(95)) & "," & CNULL(mCP(96)) & "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + ")"
'            Else
'               mStrSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp13,cp14," & _
'                 "cp16,cp17,cp18,cp19,cp20,cp26,cp31,cp32,cp33,cp34,CP48,cp55,cp56,CP64,cp89,cp90,cp91,cp92,cp93,cp94,cp95,cp96,cp150,cp140) values (" & CNULL(m_Pa(1)) & "," & CNULL(m_Pa(2)) & "," & CNULL(m_Pa(3)) & "," & CNULL(m_Pa(4)) & "," & CNULL(mCP(5)) & "," & _
'                 CNULL(mCP(6)) & "," & CNULL(mCP(7)) & "," & CNULL(mCP(9)) & "," & CNULL(mCP(10)) & "," & CNULL(mCP(11)) & "," & CNULL(mCP(13)) & "," & CNULL(mCP(14)) & "," & CNULL(mCP(16)) & "," & _
'                 CNULL(mCP(17)) & "," & CNULL(mCP(18)) & "," & CNULL(mCP(19)) & "," & CNULL(mCP(20)) & "," & CNULL(mCP(26)) & "," & CNULL(mCP(31)) & "," & CNULL(mCP(32)) & ", " & CNULL(mCP(33)) & ", " & CNULL(mCP(34)) & "," & CNULL(mCP(48), True) & "," & CNULL(mCP(55)) & "," & CNULL(mCP(56)) & "," & CNULL(ChgSQL(mCP(64))) & "," & _
'                 CNULL(mCP(89)) & "," & CNULL(mCP(90)) & "," & CNULL(mCP(91)) & "," & CNULL(mCP(92)) & "," & CNULL(mCP(93)) & "," & CNULL(mCP(94)) & "," & CNULL(mCP(95)) & "," & CNULL(mCP(96)) & "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + ")"
'            End If
'         End If
         'Modify By Sindy 2023/1/12 +,cp141,cp142,cp86
         'Modify By Sindy 2023/4/18 +,cp151
         'Modify By Sindy 2023/12/12 +,cp164
         mStrSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp13,cp14," & _
                 "cp16,cp17,cp18,cp19,cp20,cp26,cp31,cp32,cp33,cp34,CP48,cp55,cp56,CP64,cp89,cp90,cp91,cp92,cp93," & _
                 "cp94,cp95,cp96,cp150,cp140,cp141,cp142,cp86,cp151,cp164) values (" & CNULL(m_Pa(1)) & "," & CNULL(m_Pa(2)) & "," & CNULL(m_Pa(3)) & "," & CNULL(m_Pa(4)) & "," & CNULL(mCP(5)) & "," & _
                 CNULL(mCP(6)) & "," & CNULL(mCP(7)) & "," & CNULL(mCP(9)) & "," & CNULL(mCP(10)) & "," & CNULL(mCP(11)) & "," & CNULL(mCP(13)) & "," & CNULL(mCP(14)) & "," & CNULL(mCP(16)) & "," & _
                 CNULL(mCP(17)) & "," & CNULL(mCP(18)) & "," & CNULL(mCP(19)) & "," & CNULL(mCP(20)) & "," & CNULL(mCP(26)) & "," & CNULL(mCP(31)) & "," & CNULL(mCP(32)) & ", " & CNULL(mCP(33)) & "," & _
                 CNULL(mCP(34)) & "," & CNULL(mCP(48), True) & "," & CNULL(mCP(55)) & "," & CNULL(mCP(56)) & "," & CNULL(ChgSQL(mCP(64))) & "," & _
                 CNULL(mCP(89)) & "," & CNULL(mCP(90)) & "," & CNULL(mCP(91)) & "," & CNULL(mCP(92)) & "," & CNULL(mCP(93)) & "," & CNULL(mCP(94)) & "," & CNULL(mCP(95)) & "," & CNULL(mCP(96)) & "," + _
                 CNULL(mCP(150)) + "," + CNULL(mCP(140)) + "," + CNULL(mCP(141)) + "," + CNULL(mCP(142)) + "," + CNULL(mCP(86)) + "," + CNULL(mCP(151)) + "," + CNULL(mCP(164)) + ")"
         cnnConnection.Execute mStrSql
         If bolRt Then
            'Modify By Sindy 2025/1/24 + np14ForCP41,np14ForCP42
            mStrSql = "update caseprogress set CP08=" + CNULL(np13) + _
                      ",cp40=" + CNULL(ChgSQL(np14)) + _
                      ",cp41=" + CNULL(ChgSQL(np14ForCP41)) + _
                      ",cp42=" + CNULL(ChgSQL(np14ForCP42))
            mStrSql = mStrSql + " where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
         End If
         If m_Pa(23) <> "1" And mCP(31) = "Y" Then
            mStrSql = "update caseprogress set cp37=" + CNULL(ChgSQL(m_Pa(5))) + _
                      ",cp38=" + CNULL(ChgSQL(m_Pa(6))) + _
                      ",cp39=" + CNULL(ChgSQL(m_Pa(7))) + _
                      " where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
         End If
         '2023/1/12 END
         
         'Add By Sindy 2022/12/7 ¼W¥[§ó·s¥xÆWÃÒ®Ñ§Î¦¡
         If m_Pa(178) <> "" Then
            mStrSql = "update patent set pa178=" + CNULL(m_Pa(178))
            mStrSql = mStrSql & " where pa01=" + CNULL(m_Pa(1)) + " and pa02=" + CNULL(m_Pa(2)) + " and pa03=" + CNULL(m_Pa(3)) + " and pa04=" + CNULL(m_Pa(4))
            cnnConnection.Execute mStrSql
         End If
         '2022/12/7 END
         'Add By Sindy 2023/3/30 +PA179
         If m_Pa(179) <> "" Then
            mStrSql = "update patent set pa179=" + CNULL(m_Pa(179))
            mStrSql = mStrSql & " where pa01=" + CNULL(m_Pa(1)) + " and pa02=" + CNULL(m_Pa(2)) + " and pa03=" + CNULL(m_Pa(3)) + " and pa04=" + CNULL(m_Pa(4))
            cnnConnection.Execute mStrSql
         End If
         '2023/3/30 END
         
         'add by sonia 2019/7/31 Y54732000 & X30299000²Õ¦X,¥B·|½Z924µo¤å«á,·s®×Â½Ä¶201µo¤å«e¦¬¤å¤§¦^¥N902,³]¦^¥N¬ÛÃö¦¬¤å¸¹±¾·|½Z,©Ó¿ì´Á­­±¾·s®×Â½Ä¶ªº¥»©Ò´Á­­
         If m_Pa(75) = "Y54732000" And Left(m_Pa(26), 8) = "X3029900" And mCP(10) = "902" Then
            strTmp1(0) = "select c2.cp06,c1.cp09 from caseprogress c1,caseprogress c2 where c1.cp01='" & m_Pa(1) & "' and c1.cp02='" & m_Pa(2) & "' and c1.cp03='" & m_Pa(3) & "' and c1.cp04='" & m_Pa(4) & "' and c1.cp10='924' and c1.cp27>0 " & _
                        "   and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and '201'=c2.cp10(+) and c2.cp158=0"
            intJ = 1
            Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
            If intJ = 1 Then
               mStrSql = "update caseprogress set cp43='" & "" & rsQD(1) & "',cp48=" & "" & rsQD(0) & " where cp09=" & CNULL(mCP(9))
               cnnConnection.Execute mStrSql
            End If
         End If
         'end 2019/7/31
      
         'Modified by Morgan 2012/4/25 +cp71(Àu¥ýÅv¥÷¼Æ)
         'Modified by Morgan 2012/6/20 +cp118(¹q¤l°e¥ó)
         'Modified by Lydia 2022/09/20 ¦]¬°P®×¦³Trigger·|¦Û°Ê³]©w¹q¤l°e¥óCP118 = Y, ©Ò¥H§ï¦¨¨â­Ó§PÂ_
         'mStrSQL = "update caseprogress set cp12=(select st15 from staff where st01=" & CNULL(mCP(13)) & ") , cp118=" & CNULL(mCP(118)) & IIf(mCP(71) <> "", ", cp71=" & CNULL(mCP(71)), "") & " where cp09=" & CNULL(mCP(9))
         mStrSql = IIf(mCP(118) = "YY", ", CP118='Y' ", IIf(mCP(118) = "YN", ", CP118=null ", ""))
         mStrSql = "update caseprogress set cp12=(select st15 from staff where st01=" & CNULL(mCP(13)) & ") " & mStrSql & IIf(mCP(71) <> "", ", cp71=" & CNULL(mCP(71)), "") & " where cp09=" & CNULL(mCP(9))
         'end 2022/09/20
         cnnConnection.Execute mStrSql
         
         'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G¥xÆW®×B1¡BB2¤ÎC¦¬¤å®É¡A¼W¥["®×·½³æ¸¹"Äæ¦ì¤@©w­n¿é¤J¡A¨Ã±N®×·½³æ¸¹§ó·s¦Ü¸Óµ§¦¬¤åªºCP162¡C
         If intModifyKind = 0 And m_Pa(9) = "000" And (m_Pa(1) = "FCP" Or m_Pa(1) = "P") And m_LOS02 <> "" And m_LOS15 <> "" Then
              If Left(m_LOS02, 1) = "B" Or Left(m_LOS02, 1) = "C" Then
                  mStrSql = "update caseprogress set CP162='" & m_LOS15 & "' where cp09='" & mCP(9) & "' "
                  cnnConnection.Execute mStrSql
              End If
         End If
         'end 2020/05/20
         
         'Add By Sindy 2009/07/06
         If mCP(53) <> "" And mCP(54) <> "" Then
            If mCP(10) = "601" Then
               If Val(mCP(54)) > 0 Then
                  mStrSql = "update caseprogress set cp53=" & mCP(53) & ",cp54=" & mCP(54) & " where cp09=" & CNULL(mCP(9))
               End If
            Else
               mStrSql = "update caseprogress set cp53=" & mCP(53) & ",cp54=" & mCP(54) & " where cp09=" & CNULL(mCP(9))
            End If
            cnnConnection.Execute mStrSql
         End If
         '2009/07/06 End
        
          '­Y¬°±µ¬¢°O¿ý³æ(Âd¥x¦¬¤å)
          'Modify by Morgan 2007/10/26 ¶O¥Î¥i§ï®É¤~°µ¡A§_«h¤w¦¬´Ú¸ê®Æ·|³QÁÙ­ì
          If intChoose = 0 And mCP(60) = "" Then 'mCP(60) = "" => txtPatent(17).Enabled = True
              '¥¼¦¬ª÷ÃB = ¶O¥Î
              mStrSql = "update caseprogress set cp79=cp16 where cp09=" & CNULL(mCP(9))
              cnnConnection.Execute mStrSql
          End If
         'Add By Cheng 2002/05/10
         '­Y¬°¤º³¡¦¬¤å§@·~®É, ®×¥ó¶i«×ÀÉªº¬O§_¦V«È¤á¦¬´Ú³]©w¬°"N"
         If intChoose = 1 Then
            mStrSql = "Update CaseProgress Set CP20='N' Where cp09=" & CNULL(mCP(9))
            cnnConnection.Execute mStrSql
         End If
         
         'Modify By Sindy 2023/11/3 mark:¦¬¤å¤w¤£»Ý°õ¦æ¦¹¬qµ{¦¡,¦]±µ¬¢³æ¤w·|¦^¼g¶l»¼°Ï¸¹
'         mStrSql = "update customer set cu30=" & CNULL(mCU30) & " where cu01=" & CNULL(Mid(m_Pa(26), 1, 8)) & " and cu02=" & CNULL(Mid(m_Pa(26), 9, 1))
'         cnnConnection.Execute mStrSql
         '2023/11/3 END
         
        'Add by Lydia 2014/10/31 ¶}©ñ¥~±Mµ{§Ç¤H­û¥i¶i¤J±M§Q³B¨t²Î¾Þ§@FMP¾ÈµØ®×¥ó=>¼g¤J¥N²z¤H
        If InStr(mChkVal, "¾ÈµØ®×¥ó½T»{") > 0 Then
            mStrSql = "update caseprogress set cp44='Y53374000' where cp09='" & mCP(9) & "' "
            cnnConnection.Execute mStrSql
        End If
        'end. 'Add by Lydia 2014/10/31
        'Add by Morgan 2010/8/10
        '¦¬¤å¬ü±M¥¿¦¡¥Ó½Ð®×(­ì¼È®É¥Ó½Ð®×¸¹-1)®É,¨R¼È®É¥Ó½Ð®×ªº¨ä¥L´Á­­
        If m_Pa(1) = "CFP" And m_Pa(3) = "1" And m_Pa(9) = "101" And mCP(10) = "101" Then
           strTmp1(0) = "select np01,np08,np09,np22 from nextprogress where np02='" & m_Pa(1) & "' and np03='" & m_Pa(2) & "' and np04='0' and np05='" & m_Pa(4) & "' and np06 is null and np07='910' "
           intJ = 1
           Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
           If intJ = 1 Then
              mStrSql = "update caseprogress set cp06=" & rsQD("NP08") & ",cp07=" & rsQD("NP09") & " where cp09='" & mCP(9) & "'"
              cnnConnection.Execute mStrSql
              'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(mCP(9)) & "
              mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np01='" & rsQD("NP01") & "' and np22=" & rsQD("NP22")
              cnnConnection.Execute mStrSql
           End If
        End If
        'Added by Lydia 2020/11/19 CFP­^°ê²æ¼Ú®×ºÞ¨î
        If m_Pa(1) = "CFP" And mCP(31) = "Y" And pType = "CFP­^°ê²æ¼Ú®×" And m_CaseNaTmp(1) <> "" And m_CaseNaTmp(2) <> "" Then
             strTmp1(0) = "select cp09,cp30 from caseprogress where cp01='" & m_CaseNaTmp(1) & "' and cp02='" & m_CaseNaTmp(2) & "' and cp03='" & m_CaseNaTmp(3) & "' and cp04='" & m_CaseNaTmp(4) & "' " & _
                              "and substr(cp09,1,1) ='C' and cp10='1608' and cp159=0 order by cp05 desc "
             intJ = 1
             Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
             If intJ = 1 Then
                'A. ¼Ú·ù®×­Y¦³¡u³qª¾­^°ê¦Aµù¥U¡vªºCÃþ¨Ó¨ç1608¤§CP30¦s¦Ü·s­^°ê®×¤§±M§Q¸¹¼ÆPA22
                If "" & rsQD.Fields("CP30") <> "" Then
                    mStrSql = "update patent set pa22='" & ChgSQL(rsQD.Fields("cp30")) & "' where pa01='" & m_Pa(1) & "' and pa02='" & m_Pa(2) & "' and pa03='" & m_Pa(3) & "' and pa04='" & m_Pa(4) & "' "
                    cnnConnection.Execute mStrSql
                End If
                'B. ¼Ú·ù®×­Y¦³¡u³qª¾­^°ê¦Aµù¥U¡vªºCÃþ¨Ó¨ç¤]Âà¦Ü·s­^°ê®×¸¹
                If "" & rsQD.Fields("cp09") <> "" Then
                     mStrSql = "update caseprogress set cp01='" & m_Pa(1) & "', cp02='" & m_Pa(2) & "', cp03='" & m_Pa(3) & "', cp04='" & m_Pa(4) & "' where cp09='" & rsQD.Fields("cp09") & "' "
                     cnnConnection.Execute mStrSql
                End If
             End If
             'Added by Lydia 2020/12/01
             If mCP(10) = "444" Then
                  '©e¥ô¥N²z¤H¤WÄò¿ì¡F¤U¤@µ{§Ç³Æµù¥[µù¡u­^°ê®×®×¸¹¡v
                  mStrSql = "update nextprogress set np06='Y', np24='" & mCP(9) & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "­^°ê®×®×¸¹¡G" & m_Pa(1) & m_Pa(2) & m_Pa(3) & m_Pa(4) & ";'||np15 " & _
                               "where np02='" & m_CaseNaTmp(1) & "' and np03='" & m_CaseNaTmp(2) & "' and np04='" & m_CaseNaTmp(3) & "' and np05='" & m_CaseNaTmp(4) & "' and np07='444' and np06 is null "
                  cnnConnection.Execute mStrSql
                  'E. ­Y¦¬¤å¡u©e¥ô¥N²z¤H(CFP.444)¡v®É¼Ú·ù®×¤U¤@µ{§Ç¤§/¡u©µ®i¶O(­^°ê)613¡v´Á­­Âà¦Ü·s®×¸¹¨Ã§ï®×¥ó©Ê½è¬°¡u©µ®i¶O607¡v¡F¤U¤@µ{§Ç³Æµù¥[µù¡u¼Ú·ù®×®×¸¹¡v
                  'Modified by Lydia 2020/12/16 ±NNP01§ï¬°­^°ê®×¦¬¤å¸¹; §_«h¤À®×§@·~·|¿ù»~(¥»©Ò®×¸¹¤£¦P)
                  mStrSql = "update nextprogress set np01='" & mCP(9) & "', np02='" & m_Pa(1) & "', np03='" & m_Pa(2) & "', np04='" & m_Pa(3) & "', np05='" & m_Pa(4) & "', np07='607', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "¼Ú·ù®×®×¸¹¡G" & m_CaseNaTmp(1) & m_CaseNaTmp(2) & m_CaseNaTmp(3) & m_CaseNaTmp(4) & ";'||np15 " & _
                              "where np02='" & m_CaseNaTmp(1) & "' and np03='" & m_CaseNaTmp(2) & "' and np04='" & m_CaseNaTmp(3) & "' and np05='" & m_CaseNaTmp(4) & "' and np07='613' and np06 is null "
                  cnnConnection.Execute mStrSql
             Else
             'end 2020/12/01
                  'C. ¼Ú·ù®×¤U¤@µ{§Ç¤§¡u©µ®i(­^°ê)¡v(CFP.613)´Á­­¤WÄò¿ìNP06¡A¤U¤@³æ¾Ú½s¸¹NP24°O¿ý·s­^°ê®×¤§Á`¦¬¤å¸¹¡F¤U¤@µ{§Ç³Æµù¥[µù¡u­^°ê®×®×¸¹¡v
                  mStrSql = "update nextprogress set np06='Y', np24='" & mCP(9) & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "­^°ê®×®×¸¹¡G" & m_Pa(1) & m_Pa(2) & m_Pa(3) & m_Pa(4) & ";'||np15 where np02='" & m_CaseNaTmp(1) & "' and np03='" & m_CaseNaTmp(2) & "' and np04='" & m_CaseNaTmp(3) & "' and np05='" & m_CaseNaTmp(4) & "' and np07='613' and np06 is null "
                  cnnConnection.Execute mStrSql
                  'Added by Lydia 2020/12/04 ±N¼Ú·ù®×¡u©e¥ô¥N²z¤H¡v´Á­­Âà¦Ü·s­^°ê®×¡F¤U¤@µ{§Ç³Æµù¥[µù¡u¼Ú·ù®×®×¸¹¡v
                  'Modified by Lydia 2020/12/16 ±NNP01§ï¬°­^°ê®×¦¬¤å¸¹; §_«h¤À®×§@·~·|¿ù»~(¥»©Ò®×¸¹¤£¦P)
                  mStrSql = "update nextprogress set np01='" & mCP(9) & "', np02='" & m_Pa(1) & "', np03='" & m_Pa(2) & "', np04='" & m_Pa(3) & "', np05='" & m_Pa(4) & "', np15='" & ChangeTStringToTDateString(strSrvDate(2)) & "¼Ú·ù®×®×¸¹¡G" & m_CaseNaTmp(1) & m_CaseNaTmp(2) & m_CaseNaTmp(3) & m_CaseNaTmp(4) & ";'||np15 " & _
                              "where np02='" & m_CaseNaTmp(1) & "' and np03='" & m_CaseNaTmp(2) & "' and np04='" & m_CaseNaTmp(3) & "' and np05='" & m_CaseNaTmp(4) & "' and np07='444' and np06 is null "
                  cnnConnection.Execute mStrSql
                  'end 2020/12/04
             End If 'Added by Lydia 2020/12/01
             
             'D. «Ø¥ß¼Ú·ù®×¤Î­^°ê®×¤§ÃöÁp(¬ÛÃö¨÷¸¹¡B°ê¤º¥~®×¡B¦h°ê®×)
             '----¬ÛÃö¨÷¸¹
             mStrSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_Pa(1)) & ", " & CNULL(m_Pa(2)) & ", " & CNULL(m_Pa(3)) & ", " & CNULL(m_Pa(4)) & ", " & CNULL(m_CaseNaTmp(1)) & ", " & CNULL(m_CaseNaTmp(2)) & ", " & CNULL(m_CaseNaTmp(3)) & ", " & CNULL(m_CaseNaTmp(4)) & " ) "
             cnnConnection.Execute mStrSql
             mStrSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_CaseNaTmp(1)) & ", " & CNULL(m_CaseNaTmp(2)) & ", " & CNULL(m_CaseNaTmp(3)) & ", " & CNULL(m_CaseNaTmp(4)) & ", " & CNULL(m_Pa(1)) & ", " & CNULL(m_Pa(2)) & ", " & CNULL(m_Pa(3)) & ", " & CNULL(m_Pa(4)) & " ) "
             cnnConnection.Execute mStrSql
             '----°ê¤º¥~®×
             mStrSql = "insert into CaseMap(CM01,CM02,CM03,CM04,CM05,CM06,CM07,CM08,CM09,CM10,CM11) select '" & m_Pa(1) & "', '" & m_Pa(2) & "', '" & m_Pa(3) & "','" & m_Pa(4) & "',CM05,CM06,CM07,CM08,CM09,CM10,CM11 " & _
                          "from CaseMap where CM01='" & m_CaseNaTmp(1) & "' and CM02='" & m_CaseNaTmp(2) & "' and CM03='" & m_CaseNaTmp(3) & "' and CM04='" & m_CaseNaTmp(4) & "' and cm10 in ('0','3','4','5','6') "
             cnnConnection.Execute mStrSql
             mStrSql = "insert into CaseMap(CM01,CM02,CM03,CM04,CM05,CM06,CM07,CM08,CM09,CM10,CM11) select CM01,CM02,CM03,CM04,'" & m_Pa(1) & "', '" & m_Pa(2) & "', '" & m_Pa(3) & "','" & m_Pa(4) & "',CM09,CM10,CM11 " & _
                          "from CaseMap where CM05='" & m_CaseNaTmp(1) & "' and CM06='" & m_CaseNaTmp(2) & "' and CM07='" & m_CaseNaTmp(3) & "' and CM08='" & m_CaseNaTmp(4) & "' and cm10 in ('0','3','4','5','6') "
             cnnConnection.Execute mStrSql
             '----¦h°ê®×¸¹(¥ý¼W¥[¼Ú·ù®×­ì¥»ÃöÁp¡A¦A¼W¥[¼Ú·ù®×©M­^°ê®×¤§ÃöÁp)
             mStrSql = "insert into caserelation(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) select '" & m_Pa(1) & "', '" & m_Pa(2) & "', '" & m_Pa(3) & "','" & m_Pa(4) & "',cr05,cr06,cr07,cr08 " & _
                          "from caserelation where cr01='" & m_CaseNaTmp(1) & "' and cr02='" & m_CaseNaTmp(2) & "' and cr03='" & m_CaseNaTmp(3) & "' and cr04='" & m_CaseNaTmp(4) & "' "
             cnnConnection.Execute mStrSql
             mStrSql = "insert into caserelation(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) select cr01,cr02,cr03,cr04,'" & m_Pa(1) & "', '" & m_Pa(2) & "', '" & m_Pa(3) & "','" & m_Pa(4) & "' " & _
                          "from caserelation where cr05='" & m_CaseNaTmp(1) & "' and cr06='" & m_CaseNaTmp(2) & "' and cr07='" & m_CaseNaTmp(3) & "' and cr08='" & m_CaseNaTmp(4) & "' "
             cnnConnection.Execute mStrSql
             mStrSql = "insert into caserelation(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_Pa(1)) & ", " & CNULL(m_Pa(2)) & ", " & CNULL(m_Pa(3)) & ", " & CNULL(m_Pa(4)) & ", " & CNULL(m_CaseNaTmp(1)) & ", " & CNULL(m_CaseNaTmp(2)) & ", " & CNULL(m_CaseNaTmp(3)) & ", " & CNULL(m_CaseNaTmp(4)) & " ) "
             cnnConnection.Execute mStrSql
             mStrSql = "insert into caserelation(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_CaseNaTmp(1)) & ", " & CNULL(m_CaseNaTmp(2)) & ", " & CNULL(m_CaseNaTmp(3)) & ", " & CNULL(m_CaseNaTmp(4)) & ", " & CNULL(m_Pa(1)) & ", " & CNULL(m_Pa(2)) & ", " & CNULL(m_Pa(3)) & ", " & CNULL(m_Pa(4)) & " ) "
             cnnConnection.Execute mStrSql
             'Added by Lydia 2020/12/04 ¼Ú·ù®×®×¥ó³Æµù¥[µù¡u­^°ê®×®×¸¹¡v¡F·s­^°ê®×¤§·s®×¦¬¤åªº¶i«×³Æµù¥[µù¡u¼Ú·ù®×®×¸¹¡v
             mStrSql = "Update Patent set PA91=" & CNULL("­^°ê®×®×¸¹¡G" & m_Pa(1) & m_Pa(2) & m_Pa(3) & m_Pa(4) & ";") & "||PA91 where PA01='" & m_CaseNaTmp(1) & "' and PA02='" & m_CaseNaTmp(2) & "' and PA03='" & m_CaseNaTmp(3) & "' and PA04='" & m_CaseNaTmp(4) & "' "
             cnnConnection.Execute mStrSql
             mStrSql = "Update CaseProgress set CP64=" & CNULL("¼Ú·ù®×®×¸¹¡G" & m_CaseNaTmp(1) & m_CaseNaTmp(2) & m_CaseNaTmp(3) & m_CaseNaTmp(4) & ";") & "||CP64 where CP09='" & mCP(9) & "' "
             cnnConnection.Execute mStrSql
             'end 2020/12/04
             'Added by Lydia 2021/01/11 ½Æ»sÀu¥ýÅv¸ê®Æ
             mStrSql = "insert into pridate (pd01,pd02,pd03,pd04,pd05,pd06,pd07,pd08,pd09,pd10) " & _
                          "select " & CNULL(m_Pa(1)) & ", " & CNULL(m_Pa(2)) & ", " & CNULL(m_Pa(3)) & ", " & CNULL(m_Pa(4)) & ", pd05,pd06,pd07,pd08,pd09,pd10 from pridate where pd01='" & m_CaseNaTmp(1) & "' and pd02='" & m_CaseNaTmp(2) & "' and pd03='" & m_CaseNaTmp(3) & "' and pd04='" & m_CaseNaTmp(4) & "' "
             cnnConnection.Execute mStrSql
             'end 2021/01/11
             'Added by Lydia 2021/04/15 CFP­^°ê²æ¼Ú©e¥ô¥N²z¤§«áÄò³B²z¡G¦¬¤å­^°ê©µ®i¤Î©e¥ô¥N²z¤H·s®×¡A¦P®É±N¥N²z¤H¦s¤JCP44¡C
             strTmp1(0) = "select np01,np15 from nextprogress where np07='444' and np15 like '%²æ¼Ú­^°ê®×¥N²z¤H¡G%' " & _
                              "and ((np02='" & m_Pa(1) & "' and np03='" & m_Pa(2) & "' and np04='" & m_Pa(3) & "' and np05='" & m_Pa(4) & "') or (np02='" & m_CaseNaTmp(1) & "' and np03='" & m_CaseNaTmp(2) & "'  and np04='" & m_CaseNaTmp(3) & "'  and np05='" & m_CaseNaTmp(4) & "')) "
             intJ = 1
             Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
             If intJ = 1 Then
                 strTmp1(1) = Mid("" & rsQD.Fields("np15"), InStr(rsQD.Fields("np15"), "²æ¼Ú­^°ê®×¥N²z¤H¡G") + 9, 9)
                 If Left(strTmp1(1), 1) = "Y" Then
                     mStrSql = "Update CaseProgress set cp44=" & CNULL(strTmp1(1)) & " where cp09=" & CNULL(mCP(9))
                     cnnConnection.Execute mStrSql
                 End If
             End If
             'end 2021/04/15
             PUB_EUtoUK m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), m_CaseNaTmp(1), m_CaseNaTmp(2), m_CaseNaTmp(3), m_CaseNaTmp(4), mCP(9), mCP(10)  'Added by Morgan 2020/12/21 ¦^ÂÐ³æÂk¨÷
        End If
        'end 2020/11/19
        
        'Added by Lydia 2021/04/15 CFP­^°ê²æ¼Ú©e¥ô¥N²z¤§«áÄò³B²z¡G¦¬¤å­^°ê©µ®i¤Î©e¥ô¥N²z¤H·s®×¡A¦P®É±N¥N²z¤H¦s¤JCP44¡C
                                                 '¦P¤@¤Ñ±µ¬¢³æ¤§«á¦¬¤åªº³B²z
        If m_Pa(1) = "CFP" And mCP(31) <> "Y" And m_Pa(4) = "201" And (mCP(10) = "444" Or mCP(10) = "607") And m_Pa(91) <> "" And InStr(m_Pa(91), "¼Ú·ù®×®×¸¹¡G") > 0 Then
             strTmp1(0) = "select cp44 from caseprogress where cp01='" & m_Pa(1) & "' and cp02='" & m_Pa(2) & "' and cp03='" & m_Pa(3) & "' and cp04='" & m_Pa(4) & "' and cp05=" & mCP(5) & " and cp10 in ('607','444') and cp159=0 And cp31='Y' "
             intJ = 1
             Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
             If intJ = 1 Then
                 If "" & rsQD.Fields("cp44") <> "" Then
                     mStrSql = "Update CaseProgress set cp44=" & CNULL(rsQD.Fields("cp44")) & " where cp09=" & CNULL(mCP(9))
                     cnnConnection.Execute mStrSql
                 End If
             End If
        End If
        'end 2021/04/15
        
        'Modify by Morgan 2006/5/4
        'FCPªº¸É¤å¥ó202¤£­n°µ
        If Not (m_Pa(1) = "FCP" And mCP(10) = "202") Then
           mStrSql = "select np01 from nextprogress where np02 = '" & m_Pa(1) & "' and np03 = '" & m_Pa(2) & "' and np04 = '" & m_Pa(3) & "' and np05 = '" & m_Pa(4) & "' and np06 is null and np07 = '" & mCP(10) & "'"
           'Add by Morgan 2007/1/12 ¥xÆW±M§Qªº¥Ó´_©Î­×¥¿®É¤U¤@µ{§Ç¨â­Ó³£­n§ì
           If m_Pa(9) = "000" And (mCP(10) = "205" Or mCP(10) = "204") Then
              mStrSql = "select np01 from nextprogress where np02 = '" & m_Pa(1) & "' and np03 = '" & m_Pa(2) & "' and np04 = '" & m_Pa(3) & "' and np05 = '" & m_Pa(4) & "' and np06 is null and np07 IN ('204','205')"
           End If
           'end 2007/1/12
           adoquery.CursorLocation = adUseClient
           adoquery.Open mStrSql, cnnConnection, adOpenStatic, adLockReadOnly
           If adoquery.RecordCount > 0 Then
              If adoquery.RecordCount = 1 Then
                 mCP(43) = adoquery.Fields(0) 'Added by Morgan 2012/8/9
                 If IsNull(adoquery.Fields(0).Value) = False Then
                    'Add by Morgan 2010/6/30 ²§Ä³µªÅG¡BÁ|µoµªÅG­n¤@¨Ã§ó·s¹ï³y¸ê®Æ
                    If (mCP(10) = "802" Or mCP(10) = "804") Then
                       cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where CP09 = '" & mCP(9) & "'", intJ
                    Else
                    'End 2010/6/30
                       cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where CP09 = '" & mCP(9) & "'"
                    End If
                 End If
                 'add by nick 2004/09/08
                 If mCP(10) <> "411" Then
                    'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(mCP(9)) & "
                    mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np02=" & CNULL(m_Pa(1)) & " and np03=" & _
                            CNULL(m_Pa(2)) & " and np04=" & CNULL(m_Pa(3)) & " and np05=" & CNULL(m_Pa(4)) & _
                            " and np07=" & CNULL(mCP(10)) & " and np06 is null"
                         
                    'Add by Morgan 2007/1/12 ¥xÆW±M§Qªº¥Ó´_©Î­×¥¿®É¤U¤@µ{§Ç¨â­Ó³£­n§ì
                    If m_Pa(9) = "000" And (mCP(10) = "205" Or mCP(10) = "204") Then
                       'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(mCP(9)) & "
                       mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np02=" & CNULL(m_Pa(1)) & " and np03=" & CNULL(m_Pa(2)) & " and np04=" & CNULL(m_Pa(3)) & " and np05=" & CNULL(m_Pa(4)) & " and np07 in ('204','205') and np06 is null"
                    End If
                    'end 2007/1/12
                    cnnConnection.Execute mStrSql
                 End If
              End If
           Else
              adoquery.Close
              mStrSql = "select np01 from nextprogress where np02 = '" & m_Pa(1) & "' and np03 = '" & m_Pa(2) & "' and np04 = '" & m_Pa(3) & "' and np05 = '" & m_Pa(4) & "' and np06 <>'Y' and np07 = '" & mCP(10) & "'"
              'Add by Morgan 2007/1/12 ¥xÆW±M§Qªº¥Ó´_©Î­×¥¿®É¤U¤@µ{§Ç¨â­Ó³£­n§ì
              If m_Pa(9) = "000" And (mCP(10) = "205" Or mCP(10) = "204") Then
                 mStrSql = "select np01 from nextprogress where np02 = '" & m_Pa(1) & "' and np03 = '" & m_Pa(2) & "' and np04 = '" & m_Pa(3) & "' and np05 = '" & m_Pa(4) & "' and np06<>'Y' and np07 IN ('204','205')"
              End If
              'end 2007/1/12
              adoquery.CursorLocation = adUseClient
              adoquery.Open mStrSql, cnnConnection, adOpenStatic, adLockReadOnly
              If adoquery.RecordCount > 0 Then
                 If adoquery.RecordCount = 1 Then
                    mCP(43) = adoquery.Fields(0) 'Added by Morgan 2012/8/9
                    If IsNull(adoquery.Fields(0).Value) = False Then
                       'Add by Morgan 2010/6/30 ²§Ä³µªÅG¡BÁ|µoµªÅG­n¤@¨Ã§ó·s¹ï³y¸ê®Æ
                       If (mCP(10) = "802" Or mCP(10) = "804") Then
                          cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where CP09 = '" & mCP(9) & "'", intJ
                       Else
                       'End 2010/6/30
                          cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where CP09 = '" & mCP(9) & "'"
                       End If
                    End If
                    'add by nick 2004/09/08
                    If mCP(10) <> "411" Then
                         'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(mCP(9)) & "
                         mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np02=" & CNULL(m_Pa(1)) & " and np03=" & _
                            CNULL(m_Pa(2)) & " and np04=" & CNULL(m_Pa(3)) & " and np05=" & CNULL(m_Pa(4)) & _
                            " and np07=" & CNULL(mCP(10)) & " and np06 <> 'Y'"
                         'Add by Morgan 2007/1/12 ¥xÆW±M§Qªº¥Ó´_©Î­×¥¿®É¤U¤@µ{§Ç¨â­Ó³£­n§ì
                          If m_Pa(9) = "000" And (mCP(10) = "205" Or mCP(10) = "204") Then
                             'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(mCP(9)) & "
                             mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np02=" & CNULL(m_Pa(1)) & " and np03=" & CNULL(m_Pa(2)) & " and np04=" & CNULL(m_Pa(3)) & " and np05=" & CNULL(m_Pa(4)) & " and np07 in ('204','205') and np06<>'Y'"
                          End If
                          'end 2007/1/12
                         cnnConnection.Execute mStrSql
                    End If
                 End If
              End If
           End If
           adoquery.Close
        End If
        '2006/5/4 end
         
         'Add By Sindy 2024/7/12 ÀË¬d¬O§_¦³¨R¤U¤@µ{§Çªº´Á­­
         '                       ­Y¦³,¦¹µ§¦¬¤å´Á­­§ó·s¬°¤U¤@µ{§Çªº´Á­­
         If UCase(pFormName) = UCase("frm090801_New") Then
            strTmp1(0) = "select np01,np08,np09 from nextprogress where np24 = '" & mCP(9) & "'"
            intJ = 1
            Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
            If intJ = 1 Then
               'Modify By Sindy 2024/10/15 +iif§PÂ_,´Á­­¤é´Á¦³¥i¯à¬OªÅ¥Õ
               mStrSql = "update caseprogress set CP06=" & "" & IIf(rsQD.Fields("np08") > 0, rsQD.Fields("np08"), "''") & _
                                                ",CP07=" & "" & IIf(rsQD.Fields("np09") > 0, rsQD.Fields("np09"), "''")
               mStrSql = mStrSql & " where cp09=" & CNULL(mCP(9))
               cnnConnection.Execute mStrSql
            End If
         End If
         '2024/7/12 END
         
         'Add By Sindy 2024/1/4 ¹q¤l¦¬¤å¤£¶·°õ¦æ¦¹¨ç¼Æ
         If UCase(pFormName) <> UCase("frm090801_New") Then
         '2024/1/4 END
            If Cls001SetCaseProgressFee(m_Pa(1), m_Pa(9), mCP(10), mCP(9)) = False Then bolError = True
         End If
      Else
         bolError = True
      End If
   End If
   'add by nickc 2008/05/02 Àx¦s¹w©w¦¬´Ú¤é
   If bolError = False Then
       'Remove by Lydia 2018/08/22 (À³¦¬±b´ÚºÞ±±)¨ú®ø¹w©w¦¬´Ú¤é,§ï¦¨¥I´Ú¶g´Á
'       Dim rtCnt As Integer
'       'Modify by Morgan 2010/12/9
'       'If txtPatent(28) <> "" Then
'       '    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0) + 1,'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " ", rtCnt
'       If txtPatent(28) <> "" And txtPatent(28) <> txtPatent(28).Tag Then
'           cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0) + 1,'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
'       'end 2010/12/9
'           If rtCnt = 0 Then
'               cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " from dual "
'           End If
'       End If
       'end 2018/08/22
       
      'Added by Morgan 2012/8/9
      If m_Pa(1) = "FCP" Then
         'Modified by Morgan 2012/9/13 +603
         'Modified by Morgan 2012/10/24 +935
         'Modified by Morgan 2012/12/19 +125
         'Modified by Morgan 2024/11/14 +447 ¦A¼f¬d¥[³t¼f¬d¤]¹w³]µ{§Ç©Ó¿ì --±Ó²ú
         If InStr("101,102,103,105,125,401,404,416,447,601,603,605,701,702,908,929,935", mCP(10)) > 0 Then
            'Added by Morgan 2012/8/24
            '¤À³Î®×ªº¹ê¼f¤£­n¹w³]
            If mCP(10) = "416" Then
               If PUB_ChkCPExist(m_Pa, "307") Then
                  bolNoAutoCP14 = True
               End If
            End If
            'end 2012/8/24
            If bolNoAutoCP14 = False Then
               strTmp1(1) = PUB_GetFCPHandler(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), mCP(10))
               If strTmp1(1) <> "" Then
                  mStrSql = "update caseprogress set cp14='" & strTmp1(1) & "' where cp09='" & mCP(9) & "'"
                  cnnConnection.Execute mStrSql
               End If
            End If
         ElseIf mCP(43) > "C" Then
            'Modified by Lydia 2023/12/19 §ï§ì³Ì·s¤@¹D¤uµ{®v©Ó¿ì¤H­û; °Ñ¦Ò¡¨©Ó¿ì¤uµ{®v¤wÂ÷Â¾¡A©Ó¿ì¤H¹w³]¬°¸Ó®×¥ó²Õ§O¤§¥DºÞ¡¨
            'mStrSql = "update caseprogress a set cp14=(select nvl(max(b.cp14),a.cp14) from caseprogress b,staff where b.cp09=a.cp43 and st01(+)=cp14 and st03 in ('F21','F81') and st04='1') where cp09='" & mCP(9) & "'"
            'Modified by Lydia 2023/12/21 debug: ±Æ°£¾Þ§@¤H­û; ex.FCP-69492©Ó¿ì¤uµ{®v¤wÂ÷Â¾
            'mStrSql = "update caseprogress a set cp14=" & CNULL(PUB_GetFCPPromoterNo(mCP(43), mCP(10))) & " where cp09='" & mCP(9) & "'"
            'cnnConnection.Execute mStrSql
            strTmp1(0) = "SELECT CP10,CP14,ST01,ST04 FROM CASEPROGRESS,STAFF WHERE CP09='" & mCP(43) & "' AND CP14=ST01(+) AND ST03 IN ('F21','F81') "
            intJ = 1
            Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
            If intJ = 1 Then
               If "" & rsQD.Fields("st04") <> "1" Then
                  mStrSql = PUB_GetFCPPromoterNo(mCP(43), "" & rsQD.Fields("cp10"))
               Else
                  mStrSql = "" & rsQD.Fields("cp14")
               End If
               mStrSql = PUB_SetEng(mStrSql) 'Added by Lydia 2024/02/29 ¥~±M¾÷±ñ³]­p²Õ¤H­û²§°Ê½Õ¾ãµ{¦¡
               If mStrSql <> strUserNum And mStrSql <> "" And GetStaffName(mStrSql) <> "" Then
                  mStrSql = "update caseprogress a set cp14=" & CNULL(mStrSql) & " where cp09='" & mCP(9) & "'"
                  cnnConnection.Execute mStrSql
               End If
            End If
            'end 2023/12/21
         End If
         
         'Added by Morgan 2024/12/26
         'FCP»âÃÒ©Î¦~¶O©ó´_Åv´Á­­¦¬¤å®É­YµL¥¼µo¤å¤§¥Ó½Ð´_Åv®É¦Û°Ê¤º³¡¦¬¤å¨Ã¼u´£¿ô--Lisa
         If m_Pa(1) = "FCP" And (mCP(10) = "601" Or mCP(10) = "605") Then
            'Modified by Morgan 2024/12/30 ¦¬¤å¥i¯à«D­ì©l´Á­­(¤wºÞ¨î¥b¦~«á)¡A»âÃÒ601§ï§ì®Ö­ã´Á­­ Ex:FCP-061875
            'strChkDate = mCP(7)
            If mCP(10) = "601" Then
               strTmp1(0) = "SELECT CP07 FROM NEXTPROGRESS,caseprogress WHERE np02='" & m_Pa(1) & "' and np03='" & m_Pa(2) & "' and np04='" & m_Pa(3) & "' and np05='" & m_Pa(4) & "' and np07='601' and cp09(+)=np01 and cp10='1001' and cp07>0"
               intJ = 1
               Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
               If intJ = 1 Then
                  strChkDate = rsQD(0)
               End If
            Else
               strChkDate = PUB_GetNextFeeDate(m_Pa())
            End If
            
            If strChkDate = "" Then
               strAlert = "µLªk¨ú±o­ì©lªk©w´Á­­¡A½Ð¦Û¦æ§PÂ_¬O§_»Ý¦¬¤å¡u414¥Ó½Ð´_Åv¡v¡I"
            Else
            'end 2024/12/30
            
               If Val(strSrvDate(1)) > Val(strChkDate) Then
                  '»âÃÒ¹O­ìªk­­¦ý¥¼¶W¹L6­Ó¤ë
                  If mCP(10) = "601" Then
                     strFCP414BCP07 = CompDate(1, 6, strChkDate)
                     If Val(strSrvDate(1)) <= Val(strFCP414BCP07) Then
                        bolAutoFCP414 = True
                     End If
                  
                  '¦~¶O¹O­ìªk­­6­Ó¤ë¦ý¥¼¶W18­Ó¤ë(6­Ó¤ë+1¦~)
                  ElseIf mCP(10) = "605" Then
                     If Val(strSrvDate(1)) > Val(CompDate(1, 6, strChkDate)) Then
                        strFCP414BCP07 = CompDate(1, 18, strChkDate)
                        If Val(strSrvDate(1)) <= Val(strFCP414BCP07) Then
                           bolAutoFCP414 = True
                        End If
                     End If
                  End If
               End If
            End If
            If bolAutoFCP414 Then
               'ÀË¬d¬O§_¦³414¥Ó½Ð´_Åv¥¼µo¤å
               If PUB_ChkCPExist(m_Pa, "414", 1) = True Then
                  bolAutoFCP414 = False
               Else
                  strFCP414BCP06 = PUB_GetFCPOurDeadline(strFCP414BCP07, 4)
                  strFCP414BCP48 = Pub_GetHandleDay(m_Pa(1), m_Pa(9), "414", , strFCP414BCP07, , , m_Pa(1) & "-" & m_Pa(2) & "-" & m_Pa(3) & "-" & m_Pa(4))
                  strFCP414BCP17 = GetPatentOfficialFee(m_Pa(1), "414", "", m_Pa(8), m_Pa(9), m_Pa(16), m_Pa(14), m_Pa(2), m_Pa(3), m_Pa(4), mCP(118))
                  strFCP414BCP16 = Val(GetFCPFee(m_Pa(1), "414")) + Val(strFCP414BCP17)
                  ' ¶O¥Î
                  If strFCP414BCP16 > 0 Then
                     'ÂI¼Æ
                     strFCP414BCP18 = Format((Val(strFCP414BCP16) - Val(strFCP414BCP17)) / 1000, "0.0")
                  End If
      
                  strFCP414BCP09 = AutoNo("B", 6)
                  mStrSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP11,CP12,CP13,CP14,CP16,CP17,CP18,CP43,CP48,CP118)" & _
                     " select cp01,cp02,cp03,cp04,cp05," & strFCP414BCP06 & "," & strFCP414BCP07 & _
                     ",'" & strFCP414BCP09 & "','414','90',cp12,cp13,cp14," & Val(strFCP414BCP16) & "," & Val(strFCP414BCP17) & _
                     "," & Val(strFCP414BCP18) & ",CP09," & CNULL(mCP(48), True) & ",cp118 from caseprogress where cp09='" & mCP(9) & "'"
                  cnnConnection.Execute mStrSql, intW
               End If
            End If
         End If
         'end 2024/12/26
         
      End If
   End If
   
   
   'Added by Lydia 2015/12/31 ¦¬¤åFCP®×»âÃÒ601¥B¤­­Ó¥Ó½Ð¤H¦³¤@­Ó¬°X47794(¤T¬PÆp¥Û)®É¡AÀË¬d¸Ó®×¸¹ªº¦æ¨Æ¾ä´Á­­¡A­Y¨Æ¥Ñ¦³"¥i¦¬¤å»âÃÒ"®É¡A©ó¦¬¤å¦sÀÉ®É¦P®É¸Ñ°£¸Óµ§¦æ¨Æ¾ä´Á­­¡C
   If m_Pa(1) = "FCP" And mCP(10) = "601" And InStr(m_Pa(26) & "," & m_Pa(27) & "," & m_Pa(28) & "," & m_Pa(29) & "," & m_Pa(30), "X47794") > 0 Then
       strTmp1(0) = "select * from staff_calendar where sc05='" & m_Pa(1) & "' and sc06='" & m_Pa(2) & "' and sc07='" & m_Pa(3) & "' and sc08='" & m_Pa(4) & "' " & _
                   "and sc04 like '%FCP%(¤T¬PÆp¥Û)%¥i¦¬¤å»âÃÒ%' and sc18 is null "
       intJ = 1
       Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
       If intJ = 1 Then
          If rsQD.Fields("SC10") > 1 Then
             If PUB_AddFCPStaffCalendar(rsQD.Fields("SC01"), rsQD.Fields("SC10"), rsQD.Fields("SC03"), rsQD.Fields("SC04"), rsQD.Fields("SC09"), rsQD.Fields("SC10"), m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), , , rsQD.Fields("SC11")) Then
             End If
          End If
          mStrSql = "UPDATE staff_calendar SET sc17='" & strUserNum & "',sc18=" & strSrvDate(1) & ",sc19=" & CNULL(Mid(Right("000000" & ServerTime, 6), 1, 4), True) & _
                   " where sc01=" & rsQD.Fields("SC01") & " and sc02=" & rsQD.Fields("SC02")
          cnnConnection.Execute mStrSql
       End If
   End If
   'end 2015/12/31
   
   'Added by Lydia 2020/10/14 Murgitroyd§e°e´Á­­³]©w: ¤¤¶¡µ{§Ç³ø§i(¥Ó´_205¡B¦A¼f107): ¦¬¨ì¥N²z¤H«ü¥Ü7¤é¤º§¹¦¨¨Ã½Ð´Ú³ø§i¡C
   'Modified by Lydia 2020/10/19 ¼W¥[§PÂ_¨t²Î§O©M¥N²z¤H½s¸¹; ex: P-109161ªº¥Ó´_AA9043040¶i«×³Æµù¦³»~
   'Modified by Lydia 2021/01/06 +«D·s®×¦¬¤å m_PA(2) <> ""
   If (m_Pa(1) = "P" Or m_Pa(1) = "FCP") And (mCP(10) = "205" Or mCP(10) = "107") And m_Pa(75) <> "" And m_Pa(2) <> "" Then
       strTmp1(0) = Pub_GetSpecMan("¥~±MMURGITROYD³]©w")
       If strTmp1(0) <> "" And InStr(strTmp1(0), ChangeCustomerL(m_Pa(75))) > 0 Then
           strTmp1(1) = CompWorkDay(1, CompDate(2, 7, mCP(5)))
           'Added by Lydia 2022/03/17 ­Y«ü©w°e¥óªº¤é¤j©ó¥»©Ò´Á­­¡A½Ð¥H¥»©Ò´Á­­¬°·Ç
           If mCP(6) <> "" Then
              If strTmp1(1) > TransDate(mCP(6), 2) Then
                  strTmp1(1) = TransDate(mCP(6), 2)
              End If
           End If
           'end 2022/03/17
           '¦Û°Ê±a¦Ü¦¬¤å¨º¹D¶i«×³Æµù¡G¬°Murgitroyd®×»Ýxx¤ëxx¤é¡]¦¬¤å¤é¡Ï7­Ó¤é¾ä¤Ñ¡A­Y¬°°²¤é«h§ì¤U¤@­Ó¤u§@¤Ñ¡^§¹¦¨°e¥ó¨Ã³ø§i
           strTmp1(4) = "¬°Murgitroyd®×»Ý" & ChangeWStringToTDateString(strTmp1(1)) & "§¹¦¨°e¥ó¨Ã³ø§i"
           strTmp1(3) = PUB_GetFCPHandler(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4))
           'Modified by Lydia 2021/05/20 ¦b¦æ¨Æ¾ä¨Æ¥Ñ¼W¥[[¸Ñ°£ºÞ¨î¤£³qª¾]¡A±Æ°£" ¸Ñ°£¤H­û«D«Ø¥ß¦æ¨Æ¾ä¤H­û·|µoemail³qª¾«Ø¥ß¤H­û"¦æ¨Æ¾ä¤w³Q¸Ñ°£ºÞ¨î"
           If PUB_AddFCPStaffCalendar(strTmp1(1), "1", strTmp1(3), strTmp1(4) & "[¸Ñ°£ºÞ¨î¤£³qª¾]", strTmp1(3), "1", m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4)) = True Then
               'Modified by Lydia 2020/10/26 ©Ó¿ì´Á­­¡G³]¦¬¤å¤é+7­Ó¤é¾ä¤Ñ¡A­Y¬°°²¤é«h§ì¤U¤@­Ó¤u§@¤Ñ¡C
               mStrSql = "Update CaseProgress set cp64=" & CNULL(strTmp1(4)) & "||';'||cp64, cp48=" & strTmp1(1) & " where cp09=" & CNULL(mCP(9))
               cnnConnection.Execute mStrSql
           End If
       End If
   End If
   'end 2020/10/14
   'Added by Lydia 2021/01/06 ¦P¤@¤Ñ¦¬¤å´£¥Ó«á¤§¹êÅé¼f¬d+¥D°Ê­×¥¿¤§®×¥ó¡A¤ñ·Ó¤¤¶¡µ{§Ç¤§¤º³¡±±ºÞ¡C
   If (m_Pa(1) = "P" Or m_Pa(1) = "FCP") And (mCP(10) = "416" Or mCP(10) = "203") And m_Pa(75) <> "" And m_Pa(2) <> "" Then
       strTmp1(0) = Pub_GetSpecMan("¥~±MMURGITROYD³]©w")
       If strTmp1(0) <> "" And InStr(strTmp1(0), ChangeCustomerL(m_Pa(75))) > 0 Then
           strTmp1(1) = "select pa10,cp09 from patent,caseprogress where pa01='" & m_Pa(1) & "' and pa02='" & m_Pa(2) & "' and pa03='" & m_Pa(3) & "' and pa04='" & m_Pa(4) & "' " & _
                            "and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and cp05(+)=" & CNULL(mCP(5)) & " and cp10(+)=" & CNULL(IIf(mCP(10) = "416", "203", "416")) & " and cp158(+)=0 and cp159(+)=0"
           intJ = 1
           Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(1))
           If intJ = 1 Then
               strTmp1(9) = "" & rsQD.Fields("cp09")
               If "" & rsQD.Fields("pa10") <> "" And strTmp1(9) <> "" Then
                    '©ó³Ì«á¦¬¤å¤§¹êÅé¼f¬d or¥D°Ê­×¥¿®É¡A¤~²£¥Í¦æ¨Æ¾ä¨Ã¥B¤@¨Ö§ó·s¶i«×³Æµù©M©Ó¿ì´Á­­¡C
                    strTmp1(1) = CompWorkDay(1, CompDate(2, 7, mCP(5)))
                    'Added by Lydia 2022/03/17 ­Y«ü©w°e¥óªº¤é¤j©ó¥»©Ò´Á­­¡A½Ð¥H¥»©Ò´Á­­¬°·Ç
                    If mCP(6) <> "" Then
                       If strTmp1(1) > TransDate(mCP(6), 2) Then
                           strTmp1(1) = TransDate(mCP(6), 2)
                       End If
                    End If
                    'end 2022/03/17
                    strTmp1(4) = "¬°Murgitroyd®×»Ý" & ChangeWStringToTDateString(strTmp1(1)) & "§¹¦¨°e¥ó¨Ã³ø§i"
                    strTmp1(3) = PUB_GetFCPHandler(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4))
                    'Modified by Lydia 2021/06/21 ¦b¦æ¨Æ¾ä¨Æ¥Ñ¼W¥[[¸Ñ°£ºÞ¨î¤£³qª¾]¡A±Æ°£" ¸Ñ°£¤H­û«D«Ø¥ß¦æ¨Æ¾ä¤H­û·|µoemail³qª¾«Ø¥ß¤H­û"¦æ¨Æ¾ä¤w³Q¸Ñ°£ºÞ¨î"
                    If PUB_AddFCPStaffCalendar(strTmp1(1), "1", strTmp1(3), strTmp1(4) & "[¸Ñ°£ºÞ¨î¤£³qª¾]", strTmp1(3), "1", m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4)) = True Then
                        mStrSql = "Update CaseProgress set cp64=" & CNULL(strTmp1(4)) & "||';'||cp64, cp48=" & strTmp1(1) & " where cp09=" & CNULL(mCP(9))
                        cnnConnection.Execute mStrSql
                        mStrSql = "Update CaseProgress set cp64=" & CNULL(strTmp1(4)) & "||';'||cp64, cp48=" & strTmp1(1) & " where cp09=" & CNULL(strTmp1(9))
                        cnnConnection.Execute mStrSql
                    End If
               End If
           End If
       End If
   End If
   'end 2021/01/06
   
    'Added by Lydia 2020/03/30 FCP®×©MFMP®×: ¦]¬°¤À³Î®×¤]¦³¤¤»¡ÀÉ, ©Ò¥H§ï¦b³Ì«e­±·s¼W¦¬¤åDÃþEnglish_Vers©M±M§Q
    If strSrvDate(1) >= XY¯S®íÅv­­±Ò¥Î¤ébyÀÉ®× And intSaveMode = 1 And intModifyKind = 0 _
        And (m_Pa(1) = "FCP" Or (m_Pa(1) = "P" And Left(mCP(12), 1) = "F")) Then
        If PUB_ChkCPExist(m_Pa, cntEnglish_Vers, , , , "D") = False Then
              strTmp1(0) = AutoNo("D", 6)
              strTmp1(6) = PUB_GetFCPSalesNo(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4))   'FCP©Ó¿ì
              strTmp1(5) = GetSalesArea(strTmp1(6))
              strTmp1(7) = PUB_GetFCPHandler(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4)) 'FCPµ{§Ç
              
              mStrSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
                 ",cp12,cp13,cp14,cp20,cp26,cp27,cp32) values ('" & m_Pa(1) & "','" & m_Pa(2) & "','" & m_Pa(3) & "','" & m_Pa(4) & "',19221111,'" & strTmp1(0) & "','" & cntEnglish_Vers & "' " & _
                  ",'" & strTmp1(5) & "','" & strTmp1(6) & "','" & strTmp1(7) & "','N','N',19221111,'N')"
              cnnConnection.Execute mStrSql
        End If
        If PUB_ChkCPExist(m_Pa, cnt±M§Q®×¥ó, , , , "D") = False Then
              strTmp1(0) = AutoNo("D", 6)
              If strTmp1(5) = "" Or strTmp1(6) = "" Or strTmp1(7) = "" Then
                strTmp1(6) = PUB_GetFCPSalesNo(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4))   'FCP©Ó¿ì
                strTmp1(5) = GetSalesArea(strTmp1(6))
                strTmp1(7) = PUB_GetFCPHandler(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4)) 'FCPµ{§Ç
              End If
              mStrSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
                 ",cp12,cp13,cp14,cp20,cp26,cp27,cp32) values ('" & m_Pa(1) & "','" & m_Pa(2) & "','" & m_Pa(3) & "','" & m_Pa(4) & "',19221111,'" & strTmp1(0) & "','" & cnt±M§Q®×¥ó & "' " & _
                  ",'" & strTmp1(5) & "','" & strTmp1(6) & "','" & strTmp1(7) & "','N','N',19221111,'N')"
              cnnConnection.Execute mStrSql
        End If
    End If
    'end 2020/03/30
   'Added by Lydia 2017/11/14 FCP®×¥ó©R¦W¹q¤l¤Æ¡G¤¤»¡¿é¤J¬ÛÃö³]©w-¦sÀÉ
   'Modified by Lydia 2019/06/11 §PÂ_¨«©R¦W¬yµ{¤~ÀË¬d;
   'Modified by Lydia 2019/07/04 ¤À³Î®×¤£¨«©R¦W¬yµ{¡A¦ý¬O­n¯à¤Ä¿ï¨ä¥L¦¬¤å
   If mTCTVal <> "" Then
      'Added by Lydia 2023/03/09 ¿é¤J°lÂÜ¬y¤ô¸¹,¤£¨«©R¦W§@·~
      If InStr(mTCTVal, "|") = 0 Then
          mStrSql = "Update TrackingCaseName set TCN05=" & CNULL(mCP(9)) & " Where TCN01 ='" & mTCTVal & "' "
          cnnConnection.Execute mStrSql
          'Added by Lydia 2023/05/22 ¤é¤å²Õ©Ó¿ì¼W¥[¡u«È¤á¦³´£¨Ñ±m¹Ï¡vÄæ¦ì¿é¤J
          strTmp1(1) = "select tcn12 from trackingcasename where tcn01='" & mTCTVal & "' "
          intJ = 1
          Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(1))
          If intJ = 1 Then
             If "" & rsQD.Fields("tcn12") = "Y" Then
                 mStrSql = "Update Patent Set PA63=" & CNULL(rsQD.Fields("tcn12")) & " WHERE PA01='" & m_Pa(1) & "' AND  PA02='" & m_Pa(2) & "' AND PA03='" & m_Pa(3) & "' AND PA04='" & m_Pa(4) & "' "
                 cnnConnection.Execute mStrSql
             End If
          End If
          'end 2023/05/22
      Else
      'end 2023/03/09
         tmpArr = Split(mTCTVal, "|")
         '0:¤¤»¡Ãþ«¬, 1: ¤Ä¿ïªº¦¬¤å©Ê½è, 2: ©R¦W°lÂÜ¬y¤ô¸¹Tracking No, 3:¤uµ{®v²Õ§O, 4.Ä¶²¦´Á­­
         If Trim(tmpArr(2)) <> "" Then
            If intSaveMode = 1 And intModifyKind = 0 And Trim(tmpArr(3)) <> "B" Then
               'Modified by Lydia 2018/03/09 ¥u­n¤W¤w¤À®×,Trigger·|§ó·s¤À®×¤é´Á
               mStrSql = "update caseprogress set cp122='Y'  where cp09='" & mCP(9) & "' "
               cnnConnection.Execute mStrSql
            End If
            Call PUB_UpdTCTrecord(Trim(tmpArr(0)), Trim(tmpArr(1)), Trim(tmpArr(2)), mTCTList, m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), m_Pa(5), m_Pa(6), mCP(9), mCP(10), _
                  mCP(6), mCP(7), mCP(13), m_Pa(8), m_Pa(9), m_Pa(16), m_Pa(14), m_Pa(26) & "," & m_Pa(27) & "," & m_Pa(28) & "," & m_Pa(29) & "," & m_Pa(30), _
                  m_Pa(75), m_Pa(150), Trim(tmpArr(3)), Trim(tmpArr(4)))
         End If
      End If 'Added by Lydia 2023/03/09
   'Added by Lydia 2018/06/28 «á¸É:«æ¥óÂ½Ä¶±N·s®×Â½Ä¶¦¬¤å¸¹¦^¼g¨ìÂ½Ä¶¶O¥ÎÀÉ©M©R¦W°O¿ýÀÉ(TCN14)
   ElseIf (m_Pa(1) = "P" Or m_Pa(1) = "FCP") And m_Pa(2) <> "" And mCP(10) = "201" And intModifyKind = 0 Then
       'Modified by Lydia 2020/01/06 ¤£­­¨î«æ¥óÂ½Ä¶
       strTmp1(0) = "select cp09,tcn01,tcn14 from caseprogress,trackingcasename where cp01='" & m_Pa(1) & "' and cp02='" & m_Pa(2) & "' and cp03='" & m_Pa(3) & "' and cp04='" & m_Pa(4) & "' " & _
                         "And cp31='Y' and cp09=tcn05(+) "
       intJ = 1
       Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
       If intJ = 1 Then
            If "" & rsQD.Fields("tcn01") <> "" Then '¦³©R¦W°lÂÜ
                'Added by Lydia 2020/01/06 «D«æ¥óÂ½Ä¶(¥¼´£¥Ó¥ýÂ½Ä¶¦Û°Ê¤Ä¿ï)
                If "" & rsQD.Fields("tcn14") = "" Then
                     mStrSql = "Insert into TransFee(TF01,TF31) values(" & CNULL(mCP(9)) & ", 'Y' )"
                     cnnConnection.Execute mStrSql
                Else  '«æ¥óÂ½Ä¶
                'end 2020/01/06
                    'Added by Lydia 2019/12/23 FMP®×¦¬¤å¡A·s®×«ØÀÉ¥¼´£¥Ó¥ýÂ½Ä¶¦Û°Ê¤Ä¿ï
                    If m_Pa(1) = "P" Then
                         mStrSql = "update TransFee set TF01=" & CNULL(mCP(9)) & ",TF31='Y' where TF01=" & CNULL("" & rsQD.Fields("tcn14"))
                    Else
                    'end 2019/12/23
                         mStrSql = "update TransFee set TF01=" & CNULL(mCP(9)) & " where TF01=" & CNULL("" & rsQD.Fields("tcn14"))
                    End If 'end 2019/12/23
                    cnnConnection.Execute mStrSql
                    mStrSql = "update TrackingCaseName set TCN14=" & CNULL(mCP(9)) & " where TCN01=" & CNULL("" & rsQD.Fields("tcn01"))
                    cnnConnection.Execute mStrSql
                End If 'end 2020/01/06
                'Added by Lydia 2020/08/24 FMP®×¹w³]µo"¥¼´£¥Ó¥ýÂ½Ä¶"email ; °Ñ¦Òfrm060102
                strTmp1(0) = Pub_GetSpecMan("M")
                If strTmp1(0) <> "" Then
                    'Modify By Sindy 2023/3/27 +,mc13
                    mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
                       " values( '" & strUserNum & "','" & strTmp1(0) & "',to_char(sysdate,'yyyymmdd')" & _
                       ",to_char(sysdate,'hh24miss'),'" & m_Pa(1) & m_Pa(2) & IIf(m_Pa(3) & m_Pa(4) <> "000", m_Pa(3) & m_Pa(4), "") & " ¥¼´£¥Ó¥ýÂ½Ä¶" & "','¦P¥D¦®',null,'" & mCP(9) & "')"
                    cnnConnection.Execute mStrSql
                End If
                'end 2020/08/24
            End If
       End If
   'end 2018/06/28
   End If
   
   'Added by Lydia 2022/12/27 ¼W¥[FCP/P/FG®×¸¹®Éªº¨t²Î³qª¾ (½Ð¹q¸£¤¤¤ß¤ñ·Óªþ¥ó·s®×¥ß¨÷PUB_GetTCTmail³qª¾©Ó¿ì¤Î¬ÛÃö¤H­û)
   'Modified by Lydia 2024/12/13 §ï¦¨¼Ò²ÕProc_FCPNewCaseEmail
   If intSaveMode = 1 And intModifyKind = 0 And (m_Pa(1) = "FCP" Or m_Pa(1) = "P") Then
      Call Proc_FCPNewCaseEmail(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), mCP(9), mCP(10), mCP(12))
   End If
   'end 2024/12/13
   
   'Add by Sindy 2022/6/29
   'Modify By Sindy 2023/5/31
   'If InStr(pType, "¥~±M«H¥ó¨R¾P") > 0 And m_strIR01 <> "" Then
   'Modify By Sindy 2025/8/18
   'If InStr(pType, "«H¥ó¨R¾P") > 0 And m_strIR01 <> "" Then
   If m_strIR01 <> "" Then
   '2025/8/18 END
   '2023/5/31 END
      m_bolRecvOK = True
      m_strMCR11 = ""
      If m_bMRecvBatch = True Then '¦h®×¦¬¤å
         '§ó·sÁ`¦¬¤å¸¹
         mStrSql = "update multiCaseRecv set mcr11='" & mCP(9) & "'" & _
                  " where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                  " and mcr02='" & m_Pa(1) & "' and mcr03='" & m_Pa(2) & "' and mcr04='" & m_Pa(3) & "' and mcr05='" & m_Pa(4) & "'" & _
                  " and mcr06='" & mCP(10) & "'"
                  cnnConnection.Execute mStrSql
                  
         'Modify By Sindy 2022/8/26
         '¤U¸ü«H¥óÀÉ,¤W¶Ç¨÷©v°Ï
         Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, mCP(9))
                  
         'ÀË¬d¦h®×¦¬¤åª¬ªp
         strTmp1(0) = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                     " and mcr02||mcr03||mcr04||mcr05<>'" & m_Pa(1) & m_Pa(2) & m_Pa(3) & m_Pa(4) & "'" & _
                     " and mcr11 is null"
         intJ = 1
         Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
         If intJ = 1 Then
            m_bolRecvOK = False '©|¦³¥¼¦¬¤å

            'Modify By Sindy 2022/8/26 ¦¹³BMark,µ{¦¡©¹¤W²¾
'            '¤U¸ü«H¥óÀÉ,¤W¶Ç¨÷©v°Ï
'            Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, mCP(9))
         Else
            m_bolRecvOK = True '¬O§_¦¬§¹¤å
            '§ì²Ä¤@µ§ªºÁ`¦¬¤å¸¹
            strTmp1(0) = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                        " and mcr02||mcr03||mcr04||mcr05=mcr07||mcr08||mcr09||mcr10 and mcr11 is not null"
            intJ = 1
            Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
            If intJ = 1 Then
               m_strMCR11 = rsQD.Fields("mcr11")
               RetVal = RetVal & IIf(RetVal <> "", ",", "") & "MCR11:" & m_strMCR11
            Else
               MsgBox "¦h®×¦¬¤å¡AµLÅª¨ú¨ì²Ä¤@µ§®×¥óªºÁ`¦¬¤å¸¹¡A½Ð¬¢¹q¸£¤¤¤ß!!", vbExclamation '¦¹ª¬ªpÀ³¤£·|µo¥Í, ¥H¨¾¥~¤@
               GoTo ErrHand
            End If
         End If
      End If
      If m_bolRecvOK = True Then '¬O§_¦¬§¹¤å=>¥þ³¡¦¬§¹¤å
         RetVal = RetVal & IIf(RetVal <> "", ",", "") & "m_bolRecvOK = True"
         '¦h®×¦¬¤åªºÁ`¦¬¤å¸¹­n¶Ç¤J²Ä¤@µ§Á`¦¬¤å¸¹
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, _
               IIf(m_strMCR11 <> "", "¦h®×¦¬¤å", "frm010001"), _
               IIf(m_strMCR11 <> "", m_strMCR11, mCP(9))
      End If
   End If
   '2022/6/29 END
   'Added by Lydia 2023/07/28 ¥~±M-FCP±M§Q³sµ²®×ºÞ¨î¡G¦¬¤å¯S©w®×¥ó©Ê½è, ¦Û°Ê¦¬¤å¡u³qª¾¸ê°TÅÜ§ó961¡v,µo¤@«ÊEmailµ¹©Ó¿ì¤uµ{®v
   If m_Pa(1) = "FCP" And m_Pa(177) = "Y" Then
      If PUB_GetFCPlinkMC("3", mCP(5), m_Pa, mCP(9), mCP(10)) = True Then
      End If
   End If
   'end 2023/07/28
   
   If bolError Then
      'Add By Sindy 2022/9/27
      If UCase(pFormName) <> UCase("frm090801_New") Then
      '2022/9/27 END
         cnnConnection.RollbackTrans
      End If
      ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
      IsSaveData = False
   Else
      'Add By Sindy 2022/9/27
      If UCase(pFormName) <> UCase("frm090801_New") Then
      '2022/9/27 END
         cnnConnection.CommitTrans
         
         'Added by Morgan 2024/12/26
         If bolAutoFCP414 Then
            MsgBox "¦]¤w¹O´Á­­¡A¤w¦P®É¦¬¤å¤@¹DBÃþ¡u414¥Ó½Ð´_Åv¡v¡I", vbInformation
         ElseIf strAlert <> "" Then
            MsgBox strAlert, vbInformation
         End If
         'end 2024/12/26
      End If
      InsertPatentDB = True
      'mHC(2) = pa02 '²¾¨ì¥~¼hÅÜ§ó
   End If
   'mHC(2) = pa02 '²¾¨ì¥~¼hÅÜ§ó
   Set rsQD = Nothing
   Set adoquery = Nothing
   Exit Function
   
ErrHand:
   If PUB_CheckFormExist("frmpic002") = True Then Unload frmpic002 'Add By Sindy 2022/7/11
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.RollbackTrans
   End If
   ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
   IsSaveData = False
   Set rsQD = Nothing
   Set adoquery = Nothing
End Function

'Added by Lydia 2022/09/05 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G­×§ï±M§Q°ò¥»ÀÉ(±qfrm010005.UpdatePatentDatabase©â¥X¨Ó)
Private Function UpdatePatentDB(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intChoose As Integer, _
                ByRef m_Pa() As String, ByRef mCP() As String, ByVal mCU30 As String, ByVal mInventorNo As String, ByVal mChkVal As String, Optional ByRef IsSaveData As Boolean, _
                Optional ByVal pType As String, Optional ByVal pCaseNo As String, Optional ByVal mTCTVal As String, Optional ByRef mTCTList As String) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
'reTurnVal : ¦^¶Ç­È
'mChkVal¡G¶Ç¤J¨ä¥L¾Þ§@µ²ªG
Dim adoquery As New ADODB.Recordset
Dim rsQD As New ADODB.Recordset
Dim varInventorNo As Variant
Dim strInventor(100) As String
Dim m_strCPM34 As String
Dim m_CaseNaTmp() As String  '¯S®íºÞ¨î¤§ÃöÁp®×
'ªk«ß©Ò®×·½¦¬¤å
Dim m_LOS02 As String '®×·½®×¥óÃþ«¬
Dim m_LOS15 As String '®×·½³æ¸¹
Dim strCustomer(4) As String

Dim stUpdate As String
Dim tmpArr As Variant

   If IsSaveData = True Then
      Exit Function
   End If
   IsSaveData = True
   
'*********¯S®íºÞ¨îªºÅÜ¼Æ*************
   If pType = "CFP­^°ê²æ¼Ú®×" And pCaseNo <> "" Then
       ReDim m_CaseNaTmp(1 To TF_PA)
       Call ChgCaseNo(pCaseNo, m_CaseNaTmp)
   Else
       ReDim m_CaseNaTmp(1 To 4) '¹w³]°}¦CÁ×§Kµ{¦¡¥X¿ù
       'Modify By Sindy 2025/8/18
       'If pType = "LOS®×·½¦¬¤å" And pCaseNo <> "" Then
       If InStr(pType, "LOS®×·½¦¬¤å") > 0 And pCaseNo <> "" Then
       '2025/8/18 End
           m_LOS02 = Mid(pCaseNo, 1, InStr(pCaseNo, ",") - 1) '®×·½®×¥óÃþ«¬
           m_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1, 8) '®×·½³æ¸¹ 'Modify By Sindy 2025/8/18 +, 8)
       End If
       '¥~±M«H¥ó¨R¾P: ¥u¦³·s¼W¦¬¤åªº¥\¯à
   End If
'***********************************
   strTmp1(0) = "select cpm34 from casepropertymap where cpm01='" & mCP(1) & "' and cpm02='" & mCP(10) & "' "
   Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
   If intJ = 1 Then
       m_strCPM34 = "" & rsQD.Fields("cpm34")
   End If

On Error GoTo ErrHand
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.BeginTrans
   End If
   
        'Add by Lydia 2014/10/31 ¶}©ñ¥~±Mµ{§Ç¤H­û¥i¶i¤J±M§Q³B¨t²Î¾Þ§@FMP¾ÈµØ®×¥ó=>¼g¤J¥N²z¤H
        If InStr(mChkVal, "¾ÈµØ®×¥ó½T»{") > 0 Then
            mStrSql = "update caseprogress set cp44='Y53374000' where cp09='" & mCP(9) & "' "
        Else
            mStrSql = "update caseprogress set cp44='' where cp09='" & mCP(9) & "' "
        End If
        cnnConnection.Execute mStrSql
        'end. 'Add by Lydia 2014/10/31
            
            varInventorNo = Split(mInventorNo, ",")
            For intJ = 0 To UBound(varInventorNo)
               strInventor(intJ) = varInventorNo(intJ)
            Next
            For intJ = intJ + 1 To 99 '9
               strInventor(intJ) = ""
            Next
            strTmp1(0) = PUB_GetPatentInventorList(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4))
            
            'Add By Sindy 2014/11/6 §ó·s±M§Qµo©ú¤HÀÉ
            If strTmp1(0) <> mInventorNo Then
               mStrSql = "delete from patentInventor where pi01=" + CNULL(m_Pa(1)) + " and pi02=" + CNULL(m_Pa(2)) + " and pi03=" + CNULL(m_Pa(3)) + " and pi04=" + CNULL(m_Pa(4))
               Pub_SeekTbLog mStrSql 'Add By Sindy 2017/8/23
               cnnConnection.Execute mStrSql
               For intJ = 0 To 99
                  If strInventor(intJ) <> "" Then
                     mStrSql = "INSERT into patentInventor(pi01,pi02,pi03,pi04,pi05,pi06) VALUES(" & _
                              CNULL(m_Pa(1)) & "," & CNULL(m_Pa(2)) & "," & CNULL(m_Pa(3)) & "," & CNULL(m_Pa(4)) & "," & intJ + 1 & ",'" & strInventor(intJ) & "')"
                     Pub_SeekTbLog mStrSql 'Add By Sindy 2017/8/23
                     cnnConnection.Execute mStrSql
                  Else
                     Exit For
                  End If
               Next intJ
            End If
            '2014/11/6 END
            
            'Memo by Lydia 2021/08/17 §R°£ÂÂµ{¦¡½X¡G±M§Qµo©ú¤H¦b±M§Q°ò¥»ÀÉ60~69
            'Memo by Lydia 2022/08/22 + PA149, PA176
            'Modify By Sindy 2022/12/7 +PA178
            mStrSql = "update patent set pa05=" + CNULL(ChgSQL(m_Pa(5))) + ",pa06=" + CNULL(ChgSQL(m_Pa(6))) + _
               ",pa07=" + CNULL(ChgSQL(m_Pa(7))) + ",pa08=" + CNULL(m_Pa(8)) + ",pa09=" + CNULL(m_Pa(9)) + _
               ",pa26=" + CNULL(m_Pa(26)) + ",pa27=" + CNULL(m_Pa(27)) + _
               ",pa28=" + CNULL(m_Pa(28)) + ",pa29=" + CNULL(m_Pa(29)) + ",pa30=" + CNULL(m_Pa(30)) + _
               ",pa75=" + CNULL(m_Pa(75)) + ",pa77=" + CNULL(m_Pa(77)) + _
               ",pa149=" + CNULL((m_Pa(149))) + ",pa158=" + CNULL((m_Pa(158))) + ",pa176=" + CNULL((m_Pa(176))) + ",pa178=" + CNULL((m_Pa(178)))
   'Add by Morgan 2008/8/5 +PA149
   '¨Ö¤J¤W­±
   'If UCase(m_PA(149)) <> "PA149" Then
   '   mStrSQL = mStrSQL + ",PA149=" + CNULL(m_PA(149))
   'End If
   
   'Added by Lydia 2017/11/14 +PA150
   'Mark by Lydia 2022/08/22 ¦b¥~±M¤uµ{®v©R¦W¤W½u«á,¨M©w¦¬¤å¤£¿é¤J©R¦W²Õ§O
   'If fraTCT.Visible = True And txtData(2).Text <> txtData(2).Tag Then
   '   mStrSQL = mStrSQL + ",pa150=" + CNULL(IIf(txtData(2).Text = "B", "", txtData(2).Text))
   'End If
   ''end 2017/11/14
   
   'Added by Morgan 2021/7/21
   'Modified by Lydia 2022/08/22 debug
   'If PA149 <> "" Then
   '¨Ö¤J¤W­±
   'If PA176 <> "" Then
   '   mStrSQL = mStrSQL + ",PA176='" & PA176 & "'"
   'End If
   ''end 2021/7/21
   
   'Add By Sindy 2010/3/8 ¼W¥[Ápµ¸¤Hpa51~pa56Äæ¦ì
   'If bolCancel = True Then 'ª½±µ¨Ö¤J¤U­±»yªk
      mStrSql = mStrSql + ",PA51=" + CNULL(ChgSQL(m_Pa(51))) + ",PA52=" + CNULL(ChgSQL(m_Pa(52))) + _
                                ",PA53=" + CNULL(ChgSQL(m_Pa(53))) + ",PA54=" + CNULL(ChgSQL(m_Pa(54))) + _
                                ",PA55=" + CNULL(ChgSQL(m_Pa(55))) + ",PA56=" + CNULL(ChgSQL(m_Pa(56)))
   'End If
   mStrSql = mStrSql & ", PA47=" & CNULL(ChgSQL(m_Pa(47))) & " , PA48=" & CNULL(ChgSQL(m_Pa(48)))
   mStrSql = mStrSql + " where pa01=" + CNULL(m_Pa(1)) + " and pa02=" + CNULL(m_Pa(2)) + " and pa03=" + CNULL(m_Pa(3)) + " and pa04=" + CNULL(m_Pa(4))
   cnnConnection.Execute mStrSql
   strCustomer(0) = m_Pa(26)
   strCustomer(1) = m_Pa(27)
   strCustomer(2) = m_Pa(28)
   strCustomer(3) = m_Pa(29)
   strCustomer(4) = m_Pa(30)
   
   For intJ = 0 To 4
          mStrSql = "update patent set pa" + Format(31 + intJ) + "=(select cu23 from customer where cu01=" + CNULL(Mid(strCustomer(intJ), 1, 8)) + " and cu02=" + CNULL(Mid(strCustomer(intJ), 9, 1)) + _
             "),pa" + Format(36 + intJ) + "=(select cu24||' '||cu25||' '||cu26||' '||cu27||' '||cu28||' '||cu102 from customer where cu01=" + CNULL(Mid(strCustomer(intJ), 1, 8)) + " and cu02=" + CNULL(Mid(strCustomer(intJ), 9, 1)) + _
             "),pa" + Format(41 + intJ) + "=(select cu29 from customer where cu01=" + CNULL(Mid(strCustomer(intJ), 1, 8)) + " and cu02=" + CNULL(Mid(strCustomer(intJ), 9, 1)) + ") where pa01=" + CNULL(m_Pa(1)) + " and pa02=" + CNULL(m_Pa(2)) + " and pa03=" + CNULL(m_Pa(3)) + " and pa04=" + CNULL(m_Pa(4))
          cnnConnection.Execute mStrSql
   Next

      'Add by Morgan 2008/8/28 ¹w³]©Ó¿ì´Á­­
      '2010/3/11 MODIFY BY SONIA
      'If m_PA(1) = "FCP" Then
      If m_Pa(1) = "FCP" Or m_Pa(1) = "FG" Then
      'Modify by Morgan 2008/10/23
         'Ciba Y45697ªº¦~¶O©Ó¿ì´Á­­±¾15­Ó¤u§@¤Ñ
         If m_Pa(75) = "Y45697000" And mCP(10) = "605" Then
            mCP(48) = CompWorkDay(15, strSrvDate(1))
            If Val(mCP(6)) > 0 And Val(mCP(48)) > Val(mCP(6)) Then
               mCP(48) = mCP(6)
            End If
         
         'Added by Morgan 2012/7/13
         '¥[³t¼f¬d­n§PÂ_¤w¿é¤J³qª¾¹ê¼f¤é¤~±¾©Ó¿ì´Á­­
         'Modified by Morgan 2024/11/18 +477¦A¼f¬d¥[³t¼f¬d¨Ã§ï¥Î±M¥Î¼Ò²Õ§PÂ_
         'ElseIf mCP(10) = "422" Then
         '   If PUB_ChkCPExist(m_Pa, "1204") Then
         ElseIf mCP(10) = "422" Or mCP(10) = "447" Then
            If PUB_Chk1204(m_Pa) Then
         'end 2024/11/13
               mCP(48) = Pub_GetHandleDay("FCP", "000", mCP(10), , mCP(6))
            End If
         'end 2012/7/13
         
         'Add By Sindy 2021/6/24 968¦^´_»¡©ú®Ñ®Õ¾\
         ElseIf mCP(10) = "968" Then
            mCP(48) = Pub_GetHandleDay("FCP", "000", mCP(10), , mCP(6), , , m_Pa(1) & "-" & m_Pa(2) & "-" & m_Pa(3) & "-" & m_Pa(4))
            
         'end 2008/10/23
         ElseIf InStr(SkipCasePtyList, mCP(10)) = 0 Then
            mCP(48) = Pub_GetHandleDay("FCP", "000", mCP(10), , mCP(6))
         End If
         stUpdate = ",cp48=" & CNULL(mCP(48), True)
         'Y54732000 & X30299000²Õ¦X¤§¦^¥N©Ó¿ì´Á­­¤U­±¥t¦³§ó·s
         
         'Add By Sindy 2021/4/29 ¤£¬O¥DºÞ¾÷Ãö´Á­­
         If m_strCPM34 = "N" And strSrvDate(1) >= ¥~±M¥xÆW®×¬ù©w´Á­­±Ò¥Î¤é Then
            '(2)¦¬¤å®ÉµL³]¥»©Ò´Á­­¡A¥H©Ó¿ì´Á­­¡Ï5­Ó¤u§@¤Ñ¬°¥»©Ò´Á­­
            If Val(mCP(6)) = 0 Then
               mCP(6) = PUB_GetFCPOurDeadline(DBDATE(mCP(48)), , , , "N")
            '(1)¦¬¤å®É¦³³]¥»©Ò´Á­­¡A¦Û°Ê³Æµù:¥»©Ò´Á­­¬°yyy/mm/dd(¥»©Ò´Á­­)
            ElseIf InStr(mChkVal, "­ì¥»©Ò´Á­­¬°") > 0 Then '¦³²§°Ê®É
                mCP(64) = Mid(mChkVal, InStr(mChkVal, "") + 1, 15) & "¤w­×§ï;" & mCP(64)
            End If
         End If
      End If

      'Added by Morgan 2012/6/20
      'Modified by Lydia 2022/09/20 ¦]¬°P®×¦³Trigger·|¦Û°Ê³]©w¹q¤l°e¥óCP118 = Y, ©Ò¥H§ï¦¨¨â­Ó§PÂ_
      'If chkWebApp.Visible = True Then
      '   stUpdate = ",cp118='" & IIf(chkWebApp.Value = 1, "Y", "") & "'"
      'End If
      ''end 2012/6/20
      stUpdate = IIf(mCP(118) = "YY", ", CP118='Y' ", IIf(mCP(118) = "YN", ", CP118=null ", ""))
      
  'Add by Morgan 2006/6/23 (¨Ö¤J) Åý»P¤H1-5,¨üÅý¤H1-5 ; CP56, CP89, CP90, CP91, CP92
   'Modify By Sindy 2012/11/06 +CP150 ¦³¡¹¡¹ªºÀ³¦¬±b´ÚÃ±®Ö±±ºÞ
   'Modified by Lydia 2022/09/20 ®³±¼ cp118=" & CNULL(mCP(118))
   mStrSql = "update caseprogress set cp05=" + CNULL(mCP(5)) + ",cp06=" + CNULL(mCP(6)) + ",cp07=" + CNULL(mCP(7)) + ",cp10=" + CNULL(mCP(10)) + _
            ",cp11=" + CNULL(mCP(11)) + ",cp13=" + CNULL(mCP(13)) + ",cp14=" + CNULL(mCP(14)) + ",cp16=" + CNULL(mCP(16)) + ",cp17=" + CNULL(mCP(17)) + _
            ",cp18=" + CNULL(mCP(18)) + ",cp19=" + CNULL(mCP(19)) + ",cp32=" + CNULL(mCP(32)) + ",cp56=" + CNULL(mCP(56)) + _
            ",cp33=" & CNULL(mCP(33)) & ",cp34=" & CNULL(mCP(34)) & ",CP64=" + CNULL(ChgSQL(mCP(64))) + ",cp89=" + CNULL(mCP(89)) + ",cp90=" + CNULL(mCP(90)) + _
            ",cp91=" + CNULL(mCP(91)) + ",cp92=" + CNULL(mCP(92)) & stUpdate & ",cp150=" & CNULL(mCP(150)) & " where cp09='" + mCP(9) + "'"
   cnnConnection.Execute mStrSql
   
   'add by sonia 2019/7/31 Y54732000 & X30299000²Õ¦X,¥B·|½Z924µo¤å«á,·s®×Â½Ä¶201µo¤å«e¦¬¤å¤§¦^¥N902,³]¦^¥N¬ÛÃö¦¬¤å¸¹±¾·|½Z,©Ó¿ì´Á­­±¾·s®×Â½Ä¶ªº¥»©Ò´Á­­
   If m_Pa(75) = "Y54732000" And Left(m_Pa(26), 8) = "X3029900" And mCP(10) = "902" Then
      strTmp1(0) = "select c2.cp06,c1.cp09 from caseprogress c1,caseprogress c2 where c1.cp01='" & m_Pa(1) & "' and c1.cp02='" & m_Pa(2) & "' and c1.cp03='" & m_Pa(3) & "' and c1.cp04='" & m_Pa(4) & "' and c1.cp10='924' and c1.cp27>0 " & _
                  "   and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and '201'=c2.cp10(+) and c2.cp158=0"
      intJ = 1
      Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
      If intJ = 1 Then
         mStrSql = "update caseprogress set cp43='" & "" & rsQD(1) & "',cp48=" & "" & rsQD(0) & " where cp09=" & CNULL(mCP(9))
         cnnConnection.Execute mStrSql
      End If
   End If
   'end 2019/7/31
   
   'Modified by Morgan 2012/4/25 +cp71(Àu¥ýÅv¥÷¼Æ)
   mStrSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(mCP(13)) + ")" & IIf(mCP(71) <> "", ", cp71=" & CNULL(mCP(71)), "") & " where cp09=" + CNULL(mCP(9))
   'end 2012/4/25
   cnnConnection.Execute mStrSql
   
   'Add By Sindy 2009/07/06
   If mCP(53) <> "" And mCP(54) <> "" Then
      If mCP(10) = "601" Then
         If Val(mCP(54)) > 0 Then
            mStrSql = "update caseprogress set cp53=" & mCP(53) & ",cp54=" & mCP(54) & " where cp09=" & CNULL(mCP(9))
         End If
      Else
         mStrSql = "update caseprogress set cp53=" & mCP(53) & ",cp54=" & mCP(54) & " where cp09=" & CNULL(mCP(9))
      End If
      cnnConnection.Execute mStrSql
   End If
   '2009/07/06 End
           
   'Added by Lydia 2022/11/29 «D¤º³¡¦¬¤å¨Ã¥B¦³¶O¥Î¡A¥ý²Î¤@³]©wCP20=Null ;
   If intChoose = 0 And Val(mCP(16)) > 0 Then
      'Modified by Lydia 2024/05/28 §ï¦¨¼Ò²Õ
      'If (m_Pa(16) = "" And InStr("FCP062174000", m_Pa(1) & m_Pa(2) & IIf(m_Pa(3) = "", "0", m_Pa(3)) & IIf(m_Pa(4) = "", "00", m_Pa(4))) > 0) Or _
        (m_Pa(16) <> "1" And InStr("FCP067004000", m_Pa(1) & m_Pa(2) & IIf(m_Pa(3) = "", "0", m_Pa(3)) & IIf(m_Pa(4) = "", "00", m_Pa(4))) > 0) Then
      If PUB_GetCP20forSpec(m_Pa(1), m_Pa(2), IIf(m_Pa(3) = "", "0", m_Pa(3)), IIf(m_Pa(4) = "", "00", m_Pa(4)), m_Pa(16)) = "N" Then
          '±Æ°£¯S©w®×¥ó
      Else
          stUpdate = ""
          If Left(mCP(12), 1) = "F" And (m_Pa(1) = "FCP" Or m_Pa(1) = "FG" Or m_Pa(1) = "P") Then
             stUpdate = PUB_GetCP20(m_Pa(1), mCP(10))
          End If
          If stUpdate = "" Then
             mStrSql = "update caseprogress set cp20=null where cp09=" + CNULL(mCP(9))
             cnnConnection.Execute mStrSql
          End If
      End If
   End If
   'end 2022/11/29
   
           'Add By nickc 2007/08/21
           '­Y¬°±µ¬¢°O¿ý³æ(Âd¥x¦¬¤å)
           'Modify by Morgan 2007/10/26 ¶O¥Î¥i§ï®É¤~°µ¡A§_«h¤w¦¬´Ú¸ê®Æ·|³QÁÙ­ì
           If intChoose = 0 And mCP(60) = "" Then 'mCP(60) = "" => txtPatent(17).Enabled = True
               '¥¼¦¬ª÷ÃB = ¶O¥Î
               mStrSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(mCP(9))
               cnnConnection.Execute mStrSql
           End If
           
   'Add By Cheng 2002/05/10
   '­Y¬°¤º³¡¦¬¤å§@·~®É, ®×¥ó¶i«×ÀÉªº¬O§_¦V«È¤á¦¬´Ú³]©w¬°"N"
   If intChoose = 1 Then
      mStrSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(mCP(9))
      cnnConnection.Execute mStrSql
   End If
   
   'Modify By Sindy 2023/11/3 mark:¦¬¤å¤w¤£»Ý°õ¦æ¦¹¬qµ{¦¡,¦]±µ¬¢³æ¤w·|¦^¼g¶l»¼°Ï¸¹
'   mStrSql = "update customer set cu30=" + CNULL(mCU30) + " where cu01=" + CNULL(Mid(m_Pa(26), 1, 8)) + " and cu02=" + CNULL(Mid(m_Pa(26), 9, 1))
'   cnnConnection.Execute mStrSql
   '2023/11/3 EMD
   
   UpdatePatentDB = True
   adoquery.CursorLocation = adUseClient
   adoquery.Open "select np01 from nextprogress where np02 = '" & m_Pa(1) & "' and np03 = '" & m_Pa(2) & "' and np04 = '" & m_Pa(3) & "' and np05 = '" & m_Pa(4) & "' and np06 is null and np07 = '" & mCP(10) & "'", cnnConnection, adOpenStatic, adLockReadOnly
   'Modify By Cheng 2002/05/10
   '­Y¦b¤U¤@µ{§ÇÀÉ¥u§ì¨ì¤@µ§¸ê®Æ®É, ¤~­n§ì¤U¤@µ{§ÇÀÉªºÁ`¦¬¤å¸¹§ó·s®×¥ó¶i«×ÀÉªº¬ÛÃöÁ`¦¬¤å¸¹
   If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
      If IsNull(adoquery.Fields(0).Value) = False Then
         'Add by Morgan 2010/6/30 ²§Ä³µªÅG¡BÁ|µoµªÅG­n¤@¨Ã§ó·s¹ï³y¸ê®Æ
         If (mCP(10) = "802" Or mCP(10) = "804") Then
            cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where CP09 = '" & mCP(9) & "'", intJ
         Else
         'End 2010/6/30
            cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where CP09 = '" & mCP(9) & "'"
         End If
      End If
   End If
   adoquery.Close
   'add by nickc 2008/05/02 Àx¦s¹w©w¦¬´Ú¤é
   'Remove by Lydia 2018/08/22 (À³¦¬±b´ÚºÞ±±)¨ú®ø¹w©w¦¬´Ú¤é,§ï¦¨¥I´Ú¶g´Á
'   Dim rtCnt As Integer
'   'Modify by Morgan 2010/12/9
'   'If txtPatent(28) <> "" Then
'   '    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '"& mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " from receivablesday where rd01='"& mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '"& mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " ", rtCnt
'   If txtPatent(28) <> "" And txtPatent(28) <> txtPatent(28).Tag Then
'       cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '"& mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0) + 1,'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " from receivablesday where rd01='"& mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '"& mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
'   'end 2010/12/9
'       If rtCnt = 0 Then
'           cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '"& mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtPatent(28)) & " from dual "
'       End If
'   End If
   'end 2018/08/22
   
   'Added by Lydia 2017/11/14 FCP®×¥ó©R¦W¹q¤l¤Æ¡G¤¤»¡¿é¤J¬ÛÃö³]©w-¦sÀÉ
   'Modified by Lydia 2019/06/11 §PÂ_¨«©R¦W¬yµ{¤~ÀË¬d; FCP-62285¦¬307¤À³Î,­×§ï®É¼W¥[¥Ó½Ð¤H2-X76639·|¼u¥¼¿é¤J¤À®×²Õ§O
   'If fraTCT.Visible = True And fraTCT.Enabled = True Then
   'Modified by Lydia 2019/07/04 ¤À³Î®×¤£¨«©R¦W¬yµ{¡A¦ý¬O­n¯à¤Ä¿ï¨ä¥L¦¬¤å
   'If fraTCT.Visible = True And fraTCT.Enabled = True And InStr(FcpAddTct, mcp(10)) > 0 Then
   'Modified by Lydia 2023/03/09  ±Æ°£¡u¿é¤J°lÂÜ¬y¤ô¸¹,¤£¨«©R¦W§@·~¡v+InStr(mTCTVal, "|") > 0
   If mTCTVal <> "" And InStr(mTCTVal, "|") > 0 Then
       tmpArr = Split(mTCTVal, "|")
       If Trim(tmpArr(2)) <> "" Then
          Call PUB_UpdTCTrecord(Trim(tmpArr(0)), Trim(tmpArr(1)), Trim(tmpArr(2)), mTCTList, m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), m_Pa(5), m_Pa(6), mCP(9), mCP(10), _
                mCP(6), mCP(7), mCP(13), m_Pa(8), m_Pa(9), m_Pa(16), m_Pa(14), m_Pa(26) & "," & m_Pa(27) & "," & m_Pa(28) & "," & m_Pa(29) & "," & m_Pa(30), _
                m_Pa(75), m_Pa(150), Trim(tmpArr(3)), Trim(tmpArr(4)))
       End If
   End If
   
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.CommitTrans
   End If
   Set rsQD = Nothing
   Set adoquery = Nothing
   Exit Function
   
ErrHand:
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.RollbackTrans
   End If
   ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
   IsSaveData = False
   Set rsQD = Nothing
End Function

'Added by Lydia 2022/09/05 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G±M§Q¦¬¤å(±qfrm010005.SaveDatabase©â¥X¨Ó)
'Modify By Sindy 2024/11/21 + , Optional ByVal m_intCRC As Integer = 0: ¦Û°Ê¦¬¤åªº®×¥ó©Ê½è¶¶§Ç
Public Function PUB_SaveFrm010005(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intChoose As Integer, _
                ByRef m_Pa() As String, ByRef mCP() As String, ByVal mCU30 As String, ByVal mInventorNo As String, ByVal mChkVal As String, Optional ByRef IsSaveData As Boolean, _
                Optional ByVal pType As String, Optional ByVal pCaseNo As String, Optional ByRef RetVal As String, _
                Optional ByVal mTCTVal As String, Optional ByRef mTCTList As String, Optional ByVal m_intCRC As Integer = 0) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
'reTurnVal : ¦^¶Ç­È
'mChkVal¡G¶Ç¤J¨ä¥L¾Þ§@µ²ªG
'mTctVal ¡G¶Ç¤Jµe­±¦³Ãö©R¦W§@·~ªº¸ê®Æ
'mTctList ¡G¦^¶Ç©R¦W§@·~¤@¨Ö²£¥Í¤§¦¬¤å¸¹
Dim m_SalesST15 As String, m_SalesST06 As String
Dim m_SalesDeptName As String
Dim m_CP10Name As String '¦¬¤å¤§®×¥ó©Ê½è¦WºÙ
Dim m_Na01Name As String '¥Ó½Ð°ê®a¦WºÙ
Dim m_CaseNaTmp() As String  '¯S®íºÞ¨î¤§ÃöÁp®×
'ªk«ß©Ò®×·½¦¬¤å
Dim m_LOS02 As String '®×·½®×¥óÃþ«¬
Dim m_LOS15 As String '®×·½³æ¸¹
Dim strDivisionalEmp As String 'Add By Sindy 2022/9/30
Dim m_bolFMP As Boolean 'Added by Lydia 2023/05/04
Dim bolUpdCP06IsSysDate As Boolean 'Add By Sindy 2023/6/20
Dim m_bolFMP2 As Boolean 'Added by Morgan 2025/6/18 ¬O§_¾ÈµØ®×
   
'*********¯S®íºÞ¨îªºÅÜ¼Æ*************
   If pType = "CFP­^°ê²æ¼Ú®×" And pCaseNo <> "" Then
       ReDim m_CaseNaTmp(1 To TF_PA)
       Call ChgCaseNo(pCaseNo, m_CaseNaTmp)
   Else
       ReDim m_CaseNaTmp(1 To 4) '¹w³]°}¦CÁ×§Kµ{¦¡¥X¿ù
       'Modify By Sindy 2025/8/18
       'If pType = "LOS®×·½¦¬¤å" And pCaseNo <> "" Then
       If InStr(pType, "LOS®×·½¦¬¤å") > 0 And pCaseNo <> "" Then
       '2025/8/18 END
           m_LOS02 = Mid(pCaseNo, 1, InStr(pCaseNo, ",") - 1) '®×·½®×¥óÃþ«¬
           'm_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1) '®×·½³æ¸¹
           m_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1, 8) '®×·½³æ¸¹ 'Modify By Sindy 2025/8/18 + , 8)
       End If
       '¥~±M«H¥ó¨R¾P: ¥u¦³·s¼W¦¬¤åªº¥\¯à
   End If
'***********************************
   intJ = ClsPDGetCaseProperty(m_Pa(1), mCP(10), m_CP10Name, IIf(m_Pa(9) <> "000", True, False))
   m_Na01Name = PUB_GetNationName(m_Pa(9))
   m_SalesST15 = GetST15(mCP(13), m_SalesDeptName, , m_SalesST06)
   'Added by Lydia 2023/05/11 ¦]¬°PUB_ReadCaseData·|¦^¶Ç6½X«È¤á½s¸¹,©Ò¥H¥ý²Î¤@«È¤á½s¸¹
   m_Pa(26) = ChangeCustomerL(m_Pa(26))
   m_Pa(27) = ChangeCustomerL(m_Pa(27))
   m_Pa(28) = ChangeCustomerL(m_Pa(28))
   m_Pa(29) = ChangeCustomerL(m_Pa(29))
   m_Pa(30) = ChangeCustomerL(m_Pa(30))
   'end 2023/05/11
   
   RetVal = "" '¦^¶Ç­È
   
   '­Y¬°·s¼W
   If intModifyKind = 0 Then
      PUB_SaveFrm010005 = InsertPatentDB(pFormName, intSaveMode, intModifyKind, intChoose, m_Pa, mCP, mCU30, mInventorNo, mChkVal, IsSaveData, pType, pCaseNo, RetVal, mTCTVal, mTCTList)
   '­Y¬°­×§ï
   Else
      PUB_SaveFrm010005 = UpdatePatentDB(pFormName, intSaveMode, intModifyKind, intChoose, m_Pa, mCP, mCU30, mInventorNo, mChkVal, IsSaveData, pType, pCaseNo, mTCTVal, mTCTList)
   End If
   If PUB_SaveFrm010005 = False Then Exit Function 'Add By Sindy 2022/9/28 ¦sÀÉ¥¢±Ñ,«áÄò¤£ÀË¬d
'add by nickc 2007/11/09 ´ú¸Õ¸Ñ¨Mmail µo¤£¨ìªº®É­Ô·|¦s¨âµ§ªº¿ù»~
'On Error GoTo 0    'Âk¹s
   On Error GoTo ErrHand 'Add By Sindy 2022/9/29
   'Add By Sindy 2022/12/29 ­«ÅªCP,¦]«eÀYUpdate¨ç¼Æµ{¦¡¦³¥i¯àª½±µ¦sDB,¨S¦³§ó·scp³¯¦C­È
   strTmp1(0) = mCP(9)
   Erase mCP
   ReDim Preserve mCP(TF_CP) As String
   mCP(9) = strTmp1(0)
   'Modified by Lydia 2023/05/11 + false
   Call PUB_ReadCaseProgressDatabase(mCP(), 1, False)
   '2022/12/29 END
   m_bolFMP = PUB_ChkIsFMP(mCP(1), mCP(2), mCP(3), mCP(4)) 'Added by Lydia 2023/05/04
   
   If m_bolFMP Then m_bolFMP2 = PUB_GetFMP2toP(mCP(1), mCP(2), mCP(3), mCP(4)) 'Added by Morgan 2025/6/18
   
   'Added by Lydia 2021/01/08 CFP­^°ê²æ¼Ú®×¡G½Æ»s¥Nªí¹Ï
   If m_Pa(1) = "CFP" And pType = "CFP­^°ê²æ¼Ú®×" And m_CaseNaTmp(1) <> "" And m_CaseNaTmp(2) <> "" Then
      strTmp1(9) = ""
      If GetImgByteFile_Case(m_CaseNaTmp(1), m_CaseNaTmp(2), m_CaseNaTmp(3), m_CaseNaTmp(4), strTmp1(9), 0, strTmp1(5), strTmp1(6)) = True Then
          Call SaveImgByteFile(strTmp1(9), m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4), strTmp1(5), strTmp1(6))
      End If
   End If
   'end 2021/01/08
   
   'add by nickc 2005/09/05
   If intModifyKind = 0 Then
      Dim oContext As String
      Dim oMailCount As String
      Dim strTemp As String
      Dim m_strState As String
      'Add By Sindy 2021/2/1 ¤£±o¥N²zªº«áÄòÂÂ®×¦¬¤å±±ºÞ¡A³qª¾¦¬¤å¤H­û¡]CP13¡^
      If m_Pa(75) <> "" Then
        If GetAgentAndState(m_Pa(75), strTmp1(1), , , , m_Pa(1), m_strState, IIf(intSaveMode = 0, True, False)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥N²z¤H¡G " + m_Pa(75) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & m_Pa(75)
          End If
        End If
      End If
      If m_Pa(26) <> "" Then
        'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , m_Pa(2), m_Pa(3), m_Pa(4)
        If GetCustomerAndState(m_Pa(26), strTmp1(1), , , , m_Pa(1), m_strState, IIf(intSaveMode = 0, True, False), , m_Pa(2), m_Pa(3), m_Pa(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H1¡G " + m_Pa(26) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & m_Pa(26)
          End If
        End If
      End If
      If m_Pa(27) <> "" Then
        'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , m_Pa(2), m_Pa(3), m_Pa(4)
        If GetCustomerAndState(m_Pa(27), strTmp1(1), , , , m_Pa(1), m_strState, IIf(intSaveMode = 0, True, False), , m_Pa(2), m_Pa(3), m_Pa(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H2¡G " + m_Pa(27) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & m_Pa(27)
          End If
        End If
      End If
      If m_Pa(28) <> "" Then
        'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , m_Pa(2), m_Pa(3), m_Pa(4)
        If GetCustomerAndState(m_Pa(28), strTmp1(1), , , , m_Pa(1), m_strState, IIf(intSaveMode = 0, True, False), , m_Pa(2), m_Pa(3), m_Pa(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H3¡G " + m_Pa(28) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & m_Pa(28)
          End If
        End If
      End If
      If m_Pa(29) <> "" Then
        'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , m_Pa(2), m_Pa(3), m_Pa(4)
        If GetCustomerAndState(m_Pa(29), strTmp1(1), , , , m_Pa(1), m_strState, IIf(intSaveMode = 0, True, False), , m_Pa(2), m_Pa(3), m_Pa(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H4¡G " + m_Pa(29) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & m_Pa(29)
          End If
        End If
      End If
      If m_Pa(30) <> "" Then
        'Modified by Lydia 2023/03/06 ¶Ç¤J¥»©Ò®×¸¹ , , m_Pa(2), m_Pa(3), m_Pa(4)
        If GetCustomerAndState(m_Pa(30), strTmp1(1), , , , m_Pa(1), m_strState, IIf(intSaveMode = 0, True, False), , m_Pa(2), m_Pa(3), m_Pa(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H1¡G " + m_Pa(30) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & m_Pa(30)
          End If
        End If
      End If
      'Add By Sindy 2025/3/26
      If mCP(56) <> "" Then
         '¶Ç¤J®×¥ó©Ê½è , mCP(10)
         If GetCustomerAndState(mCP(56), strTmp1(1), , , , m_Pa(1), m_strState, IIf(intSaveMode = 0, True, False), , m_Pa(2), m_Pa(3), m_Pa(4), mCP(10)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "Åý»P¥Ó½Ð¤H1¡G " + mCP(56) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mCP(56)
          End If
        End If
      End If
      If mCP(89) <> "" Then
         '¶Ç¤J®×¥ó©Ê½è , mCP(10)
         If GetCustomerAndState(mCP(89), strTmp1(1), , , , m_Pa(1), m_strState, IIf(intSaveMode = 0, True, False), , m_Pa(2), m_Pa(3), m_Pa(4), mCP(10)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "Åý»P¥Ó½Ð¤H2¡G " + mCP(89) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mCP(89)
          End If
        End If
      End If
      If mCP(90) <> "" Then
         '¶Ç¤J®×¥ó©Ê½è , mCP(10)
         If GetCustomerAndState(mCP(90), strTmp1(1), , , , m_Pa(1), m_strState, IIf(intSaveMode = 0, True, False), , m_Pa(2), m_Pa(3), m_Pa(4), mCP(10)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "Åý»P¥Ó½Ð¤H3¡G " + mCP(90) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mCP(90)
          End If
        End If
      End If
      If mCP(91) <> "" Then
         '¶Ç¤J®×¥ó©Ê½è , mCP(10)
         If GetCustomerAndState(mCP(91), strTmp1(1), , , , m_Pa(1), m_strState, IIf(intSaveMode = 0, True, False), , m_Pa(2), m_Pa(3), m_Pa(4), mCP(10)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "Åý»P¥Ó½Ð¤H4¡G " + mCP(91) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mCP(91)
          End If
        End If
      End If
      If mCP(92) <> "" Then
         '¶Ç¤J®×¥ó©Ê½è , mCP(10)
         If GetCustomerAndState(mCP(92), strTmp1(1), , , , m_Pa(1), m_strState, IIf(intSaveMode = 0, True, False), , m_Pa(2), m_Pa(3), m_Pa(4), mCP(10)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "Åý»P¥Ó½Ð¤H5¡G " + mCP(92) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mCP(92)
          End If
        End If
      End If
      '2025/3/26 END
      If oContext <> "" Then
         strTemp = Mid(strTemp, 2)
         oContext = "¥»©Ò®×¸¹¡G " + m_Pa(1) + "-" + m_Pa(2) + "-" + m_Pa(3) + "-" + m_Pa(4) + vbCrLf + _
                    "®×¥ó¦WºÙ¡G " + m_Pa(5) + vbCrLf + _
                    "¦¬¤å¤é¡G " + ChangeTStringToTDateString(TransDate(mCP(5), 1)) + vbCrLf + _
                    "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf + vbCrLf + _
                    "¡i¤£±o¥N²z¡j" + vbCrLf + _
                    oContext
         oMailCount = mCP(13) & ";" & PUB_GetFCPProSup(mCP(13), True)
'         PUB_SendMail strUserNum, oMailCount, "", IIf("-" + m_Pa(3) + "-" + m_Pa(4) = "-0-00", m_Pa(1) + "-" + m_Pa(2), m_Pa(1) + "-" + m_Pa(2) + "-" + m_Pa(3) + "-" + m_Pa(4)) & _
'            " ¤w½T»{Äò¦æ¦¬¤å¡A½Ðª`·N¸Ó" & strTemp & "½s¸¹¤w³]¬°¤£±o¥N²z¡C", oContext
         'Modify By Sindy 2022/9/29
         'Modify By Sindy 2023/3/27 +,mc13
         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & IIf("-" + m_Pa(3) + "-" + m_Pa(4) = "-0-00", m_Pa(1) + "-" + m_Pa(2), m_Pa(1) + "-" + m_Pa(2) + "-" + m_Pa(3) + "-" + m_Pa(4)) & _
            " ¤w½T»{Äò¦æ¦¬¤å¡A½Ðª`·N¸Ó" & strTemp & "½s¸¹¤w³]¬°¤£±o¥N²z¡C(¤å¸¹:" & mCP(9) & ")','" & oContext & "',null,'" & mCP(9) & "')"
         cnnConnection.Execute mStrSql
         '2022/9/29 END
      End If
      '2021/2/1 END
      
      'add by nickc 2007/05/16 ¥[¤J­Y¬O¥»©Ò´Á­­¤p©óµ¥©ó·í¤Ñ¡A­nµomail  ³qª¾
      Dim oContext2 As String
      oContext = "": oContext2 = ""
      oContext = "¥»©Ò®×¸¹¡G " + m_Pa(1) + "-" + m_Pa(2) + "-" + m_Pa(3) + "-" + m_Pa(4) + vbCrLf + "®×¥ó¦WºÙ¡G " + m_Pa(5) + vbCrLf + "¦¬¤å¤é¡G " + ChangeTStringToTDateString(TransDate(mCP(5), 1)) + vbCrLf + "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf
      'add by nickc 2007/05/16 ¥[¤J­Y¬O¥»©Ò´Á­­¤p©óµ¥©ó·í¤Ñ¡A­nµomail  ³qª¾
      oContext2 = "¥»©Ò®×¸¹¡G " + m_Pa(1) + "-" + m_Pa(2) + "-" + m_Pa(3) + "-" + m_Pa(4) + vbCrLf + "®×¥ó¦WºÙ¡G " + m_Pa(5) + vbCrLf + "¥Ó½Ð°ê®a¡G" + m_Pa(4) + " " + m_Na01Name + vbCrLf + "¦¬¤å¤é¡G " + ChangeTStringToTDateString(TransDate(mCP(5), 1)) + vbCrLf + "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf
      
      'Modify By Sindy 2024/11/6 §ï¦¨¦@¥Î¨ç¼Æ: ¦¬¤å®É,ÀË¬d¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¬O§_¦³»~
      '§ï¼g­ì¥Ñ¬O¦]¬°¥Ó½Ð¤H1~5 ³v¤@ÀË¬d,¦³»~§¡­nµo mail
      'edit by nickc 2007/08/21 ­Y¥Ó½Ð¤H¥þªÅ¥Õ¡A¤£µo
      If Not (m_Pa(26) = "" And m_Pa(27) = "" And m_Pa(28) = "" And m_Pa(29) = "" And m_Pa(30) = "") Then
         'Modify By Sindy 2024/11/21 ¦Û°Ê¦¬¤åªº®×¥ó©Ê½è¶¶§Ç=1 ©Î¯È¥»¦¬¤å¥¼«ü©w
         If m_intCRC = 1 Or m_intCRC = 0 Then
         '2024/11/21 END
            Call RecvChkApplCust("¥Ó½Ð¤H1", m_Pa(26), mCP(13), Trim(m_Pa(75)), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
            Call RecvChkApplCust("¥Ó½Ð¤H2", m_Pa(27), mCP(13), Trim(m_Pa(75)), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
            Call RecvChkApplCust("¥Ó½Ð¤H3", m_Pa(28), mCP(13), Trim(m_Pa(75)), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
            Call RecvChkApplCust("¥Ó½Ð¤H4", m_Pa(29), mCP(13), Trim(m_Pa(75)), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
            Call RecvChkApplCust("¥Ó½Ð¤H5", m_Pa(30), mCP(13), Trim(m_Pa(75)), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
         End If
      End If
      '2024/11/6 END
'Modify By Sindy 2024/11/6 mark
'      'add by nick 2004/10/15  ·í¦¬¤å·~°È°Ï»P«È¤áÀÉ·~°È°Ï¤£¦P®Éµo mail  ¤Î´£¥Ü
'      Dim oStrCuSales1 As String
'      Dim oStrCuSales2 As String
'      Dim oStrCuSales3 As String
'      Dim oStrCuSales4 As String
'      Dim oStrCuSales5 As String
'      'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'      Dim IsMail  As Boolean
'      IsMail = True
'
'      oStrCuSales1 = ""
'      oStrCuSales2 = ""
'      oStrCuSales3 = ""
'      oStrCuSales4 = ""
'      oStrCuSales5 = ""
'      oMailCount = ""
'      'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'      'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'      'Modify By Sindy 2023/2/2 +, , oStrCuSales1 : ¦^¶Ç­ì´¼Åv¤H­û
'      If ChkSameCuArea(m_Pa(26), mCP(13), , , , , m_Pa(75), , oStrCuSales1) = False And mCP(13) <> "" And m_Pa(26) <> "" Then
'         'Add By Sindy 2009/10/19
'         'Modify By Sindy 2023/2/2
'         'If Left(Trim(mCP(12)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_Pa(26)), oStrCuSales1)), 1) = "F" Then
'         If Left(Trim(mCP(12)), 1) = "F" And Left(GetSalesArea(oStrCuSales1), 1) = "F" Then
'         '2023/2/2 END
'            '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'         Else
'            oMailCount = oMailCount & oStrCuSales1 & ";"
'            'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'            If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales1), 1) = "S" And _
'               InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'               oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'            End If
'            '2023/11/7 END
'            oContext = oContext & vbCrLf + "¥Ó½Ð¤H1¡G " + GetCustomerName(ChangeCustomerL(m_Pa(26))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales1)
'         End If
'      'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'      Else
'           If mCP(13) <> "" And m_Pa(26) <> "" Then
'               IsMail = False
'           End If
'      End If
'      'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_SalesST06 <> "" And m_Pa(26) <> "" And mCP(13) <> "" Then
'         'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'         'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'         If PUB_ChkOldCustomer(True, m_Pa(26), mCP(13), m_SalesST15, m_SalesST06, _
'                  IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'             IsMail = False
'         End If
'      End If
'
'      'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'      'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'      'Modify By Sindy 2023/2/2 +, , oStrCuSales2 : ¦^¶Ç­ì´¼Åv¤H­û
'      If ChkSameCuArea(m_Pa(27), mCP(13), , , , , m_Pa(75), , oStrCuSales2) = False And mCP(13) <> "" And m_Pa(27) <> "" Then
'         'Add By Sindy 2009/10/19
'         'Modify By Sindy 2023/2/2
'         'If Left(Trim(mCP(12)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL((m_Pa(27))), oStrCuSales2)), 1) = "F" Then
'         If Left(Trim(mCP(12)), 1) = "F" And Left(GetSalesArea(oStrCuSales2), 1) = "F" Then
'         '2023/2/2 END
'            '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'         Else
'            oMailCount = oMailCount & oStrCuSales2 & ";"
'            'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'            If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales2), 1) = "S" And _
'               InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'               oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'            End If
'            '2023/11/7 END
'            oContext = oContext & vbCrLf + "¥Ó½Ð¤H2¡G " + GetCustomerName(ChangeCustomerL((m_Pa(27)))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales2)
'         End If
'      'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'      Else
'           If mCP(13) <> "" And m_Pa(27) <> "" Then
'               IsMail = False
'           End If
'      End If
'      'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_SalesST06 <> "" And m_Pa(27) <> "" And mCP(13) <> "" Then
'         'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'         'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'         If PUB_ChkOldCustomer(True, m_Pa(27), mCP(13), m_SalesST15, m_SalesST06, _
'                  IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'             IsMail = False
'         End If
'      End If
'
'      'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'      'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'      'Modify By Sindy 2023/2/2 +, , oStrCuSales3 : ¦^¶Ç­ì´¼Åv¤H­û
'      If ChkSameCuArea(m_Pa(28), mCP(13), , , , , m_Pa(75), , oStrCuSales3) = False And mCP(13) <> "" And m_Pa(28) <> "" Then
'         'Add By Sindy 2009/10/19
'         'Modify By Sindy 2023/2/2
'         'If Left(Trim(mCP(12)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_Pa(28)), oStrCuSales3)), 1) = "F" Then
'         If Left(Trim(mCP(12)), 1) = "F" And Left(GetSalesArea(oStrCuSales3), 1) = "F" Then
'         '2023/2/2 END
'            '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'         Else
'            oMailCount = oMailCount & oStrCuSales3 & ";"
'            'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'            If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales3), 1) = "S" And _
'               InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'               oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'            End If
'            '2023/11/7 END
'            oContext = oContext & vbCrLf + "¥Ó½Ð¤H3¡G " + GetCustomerName(ChangeCustomerL(m_Pa(28))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales3)
'         End If
'      'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'      Else
'           If mCP(13) <> "" And m_Pa(28) <> "" Then
'               IsMail = False
'           End If
'      End If
'      'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_SalesST06 <> "" And m_Pa(28) <> "" And mCP(13) <> "" Then
'         'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'         'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'         If PUB_ChkOldCustomer(True, m_Pa(28), mCP(13), m_SalesST15, m_SalesST06, _
'                  IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'             IsMail = False
'         End If
'      End If
'
'      'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'      'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'      'Modify By Sindy 2023/2/2 +, , oStrCuSales4 : ¦^¶Ç­ì´¼Åv¤H­û
'      If ChkSameCuArea(m_Pa(29), mCP(13), , , , , m_Pa(75), , oStrCuSales4) = False And mCP(13) <> "" And m_Pa(29) <> "" Then
'         'Add By Sindy 2009/10/19
'         'Modify By Sindy 2023/2/2
'         'If Left(Trim(mCP(12)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_Pa(29)), oStrCuSales4)), 1) = "F" Then
'         If Left(Trim(mCP(12)), 1) = "F" And Left(GetSalesArea(oStrCuSales4), 1) = "F" Then
'         '2023/2/2 END
'            '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'         Else
'            oMailCount = oMailCount & oStrCuSales4 & ";"
'            'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'            If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales4), 1) = "S" And _
'               InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'               oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'            End If
'            '2023/11/7 END
'            oContext = oContext & vbCrLf + "¥Ó½Ð¤H4¡G " + GetCustomerName(ChangeCustomerL(m_Pa(29))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales4)
'         End If
'      'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'      Else
'           If mCP(13) <> "" And m_Pa(29) <> "" Then
'               IsMail = False
'           End If
'      End If
'      'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_SalesST06 <> "" And m_Pa(29) <> "" And mCP(13) <> "" Then
'         'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'         'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'         If PUB_ChkOldCustomer(True, m_Pa(29), mCP(13), m_SalesST15, m_SalesST06, _
'                  IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'             IsMail = False
'         End If
'      End If
'
'      'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'      'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'      'Modify By Sindy 2023/2/2 +, , oStrCuSales5 : ¦^¶Ç­ì´¼Åv¤H­û
'      If ChkSameCuArea(m_Pa(30), mCP(13), , , , , m_Pa(75), , oStrCuSales5) = False And mCP(13) <> "" And m_Pa(30) <> "" Then
'         'Add By Sindy 2009/10/19
'         'Modify By Sindy 2023/2/2
'         'If Left(Trim(mCP(12)), 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_Pa(30)), oStrCuSales5)), 1) = "F" Then
'         If Left(Trim(mCP(12)), 1) = "F" And Left(GetSalesArea(oStrCuSales5), 1) = "F" Then
'         '2023/2/2 END
'            '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'         Else
'            oMailCount = oMailCount & oStrCuSales5 & ";"
'            'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'            If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales5), 1) = "S" And _
'               InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'               oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'            End If
'            '2023/11/7 END
'            oContext = oContext & vbCrLf + "¥Ó½Ð¤H5¡G " + GetCustomerName(ChangeCustomerL(m_Pa(30))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales5)
'         End If
'      'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'      Else
'           If mCP(13) <> "" And m_Pa(30) <> "" Then
'               IsMail = False
'           End If
'      End If
'      'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_SalesST06 <> "" And m_Pa(30) <> "" And mCP(13) <> "" Then
'         'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'         'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'         If PUB_ChkOldCustomer(True, m_Pa(30), mCP(13), m_SalesST15, m_SalesST06, _
'                  IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'             IsMail = False
'         End If
'      End If
'
'      'edit by nickc 2007/08/21 ­Y¥Ó½Ð¤H¥þªÅ¥Õ¡A¤£µo
'      If IsMail = False Or (m_Pa(26) = "" And m_Pa(27) = "" And m_Pa(28) = "" And m_Pa(29) = "" And m_Pa(30) = "") Then
'           oMailCount = ""
'      End If
'
'      '2006/8/2 MODIFY BY SONIA ¥u§PÂ_1½X,¦]¬°FG
'      If UCase(Mid(m_Pa(1), 1, 1)) <> "F" And oMailCount <> "" Then
'         'Modify By Sindy 2010/11/26 ¥Ó½Ð¤H1~5¬° X65299 ©Î X03072 ªº©Ò¦³Ãö«Y¥ø·~³£¤£ÀË¬d·~°È°Ï
'         If Left(m_Pa(26), 6) <> "X65299" And Left(m_Pa(26), 6) <> "X03072" And _
'            Left(m_Pa(27), 6) <> "X65299" And Left(m_Pa(27), 6) <> "X03072" And _
'            Left(m_Pa(28), 6) <> "X65299" And Left(m_Pa(28), 6) <> "X03072" And _
'            Left(m_Pa(29), 6) <> "X65299" And Left(m_Pa(29), 6) <> "X03072" And _
'            Left(m_Pa(30), 6) <> "X65299" And Left(m_Pa(30), 6) <> "X03072" Then
'            'Modify By Sindy 2022/10/14
'            If UCase(pFormName) <> UCase("frm090801_New") Then
'            '2022/9/27 END
'               MsgBox "¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¤£¦P·~°È°Ï¡A·Ç³Æµo mail ¡I", , "ª`·N¡I"
'            End If
'            'edit by nickc 2005/08/10 ¥[µo¨q¬Â
'            'Modify By Sindy 2022/9/29 §ï§ì Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'            oMailCount = oMailCount & mCP(13) & ";" & Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'            oContext = oContext & vbCrLf + "¦¬¤å´¼Åv¤H­û¡G " + GetStaffName(mCP(13)) + vbCrLf + vbCrLf + "´¼Åv¤H­û(°Ï)¤£¦P¡I"
'            'Modify by Morgan 2006/6/26 ¦¬¤å¸¹©Î¤º®e¤@©w­n¦³,§_«h¤£·|±H
''            PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å³qª¾--¦¹®×¦¬¤å«D­ì´¼Åv¤H­û(°Ï)¡I", oContext
'            'Modify By Sindy 2022/9/29
'            'Modify By Sindy 2023/3/27 +,mc13
'            mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
'               " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'               ",'" & "®×¥ó¦¬¤å³qª¾--¦¹®×¦¬¤å«D­ì´¼Åv¤H­û(°Ï)¡I(¤å¸¹:" & mCP(9) & ")','" & oContext & "',null,'" & mCP(9) & "')"
'            cnnConnection.Execute mStrSql
'            '2022/9/29 END
'         End If
'      End If
'2024/11/6 mark END

      'Modify By Sindy 2022/12/28
      If UCase(pFormName) = UCase("frm090801_New") Then
         If mCP(31) = "" And mCP(1) = "P" And (mCP(10) = "601" Or mCP(10) = "605") Then
            '¦¬¤å¤é>ªk­­,ªk­­=ªk­­+¥b¦~
            'Modified by Morgan 2015/7/1 Ex.P-089468 6/30
            'If Val(cp07) <= Val(cp05) Then
            'Add By Sindy 2023/3/21
            If Val(mCP(7)) = 0 Then
               mCP(14) = ""
            Else
            '2023/3/21 END
               If Val(mCP(7)) < Val(mCP(5)) Then
                  If Val(PUB_GetWorkDay1(mCP(7), False)) < Val(mCP(5)) Then 'ªk­­¬°°²¤é(111/8/20)®É­n¥Î¤U¤@¤u§@¤Ñ(111/8/22)§PÂ_ ex:P-119436
                     mCP(7) = CompDate(1, 6, mCP(7)) 'ªk­­¦Û°Ê©µ¥b¦~
                     strSql = "update caseprogress set cp07=" & mCP(7) & " where cp09='" & mCP(9) & "'"
                     cnnConnection.Execute strSql, intJ
                     'Modify By Sindy 2023/12/6 ex:P-093857 ­pºâ¥X¨Óªºªk©w´Á­­¨S¦³¹O´Á¤~­«·s­pºâ©Ò­­
                     If Val(mCP(7)) >= Val(mCP(5)) Then
                     '2023/12/6 END
                        bolUpdCP06IsSysDate = True 'Add By Sindy 2023/6/20
                     Else
                        mCP(14) = "" 'Add By Sindy 2023/12/6
                     End If
                  End If
               End If
            End If
            
            '¦¬´Ú«á°e¥ó
            'Removed by Morgan 2024/2/22 ¤w¦¬´Ú³qª¾¤w§ï¬°³qª¾¯S®í³]©w¤H­û
            'If mCP(141) = "2" Then
            '   strSql = "insert into UndeliveredRec(UD01,UD02,UD03,UD04) VALUES('" & mCP(9) & "'," & strSrvDate(1) & ",'1','" & mCP(14) & "')"
            '   cnnConnection.Execute strSql, intj
            'End If
            'end 2024/2/22
            
            '¥xÆW»âÃÒ
            If m_Pa(9) = "000" And mCP(10) = "601" Then
               PUB_Add401ForGuiCase mCP(1), mCP(2), mCP(3), mCP(4), mCP(9), mCP(12), mCP(13)
            End If
            
            '¤w¦³©Ó¿ì¤H:
            If mCP(14) <> "" Then
               mCP(157) = strSrvDate(1) '***
               'Modified by Lydia 2025/06/18 P¤ÎCFP¤À®×®É¡A¹w³]N¤£­p¥óºÞ¨î¡G©Ó¿ì¤H¬°¤º±Mµ{§ÇP12³£¹w³]CP26=N
               strSql = "update caseprogress set cp14='" & mCP(14) & "',cp157=" & mCP(157) & IIf(PUB_GetST03(mCP(14)) = "P12", ",cp26='N'", "") & " where cp09='" & mCP(9) & "'"
               cnnConnection.Execute strSql, intJ
            End If
         End If
      'Added by Lydia 2023/04/13 FMP®×¦¬¤å(601)»âÃÒ©Î(605)Ãº¦~¶O®É¡A«D¾ÈµØ®×½Ð¹w³]©Ó¿ì¤H
      Else
         If mCP(31) = "" And m_bolFMP = True And (mCP(10) = "601" Or mCP(10) = "605") Then
            '¤w¦³©Ó¿ì¤H:
            If mCP(14) <> "" Then
               If PUB_GetFMP2toP(mCP(1), mCP(2), mCP(3), mCP(4)) = False Then 'Added by by Lydia 2025/06/19 §PÂ_«D¾ÈµØ®×
                  mCP(157) = strSrvDate(1) '***
                  'Modified by Lydia 2025/06/18 P¤ÎCFP¤À®×®É¡A¹w³]N¤£­p¥óºÞ¨î¡G©Ó¿ì¤H¬°¤º±Mµ{§ÇP12³£¹w³]CP26=N
                  strSql = "update caseprogress set cp14='" & mCP(14) & "',cp157=" & mCP(157) & IIf(PUB_GetST03(mCP(14)) = "P12", ",cp26='N'", "") & " where cp09='" & mCP(9) & "'"
                  cnnConnection.Execute strSql, intJ
               End If
            End If
         End If
      'end 2023/04/13
      End If
      '2022/12/28 END
      
'Removed by Morgan 2025/8/4 ²¾¨ìPUB_AutoRecvCRL_P(¦]¬°¦³¨Ò¥~)
'      'Added by Morgan 2025/6/18
'      'P/CFP ³]©w¬°µ{§Ç©Ó¿ì¥B¤£»Ý±M·~³¡¥DºÞ¤À®×ªº©Ê½è¡A©Ó¿ì¤H³£¦Û°Ê¹w³]¬°µ{§Ç¤H­û¡A­Y¦³»Ý­n¡A¤À®×¤H­û¦A¦Û¦æ­×§ï¡A¦ýCFP¹êÅé¼f¬d°£¥~--³¢
'      'µ{§Ç©Ó¿ìªº©Ê½è­n¦b¦¹¥ý³]©w¡A§_«h­Y±µ¬¢³æ¦P®É¦³¤uµ{®vªº®×¥ó©Ê½è·|¦b¥DºÞ¤À®×®É³Q¤@¨Ö³]©w¡A¦ý¹ê¼f³W«h¯S§O°£¥~¡C
'      If mCP(14) = "" And (mCP(1) = "P" Or mCP(1) = "CFP") And m_bolFMP2 = False And mCP(10) <> "416" Then
'         'Modified by Lydia 2025/06/19 §ï¼Ò²Õ¦WºÙ
'         'If PUB_GetCPM35byCP10(mCP(1), mCP(10)) = "2" Then
'         If PUB_GetCPMbyCP10(mCP(1), mCP(10), "cpm35") = "2" Then
'            If mCP(1) = "CFP" Then
'               mCP(14) = PUB_GetCFPHandler(mCP(1) & mCP(2) & mCP(3) & mCP(4))
'            Else
'               mCP(14) = PUB_GetPHandler(mCP(1) & mCP(2) & mCP(3) & mCP(4))
'            End If
'            '¥u§ó·s©Ó¿ì¤H
'            strSql = "update caseprogress set cp14='" & mCP(14) & "' where cp09='" & mCP(9) & "'"
'            cnnConnection.Execute strSql, intJ
'         End If
'      End If
'      'end 2025/6/18
'end 2025/8/4
      
      'add by nickc 2007/05/16 ¥[¤J­Y¬O¥»©Ò´Á­­¤p©óµ¥©ó·í¤Ñ¡A­nµomail  ³qª¾
      oMailCount = ""
      'Added by Lydia 2015/12/30 FMP¾ÈµØ®×µoµ¹FCP¤H­û
      'If m_PA(1) = "P" Or m_PA(1) = "PS" Then
      If PUB_FMPtoCheck(1, 2, "", m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4)) = True Or InStr(mChkVal, "¾ÈµØ®×¥ó½T»{") > 0 Then
           oMailCount = Pub_GetSpecMan("C")
      ElseIf m_Pa(1) = "P" Or m_Pa(1) = "PS" Then
      'end 2015/12/30
           oMailCount = Pub_GetSpecMan("A")
      ElseIf m_Pa(1) = "CFP" Or m_Pa(1) = "CPS" Then
           oMailCount = Pub_GetSpecMan("B")
      ElseIf m_Pa(1) = "FCP" Or m_Pa(1) = "FG" Then
           oMailCount = Pub_GetSpecMan("C")
      End If
      strDivisionalEmp = oMailCount 'Add By Sindy 2022/9/30 °O¿ý¤À®×¤H­û,«áÄò·|¥Î¨ì
      'Add By Sindy 2023/1/11 ­Y¤w¦³¤À©Ó¿ì¤H,¦P®É¤@¨Ö³qª¾
      If mCP(14) <> "" Then
         If oMailCount <> "" Then oMailCount = oMailCount & ";"
         oMailCount = oMailCount & mCP(14)
      End If
      '2023/1/11 END
      
      If DBDATE(mCP(6)) < strSrvDate(1) And Trim(mCP(6)) <> "" And Trim(oMailCount) <> "" Then
         '2007/8/13 MODIFY BY SONIA ¥[´¼Åv¤H­û
         'Modify By Sindy 2010/12/16 ¥[·~°È°Ï,¶O¥Î
'         PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¤w¹O¥»©Ò´Á­­¡A½Ð¾¨³t¿ì²z¡I", oContext2 & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(6))) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0")
         'Modify By Sindy 2022/9/29
         'Modified by Lydia 2022/12/23 +chgsql
         'Modify By Sindy 2023/3/27 +,mc13
         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¤w¹O¥»©Ò´Á­­¡A½Ð¾¨³t¿ì²z¡I(¤å¸¹:" & mCP(9) & ")','" & ChgSQL(oContext2) & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(6))) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & "',null,'" & mCP(9) & "')"
         cnnConnection.Execute mStrSql
         '2022/9/29 END
      End If
      If DBDATE(mCP(6)) = strSrvDate(1) And Trim(mCP(6)) <> "" And Trim(oMailCount) <> "" Then
         '2007/8/13 MODIFY BY SONIA ¥[´¼Åv¤H­û
         'Modify By Sindy 2010/12/16 ¥[·~°È°Ï,¶O¥Î
'         PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¤w©¡¥»©Ò´Á­­¡A½Ð¾¨³t¿ì²z¡I", oContext2 & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(6))) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0")
         'Modify By Sindy 2022/9/29
         'Modified by Lydia 2022/12/23 +chgsql
         'Modify By Sindy 2023/3/27 +,mc13
         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¤w©¡¥»©Ò´Á­­¡A½Ð¾¨³t¿ì²z¡I(¤å¸¹:" & mCP(9) & ")','" & ChgSQL(oContext2) & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(6))) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & "',null,'" & mCP(9) & "')"
         cnnConnection.Execute mStrSql
         '2022/9/29 END
      End If
      'Add By Sindy 2023/6/20
      If bolUpdCP06IsSysDate = True Then
         mCP(6) = strSrvDate(1)
         strSql = "update caseprogress set cp06=" & strSrvDate(1) & " where cp09='" & mCP(9) & "'"
         cnnConnection.Execute strSql, intJ
      End If
      '2023/6/20 END
      
      '2010/2/5 ADD BY SONIA ´ú¸Õ°lÂÜ¥Î
      If m_Pa(1) = "FCP" And mCP(10) = "305" Then
         'µoµ¹¨q¬Â
'         PUB_SendMail strUserNum, "83002", "", "FCP¦¬¤å§ï½ÐÁp¦X¡A½Ð°lÂÜ¥Ó½Ð®×¸¹ÅÜ¤Æ±¡§Î¡ICP30¬°¦ó·|»P·s¥Ó½Ð®×¸¹¬Û¦P?", oContext2 & vbCrLf
         'Modify By Sindy 2022/9/29
         'Modified by Lydia 2022/12/23 +chgsql
         'Modify By Sindy 2023/3/27 +,mc13
         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            " values( '" & strUserNum & "','" & Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & "FCP¦¬¤å§ï½ÐÁp¦X¡A½Ð°lÂÜ¥Ó½Ð®×¸¹ÅÜ¤Æ±¡§Î¡ICP30¬°¦ó·|»P·s¥Ó½Ð®×¸¹¬Û¦P?" & "','" & ChgSQL(oContext2) & "',null,'" & mCP(9) & "')"
         cnnConnection.Execute mStrSql
         '2022/9/29 END
      End If
      '2010/2/5 END
   End If

   'Added by Lydia 2018/09/06 Âd»OP®×(FMP)¤¤¶¡µ{§Ç©M¤¤¶¡±µ¶i¨Ó®×¥óªº¦¬¤å¡A¨t²Î¦Û°Êµoe-mail³qª¾¡C
   If InStr(mChkVal, "FMP¤¤¶¡µ{§ÇEMAIL³qª¾") > 0 Then
'      PUB_SendMail strUserNum, mCP(13), "", m_Pa(1) & "-" & m_Pa(2) & IIf(Val(m_Pa(3) & m_Pa(4)) > 0, "-" & m_Pa(3) & "-" & m_Pa(4), "") & " ¤w¦¬¤å" & m_CP10Name & " , ½Ð¶i¨÷©v°Ï²¾ÀÉ!", "¦P¥D¦®"
      'Modify By Sindy 2022/9/29
      'Modify By Sindy 2023/3/27 +,mc13
      'Added by Lydia 2024/05/29 FMP®×¡]«D¾ÈµØ®×¡^­Y¤£¨«©R¦W¬yµ{¡A¦bÂdÂi·s®×¦¬¤å®É¡A·|µo«H³qª¾¦¬¤å©Ó¿ì¤H­û¡]¦pªþÀÉ¡^¡A½Ð¤@¨ÖCCµ¹¥~±Mµ{§ÇºÞ¨î¤H­û¤Î¨ä¥DºÞ ---Phoebe
      strTmp1(0) = PUB_GetFCPHandler(m_Pa(1), m_Pa(2), m_Pa(3), m_Pa(4))
      If strTmp1(0) <> "" Then
         strTmp1(1) = PUB_GetFCPProSup(strTmp1(0))
         If strTmp1(1) <> "" Then strTmp1(0) = strTmp1(0) & ";" & strTmp1(1)
      End If
      'Modified by Lydia 2024/05/29 mc09=null¶Ç¤JstrTmp1(0)
      mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
         " values( '" & strUserNum & "','" & mCP(13) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
         ",'" & m_Pa(1) & "-" & m_Pa(2) & IIf(Val(m_Pa(3) & m_Pa(4)) > 0, "-" & m_Pa(3) & "-" & m_Pa(4), "") & " ¤w¦¬¤å" & m_CP10Name & " , ½Ð¶i¨÷©v°Ï²¾ÀÉ!" & "','¦P¥D¦®','" & strTmp1(0) & "','" & mCP(9) & "')"
      cnnConnection.Execute mStrSql
      '2022/9/29 END
   End If
   'Added by Lydia 2023/08/14 §Q¯q½Ä¬ð®×¥ó¡G­­¾\®×¥ó¸É¥R±±ºÞ: ³Q­­¾\¤H­û¦b¦¬¤å®É, ¨t²ÎÀ³¦P®Éµo³qª¾µ¹¦¬¤åªÌ+¨ä¥DºÞ
   Call ChkCufaRight(mCP(13), m_Pa(1) & m_Pa(2) & m_Pa(3) & m_Pa(4), m_Pa(26) & "," & m_Pa(27) & "," & m_Pa(28) & "," & m_Pa(29) & "," & m_Pa(30), m_Pa(75))
     
   'Add By Sindy 2022/9/27
   If UCase(pFormName) = UCase("frm090801_New") Then
      If ERecvSaveProgress(mCP, m_Pa, strDivisionalEmp, oContext2) = False Then
         GoTo ErrHand
      End If
   End If
   
   Exit Function
   
ErrHand:
   PUB_SaveFrm010005 = False 'Add By Sindy 2022/10/25
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical, "PUB_SaveFrm010005"
   End If
End Function

'Add By Sindy 2024/11/6 §ï¦¨¦@¥Î¨ç¼Æ: ¦¬¤å®É,ÀË¬d¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¬O§_¦³»~
'strApplTitle: ¥Ó½Ð¤H1/·í¨Æ¤H1
'strApplNo: ¥Ó½Ð¤H
'strFCNo: FC¥N²z¤H
'oContext: ¶l¥ó°òÂ¦¤º®e
'm_LOS04_1: LOS®×·½¦¬¤å - ¤¶²Ð¤H(²Ä¤@¦ì)
Private Sub RecvChkApplCust(ByVal strApplTitle As String, ByVal strApplNo As String, _
   ByVal strCP13 As String, ByVal strFCNo As String, ByVal m_SalesST15 As String, ByVal strCP12 As String, _
   ByVal oContext As String, ByVal m_SalesST06 As String, ByVal pFormName As String, _
   ByVal strCP01 As String, ByVal strCP02 As String, ByVal strCP03 As String, ByVal strCP04 As String, _
   ByVal strCP09 As String, Optional m_LOS04_1 As String = "")
   
Dim oStrCuSales1 As String
Dim oMailCount As String
   
   If strApplNo = "" Then Exit Sub
   oMailCount = ""
   
   'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
   'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
   'Modify By Sindy 2023/2/2 +, , oStrCuSales1 : ¦^¶Ç­ì´¼Åv¤H­û
   If ChkSameCuArea(strApplNo, strCP13, , , , , strFCNo, , oStrCuSales1) = False And strCP13 <> "" And strApplNo <> "" Then
      'Add By Sindy 2009/10/19
      'Modified by Lydia 2019/02/14
      'Modify By Sindy 2023/2/2
      'If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(strApplNo, oStrCuSales1)), 1) = "F" Then
      If Left(m_SalesST15, 1) = "F" And Left(GetSalesArea(oStrCuSales1), 1) = "F" Then
      '2023/2/2 END
         '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
      Else
         oMailCount = oMailCount & oStrCuSales1 & ";"
         'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
         If Left(strCP12, 1) <> "S" And Left(PUB_GetST03(oStrCuSales1), 1) = "S" And _
            InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
            oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
         End If
         '2023/11/7 END
         oContext = oContext & vbCrLf + strApplTitle & "¡G " + GetCustomerName(strApplNo) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales1)
      End If
   End If
   'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
   If m_SalesST06 <> "" And strApplNo <> "" And strCP13 <> "" Then
       'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
       'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
       If PUB_ChkOldCustomer(True, strApplNo, strCP13, m_SalesST15, m_SalesST06, _
               IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), strCP01 & strCP02 & strCP03 & strCP04) = True Then
           oMailCount = ""
       End If
   End If
   
   If UCase(Mid(strCP01, 1, 1)) <> "F" And oMailCount <> "" Then
      'Modify By Sindy 2010/11/26 ¥Ó½Ð¤H1~5¬° X65299 ©Î X03072 ªº©Ò¦³Ãö«Y¥ø·~³£¤£ÀË¬d·~°È°Ï
      If Left(strApplNo, 6) <> "X65299" And Left(strApplNo, 6) <> "X03072" Then
         
         'LOS®×·½¦¬¤å
         If m_LOS04_1 <> "" Then
            'Modify By Sindy 2022/10/14
            If UCase(pFormName) <> UCase("frm090801_New") Then
            '2022/9/27 END
               MsgBox "®×·½¤¶²Ð¤H­û»P«È¤á´¼Åv¤H­û¤£¦P·~°È°Ï¡A·Ç³Æµo mail ¡I", , "ª`·N¡I"
            End If
            'Modified by Lydia 2022/07/15 ³qª¾ªk«ß©Òªº´¼Åv¤H­û¨S¦³·N¸q¡AÀ³¸Ó­n§ï¬°®×·½¤¶²Ð¤H­û. ex.L-006547
            oMailCount = oMailCount & m_LOS04_1 & ";" & Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
            oContext = oContext & vbCrLf + "®×·½¤¶²Ð¤H­û¡G " + GetStaffName(m_LOS04_1) + vbCrLf + vbCrLf + "´¼Åv¤H­û(°Ï)¤£¦P¡I"
         
         Else
            'Modify By Sindy 2022/9/27
            If UCase(pFormName) <> UCase("frm090801_New") Then
            '2022/9/27 END
               MsgBox "¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¤£¦P·~°È°Ï¡A·Ç³Æµo mail ¡I", , "ª`·N¡I"
            End If
            'Modify By Sindy 2022/9/29 §ï§ì Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
            oMailCount = oMailCount & strCP13 & ";" & Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
            
            'ªk«ß©Ò¦Û¦æ¦¬¤å´¼¼z©Ò¤H­û«È¤á®É¡A¦b§PÂ_¸ó°Ï¦¬¤åµoEMAILµ¹Âù¤è´¼Åv¤H­û®É¡A¥[µoªLÁ`¡C
            If InStr(strCP01, "L") > 0 Then 'And m_LOS04_1 = ""
               oMailCount = oMailCount & PUB_ChkForLawMan(strApplNo, strCP01, strCP02, strCP03, strCP04)
            End If
            
            oContext = oContext & vbCrLf + "¦¬¤å´¼Åv¤H­û¡G " + GetStaffName(strCP13) + vbCrLf + vbCrLf + "´¼Åv¤H­û(°Ï)¤£¦P¡I"
         End If
         
'         PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å³qª¾--¦¹®×¦¬¤å«D­ì´¼Åv¤H­û(°Ï)¡I", oContext
         'Modify By Sindy 2022/9/29
         'Modified by Lydia 2022/12/23 +chgsql
         'Modify By Sindy 2023/3/27 +,mc13
         'Modify By Sindy 2024/11/21 + strApplTitle ¬°¤FÅý¥D¦®¤£¦P,­Y¦h¤H¥Ó½Ð¬í¼Æ¤S¬Û¦P®É,·|³y¦¨·s¼W­«ÂÐ¦Ó¥X¿ù
         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & "®×¥ó¦¬¤å³qª¾--¦¹®×¦¬¤å«D­ì´¼Åv¤H­û(°Ï)¡I(¤å¸¹:" & strCP09 & strApplTitle & ")','" & ChgSQL(oContext) & "',null,'" & strCP09 & "')"
         cnnConnection.Execute mStrSql
         '2022/9/29 END
      End If
   End If
End Sub

'Added by Lydia 2022/09/05 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G±M§Q¦¬¤å-§ó·sµe­±¤¤¶O¥Î¤Î³W¶OªºÄæ¦ì¤º®e(±qfrm010005.OnUpdateFee©â¥X¨Ó)
Public Function PUB_Frm010005OnUpdFee(ByVal tPA01 As String, ByVal tPA02 As String, ByVal tPA03 As String, ByVal tPA04 As String, ByVal tPA09 As String, ByVal tPA08 As String, ByVal tPA16 As String, ByVal tPA14 As String, _
                       ByVal tCP07 As String, ByVal tCP10 As String, ByVal tCP13 As String, ByVal tCP118 As String, ByVal tAppNo As String, ByRef nowCP16 As String, ByRef nowCP17 As String, ByRef nowCP18 As String) As Boolean
'tApplNo : ¥Ó½Ð¤H1~5
'PUB_Frm010005OnUpdFee : °ê¥~³¡¹w³]¶O¥Î©M³W¶O
Dim m_ST15 As String, m_ST06 As String
Dim strSG07 As String, strSG08 As String 'Add By Sindy 2012/11/22
Dim strTmpVal As String 'Added by Lydia 2018/03/01
  
   PUB_Frm010005OnUpdFee = False
   m_ST15 = GetST15(tCP13, , , m_ST06)
   
   'Added by Lydia 2018/12/06 ¹w³]¥ý²MªÅ¶O¥Îµ¥ÅÜ¼Æ, ¦]¬°¦P®É¦¬¤å­Y¹J¨ìµL¶O¥Îªº©Ê½è,©Ò¥H¥¼²MªÅ¤W¤@µ§¦¬¤åªº¼Æ­È(ex.P-121437)
   nowCP16 = ""  '¶O¥Î CP16
   nowCP17 = ""  '³W¶O CP17
   nowCP18 = ""  'ÂI¼Æ CP18
   'end 2018/12/06
   
   'edit by nickc 2008/05/30 ³¢ ½Ð§@³æ X14843050 ¤£ºÞ
   'modify by sonia 2013/11/19 ¥[X3928904,X69514 ¸­¸g²z
   'modify by sonia 2014/9/11 ¨ú®øX69514,¤wÂà¥~±M
   'Modified by Lydia 2020/01/13 X38120030¹ç»·¿¤ºÓ¹ç¹q¤l¦³­­¤½¥q,¨ä¤j³°¦~¶O605¤£­n±±¨îª÷ÃB
   If PUB_CheckExceptFrm010005(tPA01, tPA09, tCP10, tAppNo) = False Then
   'end 2020/01/13
      'Modified by Morgan 2013/2/4 +§PÂ_«D°ê¥~³¡¦¬¤åªº¤£¹w³] Ex.FCP-039919
      If tPA01 = "FCP" And Mid(m_ST15, 1, 1) = "F" Then
         If tPA03 = "" Then tPA03 = "0"
         If tPA04 = "" Then tPA04 = "00"
         '³W¶O
            'Modified by Lydia 2017/03/01 +¬O§_¹q¤l°e¥ó,+§PÂ_
            strTmpVal = GetPatentOfficialFee(tPA01, tCP10, DBDATE(tCP07), tPA08, tPA09, tPA16, tPA14, tPA02, tPA03, tPA04, tCP118)
            nowCP17 = strTmpVal
            'end 2018/03/01
            
            '¶O¥Î
            strTmpVal = Val(GetFCPFee(tPA01, tCP10)) + Val(nowCP17)
            If Val(strTmpVal) > 0 Then
               nowCP16 = strTmpVal
               'ÂI¼Æ
               nowCP18 = Format((Val(nowCP16) - Val(nowCP17)) / 1000, "0.0")
            End If
          PUB_Frm010005OnUpdFee = True
      '2009/10/15 ADD BY SONIA FMP®×¥ó¤]­n¹w³]¶O¥Î,§ìCASEFEE«h¥H'FCP'+¥Ó½Ð°ê®a+®×¥ó©Ê½è§ì
      ElseIf tPA01 = "P" And Mid(m_ST15, 1, 1) = "F" Then
'         'Add By Sindy 2012/11/22 ¨ú±o¯S®í«È¤á/¥N²z¤H¦¬¤å¶O¥Î
            '³W¶O
            nowCP17 = GetFMPOfficialFee("FCP", tCP10, tPA09)
            '¶O¥Î
            strTmpVal = GetFMPFee("FCP", tCP10, tPA09)
            If Val(strTmpVal) > 0 Then
               nowCP16 = strTmpVal
               'ÂI¼Æ
               nowCP18 = Format((Val(nowCP16) - Val(nowCP17)) / 1000, "0.0")
            End If
            PUB_Frm010005OnUpdFee = True
      '2009/10/15 END
      End If
   End If
   
   'Modified by Lydia 2024/05/28 §ï¦¨¼Ò²Õ
   ''Added by Lydia 2020/03/27 FCP-062174¼f©w«e¤£¦¬¶O±±¨î: §PÂ_°ò¥»ÀÉ¤§¥Ø«e­ã/»éPA16¬°ªÅ­È®É¡A¤£ºÞ¥ô¦ó®×¥ó©Ê½è³£¤£¥²¹w³]¦¬¤å¶O¥Î¡B³W¶O¡BÂI¼Æ¡C
   'If tPA16 = "" And InStr("FCP062174000", tPA01 & tPA02 & tPA03 & tPA04) > 0 Then
   '     nowCP16 = ""
   '     nowCP17 = ""
   '     nowCP18 = ""
   'End If
   ''Added by Lydia 2022/05/03 FCP-067004®Ö­ã«e¤£¦¬¶O±±¨î¡G¥Ó½Ð¦Ü®Ö­ã(¼È¤£¥]§t»âÃÒ)¤£¦¬¥ô¦ó¦¬¶O (¥]§t³W¶O¤ÎªA°È¶O¡B­Y«È¤á´£AEP¤]¤£¦¬¶O)
   'If tPA16 <> "1" And InStr("FCP067004000", tPA01 & tPA02 & tPA03 & tPA04) > 0 Then
   '     nowCP16 = ""
   '     nowCP17 = ""
   '     nowCP18 = ""
   'End If
   If PUB_GetCP20forSpec(tPA01, tPA02, tPA03, tPA04, tPA16) = "N" Then
        nowCP16 = ""
        nowCP17 = ""
        nowCP18 = ""
   End If
   'end 2024/05/28
End Function

'Added by Lydia 2022/08/19 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G­ì¥»frm010005.CheckExcept
Public Function PUB_CheckExceptFrm010005(ByVal tCP01 As String, ByVal tNA01 As String, ByVal tCP10 As String, ByVal tApplNo As String) As Boolean
'tCP01, tNa01 ,tCP10: ¨t²Î§O, ¥Ó½Ð°ê®a, ®×¥ó©Ê½è
'tApplNo : ¥Ó½Ð¤H1~5
     
     PUB_CheckExceptFrm010005 = False
    'modify by sonia 2013/11/19 ¥[X3928904,X69514 ¸­¸g²z
    'modify by sonia 2014/9/11 ¨ú®øX69514,¤wÂà¥~±M
    'Modified by Lydia 2020/01/13 X3812003¹ç»·¿¤ºÓ¹ç¹q¤l¦³­­¤½¥q,¨ä¤j³°¦~¶O605¤£­n±±¨îª÷ÃB
    If InStr(tApplNo, "X1484305") = 0 And InStr(tApplNo, "X3928904") = 0 Then
       'Added by Lydia 2020/01/13 X3812003¹ç»·¿¤ºÓ¹ç¹q¤l¦³­­¤½¥qªºP¤j³°®×¦~¶O605¤£­n±±¨îª÷ÃB
       If tCP01 = "P" And tNA01 = "020" And Trim(Left(tCP10, 3)) = "605" And InStr(tApplNo, "X3812003") > 0 Then
            PUB_CheckExceptFrm010005 = True
       End If
    Else
            PUB_CheckExceptFrm010005 = True
    End If
End Function

'Added by Lydia 2022/08/22 Åª¨úµo©ú¤H½s¸¹
Public Function PUB_GetPatentInventorList(ByVal pPA01 As String, ByVal pPA02 As String, ByVal pPA03 As String, ByVal pPA04 As String) As String
Dim intP As Integer, strP1 As String
Dim rsPD As New ADODB.Recordset
    
    PUB_GetPatentInventorList = ""
    strP1 = "select * from patentInventor where pi01=" & CNULL(pPA01) & " and pi02=" & CNULL(pPA02) & " and pi03=" & CNULL(pPA03) & " and pi04=" & CNULL(pPA04) & " order by pi05 asc "
    intP = 1
    Set rsPD = ClsLawReadRstMsg(intP, strP1)
    If intP = 1 Then
        rsPD.MoveFirst
        strP1 = ""
        Do While Not rsPD.EOF
            strP1 = strP1 & rsPD.Fields("pi06") & ","
            rsPD.MoveNext
        Loop
        If Right(strP1, 1) = "," Then strP1 = Mid(strP1, 1, Len(strP1) - 1)
        PUB_GetPatentInventorList = strP1
    End If
    
    Set rsPD = Nothing
End Function

'Added by Lydia 2017/11/14 FCP®×¥ó©R¦W¹q¤l¤Æ¡G¤¤»¡¿é¤J¬ÛÃö³]©w-¦sÀÉ
Public Sub PUB_UpdTCTrecord(ByVal mCnKind As String, ByVal mPtyList As String, ByVal mTCN01 As String, ByRef mRetList As String, ByVal mPA01 As String, ByVal mPA02 As String, ByVal mPA03 As String, ByVal mPA04 As String, ByVal mPA05 As String, ByVal mPA06 As String, _
              ByVal mCP09 As String, ByVal mCP10 As String, ByVal mCP06 As String, ByVal mCP07 As String, ByVal mCP13 As String, ByVal mPA08 As String, ByVal mPA09 As String, ByVal mPA16 As String, ByVal mPA14 As String, ByVal mCustList As String, ByVal mPA75 As String, ByVal mPA150 As String, _
              Optional ByVal mGrp_New As String = "B", Optional ByVal mTCT0203_new As String)
'mCnKind : ¤¤»¡Ãþ«¬
'mPtyList : ¤Ä¿ïªº¦¬¤å©Ê½è
'mCustList : ¥Ó½Ð¤H 1~5
'mGrp_Old,mGrp_New: ©R¦W°O¿ý¤§¤uµ{®v²Õ§O(B¥Nªí¥¼¤À²Õ=°hµ{§Ç)
'mTCT0203_old,mTCT0203_New: ©R¦W°O¿ý¤§Ä¶²¦¤é´Á©M®É¶¡TCT02+TCT03
Dim strPK As String, strANo As String
Dim strPKList As String
Dim m_Cp33 As Double, m_Cp34 As Double 'Added by Lydia 2018/05/07¼Ð·Ç»ù©M©³»ù
Dim m_Cp26 As String 'Added by Lydia 2018/05/10 ¬O§_ºâ®×¥ó
Dim strTF As String 'Added by Lydia 2018/06/28 ¤¤»¡²Ä¤@­Ó¦¬¤å¸¹
Dim bolExistTCT As Boolean, mGrp_old   As String, mTCT0203_old As String
'Added by Lydia 2019/12/23
Dim strReceiver As String '«æ¥óÂ½Ä¶ªº¦¬¥óªÌ
Dim strContent As String '«æ¥óÂ½Ä¶ªºemail¤º®e
Dim strCon1 As String, strCon2 As String
Dim tmpPty As Variant
Dim strNowPty As String
Dim rsQuery As New ADODB.Recordset
Dim m_bolFMP As Boolean
Dim m_TCN03st16 As String 'Added by Lydia 2023/08/11

    mRetList = ""  '¤@¨Ö²£¥Í¤§¦¬¤å¸¹
    
    'Move by Lydia 2023/02/17 ±q¼Ò²Õ³Ì¤U¤è²¾¤W¨Ó
      '¦X¨Ö¼Ò²ÕUpdateTCN01
      If mTCN01 <> "" Then
          'Modified by Lydia 2023/05/22 ¤é¤å²Õ©Ó¿ì¼W¥[¡u«È¤á¦³´£¨Ñ±m¹Ï¡vÄæ¦ì¿é¤J
          'mStrSql = "Update TrackingCaseName set TCN05=" & CNULL(mCP09) & " Where TCN01 ='" & mTCN01 & "' AND TCN05 IS NULL "
          'cnnConnection.Execute mStrSql
          'Modified by Lydia 2023/08/11 §ì´¼Åv¤H­û=ºÞ¨î¤Hªº²Õ§O
          'strTmp1(0) = "select tcn12 from trackingcasename where tcn01='" & mTCN01 & "' AND TCN05 IS NULL"
          strTmp1(0) = "select tcn12,st16 from trackingcasename,staff where tcn01='" & mTCN01 & "' AND TCN05 IS NULL and tcn03=st01(+) "
          intJ = 1
          Set rsQuery = ClsLawReadRstMsg(intJ, strTmp1(0))
          If intJ = 1 Then
             m_TCN03st16 = "" & rsQuery.Fields("st16") 'Added by Lydia 2023/08/11
             If "" & rsQuery.Fields("tcn12") = "Y" Then
                 mStrSql = "Update Patent Set PA63=" & CNULL(rsQuery.Fields("tcn12")) & " WHERE PA01='" & mPA01 & "' AND  PA02='" & mPA02 & "' AND PA03='" & mPA03 & "' AND PA04='" & mPA04 & "' "
                 cnnConnection.Execute mStrSql
             End If
             mStrSql = "Update TrackingCaseName set TCN05=" & CNULL(mCP09) & " Where TCN01 ='" & mTCN01 & "' AND TCN05 IS NULL "
             cnnConnection.Execute mStrSql
          End If
          'end 2023/05/22
      End If
    'end 2023/02/17
    
    strTmp1(0) = "select tct01,tct02,tct03,tct04 from transcasetitle where tct01='" & mCP09 & "' "
    intJ = 1
    Set rsQuery = ClsLawReadRstMsg(intJ, strTmp1(0))
    If intJ = 1 Then
        bolExistTCT = True
        mTCT0203_old = "" & rsQuery.Fields("tct02") & rsQuery.Fields("tct03")
        mGrp_old = "" & rsQuery.Fields("tct04")
        If mGrp_old = "" Then mGrp_old = "B" '°hµ{§Ç
    End If
    'Added by Lydia 2022/10/07 ¬O§_¬°FMP®×¥ó; °Ñ¦Ò¤½§i1110401-04 ÂdÂi¦¬¤å©M¤º³¡¦¬¤åFMP®×¼W¥[­pºâ©Ó¿ì´Á­­
    If PUB_ChkIsFMP(mPA01, mPA02, mPA03, mPA04) = True Or (mPA01 = "P" And mPA09 = "020") Then
       m_bolFMP = True
    End If
    'end 2022/10/07
    
    '¤¤»¡Ãþ«¬
    Select Case mCnKind
      Case "1" 'Â½Ä¶¤¤»¡201
        strPK = AutoNo("A", 6)
        strTF = strPK & "-201"
        mStrSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20) " & _
                 "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','201',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mPA01, "201")) & " from caseprogress where cp09='" & mCP09 & "' "
        cnnConnection.Execute mStrSql
        strPKList = strPKList & IIf(strPKList <> "", Right(strPK, 3), strPK) & ","
      Case "2" 'ÀËµø¤¤»¡209
        strPK = AutoNo("A", 6)
        strTF = strPK & "-209"
        mStrSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20) " & _
                 "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','209',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mPA01, "209")) & " from caseprogress where cp09='" & mCP09 & "' "
        cnnConnection.Execute mStrSql
        strPKList = strPKList & IIf(strPKList <> "", Right(strPK, 3), strPK) & ","
      Case "3", "4" '»s§@¤¤»¡210¡®¥~¤å´£¥Ó¥»242
        strPK = AutoNo("A", 6)
        strTF = strPK & "-210"
        mStrSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20) " & _
                 "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','210',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mPA01, "210")) & " from caseprogress where cp09='" & mCP09 & "' "
        cnnConnection.Execute mStrSql
        strPKList = strPKList & IIf(strPKList <> "", Right(strPK, 3), strPK) & ","
        If mCnKind = "4" Then
           strPK = AutoNo("A", 6)
           mStrSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20) " & _
                    "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','242',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mPA01, "242")) & " from caseprogress where cp09='" & mCP09 & "' "
           cnnConnection.Execute mStrSql
           strPKList = strPKList & IIf(strPKList <> "", Right(strPK, 3), strPK) & ","
        End If
      Case "5" '®Ö¹ï¤¤»¡235
        strPK = AutoNo("A", 6)
        strTF = strPK & "-235"
        mStrSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20) " & _
                 "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','235',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mPA01, "235")) & " from caseprogress where cp09='" & mCP09 & "' "
        cnnConnection.Execute mStrSql
        strPKList = strPKList & IIf(strPKList <> "", Right(strPK, 3), strPK) & ","
      'Added by Lydia 2018/05/07
      Case "6" 'ÀËµøPCT¤½¶}¥»»PFCP¬Û²§³B942
        strPK = AutoNo("A", 6)
        strTF = strPK & "-942"
        mStrSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp20) " & _
                 "select cp01,cp02,cp03,cp04,cp05,'" & strPK & "','942',cp11,cp12,cp13," & CNULL(PUB_GetCP20(mPA01, "942")) & " from caseprogress where cp09='" & mCP09 & "' "
        cnnConnection.Execute mStrSql
        strPKList = strPKList & IIf(strPKList <> "", Right(strPK, 3), strPK) & ","
    End Select
    
    If mPtyList <> "" Then
        tmpPty = Split(mPtyList, ",")
        For intJ = 0 To UBound(tmpPty)
           If Trim(tmpPty(intJ)) <> "" Then
               strANo = AutoNo("A", 6)  '·s¼W¦¬¤å¸¹
               strNowPty = Trim(tmpPty(intJ))
               strCon1 = "": strCon2 = ""  '­n·s¼WªºÄæ¦ì
               '¹w³]©Ó¿ì¤HCP14
               If InStr("416,", strNowPty) > 0 Then '¹ê¼f
                   'Modified by Morgan 2018/6/15 P®×¤£¦Û°Ê¤À®× Ex:P-120468
                   If mPA01 = "FCP" Then
                       strTmp1(1) = PUB_GetFCPHandler(mPA01, mPA02, mPA03, mPA04) '©Ó¿ì¤HCP14
                   Else
                       strTmp1(1) = ""
                   End If
                   'end 2018/6/15
               Else
                   strTmp1(1) = ""
               End If
               
               If InStr("902,924,968", strNowPty) > 0 Then '¤£¹w³]¶O¥Î: ¦^¥N(¦^ÂÐ¥N²z¤H)902, ·|½Z924,¦^´_»¡©ú®Ñ®Õ¾\968
                    strTmp1(2) = "": strTmp1(3) = "": strTmp1(4) = ""
                    strTmp1(5) = ""
                    m_Cp33 = Empty: m_Cp34 = Empty
               Else
                   'Memo --- ¦]¬°¬Oª½±µ·s¼W¨ä¥L¦¬¤å,©Ò¥H¤£³]©w¹q¤l°e¥ó
                    If PUB_Frm010005OnUpdFee(mPA01, mPA02, mPA03, mPA04, mPA09, mPA08, mPA16, mPA14, mCP07, strNowPty, mCP13, "", mCustList, strTmp1(2), strTmp1(3), strTmp1(4)) = True Then
                    End If
                    strTmp1(5) = strTmp1(2)              '¥¼¦¬ª÷ÃB CP79
                    '¼Ð·Ç»ù©M©³»ù
                    If ClsPDGetCaseLowPrice(mPA01, mPA09, strNowPty, m_Cp33, m_Cp34) = 1 Then
                    End If

                    strCon1 = strCon1 & ", cp16, cp17, cp18, cp33, cp34, cp79 "
                    strCon2 = strCon2 & ", " & CNULL(strTmp1(2)) & ", " & CNULL(strTmp1(3)) & ", " & CNULL(strTmp1(4)) & ", " & CNULL("" & m_Cp33) & ", " & CNULL("" & m_Cp34) & ", " & CNULL(strTmp1(5))
                End If
                
                '¹w³]©Ó¿ì´Á­­CP48
                'Modified by Lydia 2022/10/07 FMP®×¥ó­pºâ©Ó¿ì´Á­­; °Ñ¦Ò¤½§i1110401-04 ÂdÂi¦¬¤å©M¤º³¡¦¬¤åFMP®×¼W¥[­pºâ©Ó¿ì´Á­­
                'If InStr("203,902", strNowPty) > 0 Then
                If m_bolFMP = True Then
                    strTmp1(7) = Pub_GetHandleDay("FCP", "000", strNowPty, , TransDate(mCP06, 2)) '©Ó¿ì´Á­­
                ElseIf InStr("203,902", strNowPty) > 0 Then '¥D°Ê­×¥¿203, ¦^¥N(¦^ÂÐ¥N²z¤H)902
                'end 2022/10/07
                    If mPA01 = "FCP" Then  'Added by Lydia 2018/05/10 §PÂ_FCP®×¤~¦³©Ó¿ì´Á­­
                       strTmp1(7) = Pub_GetHandleDay("FCP", "000", strNowPty, , TransDate(mCP06, 2)) '©Ó¿ì´Á­­
                    Else 'Added by Lydia 2018/05/10 §PÂ_FCP®×¤~¦³©Ó¿ì´Á­­
                       strTmp1(7) = ""
                    End If
                    'end 2018/05/10
                Else
                    strTmp1(7) = ""
                End If
                
                'Modified by Lydia 2025/06/25
                'Pub_SetPAIsCase mPA01, strNowPty, m_Cp26 'Added by Lydia 2018/05/10 ¬O§_ºâ®×¥ó¼Æ
                If PUB_GetCPMbyCP10(mPA01, strNowPty, "cpm05") = "N" Then
                   m_Cp26 = "N"
                End If
                'end 2025/06/25
                'Modified by Lydia 2018/03/15 ¦Û°Ê¤W¤w¤À®×(CP122)
                'Modified by Lydia 2018/03/27 §PÂ_°hµ{§Ç¤£¤W¤w¤À®×
                'Modified by Lydia 2022/09/02 °Ï¤À¦³µL¹w³]¶O¥Î
                'mStrSQL = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp20,cp48,cp16,cp17,cp18,cp79,cp122,cp33,cp34,cp26) " & _
                            "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "'," & CNULL(strNowPty) & ",cp11,cp12,cp13,'" & strTmp1(1) & "'," & CNULL(PUB_GetCP20(mPA01, strNowPty)) & _
                                   ", " & CNULL(strTmp1(7), True) & ", " & CNULL(strTmp1(2)) & ", " & CNULL(strTmp1(3)) & ", " & CNULL(strTmp1(4)) & ", " & CNULL(strTmp1(5)) & _
                                    ", " & IIf(strTmp1(1) <> "" And mGrp_New <> "" And mGrp_New <> "B", "'Y'", "NULL") & ", " & CNULL("" & m_Cp33, True) & ", " & CNULL("" & m_Cp34, True) & "," & CNULL(m_Cp26) & " from caseprogress where cp09='" & mCP09 & "' "
                mStrSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp20,cp48,cp122,cp26 " & strCon1 & ") " & _
                               "select cp01,cp02,cp03,cp04,cp05,'" & strANo & "'," & CNULL(strNowPty) & ",cp11,cp12,cp13,'" & strTmp1(1) & "'," & CNULL(PUB_GetCP20(mPA01, strNowPty)) & _
                                ", " & CNULL(strTmp1(7), True) & ", " & IIf(strTmp1(1) <> "" And mGrp_New <> "" And mGrp_New <> "B", "'Y'", "NULL") & "," & CNULL(m_Cp26) & strCon2 & _
                                " from caseprogress where cp09='" & mCP09 & "' "
                cnnConnection.Execute mStrSql
                strPKList = strPKList & IIf(strPKList <> "", Right(strANo, 3), strANo) & ","
               
           End If
        Next intJ
    End If

         If strPKList <> "" Then
            'Added by Lydia 2018/06/28 «æ¥óÂ½Ä¶±N·s®×Â½Ä¶¦¬¤å¸¹¦^¼g¨ìÂ½Ä¶¶O¥ÎÀÉ©M©R¦W°O¿ýÀÉ(TCN14)
            If mTCN01 <> "" And strTF <> "" Then
                'Modified by Lydia 2020/01/09 FMP®×¦¬¤å¡A·s®×«ØÀÉ¥¼´£¥Ó¥ýÂ½Ä¶¦Û°Ê¤Ä¿ï
                mStrSql = "update TransFee set TF01=" & CNULL(Mid(strTF, 1, 9)) & IIf(mPA01 = "P", " ,TF31='Y' ", "") & _
                            " where TF01=" & CNULL(mTCN01)
                cnnConnection.Execute mStrSql, intK
                If intK > 0 Then
                     strReceiver = Pub_GetSpecMan("M") 'Added by Lydia 2019/12/23 «æ¥óÂ½Ä¶ªº¦¬¥óªÌ
                     mStrSql = "update TrackingCaseName set TCN14=" & CNULL(Mid(strTF, 1, 9)) & " where TCN01=" & CNULL(mTCN01)
                     cnnConnection.Execute mStrSql
                     If Right(strTF, 3) <> "201" Then
                         strTmp1(1) = Pub_GetSpecMan("M")
                         If strTmp1(1) <> "" Then
                            'Modified by Lydia 2019/12/23 +CC:«æ¥óÂ½Ä¶ªº¦¬¥óªÌ
                            'Modify By Sindy 2023/3/27 +,mc13
                            mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
                               " values( '" & strUserNum & "','" & strTmp1(1) & "',to_char(sysdate,'yyyymmdd') ,to_char(sysdate,'hh24miss'),'" & _
                               "«æ¥óÂ½Ä¶(" & mTCN01 & ")¥¼¦¬·s®×Â½Ä¶,½ÐÀË¬d" & mPA01 & "-" & mPA02 & IIf(mPA03 & mPA04 <> "000", "-" & mPA03 & "-" & mPA04, "") & "' ,'¦P¥D¦®','" & strReceiver & "','" & mCP09 & "')"
                            cnnConnection.Execute mStrSql
                         End If
                     End If
                     'Added by Lydia 2018/08/13 «æ¥óÂ½Ä¶¥ß®×®É,¦³¥æ½Z´Á­­²£¥Í°ê¥~³¡¦æ¨Æ¾ä(¤ñ·Ó·s®×«ØÀÉ)
                     'Modified by Lydia 2018/12/04 +¥u¥æClaims´Á­­
                     'Modified by Lydia 2019/12/23 §ï¦¨¦P®É§ì©R¦W°lÂÜ
                     strTmp1(0) = "select TF01,TF26,TF27,TF28,TF32,b2.* from TransFee a1, TrackingCaseName b2 where TF01=" & CNULL(Mid(strTF, 1, 9)) & " and tf01=tcn14(+) "
                     intJ = 1
                     Set rsQuery = ClsLawReadRstMsg(intJ, strTmp1(0))
                     If intJ = 1 Then
                         If Val("" & rsQuery.Fields("TF26")) > 0 Then
                            strTmp1(1) = PUB_GetFCPHandler(mPA01, mPA02, mPA03, mPA04)
                            strTmp1(0) = Pub_GetSpecMan("M")
                            If PUB_AddFCPStaffCalendar("" & rsQuery.Fields("TF26"), "1", strTmp1(1) & IIf(strTmp1(1) <> strTmp1(0), "," & strTmp1(0), ""), "Ä¶ªÌÂ½Ä¶¥æ½Z´Á­­", strTmp1(1), "1", mPA01, mPA02, mPA03, mPA04) = True Then
                            End If
                         End If
                         'Added by Lydia 2018/12/04 ¥u¥æClaims´Á­­
                         If Val("" & rsQuery.Fields("TF32")) > 0 Then
                            strTmp1(1) = PUB_GetFCPHandler(mPA01, mPA02, mPA03, mPA04)
                            strTmp1(0) = Pub_GetSpecMan("M")
                            If PUB_AddFCPStaffCalendar("" & rsQuery.Fields("TF32"), "1", strTmp1(1) & IIf(strTmp1(1) <> strTmp1(0), "," & strTmp1(0), ""), "Ä¶ªÌClaims¥æ½Z´Á­­", strTmp1(1), "1", mPA01, mPA02, mPA03, mPA04) = True Then
                            End If
                         End If
                         'Added by Lydia 2019/12/23 «æ¥óÂ½Ä¶¥[µù¦bemail¤º¤å
                         If Right(strTF, 3) = "201" And "" & rsQuery.Fields("TCN01") <> "" Then
                              strTmp1(9) = "" & rsQuery.Fields("tcn15") & " " & GetStaffName("" & rsQuery.Fields("tcn15"))
                              strContent = mPA01 & "-" & mPA02 & IIf(mPA03 & mPA04 <> "000", "-" & mPA03 & "-" & mPA04, "") & "¤w´£¨ÑÀÉ®×µ¹" & strTmp1(9) & "¡A¶i¦æÂ½Ä¶" & vbCrLf
                              strContent = strContent & "°lÂÜ¸¹¡G" & rsQuery.Fields("tcn01") & vbCrLf
                              strContent = strContent & "Â½Ä¶¤H­û¡G" & strTmp1(9) & vbCrLf
                              strContent = strContent & "¥æ½Z´Á­­¡G" & ChangeTStringToTDateString("" & rsQuery.Fields("tf26")) & vbCrLf
                              If "" & rsQuery.Fields("tf32") <> "" Then strContent = strContent & "¥u¥æClaims´Á­­¡G" & ChangeTStringToTDateString("" & rsQuery.Fields("tf32")) & vbCrLf
                              strContent = strContent & "­ì¤å»yºØ¡G" & Pub_GetTransFeeL("1", "" & rsQuery.Fields("tf27")) & vbCrLf
                              strContent = strContent & "Â½Ä¶»yºØ¡G" & Pub_GetTransFeeL("2", "" & rsQuery.Fields("tf28")) & vbCrLf
                              strContent = strContent & String(30, "-") & vbCrLf
                              strContent = strContent & "ºÞ¨î¤H¡G" & GetStaffName("" & rsQuery.Fields("tcn03")) & vbCrLf
                              strContent = strContent & "³Æ¡@µù¡G" & rsQuery.Fields("tcn04") & vbCrLf
                         End If
                     End If
                     'end 2018/08/13
                'Added by Lydia 2020/01/06 FMP®×¦¬¤å¡A·s®×«ØÀÉ¥¼´£¥Ó¥ýÂ½Ä¶¦Û°Ê¤Ä¿ï
                ElseIf mPA01 = "P" And mCnKind = "1" Then
                     mStrSql = "Insert into TransFee(TF01,TF31) values(" & CNULL(Mid(strTF, 1, 9)) & ", 'Y' )"
                     cnnConnection.Execute mStrSql
                     'Move by Lydia 2020/12/09 ²¾¨ì¤U¤è
                'end 2020/01/06
                End If
                'Added by Lydia 2018/08/27 ·s®×Â½Ä¶(201)¹w³]­Ó®×ªº©T©w³ø»ù(PA62)
                If mCnKind = "1" Then
                     strTmp1(0) = Pub_GetPa62Flag(mPA01 & mPA02 & mPA03 & mPA04)
                     If strTmp1(0) <> "" Then
                          mStrSql = "update patent set pa62='" & strTmp1(0) & "' where " & ChgPatent(mPA01 & mPA02 & mPA03 & mPA04)
                          cnnConnection.Execute mStrSql
                     End If
                     'Move by Lydia 2020/12/09 ±q¤W­±²¾¹L¨Ó ex.P-126344¦b¥ß¨÷«e¦³«æ¥óÂ½Ä¶
                     If mPA01 = "P" Then 'Added by Lydia 2020/12/09
                         'Added by Lydia 2020/08/24 FMP®×¹w³]µo"¥¼´£¥Ó¥ýÂ½Ä¶"email ; °Ñ¦Òfrm060102
                         strTmp1(0) = Pub_GetSpecMan("M")
                         If strTmp1(0) <> "" Then
                             'Modify By Sindy 2023/3/27 +,mc13
                             mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
                                " values( '" & strUserNum & "','" & strTmp1(0) & "',to_char(sysdate,'yyyymmdd')" & _
                                ",to_char(sysdate,'hh24miss'),'" & mPA01 & mPA02 & IIf(mPA03 & mPA04 <> "000", mPA03 & mPA04, "") & " ¥¼´£¥Ó¥ýÂ½Ä¶" & "','¦P¥D¦®',null,'" & mCP09 & "')"
                             cnnConnection.Execute mStrSql
                         End If
                         'end 2020/08/24
                     End If 'Added by Lydia 2020/12/09
                     'end------Move by Lydia 2020/12/09
                End If
                'end 2018/08/27
            End If
            'end 2018/06/28
            strPKList = Replace(Mid(strPKList, 1, Len(strPKList) - 1), ",", "¡B")
            'frm010001.lblTCT.Caption = "¤¤»¡©Î¨ä¥L¦¬¤å¸¹¡G" '²¾¨ì¥~¼hÅÜ§ó
1          'frm010001.lblTCTNO.Caption = strPKList
            mRetList = strPKList
         End If
      
      'Added by Lydia 2018/09/06 §PÂ_¦³¨«©R¦W¬yµ{ªº·s®×©Ê½è¤~²£¥Í°O¿ý
      If InStr(FcpAddTct, mCP10) > 0 Then
          '§ì¤À®×²Õ§O-¥DºÞ
          strTmp1(6) = ""
          strTmp1(6) = Pub_GetFCPGrpMan(mGrp_New)
          strTmp1(6) = PUB_GetStateForMan(strTmp1(6)) 'Added by Lydia 2022/10/12 ¯S®í±¡ªp¤§«ü©wÂ¾¥N
          If strTmp1(6) = "" Then strTmp1(6) = "B"
          'end 2019/01/09
          
          'Ä¶²¦´Á­­
          strTmp1(4) = IIf(mTCT0203_new <> "", TransDate(Left(mTCT0203_new, 7), 2), "")
          strTmp1(5) = IIf(mTCT0203_new <> "", Mid(mTCT0203_new, 8), "")
          
          mStrSql = ""
          If bolExistTCT = False Then
             mStrSql = "INSERT INTO TransCaseTitle(TCT01,TCT02,TCT03,TCT04,TCT16,TCT17,TCT112,TCT113,TCT114) " & _
                      "VALUES ('" & mCP09 & "'," & CNULL(strTmp1(4), True) & "," & CNULL(strTmp1(5), True) & "," & CNULL(IIf(strTmp1(6) <> "B", strTmp1(6), "")) & _
                      ",'" & ChgSQL(mPA05) & "','" & ChgSQL(mPA06) & "','" & strUserNum & "'," & strSrvDate(1) & "," & Mid(Format(ServerTime, "000000"), 1, 4) & ")"
          Else
             'Ä¶²¦´Á­­
             If strTmp1(4) <> TransDate(Left(mTCT0203_old, 7), 2) Then
                mStrSql = mStrSql & ", TCT02=" & CNULL(strTmp1(4), True)
             End If
             If strTmp1(5) <> Mid(mTCT0203_old, 8) Then
                mStrSql = mStrSql & ", TCT03=" & CNULL(strTmp1(5), True)
             End If
             '¤À®×
             If mPA150 = "" And mGrp_old <> strTmp1(6) Then 'PA150®×¥ó¤w¤À¤uµ{®v²Õ§O
                mStrSql = mStrSql & ", TCT04=" & CNULL(IIf(strTmp1(6) <> "B", strTmp1(6), ""))
             End If
             '®×¥ó¦WºÙ
                  mStrSql = mStrSql & ", TCT16=" & CNULL(ChgSQL(mPA05))

                  mStrSql = mStrSql & ", TCT17=" & CNULL(ChgSQL(mPA06))
             If mStrSql <> "" Then
                mStrSql = "UPDATE TransCaseTitle SET TCT112='" & strUserNum & "', TCT113=" & strSrvDate(1) & ", TCT114=" & Mid(Format(ServerTime, "000000"), 1, 4) & mStrSql & " WHERE TCT01='" & mCP09 & "' "
             End If
          End If
          If mStrSql <> "" Then
            cnnConnection.Execute mStrSql
            'µoemail³qª¾¦U²Õ¥DºÞ
            '·s¼W
            If bolExistTCT = False Then
                  'Added by Lydia 2023/02/17
                  If strSrvDate(1) >= ¥~±M·s®×»{»â±Ò¥Î¤é Then
                       If PUB_UpdateTCNstate("1", mPA01 & mPA02 & mPA03 & mPA04) = False Then
                       End If
                  End If
                  'end 2023/02/17
                  'Modified by Lydia 2019/12/23 + strContent,°Æ¥»¦¬¨üªÌ
                  'Modified by Lydia 2022/08/08 ±NDavid¥[¤J(­^¤å²Õ)©Ò¦³FCP, PÂd¥x·s®×¥ß¨÷ (101-103)³qª¾¦¬¥ó¤H¤§¤@
                  strTmp1(7) = ""
                  If InStr("101,102,103,", mCP10 & ",") > 0 And strSrvDate(1) >= ¥~±M«H¥ó¨R¾P±Ò¥Î¤é Then
                      'Modified by Lydia 2023/08/11 §ï§PÂ_´¼Åv¤H­û=ºÞ¨î¤Hªº²Õ§O
                      'If mPA75 <> "" Then  'FC¥N²z¤H
                      '    strTmp1(7) = GetPrjNationNumber(ChangeCustomerL(mPA75))
                      'ElseIf Len(mCustList) > 6 Then '¥Ó½Ð¤H1
                      '    strTmp1(7) = GetPrjNationNumber1(ChangeCustomerL(Mid(mCustList, 1, InStr(mCustList, ",") - 1)))
                      'End If
                      'If Left(strTmp1(7), 3) <> "011" Then
                      If m_TCN03st16 = "1" Then
                      'end 2023/08/11
                          strTmp1(7) = Pub_GetSpecMan("¥~±M©Ó¿ì­^¤å²Õ¥DºÞ")
                      Else '¤é¥»®×­n²MªÅ
                          strTmp1(7) = ""
                      End If
                      '«e­±+¸¹¬O¬°¤FPUB_GetTCTmail°Ï¹j«ü©w¦¬¥ó¤Hªº²Å¸¹
                      If strTmp1(7) <> "" Then strTmp1(7) = "+" & strTmp1(7)
                  End If
                  'Modified by Lydia 2023/02/17°Ï¤À¡uÂd¥x¦¬·s®×=·s®×¥ß¨÷¡v iSta=1=>0
                  'If PUB_GetTCTmail(True, 1, mPA01, mPA02, mPA03, mPA04, mCP09, strTmp1(6), strTmp1(7), , , strContent, strReceiver) Then
                  If PUB_GetTCTmail(True, 0, mPA01, mPA02, mPA03, mPA04, mCP09, strTmp1(6), strTmp1(7), , , strContent, strReceiver) Then
                  'end 2022/08/08
                  End If
            '­×§ï
            ElseIf mGrp_old <> strTmp1(6) Or mTCT0203_old <> mTCT0203_new Then
                  If mGrp_old <> strTmp1(6) Then
                      '§ï²Õ§O
                      strTmp1(6) = mGrp_old & ";" & strTmp1(6)
                      If PUB_GetTCTmail(True, 2, mPA01, mPA02, mPA03, mPA04, mCP09, , strTmp1(6), mGrp_old & "-" & mGrp_New, "­×§ï¸Éµo: ") = True Then
                      End If
                  Else
                      If PUB_GetTCTmail(True, 1, mPA01, mPA02, mPA03, mPA04, mCP09, strTmp1(6), , , "­×§ï¸Éµo: ") Then
                      End If
                  End If
            End If
          End If
      End If

      Set rsQuery = Nothing
End Sub

'Added by Lydia 2021/04/09 ÀË¬d§ó¦W«áªºÀÉ¦W¬O§_­«ÂÐ
'Move by Lydia 2022/09/01 ±qfrm010005·h¨Ó
Private Function GetNewFilename(ByVal pFielPath As String, ByVal pOldName As String, ByVal PTitle As String, Optional ByVal pMid As String) As String
Dim intQ As Integer
Dim strPass As String

     If pMid <> "" And InStr(UCase("." & pOldName), UCase(pMid)) > 0 Then
         GetNewFilename = PTitle & UCase(pMid) & Mid(pOldName, InStrRev(pOldName, ".") + 1)
         intQ = 1
         strPass = Dir(pFielPath & "\" & GetNewFilename)
         Do While strPass <> ""
              GetNewFilename = PTitle & "." & intQ & UCase(pMid) & Mid(pOldName, InStrRev(pOldName, ".") + 1)
              intQ = intQ + 1
              Sleep 1000 'Added by Lydia 2021/04/15
              strPass = Dir(pFielPath & "\" & GetNewFilename)
         Loop
         Sleep 1000 'Added by Lydia 2021/04/15
     Else
         GetNewFilename = PTitle & "." & pOldName
     End If

End Function

'Added by Lydia 2022/09/14 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡GLAÅU°Ý®×¤§0ÅU°Ý¸u¥ô¦¬¤å(±qfrm010006.SaveDatabase©â¥X¨Ó)
'Modify By Sindy 2024/11/21 + , Optional ByVal m_intCRC As Integer = 0: ¦Û°Ê¦¬¤åªº®×¥ó©Ê½è¶¶§Ç
Public Function PUB_SaveFrm010006(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intChoose As Integer, _
                ByRef mHC() As String, ByRef mCP() As String, ByVal mCU30 As String, Optional ByRef IsSaveData As Boolean, _
                Optional ByVal pType As String, Optional ByVal pCaseNo As String, Optional ByVal m_intCRC As Integer = 0) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
Dim m_SalesST15 As String, m_SalesST06 As String
Dim m_SalesDeptName As String
Dim m_CP10Name As String '¦¬¤å¤§®×¥ó©Ê½è¦WºÙ
Dim rsRD As New ADODB.Recordset
'ªk«ß©Ò®×·½¦¬¤å
Dim m_LOS01 As String '®×·½Á`¦¬¤å¸¹
Dim m_LOS01cp01 As String, m_LOS01cp02 As String, m_LOS01cp03 As String, m_LOS01cp04 As String '®×·½Á`¦¬¤å¸¹¤§¥»©Ò®×¸¹
Dim m_LOS02 As String '®×·½®×¥óÃþ«¬
Dim m_LOS15 As String '®×·½³æ¸¹
Dim m_LOS04 As String  '¤¶²Ð¤H
Dim m_LOS04_1 As String, m_LOS04_1st15 As String, m_LOS04_1st06 As String '¤¶²Ð¤H(²Ä¤@¦ì)¡B¦¬¤å³¡ªù¡B©Ò§O
Dim m_LOS05 As String  '¤¶²Ð«È¤á
Dim m_LOS12 As String  '¤¶²Ð¤é
Dim m_Los05_N As String   'LA¸É®×·½¤§¤¶²Ð¤H¤¶²Ð«È¤á
Dim oMailCount As String
      
'*********¯S®íºÞ¨îªºÅÜ¼Æ*************
    'Modify By Sindy 2025/8/18
    'If pType = "LOS®×·½¦¬¤å" And pCaseNo <> "" Then
    If InStr(pType, "LOS®×·½¦¬¤å") > 0 And pCaseNo <> "" Then
    '2025/8/18 END
        m_LOS02 = Mid(pCaseNo, 1, InStr(pCaseNo, ",") - 1) '®×·½®×¥óÃþ«¬
        m_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1, 8) '®×·½³æ¸¹ 'Modify By Sindy 2025/8/18 +, 8)
        strTmp1(0) = "select X.*,cp01,cp02,cp03,cp04 from LawOfficeSource X,caseprogress where los15=" & CNULL(m_LOS15) & " and los01=cp09(+) "
        intJ = 1
        Set rsRD = ClsLawReadRstMsg(intJ, strTmp1(0))
        If intJ = 1 Then
          '®×·½Á`¦¬¤å¸¹
          m_LOS01 = "" & rsRD.Fields("LOS01")
          '®×·½Á`¦¬¤å¸¹¤§¥»©Ò®×¸¹
          m_LOS01cp01 = "" & rsRD.Fields("cp01")
          m_LOS01cp02 = "" & rsRD.Fields("cp02")
          m_LOS01cp03 = "" & rsRD.Fields("cp03")
          m_LOS01cp04 = "" & rsRD.Fields("cp04")
          '(­ì)®×·½®×¥óÃþ«¬
          m_LOS02 = "" & rsRD.Fields("LOS02")
          '®×·½³æ¸¹
          m_LOS15 = "" & rsRD.Fields("LOS15")
          '¤¶²Ð¤H, ¤¶²Ð¤H(²Ä¤@¦ì)
          m_LOS04 = "" & rsRD.Fields("LOS04")
          If m_LOS04 <> "" And InStr(m_LOS04, ",") > 0 Then
             m_LOS04_1 = Mid(m_LOS04, 1, InStr(m_LOS04, ",") - 1)
          Else
             m_LOS04_1 = m_LOS04
          End If
          If m_LOS04_1 <> "" Then
             m_LOS04_1st15 = GetST15(m_LOS04_1, , , m_LOS04_1st06)
          End If
          '(­ì)¤¶²Ð«È¤á:
          m_LOS05 = "" & rsRD.Fields("LOS05")
          '¤¶²Ð¤é
          m_LOS12 = "" & rsRD.Fields("LOS12")
        End If
        Set rsRD = Nothing
    End If
'***********************************
   intJ = ClsPDGetCaseProperty(mHC(1), mCP(10), m_CP10Name)
   m_SalesST15 = GetST15(mCP(13), m_SalesDeptName, , m_SalesST06)
   'Added by Lydia 2023/05/11 ¦]¬°PUB_ReadCaseData·|¦^¶Ç6½X«È¤á½s¸¹,©Ò¥H¥ý²Î¤@«È¤á½s¸¹
   mHC(5) = ChangeCustomerL(mHC(5))
   mHC(24) = ChangeCustomerL(mHC(24))
   mHC(25) = ChangeCustomerL(mHC(25))
   mHC(26) = ChangeCustomerL(mHC(26))
   mHC(27) = ChangeCustomerL(mHC(27))
   'end 2023/05/11
         
   If intModifyKind = 0 Then
      PUB_SaveFrm010006 = InsertHireCaseDB(pFormName, intSaveMode, intModifyKind, intChoose, mHC, mCP, mCU30, IsSaveData, pType, pCaseNo, m_Los05_N)
   Else
      PUB_SaveFrm010006 = UpdateHireCaseDB(pFormName, intSaveMode, intModifyKind, intChoose, mHC, mCP, mCU30, IsSaveData, pType, pCaseNo)
   End If
   If PUB_SaveFrm010006 = False Then Exit Function 'Add By Sindy 2022/9/28 ¦sÀÉ¥¢±Ñ,«áÄò¤£ÀË¬d
   'add by nickc 2007/11/09 ´ú¸Õ¸Ñ¨Mmail µo¤£¨ìªº®É­Ô·|¦s¨âµ§ªº¿ù»~
   On Error GoTo 0    'Âk¹s
   On Error GoTo ErrHand 'Add By Sindy 2022/9/29
   'Add By Sindy 2022/12/29 ­«ÅªCP,¦]«eÀYUpdate¨ç¼Æµ{¦¡¦³¥i¯àª½±µ¦sDB,¨S¦³§ó·scp³¯¦C­È
   strTmp1(0) = mCP(9)
   Erase mCP
   ReDim Preserve mCP(TF_CP) As String
   mCP(9) = strTmp1(0)
   'Modified by Lydia 2023/05/11 + false
   Call PUB_ReadCaseProgressDatabase(mCP(), 1, False)
   '2022/12/29 END
   
   If intModifyKind = 0 Then
      Dim oContext As String, strCaseNo As String
      Dim strTemp As String
      Dim m_strState As String
      'Add By Sindy 2023/8/18 ¤£±o¥N²zªº«áÄòÂÂ®×¦¬¤å±±ºÞ¡A³qª¾¦¬¤å¤H­û¡]CP13¡^
      If mHC(5) <> "" Then
        If GetCustomerAndState(mHC(5), strTmp1(1), , , , mHC(1), m_strState, IIf(intSaveMode = 0, True, False), , mHC(2), mHC(3), mHC(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H1¡G " + mHC(5) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mHC(5)
          End If
        End If
      End If
      If mHC(24) <> "" Then
        If GetCustomerAndState(mHC(24), strTmp1(1), , , , mHC(1), m_strState, IIf(intSaveMode = 0, True, False), , mHC(2), mHC(3), mHC(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H2¡G " + mHC(24) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mHC(24)
          End If
        End If
      End If
      If mHC(25) <> "" Then
        If GetCustomerAndState(mHC(25), strTmp1(1), , , , mHC(1), m_strState, IIf(intSaveMode = 0, True, False), , mHC(2), mHC(3), mHC(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H3¡G " + mHC(25) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mHC(25)
          End If
        End If
      End If
      If mHC(26) <> "" Then
        If GetCustomerAndState(mHC(26), strTmp1(1), , , , mHC(1), m_strState, IIf(intSaveMode = 0, True, False), , mHC(2), mHC(3), mHC(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H4¡G " + mHC(26) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mHC(26)
          End If
        End If
      End If
      If mHC(27) <> "" Then
        If GetCustomerAndState(mHC(27), strTmp1(1), , , , mHC(1), m_strState, IIf(intSaveMode = 0, True, False), , mHC(2), mHC(3), mHC(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H5¡G " + mHC(27) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mHC(27)
          End If
        End If
      End If
      If oContext <> "" Then
         strTemp = Mid(strTemp, 2)
         strCaseNo = IIf("-" + mHC(3) + "-" + mHC(4) = "-0-00", mHC(1) + "-" + mHC(2), mHC(1) + "-" + mHC(2) + "-" + mHC(3) + "-" + mHC(4))
         oContext = "¥»©Ò®×¸¹¡G " + strCaseNo + vbCrLf + _
                    "®×¥ó¦WºÙ¡G " + mHC(6) + vbCrLf + _
                    "¦¬¤å¤é¡G " + ChangeTStringToTDateString(TransDate(mCP(5), 1)) + vbCrLf + _
                    "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf + vbCrLf + _
                    "¡i¤£±o¥N²z¡j" + vbCrLf + _
                    oContext
         oMailCount = mCP(13) & ";" & PUB_GetFCPProSup(mCP(13), True)
         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & strCaseNo & _
            " ¤w½T»{Äò¦æ¦¬¤å¡A½Ðª`·N¸Ó" & strTemp & "½s¸¹¤w³]¬°¤£±o¥N²z¡C(¤å¸¹:" & mCP(9) & ")','" & oContext & "',null,'" & mCP(9) & "')"
         cnnConnection.Execute mStrSql
      End If
      '2023/8/18 END
      
      'Added by Lydia 2020/10/05 (9/30) ­Y¸Ó¦¬¤å¸¹ÂI¼Æ>0¦ýµL®×·½(¦Û¦æ¦¬¤åªÌ)®É¡A­Y®×¥óªº«È¤á¬°«Dªk«ß©Òªº«È¤á®É«h¬°A3Ãþ®×·½¡A¤£½×·sÂÂ®×¡A¨t²Î¦Û°Ê·s¼WTT-999999®×¶i«×(BÃþ¦¬¤å)¤Îªk«ß©Ò®×·½¸ê®Æ¡C­Y¬°·s®×·~°È°Ï¤£¦PªºEmail·ÓÂÂ³qª¾¡C
      If m_Los05_N <> "" Then  '¦]¬°Âd¥xµLªk³B²z,©Ò¥H¥uµoemail
          m_LOS05 = Mid(m_Los05_N, 1, InStr(m_Los05_N, "|") - 1)
          m_LOS04_1 = Mid(m_Los05_N, InStr(m_Los05_N, "|") + 1)
          m_LOS04_1st15 = GetST15(m_LOS04_1)
          'Added by Lydia 2023/12/08 ¸É®×·½=>¤@¯ë®×·½¦¬¤å§PÂ_; °Ñ¦Ò¤é±`¤u§@\¸ê®Æ§R§ï°O¿ý\LA003060
          If InStr(pType, "LOS®×·½¦¬¤å") = 0 Then
              pType = "LOS®×·½¦¬¤å"
          End If
      End If
      'end 2020/10/05
      
      'Add By Sindy 2024/11/6
      If Not (mHC(5) = "" And mHC(24) = "" And mHC(25) = "" And mHC(26) = "" And mHC(27) = "") Then
         oContext = "¥»©Ò®×¸¹¡G " + mHC(1) + "-" + mHC(2) + "-" + mHC(3) + "-" + mHC(4) + vbCrLf + "®×¥ó¦WºÙ¡G " + mHC(6) + vbCrLf + "¦¬¤å¤é¡G " + ChangeTStringToTDateString(mCP(5)) + vbCrLf + "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf
         'Added by Lydia 2022/11/03 debug:«D®×·½(¸É®×·½¤£ºâ)§ï¥Îµe­±§PÂ_
         'Memo by Lydia 2023/12/08 §ï¦¨¸É®×·½=>¤@¯ë®×·½¦¬¤å§PÂ_
                     '°Ñ¦Ò¤å¥ó\\LINUX\PolyCOM\TaieNew\¹q¸£¤¤¤ß¤é±`¤u§@\¨Æ°È©ÒÅÜ°Ê¬ÛÃö¨Æ¶µ\ªk«ß©Ò¤Î´¼°]©Ò\¥¼§¹-®×¥ó¨t²Î¼W­×µ{¦¡.docx
                     '9/30 ­Y¸Ó¦¬¤å¸¹ÂI¼Æ>0¦ýµL®×·½(¦Û¦æ¦¬¤åªÌ)®É¡A­Y®×¥óªº«È¤á¬°«Dªk«ß©Òªº«È¤á®É«h¬°A3Ãþ®×·½¡A¤£½×·sÂÂ®×¡A¨t²Î¦Û°Ê·s¼WTT-999999®×¶i«×(BÃþ¦¬¤å)¤Îªk«ß©Ò®×·½¸ê®Æ¡C¦p¬°ÂÂ®×¥B´¿¦³A3Ãþ®×·½®É¡A¤¶²Ð¤H­û¦P³Ì«á¤@µ§®×·½¡A§_«h¥H«È¤á¥Ø«eªº´¼Åv¤H­û¬°¤¶²Ð¤H¡C­Y¬°·s®×·~°È°Ï¤£¦PªºEmail·ÓÂÂ³qª¾¡C
         If InStr(pType, "LOS®×·½¦¬¤å") = 0 Or m_LOS05 = "" Or m_LOS04_1 = "" Then
            
            'Modify By Sindy 2024/11/6 §ï¦¨¦@¥Î¨ç¼Æ: ¦¬¤å®É,ÀË¬d¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¬O§_¦³»~
            '§ï¼g­ì¥Ñ¬O¦]¬°¥Ó½Ð¤H1~5 ³v¤@ÀË¬d,¦³»~§¡­nµo mail
            'Modify By Sindy 2024/11/21 ¦Û°Ê¦¬¤åªº®×¥ó©Ê½è¶¶§Ç=1 ©Î¯È¥»¦¬¤å¥¼«ü©w
            If m_intCRC = 1 Or m_intCRC = 0 Then
            '2024/11/21 END
               Call RecvChkApplCust("«È¤á½s¸¹1", mHC(5), mCP(13), "", m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
               Call RecvChkApplCust("«È¤á½s¸¹2", mHC(24), mCP(13), "", m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
               Call RecvChkApplCust("«È¤á½s¸¹3", mHC(25), mCP(13), "", m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
               Call RecvChkApplCust("«È¤á½s¸¹4", mHC(26), mCP(13), "", m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
               Call RecvChkApplCust("«È¤á½s¸¹5", mHC(27), mCP(13), "", m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
            End If
         End If 'Added by Lydia 2022/11/03
         'Modified by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¤¶²Ð«È¤á¬°ÂÂ«È¤á¦ý»P¤¶²Ð¤H¤£¦P°Ï®ÉµoMail³qª¾¬ÛÃö¤H­û
         'Modified by Lydia 2022/11/03 ¸É®×·½¤£ºâ And InStr(pType, "LOS®×·½¦¬¤å") > 0
         If strSrvDate(1) >= ªk«ß©Ò®×·½¦¬¤å±Ò¥Î¤é And InStr(pType, "LOS®×·½¦¬¤å") > 0 And m_LOS05 <> "" And m_LOS04_1 <> "" Then
            
            'Modify By Sindy 2024/11/6 §ï¦¨¦@¥Î¨ç¼Æ: ¦¬¤å®É,ÀË¬d¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¬O§_¦³»~
            '§ï¼g­ì¥Ñ¬O¦]¬°¥Ó½Ð¤H1~5 ³v¤@ÀË¬d,¦³»~§¡­nµo mail
            'Modify By Sindy 2024/11/21 ¦Û°Ê¦¬¤åªº®×¥ó©Ê½è¶¶§Ç=1 ©Î¯È¥»¦¬¤å¥¼«ü©w
            If m_intCRC = 1 Or m_intCRC = 0 Then
            '2024/11/21 END
               Call RecvChkApplCust("«È¤á½s¸¹1", mHC(5), m_LOS04_1, "", m_LOS04_1st15, Trim(mCP(12)), oContext, m_LOS04_1st06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9), m_LOS04_1)
               Call RecvChkApplCust("«È¤á½s¸¹2", mHC(24), m_LOS04_1, "", m_LOS04_1st15, Trim(mCP(12)), oContext, m_LOS04_1st06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9), m_LOS04_1)
               Call RecvChkApplCust("«È¤á½s¸¹3", mHC(25), m_LOS04_1, "", m_LOS04_1st15, Trim(mCP(12)), oContext, m_LOS04_1st06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9), m_LOS04_1)
               Call RecvChkApplCust("«È¤á½s¸¹4", mHC(26), m_LOS04_1, "", m_LOS04_1st15, Trim(mCP(12)), oContext, m_LOS04_1st06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9), m_LOS04_1)
               Call RecvChkApplCust("«È¤á½s¸¹5", mHC(27), m_LOS04_1, "", m_LOS04_1st15, Trim(mCP(12)), oContext, m_LOS04_1st06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9), m_LOS04_1)
            End If
         End If
      End If
      '2024/11/6 END
'Modify By Sindy 2024/11/6 mark
'      'Modify By Sindy 2011/1/18
'      '·í¦¬¤å·~°È°Ï»P«È¤áÀÉ·~°È°Ï¤£¦P®Éµo mail  ¤Î´£¥Ü
'      Dim oStrCuSales1 As String
'      Dim oStrCuSales2 As String
'      Dim oStrCuSales3 As String
'      Dim oStrCuSales4 As String
'      Dim oStrCuSales5 As String
'      '¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'      Dim IsMail As Boolean
'      IsMail = True
'
'      oStrCuSales1 = ""
'      oStrCuSales2 = ""
'      oStrCuSales3 = ""
'      oStrCuSales4 = ""
'      oStrCuSales5 = ""
'
'         'Added by Lydia 2022/11/03 debug:«D®×·½(¸É®×·½¤£ºâ)§ï¥Îµe­±§PÂ_
'         'Memo by Lydia 2023/12/08 §ï¦¨¸É®×·½=>¤@¯ë®×·½¦¬¤å§PÂ_
'                     '°Ñ¦Ò¤å¥ó\\LINUX\PolyCOM\TaieNew\¹q¸£¤¤¤ß¤é±`¤u§@\¨Æ°È©ÒÅÜ°Ê¬ÛÃö¨Æ¶µ\ªk«ß©Ò¤Î´¼°]©Ò\¥¼§¹-®×¥ó¨t²Î¼W­×µ{¦¡.docx
'                     '9/30 ­Y¸Ó¦¬¤å¸¹ÂI¼Æ>0¦ýµL®×·½(¦Û¦æ¦¬¤åªÌ)®É¡A­Y®×¥óªº«È¤á¬°«Dªk«ß©Òªº«È¤á®É«h¬°A3Ãþ®×·½¡A¤£½×·sÂÂ®×¡A¨t²Î¦Û°Ê·s¼WTT-999999®×¶i«×(BÃþ¦¬¤å)¤Îªk«ß©Ò®×·½¸ê®Æ¡C¦p¬°ÂÂ®×¥B´¿¦³A3Ãþ®×·½®É¡A¤¶²Ð¤H­û¦P³Ì«á¤@µ§®×·½¡A§_«h¥H«È¤á¥Ø«eªº´¼Åv¤H­û¬°¤¶²Ð¤H¡C­Y¬°·s®×·~°È°Ï¤£¦PªºEmail·ÓÂÂ³qª¾¡C
'         If InStr(pType, "LOS®×·½¦¬¤å") = 0 Or m_LOS05 = "" Or m_LOS04_1 = "" Then
'            oContext = "¥»©Ò®×¸¹¡G " + mHC(1) + "-" + mHC(2) + "-" + mHC(3) + "-" + mHC(4) + vbCrLf + "®×¥ó¦WºÙ¡G " + mHC(6) + vbCrLf + "¦¬¤å¤é¡G " + ChangeTStringToTDateString(mCP(5)) + vbCrLf + "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf
'            oMailCount = ""
'            'Modified by Lydia 2019/02/14
'            'If GetST15(mcp(13)) <> GetCuSales(mhc(5), oStrCuSales1) And Trim(mcp(13)) <> "" And Trim(mhc(5)) <> "" Then
'            '   If Left(Trim(GetST15(mcp(13))), 1) = "F" And Left(Trim(GetCuSales(mhc(5), oStrCuSales1)), 1) = "F" Then
'            If m_SalesST15 <> GetCuSales(mHC(5), oStrCuSales1) And Trim(mCP(13)) <> "" And Trim(mHC(5)) <> "" Then
'               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mHC(5), oStrCuSales1)), 1) = "F" Then
'            'end 2019/02/14
'                  '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'               Else
'                  oMailCount = oMailCount & oStrCuSales1 & ";"
'                  'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                  If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales1), 1) = "S" And _
'                     InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                     oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                  End If
'                  '2023/11/7 END
'                  oContext = oContext & vbCrLf + "«È¤á½s¸¹1¡G " + GetCustomerName(mHC(5)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales1)
'               End If
'             '¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'             Else
'                   If Trim(mCP(13)) <> "" And Trim(mHC(5)) <> "" Then
'                       IsMail = False
'                   End If
'            End If
'            'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'            If m_SalesST06 <> "" And Trim(mHC(5)) <> "" And Trim(mCP(13)) <> "" Then
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, mHC(5), Trim(mCP(13)), m_SalesST15, m_SalesST06, , mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                   IsMail = False
'               End If
'            End If
'
'            'Modified by Lydia 2019/02/14
'            'If GetST15(mcp(13)) <> GetCuSales(mhc(24), oStrCuSales2) And Trim(mcp(13)) <> "" And mhc(24) <> "" Then
'            '   If Left(Trim(GetST15(mcp(13))), 1) = "F" And Left(Trim(GetCuSales(mhc(24), oStrCuSales2)), 1) = "F" Then
'            If m_SalesST15 <> GetCuSales(mHC(24), oStrCuSales2) And Trim(mCP(13)) <> "" And mHC(24) <> "" Then
'               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mHC(24), oStrCuSales2)), 1) = "F" Then
'            'end 2019/02/14
'                  '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'               Else
'                  oMailCount = oMailCount & oStrCuSales2 & ";"
'                  'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                  If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales2), 1) = "S" And _
'                     InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                     oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                  End If
'                  '2023/11/7 END
'                  oContext = oContext & vbCrLf + "«È¤á½s¸¹2¡G " + GetCustomerName(mHC(24)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales2)
'               End If
'             '¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'             Else
'                   If Trim(mCP(13)) <> "" And mHC(24) <> "" Then
'                       IsMail = False
'                   End If
'            End If
'            'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'            If m_SalesST06 <> "" And Trim(mHC(24)) <> "" And Trim(mCP(13)) <> "" Then
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, mHC(24), Trim(mCP(13)), m_SalesST15, m_SalesST06, , mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                   IsMail = False
'               End If
'            End If
'
'            'Modified by Lydia 2019/02/14
'            'If GetST15(mcp(13)) <> GetCuSales(mhc(25), oStrCuSales3) And Trim(mcp(13)) <> "" And mhc(25)<> "" Then
'            '   If Left(Trim(GetST15(mcp(13))), 1) = "F" And Left(Trim(GetCuSales(mhc(25), oStrCuSales3)), 1) = "F" Then
'            If m_SalesST15 <> GetCuSales(mHC(25), oStrCuSales3) And Trim(mCP(13)) <> "" And mHC(25) <> "" Then
'               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mHC(25), oStrCuSales3)), 1) = "F" Then
'            'end 2019/02/14
'                  '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'               Else
'                  oMailCount = oMailCount & oStrCuSales3 & ";"
'                  'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                  If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales3), 1) = "S" And _
'                     InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                     oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                  End If
'                  '2023/11/7 END
'                  oContext = oContext & vbCrLf + "«È¤á½s¸¹3¡G " + GetCustomerName(mHC(25)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales3)
'               End If
'             '¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'             Else
'                   If Trim(mCP(13)) <> "" And mHC(25) <> "" Then
'                       IsMail = False
'                   End If
'            End If
'            'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'            If m_SalesST06 <> "" And Trim(mHC(25)) <> "" And Trim(mCP(13)) <> "" Then
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, mHC(25), Trim(mCP(13)), m_SalesST15, m_SalesST06, , mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                   IsMail = False
'               End If
'            End If
'
'            'Modified by Lydia 2019/02/14
'            'If GetST15(mcp(13)) <> GetCuSales(mhc(26), oStrCuSales4) And Trim(mcp(13)) <> "" And mhc(26) <> "" Then
'            '   If Left(Trim(GetST15(mcp(13))), 1) = "F" And Left(Trim(GetCuSales(mhc(26), oStrCuSales4)), 1) = "F" Then
'            If m_SalesST15 <> GetCuSales(mHC(26), oStrCuSales4) And Trim(mCP(13)) <> "" And mHC(26) <> "" Then
'               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mHC(26), oStrCuSales4)), 1) = "F" Then
'            'end 2019/02/14
'                  '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'               Else
'                  oMailCount = oMailCount & oStrCuSales4 & ";"
'                  'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                  If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales4), 1) = "S" And _
'                     InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                     oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                  End If
'                  '2023/11/7 END
'                  oContext = oContext & vbCrLf + "«È¤á½s¸¹4¡G " + GetCustomerName(mHC(26)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales4)
'               End If
'             Else
'                   If Trim(mCP(13)) <> "" And mHC(26) <> "" Then
'                       IsMail = False
'                   End If
'            End If
'            'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'            If m_SalesST06 <> "" And Trim(mHC(26)) <> "" And Trim(mCP(13)) <> "" Then
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, mHC(26), Trim(mCP(13)), m_SalesST15, m_SalesST06, , mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                   IsMail = False
'               End If
'            End If
'
'            'Modified by Lydia 2019/02/14
'            'If GetST15(mcp(13)) <> GetCuSales(mhc(27), oStrCuSales5) And Trim(mcp(13)) <> "" And mhc(27) <> "" Then
'            '   If Left(Trim(GetST15(mcp(13))), 1) = "F" And Left(Trim(GetCuSales(mhc(27), oStrCuSales5)), 1) = "F" Then
'            If m_SalesST15 <> GetCuSales(mHC(27), oStrCuSales5) And Trim(mCP(13)) <> "" And mHC(27) <> "" Then
'               If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mHC(27), oStrCuSales5)), 1) = "F" Then
'            'end 2019/02/14
'                  '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'               Else
'                  oMailCount = oMailCount & oStrCuSales5 & ";"
'                  'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                  If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales5), 1) = "S" And _
'                     InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                     oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                  End If
'                  '2023/11/7 END
'                  oContext = oContext & vbCrLf + "«È¤á½s¸¹5¡G " + GetCustomerName(mHC(27)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales5)
'               End If
'             Else
'                   If Trim(mCP(13)) <> "" And mHC(27) <> "" Then
'                       IsMail = False
'                   End If
'            End If
'            'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'            If m_SalesST06 <> "" And Trim(mHC(27)) <> "" And Trim(mCP(13)) <> "" Then
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, mHC(27), Trim(mCP(13)), m_SalesST15, m_SalesST06, , mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                   IsMail = False
'               End If
'            End If
'         End If 'Added by Lydia 2022/11/03
'
'         'Modified by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¤¶²Ð«È¤á¬°ÂÂ«È¤á¦ý»P¤¶²Ð¤H¤£¦P°Ï®ÉµoMail³qª¾¬ÛÃö¤H­û
'         'Modified by Lydia 2022/11/03 ¸É®×·½¤£ºâ And InStr(pType, "LOS®×·½¦¬¤å") > 0
'         If strSrvDate(1) >= ªk«ß©Ò®×·½¦¬¤å±Ò¥Î¤é And InStr(pType, "LOS®×·½¦¬¤å") > 0 And m_LOS05 <> "" And m_LOS04_1 <> "" Then
'            oContext = "¥»©Ò®×¸¹¡G " + mHC(1) + "-" + mHC(2) + "-" + mHC(3) + "-" + mHC(4) + vbCrLf + "®×¥ó¦WºÙ¡G " + mHC(6) + vbCrLf + "¦¬¤å¤é¡G " + ChangeTStringToTDateString(mCP(5)) + vbCrLf + "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf
'            oMailCount = ""
'            'Modified by  Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¤¶²Ð«È¤á¬°ÂÂ«È¤á¦ý»P¤¶²Ð¤H¤£¦P°Ï®ÉµoMail³qª¾¬ÛÃö¤H­û
'            If mHC(5) <> "" Then
'                If ChkSameCuArea(Trim(mHC(5)), m_LOS04_1) = False Then
'                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(mHC(5), oStrCuSales1)), 1) = "F" Then
'                        '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'                    Else
'                        oMailCount = oMailCount & oStrCuSales1 & ";"
'                        'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                        If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales1), 1) = "S" And _
'                           InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                           oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                        End If
'                        '2023/11/7 END
'                        oContext = oContext & vbCrLf + "«È¤á½s¸¹1¡G " + GetCustomerName(mHC(5)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales1)
'                    End If
'                Else
'                       IsMail = False
'                End If
'                'ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á
'                'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, mHC(5), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06, _
'                        IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                    IsMail = False
'                End If
'            End If
'
'            If mHC(24) <> "" Then
'                If ChkSameCuArea(Trim(mHC(24)), m_LOS04_1) = False Then
'                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(mHC(24), oStrCuSales2)), 1) = "F" Then
'                        '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'                    Else
'                        oMailCount = oMailCount & oStrCuSales2 & ";"
'                        'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                        If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales2), 1) = "S" And _
'                           InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                           oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                        End If
'                        '2023/11/7 END
'                        oContext = oContext & vbCrLf + "«È¤á½s¸¹2 " + GetCustomerName(mHC(24)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales2)
'                    End If
'                Else
'                       IsMail = False
'                End If
'                'ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á
'                'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, mHC(24), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06, _
'                        IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                    IsMail = False
'                End If
'            End If
'
'            If mHC(25) <> "" Then
'                If ChkSameCuArea(Trim(mHC(25)), m_LOS04_1) = False Then
'                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(mHC(25), oStrCuSales3)), 1) = "F" Then
'                        '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'                    Else
'                        oMailCount = oMailCount & oStrCuSales3 & ";"
'                        'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                        If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales3), 1) = "S" And _
'                           InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                           oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                        End If
'                        '2023/11/7 END
'                        oContext = oContext & vbCrLf + "«È¤á½s¸¹3¡G " + GetCustomerName(mHC(25)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales3)
'                    End If
'                Else
'                       IsMail = False
'                End If
'                'ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á
'                'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, mHC(25), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06, _
'                        IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                    IsMail = False
'                End If
'            End If
'
'            If mHC(26) <> "" Then
'                If ChkSameCuArea(Trim(mHC(26)), m_LOS04_1) = False Then
'                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(mHC(26), oStrCuSales4)), 1) = "F" Then
'                        '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'                    Else
'                        oMailCount = oMailCount & oStrCuSales4 & ";"
'                        'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                        If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales4), 1) = "S" And _
'                           InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                           oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                        End If
'                        '2023/11/7 END
'                        oContext = oContext & vbCrLf + "«È¤á½s¸¹4¡G " + GetCustomerName(mHC(26)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales4)
'                    End If
'                Else
'                       IsMail = False
'                End If
'                'ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á
'                'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, mHC(26), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06, _
'                        IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                    IsMail = False
'                End If
'            End If
'
'            If mHC(27) <> "" Then
'                If ChkSameCuArea(Trim(mHC(27)), m_LOS04_1) = False Then
'                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(mHC(27), oStrCuSales5)), 1) = "F" Then
'                        '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'                    Else
'                        oMailCount = oMailCount & oStrCuSales5 & ";"
'                        'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                        If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales5), 1) = "S" And _
'                           InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                           oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                        End If
'                        '2023/11/7 END
'                        oContext = oContext & vbCrLf + "«È¤á½s¸¹5¡G " + GetCustomerName(mHC(27)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales5)
'                    End If
'                Else
'                       IsMail = False
'                End If
'                'ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á
'                'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, mHC(27), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06, _
'                        IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                    IsMail = False
'                End If
'            End If
'            'end 2020/05/20
'
'            '­Y¥Ó½Ð¤H¥þªÅ¥Õ¡A¤£µo
'            If IsMail = False Or (Trim(mHC(5)) = "" And Trim(mHC(24)) = "" And Trim(mHC(25)) = "" And Trim(mHC(26)) = "" And Trim(mHC(27)) = "") Then
'                 oMailCount = ""
'            End If
'         End If 'Added by Lydia 2020/03/24
'
'         'mhc(1)¥u§PÂ_1½X,¦]¬°FG
'         If UCase(Mid(mHC(1), 1, 1)) <> "F" And oMailCount <> "" Then
'            '¥Ó½Ð¤H¬° X65299 ©Î X03072 ªº©Ò¦³Ãö«Y¥ø·~³£¤£ÀË¬d·~°È°Ï
'            If Left(Trim(mHC(5)), 6) <> "X65299" And Left(Trim(mHC(5)), 6) <> "X03072" And _
'               Left(Trim(mHC(24)), 6) <> "X65299" And Left(Trim(mHC(24)), 6) <> "X03072" And _
'               Left(Trim(mHC(25)), 6) <> "X65299" And Left(Trim(mHC(25)), 6) <> "X03072" And _
'               Left(Trim(mHC(26)), 6) <> "X65299" And Left(Trim(mHC(26)), 6) <> "X03072" And _
'               Left(Trim(mHC(27)), 6) <> "X65299" And Left(Trim(mHC(27)), 6) <> "X03072" Then
'
'                '¥[µo¨q¬Â
'               'Modified by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¤¶²Ð«È¤á¬°ÂÂ«È¤á¦ý»P¤¶²Ð¤H¤£¦P°Ï®ÉµoMail³qª¾¬ÛÃö¤H­û
'               'Added by Lydia 2022/11/03 «D®×·½(¸É®×·½¤£ºâ)§ï¥Îµe­±§PÂ_
'               If InStr(pType, "LOS®×·½¦¬¤å") = 0 Or m_LOS05 = "" Or m_LOS04_1 = "" Then
'                    If UCase(pFormName) <> UCase("frm090801_New") Then
'                       MsgBox "¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¤£¦P·~°È°Ï¡A·Ç³Æµo mail ¡I", , "ª`·N¡I"
'                    End If
'                    oMailCount = oMailCount & Trim(mCP(13)) & ";" & Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'                    oMailCount = oMailCount & PUB_ChkForLawMan(mHC(5), mHC(1), mHC(2), mHC(3), mHC(4))
'                    oContext = oContext & vbCrLf + "¦¬¤å´¼Åv¤H­û¡G " + GetStaffName(mCP(13)) + vbCrLf + vbCrLf + "´¼Åv¤H­û(°Ï)¤£¦P¡I"
'               Else
'               'end 2022/11/03
'                   'Modify By Sindy 2022/10/14
'                   If UCase(pFormName) <> UCase("frm090801_New") Then
'                   '2022/9/27 END
'                      MsgBox "®×·½¤¶²Ð¤H­û»P«È¤á´¼Åv¤H­û¤£¦P·~°È°Ï¡A·Ç³Æµo mail ¡I", , "ª`·N¡I"
'                   End If
'                   'Modified by Lydia 2022/07/15 ³qª¾ªk«ß©Òªº´¼Åv¤H­û¨S¦³·N¸q¡AÀ³¸Ó­n§ï¬°®×·½¤¶²Ð¤H­û. ex.L-006547
'                   oMailCount = oMailCount & m_LOS04_1 & ";" & Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'                   oContext = oContext & vbCrLf + "®×·½¤¶²Ð¤H­û¡G " + GetStaffName(m_LOS04_1) + vbCrLf + vbCrLf + "´¼Åv¤H­û(°Ï)¤£¦P¡I"
'               End If 'Added by Lydia 2022/11/03
''                  PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å³qª¾--¦¹®×¦¬¤å«D­ì´¼Åv¤H­û(°Ï)¡I", oContext
'               'Modify By Sindy 2022/9/29
'               'Modify By Sindy 2023/3/27 +,mc13
'               mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
'                  " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'                  ",'" & "®×¥ó¦¬¤å³qª¾--¦¹®×¦¬¤å«D­ì´¼Åv¤H­û(°Ï)¡I(¤å¸¹:" & mCP(9) & ")','" & oContext & "',null,'" & mCP(9) & "')"
'               cnnConnection.Execute mStrSql
'               '2022/9/29 END
'            End If
'         End If
'2024/11/6 mark END

   End If
   
   'Add By Sindy 2022/9/27
   If UCase(pFormName) = UCase("frm090801_New") Then
      If ERecvSaveProgress(mCP, mHC, Pub_GetSpecMan("L®×¨ú®ø³¬¨÷³qª¾¤H­û"), oContext) = False Then
         GoTo ErrHand
      End If
   End If
   'Added by Lydia 2024/01/15 ÀË¬d©|¦bÅU°Ý¸u¥ô´Á¶¡ªº«È¤á¬O§_¤w³]©wÅU°Ý±M¥Î«H½c
   If intModifyKind = 0 And mCP(1) = "LA" And mCP(10) = "0" Then
      strTmp1(0) = Pub_GetChkCU199("1", mCP(9), True)
   End If
   'end 2024/01/15
   
   Exit Function
   
ErrHand:
   PUB_SaveFrm010006 = False 'Add By Sindy 2022/10/25
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical, "PUB_SaveFrm010006"
   End If
End Function

'Added by Lydia 2022/09/14 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G­×§ïÅU°Ý¸u¥ô¸ê®Æ®w(±qfrm010006.UpdateHireDatabase©â¥X¨Ó)
Private Function UpdateHireCaseDB(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intChoose As Integer, _
                ByRef mHC() As String, ByRef mCP() As String, ByVal mCU30 As String, Optional ByRef IsSaveData As Boolean, Optional ByVal pType As String, Optional ByVal pCaseNo As String) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
Dim adoquery As New ADODB.Recordset
'ªk«ß©Ò®×·½¦¬¤å
Dim m_LOS02 As String '®×·½®×¥óÃþ«¬
Dim m_LOS15 As String '®×·½³æ¸¹

If IsSaveData = True Then
    Exit Function
End If
IsSaveData = True

On Error GoTo ErrHand

'*********¯S®íºÞ¨îªºÅÜ¼Æ*************
    If pType = "LOS®×·½¦¬¤å" And pCaseNo <> "" Then
        m_LOS02 = Mid(pCaseNo, 1, InStr(pCaseNo, ",") - 1) '®×·½®×¥óÃþ«¬
        m_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1) '®×·½³æ¸¹
    End If
'***********************************

'Add By Sindy 2022/9/27
If UCase(pFormName) <> UCase("frm090801_New") Then
'2022/9/27 END
   cnnConnection.BeginTrans
End If
'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
'Add by Morgan 2008/8/5 (¨Ö¤J) +HC23
mStrSql = "update hirecase set hc05=" + CNULL(mHC(5)) + ",hc06=" + CNULL(mHC(6)) + ",hc07=" + CNULL(mHC(7)) + ",hc24=" + CNULL(mHC(24)) + ",hc25=" + CNULL(mHC(25)) + ",hc26=" + CNULL(mHC(26)) + ",hc27=" + CNULL(mHC(27)) + ",HC23=" + CNULL(mHC(23))
mStrSql = mStrSql + " where hc01=" + CNULL(mHC(1)) + " and hc02=" + CNULL(mHC(2)) + " and hc03=" + CNULL(mHC(3)) + " and hc04=" + CNULL(mHC(4))
cnnConnection.Execute mStrSql

'Modify By Sindy 2012/11/06 +CP150 ¦³¡¹¡¹ªºÀ³¦¬±b´ÚÃ±®Ö±±ºÞ
mStrSql = "update caseprogress set cp05=" + CNULL(mCP(5)) + ",cp10=" + CNULL(mCP(10)) + ",cp11=" + CNULL(mCP(11)) + ",cp53=" + CNULL(mCP(53)) + ",cp54=" + CNULL(mCP(54)) + ",cp13=" + CNULL(mCP(13)) + _
   ",cp14=" + CNULL(mCP(14)) + ",cp16=" + CNULL(mCP(16)) + ",cp32=" + CNULL(mCP(32)) + ",cp18=" & CNULL(IIf(Val(mCP(16)) / 1000 = 0, "", Val(mCP(16)) / 1000)) & ",cp150=" & CNULL(mCP(150)) & " where cp09=" + CNULL(mCP(9))
cnnConnection.Execute mStrSql
mStrSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(mCP(13)) + ") where cp09=" + CNULL(mCP(9))
cnnConnection.Execute mStrSql

        'Add By nickc 2007/08/21
        '­Y¬°±µ¬¢°O¿ý³æ(Âd¥x¦¬¤å)
        'Modify by Morgan 2007/10/26 ¶O¥Î¥i§ï®É¤~°µ¡A§_«h¤w¦¬´Ú¸ê®Æ·|³QÁÙ­ì
        If intChoose = 0 And mCP(60) = "" Then 'mCP(60) = "" =>  txtAdviser(9).Enabled = True
        'end 2007/10/26
            '¥¼¦¬ª÷ÃB = ¶O¥Î
            mStrSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
        End If
        
'Added by Lydia 2022/11/29 «D¤º³¡¦¬¤å¨Ã¥B¦³¶O¥Î¡A¥ý²Î¤@³]©wCP20=Null ;
If intChoose = 0 And Val(mCP(16)) > 0 Then
    mStrSql = "update caseprogress set cp20=null where cp09=" + CNULL(mCP(9))
    cnnConnection.Execute mStrSql
End If
'end 2022/11/29
'Add By Cheng 2002/05/10 ­Y¬°¤º³¡¦¬¤å§@·~®É, ®×¥ó¶i«×ÀÉªº¬O§_¦V«È¤á¦¬´Ú³]©w¬°"N"
If intChoose = 1 Then
   mStrSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(mCP(9))
   cnnConnection.Execute mStrSql
End If

'Modify By Sindy 2023/11/3 mark:¦¬¤å¤w¤£»Ý°õ¦æ¦¹¬qµ{¦¡,¦]±µ¬¢³æ¤w·|¦^¼g¶l»¼°Ï¸¹
'mStrSql = "update customer set cu30=" + CNULL(mCU30) + " where cu01=" + CNULL(Mid(mHC(5), 1, 8)) + " and cu02=" + CNULL(Mid(mHC(5), 9, 1))
'cnnConnection.Execute mStrSql
'2023/11/3 END

adoquery.CursorLocation = adUseClient
adoquery.Open "select np01 from nextprogress where np02 = '" & mHC(1) & "' and np03 = '" & mHC(2) & "' and np04 = '" & mHC(3) & "' and np05 = '" & mHC(4) & "' and np06 is null and np07 = '" & mCP(10) & "'", cnnConnection, adOpenStatic, adLockReadOnly
'Modify By Cheng 2002/05/10
'­Y¦b¤U¤@µ{§ÇÀÉ¥u§ì¨ì¤@µ§¸ê®Æ®É, ¤~­n§ì¤U¤@µ{§ÇÀÉªºÁ`¦¬¤å¸¹§ó·s®×¥ó¶i«×ÀÉªº¬ÛÃöÁ`¦¬¤å¸¹
If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
   If IsNull(adoquery.Fields(0).Value) = False Then
      cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & mCP(9) & "'"
   End If
End If
adoquery.Close

'add by nickc 2008/05/02 Àx¦s¹w©w¦¬´Ú¤é
'Remove by Lydia 2018/08/22 (À³¦¬±b´ÚºÞ±±)¨ú®ø¹w©w¦¬´Ú¤é,§ï¦¨¥I´Ú¶g´Á
'Dim rtCnt As Integer
''Modify by Morgan 2010/12/9
''If txtAdviser(13) <> "" Then
''    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " ", rtCnt
'If txtAdviser(13) <> "" And txtAdviser(13) <> txtAdviser(13).Tag Then
'    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0) + 1,'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
''end 2010/12/9
'    If rtCnt = 0 Then
'        cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " from dual "
'    End If
'End If
'end 2018/08/22

If m_LOS15 <> "" Then PUB_UpdateTTFee m_LOS15 'Added by Morgan 2022/4/14

'Add By Sindy 2022/9/27
If UCase(pFormName) <> UCase("frm090801_New") Then
'2022/9/27 END
   cnnConnection.CommitTrans
End If

UpdateHireCaseDB = True
Set adoquery = Nothing
Exit Function

ErrHand:
'Add By Sindy 2022/9/27
If UCase(pFormName) <> UCase("frm090801_New") Then
'2022/9/27 END
   cnnConnection.RollbackTrans
End If
ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
IsSaveData = False
Set adoquery = Nothing
End Function

'Added by Lydia 2022/09/14 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G·s¼WÅU°Ý¸u¥ô¦Ü¸ê®Æ®w(±qfrm010006.InsertHireDatabase©â¥X¨Ó)
Private Function InsertHireCaseDB(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intChoose As Integer, _
                ByRef mHC() As String, ByRef mCP() As String, ByVal mCU30 As String, Optional ByRef IsSaveData As Boolean, Optional ByVal pType As String, Optional ByVal pCaseNo As String, Optional ByRef RetVal As String) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
Dim strAutoNumber As String, bolError As Boolean
Dim adoquery As New ADODB.Recordset
Dim rsRD As New ADODB.Recordset
'ªk«ß©Ò®×·½¦¬¤å
Dim m_LOS01 As String '®×·½Á`¦¬¤å¸¹
Dim m_LOS01cp01 As String, m_LOS01cp02 As String, m_LOS01cp03 As String, m_LOS01cp04 As String '®×·½Á`¦¬¤å¸¹¤§¥»©Ò®×¸¹
Dim m_LOS02 As String '®×·½®×¥óÃþ«¬
Dim m_LOS15 As String '®×·½³æ¸¹
Dim m_LOS04 As String  '¤¶²Ð¤H
Dim m_LOS04_1 As String, m_LOS04_1st15 As String, m_LOS04_1st06 As String '¤¶²Ð¤H(²Ä¤@¦ì)¡B¦¬¤å³¡ªù¡B©Ò§O
Dim m_LOS05 As String  '¤¶²Ð«È¤á
Dim m_LOS12 As String  '¤¶²Ð¤é
Dim m_Los04_N1 As String, m_Los05_N As String  'LA¸É®×·½¤§¤¶²Ð¤H(²Ä¤@¦ì), ¤¶²Ð«È¤á
Dim strCaseNo As String, strCRL01 As String 'Add By Sindy 2023/1/17

If IsSaveData = True Then
    Exit Function
End If
IsSaveData = True

On Error GoTo ErrHand
'¶Ç¤J0¬°­«½Æ¤§¥»©Ò®×¸¹(·s¼WÂÂ®×)¡A1¬°¥¿½T¤§¥»©Ò®×¸¹(·s¼W·s®×)

'*********¯S®íºÞ¨îªºÅÜ¼Æ*************
    If pType = "LOS®×·½¦¬¤å" And pCaseNo <> "" Then
        m_LOS02 = Mid(pCaseNo, 1, InStr(pCaseNo, ",") - 1) '®×·½®×¥óÃþ«¬
        m_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1) '®×·½³æ¸¹
        strTmp1(0) = "select X.*,cp01,cp02,cp03,cp04 from LawOfficeSource X,caseprogress where los15=" & CNULL(m_LOS15) & " and los01=cp09(+) "
        intJ = 1
        Set rsRD = ClsLawReadRstMsg(intJ, strTmp1(0))
        If intJ = 1 Then
          '®×·½Á`¦¬¤å¸¹
          m_LOS01 = "" & rsRD.Fields("LOS01")
          '®×·½Á`¦¬¤å¸¹¤§¥»©Ò®×¸¹
          m_LOS01cp01 = "" & rsRD.Fields("cp01")
          m_LOS01cp02 = "" & rsRD.Fields("cp02")
          m_LOS01cp03 = "" & rsRD.Fields("cp03")
          m_LOS01cp04 = "" & rsRD.Fields("cp04")
          '(­ì)®×·½®×¥óÃþ«¬
          m_LOS02 = "" & rsRD.Fields("LOS02")
          '®×·½³æ¸¹
          m_LOS15 = "" & rsRD.Fields("LOS15")
          '¤¶²Ð¤H, ¤¶²Ð¤H(²Ä¤@¦ì)
          m_LOS04 = "" & rsRD.Fields("LOS04")
          If m_LOS04 <> "" And InStr(m_LOS04, ",") > 0 Then
             m_LOS04_1 = Mid(m_LOS04, 1, InStr(m_LOS04, ",") - 1)
          Else
             m_LOS04_1 = m_LOS04
          End If
          If m_LOS04_1 <> "" Then
             m_LOS04_1st15 = GetST15(m_LOS04_1, , , m_LOS04_1st06)
          End If
          '(­ì)¤¶²Ð«È¤á:
          m_LOS05 = "" & rsRD.Fields("LOS05")
          '¤¶²Ð¤é
          m_LOS12 = "" & rsRD.Fields("LOS12")
        End If
    End If
'***********************************
RetVal = "" '¦^¶Ç­È

'Add By Sindy 2022/9/27
If UCase(pFormName) <> UCase("frm090801_New") Then
'2022/9/27 END
   cnnConnection.BeginTrans
End If
If intSaveMode = 1 Then
   If mHC(2) = "" Then
      If ClsPDGetAutoNumber(mHC(1), strAutoNumber, True, False) Then
         mHC(2) = strAutoNumber
      Else
         bolError = True
      End If
   End If
   If bolError = False Then
      mCP(2) = mHC(2)
      'Modify by Morgan 2008/8/5 +HC23
      'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
      mStrSql = "insert into hirecase (hc01,hc02,hc03,hc04,hc05,hc06, hc07,hc23,hc24,hc25,hc26,hc27) values (" + _
          CNULL(mHC(1)) + "," + CNULL(mHC(2)) + "," + CNULL(mHC(3)) + "," + CNULL(mHC(4)) + "," + _
          CNULL(mHC(5)) + "," + CNULL(mHC(6)) + "," + CNULL(mHC(7)) + "," + CNULL(mHC(23)) + "," + _
          CNULL(mHC(24)) + "," + CNULL(mHC(25)) + "," + CNULL(mHC(26)) + "," + CNULL(mHC(27)) + ")"
      cnnConnection.Execute mStrSql
      mCP(31) = "Y"
   Else
      bolError = True
   End If
End If
If bolError = False Then
   If ClsPDGetAutoNumber(Left(mCP(9), 1), strAutoNumber, True, True) Then
      'add by nick 2005/01/07
      If mHC(7) <> "" Then
           mStrSql = "Update hirecase Set hc07='" & ChgSQL(mHC(7)) & "' Where hc01=" + CNULL(mHC(1)) + " and hc02=" + CNULL(mHC(2)) + " and hc03=" + CNULL(mHC(3)) + " and hc04=" + CNULL(mHC(4))
           cnnConnection.Execute mStrSql
      End If
       'add by nick 2004/10/29
       If mCP(10) = ÅU°Ý¸u¥ô Then
            mStrSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp01=" & CNULL(mHC(1)) & " and cp02=" & CNULL(mHC(2)) & " and cp03=" & CNULL(mHC(3)) & " and cp04=" & CNULL(mHC(4)) & " and cp27 is null and cp57 is null "
            cnnConnection.Execute mStrSql
        End If
      mCP(9) = mCP(9) + strAutoNumber

      'Modify By Sindy 2012/11/06 +CP150 ¦³¡¹¡¹ªºÀ³¦¬±b´ÚÃ±®Ö±±ºÞ
      'Modify By Sindy 2022/9/28 +,cp140
      mStrSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp14,cp53,cp54,cp13,cp16," + _
           "cp31,cp32,cp18,cp150,cp140) values (" + CNULL(mHC(1)) + "," + CNULL(mHC(2)) + "," + CNULL(mHC(3)) + "," + CNULL(mHC(4)) + "," + CNULL(mCP(5)) + "," + _
           CNULL(mCP(9)) + "," + CNULL(mCP(10)) + "," + CNULL(mCP(11)) + "," + CNULL(mCP(14)) + "," + CNULL(mCP(53)) + "," + CNULL(mCP(54)) + "," + CNULL(mCP(13)) + "," + CNULL(mCP(16)) + "," + _
           CNULL(mCP(31)) + "," + CNULL(mCP(32)) + "," + CNULL(IIf(Val(mCP(16)) / 1000 = 0, "", Val(mCP(16)) / 1000)) & "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + ")"
      cnnConnection.Execute mStrSql
      mStrSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(mCP(13)) + ") where cp09=" + CNULL(mCP(9))
      cnnConnection.Execute mStrSql
        
        '­Y¬°±µ¬¢°O¿ý³æ(Âd¥x¦¬¤å)
        'Modify by Morgan 2007/10/26 ¶O¥Î¥i§ï®É¤~°µ¡A§_«h¤w¦¬´Ú¸ê®Æ·|³QÁÙ­ì
        If intChoose = 0 And mCP(60) = "" Then 'mCP(60) = "" =>  txtAdviser(9).Enabled = True
        'end 2007/10/26
            '¥¼¦¬ª÷ÃB = ¶O¥Î
            mStrSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
        End If
        
      'Added by Lydia 2022/11/29 «D¤º³¡¦¬¤å¨Ã¥B¦³¶O¥Î¡A¥ý²Î¤@³]©wCP20=Null ;
      If intChoose = 0 And Val(mCP(16)) > 0 Then
          mStrSql = "update caseprogress set cp20=null where cp09=" + CNULL(mCP(9))
          cnnConnection.Execute mStrSql
      End If
      'end 2022/11/29
      'Add By Cheng 2002/05/10
      '­Y¬°¤º³¡¦¬¤å§@·~®É, ®×¥ó¶i«×ÀÉªº¬O§_¦V«È¤á¦¬´Ú³]©w¬°"N"
      If intChoose = 1 Then
         mStrSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(mCP(9))
         cnnConnection.Execute mStrSql
      End If
      
      'Modify By Sindy 2023/11/3 mark:¦¬¤å¤w¤£»Ý°õ¦æ¦¹¬qµ{¦¡,¦]±µ¬¢³æ¤w·|¦^¼g¶l»¼°Ï¸¹
'      mStrSql = "update customer set cu30=" + CNULL(mCU30) + " where cu01=" + CNULL(Mid(mHC(5), 1, 8)) + " and cu02=" + CNULL(Mid(mHC(5), 9, 1))
'      cnnConnection.Execute mStrSql
      '2023/11/3 END
      
      'Add By Sindy 2011/3/17 ¬°ÅU°Ý¸u¥ô(cp10=0)¥B¸u¥ô´Á¶¡>¨t²Î¤é®É,­Ycu153¬°null®É«h§ó·s¬°Y,¬°NªÌ¤£¥i§ó·s
      If mCP(10) = ÅU°Ý¸u¥ô And Val(mCP(54)) > Val(strSrvDate(2)) Then
         mStrSql = "update customer set cu153='Y' where cu01='" + Mid(mHC(5), 1, 8) + "' and cu02='" + Mid(mHC(5), 9, 1) + "' and cu153 is null"
         cnnConnection.Execute mStrSql
         mStrSql = "update potcustcont set pcc23='Y' where pcc01='" + Mid(mHC(5), 1, 8) + "' and pcc23 is null"
         cnnConnection.Execute mStrSql
      End If
      
      'Add By Sindy 2024/1/4 ¹q¤l¦¬¤å¤£¶·°õ¦æ¦¹¨ç¼Æ
      If UCase(pFormName) <> UCase("frm090801_New") Then
      '2024/1/4 END
         If Cls001SetCaseProgressFee(mHC(1), ¥xÆW°ê®a¥N¸¹, ÅU°Ý¸u¥ô, mCP(9)) = False Then bolError = True
      End If
   Else
      bolError = True
   End If
End If
adoquery.CursorLocation = adUseClient
adoquery.Open "select np01 from nextprogress where np02 = '" & mHC(1) & "' and np03 = '" & mHC(2) & "' and np04 = '" & mHC(3) & "' and np05 = '" & mHC(4) & "' and np06 is null and np07 = '" & mCP(10) & "'", cnnConnection, adOpenStatic, adLockReadOnly
'Modify By Cheng 2002/05/10
'­Y¦b¤U¤@µ{§ÇÀÉ¥u§ì¨ì¤@µ§¸ê®Æ®É, ¤~­n§ì¤U¤@µ{§ÇÀÉªºÁ`¦¬¤å¸¹§ó·s®×¥ó¶i«×ÀÉªº¬ÛÃöÁ`¦¬¤å¸¹
If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
   If IsNull(adoquery.Fields(0).Value) = False Then
      cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & mCP(9) & "'"
   End If
End If
adoquery.Close
'add by nickc 2008/05/02 Àx¦s¹w©w¦¬´Ú¤é
'Remove by Lydia 2018/08/22 (À³¦¬±b´ÚºÞ±±)¨ú®ø¹w©w¦¬´Ú¤é,§ï¦¨¥I´Ú¶g´Á
'If bolError = False Then
'   Dim rtCnt As Integer
'   'Modify by Morgan 2010/12/9
'   'If txtAdviser(13) <> "" Then
'   '    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " ", rtCnt
'   If txtAdviser(13) <> "" And txtAdviser(13) <> txtAdviser(13).Tag Then
'         cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0) + 1,'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
'   'end 2010/12/9
'         If rtCnt = 0 Then
'             cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtAdviser(13)) & " from dual "
'         End If
'   End If
'End If
'end 2018/08/22

    'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G¦sÀÉ®É®×·½³æ¸¹¦sCP162¡B®×·½Á`¦¬¤å¸¹(LOS01)¦sCP64Äæ"®×·½¡G¥»©Ò®×¸¹(Á`¦¬¤å¸¹)
    If strSrvDate(1) >= ªk«ß©Ò®×·½¦¬¤å±Ò¥Î¤é And mHC(1) = "LA" And m_LOS15 <> "" Then
        mStrSql = ""
        If m_LOS01 <> "" Then mStrSql = ",cp64=" & CNULL("®×·½¡G" & m_LOS01cp01 & "-" & m_LOS01cp02 & IIf(m_LOS01cp03 <> "0", "-" & m_LOS01cp03, "") & IIf(m_LOS01cp04 <> "00", "-" & m_LOS01cp04, "") & "(" & m_LOS01 & ");")
        mStrSql = "update caseprogress set CP162=" & CNULL(m_LOS15) & mStrSql & " where cp09=" & CNULL(mCP(9))
        cnnConnection.Execute mStrSql, intJ
       
        '¨Ã¦^¼g¦¬¤å¸¹¦Ü®×·½ÀÉªºªk«ß©ÒÁ`¦¬¤å¸¹Äæ¡C
        '5/26 ­Y¿é¤J¤§®×·½³æ¸¹¤w¦³ªk«ß©ÒÁ`¦¬¤å¸¹¥B¬°¦P®×¸¹¦P¤é¦¬¤åªÌ¡A«h¬°¦P¤@±µ¬¢³æ¤§¨ä¥L©Ê½è¡C
        mStrSql = "update LawOfficeSource set los06='" & mCP(9) & "' where los06 is null and los15=" & CNULL(m_LOS15)
        cnnConnection.Execute mStrSql, intJ
        
        'Add By Sindy 2023/1/17 ªk«ß©Ò¯È¥»¦¬¤å,Ã±®Öªº¹q¤lÀÉÂk¨÷
        strCaseNo = mCP(1) & mCP(2)
        If mCP(3) & mCP(4) <> "000" Then
            strCaseNo = strCaseNo & "-" & mCP(3)
        End If
        If mCP(4) <> "00" Then
            strCaseNo = strCaseNo & "-" & mCP(4)
        End If
        strTmp1(0) = "select los17 from LawOfficeSource where los17 is not null and los15=" & CNULL(m_LOS15)
        intJ = 1
        Set rsRD = ClsLawReadRstMsg(intJ, strTmp1(0))
        If intJ = 1 Then
            strCRL01 = rsRD.Fields("los17")
            '±µ¬¢³æ¹q¤lÀÉ§ó¦W-·s®×
            Call PUB_UpdCRLFileName(strCRL01)
            mStrSql = "update casepaperpdf set cpp01='" & mCP(9) & "',cpp10='X'" & _
                      ",cpp02='" & strCaseNo & "'||'." & mCP(10) & ".'||cpp02 where cpp11='" & strCRL01 & "'"
            cnnConnection.Execute mStrSql, intJ
        End If
        '2023/1/17 END
        
        '­Y®×·½¸ê®Æªº¤¶²Ð«È¤áLOS05¬°ªÅ®Éªí¥Ü·s«È¤á­n¦^¼g¨Ã§ó·s(¦¬¤å®É¿é¤Jªº)«È¤á´¼Åv¤H­û(CU12CU13)¬°¤¶²Ð¤H(LOS04²Ä¤@¤H)
        If m_LOS05 = "" And Trim(mHC(5) & mHC(24) & mHC(25) & mHC(26) & mHC(27)) <> "" Then
            '¨Ã¥B¦^¼g®×·½¤¶²Ð«È¤á½s¸¹LOS05
            mStrSql = "update LawOfficeSource set los05='" & mHC(5) & "' where los05 is null and los15=" & CNULL(m_LOS15)
            cnnConnection.Execute mStrSql, intJ
            If intJ > 0 Then
               strTmp1(1) = "5": strTmp1(2) = "24": strTmp1(3) = "25": strTmp1(4) = "26": strTmp1(5) = "27"
               For intJ = 1 To 5
                   If Trim(mHC(Val(strTmp1(intJ)))) <> "" Then
                        mStrSql = "update customer set cu12='" & m_LOS04_1st15 & "',cu13='" & m_LOS04_1 & "' where cu01='" & Left(mHC(Val(strTmp1(intJ))), 8) & "' and cu02='" & Right(mHC(Val(strTmp1(intJ))), 1) & "'"
                        Pub_SeekTbLog mStrSql
                        cnnConnection.Execute mStrSql
                   End If
               Next intJ
               'Added by Lydia 2022/11/10 «È¤á½s¸¹«á«Ø=m_LOS05=ªÅ¥Õ; ex.LA-003386
               m_LOS05 = mHC(5)
               RetVal = m_LOS05 & "|" & m_LOS04_1
               'end 2022/11/10
            End If
        End If
        '³Ì«á¤~°µ-->«È¤á½s¸¹¦^¼g«á¡A®×·½®×¥óÃþ«¬A¡A­YµLÂI¼Æ«h«O¯dÃþ«¬A¡A­Y¦³ÂI¼Æ«h§PÂ_¦P¤@«È¤á½s¸¹¤¶²Ð¤é«e­Y¦³A1«h¦¹µ§³]¬°A2¡A­YµL«h³]¬°A1¡C
                            '­pºâ®×·½¤§¶O¥Î¤ÎÂI¼Æ¡A§ó·s¦^®×·½Á`¦¬¤å¸¹LOS01¤§¶O¥Î¤ÎÂI¼Æ¡A¥H§Q´¼¼z©Ò¶}¥ß¦¬¾Ú¡C
                            '®×·½¬°TT-999999®É¦P®É¤Wµo¤å¤éCP27¬°¨t²Î¤é(¬°µLµo¤å¤éªÌ¤~§ó·s)¡C
                            '5/6¸ò·¨ºÊ¹î¤H½T»{°ê¥~³¡¤¶²Ð®×·½¥H¬Û¦P¤À¼í¤è¦¡­pºâ¡A¤£ºÞ°ê¥~¥N²z¤H¤´¥H«È¤á¬°¤¶²Ð°ò·Ç¡C
        If m_LOS02 = "A" And Val(mCP(16)) > 0 Then  '¶O¥Î§ï¬°ÂI¼Æ
           mStrSql = "select los02 from LawOfficeSource where los12<'" & m_LOS12 & "' and los02='A1' and los05='" & IIf(m_LOS05 <> "", m_LOS05, mHC(5)) & "' "
           intJ = 1
           Set rsRD = ClsLawReadRstMsg(intJ, mStrSql)
           If intJ = 1 Then
               mStrSql = "update LawOfficeSource set los02='A2' where los15='" & m_LOS15 & "' "
               cnnConnection.Execute mStrSql
           Else
               mStrSql = "update LawOfficeSource set los02='A1' where los15='" & m_LOS15 & "' "
               cnnConnection.Execute mStrSql
           End If
           '®×·½¬°TT-999999®É¦P®É¤Wµo¤å¤éCP27¬°¨t²Î¤é(¬°µLµo¤å¤éªÌ¤~§ó·s)¡C
           If m_LOS01cp01 & m_LOS01cp02 = "TT999999" Then
               mStrSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_LOS01 & "' and nvl(cp27,0)=0 "
               cnnConnection.Execute mStrSql
           End If
        End If
        
       '­pºâ®×·½¤§¶O¥Î¤ÎÂI¼Æ¡A§ó·s¦^®×·½Á`¦¬¤å¸¹LOS01¤§¶O¥Î¤ÎÂI¼Æ¡A¥H§Q´¼¼z©Ò¶}¥ß¦¬¾Ú¡C
       PUB_UpdateTTFee m_LOS15 'Added by Morgan 2020/9/29 ¦P®×·½³æ¸¹ªº¨C­Ó¦¬¤å©Ê½è³£­n(¶O¥Î¥[Á`)
    End If
    'end 2020/05/20
    
    'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¸Ó¦¬¤å¸¹ÂI¼Æ>0¦ýµL®×·½(¦Û¦æ¦¬¤åªÌ)®É¡A­Y®×¥óªº«È¤á¬°«Dªk«ß©Òªº«È¤á®É«h¤´ºâAÃþ®×·½(¥t¼g¨ç¼Æ°Ñ·Ó§@±b³W«h³]©w¬°A1~A4)¡C
                                       '¨t²Î¦Û°Ê·s¼WTT-999999®×¶i«×(BÃþ¦¬¤å)¤Îªk«ß©Ò®×·½¸ê®Æ(¦P³Ì«á¤@µ§®×·½ªº¸ê®Æ)¡C
    'Memo by Lydia 2020/10/05 (9/30) ­Y¸Ó¦¬¤å¸¹ÂI¼Æ>0¦ýµL®×·½(¦Û¦æ¦¬¤åªÌ)®É¡A­Y®×¥óªº«È¤á¬°«Dªk«ß©Òªº«È¤á®É«h¬°A3Ãþ®×·½¡A¤£½×·sÂÂ®×¡A¨t²Î¦Û°Ê·s¼WTT-999999®×¶i«×(BÃþ¦¬¤å)¤Îªk«ß©Ò®×·½¸ê®Æ¡C
    'Modified by Morgan 2021/1/8 ¥x¤@Ãö«Y¥ø·~ X03072 °£¥~
    If strSrvDate(1) >= ªk«ß©Ò®×·½¦¬¤å±Ò¥Î¤é And mHC(1) = "LA" And m_LOS15 = "" And Val(mCP(16)) > 0 And Left(mHC(5), 6) <> "X03072" Then
        mStrSql = "select cu01,cu02,st15,st01 from customer,staff where cu01='" & Mid(mHC(5), 1, 8) & "' and cu02='" & Mid(mHC(5), 9, 1) & "'  and cu13=st01(+) "
        intJ = 1
        Set rsRD = ClsLawReadRstMsg(intJ, mStrSql)
        If intJ = 1 Then
            strTmp1(1) = Left("" & rsRD.Fields("st15"), 1)
            If strTmp1(1) <> "L" Then
               '«Dªk«ß©ÒªºÂÂ«È¤á®É«h¤´ºâAÃþ®×·½(¥t¼g¨ç¼Æ°Ñ·Ó§@±b³W«h³]©w¬°A1~A4)
               'Modified by Lydia 2020/10/05 (9/30) ­Y¸Ó¦¬¤å¸¹ÂI¼Æ>0¦ýµL®×·½(¦Û¦æ¦¬¤åªÌ)®É¡A­Y®×¥óªº«È¤á¬°«Dªk«ß©Òªº«È¤á®É«h¬°A3Ãþ®×·½¡A¤£½×·sÂÂ®×¡A¨t²Î¦Û°Ê·s¼WTT-999999®×¶i«×(BÃþ¦¬¤å)¤Îªk«ß©Ò®×·½¸ê®Æ¡C
                   strTmp1(2) = "A3"   '®×·½Ãþ«¬
                   strTmp1(3) = "" & rsRD.Fields("st01")  'BÃþ¦¬¤å¤§´¼Åv¤H­û: ¤¶²Ð¤H²Ä¤@¤H
                   strTmp1(5) = strTmp1(3)  '¤¶²Ð¤H
                   If mHC(2) <> "" Then
                       '¦p¬°ÂÂ®×¥B´¿¦³A3Ãþ®×·½®É¡A¤¶²Ð¤H­û¦P³Ì«á¤@µ§®×·½¡A§_«h¥H«È¤á¥Ø«eªº´¼Åv¤H­û¬°¤¶²Ð¤H¡C
                       mStrSql = "Select * From Lawofficesource Where los02='A3' and los07||los08 is null and Los15 In " & _
                                   "(select max(cp162) from caseprogress where cp01='" & mHC(1) & "' and cp02='" & mHC(2) & "' and cp03='" & IIf(mHC(3) = "", "0", mHC(3)) & "' and cp04='" & IIf(mHC(4) = "", "00", mHC(4)) & "' and cp162 is not null) "
                       intJ = 1
                       Set adoquery = ClsLawReadRstMsg(intJ, mStrSql)
                       If intJ = 1 Then
                            '¥Î©ó®×·½¤§±µ¬¢¤H¨ú±o¦bÂ¾­û¤u½s¸¹©M¤¶²Ð¤H²Ä¤@¤H
                            strTmp1(5) = PUB_GetNowStaff("" & adoquery.Fields("los04"), strTmp1(3))
                       End If
                   End If
                   m_Los04_N1 = strTmp1(3)
                   If strTmp1(3) <> "" Then
                        strTmp1(1) = AutoNo("B", 6) 'TT¦¬¤å¸¹
              ' end 2020/10/05
                        'Modified by Morgan 2021/1/8 +cp20,cp27,cp32
                        'Modified by Lydia 2022/11/09 +CP27=¨t²Î¤é
                        mStrSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp11,cp12,cp13,cp20,cp32,CP162,CP27)" & _
                           " values('TT','999999','0','00'," & strSrvDate(1) & ",null ,'" & strTmp1(1) & "'" & _
                           ",'735','07','" & GetST15(strTmp1(3)) & "','" & strTmp1(3) & "','N','N',null, " & strSrvDate(1) & " )"
                        cnnConnection.Execute mStrSql
                        'ªk«ß©Ò®×·½¸ê®Æ(¦P³Ì«á¤@µ§®×·½ªº¸ê®Æ), ®×·½³æ¸¹=TTÁ`¦¬¤å¸¹
                        strTmp1(4) = AutoNo("LOS", 5, , True)
                        'Modified by Lydia 2020/10/05
                        mStrSql = "insert into LawOfficeSource(LOS01,LOS02,LOS03,LOS04,LOS05,LOS06,LOS10,LOS11,LOS12,LOS13,LOS15)" & _
                           " values ('" & strTmp1(1) & "','" & strTmp1(2) & "' ,'" & mCP(13) & "'" & _
                           ",'" & strTmp1(5) & "','" & mHC(5) & "','" & mCP(9) & "','" & strTmp1(1) & "'" & _
                           ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss'),'" & strTmp1(4) & "')"
                        m_Los05_N = mHC(5)
                        'end 2020/10/05
                        cnnConnection.Execute mStrSql
                        'Added by Lydia 2020/10/05 ¦¬¤å¤§¶i«×¥[µù®×·½
                        mStrSql = "Update CaseProgress Set cp64=" & CNULL("®×·½¡GTT-999999(" & strTmp1(1) & ");") & "||cp64, cp162='" & strTmp1(4) & "' where cp09=" & CNULL(mCP(9))
                        cnnConnection.Execute mStrSql
                        
                        '­pºâ®×·½¤§¶O¥Î¤ÎÂI¼Æ¡A§ó·s¦^®×·½Á`¦¬¤å¸¹LOS01¤§¶O¥Î¤ÎÂI¼Æ¡A¥H§Q´¼¼z©Ò¶}¥ß¦¬¾Ú¡C
                        PUB_UpdateTTFee strTmp1(4) 'Added by Morgan 2020/9/29
                        RetVal = m_Los05_N & "|" & m_Los04_N1
                   End If 'Added by Lydia 2020/10/05
               'End If 'Remove by Lydia 2020/10/05
            End If
        End If
    End If
    'end 2020/05/20

    
If bolError Then
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.RollbackTrans
   End If
   ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
   IsSaveData = False
Else
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.CommitTrans
   End If
   InsertHireCaseDB = True
   'txtCode(0) = mHC(2) '²¾¨ì¥~¼hÅÜ§ó
End If
'txtCode(0) = mHC(2)  '²¾¨ì¥~¼hÅÜ§ó
   Set adoquery = Nothing
   Set rsRD = Nothing
Exit Function

ErrHand:
'Add By Sindy 2022/9/27
If UCase(pFormName) <> UCase("frm090801_New") Then
'2022/9/27 END
   cnnConnection.RollbackTrans
End If
ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
IsSaveData = False
Set adoquery = Nothing
Set rsRD = Nothing
Resume
End Function

'Added by Lydia 2022/09/14 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡GACS®×¤§112´¼°]ÅU°Ý¦¬¤å(±qfrm010006_1.SaveDatabase©â¥X¨Ó)
'Modify By Sindy 2024/11/21 + , Optional ByVal m_intCRC As Integer = 0: ¦Û°Ê¦¬¤åªº®×¥ó©Ê½è¶¶§Ç
Public Function PUB_SaveFrm010006_1(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intCaseKind As Integer, ByVal intChoose As Integer, _
                ByRef mLC() As String, ByRef mCP() As String, ByVal mCU30 As String, Optional ByRef IsSaveData As Boolean, _
                Optional ByVal pType As String, Optional ByVal pCaseNo As String, Optional ByVal m_intCRC As Integer = 0) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intCaseKind¡A1¬°±M§Q¡A2¬°°Ó¼Ð¡A3¬°ªk°È¡A4¬°ÅU°Ý¡A5¬°±M§Q(ªA)¡A6¬°°Ó¼Ð(ªA)¡A7¬°ªk°È(ªA)¡A8¬°ÅU°Ý(ªA)
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
Dim m_SalesST15 As String, m_SalesST06 As String
Dim m_SalesDeptName As String
Dim m_CP10Name As String '¦¬¤å¤§®×¥ó©Ê½è¦WºÙ
Dim m_Na01Name As String '¥Ó½Ð°ê®a¦WºÙ
Dim oMailCount As String
   
   intJ = ClsPDGetCaseProperty(mLC(1), mCP(10), m_CP10Name)
   m_SalesST15 = GetST15(mCP(13), m_SalesDeptName, , m_SalesST06)
   'Added by Lydia 2023/05/11 ¦]¬°PUB_ReadCaseData·|¦^¶Ç6½X«È¤á½s¸¹,©Ò¥H¥ý²Î¤@«È¤á½s¸¹
   mLC(11) = ChangeCustomerL(mLC(11))
   mLC(43) = ChangeCustomerL(mLC(43))
   mLC(44) = ChangeCustomerL(mLC(44))
   mLC(45) = ChangeCustomerL(mLC(45))
   mLC(46) = ChangeCustomerL(mLC(46))
   'end 2023/05/11
     
   If intModifyKind = 0 Then
      PUB_SaveFrm010006_1 = InsertCaseDB(pFormName, intSaveMode, intModifyKind, intCaseKind, intChoose, mLC, mCP, mCU30, IsSaveData, pType, pCaseNo)
   Else
      PUB_SaveFrm010006_1 = UpdateCaseDB(pFormName, intSaveMode, intModifyKind, intCaseKind, intChoose, mLC, mCP, mCU30, IsSaveData, pType, pCaseNo)
   End If
   If PUB_SaveFrm010006_1 = False Then Exit Function '¦sÀÉ¥¢±Ñ,«áÄò¤£ÀË¬d
   '´ú¸Õ¸Ñ¨Mmail µo¤£¨ìªº®É­Ô·|¦s¨âµ§ªº¿ù»~
   On Error GoTo 0    'Âk¹s
   On Error GoTo ErrHand 'Add By Sindy 2022/9/29
   'Add By Sindy 2022/12/29 ­«ÅªCP,¦]«eÀYUpdate¨ç¼Æµ{¦¡¦³¥i¯àª½±µ¦sDB,¨S¦³§ó·scp³¯¦C­È
   strTmp1(0) = mCP(9)
   Erase mCP
   ReDim Preserve mCP(TF_CP) As String
   mCP(9) = strTmp1(0)
   'Modified by Lydia 2023/05/11 + false
   Call PUB_ReadCaseProgressDatabase(mCP(), 1, False)
   '2022/12/29 END
   
   m_Na01Name = PUB_GetNationName(mLC(15))
   If intModifyKind = 0 Then
      Dim oContext As String, strCaseNo As String
      Dim strTemp As String
      Dim m_strState As String
      'Add By Sindy 2023/8/18 ¤£±o¥N²zªº«áÄòÂÂ®×¦¬¤å±±ºÞ¡A³qª¾¦¬¤å¤H­û¡]CP13¡^
      If mLC(22) <> "" Then
        If GetAgentAndState(mLC(22), strTmp1(1), , , , mLC(1), m_strState, IIf(intSaveMode = 0, True, False)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥N²z¤H¡G " + mLC(22) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mLC(22)
          End If
        End If
      End If
      If mLC(11) <> "" Then
        If GetCustomerAndState(mLC(11), strTmp1(1), , , , mLC(1), m_strState, IIf(intSaveMode = 0, True, False), , mLC(2), mLC(3), mLC(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H1¡G " + mLC(11) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mLC(11)
          End If
        End If
      End If
      If mLC(43) <> "" Then
        If GetCustomerAndState(mLC(43), strTmp1(1), , , , mLC(1), m_strState, IIf(intSaveMode = 0, True, False), , mLC(2), mLC(3), mLC(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H2¡G " + mLC(43) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mLC(43)
          End If
        End If
      End If
      If mLC(44) <> "" Then
        If GetCustomerAndState(mLC(44), strTmp1(1), , , , mLC(1), m_strState, IIf(intSaveMode = 0, True, False), , mLC(2), mLC(3), mLC(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H3¡G " + mLC(44) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mLC(44)
          End If
        End If
      End If
      If mLC(45) <> "" Then
        If GetCustomerAndState(mLC(45), strTmp1(1), , , , mLC(1), m_strState, IIf(intSaveMode = 0, True, False), , mLC(2), mLC(3), mLC(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H4¡G " + mLC(45) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mLC(45)
          End If
        End If
      End If
      If mLC(46) <> "" Then
        If GetCustomerAndState(mLC(46), strTmp1(1), , , , mLC(1), m_strState, IIf(intSaveMode = 0, True, False), , mLC(2), mLC(3), mLC(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H5¡G " + mLC(46) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & mLC(46)
          End If
        End If
      End If
      If oContext <> "" Then
         strTemp = Mid(strTemp, 2)
         strCaseNo = IIf("-" + mLC(3) + "-" + mLC(4) = "-0-00", mLC(1) + "-" + mLC(2), mLC(1) + "-" + mLC(2) + "-" + mLC(3) + "-" + mLC(4))
         oContext = "¥»©Ò®×¸¹¡G " + strCaseNo + vbCrLf + _
                    "®×¥ó¦WºÙ¡G " + mLC(5) + vbCrLf + _
                    "¥Ó½Ð°ê®a¡G " + mLC(15) + " " + m_Na01Name + vbCrLf + _
                    "¦¬¤å¤é¡G " + ChangeTStringToTDateString(TransDate(mCP(5), 1)) + vbCrLf + _
                    "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf + vbCrLf + _
                    "¡i¤£±o¥N²z¡j" + vbCrLf + _
                    oContext
         oMailCount = mCP(13) & ";" & PUB_GetFCPProSup(mCP(13), True)
         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & strCaseNo & _
            " ¤w½T»{Äò¦æ¦¬¤å¡A½Ðª`·N¸Ó" & strTemp & "½s¸¹¤w³]¬°¤£±o¥N²z¡C(¤å¸¹:" & mCP(9) & ")','" & oContext & "',null,'" & mCP(9) & "')"
         cnnConnection.Execute mStrSql
      End If
      '2023/8/18 END
      
      oContext = "¥»©Ò®×¸¹¡G " + mLC(1) + "-" + mLC(2) + "-" + mLC(3) + "-" + mLC(4) + vbCrLf + "®×¥ó¦WºÙ¡G " + mLC(5) + vbCrLf + "¦¬¤å¤é¡G " + ChangeTStringToTDateString(mCP(5)) + vbCrLf + "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf
      'Modify By Sindy 2024/11/6 §ï¦¨¦@¥Î¨ç¼Æ: ¦¬¤å®É,ÀË¬d¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¬O§_¦³»~
      '§ï¼g­ì¥Ñ¬O¦]¬°¥Ó½Ð¤H1~5 ³v¤@ÀË¬d,¦³»~§¡­nµo mail
      'edit by nickc 2007/08/21 ­Y¥Ó½Ð¤H¥þªÅ¥Õ¡A¤£µo
      If Not (mLC(11) = "" And mLC(43) = "" And Trim(mLC(44)) = "" And Trim(mLC(45)) = "" And Trim(mLC(46)) = "") Then
         'Modify By Sindy 2024/11/21 ¦Û°Ê¦¬¤åªº®×¥ó©Ê½è¶¶§Ç=1 ©Î¯È¥»¦¬¤å¥¼«ü©w
         If m_intCRC = 1 Or m_intCRC = 0 Then
         '2024/11/21 END
            Call RecvChkApplCust("·í¨Æ¤H1", Trim(mLC(11)), mCP(13), "", m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
            Call RecvChkApplCust("·í¨Æ¤H2", Trim(mLC(43)), mCP(13), "", m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
            Call RecvChkApplCust("·í¨Æ¤H3", Trim(mLC(44)), mCP(13), "", m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
            Call RecvChkApplCust("·í¨Æ¤H4", Trim(mLC(45)), mCP(13), "", m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
            Call RecvChkApplCust("·í¨Æ¤H5", Trim(mLC(46)), mCP(13), "", m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
         End If
      End If
      '2024/11/6 END
'Modify By Sindy 2024/11/6 mark
'      '·í¦¬¤å·~°È°Ï»P«È¤áÀÉ·~°È°Ï¤£¦P®Éµo mail  ¤Î´£¥Ü
'      Dim oStrCuSales1 As String
'      Dim oStrCuSales2 As String
'      Dim oStrCuSales3 As String
'      Dim oStrCuSales4 As String
'      Dim oStrCuSales5 As String
'      '¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'      Dim IsMail As Boolean
'      IsMail = True
'
'      oStrCuSales1 = ""
'      oStrCuSales2 = ""
'      oStrCuSales3 = ""
'      oStrCuSales4 = ""
'      oStrCuSales5 = ""
'
'      oMailCount = ""
'      If m_SalesST15 <> GetCuSales(mLC(11), oStrCuSales1) And Trim(mCP(13)) <> "" And Trim(mLC(11)) <> "" Then
'         If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mLC(11), oStrCuSales1)), 1) = "F" Then
'            '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'         Else
'            oMailCount = oMailCount & oStrCuSales1 & ";"
'            'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'            If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales1), 1) = "S" And _
'               InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'               oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'            End If
'            '2023/11/7 END
'            oContext = oContext & vbCrLf + "·í¨Æ¤H1¡G " + GetCustomerName(mLC(11)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales1)
'         End If
'      '¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'      Else
'         If Trim(mCP(13)) <> "" And Trim(mLC(11)) <> "" Then
'             IsMail = False
'         End If
'      End If
'      'ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_SalesST06 <> "" And Trim(mLC(11)) <> "" And mCP(13) <> "" Then
'         'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'         'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'         If PUB_ChkOldCustomer(True, mLC(11), mCP(13), m_SalesST15, m_SalesST06, _
'                     IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'             IsMail = False
'         End If
'      End If
'
'      If m_SalesST15 <> GetCuSales(mLC(43), oStrCuSales2) And Trim(mCP(13)) <> "" And Trim(mLC(43)) <> "" Then
'         If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mLC(43), oStrCuSales2)), 1) = "F" Then
'            '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'         Else
'            oMailCount = oMailCount & oStrCuSales2 & ";"
'            'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'            If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales2), 1) = "S" And _
'               InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'               oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'            End If
'            '2023/11/7 END
'            oContext = oContext & vbCrLf + "·í¨Æ¤H2¡G " + GetCustomerName(mLC(43)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales2)
'         End If
'      '¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'      Else
'         If Trim(mCP(13)) <> "" And Trim(mLC(43)) <> "" Then
'             IsMail = False
'         End If
'      End If
'      If m_SalesST06 <> "" And Trim(mLC(43)) <> "" And mCP(13) <> "" Then
'         'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'         'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'         If PUB_ChkOldCustomer(True, mLC(43), mCP(13), m_SalesST15, m_SalesST06, _
'                  IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'             IsMail = False
'         End If
'      End If
'
'      If m_SalesST15 <> GetCuSales(mLC(44), oStrCuSales3) And Trim(mCP(13)) <> "" And Trim(mLC(44)) <> "" Then
'         If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mLC(44), oStrCuSales3)), 1) = "F" Then
'            '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'         Else
'            oMailCount = oMailCount & oStrCuSales3 & ";"
'            'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'            If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales3), 1) = "S" And _
'               InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'               oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'            End If
'            '2023/11/7 END
'            oContext = oContext & vbCrLf + "·í¨Æ¤H3¡G " + GetCustomerName(mLC(44)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales3)
'         End If
'       '¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'       Else
'             If Trim(mCP(13)) <> "" And Trim(mLC(44)) <> "" Then
'                 IsMail = False
'             End If
'      End If
'      'ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_SalesST06 <> "" And Trim(mLC(44)) <> "" And mCP(13) <> "" Then
'         'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'         'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'         If PUB_ChkOldCustomer(True, mLC(44), mCP(13), m_SalesST15, m_SalesST06, _
'                  IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'             IsMail = False
'         End If
'      End If
'
'      If m_SalesST15 <> GetCuSales(mLC(45), oStrCuSales4) And Trim(mCP(13)) <> "" And Trim(mLC(45)) <> "" Then
'         If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mLC(45), oStrCuSales4)), 1) = "F" Then
'            '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'         Else
'            oMailCount = oMailCount & oStrCuSales4 & ";"
'            'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'            If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales4), 1) = "S" And _
'               InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'               oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'            End If
'            '2023/11/7 END
'            oContext = oContext & vbCrLf + "·í¨Æ¤H4¡G " + GetCustomerName(mLC(45)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales4)
'         End If
'       Else
'             If Trim(mCP(13)) <> "" And Trim(mLC(45)) <> "" Then
'                 IsMail = False
'             End If
'      End If
'      'ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_SalesST06 <> "" And Trim(mLC(45)) <> "" And mCP(13) <> "" Then
'         'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'         'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'         If PUB_ChkOldCustomer(True, mLC(45), mCP(13), m_SalesST15, m_SalesST06, _
'                  IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'             IsMail = False
'         End If
'      End If
'
'      If m_SalesST15 <> GetCuSales(mLC(46), oStrCuSales5) And Trim(mCP(13)) <> "" And Trim(mLC(46)) <> "" Then
'         If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(mLC(46), oStrCuSales5)), 1) = "F" Then
'            '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'         Else
'            oMailCount = oMailCount & oStrCuSales5 & ";"
'            'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'            If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales5), 1) = "S" And _
'               InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'               oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'            End If
'            '2023/11/7 END
'            oContext = oContext & vbCrLf + "·í¨Æ¤H5¡G " + GetCustomerName(mLC(46)) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales5)
'         End If
'       Else
'             If Trim(mCP(13)) <> "" And Trim(mLC(46)) <> "" Then
'                 IsMail = False
'             End If
'      End If
'      'ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_SalesST06 <> "" And Trim(mLC(46)) <> "" And mCP(13) <> "" Then
'         'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'         'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'         If PUB_ChkOldCustomer(True, mLC(46), mCP(13), m_SalesST15, m_SalesST06, _
'                  IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'             IsMail = False
'         End If
'      End If
'
'      '­Y¥Ó½Ð¤H¥þªÅ¥Õ¡A¤£µo
'      If IsMail = False Or (mLC(11) = "" And mLC(43) = "" And Trim(mLC(44)) = "" And Trim(mLC(45)) = "" And Trim(mLC(46)) = "") Then
'           oMailCount = ""
'      End If
'
'      '¨t²Î§O¥u§PÂ_1½X,¦]¬°FG
'      If UCase(Mid(mLC(1), 1, 1)) <> "F" And oMailCount <> "" Then
'         '¥Ó½Ð¤H¬° X65299 ©Î X03072 ªº©Ò¦³Ãö«Y¥ø·~³£¤£ÀË¬d·~°È°Ï
'         If Left(mLC(11), 6) <> "X65299" And Left(mLC(11), 6) <> "X03072" And _
'            Left(mLC(43), 6) <> "X65299" And Left(mLC(43), 6) <> "X03072" And _
'            Left(Trim(mLC(44)), 6) <> "X65299" And Left(Trim(mLC(44)), 6) <> "X03072" And _
'            Left(Trim(mLC(45)), 6) <> "X65299" And Left(Trim(mLC(45)), 6) <> "X03072" And _
'            Left(Trim(mLC(46)), 6) <> "X65299" And Left(Trim(mLC(46)), 6) <> "X03072" Then
'            'Modify By Sindy 2022/10/14
'            If UCase(pFormName) <> UCase("frm090801_New") Then
'            '2022/9/27 END
'               MsgBox "¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¤£¦P·~°È°Ï¡A·Ç³Æµo mail ¡I", , "ª`·N¡I"
'            End If
'            '¥[µo¨q¬Â
'            'Modify By Sindy 2022/9/29 §ï§ì Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'            oMailCount = oMailCount & Trim(mCP(13)) & ";" & Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'            oContext = oContext & vbCrLf + "¦¬¤å´¼Åv¤H­û¡G " + GetStaffName(mCP(13)) + vbCrLf + vbCrLf + "´¼Åv¤H­û(°Ï)¤£¦P¡I"
''            PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å³qª¾--¦¹®×¦¬¤å«D­ì´¼Åv¤H­û(°Ï)¡I", oContext
'            'Modify By Sindy 2022/9/29
'            'Modify By Sindy 2023/3/27 +,mc13
'            mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
'               " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'               ",'" & "®×¥ó¦¬¤å³qª¾--¦¹®×¦¬¤å«D­ì´¼Åv¤H­û(°Ï)¡I(¤å¸¹:" & mCP(9) & ")','" & oContext & "',null,'" & mCP(9) & "')"
'            cnnConnection.Execute mStrSql
'            '2022/9/29 END
'         End If
'      End If
'2024/11/6 mark END

   End If
   
   'Add By Sindy 2022/9/27
   If UCase(pFormName) = UCase("frm090801_New") Then
      If ERecvSaveProgress(mCP, mLC, Pub_GetSpecMan("ACS¤À®×¤H­û"), oContext) = False Then
         GoTo ErrHand
      End If
   End If
   
   'Added by Lydia 2024/01/15 ÀË¬d©|¦bÅU°Ý¸u¥ô´Á¶¡ªº«È¤á¬O§_¤w³]©wÅU°Ý±M¥Î«H½c
   If intModifyKind = 0 And mCP(1) = "ACS" And mCP(10) = "112" Then
      strTmp1(0) = Pub_GetChkCU199("1", mCP(9), True)
   End If
   'end 2024/01/15
   
   Exit Function
   
ErrHand:
   PUB_SaveFrm010006_1 = False 'Add By Sindy 2022/10/25
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical, "PUB_SaveFrm010006_1"
   End If
End Function

'Added by Lydia 2022/09/14 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G·s¼WACS®×¤§112´¼°]ÅU°Ý¦¬¤å¦Ü¸ê®Æ®w(±qfrm010006_1.InsertCaseDatabase©â¥X¨Ó)
Private Function InsertCaseDB(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intCaseKind As Integer, ByVal intChoose As Integer, _
                ByRef mLC() As String, ByRef mCP() As String, ByVal mCU30 As String, Optional ByRef IsSaveData As Boolean, Optional ByVal pType As String, Optional ByVal pCaseNo As String) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intCaseKind¡A1¬°±M§Q¡A2¬°°Ó¼Ð¡A3¬°ªk°È¡A4¬°ÅU°Ý¡A5¬°±M§Q(ªA)¡A6¬°°Ó¼Ð(ªA)¡A7¬°ªk°È(ªA)¡A8¬°ÅU°Ý(ªA)
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
Dim strAutoNumber As String, bolError As Boolean
Dim adoquery As New ADODB.Recordset
Dim strCusReceipt As String  '¦¬¾Ú¤½¥q§O

If IsSaveData = True Then
    Exit Function
End If
IsSaveData = True

On Error GoTo ErrHand
   '¶Ç¤J0¬°­«½Æ¤§¥»©Ò®×¸¹(·s¼WÂÂ®×)¡A1¬°¥¿½T¤§¥»©Ò®×¸¹(·s¼W·s®×)
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.BeginTrans
   End If
   If intSaveMode = 1 Then
      If mLC(2) = "" Then
         If ClsPDGetAutoNumber(mLC(1), strAutoNumber, True, False) Then
            mLC(2) = strAutoNumber
         Else
            bolError = True
         End If
      End If
      If bolError = False Then
         mCP(2) = mLC(2)
         '¦¬¾Ú¤½¥q§O
         If intCaseKind <> ÅU°Ý Then
            strCusReceipt = GetReceiptCmp(Mid(mLC(11), 1, 8), Mid(mLC(11), 9, 1), mCP(1), "000")
         End If
         Select Case intCaseKind
                Case ªk°È
                    mStrSql = "insert into lawcase (lc01,lc02,lc03,lc04,lc05,lc06,lc07,lc11,lc15,lc16,lc42,lc43,lc44,lc45,lc46,lc48) " + _
                        "values (" + CNULL(mCP(1)) + "," + CNULL(mCP(2)) + "," + CNULL(mCP(3)) + "," + CNULL(mCP(4)) + "," + CNULL(ChgSQL(mLC(5))) + "," + _
                        "null, null," + CNULL(mLC(11)) + ",'000' ," + CNULL(ChgSQL(mLC(16))) + "," + CNULL(mLC(42)) + "," + CNULL(mLC(43)) + "," + CNULL(mLC(44)) + "," + CNULL(mLC(45)) + "," + CNULL(mLC(46)) + "," + CNULL(strCusReceipt) + ")"
                    cnnConnection.Execute mStrSql
         End Select
         mCP(31) = "Y"
      Else
         bolError = True
      End If
   End If
   If bolError = False Then

      If ClsPDGetAutoNumber(Left(mCP(9), 1), strAutoNumber, True, True) Then
         mCP(9) = mCP(9) + strAutoNumber
        'Modified by Lydia 2021/11/19 ¿é¤J³W¶O¡BÂI¼Æ
        'Memo by Lydia 2022/09/07 CP150: ¦³¡¹¡¹ªºÀ³¦¬±b´ÚÃ±®Ö±±ºÞ
        'Modify By Sindy 2022/9/28 +,cp140
        'Modify By Sindy 2023/4/18 +,cp151
         mStrSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp11,cp12,cp13,cp14,cp16, " + _
             "cp17,cp18,cp31,cp32,cp53,cp54,CP150,cp140,cp151) values (" + CNULL(mCP(1)) + "," + CNULL(mCP(2)) + "," + CNULL(mCP(3)) + "," + CNULL(mCP(4)) + "," + CNULL(mCP(5)) + "," + _
              CNULL(mCP(9)) + "," + CNULL(mCP(10)) + "," + CNULL(mCP(11)) + "," + CNULL(mCP(12)) + "," + CNULL(mCP(13)) + "," + CNULL(mCP(14)) + "," + CNULL(mCP(16)) + "," + _
              CNULL(mCP(17)) + "," + CNULL(mCP(18)) + "," + CNULL(mCP(31)) + "," + CNULL(mCP(32)) + "," + CNULL(mCP(53)) + "," + CNULL(mCP(54)) + "," + CNULL(mCP(150)) + "," + _
              CNULL(mCP(140)) + "," + CNULL(mCP(151)) + ")"
         cnnConnection.Execute mStrSql, intJ
         
        'Add By nickc 2007/08/21
        '­Y¬°±µ¬¢°O¿ý³æ(Âd¥x¦¬¤å)
        'Modify by Morgan 2007/10/26 ¶O¥Î¥i§ï®É¤~°µ¡A§_«h¤w¦¬´Ú¸ê®Æ·|³QÁÙ­ì
         If intChoose = 0 And mCP(60) = "" Then  'mCP(60) = "" => txtAdviser(9).Enabled = True Then
             '¥¼¦¬ª÷ÃB = ¶O¥Î
             mStrSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(mCP(9))
             cnnConnection.Execute mStrSql
         End If
         'Added by Lydia 2022/11/29 «D¤º³¡¦¬¤å¨Ã¥B¦³¶O¥Î¡A¥ý²Î¤@³]©wCP20=Null ;
         If intChoose = 0 And Val(mCP(16)) > 0 Then
             mStrSql = "update caseprogress set cp20=null where cp09=" + CNULL(mCP(9))
             cnnConnection.Execute mStrSql
         End If
         'end 2022/11/29
         'Add By Cheng 2002/05/10
         '­Y¬°¤º³¡¦¬¤å§@·~®É, ®×¥ó¶i«×ÀÉªº¬O§_¦V«È¤á¦¬´Ú³]©w¬°"N"
         If intChoose = 1 Then
            mStrSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
         End If
         
         'Modify By Sindy 2023/11/3 mark:¦¬¤å¤w¤£»Ý°õ¦æ¦¹¬qµ{¦¡,¦]±µ¬¢³æ¤w·|¦^¼g¶l»¼°Ï¸¹
'         mStrSql = "update customer set cu30=" + CNULL(mCU30) + " where cu01=" + CNULL(Mid(mLC(11), 1, 8)) + " and cu02=" + CNULL(Mid(mLC(11), 9, 1))
'         cnnConnection.Execute mStrSql
         '2023/11/3 EMD
        
        '­Y¦b¤U¤@µ{§ÇÀÉ¥u§ì¨ì¤@µ§¸ê®Æ®É, ¤~­n§ì¤U¤@µ{§ÇÀÉªºÁ`¦¬¤å¸¹§ó·s®×¥ó¶i«×ÀÉªº¬ÛÃöÁ`¦¬¤å¸¹
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select np01 from nextprogress where np02 = '" & mCP(1) & "' and np03 = '" & mCP(2) & "' and np04 = '" & mCP(3) & "' and np05 = '" & mCP(4) & "' and np06 is null and np07 = '" & mCP(10) & "'", cnnConnection, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
            If IsNull(adoquery.Fields(0).Value) = False Then
               cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & mCP(9) & "'"
            End If
         End If
         adoquery.Close
      Else
         bolError = True
      End If
   End If

   
   If bolError Then
      'Add By Sindy 2022/9/27
      If UCase(pFormName) <> UCase("frm090801_New") Then
      '2022/9/27 END
         cnnConnection.RollbackTrans
      End If
      ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
      IsSaveData = False
   Else
      'Add By Sindy 2022/9/27
      If UCase(pFormName) <> UCase("frm090801_New") Then
      '2022/9/27 END
         cnnConnection.CommitTrans
      End If
      InsertCaseDB = True
      'txtCode(0) = mCP(2)  '²¾¨ì¥~¼h
   End If
   'txtCode(0) = mCP(2)  '²¾¨ì¥~¼h
   Set adoquery = Nothing
   Exit Function

ErrHand:
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.RollbackTrans
   End If
   ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
   IsSaveData = False
   Set adoquery = Nothing
End Function

'Added by Lydia 2022/09/14 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G­×§ïACS®×¤§112´¼°]ÅU°Ý¦¬¤å(±qfrm010006_1.UpdateCaseDatabase©â¥X¨Ó)
Private Function UpdateCaseDB(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intCaseKind As Integer, ByVal intChoose As Integer, _
                ByRef mLC() As String, ByRef mCP() As String, ByVal mCU30 As String, Optional ByRef IsSaveData As Boolean, Optional ByVal pType As String, Optional ByVal pCaseNo As String) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intCaseKind¡A1¬°±M§Q¡A2¬°°Ó¼Ð¡A3¬°ªk°È¡A4¬°ÅU°Ý¡A5¬°±M§Q(ªA)¡A6¬°°Ó¼Ð(ªA)¡A7¬°ªk°È(ªA)¡A8¬°ÅU°Ý(ªA)
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
Dim adoquery As New ADODB.Recordset
Dim strCusReceipt As String  '¦¬¾Ú¤½¥q§O

If IsSaveData = True Then
    Exit Function
End If
IsSaveData = True

On Error GoTo ErrHand

 If intCaseKind <> ÅU°Ý Then
    strCusReceipt = GetReceiptCmp(Mid(mLC(11), 1, 8), Mid(mLC(11), 9, 1), mCP(1), "000")
 End If
 
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.BeginTrans
   End If
   
    Select Case intCaseKind
          Case ªk°È
                 mStrSql = "update lawcase set lc05=" + CNULL(ChgSQL(mLC(5))) + ", lc11=" + CNULL(mLC(11)) + ", lc16=" + CNULL(ChgSQL(mLC(16))) + _
                       ", lc43=" + CNULL(mLC(43)) + ", lc44=" + CNULL(mLC(44)) + ", lc45=" + CNULL(mLC(45)) + ", lc46=" + CNULL(mLC(46)) + ", lc48=" + CNULL(strCusReceipt)
                 mStrSql = mStrSql + " where lc01=" + CNULL(mCP(1)) + " and lc02=" + CNULL(mCP(2)) + " and lc03=" + CNULL(mCP(3)) + " and lc04=" + CNULL(mCP(4))
                 cnnConnection.Execute mStrSql
    End Select
    
    'Modified by Lydia 2021/11/19 ¿é¤J³W¶O¡BÂI¼Æ
    'Memo by Lydia 2022/09/07 CP150: ¦³¡¹¡¹ªºÀ³¦¬±b´ÚÃ±®Ö±±ºÞ
    mStrSql = "update caseprogress set cp05=" + CNULL(mCP(5)) + ",cp10=" + CNULL(mCP(10)) + ",cp11=" + CNULL(mCP(11)) + ",cp53=" + CNULL(mCP(53)) + ",cp54=" + CNULL(mCP(54)) + ",cp13=" + CNULL(mCP(13)) + _
       ",cp14=" + CNULL(mCP(14)) + ",cp16=" + CNULL(mCP(16)) + " ,cp17=" + CNULL(mCP(17)) + ",cp18=" + CNULL(mCP(18)) + " ,cp32=" + CNULL(mCP(32)) + " ,cp150=" & CNULL(mCP(150)) & " where cp09=" + CNULL(mCP(9))
    cnnConnection.Execute mStrSql
    mStrSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(mCP(13)) + ") where cp09=" + CNULL(mCP(9))
    cnnConnection.Execute mStrSql

    'Add By nickc 2007/08/21
    '­Y¬°±µ¬¢°O¿ý³æ(Âd¥x¦¬¤å)
    'Modify by Morgan 2007/10/26 ¶O¥Î¥i§ï®É¤~°µ¡A§_«h¤w¦¬´Ú¸ê®Æ·|³QÁÙ­ì
    If intChoose = 0 And mCP(60) = "" Then  'mCP(60) = "" => txtAdviser(9).Enabled = True Then
        '¥¼¦¬ª÷ÃB = ¶O¥Î
        mStrSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(mCP(9))
        cnnConnection.Execute mStrSql
    End If
    'Added by Lydia 2022/11/29 «D¤º³¡¦¬¤å¨Ã¥B¦³¶O¥Î¡A¥ý²Î¤@³]©wCP20=Null ;
    If intChoose = 0 And Val(mCP(16)) > 0 Then
        mStrSql = "update caseprogress set cp20=null where cp09=" + CNULL(mCP(9))
        cnnConnection.Execute mStrSql
    End If
    'end 2022/11/29
    'Add By Cheng 2002/05/10
    '­Y¬°¤º³¡¦¬¤å§@·~®É, ®×¥ó¶i«×ÀÉªº¬O§_¦V«È¤á¦¬´Ú³]©w¬°"N"
    If intChoose = 1 Then
       mStrSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(mCP(9))
       cnnConnection.Execute mStrSql
    End If
   
    'Modify By Sindy 2023/11/3 mark:¦¬¤å¤w¤£»Ý°õ¦æ¦¹¬qµ{¦¡,¦]±µ¬¢³æ¤w·|¦^¼g¶l»¼°Ï¸¹
'    mStrSql = "update customer set cu30=" + CNULL(mCU30) + " where cu01=" + CNULL(Mid(mLC(11), 1, 8)) + " and cu02=" + CNULL(Mid(mLC(11), 9, 1))
'    cnnConnection.Execute mStrSql
    '2023/11/3 END
    
    adoquery.CursorLocation = adUseClient
    adoquery.Open "select np01 from nextprogress where np02 = '" & mCP(1) & "' and np03 = '" & mCP(2) & "' and np04 = '" & mCP(3) & "' and np05 = '" & mCP(4) & "' and np06 is null and np07 = '" & mCP(10) & "'", cnnConnection, adOpenStatic, adLockReadOnly
    '­Y¦b¤U¤@µ{§ÇÀÉ¥u§ì¨ì¤@µ§¸ê®Æ®É, ¤~­n§ì¤U¤@µ{§ÇÀÉªºÁ`¦¬¤å¸¹§ó·s®×¥ó¶i«×ÀÉªº¬ÛÃöÁ`¦¬¤å¸¹
    If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
       If IsNull(adoquery.Fields(0).Value) = False Then
          cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & mCP(9) & "'"
       End If
    End If
    adoquery.Close
    
'Add By Sindy 2022/9/27
If UCase(pFormName) <> UCase("frm090801_New") Then
'2022/9/27 END
   cnnConnection.CommitTrans
End If
UpdateCaseDB = True
Set adoquery = Nothing
Exit Function

ErrHand:
'Add By Sindy 2022/9/27
If UCase(pFormName) <> UCase("frm090801_New") Then
'2022/9/27 END
   cnnConnection.RollbackTrans
End If
ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf

IsSaveData = False
Set adoquery = Nothing
End Function

'Added Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á¡A¨Ã¥B«e¤T­Ó¤ë¤º¨S¦³¨ä¥L¤H­û«Ø¥ß©¹¨Ó°O¿ý
'Move by Lydia 2022/09/08 ±qbasPublic·h¹L¨Ó
'Modify By Sindy 2022/9/27 + Optional bolTrans As Boolean = True : ­n¤UConnection Transªº§PÂ_
'Modified by Lydia 2023/12/29 +pCaseNo¥»©Ò®×¸¹
Public Function PUB_ChkOldCustomer(ByVal bolUpdate As Boolean, ByVal pCustNo As String, ByVal pST01 As String, _
   Optional ByVal pSt15 As String, Optional ByVal pST06 As String, Optional ByVal bolTrans As Boolean = True, Optional ByVal pCaseNo As String) As Boolean
'§@¥Î¡G­ì°µ¸ó°Ï¦¬¤åªº±±ºÞ(ChkSameCuArea)¡A­Y¬°¦P©Ò¤§«Ý¬¡¤Æ«È¤á¥BNVL(OCU03,0)=0¡A¥B«e¤T­Ó¤ë¤º¨S¦³¨ä¥L¤H­û«Ø¥ß©¹¨Ó°O¿ý®É¡A«hÂd¥x¦¬¤å¤£¼u°T®§©Memail¡C
           '¦sÀÉ®É¦P®É§ó·s«È¤áÀÉ´¼Åv¤H­ûCU13¬°¦¬¤å´¼Åv¤H­û¡A·~°È°ÏCU12¤]¤@¨Ö§ï¡A¦A¼gºûÅ@°O¿ýDML-LOG¡C
'bolUpdate '¬O§_§ó·sDB
Dim intB As Integer
Dim rsB As ADODB.Recordset
Dim strTmpB As String
Dim outSales As String, outOffice As String
Dim strProc As String
Dim strCase(0 To 4) As String, strCC As String  'Added by Lydia 2023/12/29
Dim rsAD As New ADODB.Recordset 'Added by Lydia 2024/05/29

    PUB_ChkOldCustomer = False
    If pCustNo = "" Then Exit Function
   
    'If bolUpdate = True Then 'Mark by Lydia 2023/12/29
        If pSt15 = "" Then
           pSt15 = GetST15(pST01, , , pST06)
        End If
    'End If  'Mark by Lydia 2023/12/29
    
    pCustNo = ChangeCustomerL(pCustNo)
    outSales = "" '«Ý¬¡¤Æ«È¤áªº´¼Åv¤H­û
    outOffice = "" '«Ý¬¡¤Æ«È¤áªº©Ò§O
    'Added by Lydia 2023/12/29
    If pCaseNo <> "" Then
      strCase(0) = Replace(pCaseNo, "-", "")
      Call ChgCaseNo(strCase(0), strCase)
      If strCase(2) <> "" Then
         strCase(0) = strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) = "000", "", "-" & strCase(3) & "-" & strCase(4))
      Else
         strCase(0) = ""
      End If
    End If
    'end 2023/12/29
    
    '«e¤T­Ó¤ë¤º¨ä¥L¤H­ûªº©¹¨Ó°O¿ý
    'Modified by Lydia 2019/10/18 debug
    'strTmpB = "SELECT COR03 AS V01,COUNT(*) AS V02 FROM CONTACTRECORD1 " & _
                     "WHERE COR03='" & pCustNo & "' AND COR02 BETWEEN " & CompDate(1, -3, strSrvDate(1)) & " AND " & strSrvDate(1) & _
                     IIf(pSt01 <> "", " AND COR06=" & CNULL(pSt01), "") & "  GROUP BY COR03"
    'Modified by Lydia 2023/12/29 ¨ú®ø" ¬°¦P©Ò¤§«Ý¬¡¤Æ«È¤á¡A¨Ã¥B«e¤T­Ó¤ë¤º¨S¦³¨ä¥L¤H­û«Ø¥ß©¹¨Ó°O¿ý®É¡A«h¶}©ñ¥i¥H¦¬¤å"­­¨î¡A¥u­n¬O¦P©Òªº«Ý¬¡¤Æ«È¤á¡A¥ô¦ó¤H³£¥i¥H¦¬¤å¡C
    'strTmpB = "SELECT COR03 AS V01,COUNT(*) AS V02 FROM CONTACTRECORD1 " & _
                     "WHERE COR03='" & pCustNo & "' AND COR02 BETWEEN " & CompDate(1, -3, strSrvDate(1)) & " AND " & strSrvDate(1) & _
                     IIf(pST01 <> "", " AND COR06<>" & CNULL(pST01), "") & " AND COR06<>'QPGMR' GROUP BY COR03"
    ''Modified by Lydia 2023/11/30 +NVL(CU04,NVL(CU05,CU06)) CUSTNAME
    'strTmpB = "SELECT OCU01,ST01,ST15,NVL(ST06,'1') ST06 ,NVL(V02,0) CNT,NVL(CU04,NVL(CU05,CU06)) CUSTNAME FROM OLDCUSTOMER,CUSTOMER,STAFF " & _
                     ", (" & strTmpB & ") VTB1 WHERE OCU01='" & Mid(pCustNo, 1, 8) & "' " & _
                     "AND OCU01=CU01 AND CU02='0' AND NVL(OCU03,0)=0 AND CU13=ST01(+) AND CU01||CU02=V01(+) "
    strTmpB = "SELECT OCU01,ST01,ST15,NVL(ST06,'1') ST06 ,NVL(CU04,NVL(CU05,CU06)) CUSTNAME FROM OLDCUSTOMER,CUSTOMER,STAFF " & _
                     "WHERE OCU01='" & Mid(pCustNo, 1, 8) & "' " & _
                     "AND OCU01=CU01 AND CU02='0' AND NVL(OCU03,0)=0 AND CU13=ST01(+) "
    'end 2023/12/29
    intB = 1
    Set rsB = ClsLawReadRstMsg(intB, strTmpB)
    If intB = 1 Then
       '«Ý¬¡¤Æ«È¤á+¨S¦³«e¤T­Ó¤ë¤º¨ä¥L¤H­ûªº©¹¨Ó°O¿ý
        outSales = "" & rsB.Fields("st01")
        outOffice = "" & rsB.Fields("st06")
        'Modified by Lydia 2019/10/18 ¬°¦P©Ò¤§«Ý¬¡¤Æ«È¤á¡A¨Ã¥B«e¤T­Ó¤ë¤º¨S¦³¨ä¥L¤H­û«Ø¥ß©¹¨Ó°O¿ý®É¡A«h¶}©ñ¥i¥H¦¬¤å
        'If "" & rsB.Fields("ocu01") <> "" And Val("" & rsB.Fields("cnt")) = 0 Then
        'Modified by Lydia 2023/12/29 ¨ú®ø" ¬°¦P©Ò¤§«Ý¬¡¤Æ«È¤á¡A¨Ã¥B«e¤T­Ó¤ë¤º¨S¦³¨ä¥L¤H­û«Ø¥ß©¹¨Ó°O¿ý®É¡A«h¶}©ñ¥i¥H¦¬¤å"­­¨î¡A¥u­n¬O¦P©Òªº«Ý¬¡¤Æ«È¤á¡A¥ô¦ó¤H³£¥i¥H¦¬¤å¡C
        'If "" & rsB.Fields("ocu01") <> "" And Val("" & rsB.Fields("cnt")) = 0 And outOffice = pST06 Then
        If "" & rsB.Fields("ocu01") <> "" And outOffice = pST06 Then
            PUB_ChkOldCustomer = True
        End If
        
        '«Ý¬¡¤Æ«È¤á¥u­n¦¬¤å->¡@OCU03¤W¨t²Î¤é¡Aªí¥Ü¤w¬¡¤Æ
        If bolUpdate = True And "" & rsB.Fields("ocu01") <> "" Then
            'Modified by Lydia 2023/12/29 Ãö«Y¥ø·~ªº«Ý¬¡¤Æ«È¤á¤]­n¤@¨Ö¬¡¤Æ
            'strProc = "Update OldCustomer SET OCU03=" & strSrvDate(1) & " Where OCU01=" & CNULL(rsB.Fields("ocu01"))
            strProc = "Update OldCustomer SET OCU03=" & strSrvDate(1) & " Where SUBSTR(OCU01,1,6)=" & CNULL(Left(rsB.Fields("ocu01"), 6)) & " AND NVL(OCU03,0)=0 "
        End If
        
        If bolUpdate = True And strProc <> "" Then
            'Modify By Sindy 2022/9/27
            If bolTrans = True Then
            '2022/9/27 END
               cnnConnection.BeginTrans
            End If
            cnnConnection.Execute strProc
            If PUB_ChkOldCustomer = True Then
                '¦P©Ò¤§«Ý¬¡¤Æ«È¤á: §ó·s«È¤áÀÉ´¼Åv¤H­ûCU13¬°¦¬¤å´¼Åv¤H­û¡A·~°È°ÏCU12¤]¤@¨Ö§ï
                If pST06 = outOffice And outSales & "" & rsB.Fields("st15") <> pST01 & pSt15 And pST01 <> "" Then
                    strProc = "Update Customer Set CU12=" & CNULL(pSt15) & ", CU13=" & CNULL(pST01) & " Where CU01||CU02=" & CNULL(pCustNo)
                    Pub_SeekTbLog strProc
                    cnnConnection.Execute "begin user_data.user_enabled:=1; " & strProc & " ; end; "
                    'Added by Lydia 2023/12/29 «È¤á§ï´¼Åv¤H­û­n¦P®É§ó·s¸Ó«È¤á©Ò¦³®×¥ó¤U¤@µ{§Ç¥¼¹L´Á¥¼Äò¿ìªº´Á­­
                    Call PUB_ChangeSaleUpdNP10(pCustNo, outSales, pST01, True)
                    'Added by Lydia 2024/05/29 ¤@¨Ö§ó·s§ó¦W«eªº½s¸¹; ex.(113/5/28) ¦¬¤åX13424
                    strTmpB = "select cu01||cu02 as custno,st15 from customer, staff where cu13=st01(+) and cu01='" & Mid(pCustNo, 1, 8) & "' and cu02<>'" & Mid(pCustNo, 9, 1) & "' order by 1,2 "
                    intB = 1
                    Set rsAD = ClsLawReadRstMsg(intB, strTmpB)
                    If intB = 1 Then
                       rsAD.MoveFirst
                       Do While Not rsAD.EOF
                          strProc = "Update Customer Set CU12=" & CNULL(pSt15) & ", CU13=" & CNULL(pST01) & " Where CU01||CU02=" & CNULL(rsAD.Fields("custno"))
                          Pub_SeekTbLog strProc
                          cnnConnection.Execute "begin user_data.user_enabled:=1; " & strProc & " ; end; "
                          Call PUB_ChangeSaleUpdNP10("" & rsAD.Fields("custno"), outSales, pST01, True)
                          rsAD.MoveNext
                       Loop
                    End If
                    'end 2024/05/28
                End If
            End If
            
            'Added by Lydia 2023/12/29 ­Y­ì´¼Åv¤H­û¬°´¼Åv³¡¤H­û¡A³£½Ð¥[µo°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v
            If Left("" & rsB.Fields("st15"), 1) = "S" Then
               strCC = Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")
            End If
            'Added by Lydia 2023/11/30 ¥[µoEMAILµ¹­ì´¼Åv¤H­û¡G
            '1. ·s¦¬¤å´¼Åv¤H­û¬°­ì´¼Åv¤H­û®É¡A¥D¦®¡G³qª¾¡G«È¤á½s¸¹+«È¤á¦WºÙ¡A¦]¦¬¤å¤w¬¡¤Æ¡C
            '2. ·s¦¬¤å´¼Åv¤H­û»P­ì´¼Åv¤H­û¤£¦P®É¡A¥D¦®¡G³qª¾¡G«È¤á½s¸¹+«È¤á¦WºÙ¡A¦]·s´¼Åv¤H­û(¤H¦W)¦¬¤å¥»©Ò®×¸¹¤w¬¡¤Æ¨ÃÂà²¾¦¹«È¤á¤§´¼Åv¤H­û¡C
            If pST01 = "" & rsB.Fields("st01") Then
               'Modified by Lydia 2023/12/29 +CC¥þ©Ò´¼Åv³¡¥DºÞ(MC09)
               strProc = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                         " values( '" & strUserNum & "','" & rsB.Fields("st01") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                         ",'" & rsB.Fields("ocu01") & " " & rsB.Fields("custname") & "¡A¦]¦¬¤å¤w¬¡¤Æ¡C','¦p¦®'," & CNULL(strCC) & ")"
               cnnConnection.Execute strProc
            Else
               'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹strCase(0), +CC¥þ©Ò´¼Åv³¡¥DºÞ(MC09)
               strProc = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                         " values( '" & strUserNum & "','" & rsB.Fields("st01") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
                         ",'" & rsB.Fields("ocu01") & " " & rsB.Fields("custname") & "¡A¦]" & GetStaffName(pST01, True) & "¦¬¤å¥»©Ò®×¸¹" & strCase(0) & "¤w¬¡¤Æ¨ÃÂà²¾¦¹«È¤á¤§´¼Åv¤H­û¡C','¦p¦®'," & CNULL(strCC) & ")"
               cnnConnection.Execute strProc
            End If
            'end 2023/11/30
            
            'Modify By Sindy 2022/9/27
            If bolTrans = True Then
            '2022/9/27 END
               cnnConnection.CommitTrans
            End If
        End If
    End If
    
    Set rsB = Nothing
    Set rsAD = Nothing 'Added by Lydia 2024/05/29
    
    Exit Function
    
ErrHandle:
   If Err.Number <> "" Then
      If strProc <> "" Then
         'Modify By Sindy 2022/9/27
         If bolTrans = True Then
         '2022/9/27 END
            cnnConnection.RollbackTrans
         End If
      End If
      MsgBox Err.Description
   End If
End Function

'Added by Lydia 2021/07/29 ¨ú±o­û¤u³Ì°ª¯ÅªººÞ¨î¤H
'Move by Lydia 2022/09/08 ±qbasPublic·h¹L¨Ó
Public Function PUB_GetSTManLimit(ByVal pUserNo As String, Optional ByVal stLv As String = "5") As String
'pUserNo ¶Ç¤J­û¤u½s¸¹
'stLv ¨ú¨ì²ÄX¯Å
Dim strA As String, intA As Integer
Dim rsAD As New ADODB.Recordset
     
     PUB_GetSTManLimit = ""
     If pUserNo = "" Or Val(stLv) < 2 Then Exit Function
     
     'Modified by Lydia 2022/01/28 +A1.st16
     strA = "SELECT A1.ST01,A1.ST16,A2.ST01 AS ST52,A2.ST04 AS ST52_K, A3.ST01 AS ST53,A3.ST04 AS ST53_K,A4.ST01 AS ST54 ,A4.ST04 AS ST54_K,A5.ST01 AS ST55 ,A5.ST04 AS ST55_K,A0908 " & _
               "FROM STAFF A1,STAFF A2,STAFF A3, STAFF A4,STAFF A5,ACC090 WHERE A1.ST01=" & CNULL(pUserNo) & _
               "AND A1.ST52=A2.ST01(+) AND A1.ST53=A3.ST01(+) AND A1.ST54=A4.ST01(+) AND A1.ST55=A5.ST01(+) AND A1.ST03=A0901(+) "
    intA = 1
    Set rsAD = ClsLawReadRstMsg(intA, strA)
    If intA = 1 Then
        'Added by Lydia 2022/01/28 ­^¤å²ÕST16='2'ªÌ§ì³Ì°ª²Ä¤G¯Å¥DºÞ(ST55°£¥~)
        If "" & rsAD.Fields("ST16") = "2" Then
            strA = ""
            For intA = Val(stLv) To 2 Step -1
                If "" & rsAD.Fields("ST5" & intA) <> "" And "" & rsAD.Fields("ST5" & intA & "_K") = "1" Then
                     If strA = "" Then
                         strA = "" & rsAD.Fields("ST5" & intA)
                     Else
                         strA = "" & rsAD.Fields("ST5" & intA)
                         Exit For
                     End If
                End If
            Next intA
            If strA = "" Then
                 PUB_GetSTManLimit = "" & rsAD.Fields("A0908")
            Else
                 PUB_GetSTManLimit = strA
            End If
        Else  '¤é¤å²ÕST16<>'2' ªÌ§ì³Ì°ª¯Å¥DºÞ(ST55°£¥~)
        'end 2022/01/28
            For intA = Val(stLv) To 2 Step -1
                If PUB_GetSTManLimit = "" And "" & rsAD.Fields("ST5" & intA) <> "" Then
                   If "" & rsAD.Fields("ST5" & intA & "_K") <> "1" Then
                        PUB_GetSTManLimit = "" & rsAD.Fields("A0908") '¤H­ûÂ÷Â¾±a¤J³¡ªù¥DºÞ
                   Else
                        PUB_GetSTManLimit = "" & rsAD.Fields("ST5" & intA)
                   End If
                   Exit For
                End If
            Next intA
        End If 'Added by Lydia 2022/01/28
    End If
    
    Set rsAD = Nothing
End Function

'Added by Lydia 2022/09/14 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡GªA°È®×¡Bªk°È®×¡BACS®×«D112¤§¦¬¤å¡BÅU°Ý®×«D0¤§¦¬¤å(±qfrm010007.SaveDatabase©â¥X¨Ó)
'Modify By Sindy 2024/11/21 + , Optional ByVal m_intCRC As Integer = 0: ¦Û°Ê¦¬¤åªº®×¥ó©Ê½è¶¶§Ç
Public Function PUB_SaveFrm010007(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intCaseKind As Integer, ByVal intChoose As Integer, _
                ByRef mOTB() As String, ByRef mCP() As String, ByVal mCU30 As String, ByVal mChkVal As String, ByVal mSaveControl As String, Optional ByRef IsSaveData As Boolean, _
                Optional ByVal pType As String, Optional ByVal pCaseNo As String, Optional ByRef RetVal As String, Optional ByVal m_intCRC As Integer = 0) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intCaseKind¡A1¬°±M§Q¡A2¬°°Ó¼Ð¡A3¬°ªk°È¡A4¬°ÅU°Ý¡A5¬°±M§Q(ªA)¡A6¬°°Ó¼Ð(ªA)¡A7¬°ªk°È(ªA)¡A8¬°ÅU°Ý(ªA)
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'mOTB¡G¨Ì¨t²Î§O¶Ç¤J°ò¥»ÀÉ(ªk°ÈLawCase¡BÅU°ÝHireCase¡BªA°ÈServicePractice)
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
'reTurnVal : ¦^¶Ç­È
'mChkVal¡G¶Ç¤J¨ä¥L¾Þ§@µ²ªG
'mSaveControl: »ô³Æ¤éºÞ¨î
Dim m_SalesST15 As String, m_SalesST06 As String
Dim m_SalesDeptName As String
Dim m_CuNo(1 To 5) As String '¥Ó½Ð¤H/·í¨Æ¤H1~5
Dim m_FaNo As String 'FC¥N²z¤H
Dim m_Na01 As String '¥Ó½Ð°ê®a
Dim m_CaseName As String '®×¥ó¦WºÙ
Dim m_CP10Name As String '¦¬¤å¤§®×¥ó©Ê½è¦WºÙ
Dim m_Na01Name As String '¥Ó½Ð°ê®a¦WºÙ
Dim rsRD As New ADODB.Recordset
'ªk«ß©Ò®×·½¦¬¤å
Dim m_LOS01 As String '®×·½Á`¦¬¤å¸¹
Dim m_LOS01cp01 As String, m_LOS01cp02 As String, m_LOS01cp03 As String, m_LOS01cp04 As String '®×·½Á`¦¬¤å¸¹¤§¥»©Ò®×¸¹
Dim m_LOS02 As String '®×·½®×¥óÃþ«¬
Dim m_LOS15 As String '®×·½³æ¸¹
Dim m_LOS04 As String  '¤¶²Ð¤H
Dim m_LOS04_1 As String, m_LOS04_1st15 As String, m_LOS04_1st06 As String '¤¶²Ð¤H(²Ä¤@¦ì)¡B¦¬¤å³¡ªù¡B©Ò§O
Dim m_LOS05 As String  '¤¶²Ð«È¤á
Dim m_LOS12 As String  '¤¶²Ð¤é
Dim m_Los05_N As String   'LA¸É®×·½¤§¤¶²Ð¤H¤¶²Ð«È¤á
'Add By Sindy 2022/9/30
Dim bolIsOverDt As Boolean, strDivisionalEmp As String
Dim strRDate As String, strRTime As String, strRestKind As String
Dim bolCP14Rest As Boolean
'2022/9/30 END
Dim oMailCount As String

'*********¯S®íºÞ¨îªºÅÜ¼Æ*************
    'Modify By Sindy 2025/8/18
    'If pType = "LOS®×·½¦¬¤å" And pCaseNo <> "" Then
    If InStr(pType, "LOS®×·½¦¬¤å") > 0 And pCaseNo <> "" Then
    '2025/8/18 END
        m_LOS02 = Mid(pCaseNo, 1, InStr(pCaseNo, ",") - 1) '®×·½®×¥óÃþ«¬
        m_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1, 8) '®×·½³æ¸¹ 'Modify By Sindy 2025/8/18 + , 8)
        strTmp1(0) = "select X.*,cp01,cp02,cp03,cp04 from LawOfficeSource X,caseprogress where los15=" & CNULL(m_LOS15) & " and los01=cp09(+) "
        intJ = 1
        Set rsRD = ClsLawReadRstMsg(intJ, strTmp1(0))
        If intJ = 1 Then
          '®×·½Á`¦¬¤å¸¹
          m_LOS01 = "" & rsRD.Fields("LOS01")
          '®×·½Á`¦¬¤å¸¹¤§¥»©Ò®×¸¹
          m_LOS01cp01 = "" & rsRD.Fields("cp01")
          m_LOS01cp02 = "" & rsRD.Fields("cp02")
          m_LOS01cp03 = "" & rsRD.Fields("cp03")
          m_LOS01cp04 = "" & rsRD.Fields("cp04")
          '(­ì)®×·½®×¥óÃþ«¬
          m_LOS02 = "" & rsRD.Fields("LOS02")
          '®×·½³æ¸¹
          m_LOS15 = "" & rsRD.Fields("LOS15")
          '¤¶²Ð¤H, ¤¶²Ð¤H(²Ä¤@¦ì)
          m_LOS04 = "" & rsRD.Fields("LOS04")
          If m_LOS04 <> "" And InStr(m_LOS04, ",") > 0 Then
             m_LOS04_1 = Mid(m_LOS04, 1, InStr(m_LOS04, ",") - 1)
          Else
             m_LOS04_1 = m_LOS04
          End If
          If m_LOS04_1 <> "" Then
             m_LOS04_1st15 = GetST15(m_LOS04_1, , , m_LOS04_1st06)
          End If
          '(­ì)¤¶²Ð«È¤á:
          m_LOS05 = "" & rsRD.Fields("LOS05")
          '¤¶²Ð¤é
          m_LOS12 = "" & rsRD.Fields("LOS12")
        End If
        Set rsRD = Nothing
    End If
    '¥~±M«H¥ó¨R¾P: ¥u¦³·s¼W¦¬¤åªº¥\¯à
'***********************************
   'Added by Lydia 2023/05/11 ¦]¬°PUB_ReadCaseData·|¦^¶Ç6½X«È¤á½s¸¹,©Ò¥H¥ý²Î¤@«È¤á½s¸¹
   If ClsPDGetSystemKind(mOTB(1), intJ) = True Then
      Select Case intJ
         Case ªk°È
              '·í¨Æ¤H1~5
              mOTB(11) = ChangeCustomerL(mOTB(11))
              mOTB(43) = ChangeCustomerL(mOTB(43))
              mOTB(44) = ChangeCustomerL(mOTB(44))
              mOTB(45) = ChangeCustomerL(mOTB(45))
              mOTB(46) = ChangeCustomerL(mOTB(46))
         Case ÅU°Ý
              mOTB(5) = ChangeCustomerL(mOTB(5))
              mOTB(24) = ChangeCustomerL(mOTB(24))
              mOTB(25) = ChangeCustomerL(mOTB(25))
              mOTB(26) = ChangeCustomerL(mOTB(26))
              mOTB(27) = ChangeCustomerL(mOTB(27))
         Case Else  'ªA°È
              mOTB(8) = ChangeCustomerL(mOTB(8))
              mOTB(58) = ChangeCustomerL(mOTB(58))
              mOTB(59) = ChangeCustomerL(mOTB(59))
              mOTB(65) = ChangeCustomerL(mOTB(65))
              mOTB(66) = ChangeCustomerL(mOTB(66))
      End Select
   End If
   'end 2023/05/11
   
   RetVal = "" '¦^¶Ç­È
   
    Select Case intCaseKind
         Case ªk°È
            m_CaseName = mOTB(5)  '®×¥ó¦WºÙ
            m_Na01 = mOTB(15)  '¥Ó½Ð°ê®a
            m_FaNo = mOTB(22)  'FC¥N²z¤H
            '¥Ó½Ð¤H/·í¨Æ¤H1~5
            m_CuNo(1) = mOTB(11):   m_CuNo(2) = mOTB(43):  m_CuNo(3) = mOTB(44):   m_CuNo(4) = mOTB(45):  m_CuNo(5) = mOTB(46)
         Case ÅU°Ý
            m_CaseName = mOTB(6)
            m_Na01 = "000"
            m_FaNo = ""
            m_CuNo(1) = mOTB(5):   m_CuNo(2) = mOTB(24):  m_CuNo(3) = mOTB(25):  m_CuNo(4) = mOTB(26):  m_CuNo(5) = mOTB(27)
         Case Else 'ªA°È
            m_CaseName = mOTB(5)
            m_Na01 = mOTB(9)
            m_FaNo = mOTB(26)
            m_CuNo(1) = mOTB(8):  m_CuNo(2) = mOTB(58):   m_CuNo(3) = mOTB(59):  m_CuNo(4) = mOTB(65):  m_CuNo(5) = mOTB(66)
    End Select
   intJ = ClsPDGetCaseProperty(mOTB(1), mCP(10), m_CP10Name, IIf(m_Na01 <> "000", True, False))
   m_Na01Name = PUB_GetNationName(m_Na01)
   m_SalesST15 = GetST15(mCP(13), m_SalesDeptName, , m_SalesST06)
   
   If intModifyKind = 0 Then
       PUB_SaveFrm010007 = InsertOtherDB(pFormName, intSaveMode, intModifyKind, intCaseKind, intChoose, mOTB, mCP, mCU30, mChkVal, mSaveControl, IsSaveData, pType, pCaseNo, m_Los05_N)
   Else
       PUB_SaveFrm010007 = UpdateOtherDB(pFormName, intSaveMode, intModifyKind, intCaseKind, intChoose, mOTB, mCP, mCU30, mChkVal, mSaveControl, IsSaveData, pType, pCaseNo)
   End If
   If PUB_SaveFrm010007 = False Then Exit Function 'Add By Sindy 2022/9/28 ¦sÀÉ¥¢±Ñ,«áÄò¤£ÀË¬d
   
   'Added by Lydia 2022/10/06 §PÂ_¦^¶Ç­È¬°¥~±M«H¥ó¨R¾Pªº¤º®e
   'Modify By Sindy 2023/5/31
   'If InStr(pType, "¥~±M«H¥ó¨R¾P") > 0 And pCaseNo <> "" Then
   If InStr(pType, "«H¥ó¨R¾P") > 0 And pCaseNo <> "" Then
   '2023/5/31 END
       RetVal = m_Los05_N
       m_Los05_N = ""
   End If
   'end 2022/10/06
   
   'add by nickc 2007/11/09 ´ú¸Õ¸Ñ¨Mmail µo¤£¨ìªº®É­Ô·|¦s¨âµ§ªº¿ù»~
   On Error GoTo 0    'Âk¹s
   On Error GoTo ErrHand 'Add By Sindy 2022/9/29
   'Add By Sindy 2022/12/29 ­«ÅªCP,¦]«eÀYUpdate¨ç¼Æµ{¦¡¦³¥i¯àª½±µ¦sDB,¨S¦³§ó·scp³¯¦C­È
   strTmp1(0) = mCP(9)
   Erase mCP
   ReDim Preserve mCP(TF_CP) As String
   mCP(9) = strTmp1(0)
   'Modified by Lydia 2023/05/11 + false
   Call PUB_ReadCaseProgressDatabase(mCP(), 1, False)
   '2022/12/29 END
   
   '¬d¦W³æ¹ïÀ³¦sÀÉ
   If pType = "T¬d¦W³æ" And pCaseNo <> "" Then
      strTmp1(1) = Mid(pCaseNo, 1, InStr(pCaseNo, "|") - 1)
      strTmp1(2) = Mid(pCaseNo, InStr(pCaseNo, "|") + 1)
      'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
      'Modified by Lydia 2024/03/14 +Fasle
      'If PUB_TMQtoCP("", mCP(9), strTmp1(2), strTmp1(1), , , IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)) = False Then
      If PUB_TMQtoCP(False, "", mCP(9), strTmp1(2), strTmp1(1), , , IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)) = False Then
      End If
   End If
   
   'add by nickc 2005/09/05
   If intModifyKind = 0 Then
      Dim oContext As String, strCaseNo As String
      Dim strTemp As String
      Dim m_strState As String
      'Add By Sindy 2023/8/18 ¤£±o¥N²zªº«áÄòÂÂ®×¦¬¤å±±ºÞ¡A³qª¾¦¬¤å¤H­û¡]CP13¡^
      If m_FaNo <> "" Then
        If GetAgentAndState(m_FaNo, strTmp1(1), , , , mOTB(1), m_strState, IIf(intSaveMode = 0, True, False)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥N²z¤H¡G " + m_FaNo + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & m_FaNo
          End If
        End If
      End If
      If m_CuNo(1) <> "" Then
        If GetCustomerAndState(m_CuNo(1), strTmp1(1), , , , mOTB(1), m_strState, IIf(intSaveMode = 0, True, False), , mOTB(2), mOTB(3), mOTB(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H1¡G " + m_CuNo(1) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & m_CuNo(1)
          End If
        End If
      End If
      If m_CuNo(2) <> "" Then
        If GetCustomerAndState(m_CuNo(2), strTmp1(1), , , , mOTB(1), m_strState, IIf(intSaveMode = 0, True, False), , mOTB(2), mOTB(3), mOTB(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H2¡G " + m_CuNo(2) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & m_CuNo(2)
          End If
        End If
      End If
      If m_CuNo(3) <> "" Then
        If GetCustomerAndState(m_CuNo(3), strTmp1(1), , , , mOTB(1), m_strState, IIf(intSaveMode = 0, True, False), , mOTB(2), mOTB(3), mOTB(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H3¡G " + m_CuNo(3) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & m_CuNo(3)
          End If
        End If
      End If
      If m_CuNo(4) <> "" Then
        If GetCustomerAndState(m_CuNo(4), strTmp1(1), , , , mOTB(1), m_strState, IIf(intSaveMode = 0, True, False), , mOTB(2), mOTB(3), mOTB(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H4¡G " + m_CuNo(4) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & m_CuNo(4)
          End If
        End If
      End If
      If m_CuNo(5) <> "" Then
        If GetCustomerAndState(m_CuNo(5), strTmp1(1), , , , mOTB(1), m_strState, IIf(intSaveMode = 0, True, False), , mOTB(2), mOTB(3), mOTB(4)) Then
          If InStr(m_strState, "¤£±o¥N²z") > 0 Then
             oContext = oContext & vbCrLf + "¥Ó½Ð¤H5¡G " + m_CuNo(5) + " " + strTmp1(1) + vbCrLf
             strTemp = strTemp & "," & m_CuNo(5)
          End If
        End If
      End If
      'Add By Sindy 2025/3/26
      If mCP(56) <> "" Then
         '¶Ç¤J®×¥ó©Ê½è , mCP(10)
         If GetCustomerAndState(mCP(56), strTmp1(1), , , , mOTB(1), m_strState, IIf(intSaveMode = 0, True, False), , mOTB(2), mOTB(3), mOTB(4), mCP(10)) Then
            If InStr(m_strState, "¤£±o¥N²z") > 0 Then
               oContext = oContext & vbCrLf + GetPrjState6(mOTB(1), mCP(10), IIf(m_Na01 = "000", "0", "1")) + "¥Ó½Ð¤H1¡G " + mCP(56) + " " + strTmp1(1) + vbCrLf
               strTemp = strTemp & "," & mCP(56)
            End If
        End If
      End If
      If mCP(89) <> "" Then
         '¶Ç¤J®×¥ó©Ê½è , mCP(10)
         If GetCustomerAndState(mCP(89), strTmp1(1), , , , mOTB(1), m_strState, IIf(intSaveMode = 0, True, False), , mOTB(2), mOTB(3), mOTB(4), mCP(10)) Then
            If InStr(m_strState, "¤£±o¥N²z") > 0 Then
               oContext = oContext & vbCrLf + GetPrjState6(mOTB(1), mCP(10), IIf(m_Na01 = "000", "0", "1")) + "¥Ó½Ð¤H2¡G " + mCP(89) + " " + strTmp1(1) + vbCrLf
               strTemp = strTemp & "," & mCP(89)
            End If
        End If
      End If
      If mCP(90) <> "" Then
         '¶Ç¤J®×¥ó©Ê½è , mCP(10)
         If GetCustomerAndState(mCP(90), strTmp1(1), , , , mOTB(1), m_strState, IIf(intSaveMode = 0, True, False), , mOTB(2), mOTB(3), mOTB(4), mCP(10)) Then
            If InStr(m_strState, "¤£±o¥N²z") > 0 Then
               oContext = oContext & vbCrLf + GetPrjState6(mOTB(1), mCP(10), IIf(m_Na01 = "000", "0", "1")) + "¥Ó½Ð¤H3¡G " + mCP(90) + " " + strTmp1(1) + vbCrLf
               strTemp = strTemp & "," & mCP(90)
            End If
        End If
      End If
      If mCP(91) <> "" Then
         '¶Ç¤J®×¥ó©Ê½è , mCP(10)
         If GetCustomerAndState(mCP(91), strTmp1(1), , , , mOTB(1), m_strState, IIf(intSaveMode = 0, True, False), , mOTB(2), mOTB(3), mOTB(4), mCP(10)) Then
            If InStr(m_strState, "¤£±o¥N²z") > 0 Then
               oContext = oContext & vbCrLf + GetPrjState6(mOTB(1), mCP(10), IIf(m_Na01 = "000", "0", "1")) + "¥Ó½Ð¤H4¡G " + mCP(91) + " " + strTmp1(1) + vbCrLf
               strTemp = strTemp & "," & mCP(91)
            End If
        End If
      End If
      If mCP(92) <> "" Then
         '¶Ç¤J®×¥ó©Ê½è , mCP(10)
         If GetCustomerAndState(mCP(92), strTmp1(1), , , , mOTB(1), m_strState, IIf(intSaveMode = 0, True, False), , mOTB(2), mOTB(3), mOTB(4), mCP(10)) Then
            If InStr(m_strState, "¤£±o¥N²z") > 0 Then
               oContext = oContext & vbCrLf + GetPrjState6(mOTB(1), mCP(10), IIf(m_Na01 = "000", "0", "1")) + "¥Ó½Ð¤H5¡G " + mCP(92) + " " + strTmp1(1) + vbCrLf
               strTemp = strTemp & "," & mCP(92)
            End If
        End If
      End If
      '2025/3/26 END
      If oContext <> "" Then
         strTemp = Mid(strTemp, 2)
         strCaseNo = IIf("-" + mOTB(3) + "-" + mOTB(4) = "-0-00", mOTB(1) + "-" + mOTB(2), mOTB(1) + "-" + mOTB(2) + "-" + mOTB(3) + "-" + mOTB(4))
         oContext = "¥»©Ò®×¸¹¡G " + strCaseNo + vbCrLf + _
                    "®×¥ó¦WºÙ¡G " + mOTB(5) + vbCrLf + _
                    "¥Ó½Ð°ê®a¡G " + m_Na01 + " " + m_Na01Name + vbCrLf + _
                    "¦¬¤å¤é¡G " + ChangeTStringToTDateString(TransDate(mCP(5), 1)) + vbCrLf + _
                    "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf + vbCrLf + _
                    "¡i¤£±o¥N²z¡j" + vbCrLf + _
                    oContext
         oMailCount = mCP(13) & ";" & PUB_GetFCPProSup(mCP(13), True)
         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & strCaseNo & _
            " ¤w½T»{Äò¦æ¦¬¤å¡A½Ðª`·N¸Ó" & strTemp & "½s¸¹¤w³]¬°¤£±o¥N²z¡C(¤å¸¹:" & mCP(9) & ")','" & oContext & "',null,'" & mCP(9) & "')"
         cnnConnection.Execute mStrSql
      End If
      '2023/8/18 END
      
      'add by nickc 2007/05/16 ¥[¤J­Y¬O¥»©Ò´Á­­¤p©óµ¥©ó·í¤Ñ¡A­nµomail  ³qª¾
      Dim oContext2 As String
      oContext2 = ""
      
      'Added by Lydia 2022/09/28 (»~§R,¸É¦^)
      oContext = "¥»©Ò®×¸¹¡G " + mOTB(1) + "-" + mOTB(2) + "-" + mOTB(3) + "-" + mOTB(4) + vbCrLf + "®×¥ó¦WºÙ¡G " + m_CaseName + vbCrLf + "¦¬¤å¤é¡G " + ChangeTStringToTDateString(mCP(5)) + vbCrLf + "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf
      
      'add by nickc 2007/05/16 ¥[¤J­Y¬O¥»©Ò´Á­­¤p©óµ¥©ó·í¤Ñ¡A­nµomail  ³qª¾
      'edit by nickc 2008/04/23 ¥[¤J°ê®a
      '2009/6/29 modify by sonia
      oContext2 = "¥»©Ò®×¸¹¡G " + mOTB(1) + "-" + mOTB(2) + "-" + mOTB(3) + "-" + mOTB(4) + vbCrLf + "®×¥ó¦WºÙ¡G " + m_CaseName + vbCrLf + "¥Ó½Ð°ê®a¡G" + m_Na01 + m_Na01Name + vbCrLf + "¦¬¤å¤é¡G " + ChangeTStringToTDateString(mCP(5)) + vbCrLf + "®×¥ó©Ê½è¡G " + m_CP10Name + vbCrLf
      
      'Added by Lydia 2020/10/05 (9/30) ­Y¸Ó¦¬¤å¸¹ÂI¼Æ>0¦ýµL®×·½(¦Û¦æ¦¬¤åªÌ)®É¡A­Y®×¥óªº«È¤á¬°«Dªk«ß©Òªº«È¤á®É¤£½×·sÂÂ®×¡A¨t²Î¦Û°Ê·s¼WTT-999999®×¶i«×(BÃþ¦¬¤å)¤Îªk«ß©Ò®×·½¸ê®Æ¡C­Y¬°·s®×·~°È°Ï¤£¦PªºEmail·ÓÂÂ³qª¾¡C
      'Modified by Lydia 2022/09/14 ­­¨îªk°È®×·½¦¬¤å ; ¦]¬°¥~±M«H¥ó¨R¾P¬O±q¨t²Î¦¬¥ó°Ï¡A©Ò¥H¤£·|¸ò®×·½¦¬¤å(frm090801)­«Å|
      'If m_Los05_N <> "" Then  '¦]¬°Âd¥xµLªk³B²z,©Ò¥H¥uµoemail
      'Modified by Lydia 2022/10/19 Lydia 2022/10/19 «È¤á½s¸¹«á«Ø=m_LOS05=ªÅ¥Õ; ex.L-006577
      'If m_Los05_N <> "" And strSrvDate(1) >= ªk«ß©Ò®×·½¦¬¤å±Ò¥Î¤é And InStr(mOTB(1), "L") > 0 And m_LOS15 = "" And Val(mCP(18)) > 0 And mOTB(2) <> "" Then
      'Modified by Lydia 2023/08/11 §PÂ_¬O§_¦³¸ó°Ï¦¬¤å®É¡A¤£­­¨î¦³¨S¦³¦¬¶O¡C
      'If m_Los05_N <> "" And strSrvDate(1) >= ªk«ß©Ò®×·½¦¬¤å±Ò¥Î¤é And InStr(mOTB(1), "L") > 0 And (m_LOS15 = "" Or m_LOS05 = "") And Val(mCP(18)) > 0 And mOTB(2) <> "" Then
      If m_Los05_N <> "" And strSrvDate(1) >= ªk«ß©Ò®×·½¦¬¤å±Ò¥Î¤é And InStr(mOTB(1), "L") > 0 And (m_LOS15 = "" Or m_LOS05 = "") And mOTB(2) <> "" Then
           m_LOS05 = Mid(m_Los05_N, 1, InStr(m_Los05_N, "|") - 1)
           m_LOS04_1 = Mid(m_Los05_N, InStr(m_Los05_N, "|") + 1)
           m_LOS04_1st15 = GetST15(m_LOS04_1)
      End If
      'end 2020/10/05
      
      'Modify By Sindy 2024/11/6 §ï¦¨¦@¥Î¨ç¼Æ: ¦¬¤å®É,ÀË¬d¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¬O§_¦³»~
      '§ï¼g­ì¥Ñ¬O¦]¬°¥Ó½Ð¤H1~5 ³v¤@ÀË¬d,¦³»~§¡­nµo mail
      'edit by nickc 2007/08/21 ­Y¥Ó½Ð¤H¥þªÅ¥Õ¡A¤£µo
      If Not (m_CuNo(1) = "" And m_CuNo(2) = "" And m_CuNo(3) = "" And m_CuNo(4) = "" And m_CuNo(5) = "") Then
         'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¤¶²Ð«È¤á¬°ÂÂ«È¤á¦ý»P¤¶²Ð¤H¤£¦P°Ï®ÉµoMail³qª¾¬ÛÃö¤H­û
         If m_LOS05 <> "" And m_LOS04_1 <> "" Then
            'Modify By Sindy 2024/11/21 ¦Û°Ê¦¬¤åªº®×¥ó©Ê½è¶¶§Ç=1 ©Î¯È¥»¦¬¤å¥¼«ü©w
            If m_intCRC = 1 Or m_intCRC = 0 Then
            '2024/11/21 END
               Call RecvChkApplCust("¥Ó½Ð¤H¡þ·í¨Æ¤H1", m_CuNo(1), m_LOS04_1, "", m_LOS04_1st15, Trim(mCP(12)), oContext, m_LOS04_1st06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9), m_LOS04_1)
               Call RecvChkApplCust("¥Ó½Ð¤H¡þ·í¨Æ¤H2", m_CuNo(2), m_LOS04_1, "", m_LOS04_1st15, Trim(mCP(12)), oContext, m_LOS04_1st06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9), m_LOS04_1)
               Call RecvChkApplCust("¥Ó½Ð¤H¡þ·í¨Æ¤H3", m_CuNo(3), m_LOS04_1, "", m_LOS04_1st15, Trim(mCP(12)), oContext, m_LOS04_1st06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9), m_LOS04_1)
               Call RecvChkApplCust("¥Ó½Ð¤H¡þ·í¨Æ¤H4", m_CuNo(4), m_LOS04_1, "", m_LOS04_1st15, Trim(mCP(12)), oContext, m_LOS04_1st06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9), m_LOS04_1)
               Call RecvChkApplCust("¥Ó½Ð¤H¡þ·í¨Æ¤H5", m_CuNo(5), m_LOS04_1, "", m_LOS04_1st15, Trim(mCP(12)), oContext, m_LOS04_1st06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9), m_LOS04_1)
            End If
         Else
            'Modify By Sindy 2024/11/21 ¦Û°Ê¦¬¤åªº®×¥ó©Ê½è¶¶§Ç=1 ©Î¯È¥»¦¬¤å¥¼«ü©w
            If m_intCRC = 1 Or m_intCRC = 0 Then
            '2024/11/21 END
               Call RecvChkApplCust("¥Ó½Ð¤H¡þ·í¨Æ¤H1", m_CuNo(1), Trim(mCP(13)), Trim(m_FaNo), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
               Call RecvChkApplCust("¥Ó½Ð¤H¡þ·í¨Æ¤H2", m_CuNo(2), Trim(mCP(13)), Trim(m_FaNo), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
               Call RecvChkApplCust("¥Ó½Ð¤H¡þ·í¨Æ¤H3", m_CuNo(3), Trim(mCP(13)), Trim(m_FaNo), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
               Call RecvChkApplCust("¥Ó½Ð¤H¡þ·í¨Æ¤H4", m_CuNo(4), Trim(mCP(13)), Trim(m_FaNo), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
               Call RecvChkApplCust("¥Ó½Ð¤H¡þ·í¨Æ¤H5", m_CuNo(5), Trim(mCP(13)), Trim(m_FaNo), m_SalesST15, Trim(mCP(12)), oContext, m_SalesST06, pFormName, mCP(1), mCP(2), mCP(3), mCP(4), mCP(9))
            End If
         End If
      End If
      '2024/11/6 END
'Modify By Sindy 2024/11/6 mark
'      'add by nick 2004/10/15  ·í¦¬¤å·~°È°Ï»P«È¤áÀÉ·~°È°Ï¤£¦P®Éµo mail  ¤Î´£¥Ü
'      Dim oStrCuSales1 As String
'      Dim oStrCuSales2 As String
'      Dim oStrCuSales3 As String
'      Dim oStrCuSales4 As String
'      Dim oStrCuSales5 As String
'      'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'      Dim IsMail As Boolean
'      IsMail = True
'
'      oStrCuSales1 = ""
'      oStrCuSales2 = ""
'      oStrCuSales3 = ""
'      oStrCuSales4 = ""
'      oStrCuSales5 = ""
'
'      oMailCount = ""
'      'Added by Lydia 2020/04/08 ´¼¼z©Ò§ó¦W¤é°_¨ú®ø´¼Åv¤H­û»P«È¤áÀÉ´¼Åv¤H­ûªº±±¨î
'      'Modified by Lydia 2020/06/05 +µÛ§@ÅvTC
'      'Modified by Lydia 2022/10/18 debug: µÛ§@ÅvTC¥Îµe­±ªº´¼Åv¤H­û§PÂ_
'      'If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And (InStr(mOTB(1), "L") > 0 Or mOTB(1) = "TC") Then
'      If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And InStr(mOTB(1), "L") > 0 Then
'            'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¤¶²Ð«È¤á¬°ÂÂ«È¤á¦ý»P¤¶²Ð¤H¤£¦P°Ï®ÉµoMail³qª¾¬ÛÃö¤H­û
'            If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(m_CuNo(1)) <> "" Then
'                If ChkSameCuArea(Trim(m_CuNo(1)), m_LOS04_1) = False Then
'                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_CuNo(1)), oStrCuSales1)), 1) = "F" Then
'                        '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'                    Else
'                        oMailCount = oMailCount & oStrCuSales1 & ";"
'                        'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                        If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales1), 1) = "S" And _
'                           InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                           oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                        End If
'                        '2023/11/7 END
'                        oContext = oContext & vbCrLf + "¥Ó½Ð¤H¡þ·í¨Æ¤H1¡G " + GetCustomerName(ChangeCustomerL(m_CuNo(1))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales1)
'                    End If
'                Else
'                       IsMail = False
'                End If
'            ElseIf Trim(m_CuNo(1)) <> "" Then
'            'end 2020/05/20
'                GoTo JumpToChk01 'Added by Lydia 2022/10/18 µL®×·½¤¶²Ð¤H¥Hµe­±¿é¤J§PÂ_; ex.L-006576(®Û­^ªº«È¤á¡Aªk«ß©Ò¦Û¦æ¦¬¤å),¦]¬°¤§«eµL®×·½©Ò¥H¤]¤£·|¸É¤W®×·½¸ê®Æ
'                IsMail = False
'            End If 'Added by Lydia 2020/05/20
'      Else
'      'end 2020/04/08
'            'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'            'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'JumpToChk01: 'Added by Lydia 2022/10/18
'            'Modify By Sindy 2023/2/2 +, , oStrCuSales1 : ¦^¶Ç­ì´¼Åv¤H­û
'            If ChkSameCuArea(Trim(m_CuNo(1)), Trim(mCP(13)), , , , , Trim(m_FaNo), , oStrCuSales1) = False And Trim(mCP(13)) <> "" And Trim(m_CuNo(1)) <> "" Then
'               'Add By Sindy 2009/10/19
'               'Modify By Sindy 2023/2/2
'               'If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_CuNo(1)), oStrCuSales1)), 1) = "F" Then
'               If Left(m_SalesST15, 1) = "F" And Left(GetSalesArea(oStrCuSales1), 1) = "F" Then
'               '2023/2/2 END
'                  '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'               Else
'                  oMailCount = oMailCount & oStrCuSales1 & ";"
'                  'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                  If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales1), 1) = "S" And _
'                     InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                     oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                  End If
'                  '2023/11/7 END
'                  oContext = oContext & vbCrLf + "¥Ó½Ð¤H¡þ·í¨Æ¤H1¡G " + GetCustomerName(ChangeCustomerL(m_CuNo(1))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales1)
'               End If
'             'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'            Else
'                   If Trim(mCP(13)) <> "" And Trim(m_CuNo(1)) <> "" Then
'                       IsMail = False
'                   End If
'            End If
'      End If 'Added by Lydia 2020/04/08
'      'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡GÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(m_CuNo(1)) <> "" Then
'            'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'            'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'            If PUB_ChkOldCustomer(True, m_CuNo(1), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06, _
'                     IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                IsMail = False
'            End If
'      Else
'      'end 2020/05/20
'            'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'            If m_SalesST06 <> "" And Trim(m_CuNo(1)) <> "" And Trim(mCP(13)) <> "" Then
'                'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, m_CuNo(1), Trim(mCP(13)), m_SalesST15, m_SalesST06, _
'                        IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                   IsMail = False
'                End If
'            End If
'      End If 'Added by Lydia 2020/05/20
'
'      'Added by Lydia 2020/04/08 ´¼¼z©Ò§ó¦W¤é°_¨ú®ø´¼Åv¤H­û»P«È¤áÀÉ´¼Åv¤H­ûªº±±¨î
'      'Modified by Lydia 2020/06/05 +µÛ§@ÅvTC
'      'Modified by Lydia 2022/10/18 debug: µÛ§@ÅvTC¥Îµe­±ªº´¼Åv¤H­û§PÂ_
'      'If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And (InStr(mOTB(1), "L") > 0 Or mOTB(1) = "TC") Then
'      If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And InStr(mOTB(1), "L") > 0 Then
'            'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¤¶²Ð«È¤á¬°ÂÂ«È¤á¦ý»P¤¶²Ð¤H¤£¦P°Ï®ÉµoMail³qª¾¬ÛÃö¤H­û
'            If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(m_CuNo(2)) <> "" Then
'                If ChkSameCuArea(Trim(m_CuNo(2)), m_LOS04_1) = False Then
'                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_CuNo(2)), oStrCuSales2)), 1) = "F" Then
'                        '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'                    Else
'                        oMailCount = oMailCount & oStrCuSales2 & ";"
'                        'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                        If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales2), 1) = "S" And _
'                           InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                           oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                        End If
'                        '2023/11/7 END
'                        oContext = oContext & vbCrLf + "¥Ó½Ð¤H¡þ·í¨Æ¤H2¡G " + GetCustomerName(ChangeCustomerL(m_CuNo(2))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales2)
'                    End If
'                Else
'                       IsMail = False
'                End If
'            ElseIf Trim(m_CuNo(2)) <> "" Then
'            'end 2020/05/20
'                GoTo JumpToChk02 'Added by Lydia 2022/10/18 µL®×·½¤¶²Ð¤H¥Hµe­±¿é¤J§PÂ_
'                IsMail = False
'            End If 'Added by Lydia 2020/05/20
'      Else
'      'end 2020/04/08
'            'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'            'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'JumpToChk02: 'Added by Lydia 2022/10/18
'            'Modify By Sindy 2023/2/2 +, , oStrCuSales2 : ¦^¶Ç­ì´¼Åv¤H­û
'            If ChkSameCuArea(Trim(m_CuNo(2)), Trim(mCP(13)), , , , , Trim(m_FaNo), , oStrCuSales2) = False And Trim(mCP(13)) <> "" And Trim(m_CuNo(2)) <> "" Then
'               'Add By Sindy 2009/10/19
'               'Modify By Sindy 2023/2/2
'               'If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_CuNo(2)), oStrCuSales2)), 1) = "F" Then
'               If Left(m_SalesST15, 1) = "F" And Left(GetSalesArea(oStrCuSales2), 1) = "F" Then
'               '2023/2/2 END
'                  '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'               Else
'                  oMailCount = oMailCount & oStrCuSales2 & ";"
'                  'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                  If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales2), 1) = "S" And _
'                     InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                     oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                  End If
'                  '2023/11/7 END
'                  oContext = oContext & vbCrLf + "¥Ó½Ð¤H¡þ·í¨Æ¤H2¡G " + GetCustomerName(ChangeCustomerL(m_CuNo(2))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales2)
'               End If
'             'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'             Else
'                   If Trim(mCP(13)) <> "" And Trim(m_CuNo(2)) <> "" Then
'                       IsMail = False
'                   End If
'            End If
'      End If 'Added by Lydia 2020/04/08
'      'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡GÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(m_CuNo(2)) <> "" Then
'            'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'            'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'            If PUB_ChkOldCustomer(True, m_CuNo(2), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06, _
'                     IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                IsMail = False
'            End If
'      Else
'      'end 2020/05/20
'            'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'            If m_SalesST06 <> "" And Trim(m_CuNo(2)) <> "" And Trim(mCP(13)) <> "" Then
'                'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, m_CuNo(2), Trim(mCP(13)), m_SalesST15, m_SalesST06, _
'                        IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                   IsMail = False
'                End If
'            End If
'      End If 'Added by Lydia 2020/05/20
'
'      'Added by Lydia 2020/04/08 ´¼¼z©Ò§ó¦W¤é°_¨ú®ø´¼Åv¤H­û»P«È¤áÀÉ´¼Åv¤H­ûªº±±¨î
'      'Modified by Lydia 2020/06/05 +µÛ§@ÅvTC
'      'Modified by Lydia 2022/10/18 debug: µÛ§@ÅvTC¥Îµe­±ªº´¼Åv¤H­û§PÂ_
'      'If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And (InStr(mOTB(1), "L") > 0 Or mOTB(1) = "TC") Then
'      If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And InStr(mOTB(1), "L") > 0 Then
'            'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¤¶²Ð«È¤á¬°ÂÂ«È¤á¦ý»P¤¶²Ð¤H¤£¦P°Ï®ÉµoMail³qª¾¬ÛÃö¤H­û
'            If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(m_CuNo(3)) <> "" Then
'                If ChkSameCuArea(Trim(m_CuNo(3)), m_LOS04_1) = False Then
'                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_CuNo(3)), oStrCuSales3)), 1) = "F" Then
'                        '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'                    Else
'                        oMailCount = oMailCount & oStrCuSales3 & ";"
'                        'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                        If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales3), 1) = "S" And _
'                           InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                           oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                        End If
'                        '2023/11/7 END
'                        oContext = oContext & vbCrLf + "¥Ó½Ð¤H¡þ·í¨Æ¤H3¡G " + GetCustomerName(ChangeCustomerL(m_CuNo(3))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales3)
'                    End If
'                Else
'                       IsMail = False
'                End If
'            ElseIf Trim(m_CuNo(3)) <> "" Then
'            'end 2020/05/20
'                GoTo JumpToChk03 'Added by Lydia 2022/10/18 µL®×·½¤¶²Ð¤H¥Hµe­±¿é¤J§PÂ_
'                IsMail = False
'            End If 'Added by Lydia 2020/05/20
'      Else
'      'end 2020/04/08
'            'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'            'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'JumpToChk03: 'Added by Lydia 2202/10/18
'            'Modify By Sindy 2023/2/2 +, , oStrCuSales3 : ¦^¶Ç­ì´¼Åv¤H­û
'            If ChkSameCuArea(Trim(m_CuNo(3)), Trim(mCP(13)), , , , , Trim(m_FaNo), , oStrCuSales3) = False And Trim(mCP(13)) <> "" And Trim(m_CuNo(3)) <> "" Then
'               'Add By Sindy 2009/10/19
'               'Modify By Sindy 2023/2/2
'               'If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_CuNo(3)), oStrCuSales3)), 1) = "F" Then
'               If Left(m_SalesST15, 1) = "F" And Left(GetSalesArea(oStrCuSales3), 1) = "F" Then
'               '2023/2/2 END
'                  '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'               Else
'                  oMailCount = oMailCount & oStrCuSales3 & ";"
'                  'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                  If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales3), 1) = "S" And _
'                     InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                     oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                  End If
'                  '2023/11/7 END
'                  oContext = oContext & vbCrLf + "¥Ó½Ð¤H¡þ·í¨Æ¤H3¡G " + GetCustomerName(ChangeCustomerL(m_CuNo(3))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales3)
'               End If
'             'add by nickc 2007/05/08 ¨q¬Â»¡¡A¨ä¤¤¤@­Ó²Å¦X´N¤£µo¤F
'             Else
'                   If Trim(mCP(13)) <> "" And Trim(m_CuNo(3)) <> "" Then
'                       IsMail = False
'                   End If
'            End If
'      End If 'Added by Lydia 2020/04/08
'      'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡GÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(m_CuNo(3)) <> "" Then
'            'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'            'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'            If PUB_ChkOldCustomer(True, m_CuNo(3), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06, _
'                     IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                IsMail = False
'            End If
'      Else
'      'end 2020/05/20
'            'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'            If m_SalesST06 <> "" And Trim(m_CuNo(3)) <> "" And Trim(mCP(13)) <> "" Then
'                'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, m_CuNo(3), Trim(mCP(13)), m_SalesST15, m_SalesST06, _
'                        IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                   IsMail = False
'                End If
'            End If
'      End If 'Added by Lydia 2020/05/20
'
'      'Added by Lydia 2020/04/08 ´¼¼z©Ò§ó¦W¤é°_¨ú®ø´¼Åv¤H­û»P«È¤áÀÉ´¼Åv¤H­ûªº±±¨î
'      'Modified by Lydia 2020/06/05 +µÛ§@ÅvTC
'      'Modified by Lydia 2022/10/18 debug: µÛ§@ÅvTC¥Îµe­±ªº´¼Åv¤H­û§PÂ_
'      'If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And (InStr(mOTB(1), "L") > 0 Or mOTB(1) = "TC") Then
'      If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And InStr(mOTB(1), "L") > 0 Then
'            'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¤¶²Ð«È¤á¬°ÂÂ«È¤á¦ý»P¤¶²Ð¤H¤£¦P°Ï®ÉµoMail³qª¾¬ÛÃö¤H­û
'            If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(m_CuNo(4)) <> "" Then
'                If ChkSameCuArea(Trim(m_CuNo(4)), m_LOS04_1) = False Then
'                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_CuNo(4)), oStrCuSales4)), 1) = "F" Then
'                        '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'                    Else
'                        oMailCount = oMailCount & oStrCuSales4 & ";"
'                        'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                        If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales4), 1) = "S" And _
'                           InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                           oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                        End If
'                        '2023/11/7 END
'                        oContext = oContext & vbCrLf + "¥Ó½Ð¤H¡þ·í¨Æ¤H4¡G " + GetCustomerName(ChangeCustomerL(m_CuNo(4))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales4)
'                    End If
'                Else
'                       IsMail = False
'                End If
'            ElseIf Trim(m_CuNo(4)) <> "" Then
'            'end 2020/05/20
'                GoTo JumpToChk04 'Added by Lydia 2022/10/18 µL®×·½¤¶²Ð¤H¥Hµe­±¿é¤J§PÂ_
'                IsMail = False
'            End If 'Added by Lydia 2020/05/20
'      Else
'      'end 2020/04/08
'            'Add By Sindy 2011/1/18
'            'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'            'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'JumpToChk04: 'Added by Lydia 2022/10/18
'            'Modify By Sindy 2023/2/2 +, , oStrCuSales4 : ¦^¶Ç­ì´¼Åv¤H­û
'            If ChkSameCuArea(Trim(m_CuNo(4)), Trim(mCP(13)), , , , , Trim(m_FaNo), , oStrCuSales4) = False And Trim(mCP(13)) <> "" And Trim(m_CuNo(4)) <> "" Then
'               'Add By Sindy 2009/10/19
'               'Modify By Sindy 2023/2/2
'               'If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_CuNo(4)), oStrCuSales4)), 1) = "F" Then
'               If Left(m_SalesST15, 1) = "F" And Left(GetSalesArea(oStrCuSales4), 1) = "F" Then
'               '2023/2/2 END
'                  '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'               Else
'                  oMailCount = oMailCount & oStrCuSales4 & ";"
'                  'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                  If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales4), 1) = "S" And _
'                     InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                     oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                  End If
'                  '2023/11/7 END
'                  oContext = oContext & vbCrLf + "¥Ó½Ð¤H¡þ·í¨Æ¤H4¡G " + GetCustomerName(ChangeCustomerL(m_CuNo(4))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales4)
'               End If
'             Else
'                   If Trim(mCP(13)) <> "" And Trim(m_CuNo(4)) <> "" Then
'                       IsMail = False
'                   End If
'            End If
'      End If 'Added by Lydia 2020/04/08
'      'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡GÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(m_CuNo(4)) <> "" Then
'            'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'            'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'            If PUB_ChkOldCustomer(True, m_CuNo(4), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06, _
'                     IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                IsMail = False
'            End If
'      Else
'      'end 2020/05/20
'            'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'            If m_SalesST06 <> "" And Trim(m_CuNo(4)) <> "" And Trim(mCP(13)) <> "" Then
'               'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'               'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'               If PUB_ChkOldCustomer(True, m_CuNo(4), Trim(mCP(13)), m_SalesST15, m_SalesST06, _
'                        IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                   IsMail = False
'               End If
'            End If
'      End If 'Added by Lydia 2020/05/20
'
'      'Added by Lydia 2020/04/08 ´¼¼z©Ò§ó¦W¤é°_¨ú®ø´¼Åv¤H­û»P«È¤áÀÉ´¼Åv¤H­ûªº±±¨î
'      'Modified by Lydia 2020/06/05 +µÛ§@ÅvTC
'      'Modified by Lydia 2022/10/18 debug: µÛ§@ÅvTC¥Îµe­±ªº´¼Åv¤H­û§PÂ_
'      'If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And (InStr(mOTB(1), "L") > 0 Or mOTB(1) = "TC") Then
'      If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And InStr(mOTB(1), "L") > 0 Then
'            'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¤¶²Ð«È¤á¬°ÂÂ«È¤á¦ý»P¤¶²Ð¤H¤£¦P°Ï®ÉµoMail³qª¾¬ÛÃö¤H­û
'            If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(m_CuNo(5)) <> "" Then
'                If ChkSameCuArea(Trim(m_CuNo(5)), m_LOS04_1) = False Then
'                    If Left(m_LOS04_1st15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_CuNo(5)), oStrCuSales5)), 1) = "F" Then
'                        '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'                    Else
'                        oMailCount = oMailCount & oStrCuSales5 & ";"
'                        'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                        If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales5), 1) = "S" And _
'                           InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                           oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                        End If
'                        '2023/11/7 END
'                        oContext = oContext & vbCrLf + "¥Ó½Ð¤H¡þ·í¨Æ¤H5¡G " + GetCustomerName(ChangeCustomerL(m_CuNo(5))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales5)
'                    End If
'                Else
'                       IsMail = False
'                End If
'            ElseIf Trim(m_CuNo(5)) <> "" Then
'            'end 2020/05/20
'                GoTo JumpToChk05 'Added by Lydia 2022/10/18 µL®×·½¤¶²Ð¤H¥Hµe­±¿é¤J§PÂ_
'                IsMail = False
'            End If 'Added by Lydia 2020/05/20
'      Else
'      'end 2020/04/08
'            'Modify by Amy 2017/01/03 ¦]¥[MCTF§PÂ_,¬G§ï§PÂ_ChkSameCuArea
'            'modify by sonia 2021/11/25 MCT®×¥[¶ÇFC¥N²z¤H¨Ó§PÂ_ChkSameCuArea
'JumpToChk05: 'Added by Lydia 2022/10/18
'            'Modify By Sindy 2023/2/2 +, , oStrCuSales5 : ¦^¶Ç­ì´¼Åv¤H­û
'            If ChkSameCuArea(Trim(m_CuNo(5)), Trim(mCP(13)), , , , , Trim(m_FaNo), , oStrCuSales5) = False And Trim(mCP(13)) <> "" And Trim(m_CuNo(5)) <> "" Then
'               'Add By Sindy 2009/10/19
'               'Modify By Sindy 2023/2/2
'               'If Left(m_SalesST15, 1) = "F" And Left(Trim(GetCuSales(ChangeCustomerL(m_CuNo(5)), oStrCuSales5)), 1) = "F" Then
'               If Left(m_SalesST15, 1) = "F" And Left(GetSalesArea(oStrCuSales5), 1) = "F" Then
'               '2023/2/2 END
'                  '­Y¦¬¤å´¼Åv¤H­û¤§ST15¬°F¦rÀY¨Ã¥B«È¤á´¼Åv¤H­û¤§ST15¤]¬°F¦rÀY«h¤£µoMail
'               Else
'                  oMailCount = oMailCount & oStrCuSales5 & ";"
'                  'Add By Sindy 2023/11/7 ­Y¦¬¤å¤H­û«D´¼Åv³¡¦Ó«È¤á´¼Åv¤H­û¬°´¼Åv³¡®É¡A¤@«ß°Æ¥»µ¹¡u¥þ©Ò´¼Åv³¡¥DºÞ¡v¡C
'                  If Left(mCP(12), 1) <> "S" And Left(PUB_GetST03(oStrCuSales5), 1) = "S" And _
'                     InStr(oMailCount, Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")) = 0 Then
'                     oMailCount = oMailCount & Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ") & ";"
'                  End If
'                  '2023/11/7 END
'                  oContext = oContext & vbCrLf + "¥Ó½Ð¤H¡þ·í¨Æ¤H5¡G " + GetCustomerName(ChangeCustomerL(m_CuNo(5))) + vbCrLf + "­ì´¼Åv¤H­û¡G " + GetPrjSalesNM(oStrCuSales5)
'               End If
'             Else
'                   If Trim(mCP(13)) <> "" And Trim(m_CuNo(5)) <> "" Then
'                       IsMail = False
'                   End If
'            End If
'            '2011/1/18 End
'      End If 'Added by Lydia 2020/04/08
'      'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡GÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'      If m_LOS05 <> "" And m_LOS04_1 <> "" And Trim(m_CuNo(5)) <> "" Then
'            'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'            'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'            If PUB_ChkOldCustomer(True, m_CuNo(5), m_LOS04_1, m_LOS04_1st15, m_LOS04_1st06, _
'                     IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                IsMail = False
'            End If
'      Else
'      'end 2020/05/20
'            'Added by Lydia 2019/09/16 ÀË¬d¬O§_¬°«Ý¬¡¤Æ«È¤á,¨Ã¥B§ó·sDB
'            If m_SalesST06 <> "" And Trim(m_CuNo(5)) <> "" And Trim(mCP(13)) <> "" Then
'                'Modify By Sindy 2022/9/27 + IIf(UCase(pFormName) = UCase("frm090801_New"), False, True)
'                'Modified by Lydia 2023/12/29 +¥»©Ò®×¸¹
'                If PUB_ChkOldCustomer(True, m_CuNo(5), Trim(mCP(13)), m_SalesST15, m_SalesST06, _
'                        IIf(UCase(pFormName) = UCase("frm090801_New"), False, True), mCP(1) & mCP(2) & mCP(3) & mCP(4)) = True Then
'                   IsMail = False
'                End If
'            End If
'      End If 'Added by Lydia 2020/05/20
'
'      'edit by nickc 2007/08/21 ­Y¥Ó½Ð¤H¥þªÅ¥Õ¡A¤£µo
'      'Modify By Sindy 2011/1/18
'      If IsMail = False Or (Trim(m_CuNo(1)) = "" And Trim(m_CuNo(2)) = "" And Trim(m_CuNo(3)) = "" And Trim(m_CuNo(4)) = "" And Trim(m_CuNo(5)) = "") Then
'         oMailCount = ""
'      End If
'
'      '2006/8/2 MODIFY BY SONIA mOTB(1)¥u§PÂ_1½X,¦]¬°FG
'      'Modified by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G¥[¤WFCL
'      If (UCase(Mid(mOTB(1), 1, 1)) <> "F" Or (UCase(mOTB(1)) = "FCL" And m_LOS05 <> "")) And oMailCount <> "" Then
'         'edit by nickc 2005/08/10
'         'MsgBox "¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¤£¦P·~°È°Ï¡A·Ç³Æµo mail ¡A½Ð©w®É§R°£¶l¥ó³Æ¥÷¡I", , "ª`·N¡I"
'         'Modify By Sindy 2010/11/26 ¥Ó½Ð¤H¬° X65299 ©Î X03072 ªº©Ò¦³Ãö«Y¥ø·~³£¤£ÀË¬d·~°È°Ï
'         'Modify By Sindy 2011/1/18
'         If Left(Trim(m_CuNo(1)), 6) <> "X65299" And Left(Trim(m_CuNo(1)), 6) <> "X03072" And _
'            Left(Trim(m_CuNo(2)), 6) <> "X65299" And Left(Trim(m_CuNo(2)), 6) <> "X03072" And _
'            Left(Trim(m_CuNo(3)), 6) <> "X65299" And Left(Trim(m_CuNo(3)), 6) <> "X03072" And _
'            Left(Trim(m_CuNo(4)), 6) <> "X65299" And Left(Trim(m_CuNo(4)), 6) <> "X03072" And _
'            Left(Trim(m_CuNo(5)), 6) <> "X65299" And Left(Trim(m_CuNo(5)), 6) <> "X03072" Then
'            'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¤¶²Ð«È¤á¬°ÂÂ«È¤á¦ý»P¤¶²Ð¤H¤£¦P°Ï®ÉµoMail³qª¾¬ÛÃö¤H­û
'            'Modified by Lydia 2020/06/05 +µÛ§@ÅvTC
'            'Modify By Sindy 2022/10/14
'            If UCase(pFormName) <> UCase("frm090801_New") Then
'            '2022/9/27 END
'               'Modified by Lydia 2022/10/18 debug: µÛ§@ÅvTC¥Îµe­±ªº´¼Åv¤H­û§PÂ_
'               'If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And (InStr(mOTB(1), "L") > 0 Or mOTB(1) = "TC") And m_LOS05 <> "" And m_LOS04_1 <> "" Then
'               If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And InStr(mOTB(1), "L") > 0 And m_LOS05 <> "" And m_LOS04_1 <> "" Then
'                  MsgBox "®×·½¤¶²Ð¤H­û»P«È¤á´¼Åv¤H­û¤£¦P·~°È°Ï¡I", , "ª`·N¡I"
'               Else
'               'end 2020/05/20
'                  MsgBox "¦¬¤å´¼Åv¤H­û»P«È¤á´¼Åv¤H­û¤£¦P·~°È°Ï¡A·Ç³Æµo mail ¡I", , "ª`·N¡I"
'               End If 'Added by Lydia 2020/05/20
'            End If
'            'edit by nickc 2005/08/10 ¥[µo¨q¬Â
'            'Added by Lydia 2022/07/15 ³qª¾ªk«ß©Òªº´¼Åv¤H­û¨S¦³·N¸q¡AÀ³¸Ó­n§ï¬°®×·½¤¶²Ð¤H­û. ex.L-006547
'            'Modified by Lydia 2022/10/24 debug: µÛ§@ÅvTC¥Îµe­±ªº´¼Åv¤H­û§PÂ_
'            'If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And (InStr(mOTB(1), "L") > 0 Or mOTB(1) = "TC") And m_LOS05 <> "" And m_LOS04_1 <> "" Then
'            If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And InStr(mOTB(1), "L") > 0 And m_LOS05 <> "" And m_LOS04_1 <> "" Then
'                 'Modify By Sindy 2022/9/29 §ï§ì Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'                 oMailCount = oMailCount & m_LOS04_1 & ";" & Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'            Else
'            'end 2022/07/15
'                 'Modify By Sindy 2022/9/29 §ï§ì Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'                 oMailCount = oMailCount & Trim(mCP(13)) & ";" & Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
'            End If 'Added by Lydia 2022/07/15
'
'            'Added by Lydia 2022/11/03 ªk°È®×¦¬¤å¨ú®ø"¬O§_ªLÁ`¦P·N¥»®×¥Ñªk«ß©Ò¦Û¦æ¦¬¤å¡H"ªº¸ß°Ý¡A§ï¬°ªk«ß©Ò¦Û¦æ¦¬¤å´¼¼z©Ò¤H­û«È¤á®É¡A¦b§PÂ_¸ó°Ï¦¬¤åµoEMAILµ¹Âù¤è´¼Åv¤H­û®É¡A¥[µoªLÁ`¡C
'            If strSrvDate(1) >= ªk«ß©Ò®×·½¦¬¤å±Ò¥Î¤é And InStr(mOTB(1), "L") > 0 And m_LOS02 = "" And m_LOS15 = "" Then
'                oMailCount = oMailCount & PUB_ChkForLawMan(m_CuNo(1), mOTB(1), mOTB(2), mOTB(3), mOTB(4))
'            End If
'            'end 2022/11/03
'
'            'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¤¶²Ð«È¤á¬°ÂÂ«È¤á¦ý»P¤¶²Ð¤H¤£¦P°Ï®ÉµoMail³qª¾¬ÛÃö¤H­û
'            'Modified by Lydia 2022/10/24 debug: µÛ§@ÅvTC¥Îµe­±ªº´¼Åv¤H­û§PÂ_
'            'If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And (InStr(mOTB(1), "L") > 0 Or mOTB(1) = "TC") And m_LOS05 <> "" And m_LOS04_1 <> "" Then
'            If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And InStr(mOTB(1), "L") > 0 And m_LOS05 <> "" And m_LOS04_1 <> "" Then
'                oContext = oContext & vbCrLf + "®×·½¤¶²Ð¤H­û¡G " + GetStaffName(m_LOS04_1) + vbCrLf + vbCrLf + "´¼Åv¤H­û(°Ï)¤£¦P¡I"
'            Else
'            'end 2020/05/20
'                oContext = oContext & vbCrLf + "¦¬¤å´¼Åv¤H­û¡G " + GetStaffName(mCP(13)) + vbCrLf + vbCrLf + "´¼Åv¤H­û(°Ï)¤£¦P¡I"
'            End If 'Added by Lydia 2020/05/20
''            PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å³qª¾--¦¹®×¦¬¤å«D­ì´¼Åv¤H­û(°Ï)¡I", oContext
'            'Modify By Sindy 2022/9/29
'            'Modify By Sindy 2023/3/27 +,mc13
'            mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
'               " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'               ",'" & "®×¥ó¦¬¤å³qª¾--¦¹®×¦¬¤å«D­ì´¼Åv¤H­û(°Ï)¡I(¤å¸¹:" & mCP(9) & ")','" & oContext & "',null,'" & mCP(9) & "')"
'            cnnConnection.Execute mStrSql
'            '2022/9/29 END
'         End If
'      End If
'2024/11/6 mark END
      
      'add by nickc 2007/05/16 ¥[¤J­Y¬O¥»©Ò´Á­­¤p©óµ¥©ó·í¤Ñ¡A­nµomail  ³qª¾
      oMailCount = ""
      If mOTB(1) = "P" Or mOTB(1) = "PS" Then
          oMailCount = Pub_GetSpecMan("A")
      ElseIf mOTB(1) = "CFP" Or mOTB(1) = "CPS" Then
           oMailCount = Pub_GetSpecMan("B")
      ElseIf mOTB(1) = "FCP" Or mOTB(1) = "FG" Then
           oMailCount = Pub_GetSpecMan("C")
      ElseIf mOTB(1) = "CFT" Or mOTB(1) = "CFC" Then
           'edit by nickc 2008/04/23
           'oMailCount = Pub_GetSpecMan("D")
           'Modified by Lydia 2021/07/30 °Ó¼Ð¤Î°Ó¼ÐªA°È·~°È¦¬¤å-¦]¥~°Ó³¯¸g²z°h¥ð¦Ó­×§ïµ{¦¡±±¨î
           'oMailCount = Pub_GetSpecMan("L")
            oMailCount = GetCFTSt16Manager(mOTB(1), mOTB(2), IIf(mOTB(3) = "", "0", mOTB(3)), IIf(mOTB(4) = "", "00", mOTB(4)))

      ElseIf mOTB(1) = "FCT" Or mOTB(1) = "S" Then
          'Added by Lydia 2021/07/30 °Ó¼Ð¤Î°Ó¼ÐªA°È·~°È¦¬¤å-¦]¥~°Ó³¯¸g²z°h¥ð¦Ó­×§ïµ{¦¡±±¨î
          If mOTB(1) = "S" Then
               If mOTB(9) = "000" Then
                   'S¥xÆW®×:¥H¥»©Ò®×¸¹©I¥sPUB_GetFCTSalesNo§ì¥X­t³dªº¤H¡A¦A§ì¸Ó­û°£ST55¤§¥~ªº³Ì°ª¥DºÞNVL(NVL(ST54,ST53),ST52)
                   strTmp1(1) = PUB_GetFCTSalesNo(mOTB(1), mOTB(2), IIf(mOTB(3) = "", "0", mOTB(3)), IIf(mOTB(4) = "", "00", mOTB(4)))
                   If strTmp1(1) = "" Then
                       oMailCount = Pub_GetSpecMan("D")
                   Else
                       oMailCount = PUB_GetSTManLimit(strTmp1(1), "4") '¼Ò²Õ¤Æ
                      'Added by Lydia 2022/01/28 µoµ¹¨t²Î¯S®í³]©w¡uD¡v¤§¤H­û©M¥DºÞ
                      strTmp1(2) = Pub_GetSpecMan("D")
                      oMailCount = oMailCount & ";" & strTmp1(2)
                      'end 2022/01/28
                   End If
               Else
                   'S«D¥xÆW®×:¥H¥»©Ò®×¸¹©I¥sGetCFTSt16Manager§ì¥DºÞ
                   oMailCount = GetCFTSt16Manager(mOTB(1), mOTB(2), IIf(mOTB(3) = "", "0", mOTB(3)), IIf(mOTB(4) = "", "00", mOTB(4)))
               End If
          Else
          'end 2021/07/30
               oMailCount = Pub_GetSpecMan("D")
          End If 'Added by Lydia 2021/07/30
          'add by nickc 2007/06/23 ¥[¤JFCT ª§Ä³®×³qª¾¤º°Ó°Óª§  84027;69008     ®×¥ó©Ê½è 202 °£¥~¡AÁÙ¬O°e¥~°Ó¡Aªü½¬»¡¦Û¤v§PÂ_¡A­Y¬°¤º°Ó®×¥ó¡A¥L·|¦AÂà¹L¨Ó
          If mOTB(1) = "FCT" Then    'add by nickc 2007/08/01 ¤£§PÂ_ªº¸Ü  CFT ¤]·|¶i¤J
'edit by nickc 2007/08/10  ¤º¥~°Ó¨óÄ³ 202 ³£µo
              If mCP(10) = "202" Then
                  oMailCount = Pub_GetSpecMan("F")
              End If

          End If
      ElseIf Mid(mOTB(1), 1, 1) = "T" Then
           oMailCount = Pub_GetSpecMan("E")
      'Add By Sindy 2022/12/16
      ElseIf Mid(mOTB(1), 1, 1) = "L" Then
           oMailCount = Pub_GetSpecMan("L®×¨ú®ø³¬¨÷³qª¾¤H­û")
           '2022/12/16 END
      End If
      strDivisionalEmp = oMailCount 'Add By Sindy 2022/9/30 °O¿ý¤À®×¤H­û,«áÄò·|¥Î¨ì
      'Add By Sindy 2023/1/11 ­Y¤w¦³¤À©Ó¿ì¤H,¦P®É¤@¨Ö³qª¾
      If mCP(14) <> "" Then
         If oMailCount <> "" Then oMailCount = oMailCount & ";"
         oMailCount = oMailCount & mCP(14)
      End If
      '2023/1/11 END
      
      If DBDATE(mCP(6)) < strSrvDate(1) And Trim(mCP(6)) <> "" And Trim(oMailCount) <> "" Then
         '2007/8/13 MODIFY BY SONIA ¥[´¼Åv¤H­û
         'Modify By Sindy 2010/12/16 ¥[·~°È°Ï,¶O¥Î
'         PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¤w¹O¥»©Ò´Á­­¡A½Ð¾¨³t¿ì²z¡I", oContext2 & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(6))) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(7))) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0")
         'Modify By Sindy 2022/9/29
         'Modified by Lydia 2022/12/23 +chgsql
         'Modify By Sindy 2023/3/27 +,mc13
         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¤w¹O¥»©Ò´Á­­¡A½Ð¾¨³t¿ì²z¡I(¤å¸¹:" & mCP(9) & ")','" & ChgSQL(oContext2) & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(6))) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(7))) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & "',null,'" & mCP(9) & "')"
         cnnConnection.Execute mStrSql
         '2022/9/29 END
      End If
      If DBDATE(mCP(6)) = strSrvDate(1) And Trim(mCP(6)) <> "" And Trim(oMailCount) <> "" Then
         '2007/8/13 MODIFY BY SONIA ¥[´¼Åv¤H­û
         'Modify By Sindy 2010/12/16 ¥[·~°È°Ï,¶O¥Î
'         PUB_SendMail strUserNum, oMailCount, "", "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¤w©¡¥»©Ò´Á­­¡A½Ð¾¨³t¿ì²z¡I", oContext2 & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(6))) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(7))) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0")
         'Modify By Sindy 2022/9/29
         'Modified by Lydia 2022/12/23 +chgsql
         'Modify By Sindy 2023/3/27 +,mc13
         mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
            " values( '" & strUserNum & "','" & oMailCount & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
            ",'" & "®×¥ó¦¬¤å ºò«æ ³qª¾--¦¹®×¤w©¡¥»©Ò´Á­­¡A½Ð¾¨³t¿ì²z¡I(¤å¸¹:" & mCP(9) & ")','" & ChgSQL(oContext2) & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(6))) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(7))) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & "',null,'" & mCP(9) & "')"
         cnnConnection.Execute mStrSql
         '2022/9/29 END
      End If
   End If
   
   'Added by Lydia 2023/08/14 §Q¯q½Ä¬ð®×¥ó¡G­­¾\®×¥ó¸É¥R±±ºÞ: ³Q­­¾\¤H­û¦b¦¬¤å®É, ¨t²ÎÀ³¦P®Éµo³qª¾µ¹¦¬¤åªÌ+¨ä¥DºÞ
   If intCaseKind <> ªk°È And intCaseKind <> ÅU°Ý Then 'Added by Lydia 2023/08/16 ­­ªA°È®×
      Call ChkCufaRight(mCP(13), mOTB(1) & mOTB(2) & mOTB(3) & mOTB(4), m_CuNo(1) & "," & m_CuNo(2) & "," & m_CuNo(3) & "," & m_CuNo(4) & "," & m_CuNo(5), mOTB(26))
   End If 'Added by Lydai 2023/08/16
   
   'Added by Lydia 2024/04/02 ªk«ß®v¦¬¤å©Ê½è¤£²Å¦X®×·½±µ¬¢³æªº®×¥ó©Ê½è¡A®×·½Ãþ§O·|¼vÅT¨ì´¼Åv¤H­ûªºÂI¼Æ¡AµoEmail³qª¾¹q¸£¤¤¤ß;
                  'L-006785ªº±µ¬¢³æ¯È¥»®×¥ó©Ê½è¬O¥Á¨Æ®×¥ó¡A¦ý¬O¤¤©Ò¦L¥X±µ¬¢³æ«á§ï®×¥ó©Ê½è¬°1101²Ä¤@¼fµ{§Ç©e¥ô«ß®v¦Óª½±µ¥H1101¦¬¤å
   If strSrvDate(1) >= ´¼¼z©Ò§ó¦W¤é And InStr(mOTB(1), "L") > 0 And m_LOS15 <> "" Then
      'Modified by Lydia 2024/04/15 debug: §ï¥Îªk«ß®×±µ¬¢³æ¸¹los17
      'mStrSql = "select * from consultreccmp where crc01 =(" & _
                "select cp140 from lawofficesource,caseprogress where los15='" & m_LOS15 & "' and los01=cp09(+)) and crc03='" & mCP(10) & "' "
      'Modified by Lydia 2024/04/16 °ê¥~³¡¨S¦³¨Ï¥Î¹q¤l¦¬¤å; ex.LIN-000266
      'mStrSql = "select a.* from consultreccmp a,lawofficesource b where los15='" & m_LOS15 & "' and los17=crc01(+) and crc03='" & mCP(10) & "' "
      mStrSql = "select crc01 as sno from consultreccmp a,lawofficesource b where los15='" & m_LOS15 & "' and los17=crc01(+) and crc03='" & mCP(10) & "' " & _
                "union select crl01 as sno from consultrecordlist a, lawofficesource b " & _
                "where los15='" & m_LOS15 & "' and los17=crl01(+) and (crl19='" & mCP(10) & "' or crl24='" & mCP(10) & "' or crl29='" & mCP(10) & "' or crl34='" & mCP(10) & "') "
      intJ = 1
      Set rsRD = ClsLawReadRstMsg(intJ, mStrSql)
      If intJ = 0 Then
         strTmp1(0) = Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
         If strTmp1(0) <> "" Then
            strTmp1(1) = "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(6))) & vbCrLf & _
                         "ªk©w´Á­­¡G" & ChangeWStringToTDateString(DBDATE(mCP(7))) & vbCrLf & _
                         "®×¥ó©Ê½è¡G" & m_CP10Name & vbCrLf & _
                         "´¼Åv¤H­û¡G" & GetStaffName(mCP(13)) & vbCrLf & _
                         "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & _
                         "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0")
            mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
               " values( '" & strUserNum & "','" & strTmp1(0) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
               ",'" & "®×¥ó¦¬¤å ºò«æ ³qª¾--" & mCP(1) & "-" & mCP(2) & IIf(mCP(3) & mCP(4) <> "000", "-" & mCP(3) & "-" & mCP(4), "") & "ªk°È®×¦¬¤å®×¥ó©Ê½è»P®×·½±µ¬¢³æ¤£¦P¡A½Ð¸ß°Ý®×·½±µ¬¢³æ¤§ªk°È¤H­û¡I(¦¬¤åÁ`¤å¸¹:" & mCP(9) & ")','" & ChgSQL(strTmp1(1)) & "',null,'" & mCP(9) & "')"
            cnnConnection.Execute mStrSql
         End If
      End If
   End If
   'end 2024/04/02
   
   'Add By Sindy 2022/9/27
   If UCase(pFormName) = UCase("frm090801_New") Then
      If Mid(mOTB(1), 1, 1) = "T" And mCP(14) <> "" Then  'Txx¦³¹w³]©Ó¿ì¤H
         '¥Ø«e¨t²Î¤é´Á®É¶¡
         strRDate = strSrvDate(1)
         strRTime = Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2)
         '­Y¹L¤U¯Z®É¶¡,ÀË¬d¤é´Á§ï¬°¤U¤@­Ó¤u§@¤Ñ,®É¶¡¹w³]¬°07:30
         'Modify By Sindy 2025/4/2 ¬°17:00«á, §ï¬°§PÂ_¹j¤é¬O§_¤H­û¥ð°²
         'If Val(Format(strRTime, "hhmm")) > 1800 Then
         If Val(Format(strRTime, "hhmm")) > 1700 Then
            strRDate = CompWorkDay(2, strSrvDate(1), 0)
            strRTime = "07:30" '¤W¤È¬O§_¦³¥ð°²
         End If
         bolCP14Rest = CheckIsPersonRest(mCP(14), strRDate, strRTime, strRestKind)
         If bolCP14Rest = True Then
            'Modify By Sindy 2025/4/2 ¤H­û¥ð°²,¤£ºÞ¬O§_¹O´Á³£¨ú®ø¹w¤À©Ó¿ì¤H,§ï¥Ñ¥DºÞ¤À®×
'            If bolIsOverDt = True Then
'               '·í¤é(¤w¹O)(±N©¡)¥»©Ò´Á­­®×¥ó, ©ÒÄÝ©Ó¿ì¤H½Ð°², ±Ä¤H¤u¤À®×
               'Add By Sindy 2025/4/2 ¤è«K¥DºÞª¾¹D³o¬O¥i¹w¤À©Ó¿ì¤H,¦ý¤H­û¥ð°²§ï¥DºÞ¤À®×
               mStrSql = "update ConsultRecCMP set CRC09='" & mCP(14) & "'" & _
                         " where CRC01='" & mCP(140) & "'" & _
                         " and CRC03='" & mCP(10) & "' and CRC08 is null and CRC09 is null"
               cnnConnection.Execute mStrSql, intI
               '2025/4/2 END
               
               mCP(14) = ""
               mStrSql = "UPDATE caseprogress SET cp14=null WHERE cp09 = '" & mCP(9) & "'"
               cnnConnection.Execute mStrSql, intI
'            Else
'               '©Ó¿ì¤H½Ð°²:µomail³qª¾ "®×¥ó¦¬¤å³qª¾,¥»©Ò´Á­­¬°xxx/xx/xx¦]©Ó¿ì¤H½Ð°²,½Ð°Æ¥»¦¬¨üªÌ¥N¬°¿ì²z¡I"¥Ø«e°Æ¥»¥[±¾³qª¾¹Å¶²©M©Ó¼z
'               'Modified by Lydia 2022/12/23 +chgsql
'               'Modify By Sindy 2023/3/27 +,mc13
'               mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
'                  " values( '" & strUserNum & "','" & mCP(14) & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
'                  ",'" & "®×¥ó¦¬¤å³qª¾,¥»©Ò´Á­­¬°" & ChangeWStringToTDateString(mCP(6)) & "¦]©Ó¿ì¤H½Ð°²,½Ð°Æ¥»¦¬¨üªÌ¥N¬°¿ì²z¡I(¤å¸¹:" & mCP(9) & ")','" & ChgSQL(oContext2) & vbCrLf & "¥»©Ò´Á­­¡G" & ChangeWStringToTDateString(mCP(6)) & vbCrLf & "ªk©w´Á­­¡G" & ChangeWStringToTDateString(mCP(7)) & vbCrLf & "´¼Åv¤H­û¡@¡G" & GetStaffName(mCP(13)) & vbCrLf & "·~°È°Ï¡@¡G" & m_SalesDeptName & vbCrLf & "¶O¥Î¡@¡@¡G" & Format(mCP(16), "##,##0") & vbCrLf & "³W¶O¡@¡@¡G" & Format(mCP(17), "##,##0") & vbCrLf & "ÂI¼Æ¡@¡@¡G" & mCP(18) & "','84027;86048','" & mCP(9) & "')"
'               cnnConnection.Execute mStrSql
'            End If
         End If

         '½T©w¤w¤À©Ó¿ì¤H
         If mCP(14) <> "" Then
            '­pºâ©Ó¿ì´Á­­
            Call PUB_CountUpdTxCP48(mCP(9), mCP(10), mCP(143), mCP(5), mCP(6), mCP(7), mCP(13), mCP(122), mOTB(1), mOTB(10), mCP(48))
         End If
      End If
      
      If ERecvSaveProgress(mCP, mOTB, strDivisionalEmp, oContext2) = False Then
         GoTo ErrHand
      End If
   End If
   
   Exit Function
   
ErrHand:
   PUB_SaveFrm010007 = False 'Add By Sindy 2022/10/25
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical, "PUB_SaveFrm010007"
   End If
End Function

'Added by Lydia 2022/09/14 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G·s¼WªA°È®×¡Bªk°È®×¡BACS®×«D112¤§¦¬¤å¡BÅU°Ý®×«D0¤§¦¬¤å(±qfrm010007.InsertOtherDatabase©â¥X¨Ó)
Private Function InsertOtherDB(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intCaseKind As Integer, ByVal intChoose As Integer, _
                ByRef mOTB() As String, ByRef mCP() As String, ByVal mCU30 As String, ByVal mChkVal As String, ByVal mSaveControl As String, Optional ByRef IsSaveData As Boolean, _
                Optional ByVal pType As String, Optional ByVal pCaseNo As String, Optional ByRef RetVal As String) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intCaseKind¡A1¬°±M§Q¡A2¬°°Ó¼Ð¡A3¬°ªk°È¡A4¬°ÅU°Ý¡A5¬°±M§Q(ªA)¡A6¬°°Ó¼Ð(ªA)¡A7¬°ªk°È(ªA)¡A8¬°ÅU°Ý(ªA)
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'mOTB¡G¨Ì¨t²Î§O¶Ç¤J°ò¥»ÀÉ(ªk°ÈLawCase¡BÅU°ÝHireCase¡BªA°ÈServicePractice)
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
'reTurnVal : ¦^¶Ç­È
'mChkVal¡G¶Ç¤J¨ä¥L¾Þ§@µ²ªG
'mSaveControl: »ô³Æ¤éºÞ¨î
Dim strAutoNumber As String
Dim np13 As String, np14 As String, bolRt As Boolean
Dim np14ForCP41 As String, np14ForCP42 As String 'Add By Sindy 2025/1/24
Dim bolError As Boolean
Dim adoquery As New ADODB.Recordset
Dim strCusReceipt As String 'Add by Amy 2018/10/11 ¦¬¾Ú¤½¥q§O
Dim rsRD As New ADODB.Recordset
Dim m_CuNo(1 To 5) As String '¥Ó½Ð¤H/·í¨Æ¤H1~5
Dim m_Na01 As String '¥Ó½Ð°ê®a
'ªk«ß©Ò®×·½¦¬¤å
Dim m_LOS01 As String '®×·½Á`¦¬¤å¸¹
Dim m_LOS01cp01 As String, m_LOS01cp02 As String, m_LOS01cp03 As String, m_LOS01cp04 As String '®×·½Á`¦¬¤å¸¹¤§¥»©Ò®×¸¹
Dim m_LOS02 As String '®×·½®×¥óÃþ«¬
Dim m_LOS15 As String '®×·½³æ¸¹
Dim m_LOS04 As String  '¤¶²Ð¤H
Dim m_LOS04_1 As String, m_LOS04_1st15 As String, m_LOS04_1st06 As String '¤¶²Ð¤H(²Ä¤@¦ì)¡B¦¬¤å³¡ªù¡B©Ò§O
Dim m_LOS05 As String  '¤¶²Ð«È¤á
Dim m_LOS12 As String  '¤¶²Ð¤é
Dim m_Los04_N1 As String, m_Los05_N As String  'LA¸É®×·½¤§¤¶²Ð¤H(²Ä¤@¦ì), ¤¶²Ð«È¤á
Dim m_bMRecvBatch As Boolean '«H¥ó¨R¾P¦h®×¦¬¤å
Dim m_bolRecvOK As Boolean 'Add By Sindy 2022/7/8 ¬O§_¦¬§¹¤å
Dim m_strMCR11 As String 'Add By Sindy 2022/7/8 ¦h®×¦¬¤å®É,²Ä¤@µ§ªºÁ`¦¬¤å¸¹
Dim m_strIR01 As String, m_strIR02 As String, m_strIR03 As String, m_strIR04 As String '«H¥ó¨R¾PPK
Dim tmpArr As Variant
Dim strCaseNo As String, strCRL01 As String 'Add By Sindy 2023/1/17
Dim rsQD As New ADODB.Recordset

'*********¯S®íºÞ¨îªºÅÜ¼Æ*************
    'Modify By Sindy 2023/5/31
    'If InStr(pType, "¥~±M«H¥ó¨R¾P") > 0 And pCaseNo <> "" Then
    If InStr(pType, "«H¥ó¨R¾P") > 0 And pCaseNo <> "" Then
    '2023/5/31 END
       'Modify By Sindy 2025/8/18
       tmpArr = Split(pCaseNo, "-")
       If InStr(pCaseNo, "-") > 0 Then
          strExc(10) = tmpArr(1)
       Else
          strExc(10) = tmpArr(0)
       End If
       tmpArr = Split(strExc(10), ",")
   '2025/8/18 END
       m_strIR01 = tmpArr(0)
       m_strIR02 = tmpArr(1)
       m_strIR03 = tmpArr(2)
       m_strIR04 = tmpArr(3)
       If InStr(pType, "¦h®×¦¬¤å") > 0 Then m_bMRecvBatch = True
       If InStr(pType, "LOS®×·½¦¬¤å") > 0 Then pType = "LOS®×·½¦¬¤å" 'Add By Sindy 2025/8/18
    'Else
    End If
    If pType = "LOS®×·½¦¬¤å" And pCaseNo <> "" Then
        m_LOS02 = Mid(pCaseNo, 1, InStr(pCaseNo, ",") - 1) '®×·½®×¥óÃþ«¬
        m_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1, 8) '®×·½³æ¸¹ 'Modify By Sindy 2025/8/18 +, 8)
        strTmp1(0) = "select X.*,cp01,cp02,cp03,cp04 from LawOfficeSource X,caseprogress where los15=" & CNULL(m_LOS15) & " and los01=cp09(+) "
        intJ = 1
        Set rsRD = ClsLawReadRstMsg(intJ, strTmp1(0))
        If intJ = 1 Then
          '®×·½Á`¦¬¤å¸¹
          m_LOS01 = "" & rsRD.Fields("LOS01")
          '®×·½Á`¦¬¤å¸¹¤§¥»©Ò®×¸¹
          m_LOS01cp01 = "" & rsRD.Fields("cp01")
          m_LOS01cp02 = "" & rsRD.Fields("cp02")
          m_LOS01cp03 = "" & rsRD.Fields("cp03")
          m_LOS01cp04 = "" & rsRD.Fields("cp04")
          '(­ì)®×·½®×¥óÃþ«¬
          m_LOS02 = "" & rsRD.Fields("LOS02")
          '®×·½³æ¸¹
          m_LOS15 = "" & rsRD.Fields("LOS15")
          '¤¶²Ð¤H, ¤¶²Ð¤H(²Ä¤@¦ì)
          m_LOS04 = "" & rsRD.Fields("LOS04")
          If m_LOS04 <> "" And InStr(m_LOS04, ",") > 0 Then
             m_LOS04_1 = Mid(m_LOS04, 1, InStr(m_LOS04, ",") - 1)
          Else
             m_LOS04_1 = m_LOS04
          End If
          If m_LOS04_1 <> "" Then
             m_LOS04_1st15 = GetST15(m_LOS04_1, , , m_LOS04_1st06)
          End If
          '(­ì)¤¶²Ð«È¤á:
          m_LOS05 = "" & rsRD.Fields("LOS05")
          '¤¶²Ð¤é
          m_LOS12 = "" & rsRD.Fields("LOS12")
        End If
        Set rsRD = Nothing
    End If
'***********************************
    Select Case intCaseKind
         Case ªk°È
            m_Na01 = mOTB(15)  '¥Ó½Ð°ê®a
            '¥Ó½Ð¤H/·í¨Æ¤H1~5
            m_CuNo(1) = mOTB(11):   m_CuNo(2) = mOTB(43):  m_CuNo(3) = mOTB(44):   m_CuNo(4) = mOTB(45):  m_CuNo(5) = mOTB(46)
         Case ÅU°Ý
            m_Na01 = "000"
            m_CuNo(1) = mOTB(5):   m_CuNo(2) = mOTB(24):  m_CuNo(3) = mOTB(25):  m_CuNo(4) = mOTB(26):  m_CuNo(5) = mOTB(27)
         Case Else 'ªA°È
            m_Na01 = mOTB(9)
            m_CuNo(1) = mOTB(8):  m_CuNo(2) = mOTB(58):   m_CuNo(3) = mOTB(59):  m_CuNo(4) = mOTB(65):  m_CuNo(5) = mOTB(66)
    End Select
    
If IsSaveData = True Then
    Exit Function
End If
IsSaveData = True

On Error GoTo ErrHand
   '¶Ç¤J0¬°­«½Æ¤§¥»©Ò®×¸¹(·s¼WÂÂ®×)¡A1¬°¥¿½T¤§¥»©Ò®×¸¹(·s¼W·s®×)
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.BeginTrans
   End If
   If intSaveMode = 1 Then
      If mOTB(2) = "" Then
         If ClsPDGetAutoNumber(mOTB(1), strAutoNumber, True, False) Then
            mOTB(2) = strAutoNumber
         Else
            bolError = True
         End If
      End If
      If bolError = False Then
         mCP(2) = mOTB(2)
         'Add by Amy 2018/10/11 ¦¬¾Ú¤½¥q§O
         'Modified by Lydia 2021/07/12 ±Æ°£ªk°È®×; ex.L-006408¤w±NLC48=J¤½¥q®³±¼
         If mOTB(1) <> "L" And mOTB(1) <> "LA" And mOTB(1) <> "LIN" And mOTB(1) <> "FCL" And mOTB(1) <> "CFL" Then
            strCusReceipt = GetReceiptCmp(Mid(m_CuNo(1), 1, 8), Mid(m_CuNo(1), 9, 1), mOTB(1), m_Na01)
         End If
         'end 2018/10/11
         Select Case intCaseKind
                      Case ªk°È
                           'Modify by Morgan 2008/8/5 +LC42
                           'Modify By Sindy 2011/1/18 +lc43,lc44,lc45,lc46
                           'Modify by Amy 2018/10/11 +lc48¦¬¾Ú¤½¥q§O
                           'Modified by Lydia 2022/09/12 +LC16,LC17 (¨Ö¤J) ¤À©Ò®×¸¹, «È¤á®×¥ó®×¸¹
                           mStrSql = "insert into lawcase (lc01,lc02,lc03,lc04,lc05,lc06,lc07,lc11,lc15,lc16,lc17,lc22,lc23,lc42,lc43,lc44,lc45,lc46,lc48) " + _
                               "values (" + CNULL(mOTB(1)) + "," + CNULL(mOTB(2)) + "," + CNULL(mOTB(3)) + "," + CNULL(mOTB(4)) + "," + CNULL(ChgSQL(mOTB(5))) + "," + _
                               CNULL(ChgSQL(mOTB(6))) + "," + CNULL(ChgSQL(mOTB(7))) + "," + CNULL(mOTB(11)) + "," + CNULL(mOTB(15)) + "," + CNULL(ChgSQL(mOTB(16))) + "," + CNULL(ChgSQL(mOTB(17))) + "," + CNULL(mOTB(22)) + "," + _
                               CNULL(mOTB(23)) + "," + CNULL(mOTB(42)) + "," + CNULL(mOTB(43)) + "," + CNULL(mOTB(44)) + "," + CNULL(mOTB(45)) + "," + CNULL(mOTB(46)) + "," + CNULL(strCusReceipt) + ")"
                           cnnConnection.Execute mStrSql
                      Case ÅU°Ý
                           'Modify by Morgan 2008/8/5 +hC23
                           'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
                           'Modified by Lydia 2022/09/12 +HC07  (¨Ö¤J) ¤À©Ò®×¸¹
                           mStrSql = "insert into hirecase (hc01,hc02,hc03,hc04,hc05,hc06,hc07,hc23,hc24,hc25,hc26,hc27) values (" + _
                               CNULL(mOTB(1)) + "," + CNULL(mOTB(2)) + "," + CNULL(mOTB(3)) + "," + CNULL(mOTB(4)) + "," + _
                               CNULL(mOTB(5)) + "," + CNULL(ChgSQL(mOTB(6))) + "," + CNULL(ChgSQL(mOTB(7))) + "," + CNULL(ChgSQL(mOTB(23))) + "," + _
                               CNULL(mOTB(24)) + "," + CNULL(mOTB(25)) + "," + CNULL(mOTB(26)) + "," + CNULL(mOTB(27)) + ")"
                           cnnConnection.Execute mStrSql
                      Case Else
                           If ClsPDGetSystemKind(mOTB(1), , , intK) Then
                              mOTB(34) = IIf(intK = 2, 2, 1)
                              'edit by nickc 2007/03/27 ¥[¤J©¼©Ò®×¸¹
                              'Modify by Morgan 2008/8/5 +SP78
                              'Modify By Sindy 2010/3/8 (¨Ö¤J)  ¼W¥[Ápµ¸¤Hsp30,sp75Äæ¦ì
                              'Modify By Sindy 2011/1/18 +sp65,sp66
                              'Moidfy by Amy 2018/10/11 +sp85  (¨Ö¤J)  ¦¬¾Ú¤½¥q§O
                              'Modified by Lydia 2022/09/12 +SP28,SP29  (¨Ö¤J)  ¤À©Ò®×¸¹, «È¤á®×¥ó®×¸¹
                              'Modify By Sindy 2024/12/30 +sp46
                              mStrSql = "insert into servicepractice (sp01,sp02,sp03,sp04,sp05,sp06,sp07,sp08,sp58,sp59,sp65,sp66,sp09,sp26, sp18,sp73,sp74,sp27,sp28,sp29,SP78,SP30,SP75,SP85,SP46) " + _
                                  "values (" + CNULL(mOTB(1)) + "," + CNULL(mOTB(2)) + "," + CNULL(mOTB(3)) + "," + CNULL(mOTB(4)) + "," + _
                                  CNULL(ChgSQL(mOTB(5))) + "," + CNULL(ChgSQL(mOTB(6))) + "," + CNULL(ChgSQL(mOTB(7))) + "," + CNULL(mOTB(8)) + "," + CNULL(mOTB(58)) + _
                                  "," + CNULL(mOTB(59)) + "," + CNULL(mOTB(65)) + "," + CNULL(mOTB(66)) + "," + CNULL(mOTB(9)) + "," + CNULL(mOTB(26)) + "," + CNULL(ChgSQL(mOTB(18))) + "," + _
                                   CNULL(mOTB(73)) + "," + CNULL(mOTB(74)) + "," + CNULL(mOTB(27)) + "," + CNULL(ChgSQL(mOTB(28))) + "," + CNULL(ChgSQL(mOTB(29))) + ", " + _
                                   CNULL(mOTB(78)) + "," + CNULL(ChgSQL(mOTB(30))) + "," + CNULL(ChgSQL(mOTB(75))) + "," + CNULL(ChgSQL(strCusReceipt)) + "," + CNULL(ChgSQL(mOTB(46))) + ")"
                              cnnConnection.Execute mStrSql
                           Else
                              bolError = True
                           End If
         End Select
         mCP(31) = "Y"
      Else
         bolError = True
      End If
   End If
   If bolError = False Then
      If ClsPDGetAutoNumber(Left(mCP(9), 1), strAutoNumber, True, True) Then
         mCP(9) = mCP(9) + strAutoNumber
         'Modify By Sindy 2025/1/24 + np14ForCP41,np14ForCP42
         bolRt = Cls001GetNextProgressData(mOTB(1), mOTB(2), mOTB(3), mOTB(4), mCP(10), np13, np14, np14ForCP41, np14ForCP42)
          
         'Add by Morgan 2008/9/23
         If mOTB(1) = "FG" Then
            mCP(48) = Pub_GetHandleDay("FG", "000", mCP(10), , mCP(6))
            mCP(20) = PUB_GetCP20(mOTB(1), mCP(10)) 'Add by Morgan 2010/4/29
         End If
        
         'Modify By Sindy 2012/11/06 +CP150 ¦³¡¹¡¹ªºÀ³¦¬±b´ÚÃ±®Ö±±ºÞ
         'Modified by Lydia 2022/09/12 (¨Ö¤J) +CP12 ·~°È°Ï§O
         'Modify By Sindy 2022/9/28 +,cp140
'         If bolRt Then
'            mStrSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp08,cp09,cp10,cp11,cp12,cp13,cp14," + _
'              "cp16,cp17,cp18,cp19,cp31,cp32,cp40,cp33,cp34,CP64,cp48,cp20,CP150,cp140) values (" + CNULL(mOTB(1)) + "," + CNULL(mOTB(2)) + "," + CNULL(mOTB(3)) + "," + CNULL(mOTB(4)) + "," + CNULL(mCP(5)) + "," + _
'              CNULL(mCP(6)) + "," + CNULL(mCP(7)) + "," + CNULL(np13) + "," + CNULL(mCP(9)) + "," + CNULL(mCP(10)) + "," + CNULL(mCP(11)) + "," + CNULL(mCP(12)) + "," + CNULL(mCP(13)) + "," + CNULL(mCP(14)) + "," + _
'              CNULL(mCP(16)) + "," + CNULL(mCP(17)) + "," + CNULL(mCP(18)) + "," + CNULL(mCP(19)) + "," + CNULL(mCP(31)) + "," + CNULL(mCP(32)) + "," + CNULL(ChgSQL(np14)) + ", " + CNULL(mCP(33)) + ", " + CNULL(mCP(34)) + "," + _
'              CNULL(ChgSQL(mCP(64))) + "," + CNULL(mCP(48), True) + ", " + CNULL(mCP(20)) + "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + ")"
'         Else
'            mStrSQL = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp12,cp13,cp14," + _
'              "cp16,cp17,cp18,cp19,cp31,cp32,cp33,cp34,CP64,cp48,cp20,CP150,cp140) values (" + CNULL(mOTB(1)) + "," + CNULL(mOTB(2)) + "," + CNULL(mOTB(3)) + "," + CNULL(mOTB(4)) + "," + CNULL(mCP(5)) + "," + _
'              CNULL(mCP(6)) + "," + CNULL(mCP(7)) + "," + CNULL(mCP(9)) + "," + CNULL(mCP(10)) + "," + CNULL(mCP(11)) + "," + CNULL(mCP(12)) + "," + CNULL(mCP(13)) + "," + CNULL(mCP(14)) + "," + _
'              CNULL(mCP(16)) + "," + CNULL(mCP(17)) + "," + CNULL(mCP(18)) + "," + CNULL(mCP(19)) + "," + CNULL(mCP(31)) + "," + CNULL(mCP(32)) + ", " + CNULL(mCP(33)) + ", " + CNULL(mCP(34)) + "," + _
'              CNULL(ChgSQL(mCP(64))) + "," + CNULL(mCP(48), True) + ", " + CNULL(mCP(20)) + "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + ")"
'         End If
         'Modify By Sindy 2023/1/12 +,cp141,cp142,cp86
         'Modify By Sindy 2023/4/18 +,cp151
         'Modify By Sindy 2023/12/12 +,cp164
         mStrSql = "insert into caseprogress (cp01,cp02,cp03,cp04,cp05,cp06,cp07,cp09,cp10,cp11,cp12,cp13,cp14," + _
              "cp16,cp17,cp18,cp19,cp31,cp32,cp33,cp34,CP64,cp48,cp20,CP150,cp140,cp141,cp142,cp86,cp151,cp164) values (" + _
              CNULL(mOTB(1)) + "," + CNULL(mOTB(2)) + "," + CNULL(mOTB(3)) + "," + CNULL(mOTB(4)) + "," + CNULL(mCP(5)) + "," + _
              CNULL(mCP(6)) + "," + CNULL(mCP(7)) + "," + CNULL(mCP(9)) + "," + CNULL(mCP(10)) + "," + CNULL(mCP(11)) + "," + CNULL(mCP(12)) + "," + CNULL(mCP(13)) + "," + CNULL(mCP(14)) + "," + _
              CNULL(mCP(16)) + "," + CNULL(mCP(17)) + "," + CNULL(mCP(18)) + "," + CNULL(mCP(19)) + "," + CNULL(mCP(31)) + "," + CNULL(mCP(32)) + ", " + CNULL(mCP(33)) + ", " + CNULL(mCP(34)) + "," + _
              CNULL(ChgSQL(mCP(64))) + "," + CNULL(mCP(48), True) + ", " + CNULL(mCP(20)) + "," + CNULL(mCP(150)) + "," + CNULL(mCP(140)) + "," + CNULL(mCP(141)) + "," + CNULL(mCP(142)) + "," + _
              CNULL(mCP(86)) + "," + CNULL(mCP(151)) + "," + CNULL(mCP(164)) + ")"
         cnnConnection.Execute mStrSql, intJ
         If bolRt Then
            'Modify By Sindy 2025/1/24 + np14ForCP41,np14ForCP42
            mStrSql = "update caseprogress set CP08=" + CNULL(np13) + _
                      ",cp40=" + CNULL(ChgSQL(np14)) + _
                      ",cp41=" + CNULL(ChgSQL(np14ForCP41)) + _
                      ",cp42=" + CNULL(ChgSQL(np14ForCP42))
            mStrSql = mStrSql + " where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
         End If
         '2023/1/12 END
         
         'MODIFY BY SONIA 2015/5/12 °ê¥~³¡¦¬¤å¤§FCL,CFL,LIN®×,©Ó¿ì¤H³£¹w³]¨t²Î¯S®í¤H­ûU2(®Û©Òªø)
         If (mOTB(1) = "FCL" Or mOTB(1) = "CFL" Or mOTB(1) = "LIN") And Left(mCP(12), 1) = "F" Then
            mStrSql = "update caseprogress set cp14='" & Pub_GetSpecMan("U2") & "' where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
         End If
'         '2015/5/12 END
         
         'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G(µÛ§@Åv)¥xÆW®×B1¡BB2¤ÎC¦¬¤å®É¡A¼W¥["®×·½³æ¸¹"Äæ¦ì¤@©w­n¿é¤J¡A¨Ã±N®×·½³æ¸¹§ó·s¦Ü¸Óµ§¦¬¤åªºCP162¡C
         If intModifyKind = 0 And "" & mOTB(9) = "000" And pType = "LOS®×·½¦¬¤å" And mOTB(1) = "TC" And m_LOS02 <> "" And m_LOS15 <> "" Then
              If Left(m_LOS02, 1) = "B" Or Left(m_LOS02, 1) = "C" Then
                  mStrSql = "update caseprogress set CP162='" & m_LOS15 & "' where cp09='" & mCP(9) & "' "
                  cnnConnection.Execute mStrSql
              End If
         End If
         'end 2020/05/20
         
         'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G¦sÀÉ®É®×·½³æ¸¹¦sCP162¡B®×·½Á`¦¬¤å¸¹(LOS01)¦sCP64Äæ"®×·½¡G¥»©Ò®×¸¹(Á`¦¬¤å¸¹)
         If strSrvDate(1) >= ªk«ß©Ò®×·½¦¬¤å±Ò¥Î¤é And pType = "LOS®×·½¦¬¤å" And InStr(mOTB(1), "L") > 0 And m_LOS15 <> "" Then
             mStrSql = " "
             If m_LOS01 <> "" And m_LOS01cp01 <> "" Then
                 mStrSql = ",cp64=" & IIf(mCP(64) <> "", "cp64||';'||", "") & CNULL("®×·½¡G" & m_LOS01cp01 & "-" & m_LOS01cp02 & IIf(m_LOS01cp03 <> "0", "-" & m_LOS01cp03, "") & IIf(m_LOS01cp04 <> "00", "-" & m_LOS01cp04, "") & "(" & m_LOS01 & ");")
             End If
             mStrSql = "update caseprogress set CP162=" & CNULL(m_LOS15) & mStrSql & " where cp09=" & CNULL(mCP(9))
             cnnConnection.Execute mStrSql
            
             'Added by Lydia 2020/06/24 ªk°È¸É¦¬´Ú78¿é¤J®×·½³æ¸¹­Y¬°B1Ãþªí¥Ü¬°A4ÂàB1¡A¦^¼g¦¬¤å¸¹¦Ü®×·½ÀÉªºªk«ß©ÒÁ`¦¬¤å¸¹Äæ2¡C
             If m_LOS02 = "B1" And mCP(10) = "78" Then
                mStrSql = "update LawOfficeSource set los21='" & mCP(9) & "' where los21 is null and los15=" & CNULL(m_LOS15)
                cnnConnection.Execute mStrSql, intJ
             Else
             'end 2020/06/24
             '¨Ã¦^¼g¦¬¤å¸¹¦Ü®×·½ÀÉªºªk«ß©ÒÁ`¦¬¤å¸¹Äæ¡C
             '5/26 ­Y¿é¤J¤§®×·½³æ¸¹¤w¦³ªk«ß©ÒÁ`¦¬¤å¸¹¥B¬°¦P®×¸¹¦P¤é¦¬¤åªÌ¡A«h¬°¦P¤@±µ¬¢³æ¤§¨ä¥L©Ê½è¡C
                mStrSql = "update LawOfficeSource set los06='" & mCP(9) & "' where los06 is null and los15=" & CNULL(m_LOS15)
                cnnConnection.Execute mStrSql, intJ
             End If 'Added by Lydia 2020/06/24
             
            'Add By Sindy 2023/1/17 ªk«ß©Ò¯È¥»¦¬¤å,Ã±®Öªº¹q¤lÀÉÂk¨÷
            strCaseNo = mCP(1) & mCP(2)
            If mCP(3) & mCP(4) <> "000" Then
                strCaseNo = strCaseNo & "-" & mCP(3)
            End If
            If mCP(4) <> "00" Then
                strCaseNo = strCaseNo & "-" & mCP(4)
            End If
            strTmp1(0) = "select los17 from LawOfficeSource where los17 is not null and los15=" & CNULL(m_LOS15)
            intJ = 1
            Set rsRD = ClsLawReadRstMsg(intJ, strTmp1(0))
            If intJ = 1 Then
                strCRL01 = rsRD.Fields("los17")
                '±µ¬¢³æ¹q¤lÀÉ§ó¦W-·s®×
                Call PUB_UpdCRLFileName(strCRL01)
                mStrSql = "update casepaperpdf set cpp01='" & mCP(9) & "',cpp10='X'" & _
                          ",cpp02='" & strCaseNo & "'||'." & mCP(10) & ".'||cpp02 where cpp11='" & strCRL01 & "'"
                cnnConnection.Execute mStrSql, intJ
            End If
            '2023/1/17 END
            
             'Added by Lydia 2020/06/10 §ó·s¨÷©v°Ï«È¤á¤å¥ó(CPP01="LOS"+LOS15)ªºÁ`¦¬¤å¸¹¬°ªk«ß®×±µ¬¢³æ¸¹(LOS17)¡AÀÉ¦W¤]­n¤@¨Ö§ó¥¿¡C
             If m_LOS02 <> "" Then
                mStrSql = ""
                Select Case Left(m_LOS02, 1)
                    Case "A"
                        mStrSql = ", cpp02=replace(cpp02,'TT999999.735.','" & mOTB(1) & mOTB(2) & IIf(mOTB(3) <> "0", "-" & mOTB(3), "") & IIf(mOTB(4) <> "00", "-" & mOTB(4), "") & "." & mCP(10) & ".') "
                    Case "B"
                        mStrSql = ", cpp02=replace(cpp02,'TT999999.736.','" & mOTB(1) & mOTB(2) & IIf(mOTB(3) <> "0", "-" & mOTB(3), "") & IIf(mOTB(4) <> "00", "-" & mOTB(4), "") & "." & mCP(10) & ".') "
                    Case "C"
                        mStrSql = ", cpp02=replace(cpp02,'TT999999.','" & mOTB(1) & mOTB(2) & IIf(mOTB(3) <> "0", "-" & mOTB(3), "") & IIf(mOTB(4) <> "00", "-" & mOTB(4), "") & "." & mCP(10) & ".') "
                End Select
                If mStrSql <> "" Then
                    mStrSql = "Update CasePaperPdf set CPP01=" & CNULL(mCP(9)) & mStrSql & " Where CPP01=" & CNULL("LOS" & m_LOS15)
                    cnnConnection.Execute mStrSql, intJ
                End If
             End If
             'end 2020/06/10
             
             '­Y¬°·s®×®Éªk°È°ò¥»ÀÉªº®×¥óÄÝ©ÊLC47¨Ì®×·½¨t²Î§O(LOS01)¦s±M§Q©Î°Ó¼Ð©ÎµÛ§@Åv
             If intCaseKind = ªk°È And intSaveMode = 1 And m_LOS01 <> "" Then
                  'Modified by Lydia 2021/09/10 §ì±µ¬¢³æ¥DÀÉªº®×¥óÄÝ©Ê
                  mStrSql = "select cp01,sk04,crl84 from caseprogress,systemkind, Consultrecordlist " & _
                              "where cp09=" & CNULL(m_LOS01) & " and cp01=sk01(+) and cp140=crl01(+) "
                  intJ = 1
                  Set rsRD = ClsLawReadRstMsg(intJ, mStrSql)
                  If intJ = 1 Then
                      mStrSql = ""
                      If InStr("" & rsRD.Fields("cp01"), "P") > 0 Then
                          mStrSql = " LC47='±M§Q' "
                      'Modified by Lydia 2020/07/08
                      'ElseIf InStr("TC,CFC", "" & rsrd.Fields("cp01")) > 0 Then
                      ElseIf "" & rsRD.Fields("CP01") = "TC" Or "" & rsRD.Fields("CP01") = "CFC" Then
                          mStrSql = " LC47='µÛ§@Åv' "
                      ElseIf InStr("" & rsRD.Fields("cp01"), "T") > 0 And "" & rsRD.Fields("cp01") <> "TT" Then
                          mStrSql = " LC47='°Ó¼Ð' "
                      'Added by Lydia 2021/09/10 ®×·½¬°TT®É¡A¹w³]ªk°È°ò¥»ÀÉªº®×¥óÄÝ©ÊLC47¬°TT±µ¬¢³æ¤§CRL84ªk°È®×¥óÄÝ©Ê¡Cex.L-006435­è¦¬¤åªº®×¥óÄÝ©Ê¬°ªÅ¥Õ
                      ElseIf "" & rsRD.Fields("cp01") = "TT" Then
                          mStrSql = " LC47='" & ChgSQL("" & rsRD.Fields("crl84")) & "' "
                      'end 2021/09/10
                      End If
                      If mStrSql <> "" Then
                            mStrSql = "Update Lawcase Set " & mStrSql & " Where lc01=" + CNULL(mOTB(1)) + " and lc02=" + CNULL(mOTB(2)) + " and lc03=" + CNULL(mOTB(3)) + " and lc04=" + CNULL(mOTB(4))
                            cnnConnection.Execute mStrSql
                      End If
                  End If
                  'Added by Lydia 2020/07/14  ·¨¥@¦w¡Gªk«ß©Ò±µ®×«á¦Û°Ê·s¼W¬ÛÃö¨÷¸¹¸ê®Æ(Âù¦V®×¸¹)¡F
                  If Left(m_LOS02, 1) = "B" Or Left(m_LOS02, 1) = "C" Then 'Added by Lydia 2020/07/29 ±Æ°£TT®×
                      mStrSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(mOTB(1)) & ", " & CNULL(mOTB(2)) & ", " & CNULL(mOTB(3)) & ", " & CNULL(mOTB(4)) & ", " & CNULL(m_LOS01cp01) & ", " & CNULL(m_LOS01cp02) & ", " & CNULL(m_LOS01cp03) & ", " & CNULL(m_LOS01cp04) & " ) "
                      cnnConnection.Execute mStrSql
                      mStrSql = "insert into caserelation1(cr01,cr02,cr03,cr04,cr05,cr06,cr07,cr08) values (" & CNULL(m_LOS01cp01) & ", " & CNULL(m_LOS01cp02) & ", " & CNULL(m_LOS01cp03) & ", " & CNULL(m_LOS01cp04) & ", " & CNULL(mOTB(1)) & ", " & CNULL(mOTB(2)) & ", " & CNULL(mOTB(3)) & ", " & CNULL(mOTB(4)) & " ) "
                      cnnConnection.Execute mStrSql
                  End If 'Added by Lydia 2020/07/29 ±Æ°£TT®×
                  'end 2020/07/14
             End If
             '¦¬¤å·s®×¥B¤£¦P¼f¯ÅªºCÃþ®×·½®É¡AÀË¬d¬Û¦PLC02ªº®×¸¹­Y¤w¦³BÃþ®×·½®ÉEmail³qª¾¨q¬Â­n½Õ¾ã¸Óªk°È®×ªº¼f¯Å¶¶§Ç¡C
             If intCaseKind = ªk°È And intSaveMode = 1 And Val(mOTB(3)) >= 1 And Left(m_LOS02, 1) = "C" Then
                mStrSql = "select los06 from LawOfficeSource where los06 in (select cp09 from caseprogress where cp01='" & mOTB(1) & "' and cp02='" & mOTB(2) & "' and cp03<='" & Val(mOTB(3)) - 1 & "' and cp04='" & mOTB(4) & "' and cp159=0) and los02 like 'B%' "
                intJ = 1
                Set rsRD = ClsLawReadRstMsg(intJ, mStrSql)
                If intJ = 1 Then
                    'Modify By Sindy 2022/9/29 §ï§ì Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û")
                    'Modify By Sindy 2023/3/27 +,mc13
                    mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc13)" & _
                             " values ('" & strUserNum & "','" & Pub_GetSpecMan("µ{¦¡ºÞ²z¤H­û") & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss')" & _
                             ",'¦¬¤å" & mOTB(1) & "-" & mOTB(2) & "-" & mOTB(3) & "-" & mOTB(4) & "¦³¤£¦P¼f¯ÅªºBÃþ®×·½¡A½ÐÀË¬d¸Óªk°È®×ªº¼f¯Å¶¶§Ç¡I','¦PºK­n','" & mCP(9) & "')"
                    cnnConnection.Execute mStrSql, intJ
                End If
             End If
             
             '­Y®×·½¸ê®Æªº¤¶²Ð«È¤áLOS05¬°ªÅ®Éªí¥Ü·s«È¤á­n¦^¼g¨Ã§ó·s(¦¬¤å®É¿é¤Jªº)«È¤á´¼Åv¤H­û(CU12CU13)¬°¤¶²Ð¤H(LOS04²Ä¤@¤H)
             If m_LOS05 = "" And Trim(m_CuNo(1) & m_CuNo(2) & m_CuNo(3) & m_CuNo(4) & m_CuNo(5)) <> "" Then
                 '¨Ã¥B¦^¼g®×·½¤¶²Ð«È¤á½s¸¹LOS05
                 mStrSql = "update LawOfficeSource set los05='" & m_CuNo(1) & "' where los05 is null and los15=" & CNULL(m_LOS15)
                 cnnConnection.Execute mStrSql, intJ
                 If intJ > 0 Then
                    For intJ = 1 To 5
                        If Trim(m_CuNo(intJ)) <> "" Then
                             mStrSql = "update customer set cu12='" & m_LOS04_1st15 & "',cu13='" & m_LOS04_1 & "' where cu01='" & Left(m_CuNo(intJ), 8) & "' and cu02='" & Right(m_CuNo(intJ), 1) & "'"
                             Pub_SeekTbLog mStrSql
                             cnnConnection.Execute mStrSql
                        End If
                    Next intJ
                 End If
                 'Added by Lydia 2022/10/19 «È¤á½s¸¹«á«Ø=m_LOS05=ªÅ¥Õ; ex.L-006577
                 m_LOS05 = m_CuNo(1)
                 RetVal = m_LOS05 & "|" & m_LOS04_1
                 'end 2022/10/19
             End If
             
             '³Ì«á¤~°µ-->«È¤á½s¸¹¦^¼g«á¡A®×·½®×¥óÃþ«¬A¡A­YµLÂI¼Æ«h«O¯dÃþ«¬A¡A­Y¦³ÂI¼Æ«h§PÂ_¦P¤@«È¤á½s¸¹¤¶²Ð¤é«e­Y¦³A1«h¦¹µ§³]¬°A2¡A­YµL«h³]¬°A1¡C
                                 '­pºâ®×·½¤§¶O¥Î¤ÎÂI¼Æ¡A§ó·s¦^®×·½Á`¦¬¤å¸¹LOS01¤§¶O¥Î¤ÎÂI¼Æ¡A¥H§Q´¼¼z©Ò¶}¥ß¦¬¾Ú¡C
                                 '®×·½¬°TT-999999®É¦P®É¤Wµo¤å¤éCP27¬°¨t²Î¤é(¬°µLµo¤å¤éªÌ¤~§ó·s)¡C
                                 '5/6¸ò·¨ºÊ¹î¤H½T»{°ê¥~³¡¤¶²Ð®×·½¥H¬Û¦P¤À¼í¤è¦¡­pºâ¡A¤£ºÞ°ê¥~¥N²z¤H¤´¥H«È¤á¬°¤¶²Ð°ò·Ç¡C
             'Modified by Lydia 2020/06/23 ©Ê½èÄÝ©óB1ÂkÄÝ©óAÃþ(A4)
             'If m_LOS02 = "A" And Val(mcp(18)) > 0 Then
             If (m_LOS02 = "A" Or m_LOS02 = "A4") And Val(mCP(18)) > 0 Then
                If m_LOS02 = "A" Then
             'end 2020/06/23
                    
                    'Modified by Morgan 2020/9/24 ¥þ³¡¥H²Ä¤@¦¸¤À¼íA1¡A¨Ò¥~¤~ºâ²Ä¤G¦¸A2¡A¬O«ü«áÄò¥Ñªk«ß©Òª½±µ¦¬¤å(¤£¬O¤¶²Ð¤H¦A¨«®×·½¬yµ{)®É«h¥HA2­pºâ
                    'mStrSQL = "select los02 from LawOfficeSource where los12<'" & m_LOS12 & "' and los02='A1' and los05='" & IIf(m_LOS05 <> "", m_LOS05, mOTB(11)) & "' "
                    'intJ = 1
                    'Set rsrd = ClsLawReadRstMsg(intJ, mStrSQL)
                    'If intJ = 1 Then
                    '    mStrSQL = "update LawOfficeSource set los02='A2' where los15='" & m_LOS15 & "' "
                    '    cnnConnection.Execute mStrSQL
                    'Else
                        mStrSql = "update LawOfficeSource set los02='A1' where los15='" & m_LOS15 & "' "
                        cnnConnection.Execute mStrSql
                    'End If
                    'end 2020/9/24
                    
                End If 'Added by Lydia 2020/06/23
                
                '®×·½¬°TT-999999®É¦P®É¤Wµo¤å¤éCP27¬°¨t²Î¤é(¬°µLµo¤å¤éªÌ¤~§ó·s)¡C
                If m_LOS01cp01 & m_LOS01cp02 = "TT999999" Then
                    mStrSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp09='" & m_LOS01 & "' and nvl(cp27,0)=0 "
                    cnnConnection.Execute mStrSql
                End If
             End If
             
             '­pºâ®×·½¤§¶O¥Î¤ÎÂI¼Æ¡A§ó·s¦^®×·½Á`¦¬¤å¸¹LOS01¤§¶O¥Î¤ÎÂI¼Æ¡A¥H§Q´¼¼z©Ò¶}¥ß¦¬¾Ú¡C
             PUB_UpdateTTFee m_LOS15 'Added by Morgan 2020/9/29 ¦P®×·½³æ¸¹ªº¨C­Ó¦¬¤å©Ê½è³£­n(¶O¥Î¥[Á`)
         End If
         'end 2020/05/20
         
         'Added by Lydia 2020/05/20 ªk«ß©Ò®×·½¦¬¤å¡G­Y¸Ó¦¬¤å¸¹ÂI¼Æ>0¦ýµL®×·½(¦Û¦æ¦¬¤åªÌ)®É¡A­Y®×¥óªº«È¤á¬°«Dªk«ß©Òªº«È¤á®É«h¤´ºâAÃþ®×·½(¥t¼g¨ç¼Æ°Ñ·Ó§@±b³W«h³]©w¬°A1~A4)¡C
                                            '¨t²Î¦Û°Ê·s¼WTT-999999®×¶i«×(BÃþ¦¬¤å)¤Îªk«ß©Ò®×·½¸ê®Æ(¦P³Ì«á¤@µ§®×·½ªº¸ê®Æ)¡C
         'Modified by Lydia 2020/10/05 +ÂÂ®×mOTB(2) <> ""
         'Modified by Lydia 2021/01/08 ®³±¼¥xÆW®×ªº­­¨î And txtOther(3) = "000"
         If strSrvDate(1) >= ªk«ß©Ò®×·½¦¬¤å±Ò¥Î¤é And InStr(mOTB(1), "L") > 0 And m_LOS15 = "" And Val(mCP(18)) > 0 And mOTB(2) <> "" Then
             'Modified by Lydia 2020/10/05 + st01
             mStrSql = "select cu01,cu02,st15,st01 from customer,staff where cu01='" & Mid(m_CuNo(1), 1, 8) & "' and cu02='" & Mid(m_CuNo(1), 9, 1) & "'  and cu13=st01(+) "
             Set rsRD = ClsLawReadRstMsg(intJ, mStrSql)
             If intJ = 1 Then
                 strTmp1(1) = Left("" & rsRD.Fields("st15"), 1)
                 If strTmp1(1) <> "L" Then
                    '«Dªk«ß©ÒªºÂÂ«È¤á®É«h¤´ºâAÃþ®×·½(¥t¼g¨ç¼Æ°Ñ·Ó§@±b³W«h³]©w¬°A1~A4)
                    'Modified by Lydia 2020/10/05 (9/30) ¦p¬°ÂÂ®×¨Ã¥B´¿¦³A1Ãþ®×·½®É¡A«h¬°A2Ãþ®×·½
                    'mStrSQL = "select * from LawOfficeSource where los12<'" & strSrvDate(1) & "' and los02 like 'A%' and los05='" & mOTB(11) & "' " & _
                                "order by los12 desc, los13 desc "
                    'Modified by Lydia 2021/01/06 ¥H³Ì«á¤@¹D®×·½¬°·Ç
                    'mStrSQL = "Select * From Lawofficesource Where los02='A1' and los07||los08 is null and Los15 In " & _
                                 "(select max(cp162) from caseprogress where cp01='" & mOTB(1) & "' and cp02='" & mOTB(2) & "' and cp162 is not null) "
                    'Modified by Lydia 2022/11/09 ¼W¥[LA®×ªºA3®×·½
                    'mStrSQL = "Select * From Lawofficesource Where los02='A1' and los07||los08 is null and Los15 In " & _
                                 "(select cp162 from caseprogress where cp01='" & mOTB(1) & "' and cp02='" & mOTB(2) & "' and cp162 is not null) order by los12 desc "
                    mStrSql = "Select * From Lawofficesource Where los02 in ('A1','A3') and los07||los08 is null and Los15 In " & _
                                 "(select cp162 from caseprogress where cp01='" & mOTB(1) & "' and cp02='" & mOTB(2) & "' and cp162 is not null) order by los12 desc "
                    intJ = 1
                    Set rsRD = ClsLawReadRstMsg(intJ, mStrSql)
                    If intJ = 1 Then
                       rsRD.MoveFirst
                       '®×·½Ãþ§O
                       'Modified by Lydia 2020/10/05 ©T©w¬°A2Ãþ®×·½
                        'If "" & rsrd.Fields("los02") = "A" Then
                        '    strTmp1(2) = "A1"
                        'Else
                        '    strTmp1(2) = "A2"
                        'End If
                        'Added by Lydia 2022/11/09 ¼W¥[LA®×ªºA3®×·½
                        If mOTB(1) = "LA" Then
                           strTmp1(2) = "A3"
                        Else
                        'end 2022/11/09
                           strTmp1(2) = "A2"
                           'end 2020/10/05
                        End If 'Added by Lydia 2022/11/09
                        'TT·s¼WBÃþ¦¬¤å
                        'Modified by Lydia 2020/10/05
                        strTmp1(3) = ""
                        If "" & rsRD.Fields("LOS04") <> "" Then '§ì¤¶²Ð¤H1
                           'Modified by Lydia 2020/10/05 ¥Î©ó®×·½¤§±µ¬¢¤H¨ú±o¦bÂ¾­û¤u½s¸¹©M¤¶²Ð¤H²Ä¤@¤H
                           strTmp1(5) = PUB_GetNowStaff("" & rsRD.Fields("los04"), strTmp1(3))
                           'end 2020/10/05
                        End If
                        'Added by Lydia 2020/10/05
                        m_Los04_N1 = strTmp1(3)
                        If strTmp1(3) <> "" Then
                            strTmp1(1) = AutoNo("B", 6) 'TT¦¬¤å¸¹ 'Move  by Lydia 2020/10/07 ±qrsrd.MoveFirst¤U¤è²¾¹L¨Ó
                        'end 2020/10/05
                            'Modified by Morgan 2021/1/8 +CP20,CP27,CP32
                            mStrSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp06,cp09,cp10,cp11,cp12,cp13,cp20,cp27,cp32,CP162)" & _
                               " values('TT','999999','0','00'," & strSrvDate(1) & "," & CNULL(mCP(6), True) & ",'" & strTmp1(1) & "'" & _
                               ",'735','07','" & GetST15(strTmp1(3)) & "','" & strTmp1(3) & "','N'," & strSrvDate(1) & ",'N',null)"
                            cnnConnection.Execute mStrSql
                            'ªk«ß©Ò®×·½¸ê®Æ(¦P³Ì«á¤@µ§®×·½ªº¸ê®Æ), ®×·½³æ¸¹=TTÁ`¦¬¤å¸¹
                            strTmp1(4) = AutoNo("LOS", 5, , True)
                            'Modified by Lydia 2020/10/05
                            mStrSql = "insert into LawOfficeSource(LOS01,LOS02,LOS03,LOS04,LOS05,LOS06,LOS10,LOS11,LOS12,LOS13,LOS15)" & _
                               " values ('" & strTmp1(1) & "','" & strTmp1(2) & "' ,'" & mCP(13) & "'" & _
                               ",'" & strTmp1(5) & "','" & m_CuNo(1) & "','" & mCP(9) & "','" & strTmp1(1) & "'" & _
                               ",'" & strUserNum & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss'),'" & strTmp1(4) & "')"
                            m_Los05_N = m_CuNo(1)
                            'end 2020/10/05
                            cnnConnection.Execute mStrSql
                            'Added by Lydia 2020/10/05 ¦¬¤å¤§¶i«×¥[µù®×·½
                            'Modified by Lydia 2021/01/08 ¸É¤W®×·½³æ¸¹CP162
                            mStrSql = "Update CaseProgress Set cp64=" & CNULL("®×·½¡GTT-999999(" & strTmp1(1) & ");") & "||cp64,CP162=" & CNULL(strTmp1(4)) & " where cp09=" & CNULL(mCP(9))
                            cnnConnection.Execute mStrSql
                        
                            '­pºâ®×·½¤§¶O¥Î¤ÎÂI¼Æ¡A§ó·s¦^®×·½Á`¦¬¤å¸¹LOS01¤§¶O¥Î¤ÎÂI¼Æ¡A¥H§Q´¼¼z©Ò¶}¥ß¦¬¾Ú¡C
                            PUB_UpdateTTFee strTmp1(4) 'Added by Morgan 2020/9/29
                            RetVal = m_Los05_N & "|" & m_Los04_N1
                        End If 'Added by Lydia 2020/10/05
                    End If
                 End If
             End If
         End If
         'end 2020/05/20
                    
           'add by nickc 2008/01/04 ¥[¤J¦^¥N®É¡A©Ó¿ì´Á­­¬°¥»©Ò¦¬¤å¤é(·í¤Ñ¤£ºâ)¤§²Ä¤G­Ó¤u§@¤Ñ
           If mCP(10) = "720" Then
               'Modify by Morgan 2008/9/23 FG §ï¦b¤W­±³]©w
               If mOTB(1) <> "FG" Then
                  mStrSql = "update caseprogress set cp48=" + CNULL(CompWorkDay(3, mCP(5), 0)) + " where cp09=" + CNULL(mCP(9))
                  cnnConnection.Execute mStrSql
               End If
           End If
   
           'Add By nickc 2007/08/21
           '­Y¬°±µ¬¢°O¿ý³æ(Âd¥x¦¬¤å)
           'Modify by Morgan 2007/10/26 ¶O¥Î¥i§ï®É¤~°µ¡A§_«h¤w¦¬´Ú¸ê®Æ·|³QÁÙ­ì
           If intChoose = 0 And mCP(60) = "" Then  'mCP(60) = "" => txtOther(12).Enabled = True
           'end 2007/10/26
               '¥¼¦¬ª÷ÃB = ¶O¥Î
               mStrSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(mCP(9))
               cnnConnection.Execute mStrSql
           End If
         'Add By Cheng 2002/05/10
         '­Y¬°¤º³¡¦¬¤å§@·~®É, ®×¥ó¶i«×ÀÉªº¬O§_¦V«È¤á¦¬´Ú³]©w¬°"N"
         If intChoose = 1 Then
            mStrSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
         End If
         
         'Modify By Sindy 2023/11/3 mark:¦¬¤å¤w¤£»Ý°õ¦æ¦¹¬qµ{¦¡,¦]±µ¬¢³æ¤w·|¦^¼g¶l»¼°Ï¸¹
'         mStrSql = "update customer set cu30=" + CNULL(mCU30) + " where cu01=" + CNULL(Mid(m_CuNo(1), 1, 8)) + " and cu02=" + CNULL(Mid(m_CuNo(1), 9, 1))
'         cnnConnection.Execute mStrSql
         '2023/11/3 END
         
           'Add by Lydia 2014/10/31 ¶}©ñ¥~±Mµ{§Ç¤H­û¥i¶i¤J±M§Q³B¨t²Î¾Þ§@FMP¾ÈµØ®×¥ó
           If InStr(mChkVal, "¾ÈµØ®×¥ó½T»{") > 0 Then
               mStrSql = "update caseprogress set cp44='Y53374000' where cp09='" & mCP(9) & "' "
               cnnConnection.Execute mStrSql
           End If
           'end. 'Add by Lydia 2014/10/31
           
         '  'Modify By Cheng 2002/05/10
         '  '­Y¦b¤U¤@µ{§ÇÀÉ¥u§ì¨ì¤@µ§¸ê®Æ®É, ¤~­n§ì¤U¤@µ{§ÇÀÉªºÁ`¦¬¤å¸¹§ó·s®×¥ó¶i«×ÀÉªº¬ÛÃöÁ`¦¬¤å¸¹
         adoquery.CursorLocation = adUseClient
         adoquery.Open "select np01 from nextprogress where np02 = '" & mOTB(1) & "' and np03 = '" & mOTB(2) & "' and np04 = '" & mOTB(3) & "' and np05 = '" & mOTB(4) & "' and np06 is null and np07 = '" & mCP(10) & "'", cnnConnection, adOpenStatic, adLockReadOnly
         If adoquery.RecordCount > 0 Then
            If adoquery.RecordCount = 1 Then
               If IsNull(adoquery.Fields(0).Value) = False Then
                  '2011/6/17 add by sonia ²§Ä³µªÅG¡Bµû©wµªÅG¡B¼o¤îµªÅG­n¤@¨Ã§ó·s¹ï³y¸ê®Æ
                  If (mCP(10) = "602" Or mCP(10) = "604" Or mCP(10) = "606") Then
                     cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp80) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42,b.cp80 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & mCP(9) & "'", intJ
                  Else
                  '2011/6/17 end
                     cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & mCP(9) & "'"
                  End If  '2011/6/17 add by sonia
               End If
               'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(mCP(9)) & "
               mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np02=" + CNULL(mOTB(1)) + " and np03=" + _
                  CNULL(mOTB(2)) + " and np04=" + CNULL(mOTB(3)) + " and np05=" + CNULL(mOTB(4)) + _
                  " and np07=" + CNULL(mCP(10)) + " and np06 is null"
               cnnConnection.Execute mStrSql
            End If
         Else
            adoquery.Close
            adoquery.CursorLocation = adUseClient
            adoquery.Open "select np01 from nextprogress where np02 = '" & mOTB(1) & "' and np03 = '" & mOTB(2) & "' and np04 = '" & mOTB(3) & "' and np05 = '" & mOTB(4) & "' and np06 <>'Y' and np07 = '" & mCP(10) & "'", cnnConnection, adOpenStatic, adLockReadOnly
            If adoquery.RecordCount > 0 Then
               If adoquery.RecordCount = 1 Then
                  If IsNull(adoquery.Fields(0).Value) = False Then
                     '2011/6/17 add by sonia ²§Ä³µªÅG¡Bµû©wµªÅG¡B¼o¤îµªÅG­n¤@¨Ã§ó·s¹ï³y¸ê®Æ
                     If (mCP(10) = "602" Or mCP(10) = "604" Or mCP(10) = "606") Then
                        cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp80) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42,b.cp80 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & mCP(9) & "'", intJ
                     Else
                     '2011/6/17 end
                        cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & mCP(9) & "'"
                     End If  '2011/6/17 add by sonia
                  End If
                  'Modify By Sindy 2016/11/2 + ,np24=" & CNULL(mCP(9)) & "
                  mStrSql = "update nextprogress set np06='Y',np24=" & CNULL(mCP(9)) & " where np02=" + CNULL(mOTB(1)) + " and np03=" + _
                     CNULL(mOTB(2)) + " and np04=" + CNULL(mOTB(3)) + " and np05=" + CNULL(mOTB(4)) + _
                     " and np07=" + CNULL(mCP(10)) + " and np06 <> 'Y'"
                  cnnConnection.Execute mStrSql
               End If
            End If
         End If
         adoquery.Close
         
         'Add By Sindy 2024/4/11 MCT¦³©¼©Ò®×¸¹®É,­n§ó·s¨äÄæ¦ì
         Dim strUpdOTBCol As String
         If mOTB(27) <> "" And mOTB(9) = "000" And (intCaseKind <> ªk°È And intCaseKind <> ÅU°Ý) _
            And intSaveMode = 0 Then
            strUpdOTBCol = strUpdOTBCol & ",sp27=" + CNULL(mOTB(27))
         End If
         If strUpdOTBCol <> "" Then
            strUpdOTBCol = Mid(strUpdOTBCol, 2)
            mStrSql = "update ServicePractice set " & strUpdOTBCol
            mStrSql = mStrSql & " where sp01=" + CNULL(mOTB(1)) + " and sp02=" + CNULL(mOTB(2)) + " and sp03=" + CNULL(mOTB(3)) + " and sp04=" + CNULL(mOTB(4))
            cnnConnection.Execute mStrSql
         End If
         '2024/4/11 END
         
         'Add By Sindy 2024/7/12 ÀË¬d¬O§_¦³¨R¤U¤@µ{§Çªº´Á­­
         '                       ­Y¦³,¦¹µ§¦¬¤å´Á­­§ó·s¬°¤U¤@µ{§Çªº´Á­­
         If UCase(pFormName) = UCase("frm090801_New") Then
            strTmp1(0) = "select np01,np08,np09 from nextprogress where np24 = '" & mCP(9) & "'"
            intJ = 1
            Set rsQD = ClsLawReadRstMsg(intJ, strTmp1(0))
            If intJ = 1 Then
               mStrSql = "update caseprogress set CP06=" & "" & rsQD.Fields("np08") & _
                                                ",CP07=" & "" & rsQD.Fields("np09")
               mStrSql = mStrSql & " where cp09=" & CNULL(mCP(9))
               cnnConnection.Execute mStrSql
            End If
         End If
         '2024/7/12 END
         
         'Add By Sindy 2024/1/4 ¹q¤l¦¬¤å¤£¶·°õ¦æ¦¹¨ç¼Æ
         If UCase(pFormName) <> UCase("frm090801_New") Then
         '2024/1/4 END
            Select Case intCaseKind
                 Case ªk°È
                    If Cls001SetCaseProgressFee(mOTB(1), mOTB(15), mCP(10), mCP(9)) = False Then bolError = True
                 Case ÅU°Ý
                    If Cls001SetCaseProgressFee(mOTB(1), "000", mCP(10), mCP(9)) = False Then bolError = True
                 Case Else 'ªA°È
                    If Cls001SetCaseProgressFee(mOTB(1), mOTB(9), mCP(10), mCP(9)) = False Then bolError = True
            End Select
         End If
      Else
         bolError = True
      End If
   End If

   'add by nickc 2008/05/02 Àx¦s¹w©w¦¬´Ú¤é
   'Remove by Lydia 2018/08/22 (À³¦¬±b´ÚºÞ±±)¨ú®ø¹w©w¦¬´Ú¤é,§ï¦¨¥I´Ú¶g´Á
'   If bolError = False Then
'       Dim rtCnt As Integer
'       'Modify by Morgan 2010/12/9
'       'If txtOther(28) <> "" Then
'       '    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtOther(28)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtOther(28)) & " ", rtCnt
'       If txtOther(28) <> "" And txtOther(28) <> txtOther(28).Tag Then
'           cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtOther(28)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
'       'end 2010/12/9
'           If rtCnt = 0 Then
'               cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtOther(28)) & " from dual "
'           End If
'       End If
'   End If
   'end 2018/08/22
   
   'Added by Lydia 2024/12/13 ¼W¥[FCP/P/FG®×¸¹®Éªº¨t²Î³qª¾ (½Ð¹q¸£¤¤¤ß¤ñ·Óªþ¥ó·s®×¥ß¨÷PUB_GetTCTmail³qª¾©Ó¿ì¤Î¬ÛÃö¤H­û)
   If intSaveMode = 1 And intModifyKind = 0 And mOTB(1) = "FG" And InStr(pType, "°lÂÜ¬y¤ô¸¹") > 0 And pCaseNo <> "" Then
      mStrSql = "Update TrackingCaseName set TCN05=" & CNULL(mCP(9)) & " Where TCN01 ='" & pCaseNo & "' "
      cnnConnection.Execute mStrSql
      Call Proc_FCPNewCaseEmail(mOTB(1), mOTB(2), mOTB(3), mOTB(4), mCP(9), mCP(10), mCP(12))
   End If
   'end 2024/12/13
   
   'Add by Sindy 2022/8/17
   'Modify By Sindy 2023/5/31
   'If InStr(pType, "¥~±M«H¥ó¨R¾P") > 0 And m_strIR01 <> "" Then
   'Modify By Sindy 2025/8/18
   'If InStr(pType, "«H¥ó¨R¾P") > 0 And m_strIR01 <> "" Then
   If m_strIR01 <> "" Then
   '2025/8/18 END
   '2023/5/31 END
      m_bolRecvOK = True
      m_strMCR11 = ""
      If m_bMRecvBatch = True Then '¦h®×¦¬¤å
         '§ó·sÁ`¦¬¤å¸¹
         mStrSql = "update multiCaseRecv set mcr11='" & mCP(9) & "'" & _
                  " where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                  " and mcr02='" & mOTB(1) & "' and mcr03='" & mOTB(2) & "' and mcr04='" & mOTB(3) & "' and mcr05='" & mOTB(4) & "'" & _
                  " and mcr06='" & mCP(10) & "'"
                  cnnConnection.Execute mStrSql
                  
         'Modify By Sindy 2022/8/26
         '¤U¸ü«H¥óÀÉ,¤W¶Ç¨÷©v°Ï
         Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, mCP(9))
         
         'ÀË¬d¦h®×¦¬¤åª¬ªp
         strTmp1(0) = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                     " and mcr02||mcr03||mcr04||mcr05<>'" & mOTB(1) & mOTB(2) & mOTB(3) & mOTB(4) & "'" & _
                     " and mcr11 is null"
         intJ = 1
         Set rsRD = ClsLawReadRstMsg(intJ, strTmp1(0))
         If intJ = 1 Then
            m_bolRecvOK = False '©|¦³¥¼¦¬¤å
            
            'Modify By Sindy 2022/8/26 ¦¹³BMark,µ{¦¡©¹¤W²¾
'            '¤U¸ü«H¥óÀÉ,¤W¶Ç¨÷©v°Ï
'            Call PUB_UploadPatentLetterFile(m_strIR01, m_strIR03, CP09)
         Else
            m_bolRecvOK = True '¬O§_¦¬§¹¤å
            '§ì²Ä¤@µ§ªºÁ`¦¬¤å¸¹
            strTmp1(0) = "select * from multiCaseRecv where mcr01='" & m_strIR01 & m_strIR03 & "'" & _
                        " and mcr02||mcr03||mcr04||mcr05=mcr07||mcr08||mcr09||mcr10 and mcr11 is not null"
            intJ = 1
            Set rsRD = ClsLawReadRstMsg(intJ, strTmp1(0))
            If intJ = 1 Then
               m_strMCR11 = rsRD.Fields("mcr11")
               RetVal = RetVal & IIf(RetVal <> "", ",", "") & "MCR11:" & m_strMCR11
            Else
               MsgBox "¦h®×¦¬¤å¡AµLÅª¨ú¨ì²Ä¤@µ§®×¥óªºÁ`¦¬¤å¸¹¡A½Ð¬¢¹q¸£¤¤¤ß!!", vbExclamation '¦¹ª¬ªpÀ³¤£·|µo¥Í, ¥H¨¾¥~¤@
               GoTo ErrHand
            End If
         End If
      End If
      If m_bolRecvOK = True Then '¬O§_¦¬§¹¤å=>¥þ³¡¦¬§¹¤å
         RetVal = RetVal & IIf(RetVal <> "", ",", "") & "m_bolRecvOK = True"
         '¦h®×¦¬¤åªºÁ`¦¬¤å¸¹­n¶Ç¤J²Ä¤@µ§Á`¦¬¤å¸¹
         PUB_UpdateEMailRec m_strIR01, m_strIR02, m_strIR03, m_strIR04, _
               IIf(m_strMCR11 <> "", "¦h®×¦¬¤å", "frm010001"), _
               IIf(m_strMCR11 <> "", m_strMCR11, mCP(9))
      End If
   End If
   '2022/8/17 END
   
   If bolError Then
      'Add By Sindy 2022/9/27
      If UCase(pFormName) <> UCase("frm090801_New") Then
      '2022/9/27 END
         cnnConnection.RollbackTrans
      End If
      ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
      IsSaveData = False
   Else
      If mSaveControl <> "" Then
           'Modified by Lydia 2022/09/29 ¶Ç¤J¨t²Î§O,°ê®a,®×¥ó©Ê½è=> mOTB(1), m_Na01, mCP(10)
           Call PUB_SaveByControl(mCP(9), mSaveControl, mOTB(1), m_Na01, mCP(10))
      End If
      
      'Add By Sindy 2022/9/27
      If UCase(pFormName) <> UCase("frm090801_New") Then
      '2022/9/27 END
         cnnConnection.CommitTrans
      End If
      InsertOtherDB = True
      'mOTB(2) = mOTB(2) '²¾¨ì¥~¼h
   End If
   'mOTB(2) = mOTB(2) '²¾¨ì¥~¼h
   
   Set rsQD = Nothing
   Set adoquery = Nothing
   Set rsRD = Nothing
   Exit Function
   
ErrHand:
   'Add By Sindy 2022/9/27
   If UCase(pFormName) <> UCase("frm090801_New") Then
   '2022/9/27 END
      cnnConnection.RollbackTrans
   End If
   Set adoquery = Nothing
   Set rsRD = Nothing
   
   ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf
   IsSaveData = False
End Function

'Added by Lydia 2022/09/14 Âd¥x¦¬¤å¼Ò²Õ¤Æ¡G­×§ïªA°È®×¡Bªk°È®×¡BACS®×«D112¤§¦¬¤å¡BÅU°Ý®×«D0¤§¦¬¤å(±qfrm010007.UpdateOtherDatabase©â¥X¨Ó)
Private Function UpdateOtherDB(ByVal pFormName As String, ByVal intSaveMode As Integer, ByVal intModifyKind As Integer, ByVal intCaseKind As Integer, ByVal intChoose As Integer, _
                ByRef mOTB() As String, ByRef mCP() As String, ByVal mCU30 As String, ByVal mChkVal As String, ByVal mSaveControl As String, Optional ByRef IsSaveData As Boolean, _
                Optional ByVal pType As String, Optional ByVal pCaseNo As String) As Boolean
'intSaveMode : 1-·s¼W
'intModifyKind=0¬°·s¼W;=1¬°­×§ï;=2¬°¬d¸ß
'intCaseKind¡A1¬°±M§Q¡A2¬°°Ó¼Ð¡A3¬°ªk°È¡A4¬°ÅU°Ý¡A5¬°±M§Q(ªA)¡A6¬°°Ó¼Ð(ªA)¡A7¬°ªk°È(ªA)¡A8¬°ÅU°Ý(ªA)
'intChoose   0:¦¬¤å   1:¤º³¡¦¬¤å
'mOTB¡G¨Ì¨t²Î§O¶Ç¤J°ò¥»ÀÉ(ªk°ÈLawCase¡BÅU°ÝHireCase¡BªA°ÈServicePractice)
'pType : ¯S®íºÞ¨î
'pCaseNo : ¯S®íºÞ¨î¤§¨Ó·½½s¸¹
'mChkVal¡G¶Ç¤J¨ä¥L¾Þ§@µ²ªG
Dim adoquery As New ADODB.Recordset
Dim stUpdate As String 'Add by Morgan 2008/8/23
'ªk«ß©Ò®×·½¦¬¤å
Dim m_LOS02 As String '®×·½®×¥óÃþ«¬
Dim m_LOS15 As String '®×·½³æ¸¹

'*********¯S®íºÞ¨îªºÅÜ¼Æ*************
   'Modify By Sindy 2025/8/18
   'If pType = "LOS®×·½¦¬¤å" And pCaseNo <> "" Then
   If InStr(pType, "LOS®×·½¦¬¤å") > 0 And pCaseNo <> "" Then
   '2025/8/18 END
        m_LOS02 = Mid(pCaseNo, 1, InStr(pCaseNo, ",") - 1) '®×·½®×¥óÃþ«¬
        m_LOS15 = Mid(pCaseNo, InStr(pCaseNo, ",") + 1, 8) '®×·½³æ¸¹ 'Modify By Sindy 2025/8/18 +, 8)
   End If
   '¥~±M«H¥ó¨R¾P: ¥u¦³·s¼W¦¬¤åªº¥\¯à
'*************************************

If IsSaveData = True Then
    Exit Function
End If
IsSaveData = True

On Error GoTo ErrHand
'Add By Sindy 2022/9/27
If UCase(pFormName) <> UCase("frm090801_New") Then
'2022/9/27 END
   cnnConnection.BeginTrans
End If

'Add by Lydia 2014/10/31 ¶}©ñ¥~±Mµ{§Ç¤H­û¥i¶i¤J±M§Q³B¨t²Î¾Þ§@FMP¾ÈµØ®×¥ó=>¼g¤J¥N²z¤H
If InStr(mChkVal, "¾ÈµØ®×¥ó½T»{") > 0 Then
    mStrSql = "update caseprogress set cp44='Y53374000' where cp09='" & mCP(9) & "' "
Else
    mStrSql = "update caseprogress set cp44='' where cp09='" & mCP(9) & "' "
End If
cnnConnection.Execute mStrSql
'end. 'Add by Lydia 2014/10/31
        
Select Case intCaseKind
             Case ªk°È
                        'edit by nickc 2007/03/27 ¥[¤J©¼©Ò®×¸¹
                        'Modify By Sindy 2011/1/18 +lc43,lc44,lc45,lc46
                        'Add by Morgan 2008/8/5 +LC42 (¨Ö¤J) Ápµ¸¤H½s¸¹
                        'Modified by Lydia 2022/09/13 +LC16, LC17 (¨Ö¤J)¤À©Ò®×¸¹,«È¤á®×¥ó®×¸¹
                        mStrSql = "update lawcase set lc05=" + CNULL(ChgSQL(mOTB(5))) + ",lc06=" + CNULL(ChgSQL(mOTB(6))) + ",lc07=" + CNULL(ChgSQL(mOTB(7))) + ",lc11=" + CNULL(mOTB(11)) + _
                             ",lc15=" + CNULL(mOTB(15)) + ",lc16=" + CNULL(ChgSQL(mOTB(16))) + ",lc17=" + CNULL(ChgSQL(mOTB(17))) + ",lc22=" + CNULL(mOTB(22)) + ",lc23=" + CNULL(mOTB(23)) + _
                             ",lc42=" + CNULL(mOTB(42)) + ",lc43=" + CNULL(mOTB(43)) + ",lc44=" + CNULL(mOTB(44)) + ",lc45=" + CNULL(mOTB(45)) + ",lc46=" + CNULL(mOTB(46))

                        mStrSql = mStrSql + " where lc01=" + CNULL(mOTB(1)) + " and lc02=" + _
                            CNULL(mOTB(2)) + " and lc03=" + CNULL(mOTB(3)) + " and lc04=" + CNULL(mOTB(4))
                        cnnConnection.Execute mStrSql
             Case ÅU°Ý
                        'Modify By Sindy 2011/1/18 +hc24,hc25,hc26,hc27
                        'Add by Morgan 2008/8/5 +HC23 (¨Ö¤J)
                        'Modified by Lydia 2022/09/13 +HC07  (¨Ö¤J) ¤À©Ò®×¸¹
                        mStrSql = "update hirecase set hc05=" + CNULL(mOTB(5)) + ",hc06=" + CNULL(ChgSQL(mOTB(6))) + ",hc07=" + CNULL(ChgSQL(mOTB(7))) + ",hc23=" + CNULL(mOTB(23)) + _
                              ",hc24=" + CNULL(mOTB(24)) + ",hc25=" + CNULL(mOTB(25)) + ",hc26=" + CNULL(mOTB(26)) + ",hc27=" + CNULL(mOTB(27))
                        mStrSql = mStrSql & " where hc01=" + CNULL(mOTB(1)) + " and hc02=" + CNULL(mOTB(2)) + " and hc03=" + CNULL(mOTB(3)) + " and hc04=" + CNULL(mOTB(4))
                        cnnConnection.Execute mStrSql
             Case Else
                        'Modify By Sindy 2011/1/18 +sp65,sp66
                        'Add by Morgan 2008/8/5 +SP78 (¨Ö¤J)
                        'Add By Sindy 2010/3/8 SP30, SP75 (¨Ö¤J)
                        'Modified by Lydia 2022/09/13 SP28, SP29(¨Ö¤J)
                        mStrSql = "update servicepractice set sp05=" + CNULL(ChgSQL(mOTB(5))) + ",sp06=" + CNULL(ChgSQL(mOTB(6))) + ",sp07=" + CNULL(ChgSQL(mOTB(7))) + ",sp08=" + CNULL(mOTB(8)) + ",sp09=" + CNULL(mOTB(9)) + _
                               ",sp58=" + CNULL(mOTB(58)) + ",sp59=" + CNULL(mOTB(59)) + ",sp65=" + CNULL(mOTB(65)) + ",sp66=" + CNULL(mOTB(66)) + ",sp26=" + CNULL(mOTB(26)) + ",sp18=" + CNULL(ChgSQL(mOTB(18))) + _
                               ",sp73=" + CNULL(mOTB(73)) + ",sp74=" + CNULL(mOTB(74)) + ",sp27=" + CNULL(mOTB(27)) + ",sp78=" + CNULL(mOTB(78)) + ",sp28=" + CNULL(ChgSQL(mOTB(28))) + ",sp29=" + CNULL(ChgSQL(mOTB(29))) + _
                               ",sp30=" + CNULL(ChgSQL(mOTB(30))) + ",sp75=" + CNULL(ChgSQL(mOTB(75)))
                        mStrSql = mStrSql & " where sp01=" + CNULL(mOTB(1)) + " and sp02=" + CNULL(mOTB(2)) + " and sp03=" + CNULL(mOTB(3)) + " and sp04=" + CNULL(mOTB(4))
                        cnnConnection.Execute mStrSql
End Select

'Add by Morgan 2008/9/23
If mOTB(1) = "FG" Then
   mCP(48) = Pub_GetHandleDay("FG", "000", mCP(10), , mCP(6))
   stUpdate = ",cp48=" & CNULL(mCP(48), True)
End If

'Modify By Sindy 2012/11/06 +CP150 ¦³¡¹¡¹ªºÀ³¦¬±b´ÚÃ±®Ö±±ºÞ
mStrSql = "update caseprogress set cp05=" + CNULL(mCP(5)) + ",cp06=" + CNULL(mCP(6)) + ",cp07=" + CNULL(mCP(7)) + ",cp10=" + CNULL(mCP(10)) + _
         ",cp11=" + CNULL(mCP(11)) + ",cp13=" + CNULL(mCP(13)) + ",cp14=" + CNULL(mCP(14)) + ",cp16=" + CNULL(mCP(16)) + ",cp17=" + CNULL(mCP(17)) + _
         ",cp18=" + CNULL(mCP(18)) + ",cp19=" + CNULL(mCP(19)) + ",cp32=" + CNULL(mCP(32)) + ",cp33=" + CNULL(mCP(33)) + ",cp34=" + CNULL(mCP(34)) + _
         ",CP64=" + CNULL(ChgSQL(mCP(64))) & stUpdate & ",cp150=" & CNULL(mCP(150)) & " where cp09=" + CNULL(mCP(9))
cnnConnection.Execute mStrSql
mStrSql = "update caseprogress set cp12=(select st15 from staff where st01=" + CNULL(mCP(13)) + ") where cp09=" + CNULL(mCP(9))
cnnConnection.Execute mStrSql
'add by nickc 2008/01/04 ¥[¤J¦^¥N®É¡A©Ó¿ì´Á­­¬°¥»©Ò¦¬¤å¤é(·í¤Ñ¤£ºâ)¤§²Ä¤G­Ó¤u§@¤Ñ
If mCP(10) = "720" Then
     mStrSql = "update caseprogress set cp48=" + CNULL(CompWorkDay(3, mCP(5), 0)) + " where cp09=" + CNULL(mCP(9))
     cnnConnection.Execute mStrSql
End If
        'Add By nickc 2007/08/21
        '­Y¬°±µ¬¢°O¿ý³æ(Âd¥x¦¬¤å)
        'Modify by Morgan 2007/10/26 ¶O¥Î¥i§ï®É¤~°µ¡A§_«h¤w¦¬´Ú¸ê®Æ·|³QÁÙ­ì
        If intChoose = 0 And mCP(60) = "" Then  'mCP(60) = "" => txtOther(12).Enabled = True
            '¥¼¦¬ª÷ÃB = ¶O¥Î
            mStrSql = "update caseprogress set cp79=cp16 where cp09=" + CNULL(mCP(9))
            cnnConnection.Execute mStrSql
        End If
        
'Added by Lydia 2022/11/29 «D¤º³¡¦¬¤å¨Ã¥B¦³¶O¥Î¡A¥ý²Î¤@³]©wCP20=Null ;
If intChoose = 0 And Val(mCP(16)) > 0 Then
    stUpdate = ""
    If mOTB(1) = "FG" Then
       stUpdate = PUB_GetCP20(mOTB(1), mCP(10))
    End If
    If stUpdate = "" Then
       mStrSql = "update caseprogress set cp20=null where cp09=" + CNULL(mCP(9))
       cnnConnection.Execute mStrSql
    End If
End If
'end 2022/11/29
'Add By Cheng 2002/05/10
'­Y¬°¤º³¡¦¬¤å§@·~®É, ®×¥ó¶i«×ÀÉªº¬O§_¦V«È¤á¦¬´Ú³]©w¬°"N"
If intChoose = 1 Then
   mStrSql = "Update CaseProgress Set CP20='N' Where cp09=" + CNULL(mCP(9))
   cnnConnection.Execute mStrSql
End If

'Modify By Sindy 2023/11/3 mark:¦¬¤å¤w¤£»Ý°õ¦æ¦¹¬qµ{¦¡,¦]±µ¬¢³æ¤w·|¦^¼g¶l»¼°Ï¸¹
'Select Case intCaseKind
'     Case ªk°È
'        mStrSql = "update customer set cu30=" + CNULL(mCU30) + " where cu01=" + CNULL(Mid(mOTB(11), 1, 8)) + " and cu02=" + CNULL(Mid(mOTB(11), 9, 1))
'     Case ÅU°Ý
'        mStrSql = "update customer set cu30=" + CNULL(mCU30) + " where cu01=" + CNULL(Mid(mOTB(5), 1, 8)) + " and cu02=" + CNULL(Mid(mOTB(5), 9, 1))
'     Case Else 'ªA°È
'        mStrSql = "update customer set cu30=" + CNULL(mCU30) + " where cu01=" + CNULL(Mid(mOTB(8), 1, 8)) + " and cu02=" + CNULL(Mid(mOTB(8), 9, 1))
'End Select
'cnnConnection.Execute mStrSql
'2023/11/3 END

adoquery.CursorLocation = adUseClient
adoquery.Open "select np01 from nextprogress where np02 = '" & mOTB(1) & "' and np03 = '" & mOTB(2) & "' and np04 = '" & mOTB(3) & "' and np05 = '" & mOTB(4) & "' and np06 is null and np07 = '" & mCP(10) & "'", cnnConnection, adOpenStatic, adLockReadOnly
'Modify By Cheng 2002/05/10
'­Y¦b¤U¤@µ{§ÇÀÉ¥u§ì¨ì¤@µ§¸ê®Æ®É, ¤~­n§ì¤U¤@µ{§ÇÀÉªºÁ`¦¬¤å¸¹§ó·s®×¥ó¶i«×ÀÉªº¬ÛÃöÁ`¦¬¤å¸¹
If adoquery.RecordCount <> 0 And adoquery.RecordCount = 1 Then
   If IsNull(adoquery.Fields(0).Value) = False Then
      '2011/6/17 add by sonia ²§Ä³µªÅG¡Bµû©wµªÅG¡B¼o¤îµªÅG­n¤@¨Ã§ó·s¹ï³y¸ê®Æ
      If (mCP(10) = "602" Or mCP(10) = "604" Or mCP(10) = "606") Then
         cnnConnection.Execute "update caseprogress a set (cp43,cp36,cp37,cp38,cp39,cp40,cp41,cp42,cp80) = (select b.cp09,b.cp36,b.cp37,b.cp38,b.cp39,b.cp40,b.cp41,b.cp42,b.cp80 from caseprogress b where b.cp09='" & adoquery.Fields(0).Value & "') where cp09 = '" & mCP(9) & "'", intJ
      Else
      '2011/6/17 end
         cnnConnection.Execute "update caseprogress set cp43 = '" & adoquery.Fields(0).Value & "' where cp09 = '" & mCP(9) & "'"
      End If  '2011/6/17 add by sonia
   End If
End If
adoquery.Close
'add by nickc 2008/05/02 Àx¦s¹w©w¦¬´Ú¤é
'Remove by Lydia 2018/08/22 (À³¦¬±b´ÚºÞ±±)¨ú®ø¹w©w¦¬´Ú¤é,§ï¦¨¥I´Ú¶g´Á
'Dim rtCnt As Integer
''Modify by Morgan 2010/12/9
''If txtOther(28) <> "" Then
''    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtOther(28)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),'" & strUserNum & "'," & DBDATE(txtOther(28)) & " ", rtCnt
'If txtOther(28) <> "" And txtOther(28) <> txtOther(28).Tag Then
'    cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),nvl(max(rd03),0)+1,'" & strUserNum & "'," & DBDATE(txtOther(28)) & " from receivablesday where rd01='" & mCP(9) & "' and rd02=to_number(to_char(sysdate,'YYYYMMDD')) group by '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')) ", rtCnt
''end 2010/12/9
'    If rtCnt = 0 Then
'        cnnConnection.Execute "insert into receivablesday (rd01,rd02,rd03,rd04,rd05) select '" & mCP(9) & "',to_number(to_char(sysdate,'YYYYMMDD')),1,'" & strUserNum & "'," & DBDATE(txtOther(28)) & " from dual "
'    End If
'End If
'end 2018/08/22

If m_LOS15 <> "" Then PUB_UpdateTTFee m_LOS15 'Added by Morgan 2022/4/14

If mSaveControl <> "" Then
    'Modified by Lydia 2022/09/29 ¶Ç¤J¨t²Î§O,°ê®a,®×¥ó©Ê½è
    'Call PUB_SaveByControl(mCP(9), mSaveControl)
    Call PUB_SaveByControl(mCP(9), mSaveControl, mOTB(1), IIf(intCaseKind = ªk°È, mOTB(15), IIf(intCaseKind = ÅU°Ý, "000", mOTB(9))), mCP(10))
End If
  
'Add By Sindy 2022/9/27
If UCase(pFormName) <> UCase("frm090801_New") Then
'2022/9/27 END
   cnnConnection.CommitTrans
End If
UpdateOtherDB = True
Set adoquery = Nothing

Exit Function
ErrHand:
'Add By Sindy 2022/9/27
If UCase(pFormName) <> UCase("frm090801_New") Then
'2022/9/27 END
   cnnConnection.RollbackTrans
End If
ShowMsg MsgText(9004) & IIf(Err.Number <> 0, vbCrLf & vbCrLf & Err.Description, "") 'Modify By Sindy 2022/10/14 + IIf

IsSaveData = False
End Function

'Add By Sindy 2022/9/26 ±µ¬¢³æ¹q¤l¦¬¤å
Public Function PUB_AutoRecvCRLMain(strSys As String, strCRL01 As String) As Boolean
Dim intCaseKind As Integer
Dim bolHC0 As Boolean, bolACS112 As Boolean
Dim rsTmp As New ADODB.Recordset
Dim rsAD As New ADODB.Recordset 'Add By Sindy 2024/5/20
Dim intA As Integer, ii As Integer
Dim strCRL06 As String, strCRL07 As String, strCRL08 As String, strCRL09 As String, strCRL10 As String
Dim strCRL15 As String
Dim strCRL() As String
Dim strCRL01_102 As String, strCRL07_102 As String, strCRL15_102 As String
Dim strCRL01_101 As String, strUpdCase As String
Dim strPA26_102 As String, strPA27_102 As String, strPA28_102 As String
Dim strPA29_102 As String, strPA30_102 As String
Dim strNation As String
Dim strUpdCase_101 As String, strUpdCase_102 As String
Dim strCRA05 As String, strFilePath As String, pSavePath As String
Dim strCRL55 As String, arrTmp As Variant, bolSpecRecv As Boolean
Dim Str01 As String, Str02 As String, Str03 As String, Str04 As String
Dim strCRC02 As String 'Add By Sindy 2024/5/16

On Error GoTo ErrHand
   
   '±µ¬¢³æ¹q¤lÀÉ§ó¦W-·s®×(·|¥h±¼¤¤¤å)
   Call PUB_UpdCRLFileName(strCRL01)
   
   ClsPDGetSystemKind strSys, intCaseKind
   
   Select Case intCaseKind
      Case ±M§Q 'frm010005
         'ÀË¬d¬O§_¦³®×¥ó©Ê½è105,125­n°µ¯S§O³B²z
         'Modify By Sindy 2023/4/19 +113¡B114¡B122
         strSql = "select * from ConsultRecCMP" & _
                  " where CRC01='" & strCRL01 & "' and CRC03 in('105','125','113','114','122')"
         intA = 1
         Set rsTmp = ClsLawReadRstMsg(intA, strSql)
         If intA = 1 Then
            'Add By Sindy 2023/5/16 ³o´X­Ó¯S®í®×¥ó©Ê½è¥i¥H­n¦¬·s®×,¨S¦³¥À¸¹(¥L©Ò¿ì²zªº)
            strSql = "select * from consultrecordlist,consultreccmp where crc01='" & strCRL01 & "' and crc03 in('105','125','113','114','122')" & _
                     " and crc01=crl01 and crl06='Y'" & _
                     " and not exists(select * from consultreccmp where crc01='" & strCRL01 & "' and crc03 in('103','101'))"
            intA = 1
            Set rsTmp = ClsLawReadRstMsg(intA, strSql)
            If intA = 1 Then
               '¦¬¤å
               If PUB_AutoRecvCRL_P(strCRL01) = False Then
                  GoTo ErrHand
               End If
            Else
            '2023/5/16 END
               'Åª¨ú±µ¬¢³æ¥DÀÉ
               strSql = "select CRL01,CRL06,CRL07,CRL08,CRL09,CRL10,CRL15 from CONSULTRECORDLIST" & _
                        " where CRL01='" & strCRL01 & "'"
               intA = 1
               Set rsTmp = ClsLawReadRstMsg(intA, strSql)
               If intA = 1 Then
                  strCRL06 = "" & rsTmp.Fields("CRL06") '·sÂÂ®×
                  strCRL07 = "" & rsTmp.Fields("CRL07")
                  strCRL08 = "" & rsTmp.Fields("CRL08")
                  strCRL09 = "" & rsTmp.Fields("CRL09")
                  strCRL10 = "" & rsTmp.Fields("CRL10")
                  strCRL15 = "" & rsTmp.Fields("CRL15")
               End If
               'ÀË¬d¬O§_¦³®×¥ó©Ê½è103¥ý³B²z
               'Modify By Sindy 2023/4/19 + 101
               strSql = "select * from ConsultRecCMP" & _
                        " where CRC01='" & strCRL01 & "' and CRC03 in('103','101')"
               intA = 1
               Set rsTmp = ClsLawReadRstMsg(intA, strSql)
               If intA = 1 Then
                  If strCRL06 <> "Y" Then
                     MsgBox "±M§Q³]­p¥Ó½ÐÀ³¦¬·s®×¡A½Ð¬¢¹q¸£¤¤¤ß¨ó§U½T»{¡I" & vbCrLf & strSql, vbExclamation, "PUB_AutoRecvCRLMain"
                     GoTo ErrHand
                  End If
                  '¦¬¤å
                  'Modify By Sindy 2023/4/19 +113¡B114¡B122
                  If PUB_AutoRecvCRL_P(strCRL01, " and CRC03 not in('105','125','113','114','122')", strCRL07, strCRL08, strCRL09, strCRL10) = False Then
                     GoTo ErrHand
                  End If
               End If
               '¦AÅª¨ú®×¥ó©Ê½è105,125¨M©w¶]´X¦¸·s®×¦¬¤å(¤ä¸¹)
               'Modify By Sindy 2023/4/19 +113¡B114¡B122
               strSql = "select CRC01,CRC02,CRC03,CPM01,CPM02,CPM03,CPM04 from ConsultRecCMP,casepropertymap" & _
                        " where CRC01='" & strCRL01 & "' and CRC03 in('105','125','113','114','122')" & _
                        " and '" & strCRL07 & "'=CPM01(+) and CRC03=CPM02(+)"
               intA = 1
               Set rsTmp = ClsLawReadRstMsg(intA, strSql)
               If intA = 1 Then
                  strCRC02 = "" & rsTmp.Fields("CRC02") 'Add By Sindy 2024/5/16
                  If strCRL07 = "" Or strCRL08 = "" Then
                     MsgBox "±M§Q" & IIf(strCRL15 = "000", "" & rsTmp.Fields("cpm03"), "" & rsTmp.Fields("cpm04")) & "¨S¦³¥À®×®×¸¹¡A½Ð¬¢¹q¸£¤¤¤ß¨ó§U½T»{¡I" & vbCrLf & strSql, vbExclamation, "PUB_AutoRecvCRLMain"
                     GoTo ErrHand
                  End If
                  
                  rsTmp.MoveFirst
                  Do While Not rsTmp.EOF
                     '¸Ó¥À®×®×¸¹ªº-1,-2¡K
                     strSql = "SELECT a.maxPA03 FROM Patent, " & _
                              "(SELECT max(PA03) as maxPA03 FROM Patent WHERE PA01='" & strCRL07 & "' AND PA02='" & strCRL08 & "') a " & _
                              "WHERE PA01='" & strCRL07 & "' AND PA02='" & strCRL08 & "' AND PA03='" & strCRL09 & "' AND PA04='" & strCRL10 & "' "
                     intA = 1
                     Set rsAD = ClsLawReadRstMsg(intA, strSql)
                     If intA = 1 Then
                        '¤ä¸¹¬°0-9,A-Z
                        If Val(rsAD.Fields(0)) >= 0 And Val(rsAD.Fields(0)) <= 8 Then
                           strCRL09 = Val(rsAD.Fields(0)) + 1
                        ElseIf Val(rsAD.Fields(0)) = 9 Then
                           strCRL09 = "A"
                        Else
                           strCRL09 = Chr(Asc(rsAD.Fields(0)) + 1)
                        End If
                        If strCRL10 = "" Then strCRL10 = "00"
                     Else
                        MsgBox "¨t²Î¦b¨ú±o¸Ó¥À®×®×¸¹(" & strCRL07 & "-" & strCRL08 & ")ªº¤ä¸¹(-1,-2¡K)®É¦³°ÝÃD¡A½Ð¬¢¹q¸£¤¤¤ß¨ó§U½T»{¡I" & vbCrLf & strSql, vbExclamation, "PUB_AutoRecvCRLMain"
                        GoTo ErrHand
                     End If
                     '¦¬¤å: ³]¬°·s®×
                     'Modify By Sindy 2023/4/19 +113¡B114¡B122
                     If PUB_AutoRecvCRL_P(strCRL01, " and (CRC03 not in('103','105','125','113','114','122') or CRC02=" & strCRC02 & ") order by decode(CRC03,'103','105','125','113','114','122',1,2) asc, CRC02 asc", strCRL07, strCRL08, strCRL09, strCRL10, True) = False Then
                        GoTo ErrHand
                     End If
                     rsTmp.MoveNext
                  Loop
               End If
            End If
         Else
            '¦¬¤å
            If PUB_AutoRecvCRL_P(strCRL01) = False Then
               GoTo ErrHand
            End If
         End If

      Case °Ó¼Ð 'frm010004
         'Modify By Sindy 2023/6/19
         'ÀË¬d¬O§_ÂÂ®×TF»â¤g©µ¦ù(104),­n°µ¯S§O³B²z
         strSql = "select * from CONSULTRECORDLIST,ConsultRecCMP" & _
                  " where CRL01='" & strCRL01 & "' and CRL01=CRC01 and CRL06 is null and CRL07='TF' and CRC03 in('104')"
         intA = 1
         Set rsTmp = ClsLawReadRstMsg(intA, strSql)
         If intA = 1 Then
            strCRL07 = "" & rsTmp.Fields("CRL07")
            strCRL08 = "" & rsTmp.Fields("CRL08")
            strCRL09 = "" & rsTmp.Fields("CRL09")
            strCRL10 = "" & rsTmp.Fields("CRL10")
            '¨ú±o¸Ó¥À®×®×¸¹ªºTM02³Ì¤j­È
            strSql = "SELECT a.maxTM02 FROM Trademark, " & _
                     "(SELECT max(TM02) as maxTM02 FROM Trademark WHERE TM01='" & strCRL07 & "' AND substr(TM02,1,5)='" & Left(strCRL08, 5) & "') a " & _
                     "WHERE TM01='" & strCRL07 & "' AND substr(TM02,1,5)='" & Left(strCRL08, 5) & "' AND TM03='" & strCRL09 & "' AND TM04='" & strCRL10 & "' "
            intA = 1
            Set rsTmp = ClsLawReadRstMsg(intA, strSql)
            If intA = 1 Then
               '»â¤g©µ¦ù¬°¬y¤ô¸¹²Ä6½X:1-9
               strCRL08 = Left(rsTmp.Fields(0), 5) & CStr(Val(Right(rsTmp.Fields(0), 1)) + 1)
               If strCRL09 = "" Then strCRL09 = "0"
               If strCRL10 = "" Then strCRL10 = "00"
            Else
               MsgBox "¨t²Î¦b¨ú±o¸Ó¥À®×®×¸¹(" & strCRL07 & "-" & strCRL08 & ")ªº»â¤g©µ¦ù®É¦³°ÝÃD¡A½Ð¬¢¹q¸£¤¤¤ß¨ó§U½T»{¡I" & vbCrLf & strSql, vbExclamation, "PUB_AutoRecvCRLMain"
               GoTo ErrHand
            End If
            '¦¬¤å: ³]¬°·s®×
            If PUB_AutoRecvCRL_T(strCRL01, strCRL07, strCRL08, strCRL09, strCRL10, True) = False Then
               GoTo ErrHand
            End If
         Else
         '2023/6/19 END
            '¦¬¤å
            If PUB_AutoRecvCRL_T(strCRL01) = False Then
               GoTo ErrHand
            End If
         End If
         
      Case Else
         If intCaseKind = ÅU°Ý Then
            strSql = "select CRC03 from ConsultRecCMP" & _
                     " where CRC01='" & strCRL01 & "' and CRC03='" & ÅU°Ý¸u¥ô & "'"
            intA = 1
            Set rsTmp = ClsLawReadRstMsg(intA, strSql)
            If intA = 1 Then
               bolHC0 = True
            End If
         ElseIf strSys = "ACS" Then
            strSql = "select CRC03 from ConsultRecCMP" & _
                     " where CRC01='" & strCRL01 & "' and CRC03='" & 112 & "'"
            intA = 1
            Set rsTmp = ClsLawReadRstMsg(intA, strSql)
            If intA = 1 Then
               bolACS112 = True
            End If
         End If
         'ÅU°Ý®×¤§0ÅU°Ý¸u¥ô¦¬¤å
         If intCaseKind = ÅU°Ý And bolHC0 = True Then 'frm010006
            If PUB_AutoRecvCRL_HireCase(strCRL01) = False Then
               GoTo ErrHand
            End If
            
         'ACS®×¤§112´¼°]ÅU°Ý¦¬¤å
         ElseIf strSrvDate(1) >= ACS_PFrateStart And _
            strSys = "ACS" And bolACS112 = True Then 'frm010006_1
            If PUB_AutoRecvCRL_Case(strCRL01) = False Then
               GoTo ErrHand
            End If
            
         '¥]§tªA°È®×¡Bªk°È®×¡BACS®×«D112¤§¦¬¤å¡BÅU°Ý®×«D0¤§¦¬¤å : frm010007
         Else
            'Modify By Sindy 2023/8/16 ­n°µ¯S§O³B²z
            'ÀË¬d¬O§_¬ÛÃö®×¸¹¬°ACS¥B´¿¦¬¤å´¼°]ÅU°Ý112¡A±µ¬¢³æ¬O¦¬¤åS©ÎTS©ÎTT¡A®×¥ó©Ê½è¦³¬d¦W001©Î°Ó¼ÐºÊ±±738¡A¥B¤Ä¿ïÂÂ®×¡A
            '¦Ó¦sÀÉ®É¥H·s®×¤è¦¡¦s¤J¨t²Î¡A¥»©Ò®×¸¹¬°¸ÓÂÂ®×ªº-1,-2¡K¡A¨Ò¦p¡G¦¬¤åS-006615¦sÀÉ§ï¬° S-006615-1¡C
            strSql = "select * from CONSULTRECORDLIST,ConsultRecCMP" & _
                     " where CRL01='" & strCRL01 & "' and CRL01=CRC01 and CRL06 is null and CRL07 in('S','TS','TT') and CRC03 in('001','738')" & _
                     " and CRL55 is not null and instr(CRL55,'ACS')>0"
            intA = 1
            Set rsTmp = ClsLawReadRstMsg(intA, strSql)
            If intA = 1 Then
               strCRL07 = "" & rsTmp.Fields("CRL07")
               strCRL08 = "" & rsTmp.Fields("CRL08")
               strCRL09 = "" & rsTmp.Fields("CRL09")
               strCRL10 = "" & rsTmp.Fields("CRL10")
               strCRL55 = "" & rsTmp.Fields("CRL55") '¬ÛÃö®×¸¹
               arrTmp = Split(strCRL55, ",")
               bolSpecRecv = False '¹w³]­È
               For ii = LBound(arrTmp) To UBound(arrTmp)
                  Str01 = SystemNumber(CStr(arrTmp(ii)), 1)
                  Str02 = SystemNumber(CStr(arrTmp(ii)), 2)
                  Str03 = SystemNumber(CStr(arrTmp(ii)), 3)
                  Str04 = SystemNumber(CStr(arrTmp(ii)), 4)
                  'ÀË¬d¬O§_¬ÛÃö®×¸¹ACS´¿¦¬¤å´¼°]ÅU°Ý112
                  If Str01 = "ACS" Then
                     strSql = "SELECT cp09 FROM Caseprogress " & _
                              "WHERE cp01='" & Str01 & "' AND cp02='" & Str02 & "' AND cp03='" & Str03 & "' AND cp04='" & Str04 & "'" & _
                              " AND cp10='112'"
                     intA = 1
                     Set rsTmp = ClsLawReadRstMsg(intA, strSql)
                     If intA = 1 Then
                        bolSpecRecv = True '­n°µ¯S§O³B²z
                        Exit For
                     End If
                  End If
               Next ii
               If bolSpecRecv = True Then
                  '¸Ó¥À®×®×¸¹ªº-1,-2¡K
                  strSql = "SELECT a.maxSP03 FROM servicepractice, " & _
                           "(SELECT max(SP03) as maxSP03 FROM servicepractice WHERE SP01='" & strCRL07 & "' AND SP02='" & strCRL08 & "') a " & _
                           "WHERE SP01='" & strCRL07 & "' AND SP02='" & strCRL08 & "' AND SP03='" & strCRL09 & "' AND SP04='" & strCRL10 & "' "
                  intA = 1
                  Set rsTmp = ClsLawReadRstMsg(intA, strSql)
                  If intA = 1 Then
                     '¤ä¸¹¬°0-9,A-Z
                     If Val(rsTmp.Fields(0)) >= 0 And Val(rsTmp.Fields(0)) <= 8 Then
                        strCRL09 = Val(rsTmp.Fields(0)) + 1
                     ElseIf Val(rsTmp.Fields(0)) = 9 Then
                        strCRL09 = "A"
                     Else
                        strCRL09 = Chr(Asc(rsTmp.Fields(0)) + 1)
                     End If
                     If strCRL10 = "" Then strCRL10 = "00"
                  Else
                     MsgBox "¨t²Î¦b¨ú±o¸Ó¥À®×®×¸¹(" & strCRL07 & "-" & strCRL08 & ")ªº¤ä¸¹(-1,-2¡K)®É¦³°ÝÃD¡A½Ð¬¢¹q¸£¤¤¤ß¨ó§U½T»{¡I" & vbCrLf & strSql, vbExclamation, "PUB_AutoRecvCRLMain"
                     GoTo ErrHand
                  End If
                  '¦¬·s®×
                  If PUB_AutoRecvCRL_Other(strCRL01, strCRL07, strCRL08, strCRL09, strCRL10, True) = False Then
                     GoTo ErrHand
                  End If
               Else
                  If PUB_AutoRecvCRL_Other(strCRL01) = False Then
                     GoTo ErrHand
                  End If
               End If
            Else
            '2023/8/16 END
               If PUB_AutoRecvCRL_Other(strCRL01) = False Then
                  GoTo ErrHand
               End If
            End If
         End If
   End Select
   
   '±N·s®×¦³³]©wÃö³sªí³æ½s¸¹ªÌ,¶ñ¤J¬Û¦P®×¸¹
   If intCaseKind = ±M§Q Then
      '±µ¬¢³æ¥DÀÉ
      ReDim Preserve strCRL(TF_CRL) As String
      strCRL(1) = strCRL01
      If ClsPDReadCRLDatabase(strCRL) = False Then
         GoTo ErrHand
      End If
      
      If strCRL(6) = "Y" And strCRL(65) <> "" And (strCRL(7) = "P" Or strCRL(7) = "CFP") Then
         'ÀË¬d¬O§_¾ã­Ó¸s²Õ¤w¥þ¦¬¤å¤F(±Æ°£ÀË¬d¦¹µ§±µ¬¢³æ),¦A¨t²Î¶ñ¤J¸ê®Æ(¦b¹q¤l¦¬¤åªº³Ì«á¤@­Ó®×¸¹¤~°µ³B²z)
         strTmp1(0) = "select CRL01 from ConsultRecordList,ConsultRecCMP where CRL65='" & strCRL(65) & "'" & _
                     " and CRL01=CRC01 and CRC08 is null" ' and CRL55 is null
         intA = 1
         Set rsTmp = ClsLawReadRstMsg(intA, strTmp1(0))
         If intA = 0 Then '¤w¥þ¦¬¤å¤F,³Ì«á¤@µ§§ó·s¤U¦C¸ê°T
            'ÀË¬d¬O§_¦³¦P®É¦¬¤@®×¨â½Ð(¦P¥Ó½Ð°ê®a,¦¬¤Fµo©ú¥Ó½Ð©M·s«¬¥Ó½Ð) ' and CRL55 is null
            '¥Ó½Ð¤H­n¬Û¦P
            strTmp1(0) = "select CRL01,CRL07,CRL08,CRL09,CRL10,CRL15,PA26,PA27,PA28,PA29,PA30 from ConsultRecordList,ConsultRecCMP,caseprogress,patent" & _
                        " where CRL65='" & strCRL(65) & "'" & _
                        " and CRL01=CRC01 and CRC03='102' and (CRL07 = 'P' Or CRL07 = 'CFP') and CRC08 is not null" & _
                        " and CRC08=CP09(+) and CP01=PA01(+) and CP02=PA02(+) and CP03=PA03(+) and CP04=PA04(+)"
            intA = 1
            Set rsTmp = ClsLawReadRstMsg(intA, strTmp1(0))
            If intA = 1 Then
               rsTmp.MoveFirst
               Do While Not rsTmp.EOF
                  strCRL01_102 = rsTmp.Fields("CRL01")
                  strCRL07_102 = rsTmp.Fields("CRL07")
                  strCRL15_102 = rsTmp.Fields("CRL15")
                  strPA26_102 = "" & rsTmp.Fields("PA26")
                  strPA27_102 = "" & rsTmp.Fields("PA27")
                  strPA28_102 = "" & rsTmp.Fields("PA28")
                  strPA29_102 = "" & rsTmp.Fields("PA29")
                  strPA30_102 = "" & rsTmp.Fields("PA30")
                  '¥xÆW,¤j³°,¼w°ê¤~¦³¤@®×¨â½Ð
                  If strCRL15_102 = "000" Or strCRL15_102 = "020" Or strCRL15_102 = "231" Then
                     strTmp1(0) = "select CRL01,CRL07,CRL08,CRL09,CRL10,CRL15 from ConsultRecordList,ConsultRecCMP,caseprogress,patent" & _
                                 " where CRL65='" & strCRL(65) & "'" & _
                                 " and CRL01=CRC01 and CRC03='101' and CRL07='" & strCRL07_102 & "' and CRL15='" & strCRL15_102 & "' and CRC08 is not null" & _
                                 " and CRC08=CP09(+) and CP01=PA01(+) and CP02=PA02(+) and CP03=PA03(+) and CP04=PA04(+)"
                     strTmp1(0) = strTmp1(0) & _
                                    " and PA26='" & strPA26_102 & "'"
                     If strPA27_102 <> "" Then
                        strTmp1(0) = strTmp1(0) & _
                                    " and PA27='" & strPA27_102 & "'"
                     Else
                        strTmp1(0) = strTmp1(0) & _
                                    " and PA27 is null"
                     End If
                     If strPA28_102 <> "" Then
                        strTmp1(0) = strTmp1(0) & _
                                    " and PA28='" & strPA28_102 & "'"
                     Else
                        strTmp1(0) = strTmp1(0) & _
                                    " and PA28 is null"
                     End If
                     If strPA29_102 <> "" Then
                        strTmp1(0) = strTmp1(0) & _
                                    " and PA29='" & strPA29_102 & "'"
                     Else
                        strTmp1(0) = strTmp1(0) & _
                                    " and PA29 is null"
                     End If
                     If strPA30_102 <> "" Then
                        strTmp1(0) = strTmp1(0) & _
                                    " and PA30='" & strPA30_102 & "'"
                     Else
                        strTmp1(0) = strTmp1(0) & _
                                    " and PA30 is null"
                     End If
                     intA = 1
                     Set rsAD = ClsLawReadRstMsg(intA, strTmp1(0))
                     If intA = 1 Then
                        strCRL01_101 = rsAD.Fields("CRL01")
                        strUpdCase = rsAD.Fields("CRL07") & "-" & rsAD.Fields("CRL08") & IIf(rsAD.Fields("CRL09") & rsAD.Fields("CRL10") = "000", "", "-" & rsAD.Fields("CRL09") & "-" & rsAD.Fields("CRL10"))
                        strSql = "update ConsultRecordList set CRL67='" & strUpdCase & "' where CRL01='" & strCRL01_102 & "'"
                        cnnConnection.Execute strSql, intA
                        strSql = "update ConsultRecordList set CRL67='" & strUpdCase & "' where CRL01='" & strCRL01_101 & "'"
                        cnnConnection.Execute strSql, intA
                     End If
                  End If
                  rsTmp.MoveNext
               Loop
            End If
            '*************************
            '§ó·s¬°¬Û¦P®×¸¹
            '*************************
            'Modify By Sindy 2023/1/30 Mark, ex:1120001480
'            '±Æ°£¦¹¸s²Õ¶È¬°¤@®×¨â½Ðªºª¬ªp
'            strtmp1(0) = "select CRL01,CRL07,CRL15 from ConsultRecordList where CRL65='" & strCRL(65) & "' and CRL55 is null and (CRL07 = 'P' Or CRL07 = 'CFP')" & _
'                        " and CRL67 is null"
'            inta = 1
'            Set rsTmp = ClsLawReadRstMsg(inta, strtmp1(0))
'            If inta = 1 Then
               strNation = "000" '¥xÆW
               strUpdCase = ""
            '***************************
            '¦³¥xÆW®×®É,
            '***************************
RunSameChk:
               'P¥xÆW®×101µo©ú,®×¸¹
               strUpdCase_101 = ""
               strTmp1(0) = "select CRL07,CRL08,CRL09,CRL10 from ConsultRecordList,ConsultRecCMP where CRL65='" & strCRL(65) & "'" & _
                           " and CRL01=CRC01 and CRL07='P' and CRC08 is not null and CRL15='" & strNation & "' and CRC03='101'" & _
                           " order by CRC03 asc"
               intA = 1
               Set rsTmp = ClsLawReadRstMsg(intA, strTmp1(0))
               If intA = 1 Then
                  strUpdCase_101 = rsTmp.Fields("CRL07") & "-" & rsTmp.Fields("CRL08") & IIf(rsTmp.Fields("CRL09") & rsTmp.Fields("CRL10") = "000", "", "-" & rsTmp.Fields("CRL09") & "-" & rsTmp.Fields("CRL10"))
               End If
               'P¥xÆW®×102·s«¬,®×¸¹
               strUpdCase_102 = ""
               strTmp1(0) = "select CRL07,CRL08,CRL09,CRL10 from ConsultRecordList,ConsultRecCMP where CRL65='" & strCRL(65) & "'" & _
                           " and CRL01=CRC01 and CRL07='P' and CRC08 is not null and CRL15='" & strNation & "' and CRC03='102'" & _
                           " order by CRC03 asc"
               intA = 1
               Set rsTmp = ClsLawReadRstMsg(intA, strTmp1(0))
               If intA = 1 Then
                  strUpdCase_102 = rsTmp.Fields("CRL07") & "-" & rsTmp.Fields("CRL08") & IIf(rsTmp.Fields("CRL09") & rsTmp.Fields("CRL10") = "000", "", "-" & rsTmp.Fields("CRL09") & "-" & rsTmp.Fields("CRL10"))
                  '¥xÆW·s«¬=¥xÆWµo©ú
                  If strUpdCase_101 <> "" Then
                     strUpdCase = strUpdCase_101 '¦³¥xÆWµo©ú
                     strSql = "update ConsultRecordList set CRL55='" & strUpdCase_101 & "',CRL56='Y' where CRL65='" & strCRL(65) & "' and CRL07 = 'P'" & _
                              " and CRL15='" & strNation & "' and CRL87='2' and CRL55 is null"
                     cnnConnection.Execute strSql, intA
                  Else
                     strUpdCase = strUpdCase_102
                  End If
                  '¨ä¥Lªº·s«¬=¥xÆW·s«¬
                  strSql = "update ConsultRecordList set CRL55='" & strUpdCase_102 & "',CRL56='Y' where CRL65='" & strCRL(65) & "' and (CRL07 = 'P' Or CRL07 = 'CFP')" & _
                           " and CRL87='2' and CRL55 is null"
                  cnnConnection.Execute strSql, intA
               End If
               '¨S¦³¥xÆW·s«¬,¦³¥xÆWµo©ú
               If strUpdCase = "" And strUpdCase_101 <> "" Then strUpdCase = strUpdCase_101
               If strUpdCase <> "" Then
                  strSql = "update ConsultRecordList set CRL55='" & strUpdCase & "',CRL56='Y' where CRL65='" & strCRL(65) & "' and (CRL07 = 'P' Or CRL07 = 'CFP')" & _
                           " and CRL55 is null"
                  cnnConnection.Execute strSql, intA
               
            '***************************
            '¥xÆW³£¨S¦³¥Ó½Ð®É,¦Ò¼{¤j³°(¦P¥xÆW³W«h)
            '***************************
               Else
                  If strNation = "000" Then
                     strNation = "020" '¤j³°
                     GoTo RunSameChk
                  End If
               End If
               
            '***************************
            '¨ä¥L
            '***************************
               If strUpdCase = "" Then
                  '¥ý§ì¥xÆW,¤¤°ê,¬ü°ê
                  strTmp1(0) = "select CRL07,CRL08,CRL09,CRL10,1 as sort from ConsultRecordList,ConsultRecCMP" & _
                              " where CRL65='" & strCRL(65) & "' and crl01=crc01" & _
                              " and CRL07='P' and CRC08 is not null and CRL15='000'" & _
                              " union select CRL07,CRL08,CRL09,CRL10,2 as sort from ConsultRecordList,ConsultRecCMP" & _
                              " where CRL65='" & strCRL(65) & "' and crl01=crc01" & _
                              " and CRL07='P' and CRC08 is not null and CRL15='020'" & _
                              " union select CRL07,CRL08,CRL09,CRL10,3 as sort from ConsultRecordList,ConsultRecCMP" & _
                              " where CRL65='" & strCRL(65) & "' and crl01=crc01" & _
                              " and CRL07='CFP' and CRC08 is not null and CRL15='101'" & _
                              " order by sort asc"
                  intA = 1
                  Set rsTmp = ClsLawReadRstMsg(intA, strTmp1(0))
                  If intA = 1 Then
                     strUpdCase = rsTmp.Fields("CRL07") & "-" & rsTmp.Fields("CRL08") & IIf(rsTmp.Fields("CRL09") & rsTmp.Fields("CRL10") = "000", "", "-" & rsTmp.Fields("CRL09") & "-" & rsTmp.Fields("CRL10"))
                  End If
                  If strUpdCase = "" Then
                     '³£¨S¦³,¤~¨Ì²Ä¤@µ§¬°¥À¸¹
                     strTmp1(0) = "select CRL07,CRL08,CRL09,CRL10 from ConsultRecordList,ConsultRecCMP where CRL65='" & strCRL(65) & "'" & _
                                 " and CRL01=CRC01 and CRC08 is not null and (CRL07 = 'P' Or CRL07 = 'CFP')" & _
                                 " order by CRC08 asc"
                     intA = 1
                     Set rsTmp = ClsLawReadRstMsg(intA, strTmp1(0))
                     If intA = 1 Then
                        strUpdCase = rsTmp.Fields("CRL07") & "-" & rsTmp.Fields("CRL08") & IIf(rsTmp.Fields("CRL09") & rsTmp.Fields("CRL10") = "000", "", "-" & rsTmp.Fields("CRL09") & "-" & rsTmp.Fields("CRL10"))
                     End If
                  End If
                  strSql = "update ConsultRecordList set CRL55='" & strUpdCase & "',CRL56='Y' where CRL65='" & strCRL(65) & "' and (CRL07 = 'P' Or CRL07 = 'CFP')" & _
                           " and CRL55 is null"
                  cnnConnection.Execute strSql, intA
               End If
'            End If

            'Add By Sindy 2023/5/31
            strTmp1(0) = "select CRL01 from ConsultRecordList where CRL65='" & strCRL(65) & "'"
            intA = 1
            Set rsTmp = ClsLawReadRstMsg(intA, strTmp1(0))
            If intA = 1 Then
               If rsTmp.RecordCount = 2 Then '³æ¯Âªº¤@°ê¤@®×¨â½Ð®É,¤£­n³]¬Û¦P®× ex:P-131585©MP-131586
                  strSql = "update consultrecordlist set crl55=null,crl56=null" & _
                           " where crl55||crl15 in(select crl55||crl15 from consultrecordlist where crl65='" & strCRL(65) & "'" & _
                           " and crl67 is not null and crl67=crl55 and instr(crl55,crl08)>0)"
                  Pub_SeekTbLog strSql
                  cnnConnection.Execute strSql, intA
               End If
            End If
            '2023/5/31 END
         End If
      End If
   
   ElseIf intCaseKind = °Ó¼Ð Then
      'Add By Sindy 2023/3/28
      If UCase(pub_DbTerminalName) = UCase(¥¿¦¡¸ê®Æ®w¹q¸£¦WºÙ) Then
         '¥x¤@¨t²Î¥D­n¥H"®×¥ó"¬°¥D¡A¹Å¶²»Ý­n¥H"«È¤á"¨¤«×¤U¸üÀÉ®×¡A¬GÀÉ®×¶×¤Jªº³W«h²§°Ê¦p¤U¡G
         '1.³z¹L¹q¤l¦¬¤å¶×¤JªºÀÉ®×¡]°e¥ó¤å®Ñ©Î¤º³¡¤å¥ó¡^¡A°£­ì¦³½Æ»s¦Ü¨÷©v°Ï¤§¥~¡A¥t«þ¨©¤A¥÷¦Üsale1.
         '2.Sale1³B¡A¨Ì®×¥ó"«È¤á½s¸¹"¡]«e¢·½X¡^·s¼W¸ê®Æ§¨¦b"°Ó¼Ð«È¤á±M°Ï"¤º¡A¨ÃµL±ø¥óÂÐ»\¦PÀÉ¦WÀÉ®×¡C
         '4.½Ð¨Ìªþ¥óÂê©w¯S©w®×¥ó©Ê½è¤~¶i¦æ¤W­z«þ¨©§@·~¡C
         pSavePath = App.path & "\" & strUserNum
         If Dir(pSavePath, vbDirectory) = "" Then MkDir pSavePath
         '§ì±µ¬¢³æ¥Ó½Ð¤H1½s¸¹, ÀË¬d¬O§_²Å¦X­n«þ¨©¹q¤lÀÉªº¯S©w®×¥ó©Ê½è
         strTmp1(0) = "select CRL01,CRA05 from ConsultRecordList,ConsultRecCMP,ConsultRecApp where CRL01='" & strCRL01 & "'" & _
                     " and CRL01=CRC01 and CRC08 is not null and CRL01=CRA01 and CRA02=1" & _
                     " and instr('" & strT000Sale1CPMList & "',CRC03)>0 and CRL07='T' and CRL15='000'"
         intA = 1
         Set rsTmp = ClsLawReadRstMsg(intA, strTmp1(0))
         If intA = 1 Then
            strCRA05 = rsTmp.Fields("CRA05") '«È¤á½s¸¹
            strTmp1(0) = "select * from casepaperpdf where CPP11='" & strCRL01 & "'"
            intA = 1
            Set rsTmp = ClsLawReadRstMsg(intA, strTmp1(0))
            If intA = 1 Then
               strFilePath = "/" & Mid(strTApp1CasePath, InStrRev(strTApp1CasePath, "\") + 1) & "/" & strCRA05
               rsTmp.MoveFirst
               Do While Not rsTmp.EOF
                  'FTP¤U¸üÀÉ®×
                  'Modified by Morgan 2025/3/28 +CPP19
                  'Modify By Sindy 2025/4/1 "" & RsTemp.Fields("CPP19") §ï "" & rsTmp.Fields("CPP19")
                  If PUB_GetFtpFile("" & rsTmp.Fields("CPP14"), pSavePath & "\" & rsTmp.Fields("CPP02"), , , , "" & rsTmp.Fields("CPP19") <> "") = False Then
                     GoTo ErrHand
                  End If
                  'FTP¤W¶ÇÀÉ®×
                  If PUB_FtpPutFile(pSavePath & "\" & rsTmp.Fields("CPP02"), strFilePath & "/" & rsTmp.Fields("CPP02"), , , Pub_GetSpecMan("FTP_VOL_IP_SALE1")) = False Then
                     GoTo ErrHand
                  End If
                  DoEvents
                  rsTmp.MoveNext
               Loop
            End If
         End If
      End If
      '2023/3/28 END
   End If
   
   PUB_AutoRecvCRLMain = True
   Set rsTmp = Nothing
   Set rsAD = Nothing 'Add By Sindy 2024/5/20
   
   Exit Function
   
ErrHand:
   PUB_AutoRecvCRLMain = False
   Set rsTmp = Nothing
   Set rsAD = Nothing 'Add By Sindy 2024/5/20
   If Err.Number <> 0 Then
      MsgBox Err.Description & vbCrLf & _
      ";strtmp1(0) = " & strTmp1(0) & vbCrLf & _
      ";strSql = " & strSql, vbCritical, "PUB_AutoRecvCRLMain"
   End If
End Function


'Add By Sindy 2022/9/26
'·s¼W°Ó¼Ð
Public Function PUB_AutoRecvCRL_T(strCRL01 As String, _
   Optional ByRef strPA01 As String, Optional ByRef strPA02 As String, Optional ByRef strPA03 As String, Optional ByRef strPA04 As String, _
   Optional ByVal bolNewCase As Boolean = False) As Boolean
Dim modCP() As String, modBase() As String '¦¬¤å ©M °ò¥»ÀÉ
Dim m_strControl As String '»ô³Æ¤éºÞ¨î
Dim mType As String, mCaseNo As String '¯S®íºÞ¨î
Dim strCRL() As String
Dim intKind As Integer, intWhere As Integer
Dim strTmpA As String
Dim intA As Integer
Dim rsAD As New ADODB.Recordset
Dim Rs As New ADODB.Recordset 'Add By Sindy 2024/5/27
Dim ii As Integer
Dim douStPrice As Double, douLowPrice As Double
Dim strAPP1 As String, strCRA10 As String, strCRA22 As String
Dim IsSaveData As Boolean
Dim modTBase() As String, k As Integer '¤À³Î®×®É¤l®×¨Ï¥Î
Dim intCRC As Integer
Dim strExSql As String 'Added by Lydia 2024/05/15
Dim strCRC09 As String 'Add By Sindy 2025/8/21

On Error GoTo ErrHand
   
   '±µ¬¢³æ¥DÀÉ
   ReDim Preserve strCRL(TF_CRL) As String
   strCRL(1) = strCRL01
   If ClsPDReadCRLDatabase(strCRL) = False Then
      GoTo ErrHand
   End If
   
   '¥Ó½Ð¤H1
   strExSql = "select * from ConsultRecApp" & _
            " where CRA01='" & strCRL01 & "' and CRA02=1"
   intA = 1
   Set Rs = ClsLawReadRstMsg(intA, strExSql)
   If intA = 1 Then
      strAPP1 = Rs.Fields("CRA05") & Rs.Fields("CRA06")
      strCRA10 = "" & Rs.Fields("CRA10") '±µ¬¢¤H
      strCRA22 = "" & Rs.Fields("CRA22") 'Ápµ¸¦a§}¶l»¼°Ï¸¹
   End If
   
'*********************
   '³]©w°}¦C
'*********************
   If ClsPDGetSystemKind(strCRL(7), intKind) = True Then
     Select Case intKind
        Case ±M§Q
           ReDim Preserve modBase(TF_PA) As String
        Case °Ó¼Ð
           ReDim Preserve modBase(TF_TM) As String
        Case ªk°È
           ReDim Preserve modBase(TF_LC) As String
        Case ÅU°Ý
           ReDim Preserve modBase(TF_HC) As String
        Case Else
           ReDim Preserve modBase(tf_SP) As String
     End Select
   End If
   ReDim Preserve modCP(TF_CP) As String
   
   If bolNewCase = True Then '«ü©w¬°·s®×³B²z(¥À®×¤§´X)
      strCRL(6) = "Y"
      modBase(1) = strPA01
      modBase(2) = strPA02
      modBase(3) = strPA03
      modBase(4) = strPA04
   Else
      '¨ú±o®×¸¹
      modBase(1) = strCRL(7)
      'ÂÂ®× ©Î ¤£¦P¼f¯Å,·|¬O·s®×¥À¸¹
      If strCRL(8) <> "" Then modBase(2) = strCRL(8)
      If strCRL(9) = "" Then
         modBase(3) = "0"
      Else
         modBase(3) = Right("0" & strCRL(9), 1)
      End If
      If strCRL(10) = "" Then
         modBase(4) = "00"
      Else
         modBase(4) = Right("00" & strCRL(10), 2)
      End If
   End If
   modCP(1) = modBase(1)
   modCP(2) = modBase(2)
   modCP(3) = modBase(3)
   modCP(4) = modBase(4)
   
   If strCRL(6) = "" Then 'ÂÂ®×
      'Modified by Lydia 2023/05/11 +false
      If PUB_ReadCaseData(modBase, intKind, intWhere, False) = False Then
         GoTo ErrHand
      End If
   '·s®×
   Else 'strCRL(6) = "Y" Then
      modBase(5) = strCRL(17) '®×¥ó¦WºÙ(¤¤)
      
      '°Ó¼ÐºØÃþ
      modBase(8) = strCRL(87)
      If modBase(1) = "CFT" And modBase(8) = "" Then modBase(8) = "1" 'Add By Sindy 2023/1/9 CFT°Ó¼ÐºØÃþ¥u¦³1.°Ó¼Ð
      'Modify By Sindy 2023/11/15 °Ó¼ÐºØÃþ<'A'®É¥u¦s°Ó¼ÐºØÃþ©óTM08¡A°Ó¼ÐºØÃþ>='A'®É«hTM08='1'¦P®É¦sTM72='°Ó¼ÐºØÃþ'¡C
      If strCRL(87) >= "A" Then
         modBase(8) = "1"
         modBase(72) = strCRL(87)
      End If
      '2023/11/15 END
      
      modBase(9) = strCRL(73) '°Ó«~Ãþ§O
      modBase(10) = strCRL(15) '¥Ó½Ð°ê®a
      modBase(35) = strCRL(16) '«È¤á®×¥ó®×¸¹
      modBase(15) = strCRL(125) '¼f©w¸¹
      modBase(12) = strCRL(71) '¥Ó½Ð®×¸¹ Add By Sindy 2024/2/2
      
      strExSql = "select * from ConsultRecApp" & _
               " where CRA01='" & strCRL01 & "'" & _
               " order by CRA02 asc"
      intA = 1
      Set Rs = ClsLawReadRstMsg(intA, strExSql)
      If intA = 1 Then
         Rs.MoveFirst
         '¥Ó½Ð¤H1~5
         For ii = 1 To 5
            If ii = 1 Then modBase(23) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            If ii = 2 Then modBase(78) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            If ii = 3 Then modBase(79) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            If ii = 4 Then modBase(80) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            If ii = 5 Then modBase(81) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            Rs.MoveNext
            If Rs.EOF Then Exit For
         Next ii
      End If
      
      'FC¥N²z¤H
      If strCRL(60) <> "" Then
         modBase(44) = ChangeCustomerL(strCRL(60) & strCRL(61))
      End If
      
      '¥Ó½Ð¤HÁpµ¸¤H½s¸¹
      If strCRA10 <> "" Then
         strExSql = "select * from potcustcont where pcc01='" & Left(strAPP1, 8) & "' and pcc05='" & strCRA10 & "'"
         intA = 1
         Set Rs = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            modBase(123) = "" & Rs.Fields("PCC02")
            '­Y­Ó®×±µ¬¢¤H»P«È¤áÀÉªº¹w³]±µ¬¢¤H¬Û¦P®É¤£¥²³]©w
            PUB_GetContact strAPP1, strTmpA, True
            If modBase(123) = strTmpA Then
               modBase(123) = ""
            End If
         End If
      End If
   End If
   modBase(45) = strCRL(77) '©¼©Ò®×¸¹
   
   '±µ¬¢°O¿ý³æ®×¥ó©Ê½è
   'Modify By Sindy 2022/10/26 ¨q¬ÂÄ±±o¥ý¨Ì´¼Åv¤H­û¿é¤Jªº¶¶§Ç§Y¥i
'   strexsql = "select CRC01,CRC02,CRC03,CRC04,CRC05,CRC06,CRC07,CRC08,1 as sort" & _
'            " from ConsultRecCMP" & _
'            " where CRC01='" & strCRL01 & "'" & _
'            " and instr('101,308',CRC03)>0" & _
'            " union select CRC01,CRC02,CRC03,CRC04,CRC05,CRC06,CRC07,CRC08,2 as sort" & _
'            " from ConsultRecCMP" & _
'            " where CRC01='" & strCRL01 & "'" & _
'            " and instr('101,308',CRC03)=0" & _
'            " order by sort asc,CRC03 asc"
   strExSql = "select *" & _
            " from ConsultRecCMP" & _
            " where CRC01='" & strCRL01 & "'" & _
            " order by CRC02 asc"
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strExSql)
   If intA = 1 Then
   rsAD.MoveFirst
   intCRC = 0
   Do While Not rsAD.EOF
      intCRC = intCRC + 1
      If intCRC > 1 Then '²Ä2µ§¦¬¤å­n²MªÅmodCP³¯¦C­È
         Erase modCP
         ReDim Preserve modCP(TF_CP) As String
         modCP(1) = modBase(1)
         modCP(2) = modBase(2)
         modCP(3) = modBase(3)
         modCP(4) = modBase(4)
      End If
      
      IsSaveData = False '*****
      modCP(9) = "A" & CompAutoNumberYear(GetTaiwanThisYear) '¦¬¤å¸¹ ex:AB1
      modCP(5) = strSrvDate(1) '¦¬¤å¤é
      modCP(10) = rsAD.Fields("CRC03") '®×¥ó©Ê½è
      'Modify By Sindy 2023/1/31
      If InStr(T®×¤£±¾´Á­­ªº®×¥ó©Ê½è, modCP(10)) > 0 And rsAD.RecordCount > 1 Then
         '¤£±¾´Á­­
      Else
      '2023/1/31 END
         modCP(6) = strCRL(12) '¥»©Ò´Á­­
         modCP(7) = strCRL(13) 'ªk©w´Á­­
      End If
      modCP(11) = "07" '®×¥ó¨Ó·½
      modCP(12) = GetST15(strCRL(3))
      modCP(13) = strCRL(3) '´¼Åv¤H­û
      
      '¤º°Ó¹w³]©Ó¿ì¤H³W«h
      modCP(14) = PUB_SetTxxCP14(modCP(1), modCP(2), modCP(3), modCP(4), strAPP1, modCP(13), modCP(10), modBase(10), strCRL01, strCRC09)
      
      'Modify By Sindy 2022/11/16 T-239673¤j³°¤À³Î±µ¬¢³æ(ÂÂ®×¬°¥À®×¤À³Î0¤¸·s®×¬°¤À³Î®×¦³¶O¥Î)
      If strCRL(15) = "020" And modCP(10) = "308" Then
         modCP(16) = 0 '¶O¥Î
         modCP(17) = 0 '³W¶O
         modCP(18) = 0 'ÂI¼Æ
      Else
      '2022/11/16 END
         modCP(16) = rsAD.Fields("CRC04") '¶O¥Î
         modCP(17) = rsAD.Fields("CRC05") '³W¶O
         modCP(18) = rsAD.Fields("CRC06") 'ÂI¼Æ
      End If
      modCP(19) = strCRL(39) '«áª÷
      'Modify By Sindy 2025/8/26 +, , strCRL(60) & strCRL(61)
      If ClsPDGetCaseLowPrice(strCRL(7), strCRL(15), modCP(10), douStPrice, douLowPrice, strCRL(87), strCRL(81), _
         strCRL(1), strCRL(5), strCRL(7), strCRL(8), strCRL(9), strCRL(10), , strCRL(60) & strCRL(61)) = 1 Then
         modCP(33) = douStPrice '¼Ð·Ç»ù
         modCP(34) = douLowPrice '©³»ù
      End If
      'Modify By Sindy 2023/5/18 Tªº1=¤£­­¨î,¤£¼g¤J¶i«×ÀÉ
      'Modify By Sindy 2024/3/18 ¤£¥Î§PÂ_, mark if
      'If strCRL(82) <> "1" Then
         modCP(141) = strCRL(82) '°e¥ó¤è¦¡
      'End If
      '2023/5/18 END
      modCP(142) = strCRL(83) '«ü©w°e¥ó¤é
      modCP(164) = strCRL(155) '«ü©w°e¥ó¤è¦¡ Add By Sindy 2023/12/12
      modCP(151) = strCRL(92) '¦¬¾Ú¦Û°Ê¦C¦L®É¶¡ÂI Add By Sindy 2023/4/18
      modCP(64) = "" & rsAD.Fields("CRC07") '³Æµù
      
      If rsAD.Fields("CRC03") = ²¾Âà Then
         'Åý»P¤H1-5,¨üÅý¤H1-5
         strExSql = "select * from ConsultRecApp" & _
                  " where CRA01='" & strCRL01 & "'" & _
                  " order by CRA02 asc"
         intA = 1
         Set Rs = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            Rs.MoveFirst
            '¥Ó½Ð¤H1~5
            For ii = 1 To 5
               If ii = 1 Then modCP(56) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               If ii = 2 Then modCP(89) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               If ii = 3 Then modCP(90) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               If ii = 4 Then modCP(91) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               If ii = 5 Then modCP(92) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               Rs.MoveNext
               If Rs.EOF Then Exit For
            Next ii
         End If
      End If
      If strCRL(59) <> "" Then modBase(136) = strCRL(59)       'ÃÒ®Ñ§Î¦¡
      modCP(140) = strCRL01 '±µ¬¢³æ½s¸¹
      
      '¹q¤l°e¥ó
      If strCRL(95) = "Y" Then
         modCP(118) = "YY"
      End If
      
      '¯S®íºÞ¨î
      mType = "": mCaseNo = ""
      'Modified by Lydia 2024/03/14 ²Ä1µ§¤~³B²z¯S®íºÞ¨î + intCRC = 1
      If intCRC = 1 And strCRL(74) <> "" And strCRL(55) <> "" Then
         If strCRL(74) = "3" Then
            If modCP(10) = ¥Ó½Ð Then
               mType = "CFT½q¨l­«·s¥Ó½Ð®×"
               mCaseNo = strCRL(55)
            End If
         ElseIf strCRL(74) = "2" Then
            'If strCRL(6) = "Y" And modBase(2) = "" Then '·s®×¥¼¨ú±o®×¸¹®É,¤~­n§ì­È
            'Modified by Lydia 2023/01/10  debug:¼Ú·ù¬O110©µ®i¶O¡]­^°ê¡^¡A­^°ê©µ®i·s®×­n§ì102©µ®i
            'If strCRL(6) = "Y" And modCP(10) = "110" Then '·s®×©µ®i¡]­^°ê¡^®É,¤~­n§ì­È
            'Modified by Lydia 2023/04/10 ¦P®É§PÂ_©e¥ô¥N²z¤H; ex. CFT-023552·s®×¬°©e¥ô¥N²z¤H
            'If strCRL(6) = "Y" And modCP(10) = "102" Then '·s®×©µ®i¡]­^°ê¡^®É,¤~­n§ì­È
            If strCRL(6) = "Y" And (modCP(10) = "102" Or modCP(10) = "710") Then
               mType = "CFT­^°ê²æ¼Ú®×"
               mCaseNo = strCRL(55)
            End If
         ElseIf strCRL(74) = "1" Then
            If strCRL(7) = "T" And InStr(TMQ_T®×, modCP(10)) > 0 Then
               mType = "T¬d¦W³æ"
               mCaseNo = strCRL(55)
               'Added Lydia 2022/11/02 ¨Ï¥ÎPUB_TQCtoTMQ¨ú±o¹ê»Úªº¬d¦W³æ¸¹,²Õ¦¨¥i¥H¨Ï¥ÎPUB_TMQtoCPªº¸ê®Æ=±µ¬¢³æ¬d¦W¥N¸¹|+¹ê»Ú¬d¦W³æ¸¹,ex.111002476|+HB1110046,
               If mCaseNo <> "" Then
                   'Modified by Lydia 2024/03/14 +False
                   Call PUB_TQCtoTMQ(False, modCP(12), modCP(13), mCaseNo, strTmpA)
                   mCaseNo = mCaseNo & "|" & strTmpA
               End If
               'end 2022/11/02
            End If
         ElseIf Len(strCRL(74)) = 2 Then
            mType = "LOS®×·½¦¬¤å"
            mCaseNo = strCRL(74) & "," & strCRL(55)
         End If
      End If
      
      '»ô³Æ¤é  --m_strControl
      m_strControl = ""
      '¤å¥ó¬O§_»ô³Æ(101¥Ó½Ð)¡B¸ê®Æ¬O§_»ô³Æ
      If strCRL(88) <> "" Then
          m_strControl = m_strControl & ",EP06|" & strCRL(88) & "|" & ""
      End If
      '¬O§_·|½Z
      If strCRL(89) <> "" Then
          m_strControl = m_strControl & ",EP34|" & strCRL(89)
      End If
      '¬O§_«æ¥ó
      If strCRL(90) <> "" Then
          m_strControl = m_strControl & ",CP122|" & strCRL(90)
      End If
      '¬d¦W¬O§_»ô³Æ
      If strCRL(137) <> "" Then
          m_strControl = m_strControl & ",CP143|" & IIf(strCRL(137) = "Y", strSrvDate(1), IIf(strCRL(137) = "N", "0", ""))
      End If
      If m_strControl <> "" Then m_strControl = Mid(m_strControl, 2)
      
'*********************
      '¦sÀÉ
'*********************
      'Modify By Sindy 2024/11/21 + intCRC
      If PUB_SaveFrm010004("frm090801_New", IIf(strCRL(6) = "Y", IIf(modBase(2) <> "" And intCRC > 1, 0, 1), 0), 0, 0, _
         modBase, modCP, strCRA22, m_strControl, IsSaveData, mType, mCaseNo, , intCRC) = False Then
         GoTo ErrHand
      Else
         'Modify By Sindy 2025/8/5 ¦¬¤å¤ÀªR±¾¬ÛÃöÁ`¦¬¤å¸¹
         If (modCP(10) = "303" Or (strCRL(74) = "4" And modCP(10) = "727")) _
            And strCRL(71) <> "" Then '§ó·s©µ´Á®×¬ÛÃöÁ`¦¬¤å¸¹=±µ¬¢³æÂI¿ïªº¤U¤@µ{§Ç¤å¸¹
            strExSql = "update caseprogress set cp43 = '" & strCRL(71) & "' where cp09 = '" & modCP(9) & "'"
            cnnConnection.Execute strExSql
         End If
         strTmp1(10) = ""
         If modCP(14) <> "" Then '¤w¦³©Ó¿ì¤H®É,¦^¼gCRC09
            strTmp1(10) = ",CRC09='" & modCP(14) & "'"
         ElseIf strCRC09 <> "" Then
            strTmp1(10) = ",CRC09='" & strCRC09 & "'"
         End If
         '§ó·s±µ¬¢°O¿ý³æ®×¥ó©Ê½èªºÁ`¦¬¤å¸¹
         'strExSql = "update ConsultRecCMP set CRC08='" & modCP(9) & "'" & IIf(modCP(14) = "", "", ",CRC09='" & modCP(14) & "'") &
         '2025/8/21 END
         strExSql = "update ConsultRecCMP set CRC08='" & modCP(9) & "'" & strTmp1(10) & _
                  " where CRC01='" & strCRL01 & "'" & _
                  " and CRC02=" & rsAD.Fields("CRC02") & " and CRC08 is null"
         cnnConnection.Execute strExSql, intI
         
         '³Ì«á¤@µ§ªºÀË¬d, §ó·s·s®×ªº®×¸¹
         If intCRC = rsAD.RecordCount Then
            strExSql = "update ConsultRecordList set CRL08='" & modBase(2) & "',CRL09='" & modBase(3) & "',CRL10='" & modBase(4) & "'" & _
                     " where CRL01='" & strCRL01 & "' and CRL06='Y' and CRL08 is null"
            cnnConnection.Execute strExSql
         End If
      End If
      
      rsAD.MoveNext
   Loop
   End If
   
   'ÀË¬d¦³µL°Ó¼Ð¤À³Î¤l®×­n³B²z
   If intKind = °Ó¼Ð And Val(strCRL(76)) > 0 Then
      'Added by Lydia 2024/11/21 T¤j³°®×¦¬¤å¤À³Î308¡A¬ÛÃöÁ`¦¬¤å¸¹=´_¼fCP43; --- from ¤º°Ó¤j³°¤§³¡¥÷®Ö»é°Ó«~²§°Ê
      If modBase(1) = "T" And modBase(10) = "020" Then
         strExSql = "update caseprogress set cp43=(select cp43 from consultreccmp, caseprogress where crc01='" & strCRL01 & "' and crc08=cp09(+) and cp10 ='401' and nvl(cp43,'N')<>'N' ) " & _
                   "where cp09=(select cp09 from consultreccmp, caseprogress where crc01='" & strCRL01 & "' and crc08=cp09(+) and cp10 ='308') and nvl(cp43,'N')='N' "
         cnnConnection.Execute strExSql
      End If
      'end 2024/11/21
      k = 0
      Do While k < Val(strCRL(76))
         k = k + 1
         Erase modTBase
         ReDim Preserve modTBase(TF_TM) As String
         Erase modCP
         ReDim Preserve modCP(TF_CP) As String
         '¬°·s®×
         modTBase(1) = strCRL(7)
         modTBase(3) = "0"
         modTBase(4) = "00"
         modCP(1) = modTBase(1)
         modCP(2) = modTBase(2)
         modCP(3) = modTBase(3)
         modCP(4) = modTBase(4)
         modTBase(5) = strCRL(17) '®×¥ó¦WºÙ(¤¤)
         modTBase(8) = strCRL(87) '°Ó¼ÐºØÃþ
         modTBase(10) = strCRL(15) '¥Ó½Ð°ê®a
         modBase(35) = strCRL(16)  '«È¤á®×¥ó®×¸¹
         modTBase(15) = strCRL(125) '¼f©w¸¹
         
         strExSql = "select * from ConsultRecApp" & _
                  " where CRA01='" & strCRL01 & "'" & _
                  " order by CRA02 asc"
         intA = 1
         Set Rs = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            Rs.MoveFirst
            '¥Ó½Ð¤H1~5
            For ii = 1 To 5
               If ii = 1 Then modTBase(23) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               If ii = 2 Then modTBase(78) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               If ii = 3 Then modTBase(79) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               If ii = 4 Then modTBase(80) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               If ii = 5 Then modTBase(81) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               Rs.MoveNext
               If Rs.EOF Then Exit For
            Next ii
         End If
         
         'FC¥N²z¤H
         If strCRL(60) <> "" Then
            modTBase(44) = ChangeCustomerL(strCRL(60) & strCRL(61))
         End If
   '      modTBase(45) = Trim(txtTrademark(33)) '©¼©Ò®×¸¹
         '¥Ó½Ð¤HÁpµ¸¤H½s¸¹
         If strCRA10 <> "" Then
            strExSql = "select * from potcustcont where pcc01='" & Left(strAPP1, 8) & "' and pcc05='" & strCRA10 & "'"
            intA = 1
            Set Rs = ClsLawReadRstMsg(intA, strExSql)
            If intA = 1 Then
               modTBase(123) = "" & Rs.Fields("PCC02")
               '­Y­Ó®×±µ¬¢¤H»P«È¤áÀÉªº¹w³]±µ¬¢¤H¬Û¦P®É¤£¥²³]©w
               PUB_GetContact strAPP1, strTmpA, True
               If modTBase(123) = strTmpA Then
                  modTBase(123) = ""
               End If
            End If
         End If
         
         '±µ¬¢°O¿ý³æ®×¥ó©Ê½è
         strExSql = "select * from ConsultRecCMP" & _
                  " where CRC01='" & strCRL01 & "' and CRC03='308'"
         intA = 1
         Set rsAD = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            modCP(9) = "A" & CompAutoNumberYear(GetTaiwanThisYear) '¦¬¤å¸¹ ex:AB1
            modCP(5) = strSrvDate(1) '¦¬¤å¤é
            modCP(6) = strCRL(12) '¥»©Ò´Á­­
            modCP(7) = strCRL(13) 'ªk©w´Á­­
            modCP(10) = rsAD.Fields("CRC03") '®×¥ó©Ê½è
            modCP(11) = "07" '®×¥ó¨Ó·½
            modCP(12) = GetST15(strCRL(3))
            modCP(13) = strCRL(3) '´¼Åv¤H­û
            
            '¤º°Ó¹w³]©Ó¿ì¤H³W«h
            modCP(14) = PUB_SetTxxCP14(modCP(1), modCP(2), modCP(3), modCP(4), strAPP1, modCP(13), modCP(10), modBase(10), strCRL01)
            
            'Modify By Sindy 2022/11/16 T-239673¤j³°¤À³Î±µ¬¢³æ(ÂÂ®×¬°¥À®×¤À³Î0¤¸·s®×¬°¤À³Î®×¦³¶O¥Î)
            If strCRL(15) = "020" And modCP(10) = "308" Then
               modCP(16) = rsAD.Fields("CRC04") '¶O¥Î
               modCP(17) = rsAD.Fields("CRC05") '³W¶O
               modCP(18) = rsAD.Fields("CRC06") 'ÂI¼Æ
            Else
            '2022/11/16 END
               modCP(16) = 0 '¶O¥Î
               modCP(17) = 0 '³W¶O
               modCP(18) = 0 'ÂI¼Æ
            End If
            modCP(19) = strCRL(39) '«áª÷
            'Modify By Sindy 2025/8/26 +, , strCRL(60) & strCRL(61)
            If ClsPDGetCaseLowPrice(strCRL(7), strCRL(15), modCP(10), douStPrice, douLowPrice, strCRL(87), strCRL(81), _
               strCRL(1), strCRL(5), strCRL(7), strCRL(8), strCRL(9), strCRL(10), , strCRL(60) & strCRL(61)) = 1 Then
               modCP(33) = douStPrice '¼Ð·Ç»ù
               modCP(34) = douLowPrice '©³»ù
            End If
            'Modify By Sindy 2024/3/18
            modCP(141) = strCRL(82) '°e¥ó¤è¦¡
            modCP(142) = strCRL(83) '«ü©w°e¥ó¤é
            modCP(164) = strCRL(155) '«ü©w°e¥ó¤è¦¡
            '2024/3/18 END
            'modCP(64) = "" & rsAD.Fields("CRC07") '³Æµù
            
            modCP(140) = strCRL01 '±µ¬¢³æ½s¸¹
            
            '¹q¤l°e¥ó
            If strCRL(95) = "Y" Then
               modCP(118) = "YY"
            End If
            
         '*********************
            '¦sÀÉ
         '*********************
            IsSaveData = False '*****
            'Modify By Sindy 2024/11/21 «e­±ÀË¬d¹L¥Ó½Ð¤H,©Ò¥H¤l®×¤£¥Î¦AÀË¬d¬G¶Ç¤J intCRC=99
            If PUB_SaveFrm010004("frm090801_New", 1, 0, 0, _
               modTBase, modCP, strCRA22, m_strControl, IsSaveData, mType, mCaseNo, , 99) = False Then
               GoTo ErrHand
            Else
               '§ó·s±µ¬¢°O¿ý³æªº¬ÛÃö®×¸¹-¤À³Î¤l®×
               strExSql = "update CONSULTRECORDLIST set CRL55=decode(CRL55,null,'',CRL55||',')||'" & modTBase(1) & "-" & modTBase(2) & "',CRL56='N'" & _
                        " where CRL01='" & strCRL01 & "'"
               cnnConnection.Execute strExSql
            End If
         End If
      Loop
   End If
   
   PUB_AutoRecvCRL_T = True
   Set rsAD = Nothing
   Set Rs = Nothing 'Add By Sindy 2024/5/27
   Exit Function
   
ErrHand:
   PUB_AutoRecvCRL_T = False
   Set rsAD = Nothing
   If Err.Number <> 0 Then
      MsgBox Err.Description & vbCrLf & _
      ";strexsql = " & strExSql, vbCritical, "PUB_AutoRecvCRL_T"
   End If
End Function

'Add By Sindy 2022/9/27
'¤º°Ó¹w³]©Ó¿ì¤H³W«h
'Modify By Sindy 2025/8/21
'   +, ByVal strCP140 As String: ¶Ç¤J·í®É±µ¬¢³æ½s¸¹
'   PUB_SetTxxCP14: ª½±µ¼g¤JCP14,¤£¸g¥DºÞ¤À®×
'   +, Optional ByRef strCRC09 As String = "": ¥u¬O,¹w¤À¥DºÞ¤À®×ªº©Ó¿ì¤H(¦s¤J±µ¬¢³æ®×¥ó©Ê½èÀÉªº¹w¤À©Ó¿ì¤H), ÁÙ¬O­n¸g¥DºÞ¤À®×
'­Y¤À®×¥DºÞ³£½Ð°²®É¡A«h¥Ñ¨t²Î¦Û°Ê¤À®×¤£»Ýµ¥¥DºÞ½T»{
Public Function PUB_SetTxxCP14(strSys As String, strCP02 As String, strCP03 As String, strCP04 As String, _
   strAPP1 As String, strCP13 As String, strCP10 As String, strNation As String, ByVal strCP140 As String, _
   Optional ByRef strCRC09 As String = "") As String
Dim i As Integer, strChkEmp As String
Dim strTM23 As String, salesNo As String, salesArea As String
Dim bolMCTF As Boolean
Dim strCPMList As String 'Add By Sindy 2023/3/28
Dim strExSql As String, intA As Integer, rsAD As New ADODB.Recordset 'Added by Lydia 2024/05/16
Dim bolIsRest1Day As Boolean, bolRest As Boolean 'Add By Sindy 2025/8/21
   
   'Add By Sindy 2025/9/9 «D¤º°Ó®×¥ó,¤£°õ¦æ
   'Modify By Sindy 2025/9/12 + And strSys <> "FCT": ¦]FCT¥¼"¦¬¤å"¹q¤l¤Æ
   If (Left(strSys, 1) <> "T" And Not (strSys = "FCT" And InStr(TMdebate, strCP10) > 0)) _
      Or strSys = "FCT" Then
      Exit Function
   End If
   '2025/9/9 END
   
   PUB_SetTxxCP14 = ""
   strCRC09 = "" 'Add By Sindy 2025/8/21
   
   strTM23 = GetPrjPeopleNum1(strSys & "-" & strCP02 & "-" & strCP03 & "-" & strCP04)
   salesArea = GetCuSales(strTM23, salesNo)
   bolMCTF = False
   If Mid(PUB_GetAKindSalesNo(strSys, strCP02, strCP03, strCP04), 1, 4) = "MCTF" Or _
      salesNo = "96029" Or salesNo = "96030" Then
      bolMCTF = True
   End If
   
   'Add By Sindy 2025/8/21
   '******************************************************************
   '³q¥Îªº§PÂ_
   '******************************************************************
   '¥Ø«e¦]¬°(303)©µ´Á©Î(201)¸É¥¿¦b¦¬¤å®É©|¨t²ÎµLªk½T»{¬O°w¹ï­þ¤@­Ó¨Ó¨çªº´Á­­°µ³B²z,
   '©Ò¥H¨t²ÎµLªk¦Û°Ê¤À®×, ¬G¼W¥[¤U¦C§PÂ_¬O§_¬°°Ó¤T²Õ®×¥óªº³W«h:
   '1.¤U¤@µ{§Ç¶È¦³¤@­Ó¥¼¦¬¤å´Á­­®É¡A§ì¸Ó´Á­­¤§¨Ó¨ç¶i«×©Ó¿ì¤H¨Ó§PÂ_°Ó¥Ó©Î°Óª§²Õ­t³d¡A¥i¹w³]¸Ó¨Ó¨ç¶i«×©Ó¿ì¤H
   '¡A°Óª§®×¤]¾A¥Î¡C
   '2.­Y¤U¤@µ{§Ç¦³¤@­Ó¥H¤W´Á­­®É¡A«h§ì¸Ó®×¥ó³Ì«á¤@¹D¨Ó¨ç¶i«×©Ó¿ì¤H¨Ó§PÂ_°Ó¥Ó©Î°Óª§²Õ­t³d¡A¥i¹w³]¸Ó¨Ó¨ç¶i«×©Ó¿ì¤H
   '¡A°Óª§®×¤]¾A¥Î¡C
   '±Æ°£µ{§Ç¦P¤¯
   '+(206)©ñ±ó±M¥ÎÅv
   If strCP10 = "303" Or strCP10 = "201" Or strCP10 = "206" Then
      strSql = "SELECT cp14,st02 FROM nextprogress,caseprogress,staff" & _
               " where NP02='" & strSys & "' and NP03='" & strCP02 & "' and NP04='" & strCP03 & "' and NP05='" & strCP04 & "'" & _
               " and NP06 is null and NP01=CP09" & _
               " and CP14=ST01(+) and ST04='1' and st03<>'P22'" & _
               " order by CP05 desc,CP09 desc"
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strSql)
      If intA = 1 Then
         strCRC09 = "" & rsAD.Fields("cp14")
      End If
      'Add By Sindy 2023/1/18
      '°Ó¥Ó¦Û°Ê¤À®×,·í¹J¡u201¸É¥¿¡B206©ñ±ó±M¥ÎÅv¡B303©µ´Á¡v®×¥ó©Ê½è®É
      '§PÂ_À³¤Àµ¹³Ì«á¤@¦ì©Ó¿ì¤H:¤£ºÞ¦³µL¦¬µo¤å¡BA,B,CÃþÁ`¦¬¤å¸¹³£ºâ¦ý­n±Æ°£µ{§Ç¦P¤¯
      If strCRC09 = "" Then 'And InStr("201,206,303", strCP10) > 0
         'Modify By Sindy 2025/8/22 + and (CP140<>'" & strCP140 & "' or CP140 is null)
         strExSql = "SELECT cp14 FROM caseprogress,staff" & _
                  " where cp01='" & strSys & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                  " and substr(cp09,1,1)<>'D' and cp159=0 and cp14 is not null" & _
                  " and cp14=st01(+) and st04='1' and st03<>'P22' and (CP140<>'" & strCP140 & "' or CP140 is null)" & _
                  " order by cp66 desc,cp67 desc"
         intA = 1
         Set rsAD = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            strCRC09 = rsAD.Fields("cp14")
         End If
      End If
      '2023/1/18 END
   End If
   '2025/8/21 END
   
   'Add By Sindy 2023/3/28 ³o¨Ç®×¥ó©Ê½è¤£¹w¤À©Ó¿ì¤H
   'T  303   ©µ´Á
   'strCPMList = "204,205,207,303,310,614,615,706,707"
   strCPMList = "204,205,207,310,614,615,706,707" 'Modify By Sindy 2025/8/21
'T  204   ·Ç³Æµ{§Ç
'T  205   ¨¥µüÅG½×
'T  207   Án©ú°Ñ³^
'T  310   ¼È½w¼f²z
'T  614   ¥¼¸É²z¥Ñ
'T  615   ¥¼µªÅG
'T  706   ¨ä¥L
'T  707   ½Õ¬d
   '******************************************************************
   '°Ó¥Óªº§PÂ_
   '******************************************************************
   'TC®×,±Ä¤H¤u¤À®×
   '«D¥xÆW®×¥ó¡B°Óª§®×¥ó,±Ä¤H¤u¤À®×
   'Modify By Sindy 2023/1/12 TS®×¦]­n¤À´¼Åv¤H­û,©Ò¥H±Ä¤H¤u¤À®×
   'Modify By Sindy 2023/3/28 µø¬°ª§Ä³®×¥ó©Ê½è¤£¹w¤À©Ó¿ì¤H + & "," & strCPMList
   'Modify By Sindy 2023/11/21 ¦³Ãö°Ó¼Ð®×¥ó¦Û°Ê¤À®×ªº±ø¥ó,½Ð½Õ¾ã¬°:
   '   T¥xÆW°Ó¥Ó®×¤ÎTC¥xÆW®×±Ä¨t²Î¦Û°Ê¤À®×
   '   °£¦¹¤§¥~ªº®×¥ó,¬Ò¦Ü¥DºÞ¤À®×°Ï, ±Ä¥DºÞ¤â°Ê¤À®×
   'If Left(strSys, 1) = "T" And strSys <> "TC" And strSys <> "TS" And InStr(TMdebate & "," & strCPMList, strCP10) = 0 And (strNation = "000" Or bolMCTF = True) Then
   If ((strSys = "T" And strNation = "000" And InStr(TMdebate & "," & strCPMList, strCP10) = 0) Or _
      (strSys = "TC" And strNation = "000")) _
      And PUB_SetTxxCP14 = "" And strCRC09 = "" Then
   '2023/11/21 END
      '1.ÂÂ®×®É, ¥ýÀË¬d¬O§_©|¦³¥¼µo¤åªº¶i«×¡A­Y¦³¡A«ö¥¼µo¤åªº©Ó¿ì¦P¤¯¤À®×¡F§Y®×¥óºû«ù¦P¤@©Ó¿ì¤H©Ó¿ì¡C
      '¦ý­Y¤H­ûÂ÷Â¾¡A¨Ì¤À°tªí¤À®×¡C
      If strCP02 <> "" Then
         'Modify By Sindy 2025/8/22 + and (CP140<>'" & strCP140 & "' or CP140 is null)
         'Modify By Sindy 2025/9/15 + and st93<>'T11' ±Æ°£°Óª§¤H­û
         strExSql = "SELECT cp14 FROM caseprogress,staff" & _
                  " where cp01='" & strSys & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                  " and substr(cp09,1,1)<>'D' and cp158=0 and cp159=0 and cp14 is not null" & _
                  " and cp14=st01(+) and st04='1' and st03<>'P22' and st93<>'T11' and (CP140<>'" & strCP140 & "' or CP140 is null)" & _
                  " order by cp66 desc,cp67 desc"
         intA = 1
         Set rsAD = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            PUB_SetTxxCP14 = rsAD.Fields("cp14")
         End If
      End If
      
      If PUB_SetTxxCP14 = "" Then
         '737.´¼°]¨ó§@§ìÅUªA²Õªº³]©w
         If strCP10 = "737" Then
            strExSql = "SELECT * FROM DutyZoneAssign,staff" & _
                     " where DZA02='W2001'" & _
                     " and DZA01=st01(+) and st04='1'" & _
                     " order by DZA01 asc"
            intA = 1
            Set rsAD = ClsLawReadRstMsg(intA, strExSql)
            If intA = 1 Then
               PUB_SetTxxCP14 = rsAD.Fields("DZA01")
            End If
         Else
            '2.¤£ºÞ·s®×¡BÂÂ®×¡A©Î¬O­ìÂÂ®×¬O§Oªº©Ó¿ì¤H©Ò©Ó¿ìªº¡A¤@«ß¨Ì¡i°Ó¥Ó©Ó¿ì¤H³d¥ô·~°È°Ï¤À°tªí¡j¤À®×¡C¤U­z¨Ò¥~±¡§Î
            'a.·í¤é¥»©Ò´Á­­®×¥ó, ©ÒÄÝ©Ó¿ì¤H½Ð°², ±Ä¤H¤u¤À®×
            'b.©Ó¿ì¤H½Ð°²:µomail³qª¾ "®×¥ó¦¬¤å³qª¾,¥»©Ò´Á­­¬°xxx/xx/xx¦]©Ó¿ì¤H½Ð°²,½Ð°Æ¥»¦¬¨üªÌ¥N¬°¿ì²z¡I"¥Ø«e°Æ¥»¥[±¾³qª¾¹Å¶²©M©Ó¼z
            'c.TC®×, ±Ä¤H¤u¤À®×
            '¥ý§ì¥Ó½Ð¤H1«e8½XÀu¥ý¡A¨S¦³¦A¤ñ¹ï6½X¡A¦AÀË¬d´¼Åv¤H­û
            strExSql = "SELECT * FROM DutyZoneAssign,staff" & _
                     " where DZA02='" & Left(strAPP1, 8) & "'" & _
                     " and DZA01=st01(+) and st04='1'" & _
                     " order by DZA01 asc"
            intA = 1
            Set rsAD = ClsLawReadRstMsg(intA, strExSql)
            If intA = 1 Then
               PUB_SetTxxCP14 = rsAD.Fields("DZA01")
            Else
               '§ì¥Ó½Ð¤H1«e6½XÅª¨ú
               strExSql = "SELECT * FROM DutyZoneAssign,staff" & _
                        " where DZA02='" & Left(strAPP1, 6) & "'" & _
                        " and DZA01=st01(+) and st04='1'" & _
                        " order by DZA01 asc"
               intA = 1
               Set rsAD = ClsLawReadRstMsg(intA, strExSql)
               If intA = 1 Then
                  PUB_SetTxxCP14 = rsAD.Fields("DZA01")
               Else
                  '¨Ì´¼Åv¤H­ûÅª¨ú
                  strExSql = "SELECT * FROM DutyZoneAssign,staff" & _
                           " where DZA02='" & strCP13 & "'" & _
                           " and DZA01=st01(+) and st04='1'" & _
                           " order by DZA01 asc"
                  intA = 1
                  Set rsAD = ClsLawReadRstMsg(intA, strExSql)
                  If intA = 1 Then
                     PUB_SetTxxCP14 = rsAD.Fields("DZA01")
                  Else
                     'ÀË¬d¬O§_MCTF´¼Åv¤H­û
                     For i = 1 To 7
                        strChkEmp = "MCTF" & Format(i, "00")
                        If InStr(Pub_GetSpecMan(strChkEmp), strCP13) > 0 Then
                           strExSql = "SELECT * FROM DutyZoneAssign,staff" & _
                                    " where DZA02='" & strChkEmp & "'" & _
                                    " and DZA01=st01(+) and st04='1'" & _
                                    " order by DZA01 asc"
                           intA = 1
                           Set rsAD = ClsLawReadRstMsg(intA, strExSql)
                           If intA = 1 Then
                              PUB_SetTxxCP14 = rsAD.Fields("DZA01")
                              Exit For
                           End If
                        End If
                     Next i
                  End If
               End If
            End If
         End If
      End If
'   Else
'      Exit Function
   End If
   
   'Add By Sindy 2025/8/21
   '******************************************************************
   '°Óª§ªº§PÂ_
   '******************************************************************
   '°Ó¤@²Õ®×¥ó¤§¥DºÞ¤À®×¨Æ©y¡A¥[¤J¹w³]³W«h¤Î¥DºÞ³£½Ð°²±¡§Î±±¨î
   '1.ÂÂ®×¹w³]¤§©Ó¿ì¤H-¬°«e¤@°Óª§µ{§Ç¤§©Ó¿ì¤H
   '2.·s®×¹w³]¤§©Ó¿ì¤H-
   '  ¤j³°®×: B0011=¼Ú¶§«C
   '  ¥xÆW®×: A5026=ªL¤_¤è
   '3.¤W¦C¹w³]¤§©Ó¿ì¤H¡A¥Ñ¤À®×¥DºÞ¦Ü¨t²Î§@½T»{¡C
   '4.¤j³°®×¹w³]¤§¥N²z¤H-¥¨¨Ê¡A¥Ñ¤À®×¥DºÞ¦Ü¨t²Î§@½T»{¡C
   If PUB_SetTxxCP14 = "" And strCRC09 = "" And _
      (InStr(TMdebate, strCP10) > 0 Or strNation <> "000") Then
      'ÂÂ®×
      If strCP02 <> "" Then
         strExSql = "SELECT cp14 FROM caseprogress,staff" & _
                  " where cp01='" & strSys & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                  " and substr(cp09,1,1)<>'D' and cp14 is not null" & _
                  " and cp14=st01(+) and st04='1' and st03<>'P22' and st93='T11' and (CP140<>'" & strCP140 & "' or CP140 is null)" & _
                  " order by cp66 desc,cp67 desc"
         intA = 1
         Set rsAD = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            strCRC09 = rsAD.Fields("cp14")
         End If
      '·s®×
      'Modify By Sindy 2025/9/11 ÂÂ®×³W«h§ì¤£¨ì©Ó¿ì¤H®É,´N±Ä·s®×ªº³W«h
      'Else
      End If
      If strCRC09 = "" Then
      '2025/9/11 End
         If strNation <> "000" Then
            strCRC09 = "B0011"
         Else
            strCRC09 = "A5026"
         End If
      End If
   End If
   
   'ÀË¬d¤À®×¥DºÞ¬O§_¦³¥þ³¡³£½Ð°²ªºª¬ªp¡A­Y¦³,«h§ï¥Ñ¨t²Î¦Û°Ê¤À®×¤£»Ýµ¥¥DºÞ½T»{
   If strCRC09 <> "" Then
      '°Ó¤T²Õ°Ó¥Ó¥DºÞ:«Dª§Ä³®×¥ó,¥xÆW®×
      If PUB_GetST93(strCRC09) = "T31" Then
         strSql = "select st93,st02,staff_right.* from staff_right,staff where sr02='frm210156' and sr01=st01(+)" & _
                  " and st93='T31'" & _
                  " order by st01 asc"
      '°Ó¼Ð¤@²Õ¥DºÞ
      Else
         strSql = "select st93,st02,staff_right.* from staff_right,staff where sr02='frm210156' and sr01=st01(+)" & _
                  " and st93='T11'" & _
                  " order by st01 asc"
      End If
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strSql)
      If intA = 1 Then
         rsAD.MoveFirst
         Do While rsAD.EOF = False
            strExc(10) = rsAD.Fields("sr01") '¤À®×¥DºÞ
            'ÀË¬d¬O§_¦³¥ð°²
            bolRest = CheckIsPersonRest(strExc(10), strSrvDate(1), Left(Right("000000" & ServerTime, 6), 2) & ":" & Mid(Right("000000" & ServerTime, 6), 3, 2), , bolIsRest1Day)
            If bolRest = True Then '·í®É¥ð°²
               If bolIsRest1Day = True Then '¾ã¤é¥ð°²
                  PUB_SetTxxCP14 = strCRC09
               Else
                  '(¦¬¤å·í®É¤Î·í¤Ñ16:55¤À¥DºÞ¦P³B©ó½Ð°²ªºª¬ºA®É)¡A«h¨Ì¹w³]©Ó¿ì¤H³W«hª½±µ°µ¥DºÞ¤À®×ªº½T»{¡C
                  bolRest = CheckIsPersonRest(strExc(10), strSrvDate(1), "16:55")
                  If bolRest = True Then
                     PUB_SetTxxCP14 = strCRC09
                  Else
                     PUB_SetTxxCP14 = ""
                     Exit Do
                  End If
               End If
            Else
               PUB_SetTxxCP14 = ""
               Exit Do
            End If
            rsAD.MoveNext
         Loop
      End If
   End If
   '2025/8/21 END
   
   Set rsAD = Nothing 'Added by Lydia 2024/05/16
End Function

'Add By Sindy 2022/10/10
'±M§Q
Public Function PUB_AutoRecvCRL_P(strCRL01 As String, Optional ByVal strCRCWhere As String = "", _
   Optional ByRef strPA01 As String, Optional ByRef strPA02 As String, Optional ByRef strPA03 As String, Optional ByRef strPA04 As String, _
   Optional ByVal bolNewCase As Boolean = False) As Boolean
Dim modCP() As String, modBase() As String '¦¬¤å ©M °ò¥»ÀÉ
Dim m_strControl As String '»ô³Æ¤éºÞ¨î
Dim mType As String, mCaseNo As String '¯S®íºÞ¨î
Dim strCRL() As String
Dim intKind As Integer, intWhere As Integer
Dim strTmpA As String
Dim intA As Integer
Dim rsAD As New ADODB.Recordset
Dim rsPD As New ADODB.Recordset
Dim intP As Integer
Dim ii As Integer
Dim douStPrice As Double, douLowPrice As Double
Dim strAPP1 As String, strCRA10 As String, strCRA22 As String
Dim mChkStr As String, mRetVal As String
Dim IsSaveData As Boolean
Dim strInventorNo As String
Dim mTCTVal As String '¶Ç¤Jµe­±¦³Ãö©R¦W§@·~ªº¸ê®Æ
Dim mTCTList As String  '¦^¶Ç©R¦W§@·~¤@¨Ö²£¥Í¤§¦¬¤å¸¹
Dim intCRC As Integer
Dim strCRC09 As String, strCPM35 As String
Dim bolUpdPA179 As String 'Add By Sindy 2023/3/30
Dim strExSql As String 'Added by Lydia 2024/05/16

On Error GoTo ErrHand

   '±µ¬¢³æ¥DÀÉ
   ReDim Preserve strCRL(TF_CRL) As String
   strCRL(1) = strCRL01
   If ClsPDReadCRLDatabase(strCRL) = False Then
      GoTo ErrHand
   End If
   
   'Add By Sindy 2023/1/31
   '·s¼W±µ¬¢³æ«È¤áµo©ú¤H¸ê®Æ
   If PUB_InsertCRLInventor(strCRL01) = False Then
      GoTo ErrHand
   End If
   
   '¥Ó½Ð¤H1
   strExSql = "select * from ConsultRecApp" & _
            " where CRA01='" & strCRL01 & "' and CRA02=1"
   intA = 1
   Set RsTemp = ClsLawReadRstMsg(intA, strExSql)
   If intA = 1 Then
      strAPP1 = RsTemp.Fields("CRA05") & RsTemp.Fields("CRA06")
      strCRA10 = "" & RsTemp.Fields("CRA10") '±µ¬¢¤H
      strCRA22 = "" & RsTemp.Fields("CRA22") 'Ápµ¸¦a§}¶l»¼°Ï¸¹
   End If

'*********************
   '³]©w°}¦C
'*********************
   If ClsPDGetSystemKind(strCRL(7), intKind) = True Then
     Select Case intKind
        Case ±M§Q
           ReDim Preserve modBase(TF_PA) As String
        Case °Ó¼Ð
           ReDim Preserve modBase(TF_TM) As String
        Case ªk°È
           ReDim Preserve modBase(TF_LC) As String
        Case ÅU°Ý
           ReDim Preserve modBase(TF_HC) As String
        Case Else
           ReDim Preserve modBase(tf_SP) As String
     End Select
   End If
   ReDim Preserve modCP(TF_CP) As String
   
   If bolNewCase = True Then '«ü©w¬°·s®×³B²z(¥À®×¤§´X)
      strCRL(6) = "Y"
      modBase(1) = strPA01
      modBase(2) = strPA02
      modBase(3) = strPA03
      modBase(4) = strPA04
   Else
      '¨ú±o®×¸¹
      modBase(1) = strCRL(7)
      If strCRL(8) <> "" Then modBase(2) = strCRL(8)
      If strCRL(9) = "" Then
         modBase(3) = "0"
      Else
         modBase(3) = Right("0" & strCRL(9), 1)
      End If
      If strCRL(10) = "" Then
         modBase(4) = "00"
      Else
         modBase(4) = Right("00" & strCRL(10), 2)
      End If
   End If
   modCP(1) = modBase(1)
   modCP(2) = modBase(2)
   modCP(3) = modBase(3)
   modCP(4) = modBase(4)
   
   If strCRL(6) = "" Then  'ÂÂ®×
      'Modified by Lydia 2023/05/11 +false
      If PUB_ReadCaseData(modBase, intKind, intWhere, False) = False Then
         GoTo ErrHand
      End If
   '·s®×
   Else 'strCRL(6) = "Y" Then
      '°ò¥»ÀÉ
      modBase(5) = strCRL(17) '®×¥ó¦WºÙ(¤¤)
'      modBase(6) = txtPatent(6)  '®×¥ó¦WºÙ(­^)
'      modBase(7) = txtPatent(7)  '®×¥ó¦WºÙ(¤é)
      modBase(8) = strCRL(87) '±M§QºØÃþ
      modBase(9) = strCRL(15) '¥Ó½Ð°ê®a
      
      strExSql = "select * from ConsultRecApp" & _
               " where CRA01='" & strCRL01 & "'" & _
               " order by CRA02 asc"
      intA = 1
      Set RsTemp = ClsLawReadRstMsg(intA, strExSql)
      If intA = 1 Then
         RsTemp.MoveFirst
         '¥Ó½Ð¤H1~5
         For ii = 1 To 5
            If ii = 1 Then modBase(26) = ChangeCustomerL(RsTemp.Fields("CRA05") & RsTemp.Fields("CRA06"))
            If ii = 2 Then modBase(27) = ChangeCustomerL(RsTemp.Fields("CRA05") & RsTemp.Fields("CRA06"))
            If ii = 3 Then modBase(28) = ChangeCustomerL(RsTemp.Fields("CRA05") & RsTemp.Fields("CRA06"))
            If ii = 4 Then modBase(29) = ChangeCustomerL(RsTemp.Fields("CRA05") & RsTemp.Fields("CRA06"))
            If ii = 5 Then modBase(30) = ChangeCustomerL(RsTemp.Fields("CRA05") & RsTemp.Fields("CRA06"))
            RsTemp.MoveNext
            If RsTemp.EOF Then Exit For
         Next ii
      End If
      modBase(48) = strCRL(16) '«È¤á®×¥ó®×¸¹
      'FC¥N²z¤H
      If strCRL(60) <> "" Then
         modBase(75) = ChangeCustomerL(strCRL(60) & strCRL(61))
      End If
      modBase(77) = strCRL(77) '©¼©Ò®×¸¹
'      'Ápµ¸¤H1~2 =>  frm010007_1.bolOK
'      If bolCancel = True Then
'          modBase(51) = strPA51s
'          modBase(52) = strPA52s
'          modBase(53) = strPA53s
'          modBase(54) = strPA54s
'          modBase(55) = strPA55s
'          modBase(56) = strPA56s
'      End If
      modBase(158) = strCRL(81) '®×¥óÄÝ©Ê
      
      '¤j³°µo©ú¥ÍÂå®×¬O§_·sÃÄ±M§Q³]©w
      If strCRL(140) <> "" Then
         modBase(176) = strCRL(140)
      End If
      
      '¥Ó½Ð¤HÁpµ¸¤H½s¸¹
      If strCRA10 <> "" Then
         strExSql = "select * from potcustcont where pcc01='" & Left(strAPP1, 8) & "' and pcc05='" & strCRA10 & "'"
         intA = 1
         Set RsTemp = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            modBase(149) = "" & RsTemp.Fields("PCC02")
            '­Y­Ó®×±µ¬¢¤H»P«È¤áÀÉªº¹w³]±µ¬¢¤H¬Û¦P®É¤£¥²³]©w
            PUB_GetContact strAPP1, strTmpA, True
            If modBase(149) = strTmpA Then
               modBase(149) = ""
            End If
         End If
      End If
   End If
   
   '±µ¬¢°O¿ý³æ®×¥ó©Ê½è
   'Modify By Sindy 2022/10/26 ¨q¬ÂÄ±±o¥ý¨Ì´¼Åv¤H­û¿é¤Jªº¶¶§Ç§Y¥i
'   strexsql = "select CRC01,CRC02,CRC03,CRC04,CRC05,CRC06,CRC07,CRC08,1 as sort" & _
'            " from ConsultRecCMP" & _
'            " where CRC01='" & strCRL01 & "'" & _
'            " and instr('" & NewCasePtyList & "',CRC03)>0" & strCRCWhere & _
'            " union select CRC01,CRC02,CRC03,CRC04,CRC05,CRC06,CRC07,CRC08,2 as sort" & _
'            " from ConsultRecCMP" & _
'            " where CRC01='" & strCRL01 & "'" & _
'            " and instr('" & NewCasePtyList & "',CRC03)=0" & strCRCWhere & _
'            " order by sort asc,CRC03 asc"
   strExSql = "select *" & _
            " from ConsultRecCMP" & _
            " where CRC01='" & strCRL01 & "'" & strCRCWhere
   'Modify By Sindy 2023/5/10
   If InStr(UCase(strCRCWhere), UCase("order by")) = 0 Then
      strExSql = strExSql & " order by CRC02 asc"
   End If
   '2023/5/10 END
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strExSql)
   If intA = 1 Then
   rsAD.MoveFirst
   intCRC = 0
   Do While Not rsAD.EOF
      intCRC = intCRC + 1
      If intCRC > 1 Then '²Ä2µ§¦¬¤å­n²MªÅmodCP³¯¦C­È
         Erase modCP
         ReDim Preserve modCP(TF_CP) As String
         modCP(1) = modBase(1)
         modCP(2) = modBase(2)
         modCP(3) = modBase(3)
         modCP(4) = modBase(4)
      End If
      
      '¦¬¤åCaseProgress
      IsSaveData = False '*****
      modCP(9) = "A" & CompAutoNumberYear(GetTaiwanThisYear) '¦¬¤å¸¹ ex:AB1
      modCP(5) = strSrvDate(1) '¦¬¤å¤é
      modCP(10) = rsAD.Fields("CRC03") '®×¥ó©Ê½è
      'Modify By Sindy 2023/1/31
      If InStr(P®×¤£±¾´Á­­ªº®×¥ó©Ê½è, modCP(10)) > 0 And rsAD.RecordCount > 1 Then
         '¤£±¾´Á­­
      Else
      '2023/1/31 END
         modCP(6) = strCRL(12) '¥»©Ò´Á­­
         modCP(7) = strCRL(13) 'ªk©w´Á­­
      End If
      modCP(11) = "07" '®×¥ó¨Ó·½
      modCP(12) = GetST15(strCRL(3))
      modCP(13) = strCRL(3) '´¼Åv¤H­û
      'Add By Sindy 2023/1/12
      strCPM35 = PUB_GetCPM35(strCRL01, strCRL(7))
      If strCPM35 = "2" Or strCPM35 = "3" Then
         modCP(14) = PUB_SetPxxCP14(modCP(1), modCP(2), modCP(3), modCP(4), modCP(10))
      End If
      '2023/1/12 END
      
      'ÂÂ®×:P®×»âÃÒ¤ÎÃº¦~¶O,¦~¶O
      'Morgan»¡¤º±Mªº¤º±M»âÃÒ¦~¶O¾ã§åµo¤åfrm040104_i¡A¤w¸g¤£¦Ò¼{CP157¤ÎCP140¡A©Ò¥H»âÃÒ¦~¶Oªº¯S®íª¬ªp¥u­n¤£¹w³]©Ó¿ì¤H§Y¥i¡C
      'Modify By Sindy 2023/4/18 P®×¦¬¤å(601)»âÃÒ¡B(605)Ãº¦~¶O¥u­n¬O1.ÂÂ®×2.«D³¬¨÷¡A
      '                          ´¼Åv¦P¤¯¦¬¤å§¹¦¨¸g¥Ñ¥DºÞÃ±¥iªº®×¥ó¡A§¡¥Ñ¨t²Î¦Û°Ê¤À®×¦Ü¹w³]ªº©Ó¿ì¤H¡A¤£¥²¦A¸g¥Ñµ{§Ç¤H­û¶i¦æ¤À®×§@·~¡C
      'Modify By Sindy 2025/1/14 ¬Â¬Â´£,®×¥ó©Ê½è(606)ºû«ù¶O¡A¥i¦Û°Ê¦¬¤å¤Î¤À®×¡A³W«h¦P(601)»âÃÒ,(605)¦~¶O
      If strCRL(6) = "" And strCRL(7) = "P" And (modCP(10) = "601" Or modCP(10) = "605" Or modCP(10) = "606") And _
         Val(strCRL(144)) > 0 And Val(strCRL(145)) > 0 Then 'And strCRL(69) = ""
'         If InStr(strCRL(70), "Ãº¶O¦~«×¡G") = 0 Then
'            strtmp1(10) = Trim(strCRL(70))
'         Else
'            strtmp1(10) = Trim(Replace(strCRL(70), Mid(strCRL(70), InStr(strCRL(70), "Ãº¶O¦~«×¡G"), InStr(strCRL(70), vbCrLf) + 1), ""))
'         End If
'         If Len(strtmp1(10)) = 0 Then
         strExSql = "select pa01" & _
                  " from Patent" & _
                  " where PA01='" & modCP(1) & "' and PA02='" & modCP(2) & "'" & _
                  " and PA03='" & modCP(3) & "' and PA04='" & modCP(4) & "'" & _
                  " and PA57 is null"
         intA = 1
         Set RsTemp = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            'Added by Morgan 2025/1/22
            If strSrvDate(1) >= P·~°È°Ï¹º¤À±Ò¥Î¤é Then
               modCP(14) = PUB_GetPHandler(modCP(1) & modCP(2) & modCP(3) & modCP(4))
            Else
            'end 2025/1/22
               If strCRL(15) = "000" Then
                  modCP(14) = Pub_GetSpecMan("A1") 'P»OÆW®×»âÃÒ¡BÃº¦~¶O
               Else
                  modCP(14) = Pub_GetSpecMan("A111") 'P«D»OÆW®×»âÃÒ¡BÃº¦~¶O
               End If
            End If
         End If
      
      'Memo by Morgan 2025/8/4 ±q¤U­±²¾¤W¨Ó
      'Add By Sindy 2025/1/2 ª´­µ¥D¥ô´£:¥Ø«eP®×ªº¦~¶O¬O¥i¥H¦Û°Ê±¾©Ó¿ì¤Hªº¡A
      '½Ð¨ó§U½T»{CFP®×¤ñ·Ó¿ì²z¡A¦~¶O¡Bºû«ù¶O¡B©µ®i¶O¦¬¤å«á¦Û°Ê¨Ì¾Ú·~°È¤À°Ï±¾©Ó¿ì¤H
      ElseIf strCRL(6) = "" And strCRL(7) = "CFP" And (modCP(10) = ¦~¶O Or modCP(10) = ºû«ù¶O Or modCP(10) = ©µ®i¶O) _
         And Val(strCRL(144)) > 0 And Val(strCRL(145)) > 0 Then
         modCP(14) = PUB_GetCFPHandler(strCRL(7) & "-" & strCRL(8) & "-" & strCRL(9) & "-" & strCRL(10))
      '2025/1/2 END
      
      'Added by Morgan 2025/8/4 ±qPUB_SaveFrm010005²¾¨Ó
      'P/CFP ³]©w¬°µ{§Ç©Ó¿ì¥B¤£»Ý±M·~³¡¥DºÞ¤À®×ªº©Ê½è¡A©Ó¿ì¤H³£¦Û°Ê¹w³]¬°µ{§Ç¤H­û¡A­Y¦³»Ý­n¡A¤À®×¤H­û¦A¦Û¦æ­×§ï¡A¦ýCFP¹êÅé¼f¬d°£¥~--³¢
      'µ{§Ç©Ó¿ìªº©Ê½è­n¦b¦¹¥ý³]©w¡A§_«h­Y±µ¬¢³æ¦P®É¦³¤uµ{®vªº®×¥ó©Ê½è·|¦b¥DºÞ¤À®×®É³Q¤@¨Ö³]©w¡A¦ý¹ê¼f³W«h¯S§O°£¥~¡C
      'Modified by Morgan 2025/8/25 +ÂÂ®×±ø¥ó(·s®×¦¹³BµL¥»©Ò¸¹·|§ì¿ù¡A§ï¤À®×µe­±¦A¹w³]¡A¥B¤]»Ý¤À®×¤H­û½T»{¡A¤£¥i¦Û°Ê¤W¤À®×¤é)
      ElseIf strCRL(6) = "" And modCP(14) = "" And strCPM35 = "2" And (modCP(1) = "P" Or modCP(1) = "CFP") And modCP(10) <> "416" Then
         If modCP(1) = "CFP" Then
            modCP(14) = PUB_GetCFPHandler(modCP(1) & modCP(2) & modCP(3) & modCP(4))
         Else
            modCP(14) = PUB_GetPHandler(modCP(1) & modCP(2) & modCP(3) & modCP(4))
         End If
      'end 2025/8/4

      End If
      '2023/4/18 END
      
'Removed by Morgan 2025/8/4 ²¾¨ì¤W­±
'      'Add By Sindy 2025/1/2 ª´­µ¥D¥ô´£:¥Ø«eP®×ªº¦~¶O¬O¥i¥H¦Û°Ê±¾©Ó¿ì¤Hªº¡A
'      '½Ð¨ó§U½T»{CFP®×¤ñ·Ó¿ì²z¡A¦~¶O¡Bºû«ù¶O¡B©µ®i¶O¦¬¤å«á¦Û°Ê¨Ì¾Ú·~°È¤À°Ï±¾©Ó¿ì¤H
'      If strCRL(6) = "" And strCRL(7) = "CFP" And (modCP(10) = ¦~¶O Or modCP(10) = ºû«ù¶O Or modCP(10) = ©µ®i¶O) _
'         And Val(strCRL(144)) > 0 And Val(strCRL(145)) > 0 Then
'         modCP(14) = PUB_GetCFPHandler(strCRL(7) & "-" & strCRL(8) & "-" & strCRL(9) & "-" & strCRL(10))
'      End If
'      '2025/1/2 END
'end 2025/8/4
      
      
      'Added by Morgan 2025/5/16
      '¤j³°±M§QÅvµû»ù³ø§i423¦Û°Ê±¾µ{§Ç¤H­û--³¢
      'Removed by Morgan 2025/5/23 ¨ú®ø¡A§ï¦b¤À®×²Î¤@¹w³]
      'If strCRL(6) = "" And strCRL(7) = "P" And (strCRL(15) = "020" And modCP(10) = "423") Then
      '   modCP(14) = PUB_GetPHandler(modCP(1) & modCP(2) & modCP(3) & modCP(4))
      'End If
      'end 2025/5/23
      'end 2025/51/6
      
      modCP(16) = rsAD.Fields("CRC04") '¶O¥Î
      modCP(17) = rsAD.Fields("CRC05") '³W¶O
      modCP(18) = rsAD.Fields("CRC06") 'ÂI¼Æ
      modCP(19) = strCRL(39) '«áª÷
      'Modify By Sindy 2025/8/26 +, , strCRL(60) & strCRL(61)
      If ClsPDGetCaseLowPrice(strCRL(7), strCRL(15), modCP(10), douStPrice, douLowPrice, strCRL(87), strCRL(81), _
         strCRL(1), strCRL(5), strCRL(7), strCRL(8), strCRL(9), strCRL(10), , strCRL(60) & strCRL(61)) = 1 Then
         modCP(33) = douStPrice '¼Ð·Ç»ù
         modCP(34) = douLowPrice '©³»ù
      End If
      If strCRL(59) <> "" Then modBase(178) = strCRL(59) 'ÃÒ®Ñ§Î¦¡
      modCP(141) = strCRL(82) '°e¥ó¤è¦¡
      modCP(142) = strCRL(83) '«ü©w°e¥ó¤é
      modCP(164) = strCRL(155) '«ü©w°e¥ó¤è¦¡ Add By Sindy 2023/12/12
      modCP(151) = strCRL(92) '¦¬¾Ú¦Û°Ê¦C¦L®É¶¡ÂI Add By Sindy 2023/4/18
      modCP(64) = "" & rsAD.Fields("CRC07") '³Æµù
      modCP(86) = "N" '¦¬¨ì¤À©Ò±µ¬¢³æ:N.¦Û°Ê¦¬¤å
      
      'Add By Sindy 2023/10/30 405.¥Ó½ÐÀu¥ýÅvÃÒ©ú®Ñ¤~»Ý­nÀx¦s¦¹Äæ¦ì­È
      If (modBase(1) = "P" Or modBase(1) = "CFP") And _
         modCP(10) = "405" Then
      '2023/10/30 END
         modCP(71) = strCRL(58) 'Àu¥ýÅv¥÷¼Æ
      End If
      modCP(140) = strCRL01 '±µ¬¢³æ½s¸¹
      
      'Add By Sindy 2023/3/30 ªk°ê(203)³]­p¥Ó½Ð(103)ªº¶Â¥Õ¹Ï¦¡/±m¦â¹Ï¦¡³]©w¤£¬O¨­¤ÀºØÃþ¡A¤£­n¦^¼g PA179¡C
      If strSrvDate(1) >= PA179±Ò¥Î¤é Then
         If strCRL(94) <> "" Then
            bolUpdPA179 = True
            If strCRL(15) = "203" Then
               strExSql = "select *" & _
                        " from ConsultRecCMP" & _
                        " where CRC01='" & strCRL01 & "' and CRC03='103'"
               intA = 1
               Set RsTemp = ClsLawReadRstMsg(intA, strExSql)
               If intA = 1 Then
                  bolUpdPA179 = False
               End If
            End If
            'Modify By Sindy 2024/7/8 ¼W¥[§PÂ_P¤j³°®×¡A¤£­n¦^¼g PA179¡C
            If modBase(1) = "P" And strCRL(15) = "020" Then
               bolUpdPA179 = False
            End If
            '2024/7/8 END
            If bolUpdPA179 = True Then
               modBase(179) = strCRL(94)
            End If
         End If
      End If
      '2023/3/30 END
      
      If rsAD.Fields("CRC03") = Åý»P Or rsAD.Fields("CRC03") = ±M§QÅvÅý»P Or rsAD.Fields("CRC03") = ¦X¨Ö Or rsAD.Fields("CRC03") = Ä~©Ó Then
         'Åý»P¤H1-5,¨üÅý¤H1-5
         strExSql = "select * from ConsultRecApp" & _
                  " where CRA01='" & strCRL01 & "'" & _
                  " order by CRA02 asc"
         intA = 1
         Set RsTemp = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            RsTemp.MoveFirst
            '¥Ó½Ð¤H1~5
            For ii = 1 To 5
               If ii = 1 Then modCP(56) = ChangeCustomerL(RsTemp.Fields("CRA05") & RsTemp.Fields("CRA06"))
               If ii = 2 Then modCP(89) = ChangeCustomerL(RsTemp.Fields("CRA05") & RsTemp.Fields("CRA06"))
               If ii = 3 Then modCP(90) = ChangeCustomerL(RsTemp.Fields("CRA05") & RsTemp.Fields("CRA06"))
               If ii = 4 Then modCP(91) = ChangeCustomerL(RsTemp.Fields("CRA05") & RsTemp.Fields("CRA06"))
               If ii = 5 Then modCP(92) = ChangeCustomerL(RsTemp.Fields("CRA05") & RsTemp.Fields("CRA06"))
               RsTemp.MoveNext
               If RsTemp.EOF Then Exit For
            Next ii
         End If
      End If
      
      '¹q¤l°e¥ó
      If strCRL(95) = "Y" Then
         modCP(118) = "YY"
      End If
      '¦~¶O´Á¶¡
      If ((modCP(1) = "P" Or modCP(1) = "CFP") And (modCP(10) = "605" Or modCP(10) = "606" Or modCP(10) = "607")) Or _
         (modCP(1) = "P" And modCP(10) = "601") Or _
         (modCP(1) = "CFP" And modCP(10) = "613") Then
         modCP(53) = strCRL(144)
         modCP(54) = strCRL(145)
      End If
      
      '¯S®íºÞ¨î
      mType = "": mCaseNo = ""
      'Modified by Lydia 2024/03/14 ²Ä1µ§¤~³B²z¯S®íºÞ¨î + intCRC = 1
      If intCRC = 1 And strCRL(74) <> "" And strCRL(55) <> "" Then
         If strCRL(7) = "CFP" And strCRL(74) = "2" Then
            'If strCRL(6) = "Y" And modBase(2) = "" Then '·s®×¥¼¨ú±o®×¸¹®É,¤~­n§ì­È
            'Modified by Lydia 2023/01/10  debug:¼Ú·ù¬O613©µ®i¶O¡]­^°ê¡^¡A­^°ê©µ®i·s®×­n§ì607©µ®i
            'If strCRL(6) = "Y" And modCP(10) = "613" Then '·s®×©µ®i¶O¡]­^°ê¡^®É,¤~­n§ì­È
            'Modified by Lydia 2023/04/10 ¦P®É§PÂ_©e¥ô¥N²z¤H; ex. CFT-023552·s®×¬°©e¥ô¥N²z¤H
            'If strCRL(6) = "Y" And modCP(10) = "607" Then '·s®×©µ®i¶O¡]­^°ê¡^®É,¤~­n§ì­È
            If strCRL(6) = "Y" And (modCP(10) = "607" Or modCP(10) = "444") Then
               mType = "CFP­^°ê²æ¼Ú®×"
               mCaseNo = strCRL(55)
            End If
         ElseIf Len(strCRL(74)) = 2 Then
            mType = "LOS®×·½¦¬¤å"
            mCaseNo = strCRL(74) & "," & strCRL(55)
         End If
      End If
      
      'µo©ú¤H¸ê®Æ
      strInventorNo = ""
      'Modify By Sindy 2023/5/10 ²Ä¤@µ§¤~¶··s¼Wµo©ú¤H¸ê®Æ
      If intCRC = 1 Then
      '2023/5/10 END
         strExSql = "select * from consultrecinv where cri01='" & strCRL01 & "' order by cri02 asc"
         intP = 1
         Set rsPD = ClsLawReadRstMsg(intP, strExSql)
         If intP = 1 Then
            rsPD.MoveFirst
            Do While Not rsPD.EOF
               If "" & rsPD.Fields("cri03") <> "" And "" & rsPD.Fields("cri04") <> "" Then
                  strInventorNo = strInventorNo & rsPD.Fields("cri03") & rsPD.Fields("cri04") & ","
               End If
               rsPD.MoveNext
            Loop
            If Right(strInventorNo, 1) = "," Then strInventorNo = Mid(strInventorNo, 1, Len(strInventorNo) - 1)
         End If
      End If
      
      'Modify By Sindy 2024/11/21 + intCRC
      If PUB_SaveFrm010005("frm090801_New", IIf(strCRL(6) = "Y", IIf(modBase(2) <> "" And intCRC > 1, 0, 1), 0), 0, 0, _
         modBase, modCP, strCRA22, strInventorNo, mChkStr, IsSaveData, mType, mCaseNo, mRetVal, mTCTVal, mTCTList, intCRC) = False Then
         GoTo ErrHand
      Else
         '¦^¶Ç­È
         strPA01 = modBase(1)
         strPA02 = modBase(2)
         strPA03 = modBase(3)
         strPA04 = modBase(4)
         
         '±M§Q³B¹w³]©Ó¿ì¤H³W«h : ¦s¤J±µ¬¢³æ®×¥ó©Ê½èÀÉªº¹w¤À©Ó¿ì¤H
         strCRC09 = PUB_SetPxxCP14(modCP(1), modCP(2), modCP(3), modCP(4), modCP(10))
         strTmp1(10) = ""
         If modCP(14) <> "" Then '¦~¶O¦^¼gCRC09
            strTmp1(10) = ",CRC09='" & modCP(14) & "'"
         ElseIf strCRC09 <> "" Then
            strTmp1(10) = ",CRC09='" & strCRC09 & "'"
         End If
         '§ó·s±µ¬¢°O¿ý³æ®×¥ó©Ê½èªºÁ`¦¬¤å¸¹
         strExSql = "update ConsultRecCMP set CRC08='" & modCP(9) & "'" & strTmp1(10) & _
                  " where CRC01='" & strCRL01 & "'" & _
                  " and CRC02='" & rsAD.Fields("CRC02") & "' and CRC08 is null"
         cnnConnection.Execute strExSql, intA
         
         '§ó·s·s®×ªº®×¸¹
         If intCRC = rsAD.RecordCount Then
            strExSql = "update ConsultRecordList set CRL08='" & modBase(2) & "',CRL09='" & modBase(3) & "',CRL10='" & modBase(4) & "'" & _
                     " where CRL01='" & strCRL01 & "' and CRL06='Y' and CRL08 is null"
            cnnConnection.Execute strExSql, intA
         End If
      End If

      rsAD.MoveNext
   Loop
   End If

   PUB_AutoRecvCRL_P = True
   Set rsAD = Nothing
   Set rsPD = Nothing
   Exit Function
   
ErrHand:
   PUB_AutoRecvCRL_P = False
   Set rsAD = Nothing
   Set rsPD = Nothing
   If Err.Number <> 0 Then
      MsgBox Err.Description & vbCrLf & _
      ";strexsql = " & strExSql, vbCritical, "PUB_AutoRecvCRL_P"
   End If
End Function

'Add By Sindy 2023/1/9
'±M§Q³B¹w³]©Ó¿ì¤H³W«h
Public Function PUB_SetPxxCP14(strSys As String, strCP02 As String, strCP03 As String, strCP04 As String, _
   strCP10 As String) As String
Dim strExSql As String, intA As Integer, rsAD As New ADODB.Recordset 'Added by Lydia 2024/05/16

   If strSys <> "P" And strSys <> "PS" And strSys <> "CFP" And strSys <> "CPS" Then Exit Function
   
   PUB_SetPxxCP14 = ""
   
   '¯S©wªº®×¥ó©Ê½è¤~¹w¤À©Ó¿ì¤H
   '106 ¥D±i°ê»ÚÀu¥ýÅv
   '123 ¥D±iÀu´f´Á
   '416 ¹êÅé¼f¬d
   '920 «æ¥ó¶O
   '938 ¶W­¶¶O
   '939 ¶W¶µ¶O
   '944 ·|½Z­×§ï
   'PS.«á¦¬¤å¥D±i°ê¤ºÀu¥ýÅv«Ü©_©Ç , ¦]¦¹§Ú¶É¦V§¹¥þ¤â°Ê³B²z, ¦Ó¤£­n¦Û°Ê±a¤J
   If InStr("106,123,416,920,938,939,944", strCP10) > 0 Then
      '¦P¤@®×¸¹«áÄò¦¬¤åªº®×¥ó©Ê½è, ¥D°Ê¶ñ¤J»P·s®×¦P¤@¤uµ{®v
      '¥u¬O, ¥DºÞ¤À®×®É·|¦Û°Ê¶ñ¤J¤uµ{®v¦Ó¤w(¨Ã«D¦Û°Ê¤À®×), ÁÙ¬O­n¸g¥DºÞ¤À®×
      '§PÂ_¬OÂ÷Â¾¤H­û®É´N¤£±a¤J
      strExSql = "SELECT cp14 FROM caseprogress,staff" & _
               " where cp01='" & strSys & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
               " and cp31='Y' and cp157>0 and cp159=0 and cp14 is not null and cp14=st01(+) and st04='1'" & _
               " order by cp66 desc,cp67 desc"
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strExSql)
      If intA = 1 Then
         PUB_SetPxxCP14 = "" & rsAD.Fields("cp14")
         
         'Add By Sindy 2023/2/1 ¬Â¬Â´£¹êÅé¼f¬d·s®×¤wµo¤å,´N¤£»Ý­n¹w¤À©Ó¿ì¤H
         If InStr("416", strCP10) > 0 Then
            strExSql = "SELECT cp14 FROM caseprogress" & _
                     " where cp01='" & strSys & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                     " and cp31='Y' and cp158>0 and cp159=0" & _
                     " order by cp66 desc,cp67 desc"
            intA = 1
            Set rsAD = ClsLawReadRstMsg(intA, strExSql)
            If intA = 1 Then
               PUB_SetPxxCP14 = ""
            End If
         End If
         '2023/2/1 END
      End If
   End If
   Set rsAD = Nothing 'Added by Lydia 2024/05/16
End Function

'Add By Sindy 2022/9/30
'·s¼W ªA°È®×¡Bªk°È®×¡BACS®×«D112¤§¦¬¤å¡BÅU°Ý®×«D0¤§¦¬¤å
Public Function PUB_AutoRecvCRL_Other(strCRL01 As String, _
   Optional ByRef strSP01 As String, Optional ByRef strSP02 As String, Optional ByRef strSP03 As String, Optional ByRef strSP04 As String, _
   Optional ByVal bolNewCase As Boolean = False) As Boolean
Dim modCP() As String, modBase() As String '¦¬¤å ©M °ò¥»ÀÉ
Dim m_strControl As String '»ô³Æ¤éºÞ¨î
Dim mType As String, mCaseNo As String '¯S®íºÞ¨î
Dim strCRL() As String
Dim intKind As Integer, intWhere As Integer
Dim strTmpA As String
Dim intA As Integer
Dim rsAD As New ADODB.Recordset
Dim Rs As New ADODB.Recordset 'Add By Sindy 2024/5/27
Dim ii As Integer
Dim douStPrice As Double, douLowPrice As Double
Dim strAPP1 As String, strCRA10 As String, strCRA22 As String
Dim mChkStr As String, mRetVal As String
Dim IsSaveData As Boolean
Dim intCRC As Integer
Dim strCRC09 As String
Dim strExSql As String 'Added by Lydia 2024/05/16
Dim strSP46() As String 'Add By Sindy 2024/12/30

On Error GoTo ErrHand

   '±µ¬¢³æ¥DÀÉ
   ReDim Preserve strCRL(TF_CRL) As String
   strCRL(1) = strCRL01
   If ClsPDReadCRLDatabase(strCRL) = False Then
      GoTo ErrHand
   End If

   '¥Ó½Ð¤H1
   strExSql = "select * from ConsultRecApp" & _
            " where CRA01='" & strCRL01 & "' and CRA02=1"
   intA = 1
   Set Rs = ClsLawReadRstMsg(intA, strExSql)
   If intA = 1 Then
      strAPP1 = Rs.Fields("CRA05") & Rs.Fields("CRA06")
      strCRA10 = "" & Rs.Fields("CRA10") '±µ¬¢¤H
      strCRA22 = "" & Rs.Fields("CRA22") 'Ápµ¸¦a§}¶l»¼°Ï¸¹
   End If

'*********************
   '³]©w°}¦C
'*********************
   If ClsPDGetSystemKind(strCRL(7), intKind) = True Then
     Select Case intKind
        Case ±M§Q
           ReDim Preserve modBase(TF_PA) As String
        Case °Ó¼Ð
           ReDim Preserve modBase(TF_TM) As String
        Case ªk°È
           ReDim Preserve modBase(TF_LC) As String
        Case ÅU°Ý
           ReDim Preserve modBase(TF_HC) As String
        Case Else
           ReDim Preserve modBase(tf_SP) As String
     End Select
   End If
   ReDim Preserve modCP(TF_CP) As String
   
   'Add By Sindy 2023/8/16
   If bolNewCase = True Then '«ü©w¬°·s®×³B²z(¥À®×¤§´X)
      strCRL(6) = "Y"
      modBase(1) = strSP01
      modBase(2) = strSP02
      modBase(3) = strSP03
      modBase(4) = strSP04
   Else
   '2023/8/16 END
      '¨ú±o®×¸¹
      modBase(1) = strCRL(7)
      'ÂÂ®× ©Î ("¤£¦P¼f¯Å"·|¬O·s®×¥À¸¹)
      If strCRL(8) <> "" Then modBase(2) = strCRL(8)
      If strCRL(9) = "" Then
         modBase(3) = "0"
      Else
         modBase(3) = Right("0" & strCRL(9), 1)
      End If
      If strCRL(10) = "" Then
         modBase(4) = "00"
      Else
         modBase(4) = Right("00" & strCRL(10), 2)
      End If
   End If
   modCP(1) = modBase(1)
   modCP(2) = modBase(2)
   modCP(3) = modBase(3)
   modCP(4) = modBase(4)
   
   If strCRL(6) = "" Then 'ÂÂ®×
      'Modified by Lydia 2023/05/11 +false
      If PUB_ReadCaseData(modBase, intKind, intWhere, False) = False Then
         GoTo ErrHand
      End If
      
      'Add By Sindy 2024/4/11
      Select Case intKind
           Case ªk°È
               modBase(23) = strCRL(77) '¥N²z¤H©¼©Ò®×¸¹
               
           Case ÅU°Ý
               
           Case Else  'ªA°È
               modBase(27) = strCRL(77) '¥N²z¤H©¼©Ò®×¸¹
      End Select
      '2024/4/11 END
      
   '·s®×
   Else 'strCRL(6) = "Y" Then
      '°ò¥»ÀÉ
      strTmp1(1) = "": strTmp1(2) = ""
      
      If ClsPDGetSystemKind(strCRL(7), intKind) = True Then
        Select Case intKind
           Case ªk°È
               modBase(5) = strCRL(17) '®×¥ó¦WºÙ(¤¤)
                
               strExSql = "select * from ConsultRecApp" & _
                        " where CRA01='" & strCRL01 & "'" & _
                        " order by CRA02 asc"
               intA = 1
               Set Rs = ClsLawReadRstMsg(intA, strExSql)
               If intA = 1 Then
                  Rs.MoveFirst
                  '·í¨Æ¤H1~5
                  For ii = 1 To 5
                     If ii = 1 Then modBase(11) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     If ii = 2 Then modBase(43) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     If ii = 3 Then modBase(44) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     If ii = 4 Then modBase(45) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     If ii = 5 Then modBase(46) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     Rs.MoveNext
                     If Rs.EOF Then Exit For
                  Next ii
               End If
               
               strTmp1(1) = "11"  '·í¨Æ¤H1
               strTmp1(2) = "42" 'Ápµ¸¤H½s¸¹
               modBase(15) = strCRL(15) '¥Ó½Ð°ê®a
               modBase(17) = strCRL(16) '«È¤á®×¥ó®×¸¹
               'FC¥N²z¤H
               If strCRL(60) <> "" Then
                  modBase(22) = ChangeCustomerL(strCRL(60) & strCRL(61))
               End If
               modBase(23) = strCRL(77) '¥N²z¤H©¼©Ò®×¸¹
               
           Case ÅU°Ý
               modBase(6) = strCRL(17) '®×¥ó¦WºÙ(¤¤)
               
               strExSql = "select * from ConsultRecApp" & _
                        " where CRA01='" & strCRL01 & "'" & _
                        " order by CRA02 asc"
               intA = 1
               Set Rs = ClsLawReadRstMsg(intA, strExSql)
               If intA = 1 Then
                  Rs.MoveFirst
                  '·í¨Æ¤H1~5
                  For ii = 1 To 5
                     If ii = 1 Then modBase(5) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     If ii = 2 Then modBase(24) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     If ii = 3 Then modBase(25) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     If ii = 4 Then modBase(26) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     If ii = 5 Then modBase(27) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     Rs.MoveNext
                     If Rs.EOF Then Exit For
                  Next ii
               End If
               
               strTmp1(1) = "5"  '·í¨Æ¤H1
               strTmp1(2) = "23" 'Ápµ¸¤H½s¸¹
               
           Case Else  'ªA°È
               modBase(5) = strCRL(17) '®×¥ó¦WºÙ(¤¤)
               
               strExSql = "select * from ConsultRecApp" & _
                        " where CRA01='" & strCRL01 & "'" & _
                        " order by CRA02 asc"
               intA = 1
               Set Rs = ClsLawReadRstMsg(intA, strExSql)
               If intA = 1 Then
                  Rs.MoveFirst
                  '¥Ó½Ð¤H1~5
                  For ii = 1 To 5
                     If ii = 1 Then modBase(8) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     If ii = 2 Then modBase(58) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     If ii = 3 Then modBase(59) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     If ii = 4 Then modBase(65) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     If ii = 5 Then modBase(66) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
                     Rs.MoveNext
                     If Rs.EOF Then Exit For
                  Next ii
               End If
               
               modBase(29) = strCRL(16) '«È¤á®×¥ó®×¸¹
               strTmp1(1) = "8"  '¥Ó½Ð¤H1
               strTmp1(2) = "78" 'Ápµ¸¤H½s¸¹
               modBase(9) = strCRL(15) '¥Ó½Ð°ê®a
               'FC¥N²z¤H
               If strCRL(60) <> "" Then
                  modBase(26) = ChangeCustomerL(strCRL(60) & strCRL(61))
               End If
               modBase(27) = strCRL(77) '¥N²z¤H©¼©Ò®×¸¹
               modBase(73) = strCRL(73) '°Ó«~Ãþ§O
               'Add By Sindy 2024/12/30 §@«~ºØÃþ
               If strCRL(94) = "1" Or strCRL(94) = "2" Then
                  strExSql = "select *" & _
                             " from ConsultRecCMP" & _
                             " where CRC01='" & strCRL01 & "'" & _
                             " order by CRC02 asc"
                  intA = 1
                  Set Rs = ClsLawReadRstMsg(intA, strExSql)
                  If intA = 1 Then
                     Call Pub_GetCF209(modBase(1), modBase(9), Rs.Fields("CRC03"), strSP46)
                     modBase(46) = strSP46(strCRL(94) - 1)
                  End If
               End If
               '2024/12/30 END
               'Ápµ¸¤H1~2 =>  frm010007_1.bolOK
'               If bolCancel = True Then
'                  modBase(30) = strSP30s
'                  modBase(75) = strSP75s
'               End If
        End Select
      End If

      '¥Ó½Ð¤HÁpµ¸¤H½s¸¹
      If strCRA10 <> "" Then
         strExSql = "select * from potcustcont where pcc01='" & Left(strAPP1, 8) & "' and pcc05='" & strCRA10 & "'"
         intA = 1
         Set Rs = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            modBase(Val(strTmp1(2))) = "" & Rs.Fields("PCC02")
            '­Y­Ó®×±µ¬¢¤H»P«È¤áÀÉªº¹w³]±µ¬¢¤H¬Û¦P®É¤£¥²³]©w
            PUB_GetContact strAPP1, strTmpA, True
            If modBase(Val(strTmp1(2))) = strTmpA Then
               modBase(Val(strTmp1(2))) = ""
            End If
         End If
      End If
   End If
   
   '±µ¬¢°O¿ý³æ®×¥ó©Ê½è
   'Modify By Sindy 2022/10/26 ¨q¬ÂÄ±±o¥ý¨Ì´¼Åv¤H­û¿é¤Jªº¶¶§Ç§Y¥i
   strExSql = "select *" & _
            " from ConsultRecCMP" & _
            " where CRC01='" & strCRL01 & "'" & _
            " order by CRC02 asc"
            '" order by CRC03 asc"
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strExSql)
   If intA = 1 Then
   rsAD.MoveFirst
   intCRC = 0
   Do While Not rsAD.EOF
      intCRC = intCRC + 1
      If intCRC > 1 Then '²Ä2µ§¦¬¤å­n²MªÅmodCP³¯¦C­È
         Erase modCP
         ReDim Preserve modCP(TF_CP) As String
         modCP(1) = modBase(1)
         modCP(2) = modBase(2)
         modCP(3) = modBase(3)
         modCP(4) = modBase(4)
      End If
      
      '¦¬¤åCaseProgress
      IsSaveData = False '*****
      modCP(9) = "A" & CompAutoNumberYear(GetTaiwanThisYear) '¦¬¤å¸¹ ex:AB1
      modCP(5) = strSrvDate(1) '¦¬¤å¤é
      modCP(10) = rsAD.Fields("CRC03") '®×¥ó©Ê½è
      modCP(6) = strCRL(12) '¥»©Ò´Á­­
      modCP(7) = strCRL(13) 'ªk©w´Á­­
      modCP(11) = "07" '®×¥ó¨Ó·½
      modCP(12) = GetST15(strCRL(3))
      modCP(13) = strCRL(3) '´¼Åv¤H­û
      
      '¤º°Ó¹w³]©Ó¿ì¤H³W«h
      'If intKind <> ±M§Q And intKind <> ªk°È And intKind <> ÅU°Ý Then
      If intKind = 6 And Left(modCP(12), 1) <> "F" Then 'ªA°È°Ó¼Ð
         modCP(14) = PUB_SetTxxCP14(modCP(1), modCP(2), modCP(3), modCP(4), strAPP1, modCP(13), modCP(10), modBase(9), strCRL01, strCRC09)
      End If
         
      modCP(16) = rsAD.Fields("CRC04") '¶O¥Î
      modCP(17) = rsAD.Fields("CRC05") '³W¶O
      modCP(18) = rsAD.Fields("CRC06") 'ÂI¼Æ
      modCP(19) = strCRL(39) '«áª÷
      'Modify By Sindy 2025/8/26 +, , strCRL(60) & strCRL(61)
      If ClsPDGetCaseLowPrice(strCRL(7), strCRL(15), modCP(10), douStPrice, douLowPrice, strCRL(87), strCRL(81), _
         strCRL(1), strCRL(5), strCRL(7), strCRL(8), strCRL(9), strCRL(10), , strCRL(60) & strCRL(61)) = 1 Then
         modCP(33) = douStPrice '¼Ð·Ç»ù
         modCP(34) = douLowPrice '©³»ù
      End If
      'Modify By Sindy 2023/5/18 Txªº1=¤£­­¨î,¤£¼g¤J¶i«×ÀÉ
      'Modify By Sindy 2024/3/18 ¤£¥Î§PÂ_, mark if
      'If Left(modCP(1), 1) = "T" And strCRL(82) <> "1" Then
         modCP(141) = strCRL(82) '°e¥ó¤è¦¡
      'End If
      '2023/5/18 END
      modCP(142) = strCRL(83) '«ü©w°e¥ó¤é
      modCP(164) = strCRL(155) '«ü©w°e¥ó¤è¦¡ Add By Sindy 2023/12/12
      modCP(151) = strCRL(92) '¦¬¾Ú¦Û°Ê¦C¦L®É¶¡ÂI Add By Sindy 2023/4/18
      
      'ÂÂ®×®É¶i«×³Æµù=¥DÃD
      If strCRL(6) = "" Then 'ÂÂ®×
         If (strCRL(7) = "L" And modBase(5) <> strCRL(17)) Or _
            (strCRL(7) = "LA" And modBase(6) <> strCRL(17)) Then
            modCP(64) = strCRL(17)
         End If
      End If
      If modCP(64) = "" Then
         modCP(64) = "" & rsAD.Fields("CRC07") '³Æµù
      End If
      
      modCP(140) = strCRL01 '±µ¬¢³æ½s¸¹
      
      'Add By Sindy 2025/3/26 ªA°È°Ó¼Ð
      'TD-Âà§}
      'CFC(¥~°ÓµÛ§@Åv)-²¾Âà;TB(±ø½X)-ÂàÅý;TC(¤º°ÓµÛ§@Åv)-Åý»P
      If intKind = 6 And modCP(1) <> "TD" And rsAD.Fields("CRC03") = ²¾Âà Then
         'Åý»P¤H1-5,¨üÅý¤H1-5
         strExSql = "select * from ConsultRecApp" & _
                  " where CRA01='" & strCRL01 & "'" & _
                  " order by CRA02 asc"
         intA = 1
         Set Rs = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            Rs.MoveFirst
            '¥Ó½Ð¤H1~5
            For ii = 1 To 5
               If ii = 1 Then modCP(56) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               If ii = 2 Then modCP(89) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               If ii = 3 Then modCP(90) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               If ii = 4 Then modCP(91) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               If ii = 5 Then modCP(92) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
               Rs.MoveNext
               If Rs.EOF Then Exit For
            Next ii
         End If
      End If
      '2025/3/26 END
      
      '¯S®íºÞ¨î
      mType = "": mCaseNo = ""
      'Modified by Lydia 2024/03/14 ²Ä1µ§¤~³B²z¯S®íºÞ¨î + intCRC = 1
      If intCRC = 1 And strCRL(74) <> "" And strCRL(55) <> "" Then
         If Len(strCRL(74)) = 2 Then
            mType = "LOS®×·½¦¬¤å"
            mCaseNo = strCRL(74) & "," & strCRL(55)
         ElseIf strCRL(74) = "1" Then
            If strCRL(7) = "TS" And InStr(TMQ_TS®×, modCP(10)) > 0 Then
               mType = "T¬d¦W³æ"
               mCaseNo = strCRL(55)
               'Added Lydia 2023/01/16 ¨Ï¥ÎPUB_TQCtoTMQ¨ú±o¹ê»Úªº¬d¦W³æ¸¹,²Õ¦¨¥i¥H¨Ï¥ÎPUB_TMQtoCPªº¸ê®Æ=±µ¬¢³æ¬d¦W¥N¸¹|+¹ê»Ú¬d¦W³æ¸¹,ex.111002476|+HB1110046,
               If mCaseNo <> "" Then
                   'Modified by Lydia 2024/03/14 +False
                   Call PUB_TQCtoTMQ(False, modCP(12), modCP(13), mCaseNo, strTmpA)
                   mCaseNo = mCaseNo & "|" & strTmpA
               End If
               'end 2023/01/16
            End If
         End If
      End If
      
'      '¶Ç¤J¨ä¥L¾Þ§@µ²ªG
'      mChkStr = ","
'      If Trim(txtSystem) = "PS" And (mCP31 = "Y" Or frm010001.intSaveMode = 1) And m_LOS15 = "" And mFMPchk = True Then
'          mChkStr = mChkStr & "¾ÈµØ®×¥ó½T»{,"
'      End If

      '»ô³Æ¤é  --m_strControl
      m_strControl = ""
      '¸ê®Æ¬O§_»ô³Æ
      If strCRL(88) <> "" Then
          m_strControl = m_strControl & ",EP06|" & strCRL(88) & "|" & ""
      End If
      '¬O§_·|½Z
      If strCRL(89) <> "" Then
          m_strControl = m_strControl & ",EP34|" & strCRL(89)
      End If
      If m_strControl <> "" Then m_strControl = Mid(m_strControl, 2)

'*********************
      '¦sÀÉ
'*********************
      '"¤£¦P¼f¯Å"·|¬O·s®×¥À¸¹
      'Modify By Sindy 2024/11/21 + intCRC
      If PUB_SaveFrm010007("frm090801_New", IIf(strCRL(6) = "Y", IIf(modBase(2) <> "" And intCRC > 1, 0, 1), 0), 0, intKind, 0, _
         modBase, modCP, strCRA22, mChkStr, m_strControl, IsSaveData, mType, mCaseNo, mRetVal, intCRC) = False Then
         GoTo ErrHand
      Else
         'Modify By Sindy 2025/8/21
         If modCP(14) = "" And strCRC09 = "" Then
         '2025/8/21 END
            '±M§Q³B¹w³]©Ó¿ì¤H³W«h : ¦s¤J±µ¬¢³æ®×¥ó©Ê½èÀÉªº¹w¤À©Ó¿ì¤H
            strCRC09 = PUB_SetPxxCP14(modCP(1), modCP(2), modCP(3), modCP(4), modCP(10))
         End If
         strTmp1(10) = ""
         If modCP(14) <> "" Then '¤w¦³©Ó¿ì¤H®É,¦^¼gCRC09
            strTmp1(10) = ",CRC09='" & modCP(14) & "'"
         ElseIf strCRC09 <> "" Then
            strTmp1(10) = ",CRC09='" & strCRC09 & "'"
         End If
         '§ó·s±µ¬¢°O¿ý³æ®×¥ó©Ê½èªºÁ`¦¬¤å¸¹
         strExSql = "update ConsultRecCMP set CRC08='" & modCP(9) & "'" & strTmp1(10) & _
                  " where CRC01='" & strCRL01 & "'" & _
                  " and CRC02='" & rsAD.Fields("CRC02") & "' and CRC08 is null"
         cnnConnection.Execute strExSql
         
         '§ó·s·s®×ªº®×¸¹
         If intCRC = rsAD.RecordCount Then
            strExSql = "update ConsultRecordList set CRL08='" & modBase(2) & "',CRL09='" & modBase(3) & "',CRL10='" & modBase(4) & "'" & _
                     " where CRL01='" & strCRL01 & "' and CRL06='Y' and CRL08 is null"
            cnnConnection.Execute strExSql
         End If
      End If

      rsAD.MoveNext
   Loop
   End If

   PUB_AutoRecvCRL_Other = True
   Set rsAD = Nothing
   Set Rs = Nothing 'Add By Sindy 2024/5/27
   Exit Function
   
ErrHand:
   PUB_AutoRecvCRL_Other = False
   Set rsAD = Nothing
   If Err.Number <> 0 Then
      MsgBox Err.Description & vbCrLf & _
      ";strexsql = " & strExSql, vbCritical, "PUB_AutoRecvCRL_Other"
   End If
End Function

'Add By Sindy 2022/9/30
'·s¼W ÅU°Ý®×¤§0ÅU°Ý¸u¥ô¦¬¤å
Public Function PUB_AutoRecvCRL_HireCase(strCRL01 As String) As Boolean
Dim modCP() As String, modBase() As String '¦¬¤å ©M °ò¥»ÀÉ
Dim mType As String, mCaseNo As String '¯S®íºÞ¨î
Dim strCRL() As String
Dim intKind As Integer, intWhere As Integer
Dim strTmpA As String
Dim intA As Integer
Dim rsAD As New ADODB.Recordset
Dim Rs As New ADODB.Recordset 'Add By Sindy 2024/5/27
Dim ii As Integer
Dim strAPP1 As String, strCRA10 As String, strCRA22 As String
Dim IsSaveData As Boolean
Dim intCRC As Integer
Dim strExSql As String 'Added by Lydia 2024/05/16

On Error GoTo ErrHand

   '±µ¬¢³æ¥DÀÉ
   ReDim Preserve strCRL(TF_CRL) As String
   strCRL(1) = strCRL01
   If ClsPDReadCRLDatabase(strCRL) = False Then
      GoTo ErrHand
   End If

   '¥Ó½Ð¤H1
   strExSql = "select * from ConsultRecApp" & _
            " where CRA01='" & strCRL01 & "' and CRA02=1"
   intA = 1
   Set Rs = ClsLawReadRstMsg(intA, strExSql)
   If intA = 1 Then
      strAPP1 = Rs.Fields("CRA05") & Rs.Fields("CRA06")
      strCRA10 = "" & Rs.Fields("CRA10") '±µ¬¢¤H
      strCRA22 = "" & Rs.Fields("CRA22") 'Ápµ¸¦a§}¶l»¼°Ï¸¹
   End If

'*********************
   '³]©w°}¦C
'*********************
   If ClsPDGetSystemKind(strCRL(7), intKind) = True Then
     Select Case intKind
        Case ±M§Q
           ReDim Preserve modBase(TF_PA) As String
        Case °Ó¼Ð
           ReDim Preserve modBase(TF_TM) As String
        Case ªk°È
           ReDim Preserve modBase(TF_LC) As String
        Case ÅU°Ý
           ReDim Preserve modBase(TF_HC) As String
        Case Else
           ReDim Preserve modBase(tf_SP) As String
     End Select
   End If
   ReDim Preserve modCP(TF_CP) As String
   
   '¨ú±o®×¸¹
   modBase(1) = strCRL(7)
   'ÂÂ®× ©Î ¤£¦P¼f¯Å,·|¬O·s®×¥À¸¹
   If strCRL(8) <> "" Then modBase(2) = strCRL(8)
   If strCRL(9) = "" Then
      modBase(3) = "0"
   Else
      modBase(3) = Right("0" & strCRL(9), 1)
   End If
   If strCRL(10) = "" Then
      modBase(4) = "00"
   Else
      modBase(4) = Right("00" & strCRL(10), 2)
   End If
   modCP(1) = modBase(1)
   modCP(2) = modBase(2)
   modCP(3) = modBase(3)
   modCP(4) = modBase(4)
   
   If strCRL(6) = "" Then 'ÂÂ®×
      'Modified by Lydia 2023/05/11 +false
      If PUB_ReadCaseData(modBase, intKind, intWhere, False) = False Then
         GoTo ErrHand
      End If
   
   '·s®×
   Else 'strCRL(6) = "Y" Then
      '°ò¥»ÀÉ
      modBase(6) = strCRL(17) '®×¥ó¦WºÙ(¤¤)
      strExSql = "select * from ConsultRecApp" & _
               " where CRA01='" & strCRL01 & "'" & _
               " order by CRA02 asc"
      intA = 1
      Set Rs = ClsLawReadRstMsg(intA, strExSql)
      If intA = 1 Then
         Rs.MoveFirst
         '·í¨Æ¤H1~5
         For ii = 1 To 5
            If ii = 1 Then modBase(5) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            If ii = 2 Then modBase(24) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            If ii = 3 Then modBase(25) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            If ii = 4 Then modBase(26) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            If ii = 5 Then modBase(27) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            Rs.MoveNext
            If Rs.EOF Then Exit For
         Next ii
      End If
      
      '¥Ó½Ð¤HÁpµ¸¤H½s¸¹
      If strCRA10 <> "" Then
         strExSql = "select * from potcustcont where pcc01='" & Left(strAPP1, 8) & "' and pcc05='" & strCRA10 & "'"
         intA = 1
         Set Rs = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            modBase(23) = "" & Rs.Fields("PCC02")
            '­Y­Ó®×±µ¬¢¤H»P«È¤áÀÉªº¹w³]±µ¬¢¤H¬Û¦P®É¤£¥²³]©w
            PUB_GetContact strAPP1, strTmpA, True
            If modBase(23) = strTmpA Then
               modBase(23) = ""
            End If
         End If
      End If
   End If
   
   '±µ¬¢°O¿ý³æ®×¥ó©Ê½è
   'Modify By Sindy 2022/10/26 ¨q¬ÂÄ±±o¥ý¨Ì´¼Åv¤H­û¿é¤Jªº¶¶§Ç§Y¥i
   strExSql = "select *" & _
            " from ConsultRecCMP" & _
            " where CRC01='" & strCRL01 & "'" & _
            " order by CRC02 asc"
            '" order by CRC03 asc"
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strExSql)
   If intA = 1 Then
   rsAD.MoveFirst
   intCRC = 0
   Do While Not rsAD.EOF
      intCRC = intCRC + 1
      If intCRC > 1 Then '²Ä2µ§¦¬¤å­n²MªÅmodCP³¯¦C­È
         Erase modCP
         ReDim Preserve modCP(TF_CP) As String
         modCP(1) = modBase(1)
         modCP(2) = modBase(2)
         modCP(3) = modBase(3)
         modCP(4) = modBase(4)
      End If
      
      '¦¬¤åCaseProgress
      IsSaveData = False '*****
      modCP(9) = "A" & CompAutoNumberYear(GetTaiwanThisYear) '¦¬¤å¸¹ ex:AB1
      modCP(5) = strSrvDate(1) '¦¬¤å¤é
      modCP(10) = rsAD.Fields("CRC03") '®×¥ó©Ê½è
      modCP(11) = "07" '®×¥ó¨Ó·½
      modCP(12) = GetST15(strCRL(3))
      modCP(13) = strCRL(3) '´¼Åv¤H­û
      'modCP(14) = Trim(txtAdviser(11))    '©Ó¿ì¤H
      modCP(16) = rsAD.Fields("CRC04") '¶O¥Î
      'modCP(18) = rsAD.Fields("CRC06") 'ÂI¼Æ ---¦sÀÉ®É­pºâ
      modCP(140) = strCRL01 '±µ¬¢³æ½s¸¹
      
      '¸u¥ô´Á¶¡
      modCP(53) = strCRL(144)
      modCP(54) = strCRL(145)

      '¯S®íºÞ¨î
      mType = "": mCaseNo = ""
      'Modified by Lydia 2024/03/14 ²Ä1µ§¤~³B²z¯S®íºÞ¨î + intCRC = 1
      If intCRC = 1 And Len(strCRL(74)) = 2 And strCRL(55) <> "" Then
         mType = "LOS®×·½¦¬¤å"
         mCaseNo = strCRL(74) & "," & strCRL(55)
      End If
      
'*********************
      '¦sÀÉ
'*********************
      'Modify By Sindy 2024/11/21 + intCRC
      If PUB_SaveFrm010006("frm090801_New", IIf(strCRL(6) = "Y", IIf(modBase(2) <> "" And intCRC > 1, 0, 1), 0), 0, 0, _
         modBase, modCP, strCRA22, IsSaveData, mType, mCaseNo, intCRC) = False Then
         GoTo ErrHand
      Else
         '§ó·s±µ¬¢°O¿ý³æ®×¥ó©Ê½èªºÁ`¦¬¤å¸¹
         strExSql = "update ConsultRecCMP set CRC08='" & modCP(9) & "'" & IIf(modCP(14) = "", "", ",CRC09='" & modCP(14) & "'") & _
                  " where CRC01='" & strCRL01 & "'" & _
                  " and CRC02='" & rsAD.Fields("CRC02") & "' and CRC08 is null"
         cnnConnection.Execute strExSql
         
         '§ó·s·s®×ªº®×¸¹
         If intCRC = rsAD.RecordCount Then
            strExSql = "update ConsultRecordList set CRL08='" & modBase(2) & "',CRL09='" & modBase(3) & "',CRL10='" & modBase(4) & "'" & _
                     " where CRL01='" & strCRL01 & "' and CRL06='Y' and CRL08 is null"
            cnnConnection.Execute strExSql
         End If
      End If

      rsAD.MoveNext
   Loop
   End If
   
   PUB_AutoRecvCRL_HireCase = True
   Set rsAD = Nothing
   Set Rs = Nothing 'Add By Sindy 2024/5/27
   Exit Function
   
ErrHand:
   PUB_AutoRecvCRL_HireCase = False
   Set rsAD = Nothing
   If Err.Number <> 0 Then
      MsgBox Err.Description & vbCrLf & _
      ";strexsql = " & strExSql, vbCritical, "PUB_AutoRecvCRL_HireCase"
   End If
End Function

'Add By Sindy 2022/9/30
'·s¼W ACS®×¤§112´¼°]ÅU°Ý¦¬¤å
Public Function PUB_AutoRecvCRL_Case(strCRL01 As String) As Boolean
Dim modCP() As String, modBase() As String '¦¬¤å ©M °ò¥»ÀÉ
Dim mType As String, mCaseNo As String '¯S®íºÞ¨î
Dim strCRL() As String
Dim intKind As Integer, intWhere As Integer
Dim strTmpA As String
Dim intA As Integer
Dim rsAD As New ADODB.Recordset
Dim Rs As New ADODB.Recordset 'Add By Sindy 2024/5/27
Dim ii As Integer
Dim strAPP1 As String, strCRA10 As String, strCRA22 As String
Dim IsSaveData As Boolean
Dim intCRC As Integer
Dim strExSql As String 'Added by Lydia 2024/05/16

On Error GoTo ErrHand

   '±µ¬¢³æ¥DÀÉ
   ReDim Preserve strCRL(TF_CRL) As String
   strCRL(1) = strCRL01
   If ClsPDReadCRLDatabase(strCRL) = False Then
      GoTo ErrHand
   End If

   '¥Ó½Ð¤H1
   strExSql = "select * from ConsultRecApp" & _
            " where CRA01='" & strCRL01 & "' and CRA02=1"
   intA = 1
   Set Rs = ClsLawReadRstMsg(intA, strExSql)
   If intA = 1 Then
      strAPP1 = Rs.Fields("CRA05") & Rs.Fields("CRA06")
      strCRA10 = "" & Rs.Fields("CRA10") '±µ¬¢¤H
      strCRA22 = "" & Rs.Fields("CRA22") 'Ápµ¸¦a§}¶l»¼°Ï¸¹
   End If

'*********************
   '³]©w°}¦C
'*********************
   If ClsPDGetSystemKind(strCRL(7), intKind) = True Then
     Select Case intKind
        Case ±M§Q
           ReDim Preserve modBase(TF_PA) As String
        Case °Ó¼Ð
           ReDim Preserve modBase(TF_TM) As String
        Case ªk°È
           ReDim Preserve modBase(TF_LC) As String
        Case ÅU°Ý
           ReDim Preserve modBase(TF_HC) As String
        Case Else
           ReDim Preserve modBase(tf_SP) As String
     End Select
   End If
   ReDim Preserve modCP(TF_CP) As String
   
   '¨ú±o®×¸¹
   modBase(1) = strCRL(7)
   'ÂÂ®× ©Î ¤£¦P¼f¯Å,·|¬O·s®×¥À¸¹
   If strCRL(8) <> "" Then modBase(2) = strCRL(8)
   If strCRL(9) = "" Then
      modBase(3) = "0"
   Else
      modBase(3) = Right("0" & strCRL(9), 1)
   End If
   If strCRL(10) = "" Then
      modBase(4) = "00"
   Else
      modBase(4) = Right("00" & strCRL(10), 2)
   End If
   modCP(1) = modBase(1)
   modCP(2) = modBase(2)
   modCP(3) = modBase(3)
   modCP(4) = modBase(4)
   
   If strCRL(6) = "" Then 'ÂÂ®×
      'Modified by Lydia 2023/05/11 +false
      If PUB_ReadCaseData(modBase, intKind, intWhere, False) = False Then
         GoTo ErrHand
      End If
      
   '·s®×
   Else 'strCRL(6) = "Y" Then
      '°ò¥»ÀÉ
      modBase(5) = strCRL(17) '®×¥ó¦WºÙ(¤¤)
      strExSql = "select * from ConsultRecApp" & _
               " where CRA01='" & strCRL01 & "'" & _
               " order by CRA02 asc"
      intA = 1
      Set Rs = ClsLawReadRstMsg(intA, strExSql)
      If intA = 1 Then
         Rs.MoveFirst
         '·í¨Æ¤H1~5
         For ii = 1 To 5
            If ii = 1 Then modBase(11) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            If ii = 2 Then modBase(43) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            If ii = 3 Then modBase(44) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            If ii = 4 Then modBase(45) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            If ii = 5 Then modBase(46) = ChangeCustomerL(Rs.Fields("CRA05") & Rs.Fields("CRA06"))
            Rs.MoveNext
            If Rs.EOF Then Exit For
         Next ii
      End If
      
      '¥Ó½Ð¤HÁpµ¸¤H½s¸¹
      If strCRA10 <> "" Then
         strExSql = "select * from potcustcont where pcc01='" & Left(strAPP1, 8) & "' and pcc05='" & strCRA10 & "'"
         intA = 1
         Set Rs = ClsLawReadRstMsg(intA, strExSql)
         If intA = 1 Then
            modBase(42) = "" & Rs.Fields("PCC02")
            '­Y­Ó®×±µ¬¢¤H»P«È¤áÀÉªº¹w³]±µ¬¢¤H¬Û¦P®É¤£¥²³]©w
            PUB_GetContact strAPP1, strTmpA, True
            If modBase(42) = strTmpA Then
               modBase(42) = ""
            End If
         End If
      End If
   End If
   
   '±µ¬¢°O¿ý³æ®×¥ó©Ê½è
   'Modify By Sindy 2022/10/26 ¨q¬ÂÄ±±o¥ý¨Ì´¼Åv¤H­û¿é¤Jªº¶¶§Ç§Y¥i
   strExSql = "select *" & _
            " from ConsultRecCMP" & _
            " where CRC01='" & strCRL01 & "'" & _
            " order by CRC02 asc"
            '" order by CRC03 asc"
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strExSql)
   If intA = 1 Then
   rsAD.MoveFirst
   intCRC = 0
   Do While Not rsAD.EOF
      intCRC = intCRC + 1
      If intCRC > 1 Then '²Ä2µ§¦¬¤å­n²MªÅmodCP³¯¦C­È
         Erase modCP
         ReDim Preserve modCP(TF_CP) As String
         modCP(1) = modBase(1)
         modCP(2) = modBase(2)
         modCP(3) = modBase(3)
         modCP(4) = modBase(4)
      End If
      
      '¦¬¤åCaseProgress
      IsSaveData = False '*****
      modCP(9) = "A" & CompAutoNumberYear(GetTaiwanThisYear) '¦¬¤å¸¹ ex:AB1
      modCP(5) = strSrvDate(1) '¦¬¤å¤é
      modCP(10) = rsAD.Fields("CRC03") '®×¥ó©Ê½è
      modCP(11) = "07" '®×¥ó¨Ó·½
      modCP(12) = GetST15(strCRL(3))
      modCP(13) = strCRL(3) '´¼Åv¤H­û
      'modCP(14) = Trim(txtOther(20))    '©Ó¿ì¤H
      modCP(16) = rsAD.Fields("CRC04") '¶O¥Î
      modCP(17) = rsAD.Fields("CRC05") '³W¶O
      modCP(18) = rsAD.Fields("CRC06") 'ÂI¼Æ
      modCP(140) = strCRL01 '±µ¬¢³æ½s¸¹
      modCP(151) = strCRL(92) '¦¬¾Ú¦Û°Ê¦C¦L®É¶¡ÂI Add By Sindy 2023/4/18
      
      '¸u¥ô´Á¶¡
      modCP(53) = strCRL(144)
      modCP(54) = strCRL(145)

      '¯S®íºÞ¨î
      mType = "": mCaseNo = ""
      
'*********************
      '¦sÀÉ
'*********************
      'Modify By Sindy 2024/11/21 + intCRC
      If PUB_SaveFrm010006_1("frm090801_New", IIf(strCRL(6) = "Y", IIf(modBase(2) <> "" And intCRC > 1, 0, 1), 0), 0, intKind, 0, _
         modBase, modCP, strCRA22, IsSaveData, mType, mCaseNo, intCRC) = False Then
         GoTo ErrHand
      Else
         '§ó·s±µ¬¢°O¿ý³æ®×¥ó©Ê½èªºÁ`¦¬¤å¸¹
         strExSql = "update ConsultRecCMP set CRC08='" & modCP(9) & "'" & IIf(modCP(14) = "", "", ",CRC09='" & modCP(14) & "'") & _
                  " where CRC01='" & strCRL01 & "'" & _
                  " and CRC02='" & rsAD.Fields("CRC02") & "' and CRC08 is null"
         cnnConnection.Execute strExSql
         
         '§ó·s·s®×ªº®×¸¹
         If intCRC = rsAD.RecordCount Then
            strExSql = "update ConsultRecordList set CRL08='" & modBase(2) & "',CRL09='" & modBase(3) & "',CRL10='" & modBase(4) & "'" & _
                     " where CRL01='" & strCRL01 & "' and CRL06='Y' and CRL08 is null"
            cnnConnection.Execute strExSql
         End If
         
         'Added by Lydia 2023/10/06 ´¼°]ÅU°Ý¤§±M·~®É¼Æ½Õ¾ã¡G´¼Åv¤H­û¦¬¤å´¼°]ÅU°Ý®É¡A¨t²Î¦Û°Ê¶Ç°e«H¨ç¦Ü§ù¿P¤å¤§«H½c¡C
         'Modified by Lydia 2023/11/28 §ï¦bfrm090801_12
         'If intCRC = 1 Then
         '   strTmp1(0) = Pub_GetSpecMan("¥þ©Ò´¼Åv³¡¥DºÞ")
         '   If strTmp1(0) <> "" Then
         '      mstrexsql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc13)" & _
         '                " values ('" & strUserNum & "','" & strTmp1(0) & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss')" & _
         '                ",'¦¬¤å´¼°]ÅU°Ý®×" & modBase(1) & "-" & modBase(2) & IIf(modBase(3) & modBase(4) <> "000", "-" & modBase(3) & "-" & modBase(4), "") & "','" & modCP(9) & "')"
         '      cnnConnection.Execute mstrexsql
         '   End If
         'End If
         'end 2023/10/06
      End If

      rsAD.MoveNext
   Loop
   End If
   
   PUB_AutoRecvCRL_Case = True
   Set rsAD = Nothing
   Set Rs = Nothing 'Add By Sindy 2024/5/27
   Exit Function
   
ErrHand:
   PUB_AutoRecvCRL_Case = False
   Set rsAD = Nothing
   If Err.Number <> 0 Then
      MsgBox Err.Description & vbCrLf & _
      ";strexsql = " & strExSql, vbCritical, "PUB_AutoRecvCRL_Case"
   End If
End Function

'Added by Lydia 2022/11/03 ªk°È®×¦¬¤å¨ú®ø"¬O§_ªLÁ`¦P·N¥»®×¥Ñªk«ß©Ò¦Û¦æ¦¬¤å¡H"ªº¸ß°Ý¡A§ï¬°ªk«ß©Ò¦Û¦æ¦¬¤å´¼¼z©Ò¤H­û«È¤á®É¡A¦b§PÂ_¸ó°Ï¦¬¤åµoEMAILµ¹Âù¤è´¼Åv¤H­û®É¡A¥[µoªLÁ`¡C
Public Function PUB_ChkForLawMan(ByVal PCU01 As String, ByVal pCP01 As String, pCP02 As String, ByVal pCP03 As String, pCP04 As String) As String
Dim rsQuery As New ADODB.Recordset
Dim intQ As Integer, strQ1 As String, strQ2 As String

    PUB_ChkForLawMan = ""
    'Added by Lydia 2020/08/31 ±Æ°£¥x¤@
    If PCU01 <> "" And InStr(PCU01, "X03072") > 0 Then
    Else
    'end 2020/08/31
       'Modified by Lydia 2020/08/03 ±Æ°£LA999999
       'Modified by Lydia 2020/11/05 ªk°È³BP31
       strQ1 = GetCuSales(ChangeCustomerL(PCU01))
       If Trim(PCU01) <> "" And Left(strQ1, 1) <> "L" And strQ1 <> "P31" And pCP01 & pCP02 <> "LA999999" Then
       'end 2020/11/05
           'Added by Lydia 2021/01/08 ¥Ñªk°È¤H­ûª½±µ¦¬A2Ãþ®×·½ªº§PÂ_(¦b¦sÀÉ®É·|¦Û°Ê¸É¤W®×·½)
           strTmp1(0) = ""
           If pCP02 <> "" Then
               mStrSql = "Select * From Lawofficesource Where los02='A1' and los07||los08 is null and Los15 In " & _
                            "(select cp162 from caseprogress where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp162 is not null)  order by los12 desc "
               intQ = 1
               Set rsQuery = ClsLawReadRstMsg(intQ, mStrSql)
               If intQ = 1 Then
                   strQ2 = "A2"
               End If
           End If
           If strQ2 = "" Then
           'end 2021/01/08
             'Added by Morgan 2021/5/28 ¦³ªLÁ`¦P·Nªk«ß©Ò¦Û¦æ¦¬¤åªº¨Ò¥~ Ex:L-6396
             'Modified by Morgan 2022/5/19 °ê¥~³¡«È¤á°£¥~ Ex:FCL-010967 --¨q¬Â
             If Left(strQ1, 1) <> "F" Then
                 '§ï¬°ªk«ß©Ò¦Û¦æ¦¬¤å´¼¼z©Ò¤H­û«È¤á®É¡A¦b§PÂ_¸ó°Ï¦¬¤åµoEMAILµ¹Âù¤è´¼Åv¤H­û®É¡A¥[µoªLÁ`¡C
                 PUB_ChkForLawMan = ";94007"
             End If
             'end 2022/5/19
           End If 'Added by Lydia 2021/01/08
       End If
End If 'Added by Lydia 2020/08/31
End Function

'Modify By Sindy 2022/11/15 ±qfrm010001·h¨ì¦¹³B§ï¬°¦@¥Î
'2007/7/4 add by sonia ÀË¬d¬O§_¦P·N­«·s©e¥ô
'bolRecv: ¬O§_¦¬¤å§@·~
Public Function PUB_ChkAgree928(strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String, _
   Optional bolRecv As Boolean = True) As Boolean
Dim StrSQLa As String, StrSqlB As String
Dim rsA As New ADODB.Recordset, rsB As New ADODB.Recordset
Dim iCol As Integer

   PUB_ChkAgree928 = True
   StrSQLa = "Select PA26,PA27,PA28,PA29,PA30 From patent where pa01='" & strPA01 & "' and pa02='" & strPA02 & "' and pa03='" & IIf(strPA03 = "", "0", strPA03) & "' and pa04='" & IIf(strPA04 = "", "00", strPA04) & "' "
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      For iCol = 0 To 4
         If Not IsNull(rsA.Fields(iCol)) Then
            StrSqlB = "Select * From LinReasignRec Where LR01='" & rsA.Fields(iCol) & "' "
            rsB.CursorLocation = adUseClient
            rsB.Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
            If rsB.RecordCount > 0 Then
               If rsB("LR09").Value = "N" Then
                  MsgBox "¦¹®×¥ó«È¤á¤£¦P·N­«·s©e¥ô" & IIf(bolRecv = True, ", ½Ð°h¦^­ì´¼Åv¤H­û", "") & "!!!", vbExclamation + vbOKOnly
                  PUB_ChkAgree928 = False
                  If rsB.State <> adStateClosed Then rsB.Close
                  Set rsB = Nothing
                  Exit Function
               End If
            End If
            If rsB.State <> adStateClosed Then rsB.Close
            Set rsB = Nothing
         Else
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
            Exit Function
         End If
      Next
   Else
      MsgBox "µL¦¹®×¸¹°ò¥»¸ê®Æ !!!", vbExclamation + vbOKOnly
      PUB_ChkAgree928 = False
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function
'2007/7/4 end

'Modify By Sindy 2022/11/15 ±qfrm010001·h¨ì¦¹³B§ï¬°¦@¥Î
Public Function PUB_RecvCheck412(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String) As Boolean

   PUB_RecvCheck412 = False
   
   strSql = "Select PA09, PA14, CP27  From Patent, caseprogress Where " & ChgPatent(strCP01 & strCP02 & strCP03 & strCP04) & " and CP01(+)=PA01 AND CP02(+)=PA02 AND CP03(+)=PA03 AND CP04(+)=PA04 AND CP10(+)='601'"
   
On Error GoTo ErrHnd
   
   CheckOC
   With adoRecordset
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      '­Y¦³¸ê®Æ
      If .RecordCount > 0 Then
         If ("" & .Fields("PA09")) = "000" Then
            '¤w¤½§i
            If Val("" & .Fields("PA14")) > 0 Then
               MsgBox "¸Ó®×¤w©ó " & ChangeTStringToTDateString(.Fields("PA14") - 19110000) & " ¤½§i¡A¤£¥i¦¬©µ½w¤½§i¡I", vbCritical
               
            'Modify by Morgan 2004/12/20 ¨ú®ø±±¨î
'            '»âÃÒ¤wµo¤å
'            ElseIf Not IsNull(.Fields("CP27")) Then
'               MsgBox "¸Ó®×»âÃÒ¤w©ó " & ChangeTStringToTDateString(.Fields("CP27") - 19110000) & " µo¤å¡A¤£¥i¦¬" & lblCasePropertyName & "¡I", vbCritical

            Else
               PUB_RecvCheck412 = True
            End If
         '«D¥xÆW®×
         Else
            PUB_RecvCheck412 = True
         End If
      End If
   
   End With
   CheckOC
   
ErrHnd:

   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Add By Sindy 2022/11/21
'±µ¬¢³æ·s®×¦³¿éªk©w´Á­­®É¡A¥»©Ò´Á­­«h¨Ì¤U¦C³W«hÀË¬d¡A¿é¤J¤§¥»©Ò´Á­­¤£¥i¤p©ó¦¹¼Ò²Õ­pºâ¥X¨Óªº¥»©Ò´Á­­¡C
'­Y¤£²Å³W«h®É¼u°T®§´£¿ô¦ý¤´¥i¾Þ§@¡A°ß¥»©Ò´Á­­¤£¥i<¨t²Î¤é¡G½Ð¼g¦¨¼Ò²Õ¥H«K«áÄò¥i§Q¥Î
'¦U³¡ªù¤À®×µ{¦¡¤]­n¥[¤J¤W­zÀË¬d¡A¤@¼Ë¤]¬O·s®×¦³ªk©w´Á­­®É¡AÀË¬d¿é¤J¤§¥»©Ò´Á­­¤£¥i¤p©ó¦¹¼Ò²Õ­pºâ¥X¨Óªº¥»©Ò´Á­­¡C
'­Y¤£²Å³W«h®É¼u°T®§´£¿ô¦ý¤´¥i¾Þ§@¡A°ß¥»©Ò´Á­­¤£¥i<¨t²Î¤é¡C
'Modify By Sindy 2023/3/8 , Optional ByRef strOurDeadline As String = "" : ¨t²Î­pºâ¥X¨Óªº¥»©Ò´Á­­
'                         , Optional ByVal bolChkShowMsg As Boolean = True : ¶·ÀË¬d¤é´Á¦³°ÝÃD¯Â¼u°T®§
'Modify By Sindy 2023/3/9 , Optional ByVal bolOnlyCountCP06 As Boolean = False : ¶È¬°­pºâ¥»©Ò´Á­­
'Modify By Sindy 2023/4/12 , Optional ByVal strCP09 As String = "" : Á`¦¬¤å¸¹
Public Function PUB_CRLUseCP07CheckCP06(strNewCase As String, strNA01 As String, strCP01 As String, strCP10 As String, _
   strCP06 As String, strCP07 As String, Optional ByRef strOurDeadline As String = "", Optional ByVal bolChkShowMsg As Boolean = True _
   , Optional ByVal bolOnlyCountCP06 As Boolean = False, Optional ByVal strCP09 As String = "") As Boolean
Dim strExSql As String, intA As Integer, rsAD As New ADODB.Recordset 'Added by Lydia 2024/05/16

   PUB_CRLUseCP07CheckCP06 = True
   If strNA01 = "" Then GoTo EXITSUB
   If strCP01 = "" Then GoTo EXITSUB
   If strCP10 = "" Then GoTo EXITSUB
   If Val(DBDATE(strCP07)) = 0 Then GoTo EXITSUB
   
   'Add By Sindy 2023/4/12 ÀË¬d¤å¸¹­Y¬°¤U¤@´Á­­¦¬¤åªº´N¤£¥ÎÀË¬d´Á­­ ex:P-123113(AB2013581)
   If strCP09 <> "" Then
      strExSql = "Select NP01 From NextProgress Where NP24='" & strCP09 & "'"
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strExSql)
      If intA = 1 Then
         GoTo EXITSUB
      End If
   End If
   '2023/4/12 END
   
   strOurDeadline = ""
   'Modify By Sindy 2023/3/28
   '·s®×¦³¿éªk©w´Á­­®É¡A¥»©Ò´Á­­«h¨Ì¤U¦C³W«hÀË¬d...
'   If strNewCase = "Y" Then
   '2023/3/28 END
   
      '©Ò¦³¥xÆW®×¥»©Ò´Á­­=ªk©w´Á­­¡Ð2­Ó¤u§@¤Ñ¡]¤£§t·í¤é¡^
      'T¤j³°Äò®i®×¥ó
      'Add By Sindy 2025/9/22 ¨q¬Â:±µ¬¢³æT®×¤¤¶¡±µ¶i¨Ó¤§·s®×¡A­Y¦³¿é¤Jªk©w´Á­­¡A¥»©Ò´Á­­¤@«ß¹w³]ªk©w´Á­­-2­Ó¤u§@¤Ñ¡C
      '                       µ{§Ç¤À®×®É­YÄ±±o¤£§´¦A¦Û¦æ­×§ï¡Aµ{¦¡·|´£¿ô»P±µ¬¢³æ¤£¦P¡C
      If strNA01 = "000" Or _
         (strCP01 = "T" And strNA01 = "020" And InStr(strCP10, "102") > 0) Or _
         (strCP01 = "T" And strNewCase = "Y" And strCP10 <> "101") Then
         strOurDeadline = PUB_GetOurDeadline(strCP07)
         
      'P«D¥xÆW®×¥»©Ò´Á­­=ªk©w´Á­­¡Ð10¤Ñ¡]¤£§t·í¤é¡^¡A¦A©¹«e±À¤u§@¤é
      ElseIf strCP01 = "P" And strNA01 <> "000" Then
         strOurDeadline = CompDate(2, -10, TransDate(strCP07, 2))
         strOurDeadline = PUB_GetWorkDay1(strOurDeadline, True)
         
      'CFP®×
      ElseIf strCP01 = "CFP" Then
         If strNA01 = "102" And InStr(strCP10, "107") > 0 Then '102¥[®³¤j,107µªÅG
            '¥[®³¤jµªÅG´Á­­¥»©Ò´Á­­=ªk©w´Á­­¡Ð14¤Ñ¡Ð1­Ó¤ë¡]¤£§t·í¤é¡^¡A¦A©¹«e±À¤u§@¤é
            strOurDeadline = CompDate(2, -14, TransDate(strCP07, 2))
            strOurDeadline = CompDate(1, -1, TransDate(strOurDeadline, 2))
            strOurDeadline = PUB_GetWorkDay1(strOurDeadline, True)
         Else
            'CFP¥»©Ò´Á­­=ªk©w´Á­­¡Ð14¤Ñ¡]¤£§t·í¤é¡^¡A¦A©¹«e±À¤u§@¤é
            strOurDeadline = CompDate(2, -14, TransDate(strCP07, 2))
            strOurDeadline = PUB_GetWorkDay1(strOurDeadline, True)
         End If
         
      'Modify By Sindy 2023/3/8 mark
'      'T¤j³°¥»©Ò´Á­­=ªk©w´Á­­¡Ð15¤Ñ¡]¤£§t·í¤é¡^¡A¦A©¹«e±À¤u§@¤é
'      ElseIf strCP01 = "T" And strNA01 <> "000" Then
'         strOurDeadline = CompDate(2, -15, TransDate(strCP07, 2))
'         strOurDeadline = PUB_GetWorkDay1(strOurDeadline, True)
         
      'TF®×
      ElseIf strCP01 = "TF" Then
         'TF¨Ï¥Î«Å»}105¡G¥»©Ò´Á­­=ªk©w´Á­­¡Ð2­Ó¤ë¡A¦A©¹«e±À¤u§@¤é
         If InStr(strCP10, "105") > 0 Then
            strOurDeadline = CompDate(1, -2, TransDate(strCP07, 2))
            strOurDeadline = PUB_GetWorkDay1(strOurDeadline, True)
         'TF©µ®i102¡G¥»©Ò´Á­­=ªk©w´Á­­¡Ð1­Ó¤ë¡A¦A©¹«e±Àªº¤u§@¤Ñ
         ElseIf InStr(strCP10, "102") > 0 Then
            strOurDeadline = CompDate(1, -1, TransDate(strCP07, 2))
            strOurDeadline = PUB_GetWorkDay1(strOurDeadline, True)
         'Modify By Sindy 2023/3/8 mark
'         Else
'            '¨ä¥L®×¥ó©Ê½è¥»©Ò´Á­­=ªk©w´Á­­¡Ð2¤Ñ¡]¤£§t·í¤é¡^¡A¦A©¹«e±À¤u§@¤é
'            strOurDeadline = CompDate(2, -2, TransDate(strCP07, 2))
'            strOurDeadline = PUB_GetWorkDay1(strOurDeadline, True)
         End If
         
      'CFT®×
      ElseIf strCP01 = "CFT" Then
         'CFT©µ®i102¡B¨Ï¥Î«Å»}105¡G¥»©Ò´Á­­=ªk©w´Á­­¡Ð2­Ó¤ë¡A¦A©¹«e±À¤u§@¤é
         If InStr(strCP10, "102") > 0 Or InStr(strCP10, "105") > 0 Then
            strOurDeadline = CompDate(1, -2, TransDate(strCP07, 2))
            strOurDeadline = PUB_GetWorkDay1(strOurDeadline, True)
'         Else
'            '¨ä¥L®×¥ó©Ê½è¥»©Ò´Á­­¤Îªk©w´Á­­¥Ñ¾Þ§@ªÌ¦Û¦æ¿é¤J¡A¥»©Ò´Á­­Äæ¨t²Î¦A©¹«e±À¤u§@¤é
'            If DBDATE(strCP06) > 0 Then
'               strOurDeadline = PUB_GetWorkDay1(DBDATE(strCP06), True)
'            End If
         End If
      End If
      
      'Modify By Sindy 2023/10/19 Val(strCP07) => Val(DBDATE(strCP07))
      If Val(DBDATE(strCP07)) >= strSrvDate(1) Then 'ªk©w´Á­­¨S¦³¹L´Á®É Add By Sindy 2023/6/20 + if
         If Val(strOurDeadline) > 0 Then
            '¥»©Ò´Á­­¤p©ó¨t²Î¤é,¥»©Ò´Á­­=¨t²Î¤é
            If Val(strOurDeadline) < strSrvDate(1) Then
               strOurDeadline = strSrvDate(1)
            End If
         End If
      End If
'   End If
   
   If bolOnlyCountCP06 = False Then
      If Val(DBDATE(strCP07)) <> Val(strSrvDate(1)) Then
         If Val(DBDATE(strCP06)) > Val(DBDATE(strCP07)) Then
            MsgBox "¥»©Ò´Á­­¤£¥i¤j©óªk©w´Á­­¡I", vbExclamation
            PUB_CRLUseCP07CheckCP06 = False
            GoTo EXITSUB
         End If
      End If
      
      '¨t²Î¨S¦³­pºâ¥X¥»©Ò´Á­­ ©Î ÀË¬d¤é´Á¦³°ÝÃD¯Â¼u°T®§®É,­nÀË¬d...
      If Val(strOurDeadline) = 0 Or bolChkShowMsg = True Then
         If Val(DBDATE(strCP06)) = 0 Then  '¦³¿éªk©w´Á­­,¥»©Ò´Á­­¤£¥iªÅ¥Õ
            MsgBox "¥»©Ò´Á­­" & MsgText(9006), vbExclamation
            PUB_CRLUseCP07CheckCP06 = False
            GoTo EXITSUB
            
         '¿é¤J¤§¥»©Ò´Á­­¤£¥i¤p©ó¦¹¼Ò²Õ­pºâ¥X¨Óªº¥»©Ò´Á­­,­Y¤£²Å³W«h®É¼u°T®§´£¿ô¦ý¤´¥i¾Þ§@¡A°ß¥»©Ò´Á­­¤£¥i<¨t²Î¤é
         'Modify By Sindy 2023/2/17 ½Ð±±¨îT©µ®i102¡BFCT©µ®i102¡BP»âÃÒ601¡BP¦~¶O605¡A½Ð²¤¹L¥»©Ò´Á­­¤£¥i¤p©ó¨t²Î¤éªºÀË¬d¡A
         '                          ¦]¬°³o¨Ç®×¥ó©Ê½è¦b¦sÀÉ®É¡Aµ{¦¡·|¦Û°Ê­pºâ¹O´Á«áªº·s´Á­­§ó·s¦^¥h¡A©Ò¥H¤£¯àÀË¬d¾×¦í¦sÀÉ¡C
         ElseIf Val(DBDATE(strCP06)) < Val(strSrvDate(1)) And _
            Not (strCP01 = "P" And (InStr(strCP10, "601") > 0 Or InStr(strCP10, "605") > 0)) And _
            Not ((strCP01 = "T" Or strCP01 = "FCT") And InStr(strCP10, "102") > 0) Then
            
            MsgBox "¥»©Ò´Á­­¤£¥i¤p©ó¨t²Î¤é¡I", vbExclamation
            PUB_CRLUseCP07CheckCP06 = False
            GoTo EXITSUB
            
         ElseIf Val(DBDATE(strCP06)) < Val(strOurDeadline) Then
            If MsgBox("¥»©Ò´Á­­¤£¥i¤p©ó " & TransDate(Val(strOurDeadline), 1) & " (¨t²Î¨Ì³W©w¥Hªk©w´Á­­­pºâ¤§µ²ªG)¡A½T©wÄ~Äò¾Þ§@¶Ü¡H", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then
               PUB_CRLUseCP07CheckCP06 = False
               GoTo EXITSUB
            End If
         Else
            '§PÂ_¬O§_¬°¤u§@¤Ñ
            If ChkWorkDay(DBDATE(strCP06)) = False Then
               MsgBox "¥»©Ò´Á­­¥²¶·¬°¤u§@¤Ñ¡I", vbExclamation
               PUB_CRLUseCP07CheckCP06 = False
               GoTo EXITSUB
            End If
         End If
      End If
   End If
   
   '¦è¤¸¤é´Á§ï¬°¥Á°ê¤é´Á
   If Val(strOurDeadline) > 0 Then strOurDeadline = TransDate(strOurDeadline, 1)
   
'Added by Lydia 2024/05/16
EXITSUB:
   Set rsAD = Nothing
   
End Function

'Add By Sindy 2024/1/30 ¦U³¡ªù¤À®×®É¡A­Y¥»©Ò´Á­­»Pªk©w´Á­­»P±µ¬¢³æªº¥»©Ò´Á­­»Pªk©w´Á­­¤£¦P®É¡A­n´£¿ô
Public Function PUB_ChkCRLdtCP06CP07(strCP09 As String)
Dim strCP06 As String, strCRLCP06 As String
Dim strCP07 As String, strCRLCP07 As String
Dim strCP140 As String
Dim strCP157 As String
Dim strExSql As String, intA As Integer, rsAD As New ADODB.Recordset 'Added by Lydia 2024/05/16

   strExSql = "Select cp06,cp07,cp140,cp157 From CaseProgress Where CP09='" & strCP09 & "'"
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strExSql)
   If intA = 1 Then
      strCP06 = "" & rsAD.Fields("cp06")
      strCP07 = "" & rsAD.Fields("cp07")
      strCP140 = "" & rsAD.Fields("cp140")
      strCP157 = "" & rsAD.Fields("cp157")
   End If
   'Modify By Sindy 2024/1/31 ¼W¥[§PÂ_¥¼¤À®×¤~ÀË¬d
   If strCP140 <> "" And Val(strCP157) = 0 Then
      strExSql = "Select * From ConsultRecordList Where CRL01='" & strCP140 & "'"
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strExSql)
      If intA = 1 Then
         strCRLCP06 = "" & rsAD.Fields("CRL12")
         strCRLCP07 = "" & rsAD.Fields("CRL13")
      End If
      If strCP06 <> strCRLCP06 Or strCP07 <> strCRLCP07 Then
         MsgBox "¦¹¶i«×¤§¥»©Ò´Á­­»Pªk©w´Á­­»P±µ¬¢³æªº¥»©Ò´Á­­»Pªk©w´Á­­¤£¦P¡A" & vbCrLf & vbCrLf & "½Ðµø®×¥óª¬ªp­×§ï®×¥ó°ò¥»¸ê®Æ©Î¶i«×¸ê®Æ¡I", vbInformation, "½Ðª`·N"
      End If
   End If
   
   Set rsAD = Nothing 'Added by Lydia 2024/05/16
End Function

'Add By Sindy 2023/12/13 ÀË¬d±µ¬¢³æªºFlow¬O§_­nµ²§ô
'Modify By Sindy 2024/11/20 + Optional ByVal oForm As Form, Optional ByVal bolEPC218 As Boolean = False
Public Function PUB_UpdateCRLFlowClose(strCP140 As String, strCP09 As String, _
   Optional ByVal oForm As Form, Optional ByVal bolEPC218 As Boolean = False)
Dim rsA As New ADODB.Recordset
Dim strTmp As String, intA As Integer 'Added by Lydia 2024/05/15
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim strCP13 As String, strCC As String
   
   'Add By Sindy 2024/11/20 ¸Ñ°£´Á­­ ©M ¨ú®ø¦¬¤å
   If UCase(TypeName(oForm)) = UCase("frm110101_2") Or _
      UCase(TypeName(oForm)) = UCase("frm110102_2") Then
      strTmp = "select * from caseprogress,patent where cp09='" & strCP09 & "' and cp01='CFP'" & _
               " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)" & _
               " and pa09='221'"
      intA = 1
      Set rsA = ClsLawReadRstMsg(intA, strTmp)
      If intA = 1 Then
         strCP01 = rsA.Fields("cp01")
         strCP02 = rsA.Fields("cp02")
         strCP03 = rsA.Fields("cp03")
         strCP04 = rsA.Fields("cp04")
         strCP13 = rsA.Fields("cp13")
         strCC = GetDeptMan(PUB_GetST93(strCP13))
         '­Y¦³EPC®×¥ó¤w¦¬¤å¹êÅé¼f¬d¤Î«ü©w¶O¦ý¥¼µo¤å¡A©ó¶i¦æ¦^ÂÐÀË¯Á³ø§i¤§µ²®×µ{§Ç®É(µ{§Ç¤H­û¦b¾Þ§@µ²®×½T©w®É)
         '¥X²{´£¿ôµøµ¡¶i¦æ¹êÅé¼f¬d¤Î«ü©w¶O¤§®×¥ó¤Î¶O¥Î³B²z¡A­Y¬°´¼Åv³¡¤§¤H­û¡A¨Ã½Ð¦P®Éµo³qª¾¤©°Ï¥DºÞ¤Î¥»¤H¡C
         If rsA.Fields("cp10") = "218" Or bolEPC218 = True Then '218=¦^ÂÐÀË¯Á³ø§i
            strTmp = "select cpm04 from caseprogress,CASEPROPERTYMAP" & _
                     " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                     " and cp10 in('416','215')" & _
                     " and cp158=0 and cp159=0" & _
                     " and cp01=cpm01 and cp10=cpm02"
            intA = 1
            Set rsA = ClsLawReadRstMsg(intA, strTmp)
            If intA = 1 Then
               strExc(10) = ""
               rsA.MoveFirst
               Do While Not rsA.EOF
                  strExc(10) = strExc(10) & "¡B" & rsA.Fields("cpm04")
                  rsA.MoveNext
               Loop
               strExc(10) = Mid(strExc(10), 2)
               strTmp = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                  " VALUES ( '" & strUserNum & "','" & strCP13 & "',to_char(sysdate,'yyyymmdd')" & _
                  ",to_char(sysdate,'hh24miss'),'EPC¦^ÂÐÀË¯Á³ø§iµ²®×¡A½Ð¶i¦æ¥¼µo¤å¤§" & strExc(10) & "¤§®×¥ó¤Î¶O¥Î³B²z !','¦p¦®','" & strCC & "')"
               cnnConnection.Execute strTmp
               MsgBox "¨t²Î±N³qª¾´¼Åv¤H­û³B²z¥¼µo¤å¤§" & strExc(10) & " !" & vbCrLf & _
                      "¦P®Éµo³qª¾¤©°Ï¥DºÞ¤Î¥»¤H¡C", vbExclamation
            End If
         End If
      End If
   End If
   '2024/11/20 END
   
   'Add By Sindy 2023/3/7
   If Len(strCP140) = 10 Then
      strTmp = "select f0301 from flow003 where f0301='" & strCP140 & "' and f0308='A7' and f0309='" & Flow_³B²z¤¤ & "'"
      intA = 1
      Set rsA = ClsLawReadRstMsg(intA, strTmp)
      If intA = 1 Then
         'ÀË¬d±µ¬¢³æ¥þ³¡®×¥ó©Ê½è¬O§_¥þ³¡¤À®×§¹¦¨
         'Modify By Sindy 2024/3/8 +, strCP140
         If PUB_GetCP140CP157IsOK(strCP09, strCP140) = True Then
             'Ã±®ÖÀÉ(¤w³B²z)
             strSql = "update FLOW002 set " & _
                    "F0205='" & strSrvDate(1) & "'" & _
                    ",F0206='" & Right("000000" & ServerTime, 6) & "'" & _
                    ",F0207='3',F0204='" & strUserNum & "'" & _
                    " where F0201='" & strCP140 & "' and F0202='A7' and F0207 is null "
             cnnConnection.Execute strSql
             'ªí³æ¥DÀÉ
             strSql = "update FLOW003 set " & _
                     "F0309=" & CNULL(Flow_¤w¤À®×) & _
                     " where F0301='" & strCP140 & "' "
             cnnConnection.Execute strSql
         End If
      End If
   End If
   '2023/3/7 END
   Set rsA = Nothing
End Function

'Add By Sindy 2022/11/22 ÀË¬d±µ¬¢³æ¥þ³¡®×¥ó©Ê½è¬O§_¥þ³¡¤À®×§¹¦¨
'Modify By Sindy 2024/3/8 +, Optional ByVal strCP140 As String
Public Function PUB_GetCP140CP157IsOK(ByVal stCP09 As String, Optional ByVal strCP140 As String) As Boolean
   Dim RsQ As New ADODB.Recordset
   Dim stQ As String, intQ As Integer
   'Dim strCP140 As String
   
   PUB_GetCP140CP157IsOK = True
   'Modify By Sindy 2024/3/8
   If strCP140 = "" Then
   '2024/3/8 END
      stQ = "Select CP140 From CaseProgress Where CP09='" & stCP09 & "'"
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, stQ)
      If intQ = 1 Then
         strCP140 = "" & RsQ.Fields("CP140")
      End If
   End If
   If strCP140 <> "" Then
      '¤@±i±µ¬¢³æ¥i¥H¦h¹D¦¬¤å
      'Modify By Sindy 2023/12/13 ±Æ°£¤w¨ú®ø¦¬¤å => and CP159 is null
      stQ = "Select CP140 From CaseProgress Where CP140='" & strCP140 & "' and CP157 is null and CP159=0"
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, stQ)
      If intQ = 1 Then
         PUB_GetCP140CP157IsOK = False
      End If
   End If
   
   Set RsQ = Nothing
End Function

'Add By Sindy 2022/11/22 µ{§Ç¤À®×§@·~­nÅã¥Üªº¥Ø«eªí³æª¬ºA
Public Function PUB_GetCP157forF0309(ByVal stCP09 As String) As String
   Dim RsQ As New ADODB.Recordset
   Dim stQ As String, intQ As Integer
   Dim strCP140 As String
   
   PUB_GetCP157forF0309 = ""
   
   'Modify By Sindy 2024/1/26
'   stQ = "Select CP09,CP157,CP140,Decode(F0309,'" & Flow_³B²z¤¤ & "',Decode(Decode(Decode(cp157,null,0,1),1,'¤w¤À®×',F0309),'¤w¤À®×','¤w¤À®×'," & ShowFlowªí³æª¬ºA¤¤¤å & ")" & _
'                                          "," & ShowFlowªí³æª¬ºA¤¤¤å & ") as Status" & _
'         " From CaseProgress,flow003 Where CP09='" & stCP09 & "' and CP140=F0301"
   stQ = "Select CP09,CP157,CP140,Decode(F0309," & ShowFlowªí³æª¬ºA¤¤¤å & ") as Status" & _
         " From CaseProgress,flow003 Where CP09='" & stCP09 & "' and CP140=F0301"
   '2024/1/26 END
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stQ)
   If intQ = 1 Then
      PUB_GetCP157forF0309 = "" & RsQ.Fields("Status")
   'Add By Sindy 2024/1/26
   Else
      'µLÃ±®Ö¬yµ{
      stQ = "Select CP09,CP157 From CaseProgress Where CP09='" & stCP09 & "'"
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, stQ)
      If intQ = 1 Then
         If Val("" & RsQ.Fields("CP157")) > 0 Then '¦³¥_©Ò¤À®×¤é
            PUB_GetCP157forF0309 = "¤w¤À®×"
         End If
      End If
      '2024/1/26 END
   End If
   
   Set RsQ = Nothing
End Function

'Add By Sindy 2022/11/23 ÀË¬d¬O§_¬°¹q¤l¦¬¤åªº±µ¬¢³æ½s¸¹
Public Function Pub_GetIsFlowCP140(ByVal stCP09 As String) As String
    Dim RsQ As ADODB.Recordset
    Dim stQ As String, intQ As Integer
    
    Pub_GetIsFlowCP140 = ""
    stQ = "Select f0301 From CaseProgress,flow003 Where cp09='" & stCP09 & "' and cp140 is not null and cp140=f0301(+) and f0301 is not null"
    intQ = 1
    Set RsQ = ClsLawReadRstMsg(intQ, stQ)
    If intQ = 1 Then
        Pub_GetIsFlowCP140 = "" & RsQ.Fields("f0301")
    End If
    
    Set RsQ = Nothing
End Function

'Add By Sindy 2022/11/23 ¨ú±o±µ¬¢³æªº¬Û¦P®×¸¹¸ê®Æ
'strType= 0.¬Û¦P®×¸¹ 3.¤@®×¨â½Ð 6.ÀÀ¨î³à¥¢·s¿o©Ê
'strSys= ¨t²Î§O
Public Function Pub_GetCRLCaseMap(ByVal strCRL01 As String, ByVal strType As String, ByVal strSys As String, _
   ByVal strCP01 As String, ByVal strCP02 As String, ByVal strCP03 As String, ByVal strCP04 As String) As String
Dim RsQ As ADODB.Recordset
Dim stQ As String, intQ As Integer
Dim arrTmp As Variant, ii As Integer
   
   Pub_GetCRLCaseMap = ""
   Select Case strType
      Case "0" '¬Û¦P®×¸¹
         'Modify By Sindy 2023/1/10 + and CRL56='Y':¬Û¦PÃö«Yªº®×¸¹¤~§ì
         stQ = "Select CRL55 From consultrecordlist Where CRL01='" & strCRL01 & "' and CRL55 is not null and CRL74 is null and CRL56='Y' and CRL55<>nvl(CRL67,' ')"
         intQ = 1
         Set RsQ = ClsLawReadRstMsg(intQ, stQ)
         If intQ = 1 Then
            arrTmp = Split(RsQ.Fields("CRL55"), ",")
            For ii = LBound(arrTmp) To UBound(arrTmp)
               If arrTmp(ii) <> "" And _
                  Not (SystemNumber(CStr(arrTmp(ii)), 1) = strCP01 And SystemNumber(CStr(arrTmp(ii)), 2) = strCP02 And _
                       SystemNumber(CStr(arrTmp(ii)), 3) = strCP03 And SystemNumber(CStr(arrTmp(ii)), 4) = strCP04) Then
                  If strSys = "P" Then
                     'Modify By Sindy 2023/11/7 P-132503ªº¬Û¦P®×¸¹¬°CFP-033632
                     'If SystemNumber(CStr(arrTmp(ii)), 1) = "P" Then
                        Pub_GetCRLCaseMap = arrTmp(ii)
                        Exit For
                     'End If
                  Else
                     If SystemNumber(CStr(arrTmp(ii)), 1) = "CFP" Then
                        Pub_GetCRLCaseMap = arrTmp(ii)
                        Exit For
                     End If
                  End If
               End If
            Next
         End If
         
      Case "3" '¤@®×¨â½Ð
         stQ = "Select CRL67 From consultrecordlist Where CRL01='" & strCRL01 & "' and CRL67 is not null"
         intQ = 1
         Set RsQ = ClsLawReadRstMsg(intQ, stQ)
         If intQ = 1 Then
            If Not (SystemNumber(RsQ.Fields("CRL67"), 1) = strCP01 And SystemNumber(RsQ.Fields("CRL67"), 2) = strCP02 And _
                    SystemNumber(RsQ.Fields("CRL67"), 3) = strCP03 And SystemNumber(RsQ.Fields("CRL67"), 4) = strCP04) Then
               Pub_GetCRLCaseMap = RsQ.Fields("CRL67")
            End If
         End If
         
      Case "6" 'ÀÀ¨î³à¥¢·s¿o©Ê
         stQ = "Select CRL68 From consultrecordlist Where CRL01='" & strCRL01 & "' and CRL68 is not null"
         intQ = 1
         Set RsQ = ClsLawReadRstMsg(intQ, stQ)
         If intQ = 1 Then
            If Not (SystemNumber(RsQ.Fields("CRL68"), 1) = strCP01 And SystemNumber(RsQ.Fields("CRL68"), 2) = strCP02 And _
                    SystemNumber(RsQ.Fields("CRL68"), 3) = strCP03 And SystemNumber(RsQ.Fields("CRL68"), 4) = strCP04) Then
               Pub_GetCRLCaseMap = RsQ.Fields("CRL68")
            End If
         End If
         
   End Select
   
   Set RsQ = Nothing
End Function

'Add By Sindy 2022/11/24 ¹q¤l¦¬¤å«e¥ýµ¹³sÄò®×¸¹
'Modify By Sindy 2023/5/4 + strCRL01 As String, Optional bolShowMsg As Boolean
Public Function Pub_GetContinuousCaseNo(strCRL01 As String, Optional bolShowMsg As Boolean = True) As Boolean
   Dim RsQ As ADODB.Recordset
   Dim stQ As String, intQ As Integer
   Dim RsUpd As ADODB.Recordset
   Dim stUpd As String, intUpd As Integer
   Dim strCRL65 As String, strCRL07 As String
   Dim intRow As Integer, dblMaxCase As Double, dblStarCase As Double
   Dim strUpdCRL01 As String
   Dim strAutoNumber As String 'Add By Sindy 2024/7/26
   
On Error GoTo CheckingErr
   
   Pub_GetContinuousCaseNo = False 'Add By Sindy 2023/5/4
   
   '¥²¶·¬°·s®×ªÅ¸¹¦³³]Ãö³s¸s²Õ:
   stQ = "Select CRL65 From consultrecordlist Where CRL01='" & strCRL01 & "' and CRL06='Y' and CRL08 is null and CRL65 is not null"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, stQ)
   If intQ = 1 Then
      'Ãö³sªí³æ½s¸¹(³]¸s²Õ)
      strCRL65 = RsQ.Fields("CRL65")
      
      stQ = "Select CRL07,count(*) From consultrecordlist Where CRL65='" & strCRL65 & "' and CRL08 is null group by CRL07"
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, stQ)
      If intQ = 1 Then
         RsQ.MoveFirst
         Do While Not RsQ.EOF
            '¨t²Î§O
            strCRL07 = RsQ.Fields("CRL07")
            intRow = RsQ.Fields(1)
            
            cnnConnection.BeginTrans
            
            '¥ý·m¸¹
            strSql = "update autonumber set au03=au03+" & intRow & " where au01='" & strCRL07 & "'"
            cnnConnection.Execute strSql
            '¨ú±o³sÄò®×¸¹
            strSql = "select au03 from autonumber where au01='" & strCRL07 & "'"
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               dblMaxCase = RsTemp.Fields("au03")
               dblStarCase = dblMaxCase - intRow
            End If
            
            '§ó·s·sªº®×¸¹
            stUpd = "Select CRL01 From consultrecordlist Where CRL65='" & strCRL65 & "' and CRL07='" & strCRL07 & "' and CRL08 is null order by CRL01 asc"
            intUpd = 1
            Set RsUpd = ClsLawReadRstMsg(intUpd, stUpd)
            If intUpd = 1 Then
               If intRow = RsUpd.RecordCount Then
                  RsUpd.MoveFirst
                  Do While Not RsUpd.EOF
                     strUpdCRL01 = RsUpd.Fields("CRL01")
                     dblStarCase = dblStarCase + 1
                     'Modify By Sindy 2024/7/26
                     If strCRL07 = °¨¼w¨½®× Then
                        strAutoNumber = Format(dblStarCase, "00000") + "0"
                     Else
                        strAutoNumber = Format(dblStarCase, "000000")
                     End If
                     '2024/7/26 END
                     
                     strSql = "update consultrecordlist set CRL08='" & strAutoNumber & "',CRL09='0',CRL10='00' where CRL01='" & strUpdCRL01 & "'"
                     cnnConnection.Execute strSql
                     
                     RsUpd.MoveNext
                  Loop
               End If
            End If
            
            cnnConnection.CommitTrans
            
            RsQ.MoveNext
         Loop
      End If
   Else
      If bolShowMsg = True Then MsgBox "±µ¬¢³æ½s¸¹¥²¶·¬°·s®×ªÅ¸¹¨Ã¥B¦³³]©wÃö³s¸s²Õ¡I", vbExclamation
      Exit Function
   End If
   If bolShowMsg = True Then MsgBox "¤w¨ú±o³s¸¹¡I¡I", vbInformation
   
   Set RsQ = Nothing
   
   Pub_GetContinuousCaseNo = True 'Add By Sindy 2023/5/4
   Exit Function
   
CheckingErr:
   cnnConnection.RollbackTrans
   If Err.Description <> "" Then
      MsgBox (Err.Description), , "¨ú±o³sÄò®×¸¹"
   Else
      MsgBox "¨ú±o³sÄò®×¸¹¥¢±Ñ¡I", vbExclamation, "¨ú±o³sÄò®×¸¹"
   End If
End Function
'2022/11/24 END

'Add By Sindy 2023/1/19 ·s¼W±µ¬¢³æ«È¤áµo©ú¤H¸ê®Æ
Public Function PUB_InsertCRLInventor(strCRL01 As String) As Boolean
Dim rsPD As New ADODB.Recordset
Dim strCRA05 As String, strCU23 As String
Dim strCRI01 As String, strCRI02 As String
Dim strIN02 As String, strIN03 As String, strIN04 As String
Dim strIN05 As String, strIN07 As String, strIN11 As String
   
   PUB_InsertCRLInventor = False
   
   'ÀË¬d±µ¬¢³æµo©ú¤H¸ê®ÆÀÉ,¬O§_¦³·sªºµo©ú¤H­n«ØÀÉ
   strSql = "select * from consultrecinv where cri01='" & strCRL01 & "' and cri03 is null" & _
            " order by cri02 asc"
   intI = 1
   Set rsPD = ClsLawReadRstMsg(intI, strSql)
   If intI = 1 Then
      '§ì¥Ó½Ð¤H1
      strSql = "select cra05,cu23 from consultrecApp,customer where cra01='" & strCRL01 & "' and cra02=1" & _
               " and cu01=cra05 and cu02='0'"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         strCRA05 = RsTemp.Fields("cra05")
         strCU23 = "" & RsTemp.Fields("cu23") '«È¤á¤¤¤å¦a§}
      End If
      If strCRA05 <> "" Then
         rsPD.MoveFirst
         Do While Not rsPD.EOF
            strCRI01 = rsPD.Fields("CRI01")
            strCRI02 = rsPD.Fields("CRI02")
            strIN02 = PUB_GetNewIN02(strCRA05) 'µo©ú¤H¥N¸¹-¨ú¸¹
            strIN03 = Trim("" & rsPD.Fields("cri05")) 'ID
            strIN04 = Trim("" & rsPD.Fields("cri06")) '¤¤¤å¦WºÙ
            strIN05 = "" '­^¤å¦WºÙ
            Call PUB_analyzeINVname(strIN04, strIN04, strIN05) '¸ÑªRµo©ú¤H¦WºÙ
            
            '¤¤¤å¦a§}
            If Trim("" & rsPD.Fields("cri07")) = "µo©ú¤H¦a§}¦P¥Ó½Ð¤H¦a§}" Then
               strIN07 = strCU23
            Else
               strIN07 = Trim("" & rsPD.Fields("cri07"))
            End If
            '°êÄy
            strIN11 = Trim("" & rsPD.Fields("cri08"))
            If strIN11 <> "" Then strIN11 = Left(strIN11, 3) '¨ú«e3½X
            
            '¥ýÀË¬d¸ê®Æ¬O§_¤w¦s¦b
            strSql = "select * from inventor where in01='" & strCRA05 & "'"
            If strIN03 <> "" Then 'ID
               strSql = strSql & " and in03='" & strIN03 & "'"
            End If
            If strIN04 <> "" Then '¤¤¤å¦WºÙ
               strSql = strSql & " and in04='" & strIN04 & "'"
            Else '­^¤å¦WºÙ
               strSql = strSql & " and in05='" & strIN05 & "'"
            End If
            intI = 1
            Set RsTemp = ClsLawReadRstMsg(intI, strSql)
            If intI = 1 Then
               strIN02 = RsTemp.Fields("in02")
            Else
               '·s¼W«È¤áµo©ú¤HÀÉ
               strSql = "insert into inventor(in01,in02,in03,in04,in05,in07,in11)" & _
                        " Values('" & strCRA05 & "','" & strIN02 & "','" & strIN03 & "'" & _
                        ",'" & strIN04 & "','" & strIN05 & "','" & strIN07 & "','" & strIN11 & "')"
               cnnConnection.Execute strSql, intI
            End If
            
            '§ó·s±µ¬¢³æµo©ú¤H¸ê®ÆÀÉ
            strSql = "update consultrecinv set cri03='" & strCRA05 & "',cri04='" & strIN02 & "'" & _
                     " where cri01='" & strCRI01 & "' and cri02=" & strCRI02
            cnnConnection.Execute strSql, intI
            
            rsPD.MoveNext
         Loop
      End If
   End If
   
   PUB_InsertCRLInventor = True
   Set rsPD = Nothing
End Function

'Add By Sindy 2023/2/2 ¸ÑªRµo©ú¤H¦WºÙ
Public Sub PUB_analyzeINVname(strName As String, _
   Optional ByRef strChName As String, Optional ByRef strEngName As String)
   
Dim bolIsEng As Boolean
Dim strTmp As String, ii As Integer

   '***** ÀË¬d¬O§_¦³­^¤å¦WºÙ *****
   'ÀË¬d²Ä1½X¬O§_¬°­^¤å,­Y¬O«h¬°­^¤å¦WºÙ
   If PUB_GetSimpleName(Mid(strName, 1, 1)) <> "" Then
      strEngName = strName '­^¤å¦WºÙ
      strChName = "" '¤¤¤å¦WºÙ
   Else
      bolIsEng = False
      For ii = 1 To Len(strName)
         strTmp = Mid(strName, ii, 1)
         'ªÅ¥Õ®æ«á­±´N¥Ü¬°­^¤å ©Î PUB_GetSimpleName¶Ç¦^"«DªÅ¥Õ­È"«h¥Ü¬°­^¤å
         'ex:±i´Âªi ZHANG,Zhang,Chaobo
         'ex:HOU,Yingjie
         'ex:ÂÅ±R¤å
         If strTmp = " " Then
            If PUB_GetSimpleName(Mid(strName, ii + 1, 1)) <> "" Then
               bolIsEng = True
            End If
         ElseIf PUB_GetSimpleName(strTmp) <> "" Then
            bolIsEng = True
         End If
         If bolIsEng = True Then
            strTmp = strName
            strChName = Trim(Mid(strTmp, 1, ii)) '¤¤¤å¦WºÙ
            strEngName = Trim(Mid(strTmp, ii)) '­^¤å¦WºÙ
            Exit For
         End If
      Next ii
   End If
End Sub

'Added by Lydia 2023/08/14 §Q¯q½Ä¬ð®×¥ó¡G­­¾\®×¥ó¸É¥R±±ºÞ: ³Q­­¾\¤H­û¦b¦¬¤å®É, ¨t²ÎÀ³¦P®Éµo³qª¾µ¹¦¬¤åªÌ+¨ä¥DºÞ
Private Sub ChkCufaRight(ByVal pUserNo As String, ByVal pCaseNo As String, ByVal pCustNo As String, ByVal pFCno As String)
Dim strBCase(0 To 4) As String
Dim rsBD As New ADODB.Recordset

   strBCase(0) = pCaseNo
   Call ChgCaseNo(strBCase(0), strBCase)
   mStrSql = "select count(*) cnt from cufa_right where cfr02='" & strBCase(1) & "' and cfr03='M51' "
   intJ = 1
   Set rsBD = ClsLawReadRstMsg(intJ, mStrSql)
   If intJ = 1 Then
      If Val("" & rsBD.Fields("cnt")) > 0 Then
         If PUB_ChkCufaByCase("Service1", strBCase(1), pCaseNo, pCustNo, pFCno, pUserNo) = False Then
            strTmp1(1) = PUB_GetFCPProSup(pUserNo)
            If strTmp1(1) <> "" Then
               mStrSql = "select a.*,b.st02 from cufa_right a,staff b where cfr02='" & strBCase(1) & "' and cfr03='M51' and instr('" & pCustNo & "',cfr01)>0 and cfr05=st01(+) " & _
                         "union select a.*,b.st02 from cufa_right a,staff b where cfr02='FCP' and cfr03='M51' and instr('" & pFCno & "',cfr01)>0 and cfr05=st01(+) "
               intJ = 1
               Set rsBD = ClsLawReadRstMsg(intJ, mStrSql)
               If intJ = 1 Then
                  Sleep 100
                  strTmp1(2) = strBCase(1) & "-" & strBCase(2) & IIf(strBCase(3) & strBCase(4) <> "000", "-" & strBCase(3) & "-" & strBCase(4), "") & _
                       "¤w³Q" & rsBD.Fields("st02") & "©ó" & ChangeWStringToTDateString("" & rsBD.Fields("cfr06")) & "´£¥X­­¾\±±ºÞ¡A±zµLÅv­­¬d¸ß®×¥ó¸ê®Æ¡C"
                  mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                      " values('" & strUserNum & "','" & pUserNo & ";" & strTmp1(1) & "',to_char(sysdate,'yyyymmdd')" & _
                      ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strTmp1(2)) & "','¦P¥D¦®',null)"
                  cnnConnection.Execute mStrSql
               End If
            End If
         End If
      End If
   End If
   Set rsBD = Nothing
End Sub

'Added by Lydia 2023/08/18 ¥~±M®×¥ó©R¦W°lÂÜ(TrackingNo)¡G§ï¥ÎFTP(­ì©lÀÉ°Ï)¦s©ñÀÉ®×¡AÀË¬d¬O§_¦³¤W¶ÇÀÉ®×
Public Function PUB_ChkTCNfileExist(ByVal pTCN01 As String, Optional ByVal pKey02 As String) As Boolean
Dim intQ As Integer, strQuery As String
Dim rsQuery As New ADODB.Recordset
   
   PUB_ChkTCNfileExist = False
   If Trim(pTCN01) = "" Then Exit Function
   
   strQuery = "select cpf01,cpf02 from casepaperfile where cpf01='" & Trim(pTCN01) & "' and cpf10<>'D' and substr(upper(cpf02),-4)<>upper('.del') "
   If pKey02 <> "" Then
      strQuery = strQuery & "AND (UPPER(CPF02) LIKE '%" & UCase(pKey02) & "')"
   End If
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
      PUB_ChkTCNfileExist = True
   Else
      PUB_ChkTCNfileExist = False
   End If
   Set rsQuery = Nothing
End Function

'Added by Lydia 2023/08/18 FCP·s®×¥ß¨÷: Tracking NO.§ï¥ÎFTP(­ì©lÀÉ°Ï)¦s©ñÀÉ®×
Public Sub PUB_UpdTCNfile(ByVal mTCNKey As String, ByVal mCaseNo As String, ByVal mCP09 As String, ByVal mCP05 As String, Optional ByRef mSaveDir As String, Optional ByRef bolMoveOK As Boolean)
'mTCNKey: ©R¦W°lÂÜ¬y¤ô¸¹
'mCaseNo: ¥»©Ò®×¸¹
'mCP09,mCP05: ·s®×¦¬¤å¸¹, ¦¬¤å¤é
Dim m_TempDir As String
Dim stCaseName As String 'ÀÉ¦W¶}ÀY
Dim strOldName As String
Dim strNewName As String
Dim strTempA As String
Dim strTempB As String
Dim stEngCP09 As String 'English_Vers¦¬¤å¸¹
Dim bolHaveSEQ As Boolean 'Added by Lydia 2021/04/09
Dim intR As Integer, strR1 As String
Dim rsRD As New ADODB.Recordset
Dim strBC(1 To 4) As String
Dim strCPF02Name As String, stReName As String
Dim strErrMail As String
Dim strSub As String, strTo As String, strCC As String 'Added by Lydia 2024/12/13
   
   Call ChgCaseNo(mCaseNo, strBC)
   
   If mTCNKey = "" Or Len(strBC(2)) <> 6 Then Exit Sub

   If mSaveDir = "" Then
      mSaveDir = App.path & "\" & strUserNum & "\¼È¦s°Ï"
   End If
   m_TempDir = mSaveDir & "\" & mTCNKey
   
   stCaseName = PUB_FCPCaseNo2FileName(strBC(1), strBC(2), strBC(3), strBC(4))
   
On Error GoTo ErrHandle
   
   'Added by Lydia 2024/12/13 FG®×¿é¤J°lÂÜ¬y¤ô¸¹TrackingNo
   If strBC(1) = "FG" Then
      'ª½±µ©ñ¦b·s®×¶i«×
      stEngCP09 = mCP09
   Else
   'end 2024/12/13
      If PUB_ChkCPExist(strBC, cntEnglish_Vers, , stEngCP09, , "D") = True Then 'English_Vers992
      Else   '²£¥ÍEnglish_Vers¦¬¤å¸¹
         stEngCP09 = AutoNo("D", 6)
         strTempA = PUB_GetFCPSalesNo(strBC(1), strBC(2), strBC(3), strBC(4))   'FCP©Ó¿ì
         strTempB = PUB_GetFCPHandler(strBC(1), strBC(2), strBC(3), strBC(4)) 'FCPµ{§Ç
         
         mStrSql = "insert into caseprogress( cp01,cp02,cp03,cp04,cp05,cp09,cp10" & _
            ",cp12,cp13,cp14,cp20,cp26,cp27,cp32 ) values ('" & strBC(1) & "'" & _
             ",'" & strBC(2) & "','" & strBC(3) & "','" & strBC(4) & "'," & mCP05 & ",'" & stEngCP09 & "','" & cntEnglish_Vers & "' " & _
             ",'" & GetSalesArea(strTempA) & "','" & strTempA & "','" & strTempB & "','N','N'," & mCP05 & ",'N')"
         cnnConnection.Execute mStrSql
      End If
   End If
   '³v¤@ÀÉ®×§ó¦W
   strR1 = "select * from casepaperfile where cpf01='" & mTCNKey & "' order by cpf15,cpf16 "
   intR = 1
   Set rsRD = ClsLawReadRstMsg(intR, strR1)
   If intR = 1 Then
      rsRD.MoveFirst
      Do While Not rsRD.EOF
         strOldName = "" & rsRD.Fields("cpf02")
         If Right(UCase(strOldName), Len(FcpTcnFKey02)) = FcpTcnFKey02 Then   '¥~¤å­ì¤å¥»
             strNewName = stCaseName & FcpTcnFKey02
         ElseIf Right(UCase(strOldName), Len(FcpTcnFKey01)) = UCase(FcpTcnFKey01) Then 'msgÀÉ
             If strTempA = "" Then strTempA = Format(ServerTime, "000000") '®É¶¡
             'Á×§K­«ÂÐÀÉ¦W
             If strTempB = strTempA Then
                 strTempA = Format(Val(strTempA) + 1, "000000")
             End If
             strNewName = stCaseName & "." & strSrvDate(1) & strTempA & ".rx" & FcpTcnFKey01
             strTempB = strTempA
         '¡uSEQ.*¡BPRI.*¡BPOA.*¡v(¤£­­ÀÉ®×«¬ºA¡A¥]§tPDF¡BWord¡BTxtÀÉµ¥)¡A³W«h¦P­ì¤å¥»ªº©R¦W¤è¦¡(ORI.PDF)¡C
         ElseIf InStr(UCase("." & strOldName), ".SEQ.") > 0 Then
             strNewName = GetNewFilename(m_TempDir, strOldName, stCaseName, ".SEQ.")
             bolHaveSEQ = True
         ElseIf InStr(UCase("." & strOldName), ".PRI.") > 0 Then
             strNewName = GetNewFilename(m_TempDir, strOldName, stCaseName, ".PRI.")
         ElseIf InStr(UCase("." & strOldName), ".POA.") > 0 Then
             strNewName = GetNewFilename(m_TempDir, strOldName, stCaseName, ".POA.")
         Else
             strNewName = stCaseName & "." & strOldName
         End If
         '¶l¥ó*.msg¥t¥~¤W¶Ç¨ì¨÷©v°Ï
         If Right(UCase(strNewName), Len(FcpTcnFKey01)) = UCase(FcpTcnFKey01) Then
            If PUB_GetAttachFile_Org(mTCNKey, strOldName, m_TempDir & "\" & strNewName, True) = False Then
               MsgBox "ÀÉ®×(" & strOldName & ")¤U¸ü¥¢±Ñ¡I", vbCritical
            Else
               If SaveAttFile_PDF(mCP09, m_TempDir & "\" & strNewName, strNewName, Val(strSrvDate(1)), Val(ServerTime), False) = False Then
                   strErrMail = strErrMail & IIf(strErrMail <> "", vbCrLf, "") & "Âd¥x¦¬¤åµLªk¤W¶Ç" & strNewName & "¨ì¨÷©v°Ï¡F"
               Else
                   PUB_DelPCOrgFile m_TempDir & "\" & strNewName
               End If
            End If
         End If
         
         strCPF02Name = PUB_GetReNameCPF02(strNewName, strBC(1), strBC(2), strBC(3), strBC(4), cntEnglish_Vers, "M")
         'Modified by Lydia 2024/12/13 ¥i¥H¤W¶ÇPDF
         'If PUB_GetEmpFlowReNameFile(strBC(1), strBC(2), strBC(3), strBC(4), cntEnglish_Vers, strCPF02Name, stReName, True, 0) = False Then
         If PUB_GetEmpFlowReNameFile(strBC(1), strBC(2), strBC(3), strBC(4), cntEnglish_Vers, strCPF02Name, stReName, False, 0) = False Then
            Exit Sub
         End If
         intK = 1
         
CheckFileName:
         strR1 = "SELECT CPF01 FROM CasePaperFile WHERE CPF01='" & stEngCP09 & "' and upper(CPF02)=upper('" & stReName & "')"
         intJ = 1
         Set RsTemp = ClsLawReadRstMsg(intJ, strR1)
         If intJ = 1 Then
            stReName = Mid(stReName, 1, InStrRev(stReName, ".") - 1) & "." & intK & Mid(stReName, InStrRev(stReName, "."))
            GoTo CheckFileName
         End If
         mStrSql = "update casepaperfile set CPF01='" & stEngCP09 & "', CPF02='" & stReName & "' WHERE CPF01='" & mTCNKey & "' and CPF02 ='" & rsRD.Fields("cpf02") & "'"
         cnnConnection.Execute mStrSql, intJ
         
         rsRD.MoveNext
      Loop
   End If
   Set rsRD = Nothing

   '¦b·s®×¦¬¤å¥ß¨÷®É¡A§PÂ_·h¤JEnglish Vers­ì©lÀÉ°Ï¦³.SEQ.ÀÉ®×¡A¹w³]¸ÓÄæ¦ì¬°¤Ä¿ï=Y¦³§Ç¦Cªí
   If bolHaveSEQ = True Then
       mStrSql = " update patent set pa175='Y' where pa01='" & strBC(1) & "' and pa02='" & strBC(2) & "' and pa03='" & strBC(3) & "' and pa04='" & strBC(4) & "' "
       cnnConnection.Execute mStrSql, intJ
       mStrSql = " update transcasetitle set tct119='Y' where tct01='" & mCP09 & "' "
       cnnConnection.Execute mStrSql, intJ
   End If
   
   mSaveDir = m_TempDir
   If strErrMail = "" Then
      bolMoveOK = True
   Else
      'Added by Lydia 2024/12/13 ±N­ì¥»©ñ¦b¨Ó·½ªí³æªº³qª¾Email²¾¨ì¼Ò²Õ
      '¶l¥ó¥D¦®
      strR1 = strBC(1) & "-" & strBC(2) & IIf(Val(strBC(3) & strBC(4)) > 0, "-" & strBC(3) & "-" & strBC(4), "")
      strSub = "³qª¾¡I½ÐÀË¬d" & strR1 & "ªº¸ê®Æ§¨ÀÉ®×"
      '¦¬¥ó¤H
      strTo = PUB_GetFCPSalesNo(strBC(1), strBC(2), strBC(3), strBC(4))
      strCC = PUB_GetFCPProSup(strTo)
      mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
           " values( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
           ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & ChgSQL(strSub & vbCrLf & strErrMail) & "'," & strCC & ",'" & mCP09 & "')"
      cnnConnection.Execute mStrSql
      'end 2024/12/13
   End If
    
ErrHandle:
    If Err.Number <> 0 Then
        MsgBox "FCP·s®×¥ß¨÷(¤W¶ÇÀÉ®×)¡G" & vbCrLf & Err.Description
        Resume Next
    End If
    
End Sub

'Added by Lydia 2015/02/04 ©Ò¦³¤º³¡¦¬¤å, ­Y¦³¿é¤J¥»©Ò´Á­­©Îªk©w´Á­­ªÌ, ÀË¬d´Á­­¤£¥i¤p©ó¨t²Î¤é
'Modified by Lydia 2017/07/31 byval => byref
'Public Function PUB_CheckCP0607(ByVal cType As Integer, ByVal txt06 As String, ByVal txt07 As String) As Boolean
'Modify by Amy 2021/12/17 §ïForm2.0 «á·|¿ù,¬GTextBox->Control
'Move by Lydia 2023/11/08 ±qbasUpdate·h¹L¨Ó
'Modified by Lydia 2023/11/08 +strNewCase·s®×=Y, strNA01=¥Ó½Ð°ê, strCP01,strCP10=¨t²Î§O+®×¥ó©Ê½è
Public Function PUB_CheckCP0607(ByVal cType As Integer, ByRef tB06 As Control, ByRef tB07 As Control, ByVal strNewCase As String, ByVal strNA01 As String, ByVal strCP01 As String, strCP10 As String) As Boolean
'cType : 0 ¥iªÅ­È,1 ¤£¥iªÅ
Dim txt06 As String, txt07 As String
Dim strOurDeadline As String  'Added by Lydia 2023/11/08
    txt06 = "" & tB06.Text
    txt07 = "" & tB07.Text
'end 2017/07/31

    PUB_CheckCP0607 = False
    If cType = 1 Then
       If Len(txt06) = 0 Then
          ShowMsg "½Ð¿é¤J¥»©Ò´Á­­! "
       ElseIf Len(txt06) = 0 Then
          ShowMsg "½Ð¿é¤Jªk©w´Á­­! "
       End If
       If Len(txt06) = 0 Or Len(txt07) = 0 Then Exit Function
    End If

       If Len(txt06) > 0 Then
            If CheckIsTaiwanDate(txt06) Then
               txt06 = ChangeWDateStringToWString(txt06)
            ElseIf CheckIsDate(txt06) Then
               txt06 = ChangeWDateStringToWString(txt06)
            End If
            If Val(txt06) < strSrvDate(1) Then
               'Modified by Lydia 2015/02/24 ­ì¥»¬°¤£¥i¿é¤J¡A²{§ï¬°¼u°T®§¸ß°ÝUser
               'ShowMsg "¥»©Ò´Á­­¤£¥i¤p©ó¨t²Î¤é! " :Exit Function
               If MsgBox("¥»©Ò´Á­­¤p©ó¨t²Î¤é,¬O§_­n­«·s¿é¤J¡H", vbYesNo) = vbYes Then Exit Function
            End If
       End If
       If Len(txt07) > 0 Then
            If CheckIsTaiwanDate(txt07) Then
               txt07 = ChangeWDateStringToWString(txt07)
            ElseIf CheckIsDate(txt07) Then
               txt07 = ChangeWDateStringToWString(txt07)
            End If
            If Val(txt07) < strSrvDate(1) Then
               'Modified by Lydia 2015/02/24 ­ì¥»¬°¤£¥i¿é¤J¡A²{§ï¬°¼u°T®§¸ß°ÝUser
               'ShowMsg "ªk©w´Á­­¤£¥i¤p©ó¨t²Î¤é! " :Exit Function
               If MsgBox("ªk©w´Á­­¤p©ó¨t²Î¤é,¬O§_­n­«·s¿é¤J¡H", vbYesNo) = vbYes Then Exit Function
            End If
            'Added by Lydia 2017/08/08 ­Yªk©w´Á­­¦³­È¦ý¥»©Ò´Á­­¨S¦³­È®É,¥»©Ò´Á­­¹w³]¬°ªk©w´Á­­.
                                    '¦ý¥»©Ò´Á­­¦³­È®É,ªk©w´Á­­¥i¥H¨S¦³­È,¤£¯à¹w³].
            'Modified by Lydia 2023/11/08 ½Ð­×§ï¦³¿é¤Jªk©w´Á­­¦ý¨S¦³¿é¤J¥»©Ò´Á­­®É¡A§ï©I¥sService1 : PUB_CRLUseCP07CheckCP06¨Ó­pºâ¥»©Ò´Á­­¡C
            'If Len(txt06) = 0 Or Val(txt06) = 0 Then
            '   txt06 = TransDate(txt07, 1)
            If Len(txt06) = 0 Or Val(txt06) = 0 Then
               If PUB_CRLUseCP07CheckCP06(strNewCase, strNA01, strCP01, strCP10, txt06, txt07, strOurDeadline, False) = True Then
                  txt06 = TransDate(strOurDeadline, 1)
               End If
               If Len(txt06) = 0 Then txt06 = TransDate(txt07, 1)
            'end 2023/11/08
            End If
            'end 2017/08/08
       End If
       
       If Len(txt06) > 0 And Len(txt07) > 0 And Val(txt07) < Val(txt06) Then
           ShowMsg "ªk©w´Á­­¤£¥i¤p©ó¥»©Ò´Á­­! "
           Exit Function
       End If
    'End If
    
    'Added by Lydia 2017/07/31 ¦^¶Ç©Ò­­©Mªk­­
    If txt06 <> "" Then tB06.Text = TransDate(txt06, 1)
    If txt07 <> "" Then tB07.Text = TransDate(txt07, 1)
    'end 2017/07/31
    
    PUB_CheckCP0607 = True
End Function

'Added by Lydia 2024/12/13 ¼W¥[FCP/P/FG®×¸¹®Éªº¨t²Î³qª¾ (½Ð¹q¸£¤¤¤ß¤ñ·Óªþ¥ó·s®×¥ß¨÷PUB_GetTCTmail³qª¾©Ó¿ì¤Î¬ÛÃö¤H­û)
Private Sub Proc_FCPNewCaseEmail(ByVal pCP01 As String, pCP02 As String, pCP03 As String, pCP04 As String, pCP09 As String, pCP10 As String, pCP12 As String)
Dim intA As Integer, strA1 As String
Dim strB1 As String, strB2 As String, strB3 As String, strB4
Dim rsAD As New ADODB.Recordset
   
   If pCP01 = "FG" Or ((pCP01 = "FCP" Or (pCP01 = "P" And Left(pCP12, 2) = "F2")) And InStr(FcpNewCaseEmail, pCP10) > 0) Then
       strA1 = "select cp13,fa10 as pa75area,pa75, nvl(fa05,nvl(fa04,fa06)) as pa75n,cu10 pa26area, pa26, nvl(cu05,nvl(cu04,cu06)) as pa26n,decode(pa09,'000',nvl(cpm03,cpm04),nvl(cpm04,cpm03)) as cpm0304 " & _
                        "from caseprogress,patent, fagent, customer,casepropertymap where cp09='" & pCP09 & "' and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) " & _
                        "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) "
       strA1 = strA1 & " union select cp13,fa10 as sp26area,sp26, nvl(fa05,nvl(fa04,fa06)) as sp26n,cu10 sp08area, sp08, nvl(cu05,nvl(cu04,cu06)) as sp08n,decode(sp09,'000',nvl(cpm03,cpm04),nvl(cpm04,cpm03)) as cpm0304 " & _
                        " from caseprogress,servicepractice, fagent, customer,casepropertymap where cp09='" & pCP09 & "' and cp01=sp01(+) and cp02=sp02(+) and cp03=sp03(+) and cp04=sp04(+) " & _
                        " and substr(sp26,1,8)=fa01(+) and substr(sp26,9,1)=fa02(+) and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+) and cp01=cpm01(+) and cp10=cpm02(+) "
       intA = 1
       strB1 = "": strB2 = ""
       Set rsAD = ClsLawReadRstMsg(intA, strA1)
       If intA = 1 Then
           'FCPµ{§ÇºÞ¨î¤H©M¥DºÞ
           strB3 = PUB_GetFCPHandler(pCP01, pCP02, pCP03, pCP04)
           If strB3 <> "" Then
              strB4 = PUB_GetFCPProSup(strB3)
              strB1 = strB1 & ";" & strB3
              strB2 = strB2 & ";" & strB4
           End If
           'FCP©Ó¿ìºÞ¨î¤H©M¥DºÞ
           strB3 = PUB_GetFCPSalesNo(pCP01, pCP02, pCP03, pCP04)
           If strB3 <> "" Then
              strB4 = PUB_GetFCPProSup(strB3)
              strB1 = strB1 & ";" & strB3
              strB2 = strB2 & ";" & strB4
           End If
           '±NDavid¥[¤J(­^¤å²Õ)©Ò¦³FCP, PÂd¥x·s®×¥ß¨÷ (101-103)³qª¾¦¬¥ó¤H¤§¤@
           If "" & rsAD.Fields("pa75area") <> "" Then
              If Left("" & rsAD.Fields("pa75area"), 3) <> "011" Then
                  strB1 = strB1 & ";" & Pub_GetSpecMan("¥~±M©Ó¿ì­^¤å²Õ¥DºÞ")
              End If
           ElseIf "" & rsAD.Fields("pa26area") <> "" Then
              If Left("" & rsAD.Fields("pa26area"), 3) <> "011" Then
                  strB1 = strB1 & ";" & Pub_GetSpecMan("¥~±M©Ó¿ì­^¤å²Õ¥DºÞ")
              End If
           End If
           '±N¦¬¤å¤§´¼Åv¤H­û¦C¬°¦¬¥ó¤H
           If InStr(strB1 & ";" & strB2, rsAD.Fields("CP13")) = 0 And rsAD.Fields("CP13") <> "" Then
              strB1 = strB1 & ";" & rsAD.Fields("CP13")
           End If
           strB3 = "¥N²z¤H¡G" & IIf(rsAD.Fields("pa75") <> "", rsAD.Fields("pa75") & " " & rsAD.Fields("PA75N"), "¡]ªÅ¥Õ¡^") & vbCrLf & _
                        "¥Ó½Ð¤H¡G" & IIf(rsAD.Fields("pa26") <> "", rsAD.Fields("pa26") & " " & rsAD.Fields("PA26N"), "¡]ªÅ¥Õ¡^") & vbCrLf
           strB4 = pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 <> "", "-" & pCP03 & "-" & pCP04, "")
           If InStr("401ÅÜ§ó,935®×¥óÂà¦Ü¥»©Ò", pCP10) > 0 Then
               strB4 = strB4 & "®×¥óÂà¦Ü¥»©Ò¥ß¨÷"
           ElseIf InStr("605¦~¶O,", pCP10) > 0 Then
               strB4 = strB4 & rsAD.Fields("cpm0304") & "¥NÃº¥ß¨÷"
           Else
               strB4 = strB4 & rsAD.Fields("cpm0304") & "®×¥ß¨÷"
           End If
           'Modify By Sindy 2023/3/27 +,mc13
           mStrSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09,mc13)" & _
                " values( '" & strUserNum & "','" & Mid(strB1, 2) & "',to_char(sysdate,'yyyymmdd')" & _
                ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strB4) & "','" & ChgSQL(strB3) & "'," & CNULL(Mid(strB2, 2)) & ",'" & pCP09 & "')"
           cnnConnection.Execute mStrSql
       End If
   End If
   Set rsAD = Nothing
End Sub

'Move by Lydia 2024/12/13 ±qfrm010005·h¹L¨Ó
Public Function Pub_GetTCN01(ByVal pCP09 As String) As String
Dim strP1 As String, intP As Integer
Dim rsPD As New ADODB.Recordset

    Pub_GetTCN01 = ""
    strP1 = "Select TCN01 From TrackingCaseName Where TCN05='" & pCP09 & "' And TCN05<>'111111' Order by TCN01"
    If intP = 1 Then
       Pub_GetTCN01 = "" & rsPD.Fields("tcn01")
    End If
    Set rsPD = Nothing
End Function

'Added by Lydia 2024/12/13 ÀË¬dTracking No.©M´¼Åv¤H­û¬O§_¤@­P
Public Function Pub_ChkTCN01Status(ByVal pTCN01 As String, ByVal pSalesNo As String) As Boolean
Dim strP1 As String, intP As Integer
Dim rsPD As New ADODB.Recordset

   If Val(pTCN01) = 0 Then Exit Function

   '½T»{´¼Åv¤H­n¬°ºÞ¨î¤H©Î·s¼W¤Hªº¨ä¤¤¤@­Ó
   strP1 = "select tcn05 From TrackingCaseName Where TCN01=" & Val(pTCN01) & " and instr(TCN03||','||TCN06,'" & pSalesNo & "') > 0 "
   intP = 1
   Set rsPD = ClsLawReadRstMsg(intP, strP1)
   If intP = 0 Then
      MsgBox "°lÂÜ¬y¤ô¸¹¡G" & pTCN01 & " ¤§ºÞ¨î¤H©Î·s¼W¤H»P´¼Åv¤H¤£²Å¡I", vbCritical
      GoTo EXITSUB
   Else
      '§PÂ_¤£¥i¿é¤J¤w¦³¦¬¤å¸¹ªº°lÂÜ¬y¤ô¸¹
      If "" & rsPD.Fields("tcn05") <> "" Then
         MsgBox "°lÂÜ¬y¤ô¸¹¡G" & pTCN01 & " ¤w¦³¦¬¤å¸¹¡I", vbCritical
         GoTo EXITSUB
      Else
         Pub_ChkTCN01Status = True
      End If
   End If
   
EXITSUB:
   Set rsPD = Nothing
             
End Function

'Added by Morgan 2025/5/23
'Modified by Lydia 2025/06/19 §ï¦¨¥i«ü©wÄæ¦ì; ¿é¤J¨t²Î§O¡B®×¥ó©Ê½è¡B«ü©wÄæ¦ì¨ú±oCasePropertyMap®×¥ó©Ê½èªíªºÄæ¦ì­È
'Public Function PUB_GetCPM35byCP10(pCP01 As String, pCP10 As String) As String
Public Function PUB_GetCPMbyCP10(pCP01 As String, pCP10 As String, pFieldName As String) As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   'Added by Lydia 2025/06/19
   pFieldName = UCase(pFieldName)
   If InStr(pFieldName, "CPM") = 0 Then Exit Function
   'end 2025/06/19
   
   'Modified by Lydia 2025/06/19  cpm35=>pFieldName
   stSQL = "select " & pFieldName & " from casepropertymap where cpm01='" & pCP01 & "' and cpm02='" & pCP10 & "'"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      'Modified by Lydia 2025/06/19 PUB_GetCPM35byCP10=>PUB_GetCPMbyCP10
      PUB_GetCPMbyCP10 = "" & rsQuery(0)
   End If
   Set rsQuery = Nothing
End Function

