Attribute VB_Name = "basPublic"
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/5 日期欄已修改
Option Explicit

Public Const GWL_WNDPROC = (-4)
Public Const EM_GETLINE = &HC4F
Public Const EM_LINELENGTH = &HC1
Public Const EM_LINEINDEX = &HBB
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal MSG As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Sub RtlMoveMemory Lib "kernel32" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Public prevWndProc As Long
Public intRecount As Integer
Private Declare Function WideCharToMultiByte Lib "kernel32" ( _
    ByVal CodePage As Long, _
    ByVal dwFlags As Long, _
    ByVal lpWideCharStr As Long, _
    ByVal cchWideChar As Long, _
    ByVal lpMultiByteStr As Long, _
    ByVal cbMultiByte As Long, _
    ByVal lpDefaultChar As Long, _
    ByVal lpUsedDefaultChar As Long) As Long
Public m_strCP10Data(1 To 20) As String '案件性質先開陣列到20個

'Added by Morgan 2024/8/14 API for base64 encoding and decoding
Public Declare Function CryptBinaryToString Lib "crypt32" Alias "CryptBinaryToStringW" (ByVal pbBinary As Long, ByVal cbBinary As Long, ByVal dwFlags As Long, ByVal pszString As Long, pcchString As Long) As Long
Public Declare Function CryptStringToBinary Lib "crypt32" Alias "CryptStringToBinaryW" (ByVal pszString As Long, ByVal cchString As Long, ByVal dwFlags As Long, ByVal pbBinary As Long, pcbBinary As Long, Optional ByVal pdwSkip As Long, Optional ByVal pdwFlags As Long) As Long

'Memo by Lydia 2019/11/13 取得TF案和CFT案註冊證的繳費起算日及繳費年度
'取得繳費起算日及繳費年度，
'InPut : strNation : 國家代碼，strTxt() 0:收文號 1-4:本所案號
'OutPut : strDate 正確起算日期/專用起始日，strYear 繳費年度，strEnd 專用結束日
Public Function TFGetMoneyDate( _
   ByVal strNation As String, _
   ByRef strTxt() As String, _
   ByRef strDate As String, _
   ByRef strYear As String, _
   Optional ByRef strEnd As String) As Boolean
Dim i As Integer, rsTmp1 As New ADODB.Recordset
Dim intP As Integer, strConP As String, rsQ1 As New ADODB.Recordset 'Added by Lydia 2024/03/29

On Error GoTo ErrHand
   TFGetMoneyDate = False
   
   'add by nick 2004/12/16
   NickTmNa12 = 0
   
   strDate = ""
   strEnd = ""
   'Modified by Lydia 2019/11/13 +別名
   'strConP = "SELECT na12,'',NVL(na13,0) FROM NATION WHERE NA01=" + CNULL(strNation)
   strConP = "SELECT na12,'' as a02,NVL(na13,0) as na13 FROM NATION WHERE NA01=" + CNULL(strNation)
   intP = 1
   Set rsQ1 = ClsLawReadRstMsg(intP, strConP)
   With rsQ1
   If intP = 1 Then
      If Not IsNull(.Fields("na12")) Then
         i = 1
         Select Case .Fields("na12")
            Case 收文日
               strConP = "CP05"
               '910919 nick cft 用
               NickTmNa12 = 1
               NickTmNa13 = Val(CheckStr(.Fields("na13")))
            Case 申請日
               strConP = "TM11"
               i = 2
               '910919 nick cft 用
               NickTmNa12 = 2
               NickTmNa13 = Val(CheckStr(.Fields("na13")))
            Case 發文日
               strConP = "CP27"
               '910919 nick cft 用
               NickTmNa12 = 3
               NickTmNa13 = Val(CheckStr(.Fields("na13")))
            Case 准駁日
               strConP = "CP25"
               '910919 nick cft 用
               NickTmNa12 = 4
               NickTmNa13 = Val(CheckStr(.Fields("na13")))
            Case 公告日
               strConP = "TM14"
               i = 2
               '910919 nick cft 用
               NickTmNa12 = 5
               NickTmNa13 = Val(CheckStr(.Fields("na13")))
            Case 發證日
               strConP = "TM20"
               i = 2
               '910919 nick cft 用
               NickTmNa12 = 6
               NickTmNa13 = Val(CheckStr(.Fields("na13")))
         End Select
      Else
         TFGetMoneyDate = True
         '2008/11/25 ADD BY SONIA TF-000570
         If Not IsNull(.Fields("na13")) Then
            NickTmNa13 = Val(CheckStr(.Fields("na13")))
         End If
         '2008/11/25 END
         Exit Function
'         strConP = "TM11"
'         i = 2
      End If
      If i = 1 Then
         If strTxt(0) = "" Then
            strConP = "SELECT " & strConP & " FROM CASEPROGRESS WHERE CP09 IS NULL"
         Else
            strConP = "SELECT " & strConP & " FROM CASEPROGRESS WHERE CP09='" & strTxt(0) & "'"
         End If
      Else
         strConP = "SELECT " & strConP & " FROM TRADEMARK WHERE TM01='" & strTxt(1) & "' AND " & _
            "TM02='" & strTxt(2) & "' AND TM03='" & strTxt(3) & "' AND TM04='" & strTxt(4) & "'"
      End If
      intP = 1
      Set rsTmp1 = ClsLawReadRstMsg(intP, strConP) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intp, strConP)
      If intP = 1 Then
         If Not IsNull(rsTmp1.Fields(0)) Then
            strDate = "" & rsTmp1.Fields(0)
            '92.10.31 MODIFY BY SONIA
            'strEnd = Val(Left(strDate, 4)) + .Fields("na13") & Right(strDate, 4)
            Select Case strNation
               Case "013"   '香港新舊法專用期不同
                  If strDate < 20030404 Then
                     'Modified by Lydia 2019/11/13
                     'strEnd = Val(Left(strDate, 4)) + 7 & Right(strDate, 4)
                     strEnd = PUB_GetEndDate(strDate, 7, "N")
                  Else
                     'Modified by Lydia 2019/11/13 因為NA85尚未設定,一律設「計算商標專用期是否減1天」=N
                     'strEnd = Val(Left(strDate, 4)) + .Fields("na13") & Right(strDate, 4)
                     strEnd = PUB_GetEndDate(strDate, Val("" & .Fields("na13")), "N")
                  End If
               Case "238"   '馬德里新舊法專用期不同
                  If strDate < 19980101 Then
                     'Modified by Lydia 2019/11/13
                     'strEnd = Val(Left(strDate, 4)) + 20 & Right(strDate, 4)
                     strEnd = PUB_GetEndDate(strDate, 20, "N")
                  Else
                     'Modified by Lydia 2019/11/13
                     'strEnd = Val(Left(strDate, 4)) + .Fields("na13") & Right(strDate, 4)
                     strEnd = PUB_GetEndDate(strDate, Val("" & .Fields("na13")), "N")
                  End If
               'add by nickc 2005/08/02 '紐西蘭新舊法不同
               Case "016"
                  If strDate < 20030820 Then
                     'Modified by Lydia 2019/11/13
                     'strEnd = Val(Left(strDate, 4)) + 7 & Right(strDate, 4)
                     strEnd = PUB_GetEndDate(strDate, 7, "N")
                  Else
                     'Modified by Lydia 2019/11/13
                     'strEnd = Val(Left(strDate, 4)) + .Fields("na13") & Right(strDate, 4)
                     strEnd = PUB_GetEndDate(strDate, Val("" & .Fields("na13")), "N")
                  End If
               'add end
               'Add By Sindy 2012/11/27
               Case "113" '委內瑞拉
                  '註冊日期在2008年9月17日之前者,專用期限自註冊日起算10年
                  '註冊日期在2008年9月17日之後者,專用期限自註冊日起算15年(國家檔設定)
                  If strDate < 20080917 Then
                     'Modified by Lydia 2019/11/13
                     'strEnd = Val(Left(strDate, 4)) + 10 & Right(strDate, 4)
                     strEnd = PUB_GetEndDate(strDate, 10, "N")
                  Else
                     'Modified by Lydia 2019/11/13
                     'strEnd = Val(Left(strDate, 4)) + .Fields("na13") & Right(strDate, 4)
                     strEnd = PUB_GetEndDate(strDate, Val("" & .Fields("na13")), "N")
                  End If
               '2012/11/27 End
               Case Else
                  'Modified by Lydia 2019/11/13
                  'strEnd = Val(Left(strDate, 4)) + .Fields("na13") & Right(strDate, 4)
                  strEnd = PUB_GetEndDate(strDate, Val("" & .Fields("na13")), "N")
            End Select
            '92.10.31 END
         End If
      End If
      TFGetMoneyDate = True
   End If
   End With
   'Added by Lydia 2024/03/29
   Set rsTmp1 = Nothing
   Set rsQ1 = Nothing
   'end 2024/03/29
   
   Exit Function
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'Memo by Lydia 2019/11/13 取得CFT案延展後專用期止日
'取得延展後專用期止日
'InPut : strNation : 國家代碼，strTxt() 0:收文號 1-4:本所案號
'OutPut : strEnd 延展後專用期止日，strPreEnd 延展前專用期止日
Public Function CFTGetNewDate( _
   ByVal strNation As String, _
   ByRef strTxt() As String, _
   ByRef strDate As String, _
   Optional ByRef strEnd As String, Optional ByRef strPreEnd As String) As Boolean
Dim i As Integer, rsTmp1 As New ADODB.Recordset
Dim intP As Integer, strConP As String, rsQ1 As New ADODB.Recordset 'Added by Lydia 2024/03/29

On Error GoTo ErrHand
   CFTGetNewDate = False
   
   strDate = ""
  strEnd = "": strPreEnd = ""
   
   strConP = "SELECT NA14 FROM NATION WHERE NA01=" + CNULL(strNation)
   intP = 1
   Set rsQ1 = ClsLawReadRstMsg(intP, strConP) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intp, strConP)
   With rsQ1
   If intP = 1 Then
      '2009/7/1 MODIFY BY SONIA 因98/6/18之需求於延展代理人已提申時先更新專用期間,故此情形以CP54為新專用期止日
      'strConP = "SELECT TM21,TM22,TM10 FROM TRADEMARK WHERE TM01='" & strTxt(1) & "' AND " & _
      '   "TM02='" & strTxt(2) & "' AND TM03='" & strTxt(3) & "' AND TM04='" & strTxt(4) & "'"
      'modify by sonia 2014/10/31 加CP53延展前專用期止日CFT-14866莫三比克延展後五年要提使用宣誓
      strConP = "SELECT TM21,TM22,TM10,CP47,CP54,TM20,CP53 FROM TRADEMARK,CASEPROGRESS WHERE TM01='" & strTxt(1) & "' AND " & _
         "TM02='" & strTxt(2) & "' AND TM03='" & strTxt(3) & "' AND TM04='" & strTxt(4) & "' AND " & _
         "TM01=CP01(+) AND TM02=CP02(+) AND TM03=CP03(+) AND TM04=CP04(+) AND '" & strTxt(0) & "'=CP09(+)"
      '2009/7/1 END
      intP = 1
      Set rsTmp1 = ClsLawReadRstMsg(intP, strConP) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intp, strConP)
      If intP = 1 Then
         If Not IsNull(rsTmp1.Fields("TM21")) Then
            strDate = rsTmp1.Fields("TM21")
         End If
         'add by sonia 2014/10/31
         If Not IsNull(rsTmp1.Fields("CP53")) Then
            strPreEnd = rsTmp1.Fields("CP53")
         End If
         'end 2014/10/31
         If Not IsNull(rsTmp1.Fields("TM22")) Then
            'Modified by Lydia 2019/11/13 因為NA85尚未設定,一律設「計算商標專用期是否減1天」=N
            'strEnd = Val(Left(rsTmp1.Fields("TM22"), 4)) + .Fields("NA14") & Right(rsTmp1.Fields("TM21"), 4)
            'Modified by Lydia 2019/11/21 CFT是用專用期限止+延展年
            'strEnd = PUB_GetEndDate(rsTmp1.Fields("TM21"), Val("" & .Fields("NA14")), "N")
            'Modified by Lydia 2019/12/06 CFT是用專用期限止+延展年=>取年份+專用期限起日的月日
            'strEnd = PUB_GetEndDate(rsTmp1.Fields("TM22"), Val("" & .Fields("NA14")), "N")
            strEnd = Left(PUB_GetEndDate(rsTmp1.Fields("TM22"), Val("" & .Fields("NA14")), "N"), 4) & Right(rsTmp1.Fields("TM21"), 4)
            '2009/3/27 MODIFY BY SONIA 葉易雲說2004年4月12日前期滿應延展案件為15年,之後為10年(國家檔設定)
            If rsTmp1.Fields("TM10") = "038" And rsTmp1.Fields("TM22") < 20040412 Then
               'Modified by Lydia 2019/11/13
               'strEnd = Val(Left(rsTmp1.Fields("TM22"), 4)) + 15 & Right(rsTmp1.Fields("TM21"), 4)
               'Modified by Lydia 2019/11/21 CFT是用專用期限止+延展年
               'strEnd = PUB_GetEndDate(rsTmp1.Fields("TM21"), 15, "N")
               'Modified by Lydia 2019/12/06 CFT是用專用期限止+延展年=>取年份+專用期限起日的月日
               'strEnd = PUB_GetEndDate(rsTmp1.Fields("TM22"), 15, "N")
               strEnd = Left(PUB_GetEndDate(rsTmp1.Fields("TM22"), 15, "N"), 4) & Right(rsTmp1.Fields("TM21"), 4)
            '2009/3/27 END
            'Add By Sindy 2012/11/27
            ElseIf rsTmp1.Fields("TM10") = "113" And rsTmp1.Fields("TM20") < 20080917 Then
               '註冊日期在2008年9月17日之前者,專用期限自註冊日起算10年
               '註冊日期在2008年9月17日之後者,專用期限自註冊日起算15年(國家檔設定)
               'Modified by Lydia 2019/11/13
               'strEnd = Val(Left(rsTmp1.Fields("TM22"), 4)) + 10 & Right(rsTmp1.Fields("TM21"), 4)
               'Modified by Lydia 2019/11/21 CFT是用專用期限止+延展年
               'strEnd = PUB_GetEndDate(rsTmp1.Fields("TM21"), 10, "N")
               'Modified by Lydia 2019/12/06 CFT是用專用期限止+延展年=>取年份+專用期限起日的月日
               'strEnd = PUB_GetEndDate(rsTmp1.Fields("TM22"), 10, "N")
               strEnd = Left(PUB_GetEndDate(rsTmp1.Fields("TM22"), 10, "N"), 4) & Right(rsTmp1.Fields("TM21"), 4)
            '2012/11/27 End
            End If
            '2009/7/1 MODIFY BY SONIA 因98/6/18之需求於延展代理人已提申時先更新專用期間,故此情形以CP54為新專用期止日
            If Not IsNull(rsTmp1.Fields("CP47")) Then
               strEnd = "" & rsTmp1.Fields("CP54")
            End If
            '2009/7/1 END
         End If
      End If
      CFTGetNewDate = True
   End If
   End With
   'Added by Lydia 2024/03/29
   Set rsTmp1 = Nothing
   Set rsQ1 = Nothing
   'end 2024/03/29
   
   Exit Function
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'申請案號檢查 iSitu=0 國內 iSitu=1 大陸
'2005/6/9 MODIFY BY SONIA 加判斷卷宗性質
'Public Function ChkAppNo(ByVal strTemp As String, ByVal PA08 As Integer, Optional iSitu As Integer = 0) As Boolean
'Modify by Morgan 2010/8/20 +pa10
'Modified by Morgan 2012/10/8 +CP10 判斷追加,聯合或衍生
Public Function ChkAppNo(ByVal strTemp As String, ByVal PA08 As Integer, Optional iSitu As Integer = 0, Optional pa23 As Integer = 0, Optional PA10 As String, Optional CP10 As String) As Boolean
   Dim i As Integer, bolChk As Boolean, j As Integer, strWork As String, stErrMsg As String
  
   strTemp = UCase(strTemp)
   bolChk = True
   If iSitu = 0 Then
      'Add by Morgan 2010/8/20
      '國內專利申請號改輸9碼
      If bolNewAppNoFormat Then
         stErrMsg = ""
         'Modified by Morgan 2025/2/19 +15碼(Ex:衍生設計被舉發 P-135182)
         If Len(strTemp) = 9 Or Len(strTemp) = 12 Or Len(strTemp) = 15 Then
            '前9碼必須為數字
            For i = 1 To 9
               If Not IsNumeric(Mid(strTemp, i, 1)) Then
                  stErrMsg = "( 前9碼必須為數字 )"
                  bolChk = False
                  GoTo A0
               End If
            Next
            
            If PA10 <> "" Then
               '非改請前3碼必須為申請年度
               If Val(Left(strTemp, 3)) > (TransDate(PA10, 1) \ 10000) Then
                  stErrMsg = "( 前3碼必須為申請年度 )"
                  bolChk = False
                  GoTo A0
               End If
            Else
               '前3碼為年度且不可大於系統年
               If Val(Left(strTemp, 3)) > (strSrvDate(2) \ 10000) Then
                  stErrMsg = "( 前3碼為年度且不可大於系統年 )"
                  bolChk = False
                  GoTo A0
               End If
            End If
            '第4碼必須為專利種類
            If Mid(strTemp, 4, 1) <> Format(PA08) Then
               stErrMsg = "( 第4碼必須為專利種類 )"
               bolChk = False
               GoTo A0
            End If
            'Added by Morgan 2012/10/8
            '+衍生設計(申請案號輸入-有傳案件性質)
            If CP10 = "125" Then
               If Mid(strTemp, 10, 1) <> "D" Then
                  stErrMsg = "( 衍生設計案第10碼必須為 D )"
                  bolChk = False
                  GoTo A0
               End If
            Else
            'end 2012/8/10
            
               'Added by Morgan 2025/2/19 衍生設計也可能被舉發 EX:P135182--玲玲
               If PA08 = "3" And Mid(strTemp, 10, 1) = "D" Then
                  If Len(strTemp) = 15 Then
                     If pa23 = 3 Then
                        If Mid(strTemp, 13, 1) <> "N" Then
                           stErrMsg = "( 衍生設計的舉發案第13碼必須為 N )"
                           bolChk = False
                           GoTo A0
                        End If
                     Else
                        stErrMsg = "( 申請案號必須為12碼 )"
                        bolChk = False
                        GoTo A0
                     End If
                  End If
               'end 2025/2/19
               ElseIf Len(strTemp) = 12 Then
                     '異議
                  If pa23 = 2 Then
                     If Mid(strTemp, 10, 1) <> "P" Then
                        stErrMsg = "( 異議案第10碼必須為 P )"
                        bolChk = False
                        GoTo A0
                     End If
                  '舉發
                  ElseIf pa23 = 3 Then
                     If Mid(strTemp, 10, 1) <> "N" Then
                        stErrMsg = "( 舉發案第10碼必須為 N )"
                        bolChk = False
                        GoTo A0
                     End If
                  '追加,聯合
                  Else
                     'Modified by Morgan 2012/10/8
                     '+衍生設計(基本資料維護-沒傳案件性質)
                     If Mid(strTemp, 10, 1) <> "A" And Mid(strTemp, 10, 1) <> "U" And Mid(strTemp, 10, 1) <> "D" Then
                        stErrMsg = "( 追加、聯合或衍生設計案第10碼必須為 A 、 U 或 D )"
                        bolChk = False
                        GoTo A0
                     End If
                  End If
               End If
            End If
         Else
            stErrMsg = "( 申請案號必須為9或12碼 )"
            bolChk = False
            GoTo A0
         End If
      Else
      'end 2010/8/20
      
         If Len(strTemp) = 8 Then
            For i = 1 To 8
               If Not IsNumeric(Mid(strTemp, i, 1)) Then
                  bolChk = False
                  GoTo A0
               End If
            Next
         ElseIf Len(strTemp) = 11 Then
            For i = 1 To 11
               If i = 9 Then
                  If Mid(strTemp, i, 1) <> "A" And Mid(strTemp, i, 1) <> "U" Then
                     '2005/6/9 MODIFY BY SONIA
                     'bolChk = False
                     'GoTo A0
                     Select Case pa23
                     Case 1
                        bolChk = False
                        GoTo A0
                     Case 2
                        If Mid(strTemp, i, 1) <> "P" Then
                           bolChk = False
                           GoTo A0
                        End If
                     Case 3
                        If Mid(strTemp, i, 1) <> "N" Then
                           bolChk = False
                           GoTo A0
                        End If
                     Case Else
                        bolChk = False
                        GoTo A0
                     End Select
                     '2005/6/9 END
                  End If
               Else
                  If Not IsNumeric(Mid(strTemp, i, 1)) Then
                     bolChk = False
                     GoTo A0
                  End If
               End If
            Next
         Else
            bolChk = False
            GoTo A0
         End If
         'Modify by Morgan 2003/12/31
         'If Left(strTemp, 2) > GetTaiwanThisYear Then
         If Left(strTemp, 2) > (strSrvDate(2) \ 10000) Then
            bolChk = False
            GoTo A0
         End If
         
         If Mid(strTemp, 3, 1) <> Format(PA08) Then
            bolChk = False
            GoTo A0
         End If
      End If
   Else
      If iSitu = 1 Then
         If Len(strTemp) = 10 Then
            For i = 1 To 8
               If Not IsNumeric(Mid(strTemp, i, 1)) Then
                  bolChk = False
                  GoTo A0
               End If
            Next
            If Mid(strTemp, 9, 1) <> "." Then
               bolChk = False
               GoTo A0
            End If
            j = 0
            For i = 1 To 8
               j = j + (i + 1) * Val(Mid(strTemp, i, 1))
            Next
            'If Format(j Mod 11) <> Right(strTemp, 1) Then
            'modify by sonia 90.12.14
            strWork = Format(j Mod 11)
            If strWork = 10 Then strWork = "X"
            If strWork <> Right(strTemp, 1) Then
               bolChk = False
               GoTo A0
            End If
         Else
         '92.10.23 MODIFY BY SONIA
         '   bolChk = False
         '   GoTo A0
         'End If
            If Len(strTemp) = 14 Then
               For i = 1 To 12
                  If Not IsNumeric(Mid(strTemp, i, 1)) Then
                     bolChk = False
                     GoTo A0
                  End If
               Next
               
               'Add by Morgan 2006/3/7 12碼檢查
               j = 0
               For i = 1 To 8
                  j = j + (i + 1) * Val(Mid(strTemp, i, 1))
               Next
               For i = 9 To 12
                  j = j + (i - 7) * Val(Mid(strTemp, i, 1))
               Next
               strWork = Format(j Mod 11)
               If strWork = 10 Then strWork = "X"
               If strWork <> Right(strTemp, 1) Then
                  bolChk = False
                  GoTo A0
               End If
               '2006/3/7 end
               
               If Mid(strTemp, 13, 1) <> "." Then
                  bolChk = False
                  GoTo A0
               End If
               'Add by Morgan 2004/2/3
               '大陸申請案號 92.10.1 後為14碼且前4碼為西元年後第5碼為專利種類
               If Left(strTemp, 4) < 2003 Or Left(strTemp, 4) > Left(strSrvDate(1), 4) Then
                   bolChk = False
                   GoTo A0
               ElseIf Mid(strTemp, 5, 1) <> Format(PA08) Then
                   bolChk = False
                   GoTo A0
               End If
               'Add End 2004/2/3
            Else
               bolChk = False
               GoTo A0
            End If
         End If
      End If
   End If
   
A0:
   If bolChk = False Then
      ChkAppNo = False
      MsgBox "申請案號錯誤，請重新輸入 !" & IIf(stErrMsg <> "", vbCrLf, "") & stErrMsg, vbCritical
   Else
      ChkAppNo = True
   End If
End Function

Public Function MergeString(ByVal strTmp1 As String, ByVal strTmp2 As String, ByVal strTmp3 As String, ByVal strTmp4 As String) As String
   If strTmp3 = "0" And strTmp4 = "00" Then
      MergeString = strTmp1 & " - " & strTmp2
   Else
      MergeString = strTmp1 & " - " & strTmp2 & " - " & strTmp3 & " - " & strTmp4
   End If
End Function

Public Sub InitGrid(ByVal iRow As Integer, MsGrid As MSHFlexGrid)
 Dim strTmp As String, intA As Integer
 Dim intQ As Integer, strConQ As String, rsQ1 As New ADODB.Recordset 'Added by Lydia 2024/03/29
 
   strTmp = ""
   For intA = 1 To iRow
      ' 90.07.24 modify by louis
      'strTmp = strTmp & "'',"
      strTmp = strTmp & "' ',"
   Next
   If Right(strTmp, 1) = "," Then strTmp = Left(strTmp, Len(strTmp) - 1)
   strConQ = "SELECT " & strTmp & " FROM DUAL WHERE ROWNUM<1"
   intQ = 1
   Set rsQ1 = ClsLawReadRstMsg(intQ, strConQ)
   If intQ <> 2 Then
     Set MsGrid.Recordset = rsQ1
   End If
   Set rsQ1 = Nothing 'Added by Lydia 2024/03/29
End Sub

Public Sub OnlyOneRec(MsGrid As MSHFlexGrid, ByVal iCol As Integer)
   MsGrid.row = 1
   MsGrid.col = iCol
   MsGrid.Text = "v"
End Sub


'年費單筆不跑
Public Function CU73FA40(ByVal pa26 As String, ByVal pa75 As String, Optional ByVal pa76 As String) As Boolean
 Dim adoRst As New ADODB.Recordset
 Dim intQ As Integer, strConQ As String 'Added by Lydia 2024/03/29
   CU73FA40 = False

   If pa76 <> "" Then
      strConQ = "SELECT FA40 FROM FAGENT WHERE " & ChgFagent(pa76)
   ElseIf pa75 <> "" Then
      strConQ = "SELECT FA40 FROM FAGENT WHERE " & ChgFagent(pa75)
   Else
      strConQ = "SELECT CU73 FROM CUSTOMER WHERE " & ChgCustomer(pa26)
   End If
   intQ = 1
   Set adoRst = ClsLawReadRstMsg(intQ, strConQ)
   If intQ = 1 Then
      If adoRst.Fields(0) = "Y" Then CU73FA40 = True
   End If
   Set adoRst = Nothing 'Added by Lydia 2024/03/29
End Function

'聯絡人 1 2
Public Function PA51CU58FA07(pa() As String) As Boolean
 Dim adoRst As New ADODB.Recordset, i As Integer, bolChk As Boolean
 Dim intQ As Integer, strConQ As String 'Added by Lydia 2024/03/29
   
   PA51CU58FA07 = False
   bolChk = False
   
   If pa(1) = "FCP" Then
      Select Case pa(85)
         Case 1
            If pa(51) = "" And pa(54) = "" Then bolChk = True
         Case 2
            If pa(52) = "" And pa(55) = "" Then bolChk = True
         Case 3
            If pa(53) = "" And pa(56) = "" Then bolChk = True
      End Select
      If bolChk Then
         'Modify by Morgan 2006/10/18 非寄送用，日文聯絡人不必帶部門
         If pa(75) = "" Then
            strConQ = "SELECT CU58,CU59,CU60,CU61,CU62,CU63 FROM CUSTOMER WHERE " & ChgCustomer(pa(26))
         Else
            strConQ = "SELECT FA07,FA08,FA09,FA52,FA53,FA54 FROM FAGENT WHERE " & ChgFagent(pa(75))
         End If
         intQ = 1
         Set adoRst = ClsLawReadRstMsg(intQ, strConQ)
         If intQ = 1 Then
            For i = 0 To 5
                pa(51 + i) = "" & adoRst.Fields(i)
            Next
         End If
      End If
   Else
      If pa(30) = "" Then bolChk = True
      If bolChk Then
         If pa(26) = "" Then
            strConQ = "SELECT CU58 FROM CUSTOMER WHERE " & ChgCustomer(pa(8))
         Else
            strConQ = "SELECT FA07 FROM FAGENT WHERE " & ChgFagent(pa(26))
         End If
         intQ = 1
         Set adoRst = ClsLawReadRstMsg(intQ, strConQ) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intq, strconq)
         If intQ = 1 Then
            If Not IsNull(adoRst.Fields(0)) Then
               pa(30) = ""
            Else
               pa(30) = adoRst.Fields(0)
            End If
         End If
      End If
   End If
   
   PA51CU58FA07 = True
   Set adoRst = Nothing 'Added by Lydia 2024/03/29
End Function

'新增資料至1K0
Public Function SaveNew1K0(ByRef A1K() As String) As Boolean
   Dim strTxt(1) As String, i As Integer
 On Error GoTo ErrHand
   A1K(3) = ChangeCustomerL(A1K(3))
    'Modify By Cheng 2002/12/24
'   strTxt(1) = "insert into 1k0 values ("
   'Modified by Morgan 2017/6/20 要和下面的值欄位一致
   'strTxt(1) = "insert into ACC1k0 (A1K01,A1K02,A1K03,A1K04,A1K05,A1K06,A1K07,A1K08,A1K09,A1K10," & _
                    "A1K11,A1K12,A1K13,A1K14,A1K15,A1K16,A1K17,A1K18,A1K19,A1K20," & _
                    "A1K21,A1K22,A1K23,A1K24,A1K25,A1K26,A1K27,A1K28,A1K29,A1K30) values ("
   strTxt(1) = "insert into ACC1k0 (A1K01"
   For i = 2 To TF_1K0
      strTxt(1) = strTxt(1) & ",A1K" & Format(i, "00")
   Next
   strTxt(1) = strTxt(1) & ") values ("
   'end 2017/6/20
   
   'Modified by Morgan 2017/6/
   For i = 1 To TF_1K0 - 1 'edit by nickc 2007/02/02 T_1K0 - 1
      strTxt(1) = strTxt(1) + CNULL(A1K(i)) + ","
   Next
   strTxt(1) = strTxt(1) + CNULL(A1K(i)) + ")"
   'edit by nickc 2007/02/05 不用 dll 了
   'If objLawDll.ExecSQL(1, strTxt) Then SaveNew1K0 = True
   If ClsLawExecSQL(1, strTxt) Then SaveNew1K0 = True
   Exit Function
ErrHand:
   ShowMsg MsgText(9130)
   SaveNew1K0 = False
End Function

Public Function GetNP07(ByVal NA01 As String, PA08 As String, NP07) As Boolean
 Dim strTmp As String
 Dim intQ As Integer, strConQ As String, rsQ1 As New ADODB.Recordset 'Added by Lydia 2024/03/29
 
   NP07 = ""
   GetNP07 = False
   Select Case PA08
      Case 1
         strTmp = "NA07,NA20"
      Case 2
         strTmp = "NA09,NA22"
      Case 3
         strTmp = "NA11,NA24"
   End Select
   strConQ = "SELECT " & strTmp & " FROM NATION WHERE NA01='" & NA01 & "'"
   intQ = 1
   Set rsQ1 = ClsLawReadRstMsg(intQ, strConQ)
   If intQ = 1 And Not IsNull(rsQ1.Fields(0)) Then
      If IsNull(rsQ1.Fields(1)) Then

      Else
         NP07 = rsQ1(1)
         GetNP07 = True
      End If
   Else
      MsgBox "此國家無此專利種類 !", vbCritical
   End If
   
   Set rsQ1 = Nothing 'Added by Lydia 2024/03/29
End Function

'再次確定是否閉卷
'Modify by Amy 2018/08/29 +stMsg 參數
Public Function CheckCloseFile(Optional ByVal stMsgTxt As String = "") As Boolean
Dim stMsg As String 'Add by Amy 2018/08/29

'Add by Amy 2018/08/29
stMsg = "是否確定閉卷？"
If stMsgTxt <> "" Then stMsg = stMsgTxt
'end 2018/08/29
If MsgBox(stMsg, vbInformation + vbDefaultButton1 + vbOKCancel) = vbOK Then
   CheckCloseFile = True
Else
   CheckCloseFile = False
End If
End Function

'將資料放入Grid內
Public Sub MakeGrdData(ByRef rsResTemp As ADODB.Recordset, ByRef grdTemp As MSFlexGrid)
Dim i As Integer, j As Integer

On Error GoTo ErrHand
If Not rsResTemp.EOF Then
   Do While Not rsResTemp.EOF
          i = i + 1
          grdTemp.Rows = i + 1
          For j = 0 To rsResTemp.Fields.Count - 1
                 grdTemp.TextMatrix(i, j) = IIf(IsNull(rsResTemp.Fields(j)), "", rsResTemp.Fields(j))
          Next
          rsResTemp.MoveNext
          grdTemp.MergeRow(i) = True
   Loop
Else
   grdTemp.Rows = 1
End If
Exit Sub
ErrHand:
End Sub

'取得 UserID
Public Function GetUserID() As String
Dim lngRt As Long, strUser As String * 100, strUserNum As String

lngRt = WNetGetUser("", strUser, 100)
If lngRt = 0 Then
   GetUserID = UCase(MyTrim(strUser))
End If
End Function
'取得案件種類
Public Function GetCaseNumSysKind(ByVal strCP As String) As String
Dim i As Long
For i = 1 To Len(strCP)
     If Mid(strCP, i, 1) = "-" Then
        GetCaseNumSysKind = Mid(strCP, 1, i - 1)
        Exit For
     End If
 Next
End Function
'移動線段
Public Function MoveLine(ByRef strlc As String) As String
Dim i As Long
For i = 1 To Len(strlc)
     If Mid(strlc, i, 1) <> "-" Then
        MoveLine = MoveLine + Mid(strlc, i, 1)
     End If
 Next
End Function

'取得新增或修改日期
Public Function GetCreateUpdateDate(Id As String, dt As String, tm As String) As Boolean
GetCreateUpdateDate = True
 Id = GetUserID
 Id = "SnowFi"
 'edit by nickc 2006/03/17
 'dt = Trim(str(Val(ChangeWDateStringToTString(Date)) + 19110000))
 dt = Trim(strSrvDate(1))
 tm = Trim(GetTime)
End Function

'選擇符號切換
'Modify by Amy 2021/07/09 +intChkCol as integer
Public Function CheckGridChoese(ByRef grdTemp As MSHFlexGrid, ByRef intLastRow As Integer, ByVal intCols As Integer, Optional intChkCol As Integer = 1) As Boolean
ShowBar grdTemp, intLastRow, intCols
With grdTemp
    .col = intChkCol 'Moidfy by Amy 2021/07/09 原:1
         If .Text = "" Then CheckGridChoese = False: Exit Function
    .col = 0
        If .Text = "v" Then
           .Text = ""
        Else
           .Text = "v"
        End If
        CheckGridChoese = True
End With
End Function

'Add By Cheng 2002/06/10
'Copy From prjTaieLawDll : GetNextPayDate
'取得下次繳費日,利用array傳值,array(1-4)為本所案號,
'若系統類別為"P"時, 則帶出本所期限, 否則帶出法定期限
Public Function PUB_GetNextPayDate(ByRef pa() As String, ByRef strName As String) As Boolean
Dim strQty As String
Dim RsAdo As New ADODB.Recordset
On Error GoTo ErrHand
   PUB_GetNextPayDate = False
   strName = ""
   If pa(1) = "P" Then
      strQty = "SELECT MAX(NP08) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "')"
   Else
      strQty = "SELECT MAX(NP09) FROM NEXTPROGRESS WHERE " & ChgNextProgress(pa(1) & pa(2) & pa(3) & pa(4)) & _
         " AND NP07 IN ('" & 年費 & "','" & 維持費 & "','" & 延展費 & "')"
   End If
   RsAdo.Open strQty, cnnConnection
   Do While Not RsAdo.EOF
      If Not IsNull(RsAdo.Fields(0)) Then strName = RsAdo.Fields(0)
      PUB_GetNextPayDate = True
      Exit Do
   Loop
   RsAdo.Close
   Exit Function
ErrHand:
   MsgBox "錯誤 : " & Err.Description, vbCritical
End Function

'Add By Cheng 2002/07/04
Public Function PUB_GetLawDay(dblDate As Double) As Double
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String

PUB_GetLawDay = dblDate - 19110000
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
StrSQLa = "Select * From WorkDay Where WD01>=" & dblDate & " Order By WD01 "
rsA.CursorLocation = adUseClient
'Add by Morgan 2003/12/31
rsA.MaxRecords = 1
   
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   PUB_GetLawDay = rsA.Fields(0).Value - 19110000
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add By Cheng 2002/07/12
'以本所案號來查詢案件是否閉卷
'Modify By Sindy 2014/7/28 +bolIsExists 檢查此本所案號是否存在
'                          +bolIsChkClose 是否檢查閉卷狀況
Public Function PUB_CaseClosed(Str01 As String, Str02 As String, Str03 As String, Str04 As String, _
                               Optional ByRef bolIsExists As Boolean = False, Optional bolIsChkClose As Boolean = True) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

PUB_CaseClosed = False
bolIsExists = False 'Add By Sindy 2022/10/21

If Str03 = "" Then Str03 = "0"
If Str04 = "" Then Str04 = "00"
StrSQLa = "SELECT PA57 FROM PATENT WHERE PA01='" & Str01 & "' AND PA02='" & Str02 & "' AND PA03='" & Str03 & "' AND PA04='" & Str04 & "'" & _
         " UNION SELECT TM29 FROM TRADEMARK WHERE TM01='" & Str01 & "' AND TM02='" & Str02 & "' AND TM03='" & Str03 & "' AND TM04='" & Str04 & "'" & _
         " UNION SELECT LC08 FROM LAWCASE WHERE LC01='" & Str01 & "' AND LC02='" & Str02 & "' AND LC03='" & Str03 & "' AND LC04='" & Str04 & "'" & _
         " UNION SELECT HC09 FROM HIRECASE WHERE HC01='" & Str01 & "' AND HC02='" & Str02 & "' AND HC03='" & Str03 & "' AND HC04='" & Str04 & "'" & _
         " UNION SELECT SP15 FROM SERVICEPRACTICE WHERE SP01='" & Str01 & "' AND SP02='" & Str02 & "' AND SP03='" & Str03 & "' AND SP04='" & Str04 & "'"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   'Add By Sindy 2014/7/28
   bolIsExists = True
   If bolIsChkClose = True Then
   '2014/7/28 END
      If "" & rsA.Fields(0).Value = "Y" Then
         MsgBox "此案件已閉卷, 不可執行發文作業!!!", vbExclamation + vbOKOnly
         PUB_CaseClosed = True
      End If
   End If
Else
   MsgBox "無此案件基本資料, 請檢查是否案號輸入錯誤!!!", vbExclamation + vbOKOnly
   PUB_CaseClosed = True
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'取得來函記錄檔之系統類別
Public Function GetCkindSys(ByRef strDate As String, ByRef strSys() As String) As Boolean
Dim strSql As String, i As Integer, rsRecordset As New ADODB.Recordset

On Error GoTo ErrHand
strSql = "select distinct mr12 from mailrec where mr02=" + strDate
Set rsRecordset = New ADODB.Recordset
rsRecordset.CursorLocation = adUseClient
rsRecordset.Open strSql, cnnConnection
If rsRecordset.RecordCount > 0 Then
   Do While Not rsRecordset.EOF
         ReDim Preserve strSys(i) As String
         strSys(i) = rsRecordset.Fields(0)
         rsRecordset.MoveNext
         i = i + 1
   Loop
   GetCkindSys = True
Else
   MsgBox "Mailrec檔案無資料!!", vbCritical
End If
rsRecordset.Close
Exit Function
ErrHand:
MsgBox "讀取Mailrec檔案失敗!!", vbCritical
End Function
'Add by Morgan 2004/3/15
'Copy from prjTaieDll003.cls003
'讀取國內外案件關聯表之Recordset
'iSitu 0 國內外案件資料維護資料(CFP)
'edit by nickc 2005/06/23
'Public Function ReadCaseRelationRst(ByRef strEnginer As String, ByRef StrDate1 As String, ByRef StrDate2 As String) As ADODB.Recordset
Public Function ReadCaseRelationRst(ByRef strEnginer As String, ByRef strDate1 As String, ByRef StrDate2 As String, Optional iSitu As Integer = 0) As ADODB.Recordset
    Dim i As Integer, rsRecordset As New ADODB.Recordset
    Dim strSql As String, strSQL1 As String, StrSQL3 As String
    Dim m_Title As String 'Added by Lydia 2015/07/27
On Error GoTo ErrHand
    'Add by Lydia 2014/10/31 開放外專程序人員(31,33,34)可進入專利處系統操作FMP寰華案件，但非此類案件時外專程序人員不可操作。
     Dim midSql As String
     If FMP2open = True Then
        '國內外案件關聯判斷國外案號(cm01~cm04),香港大陸案判斷香港案號(cm01~cm04)
        midSql = Replace(FMP2openSQL, "f0.CP", "pa1.PA")
     End If
   Select Case iSitu
      Case "0" '國內外
         '國外案不抓聯合申請
         'edit by nickc 2005/10/24
         'strSQL1 = " and cp1.cp10 in (" + CNULL(發明申請) + "," + CNULL(新型申請) + "," + CNULL(設計申請) + "," + CNULL(追加申請) + "," + CNULL(翻譯) + ",'" & 113 & "','" & 114 & "','" & 307 & "')"
         '2006/3/7 MODIFY BY SONIA 加案件性質 109
         'Modify by Morgan 2010/10/8 +聯合申請改收不同案號故要加回
         'strSQL1 = " and cp1.cp10 in (" + CNULL(發明申請) + "," + CNULL(新型申請) + "," + CNULL(設計申請) + "," + CNULL(追加申請) + "," + CNULL(翻譯) + ",'" & 113 & "','" & 114 & "','" & 307 & "','109','110','112')"
         strSQL1 = " and cp1.cp10 in (" + CNULL(發明申請) + "," + CNULL(新型申請) + "," + CNULL(設計申請) + "," + CNULL(追加申請) + "," + CNULL(聯合申請) + "," + CNULL(翻譯) + ",'" & 113 & "','" & 114 & "','" & 307 & "','109','110','112')"
         
         If strEnginer <> "" Then
            strSQL1 = strSQL1 & " and cp1.cp14=" + CNULL(strEnginer)
         End If
         
         'StrSQL3 = " and cp2.cp10 in (" + CNULL(發明申請) + "," + CNULL(新型申請) + "," + CNULL(設計申請) + "," + CNULL(追加申請) + "," + CNULL(聯合申請) + "," + CNULL(翻譯) + ",'" & 113 & "','" & 114 & "','" & 307 & "')"
         '2006/3/7 MODIFY BY SONIA 加案件性質 109
         StrSQL3 = " and cp2.cp10 in (" + CNULL(發明申請) + "," + CNULL(新型申請) + "," + CNULL(設計申請) + "," + CNULL(追加申請) + "," + CNULL(聯合申請) + "," + CNULL(翻譯) + ",'" & 113 & "','" & 114 & "','" & 307 & "','109','110','112')"
         
         If strDate1 <> "" And StrDate2 <> "" Then
            StrSQL3 = StrSQL3 & " and cp2.cp27 between " + strDate1 + " and " + StrDate2
         ElseIf StrDate2 <> "" Then
            StrSQL3 = StrSQL3 & " and (cp2.cp27<=" + StrDate2 & " or cp2.cp27 is null)"
         End If
         
         'Modify by Morgan 2004/12/27 加國外案發文日CP271
         strSql = "select distinct cm01 國外案號,cm02 國外案號,cm03 國外案號,cm04 國外案號,nvl(pa1.pa05,nvl(pa1.pa06,pa1.pa07)) 案件名稱" & _
            ",st1.st02 承辦人,st2.st02 智權人員" & _
            ",DECODE(CP271,'','',SUBSTR(CP271,1,4)-1911||'/'||SUBSTR(CP271,5,2)||'/'||SUBSTR(CP271,7,2)) 發文日" & _
            ",cm05 國內案號,cm06 國內案號,cm07 國內案號,cm08 國內案號,nvl(pa2.pa05,nvl(pa2.pa06,pa2.pa07)) 案件名稱" & _
            ",st3.st02 承辦人,st4.st02 智權人員" & _
            ",DECODE(CP27,'','',SUBSTR(CP27,1,4)-1911||'/'||SUBSTR(CP27,5,2)||'/'||SUBSTR(CP27,7,2)) 發文日" & _
            ",DECODE(CP57,'','',SUBSTR(CP57,1,4)-1911||'/'||SUBSTR(CP57,5,2)||'/'||SUBSTR(CP57,7,2)) 取消收文日" & _
            ",cm18 記錄" & _
            " FROM (" & _
            " select distinct cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08,cm18" & _
            ",cp1.cp14 cp141,cp1.cp13 cp131,cp2.cp14 cp142, cp2.cp13 cp132,cp2.cp27,cp2.cp57,cp1.cp27 cp271" & _
            " from caseprogress cp2,casemap,caseprogress cp1" & _
            " where (cp2.cp01||''='P' OR cp2.CP01||''='CFP')" & StrSQL3 & _
            " and cm05(+)=cp2.cp01 and cm06(+)=cp2.cp02  and cm07(+)=cp2.cp03 and cm08(+)=cp2.cp04 and CM10='" & iSitu & "'" & _
            " AND cp1.cp01(+)=cm01 and cp1.cp02(+)=cm02 and cp1.cp03(+)=cm03 and cp1.cp04(+)=cm04" & _
            " and (cp1.cp01||''='P' OR cp1.CP01||''='CFP')" & strSQL1 & _
            " and cp1.CP57 IS NULL ),patent pa1,patent pa2,staff st1,staff st2,staff st3,staff st4" & _
            " where pa1.pa01(+)=cm01 and pa1.pa02(+)=cm02 and pa1.pa03(+)=cm03 and pa1.pa04(+)=cm04" & _
            " and st1.st01(+)=cp141 and st2.st01(+)=cp131" & _
            " and pa2.pa01(+)=cm05 and pa2.pa02(+)=cm06 and pa2.pa03(+)=cm07 and pa2.pa04(+)=cm08" & _
            " and st3.st01(+)=cp142 and st4.st01(+)=cp132" & midSql & _
            " ORDER BY cm01,cm02,cm03,CM04"
         '2004/9/7 END
      'Add by Morgan 2007/4/26
      'Modified by Lydia 2015/07/27 +澳門案"5"
'      Case "4" '香港大陸
'
'         strSQL1 = " and cp1.cp10='110'"
      Case "4", "5"
         If iSitu = 4 Then
            strSQL1 = " and cp1.cp10='110'": m_Title = "香港"
         Else
            strSQL1 = " and cp1.cp10 in (" & CaseMapIn & ")": m_Title = "澳門"
         End If
      'end 2015/07/27
         If strEnginer <> "" Then
            strSQL1 = strSQL1 & " and cp1.cp14=" + CNULL(strEnginer)
         End If
      
         StrSQL3 = " and cp2.cp10 in (" + CNULL(發明申請) + "," + CNULL(新型申請) + "," + CNULL(設計申請) + "," + CNULL(追加申請) + "," + CNULL(聯合申請) + "," + CNULL(翻譯) + ",'" & 113 & "','" & 114 & "','" & 307 & "','109','110','112')"
         If strDate1 <> "" And StrDate2 <> "" Then
            StrSQL3 = StrSQL3 & " and cp2.cp27 between " + strDate1 + " and " + StrDate2
         ElseIf StrDate2 <> "" Then
            StrSQL3 = StrSQL3 & " and (cp2.cp27<=" + StrDate2 & " or cp2.cp27 is null)"
         End If
         'Modified by Lydia 2015/07/27 改名稱
'         strSql = "select distinct cm01 香港案號,cm02 香港案號,cm03 香港案號,cm04 香港案號,nvl(pa1.pa05,nvl(pa1.pa06,pa1.pa07)) 案件名稱" & _
            ",st1.st02 承辦人,st2.st02 智權人員" & _
            ",DECODE(CP271,'','',SUBSTR(CP271,1,4)-1911||'/'||SUBSTR(CP271,5,2)||'/'||SUBSTR(CP271,7,2)) 發文日" & _
            ",cm05 大陸案號,cm06 大陸案號,cm07 大陸案號,cm08 大陸案號,nvl(pa2.pa05,nvl(pa2.pa06,pa2.pa07)) 案件名稱" & _
            ",st3.st02 承辦人,st4.st02 智權人員" & _
            ",DECODE(CP27,'','',SUBSTR(CP27,1,4)-1911||'/'||SUBSTR(CP27,5,2)||'/'||SUBSTR(CP27,7,2)) 發文日" & _
            ",DECODE(CP57,'','',SUBSTR(CP57,1,4)-1911||'/'||SUBSTR(CP57,5,2)||'/'||SUBSTR(CP57,7,2)) 取消收文日" & _
            ",cm18 記錄" & _
            " FROM (" & _
            " select distinct cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08,cm18" & _
            ",cp1.cp14 cp141,cp1.cp13 cp131,cp2.cp14 cp142, cp2.cp13 cp132,cp2.cp27,cp2.cp57,cp1.cp27 cp271" & _
            " from caseprogress cp2,casemap,caseprogress cp1" & _
            " where (cp2.cp01||''='P' OR cp2.CP01||''='CFP')" & StrSQL3 & _
            " and cm05(+)=cp2.cp01 and cm06(+)=cp2.cp02  and cm07(+)=cp2.cp03 and cm08(+)=cp2.cp04 and CM10='" & iSitu & "'" & _
            " AND cp1.cp01(+)=cm01 and cp1.cp02(+)=cm02 and cp1.cp03(+)=cm03 and cp1.cp04(+)=cm04" & _
            " and (cp1.cp01||''='P' OR cp1.CP01||''='CFP')" & strSQL1 & _
            " and cp1.CP57 IS NULL ),patent pa1,patent pa2,staff st1,staff st2,staff st3,staff st4" & _
            " where pa1.pa01(+)=cm01 and pa1.pa02(+)=cm02 and pa1.pa03(+)=cm03 and pa1.pa04(+)=cm04" & _
            " and st1.st01(+)=cp141 and st2.st01(+)=cp131" & _
            " and pa2.pa01(+)=cm05 and pa2.pa02(+)=cm06 and pa2.pa03(+)=cm07 and pa2.pa04(+)=cm08" & _
            " and st3.st01(+)=cp142 and st4.st01(+)=cp132" & midSql & _
            " ORDER BY cm01,cm02,cm03,CM04"
         strSql = "select distinct cm01 " & m_Title & "案號,cm02 " & m_Title & "案號,cm03 " & m_Title & "案號,cm04 " & m_Title & "案號,nvl(pa1.pa05,nvl(pa1.pa06,pa1.pa07)) 案件名稱" & _
            ",st1.st02 承辦人,st2.st02 智權人員" & _
            ",DECODE(CP271,'','',SUBSTR(CP271,1,4)-1911||'/'||SUBSTR(CP271,5,2)||'/'||SUBSTR(CP271,7,2)) 發文日" & _
            ",cm05 大陸案號,cm06 大陸案號,cm07 大陸案號,cm08 大陸案號,nvl(pa2.pa05,nvl(pa2.pa06,pa2.pa07)) 案件名稱" & _
            ",st3.st02 承辦人,st4.st02 智權人員" & _
            ",DECODE(CP27,'','',SUBSTR(CP27,1,4)-1911||'/'||SUBSTR(CP27,5,2)||'/'||SUBSTR(CP27,7,2)) 發文日" & _
            ",DECODE(CP57,'','',SUBSTR(CP57,1,4)-1911||'/'||SUBSTR(CP57,5,2)||'/'||SUBSTR(CP57,7,2)) 取消收文日" & _
            ",cm18 記錄" & _
            " FROM (" & _
            " select distinct cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08,cm18" & _
            ",cp1.cp14 cp141,cp1.cp13 cp131,cp2.cp14 cp142, cp2.cp13 cp132,cp2.cp27,cp2.cp57,cp1.cp27 cp271" & _
            " from caseprogress cp2,casemap,caseprogress cp1" & _
            " where (cp2.cp01||''='P' OR cp2.CP01||''='CFP')" & StrSQL3 & _
            " and cm05(+)=cp2.cp01 and cm06(+)=cp2.cp02  and cm07(+)=cp2.cp03 and cm08(+)=cp2.cp04 and CM10='" & iSitu & "'" & _
            " AND cp1.cp01(+)=cm01 and cp1.cp02(+)=cm02 and cp1.cp03(+)=cm03 and cp1.cp04(+)=cm04" & _
            " and (cp1.cp01||''='P' OR cp1.CP01||''='CFP')" & strSQL1 & _
            " and cp1.CP57 IS NULL ),patent pa1,patent pa2,staff st1,staff st2,staff st3,staff st4" & _
            " where pa1.pa01(+)=cm01 and pa1.pa02(+)=cm02 and pa1.pa03(+)=cm03 and pa1.pa04(+)=cm04" & _
            " and st1.st01(+)=cp141 and st2.st01(+)=cp131" & _
            " and pa2.pa01(+)=cm05 and pa2.pa02(+)=cm06 and pa2.pa03(+)=cm07 and pa2.pa04(+)=cm08" & _
            " and st3.st01(+)=cp142 and st4.st01(+)=cp132" & midSql & _
            " ORDER BY cm01,cm02,cm03,CM04"
   End Select
   
    rsRecordset.CursorLocation = adUseClient
    rsRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
    Set ReadCaseRelationRst = rsRecordset
    Set rsRecordset = Nothing
  
ErrHand:
    If Err.Number <> 0 Then
        MsgBox Err.Description
    End If
End Function

'新增國內外案件關聯表
'iSitu 0 國內外案件資料維護資料
'Modify by Morgan 2007/10/16 只有香港大陸案有用，國內外程式移除不再維護
'Modified by Lydia 2015/07/27 +澳門案
Public Function InsertCaseRelationData(ByRef strCode() As String, Optional ByVal iSitu As Integer = 0) As Boolean
     
On Error GoTo ErrHand

    Dim strSqlEx As String, strTxt(0 To 3) As String, i As Integer
    Dim rsRecordset As New ADODB.Recordset, strTmp As String, strPromoteDate As String
     
    strCode(2) = IIf(strCode(2) = "", "0", strCode(2))
    strCode(3) = IIf(strCode(3) = "", "00", strCode(3))
    strCode(6) = IIf(strCode(6) = "", "0", strCode(6))
    strCode(7) = IIf(strCode(7) = "", "00", strCode(7))
    
    Select Case iSitu
       'edit by nickc 2005/06/28
       Case 0
'Remove by Morgan 2007/10/16
'          For i = 4 To 7
'             strTxt(i - 4) = strCode(i)
'          Next
'          '檢查若內案是否有已發文資料
'          If Not ChkDocu(strTxt, "CP27", False, True) Then
'            '檢查國外案是否有未發文資料
'            '2006/3/7 MODIFY BY SONIA 加案件性質 109,110,112
'            strSqlEx = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & strCode(0) & "' AND CP02='" & strCode(1) & "'" & _
'                " AND CP03='" & strCode(2) & "' AND CP04='" & strCode(3) & "'" & _
'                " AND CP10 in ('" & 發明申請 & "','" & 新型申請 & "','" & 設計申請 & "','" & 追加申請 & "','" & 聯合申請 & "','" & 翻譯 & "','113','114','307','109','110','112')" & _
'                " AND CP27 IS NULL"
'            rsRecordset.CursorLocation = adUseClient
'            rsRecordset.Open strSqlEx, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
'             If Not rsRecordset.BOF And Not rsRecordset.EOF Then
'                strTmp = "" & rsRecordset.Fields(0).Value
'             Else
'                strTmp = ""
'             End If
'             rsRecordset.Close
'             If strTmp <> "" Then
'                '更新國外案之文件齊備日
'                strSqlEx = "SELECT COUNT(*) FROM ENGINEERPROGRESS WHERE EP02=" & CNULL(strTmp)
'                rsRecordset.CursorLocation = adUseClient
'                rsRecordset.Open strSqlEx, cnnConnection, adOpenStatic, adLockReadOnly   'edit by nickc 2005/08/04 原先是  動態開啟
'                If rsRecordset.Fields(0) = 0 Then
'                        MsgBox "無法更新國外案之文件齊備日 !", vbInformation
'                Else
'                    strSqlEx = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02=" & CNULL(strTmp) & " AND EP06 IS NULL"
'                    cnnConnection.Execute strSqlEx, intI
'                    '若文件齊備日為工作天
'                    'Modify by Morgan 2007/10/16 加判斷有齊備日有更新才做
'                    'If ChkWorkDay(strSrvDate(1)) = True Then
'                    If ChkWorkDay(strSrvDate(1)) = True And intI = 1 Then
'                        strSqlEx = "Select NVL(CF04,0) CF04, CP06,CP01,PA09,CP10 From CaseProgress, Patent, Casefee Where CP09=" & CNULL(strTmp) & " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+)  and PA01=cf01(+) and pa09=cf02(+) and cp10=cf03 "
'                        intI = 1
'                        Set rsRecordset = ClsLawReadRstMsg(intI, strSqlEx)
'                        If intI = 1 Then
'                            With rsRecordset
'                            'Modify by Morgan 2007/10/11 承辦期限改呼叫共用函數計算
'                            'If Val("" & rsRecordset.Fields(0).Value) > 0 Then
'                            '    strPromoteDate = CompWorkDay(rsRecordset.Fields(0).Value, strSrvDate(1), 0)
'                            '    '若有計算出承辦期限且有本所期限
'                            '    If strPromoteDate <> "" And ("" & rsRecordset.Fields("CP06").Value) <> "" Then
'                            '        '若承辦期限大於本所期限
'                            '        '轉換成西元日期比較大小
'                            '        If Val(strPromoteDate) > Val("" & rsRecordset.Fields("CP06").Value) Then
'                            '            strPromoteDate = "" & Val("" & rsRecordset.Fields("CP06").Value)
'                            '        End If
'                            '    End If
'                            strPromoteDate = Pub_GetHandleDay(.Fields("CP01"), .Fields("PA09"), .Fields("CP10"), , "" & .Fields("CP06"), strTmp)
'                            If strPromoteDate <> "" Then
'                            'end 2007/10/16
'                                strSqlEx = "Update CaseProgress Set CP48=" & CNULL(strPromoteDate) & " Where CP09=" & CNULL(strTmp)
'                                cnnConnection.Execute strSqlEx
'                            End If
'                            End With
'                        End If
'                    End If
'                End If
'            End If
'          End If
'          strSqlEx = "insert into casemap (cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08,cm10) values (" + CNULL(strCode(0)) + "," + CNULL(strCode(1)) + "," + CNULL(strCode(2)) + "," + CNULL(strCode(3)) + "," + CNULL(strCode(4)) + "," + CNULL(strCode(5)) + "," + CNULL(strCode(6)) + "," + CNULL(strCode(7)) + ",'" & iSitu & "')"
'         cnnConnection.Execute strSqlEx
'
'         'Add by Morgan 2005/8/24 若國內外案之國內案領證已發文且未公告時需更新國外新案期限
'         strExc(1) = strCode(0): strExc(2) = strCode(1): strExc(3) = strCode(2): strExc(4) = strCode(3)
'         strExc(5) = strCode(4): strExc(6) = strCode(5): strExc(7) = strCode(6): strExc(8) = strCode(7)
'         PUB_UpdateCP0607 strExc
         
      'add by nickc 2005/06/28
      'Modified by Lydia 2015/07/27
      'Case 4
      Case 4, 5
         For i = 4 To 7
            strTxt(i - 4) = strCode(i)
         Next
         strSqlEx = "insert into casemap (cm01,cm02,cm03,cm04,cm05,cm06,cm07,cm08,cm10) values (" + CNULL(strCode(0)) + "," + CNULL(strCode(1)) + "," + CNULL(strCode(2)) + "," + CNULL(strCode(3)) + "," + CNULL(strCode(4)) + "," + CNULL(strCode(5)) + "," + CNULL(strCode(6)) + "," + CNULL(strCode(7)) + ",'" & iSitu & "')"
         cnnConnection.Execute strSqlEx
    End Select
    InsertCaseRelationData = True
    
ErrHand:
    Set rsRecordset = Nothing
    If Err.Number <> 0 Then ShowMsg MsgText(1005)
End Function

'讀取國內案件關聯表之資料
'iSitu 0 國內外案件資料維護資料
Public Function GetCaseRelationDataIn(ByRef strCode1 As String, ByRef strCode2 As String, _
   ByRef strCode3 As String, ByRef strCode4 As String, ByRef strPromoter As String, _
   ByRef strSendDay As String, Optional ByVal iSitu As Integer = 0, Optional strSales As String) As Boolean
   
On Error GoTo ErrHand

    Dim strSql As String, rsRecordset As New ADODB.Recordset, strWhich As String
    Dim strTmp As String
    
    Select Case iSitu
       Case 0
          strWhich = "國內"
          strTmp = "cp27"
      'add by nickc 2005/06/28
      Case 4
          strWhich = "大陸香港"
          strTmp = "cp27"
      'Added by Lydia 2015/07/27 +澳門案
      Case 5
          strWhich = "大陸澳門"
          strTmp = "cp27"
    End Select
    
    strPromoter = ""
    strSales = ""
    strSendDay = ""
    
    strCode3 = IIf(strCode3 = "", "0", strCode3)
    strCode4 = IIf(strCode4 = "", "00", strCode4)
    'Modify by Morgan 2006/1/19 加 109 PCT申請
    '2006/3/7 MODIFY BY SONIA 加案件性質 110,112
    'Modify by Morgan 2006/4/14 案件性質改用常數控制
    strSql = "select s1.st02," & strTmp & ",cp10,s2.st02 from caseprogress,staff s1,staff s2 where " & _
       "cp01=" + CNULL(strCode1) + " and " & "cp02=" + CNULL(strCode2) + _
       " and cp03=" + CNULL(strCode3) + " and cp04=" + CNULL(strCode4) + _
       " and cp10 in (" & CaseMapIn & ") and cp14=s1.st01(+) and cp13=s2.st01(+)"
    
    rsRecordset.CursorLocation = adUseClient
    rsRecordset.Open strSql, cnnConnection, adOpenDynamic
    If Not rsRecordset.EOF And Not rsRecordset.BOF Then
        strPromoter = IIf(IsNull(rsRecordset.Fields(0)), "", rsRecordset.Fields(0))
        strSendDay = IIf(IsNull(rsRecordset.Fields(1)), "", rsRecordset.Fields(1))
        If Not IsNull(rsRecordset.Fields(3)) Then strSales = rsRecordset.Fields(3)
        GetCaseRelationDataIn = True
    Else
       ShowMsg strWhich + MsgText(1010)
    End If
    Exit Function
ErrHand:
    ShowMsg MsgText(1008)
End Function

'bolTmp False 新案 True 檢索報告
Private Function ChkDocu(ByRef strCode() As String, ByVal strTmp As String, Optional bolTmp As Boolean = False, Optional bolShowMsg As Boolean = False) As Boolean
 Dim strSql As String, rsRecordset As New ADODB.Recordset, strTmp1 As String, i As Integer
On Error GoTo ErrHand
   ChkDocu = False
   Select Case strTmp
      Case "CP25"
         strTmp1 = "准駁"
      Case "CP27"
         strTmp1 = "發文"
      Case "CP57"
         strTmp1 = "取消收文"
   End Select
   If bolTmp Then
      strSql = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & strCode(0) & "' AND CP02='" & strCode(1) & "' AND CP03='" & strCode(2) & _
         "' AND CP04='" & strCode(3) & "' AND CP10='" & 檢索報告 & "' AND " & strTmp & " IS NOT NULL"
   Else
      strSql = "SELECT CP09 FROM CASEPROGRESS WHERE CP01='" & strCode(0) & "' AND CP02='" & strCode(1) & _
         "' AND CP03='" & strCode(2) & "' AND CP04='" & strCode(3) & "' AND " & _
         "CP10 in ('" & 發明申請 & "','" & 新型申請 & "','" & 設計申請 & "','" & 追加申請 & "','" & 聯合申請 & "','" & 翻譯 & "','" & 113 & "','" & 114 & "','" & 307 & "')" & _
         " AND " & strTmp & " IS NOT NULL"
   End If
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection, adOpenDynamic
   If Not rsRecordset.BOF And Not rsRecordset.EOF Then
        If bolShowMsg Then ShowMsg strCode(0) & strCode(1) & strCode(2) & strCode(3) & "已有" & strTmp1 & "日 !"
   Else
      ChkDocu = True
   End If
Exit Function
ErrHand:
   ShowMsg Err.Description
End Function
'Add end 2004/3/15

'Add by Morgan 2004/3/19
'檢查基本檔是否已核准
'Modify By Sindy 2018/11/28 + , Optional ByVal strMsgText As String=""
Public Function PUB_ApproveCheck(ByVal stRecNo As String, Optional ByVal strMsgText As String = "") As Boolean

    Dim stSQL As String, rsQuery As New ADODB.Recordset
    
On Error GoTo flgErr

    stSQL = "SELECT PA16 FROM CASEPROGRESS, PATENT WHERE PA01=CP01 AND PA02=CP02 AND PA03=CP03 AND PA04=CP04 AND CP09='" & stRecNo & "'"
    rsQuery.CursorLocation = adUseClient
    rsQuery.Open stSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
    If rsQuery.RecordCount > 0 Then
        If IsNull(rsQuery.Fields(0)) Then
            MsgBox "基本檔尚無准駁，" & IIf(strMsgText <> "", strMsgText, "不可發文") & "！", vbCritical
        ElseIf (rsQuery.Fields(0) = "2") Then
            MsgBox "基本檔目前為核駁，" & IIf(strMsgText <> "", strMsgText, "不可發文") & "！", vbCritical
        ElseIf (rsQuery.Fields(0) = "1") Then
            PUB_ApproveCheck = True
        End If
    Else
        MsgBox "無法讀取基本檔資料", vbCritical
    End If
    
flgErr:
    
    If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
    Set rsQuery = Nothing
    
End Function

'Add by Morgan 2004/3/23
'Modified by Morgan 2015/9/10 日期不必再民國或西元
'計算實體審查法定期限
'stRefDate=收/發文日,stLawDate=原案實體審查法定期限
Public Function PUB_Get416LawLimit(ByVal stRefDate As String, ByRef stLawDate As String) As String
   Dim dtTemp1 As String, dtTemp2 As String, stDate As String
   '原案實體審查法定期限
   dtTemp1 = DBDATE(stLawDate)
   '收/發文日
   dtTemp2 = DBDATE(stRefDate)
   '原案實體審查法定期限>收/發文日-->法定期限=原案實體審查法定期限
   If dtTemp1 > dtTemp2 Then
      stDate = dtTemp1
   '原案實體審查法定期限<=收/發文日-->法定期限='收/發文日+30天
   Else
      '收/發文日+30天
      stDate = CompDate(2, 30, dtTemp2)
   End If
   PUB_Get416LawLimit = stDate
End Function

'Add by Morgan 2004/3/23
'檢查是否有未取消收文的實體審查
'stCP09 實體審查收文號,
Public Function PUB_Get416CP09(ByRef stCP09 As String, ByRef stLimit As String, ByRef pa() As String, Optional ByVal bolMsg As Boolean = True) As Boolean
   Dim stSQL As String, rsQuery As New ADODB.Recordset
On Error GoTo flgErr
   stSQL = " SELECT CP09 FROM CASEPROGRESS" & _
      " WHERE CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "' AND CP10='416' AND CP57 IS NULL"
   rsQuery.CursorLocation = adUseClient
   rsQuery.Open stSQL, cnnConnection, adOpenStatic, adLockReadOnly
   If rsQuery.RecordCount > 0 Then
      stCP09 = rsQuery.Fields(0).Value
   ElseIf bolMsg = True Then
      MsgBox "此分割案尚未收文實體審查，期限為" & stLimit & "，請提醒智權人員!!!", vbExclamation
   End If
   PUB_Get416CP09 = True
   
flgErr:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
End Function

'Remove by Morgan 2009/5/21 不再使用
''Add by Morgan 2005/8/24
''讀取國外新案收文號,期限
'Public Function PUB_GetUpdateInfo(ByRef p_CM() As String, ByRef p_CP09 As String, ByRef p_CP06 As String, ByRef p_CP07 As String) As Boolean
'
'   Dim stCP27 As String, stCP71 As String, stDate(1 To 3) As String
'
'   If p_CM(1) <> "CFP" Or p_CM(5) <> "P" Then Exit Function
'
'   '抓國外案新案未發文無期限,台灣案領證已發文未公告,考慮有延緩公告
'   strSQL = "SELECT A.CP09,B.CP27,B.CP71 FROM CASEMAP,CASEPROGRESS A, PATENT,CASEPROGRESS B" & _
'      " WHERE CM01='" & p_CM(1) & "' AND CM02='" & p_CM(2) & "' AND CM03='" & p_CM(3) & "' AND CM04='" & p_CM(4) & "'" & _
'      " AND CM05='" & p_CM(5) & "' AND CM06='" & p_CM(6) & "' AND CM07='" & p_CM(7) & "' AND CM08='" & p_CM(8) & "' AND CM10='0'" & _
'      " AND A.CP01(+)=CM01 AND A.CP02(+)=CM02  AND A.CP03(+)=CM03  AND A.CP04(+)=CM04" & _
'      " AND A.CP06 IS NULL AND A.CP27 IS NULL AND A.CP31='Y' AND A.CP57 IS NULL" & _
'      " AND PA01(+)=CM05 AND PA02(+)=CM06 AND PA03(+)=CM07 AND PA04(+)=CM08 AND PA09='000' AND PA14 IS NULL" & _
'      " AND B.CP01(+)=PA01 AND B.CP02(+)=PA02  AND B.CP03(+)=PA03  AND B.CP04(+)=PA04" & _
'      " AND (B.CP10='601' OR B.CP10='412') AND B.CP27>0 AND B.CP57 IS NULL" & _
'      " ORDER BY 3"
'
'   CheckOC3
'   With AdoRecordSet3
'      .CursorLocation = adUseClient
'      .Open strSQL, cnnConnection, adOpenForwardOnly, adLockReadOnly
'      If .RecordCount > 0 Then
'         stCP27 = .Fields(1)
'         stCP71 = "" & .Fields(2)
'         '抓
'         PUB_Get605NP stCP27, 0, stDate(), stCP71
'         p_CP09 = .Fields(0)
'         '法定期限=預估公告日
'         p_CP07 = stDate(3)
'         '本所期限=法定期限-10天
'         p_CP06 = PUB_GetWorkDay1(CompDate(2, -10, p_CP07), True)
'         If p_CP06 < strSrvDate(1) Then
'            p_CP06 = strSrvDate(1)
'         End If
'         PUB_GetUpdateInfo = True
'      End If
'   End With
'
'End Function

''Add by Morgan 2005/8/24
''若國內外案之國內案領證已發文且未公告時需更新國外新案期限
'Public Function PUB_UpdateCP0607(ByRef p_CM() As String) As Boolean
'   Dim stCP09 As String, stCP06 As String, stCP07 As String
'   If p_CM(1) = "CFP" And p_CM(5) = "P" Then
'      If PUB_GetUpdateInfo(p_CM, stCP09, stCP06, stCP07) = True Then
'         strSQL = "Update caseprogress set CP06=" & stCP06 & ",CP07=" & stCP07 & _
'            " WHERE CP06 IS NULL AND CP09='" & stCP09 & "' AND CP27 IS NULL AND CP31='Y' AND CP57 IS NULL"
'         cnnConnection.Execute strSQL
'      End If
'   End If
'
'End Function
'end 2009/5/21

'Add by Morgan 2006/4/28
'1.已收文存活證明221->更新期限(未發文) or 2.有掛存活證明221期限於NP->更新期限(未續辦) or 3.新增NP
'Modify By Sindy 2021/4/27 + , p_NP23 As String:約定期限
Public Function PUB_Get221SQL(ByRef p_CP() As String, p_CP06 As String, p_CP07 As String, _
   p_CP13 As String, p_NP01 As String, Optional p_NP23 As String)
   
   'Modified by Lydia 2022/09/15 p_NP23 => CNULL(p_NP23, True)
   PUB_Get221SQL = "declare V_221CP09 VARCHAR2(9); V_221NP22 NUMERIC(10); V_NewNP22 NUMERIC(10);" & _
   " begin" & _
   "  SELECT MAX(CP09) INTO V_221CP09 from caseprogress where cp01='" & p_CP(1) & "' and cp02='" & p_CP(2) & "' and cp03='" & p_CP(3) & "' and cp04='" & p_CP(4) & "' and cp10='221';" & _
   "  If V_221CP09 Is Not Null Then" & _
   "   UPDATE CASEPROGRESS SET CP06=" & p_CP06 & ",CP07=" & p_CP07 & " WHERE CP09=V_221CP09 AND CP27 IS NULL;" & _
   "  Else" & _
   "   SELECT MAX(NP22) INTO V_221NP22 from nextprogress where np02='" & p_CP(1) & "' and np03='" & p_CP(2) & "' and np04='" & p_CP(3) & "' and np05='" & p_CP(4) & "' and NP07='221';" & _
   "   If V_221NP22 Is Not Null Then" & _
   "    UPDATE NEXTPROGRESS SET NP08=" & p_CP06 & ",NP09=" & p_CP07 & ",NP23=" & CNULL(p_NP23, True) & " WHERE NP22=V_221NP22 AND (NP06 IS NULL OR NP06='N');" & _
   "   Else" & _
   "    select nvl(max(np22),0)+1 into V_NewNP22 from nextprogress;" & _
   "    INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22,NP23)" & _
   "    values ('" & p_NP01 & "','" & p_CP(1) & "','" & p_CP(2) & "','" & p_CP(3) & "','" & p_CP(4) & "','221'," & p_CP06 & "," & p_CP07 & ",'" & p_CP13 & "',V_NewNP22," & CNULL(p_NP23, True) & ");" & _
   "   end if;  END IF; end;"
End Function

'*************************************************
'  會計科目名稱查詢
'
'*************************************************
Public Function A0102Query(InputNo As String) As String
Dim adoacc010 As New ADODB.Recordset
   adoacc010.CursorLocation = adUseClient
   adoacc010.Open "select * from acc010 where a0101 = '" & InputNo & "'", adoTaie, adOpenStatic, adLockReadOnly
   If adoacc010.RecordCount <> 0 Then
      If IsNull(adoacc010.Fields("a0102").Value) Then
         A0102Query = MsgText(601)
      Else
         A0102Query = adoacc010.Fields("a0102").Value
      End If
   Else
      A0102Query = MsgText(601)
   End If
   adoacc010.Close
End Function

'Add by Morgan 2009/8/17
'若歐盟設計的其他多國皆有申請日時提醒該案已可發文
Public Sub chk103in239OK(cp() As String)
   Dim stSQL As String
   Dim adoRst As ADODB.Recordset
   Dim intR As Integer
   Dim stMsg As String
   Dim stVTB As String
   
   
   'Modify by Morgan 2009/10/21 +集體設計105(控制所有相同案都已有申請日)
   'Modified by Morgan 2019/6/24 相關案改用函數抓
   'stVTB = "select '" & cp(1) & "' cr01,'" & cp(2) & "' cr02,'" & cp(3) & "' cr03,'" & cp(4) & "' cr04 from dual" & _
      " Union select cr05,cr06,cr07,cr08 from caserelation" & _
      " where cr01='" & cp(1) & "' and cr02='" & cp(2) & "' and cr03='" & cp(3) & "' and cr04='" & cp(4) & "'" & _
      " Union select cm05,cm06,cm07,cm08 from casemap" & _
      " where cm01='" & cp(1) & "' and cm02='" & cp(2) & "' and cm03='" & cp(3) & "' and cm04='" & cp(4) & "'" & _
      " Union select cm01,cm02,cm03,cm04 from casemap" & _
      " where cm05='" & cp(1) & "' and cm06='" & cp(2) & "' and cm07='" & cp(3) & "' and cm08='" & cp(4) & "'" & _
      " Union select b.cm01,b.cm02,b.cm03,b.cm04 from casemap a,casemap b" & _
      " where a.cm01='" & cp(1) & "' and a.cm02='" & cp(2) & "' and a.cm03='" & cp(3) & "' and a.cm04='" & cp(4) & "'" & _
      " and b.cm05=a.cm05 and b.cm06=a.cm06 and b.cm07=a.cm07 and b.cm08=a.cm08" & _
      " Union select b.cm05,b.cm06,b.cm07,b.cm08 from casemap a,casemap b" & _
      " where a.cm05='" & cp(1) & "' and a.cm06='" & cp(2) & "' and a.cm07='" & cp(3) & "' and a.cm08='" & cp(4) & "'" & _
      " and b.cm01=a.cm01 and b.cm02=a.cm02 and b.cm03=a.cm03 and b.cm04=a.cm04"
   
   stVTB = PUB_GetRefCaseSQL(cp)
   'end 2019/6/24
   
   'Modified by Morgan 2017/9/12 +英國201、澳洲015,通知順序 澳洲>歐盟>英國 (一次只會有一個國家要通知)
   stSQL = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CaseNo,cp09,na03,decode(pa09,'015',1,'239',2,'201',3) Seq,cp10" & _
      " from (" & stVTB & ") X,patent,caseprogress,nation" & _
      " where pa01(+)=cr01 and pa02(+)=cr02 and pa04(+)=cr04" & _
      " and pa08='3' and pa09 in ('239','201','015') and pa10 is null and pa57 is null and na01(+)=pa09" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
      " and cp10 in ('103','105','125') and cp27 is null and cp57 is null"
      
   'Added by Morgan 2014/12/4
   'EU有主張優先權者不必通知 P-110172 -> CFP-27348
   stSQL = stSQL & " and not exists(select * from pridate where pd01=pa01 and pd02=pa02 and pd03=pa03 and pa04=pa04)"
   'end 2014/12/4
   
   'Added by Morgan 2019/6/24 +德國新型
   stSQL = stSQL & " union select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CaseNo,cp09,na03,4 Seq,cp10" & _
      " from (" & stVTB & ") X,patent,caseprogress,nation" & _
      " where pa01(+)=cr01 and pa02(+)=cr02 and pa04(+)=cr04" & _
      " and pa08='2' and pa09='231' and pa10 is null and pa57 is null and na01(+)=pa09" & _
      " and cp01(+)=pa01 and cp02(+)=pa02 and cp03(+)=pa03 and cp04(+)=pa04" & _
      " and cp10='102' and cp27 is null and cp57 is null" & _
      " and not exists(select * from pridate where pd01=pa01 and pd02=pa02 and pd03=pa03 and pa04=pa04)"
   'end 2019/6/24
   
   stSQL = stSQL & " order by Seq"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With adoRst
      If PUB_Chk103in239(.Fields("cp09"), False) = False Then
         'Modified by Morgan 2019/6/24
         'stMsg = .Fields("na03") & "設計案" & .Fields("CaseNo")
         stMsg = .Fields("na03") & IIf(.Fields("cp10") = "102", "新型案", "設計案") & .Fields("CaseNo")
         
         'MsgBox "所有相同案皆已有申請日，請通知相關人員歐盟設計 " & stMsg & " 已可發文!!"
         '發Mail通知
         stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08) values ('" & strUserNum & "','79017',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'所有關連案皆已有申請日或閉卷，" & stMsg & " 已可發文!!',' ')"
         cnnConnection.Execute stSQL, intR
      End If
      End With
   End If
   Set adoRst = Nothing
End Sub
'Add by Morgan 2009/10/1
'來函有收文實審或再審的退費提醒
Public Sub Check908(cp() As String)
   Dim stSQL As String
   Dim adoRst As ADODB.Recordset
   Dim intR As Integer
   Dim stMsg As String
   
   stSQL = "select 1 from caseprogress c1 where cp01='" & cp(1) & "'" & _
      " and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='908' and cp27 is null and cp57 is null" & _
      " and exists(select * from caseprogress c2 where c2.cp09=c1.cp43 and c2.cp10 in ('416','107'))"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      MsgBox "本案有收文實審或再審的退費且尚未發文!!"
   End If
   Set adoRst = Nothing
End Sub

'Add by Morgan 2009/7/21
'歐盟設計發文需控制其他多國都已有申請日
Public Function PUB_Chk103in239(strCP09 As String, Optional bMsg As Boolean = True) As Boolean
   Dim stSQL As String
   Dim adoRst As ADODB.Recordset
   Dim intR As Integer
   Dim stList As String
   Dim stCon As String
   Dim stCountryName As String
   Dim stExceptCountry As String
   'Added by Morgan 2018/12/12
   Dim arrNum(4) As String
   Dim stVTB As String
   
   'Modify by Morgan 2009/10/21 +集體設計105(控制所有相同案都已有申請日)
   'Modified by Morgan 2017/9/11 +英國201、澳洲015
   'Modified by Morgan 2018/11/29 +德國新型外翻案件
   'Modified by Morgan 2019/6/24 德國新型不必再限制外翻案件(第2句拿掉 and cp14 like 'F%')--禧佩 Ex:CFP030920(德),CFP030919(日)
   stSQL = "select cp10,pa09,cp01,cp02,cp03,cp04,na03 from caseprogress,patent,nation where cp09='" & strCP09 & "' and cp10 in ('103','105')" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09 in ('201','239','015') and na01(+)=pa09" & _
      " union select cp10,pa09,cp01,cp02,cp03,cp04,na03 from caseprogress,patent,nation where cp09='" & strCP09 & "' and cp10='102'" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 and pa09='231' and na01(+)=pa09"
      
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With adoRst
      'Added by Morgan 2017/9/11 核准速度:英國>歐盟>澳洲,故申請歐盟時要排除英國,澳洲要排除英國及歐盟
      stCountryName = "" & .Fields("na03") & IIf(.Fields("cp10") = "102", "新型", "設計")
      If .Fields("pa09") = "239" Then
         stExceptCountry = " and pa09<>'201'"
      ElseIf .Fields("pa09") = "015" Then
         stExceptCountry = " and pa09<>'201' and pa09<>'239'"
      End If
      'end 2017/9/11
      
      'Added by Morgan 2018/12/12

      'end 2018/12/12
      
      'Modify by Morgan 2009/8/17 國內案的其他國外案(可能有其他的P案)也要檢查
      'Modify by Morgan 2010/8/18 排除有主張優先權的案件  CFP-23270
      'Modify by Morgan 2011/10/19 排除集體子案
      'Modified by Morgan 2018/12/12 相關案改用函數抓
      'stCon = " and not (cr05='" & .Fields("cp01") & "' and cr06='" & .Fields("cp02") & "' and cr08='" & .Fields("cp04") & "')"
      'stSQL = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CaseNo,na03" & _
         " from (select cr05,cr06,cr07,cr08 from caserelation" & _
         " where cr01='" & .Fields("cp01") & "' and cr02='" & .Fields("cp02") & "'" & _
         " and cr04='" & .Fields("cp04") & "'" & _
         " union select cm05,cm06,cm07,cm08 from casemap" & _
         " where cm01='" & .Fields("cp01") & "' and cm02='" & .Fields("cp02") & "'" & _
         " and cm04='" & .Fields("cp04") & "'" & _
         " union select b.cm01,b.cm02,b.cm03,b.cm04 from casemap a,casemap b" & _
         " where a.cm01='" & .Fields("cp01") & "' and a.cm02='" & .Fields("cp02") & "'" & _
         " and a.cm04='" & .Fields("cp04") & "'" & _
         " and b.cm05=a.cm05 and b.cm06=a.cm06 and b.cm07=a.cm07 and b.cm08=a.cm08" & _
         " and not (b.cm01=a.cm01 and b.cm02=a.cm02 and b.cm03=a.cm03 and b.cm04=a.cm04)" & _
         " ) A,patent,nation where pa01(+)=cr05 and pa02(+)=cr06 and pa03(+)=cr07 and pa04(+)=cr08" & stCon & _
         " and pa10 is null and pa57 is null" & stExceptCountry & " and na01(+)=pa09" & _
         " and not exists(select * from pridate where pd01=pa01 and pd02=pa02 and pd03=pa03 and pd04=pa04)"
      
      arrNum(1) = .Fields("cp01")
      arrNum(2) = .Fields("cp02")
      arrNum(3) = .Fields("cp03")
      arrNum(4) = .Fields("cp04")
      stVTB = PUB_GetRefCaseSQL(arrNum)
      stCon = " and not (cr01='" & .Fields("cp01") & "' and cr02='" & .Fields("cp02") & "' and cr04='" & .Fields("cp04") & "')" '排除集體子案(第3碼不同者)
      stSQL = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CaseNo,na03" & _
         " from (" & stVTB & ") A,patent,nation" & _
         " where pa01(+)=cr01 and pa02(+)=cr02 and pa03(+)=cr03 and pa04(+)=cr04" & stCon & _
         " and pa10 is null and pa57 is null" & stExceptCountry & " and na01(+)=pa09" & _
         " and not exists(select * from pridate where pd01=pa01 and pd02=pa02 and pd03=pa03 and pd04=pa04)"
      'end 2018/12/12
      intR = 1
      Set adoRst = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         PUB_Chk103in239 = True
         If bMsg Then
            With adoRst '一定要重設否則會抓到前一次的查詢結果
            Do While Not .EOF
               stList = stList & vbCrLf & "　　" & .Fields("CaseNo") & " " & .Fields("na03")
               .MoveNext
            Loop
            End With
            'Modify by Morgan 2010/10/8 提醒但可選擇發文--甄妮,郭
            'MsgBox "因下列相同案尚無申請日：" & vbCrLf & stList & vbCrLf & vbCrLf & "歐盟設計不可先發文!!"
            If MsgBox("因下列相同案尚無申請日：" & vbCrLf & stList & vbCrLf & vbCrLf & stCountryName & "案應不可先發文!!是否仍要繼續發文？", vbYesNo + vbExclamation + vbDefaultButton2) = vbYes Then
               PUB_Chk103in239 = False
            End If
         End If
      End If
      End With
   End If
   Set adoRst = Nothing
End Function
'Add by Morgan 2009/11/3
'取得相關案語法
Public Function PUB_GetRefCaseSQL(pCN() As String) As String
   Dim stVTable As String
   
   '多國案
   'Modified by Morgan 2019/8/5
   'stVTable = " SELECT CR01,CR02,CR03,CR04 FROM CaseRelation" & _
      " WHERE CR01='" & pCN(1) & "' AND CR02='" & pCN(2) & "' AND CR03='" & pCN(3) & "' AND CR04='" & pCN(4) & "'"
   stVTable = " SELECT CR01,CR02,CR03,CR04 FROM CaseRelation" & _
      " WHERE CR05='" & pCN(1) & "' AND CR06='" & pCN(2) & "' AND CR07='" & pCN(3) & "' AND CR08='" & pCN(4) & "'"
   'end 2019/8/5
   '國內案
   stVTable = stVTable & " UNION SELECT CM05,CM06,CM07,CM08 FROM CASEMAP" & _
      " WHERE CM01='" & pCN(1) & "' AND CM02='" & pCN(2) & "' AND CM03='" & pCN(3) & "' AND CM04='" & pCN(4) & "'"
   '國內案的其他國外案
   stVTable = stVTable & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
      " (SELECT CM05,CM06,CM07,CM08 FROM CASEMAP" & _
      " WHERE CM01='" & pCN(1) & "' AND CM02='" & pCN(2) & "' AND CM03='" & pCN(3) & "' AND CM04='" & pCN(4) & "')"
   '國內案的其他國外案的國外案
   stVTable = stVTable & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
      " (SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
      " (SELECT CM05,CM06,CM07,CM08 FROM CASEMAP" & _
      " WHERE CM01='" & pCN(1) & "' AND CM02='" & pCN(2) & "' AND CM03='" & pCN(3) & "' AND CM04='" & pCN(4) & "'))"
   '國外案
   stVTable = stVTable & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP" & _
      " WHERE CM05='" & pCN(1) & "' AND CM06='" & pCN(2) & "' AND CM07='" & pCN(3) & "' AND CM08='" & pCN(4) & "'"
   '國外案的國外案
   stVTable = stVTable & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
      " (SELECT CM01,CM02,CM03,CM04 FROM CASEMAP" & _
      " WHERE CM05='" & pCN(1) & "' AND CM06='" & pCN(2) & "' AND CM07='" & pCN(3) & "' AND CM08='" & pCN(4) & "')"
   
   'Added by Morgan 2018/12/12
   '國內案的國內案
   stVTable = stVTable & " UNION SELECT CM05,CM06,CM07,CM08 FROM CASEMAP WHERE (CM01,CM02,CM03,CM04) IN" & _
      " (SELECT CM05,CM06,CM07,CM08 FROM CASEMAP" & _
      " WHERE CM01='" & pCN(1) & "' AND CM02='" & pCN(2) & "' AND CM03='" & pCN(3) & "' AND CM04='" & pCN(4) & "')"
   
   '國內案的國內案的國外案
   stVTable = stVTable & " UNION SELECT CM01,CM02,CM03,CM04 FROM CASEMAP WHERE (CM05,CM06,CM07,CM08) IN" & _
      " (SELECT CM05,CM06,CM07,CM08 FROM CASEMAP WHERE (CM01,CM02,CM03,CM04) IN" & _
      " (SELECT CM05,CM06,CM07,CM08 FROM CASEMAP" & _
      " WHERE CM01='" & pCN(1) & "' AND CM02='" & pCN(2) & "' AND CM03='" & pCN(3) & "' AND CM04='" & pCN(4) & "'))"
   'end 2018/12/12
   
   PUB_GetRefCaseSQL = stVTable
End Function

'Add By Sindy 2010/8/18 檢查申請案號或審定號所輸入的長度是否正確
'2011/1/14 MODIFY BY SONIA 台灣審定號核准8碼,核駁7碼,故加傳准駁strTM16
'strType : 1.申請案號 2.審定號
'strTM16 : 1或NULL.准       2.駁
'Modify By Sindy 2017/5/17 strReturnText : 回傳,號數值,去掉空白及跳行符號
Public Function PUB_ChkTm12Tm15Length(ByVal strType As String, ByVal strChkText As String, _
   ByVal strTM01 As String, ByVal strTM02 As String, ByVal strTM03 As String, ByVal strTM04 As String, _
   Optional ByVal strTM10 As String, Optional ByVal strTM16 As String, _
   Optional ByVal bolChooseMsg As Boolean = False, Optional ByRef strReturnText As String) As Boolean
Dim i As Integer
Dim bol308Case As Boolean 'Add By Sindy 2015/3/25 是否為分割案
Dim str308MonCaseApplno As String 'Add By Sindy 2015/3/25 分割母案的申請案號
Dim str308SonCaseApplno As String 'Add By Sindy 2015/3/25 分割子案的申請案號
Dim intQ As Integer, strCon1 As String, rsQ1 As New ADODB.Recordset 'Added by Lydia 2024/03/29

   PUB_ChkTm12Tm15Length = True
   If strChkText = "" Then Exit Function
   'Add By Sindy 2017/5/17
   strChkText = Trim(strChkText) '去掉空白
   strChkText = PUB_StringFilter(strChkText) '去掉跳行符號
   strReturnText = strChkText
   '2017/5/17 END
   
   If bolNewAppNoFormat Then
      '取得申請國家
      If strTM10 = "" And strTM02 <> "" Then
         If strTM03 = "" Then strTM03 = "0"
         If strTM04 = "" Then strTM04 = "00"
         strCon1 = "SELECT TM10 FROM Trademark WHERE TM01='" & strTM01 & "' AND TM02='" & strTM02 & "' AND TM03='" & strTM03 & "' AND TM04='" & strTM04 & "' "
         intQ = 1
         Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
         If intQ = 1 Then
            strTM10 = "" & rsQ1("TM10")
         End If
      End If
      
      'Add By Sindy 2015/3/25 大陸案若存在於分割案件關係檔DivisionCase之分割案號, 申請案號則可多一碼且必須為母案申請案號+A
      bol308Case = False '非分割案
      str308MonCaseApplno = "" '分割母案的申請案號
      str308SonCaseApplno = "" '分割子案的申請案號
      If strTM10 = "020" Then
         strCon1 = "select DivisionCase.*,t1.tm12 T1_TM12,t2.tm12 T2_TM12" & _
                  " From DivisionCase,Trademark t1,Trademark t2" & _
                  " where DC01='" & strTM01 & "' and DC02='" & strTM02 & "' and DC03='" & strTM03 & "' and DC04='" & strTM04 & "'" & _
                  " and DC05=t1.tm01(+) and DC06=t1.tm02(+) and DC07=t1.tm03(+) and DC08=t1.tm04(+)" & _
                  " and DC01=t2.tm01(+) and DC02=t2.tm02(+) and DC03=t2.tm03(+) and DC04=t2.tm04(+)"
         intQ = 1
         Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
         If intQ = 1 Then
            bol308Case = True '分割案
            str308MonCaseApplno = "" & rsQ1.Fields("T1_TM12")
            str308SonCaseApplno = "" & rsQ1.Fields("T2_TM12")
         End If
      End If
      '2015/3/25 END
      
      If strType = "1" Then '1.申請案號
         If (strTM01 = "T" Or strTM01 = "FCT") And strTM10 = "000" Then
            If Len(Trim(strChkText)) <> 9 Then
               If bolChooseMsg = True Then
                  If MsgBox("申請國家為台灣，申請案號只可為9碼數字，不足9碼請在前面補0！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                     PUB_ChkTm12Tm15Length = False
                     Exit Function
                  End If
               Else
                  MsgBox "申請國家為台灣，申請案號只可為9碼數字，不足9碼請在前面補0！", vbExclamation + vbOKOnly
                  PUB_ChkTm12Tm15Length = False
                  Exit Function
               End If
            End If
         End If
         If strTM01 = "T" And strTM10 = "020" Then
            'Modify By Sindy 2011/10/27 大陸開放可以輸入8碼
            'If Len(Trim(strChkText)) > 7 Or Left(Trim(strChkText), 1) = "0" Then
            'Add By Sindy 2015/3/25 大陸增加分割案, 其申請案號為母案之申請案號+A 故碼數會多一碼, 審定號也是.
            If bol308Case = True Then '分割案
               If Trim(strChkText) <> Trim(str308MonCaseApplno) & "A" Then
                  If bolChooseMsg = True Then
                     If MsgBox("申請國家為大陸，子案的申請案號必須為母案申請案號+A！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                        PUB_ChkTm12Tm15Length = False
                        Exit Function
                     End If
                  Else
                     MsgBox "申請國家為大陸，子案的申請案號必須為母案申請案號+A！", vbExclamation + vbOKOnly
                     PUB_ChkTm12Tm15Length = False
                     Exit Function
                  End If
               End If
               If Len(Trim(strChkText)) > 9 Or Left(Trim(strChkText), 1) = "0" Or Right(Trim(strChkText), 1) <> "A" Then
                  If bolChooseMsg = True Then
                     If MsgBox("申請國家為大陸，申請案號不可大於9碼，第一碼不可0且只可為數字，最後一碼必須為A！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                        PUB_ChkTm12Tm15Length = False
                        Exit Function
                     End If
                  Else
                     MsgBox "申請國家為大陸，申請案號不可大於9碼，第一碼不可0且只可為數字，最後一碼必須為A！", vbExclamation + vbOKOnly
                     PUB_ChkTm12Tm15Length = False
                     Exit Function
                  End If
               Else
                  '第一碼不可0且只可為數字1-9
                  If Asc(Mid(Trim(strChkText), 1, 1)) < 49 Or _
                     Asc(Mid(Trim(strChkText), 1, 1)) > 57 Then
                     If bolChooseMsg = True Then
                        If MsgBox("申請國家為大陸，申請案號不可大於9碼，第一碼不可0且只可為數字，最後一碼必須為A！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                           PUB_ChkTm12Tm15Length = False
                           Exit Function
                        End If
                     Else
                        MsgBox "申請國家為大陸，申請案號不可大於9碼，第一碼不可0且只可為數字，最後一碼必須為A！", vbExclamation + vbOKOnly
                        PUB_ChkTm12Tm15Length = False
                        Exit Function
                     End If
                  End If
               End If
            Else
            '2015/3/25 END
               If Len(Trim(strChkText)) > 8 Or Left(Trim(strChkText), 1) = "0" Then
                  'MsgBox "申請國家為大陸，申請案號只可為7碼，第一碼不可0且只可為數字！", vbExclamation + vbOKOnly
                  If bolChooseMsg = True Then
                     'If MsgBox("申請國家為大陸，申請案號不可大於7碼，第一碼不可0且只可為數字！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                     If MsgBox("申請國家為大陸，申請案號不可大於8碼，第一碼不可0且只可為數字！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                        PUB_ChkTm12Tm15Length = False
                        Exit Function
                     End If
                  Else
                     'MsgBox "申請國家為大陸，申請案號不可大於7碼，第一碼不可0且只可為數字！", vbExclamation + vbOKOnly
                     MsgBox "申請國家為大陸，申請案號不可大於8碼，第一碼不可0且只可為數字！", vbExclamation + vbOKOnly
                     PUB_ChkTm12Tm15Length = False
                     Exit Function
                  End If
               Else
                  For i = 1 To 1
                     If i = 1 Then
                        '第一碼不可0且只可為數字1-9
                        If Asc(Mid(Trim(strChkText), i, 1)) < 49 Or _
                           Asc(Mid(Trim(strChkText), i, 1)) > 57 Then
                           'MsgBox "申請國家為大陸，申請案號只可為7碼，第一碼不可0且只可為數字！", vbExclamation + vbOKOnly
                           If bolChooseMsg = True Then
                              'If MsgBox("申請國家為大陸，申請案號不可大於7碼，第一碼不可0且只可為數字！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                              If MsgBox("申請國家為大陸，申請案號不可大於8碼，第一碼不可0且只可為數字！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                                 PUB_ChkTm12Tm15Length = False
                                 Exit Function
                              End If
                           Else
                              'MsgBox "申請國家為大陸，申請案號不可大於7碼，第一碼不可0且只可為數字！", vbExclamation + vbOKOnly
                              MsgBox "申請國家為大陸，申請案號不可大於8碼，第一碼不可0且只可為數字！", vbExclamation + vbOKOnly
                              PUB_ChkTm12Tm15Length = False
                              Exit Function
                           End If
                        End If
                     End If
                  Next i
               End If
            End If
         End If
      
      ElseIf strType = "2" Then '2.審定號
         '2011/1/14 MODIFY BY SONIA 改寫法
         Select Case strTM10
         Case "000"
            '2011/1/14 ADD BY SONIA 台灣核駁審定號0+6碼數字
            'MODIFY BY SONIA 2014/4/2 台灣核駁審定號改為T0+6碼數字,因有舊資料改為二者都可
            Select Case strTM16
               Case "2"
                  If Len(Trim(strChkText)) <> 7 And Len(Trim(strChkText)) <> 8 Then
                     If bolChooseMsg = True Then
                        If MsgBox("申請國家為台灣，核駁審定號必須為7碼或8碼！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                           PUB_ChkTm12Tm15Length = False
                           Exit Function
                        End If
                     Else
                        MsgBox "申請國家為台灣，核駁審定號必須為7碼或8碼！", vbExclamation + vbOKOnly
                        PUB_ChkTm12Tm15Length = False
                        Exit Function
                     End If
                  ElseIf Mid(Trim(strChkText), 1, 1) <> "0" And Mid(Trim(strChkText), 1, 1) <> "T" Then
                     If bolChooseMsg = True Then
                        If MsgBox("申請國家為台灣，核駁審定號第一碼必為0或T！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                           PUB_ChkTm12Tm15Length = False
                           Exit Function
                        End If
                     Else
                        MsgBox "申請國家為台灣，核駁審定號第一碼必為0或T！", vbExclamation + vbOKOnly
                        PUB_ChkTm12Tm15Length = False
                        Exit Function
                     End If
                  'MODIFY BY SONIA 2014/4/2
                  'ElseIf IsNumeric(Mid(Trim(strChkText), 2)) = False Then
                  ElseIf IsNumeric(Right(Trim(strChkText), 6)) = False Then
                     If bolChooseMsg = True Then
                        If MsgBox("申請國家為台灣，核駁審定號後6碼必須為數字！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                           PUB_ChkTm12Tm15Length = False
                           Exit Function
                        End If
                     Else
                        MsgBox "申請國家為台灣，核駁審定號後6碼必須為數字！", vbExclamation + vbOKOnly
                        PUB_ChkTm12Tm15Length = False
                        Exit Function
                     End If
                  End If
               Case Else
               '2011/1/14 END
                  If Len(Trim(strChkText)) <> 8 Then
                     If bolChooseMsg = True Then
                        If MsgBox("申請國家為台灣，核准審定號必須為8碼！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                           PUB_ChkTm12Tm15Length = False
                           Exit Function
                        End If
                     Else
                        MsgBox "申請國家為台灣，核准審定號必須為8碼！", vbExclamation + vbOKOnly
                        PUB_ChkTm12Tm15Length = False
                        Exit Function
                     End If
                  End If
            End Select
         Case "020"
            'Modify By Sindy 2011/10/27 大陸開放可以輸入8碼
            'If Len(Trim(strChkText)) > 7 Or Left(Trim(strChkText), 1) = "0" Then
            'Add By Sindy 2015/3/25 大陸增加分割案, 其申請案號為母案之申請案號+A 故碼數會多一碼, 審定號也是.
            If bol308Case = True Then '分割案:審定號欄必須與申請案號相同
               If Trim(strChkText) <> str308SonCaseApplno Then
                  If bolChooseMsg = True Then
                     If MsgBox("申請國家為大陸，審定號欄必須與申請案號相同！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                        PUB_ChkTm12Tm15Length = False
                        Exit Function
                     End If
                  Else
                     MsgBox "申請國家為大陸，審定號欄必須與申請案號相同！", vbExclamation + vbOKOnly
                     PUB_ChkTm12Tm15Length = False
                     Exit Function
                  End If
               End If
               If Len(Trim(strChkText)) > 9 Or Left(Trim(strChkText), 1) = "0" Or Right(Trim(strChkText), 1) <> "A" Then
                  If bolChooseMsg = True Then
                     If MsgBox("申請國家為大陸，審定號不可大於9碼，第一碼不可0，第一碼可為數字或G，最後一碼必須為A！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                        PUB_ChkTm12Tm15Length = False
                        Exit Function
                     End If
                  Else
                     MsgBox "申請國家為大陸，審定號不可大於9碼，第一碼不可0，第一碼可為數字或G，最後一碼必須為A！", vbExclamation + vbOKOnly
                     PUB_ChkTm12Tm15Length = False
                     Exit Function
                  End If
               Else
                  For i = 1 To 8
                     If i = 1 Then
                        '第一碼可為數字1-9或G(不可0)
                        If Asc(Mid(Trim(strChkText), i, 1)) < 49 Or _
                           (Asc(Mid(Trim(strChkText), i, 1)) > 57 And Asc(Mid(Trim(strChkText), i, 1)) <> 71) Then
                           If bolChooseMsg = True Then
                              If MsgBox("申請國家為大陸，審定號不可大於9碼，第一碼不可0，第一碼可為數字或G，最後一碼必須為A！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                                 PUB_ChkTm12Tm15Length = False
                                 Exit Function
                              End If
                           Else
                              MsgBox "申請國家為大陸，審定號不可大於9碼，第一碼不可0，第一碼可為數字或G，最後一碼必須為A！", vbExclamation + vbOKOnly
                              PUB_ChkTm12Tm15Length = False
                              Exit Function
                           End If
                        End If
                     Else
                        '其他碼只可為數字0-9
                        If Mid(Trim(strChkText), i, 1) <> "" Then
                           If Asc(Mid(Trim(strChkText), i, 1)) < 48 Or _
                              Asc(Mid(Trim(strChkText), i, 1)) > 57 Then
                              If bolChooseMsg = True Then
                                 If MsgBox("申請國家為大陸，審定號不可大於9碼，第一碼不可0，第一碼可為數字或G，最後一碼必須為A！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                                    PUB_ChkTm12Tm15Length = False
                                    Exit Function
                                 End If
                              Else
                                 MsgBox "申請國家為大陸，審定號不可大於9碼，第一碼不可0，第一碼可為數字或G，最後一碼必須為A！", vbExclamation + vbOKOnly
                                 PUB_ChkTm12Tm15Length = False
                                 Exit Function
                              End If
                           End If
                        End If
                     End If
                  Next i
               End If
            Else
            '2015/3/25 END
               'Modified by Morgan 2018/3/27 改若為9碼時，第9碼只能為A(分割案)--桂英
               'If Len(Trim(strChkText)) > 8 Or Left(Trim(strChkText), 1) = "0" Then
               If (Len(Trim(strChkText)) > 8 And Not (Len(Trim(strChkText)) = 9 And Right(Trim(strChkText), 1) = "A")) Or Left(Trim(strChkText), 1) = "0" Then
                  'MsgBox "申請國家為大陸，審定號只可為7碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！", vbExclamation + vbOKOnly
                  If bolChooseMsg = True Then
                     'If MsgBox("申請國家為大陸，審定號不可大於7碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                     If MsgBox("申請國家為大陸，審定號不可大於8碼(若為9碼時，第9碼只能為A)，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                        PUB_ChkTm12Tm15Length = False
                        Exit Function
                     End If
                  Else
                     'MsgBox "申請國家為大陸，審定號不可大於7碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！", vbExclamation + vbOKOnly
                     MsgBox "申請國家為大陸，審定號不可大於8碼(若為9碼時，第9碼只能為A)，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！", vbExclamation + vbOKOnly
                     PUB_ChkTm12Tm15Length = False
                     Exit Function
                  End If
               Else
                  'Modify By Sindy 2011/10/27 大陸開放可以輸入8碼
                  'For i = 1 To 7
                  For i = 1 To 8
                     If i = 1 Then
                        '第一碼可為數字1-9或G(不可0)
                        If Asc(Mid(Trim(strChkText), i, 1)) < 49 Or _
                           (Asc(Mid(Trim(strChkText), i, 1)) > 57 And Asc(Mid(Trim(strChkText), i, 1)) <> 71) Then
                           'MsgBox "申請國家為大陸，審定號只可為7碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！", vbExclamation + vbOKOnly
                           If bolChooseMsg = True Then
                              'If MsgBox("申請國家為大陸，審定號不可大於7碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                              If MsgBox("申請國家為大陸，審定號不可大於8碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                                 PUB_ChkTm12Tm15Length = False
                                 Exit Function
                              End If
                           Else
                              'MsgBox "申請國家為大陸，審定號不可大於7碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！", vbExclamation + vbOKOnly
                              MsgBox "申請國家為大陸，審定號不可大於8碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！", vbExclamation + vbOKOnly
                              PUB_ChkTm12Tm15Length = False
                              Exit Function
                           End If
                        End If
                     Else
                        '其他碼只可為數字0-9
                        If Mid(Trim(strChkText), i, 1) <> "" Then
                           If Asc(Mid(Trim(strChkText), i, 1)) < 48 Or _
                              Asc(Mid(Trim(strChkText), i, 1)) > 57 Then
                              'MsgBox "申請國家為大陸，審定號只可為7碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！", vbExclamation + vbOKOnly
                              If bolChooseMsg = True Then
                                 'If MsgBox("申請國家為大陸，審定號不可大於7碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                                 If MsgBox("申請國家為大陸，審定號不可大於8碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！是否要修改？", vbYesNo + vbDefaultButton2) = vbYes Then
                                    PUB_ChkTm12Tm15Length = False
                                    Exit Function
                                 End If
                              Else
                                 'MsgBox "申請國家為大陸，審定號不可大於7碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！", vbExclamation + vbOKOnly
                                 MsgBox "申請國家為大陸，審定號不可大於8碼，第一碼不可0，第一碼可為數字或G，其他碼只可為數字！", vbExclamation + vbOKOnly
                                 PUB_ChkTm12Tm15Length = False
                                 Exit Function
                              End If
                           End If
                        End If
                     End If
                  Next i
               End If
            End If
         End Select
      End If
   End If
   
   Set rsQ1 = Nothing 'Added by Lydia 2024/03/29
End Function

'Add by Morgan 2010/4/27
'檢查發明人與申請人資料是否搭配
Public Function PUB_ChkInventor(ByRef p_Inventors As String, ByVal p_Applicants As String, Optional bolMsg As Boolean = False) As Boolean
   Dim arrInv
   Dim arrApp
   Dim ii As Integer, jj As Integer, bFound As Boolean, bChanged As Boolean
   Dim strNewInventors As String
      
   bChanged = False
   If p_Inventors <> "" And p_Applicants <> "" Then
      arrInv = Split(p_Inventors, ",")
      arrApp = Split(p_Applicants, ",")
      For ii = LBound(arrInv) To UBound(arrInv)
         If Trim(arrInv(ii) <> "") Then
            bFound = False
            For jj = LBound(arrApp) To UBound(arrApp)
               If Left(arrInv(ii), 8) = Left(arrApp(jj) & "00", 8) Then
                  bFound = True
                  strNewInventors = strNewInventors & "," & arrInv(ii)
                  Exit For
               End If
            Next
         Else
            bFound = True
         End If
         If Not bFound Then bChanged = True
      Next
      strNewInventors = Mid(strNewInventors, 2)
      If bChanged Then
         p_Inventors = strNewInventors
         If bolMsg Then MsgBox "申請人已變更，原申請人之發明人設定將清除！"
      End If
   End If
   PUB_ChkInventor = Not bChanged
End Function
'Add by Morgan 2010/11/22
'設定補文件選項
Public Sub PUB_SetCombo202(oCombo As Object, pCP10 As String)
   oCombo.Clear
   Select Case pCP10
      'Modified by Morgan 2013/5/22 +125
      Case "101", "102", "103", "104", "105", "125" '申請案 '3
          'Memo by Lydia 2021/02/22 若有文字變動，請檢查frm060120補文件抓下一程序備註是否要配合修改；目前設定如下
            '5.優先權證明書 => 優先權
            '6.委任狀 => 委任書
            '7.代表人資訊 => 代表人
            '8.申請人資訊 => 申請人
            '9.發明人資訊 => 發明人
            '10.非WTO會員國之住所證明 => 國籍證明
          'end 2021/02/22
'MODIFY BY SONIA 2014/5/8
'原來
'         oCombo.AddItem "在國外申請日及申請案號說明頁一式三份"
'         oCombo.AddItem "申請權證明書正本一份"
'         oCombo.AddItem "代理人委任書正本一份"
'         oCombo.AddItem "專利申請書一份"
'         oCombo.AddItem "優先權證明文件正本及首頁影本各一份(含首頁中譯文二份)"
         
'修改後
         oCombo.AddItem "代理人委任書正本一份"
         '2015/2/24 MODIFY BY SONIA 智慧局2015/2/23發布訊息
         'oCombo.AddItem "優先權證明文件正本及首頁影本各一份(含首頁中譯文二份)"
         oCombo.AddItem "優先權證明文件正本"
         '2015/2/24 END
         oCombo.AddItem "基本資料表一份(代表人)"     'add by sonia 2019/2/13
         oCombo.AddItem "專利申請書一份(代表人)"
         oCombo.AddItem "專利申請書一份(優先權)"
         oCombo.AddItem "專利申請書一份(發明人中譯名)"
         oCombo.AddItem "專利申請書一份(發明人國籍)"
         oCombo.AddItem "專利申請書一份(申請人中譯名)"
         oCombo.AddItem "專利申請書一份(申請人國籍)"
         oCombo.AddItem "英文摘要一份"               'add by sonia 2019/1/25
'END 2014/5/8
         oCombo.AddItem "發明人拒簽之切結書正本一份"
         'Modify by Morgan 2010/11/15 David
         'oCombo.AddItem "僱傭契約或經認證之讓與文件正本一份"
         oCombo.AddItem "僱傭契約一份"
         oCombo.AddItem "美國讓與文件一份"
         'end 2010/11/15
         oCombo.AddItem "國內寄存證明正本一份"
         oCombo.AddItem "國外寄存證明正本一份"
         oCombo.AddItem "死亡證明正本一份"
         oCombo.AddItem "繼承證明正本一份"
         oCombo.AddItem "法人地位證明書正本一份"
         oCombo.AddItem "國籍證明書正本一份"
         oCombo.AddItem "台籍發明人ID碼" 'Added by Morgan 2021/10/12
         oCombo.AddItem "請求資訊不公開之聲明書" 'Added by Morgan 2022/1/13
         oCombo.AddItem "全部發明人名" 'Added by Morgan 2022/1/17
         oCombo.AddItem "發明人名有特殊字" 'Added by Morgan 2022/1/17
         
         
      Case "107" '再審申請 '5
         oCombo.AddItem "再審理由書"
         
      Case "205" '申復 '2
         oCombo.AddItem "申復理由書"
         oCombo.AddItem "說明書修正本"
         oCombo.AddItem "設計圖說修正本"
         oCombo.AddItem "原文說明書修正本"
         oCombo.AddItem "申請專利範圍修正本"
         oCombo.AddItem "原文申請專利範圍修正本"
         oCombo.AddItem "圖式修正本"
         oCombo.AddItem "設計圖卡修正本"
         oCombo.AddItem "補充說明書"
         
      Case "501" '訴願 '6
         oCombo.AddItem "代理人委任書"
         
      Case "503", "504", "509" '行政訴訟,行政再審,抗告 '7
         oCombo.AddItem "代理人委任書"
         
      Case "507" '行政訴訟上訴 '8
         oCombo.AddItem "代理人委任書"
         
      Case "801" '異議 '9
         oCombo.AddItem "異議理由書一式三份"
         oCombo.AddItem "代理人委任書"
         oCombo.AddItem "法人地位證明書"
         
      Case "802" '異議答辯 '11
         oCombo.AddItem "異議答辯書"
         
      Case "803" '舉發 '10
         oCombo.AddItem "舉發理由書一式四份"
         oCombo.AddItem "代理人委任書"
         oCombo.AddItem "法人地位證明書"
         
      Case "804" '舉發答辯 '12
         oCombo.AddItem "舉發答辯書"
         
      Case Else
         '申請權/專利權異動 '1
         If Left(pCP10, 1) = "7" Then
            oCombo.AddItem "申請權證明書傳真本"
            oCombo.AddItem "法人地位證明書正本"
            oCombo.AddItem "國籍證明書正本"
            oCombo.AddItem "證書正本"
            oCombo.AddItem "讓與契約書正本"
            oCombo.AddItem "授權契約書正本"
         '4
         Else
            oCombo.AddItem "申請專利說明書一式二份"
            oCombo.AddItem "設計圖說一式二份"
            oCombo.AddItem "圖式一式二份"
            oCombo.AddItem "在國外申請日及申請案號"
            oCombo.AddItem "宣誓書正本一份   *發明人已同意簽署申請文件"
            oCombo.AddItem "申請權證明書正本一份   *發明人已同意簽署申請文件"
            oCombo.AddItem "宣誓書傳真本一份   *發明人已同意簽署申請文件"
            oCombo.AddItem "申請權證明書傳真本一份"
            oCombo.AddItem "代理人委任書正本一份"
            oCombo.AddItem "專利申請書一份"
            oCombo.AddItem "優先權證明文件"
            oCombo.AddItem "原文申請專利範圍修正本一式三份"
            oCombo.AddItem "掛號回執影本一份"
            oCombo.AddItem "圖卡一式二份"
            oCombo.AddItem "發明人拒簽之切結書正本一份"
            oCombo.AddItem "僱傭契約或經認證之讓與文件正本一份"
            oCombo.AddItem "死亡證明正本一份"
            oCombo.AddItem "繼承證明正本一份"
            oCombo.AddItem "法人地位證明書正本一份"
            oCombo.AddItem "國籍證明書正本一份"
            oCombo.AddItem "原文說明書一式二份"
            oCombo.AddItem "圖式修正本一式三份"
            oCombo.AddItem "切結書正本一份"
            oCombo.AddItem "證書正本一份"
            oCombo.AddItem "讓與契約書正本一份"
            oCombo.AddItem "授權契約書正本一份"
            oCombo.AddItem "終止授權契約書正本一份"
            
         End If
   End Select
End Sub

'Add By Sindy 2013/11/14
'檢查是否有承辦歷程是否有產生承辦單可以發文
Public Function PUB_IsEmpFlowIsSend(strCP09 As String) As Boolean
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim strCP163 As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, strCPM29 As String

PUB_IsEmpFlowIsSend = True

'Add By Sindy 2020/2/26 此二帳號不檢查歷程是否完成,可直接發文
'             同時將承辦人工作進度中相關日期欄位填入系統日(ENGINEERPROGRESS_BEFORE8)
'P1090　P程序　　puser
'P1091　CFP程序　cfpuser
If strUserNum = "P1090" Or strUserNum = "P1091" Then Exit Function

'Modify By Sindy 2015/5/20 ex.P-109015 AA3028189
'StrSQLa = "select count(*) from empelectronfile where eef01='" & strCP09 & "'"
'rsA.CursorLocation = adUseClient
'rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'If rsA.RecordCount > 0 Then
'   '有承辦歷程
'   If rsA.Fields(0) > 0 Then
'      'Add By Sindy 2013/11/20 增加檢查流程是否有進行到待送件區
'      If PUB_ChkEmpFlowExists(strCP09, EMP_判發) = True Or _
'         PUB_ChkEmpFlowExists(strCP09, EMP_退件重送) = True Then
'      '2013/11/20 END
'         '檢查是否已有產生承辦單
'         rsA.Close
'         StrSQLa = "select eef03 from empelectronfile where eef01='" & strCP09 & "' and instr(upper(eef03),upper('" & EMP_承辦單 & "'))>0"
'         rsA.CursorLocation = adUseClient
'         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount = 0 Then
'            PUB_IsEmpFlowIsSend = False
'            MsgBox "此案件有承辦歷程但未產生電子承辦單, 不可執行發文作業!!!", vbExclamation + vbOKOnly
'         Else
'            'Add By Sindy 2013/11/19
'            '檢查是否已歸檔
'            rsA.Close
'            StrSQLa = "select cpp02 from casepaperpdf where cpp01='" & strCP09 & "' and instr(upper(cpp02),upper('" & EMP_承辦單 & "'))>0"
'            rsA.CursorLocation = adUseClient
'            rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'            If rsA.RecordCount = 0 Then
'               PUB_IsEmpFlowIsSend = False
'               MsgBox "此案件有承辦歷程但尚未歸檔完成, 不可執行發文作業!!!", vbExclamation + vbOKOnly
'            End If
'         End If
'      End If
'   End If
'End If

'Add By Sindy 2020/11/26 是否為多案歷程案件
StrSQLa = "select cp01,cp02,cp03,cp04,cp09,cp163,cpm29 from caseprogress,casepropertymap where cp09='" & strCP09 & "' and cp01=cpm01(+) and cp10=cpm02(+)"
If rsA.State <> 0 Then rsA.Close
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   strCP163 = "" & rsA.Fields("cp163") '多案歷程
   strCP01 = "" & rsA.Fields("cp01")
   'Add By Sindy 2023/9/11
   strCP02 = "" & rsA.Fields("cp02")
   strCP03 = "" & rsA.Fields("cp03")
   strCP04 = "" & rsA.Fields("cp04")
   strCPM29 = "" & rsA.Fields("cpm29")
   '2023/9/11 END
End If
'2020/11/26 END

StrSQLa = "select count(*) from empelectronProcess where eep01='" & strCP09 & "' and eep04 not in(" & EMP_流程控制除外的狀態 & ")"
If rsA.State <> 0 Then rsA.Close
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   '有承辦歷程
   If rsA.Fields(0) > 0 Then
      'Add By Sindy 2013/11/20 增加檢查流程是否有進行到待送件區
      'Modify By Sindy 2018/5/3 + or PUB_ChkEmpFlowExists(strCP09, EMP_送件) = True
      'Modify By Sindy 2018/8/8 + or PUB_ChkEmpFlowExists(strCP09, EMP_發文歸檔) = True
      If PUB_ChkEmpFlowExists(strCP09, EMP_判發) = True Or _
         PUB_ChkEmpFlowExists(strCP09, EMP_送件) = True Or _
         PUB_ChkEmpFlowExists(strCP09, EMP_退件重送) = True Or _
         PUB_ChkEmpFlowExists(strCP09, EMP_發文歸檔) = True Then
      '2013/11/20 END
         'Modify By Sindy 2015/10/2 Mark:承辦單檢查卷宗區有無存在即可,因若檢查歷程附件檔會導至發文日隔天之後,
         '                               若要重新發文時,因歷程附件已被刪光,而無法執行發文
'         '檢查是否已有產生承辦單
'         rsA.Close
'         StrSQLa = "select eef03 from empelectronfile where eef01='" & strCP09 & "' and instr(upper(eef03),upper('" & EMP_承辦單 & "'))>0"
'         rsA.CursorLocation = adUseClient
'         rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
'         If rsA.RecordCount = 0 Then
'            PUB_IsEmpFlowIsSend = False
'            MsgBox "此案件有承辦歷程但未產生電子承辦單, 不可執行發文作業!!!", vbExclamation + vbOKOnly
'         Else
            
            'Modify By Sindy 2021/10/15
            If strCP01 = "ACS" Then
               '檢查是否已歸檔
               StrSQLa = "select count(*) from empelectronProcess where eep01='" & strCP09 & "' and eep04='" & EMP_判發 & "' and eep13='Y'"
               If rsA.State <> 0 Then rsA.Close
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount = 0 Then
                  PUB_IsEmpFlowIsSend = False
                  MsgBox "此案件有承辦歷程但尚未判發完成, 不可執行發文作業!!!", vbExclamation + vbOKOnly
               End If
            'Add By Sindy 2023/11/13 外專
            ElseIf Left(PUB_GetST03(strUserNum), 2) = "F2" Then
               '檢查是否已歸檔
               StrSQLa = "select cpp02 from casepaperpdf where cpp01='" & strCP09 & "' and instr(upper(cpp02),upper('" & EMP_承辦單 & ".menu'))>0"
               If rsA.State <> 0 Then rsA.Close
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount = 0 Then
                  PUB_IsEmpFlowIsSend = False
                  MsgBox "此案件有承辦歷程但尚未歸檔完成, 不可執行發文作業!!!", vbExclamation + vbOKOnly
               End If
            '2023/11/13 END
            Else
            '2021/10/15 END
               'Add By Sindy 2013/11/19
               '檢查是否已歸檔
               'Modify By Sindy 2020/9/30 + EMP_多案承辦單
               StrSQLa = "select cpp02 from casepaperpdf where cpp01='" & strCP09 & "' and (instr(upper(cpp02),upper('" & EMP_承辦單 & "'))>0 or instr(upper(cpp02),upper('" & EMP_多案承辦單 & "'))>0)"
               If rsA.State <> 0 Then rsA.Close
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
               If rsA.RecordCount = 0 Then
                  PUB_IsEmpFlowIsSend = False
                  MsgBox "此案件有承辦歷程但尚未歸檔完成, 不可執行發文作業!!!", vbExclamation + vbOKOnly
               End If
            End If
'         End If
      Else
         'Modify By Sindy 2020/11/26 增加多案歷程案件,調整檢查
         If strCP163 = "" Or (strCP163 <> "" And Len(strCP163) = 9 And strCP163 = strCP09) Then '主操作案件
            PUB_IsEmpFlowIsSend = False
            MsgBox "此案件有承辦歷程尚未完成, 不可執行發文作業!!!", vbExclamation + vbOKOnly
         End If
      End If
   
   'Add By Sindy 2023/9/11 內商發文要檢查無歷程,案件不可發文(台灣案)!以免誤發期限案件!!
   'Modify By Sindy 2023/9/14 排除多案歷程不需管制 + And strCP163 = ""
   ElseIf Left(Pub_StrUserSt03, 2) = "P2" And _
      GetPrjNation1(strCP01 & "-" & strCP02 & "-" & strCP03 & "-" & strCP04) = "000" And _
      strCPM29 <> "N" And strCP163 = "" Then
      PUB_IsEmpFlowIsSend = False
      MsgBox "此案件尚無承辦歷程, 不可執行發文作業!!!", vbExclamation + vbOKOnly
   '2023/9/11 END
   End If
End If

If strCP163 <> "" And Len(strCP163) = 9 And strCP163 <> strCP09 Then '附屬案件
   '檢查主操作案件是否已歸檔
   rsA.Close
   StrSQLa = "select cpp02 from casepaperpdf where cpp01='" & strCP163 & "' and instr(upper(cpp02),upper('" & EMP_多案承辦單 & "'))>0"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount = 0 Then
      PUB_IsEmpFlowIsSend = False
      MsgBox "此案屬【多案歷程】，【主案】尚未歸檔完成, 不可執行發文作業!!!", vbExclamation + vbOKOnly
   Else
      '檢查附屬案件是否已有承辦單
      rsA.Close
      StrSQLa = "select cpp02 from casepaperpdf where cpp01='" & strCP09 & "' and instr(upper(cpp02),upper('" & EMP_多案承辦單 & "'))>0"
      rsA.CursorLocation = adUseClient
      rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
      If rsA.RecordCount = 0 Then
         PUB_IsEmpFlowIsSend = False
         MsgBox "無多案承辦單，請至待送件區發文!!!", vbExclamation + vbOKOnly
      End If
   End If
End If

If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Function

'Add By Sindy 2014/9/3
'公報代理人事務所名稱欄空白清單
Public Function ReadTagentTa04IsNull(Optional strTA05 As String = "") As Boolean
Dim rsTmp As New ADODB.Recordset
Dim strSql As String
Dim strFileName As String
Dim PLeft(1 To 4) As Integer
Dim strTemp(1 To 5) As String, i As Integer
Dim ff1 As Integer
Dim iLine As Integer
   
   PLeft(1) = 500
   PLeft(2) = 2000
   PLeft(3) = 5000
   PLeft(4) = 7000
   ReadTagentTa04IsNull = False
   iLine = 0
   
   strSql = "SELECT * FROM Tagent WHERE ta01='P'"
   If Val(strTA05) > 0 Then
      strSql = strSql & " and (ta04 is null or ta04='' or Ta05=" & DBDATE(strTA05) & ")"
   Else
      strSql = strSql & " and (ta04 is null or ta04='')"
   End If
   strSql = strSql & " order by ta02 asc"
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   ' 有事務所名稱空白時才產生檢核表
   If rsTmp.RecordCount > 0 Then
      Screen.MousePointer = vbHourglass
      rsTmp.MoveFirst
      Do While Not rsTmp.EOF
         '產生文字檔
         If ReadTagentTa04IsNull = False Then
            ReadTagentTa04IsNull = True
            If ff1 > 0 Then Close #ff1
            ff1 = FreeFile
            strFileName = "公報代理人事務所名稱欄空白清單.txt"
            Open PUB_Getdesktop & "\" & strFileName For Output As ff1
            'Print #ff1, "備註：改字型Fixedsys標準11號字以橫式上下左右各10MM列印"
            Print #ff1, "代理人代號  代理人名稱                      建檔時公告日  事務所名稱"
            Print #ff1, "==========  ==============================  ============  =========="
         End If
         For i = 1 To 5
            strTemp(i) = ""
         Next i
         
         strTemp(1) = Trim(rsTmp.Fields("ta01"))
         strTemp(2) = Trim(rsTmp.Fields("ta02"))
         strTemp(3) = Trim(rsTmp.Fields("ta03"))
         strTemp(4) = "" & Trim(rsTmp.Fields("ta04"))
         strTemp(5) = "" & Trim(rsTmp.Fields("ta05"))
         
         strTemp(1) = strTemp(1)
         strTemp(2) = convForm(CheckStr(strTemp(2)), 10)
         strTemp(3) = convForm(CheckStr(strTemp(3)), 30)
         strTemp(4) = strTemp(4)
         strTemp(5) = convForm(CheckStr(strTemp(5)), 12)
         Print #ff1, strTemp(2) & "  " & strTemp(3) & "  " & strTemp(5) & "  " & strTemp(4)
         
         '直接列印出來
         If iLine > 52 Or iLine = 0 Then
            If iLine = 0 Then
               Printer.Orientation = 1 '1.直印 2.橫印
            Else
               Printer.NewPage
            End If
            '列印表頭
            iLine = 1
            Printer.Font.Size = 16
            Printer.Font.Underline = False
            Printer.FontBold = False
            
            Printer.CurrentX = Printer.ScaleWidth / 2 - (Printer.TextWidth("公報代理人事務所名稱欄空白清單") / 2)
            Printer.CurrentY = iLine * 300
            Printer.Print "公報代理人事務所名稱欄空白清單"
            
            Printer.Font.Size = 12
            Printer.Font.Underline = False
            Printer.FontBold = False
            
            iLine = iLine + 1
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = 900
            Printer.Print "列印人員：" & strUserName
            Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
            Printer.CurrentY = 900
            Printer.Print "列印日期：" & ChangeTStringToTDateString(strSrvDate(2))
            iLine = iLine + 1
            Printer.CurrentX = Printer.ScaleWidth - Printer.TextWidth("列印日期：" & ChangeTStringToTDateString(strSrvDate(2))) - 500
            Printer.CurrentY = 1200
            Printer.Print "頁　　次：" & Printer.Page
            
            iLine = 5
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iLine * 300
            Printer.Print "代理人代號"
            Printer.CurrentX = PLeft(2)
            Printer.CurrentY = iLine * 300
            Printer.Print "代理人名稱"
            Printer.CurrentX = PLeft(3)
            Printer.CurrentY = iLine * 300
            Printer.Print "建檔時公告日"
            Printer.CurrentX = PLeft(4)
            Printer.CurrentY = iLine * 300
            Printer.Print "事務所名稱"
            
            iLine = iLine + 1
            Printer.CurrentX = PLeft(1)
            Printer.CurrentY = iLine * 300
            Printer.Print String(145, "-")
            iLine = iLine + 1
         End If
         '列印明細
         For i = 1 To 4
            Printer.CurrentX = PLeft(i)
            Printer.CurrentY = iLine * 300
            If i = 1 Then Printer.Print strTemp(2)
            If i = 2 Then Printer.Print strTemp(3)
            If i = 3 Then Printer.Print strTemp(5)
            If i = 4 Then Printer.Print strTemp(4)
         Next i
         iLine = iLine + 1
         
         rsTmp.MoveNext
      Loop
   End If
   rsTmp.Close
   
   If ReadTagentTa04IsNull = True Then
      Close ff1
      Printer.EndDoc
   End If
      
   Screen.MousePointer = vbDefault
   Set rsTmp = Nothing
End Function

'Add by Lydia 2014/11/18 台灣案主管機關來函輸入，若此案有工程師未發文的程序，發E-MAIL通知工程師收到來函的內容
'Modified by Lydia 2015/04/20 + tDualNO,tDualNO2
'Modified by Lydia 2022/08/16 +申請國家tNA01
'Modified by Morgan 2023/6/27 +bNoCurNo 是否排除本次來函
Public Sub PUB_TaiwanCInputMsg(ByVal tCP01 As String, ByVal tCP02 As String, ByVal tCP03 As String, ByVal tCP04 As String, ByVal tCP10 As String, ByVal tNA01 As String, _
                 Optional tRecNO As String, Optional tDualNo As String, Optional tDualNo2 As String, Optional bNoCurNo As Boolean = False)
'Memo by Lydia 2019/05/20 主要相關請作單
'1.103112403: 最初的請作單
'2.104042102:一案兩請案件加註
'3.106033001: 改郵件主旨:傳入的官方來函-案件性質
'4.108052102: 改Email主旨:P-XXXXXX 尚有未發文(未發文案件性質)，收到(官方來函性質)函，來函請參照卷宗區
Dim rsA1 As ADODB.Recordset, strA1 As String, strA2 As String, strB1 As String, strB2 As String, strB3 As String
Dim mMsgbol As Boolean
Dim intQ As Integer 'Added by Lydia 2024/03/29

On Error GoTo flgErr

If IsNull(tCP03) Then tCP03 = "0"
If IsNull(tCP04) Then tCP04 = "00"

'Added by Lydia 2015/04/20+ tDualNO,tDualNO2
If tCP10 = "1506" Or (tDualNo <> "" And tDualNo2 <> "") Then '1506(智慧局答辯函)一律要發Mail
    mMsgbol = True
Else
    'Remove by Lydia 2017/03/29 有已收文未發文,就發mail通知 (ex.P-107364 少發email,現在改成抓所有已收文未發文的進度的承辦人發信，主旨:傳入的官方來函-案件性質，當有其他案件性質時，內文增加承辦進度的性質。by 玲玲)
    'If Len(tRecNO) > 0 Then strA1 = " and CP09<>'" & tRecNO & "' and CP43<>'" & tRecNO & "' " '排除現在輸入來函和相關總收文號
    'Added by Lydia 2022/08/16 因為P大陸案有核對的機制不會直接上發文,所以排除現在來函的收文號
    'Modified by Morgan 2024/9/27 沒有相關收文號也要通知
    'If tNA01 <> "000" And Len(tRecNO) > 0 Then strA1 = " and CP09<>'" & tRecNO & "' and CP43<>'" & tRecNO & "'"
    If tNA01 <> "000" And Len(tRecNO) > 0 Then strA1 = " and CP09<>'" & tRecNO & "' and (CP43<>'" & tRecNO & "' or CP43 is null)"
    
    If bNoCurNo And tRecNO <> "" Then strA1 = " and CP09<>'" & tRecNO & "'" 'Added by Morgan 2023/6/27
    
    'Memo by Lydia 2018/07/16 P-117827 一般來函收通知補文件(1003)並產生補文件(202),原本就會發通知信,並且符合這個模組所以又發Email通知;與玲玲確認保持現況。
        strA1 = "select distinct cp14 from caseprogress" & _
                " where cp01='" & tCP01 & "' and cp02='" & tCP02 & "' and cp03='" & tCP03 & "' and cp04='" & tCP04 & "' " & _
                " and cp27 is null and cp57 is null " & strA1
    intQ = 1
    Set rsA1 = ClsLawReadRstMsg(intQ, strA1)
    If intQ = 1 Then
       If rsA1.RecordCount > 0 Then
          mMsgbol = True
          'Remove by Lydia 2017/03/29 有已收文未發文,就發mail通知
          'rsA1.MoveFirst
          'Do While Not rsA1.EOF
          '   strB3 = strB3 & rsA1!cp14 & ";"
          '   rsA1.MoveNext
          'Loop
          'end 2017/03/29
       End If
    End If
End If
If mMsgbol = True Then
    'Modified by Morgan 2023/2/8 +m0.cp06
    strA1 = "select m0.cp05,m0.cp09,m0.cp10,cpm03,m0.cp12,m0.cp13,m0.cp14,nvl(pa05,nvl(pa06,pa07)) as pname,pa26,m0.cp64 " & _
            ",m1.cp09 mno1,m1.cp14 as mno2,m0.cp06 From caseprogress m0, casepropertymap, patent,caseprogress m1 " & _
            "where m0.cp01=pa01(+) and m0.cp02=pa02(+) and m0.cp03=pa03(+)and m0.cp04=pa04(+) and m0.cp43=m1.cp09(+) " & _
            "and m0.cp01=cpm01(+) and m0.cp10=cpm02(+) and m0.cp01='" & tCP01 & "' and m0.cp02='" & tCP02 & "' and m0.cp03='" & tCP03 & "' and m0.cp04='" & tCP04 & "' "
     'Remove by Lydia 2017/03/29 有已收文未發文,就發mail通知
    'If Len(tRecNO) > 0 Then
    '   strA1 = strA1 & " and m0.CP09='" & tRecNO & "' " '來函收文號
    'Else
        'Modified by Lydia 2025/01/24 排除
        'strA1 = strA1 & " and m0.cp27 is null and m0.cp57 is null "
        'Modified by Morgan 2025/2/13 只要排除自己，通知補文件可能會更新之前已收文的相關收文號(Ex:P-135004)
        'strA1 = strA1 & " and m0.cp27 is null and m0.cp57 is null " & IIf(tRecNO <> "", " and m0.CP09<>'" & tRecNO & "' and nvl(m0.CP43,'N') <>'" & tRecNO & "' ", "")
        strA1 = strA1 & " and m0.cp27 is null and m0.cp57 is null " & IIf(tRecNO <> "", " and m0.CP09<>'" & tRecNO & "'", "")
    'End If
    'end 2017/03/29
    intQ = 1
    Set rsA1 = ClsLawReadRstMsg(intQ, strA1)
    If intQ = 1 Then
       If rsA1.RecordCount > 0 Then '發mail
            'Modified by Lydia 2017/03/29
            'strA2 = rsA1!cpm03 & PUB_GetRelateCasePropertyName(tRecNO, "1")  '求相關案件性質
            strA2 = GetPrjState6(tCP01, tCP10)
            strA2 = strA2 & PUB_GetRelateCasePropertyName(tRecNO, "1")
            'end 2017/03/29
         
            'Mark by Lydia 2019/05/20
            'strB1 = "已收到 " & tCP01 & "-" & tCP02 & "-" & tCP03 & "-" & tCP04 & " " & LTrim(RTrim(rsA1!cpm03)) & IIf(Right(LTrim(RTrim(rsA1!cpm03)), 1) = "函", "", " 函")
            'Modified by Lydia 2017/03/29 改成抓所有已收文未發文進度的承辦人發信(主旨:傳入的官方來函-案件性質，當有其他案件性質時，內文增加那道進度的性質)
            rsA1.MoveFirst
            strA1 = strB1
            Do While Not rsA1.EOF 'Added by Lydia 2017/03/29
                'Modified by Lydia 2019/05/20 主旨改成: P-XXXXXX 尚有未發文(未發文案件性質)，收到(官方來函性質)函，來函請參照卷宗區
                'strB1 = strA1 'Added by Lydia 2017/03/29 主旨:傳入的官方來函-案件性質 'Memo by Lydia 2019/05/20 輸入來函的案件性質
                strB1 = tCP01 & "-" & tCP02 & "-" & tCP03 & "-" & tCP04 & " 尚有未發文(" & LTrim(RTrim(rsA1!cpm03)) & _
                           ")，收到(" & LTrim(strA2) & IIf(Right(strA2, 1) = "函", "", " 函") & ")"
                'Mark by Lydia 2019/05/20 主旨不用+進度備註
                'If Len(rsA1!CP64) > 0 Then strB1 = strB1 & "(進度備註：" & LTrim(RTrim(rsA1!CP64)) & ")"
                strB1 = strB1 & "，內容請參照卷宗區的電子檔。"
                
                'oContcp
                'Modified by Lydia 2017/03/29 增加案件性質
                'strB2 = "本所案號：" & tCP01 & "-" & tCP02 & "-" & tCP03 & "-" & tCP04 & vbCrLf & _
                        "案件名稱：" & LTrim(RTrim(rsA1!pName)) & vbCrLf & _
                        "申請人　：" & GetCustomerName(rsA1!pa26) & vbCrLf & _
                        "來函日期：" & ChangeWStringToTDateString(rsA1!cp05) & vbCrLf & _
                        "來函性質：" & strA2 & vbCrLf & _
                        "進度備註：" & rsA1!CP64
                strB2 = "本所案號：" & tCP01 & "-" & tCP02 & "-" & tCP03 & "-" & tCP04 & vbCrLf & _
                        "案件名稱：" & LTrim(RTrim(rsA1!pName)) & vbCrLf & _
                        "申請人　：" & GetCustomerName(rsA1!pa26) & vbCrLf & _
                        "來函日期：" & ChangeWStringToTDateString(rsA1!cp05) & vbCrLf & _
                        "來函性質：" & strA2 & vbCrLf
                If tRecNO <> "" And rsA1.Fields("cp09") <> tRecNO Then
                   strB2 = strB2 & "承辦進度：" & rsA1.Fields("cp09") & "　" & rsA1.Fields("cpm03") & vbCrLf
                End If
                strB2 = strB2 & "進度備註：" & rsA1!CP64
                strB3 = "" & rsA1.Fields("cp14")
                'end 2017/03/29
                'Modified by Lydia 2017/03/29 限來函的部份
                'If tCP10 = "1506" Then strB3 = rsA1!mno2 '1506(智慧局答辯函)掛相關總收文號(來函)那道的承辦工程師
                If tCP10 = "1506" And rsA1.Fields("cp09") = tRecNO Then strB3 = rsA1!mno2
                'Added by Lydia 2015/04/20 一案兩請案件若新型已收到通知修正(1201)來函,請於通知承辦工程師已作內部收文的E-MAIL中提醒(tDualNO,tDualNO2)
                If tDualNo <> "" And tDualNo2 <> "" Then
                   strB1 = "(一案兩請)" & strB1
                   strB2 = strB2 & vbCrLf & "　　　　　本案為一案兩請案件，另一" & tDualNo2 & "案為" & tDualNo & "，請研議是否一併辦理修正。"
                   strB3 = rsA1!mno2
                End If
                'Modified by Lydia 2017/03/29 單獨發信給工程師(有承辦人)
                'PUB_SendMail strUserNum, strB3, "", strB1, strB2 '發信給工程師
                If strB3 > "" Then PUB_SendMail strUserNum, strB3, "", strB1, strB2
                
                'Added by Morgan 2023/2/8
                If rsA1.Fields("cp09") < "C" And IsNull(rsA1.Fields("cp06")) Then
                  'Modified by Morgan 2023/8/21 排除(211)準備程序，(212)言詞辯論--玲玲
                  cnnConnection.Execute "update caseprogress set cp06=(select cp06 from caseprogress where cp09='" & tRecNO & "') where cp09='" & rsA1.Fields("cp09") & "' and cp27 is null and cp06 is null and cp10 not in('211','212')"
                End If
                'end 2023/2/8
                rsA1.MoveNext
            Loop
            'end 2017/03/29
       End If
    End If
End If

Set rsA1 = Nothing 'Added by Lydia 2024/03/29
flgErr:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
End Sub

'Added by Lydia 2015/01/05
' 檢查專利年費資料檔記錄是否已經存在
'含大陸領證,年費報價資料維護
Public Function PUB_PYFIsExists(ByVal strYF01 As String, ByVal strYF02 As String, ByVal strYF03 As String, ByVal strYF04 As String, ByVal strYF05 As String) As Boolean
   Dim rsTmp As New ADODB.Recordset
   Dim strSql As String
   
   PUB_PYFIsExists = False
   strSql = "SELECT * FROM PATENTYEARFEE " & _
            "WHERE YF01 = '" & strYF01 & "' AND " & _
                  "YF02 = '" & strYF02 & "' AND " & _
                  "YF03 = '" & strYF03 & "' AND " & _
                  "YF04 = '" & strYF04 & "' AND " & _
                  "YF05 = '" & strYF05 & "' "
                  
   ' 讀取資料庫
   rsTmp.CursorLocation = adUseClient
   rsTmp.Open strSql, cnnConnection, adOpenDynamic
   ' 檢查讀取的資料筆數
   If rsTmp.RecordCount > 0 Then
      PUB_PYFIsExists = True
   Else
      PUB_PYFIsExists = False
   End If
   rsTmp.Close
   Set rsTmp = Nothing
End Function

'Add By Sindy 2015/5/28 解析檔名中的本所案號
Public Function PUB_AnalysisFileNmGetCaseNO(ByVal strFileName As String) As String
Dim ii As Integer, jj As Integer
Dim strSys As String, strCase1 As String, strCase2 As String, strCase3 As String
   
   '65 :A ~ 90 :Z
   '97 :a ~ 122:z
   '45 :-
   '48 :0 ~ 57 :9
   For ii = 1 To Len(strFileName)
      '英文字母
      If (Asc(Mid(strFileName, ii, 1)) >= 65 And Asc(Mid(strFileName, ii, 1)) <= 90) Or _
         (Asc(Mid(strFileName, ii, 1)) >= 97 And Asc(Mid(strFileName, ii, 1)) <= 122) Then
         '系統別
         strSys = strSys & Mid(strFileName, ii, 1)
      '-符號
      ElseIf Asc(Mid(strFileName, ii, 1)) = 45 Then
         jj = jj + 1
      '數字
      ElseIf (Asc(Mid(strFileName, ii, 1)) >= 48 And Asc(Mid(strFileName, ii, 1)) <= 57) Then
         If Len(strCase1) = 6 And jj <= 1 Then jj = 2
         If Len(strCase2) = 1 And jj <= 2 Then jj = 3
         If Len(strCase3) = 2 Then Exit For
         If jj = 0 Or jj = 1 Then
            jj = 1
            strCase1 = strCase1 & Mid(strFileName, ii, 1)
         ElseIf jj = 2 Then
            strCase2 = strCase2 & Mid(strFileName, ii, 1)
         ElseIf jj = 3 Then
            strCase3 = strCase3 & Mid(strFileName, ii, 1)
         Else
            Exit For
         End If
      Else
         Exit For
      End If
   Next ii
   strCase1 = Right("000000" & strCase1, 6)
   strCase2 = Right("0" & strCase2, 1)
   strCase3 = Right("00" & strCase3, 2)
   If strSys <> "" Then
      PUB_AnalysisFileNmGetCaseNO = UCase(strSys) & "-" & strCase1 & "-" & strCase2 & "-" & strCase3
   End If
End Function

'Add By Sindy 2013/6/17 產生申請書
'Modified by Lydia 2019/03/28 +strCP09 收文號
'Modified by Lydia 2019/05/09 +strChkVal勾選項
'Modified by Lydia 2023/07/05 +strSpecVal 特殊
Public Function PUB_GetApplBook(strCaseNo As String, strCP10 As String, _
Optional strApplPer1 As String = "", Optional strApplPer2 As String = "", Optional strApplPer3 As String = "", _
Optional strApplPer4 As String = "", Optional strApplPer5 As String = "", Optional strCP09 As String, Optional strChkVal As String, Optional strSpecVal As String) As Boolean
Dim m_FileName As String ', m_TempFileName As String
Dim m_TM05 As String
Dim m_TM08 As String
Dim mChkType As String 'Added by Lydia 2023/11/14 確定商標種類：特殊商標TM72>商標種類TM08;
Dim m_TM12 As String
Dim m_TM15 As String
Dim m_TM23 As String
Dim m_TM79 As String
Dim m_TM78 As String
Dim m_TM80 As String
Dim m_TM81 As String
Dim strName As String, strText As String
Dim intAppCnt As Integer, intAppCnt2 As Integer
Dim m_CU10 As String
Dim m_CU15 As String
Dim m_CU11 As String
Dim m_CU04 As String
Dim m_CU05 As String
Dim m_CU16 As String
Dim m_CU17 As String
Dim m_CU18 As String
Dim m_CU19 As String
Dim m_CU07 As String
Dim m_CU103 As String
Dim m_CU23 As String, m_CU112 As String
Dim m_CU24 As String
Dim strKey As String
Dim strLineText As String
Dim i As Integer, k As Integer
Dim m_TM24 As String, m_tm25 As String, m_TM82 As String, m_TM83 As String
Dim m_TM84 As String, m_TM85 As String, m_TM86 As String, m_TM87 As String
Dim m_TM88 As String, m_TM89 As String
'Added by Lydia 2019/03/28
Dim m_TM01 As String, m_TM02 As String, m_TM03 As String, m_TM04 As String
Dim m_CP17 As String, m_CP110 As String
Dim rsAD As New ADODB.Recordset '取代rsTemp
Dim intA As Integer
'Added by Lydia 2019/04/10
Dim m_TM09 As String
Dim tmpArr As Variant
'Added by Lydia 2019/09/20 案件備註
Dim m_TM58 As String
'Added by Lydia 2020/02/05
Dim intP As Integer
Dim bolIsManyCase As Boolean 'Add By Sindy 2023/6/28
Dim bolFontBorders As Boolean 'Added by Lydia 2023/07/05 +字元框線:商標種類代號
Dim strCon1 As String, strCon2 As String 'Added by Lydia 2024/03/29

On Error GoTo ErrHand
   
   'Add By Sindy 2023/6/28 變更和移轉才詢問
   bolIsManyCase = False
   If strCP10 = "301" Or strCP10 = "501" Then
      'Added by Lydia 2023/07/05 FCT案改在產生文件前確認
      If Left(strSpecVal, 1) = "M" Then
         bolIsManyCase = True
      Else
      'end 2023/07/05
         If MsgBox("是否為「一文多案」申請書？", vbYesNo, "詢問") = vbYes Then
            bolIsManyCase = True
         End If
      End If 'Added by Lydia 2023/07/05
   End If
   '2023/6/28 END
   
   PUB_GetApplBook = False
   'Modified by Lydia 2019/04/10 +101商申
   If strCP10 = "102" Or strCP10 = "301" Or strCP10 = "501" Or strCP10 = "103" Or strCP10 = "101" Then
      '讀取基本檔資料
      m_TM05 = "": m_TM08 = "": m_TM12 = "": m_TM15 = ""
      m_TM23 = "": m_TM78 = "": m_TM79 = "": m_TM80 = "": m_TM81 = ""
      'Added by Lydia 2019/03/28
      m_TM01 = SystemNumber(strCaseNo, 1)
      m_TM02 = SystemNumber(strCaseNo, 2)
      m_TM03 = SystemNumber(strCaseNo, 3)
      m_TM04 = SystemNumber(strCaseNo, 4)
      'end 2019/03/28
      intAppCnt = 0: intAppCnt2 = 0
      'Modified by Lydia 2019/03/28 SystemNumber=>m_TM01~04
      'strcon1 = "select tm05,tm08,tm12,tm15,tm23,tm78,tm79,tm80,tm81" & _
                  ",tm24,tm25,tm82,tm83,tm84,tm85,tm86,tm87,tm88,tm89" & _
                  " from trademark where tm01=" & CNULL(m_TM01) & _
                  " and tm02=" & CNULL(m_TM02) & _
                  " and tm03=" & CNULL(m_TM03) & _
                  " and tm04=" & CNULL(m_TM04)
      'Modified by Lydia 2019/04/10 +tm09
      'Modified by Lydia 2019/09/20 + tm58
      'Modified by Lydia 2023/11/14
      strCon1 = "select tm05,tm08,tm12,tm15,tm23,tm78,tm79,tm80,tm81" & _
                  ",tm24,tm25,tm82,tm83,tm84,tm85,tm86,tm87,tm88,tm89,cp17,cp110,tm09,tm58,tm72 " & _
                  " from trademark,caseprogress where tm01=" & CNULL(m_TM01) & " and tm02=" & CNULL(m_TM02) & " and tm03=" & CNULL(m_TM03) & " and tm04=" & CNULL(m_TM04) & _
                  " and tm01=cp01(+) and tm02=cp02(+) and tm03=cp03(+) and tm04=cp04(+) and cp09='" & strCP09 & "'"
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strCon1)
      If intA = 1 Then
         m_TM05 = Trim("" & rsAD.Fields("tm05"))
         m_TM08 = Trim("" & rsAD.Fields("tm08"))
         'Added by Lydia 2023/11/14 確定商標種類：特殊商標TM72>商標種類TM08;
         mChkType = IIf(Trim("" & rsAD.Fields("tm72")) <> "", Trim("" & rsAD.Fields("tm72")), Trim("" & rsAD.Fields("tm08")))
         If mChkType = "" Then mChkType = "1"
         'end 2023/11/14
         m_TM09 = Trim("" & rsAD.Fields("tm09")) 'Added by Lydia 2019/04/10
         m_TM12 = Trim("" & rsAD.Fields("tm12"))
         m_TM15 = Trim("" & rsAD.Fields("tm15"))
         m_TM23 = Trim("" & rsAD.Fields("tm23"))
         m_TM58 = Trim("" & rsAD.Fields("tm58")) 'Added by Lydia 2019/09/20
         m_TM78 = Trim("" & rsAD.Fields("tm78"))
         m_TM79 = Trim("" & rsAD.Fields("tm79"))
         m_TM80 = Trim("" & rsAD.Fields("tm80"))
         m_TM81 = Trim("" & rsAD.Fields("tm81"))
         m_TM24 = Trim("" & rsAD.Fields("tm24")) 'Add By Sindy 2017/3/15
         m_tm25 = Trim("" & rsAD.Fields("tm25")) 'Add By Sindy 2017/3/15
         m_TM82 = Trim("" & rsAD.Fields("tm82")) 'Add By Sindy 2017/3/15
         m_TM83 = Trim("" & rsAD.Fields("tm83")) 'Add By Sindy 2017/3/15
         m_TM84 = Trim("" & rsAD.Fields("tm84")) 'Add By Sindy 2017/3/15
         m_TM85 = Trim("" & rsAD.Fields("tm85")) 'Add By Sindy 2017/3/15
         m_TM86 = Trim("" & rsAD.Fields("tm86")) 'Add By Sindy 2017/3/15
         m_TM87 = Trim("" & rsAD.Fields("tm87")) 'Add By Sindy 2017/3/15
         m_TM88 = Trim("" & rsAD.Fields("tm88")) 'Add By Sindy 2017/3/15
         m_TM89 = Trim("" & rsAD.Fields("tm89")) 'Add By Sindy 2017/3/15
         'Added by Lydia 2019/03/28
         m_CP17 = Format(Val("" & rsAD.Fields("cp17")), "#,##0")
         'Added by Lydia 2023/07/05 FCT案改在產生文件前確認
         If Left(strSpecVal, 1) = "M" Then
            m_CP17 = Format(Val(Mid(strSpecVal, 2)), "#,##0")
         End If
         'end 2023/07/05
         m_CP110 = "" & rsAD.Fields("CP110")
         '出名代理人(若無,則改為預設)
         If m_CP110 = "" Then m_CP110 = "94007,81040"
         'end 2019/03/28
         '有輸入申請人時,以輸入資料為主,不然才抓基本檔申請人資料
         If Trim(strApplPer1) = "" Then
            strApplPer1 = m_TM23
            strApplPer2 = m_TM78
            strApplPer3 = m_TM79
            strApplPer4 = m_TM80
            strApplPer5 = m_TM81
         End If
         '共幾人
         If strApplPer1 <> "" Then intAppCnt = intAppCnt + 1
         If strApplPer2 <> "" Then intAppCnt = intAppCnt + 1
         If strApplPer3 <> "" Then intAppCnt = intAppCnt + 1
         If strApplPer4 <> "" Then intAppCnt = intAppCnt + 1
         If strApplPer5 <> "" Then intAppCnt = intAppCnt + 1
         '原申請人共幾人
         If m_TM23 <> "" Then intAppCnt2 = intAppCnt2 + 1
         If m_TM78 <> "" Then intAppCnt2 = intAppCnt2 + 1
         If m_TM79 <> "" Then intAppCnt2 = intAppCnt2 + 1
         If m_TM80 <> "" Then intAppCnt2 = intAppCnt2 + 1
         If m_TM81 <> "" Then intAppCnt2 = intAppCnt2 + 1
      End If
      
       '取得樣本檔
      Select Case strCP10
         Case "102"
            m_FileName = "延展_樣本.doc"
            Call PUB_GetSampleFile(m_FileName, "M51-000100-0-01")
         Case "301"
            If m_TM15 = "" Then
               m_FileName = "註冊前變更_樣本.doc"
               'Add By Sindy 2023/6/28
               If bolIsManyCase = True Then
                  Call PUB_GetSampleFile(m_FileName, "M51-000100-1-02")
               Else
               '2023/6/28 END
                  Call PUB_GetSampleFile(m_FileName, "M51-000100-0-02")
               End If
            Else
               m_FileName = "註冊變更_樣本.doc"
               'Add By Sindy 2023/6/28
               If bolIsManyCase = True Then
                  Call PUB_GetSampleFile(m_FileName, "M51-000100-1-03")
               Else
               '2023/6/28 END
                  Call PUB_GetSampleFile(m_FileName, "M51-000100-0-03")
               End If
            End If
         Case "501"
            m_FileName = "移轉_樣本.doc"
            'Add By Sindy 2023/6/28
            If bolIsManyCase = True Then
               Call PUB_GetSampleFile(m_FileName, "M51-000100-1-04")
            Else
            '2023/6/28 END
               Call PUB_GetSampleFile(m_FileName, "M51-000100-0-04")
            End If
         Case "103"
            m_FileName = "補換發註冊證_樣本.doc"
            Call PUB_GetSampleFile(m_FileName, "M51-000100-0-05")
         'Added by Lydia 2019/04/10
         Case "101" '商申-註冊
            'Modified by Lydia 2023/11/14 m_TM08=>mChkType
            Select Case mChkType
                 Case "7"  '證明標章
                    m_FileName = "證明標章註冊_樣本.doc"
                    Call PUB_GetSampleFile(m_FileName, "M51-000100-0-07")
                 Case "A"  '立體商標
                    m_FileName = "立體商標註冊_樣本.doc"
                    Call PUB_GetSampleFile(m_FileName, "M51-000100-0-08")
                 'Modified by Lydia 2023/11/14 "C"=>"B"
                 Case "B"  '顏色商標
                    m_FileName = "顏色註冊_樣本.doc"
                    Call PUB_GetSampleFile(m_FileName, "M51-000100-0-09")
                 Case Else '商標
                    m_FileName = "商標註冊_樣本.doc"
                    Call PUB_GetSampleFile(m_FileName, "M51-000100-0-06")
            End Select
      End Select
      
      Set rsAD = Nothing 'Added by Lydia 2019/03/28
      
      If Dir(App.path & "\" & m_FileName) <> "" Then
         Screen.MousePointer = vbHourglass
         '判斷word是否已開啟
         If g_WordAp Is Nothing Then
RestarWord:
            Set g_WordAp = New Word.Application
            g_WordAp.Visible = True 'False
         End If
'         If Dir(PUB_Getdesktop & "\" & m_TempFileName) <> "" Then
'            Kill PUB_Getdesktop & "\" & m_TempFileName
'         End If
         g_WordAp.Documents.Open App.path & "\" & m_FileName
'         g_WordAp.ActiveDocument.SaveAs PUB_Getdesktop & "\" & m_TempFileName
'         g_WordAp.ActiveDocument.Close
'         g_WordAp.Documents.Open PUB_Getdesktop & "\" & m_TempFileName
         With g_WordAp
            .Selection.WholeStory
            .Selection.Copy
            'Modified by Lydia 2019/03/28 +晉字,出名代理人,規費,代理人簽章
            'For i = 0 To 7
            'Modified by Lydia 2019/04/10 +優先權聲明,商品類別,商品類別組群
            'For i = 0 To 11
            'Modified by Lydia 2019/05/08 事務所案件編號(本所案號)從申請人1移到第1頁的上方
            'For i = 0 To 14
            'Modified by Lydia 2019/05/15 附件勾選項+商標顏色
            'For i = 0 To 15
            'Modified by Lydia 2019/07/31 附件之含中譯本
            'For i = 0 To 20
            For i = 0 To 23
               strName = ""
               strText = ""
               strLineText = ""
               bolFontBorders = False 'Added by Lydia 2023/07/05
               'Modified by Lydia 2019/03/28
               'If i = 0 Then
               If i = 0 Then
                  strName = "晉字"
                  If m_TM01 = "T" Then
                      strText = "晉商"
                  Else
                      strText = "晉外"
                  End If
               ElseIf i = 1 Then
                  strName = "出名代理人"
                  '(註冊前變更)=>多位代理人時，應將本欄位完整複製後依序填寫
                  'Modified by Lydia 2019/04/10 +101申請(註冊)
                  'If strCP10 = "301" And m_TM15 = "" Then
                  If (strCP10 = "301" And m_TM15 = "") Or strCP10 = "101" Then
                      strText = PUB_GetAgentCP110(strCP09, m_CP110, m_TM01, "3")
                  Else
                      strText = PUB_GetAgentCP110(strCP09, m_CP110, m_TM01, "2")
                  End If
               ElseIf i = 2 Then
                  strName = "規費"
                  strText = m_CP17
               ElseIf i = 3 Then
               'end 2019/03/28
                  strName = "商標種類"
                  If strCP10 = "102" Then '使用於延展
                     'Modified by Lydia 2019/09/20 與電子送件一致,從備註判斷
                     'If m_TM08 = "1" Or m_TM08 = "2" Or m_TM08 = "3" Then
                     If m_TM58 <> "" And InStr(m_TM58, "原為服務標章") > 0 Then
                        strText = "商標種類：□商標  ■商標（92年修正前服務標章）" & vbCrLf & "          "
                     ElseIf m_TM08 = "1" Or m_TM08 = "2" Or m_TM08 = "3" Then
                     'end 2019/09/20
                        strText = "商標種類：■商標  □商標（92年修正前服務標章）" & vbCrLf & "          "
                     ElseIf m_TM08 = "4" Or m_TM08 = "5" Or m_TM08 = "6" Then
                        strText = "商標種類：□商標  ■商標（92年修正前服務標章）" & vbCrLf & "          "
                     Else
                        strText = "商標種類：□商標  □商標（92年修正前服務標章）" & vbCrLf & "          "
                     End If
                     If m_TM08 = "8" Then
                        strText = strText & "■團體標章  "
                     Else
                        strText = strText & "□團體標章  "
                     End If
                     If m_TM08 = "7" Then
                        strText = strText & "■證明標章  "
                     Else
                        strText = strText & "□證明標章  "
                     End If
                     If m_TM08 = "9" Then
                        strText = strText & "■團體商標"
                     Else
                        strText = strText & "□團體商標"
                     End If
                  ElseIf strCP10 = "301" And m_TM15 = "" Then '使用於註冊前變更
                     'Add By Sindy 2023/6/28 變更和移轉一文多案時填代碼
                     If bolIsManyCase = True Then
                        strText = m_TM08
                        If m_TM08 = "1" Or m_TM08 = "2" Or m_TM08 = "3" Or _
                           m_TM08 = "4" Or m_TM08 = "5" Or m_TM08 = "6" Then
                           strText = "T"
                        ElseIf m_TM08 = "9" Then
                           strText = "G"
                        End If
                        bolFontBorders = True 'Added by Lydia 2023/07/05
                     Else
                     '2023/6/28 END
                        If m_TM08 = "1" Or m_TM08 = "2" Or m_TM08 = "3" Or _
                           m_TM08 = "4" Or m_TM08 = "5" Or m_TM08 = "6" Then
                           strText = strText & "■商標  "
                        Else
                           strText = strText & "□商標  "
                        End If
                        If m_TM08 = "8" Then
                           strText = strText & "■團體標章  "
                        Else
                           strText = strText & "□團體標章  "
                        End If
                        If m_TM08 = "7" Then
                           strText = strText & "■證明標章  "
                        Else
                           strText = strText & "□證明標章  "
                        End If
                         If m_TM08 = "9" Then
                           strText = strText & "■團體商標"
                        Else
                           strText = strText & "□團體商標"
                        End If
                     End If
                  'Added by Lydia 2019/04/10
                  ElseIf strCP10 = "101" Then '申請(註冊)
                         GoTo ReadNext
                  Else 'If strCP10 = "301" And m_TM15 <> "" Then '使用於註冊變更
                     'Add By Sindy 2023/6/28 變更和移轉一文多案時填代碼
                     If bolIsManyCase = True Then
                        strText = m_TM08
                        If m_TM08 = "1" Or m_TM08 = "2" Or m_TM08 = "3" Or _
                           m_TM08 = "4" Or m_TM08 = "5" Or m_TM08 = "6" Then
                           strText = "T"
                        ElseIf m_TM08 = "9" Then
                           strText = "G"
                        End If
                        bolFontBorders = True 'Added by Lydia 2023/07/05
                     Else
                     '2023/6/28 END
                        'Modified by Lydia 2019/09/23 與電子送件一致,從備註判斷
                        'If m_TM08 = "1" Or m_TM08 = "2" Or m_TM08 = "3" Then
                        If (m_TM08 = "1" Or m_TM08 = "2" Or m_TM08 = "3") And _
                              Not (m_TM58 <> "" And InStr(m_TM58, "原為服務標章") > 0) Then
                           strText = "■商標  "
                        Else
                           strText = "□商標  "
                        End If
        
                        'Modified by Lydia 2019/09/23 與電子送件一致,從備註判斷
                        'If m_TM08 = "4" Or m_TM08 = "5" Or m_TM08 = "6" Then
                        If m_TM08 = "4" Or m_TM08 = "5" Or m_TM08 = "6" Or _
                             (m_TM58 <> "" And InStr(m_TM58, "原為服務標章") > 0) Then
                           strText = strText & "■商標（92年修正前服務標章）  "
                        Else
                           strText = strText & "□商標（92年修正前服務標章）  "
                        End If
      
                        If m_TM08 = "7" Then
                           strText = strText & "■證明標章  "
                        Else
                           strText = strText & "□證明標章  "
                        End If
                        If m_TM08 = "8" Then
                           strText = strText & "■團體標章  "
                        Else
                           strText = strText & "□團體標章  "
                        End If
                        If m_TM08 = "9" Then
                           strText = strText & "■團體商標"
                        Else
                           strText = strText & "□團體商標"
                        End If
                     End If
                  End If
               'Modified by Lydia 2019/03/28 原本是1,後續編號+3
               ElseIf i = 4 Then
                  strName = "註冊號數"
                  strText = m_TM15
               ElseIf i = 5 Then
                  strName = "商標名稱"
                  strText = m_TM05
               ElseIf i = 6 Then
                  strName = "共幾人"
                  strText = intAppCnt
               ElseIf i = 7 Then
                  strName = "申請人"
                  For k = 1 To 5
                     strKey = ""
                     m_CU10 = "": m_CU15 = "": m_CU11 = "": m_CU04 = "": m_CU05 = ""
                     m_CU16 = "": m_CU17 = "": m_CU18 = "": m_CU19 = ""
                     m_CU07 = "": m_CU103 = "": m_CU23 = "": m_CU24 = "": m_CU112 = ""
                     If k = 1 And strApplPer1 <> "" Then strKey = strApplPer1
                     If k = 2 And strApplPer2 <> "" Then strKey = strApplPer2
                     If k = 3 And strApplPer3 <> "" Then strKey = strApplPer3
                     If k = 4 And strApplPer4 <> "" Then strKey = strApplPer4
                     If k = 5 And strApplPer5 <> "" Then strKey = strApplPer5
                     If strKey <> "" Then
                        'Modified by Lydia 2020/02/05 +代表人1~6 (CU39,CU40,CU42,CU43,CU45,CU46,CU48,CU49,CU51,CU52,CU54,CU55
                        strCon1 = "select cu10,cu15,cu11,cu04,decode(cu05,null,'',nvl(cu05,'')||' '||nvl(cu88,'')||' '||nvl(cu89,'')||' '||nvl(cu90,'')) as cu05,cu16,cu17,cu18,cu19," & _
                                    "cu07,cu103,cu23,decode(cu24,null,'',nvl(cu24,'')||' '||nvl(cu25,'')||' '||nvl(cu26,'')||' '||nvl(cu27,'')||' '||nvl(cu28,'')||' '||nvl(cu102,'')) as cu24,cu112 " & _
                                    ",CU39,CU40,CU42,CU43,CU45,CU46,CU48,CU49,CU51,CU52,CU54,CU55 " & _
                                    "from customer where cu01='" & Left(strKey, 8) & "'" & _
                                    " and cu02='" & Mid(strKey, 9) & "'"
                        intA = 1
                        Set rsAD = ClsLawReadRstMsg(intA, strCon1)
                        If intA = 1 Then
                           m_CU10 = Trim("" & rsAD.Fields("cu10"))
                           m_CU15 = Trim("" & rsAD.Fields("cu15"))
                           m_CU11 = Trim("" & rsAD.Fields("cu11"))
                           m_CU04 = Trim("" & rsAD.Fields("cu04"))
                           m_CU05 = Trim("" & rsAD.Fields("cu05"))
                           m_CU16 = Trim("" & rsAD.Fields("cu16"))
                           m_CU17 = Trim("" & rsAD.Fields("cu17"))
                           m_CU18 = Trim("" & rsAD.Fields("cu18"))
                           m_CU19 = Trim("" & rsAD.Fields("cu19"))
                           'Added by Lydia 2020/02/05 FCT案比照電子送件抓代表人; 因為申請書在發文前產生,發文後會自動更新商標基本檔的代表人和地址 by 阿蓮
                           If m_TM01 = "FCT" Then
                                intP = 0
                                For intA = 1 To 6
                                    If "" & rsAD.Fields("CU" & CStr(39 + 3 * (intA - 1))) & rsAD.Fields("CU" & CStr(40 + 3 * (intA - 1))) <> "" Then
                                        intP = intP + 1
                                        m_CU07 = m_CU07 & " " & intP & "." & IIf("" & rsAD.Fields("CU" & CStr(39 + 3 * (intA - 1))) <> "", rsAD.Fields("CU" & CStr(39 + 3 * (intA - 1))), "（容後補呈）")
                                        m_CU103 = m_CU103 & " " & intP & "." & IIf("" & rsAD.Fields("CU" & CStr(40 + 3 * (intA - 1))) <> "", rsAD.Fields("CU" & CStr(40 + 3 * (intA - 1))), "（容後補呈）")
                                    End If
                                Next intA

                                If intP = 0 Then
                                    m_CU07 = "（容後補呈）"
                                    m_CU103 = "（容後補呈）"
                                ElseIf intP = 1 Then '只有代表人1，去掉 1.
                                   m_CU07 = Mid(m_CU07, 4)
                                   m_CU103 = Mid(m_CU103, 4)
                                Else
                                   m_CU07 = Mid(m_CU07, 2)
                                   m_CU103 = Mid(m_CU103, 2)
                                End If
                           Else
                           'end 2020/02/05
                                m_CU07 = Trim("" & rsAD.Fields("cu07"))
                                m_CU103 = Trim("" & rsAD.Fields("cu103"))
                           End If  'end 2020/02/05
                           m_CU23 = Trim("" & rsAD.Fields("cu23"))
                           m_CU24 = Trim("" & rsAD.Fields("cu24"))
                           m_CU112 = Trim("" & rsAD.Fields("cu112"))
                        End If
                        'Modify By Sindy 2017/3/15
                        'T-152054 除了301變更501移轉抓客戶地址外
                        '         其餘應抓商標基本檔之申請地址
                        'Modified by Lydia 2020/02/05 排除FCT案 ; FCT案統一抓客戶地址
                        'If strCP10 <> "301" And strCP10 <> "501" Then
                        If m_TM01 <> "FCT" And strCP10 <> "301" And strCP10 <> "501" Then
                           m_CU23 = ""
                           If k = 1 Then
                              m_CU112 = m_TM24: m_CU24 = m_tm25
                           ElseIf k = 2 Then
                              m_CU112 = m_TM82: m_CU24 = m_TM86
                           ElseIf k = 3 Then
                              m_CU112 = m_TM83: m_CU24 = m_TM87
                           ElseIf k = 4 Then
                              m_CU112 = m_TM84: m_CU24 = m_TM88
                           ElseIf k = 5 Then
                              m_CU112 = m_TM85: m_CU24 = m_TM89
                           End If
                        End If
                        '2017/3/15 END
                        strText = strText & "   （第" & k & "申請人）" & vbCrLf
                        If m_CU10 = "" Then
                           strText = strText & "國    籍：□中華民國 □大陸地區（□大陸、□香港、□澳門）" & vbCrLf
                           strText = strText & "          □外國籍：" & vbCrLf
                        Else
                           If m_CU10 <= "010" Then
                              strText = strText & "國    籍：■中華民國 □大陸地區（□大陸、□香港、□澳門）" & vbCrLf
                              strText = strText & "          □外國籍：" & vbCrLf
                           ElseIf m_CU10 = "020" Then
                              strText = strText & "國    籍：□中華民國 ■大陸地區（■大陸、□香港、□澳門）" & vbCrLf
                              strText = strText & "          □外國籍：" & vbCrLf
                           ElseIf m_CU10 = "013" Then
                              strText = strText & "國    籍：□中華民國 ■大陸地區（□大陸、■香港、□澳門）" & vbCrLf
                              strText = strText & "          □外國籍：" & vbCrLf
                           ElseIf m_CU10 = "044" Then
                              strText = strText & "國    籍：□中華民國 ■大陸地區（□大陸、□香港、■澳門）" & vbCrLf
                              strText = strText & "          □外國籍：" & vbCrLf
                           Else
                              strText = strText & "國    籍：□中華民國 □大陸地區（□大陸、□香港、□澳門）" & vbCrLf
                              strText = strText & "          ■外國籍：" & GetPrjNationName(m_CU10) & vbCrLf
                              strLineText = GetPrjNationName("m_CU10")
                           End If
                        End If
                        If m_CU15 = "0" Then
                           strText = strText & "身分種類：■自然人               □法人、公司、機關、學校" & vbCrLf
                           strText = strText & "          □商號、行號、工廠" & vbCrLf
                        Else
                           strText = strText & "身分種類：□自然人               ■法人、公司、機關、學校" & vbCrLf
                           strText = strText & "          □商號、行號、工廠" & vbCrLf
                        End If
                        strText = strText & "ID：　" & m_CU11 & vbCrLf
                        'Added by Lydia 2019/07/10 對智慧局文件要+國籍(阿蓮: 修法後一直用人工修改)
                        'Modified by Lydia 2019/07/31 +公司別
                        'If m_CU10 > "010" Then
                        If m_CU10 > "010" And m_CU15 = "1" Then
                            'Remove by Lydia 2019/08/29 大陸改成大陸商
'                            If m_CU10 = "020" Then
'                                'Modified by Lydia 2019/08/05 ８月份有收到智慧局行文於七月討論結果，決定外商後不再加點或空格，並於８月１日開始施行。
'                                'strText = strText & "申請人名稱（中文）：" & "大陸地區．" & m_CU04 & vbCrLf
'                                strText = strText & "申請人名稱（中文）：" & "大陸地區" & m_CU04 & vbCrLf
'                            Else
                                'Modified by Lydia 2019/07/31 外商:國籍xx商與申請人名稱之間 , 固定都加‧區隔
                                'Modified by Lydia 2019/08/05 ８月份有收到智慧局行文於七月討論結果，決定外商後不再加點或空格，並於８月１日開始施行。
                                'strText = strText & "申請人名稱（中文）：" & GetPrjNationName(m_CU10, "NA81") & "‧" & m_CU04 & vbCrLf
                                strText = strText & "申請人名稱（中文）：" & GetPrjNationName(m_CU10, "NA81") & m_CU04 & vbCrLf
'                            End If 'end 2019/08/29
                        'Added by Lydia 2019/08/29 FCT和T的自然人+XX籍(紙本)
                        ElseIf m_CU10 > "010" And m_CU15 = "0" Then
                                strText = strText & "申請人名稱（中文）：" & Replace(GetPrjNationName(m_CU10, "NA81"), "商", "籍") & m_CU04 & vbCrLf
                        
                        Else
                        'end 2019/07/10
                            strText = strText & "申請人名稱（中文）：" & m_CU04 & vbCrLf
                        End If
                        
                        'Remove by Lydia 2019/05/08
                        'If k = 1 Then '第1個申請人時才需要出現案號
                        '   'Modified by Lydia 2019/03/28 SystemNumber=>m_TM02
                        '   'Added by Lydia 2019/04/10 外商的編號為2
                        '   If m_TM01 = "FCT" Then
                        '        strText = strText & "                     ※事務所案件編號2-" & m_TM02 & vbCrLf
                        '   Else
                        '   'end 2019/04/10
                        '        strText = strText & "                     ※事務所案件編號1-" & m_TM02 & vbCrLf
                        '   End If
                        'End If
                        'end 2019/05/08
                        strText = strText & "          （英文）：" & m_CU05 & vbCrLf
                        strText = strText & "    代表人（中文）：" & m_CU07 & vbCrLf
                        strText = strText & "          （英文）：" & m_CU103 & vbCrLf
                        strText = strText & "地      址（中文）：" & PUB_ChgNumeralStyle(IIf(m_CU112 <> "", m_CU112 & " ", "") & m_CU23) & vbCrLf 'Modify By Sindy 2015/9/21 全形轉半形 m_CU23 ==> PUB_ChgNumeralStyle(m_CU112 & " " & m_CU23)
                        strText = strText & "          （英文）：" & m_CU24 & vbCrLf
                        If intAppCnt > 1 And k = 1 Then '多申請人且為第1個申請人時才需勾選
                           strText = strText & "■此申請人為選定代表人" & vbCrLf
                        Else
                           strText = strText & "□此申請人為選定代表人" & vbCrLf
                        End If
                        'Modify By Sindy 2015/10/27 嘉雯:所有申請書內容皆不載入申請人聯絡電話及傳真號碼
                        strText = strText & "聯絡電話及分機：" & vbCrLf '& IIf(m_CU16 <> "", m_CU16, m_CU17) & vbCrLf
                        'Modified by Lydia 2019/07/31 +vbCrLf
                        strText = strText & "傳  真：" & vbCrLf '& IIf(m_CU18 <> "", m_CU18, m_CU19) '& vbCrLf
                        'strText = strText & "E-MAIL：" 'E-Mail不需要帶資料
                     End If
                  Next k
               ElseIf i = 8 Then
                  strName = "申請案號"
                  strText = m_TM12
                  If strCP10 = "101" Then GoTo ReadNext 'Added by Lydia 2019/04/10
               ElseIf i = 9 Then
                  strName = "原申請人共幾人"
                  strText = intAppCnt2
                  If strCP10 = "101" Then GoTo ReadNext 'Added by Lydia 2019/04/10
               ElseIf i = 10 Then
                  strName = "原申請人"
                  If strCP10 = "101" Then GoTo ReadNext 'Added by Lydia 2019/04/10
                  For k = 1 To 5
                     strKey = ""
                     m_CU04 = "": m_CU05 = "": m_CU11 = ""
                     If k = 1 And m_TM23 <> "" Then strKey = m_TM23
                     If k = 2 And m_TM78 <> "" Then strKey = m_TM78
                     If k = 3 And m_TM79 <> "" Then strKey = m_TM79
                     If k = 4 And m_TM80 <> "" Then strKey = m_TM80
                     If k = 5 And m_TM81 <> "" Then strKey = m_TM81
                     If strKey <> "" Then
                        strCon1 = "select cu10,cu15,cu11,cu04,decode(cu05,null,'',nvl(cu05,'')||' '||nvl(cu88,'')||' '||nvl(cu89,'')||' '||nvl(cu90,'')) as cu05,cu16,cu17,cu18,cu19," & _
                                    "cu07,cu103,cu23,decode(cu24,null,'',nvl(cu24,'')||' '||nvl(cu25,'')||' '||nvl(cu26,'')||' '||nvl(cu27,'')||' '||nvl(cu28,'')||' '||nvl(cu102,'')) as cu24 " & _
                                    "from customer where cu01='" & Left(strKey, 8) & "'" & _
                                    " and cu02='" & Mid(strKey, 9) & "'"
                        intA = 1
                        Set rsAD = ClsLawReadRstMsg(intA, strCon1)
                        If intA = 1 Then
                           m_CU04 = Trim("" & rsAD.Fields("cu04"))
                           m_CU05 = Trim("" & rsAD.Fields("cu05"))
                           m_CU11 = Trim("" & rsAD.Fields("cu11"))
                        End If
                        strText = strText & "   （第" & k & "申請人）" & vbCrLf
                        strText = strText & "ID：　" & m_CU11 & vbCrLf
                        strText = strText & "申請人名稱（中文）：" & m_CU04 & vbCrLf
                        strText = strText & "          （英文）：" & m_CU05 & vbCrLf
                     End If
                  Next k
               'Added by Lydia 2019/03/28
               ElseIf i = 11 Then
                    strName = "代理人簽章"
                    'Mark by Lydia 2019/07/31 預設都有代理人簽章
                    'If InStr("301,501,103", strCP10) > 0 Then
                        strText = PUB_GetAgentCP110(strCP09, m_CP110, m_TM01, "5")
                    'Else
                    '    GoTo ReadNext
                    'End If
               'end 2019/03/28
               'Added by Lydia 2019/04/10
               ElseIf i = 12 Then
                    strName = "優先權聲明"
                    If strCP10 <> "101" Then GoTo ReadNext
                    
                    strCon1 = "select PD05 ,PD06,NA03,NVL(A1.TM01||A1.TM02||A1.TM03||A1.TM04,A2.TM01||A2.TM02||A2.TM03||A2.TM04) AS caseno,PD10 " & _
                             "from PRIDATE,NATION,TRADEMARK A1,TRADEMARK A2 " & _
                             "WHERE PD01='" & m_TM01 & "' AND PD02='" & m_TM02 & "' AND PD03='" & m_TM03 & "' AND PD04 ='" & m_TM04 & "' " & _
                             "AND PD06=A1.TM12(+) AND PD05=A1.TM11(+) AND PD07=A1.TM10(+) " & _
                             "AND PD06=A2.TM15(+) AND PD05=A2.TM11(+) AND PD07=A2.TM10(+) " & _
                             "AND PD07=NA01(+) " & _
                             "ORDER BY PD01,PD02,PD03,PD04"
                    intA = 1
                    Set rsAD = ClsLawReadRstMsg(intA, strCon1)
                    If intA = 1 Then
                        With rsAD
                            .MoveFirst
                            Do While Not .EOF
                                 strCon2 = Val(Left(PUB_DBYEAR("" & .Fields("pd05")), 4)) - 1911
                                 strText = strText & IIf(strText <> "", vbCrLf, "") & "優先權日：民國" & strCon2 & "年" & PUB_DBMONTH("" & .Fields("pd05")) & "月" & PUB_DBDAY("" & .Fields("pd05")) & "日" & vbCrLf & _
                                               "第一次申請國家 (地區)：" & .Fields("na03") & vbCrLf & _
                                               "案　　號：" & .Fields("PD06")
                                 .MoveNext
                            Loop
                        End With
                    Else
                        strText = "優先權日：民國   年   月   日" & vbCrLf & _
                                      "第一次申請國家 (地區)：" & vbCrLf & _
                                      "案　　號："
                    End If
               ElseIf i = 13 Then
                    strName = "商品類別"
                    If strCP10 <> "101" Then GoTo ReadNext
                    'Modified by Lydia 2023/11/14
                    'If InStr("1,A,C", IIf(m_TM08 = "", "1", m_TM08)) = 0 Then GoTo ReadNext
                    If InStr("1,A,B", mChkType) = 0 Then GoTo ReadNext
                    strText = m_TM09
               ElseIf i = 14 Then
                    strName = "商品類別組群"
                    If strCP10 <> "101" Then GoTo ReadNext
                    'Modified by Lydia 2023/11/14
                    'If InStr("1,A,C", IIf(m_TM08 = "", "1", m_TM08)) = 0 Then GoTo ReadNext
                    If InStr("1,A,B", IIf(m_TM08 = "", "1", m_TM08)) = 0 Then GoTo ReadNext
                    strCon1 = BeforePrintGetDBData("TMGoods:" & m_TM01 & "-" & m_TM02 & "-" & m_TM03 & "-" & m_TM04 & "-||區隔", True)
                    If Trim(strCon1) <> "" Then
                         tmpArr = Empty
                         tmpArr = Split(strCon1, "||")
                         For intA = 0 To UBound(tmpArr)
                             strCon2 = Trim(tmpArr(intA))
                             If strCon2 <> "" Then
                                  'Modified by Lydia 2019/07/31 阿蓮:FCT不要組群代碼
                                  strText = strText & IIf(strText <> "", vbCrLf, "") & _
                                              "類別：" & Mid(strCon2, 1, InStr(strCon2, "：") - 1) & vbCrLf & _
                                              "商品／服務名稱：" & Mid(strCon2, InStr(strCon2, "：") + 1) & vbCrLf & _
                                             IIf(m_TM01 = "T", "※組群代碼：" & vbCrLf, "")
                             End If
                         Next intA
                    'Added by Lydia 2019/07/31 阿蓮: 因為在產生申請書時才撰寫商品服務,所以依收文的類別產生
                    'Memo by Lydia 2019/07/31 嘉雯: 內商的商品服務類別內容是用智慧局插件產生,代空白也可以,而外商多半無法用智慧局所以才用人工撰寫
                    ElseIf m_TM09 <> "" Then
                         tmpArr = Empty
                         tmpArr = Split(m_TM09, ",")
                         For intA = 0 To UBound(tmpArr)
                             strCon2 = Trim(tmpArr(intA))
                             If strCon2 <> "" Then
                                  strText = strText & IIf(strText <> "", vbCrLf, "") & _
                                              "類別：" & strCon2 & vbCrLf & _
                                              "商品／服務名稱：" & vbCrLf & vbCrLf & vbCrLf & _
                                             IIf(m_TM01 = "T", "※組群代碼：" & vbCrLf, "")
                             End If
                         Next intA
                    'end 2019/07/31
                    Else
                         strText = "類別：" & vbCrLf & _
                                      "商品／服務名稱：" & vbCrLf & vbCrLf & vbCrLf & _
                                      "※組群代碼："
                    End If
               'end 2019/04/10
               End If
               
               '必須依各申請書不會使用到的置換資料,設定成略過
               '註冊前變更
               'Modified by Lydia 2019/05/09 改判斷
               'If strCP10 = "301" And m_TM15 = "" Then
               '   'Modified by Lydia 2019/03/28
               '   'If i = 1 Then GoTo ReadNext
               '   If i = 4 Then GoTo ReadNext
               If i = 4 Then '註冊號數
                   If strCP10 = "301" And m_TM15 = "" Then
                       GoTo ReadNext
                   End If
               'Added by Lydia 2019/03/28 判斷是否略過代理人簽章
               ElseIf i = 11 Then
               'If i = 11 Then
                  'Mark by Lydia 2019/07/31 預設都有代理人簽章
                  'If InStr("301,501,103", strCP10) = 0 Then GoTo ReadNext
               'end 2019/03/28

               'Added by Lydia 2019/05/08 全部申請(事務所案件編號)
               ElseIf i = 15 Then
                    strName = "案號"
                    If m_TM01 = "FCT" Then
                         strText = strText & "2-" & m_TM02
                    Else
                         strText = strText & "1-" & m_TM02
                    End If
               'Added by Lydia 2019/07/31 附件之含中譯本
               ElseIf i >= 21 And i <= 23 Then
                   If m_TM01 = "T" Then '內商預設(嘉雯: 不用中譯本)
                        If strCP10 = "101" Then
                            strChkVal = "B1、B2、B3"
                        Else
                            strChkVal = "B1"
                        End If
                   End If
                    strText = ""
                    If strChkVal <> "" Then
                        If InStr(strChkVal, "A" & i - 20) > 0 Then
                           strName = "C" & i - 20
                           strText = "■"  '阿蓮:附件都要含中譯本
                        ElseIf InStr(strChkVal, "B" & i - 20) > 0 Then
                           strName = "C" & i - 20
                           strText = "□"
                        End If
                        If strCP10 <> "101" And i > 21 Then '非101申請只有委任書有中譯本
                            strText = ""
                        End If
                    End If
                    If strText = "" Then GoTo ReadNext
               'Added by Lydia 2019/05/09 附件勾選項
               ElseIf i >= 16 And i <= 19 Then
                  '101(註冊)申請:A1委任書、A2優先權、A3展覽會優先權
                  '102延展: A1委任書、A2變更證明文件
                  '301註冊前變更: A1委任書、A2具結書
                  '301註冊變更: A1委任書
                  '103補換發證書: A1委任書
                  If m_TM01 = "FCT" Then
                     strText = ""
                     'Modify By Sindy 2023/6/28 移轉沒有A1 + And strCP10 <> "501"
                     If strChkVal <> "" And strCP10 <> "501" Then
                           If InStr(strChkVal, "A" & i - 15) > 0 Then
                              strName = "A" & i - 15
                              strText = "■"
                           ElseIf InStr(strChkVal, "B" & i - 15) > 0 Then
                              strName = "A" & i - 15
                              strText = "□"
                           End If
                     End If
                     If strText = "" Then GoTo ReadNext
                  Else
                     'Added by Lydia 2019/07/04 T只有3個勾選項
                     If m_TM01 = "T" And i >= 19 Then
                         If strText = "" Then GoTo ReadNext
                     End If
                     
                     If strCP10 = "102" Or (strCP10 = "301" And m_TM15 = "") Then
                          strChkVal = "A1、A2"
                     'Modified by Lydia 2019/07/31 嘉雯: 申請只要預設委任書(strCP10 = "101")
                     ElseIf strCP10 = "101" Or strCP10 = "103" Or (strCP10 = "301" And m_TM15 <> "") Then
                          strChkVal = "A1"
                     End If
                     If strChkVal <> "" And InStr(strChkVal, "A" & i - 15) > 0 Then
                           strName = "A" & i - 15
                           strText = "■"
                     'Added by Lydia 2019/07/04
                     ElseIf strCP10 = "101" Then
                           strName = "A" & i - 15
                           strText = "□"
                     End If
                  End If
               'Added by Lydia 2019/05/15 商標顏色
               ElseIf i = 20 Then
                  If strCP10 <> "101" Then GoTo ReadNext
                  
                  strName = "商標顏色"
                  If m_TM08 = "C" Then '顏色商標
                       strText = "彩色"
                  Else
                        If m_TM01 = "T" Then
                            strText = "■墨色  □彩色"
                        ElseIf m_TM01 = "FCT" Then
                            If InStr(strChkVal, "彩") > 0 Then
                                 strText = "□墨色  ■彩色"
                            Else
                                 strText = "■墨色  □彩色"
                            End If
                        End If
                  End If
               Else
                  'Modified by Lydia 2019/03/28
                  'If i >= 5 Then GoTo ReadNext
                  'Modified by Lydia 2019/04/10
                  'If i >= 8 Then GoTo ReadNext
                  'Modified by Lydia 2019/05/09 改判斷
                  'If strCP10 <> "101" And i >= 8 Then GoTo ReadNext
                  If strCP10 = "301" And m_TM15 = "" Then
                  Else
                       If strCP10 <> "101" And i >= 8 Then GoTo ReadNext
                  End If
               End If
               
               'Find並且置換
               If Trim(strName) <> "" Then
                  .Selection.Find.ClearFormatting
                  .Selection.Find.Text = "|#" & strName & "#|"
                  .Selection.Find.Replacement.Text = ""
                  .Selection.Find.Forward = True
                  .Selection.Find.Wrap = wdFindContinue
                  .Selection.Find.Format = False
                  .Selection.Find.MatchCase = False
                  .Selection.Find.MatchWholeWord = False
                  .Selection.Find.MatchWildcards = False
                  .Selection.Find.MatchSoundsLike = False
                  .Selection.Find.MatchAllWordForms = False
                  .Selection.Find.MatchByte = True
                  .Selection.Find.Execute
                  .Selection.Delete
                  'Added by Lydia 2023/07/05 變更和移轉一文多案時填商標種類代碼，字元要加框線
                  If bolFontBorders = True Then
                    .Selection.Font.Borders(1).LineStyle = wdLineStyleSingle
                  End If
                  'end 2023/07/05
                  .Selection.TypeText strText
                  If strLineText <> "" Then
                     .Selection.HomeKey
                     .Selection.Find.ClearFormatting
                     With .Selection.Find
                         .Text = strLineText
                         .Replacement.Text = ""
                         .Forward = True
                         .Wrap = wdFindContinue
                         .Format = False
                         .MatchCase = False
                         .MatchWholeWord = False
                         .MatchWildcards = False
                         .MatchSoundsLike = False
                         .MatchAllWordForms = False
                         .MatchByte = True
                     End With
                     .Selection.Find.Execute
                     .Selection.Font.Underline = wdUnderlineSingle
                  End If
               End If
ReadNext:
            Next i
         End With
         Screen.MousePointer = vbDefault
'         g_WordAp.ActiveDocument.Save
'         g_WordAp.ActiveDocument.Close
'         MsgBox "檔案已存放在：" & PUB_Getdesktop & "\" & m_TempFileName
         MsgBox "資料已產生完畢!!!"
         PUB_GetApplBook = True
      Else
         MsgBox "無申請書的樣本!!!"
      End If
   End If
   
   Exit Function
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   End If
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function

'Add By Sindy 2015/10/27 產生申請書
'Add By Sindy 2018/4/19 + pa() As String
Public Function PUB_GetApplBook_FCP(pa() As String, strCaseNo As String, strCP10 As String, _
Optional strApplPer1 As String = "", Optional strApplPer2 As String = "", Optional strApplPer3 As String = "", _
Optional strApplPer4 As String = "", Optional strApplPer5 As String = "", Optional strCP09 As String = "") As Boolean

Dim m_FileName As String
Dim m_PA05 As String, m_PA06 As String, m_PA09 As String, m_PA11 As String
Dim strName As String, strText As String
Dim intAppCnt As Integer, intAppCnt2 As Integer
Dim strKey As String
Dim strLineText As String
Dim i As Integer, k As Integer
Dim m_CP84 As String, m_CP27 As String, m_CP28 As String
Dim m_CP110 As String, m_CP135 As String, m_CP136 As String
Dim strApplLineText As String
Dim strApplData As String
Dim strCon1 As String, intQ As Integer, rsQ1 As New ADODB.Recordset 'Added by Lydia 2024/03/29

On Error GoTo ErrHand
   
   PUB_GetApplBook_FCP = False
   
   m_MySt(1) = SystemNumber(strCaseNo, 1)
   m_MySt(2) = SystemNumber(strCaseNo, 2)
   m_MySt(3) = SystemNumber(strCaseNo, 3)
   m_MySt(4) = SystemNumber(strCaseNo, 4)
   m_SysKind = CheckSys(m_MySt(1))
   SetLetterSt
   
   If strCP10 = 實體審查 Then
      '讀取基本檔資料
      strCon1 = "select * from patent where pa01=" & CNULL(m_MySt(1)) & _
                  " and pa02=" & CNULL(m_MySt(2)) & _
                  " and pa03=" & CNULL(m_MySt(3)) & _
                  " and pa04=" & CNULL(m_MySt(4))
      intQ = 1
      Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
      If intQ = 1 Then
         m_PA05 = Trim("" & rsQ1.Fields("PA05"))
         m_PA06 = Trim("" & rsQ1.Fields("PA06"))
         m_PA09 = Trim("" & rsQ1.Fields("PA09")) '申請國家
         m_PA11 = Trim("" & rsQ1.Fields("PA11")) '申請案號
         '取得申請人資料
         'Modify By Sindy 2016/3/9
         'strApplData = PUB_GetApplData(m_MySt(1), m_MySt(2), m_MySt(3), m_MySt(4), strApplPer1, strApplPer2, strApplPer3, strApplPer4, strApplPer5, strApplLineText, intAppCnt, intAppCnt2)
         strApplData = PUB_GetApplData(pa(), m_MySt(1), m_MySt(2), m_MySt(3), m_MySt(4), strApplPer1, strApplPer2, strApplPer3, strApplPer4, strApplPer5, strApplLineText, intAppCnt, intAppCnt2, strCP10)
         '2016/3/9 END
      End If
      
      If strCP09 <> "" Then
         strCon1 = "select * from caseprogress where cp09='" & strCP09 & "'"
         intQ = 1
         Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
         If intQ = 1 Then
            m_CP84 = Trim("" & rsQ1.Fields("CP84")) '發文規費
            m_CP27 = Trim("" & rsQ1.Fields("cp27")) '發文日期
            m_CP28 = Trim("" & rsQ1.Fields("cp28")) '發文字號
            m_CP110 = Trim("" & rsQ1.Fields("cp110")) '出名代理人
            m_CP135 = Trim("" & rsQ1.Fields("cp135")) '頁數
            m_CP136 = Trim("" & rsQ1.Fields("cp136")) '項數
         End If
      End If
      
      '取得樣本檔
      Select Case strCP10
         Case 實體審查
            m_FileName = "實體審查_樣本.doc"
            Call PUB_GetSampleFile(m_FileName, "M51-000200-0-01")
      End Select
      
      If Dir(App.path & "\" & m_FileName) <> "" Then
         Screen.MousePointer = vbHourglass
         '判斷word是否已開啟
         If g_WordAp Is Nothing Then
RestarWord:
            Set g_WordAp = New Word.Application
            g_WordAp.Visible = True 'False
         End If
'         If Dir(PUB_Getdesktop & "\" & m_TempFileName) <> "" Then
'            Kill PUB_Getdesktop & "\" & m_TempFileName
'         End If
         g_WordAp.Documents.Open App.path & "\" & m_FileName
'         g_WordAp.ActiveDocument.SaveAs PUB_Getdesktop & "\" & m_TempFileName
'         g_WordAp.ActiveDocument.Close
'         g_WordAp.Documents.Open PUB_Getdesktop & "\" & m_TempFileName
         With g_WordAp
            .Selection.WholeStory
            .Selection.Copy
            For i = 0 To 9
               strName = ""
               strText = ""
               strLineText = ""
               If i = 0 Then
                  strName = "申請案號"
                  strText = m_PA11
               ElseIf i = 1 Then
                  strName = "案號"
                  strText = m_MySt(2) & IIf(m_MySt(3) & m_MySt(4) = "000", "", "-" & m_MySt(3) & "-" & m_MySt(4))
               ElseIf i = 2 Then
                  strName = "發明名稱"
                  strText = m_PA05 & IIf(m_PA05 <> "" And m_PA06 <> "", vbCrLf, "") & m_PA06
               ElseIf i = 3 Then
                  strName = "共幾人"
                  strText = intAppCnt
               ElseIf i = 4 Then
                  strName = "申請人"
                  strText = strApplData
                  strLineText = strApplLineText
               ElseIf i = 5 Then
                  strName = "出名代理人"
                  strText = PUB_GetAgentCP110(strCP09)
               ElseIf i = 6 Then
                  strName = "頁數"
                  strText = IIf(m_CP135 = "", "　　", m_CP135)
               ElseIf i = 7 Then
                  strName = "項數"
                  strText = IIf(m_CP136 = "", "　　", m_CP136)
               ElseIf i = 8 Then
                  strName = "發文規費"
                  strText = IIf(m_CP84 = "", "　　　", Format(m_CP84, "#,##0"))
               ElseIf i = 9 Then
                  strName = "發文字號"
                  If m_CP27 = "" Then
                     strText = "發文字號： 　 年 　 月 　 日(" & Left(strSrvDate(2), 3) & ")"
                  Else
                     strText = "發文字號： " & Val(Left(m_CP27, 4)) - 1911 & " 年 " & _
                                            Mid(m_CP27, 5, 2) & " 月 " & _
                                            Right(m_CP27, 2) & " 日(" & Left(m_CP27, 4) - 1911 & ")"
                  End If
                  'strText = strText & "晉專 " & ExceptFieldData("<管制人專業代號>") & " 字第 " & m_CP28 & " 號"
                  strText = strText & "晉專外字第　　　　　　　號"
               End If
               'Find並且置換
               If Trim(strName) <> "" Then
                  .Selection.Find.ClearFormatting
                  .Selection.Find.Text = "|#" & strName & "#|"
                  .Selection.Find.Replacement.Text = ""
                  .Selection.Find.Forward = True
                  .Selection.Find.Wrap = wdFindContinue
                  .Selection.Find.Format = False
                  .Selection.Find.MatchCase = False
                  .Selection.Find.MatchWholeWord = False
                  .Selection.Find.MatchWildcards = False
                  .Selection.Find.MatchSoundsLike = False
                  .Selection.Find.MatchAllWordForms = False
                  .Selection.Find.MatchByte = True
                  .Selection.Find.Execute
                  .Selection.Delete
                  .Selection.TypeText strText
                  If strLineText <> "" Then
                     .Selection.HomeKey
                     .Selection.Find.ClearFormatting
                     With .Selection.Find
                         .Text = strLineText
                         .Replacement.Text = ""
                         .Forward = True
                         .Wrap = wdFindContinue
                         .Format = False
                         .MatchCase = False
                         .MatchWholeWord = False
                         .MatchWildcards = False
                         .MatchSoundsLike = False
                         .MatchAllWordForms = False
                         .MatchByte = True
                     End With
                     .Selection.Find.Execute
                     .Selection.Font.Underline = wdUnderlineSingle
                  End If
                  If Trim(strName) = "出名代理人" Then
                     ChgWordFormat g_WordAp.Application, strText
                  End If
               End If
ReadNext:
            Next i
         End With
         Screen.MousePointer = vbDefault
'         g_WordAp.ActiveDocument.Save
'         g_WordAp.ActiveDocument.Close
'         MsgBox "檔案已存放在：" & PUB_Getdesktop & "\" & m_TempFileName
         MsgBox "資料已產生完畢!!!"
         PUB_GetApplBook_FCP = True
      Else
         MsgBox "無申請書的樣本!!!"
      End If
   End If
   
   Set rsQ1 = Nothing 'Added by Lydia 2024/03/29
   Exit Function
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   End If
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
      Set rsQ1 = Nothing 'Added by Lydia 2024/03/29
   End If
End Function

'Add By Sindy 2015/11/26
'申請書:申請人資料
'strPer1_C,strPer1_E ~ strPer10_C,strPer10_E 代表人1~10 :C.中文 E.英文
'Add By Sindy 2018/4/19 + pa() As String
'Add By Sindy 2018/6/22 + Optional ByVal strAppType As String = "" : 空白.紙本/E.電子送件
'                       , Optional ByVal ET01 As String = "", Optional ByVal ET03 As String = "", Optional ByVal strReceiveNo As String = ""
Public Function PUB_GetApplData(pa() As String, strPA01 As String, strPA02 As String, strPA03 As String, strPA04 As String, _
                                Optional ByVal strApplPer1 As String = "", Optional ByVal strApplPer2 As String = "", _
                                Optional ByVal strApplPer3 As String = "", Optional ByVal strApplPer4 As String = "", _
                                Optional ByVal strApplPer5 As String = "", Optional ByRef strLineText As String, _
                                Optional ByRef intAppCnt As Integer, Optional ByRef intAppCnt2 As Integer, _
                                Optional ByVal strCP10 As String = "", _
Optional ByVal strPer1_C As String = "", Optional ByVal strPer1_E As String = "", _
Optional ByVal strPer2_C As String = "", Optional ByVal strPer2_E As String = "", _
Optional ByVal strPer3_C As String = "", Optional ByVal strPer3_E As String = "", _
Optional ByVal strPer4_C As String = "", Optional ByVal strPer4_E As String = "", _
Optional ByVal strPer5_C As String = "", Optional ByVal strPer5_E As String = "", _
Optional ByVal strPer6_C As String = "", Optional ByVal strPer6_E As String = "", _
Optional ByVal strPer7_C As String = "", Optional ByVal strPer7_E As String = "", _
Optional ByVal strPer8_C As String = "", Optional ByVal strPer8_E As String = "", _
Optional ByVal strPer9_C As String = "", Optional ByVal strPer9_E As String = "", _
Optional ByVal strPer10_C As String = "", Optional ByVal strPer10_E As String = "", _
Optional ByVal strAppType As String = "", Optional ByVal ET01 As String = "", _
Optional ByVal ET03 As String = "", Optional ByVal strReceiveNo As String = "") As String
Dim k As Integer
Dim strKey As String
Dim m_PA26 As String
Dim m_PA27 As String
Dim m_PA28 As String
Dim m_PA29 As String
Dim m_PA30 As String
Dim strText As String
Dim m_CU10 As String
Dim m_CU15 As String
Dim m_CU11 As String
Dim m_CU04 As String
Dim m_CU05 As String
Dim m_CU23 As String, m_CU112 As String
Dim m_CU24 As String, m_CU07 As String, m_CU103 As String
Dim m_PA79 As String, m_PA80 As String, m_PA82 As String
Dim m_PA83 As String, m_PA109 As String, m_PA110 As String
Dim m_PA112 As String, m_PA113 As String, m_PA115 As String
Dim m_PA116 As String, m_PA118 As String, m_PA119 As String
Dim m_PA121 As String, m_PA122 As String, m_PA124 As String
Dim m_PA125 As String, m_PA127 As String, m_PA128 As String
Dim m_PA130 As String, m_PA131 As String
Dim m_Person1_C As String, m_Person1_E As String, m_Person2_C As String, m_Person2_E As String
Dim strCWord1 As String, strCWord2 As String, strEWord1 As String, strEWord2 As String 'Add By Sindy 2015/12/22
Dim varTemp As Variant 'Add By Sindy 2015/12/22
Dim strPerType As String 'Add By Sindy 2016/5/4
Dim ii As Integer, strTxt(110) As String, strTmp As String 'Add By Sindy 2018/6/22
Dim m_CUX1 As String, m_CUX2 As String 'Add By Sindy 2018/6/22
Dim strChaName As String, strEngName As String 'Add By Sindy 2018/6/25
Dim bolTrans As Boolean 'Add By Sindy 2019/2/27
Dim intRow As Integer, kk As Integer 'Add By Sindy 2019/3/12
Dim strCon1 As String, intQ As Integer, rsQ1 As New ADODB.Recordset 'Added by Lydia 2024/03/29

   'Add By Sindy 2016/5/4
   'Modify By Sindy 2019/2/23 + Or strCP10 = 專利權讓與
   bolTrans = False 'Add By Sindy 2019/2/27
   If strCP10 = 讓與 Or strCP10 = 合併 Or strCP10 = 專利權讓與 Then
      bolTrans = True 'Add By Sindy 2019/2/27
      If strApplPer1 <> "" Then
         strPerType = "受讓人"
      Else
         strPerType = "讓與人"
      End If
   'Add By Sindy 2019/11/15
   ElseIf strCP10 = 授權 Or strCP10 = 終止授權 Then
      If strApplPer1 <> "" Then
         strPerType = "被授權人"
      End If
   '2019/11/15 END
   End If
   '2016/5/4 END
   If strPerType = "" Then strPerType = "申請人"
   
   'Add By Sindy 2019/1/23
   If strApplPer1 <> "" Then strApplPer1 = ChangeCustomerL(strApplPer1)
   If strApplPer2 <> "" Then strApplPer2 = ChangeCustomerL(strApplPer2)
   If strApplPer3 <> "" Then strApplPer3 = ChangeCustomerL(strApplPer3)
   If strApplPer4 <> "" Then strApplPer4 = ChangeCustomerL(strApplPer4)
   If strApplPer5 <> "" Then strApplPer5 = ChangeCustomerL(strApplPer5)
   
   '讀取基本檔資料
   strCon1 = "select * from patent where pa01=" & CNULL(strPA01) & _
               " and pa02=" & CNULL(strPA02) & _
               " and pa03=" & CNULL(strPA03) & _
               " and pa04=" & CNULL(strPA04)
   intQ = 1
   Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      m_PA26 = Trim("" & rsQ1.Fields("PA26"))
      m_PA27 = Trim("" & rsQ1.Fields("PA27"))
      m_PA28 = Trim("" & rsQ1.Fields("PA28"))
      m_PA29 = Trim("" & rsQ1.Fields("PA29"))
      m_PA30 = Trim("" & rsQ1.Fields("PA30"))
      '代表人:
      'Add By Sindy 2019/11/15
      If strPerType <> "被授權人" Then
      '2019/11/15 END
         'Add By Sindy 2016/5/4
         If strPerType = "受讓人" Then
            m_PA79 = strPer1_C
            m_PA80 = strPer1_E
            m_PA82 = strPer2_C
            m_PA83 = strPer2_E
            m_PA109 = strPer3_C
            m_PA110 = strPer3_E
            m_PA112 = strPer4_C
            m_PA113 = strPer4_E
            m_PA115 = strPer5_C
            m_PA116 = strPer5_E
            m_PA118 = strPer6_C
            m_PA119 = strPer6_E
            m_PA121 = strPer7_C
            m_PA122 = strPer7_E
            m_PA124 = strPer8_C
            m_PA125 = strPer8_E
            m_PA127 = strPer9_C
            m_PA128 = strPer9_E
            m_PA130 = strPer10_C
            m_PA131 = strPer10_E
         Else
         '2016/5/4 END
            m_PA79 = Trim("" & rsQ1.Fields("PA79"))
            m_PA80 = Trim("" & rsQ1.Fields("PA80"))
            m_PA82 = Trim("" & rsQ1.Fields("PA82"))
            m_PA83 = Trim("" & rsQ1.Fields("PA83"))
            m_PA109 = Trim("" & rsQ1.Fields("PA109"))
            m_PA110 = Trim("" & rsQ1.Fields("PA110"))
            m_PA112 = Trim("" & rsQ1.Fields("PA112"))
            m_PA113 = Trim("" & rsQ1.Fields("PA113"))
            m_PA115 = Trim("" & rsQ1.Fields("PA115"))
            m_PA116 = Trim("" & rsQ1.Fields("PA116"))
            m_PA118 = Trim("" & rsQ1.Fields("PA118"))
            m_PA119 = Trim("" & rsQ1.Fields("PA119"))
            m_PA121 = Trim("" & rsQ1.Fields("PA121"))
            m_PA122 = Trim("" & rsQ1.Fields("PA122"))
            m_PA124 = Trim("" & rsQ1.Fields("PA124"))
            m_PA125 = Trim("" & rsQ1.Fields("PA125"))
            m_PA127 = Trim("" & rsQ1.Fields("PA127"))
            m_PA128 = Trim("" & rsQ1.Fields("PA128"))
            m_PA130 = Trim("" & rsQ1.Fields("PA130"))
            m_PA131 = Trim("" & rsQ1.Fields("PA131"))
         End If
      End If
      '有輸入申請人時,以輸入資料為主,不然才抓基本檔申請人資料
      If Trim(strApplPer1) = "" Then
         strApplPer1 = m_PA26
         strApplPer2 = m_PA27
         strApplPer3 = m_PA28
         strApplPer4 = m_PA29
         strApplPer5 = m_PA30
      End If
      '共幾人
      If strApplPer1 <> "" Then intAppCnt = intAppCnt + 1
      If strApplPer2 <> "" Then intAppCnt = intAppCnt + 1
      If strApplPer3 <> "" Then intAppCnt = intAppCnt + 1
      If strApplPer4 <> "" Then intAppCnt = intAppCnt + 1
      If strApplPer5 <> "" Then intAppCnt = intAppCnt + 1
      '原申請人共幾人
      If m_PA26 <> "" Then intAppCnt2 = intAppCnt2 + 1
      If m_PA27 <> "" Then intAppCnt2 = intAppCnt2 + 1
      If m_PA28 <> "" Then intAppCnt2 = intAppCnt2 + 1
      If m_PA29 <> "" Then intAppCnt2 = intAppCnt2 + 1
      If m_PA30 <> "" Then intAppCnt2 = intAppCnt2 + 1
   End If
   
   'Add By Sindy 2019/11/15
   If strPerType = "被授權人" And Left(strApplPer1, 1) <> "X" Then
      ii = ii + 1
      strTmp = strPerType & "1-國籍"
      strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','　　')"
      If strPer1_C <> "" Then
         ii = ii + 1
         strTmp = strPerType & "1-中文名稱"
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & strPer1_C & "')"
      End If
      If Trim(strPer1_E) <> "" Then
         ii = ii + 1
         strTmp = strPerType & "1-英文名稱"
         strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
            " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & strPer1_E & "')"
      End If
   Else
   '2019/11/15 END
      For k = 1 To 5
         strKey = ""
         m_CU10 = "": m_CU15 = "": m_CU11 = "": m_CU04 = "": m_CU05 = ""
         m_CU23 = "": m_CU24 = "": m_CU112 = ""
         m_Person1_C = "": m_Person1_E = "": m_Person2_C = "": m_Person2_E = ""
         If k = 1 And strApplPer1 <> "" Then
            strKey = strApplPer1
            m_Person1_C = m_PA79
            m_Person1_E = m_PA80
           m_Person2_C = m_PA82
            m_Person2_E = m_PA83
         ElseIf k = 2 And strApplPer2 <> "" Then
            strKey = strApplPer2
            m_Person1_C = m_PA109
            m_Person1_E = m_PA110
            m_Person2_C = m_PA112
            m_Person2_E = m_PA113
         ElseIf k = 3 And strApplPer3 <> "" Then
            strKey = strApplPer3
            m_Person1_C = m_PA115
            m_Person1_E = m_PA116
            m_Person2_C = m_PA118
            m_Person2_E = m_PA119
         ElseIf k = 4 And strApplPer4 <> "" Then
            strKey = strApplPer4
            m_Person1_C = m_PA121
            m_Person1_E = m_PA122
            m_Person2_C = m_PA124
            m_Person2_E = m_PA125
         ElseIf k = 5 And strApplPer5 <> "" Then
            strKey = strApplPer5
            m_Person1_C = m_PA127
            m_Person1_E = m_PA128
            m_Person2_C = m_PA130
            m_Person2_E = m_PA131
         End If
        If strKey <> "" Then
   '         strcon1 = "select cu10,cu15,cu11,cu04,decode(cu05,null,'',nvl(cu05,'')||' '||nvl(cu88,'')||' '||nvl(cu89,'')||' '||nvl(cu90,'')) as cu05,cu16,cu17,cu18,cu19," & _
   '                     "cu07,cu103,cu23,cu39,cu40,decode(cu24,null,'',nvl(cu24,'')||' '||nvl(cu25,'')||' '||nvl(cu26,'')||' '||nvl(cu27,'')||' '||nvl(cu28,'')) as cu24,cu112," & _
   '                     "N1.NA72 X1,N2.NA72 X2" & _
   '                     " from customer,NATION N1,NATION N2 where cu01='" & Left(strKey, 8) & "'" & _
   '                     " and cu02='" & Mid(strKey, 9) & "' AND N1.NA01(+)=CU10 AND N2.NA01(+)=CU87"
            'Modify By Sindy 2019/3/12 + 代表人
            strCon1 = "select cu10,cu15,cu11,cu04,decode(cu05,null,'',nvl(cu05,'')||' '||nvl(cu88,'')||' '||nvl(cu89,'')||' '||nvl(cu90,'')) as cu05,cu16,cu17,cu18,cu19" & _
                        ",cu07,cu103,cu23" & _
                        ",cu39,cu40,cu41,cu42,cu43,cu44,cu45,cu46,cu47,cu48,cu49,cu50" & _
                        ",cu51,cu52,cu53,cu54,cu55,cu56" & _
                        ",decode(cu24,null,'',nvl(cu24,'')||' '||nvl(cu25,'')||' '||nvl(cu26,'')||' '||nvl(cu27,'')||' '||nvl(cu28,'')||' '||nvl(cu102,'')) as cu24,cu112" & _
                        ",N1.NA72 X1,N2.NA72 X2" & _
                        " from customer,NATION N1,NATION N2 where cu01='" & Left(strKey, 8) & "'" & _
                        " and cu02='" & Mid(strKey, 9) & "' AND N1.NA01(+)=CU10 AND N2.NA01(+)=CU87"
            intQ = 1
            Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
            If intQ = 1 Then
               m_CU10 = Trim("" & rsQ1.Fields("cu10"))
               m_CU15 = Trim("" & rsQ1.Fields("cu15"))
               m_CU11 = Trim("" & rsQ1.Fields("cu11"))
               m_CU04 = ChgSQL(Trim("" & rsQ1.Fields("cu04")))
               m_CU05 = ChgSQL(Trim("" & rsQ1.Fields("cu05")))
               m_CU23 = ChgSQL(Trim("" & rsQ1.Fields("cu23")))
               m_CU24 = ChgSQL(Trim("" & rsQ1.Fields("cu24")))
               m_CU112 = Trim("" & rsQ1.Fields("cu112"))
               m_CUX1 = Trim("" & rsQ1.Fields("X1"))
               m_CUX2 = Trim("" & rsQ1.Fields("X2"))
               m_CU07 = Trim("" & rsQ1.Fields("cu07")) 'Add By Sindy 2019/4/12 公司負責人
               m_CU103 = Trim("" & rsQ1.Fields("cu103")) 'Add By Sindy 2019/4/12 公司英文負責人
            End If
            
            'Add By Sindy 2018/6/22
            If UCase(strAppType) = "E" Then 'E.電子送件
               ii = ii + 1
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-國籍','" & m_CUX1 & "')"
               ii = ii + 1
               If m_CU15 = "0" Then
                  strTmp = "自然人"
               Else
                  strTmp = "法人公司機關學校"
               End If
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-身分種類','" & strTmp & "')"
                     
               ii = ii + 1
               If m_CU10 < "010" Then
                  If m_CU15 = "0" And "" & m_CU11 = "" Then '個人無ID時也要顯示標題
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-ID','♀')"
                  Else
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-ID','" & m_CU11 & "')"
                  End If
               End If
                  
               ii = ii + 1
               If m_CU15 = "0" Then
                  strTmp = strPerType & k & "-中文姓名"
               Else
                  strTmp = strPerType & k & "-中文名稱"
               End If
               '修法:106/12/01開始中文名稱要加外商國名
               If Val(strSrvDate(2)) >= 1061201 And m_CU15 = "1" Then '1.公司
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & GetPrjNationName(m_CU10, "NA81", pa(1)) & m_CU04 & "')"
               Else
                  '柏翰提個人的姓和名中間要有,號
                  If m_CU15 = "0" Then '自然人
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & PUB_ConvertNameFormat(m_CU04) & "')"
                  Else
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & m_CU04 & "')"
                  End If
               End If
               
               ii = ii + 1
               If m_CU15 = "0" Then
                  strTmp = strPerType & k & "-英文姓名"
               Else
                  strTmp = strPerType & k & "-英文名稱"
               End If
               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & m_CU05 & "')"
               
               '目前抓客戶基本檔資料,等基本檔加欄位後需改抓
               'Modify By Sindy 2019/5/20
               If Left(Pub_StrUserSt03, 1) = "F" Then '外專抓取地址國籍
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-居住國','" & m_CUX2 & "')"
               Else
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-居住國','" & m_CUX1 & "')"
               End If
               
               '讓與時應該抓客戶資料
               'Modify By Sindy 2018/7/10 ex:FCP-047260:701
               'If strCP10 = 讓與 Then
               'If strCP10 = 讓與 And strPerType = "受讓人" Then
               'Modify By Sindy 2019/2/27 + bolTrans = True
               If bolTrans = True And strPerType = "受讓人" Then
               '2018/7/10 END
                  If m_CU10 < "010" Then
                     ii = ii + 1
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-郵遞區號','" & PUB_ChgNumeralStyle(m_CU112) & "')"
                  End If
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-中文地址','" & PUB_ChgNumeralStyle(m_CU23) & "')"
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-英文地址','" & m_CU24 & "')"
               Else
                  'Add By Sindy 2019/2/27
                  If m_CU10 < "010" And Trim(PUB_ChgNumeralStyle(ChgSQL(m_CU23))) = Trim(PUB_ChgNumeralStyle(ChgSQL(pa(30 + k)))) Then
                     ii = ii + 1
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-郵遞區號','" & PUB_ChgNumeralStyle(m_CU112) & "')"
                  'Add By Sindy 2019/8/14
                  ElseIf m_CU10 < "010" Then
                     ii = ii + 1
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-郵遞區號','♀')"
                  '2019/8/14 END
                  End If
                  '2019/2/27 END
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-中文地址','" & PUB_ChgNumeralStyle(ChgSQL(pa(30 + k))) & "')"
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-英文地址','" & ChgSQL(pa(35 + k)) & "')"
               End If
               
               strChaName = "": strEngName = ""
               If m_CU15 <> "0" Then '非自然人才要帶出代表人資料
                  If Left(Pub_StrUserSt03, 1) = "F" Then '國外部
                     strChaName = strChaName & " " & IIf(m_Person2_C <> "", "1.", "") & m_Person1_C
                     If m_Person2_C <> "" Then
                        strChaName = strChaName & " 2." & m_Person2_C
                     End If
                     strChaName = Trim(strChaName)
                     strEngName = strEngName & " " & IIf(m_Person2_E <> "", "1.", "") & m_Person1_E
                     If m_Person2_E <> "" Then
                        strEngName = strEngName & " 2." & m_Person2_E
                     End If
                     strEngName = Trim(strEngName)
                  'Modify By Sindy 2019/3/12 代表人:P案要抓客戶檔
                  Else
                     'Modify By Sindy 2019/5/14
                     'Modify By Sindy 2019/4/12
                     '公司負責人
                     If m_CU07 <> "" Then
                        If Len(m_CU07) = 3 Then
                           strChaName = PUB_ConvertNameFormat(m_CU07)
                        Else
                           strChaName = m_CU07
                        End If
                     End If
                     '2019/5/14 END
                     '公司英文負責人
                     strEngName = m_CU103
                     
                     '2019/4/12 END
   '                  intRow = 0
   '                  For kk = 1 To 6
   '                     intRow = intRow + 1
   '                     '代表人中文姓名-->非自然人時為必要欄位
   '                     strTmp = "" & rsq1("CU" & CStr(39 + 3 * (kk - 1)))
   '                     If strTmp <> "" Then
   '                        If Len(strTmp) = 3 Then strTmp = PUB_ConvertNameFormat(strTmp)
   '                        strChaName = strChaName & " " & intRow & "." & strTmp
   '                     Else
   '                        'Modify By Sindy 2018/1/17 只有一個代表人時不要有1.
   '                        If strChaName <> "" Then
   '                           strChaName = Replace(strChaName, "1.", "")
   '                        End If
   '                        '2018/1/17 END
   '                     End If
   '                     '代表人英文姓名-->非必要欄位
   '                     strTmp = "" & rsq1("CU" & CStr(40 + 3 * (kk - 1)))
   '                     If strTmp <> "" Then
   '                        strEngName = strEngName & " " & intRow & "." & strTmp
   '                     Else
   '                        'Modify By Sindy 2018/1/17 只有一個代表人時不要有1.
   '                        If strEngName <> "" Then
   '                           strEngName = Replace(strEngName, "1.", "")
   '                        End If
   '                        '2018/1/17 END
   '                     End If
   '                  Next kk
                  '2019/3/12 END
                  End If
                  '代表人中文姓名
                  If strChaName = "" Then
                     If Left(Pub_StrUserSt03, 1) = "F" Then '外專要帶後補2個字
                        strChaName = "後補"
                     Else
                        strChaName = "♀"
                     End If
                  'Modify By Sindy 2018/7/10 ex:FCP-047260:701
   '               Else
   '                  strChaName = Mid(strChaName, 2)
                  End If
                  ii = ii + 1
                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-代表人中文姓名','" & ChgSQL(strChaName) & "')"
                  '代表人英文姓名
                 'Add By Sindy 2019/8/29 敏莉說中文有後補,英文無資料才要帶後補
                  If Trim(strEngName) = "" And Left(Pub_StrUserSt03, 1) = "F" And strChaName = "後補" Then '外專要帶後補2個字
                     strEngName = "後補"
                  End If
                  '2019/8/29 END
                  If strEngName <> "" Then
                     ii = ii + 1
                     strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & k & "-代表人英文姓名','" & ChgSQL(strEngName) & "')"
                  End If
   '               strChaName = "": strEngName = ""
   '               '讓與時應該抓客戶資料
   '               If strCP10 = 讓與 Then
   '                  intRow = 0
   '                  For kk = 1 To 6
   '                     intRow = intRow + 1
   '                     '代表人中文姓名-->非自然人時為必要欄位
   '                     strTmp = "" & rsq1("CU" & CStr(39 + 3 * (kk - 1)))
   '                     If strTmp <> "" Then
   '                        If Len(strTmp) = 3 Then strTmp = PUB_ConvertNameFormat(strTmp)
   '                        strChaName = strChaName & " " & intRow & "." & strTmp
   '                     Else
   '                        '只有一個代表人時不要有1.
   '                        If strChaName <> "" Then
   '                           strChaName = Replace(strChaName, "1.", "")
   '                        End If
   '                     End If
   '                     '代表人英文姓名-->非必要欄位
   '                     strTmp = "" & rsq1("CU" & CStr(40 + 3 * (kk - 1)))
   '                     If strTmp <> "" Then
   '                        strEngName = strEngName & " " & intRow & "." & strTmp
   '                     Else
   '                        '只有一個代表人時不要有1.
   '                        If strEngName <> "" Then
   '                           strEngName = Replace(strEngName, "1.", "")
   '                        End If
   '                     End If
   '                  Next kk
   '               Else
   '                  If jj < 3 Then
   '                     k_star = 1: k_end = 2
   '                  ElseIf jj = 3 Then
   '                     k_star = 3: k_end = 4
   '                  ElseIf jj = 4 Then
   '                     k_star = 5: k_end = 6
   '                  ElseIf jj = 5 Then
   '                     k_star = 7: k_end = 8
   '                  End If
   '                  intRow = 0
   '                  For kk = k_star To k_end
   '                     intRow = intRow + 1
   '                     '代表人中文姓名-->非自然人時為必要欄位
   '                     If jj = 1 Then
   '                        strTmp = pa(79 + 3 * (kk - 1))
   '                     Else
   '                        strTmp = pa(109 + 3 * (kk - 1))
   '                     End If
   '                     If strTmp <> "" Then
   '                        strChaName = strChaName & " " & intRow & "." & strTmp
   '                     Else
   '                        '只有一個代表人時不要有1.
   '                        If strChaName <> "" Then
   '                           strChaName = Replace(strChaName, "1.", "")
   '                        End If
   '                     End If
   '                     '代表人英文姓名-->非必要欄位
   '                     If jj = 1 Then
   '                        strTmp = pa(80 + 3 * (kk - 1))
   '                     Else
   '                        strTmp = pa(110 + 3 * (kk - 1))
   '                     End If
   '                     If strTmp <> "" Then
   '                        strEngName = strEngName & " " & intRow & "." & strTmp
   '                     Else
   '                        '只有一個代表人時不要有1.
   '                        If strEngName <> "" Then
   '                           strEngName = Replace(strEngName, "1.", "")
   '                        End If
   '                     End If
   '                  Next kk
   '               End If
   '               '代表人中文姓名
   '               If strChaName = "" Then
   '                  If Left(Pub_StrUserSt03, 2) = "F2" Then '外專要帶後補2個字
   '                     strChaName = "後補"
   '                  Else
   '                     strChaName = "♀"
   '                  End If
   '               Else
   '                  strChaName = Mid(strChaName, 2)
   '               End If
   '               ii = ii + 1
   '               strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-代表人中文姓名','" & ChgSQL(strChaName) & "')"
   '               '代表人英文姓名
   '               If strEngName <> "" Then
   '                  strEngName = Mid(strEngName, 2)
   '                  ii = ii + 1
   '                  strTxt(ii) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
   '                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','申請人" & jj & "-代表人英文姓名','" & ChgSQL(strEngName) & "')"
   '               End If
               End If
            Else
            '2018/6/22 END
               
               'Modify By Sindy 2016/3/9
               'strText = strText & IIf(k > 1, vbCrLf, "") & "（第" & k & "申請人）■為專利申請人　□非專利申請人" & vbCrLf
               If strCP10 = 實體審查 Then
                  strText = strText & IIf(k > 1, vbCrLf, "") & "（第" & k & "申請人）■為專利申請人　□非專利申請人" & vbCrLf
               Else
                  'Modify By Sindy 2016/5/4
                  If strPerType <> "" Then
                     strText = strText & IIf(k > 1, vbCrLf, "") & "（第" & k & strPerType & "）" & vbCrLf
                  Else
                  '2016/5/4 END
                     strText = strText & IIf(k > 1, vbCrLf, "") & "（第" & k & "申請人）" & vbCrLf
                  End If
               End If
               '2016/3/9 END
               If m_CU10 = "" Then
                  strText = strText & "國　　籍：□中華民國 □大陸地區（□大陸、□香港、□澳門）" & vbCrLf
                  strText = strText & "　　　　　□外國籍：" & vbCrLf
               Else
                  If m_CU10 <= "010" Then
                     strText = strText & "國　　籍：■中華民國 □大陸地區（□大陸、□香港、□澳門）" & vbCrLf
                     strText = strText & "　　　　　□外國籍：" & vbCrLf
                  ElseIf m_CU10 = "020" Then
                     strText = strText & "國　　籍：□中華民國 ■大陸地區（■大陸、□香港、□澳門）" & vbCrLf
                     strText = strText & "　　　　　□外國籍：" & vbCrLf
                  ElseIf m_CU10 = "013" Then
                     strText = strText & "國　　籍：□中華民國 ■大陸地區（□大陸、■香港、□澳門）" & vbCrLf
                     strText = strText & "　　　　　□外國籍：" & vbCrLf
                  ElseIf m_CU10 = "044" Then
                     strText = strText & "國　　籍：□中華民國 ■大陸地區（□大陸、□香港、■澳門）" & vbCrLf
                     strText = strText & "　　　　　□外國籍：" & vbCrLf
                  Else
                     strText = strText & "國　　籍：□中華民國 □大陸地區（□大陸、□香港、□澳門）" & vbCrLf
                     strText = strText & "　　　　　■外國籍：" & GetPrjNationName(Left(m_CU10, 3)) & vbCrLf
                     strLineText = GetPrjNationName(Left(m_CU10, 3))
                  End If
               End If
             If m_CU15 = "0" Then
                  strText = strText & "身分種類：■自然人               □法人、公司、機關、學校" & vbCrLf
               Else
                  strText = strText & "身分種類：□自然人               ■法人、公司、機關、學校" & vbCrLf
               End If
               strText = strText & "ID：　" & m_CU11 & vbCrLf
               'Modify By Sindy 2015/12/22
      '         strText = strText & "姓　名：　　姓：　　　　　　名：" & vbCrLf
      '         strText = strText & "　　　　　　Family　　　　　Given" & vbCrLf
      '         strText = strText & "　　　　　　name：　　　　　name：" & vbCrLf
               'Modify By Sindy 2016/2/3
               If m_CU15 = "0" Then '個人
                  '中文名稱
                  strCWord1 = "姓：": strCWord2 = "名："
                  If m_CU04 <> "" Then
                     varTemp = Split(m_CU04, ",")
                    If UBound(varTemp) = 0 Then
                        strCWord1 = strCWord1 & m_CU04
                     Else
                        strCWord1 = strCWord1 & Trim(varTemp(0))
                        strCWord2 = strCWord2 & Trim(varTemp(1))
                     End If
                     If LenB(Trim(strCWord1)) > 23 Then
                        strCWord1 = strCWord1 & " "
                     Else
                        strCWord1 = convForm(CheckStr(Trim(strCWord1)), 23)
                     End If
                     strText = strText & "姓　名：　　" & strCWord1 & strCWord2 & vbCrLf
                  Else
                     strText = strText & "姓　名：　　姓：　　　　　　       名：" & vbCrLf
                  End If
                  '英文名稱
                  strEWord1 = "Family name：": strEWord2 = "Given name："
                  If m_CU05 <> "" Then
                     varTemp = Split(m_CU05, ",")
                     If UBound(varTemp) = 0 Then
                        strEWord1 = strEWord1 & m_CU05
                     Else
                        strEWord1 = strEWord1 & Trim(varTemp(0))
                        strEWord2 = strEWord2 & Trim(varTemp(1))
                     End If
                     If Len(Trim(strEWord1)) > 21 Then
                        strEWord1 = strEWord1 & " "
                     Else
                       strEWord1 = convForm(CheckStr(Trim(strEWord1)), 21)
                     End If
                     strText = strText & "　　　　　　" & strEWord1 & strEWord2 & vbCrLf
                  Else
                     strText = strText & "　　　　　　Family name：　　　　　Given name：" & vbCrLf
                  End If
               Else
               '2016/2/3 END
               '2015/12/22 END
                  'Add By Sindy 2017/11/14 修法:106/12/01開始中文名稱要加外商國名
                  If Val(strSrvDate(2)) >= 1061201 And m_CU15 = "1" Then   '1.公司
                     strText = strText & "名　稱（中文）：" & GetPrjNationName(m_CU10, "NA81", pa(1)) & m_CU04 & vbCrLf
                  Else
                  '2017/11/14 END
                     strText = strText & "名　稱（中文）：" & m_CU04 & vbCrLf
                  End If
                  strText = strText & "　　　（英文）：" & m_CU05 & vbCrLf
               End If
               If m_Person1_C = "" And m_Person1_E = "" And m_Person2_C = "" And m_Person2_E = "" Then
                  strText = strText & "代表人（中文）：" & vbCrLf
                  strText = strText & "　　　（英文）：" & vbCrLf
               Else
                  strText = strText & "代表人（中文）：" & m_Person1_C & vbCrLf
                  strText = strText & "　　　（英文）：" & m_Person1_E & vbCrLf
                  If m_Person2_C <> "" Or m_Person2_E <> "" Then
                     strText = strText & "　　　（中文）：" & m_Person2_C & vbCrLf
                     strText = strText & "　　　（英文）：" & m_Person2_E & vbCrLf
                  End If
               End If
               'Add By Sindy 2018/4/19 地址改抓個案地址
      '         strText = strText & "地　址（中文）：" & PUB_ChgNumeralStyle(IIf(m_CU112 <> "", m_CU112 & " ", "") & m_CU23) & vbCrLf
      '         strText = strText & "　　　（英文）：" & m_CU24 & vbCrLf
               'Modify By Sindy 2018/5/3 讓與時應該抓客戶資料
               'Modify By Sindy 2018/7/10 ex:FCP-047260:701
               'If strCP10 = 讓與 Then
               'If strCP10 = 讓與 And strPerType = "受讓人" Then
               'Modify By Sindy 2019/2/27 + bolTrans = True
               If bolTrans = True And strPerType = "受讓人" Then
               '2018/7/10 END
                  strText = strText & "地　址（中文）：" & PUB_ChgNumeralStyle(IIf(m_CU112 <> "", m_CU112 & " ", "") & m_CU23) & vbCrLf
                  strText = strText & "　　　（英文）：" & m_CU24 & vbCrLf
               Else
               '2018/5/3 END
                  strText = strText & "地　址（中文）：" & PUB_ChgNumeralStyle(pa(30 + k)) & vbCrLf
                  strText = strText & "　　　（英文）：" & pa(35 + k) & vbCrLf
               End If
               '2018/4/19 END
               'Modify By Sindy 2016/5/4 + Or strPerType = "受讓人"
               If strCP10 = 發明申請 Or strCP10 = 新型申請 Or strCP10 = 設計申請 Or strPerType = "受讓人" Then
                  strText = strText & "□註記此申請人為應受送達人" & vbCrLf
               End If
               'Modify By Sindy 2016/5/4
               'strText = strText & "聯絡電話及分機：" & vbCrLf
               strText = strText & "聯絡電話及分機："
               If strCP10 <> 讓與 Then
                  strText = strText & vbCrLf
               '2016/5/4 END
                  strText = strText & "傳  真："
               End If
            End If
         End If
      Next k
   End If
   
   Set rsQ1 = Nothing 'Added by Lydia 2024/03/29
   
   'Add By Sindy 2018/6/22
   If UCase(strAppType) = "E" Then 'E.電子送件
      If Not ClsLawExecSQL(ii, strTxt) Then
         MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
      End If
      PUB_GetApplData = ""
   Else
   '2018/6/22 END
      PUB_GetApplData = strText
   End If
End Function

'Add By Sindy 2015/11/26
'申請書:發明人資料
Public Function PUB_GetApplInventor(strPA01 As String, strPA02 As String, strPA03 As String, _
                                    strPA04 As String, strTypeName As String, _
                                    Optional ByRef intInvCnt As Integer) As String
Dim ii As Integer
Dim rsA As New ADODB.Recordset
Dim varTemp As Variant
Dim strCWord1 As String, strCWord2 As String
Dim strEWord1 As String, strEWord2 As String
Dim intQ As Integer, strCon1 As String, strCon2 As String 'Added by Lydia 2024/03/29

   PUB_GetApplInventor = ""
   intInvCnt = 0
   strCon1 = "select in03,in04,in05,na03 from PatentInventor,Inventor,nation" & _
               " where pi01='" & strPA01 & "' and pi02='" & strPA02 & "' and pi03='" & strPA03 & "' and pi04='" & strPA04 & "'" & _
               " and substr(pi06,1,8)=in01(+) and substr(pi06,9)=in02(+)" & _
               " and in11=na01(+)" & _
               " order by pi05 asc"
   intQ = 1
   Set rsA = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      rsA.MoveFirst
      ii = 0
      Do While rsA.EOF = False
         ii = ii + 1
         '中文名稱
         strCWord1 = "　　　　　　　　　    " 'Add By Sindy 2016/8/10
         strCWord2 = "" 'Add By Sindy 2016/8/10
         If Trim("" & rsA.Fields("in04")) <> "" Then
            varTemp = Split(rsA.Fields("in04"), ",")
            If LenB(Trim(varTemp(0))) > 24 Then
               strCWord1 = Trim(varTemp(0)) & " "
            Else
               strCWord1 = convForm(CheckStr(Trim(varTemp(0))), 24)
            End If
            If UBound(varTemp) > 0 Then
               strCWord2 = Trim(varTemp(1))
            End If
         End If
         '英文名稱
         strEWord1 = "               " 'Add By Sindy 2016/8/10
         strEWord2 = "" 'Add By Sindy 2016/8/10
         If Trim("" & rsA.Fields("in05")) <> "" Then
            'Modify By Sindy 2018/10/25 增加英文名稱格式化 PUB_FCPIN05Format_EName
            Call PUB_FCPIN05Format_EName(rsA.Fields("in05"), "" & rsA.Fields("na03"), strEWord1, strEWord2)
            
'            varTemp = Split(rsA.Fields("in05"), ",")
'            If LenB(Trim(varTemp(0))) > 12 Then
'               'Modify By Sindy 2016/10/18 姓:全部大寫,日本除外
'               If rsA.Fields("na03") = "日本" Then
'                  strEWord1 = Trim(varTemp(0)) & " "
'               Else
'                  If UBound(varTemp) > 0 Then
'                     strEWord1 = UCase(Trim(varTemp(0))) & " "
'                  Else
'                     strEWord1 = Trim(varTemp(0)) & " "
'                  End If
'               End If
'               '2016/10/18 END
'            Else
'               'Modify By Sindy 2016/10/18 姓:全部大寫,日本除外
'               If rsA.Fields("na03") = "日本" Then
'                  strEWord1 = convForm(CheckStr(Trim(varTemp(0))), 12)
'               Else
'                  If UBound(varTemp) > 0 Then
'                     strEWord1 = convForm(CheckStr(UCase(Trim(varTemp(0)))), 12)
'                  Else
'                     strEWord1 = convForm(CheckStr(Trim(varTemp(0))), 12)
'                  End If
'               End If
'               '2016/10/18 END
'            End If
'            If UBound(varTemp) > 0 Then
'               'Modify By Sindy 2016/10/18 名:第一個字大寫其他全部小寫,日本除外
'               If rsA.Fields("na03") = "日本" Then
'                  strEWord2 = Trim(varTemp(1))
'               Else
'                  strEWord2 = StrConv(Trim(varTemp(1)), vbProperCase)
'                  'Modify By Sindy 2017/8/16 "-"後的第一個英文字也要大寫
'                  If InStr(strEWord2, "-") > 0 Then
'                     strEWord2 = Left(strEWord2, InStr(strEWord2, "-") - 1) & "-" & StrConv(Mid(strEWord2, InStr(strEWord2, "-") + 1), vbProperCase)
'                  End If
'                  '2017/8/16 END
'               End If
'               '2016/10/18 END
'            End If
         End If
         strCon2 = "" & rsA.Fields("in03")
         If strCon2 <> "" Then
            strCon2 = convForm(CheckStr(strCon2), 30)
         Else
            strCon2 = convForm(CheckStr(strCon2), 32)
         End If
         'Modify By Sindy 2016/3/11 國籍=台灣時顯示為中華民國
         PUB_GetApplInventor = PUB_GetApplInventor & IIf(ii > 1, vbCrLf, "") & _
                             "（第" & ii & strTypeName & "）" & vbCrLf & _
                             "ID：" & strCon2 & "國籍：" & IIf(rsA.Fields("na03") = "台灣", "中華民國", rsA.Fields("na03")) & vbCrLf & _
                             "姓名：姓：" & strCWord1 & "名：" & strCWord2 & vbCrLf & _
                             "　　　Family name：" & strEWord1 & "Given name：" & strEWord2
         rsA.MoveNext
      Loop
      intInvCnt = rsA.RecordCount
   Else
      PUB_GetApplInventor = "（第   " & strTypeName & "）" & vbCrLf & _
                            "ID：　　　　　　　　　　　　　　國籍：" & vbCrLf & _
                            "姓名：姓：　　　　　　　　　　　名：" & vbCrLf & _
                            "　　　Family name：　　　　　　  Given name："
   End If
   rsA.Close
   Set rsA = Nothing
End Function

'申請書:出名代理人資料
'Modify By Sindy 2016/4/29 + Optional strCP110 As String, Optional strCP01 As String = ""
'Modified by Lydia 2019/03/27 + strType 回傳資料方式(1.FCP紙本送件,2/3.T台灣案+FCT紙本送件,4.T台灣案+FCT電子送件,5.代理人簽章)
Public Function PUB_GetAgentCP110(strCP09 As String, Optional strCP110 As String = "", _
                                  Optional strCP01 As String = "", Optional strType As String = "1") As String
Dim ii As Integer
Dim rsA As New ADODB.Recordset
Dim strA1 As String, strA2 As String 'Added by Lydia 2018/03/26
Dim strST02 As String 'Added by Lydia 2019/03/29
Dim intQ As Integer, strCon1 As String 'Added by Lydia 2024/03/29
Dim strA3 As String 'Added by Lydia 2024/05/20

   PUB_GetAgentCP110 = ""
   '|#(小字)#|
   If strCP09 = "" Then
      'Added by Lydia 2019/03/27 ＋判斷
      Select Case strType
           Case "1" '(原)FCP紙本送件
                PUB_GetAgentCP110 = "|#(18字)◎代理人：#||#(小字)(多位代理人時，應將本欄位完整複製後依序填寫)#|" & vbCrLf & _
                                    "ID：" & vbCrLf & _
                                    "姓名：　姓：　　　名：" & vbCrLf & _
                                    "證書字號：" & vbCrLf & _
                                    "事務所名稱：" & vbCrLf & _
                                    "地址：" & vbCrLf & _
                                    "聯絡電話及分機："
           Case "2", "3" 'T台灣案+FCT紙本送件
                'Modified by Lydia 2024/05/20 +登錄字號
                PUB_GetAgentCP110 = "ID：" & vbCrLf & _
                                    "姓    名：" & vbCrLf & _
                                    "登錄字號：" & vbCrLf & _
                                    "地  　址：" & vbCrLf & _
                                    "聯絡電話及分機：(02)2506-1023 分機" & vbCrLf & _
                                    "傳    真："
           Case "4" 'T台灣案+FCT電子送件
                PUB_GetAgentCP110 = ""
      End Select
      'end 2019/03/27
   Else
      'Add By Sindy 2016/4/29 直接傳入出名代理人代碼
      If strCP110 <> "" Then
        'Modified by Lydia 2019/03/27 +OA08(電子送件)證書號
         strCon1 = "SELECT s1.ST02,s1.ST26,OA05,NVL(OA06,'專利代理人') OA06,OA08" & _
                     " FROM STAFF s1,OURAGENT" & _
                     " WHERE INSTR('" & strCP110 & "',s1.ST01)>0 AND OA01='" & strCP01 & "' AND OA02=s1.ST01" & _
                     " order by OA03"
      Else
      '2016/4/29 END
         'Modified by Lydia 2019/03/27 +OA08(電子送件)證書號
         strCon1 = "SELECT s1.ST02,s1.ST26,OA05,NVL(OA06,'專利代理人') OA06,OA08" & _
                     " FROM CASEPROGRESS,STAFF s1,OURAGENT" & _
                     " WHERE INSTR(CP110,s1.ST01)>0 AND OA01=CP01 AND OA02=s1.ST01 AND CP09='" & strCP09 & "'" & _
                     " order by OA03"
      End If
      intQ = 1
      Set rsA = ClsLawReadRstMsg(intQ, strCon1)
      If intQ = 1 Then
         rsA.MoveFirst
         ii = 0
         Do While rsA.EOF = False
            ii = ii + 1
            strST02 = PUB_Big5toUnicode("" & rsA.Fields("st02")) 'Added by Lydia 2019/03/29 更換造字為Unicode字
            'Added by Lydia 2019/03/27 ＋判斷
            Select Case strType
                 Case "1" '(原)FCP紙本送件
                    'Modify By Sindy 2018/4/10 出名代理人的ID不要顯示資料 "ID：" & rsA.Fields(1) & vbCrLf ==> "ID：" & vbCrLf
                    'Modify By Sindy 2020/3/26 "事務所名稱：台一國際專利法律事務所" => CompNameQuery("2")
                    PUB_GetAgentCP110 = PUB_GetAgentCP110 & IIf(ii > 1, vbCrLf, "") & _
                                        "|#(18字)◎代理人" & IIf(rsA.RecordCount > 1, ii, "") & "：#||#(小字)(多位代理人時，應將本欄位完整複製後依序填寫)#|" & vbCrLf & _
                                        "ID：" & vbCrLf & _
                                        "姓名：　姓：" & Left(rsA.Fields(0), 1) & "　　名：" & Mid(rsA.Fields(0), 2) & vbCrLf & _
                                        "證書字號：" & rsA.Fields(2) & vbCrLf & _
                                        "事務所名稱：" & CompNameQuery("2") & vbCrLf & _
                                        "地址：臺北市長安東路二段112號9樓" & vbCrLf & _
                                        "聯絡電話及分機：(02)2506-1023/分機" & ExceptFieldData("FCP管制人分機")

                 Case "2" 'T台灣案+FCT紙本送件: 同一行顯示
                      strA1 = strA1 & "、" & rsA.Fields("ST26") 'ID
                      strA2 = strA2 & "、" & strST02
                      'Added by Lydia 2024/05/20 +登錄字號
                      strA3 = strA3 & "、" & rsA.Fields("OA05")
                 Case "3"  'T台灣案+FCT紙本送件(註冊前變更)=>多位代理人時，應將本欄位完整複製後依序填寫
                      'Modified by Lydia 2024/05/20 +登錄字號
                      PUB_GetAgentCP110 = PUB_GetAgentCP110 & IIf(ii > 1, vbCrLf & vbCrLf, "") & _
                                           "ID：" & rsA.Fields("ST26") & vbCrLf & _
                                           "姓    名：" & strST02 & vbCrLf & _
                                           "登錄字號：" & rsA.Fields("OA05") & vbCrLf & _
                                           "地    址：臺北市長安東路二段112號9樓" & vbCrLf & _
                                           "聯絡電話及分機：(02)2506-1023 分機" & vbCrLf & _
                                           "傳    真：(02)2501-1666" & vbCrLf & _
                                           "E-MAIL："
                 Case "4" 'T台灣案+FCT電子送件:回傳”證書字號OA08,ID,中文姓名”，出名代理人之間以|區隔
                       'Modified by Lydia 2024/05/16 因應商標法修法緣故,於商標申請書中需載明代理人登錄字號OA05
                       'PUB_GetAgentCP110 = PUB_GetAgentCP110 & IIf(ii > 1, "|", "") & _
                                  rsA.Fields("oa08") & "," & rsA.Fields("st26") & "," & rsA.Fields("st02")
                       PUB_GetAgentCP110 = PUB_GetAgentCP110 & IIf(ii > 1, "|", "") & _
                                  rsA.Fields("oa05") & "," & rsA.Fields("st26") & "," & rsA.Fields("st02")
                 Case "5" '代理人簽章
                       PUB_GetAgentCP110 = PUB_GetAgentCP110 & IIf(ii > 1, "、", "") & strST02
            End Select
            'end 2019/03/27
            rsA.MoveNext
         Loop
      End If
      
      'Added by Lydia 2019/03/27 T台灣案+FCT紙本送件: 同一行顯示
      If strType = "2" And strA1 <> "" Then
         'Modified by Lydia 2024/05/20 +登錄字號
         PUB_GetAgentCP110 = "ID：" & Mid(strA1, 2) & vbCrLf & _
                             "姓    名：" & Mid(strA2, 2) & vbCrLf & _
                             "登錄字號：" & Mid(strA3, 2) & vbCrLf & _
                             "地    址：臺北市長安東路二段112號9樓" & vbCrLf & _
                             "聯絡電話及分機：(02)2506-1023 分機" & vbCrLf & _
                             "傳   真：(02)2501-1666"
      End If
      'end 2019/03/27
      rsA.Close
   End If
   Set rsA = Nothing
End Function

'Added by Lydia 2015/12/08 將公報年月轉換為卷期 ex.920101 : 3001, 920116 : 3002 ...(每年會有24期)
Public Function Pub_ChgDateToTMBM07(strData As String, ByRef strVol1 As String, ByRef strVol2 As String) As String
   Pub_ChgDateToTMBM07 = ""
   strData = Right("00000" & Trim(strData), 5)
   If strData = "00000" Then Exit Function
   
   Pub_ChgDateToTMBM07 = CStr(Val(Left(strData, 3)) - 62)
   
   Select Case Right(strData, 2)
      Case "01"
         strVol1 = Pub_ChgDateToTMBM07 & "01"
         strVol2 = Pub_ChgDateToTMBM07 & "02"
         Pub_ChgDateToTMBM07 = Pub_ChgDateToTMBM07 & "02"
      Case "02"
         strVol1 = Pub_ChgDateToTMBM07 & "03"
         strVol2 = Pub_ChgDateToTMBM07 & "04"
         Pub_ChgDateToTMBM07 = Pub_ChgDateToTMBM07 & "04"
      Case "03"
         strVol1 = Pub_ChgDateToTMBM07 & "05"
         strVol2 = Pub_ChgDateToTMBM07 & "06"
         Pub_ChgDateToTMBM07 = Pub_ChgDateToTMBM07 & "06"
      Case "04"
         strVol1 = Pub_ChgDateToTMBM07 & "07"
         strVol2 = Pub_ChgDateToTMBM07 & "08"
         Pub_ChgDateToTMBM07 = Pub_ChgDateToTMBM07 & "08"
      Case "05"
         strVol1 = Pub_ChgDateToTMBM07 & "09"
         strVol2 = Pub_ChgDateToTMBM07 & "10"
         Pub_ChgDateToTMBM07 = Pub_ChgDateToTMBM07 & "10"
      Case "06"
         strVol1 = Pub_ChgDateToTMBM07 & "11"
         strVol2 = Pub_ChgDateToTMBM07 & "12"
         Pub_ChgDateToTMBM07 = Pub_ChgDateToTMBM07 & "12"
      Case "07"
         strVol1 = Pub_ChgDateToTMBM07 & "13"
         strVol2 = Pub_ChgDateToTMBM07 & "14"
         Pub_ChgDateToTMBM07 = Pub_ChgDateToTMBM07 & "14"
      Case "08"
         strVol1 = Pub_ChgDateToTMBM07 & "15"
         strVol2 = Pub_ChgDateToTMBM07 & "16"
         Pub_ChgDateToTMBM07 = Pub_ChgDateToTMBM07 & "16"
      Case "09"
         strVol1 = Pub_ChgDateToTMBM07 & "17"
         strVol2 = Pub_ChgDateToTMBM07 & "18"
         Pub_ChgDateToTMBM07 = Pub_ChgDateToTMBM07 & "18"
      Case "10"
         strVol1 = Pub_ChgDateToTMBM07 & "19"
         strVol2 = Pub_ChgDateToTMBM07 & "20"
         Pub_ChgDateToTMBM07 = Pub_ChgDateToTMBM07 & "20"
      Case "11"
         strVol1 = Pub_ChgDateToTMBM07 & "21"
         strVol2 = Pub_ChgDateToTMBM07 & "22"
         Pub_ChgDateToTMBM07 = Pub_ChgDateToTMBM07 & "22"
      Case "12"
         strVol1 = Pub_ChgDateToTMBM07 & "23"
         strVol2 = Pub_ChgDateToTMBM07 & "24"
         Pub_ChgDateToTMBM07 = Pub_ChgDateToTMBM07 & "24"
   End Select
End Function
'Added by Lydia 2015/12/17 取得國內公報代理人的事務所名稱
Public Function Pub_GetTA04(ByRef pTxt As String) As String
Dim rsXR As New ADODB.Recordset
Dim strX As String, inD As Integer
   Pub_GetTA04 = "": strX = ""
   If pTxt <> "" Then
      inD = 1
      strX = "select ta04 from tagent where ta01='T' and (ta03 like '%" & pTxt & "%' or ta04 like '%" & pTxt & "%') "
      Set rsXR = ClsLawReadRstMsg(inD, strX)
      strX = ""
      If inD = 1 Then
         rsXR.MoveFirst
         Do While Not rsXR.EOF
             If "" & rsXR.Fields("ta04") <> "" Then
                 If InStr(strX, Trim(rsXR.Fields("ta04"))) = 0 Then
                    strX = strX & Trim(rsXR.Fields("ta04")) & ","
                 End If
             End If
         rsXR.MoveNext
         Loop
         
         If strX <> "" Then strX = Mid(strX, 1, Len(strX) - 1)
      End If
   End If
   If strX = "" Then
      Pub_GetTA04 = "'" & pTxt & "'"
   Else
      Pub_GetTA04 = "'" & Replace(strX, ",", "','") & "'"
   End If
End Function

'Add By Sindy 2015/12/28 外專期限彈跳和期限管制表的備註
'Modify By Sindy 2024/5/28 +, Optional strCP176 As String = ""
Public Function PUB_GetFCPAddQuyNotes(strCP01 As String, strCP09 As String, strCP10 As String, _
                                   strCP43 As String, Optional strCnt As String = "", _
                                   Optional strNP22 As String = "", Optional strCP142 As String = "", _
                                   Optional strCP164 As String = "", Optional strCP176 As String = "") As String
Dim rsXR As New ADODB.Recordset
Dim strX As String, inD As Integer
   
   PUB_GetFCPAddQuyNotes = ""
   
   '讀取目前延期次數
   If strCnt = "" Then
      inD = 1
      'Modify By Sindy 2016/3/11 + FCP-46153 104/12/14尚未收申復
      '                            未延期:增加判斷若做CP43去檢查是否有延期時,還要過濾掉CP09不可為C類 ==> and cp09<'C'
      strX = "select DL01,sum(DLCnt) Cnt from" & _
           " (select DL01,count(*) DLCnt from datelimit" & _
           " where DL01='" & strCP09 & "' group by DL01" & _
           " Union all " & _
           " select cp09 DL01,count(*) DLCnt from caseprogress,datelimit" & _
           " where cp09='" & strCP09 & "' and cp43 is not null and cp43=DL01 and cp09<'C' group by cp09)" & _
           " group by DL01"
      Set rsXR = ClsLawReadRstMsg(inD, strX)
      If inD = 1 Then
         strCnt = rsXR.Fields(1)
      End If
   End If
   
   'Add By Sindy 2015/10/22
   '201.新案翻譯,新案202.補文件,205.申復,501.訴願等案性質不能延期二次,延期過一次,備註"不得延期"
   'Modified by Lydia 2017/01/26 +209 檢視中說、235 核對中說格式
   If strCP01 = "FCP" Then
      If (strCP10 = "201" Or strCP10 = "202" Or strCP10 = "209" Or strCP10 = "205" Or strCP10 = "235" Or strCP10 = "501") _
         And Val(strCnt) >= 1 Then
         If strCP10 = "202" And "" & strCP43 <> "" Then
            '檢查是否為新案補文件
            strX = "select cp09,cp10 from caseprogress" & _
               " where cp09='" & strCP43 & "' and cp10 in(" & NewCasePtyList & ")"
            inD = 1
            Set rsXR = ClsLawReadRstMsg(inD, strX)
            If inD = 1 Then
               PUB_GetFCPAddQuyNotes = "不得延期;"
            End If
         Else
            PUB_GetFCPAddQuyNotes = "不得延期;"
         End If
      '107.再審不能延期三次,延期過二次,備註"不得延期"
      ElseIf strCP10 = "107" And Val(strCnt) >= 2 Then
         PUB_GetFCPAddQuyNotes = "不得延期;"
      '503.行政訴訟不能延期,備註"不得延期"
      ElseIf strCP10 = "503" Then
         PUB_GetFCPAddQuyNotes = "不得延期;"
      End If
   '2015/10/22 END
   'Add By Sindy 2015/10/29 205.陳述意見
   ElseIf strCP01 = "P" And strCP10 = "205" Then
      If Val(strNP22) > 0 Then 'R12=NP22:下一程序
         strX = "select np01,np09,cp07,to_char(add_months(to_date(cp07,'YYYYMMDD'),2),'YYYYMMDD')" & _
                     " From nextprogress,caseprogress" & _
                     " where np01='" & strCP09 & "' and np22=" & strNP22 & _
                     " and np01=cp09(+)" & _
                     " and cp07 is not null"
         inD = 1
         Set rsXR = ClsLawReadRstMsg(inD, strX)
         If inD = 1 Then
            If Val("" & rsXR(1)) >= Val("" & rsXR(3)) Then
               PUB_GetFCPAddQuyNotes = "不得延期;"
            End If
         End If
      Else '進度檔
         strX = "select c1.cp43,c1.cp07,c2.cp07,to_char(add_months(to_date(c2.cp07,'YYYYMMDD'),2),'YYYYMMDD')" & _
                     " from caseprogress c1,caseprogress c2" & _
                     " where c1.cp09='" & strCP09 & "'" & _
                     " and substr(c1.cp43,1,1)='C'" & _
                     " and c1.cp43=c2.cp09(+)" & _
                     " and c2.cp07 is not null"
         inD = 1
         Set rsXR = ClsLawReadRstMsg(inD, strX)
         If inD = 1 Then
            If Val("" & rsXR(1)) >= Val("" & rsXR(3)) Then
               PUB_GetFCPAddQuyNotes = "不得延期;"
            End If
         End If
      End If
   End If
   '2015/10/29 END
   
   'Add By Sindy 2024/5/28 期限彈跳通知：【暫不送】勾選時備註顯示"待指示"(於指定送件日後)
   If strCP176 = "Y" Then
      PUB_GetFCPAddQuyNotes = "待指示;" & PUB_GetFCPAddQuyNotes
   End If
   '2024/5/28 END
   'Add By Sindy 2015/12/15
   If Val(strCP142) > 0 Then 'CP142:指定送件日期
      'Modified by Morgan 2016/6/16
      'PUB_GetFCPAddQuyNotes = "當天;"
      'Modify By Sindy 2021/4/21
      'PUB_GetFCPAddQuyNotes = ChangeWStringToTDateString(strCP142) & "當天;"
      'Modify By Sindy 2021/10/20 + , IIf(strCP164 = "3", "之後", "")
      PUB_GetFCPAddQuyNotes = "客戶指定" & ChangeWStringToTDateString(strCP142) & _
            IIf(strCP164 = "1", "當天", IIf(strCP164 = "2", "之前", IIf(strCP164 = "3", "之後", ""))) & "送件;" & _
            PUB_GetFCPAddQuyNotes
   End If
   '2015/12/15 END
   
   'Add By Sindy 2022/6/24 檢查是否已收延期,若是,下一程序+(已收延期)
   If Val(strNP22) > 0 Then 'R12=NP22:下一程序 = 未收文
      strX = "select cp09 from caseprogress" & _
                  " where cp43='" & strCP09 & "'" & _
                  " and cp10=404" & _
                  " and cp27||cp57 is null"
      inD = 1
      Set rsXR = ClsLawReadRstMsg(inD, strX)
      If inD = 1 Then
         PUB_GetFCPAddQuyNotes = "已收延期;" & PUB_GetFCPAddQuyNotes
      End If
   End If
   '2022/6/24 END
   
   Set rsXR = Nothing
End Function


'Added by Lydia 2016/02/18 以優先權檔判斷是否要主張國際優先權(106)或國內優先權(121)
Public Function PUB_CheckPDMsg(ByVal strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, strCP10 As String, strPA09 As String, strTit As String, Optional ByVal strPridata As String) As Boolean
Dim tmpBol As Boolean
Dim rsR1 As New ADODB.Recordset
Dim intA As Integer
Dim strA1 As String
Dim strA2 As String

    strA1 = "SELECT cp10 FROM CASEPROGRESS WHERE CP01='" & strCP01 & "' AND CP02='" & strCP02 & "' AND CP03='" & strCP03 & "' AND CP04='" & strCP04 & "' and cp10 in ('106','121') and cp57 is null "
    intA = 1
    Set rsR1 = ClsLawReadRstMsg(intA, strA1)
    '有收
    If intA = 1 Then
       '尚未輸入優先權資料,可先主張優先權
       If Len("" & strPridata) = 0 Then
          tmpBol = True
       '有輸入優先權資料
       Else
           strA1 = Replace(strPridata, "，", ",")
           '新案
           If InStr("101,102", strCP10) > 0 Then
                '國內優先權
                If InStr(strA1, strPA09) > 0 Then
                   strA2 = strA2 & "121" & ","
                   strA1 = Replace(Replace(strA1, strPA09, ""), ",", "") '清除國內資料
                End If
                '國際優先權
                If Len(strA1) > 0 Then
                   strA2 = strA2 & "106" & ","
                End If
                
                strA2 = ChgNewStr(strA2)
                strA1 = strA2 '保留案件性質
                rsR1.MoveFirst
                Do While Not rsR1.EOF
                   strA2 = Replace(strA2, "" & rsR1.Fields("cp10"), "")
                   rsR1.MoveNext
                Loop
                '全部-已收文
                If Len(strA2) < 3 Then
                   tmpBol = True
                   strA2 = strA1 '保留案件性質
                End If
                '要收文的案件性質名稱
                If InStr(strA2, 主張優先權) > 0 Then
                   strTit = strTit & "主張國際優先權、"
                End If
                If InStr(strA2, "121") > 0 Then
                   strTit = strTit & "主張國內優先權、"
                End If
                strTit = Mid(strTit, 1, Len(strTit) - 1)
           '非新案
           Else
              If strCP10 = "121" Then
                 strTit = "主張國內優先權"
                 If InStr(strA1, strPA09) > 0 Then
                    tmpBol = True
                 Else
                    tmpBol = False
                 End If
              Else
                 strTit = "主張國際優先權"
                 strA2 = Replace(strA1, strPA09, "")
                 If InStr(strA1, strPA09) = 0 And Len(Replace(strA1, ",", "")) >= 3 Then
                    tmpBol = True
                 Else
                    tmpBol = False
                 End If
              End If
           End If
       End If
    '沒收
    Else
       tmpBol = False
    End If
    
    Set rsR1 = Nothing
    PUB_CheckPDMsg = tmpBol
End Function

'Added by Morgan 2016/3/28
'指示信EMail內容
Public Function PUB_GetOrderLetterContent(Optional pSys As String = "P") As String
   'Modified by Morgan 2018/8/13 +CFP
   If pSys = "CFP" Then
      'Modified by Morgan 2019/11/20 落款改寄件時載入(有圖)
      'PUB_GetOrderLetterContent = "<span style=""font-size:12pt;font-family:&quot;Times New Roman&quot;,&quot;serif&quot;"">&nbsp;" & vbCrLf & _
         "Dear Associates," & vbCrLf & vbCrLf & _
         "Please refer to the attachments for the subject case." & vbCrLf & vbCrLf & _
         "Please acknowledge receipt of this email. Thank you for your assistance." & vbCrLf & vbCrLf & vbCrLf & _
         "Best Regards," & vbCrLf & _
         "Jerry C. Y. Lin" & vbCrLf & _
         "Patent Attorney" & vbCrLf & vbCrLf & _
         "CYL/" & PUB_GetST07(strUserNum) & vbCrLf & vbCrLf & _
         "</span><span style=""font-size:7.5pt;font-family:&quot;Times New Roman&quot;,&quot;serif&quot;"">" & vbCrLf & _
         "<i><b>Tai E International Patent & Law Office</b>" & vbCrLf & _
         "9F, 112, Section 2, Chang-An East Road Taipei, Taiwan, R.O.C. ;P.O. Box 46-478, Taipei, Taiwan, R.O.C." & vbCrLf & _
         "Tel: 886-2-25061023, 886-2-25081531 Fax: 886-2-25068147, 886-2-25076571, 886-2-25090804," & vbCrLf & _
         "886-2-25064319 URL:http://www.taie.com.tw E-mail:patent@taie.com.tw </i>" & vbCrLf & _
         "************* Email Confidentiality Notice ********************" & vbCrLf & _
         "This e-mail transmission is intended only for the use of the individual" & vbCrLf & _
         "or entity to which it is addressed, and may contain information that is" & vbCrLf & _
         "privileged, confidential and exempt from disclosure under applicable" & vbCrLf & _
         "law.  If the reader is not the intended recipient, you are hereby" & vbCrLf & _
         "notified that any dissemination, distribution or copying of this" & vbCrLf & _
         "communication is strictly prohibited. If you have received this" & vbCrLf & _
         "transmission in error, please notify us immediately, and return the" & vbCrLf & _
         "original message to us at the above address. We greatly appreciate your" & vbCrLf & _
         "cooperation.</span>"
      'Modified by Morgan 2024/4/10 對外統一用 Dear Colleagues --林總
      'Dear Associates --> Dear Colleagues
      PUB_GetOrderLetterContent = "Dear Colleagues," & vbCrLf & vbCrLf & _
         "Please refer to the attachments for the subject case." & vbCrLf & vbCrLf & _
         "Please acknowledge receipt of this email. Thank you for your assistance." & vbCrLf & vbCrLf & vbCrLf & _
         "Best Regards," & vbCrLf & _
         "Jerry C. Y. Lin" & vbCrLf & _
         "CEO, Partner, Patent Attorney" & vbCrLf & vbCrLf & _
         "CYL/" & Pub_StrUserSt17 & vbCrLf & vbCrLf
   Else
      'Modified by Morgan 2019/8/26 本文統一"** 主旨如附檔所示,請確認收到此郵件 **"--品薇
      'Modified by Morgan 2019/11/20 落款改寄件時載入(有圖)
      'PUB_GetOrderLetterContent = vbCrLf & "** 主旨如附檔所示,請確認收到此郵件 **" & vbCrLf & vbCrLf & vbCrLf & _
         "台一國際專利法律事務所  / " & PUB_GetST07(strUserNum) & vbCrLf & _
         "電　話：(02)25061023" & vbCrLf & _
         "傳　真：(02)25011666" & vbCrLf & _
         "URL:https://www.taie.com.tw" & vbCrLf & _
         "*************保密警語********************" & vbCrLf & _
         "本信件僅授權於指定之收信人取閱之用，信件中可能含有機密性資訊。" & vbCrLf & _
         "如果您並非被指定之收信人，任何未經授權而擅自使用此信件所含之機密資訊的行為是被嚴格禁止的。" & vbCrLf & _
         "如果您在任何未經授權的情形之下收到本信件，煩請您立即告知原發信人並將此信件回傳至以上地址。" & vbCrLf & _
         "謝謝您的合作。"
      'Modified by Morgan 2020/2/25 +時　　祺
      'Modified by Morgan 2023/5/3 2023/5/11起P案雙署名改郭雅娟
      PUB_GetOrderLetterContent = "敬啟者：" & vbCrLf & vbCrLf & _
         "主旨如附檔所示,請確認收到此郵件。" & vbCrLf & vbCrLf & vbCrLf & _
         "耑此　順頌" & vbCrLf & vbCrLf & _
         "時　　祺" & vbCrLf & vbCrLf & _
         "林景郁、" & IIf(strSrvDate(1) >= "20230511", "郭雅娟", "王錦寬") & " /<font style=""font-size:11pt"">&nbsp;" & Pub_StrUserSt17 & "</font>"
   End If
End Function

'Added by Morgan 2016/3/28
Public Function PUB_SendOrderLetterP(pRecNo As String, Optional pSubject As String, Optional pAutoMail As Boolean = False) As Boolean
   Dim strPath As String, strFile As String
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stCP01 As String, stCP02 As String, stCP03 As String, stCP04 As String, stCP44 As String, stCP116 As String
   Dim stCP10 As String, stCP12 As String, stAF13 As String, stCPM04 As String
   Dim lngMousePointer As Long
   Dim stCP140 As String 'Added by Morgan 2018/8/30
   
   lngMousePointer = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   
On Error GoTo ErrHnd
   
   'Modified by Morgan 2018/9/17 指示信要剔除已刪除
   stSQL = "select cp01,cp02,cp03,cp04,cp10,cp12,cp44,cp116,cpm04,cpp02,af13,cp140 from caseprogress,patent,casepropertymap,casepaperpdf,appform" & _
      " where cp09='" & pRecNo & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 and cpp01(+)=cp09 and substr(upper(cpp02(+)),-9)='.DATA.PDF' and af01(+)=cp09"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With rsQuery
      
      If IsNull(.Fields("cpp02")) Then
         MsgBox "卷宗區沒有指示信PDF檔！", vbCritical
         GoTo ErrHnd
      End If
      '指示信
      strFile = "" & .Fields("cpp02")
      
      '下載指示信
      strPath = App.path & "\" & strUserNum
      If PUB_GetAttachFile_CPP(pRecNo, strFile, strPath) = False Then
         MsgBox "指示信PDF檔下載失敗！", vbCritical
         GoTo ErrHnd
      End If
      
      stCP01 = .Fields("cp01")
      stCP02 = .Fields("cp02")
      stCP03 = .Fields("cp03")
      stCP04 = .Fields("cp04")
      stCP44 = "" & .Fields("cp44")
      stCP116 = "" & .Fields("cp116")
      stCP10 = "" & .Fields("cp10")
      stCP12 = "" & .Fields("cp12")
      stAF13 = "" & .Fields("af13")
      stCPM04 = "" & .Fields("cpm04")
      stCP140 = "" & .Fields("cp140") 'Added by Morgan 2018/8/30
      End With
      
      With frm880019
      .m_isCFFagent = True 'Add By Sindy 2022/3/14 讓寄信辨識是CF代理人
      .m_bolAutoMail = pAutoMail '自動寄送
      .m_bolSaveMail = True
      .m_CP01 = stCP01
      .m_CP02 = stCP02
      .m_CP03 = stCP03
      .m_CP04 = stCP04
      .m_CP09 = pRecNo
      .m_CP10 = stCP10
      
      'FMP時列印寄件備份歸檔存卷
      If Left(stCP12, 1) = "F" Then
         .chkPrint.Visible = True
         '.chkPrint.Value = 1 'Removed by Morgan 2020/5/6 FMP案無須再印寄件備份--潘韻丞(外專用Outlook寄，自動分信到卷宗區--敏莉)
      End If
      '主旨
      If pSubject = "" Then
         'Modified by Morgan 2016/5/12 先抓AF13
         pSubject = stAF13
         If pSubject = "" Then
            'Modify By Sindy 2018/1/2 + & "-" & : 本所案號要加P-, CFP-
            pSubject = stCP01 & "-" & stCP02 & IIf(stCP03 & stCP04 = "000", "", "-" & stCP03) & IIf(stCP04 = "00", "", "-" & stCP04) & "案" & stCPM04 & PUB_GetRelateCasePropertyName(pRecNo, "1")
         End If
      End If
      .txtSubject = pSubject
      
      '本文
      .txtContent = PUB_GetOrderLetterContent(stCP01)
      
      '信箱(CF代理人)
      If stCP44 = "" Then
         '抓AB類收文號的代理人，預設最後發文日最大收文號的代理人...同發文作業預設的代理人(AddAgent)
         PUB_GetCP44 stCP01, stCP02, stCP03, stCP04, stCP44, stCP116
      End If
      
      If stCP44 <> "" Then
         .SetEmail "", "", stCP44, stCP116, True
      End If
      
      If stCP01 = "CFP" Then
         '.m_bolCfpLetter = True 'Removed by Morgan 2019/11/20
         .txtBCC = strUserNum
         .chkImportant.Visible = True 'Added by Morgan 2018/9/12
      End If
      
      .m_bolPLetter = True 'Added by Morgan 2019/11/20
      .SetAttach strFile
      'Modified by Morgan 2016/5/20 改可點選卷宗區檔案
      '.cmdAttach.Visible = False
      .m_bolAttFromCpp = True
      'end 2016/5/20
         
      End With
      
      Screen.MousePointer = vbDefault
      frm880019.Show vbModal
      If frm880019.m_bolDone Then
         'Removed by Morgan 2019/11/22 改在 frm880019
         'If frm880019.m_SMB02 <> "" Then 'Added by Morgan 2018/9/13 CFP指示信有可能不寄
         '   strSql = "update appform set af11=" & frm880019.m_SMB02 & ",af12=" & frm880019.m_SMB03 & ",af14='" & strUserNum & "' where af01='" & pRecNo & "'"
         '   cnnConnection.Execute strSql, intI
         'End If
         'end 2019/11/22
         
         'Added by Morgan 2016/5/17
         '催收達、催提申、催公開指示信寄送完成要更新發文日
         strSql = "update caseprogress set cp27=" & strSrvDate(1) & " where cp28='" & pRecNo & "' and cp10 in ('952','953','954') and cp27 is null"
         cnnConnection.Execute strSql, intQ
         strSql = "update caseprogress set cp27=" & strSrvDate(1) & ",cp28=cp09 where cp09='" & pRecNo & "' and cp10 ='411' and cp27 is null"
         cnnConnection.Execute strSql, intQ
         'end 2016/5/17
         
         If stCP140 <> "" Then OrderLetterFlowStatusUpdate stCP140 'Added by Morgan 2018/8/30
         
         PUB_SendOrderLetterP = True
      End If
      Unload frm880019
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
   Set rsQuery = Nothing
   Screen.MousePointer = lngMousePointer
End Function

'Added by Morgan 2016/3/28
'列印定稿成Pdf檔
Public Function PUB_PrintDocAsPdf(ByVal pReceiveNo As String, ByVal pletterSitu As String, ByVal pSitu As String, ByVal pLetterRecNo As String, Optional ByRef pSavePath As String, Optional ByRef pSaveName As String) As Boolean
   Dim Os_Printer As String
   Dim strFullFileName As String
   
On Error GoTo ErrHnd
   
   If pSavePath = "" Then pSavePath = App.path & "\" & strUserNum
   If Dir(pSavePath, vbDirectory) = "" Then MkDir pSavePath
   If pSaveName = "" Then pSaveName = "$" & pReceiveNo & ".PDF"
   strFullFileName = pSavePath & "\" & pSaveName
   If Dir(strFullFileName) <> "" Then Kill strFullFileName
   
   'Added by Morgan 2018/5/29
   If pub_Word2Pdf Then
      NowPrint pReceiveNo, pletterSitu, pSitu, True, strUserNum, , , , , 1, , , , , , , , pLetterRecNo
      g_WordAp.ActiveDocument.ExportAsFixedFormat OutputFileName:=strFullFileName, ExportFormat:=17, OpenAfterExport:=False
      g_WordAp.Quit wdDoNotSaveChanges
      Set g_WordAp = Nothing
   Else
   'end 2018/5/29
   
      Os_Printer = PUB_GetOsDefaultPrinter
      frmPDF.Show
      frmPDF.StartProcess pSavePath, pSaveName
      PUB_SetOsDefaultPrinter Printer.DeviceName
      PUB_SetWordActivePrinter
      NowPrint pReceiveNo, pletterSitu, pSitu, False, strUserNum, , , , , 1, , , , , True, , , pLetterRecNo
      frmPDF.EndtProcess
      Unload frmPDF
      PUB_SetOsDefaultPrinter Os_Printer
      
   End If 'Added by Morgan 2018/5/29
   
   If Dir(strFullFileName) <> "" Then
      PUB_PrintDocAsPdf = True
   End If
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function

'Added by Morgan 2016/3/28
'上傳指示信pdf到卷宗區
Public Function PUB_UploadOrderLetter(pUploadFilePath As String, pRecNo As String, Optional pPdfName As String) As Boolean
   
   Dim rsQuery As ADODB.Recordset, intQ As Integer, stSQL As String
   Dim oFileSys As New FileSystemObject
   Dim oFile
   Dim boInTrans As Boolean
   
On Error GoTo ErrHnd

   If pPdfName = "" Then
      stSQL = "select cp01,cp02,cp03,cp04,cp10 from caseprogress where cp09='" & pRecNo & "'"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         pPdfName = PUB_CaseNo2FileName(rsQuery("cp01"), rsQuery("cp02"), rsQuery("cp03"), rsQuery("cp04")) & "." & rsQuery("cp10") & ".DATA.PDF"
      End If
   End If

   '寫回卷宗區並更新AppForm
   If Dir(pUploadFilePath) <> "" Then
      Set oFile = oFileSys.GetFile(pUploadFilePath)
      
      cnnConnection.BeginTrans
      boInTrans = True
      '鎖定
      strSql = "update AppForm set af02='" & strUserNum & "' where af01='" & pRecNo & "'"
      cnnConnection.Execute strSql, intQ
      '刪除卷宗區
      strSql = "delete from CasePaperPDF where cpp01='" & pRecNo & "' and upper(cpp02)=upper('" & pPdfName & "')"
      cnnConnection.Execute strSql, intQ
      '刪除ftp
      PUB_DelFtpFile2 pRecNo, " and upper(cpp02)=upper('" & pPdfName & "')"
      '上傳ftp並回寫卷宗區
      SaveAttFile_PDF pRecNo, pUploadFilePath, pPdfName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False, , , True
      '更新上傳日期
      'Modify by Morgan 2016/5/12 自行判發同時更新判發日
      strSql = "update AppForm set af03=to_char(sysdate,'yyyymmdd'),af07=decode(af06,af02,to_char(sysdate,'yyyymmdd'),af07) where af01='" & pRecNo & "' and af02='" & strUserNum & "'"
      cnnConnection.Execute strSql, intQ
      
      cnnConnection.CommitTrans
      boInTrans = False
      PUB_UploadOrderLetter = True
   End If
   
ErrHnd:
   If Err.Number > 0 Then
      If boInTrans Then cnnConnection.RollbackTrans
      MsgBox Err.Description, vbCritical
   End If
   
   Set oFileSys = Nothing
   Set oFile = Nothing
   Set rsQuery = Nothing
End Function

'Add By Sindy 2016/5/9 內商電子收文請款提醒訊息
Public Sub PUB_TCaseEFeeRemind(strCP09 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset

StrSQLa = "SELECT cp01,cp02,cp03,cp04,cp09,cp16,cp17,cp18,cp64" & _
          " FROM caseprogress WHERE CP09='" & strCP09 & "' and cp60 is null and nvl(cp16,0)>0 and cp140 is not null"
rsA.CursorLocation = adUseClient
rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
If rsA.RecordCount > 0 Then
   MsgBox "收文費用 = " & rsA("cp16") & "  規費 = " & rsA("cp17") & "  點數 = " & rsA("cp18") & vbCrLf & vbCrLf & _
          "進度備註：" & rsA("cp64"), vbExclamation + vbOKOnly, "電子收文請款提醒訊息"
End If
If rsA.State <> adStateClosed Then rsA.Close
Set rsA = Nothing
End Sub

'Added by Lydia 2016/06/16 外專分案(工程師主管分案)-檢查分案的系統別
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Function PUB_CheckFCPsys(strPA01 As String) As Boolean
   If strPA01 <> "FCP" And strPA01 <> "FG" And strPA01 <> "P" And strPA01 <> "PS" And strPA01 <> "CFP" And strPA01 <> "CPS" Then
      PUB_CheckFCPsys = False
   Else
      PUB_CheckFCPsys = True
   End If
End Function

'Added by Lydia 2016/06/17 外專分案(工程師主管分案)-設定期限
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Function PUB_GetFCPsetDate(iCP01 As String, iCP10 As String) As String
'iCP01 = 系統別, iCP10 =案件性質
Dim strTmp As String

   If iCP01 = "FCP" And iCP10 = "926" Then
      strTmp = CompWorkDay(12, strSrvDate(1), 0)
   End If
   PUB_GetFCPsetDate = strTmp
End Function

'Add By Sindy 2016/6/20 抓進度備註欄專利個體
'Modified by Morgan 2023/3/29 PA179啟用後改 casefee2.cf209
Public Function PUB_GetP_PA91Individual(strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String) As String
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
   
   PUB_GetP_PA91Individual = ""
   'Modified by Morgan 2023/3/29 +casefee2.cf209
   'Modified by Morgan 2025/5/12 +判斷PCT的案件性質為109
   If rsA.State <> adStateClosed Then rsA.Close
   StrSQLa = "SELECT pa91,cf209" & _
             " FROM patent,casefee2 WHERE PA01='" & strCP01 & "' and PA02='" & strCP02 & "'" & _
             " and PA03='" & strCP03 & "' and PA04='" & strCP04 & "' and cf201(+)=pa01" & _
             " and cf202(+)=pa09 and cf205(+)=pa179 and cf203(+)=decode(pa09,'056','109','10'||pa08)" & _
             " and cf210(+)<=" & strSrvDate(1) & " order by cf210 desc"
   rsA.CursorLocation = adUseClient
   rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
   If rsA.RecordCount > 0 Then
      'Added by Morgan 2023/3/29
      If strSrvDate(1) >= PA179啟用日 Then
         PUB_GetP_PA91Individual = "" & rsA.Fields("cf209")
      Else
      'end 2023/3/29
         If InStr(1, "" & rsA.Fields("pa91"), "大個體", 1) > 0 Then
            PUB_GetP_PA91Individual = "大個體"
         ElseIf InStr(1, "" & rsA.Fields("pa91"), "小個體", 1) > 0 Then
            PUB_GetP_PA91Individual = "小個體"
         ElseIf InStr(1, "" & rsA.Fields("pa91"), "微個體", 1) > 0 Then
            PUB_GetP_PA91Individual = "微個體"
         End If
      End If 'Added by Morgan 2023/3/29
   End If
   If rsA.State <> adStateClosed Then rsA.Close
   Set rsA = Nothing
End Function

'Added  by Morgan 2016/06/20 外專分案(工程師主管分案)-設定承辦期限
'Add By Sindy 2021/9/6 + Optional ByVal strSelCP06 As String :點選已收文或未收文程序的所限
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Sub PUB_GetFCPsetCP48(ByRef bolRun As Boolean, ByRef nPA() As String, ByRef nCP27 As String, _
   ByRef nCP10 As Object, ByRef nCP14 As Object, ByRef nCP122 As String, ByRef nCP06 As Object, ByRef nCP07 As Object, _
   ByRef nCP48 As Object, Optional ByRef nEP06 As Object, Optional nCMB As Object, Optional ByRef nCP05 As Object, _
   Optional ByVal strSelCP06 As String)
'bolRun 讀檔時不必觸發
'nPA 專利基本檔
'nCP27 發文日
'nCP10 / nCP10m  : 案件性質 / 名稱
'nCP14 承辦人
'nCP122 是否分案
'nCP06 / nCP07 : 所限/法限
'nCP48 承辦期限
'nEP06 文件齊備日
'nCMB 承辦期限下拉選單
'nCP05 收文日
Dim strR1 As String
Dim intQ As Integer
Dim adoR1 As New ADODB.Recordset
Dim bolExcept As Boolean
Dim strCPM34 As String
Dim strInStaffID As String 'Added by Morgan 2024/6/18 內翻人員對應的所內編號
   
   'Add By Sindy 2021/11/5 AEP(加速審查422)、PPH(高速審查431)分案時，本所及承辦期限應為"空"，待輸入通知實審日再依原設定規則掛期限
   'Modified by Morgan 2024/11/14 +447再審查加速審查
   If nCP10 = "422" Or nCP10 = "431" Or nCP10 = "447" Then Exit Sub
      
   '讀檔時不必觸發,收文自動分案的也要執行
   'Modify By Sindy 2021/9/3 + Or (nCP14.Text = "" And nCP10 = "924" And nCP122 = "")
   If (bolRun = True Or (nCP14.Text <> "" And nCP122 = "") Or (nCP14.Text = "" And nCP10 = "924" And nCP122 = "")) And nCP27 = "" Then
      '翻譯
      If nCP10.Text = "201" Then
         If nCP14.Text <> nCP14.Tag Then
            '設定舜禹(F5588)、捷恩凱(F5653)時預設1個月
            'Modified by Lydia 2018/01/04 F5588-> 外翻_舜禹 F5653-> 外翻_捷恩凱 ,+迅達
            'Modified by Lydia 2025/03/13 改用模組取得
            'If nCP14.Text = 外翻_舜禹 Or nCP14.Text = 外翻_捷恩凱 Or nCP14.Text = 外翻_迅達 Then
            If InStr(Pub_SetF51Order("F", ""), nCP14.Text) > 0 Then
               nCP48.Text = TransDate(PUB_GetWorkDay1(CompDate(1, 1, strSrvDate(1)), False), 1)
            '內翻預設75天
            'Modified by Lydia 2018/09/28 只要是所內員工,不論上下班都預設承辦天數
            'ElseIf PUB_GetMapID(nCP14.Text) <> "" Then
            ElseIf GetStaffName(nCP14.Text) <> "" Then
                'Modified by Lydia 2018/11/01 國內所外(非員工)譯者(F51): 45天
                'nCP48.Text = TransDate(PUB_GetWorkDay1(CompDate(2, 75, strSrvDate(1)), False), 1)
                strR1 = PUB_GetST03(nCP14.Text)
                strInStaffID = Pub_GetField("staff_idmap", "sim02='" & nCP14.Text & "'", "sim01")
                'Modified by Morgan 2024/6/18 約定薪資的翻譯人員比照外翻人員的規則
                'If strR1 = "F51" Then
                If strR1 = "F51" Or Right(strInStaffID, 2) >= "9A" Then
                'end 2024/6/18
                    nCP48.Text = TransDate(PUB_GetWorkDay1(CompDate(2, 45, strSrvDate(1)), False), 1)
                ElseIf strR1 = "F52" Then '所內員工下班翻譯(F52): 75天
                    nCP48.Text = TransDate(PUB_GetWorkDay1(CompDate(2, 75, strSrvDate(1)), False), 1)
                Else                                 '所內員工上班翻譯(員工編號): 自設期限
                    nCP48.Text = ""
                End If
                'end 2018/11/01
            Else
               nCP48.Text = ""
            End If
         End If
      '檢視中說、210 製作中說、235核對中說格式
      ElseIf InStr("209,210,235", nCP10.Text) > 0 And nCP10.Text <> "" Then
         If nEP06.Text <> "" And nEP06.Text <> nEP06.Tag Then
            nCP48.Text = TransDate(Pub_GetHandleDay(nPA(1), nPA(9), nCP10, DBDATE(nEP06.Text)), 1)
         End If
         nEP06.Tag = nEP06.Text
      'Add By Sindy 2021/9/3
      '亭妙,淑華:
      '會稿分案時, 檢查是否有中說(201翻譯、209檢視中說、210製作中說、235核對中說格式)未發文, 若有, 不計算, 也不帶承辦工程師
      '若無 , 則計算:
      '本所期限=系統日+14個工作天(不可大於有點未收文或未發文的所限, 若超過同其所限)
      '承辦期限=所限-2個工作天
      '自動帶出承辦工程師
      ElseIf nCP10.Text = "924" Then '會稿
         strR1 = "select * from caseprogress where cp01='" & nPA(1) & "' and cp02='" & nPA(2) & "' and  cp03='" & nPA(3) & "' and cp04='" & nPA(4) & "'" & _
                 " and cp10 in ('201','209','210','235') and cp57||cp27 is null"
         intQ = 1
         Set adoR1 = ClsLawReadRstMsg(intQ, strR1)
         If intQ = 0 Then
            '本所期限=系統日+14個工作天
            nCP06.Text = TransDate(CompWorkDay(14, CompDate(2, 1, strSrvDate(1)), 0), 1)
            '不可大於有點未收文或未發文的所限, 若超過同其所限
            If Val(strSelCP06) > 0 Then
               If Val(DBDATE(nCP06.Text)) > Val(DBDATE(strSelCP06)) Then
                  nCP06.Text = TransDate(DBDATE(strSelCP06), 1)
               End If
            End If
            '承辦期限=所限-2個工作天
            nCP48.Text = TransDate(CompWorkDay(2, CompDate(2, -1, DBDATE(nCP06.Text)), 1), 1)
            '自動帶出承辦工程師
            strR1 = nCP14
            If PUB_GetFCPCP14_F21(nPA, strR1) = True Then '抓承辦人為工程師
               nCP14 = strR1
            End If
         End If
         
      '其他
      Else
         '若實審"沒有"發明或分割未發文設承辦期限15個工作天
         If nCP10.Text = "416" And nCP14.Text <> "" Then
            strR1 = "select * from caseprogress where cp01='" & nPA(1) & "' and cp02='" & nPA(2) & "' and  cp03='" & nPA(3) & "' and cp04='" & nPA(4) & "' and cp10 in ('101','307') and cp57||cp27 is null"
            intQ = 1
            Set adoR1 = ClsLawReadRstMsg(intQ, strR1)
            If intQ = 0 Then
               nCP48.Text = TransDate(CompWorkDay(15, DBDATE(nCP05.Text)), 1)
               bolExcept = True
            End If
         End If
         
         If nCP10.Text <> nCP10.Tag Then
            If bolExcept = False Then
               nCP48.Text = TransDate(Pub_GetHandleDay(nPA(1), nPA(9), nCP10.Text, DBDATE(nCP05.Text), , , , nPA(1) & "-" & nPA(2) & "-" & nPA(3) & "-" & nPA(4)), 1)
            End If
         End If
      End If
      
      'Add By Sindy 2021/8/12 不是主管機關期限,本所期限=承辦期限＋5個工作天,法定期限空白
      CheckOC3
      strCPM34 = ""
      strSql = "select cpm34 from casepropertymap where cpm01='" & nPA(1) & "' and cpm02='" & nCP10.Text & "'"
      AdoRecordSet3.CursorLocation = adUseClient
      AdoRecordSet3.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If AdoRecordSet3.RecordCount > 0 Then
         strCPM34 = "" & AdoRecordSet3.Fields(0)
      End If
      'Modify By Sindy 2022/3/15 亭妙:當 非智慧局期限之案件性質，本身已有法定期限資料的狀況下，請不要重新更新期限，以自身的期限為主。
      If strCPM34 = "N" And Val(nCP07) = 0 Then
         nCP06 = TransDate(PUB_GetFCPOurDeadline(DBDATE(nCP48), , , , "N"), 1)
         nCP07 = ""
      End If
      '2021/8/12 END
      
      If nCP48.Text <> "" And nCP06.Text <> "" And Val(nCP48.Text) > Val(nCP06.Text) Then
         nCP48.Text = nCP06.Text
      End If
      nCMB.ListIndex = -1
      nCP14.Tag = nCP14.Text
      nCP10.Tag = nCP10.Text
   End If
   
   Set adoR1 = Nothing
End Sub

'Add By Sindy 2021/9/3 該案號最大收文日最大總收文號,承辦人為工程師
Public Function PUB_GetFCPCP14_F21(strCase As Variant, ByRef strCP14 As String) As Boolean
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String, intA As Integer

   If Trim(strCP14) = "" Then
      '該案號最大收文日最大總收文號,承辦人為工程師
      'Modified by Lydia 2022/07/05 排除F4102,F4104,F4105 =>  st01<>'F4102' and st01<>'F4104' and st01<>'F4105'
      'Modified by Lydia 2024/01/04 判斷中說性質
      'strSqlA = "select cp09,cp14 from caseprogress,staff" & _
                  " where cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "'" & _
                  " and cp14=st01(+) and st03='F21' and cp14 is not null and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' " & _
                  " order by SQLDatet2(CP05) DESC, CP66 DESC, CP67 DESC, CP09 DESC"
      StrSQLa = "select cp09,cp10,cp14 from caseprogress,staff" & _
                  " where cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "'" & _
                  " and cp14=st01(+) and st03 in ('F21','F51','F52') and cp14 is not null and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' " & _
                  " order by SQLDatet2(CP05) DESC, CP66 DESC, CP67 DESC, CP09 DESC"
      intA = 1
      Set rsA = ClsLawReadRstMsg(intA, StrSQLa)
      If intA = 1 Then
         rsA.MoveFirst
         strCP14 = "" & rsA.Fields("cp14")
         'Added by Lydia 2024/01/04 判斷中說性質
         If InStr("201,209,235,210", "" & rsA.Fields("cp10")) > 0 Then
            strCP14 = PUB_GetSpecCP14(strCase(1) & strCase(2) & strCase(3) & strCase(4))
         End If
         'end 2024/01/04
         
         strCP14 = PUB_SetEng(strCP14) 'Added by Lydia 2024/03/06 外專機械設計組人員異動調整程式
         
         '檢查人員是否存在或離職
         If ChkStaffST04(strCP14, False) = True Then
            '是離職承辦工程師，承辦人請自動帶該組別副理（日文組案件時，化學案就帶簡副理(99037)；電機案就帶林副理(94012)）
            'Modified by Lydia 2023/01/12 日文組只需單純回傳離職人員的主管
            'strCP14 = PUB_GetFCPEngSup(strCP14, True)
            strCP14 = PUB_GetFCPEngSup(strCP14, True, True)
         End If
         PUB_GetFCPCP14_F21 = True
      End If
   End If
   
   Set rsA = Nothing
End Function

'Add By Sindy 2021/9/6 該案號最大收文日最大總收文號,承辦人為工程師
Public Function PUB_GetCP14_P11(strCase As Variant, ByRef strCP14 As String) As Boolean
Dim rsA As New ADODB.Recordset
Dim StrSQLa As String, intA As Integer

   If Trim(strCP14) = "" Then
      '該案號最大收文日最大總收文號,承辦人為工程師
      'Modified by Lydia 2024/01/04 判斷中說性質
      'strSqlA = "select cp09,cp14,st04 from caseprogress,staff" & _
                  " where cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "'" & _
                  " and cp14=st01(+) and st03='P11' and cp14 is not null" & _
                  " order by SQLDatet2(CP05) DESC, CP66 DESC, CP67 DESC, CP09 DESC"
      StrSQLa = "select cp09,cp10,cp12,cp13,cp14,st04 from caseprogress,staff" & _
                  " where cp01='" & strCase(1) & "' and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "'" & _
                  " and cp14=st01(+) and st03 in ('P11','F21','F51','F52') and cp14 is not null" & _
                  " order by SQLDatet2(CP05) DESC, CP66 DESC, CP67 DESC, CP09 DESC"
      intA = 1
      Set rsA = ClsLawReadRstMsg(intA, StrSQLa)
      If intA = 1 Then
         rsA.MoveFirst
         strCP14 = "" & rsA.Fields("cp14")
         'Added by Lydia 2024/01/04 判斷中說性質
         If InStr("201,209,235,210,910,942", "" & rsA.Fields("cp10")) > 0 And Left("" & rsA.Fields("cp12"), 1) = "F" Then
            strCP14 = PUB_GetSpecCP14(strCase(1) & strCase(2) & strCase(3) & strCase(4))
         End If
         'end 2024/01/04
         
'         '檢查人員是否存在或離職
'         If ChkStaffST04(strCP14, False) = True Then
'            '是離職承辦工程師，承辦人請自動帶該組別副理（日文組案件時，化學案就帶簡副理(99037)；電機案就帶林副理(94012)）
'            strCP14 = PUB_GetFCPEngSup(strCP14, True)
'         End If
         PUB_GetCP14_P11 = True
      End If
   End If
   
   Set rsA = Nothing
End Function

'Added by Lydia 2016/06/20 外專分案(工程師主管分案)-提醒是否重新計算承辦期限
'Modify By Sindy 2021/9/8 + , Optional ByRef iCP64 As Object
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Function PUB_CheckFCPshowMsg(ByVal bolRun As Boolean, ByRef iPA() As String, ByRef iCP27 As String, _
   ByRef iCP10 As Object, ByRef iCP14 As Object, ByRef iCP122 As String, ByRef iCP06 As Object, ByRef iCP07 As Object, _
   ByRef iCp48 As Object, Optional ByRef iCP10m As String, Optional ByRef iEP06 As Object, Optional nCMB As Object, _
   Optional ByRef iCP05 As Object, Optional ByRef m203CP48 As String, Optional ByRef iCP64 As Object) As Boolean
'm203CP48 主動修正的預設承辦期限
Dim inA As Integer
Dim rsA1 As New ADODB.Recordset
Dim strA1 As String, strA2 As String, strA3 As String

If iPA(1) <> "FCP" Then Exit Function

   '102新法,主動修正無期限,但若新案已發文且201,209,210,235未發文則更新承辦期限預設為上述程序之所限
   m203CP48 = ""
   '已提申
   If iCP27 = "" And iPA(10) <> "" And (iCP10 = "203" Or iCP10 = "206") Then '206.補充說明
      strA1 = "select cp06,cp07,cpm03 from caseprogress,casepropertymap" & _
              " where cp01='" & iPA(1) & "' and cp02='" & iPA(2) & "' and cp03='" & iPA(3) & "' and cp04='" & iPA(4) & "'" & _
              " and cp10 in('201','209','235','210') and cp27||cp57 is null and cp06>0 and cpm01(+)=cp01 and cpm02(+)=cp10"
      inA = 1
      Set rsA1 = ClsLawReadRstMsg(inA, strA1)
      If inA = 1 Then
'         MsgBox "有" & rsA1.Fields("cpm03") & "未發文，" & iCP10m & "的承辦期限將設定為" & rsA1.Fields("cpm03") & "的本所期限！", vbInformation
'         iCp48 = TransDate("" & rsA1.Fields("cp06"), 1) '新案翻譯等的所限
'         'Add By Sindy 2021/8/11 承辦期限＋5個工作天為本所期限
'         iCP06 = TransDate(PUB_GetFCPOurDeadline(DBDATE(iCp48), , , , "N"), 1)

         'Modify By Sindy 2021/8/18
         'A. 若新案翻譯、檢視中說、核對中說格式、製作中說 未發文：
         '主動修正的本所期限 = 新案翻譯的本所期限
         '          承辦期限 = 本所期限往前-5個工作天
         MsgBox "有" & rsA1.Fields("cpm03") & "未發文，" & vbCrLf & _
                iCP10m & "的本所期限將設定為" & rsA1.Fields("cpm03") & "的本所期限，承辦期限 = 本所期限往前-5個工作天！", vbInformation
         iCP06 = TransDate("" & rsA1.Fields("cp06"), 1)
         iCp48 = TransDate(CompWorkDay(5, CompDate(2, -1, DBDATE(iCP06)), 1), 1)
         '2021/8/18 END
         m203CP48 = iCp48
         
      '若有發明實審未發文時,則所限與法限設相同
      ElseIf iPA(8) = "1" Then 'And iCP07 = ""
         strA1 = "select cp06,cp07,cpm03 from caseprogress,casepropertymap" & _
                 " where cp01='" & iPA(1) & "' and cp02='" & iPA(2) & "' and cp03='" & iPA(3) & "' and cp04='" & iPA(4) & "'" & _
                 " and cp10='416' and cp27||cp57 is null and cp06>0 and cpm01(+)=cp01 and cpm02(+)=cp10"
         inA = 1
         Set rsA1 = ClsLawReadRstMsg(inA, strA1)
         If inA = 1 Then
            'MsgBox "有" & rsA1.Fields("cpm03") & "未發文，" & iCP10m & "的本所期限與法定期限將設定為相同！", vbInformation
            MsgBox "有" & rsA1.Fields("cpm03") & "未發文，" & vbCrLf & _
                   iCP10m & "的本所期限將設定為相同，承辦期限 = 本所期限往前-5個工作天！", vbInformation
            iCP06 = TransDate("" & rsA1.Fields("cp06"), 1)
            'Modify By Sindy 2021/8/18 淑華說,不應該有法限,所以Mark
'            iCP07 = TransDate("" & rsA1.Fields("cp07"), 1) '
            iCP07 = ""
            '承辦期限 = 本所期限往前-5個工作天
            iCp48 = TransDate(CompWorkDay(5, CompDate(2, -1, DBDATE(iCP06)), 1), 1)
            m203CP48 = iCp48
            '2021/8/18 END
         End If
      End If
      
      'Add By Sindy 2021/8/18
      If m203CP48 = "" Then
         strA1 = "select cp06,cp07,cpm03 from caseprogress,casepropertymap" & _
                " where cp01='" & iPA(1) & "' and cp02='" & iPA(2) & "' and cp03='" & iPA(3) & "' and cp04='" & iPA(4) & "'" & _
                " and cp10 in('201','209','235','210') and cp27 is not null and cp57 is null and cp06>0 and cpm01(+)=cp01 and cpm02(+)=cp10"
         inA = 1
         Set rsA1 = ClsLawReadRstMsg(inA, strA1)
         If inA = 1 Then
            'B. 若新案翻譯、檢視中說、核對中說格式、製作中說 已發文：照原設定的期限規則設定
            '主動修正的承辦期限為原來設定的工作天數
            '          本所期限 = 承辦期限 + 再加5個工作天
            MsgBox "有" & rsA1.Fields("cpm03") & "已發文，" & vbCrLf & _
                   iCP10m & "的承辦期限為原來設定的工作天數，本所期限 = 承辦期限+再加5個工作天！", vbInformation
            iCP06 = TransDate(PUB_GetFCPOurDeadline(DBDATE(iCp48), , , , "N"), 1)
            '2021/8/18 END
            m203CP48 = iCp48
         End If
      End If
      '2021/8/18 END
      
      If m203CP48 <> "" Then PUB_CheckFCPshowMsg = True
   End If
   
    If iCP27 = "" Then
       '判斷未分案確認(有些程序會在收文時自動分案)
       If iCP122 = "" Then Call PUB_GetFCPsetCP48(bolRun, iPA, iCP27, iCP10, iCP14, iCP122, iCP06, iCP07, iCp48, iEP06, nCMB, iCP05)

       '告代分案無承辦人時, 檢查若新案已發文提醒是否重新計算承辦期限
       If iPA(1) = "FCP" And iCP10 = "901" And iCP14 = "" Then
          strA1 = "select cp06 from caseprogress where cp01='" & iPA(1) & "' and cp02='" & iPA(2) & "'" & _
                  " and cp03='" & iPA(3) & "' and cp04='" & iPA(4) & "' and cp31='Y' and cp27>0"
          inA = 1
          Set rsA1 = ClsLawReadRstMsg(inA, strA1)
          If inA = 1 Then
             If MsgBox("新案已發文，是否重算【" & iCP10m & "】承辦期限？", vbYesNo) = vbYes Then
               '承辦期限 = 告代收文日起+工作天數
               iCp48 = TransDate(Pub_GetHandleDay("FCP", "000", iCP10, DBDATE(iCP05), DBDATE(iCP06)), 1)
               'Add By Sindy 2021/8/11 承辦期限＋5個工作天為本所期限
               iCP06 = TransDate(PUB_GetFCPOurDeadline(DBDATE(iCp48), , , , "N"), 1)
               '2021/8/11 END
               PUB_CheckFCPshowMsg = True
             End If
          End If
       End If

       '主動修正、修正分案檢查若新案未發文提醒是否更改承辦期限與提申期限一致
       'Modify By Sindy 2025/7/16 增加告代 +,901
       If iPA(1) = "FCP" And InStr("203,204,901", iCP10) > 0 And iCP10.Text <> "" Then
          'Modify By Sindy 2025/7/16 +,cp141,cp142,cp164
          strA1 = "select cp06,cp141,cp142,cp164 from caseprogress where cp01='" & iPA(1) & "' and cp02='" & iPA(2) & "' and cp03='" & iPA(3) & "' and cp04='" & iPA(4) & "' and cp31='Y' and cp27 is null and cp06>0"
          inA = 1
          Set rsA1 = ClsLawReadRstMsg(inA, strA1)
          If inA = 1 Then
            'Modify By Sindy 2021/9/8 亭妙:可能是因為以前主動修正都只有掛承辦期限的原因，所以控管方式變成用主動修正的承辦期限來控管
            '(調整成)→ 只要新案那道未上發文，皆在分案的時候彈訊息詢問。
            If iCP10 = 主動修正 Then
               If MsgBox("新案未發文，【主動修正】是否為提申前主動修正？", vbYesNo) = vbYes Then
                  'Y:是（提申前主動修正）→設定更新本所期限與新案一致 →進度備註: 分案-提申前主動修正
                  If InStr(iCP64, "分案-提申前主動修正") = 0 Then
                     iCP64 = "分案-提申前主動修正;" & iCP64
                  End If
                  'Add By Sindy 2025/7/16 若有指定送件日當天or之前(排除指定送件日之後)
                  '                       本所期限若超過指定送件日, 則以指定送件日當本所期限
                  If "" & rsA1.Fields("cp141") = "3" And "" & rsA1.Fields("cp164") <> "3" And _
                     Val("" & rsA1.Fields("cp06")) > Val("" & rsA1.Fields("cp142")) Then
                     iCP06 = TransDate(rsA1.Fields("cp142"), 1)
                  Else
                  '2025/7/16 END
                     'Modify By Sindy 2021/8/18
                     '主動修正的本所期限 = 提申的本所期限
                     iCP06 = TransDate(rsA1.Fields(0), 1)
                  End If
                  '承辦期限 = 本所期限往前-5個工作天
                  iCp48 = TransDate(CompWorkDay(5, CompDate(2, -1, DBDATE(iCP06)), 1), 1)
                  '2021/8/18 END
                  PUB_CheckFCPshowMsg = True
               'Add By Sindy 2021/8/11
               '點選否, 表示提申後修正:
               '主動修正的承辦期限及本所期限暫先空白(等新申請案發文時會有其控制)
               Else
                  'N:否（提申後主動修正）→進度備註: 分案-提申後主動修正
                  '→ 待新案發文後，update主動修正與中說那道本所期限一致
                  If InStr(iCP64, "分案-提申後主動修正") = 0 Then
                     iCP64 = "分案-提申後主動修正;" & iCP64
                  End If
                  iCp48 = ""
                  iCP06 = ""
                  PUB_CheckFCPshowMsg = True
               '2021/8/11 END
               End If
            
            'Add By Sindy 2025/7/16 增加告代
            ElseIf iCP10 = 告知代理人 Then
               If MsgBox("新案未發文，【告代】是否為提申前告代？", vbYesNo) = vbYes Then
                  'Y:是（提申前告代）
                  If InStr(iCP64, "分案-提申前告代") = 0 Then
                     iCP64 = "分案-提申前告代;" & iCP64
                  End If
                  '承辦期限 = 告代收文日起+工作天數
                  iCp48 = TransDate(Pub_GetHandleDay("FCP", "000", iCP10, DBDATE(iCP05), DBDATE(iCP06)), 1)
                  '承辦期限＋5個工作天為本所期限
                  iCP06 = TransDate(PUB_GetFCPOurDeadline(DBDATE(iCp48), , , , "N"), 1)
                  '若有指定送件日當天or之前(排除指定送件日之後)
                  '本所期限若超過指定送件日, 則以指定送件日當本所期限
                  If "" & rsA1.Fields("cp141") = "3" And "" & rsA1.Fields("cp164") <> "3" And _
                     Val(DBDATE(iCP06)) > Val("" & rsA1.Fields("cp142")) Then
                     iCP06 = TransDate(rsA1.Fields("cp142"), 1)
                  Else
                     '本所期限若超過提申所限, 則以提申所限當本所期限
                     If Val(DBDATE(iCP06)) > Val("" & rsA1.Fields("cp06")) Then
                        iCP06 = TransDate(rsA1.Fields("cp06"), 1)
                     End If
                  End If
                  PUB_CheckFCPshowMsg = True
               '點選否, 表示提申後告代:
               Else
                  'N:否（提申後告代）
                  If InStr(iCP64, "分案-提申後告代") = 0 Then
                     iCP64 = "分案-提申後告代;" & iCP64
                  End If
                  iCp48 = ""
                  iCP06 = ""
                  PUB_CheckFCPshowMsg = True
               End If
               '2025/7/16 END
            '修正
            Else
            '2021/9/8 END
               '無承辦期限, 或承辦期限大於所限
               If iCp48 = "" Or Val(DBDATE(iCp48)) > Val("" & rsA1.Fields(0)) Then
                  'If MsgBox("新案未發文，是否更新【" & iCP10m & "】承辦期限與提申期限一致？", vbYesNo) = vbYes Then
                  If MsgBox("新案未發文，是否更新【" & iCP10m & "】本所期限與提申期限一致？", vbYesNo) = vbYes Then
                  '   iCp48 = TransDate(rsA1.Fields(0), 1)
                  '   'Add By Sindy 2021/8/11 承辦期限＋5個工作天為本所期限
                  '   iCP06 = TransDate(PUB_GetFCPOurDeadline(DBDATE(iCp48), , , , "N"), 1)
                  '   '2021/8/11 END
                     'Modify By Sindy 2021/8/18
                     iCP06 = TransDate(rsA1.Fields(0), 1)
                     '承辦期限 = 本所期限往前-5個工作天
                     iCp48 = TransDate(CompWorkDay(5, CompDate(2, -1, DBDATE(iCP06)), 1), 1)
                     '2021/8/18 END
                     'Added by Lydia 2023/05/09
                     If InStr(iCP64, "分案-所限與提申期限一致(" & iCP06 & ")") = 0 Then
                       iCP64 = "分案-所限與提申期限一致(" & iCP06 & ");" & iCP64
                     End If
                     'end 2023/05/09
                     PUB_CheckFCPshowMsg = True
                  End If
               End If
            End If
          End If
       End If
    End If
    '新案翻譯無期限
    If iCP10 = "201" And iCP07 = "" Then
       strA1 = "select cp27 from caseprogress where cp01='" & iPA(1) & "'" & _
          " and cp02='" & iPA(2) & "' and cp03='" & iPA(3) & "' and cp04='" & iPA(4) & "'" & _
          " and cp10='101' and cp27>0"
       inA = 1
       Set rsA1 = ClsLawReadRstMsg(inA, strA1)
       If inA = 1 Then
          MsgBox "翻譯無期限！"
          '法限=申請案發文日+4個月
          strA2 = CompDate(1, 4, rsA1(0))
          iCP07 = TransDate(strA2, 1)
          '所限=法限-4天
          strA3 = CompDate(2, -4, strA2)
          iCP06 = TransDate(strA3, 1)
          iCP06.Tag = iCP06.Text
          PUB_CheckFCPshowMsg = True
       End If
    End If
End Function

'Added by Lydia 2016/06/20 外專分案(工程師主管分案)-檢查承辦人或核稿人
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Function PUB_CheckFCPtxtValidate(ByRef iPA() As String, ByRef tCP14 As Object, ByRef tCP10 As Object, ByRef iEP04 As Object, ByRef tCP07 As Object, Optional ByRef bolCancel As Boolean, Optional ByRef m203CP As String, Optional ByVal iCP27 As String) As Boolean
'm203CP 實審分案時若有主動修正(203,206)未發文
Dim intA As Integer
Dim strA1 As String
Dim rsAD As New ADODB.Recordset

PUB_CheckFCPtxtValidate = False

   If Left(tCP14.Text, 1) = "F" And InStr("201,927,236", tCP10.Text) = 0 And tCP10.Text <> "" Then
      MsgBox "只有當案件性質為<翻譯>時承辦人才可輸入外譯編號！"
      bolCancel = True
      Exit Function
   End If
   '設計的201翻譯、209檢視中說、210製作中說、235核對中說格式 不掛核稿
   If iPA(8) = "3" And InStr("201,209,210,235", tCP10.Text) > 0 And tCP10.Text <> "" Then
       If iEP04 <> "" Then
           If MsgBox("設計的翻譯、檢視中說、核對中說格式、製作中說不掛核稿人，是否清空核稿人？", vbQuestion + vbYesNo) = vbYes Then
               iEP04 = ""
           Else
               Exit Function
           End If
       End If
   End If
   '案件性質改為非新案翻譯時核稿人檢查 Ex.FCP-51295
   If iEP04 <> "" And InStr("209,210,235", tCP10.Text) > 0 And tCP10.Text <> "" Then
      If MsgBox("檢視中說、核對中說格式、製作中說不掛核稿人，是否清空核稿人？", vbQuestion + vbYesNo) = vbYes Then
         iEP04 = ""
      Else
         Exit Function
      End If
   End If
   If tCP10.Text = "201" And tCP14.Text <> "" And tCP14.Text < "F" Then
      If MsgBox("承辦人為外專工程師之員工編號是否要繼續?", vbYesNo + vbDefaultButton2) = vbNo Then
         bolCancel = True
         Exit Function
      End If
   End If
   
   '實審分案時若有主動修正未發文時提醒更新期限
   m203CP = ""
   If tCP10 = "416" And tCP07 <> "" And iPA(10) <> "" And iCP27 = "" Then
      '判斷非來函通知的補充說明206
      'Modified by Lydia 2018/11/12 增加判斷中說已發文才更新主動修正的期限(ex.FCP-59553)
      'strA1 = "select cp09,cpm03 from caseprogress,nextprogress a,casepropertymap where cp01='" & iPA(1) & "' and cp02='" & iPA(2) & "' and cp03='" & iPA(3) & "' and cp04='" & iPA(4) & "' and cp10 in ('203','206') and cp27 is null and cp57 is null" & _
         " and np02(+)=cp43 and np07(+)=cp10 and np09 is null and cpm01(+)=cp01 and cpm02(+)=cp10"
      'Modify By Sindy 2023/11/17 分割案提申後之收文 "實審+主動修正"之分案，請新增於分案"實審" 時將其期限update"主動修正"之所限及法限
      strA1 = "select cp09,cpm03 from caseprogress,nextprogress a,casepropertymap where cp01='" & iPA(1) & "' and cp02='" & iPA(2) & "' and cp03='" & iPA(3) & "' and cp04='" & iPA(4) & "' and cp10 in ('203','206') and cp27 is null and cp57 is null" & _
                  " and (cp01,cp02,cp03,cp04) in (select cp01,cp02,cp03,cp04 from caseprogress where cp01='" & iPA(1) & "' and cp02='" & iPA(2) & "' and cp03='" & iPA(3) & "' and cp04='" & iPA(4) & "' and cp10 in('201','209','235','210','307') and cp158 > 0 ) " & _
                  " and np02(+)=cp43 and np07(+)=cp10 and np09 is null and cpm01(+)=cp01 and cpm02(+)=cp10"
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strA1)
      If intA = 1 Then
         m203CP = rsAD.Fields(0)
         MsgBox "有收文" & rsAD.Fields(1) & "，期限將改與實審相同！", vbInformation
      End If
   End If
   
PUB_CheckFCPtxtValidate = True

End Function

'Added by Lydia 2016/06/20 外專分案(工程師主管分案)-判斷承辦人或核稿人資料
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Function PUB_FCPGetCP14EP04(ByVal iTyp As String, ByRef iPA() As String, ByRef iText As Object, ByRef iName As Object, Optional ByRef bolCancel As Boolean) As Boolean
Dim iClass As String
Dim strName As String
Dim strDept As String
   iName.Caption = ""
   If iText.Text = "" Then
      iName.Caption = ""
   Else
      Select Case iPA(1)
          Case "PS", "FG", "CPS": iClass = iPA(79)
          Case Else: iClass = iPA(150)
      End Select
        If ClsPDGetStaff(iText, strName) Then
            '林信昌因分組故自動帶與案件組別的編號
            If iTyp <> "" Then '判斷承辦人或核稿人
                If InStr(strName, "林信昌") > 0 Then
                   Select Case iClass
                      Case "1"
                         If Left(iText, 1) = "6" Then iText = "68091"
                         If Left(iText, 1) = "F" Then iText = "F5644"
                      Case "2"
                         If Left(iText, 1) = "6" Then iText = "68092"
                         If Left(iText, 1) = "F" Then iText = "F5645"
                      Case Else
                         If Left(iText, 1) = "6" Then iText = "68007"
                         If Left(iText, 1) = "F" Then iText = "F5162"
                   End Select
                   If ClsPDGetStaff(iText, strName) Then
                   End If
                End If
            End If
            iName = strName
            strDept = GetStaffDepartment(iText)
            PUB_FCPGetCP14EP04 = True
            
            '核稿人控制只能為外專工程師
            If iTyp = "EP04" And InStr("F21,F52,F81", strDept) = 0 Then
               MsgBox "核稿人僅能輸外專工程師！"
               bolCancel = True
            End If
        Else
            iName = ""
            bolCancel = True
        End If
   End If
End Function

'Added By Lydia 2016/06/20 外專分案(工程師主管分案)-本所期限輸入約定期限 或 加打指定日期 計算承辦期限
'strLimitDT :約定期限
'tCP142     :指定送件日
'Modify By Sindy 2021/10/20 + ByVal strCP164 As String, ByVal strCP01 As String
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
'Modify By Sindy 2024/12/19 + , ByVal strCP09 As String
Public Sub PUB_GetFCPsetCP48Limit(ByRef strLimitDT As String, ByRef tCP14 As Object, _
   ByRef tCP06 As Object, ByRef tCP48 As Object, ByRef tCP142 As Object, ByRef tCP10 As Object, _
   ByVal strCP164 As String, ByVal strCP01 As String, ByVal strCP09 As String)
Dim StrStr1 As String
Dim intS As Integer
Dim AdoRs As New ADODB.Recordset
Dim strCPM34 As String
   
   'Add By Sindy 2021/10/20
   StrStr1 = "select cpm34 from casepropertymap where cpm01='" & strCP01 & "' and cpm02='" & tCP10.Text & "'"
   intS = 1
   Set AdoRs = ClsLawReadRstMsg(intS, StrStr1)
   If intS = 1 Then
      strCPM34 = "" & AdoRs.Fields(0)
   End If
   '2021/10/20 END
   
   If tCP14 <> "" Then '有輸入承辦人
      'Modify By Sindy 2021/8/30 淑華覺得不適用,先Mark
      'Modify By Sindy 2021/9/2 淑華:改排除201.新案翻譯
      'Modify By Sindy 2021/11/5 AEP(加速審查422)、PPH(高速審查431)分案時，本所及承辦期限應為"空"，待輸入通知實審日再依原設定規則掛期限
      'Modified by Morgan 2024/11/14 +447再審查加速審查
      If tCP10 <> 翻譯 And tCP10 <> "422" And tCP10 <> "431" And tCP10 <> "447" Then
      '2021/9/2 END
         StrStr1 = "SELECT st01,st15,st52 FROM staff WHERE st01='" & tCP14 & "' and substr(st01,1,1)<>'F'" & _
                     " Union" & _
                     " SELECT st01,st15,st52 FROM staff WHERE st26 in(" & _
                     " SELECT st26 FROM staff WHERE st01='" & tCP14 & "' and substr(st01,1,1)='F'" & _
                     " and st26 is not null)" & _
                     " and substr(st01,1,1)<>'F'"
         intS = 1
         Set AdoRs = ClsLawReadRstMsg(intS, StrStr1)
         If intS = 0 Then
            StrStr1 = "SELECT st01,st15,st52 FROM staff WHERE st01='" & Pub_GetSpecMan("M") & "'"
            intS = 1
            Set AdoRs = ClsLawReadRstMsg(intS, StrStr1)
         End If
         If intS = 1 Then
            'Modify By Sindy 2021/10/20 Mark
'            If "" & AdoRs.Fields("st15") = "F21" Or "" & AdoRs.Fields("st15") = "F22" Then
'               If "" & AdoRs.Fields("st15") = "F21" Then
'                  '有指定送件日:
'                  '當承辦人為工程師,其承辦期限更新為指定日期前4個工作天
'                  If tCP142.Text <> "" Then
'                     strLimitDT = DBDATE(tCP142.Text)
'                     tCP48 = CompWorkDay(5, DBDATE(tCP142), 1) - 19110000
'                  '當承辦人為工程師,其承辦期限更新為本所前4個工作天
'                  Else
'                     strLimitDT = DBDATE(tCP06.Text)
'                     tCP48 = CompWorkDay(5, DBDATE(tCP06), 1) - 19110000
'                  End If
'               ElseIf "" & AdoRs.Fields("st15") = "F22" Then
'                  '有指定送件日:
'                  '當承辦人為程序同仁,其承辦期限更新為指定日期
'                  If tCP142.Text <> "" Then
'                     strLimitDT = DBDATE(tCP142.Text)
'                     tCP48 = tCP142
'                  '當承辦人為程序同仁,其承辦期限更新為本所
'                  Else
'                     strLimitDT = DBDATE(tCP06.Text)
'                     tCP48 = tCP06
'                  End If
'               End If
            'Modify By Sindy 2021/10/20
'            '只有智慧局期限,承辦人=程序時; 指定日期方式為當天或之後, 才要更新承辦期限=指定送件日
'            If strCPM34 = "Y" And
            '程序
            If "" & AdoRs.Fields("st15") = "F22" And Val(tCP142.Text) > 0 Then
               'Modify By Sindy 2021/10/27 再調整如下
               '有指定送件日,當天或之後:承辦期限=指定送件日
               If strCP164 = "1" Or strCP164 = "3" Then
                  tCP48 = tCP142.Text
               '之前
               Else
                  '指定送件日<原預設承辦期限→承辦期限更新為指定送件日當天
                  If Val(tCP142.Text) < Val(tCP48.Text) Then
                     tCP48 = tCP142.Text
                  Else
                  '指定送件日>=原預設承辦期限→維持原承辦期限
                  End If
               End If
               '2021/10/27 END
               
            '非程序
            ElseIf "" & AdoRs.Fields("st15") <> "F22" Then
               '非智慧局期限案件並且承辦人<>F22.程序人員，在輸指定送件日(當天+之前+之後)時，請同時更新本所期限，如下：
               '有指定送件日<=本所期限：指定日期方式為之前或當天;更改本所期限=指定送件日
               '有指定送件日> 本所期限：指定日期方式為當天或之後;更改本所期限=指定送件日
               If strCPM34 = "N" And Val(tCP142.Text) > 0 And Val(tCP06.Text) > 0 Then
                  If (Val(tCP142.Text) <= Val(tCP06.Text) And (strCP164 = "1" Or strCP164 = "2")) Or _
                     (Val(tCP142.Text) > Val(tCP06.Text) And (strCP164 = "1" Or strCP164 = "3")) Then
                     tCP06.Text = tCP142.Text
                  End If
               End If
               
               'Add By Sindy 2021/10/27
               If Val(tCP142.Text) > 0 Then
                  '指定送件日<原預設承辦期限
                  If Val(tCP142.Text) < Val(tCP48.Text) Then
                     '之後:維持原承辦期限
                     '當天或之前: 承辦期限更新為指定送件日-5工作天
                     If strCP164 = "1" Or strCP164 = "2" Then
                        tCP48 = CompWorkDay(6, DBDATE(tCP142.Text), 1) - 19110000
                     End If
                  '指定送件日>=原預設承辦期限
                  Else
                     '之後:承辦期限更新為指定送件日-5工作天
                     '當天:維持原承辦期限
                     '之前:指定送件日-5工作天後若>=原預設承辦期限→維持原承辦期限
                     '                                        反之,更新承辦期限為指定送件日-5工作天
                     If strCP164 = "3" Then
                        tCP48 = CompWorkDay(6, DBDATE(tCP142.Text), 1) - 19110000
                     ElseIf strCP164 = "2" Then
                        If Val(CompWorkDay(6, DBDATE(tCP142.Text), 1) - 19110000) < Val(tCP48.Text) Then
                           tCP48.Text = Val(CompWorkDay(6, DBDATE(tCP142.Text), 1) - 19110000)
                        End If
                     End If
                  End If
               End If
               '2021/10/27 END
            End If
            '2021/10/20 END
            
            'Add By Sindy 2024/12/19 會稿計算本所期限的規則
            '分案，當案件性質為「會稿924」時，當有設定指定期限 之前/當天/之後 時，皆更新本所期限=指定期限
            '請排除會稿的相關收文號為中說性質「201新案翻譯、209檢視中說、235核對中說格式、210製作中說」
            If tCP10 = "924" And Val(tCP142.Text) > 0 Then
               StrStr1 = "select * from caseprogress c1,caseprogress c2" & _
                         " where c1.cp09='" & strCP09 & "' and c1.cp43 is not null" & _
                         " and c1.cp43=c2.cp09 and c2.cp10 in('201','209','235','210')"
               intS = 1
               Set AdoRs = ClsLawReadRstMsg(intS, StrStr1)
               If intS = 0 Then
                  tCP06.Text = tCP142.Text
               End If
            End If
            '2024/12/19 END
            
            '承辦期限小於等於系統日期時,承辦期限等於系統日
            If tCP48 <> "" And Val(tCP48) <= Val(strSrvDate(2)) Then
               tCP48 = strSrvDate(2)
            End If
'            End If
            
            If tCP142.Text <> "" Then
               strLimitDT = DBDATE(tCP142.Text)
            ElseIf tCP06.Text <> "" Then
               strLimitDT = DBDATE(tCP06.Text)
            End If
         End If
      End If
   End If
End Sub

'Added by Lydia 2016/06/21 外專分案(工程師主管分案)-承辦人欄位控制,並且影響核稿人欄位
'Modified by Morgan 2022/2/7 核稿人改Optional並取消回寫
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Function PUB_FCPCheckCP14(ByRef tPA() As String, ByRef tCP10 As Object, ByRef tCP14 As Object, Optional ByRef tEP04 As Object, Optional ByRef tEP04n As Object) As Boolean
Dim intT As Integer
Dim rsT1 As New ADODB.Recordset
Dim stDept As String
Dim StrStr As String
Dim m_Fagent As String

   If tCP14.Text = "" Then Exit Function
   
   stDept = GetStaffDepartment(tCP14)
   '209,210,235 承辦只為工程師 F21
   If InStr("209,210,235", tCP10.Text) > 0 And tCP10.Text <> "" Then
      If tCP14.Text <> "" And tCP14.Text <> tCP14.Tag Then
         If stDept <> "F21" And stDept <> "F81" Then
            MsgBox "該案件性質的承辦人只可為工程師!!!"
            Exit Function
         End If
      End If
   End If
   '非設計案的翻譯(與抓核稿人的案件性質相同)
   If tPA(8) <> "3" And InStr(FCPHaveEP04, tCP10.Text) > 0 And tCP10.Text <> "" Then
       'Modified by Morgan 2022/2/7 取消,分案改為只顯示,核稿人只能由完稿輸入修改--Sharon
       'If tCP14.Text <> tCP14.Tag Or tEP04.Text <> tEP04.Tag Then
       If tCP14.Text <> tCP14.Tag Then
       'end 2022/2/7
       
           Select Case stDept
               Case "F22"
                   MsgBox "該案件性質的承辦人不可為程序!!!"
                   Exit Function
               Case "F51"
                   'Removed by Morgan 2022/2/7 取消,分案改為只顯示,核稿人只能由完稿輸入修改--Sharon
                   'tEP04.Text = ""
                   'tEP04n.Caption = ""
                   'end 2022/2/7
               Case Else
                  '新案翻譯 201 承辦人為國外部工程師的才要預設核稿人 ; 所內員工不用考慮對照(新進同仁無外譯編號)
                  'Mark by Lydia 2019/05/02 不用預設核稿人; 參考107/12/7 新案翻譯之發文,若承辦人與核稿人為同一人,彈訊息"承辦人與核稿人為同一人,不可發文",且不得發文。
                  'If tCP10.Text = "201" Then
                  '   StrStr = "select st01 from staff where st01='" & tCP14 & "' and ST15='F21' and st04='1'" & _
                        " union select st01 from staff_idmap,staff where sim02='" & tCP14 & "'" & _
                        " and st01(+)=sim01 and ST15='F21' and st04='1'"
                        
                  '   intT = 1
                  '   Set rsT1 = ClsLawReadRstMsg(intT, StrStr)
                  '   If intT = 1 Then
                  '      tEP04.Text = tCP14.Text
                  '      tEP04n.Caption = GetStaffName(tCP14.Text)
                  '   Else
                        'Removed by Morgan 2022/2/7 取消,分案改為只顯示,核稿人只能由完稿輸入修改--Sharon
                        'tEP04.Text = ""
                        'tEP04n.Caption = ""
                        'end 2022/2/7
                  '   End If
                  'End If
                  'end 2019/05/02
               End Select
       End If
   Else
       'Removed by Morgan 2022/2/7 取消,分案改為只顯示,核稿人只能由完稿輸入修改--Sharon
       'tEP04.Text = ""
       'tEP04n.Caption = ""
       'end 2022/2/7
       
       '年費依有無年費代理人檢查承辦人且須為該國管制人
       If tCP10.Text = "605" Then
          '改比照期限管制規則(同信函收件人)
          m_Fagent = PUB_GetReceiver("" & ChgSQL(tPA(1)), "" & ChgSQL(tPA(2)), "" & ChgSQL(tPA(3)), "" & ChgSQL(tPA(4)), "605", "1")
          'Modified by Lydia 2017/02/13 +FMP管制人
          If strSrvDate(1) < FMP管制人啟用日 Then
              StrStr = "select na16 from fagent,nation where fa01='" & ChgSQL(Mid(m_Fagent, 1, 8)) & "' and fa02='" & ChgSQL(Mid(m_Fagent, 9, 1)) & "' and fa10=na01(+) "
          Else
              StrStr = "select na16,na79 from fagent,nation where fa01='" & ChgSQL(Mid(m_Fagent, 1, 8)) & "' and fa02='" & ChgSQL(Mid(m_Fagent, 9, 1)) & "' and fa10=na01(+) "
          End If
          'end 2017/02/13
          
          intT = 1
          Set rsT1 = ClsLawReadRstMsg(intT, StrStr)
          If intT = 1 Then
            If tCP14.Text <> "" And tCP14.Text <> "" & rsT1.Fields(0) Then
               'Modified by Lydia 2017/02/13 +FMP管制人
               'If MsgBox("年費的承辦人錯誤, 應為 " & rsT1.Fields(0) & GetStaffName(rsT1.Fields(0)) & "，是否確定要繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
               If strSrvDate(1) < FMP管制人啟用日 Then
                  StrStr = rsT1.Fields("na16") & GetStaffName(rsT1.Fields("na16"))
               Else
                    If (ChgSQL(tPA(1)) = "P" Or ChgSQL(tPA(1)) = "PS") And "" & rsT1.Fields("na79") <> "" Then
                        StrStr = rsT1.Fields("na79") & GetStaffName(rsT1.Fields("na79"))
                    Else
                        StrStr = rsT1.Fields("na16") & GetStaffName(rsT1.Fields("na16"))
                    End If
               End If
               If MsgBox("年費的承辦人錯誤, 應為 " & StrStr & "，是否確定要繼續？", vbYesNo + vbDefaultButton1) = vbNo Then
               'end 2017/02/13
                  Exit Function
               End If
            End If
         End If
       End If
   End If
   
   '重新核稿承辦人不可為原翻譯的核稿人
   If tCP10.Text = "229" Then
      If Left(tCP14, 1) = "F" Then
         StrStr = "select * from staff_idmap where sim02='" & tCP14 & "' and sim01=ep04"
      Else
         StrStr = "select * from staff_idmap where sim01='" & tCP14 & "' and sim02=ep04"
      End If
      StrStr = "select ep04 from caseprogress,engineerprogress" & _
         " where cp01='" & tPA(1) & "' and cp02='" & tPA(2) & "'" & _
         " and cp03='" & tPA(3) & "' and cp04='" & tPA(4) & "' and cp10='201' and ep02(+)=cp09" & _
         " and (ep04='" & tCP14 & "' or exists(" & StrStr & "))"
      intT = 1
      Set rsT1 = ClsLawReadRstMsg(intT, StrStr)
      If intT = 1 Then
         MsgBox "重新核稿承辦人不可為原翻譯的核稿人!!", vbExclamation
         Exit Function
      End If
   End If
   
   PUB_FCPCheckCP14 = True
End Function

'Added by Lydia 2024/03/27 一併修改相關收文號之承辦人
Public Sub PUB_SaveFCPcp14Ex(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pCP09 As String, ByVal pCP10 As String, pNowNo As String)
Dim strSqlEx As String
   

    '審查意見或核駁修改承辦人時一併修改相關收文號之告代承辦人 Ex.FCP-45516
    If pNowNo <> "" And (pCP10 = "1202" Or pCP10 = "1002" Or pCP10 = "1227") Then
       strSqlEx = "update caseprogress set cp14='" & pNowNo & "' where cp43='" & pCP09 & "' and cp10='901' and cp27 is null"
       cnnConnection.Execute strSqlEx
    End If
    '113/3/26 Wilison
    '關於"核對已准專利"及"核准"此兩道收文，工程師應是同一人
    '因此，請設定更改"核對已准專利"的工程師時，"核准"的工程師也同時修改為同一人
    '另外，更改"核對"的工程師時，"核准已准專利"的工程師也同時修改為同一人
    If pCP01 = "FCP" And pNowNo <> "" And (pCP10 = "1001" Or pCP10 = "926") Then
       If pCP10 = "1001" Then
         strSqlEx = "update caseprogress set cp14='" & pNowNo & "' where cp43='" & pCP09 & "' and cp10='926' and cp27 is null"
         cnnConnection.Execute strSqlEx
       ElseIf pCP10 = "926" Then '因為核准自動上發文日，所以不限制未發文
         strSqlEx = "update caseprogress set cp14='" & pNowNo & "' where cp09=(select cp43 from caseprogress where cp09='" & pCP09 & "') and cp10='1001' "
         cnnConnection.Execute strSqlEx
       End If
    End If
End Sub

'Added by Lydia 2016/06/21 外專分案(工程師主管分案)-與承辦人、核稿人相關欄位存檔
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
'Modified by Morgan 2024/5/21 sSubject EMail主旨，目前為訴願的補充說明分案通知用
Public Sub PUB_SaveFCPcp14(ByRef sPA() As String, ByRef tCP48 As Object, ByVal sCP09 As String, ByVal sCP10 As String, ByVal sCP14 As String, ByVal oldCP14 As String, _
                           ByVal sCP06 As String, ByVal sCP07 As String, ByVal sCP27 As String, ByVal sCP57 As String, ByVal sCP60 As String, ByVal sPA75 As String, _
                           ByVal sTeam As String, ByVal sEP04 As String, ByVal oldEP04 As String, ByRef tEP06 As Object, ByVal strLimitDT As String, _
                           ByVal str203CP09 As String, ByVal sCPM As String, ByVal sPrjName As String, Optional ByVal sSubject As String)
Dim strS1 As String, strS2 As String
Dim intS As Integer
Dim rsRD As New ADODB.Recordset
Dim strMid As String
Dim strEP08 As String
Dim strEP09 As String
Dim tmpTo As String, tmpSub As String, tmpCont As String
Dim tmpName As String
Dim tmpEmpMan  As String 'Added by Lydia 2020/02/10

    '審查意見或核駁修改承辦人時一併修改相關收文號之告代承辦人 Ex.FCP-45516
    'Modified by Lydia 2024/03/27 另外抽成一個模組
    'If sCP14 <> "" And (sCP10 = "1202" Or sCP10 = "1002" Or sCP10 = "1227") Then
    '   strS1 = "update caseprogress set cp14='" & sCP14 & "' where cp43='" & sCP09 & "' and cp10='901' and cp27 is null"
    '   cnnConnection.Execute strS1, intS
    'End If
    Call PUB_SaveFCPcp14Ex(sPA(1), sPA(2), sPA(3), sPA(4), sCP09, sCP10, sCP14)
    'end 2024/03/27

    '更新齊備日
    If tEP06.Locked = False Then
       strMid = ",EP06=" & CNULL(DBDATE(tEP06.Text), True)
    End If
    
    'Added by Lydia 2020/02/10 承辦人(含外翻編號F編號)為外專工程師，通知外專工程師主管
    strS1 = "SELECT st01,st15,st52 FROM staff WHERE st26 in(" & _
                " SELECT st26 FROM staff WHERE st01='" & sCP14 & "' and st26 is not null)" & _
                " and substr(st01,1,1)<>'F'"
    'Added by Lydia 2020/02/20 增加判斷工程師組別; 因為林信昌有三個組別
    tmpName = PUB_GetStaffST16(sCP14)
    If tmpName <> "" Then strS1 = strS1 & " and st16='" & tmpName & "' "
    'end 2020/02/20
    intS = 1
    Set rsRD = ClsLawReadRstMsg(intS, strS1)
    If intS = 1 Then
        If "" & rsRD.Fields("st15") = "F21" Then
            tmpEmpMan = PUB_GetFCPEngSup("" & rsRD.Fields("st01"))
        End If
    End If
    'end 2020/02/10
    
    'Modify By Sindy 2021/8/30 淑華覺得不適用,先Mark
'    '有約定期限或指定日期
'    If sCP27 = "" And sCP57 = "" Then
'       If strLimitDT <> "" And sCP10 = "201" Then
'          '新案翻譯的核稿期限設定為前7個工作天,不能早於完稿日
'          strEP08 = DBDATE(CompWorkDay(8, DBDATE(strLimitDT), 1))
'          strEP09 = DBDATE(strEP09) '原完稿日
'          If Val(strEP08) < Val(strEP09) Then
'             strEP08 = strEP09
'          End If
'          strMid = ",EP08=" & CNULL(strEP08, True)
'          'e核稿人及工程師主管
'          If sEP04 <> "" Then
'             tmpTo = sEP04 & ";"
'          End If
'
'          If Left(tmpTo, 1) = "F" Then 'Memo by Lydia 2020/02/10 核稿人為外翻編號F編號
'             strS1 = "SELECT st01,st15,st52 FROM staff WHERE st26 in(" & _
'                         " SELECT st26 FROM staff WHERE st01='" & Left(tmpTo, 5) & "' and st26 is not null)" & _
'                         " and substr(st01,1,1)<>'F'"
'             intS = 1
'             Set rsRd = ClsLawReadRstMsg(intS, strS1)
'             If intS = 1 Then
'                tmpTo = rsRd.Fields("st01") & ";"
'                'Added by Lydia 2020/02/10 核稿人為外專工程師
'                If "" & rsRd.Fields("st15") = "F21" Then
'                    strS2 = PUB_GetFCPEngSup("" & rsRd.Fields("st01"))
'                End If
'                'end 2020/02/10
'             End If
'          End If
'
'          'Modified by Lydia 2020/02/10
'          'If sTeam <> "" Then
'          If tmpEmpMan <> "" Then
'               tmpTo = tmpTo & tmpEmpMan & ";"
'          ElseIf strS2 <> "" Then
'               tmpTo = tmpTo & strS2 & ";"
'          ElseIf sTeam <> "" Then
'          'end 2020/02/10
'               strS2 = IIf(sTeam = "1", Pub_GetSpecMan("T"), IIf(sTeam = "2", Pub_GetSpecMan("R"), IIf(sTeam = "3", Pub_GetSpecMan("S"), Pub_GetSpecMan("T1"))))
'               If InStr(tmpTo, strS2) = 0 Then tmpTo = tmpTo & strS2
'          End If
'
'          tmpSub = "請儘速辦理中說(翻譯/核稿)"
'          tmpCont = "FCP" & sPA(2) & "因客戶催辦，指定於" & ChangeTStringToTDateString(strLimitDT) & IIf(sCP06 <> "" And DBDATE(sCP06) <> strLimitDT, "(本所期限：" & ChangeTStringToTDateString(sCP06) & ")", "") & vbCrLf & _
'                       "前呈送智慧局，請儘速辦理，謝謝。"
'          strS1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
'                   " VALUES ( '" & strUserNum & "','" & tmpTo & "',to_char(sysdate,'yyyymmdd')" & _
'                   ",to_char(sysdate,'hh24miss'),'" & tmpSub & "','" & tmpCont & "')"
'          cnnConnection.Execute strS1
'       End If
'    End If

    '更新核稿人
    strS1 = "Update EngineerProgress Set EP04='" & sEP04 & "'" & strMid & " Where EP02='" & sCP09 & "'"
    cnnConnection.Execute strS1

    '實審分案時若有主動修正(203,206)未發文時,一併更新期限
    If str203CP09 <> "" And sCP07 <> "" Then
       strS1 = "update caseprogress set cp06=" & DBDATE(sCP06) & ",cp07=" & DBDATE(sCP07) & " where cp09='" & str203CP09 & "'"
       cnnConnection.Execute strS1, intS
    End If

    '若實審分案時有主動修正未發文且承辦期限較早者更新為相同
    If sCP27 = "" And (sCP10 = "416" Or sCP10 = "203") And tCP48 <> "" Then '判斷申請案已送件
       '當主動修正後收文時,承辦期限更新至實審,若實審後收文則以實審承辦期限更新至主動修正-->以後收文者為準,不用比較大小
       strS1 = "select cp09,cp06 from caseprogress a where cp01='" & sPA(1) & "' and cp02='" & sPA(2) & "' and cp03='" & sPA(3) & "' and cp04='" & sPA(4) & "'" & _
          " and cp10='" & IIf(sCP10 = "203", "416", "203") & "' and cp27||cp57 is null" & _
          " and exists(select * from caseprogress b where b.cp01=a.cp01 and b.cp02=a.cp02 and b.cp03=a.cp03 and b.cp04=a.cp04 and b.cp10='101' and b.cp27>0)"
       intS = 1
       Set rsRD = ClsLawReadRstMsg(intS, strS1)
       If intS = 1 Then
          strS2 = DBDATE(tCP48)
          If Not IsNull(rsRD.Fields("cp06")) And Val("" & rsRD.Fields("cp06")) < strS2 Then
             strS2 = rsRD.Fields("cp06")
          End If
          strS1 = "update caseprogress set cp48=" & strS2 & " where cp09='" & rsRD("cp09") & "'"
          cnnConnection.Execute strS1, intS
       End If
    End If
    
    '若已開請款單則換承辦人或核稿人時發Mail通知相關人員；
    If sCP60 > "X" Then
       'Modified by Lydia 2019/10/17 本所案號+"-"
       'PUB_PointReAssignInform sPA(1) & sPA(2) & sPA(3) & sPA(4), sCP60, oldCP14, sCP14, oldEP04, sEP04
       PUB_PointReAssignInform sPA(1) & "-" & sPA(2) & IIf(sPA(3) & sPA(4) = "000", "", "-" & sPA(3) & "-" & sPA(4)), sCP60, oldCP14, sCP14, oldEP04, sEP04
    End If
      
    'Email通知
    tmpName = GetStaffName(sCP14, True)
    If sCP27 = "" And sCP57 = "" Then
       'Added by Morgan 2024/5/21
       
       'end 2024/5/21
    
       '達法定期限當天
       'Modified by Lydia 2017/03/08 增加"已過法定期限"通知
       'If sCP07 <> "" And strSrvDate(1) = DBDATE(sCP07) Then
       If sCP07 <> "" And strSrvDate(1) >= DBDATE(sCP07) Then
       'end 2017/03/08
          'Modified by Lydia 2017/02/13 +FMP管制人
          If (sPA(1) = "P" Or sPA(2) = "PS") And strSrvDate(1) > FMP管制人啟用日 Then
            strS1 = "SELECT s1.st15 aST15,s1.st52 aST52,na01,nvl(na79,na16) na16,s2.st01 bST01,s2.st52 bST52" & _
                  " FROM staff s1,nation,staff s2,fagent" & _
                  " WHERE s1.st01='" & sCP14 & "' and fa01='" & Left(ChangeCustomerL(sPA75), 8) & "' and fa02='" & Mid(ChangeCustomerL(sPA75), 9, 1) & "'" & _
                  " and na01(+)=fa10 and s2.st01(+)=nvl(na79,na16)"
          Else
            strS1 = "SELECT s1.st15 aST15,s1.st52 aST52,na01,na16,s2.st01 bST01,s2.st52 bST52" & _
                  " FROM staff s1,nation,staff s2,fagent" & _
                  " WHERE s1.st01='" & sCP14 & "' and fa01='" & Left(ChangeCustomerL(sPA75), 8) & "' and fa02='" & Mid(ChangeCustomerL(sPA75), 9, 1) & "'" & _
                  " and na01(+)=fa10 and s2.st01(+)=na16"
          End If
          'end 2017/02/13
          intS = 1
          Set rsRD = ClsLawReadRstMsg(intS, strS1)
          tmpTo = ""
          If intS = 1 Then
             '工程師
             If "" & rsRD.Fields("aST15") = "F21" Then
                tmpTo = sCP14 & ";"
                'Modified by Lydia 2020/02/10 承辦人(含外翻編號F編號)為外專工程師，通知外專工程師主管
                'If sTeam <> "" Then
                If tmpEmpMan <> "" Then
                   tmpTo = tmpTo & tmpEmpMan & ";"
                ElseIf sTeam <> "" Then
                'end 2020/02/10
                   strS2 = IIf(sTeam = "1", Pub_GetSpecMan("T"), IIf(sTeam = "2", Pub_GetSpecMan("R"), IIf(sTeam = "3", Pub_GetSpecMan("S"), Pub_GetSpecMan("T1"))))
                   If InStr(tmpTo, strS2) = 0 Then tmpTo = tmpTo & strS2 & ";"
                End If
                '加發管制人
                If "" & rsRD.Fields("na16") <> "" Then
                   If InStr(tmpTo, rsRD.Fields("na16")) = 0 Then tmpTo = tmpTo & rsRD.Fields("na16") & ";"
                   If "" & rsRD.Fields("bST52") <> "" Then
                      If InStr(tmpTo, rsRD.Fields("bST52")) = 0 Then tmpTo = tmpTo & rsRD.Fields("bST52") & ";"
                   End If
                End If
                If InStr(tmpTo, Pub_GetSpecMan("C")) = 0 Then tmpTo = tmpTo & Pub_GetSpecMan("C") 'Modify By Sindy 2022/8/11 N改為C
             '程序人員
             ElseIf "" & rsRD.Fields("aST15") = "F22" Then
                tmpTo = sCP14 & ";"
                '一級主管
                If "" & rsRD.Fields("aST52") <> "" Then
                   If InStr(tmpTo, rsRD.Fields("aST52")) = 0 Then tmpTo = tmpTo & rsRD.Fields("aST52") & ";"
                End If
                '分案程序主管
                If InStr(tmpTo, Pub_GetSpecMan("C")) = 0 Then tmpTo = tmpTo & Pub_GetSpecMan("C") 'Modify By Sindy 2022/8/11 N改為C
             End If
             If tmpTo <> "" Then
                'Modified by Lydia 2017/03/08 增加"已過法定期限"通知
                If strSrvDate(1) > DBDATE(sCP07) Then
                   tmpSub = "已過法定期限"
                Else
                   tmpSub = "已達法定當天"
                End If
                'end 2017/03/08
                tmpCont = "本所案號：" + sPA(1) & "-" & sPA(2) & "-" & sPA(3) & "-" & sPA(4) + vbCrLf + _
                          "案件名稱：" + sPrjName + vbCrLf + _
                          "案件性質：" + sCPM + vbCrLf + _
                          "本所期限：" + ChangeTStringToTDateString(sCP06) + vbCrLf + _
                          "法定期限：" + ChangeTStringToTDateString(sCP07) + vbCrLf + _
                          "承 辦 人：" + tmpName + vbCrLf
                strS1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                   " VALUES ( '" & strUserNum & "','" & tmpTo & "',to_char(sysdate,'yyyymmdd')" & _
                   ",to_char(sysdate,'hh24miss'),'" & tmpSub & "','" & tmpCont & "')"
                cnnConnection.Execute strS1
             End If
          End If
          
       '已達本所前二個工作天
       'Modified by Morgan 2024/5/21 訴願的補充說明分案通知也用相同的設定
       'ElseIf sCP06 <> "" And strSrvDate(1) >= CompWorkDay(3, DBDATE(sCP06), 1) Then
       ElseIf sCP06 <> "" And (strSrvDate(1) >= CompWorkDay(3, DBDATE(sCP06), 1) Or sSubject <> "") Then
          strS1 = "SELECT st15,st52 FROM staff WHERE st01='" & sCP14 & "'"
          intS = 1
          Set rsRD = ClsLawReadRstMsg(intS, strS1)
          tmpTo = ""
          If intS = 1 Then
             '工程師
             If "" & rsRD.Fields("st15") = "F21" Then
                '工程師本人
                tmpTo = sCP14 & ";"
                'Modified by Lydia 2020/02/10 承辦人(含外翻編號F編號)為外專工程師，通知外專工程師主管
                'If sTeam = "" Then
                If tmpEmpMan <> "" Then
                   tmpTo = tmpTo & tmpEmpMan & ";"
                ElseIf sTeam = "" Then
                'end 2020/02/10
                   '分案程序主管
                   If InStr(tmpTo, Pub_GetSpecMan("C")) = 0 Then tmpTo = tmpTo & Pub_GetSpecMan("C") 'Modify By Sindy 2022/8/11 N改為C
                Else
                   '工程師主管
                   strS2 = IIf(sTeam = "1", Pub_GetSpecMan("T"), IIf(sTeam = "2", Pub_GetSpecMan("R"), IIf(sTeam = "3", Pub_GetSpecMan("S"), Pub_GetSpecMan("T1"))))
                   If InStr(tmpTo, strS2) = 0 Then tmpTo = tmpTo & strS2
                End If
             End If
             If tmpTo <> "" Then
                'Added by Morgan 2024/5/21
                If sSubject <> "" Then
                  tmpSub = sSubject
                Else
                'end 2024/5/21
                  tmpSub = "已達本所前二個工作天"
                End If
                tmpCont = "本所案號：" + sPA(1) & "-" & sPA(2) & "-" & sPA(3) & "-" & sPA(4) + vbCrLf + _
                          "案件名稱：" + sPrjName + vbCrLf + _
                          "案件性質：" + sCPM + vbCrLf + _
                          "本所期限：" + ChangeTStringToTDateString(sCP06) + vbCrLf + _
                          "法定期限：" + ChangeTStringToTDateString(sCP07) + vbCrLf + _
                          "承 辦 人：" + tmpName + vbCrLf
                strS1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                   " VALUES ( '" & strUserNum & "','" & tmpTo & "',to_char(sysdate,'yyyymmdd')" & _
                   ",to_char(sysdate,'hh24miss'),'" & tmpSub & "','" & tmpCont & "')"
                cnnConnection.Execute strS1
             End If
          End If
          
       '達本所期限當天
       ElseIf sCP06 <> "" And strSrvDate(1) = DBDATE(sCP06) Then
          strS1 = "SELECT st15,st52 FROM staff WHERE st01='" & sCP14 & "'"
          intS = 1
          Set rsRD = ClsLawReadRstMsg(intS, strS1)
          tmpTo = ""
          If intS = 1 Then
             '程序人員
             If "" & rsRD.Fields("st15") = "F22" Then
                tmpTo = sCP14 & ";"
                '一級主管
                If "" & rsRD.Fields("st52") <> "" Then
                   If InStr(tmpTo, rsRD.Fields("st52")) = 0 Then tmpTo = tmpTo & rsRD.Fields("st52") & ";"
                Else
                   '分案程序主管
                   If InStr(tmpTo, Pub_GetSpecMan("C")) = 0 Then tmpTo = tmpTo & Pub_GetSpecMan("C") 'Modify By Sindy 2022/8/11 N改為C
                End If
             End If
             If tmpTo <> "" Then
                tmpSub = "已達本所當天"
                tmpCont = "本所案號：" + sPA(1) & "-" & sPA(2) & "-" & sPA(3) & "-" & sPA(4) + vbCrLf + _
                          "案件名稱：" + sPrjName + vbCrLf + _
                          "案件性質：" + sCPM + vbCrLf + _
                          "本所期限：" + ChangeTStringToTDateString(sCP06) + vbCrLf + _
                          "法定期限：" + ChangeTStringToTDateString(sCP07) + vbCrLf + _
                          "承 辦 人：" + tmpName + vbCrLf
                strS1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                   " VALUES ( '" & strUserNum & "','" & tmpTo & "',to_char(sysdate,'yyyymmdd')" & _
                   ",to_char(sysdate,'hh24miss'),'" & tmpSub & "','" & tmpCont & "')"
                cnnConnection.Execute strS1
             End If
          End If
       End If
    End If
End Sub

'Added by Lydia 2016/06/21 外專分案(工程師主管分案)-核稿期限設定(原FCP-翻譯完稿輸入)
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Sub PUB_FCPsetEP08(ByVal sCP01 As String, ByVal sCP02 As String, ByVal sCP03 As String, ByVal sCP04 As String, ByVal sCP06 As String, ByVal sCP10 As String, ByVal sPA10 As String, ByVal sCP14 As String, ByVal sEP04 As String, ByRef tEP08T As Object, ByRef tEP09T As Object)
   Dim dtEP09 As Date, dtEP08 As Date, dtTmp1 As Date, dtTmp2 As Date
   Dim intE As Integer, strE As String
   Dim rRS As New ADODB.Recordset
   
   '若原先已有核稿期限時不再重算
   If tEP08T.Tag = "" Then
   
      tEP08T = ""
      dtEP09 = ChangeTStringToWDateString(tEP09T.Text)
      If sPA10 = "" Then
          MsgBox "尚未輸入申請日！", vbExclamation
          dtTmp1 = 0
      Else
          dtTmp1 = DateAdd("m", 6, ChangeTStringToWDateString(sPA10)) - 4
      End If
      Select Case sCP10
          Case "201" '翻譯
              'P案翻譯核稿期限=完稿日+5個工作天
              If sCP01 = "P" Then
                 dtTmp2 = CDate(ChangeWStringToWDateString(CompWorkDay(5, DBDATE(tEP09T.Text))))
              Else
                 '承辦人與核稿人不同(外翻):核稿承辦期限=完稿日+22個工作天
                 If sCP14 <> sEP04 Then
                     dtTmp2 = CDate(ChangeWStringToWDateString(CompWorkDay(22, DBDATE(tEP09T.Text))))
                 '承辦人與核稿人相同(內翻):核稿承辦期限=完稿日+10個工作天
                 Else
                     dtTmp2 = CDate(ChangeWStringToWDateString(CompWorkDay(10, DBDATE(tEP09T.Text))))
                 End If
              End If
          'Added by Lydia 2016/07/19
          Case "927" '其他翻譯
              '其他翻譯核稿期限=完稿日+3個工作天
              dtTmp2 = CDate(ChangeWStringToWDateString(CompWorkDay(3, DBDATE(tEP09T.Text))))
      End Select
           
      '核稿承辦期限<=申請日+6個月-4天
      If dtTmp2 <> 0 Then
         'P案不必控制
         'Modifeid by Lydia 2016/07/19 + 其他翻譯
         If sCP01 = "P" Or sCP10 = "927" Then
            dtEP08 = dtTmp2
         Else
            If dtTmp1 = 0 Then
                dtEP08 = dtTmp2
            ElseIf DateDiff("d", dtTmp1, dtTmp2) > 0 Then
                dtEP08 = dtTmp1
            Else
                dtEP08 = dtTmp2
            End If
         End If
         
         '承辦期限不可大於本所期限
         If sCP06 <> "" Then
             If dtEP08 > ChangeWStringToWDateString(sCP06) Then
                dtEP08 = ChangeWStringToWDateString(sCP06)
             End If
         End If
         tEP08T = Format(Val(Format(dtEP08, "YYYYMMDD")) - 19110000)
      End If
      
   End If

   '若核稿期限大於會稿、寄中說的本所期限時提醒會更新
   If tEP08T <> "" Then
      strE = "select cp06 from caseprogress where cp01='" & sCP01 & "'" & _
         " and cp02='" & sCP02 & "' and cp03='" & sCP03 & "'" & _
         " and cp04='" & sCP04 & "' and cp10 in ('924','949') and cp57||cp27 is null and cp06<" & DBDATE(tEP08T)
      intE = 1
      Set rRS = ClsLawReadRstMsg(intE, strE)
      If intE = 1 Then
         MsgBox "將更新核稿期限為【會稿】或【寄中說】之本所期限！"
         tEP08T = TransDate(rRS(0), 1)
      End If
   End If
End Sub

'Added by Lydia 2021/01/06 Murgitroyd案：新案發文時一併設定中說核稿期限
Public Function PUB_FCPsetEP08M(ByVal sCP01 As String, ByVal sCP02 As String, ByVal sCP03 As String, ByVal sCP04 As String, ByVal sCP06 As String, ByVal sCP10 As String, ByVal sPA10 As String, ByRef dtEP08 As String, Optional ByVal bShowMsg As Boolean = True) As Boolean
Dim intE As Integer, strE As String
Dim rRS As New ADODB.Recordset
   
    dtEP08 = ""
    If sPA10 = "" Then
        MsgBox "尚未輸入申請日！", vbExclamation
        Exit Function
    End If
    '(新案翻譯)中說核稿期限：新案發文日＋2個月+15個日曆天再往前推3個工作天
    dtEP08 = CompWorkDay(4, CompDate(2, 15, CompDate(1, 2, DBDATE(sPA10))), 1)
    '(核稿期限)承辦期限不可大於本所期限
    If sCP06 <> "" Then
        If dtEP08 > DBDATE(sCP06) Then
           dtEP08 = DBDATE(sCP06)
        End If
    End If

   '若(核稿期限)核稿期限大於會稿、寄中說的本所期限時提醒會更新
   If dtEP08 <> "" Then
      strE = "select cp06 from caseprogress where cp01='" & sCP01 & "'" & _
         " and cp02='" & sCP02 & "' and cp03='" & sCP03 & "'" & _
         " and cp04='" & sCP04 & "' and cp10 in ('924','949') and cp57||cp27 is null and cp06<" & dtEP08
      intE = 1
      Set rRS = ClsLawReadRstMsg(intE, strE)
      If intE = 1 Then
         If bShowMsg = True Then MsgBox "將更新核稿期限為【會稿】或【寄中說】之本所期限！"
         dtEP08 = "" & rRS(0)
      End If
   End If
   If dtEP08 <> "" Then
       PUB_FCPsetEP08M = True
   End If
   
End Function

'Added by Lydia 2016/06/21 外專分案(工程師主管分案)-更新核稿人至會稿或寄中說的承辦人
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Sub PUB_UpdateFCP924(ByVal iCP01 As String, ByVal iCP02 As String, ByVal iCP03 As String, ByVal iCP04 As String, ByVal iCP10 As String, ByRef tEP04 As Object)
Dim strTmp As String
Dim ii As Integer

   If tEP04 <> "" And InStr(FCPHaveEP09, iCP10) > 0 And iCP10 <> "" Then
      strTmp = "update caseprogress set cp14='" & tEP04 & "',cp122='Y' where cp01='" & iCP01 & "'" & _
         " and cp02='" & iCP02 & "' and cp03='" & iCP03 & "'" & _
         " and cp04='" & iCP04 & "' and cp10 in ('924','949') and cp57||cp27 is null and (cp14 is null or cp14<>'" & tEP04 & "')"
      cnnConnection.Execute strTmp, ii
   End If
End Sub

'Added by Lydia 2016/07/07 內專分案(工程師主管分案)-FMP案翻譯分案時,員工編號改外譯編號
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Sub PUB_GetPfmpCP14(ByVal sCP10 As String, ByRef tCP14t As Object)
Dim strQ As String

    If Left(tCP14t.Text, 1) <> "F" And sCP10 = "201" Then
       strQ = PUB_GetMapID(tCP14t.Text, 0)
       If strQ <> "" Then
          tCP14t.Text = strQ
       End If
    End If
End Sub

'Added by Lydia 2016/07/07 內專分案(工程師主管分案)-超頁、超項費、急件費、補收款(917,920,911,938,939)若新案已發文則詢問是否上發文日
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Function PUB_ProcessPchk(ByVal iCP01 As String, ByVal iCP02 As String, ByVal iCP03 As String, ByVal iCP04 As String, ByVal iCP10 As String, ByVal iPA09 As String, ByRef bolUpd27 As Boolean) As Boolean
Dim inA As Integer, strA1 As String
Dim rsA As New ADODB.Recordset

    If iCP10 = "917" Or iCP10 = "920" Or iCP10 = "911" Or iCP10 = "938" Or iCP10 = "939" Then
       strA1 = "select * from caseprogress where cp01='" & iCP01 & "' and cp02='" & iCP02 & "' and cp03='" & iCP03 & "' and cp04='" & iCP04 & "' AND CP10 IN (" & CaseMapIn & ") and cp27>0"
       inA = 1
       Set rsA = ClsLawReadRstMsg(inA, strA1)
       If inA = 1 Then
          strA1 = "select * from caseprogress where cp01='" & iCP01 & "' and cp02='" & iCP02 & "' and cp03='" & iCP03 & "' and cp04='" & iCP04 & "' AND CP10 IN (" & CaseMapIn & ") and cp27>0"
          inA = 1
          '台灣案已收再審案則不必詢問,P-090238
             If iPA09 <> "000" Then
             If MsgBox("新申請案已發文，請問本程序是否要上發文日？", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
                bolUpd27 = True
             End If
          Else
             strA1 = "select * from caseprogress where cp01='" & iCP01 & "' and cp02='" & iCP02 & "' and cp03='" & iCP03 & "' and cp04='" & iCP04 & "' AND CP10='107' and nvl(cp57,0)=0"
             inA = 1
             Set rsA = ClsLawReadRstMsg(inA, strA1)
             If inA = 0 Then
                If MsgBox("新申請案已發文，請問本程序是否要上發文日？", vbYesNo + vbDefaultButton1 + vbQuestion) = vbYes Then
                   bolUpd27 = True
                End If
             End If
          End If
       End If
       Set rsA = Nothing
    End If
End Function

'Added by Lydia 2016/07/07 內專分案(工程師主管分案)-判斷香港案或澳門案與大陸案之關聯
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Function PUB_GetPcm10(ByVal oldNo As String, ByVal pPA09 As String, ByVal nPA09 As String, ByVal pKind As String) As String

    If oldNo <> "" Then
       PUB_GetPcm10 = oldNo
    Else
       PUB_GetPcm10 = "0"
    End If
    '香港發明案
    If (pPA09 = "020" Or pPA09 = "221" Or pPA09 = "201") And nPA09 = "013" And pKind = "1" Then
       PUB_GetPcm10 = "4"
    End If
    '澳門發明案
    If pPA09 = "020" And nPA09 = "044" And pKind = "1" Then
       PUB_GetPcm10 = "5"
    End If
End Function

'Added by Lydia 2016/07/07 內專分案(工程師主管分案)-更新關聯案
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Sub PUB_SavePtoUpd1(ByVal bFMP As Boolean, ByRef mpa() As String, ByVal mCP09 As String, ByVal mCP10 As String, ByVal mCP06 As String, ByVal mPCTdate As String, ByVal InCNo As String, ByVal InCM10 As String)
Dim stInCNo(1 To 4) As String '關聯案 案號
Dim rsR1 As New ADODB.Recordset
Dim strR1 As String, strR2 As String
Dim stCP48 As String '承辦期限
Dim inR As Integer

    '有關聯案
    If InCNo <> "" Then
        ChgCaseNo InCNo, stInCNo
        '非PCT案,非關聯香港案、澳門案
        If mPCTdate = "" And Not (InCM10 = "4" Or InCM10 = "5") Then
           '若國內案已發文則文件齊備日上系統日
           strR1 = "SELECT CP09 FROM CASEPROGRESS" & _
               " WHERE CP01='" & stInCNo(1) & "' AND CP02='" & stInCNo(2) & "'" & _
               " AND CP03='" & stInCNo(3) & "' AND CP04='" & stInCNo(4) & "'" & _
               " AND CP10 in (" & CaseMapIn & ") AND CP27>0"
              
           inR = 1
           Set rsR1 = ClsLawReadRstMsg(inR, strR1)
           If inR = 1 Then
              strR2 = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & _
                 " WHERE EP02='" & mCP09 & "' AND EP06 IS NULL"
              cnnConnection.Execute strR2, inR
              '重新計算承辦期限
              If inR = 1 Then '齊備日有更新才做
                 If bFMP Or PUB_IfSetCP48() Then
                    stCP48 = Pub_GetHandleDay(mpa(1), mpa(9), mCP10, , TransDate(mCP06, 2), mCP09)
                    If stCP48 <> "" Then
                       strR2 = "UPDATE CASEPROGRESS SET CP48=" & stCP48 & _
                          " WHERE CP09='" & mCP09 & "'"
                       cnnConnection.Execute strR2, inR
                    End If
                 End If
              End If
           End If
           
           '瓊玉說不應限制設計P-090127,只要是大陸新申請案無繪圖人員都要帶國內案繪圖人員且草墨圖都不計件
           If mpa(9) = "020" And InStr(CaseMapIn, mCP10) > 0 Then
              Call PUB_UpdateEP13(mCP09, stInCNo())
           End If
        End If
         
        '依大陸案更新香港澳門期限
        '大陸-香港案
        If mpa(9) = "013" And mCP10 = "110" Then
           Call PUB_UpdCP07by020(stInCNo, bFMP, "4", strSrvDate(1))
           '更新大陸案的標準專利紀錄請求期限(NP)為續辦
           strR2 = "Update nextprogress set np06='Y' where np02='" & stInCNo(1) & "' and np03='" & stInCNo(2) & "' and np04='" & stInCNo(3) & "' and np05='" & stInCNo(4) & "' and np06 is null and np07='110'"
           cnnConnection.Execute strR2, inR
        End If
        '大陸-澳門案
        If mpa(9) = "044" And mCP10 = "101" Then
           Call PUB_UpdCP07by020(stInCNo, bFMP, "5")
        End If
    End If
End Sub

'Added by Lydia 2016/07/07 內專分案(工程師主管分案)-更新其他
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Sub PUB_SavePtoUpd3(ByRef mpa() As String, ByRef tCP14 As Object, ByVal mCP09 As String, ByVal mCP10 As String, ByVal mCP14 As String, ByVal mCP60 As String, ByVal mEP04 As String)
Dim stInCNo(1 To 4) As String '國內案案號
Dim rsR1 As New ADODB.Recordset
Dim strR1 As String, strR2 As String
Dim strCP48 As String '承辦期限
Dim n_EP04  As String  '更新後的核稿人
Dim inR As Integer

    '案件性質 941分析, 分案時自動上齊備日
    If mCP10 = "941" Then
       strR2 = "UPDATE ENGINEERPROGRESS SET EP06=" & strSrvDate(1) & " WHERE EP02='" & mCP09 & "' AND EP06 IS NULL"
       cnnConnection.Execute strR2, inR
    End If
    
    '若已開請款單則換承辦人或核稿人時發Mail通知靜芳
    If mCP60 > "X" Then
       n_EP04 = ""
       If mCP10 = "201" Then
          strR1 = "select ep04 from engineerprogress where ep02='" & mCP09 & "'"
          inR = 1
          Set rsR1 = ClsLawReadRstMsg(inR, strR1)
          If inR = 1 Then
             n_EP04 = "" & rsR1(0)
          End If
       End If
       'Modified by Lydia 2019/10/17 本所案號+"-"
       'PUB_PointReAssignInform mpa(1) & mpa(2) & mpa(3) & mpa(4), mCP60, mCP14, tCP14.Text, mEP04, n_EP04
       PUB_PointReAssignInform mpa(1) & "-" & mpa(2) & IIf(mpa(3) & mpa(4) = "000", "", "-" & mpa(3) & "-" & mpa(4)), mCP60, mCP14, tCP14.Text, mEP04, n_EP04
    End If
End Sub

'Added by Lydia 2016/07/07 內專分案(工程師主管分案)取得國內案的案號
'Memo by Lydia 2024/03/20 因為工程師主管已可以直接修改承辦人,所以移除工程師主管分案表單
Public Sub PUB_GetPcmNo(ByRef tNo As Object, ByRef sCM() As String, ByVal iNa01 As String, ByVal iCP10 As String)

   If InStr(CaseMapIn, iCP10) > 0 Then
      '國內案
      If Cls003GetCaseMap(sCM) = True Then tNo.Text = sCM(4) & sCM(5) & sCM(6) & sCM(7)
      '香港案
      If tNo.Text = "" And iNa01 = "013" Then
         If Cls003GetCaseMap(sCM, 4) = True Then tNo.Text = sCM(4) & sCM(5) & sCM(6) & sCM(7)
      End If
      '澳門案
      If tNo.Text = "" And iNa01 = "044" Then
         If Cls003GetCaseMap(sCM, 5) = True Then tNo.Text = sCM(4) & sCM(5) & sCM(6) & sCM(7)
      End If
   End If
End Sub

'Add By Sindy 2016/8/1 檢查專利代理人/申請人是否有上傳平台帳號
Public Function PUB_ChkCustWebExist(ByVal strCP01 As String, ByVal strCP02 As String, _
                                    ByVal strCP03 As String, ByVal strCP04 As String) As String
Dim rRS As New ADODB.Recordset
Dim intQ As Integer, strCon1 As String 'Added by Lydia 2024/03/29

   PUB_ChkCustWebExist = ""
   strCon1 = "select pa75 from patent,custweb where pa01='" & strCP01 & "' and pa02='" & strCP02 & "' and pa03='" & strCP03 & "' and pa04='" & strCP04 & "' and instr(cw04,pa75)>0" & _
            " Union select pa26 from patent,custweb where pa01='" & strCP01 & "' and pa02='" & strCP02 & "' and pa03='" & strCP03 & "' and pa04='" & strCP04 & "' and instr(cw04,pa26)>0" & _
            " Union select pa27 from patent,custweb where pa01='" & strCP01 & "' and pa02='" & strCP02 & "' and pa03='" & strCP03 & "' and pa04='" & strCP04 & "' and instr(cw04,pa27)>0" & _
            " Union select pa28 from patent,custweb where pa01='" & strCP01 & "' and pa02='" & strCP02 & "' and pa03='" & strCP03 & "' and pa04='" & strCP04 & "' and instr(cw04,pa28)>0" & _
            " Union select pa29 from patent,custweb where pa01='" & strCP01 & "' and pa02='" & strCP02 & "' and pa03='" & strCP03 & "' and pa04='" & strCP04 & "' and instr(cw04,pa29)>0" & _
            " Union select pa30 from patent,custweb where pa01='" & strCP01 & "' and pa02='" & strCP02 & "' and pa03='" & strCP03 & "' and pa04='" & strCP04 & "' and instr(cw04,pa30)>0"
   intQ = 1
   Set rRS = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      rRS.MoveFirst
      Do While Not rRS.EOF
         PUB_ChkCustWebExist = PUB_ChkCustWebExist & "、" & rRS.Fields(0)
         rRS.MoveNext
      Loop
      If PUB_ChkCustWebExist <> "" Then PUB_ChkCustWebExist = Mid(PUB_ChkCustWebExist, 2)
   End If
End Function

'Add By Sindy 2016/12/6 檢查是否有變更事項,及取得欄位值
Public Function PUB_FCTchkChangeEventData(ByVal strCE01 As String, ByVal strCol As String, _
      Optional ByRef strVal As String) As Boolean
Dim rRS As New ADODB.Recordset
Dim intQ As Integer, strCon1 As String 'Added by Lydia 2024/03/29

   PUB_FCTchkChangeEventData = False
   If Trim(strCE01) = "" Then Exit Function
   strVal = ""
   strCon1 = "select " & strCol & " from ChangeEvent where CE01='" & strCE01 & "'"
   intQ = 1
   Set rRS = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      PUB_FCTchkChangeEventData = True
      If "" & rRS.Fields(0) <> "" Then
         strVal = rRS.Fields(0)
      End If
   End If
End Function

'Added by Lydia 2016/12/19 商標非TF案的申請或分割尚未設定審查時間,發MAIL通知
Public Sub PUB_SetChkResultDateT(ByVal iCF01 As String, ByVal iCF02 As String, ByVal iCF03 As String, ByRef sFromDate As String, ByRef sToDate As String, ByVal iCase02 As String, ByVal iCase03 As String, ByVal iCase04 As String)
'sToDate 催審期限
Dim rsB1 As New ADODB.Recordset
Dim intB As Integer
Dim tmpSub As String, tmpCont As String, tmpTo As String
Dim strTmpA As String
Dim strMid As String 'Added by Lydia 2018/09/17

   sToDate = ""
   If iCF01 = "TF" Or sFromDate = "" _
      Or (iCF03 <> "101" And iCF03 <> "308") Then
      Exit Sub
   End If
   
   strTmpA = "SELECT NVL(CF05,0) FROM CASEFEE WHERE CF01=" & CNULL(iCF01) & " AND CF02=" & CNULL(iCF02) & " AND CF03=" & CNULL(iCF03)
   intB = 1
   Set rsB1 = ClsLawReadRstMsg(intB, strTmpA)
   If intB = 1 Then
      If Val("" & rsB1(0)) > 0 Then
         sToDate = CompDate(2, Val("" & rsB1(0)), sFromDate)
         'Added by Lydia 2018/09/17 若期限超過催審提醒範圍,改成下一催審提醒區間(1,16號) (ex.T-216201)
         If Right(strSrvDate(1), 2) >= "16" Then
            strMid = Left(CompDate(1, 1, strSrvDate(1)), 6) & "01"
         Else
            strMid = Left(strSrvDate(1), 6) & "16"
         End If
         If sToDate < strMid Then
             sToDate = strMid
         End If
         'end 2018/09/17
         Exit Sub
      End If
   End If
   
   strTmpA = "select na03,tm23,nvl(cu04,nvl(cu05,cu06)) cname from trademark,nation,customer where tm01='" & iCF01 & "' and tm02='" & iCase02 & "' and tm03='" & iCase03 & "' and tm04='" & iCase04 & "' " & _
            " and tm10=na01(+) and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+) "
   intB = 1
   Set rsB1 = ClsLawReadRstMsg(intB, strTmpA)
   If intB = 1 Then
      'Modofied by Lydia 2021/06/23 先改為80030洪琬姿
      'tmpTo = Pub_GetSpecMan("V")
      'Modified by Lydia 2021/06/30 改帶CFT主管
      'tmpTo = "80030"
      tmpTo = GetCFTSt16Manager(iCF01, iCase02, iCase03, iCase04)
      If tmpTo <> "" Then
      'end 2021/06/30
        tmpSub = Trim("" & rsB1.Fields("na03")) & "的" & IIf(iCF03 = "101", "申請", "分割") & "案尚未設定審查時間，以致" & iCF01 & "-" & iCase02 & IIf(iCase03 & iCase04 = "000", "", "-" & iCase03 & "-" & iCase04) & "未掛催審期限!"
        tmpCont = "本所案號：" & iCF01 & "-" & iCase02 & "-" & iCase03 & "-" & iCase04 & vbCrLf & _
                  "申請國家：" & Trim("" & rsB1.Fields("na03")) & vbCrLf & _
                  "申請人１：" & Trim("" & rsB1.Fields("cname")) & vbCrLf & vbCrLf
        tmpCont = tmpCont & "請儘速設定申請案之審查時間(天)，並請程序人員輸入催審期限，以便系統管制。"
        PUB_SendMail strUserNum, tmpTo, "", tmpSub, tmpCont
      End If 'Added by Lydia 2021/06/30
   End If

   Set rsB1 = Nothing
End Sub

'Added by Lydia 2017/02/07 期間內的公報資料若無國家地區設定，則提醒使用者先補完資料，才能跑報表
Public Function Pub_ChkTMBMValidate(ByVal bNo As String, ByVal eNo As String) As Boolean
Dim strRead As String
Dim rsRD As New ADODB.Recordset
Dim intR As Integer
   
   Pub_ChkTMBMValidate = False
   If bNo = "" Or eNo = "" Then Exit Function
   
   strRead = "select '1' ord1,count(*) cnt from TMBulletin where tmbm07='" & bNo & "' and rownum <2 " & _
             "union select '2' ord1,count(*) cnt from TMBulletin where tmbm07='" & eNo & "' and rownum <2 "
   intR = 1
   Set rsRD = ClsLawReadRstMsg(intR, strRead)
   If intR = 1 Then
      rsRD.MoveFirst
      Do While Not rsRD.EOF
         If Val("" & rsRD.Fields("cnt")) = 0 Then
            MsgBox "期別:" & IIf(rsRD.Fields("ord1") = "1", bNo, eNo) & "，尚未有公報資料！", vbCritical
            GoTo JumpExit
         End If
         rsRD.MoveNext
      Loop
   End If
   
   strRead = "SELECT TMBM01 from TMBULLETIN where tmbm07>='" & bNo & "' And tmbm07<='" & eNo & "' and TMBM05 IS NULL"
   intR = 1
   Set rsRD = ClsLawReadRstMsg(intR, strRead)
   If intR = 1 Then
      rsRD.MoveFirst
      strRead = ""
      Do While Not rsRD.EOF
         strRead = strRead & vbCrLf & rsRD.Fields("TMBM01")
         rsRD.MoveNext
      Loop
      If strRead <> "" Then
         MsgBox "下列審定號尚未輸入地區編號，請輸入後再次執行程式:" & strRead, vbCritical
         GoTo JumpExit
      End If
   End If
   
   Pub_ChkTMBMValidate = True
   
JumpExit:
   Set rsRD = Nothing
End Function

'Added by Lydia 2017/02/24 傳收文號抓核稿人
'Modifie by Lydia 2021/05/18 +承辦工程師pEP05, 完稿日pEP09
Public Function PUB_GetEP04id(ByVal pCP09 As String, Optional ByVal bolMap As Boolean = False, Optional ByRef pEP05 As String, Optional ByRef pEP09 As String) As String
'bolMap 若核稿人為外翻編號,則轉為所內編號
Dim strB As String, intB As Integer
Dim rsBD As New ADODB.Recordset
   
   PUB_GetEP04id = ""
   pEP09 = "": pEP05 = "" 'Added by Lydia 2021/05/18
   If pCP09 = "" Then Exit Function
   
   'Modified by Lydia 2021/05/18
   'strB = "select cp01,cp02,cp03,cp04,cp10,cp12,cp13,cp14,b.*,c.* from caseprogress a,EngineerProgress b,staff_idmap c " & _
          "where cp09='" & pCP09 & "' and cp09=ep02(+) and ep04=sim02(+) "
   strB = "select cp01,cp02,cp03,cp04,cp10,cp12,cp13,cp14,b.*,c.sim01 as ep05c,c.sim01 as ep04c " & _
            "from caseprogress a,EngineerProgress b,staff_idmap c,staff_idmap d " & _
          "where cp09='" & pCP09 & "' and cp09=ep02(+) and ep05=c.sim02(+) and ep04=d.sim02(+)"
   intB = 1
   
   Set rsBD = ClsLawReadRstMsg(intB, strB)
   If intB = 1 Then
      PUB_GetEP04id = "" & rsBD.Fields("ep04")
      'Modified by Lydia 2021/05/18
      'If Mid("" & rsBD.Fields("ep04"), 1, 1) = "F" And bolMap = True And "" & rsBD.Fields("sim01") <> "" Then
      '   PUB_GetEP04id = "" & rsBD.Fields("sim01")
      'End If
      pEP05 = "" & rsBD.Fields("ep05")  '承辦人
      pEP09 = "" & rsBD.Fields("ep09") '完稿日
      If bolMap = True Then
          If Mid("" & rsBD.Fields("ep05"), 1, 1) = "F" And "" & rsBD.Fields("ep05c") <> "" Then
              pEP05 = "" & rsBD.Fields("ep05c")
          End If
          '核稿人
          If Mid("" & rsBD.Fields("ep04"), 1, 1) = "F" And "" & rsBD.Fields("ep04c") <> "" Then
              PUB_GetEP04id = "" & rsBD.Fields("ep04c")
          End If
      End If
      'end 2021/05/18
   End If
     
   Set rsBD = Nothing
End Function

'Added by Lydia 2017/03/24 新增委任契約書用印記錄
'Modified by Lydia 2017/04/11 +pKind 受任人
Public Function PUB_AddRecSeal(ByVal pType As String, ByVal pCnt As String, ByVal pSpace As String, ByVal pContent As String, ByVal pKind As String)
'pType: 委任書種類
'pCnt: 列印份數
'pSpace: 是否為空白委任書(Y/null)
'pContent:委任書內容
Dim strAdd As String
Dim strKind As String 'Added by Lydia 2017/04/11
Dim intQ As Integer, rsQuery As New ADODB.Recordset
On Error GoTo ErrHandle

'Added by Lydia 2020/03/25 改成與公司別ACC080一致；舊資料不動
If strSrvDate(1) >= 智慧所更名日 Then
    strAdd = "select a0801,a0802 from acc080 where a0802='" & pKind & "' order by 1 "
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, strAdd)
    If intQ = 1 Then
       strKind = "" & rsQuery.Fields("a0801")
    Else
       GoTo JumpToExcept
    End If
    Set rsQuery = Nothing
Else
JumpToExcept: 'Added by Lydia 2020/03/25
    'Added by Lydia 2017/04/11
    If InStr(pKind, "智權") > 0 Then
       strKind = "3"
    ElseIf InStr(pKind, "法律") > 0 Then
       strKind = "2"
    ElseIf InStr(pKind, "商標") > 0 Then
       strKind = "1"
    End If
    'end 2017/04/11
End If

    If pType = "" Or Val(pCnt) = 0 Then Exit Function
    
    'Modified by Lydia 2017/04/11 +RS08
    strAdd = "INSERT INTO RECSEAL(RS01,RS02,RS03,RS04,RS05,RS06,RS07,RS08) VALUES ('" & strUserNum & "'," & strSrvDate(1) & "," & Format(ServerTime, "000000") & _
             ",'" & pType & "'," & pCnt & ",'" & pSpace & "','" & ChgSQL(pContent) & "','" & strKind & "') "
             
    cnnConnection.Execute strAdd
    
    PUB_AddRecSeal = True
    Exit Function
    
ErrHandle:
    If Err.Number <> 0 Then
       MsgBox Err.Number & " " & Err.Description
       Resume Next
    End If
End Function

'Added by Lydia 2017/05/09 專利：回傳優先權基礎案的狀態
Public Function PUB_ReadPDStateNew(ByRef strPA() As String, Optional ByVal oCp10 As String, Optional ByVal bolCls88 As Boolean = False) As ADODB.Recordset
Dim intA As Integer
Dim rsA As New ADODB.Recordset
Dim strA1 As String
  
   strA1 = "select '' AS V,PD05 AS 優先權日,PD06 AS 優先權號,NA03 AS 優先權國家,PD09 as 優先權存取碼,PA01||PA02||PA03||PA04 AS 本所案號,PD07 " & _
           "From PRIDATE, Nation, PATENT WHERE PD01='" & strPA(1) & "' AND PD02='" & strPA(2) & "' AND PD03='" & strPA(3) & "' AND PD04 ='" & strPA(4) & "' AND PD07=NA01(+) AND PD06=PA11(+) AND PD05=PA10(+) AND PD07=PA09(+) "
   If bolCls88 Then '抓已閉卷，並且閉卷原因為88被主張國內優先權
      strA1 = strA1 & "AND PA09='" & strPA(9) & "' AND PA57='Y' AND PA59='88' "
   Else
      Select Case oCp10
         Case "106" '主張國際優先權
             strA1 = strA1 & "AND PD07<>'" & strPA(9) & "' "
         Case "121" '主張國內優先權
             strA1 = strA1 & "AND PD07='" & strPA(9) & "' "
      End Select
   End If
   strA1 = strA1 & "ORDER BY PD07,PD05,PD06 "
   intA = 1
   Set rsA = ClsLawReadRstMsg(intA, strA1)

   Set PUB_ReadPDStateNew = rsA
   Set rsA = Nothing
End Function

'Added by Lydia 2017/05/09 區分出優先權號|優先權日|優先權國家
Public Sub PUB_GetPD060507(ByVal sList As String, ByRef oPD06 As String, ByRef oPD05 As String, ByRef oPD07 As String)
Dim ii As Integer
Dim tmpStr As Variant

   oPD06 = ""
   oPD05 = ""
   oPD07 = ""
      
   If sList = "" Or sList = "||" Then
      Exit Sub
   Else
      tmpStr = Split(sList, "|")
      For ii = 0 To UBound(tmpStr)
         Select Case ii
             Case 0: oPD06 = Trim(tmpStr(ii))
             Case 1: oPD05 = Trim(tmpStr(ii))
             Case 2: oPD07 = Trim(tmpStr(ii))
         End Select
      Next ii
   End If
End Sub

'Added by Lydia 2017/05/19 委任契約書表單大於主表單，控制主表單放大。
Public Sub PUB_InitForm210114(ByRef fForm As Form, ByRef cForm As Form)
   Dim iWidth As Long, iHeight As Long
   'Added by Lydia 2022/07/05
   '取得螢幕解析度
   Dim ScreenX As Long, ScreenY As Long
   ScreenX = Screen.Width \ Screen.TwipsPerPixelX
   ScreenY = Screen.Height \ Screen.TwipsPerPixelY
   If cForm.Width + (60 * Screen.TwipsPerPixelX) >= ScreenX * Screen.TwipsPerPixelX Then
      fForm.WindowState = 2
   End If
   If fForm.WindowState = 2 Then Exit Sub '判斷mdiMain最大化，不用調整
   'end 2022/07/05
   
   iWidth = fForm.Width - cForm.Width
   iHeight = fForm.Height - cForm.Height
   'Modified by Lydia 2022/07/05
   'Do While iWidth < 600 Or iHeight < 600
   '   fForm.Width = fForm.Width + 150
   '   fForm.Height = fForm.Height + 150
   '   iWidth = fForm.Width - cForm.Width
   '   iHeight = fForm.Height - cForm.Height
   'Loop
   If iWidth > 60 * Screen.TwipsPerPixelX And iHeight > 80 * Screen.TwipsPerPixelY Then Exit Sub
   If iWidth < 60 * Screen.TwipsPerPixelX Then
      iWidth = cForm.Width + (60 * Screen.TwipsPerPixelX)
      fForm.Width = iWidth
   End If
   If iHeight < 80 * Screen.TwipsPerPixelY Then
      If cForm.Height + (80 * Screen.TwipsPerPixelY) >= ScreenY * Screen.TwipsPerPixelY Then '含功能表的高度
         '超過螢幕解析度時，只要放大到最大解析度
         iHeight = ScreenY * Screen.TwipsPerPixelY
      Else
         iHeight = cForm.Height + (80 * Screen.TwipsPerPixelY)
      End If
      fForm.Height = iHeight
   End If
   'end 2022/07/05
End Sub

'Add By Sindy 2017/9/13　檢查是否有安裝OUTLOOK.EXE
Public Sub PUB_ChkFileTypeOpenExE(strFileName As String)
Dim strReaderPath As String
   
   Select Case Right(UCase(strFileName), 4)
      Case UCase(".msg")
         '預設用 Reader,找不到才檢查關聯設定
         If PUB_CheckIsRunning("OUTLOOK.EXE") = False Then
            strReaderPath = PUB_FindFirstFileAPI("C:\Program Files\", "OUTLOOK.EXE")
            DoEvents
            If strReaderPath <> "" Then
               Shell strReaderPath
               DoEvents
            End If
         End If
   End Select
End Sub

'Added by Lydia 2017/12/21 櫃台收文-檢查是否有不續辦相同性質且未到期的期限，若有則提醒操作人員注意要輸入接洽單上填寫的期限
Public Function Pub_GetNPDoubleMsg(ByVal nDate As String, ByVal np02 As String, ByVal np03 As String, ByVal np04 As String, ByVal np05 As String, ByVal NP07 As String) As String
'ex.FCP-53639 年費,最初客戶表示不續辦( 解除期限的管制下次期限=Y，會產出新的期限)
'               後來客戶又要續辦,收文預設期限要增加判斷
Dim intM As Integer
Dim rsAD As New ADODB.Recordset
Dim strCon As String
    Pub_GetNPDoubleMsg = ""
    
    If nDate = "" Then nDate = strSrvDate(1)
   '馬德里使用宣誓,抓母號
   If np02 = "TF" And NP07 = "105" Then
        strSql = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and substr(np03,1,5)=" + _
                  CNULL(Left(np03, 5)) + _
                  " and np07 in(" + NP07 + ") and np06='N' and np09>=" + CNULL(nDate, True)
   Else
        'Modified by Lydia 2018/06/25 拿掉案件性質的Cnull (ex. P117898收204修正會出錯)
        'strCon = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
                  CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
                  " and np07 in (" + CNULL(NP07) + ") and np06='N' and np09>=" + CNULL(nDate, True)
        strCon = "select np08,np09,NP15,NP10 from nextprogress where np02=" + CNULL(np02) + " and np03=" + _
                  CNULL(np03) + " and np04=" + CNULL(np04) + " and np05=" + CNULL(np05) + _
                  " and np07 in (" + NP07 + ") and np06='N' and np09>=" + CNULL(nDate, True)
   End If
    intM = 1
    Set rsAD = ClsLawReadRstMsg(intM, strCon)
    If intM = 1 Then
         Pub_GetNPDoubleMsg = "有不續辦相同性質且未到期的期限，請輸入接洽單上填寫的期限!!"
    End If
    Set rsAD = Nothing
End Function

'Add By Sindy 2023/3/16 讀取專利說明書頁數明細
Public Function PUB_ReadPageDetail(ByRef strPD01 As String, ByRef pageD() As String, _
   Optional ByRef strAddPage As String, Optional ByRef strCP167 As String, Optional ByRef strCP168 As String, _
   Optional ByRef strChangePA64 As String, Optional ByRef strChangePA65 As String, _
   Optional ByRef strChangePA67 As String, Optional ByRef strChangePA68 As String, _
   Optional ByVal strPD20 As String) As Boolean
   
Dim strSql As String, rsRecordset As New ADODB.Recordset, i As Integer

On Error GoTo ErrHand
   
   PUB_ReadPageDetail = False
   If strPD20 <> "" Then '中說一併修正
      strSql = "select * from pagedetail where pd20='" & strPD20 & "' order by pd15 desc,pd16 desc"
   Else
      strSql = "select * from pagedetail where pd01='" & strPD01 & "' order by pd15 desc,pd16 desc"
   End If
   rsRecordset.CursorLocation = adUseClient
   rsRecordset.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
   If rsRecordset.RecordCount > 0 Then
      PUB_ReadPageDetail = True
      strPD01 = "" & rsRecordset.Fields("pd01") '中說一併修正,發文時要取得修正的總收文號
      With rsRecordset
         For i = 0 To 20
            If IsNull(.Fields(i).Value) Then
               pageD(i + 1) = ""
            Else
               pageD(i + 1) = .Fields(i).Value
            End If
         Next
      End With
   Else
      PUB_ReadPageDetail = False
      rsRecordset.Close
      Exit Function
   End If
   
   '增減後:
   '摘要頁數
   If Val(pageD(2)) - Val(pageD(6)) - Val(pageD(10)) <> 0 Then
      strChangePA64 = Val(pageD(2)) - Val(pageD(6)) - Val(pageD(10))
   Else
      strChangePA64 = "0"
   End If
   '說明書頁數
   If Val(pageD(3)) - Val(pageD(7)) - Val(pageD(11)) <> 0 Then
      strChangePA65 = Val(pageD(3)) - Val(pageD(7)) - Val(pageD(11))
   Else
      strChangePA65 = "0"
   End If
   '申請專利範圍頁數
   If Val(pageD(4)) - Val(pageD(8)) - Val(pageD(12)) <> 0 Then
      strChangePA67 = Val(pageD(4)) - Val(pageD(8)) - Val(pageD(12))
   Else
      strChangePA67 = "0"
   End If
   '圖式頁數
   If Val(pageD(5)) - Val(pageD(9)) - Val(pageD(13)) <> 0 Then
      strChangePA68 = Val(pageD(5)) - Val(pageD(9)) - Val(pageD(13))
   Else
      strChangePA68 = "0"
   End If
   
   '合計:
   '增加頁數:
   If Val(pageD(2)) + Val(pageD(3)) + Val(pageD(4)) + Val(pageD(5)) > 0 Then
      strAddPage = Val(pageD(2)) + Val(pageD(3)) + Val(pageD(4)) + Val(pageD(5))
   Else
      strAddPage = ""
   End If
   '刪除未審頁數:
   If Val(pageD(6)) + Val(pageD(7)) + Val(pageD(8)) + Val(pageD(9)) > 0 Then
      strCP167 = Val(pageD(6)) + Val(pageD(7)) + Val(pageD(8)) + Val(pageD(9))
   Else
      strCP167 = ""
   End If
   '刪除已審頁數:
   If Val(pageD(10)) + Val(pageD(11)) + Val(pageD(12)) + Val(pageD(13)) > 0 Then
      strCP168 = Val(pageD(10)) + Val(pageD(11)) + Val(pageD(12)) + Val(pageD(13))
   Else
      strCP168 = ""
   End If
   
   Exit Function
   
ErrHand:
   MsgBox Err.Description
End Function

'Add By Sindy 2023/3/9
'取得總頁數/總項數
Public Sub PUB_GetAllPageItem(ByVal strCP09 As String, ByRef cp() As String, ByRef pa() As String, _
   ByRef m_allPage As String, ByRef m_allItem As String)
Dim intQ As Integer, strCon1 As String, strCon2 As String, rsQ1 As New ADODB.Recordset 'Added by Lydia 2024/03/29

On Error GoTo ErrHand
   
   '原總頁數:最近一筆進度的頁數
   'Modify By Sindy 2018/7/27 + and cp158>0
   'Modify By Sindy 2018/7/27 + 去掉and cp158>0 ex:FCP-058875 主動修正一併送中說
   strCon1 = "select cp09,cp10,nvl(cp135,0) from caseprogress" & _
               " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
               " and cp159=0 and nvl(cp135,0)>0"
   If strCP09 <> "" Then
      strCon1 = strCon1 & _
                  " and cp09<>'" & strCP09 & "'"
   End If
   strCon1 = strCon1 & _
               " ORDER BY CP27 DESC,CP82 DESC" '發文日期時間大到小
   intQ = 1
   Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      m_allPage = Val("" & rsQ1.Fields(2))
   End If
   '原總項數:增加項數-刪除未審項數-刪除已審項數
   'Modify By Sindy 2018/7/27 + and cp158>0
   'Modify By Sindy 2018/8/27 若已有再審,以再審後的項數做計算
   strCon2 = ""
   If PUB_ChkCPExist(cp, "107", 2) Then
      strCon1 = "select cp66,cp67 from caseprogress" & _
                  " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
                  " and cp10='107'"
      intQ = 1
      Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
      If intQ = 1 Then
         strCon2 = " and (cp66>" & rsQ1.Fields("cp66") & " or(cp66=" & rsQ1.Fields("cp66") & " and cp67>=" & rsQ1.Fields("cp67") & "))"
      End If
   End If
   'Modify By Sindy 2018/7/27 + 去掉and cp158>0 ex:FCP-058875 主動修正一併送中說
   'Modify By Sindy 2023/3/9 +,cp135,cp167,cp168
   strCon1 = "select sum(nvl(cp136,0)) as cp136,sum(nvl(cp137,0)) as cp137,sum(nvl(cp138,0)) as cp138" & _
               ",sum(nvl(cp135,0)) as cp135,sum(nvl(cp167,0)) as cp167,sum(nvl(cp168,0)) as cp168" & _
               " from caseprogress" & _
               " where CP01='" & pa(1) & "' AND CP02='" & pa(2) & "' AND CP03='" & pa(3) & "' AND CP04='" & pa(4) & "'" & _
               " and cp159=0" & strCon2
   If strCP09 <> "" Then
      strCon1 = strCon1 & _
                  " and cp09<>'" & strCP09 & "'"
   End If
   'Modify By Sindy 2023/11/28 修正改用 having
   strCon1 = strCon1 & _
               " having (sum(nvl(cp136,0))>0 or sum(nvl(cp137,0))>0 or sum(nvl(cp138,0))>0 or sum(nvl(cp135,0))>0 or sum(nvl(cp167,0))>0 or sum(nvl(cp168,0))>0)"
   '2018/8/27 END
   intQ = 1
   Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      m_allItem = Val("" & rsQ1.Fields("cp136")) - Val("" & rsQ1.Fields("cp137")) - Val("" & rsQ1.Fields("cp138"))
      'Add By Sindy 2023/3/9
      m_allPage = Val("" & rsQ1.Fields("cp135")) - Val("" & rsQ1.Fields("cp167")) - Val("" & rsQ1.Fields("cp168"))
   End If
   
   Set rsQ1 = Nothing  'Added by Lydia 2024/03/29
   
   Exit Sub
   
ErrHand:
   MsgBox Err.Description
End Sub

'Add by Morgan 2010/1/5
'Modify by Morgan 2011/6/29 改和內專相同,要先扣除已收未發的超頁超項費
'檢查規費
'Modify By Sindy 2023/3/14 + , Optional txtCP167 As Object, Optional txtCP168 As Object
'Modify By Sindy 2023/3/27 + , Optional bolIsSend As Boolean = False : 發文作業才檢查
Public Function PUB_CheckOfficialFee_P(ByRef cp() As String, _
   ByVal m_bolChkPageItem As Boolean, ByVal m_bolChkItem As Boolean, _
   txtCP135 As Object, txtCP136 As Object, Optional txtCP137 As Object, Optional txtCP138 As Object, Optional txtCP84 As Object, _
   Optional ByRef m_lngRecOverPageFee As Long, Optional ByRef m_lngRecOverItemFee As Long, _
   Optional ByRef m_FeeMemo As String, Optional ByRef m_lngOverPageFee As Long, Optional ByRef m_lngOverItemFee As Long, _
   Optional ByRef m_lngOverPageFeeDiff As Long, Optional ByRef m_lngOverItemFeeDiff As Long, _
   Optional txtCP167 As Object, Optional txtCP168 As Object, _
   Optional bolIsSend As Boolean = False) As Boolean

Dim strMsg As String, bolBilled As Boolean
Dim dblCP84 As Double
Dim intQ As Integer, strCon1 As String, rsQ1 As New ADODB.Recordset 'Added by Lydia 2024/03/29

   PUB_CheckOfficialFee_P = True
   m_lngRecOverPageFee = 0
   m_lngRecOverItemFee = 0
   
   If m_bolChkPageItem Then
      If txtCP135 = "" Then
         MsgBox "頁數不可空白！", vbExclamation
         If txtCP135.Enabled Then txtCP135.SetFocus
         PUB_CheckOfficialFee_P = False
         Exit Function
      ElseIf txtCP136 = "" Then
         MsgBox "項數不可空白！", vbExclamation
         If txtCP136.Enabled Then txtCP136.SetFocus
         PUB_CheckOfficialFee_P = False
         Exit Function
      End If
   End If
   
   'Add by Morgan 2010/9/27
   If m_bolChkItem Then
      If TypeName(txtCP137) <> "Nothing" And TypeName(txtCP138) <> "Nothing" Then
         If txtCP136 = "" And txtCP137 = "" And txtCP138 = "" Then
            If MsgBox("增加項數及刪除項數皆為空白，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
               If txtCP136.Enabled Then txtCP136.SetFocus
               PUB_CheckOfficialFee_P = False
               Exit Function
            End If
         End If
      End If
      'Add By Sindy 2023/3/14
      If TypeName(txtCP167) <> "Nothing" And TypeName(txtCP168) <> "Nothing" Then
         If txtCP135 = "" And txtCP167 = "" And txtCP168 = "" Then
            If MsgBox("增加頁數及刪除頁數皆為空白，是否要繼續？", vbYesNo + vbDefaultButton2) = vbNo Then
               If txtCP135.Enabled Then txtCP135.SetFocus
               PUB_CheckOfficialFee_P = False
               Exit Function
            End If
         End If
      End If
      '2023/3/14 END
   End If
   'end 2010/9/27
   
   '退規費
   dblCP84 = 0
   If TypeName(txtCP84) <> "Nothing" Then dblCP84 = Val(txtCP84)
   If dblCP84 < 0 And TypeName(txtCP84) <> "Nothing" Then
      m_FeeMemo = "退規費 " & Format(-1 * Val(txtCP84), DAmount) & " 元;"
      MsgBox "本次發文可退規費【" & Format(-1 * Val(txtCP84), DAmount) & "】元！", vbInformation
      txtCP84 = 0
      
   ElseIf m_lngOverPageFee + m_lngOverItemFee > 0 Then
   
      m_lngOverPageFeeDiff = m_lngOverPageFee
      m_lngOverItemFeeDiff = m_lngOverItemFee
      
      'Add by Morgan 2011/6/29
      strCon1 = "select cp10,sum(nvl(cp17,0)-nvl(cp77,0)) Fee,max(cp60) BNo from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='938' and cp27||cp57 is null group by cp10" & _
         " union select cp10,sum(nvl(cp17,0)-nvl(cp77,0)) Fee,max(cp60) BNo from caseprogress where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "' and cp10='939' and cp27||cp57 is null group by cp10"
      intQ = 1
      Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
      If intQ = 1 Then
         Do While Not rsQ1.EOF
            If rsQ1.Fields("cp10") = "938" Then m_lngRecOverPageFee = m_lngRecOverPageFee + Val("" & rsQ1.Fields("Fee"))
            If rsQ1.Fields("cp10") = "939" Then m_lngRecOverItemFee = m_lngRecOverItemFee + Val("" & rsQ1.Fields("Fee"))
            If Not IsNull(rsQ1.Fields("BNo")) Then bolBilled = True
            rsQ1.MoveNext
         Loop
         m_lngOverPageFeeDiff = m_lngOverPageFee - m_lngRecOverPageFee
         m_lngOverItemFeeDiff = m_lngOverItemFee - m_lngRecOverItemFee
      End If
      
      If bolBilled And (m_lngOverPageFeeDiff <> 0 Or m_lngOverItemFeeDiff <> 0) Then
         MsgBox "本案已收文超頁費或超項費且已請款但金額不符，請修正後再發文！"
         PUB_CheckOfficialFee_P = False
         Exit Function
      End If
      'end 2011/6/29
      
      strMsg = ""
      If m_lngOverPageFee > 0 Then
         strMsg = "超頁費【" & Format(m_lngOverPageFee, DDollar) & "】元"
      End If
      If m_lngOverItemFee > 0 Then
         strMsg = strMsg & IIf(strMsg <> "", "及", "") & "超項費【" & Format(m_lngOverItemFee, DDollar) & "】元"
      End If
      If strMsg <> "" Then strMsg = strMsg & "，"
      'Modify By Sindy 2023/4/12
      If m_lngOverPageFee < 0 Or m_lngOverItemFee < 0 Then
         strMsg = strMsg & "代辦退費【" & Format(IIf(m_lngOverPageFee < 0, m_lngOverPageFee * -1, 0) + IIf(m_lngOverItemFee < 0, m_lngOverItemFee * -1, 0), DDollar) & "】元"
      End If
      '2023/4/12 END
      'Modify By Sindy 2023/3/27
      If bolIsSend = True Then
      '2023/3/27 END
         'Modify by Morgan 2011/6/29
         'If MsgBox("本案須繳" & strMsg & "共【" & Format(m_lngOverPageFee + m_lngOverItemFee, DDollar) & "】元，存檔時會自動做內部收文並同時上發文日，是否確定要繼續？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
         'If MsgBox("本案須繳" & strMsg & "共【" & Format(m_lngOverPageFee + m_lngOverItemFee, DDollar) & "】元，存檔時會自動做內部收文並同時上發文日(若已收文將更新收文金額並同時上發文日)，是否確定要繼續？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
         'Modify By Sindy 2023/4/12
         If MsgBox("本案須繳" & strMsg & "，存檔時會自動做內部收文並同時上發文日(若已收文將更新收文金額並同時上發文日)，是否確定要繼續？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
         '2023/4/12 END
         'end 2011/6/29
            PUB_CheckOfficialFee_P = False
            Exit Function
         End If
      End If
   End If
   Set rsQ1 = Nothing 'Added by Lydia 2024/03/29
   
End Function

'參考frm060104_3
'計算實審,修正規費
'Modify By Sindy 2018/5/8 + Optional ByVal m_bolESet As Boolean = False : 是否電子送件
'                         , Optional ByVal m_bolSend416 As Boolean = False : 是否一併提實審
'm_lngOfficialFee : 原始規費
'Modify By Sindy 2020/10/19 + , Optional ByVal m_strCP27 As String : 傳入畫面上的發文日
'Modify By Sindy 2023/3/14 + , Optional txtAddPageFee As Object, Optional txtDecreasePageFee As Object, Optional txtCP167 As Object
'Modify By Sindy 2023/4/7 + Optional m_WriteNote As String = "" : Y.一併送中說
'Modify By Sindy 2023/4/7 +, Optional ByRef m_str938RecvNo As String = "" : 回傳退費的超頁費文號
'                          , Optional ByRef m_str939RecvNo As String = "" : 回傳退費的超項費文號
Public Sub PUB_SetOfficialFee_P(ByRef cp() As String, ByRef pa() As String, _
   ByVal bolDelay As Boolean, ByVal m_strDelayCP09 As String, ByVal m_strReExamCP27 As String, _
   txtCP135 As Object, txtCP136 As Object, Optional txtCP137 As Object, Optional txtCP84 As Object, _
   Optional txtAddItemFee As Object, Optional txtDecreaseItemFee As Object, _
   Optional ByRef m_lngOverPageFee As Long, Optional ByRef m_lngOverItemFee As Long, _
   Optional ByVal m_bolESet As Boolean = False, Optional ByVal m_bolSend416 As Boolean = False, _
   Optional ByRef m_lngOfficialFee As Long, Optional ByVal m_strCP27 As String, _
   Optional txtAddPageFee As Object, Optional txtDecreasePageFee As Object, Optional txtCP167 As Object, _
   Optional m_WriteNote As String = "", Optional ByRef m_str938RecvNo As String = "", Optional ByRef m_str939RecvNo As String = "")
   
Dim iItems As Integer, iItemsOld As Integer, iItemsAdd As Integer
Dim strCPM02 As String 'Add By Sindy 2018/7/30
Dim iPages As Integer, iPagesOld As Integer, iPagesAdd As Integer 'Add By Sindy 2023/3/14
Dim intQ As Integer, strCon1 As String, strN1 As String, strN2 As String, rsQ1 As New ADODB.Recordset 'Added by Lydia 2024/03/29

   m_lngOverPageFee = 0
   m_lngOverItemFee = 0
         
   'Added by Morgan 2013/1/3
   '421.申請技術報告 807.申請第三人技術報告
   'Modify By Sindy 2024/8/7 申請技術報告也會有超項費
   'If cp(10) = "421" Or cp(10) = "807" Then
   If cp(10) = "807" Then
   '2024/8/7 END
      '發文規費
      If TypeName(txtCP84) <> "Nothing" Then
         txtCP84 = PUB_GetReportFee(pa(1), pa(9), cp(10), Val(txtCP136))
      End If
   Else
   'end 2013/1/3
         
      'Added by Morgan 2013/1/10
      If cp(10) = "107" Then
         m_lngOfficialFee = GetPatentOfficialFee(cp(1), cp(10), cp(7), pa(8), pa(9), pa(16))
         '有延期過則須扣除延期的發文規費
         If bolDelay = True Then
            strCon1 = "select cp84 from caseprogress where cp09='" & m_strDelayCP09 & "'"
            intQ = 1
            Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
            If intQ = 1 Then
               m_lngOfficialFee = m_lngOfficialFee - Val("" & rsQ1("cp84"))
            End If
         End If
      'Add By Sindy 2018/5/8 分割,新申請案
      ElseIf cp(10) = "307" Or cp(10) = "101" Or cp(10) = "102" Or cp(10) = "103" Then
'         If pa(8) = "1" Then
'            If m_bolSend416 = True Then
'               m_lngOfficialFee = "9900"
'            Else
'               m_lngOfficialFee = "2900"
'            End If
'         Else
'            m_lngOfficialFee = "2400"
'         End If
         If pa(8) = "1" Or cp(10) = "101" Then
            strCPM02 = "101"
         ElseIf pa(8) = "2" Or cp(10) = "102" Then
            strCPM02 = "102"
         Else
            strCPM02 = "103"
         End If
         strCon1 = "select cf08 from casefee where cf01='" & pa(1) & "' and cf02='" & pa(9) & "' and cf03='" & strCPM02 & "'"
         intQ = 1
         Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
         If intQ = 1 Then
            '原始規費
            m_lngOfficialFee = Val("" & rsQ1.Fields(0))
         End If
         '2018/5/8 END
         'Add By Sindy 2018/7/30
         If m_bolESet = True Then '電子送件
            m_lngOfficialFee = m_lngOfficialFee - 600
         End If
      Else
      'end 2013/1/10
         strCon1 = "select cf08 from casefee where cf01='" & pa(1) & "' and cf02='" & pa(9) & "' and cf03='" & cp(10) & "'"
         intQ = 1
         Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
         If intQ = 1 Then
            '原始規費
            m_lngOfficialFee = Val("" & rsQ1.Fields(0))
         End If
      End If 'Added by Morgan 2013/1/9
      'Add By Sindy 2018/7/30
      'Modified by Morgan 2022/5/12 排除435續行母案再審
      'If m_bolSend416 = True Then '一併提實審
      If m_bolSend416 = True And cp(10) <> "435" Then '一併提實審
      'end 2022/5/12
         m_lngOfficialFee = m_lngOfficialFee + 7000
      End If
      
      'Modify By Sindy 2023/3/23 Mark,改在下列程式同項規則做計算
'      '超頁費
'      'Modify By Sindy 2018/5/30 發明才會有超頁費
'      If Val(txtCP135) > 50 And pa(8) = "1" Then
'         m_lngOverPageFee = 500# * ((Val(txtCP135) - 1) \ 50)
'      End If
      
      '超項費
      'Modified by Morgan 2013/1/8 +再審107
      'If cp(10) = "416" Or cp(10) = "201" Or cp(10) = "209" Or cp(10) = "210" Then
      'Modified by Morgan 2013/10/18 +435續行母案再審
      'Modified by Morgan 2013/11/6 +235核對中說格式
      '416.實體審查 201.新案翻譯 209.檢視中說 210.製作中說
      'Modify by Sindy 2018/7/30 +307.分割
      If cp(10) = "416" Or cp(10) = "435" Or cp(10) = "107" Or cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or cp(10) = "210" Or cp(10) = "307" Then
         iItems = Val(txtCP136)
         'Modify By Sindy 2018/5/30 發明才會有超項費
         If iItems > 10 Then
            If pa(8) = "1" Then
               m_lngOverItemFee = 800# * (iItems - 10)
            End If
         End If
         'Add By Sindy 2023/3/14
         iPages = Val(txtCP135)
         If iPages > 50 And pa(8) = "1" Then
            m_lngOverPageFee = 500# * ((iPages - 1) \ 50)
         End If
         '2023/3/14 END
         
      Else
         If TypeName(txtCP137) <> "Nothing" Then
            iItemsAdd = Val(txtCP136) - Val(txtCP137)
         Else
            iItemsAdd = Val(txtCP136)
         End If
         'Add By Sindy 2023/3/14
         If TypeName(txtCP167) <> "Nothing" Then
            iPagesAdd = Val(txtCP135) - Val(txtCP167)
         Else
            iPagesAdd = Val(txtCP135)
         End If
         '2023/3/14 END
         
         'Add By Sindy 2024/8/7 不用管其他進度的項數
         If cp(10) = "421" Then
            iItemsOld = 0
            iPagesOld = 0
         Else
         '2024/8/7 END
            strCon1 = "select sum(cp136),sum(cp137),sum(cp138),sum(cp135) as totCP135,sum(cp167) as totCP167,sum(cp168) as totCP168" & _
                        " from caseprogress" & _
                        " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'" & _
                        " and cp27>0 and cp57 is null"
            'Added by Morgan 2013/1/10
            '若為再審後修正時只能抓再審以後發文的程序加總
            If m_strReExamCP27 <> "" Then
               strCon1 = strCon1 & " and cp27>=" & m_strReExamCP27
            End If
            'end 2013/1/10
            intQ = 1
            Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
            If intQ = 1 Then
               iItemsOld = Val("" & rsQ1.Fields(0)) - Val("" & rsQ1.Fields(1))
               iPagesOld = Val("" & rsQ1.Fields("totCP135")) - Val("" & rsQ1.Fields("totCP167")) 'Add By Sindy 2023/3/14
            End If
         End If
         iItems = iItemsOld + iItemsAdd
         '項數增加
         If iItemsAdd > 0 Then
            '超過10項
            'Modify By Sindy 2018/5/30 發明才會有超項費
            If iItemsOld >= 10 Then
               If pa(8) = "1" Then
                  m_lngOverItemFee = 800# * iItemsAdd
               'Modify By Sindy 2024/8/7 申請技術報告也會有超項費
               ElseIf cp(10) = "421" Then
                  m_lngOverItemFee = 600# * iItemsAdd
               '2024/8/7 END
               End If
            ElseIf iItems >= 10 Then
               If pa(8) = "1" Then
                  m_lngOverItemFee = 800# * (iItems - 10)
               'Modify By Sindy 2024/8/7 申請技術報告也會有超項費
               ElseIf cp(10) = "421" Then
                  m_lngOverItemFee = 600# * (iItems - 10)
               '2024/8/7 END
               End If
            End If
         '項數減少
         ElseIf iItemsAdd < 0 Then
            '超過10項
            'Modify By Sindy 2018/5/30 發明才會有超項費
            If iItems >= 10 Then
               If pa(8) = "1" Then
                  m_lngOverItemFee = 800# * iItemsAdd
               'Modify By Sindy 2024/8/7 申請技術報告也會有超項費
               ElseIf cp(10) = "421" Then
                  m_lngOverItemFee = 600# * iItemsAdd
               '2024/8/7 END
               End If
            '刪減後少於10項,但原來總項數>10,則可退原繳超項費=800*(原來總項數-10)
            ElseIf iItemsOld > 10 Then
               If pa(8) = "1" Then
                  m_lngOverItemFee = -1 * 800# * (iItemsOld - 10)
               'Modify By Sindy 2024/8/7 申請技術報告也會有超項費
               ElseIf cp(10) = "421" Then
                  m_lngOverItemFee = -1 * 600# * (iItemsOld - 10)
               '2024/8/7 END
               End If
            End If
         End If
         'Add By Sindy 2023/3/14
         iPages = iPagesOld + iPagesAdd
         '頁數增加
         If iPagesAdd > 0 Then
            '超過50頁,每50頁再加收500元
            '發明才會有超頁費
            If iPages > 50 And pa(8) = "1" Then
               'Modify By Sindy 2023/8/10 EX:FCP-61900申復
               strN1 = 500# * ((iPagesOld - 1) \ 50)
               strN2 = 500# * ((iPages - 1) \ 50)
               m_lngOverPageFee = Val(strN2) - Val(strN1)
               '2023/8/10 END
'            ElseIf iPages > 50 And pa(8) = "1" Then
'               m_lngOverPageFee = 500# * ((iPages - 1) \ 50)
            End If
         '頁數減少
         ElseIf iPagesAdd < 0 Then
            '發明才會有超頁費
            'Modify By Sindy 2023/8/10
'            If iPages > 50 And pa(8) = "1" Then
            If iPagesOld > 50 And pa(8) = "1" Then
               'Modify By Sindy 2023/8/10
               strN1 = 500# * ((iPagesOld - 1) \ 50)
               strN2 = 500# * ((iPages - 1) \ 50)
               m_lngOverPageFee = Val(strN2) - Val(strN1)
               '2023/8/10 END
'            '刪減後少於50頁,但原來總頁數>50,則可退原繳超頁費
'            ElseIf iPagesOld > 50 And pa(8) = "1" Then
'               m_lngOverPageFee = -1 * (500# * ((iPagesOld - 1) \ 50))
            End If
            '2023/8/10 END
         End If
         '2023/3/14 END
      End If
      '發文規費
      If TypeName(txtCP84) <> "Nothing" Then
         'Modify By Sindy 2023/3/29 敏莉說費用是負數的就不要加入
         'txtCP84 = m_lngOfficialFee + m_lngOverPageFee + m_lngOverItemFee
         txtCP84 = m_lngOfficialFee
         If Val(m_lngOverPageFee) > 0 Then
            txtCP84 = txtCP84 + m_lngOverPageFee
         End If
         If Val(m_lngOverItemFee) > 0 Then
            txtCP84 = txtCP84 + m_lngOverItemFee
         End If
         '2023/3/29 END
      End If
      If TypeName(txtAddItemFee) <> "Nothing" Then
         'Modify By Sindy 2019/1/22 應加收規費不含超頁費 ex:FCP-060147
         'txtAddItemFee = m_lngOverPageFee + m_lngOverItemFee '本次應加收規費
         txtAddItemFee = m_lngOverItemFee  '本次項數應加收規費
      End If
      'If Val(txtCP137) > 0 Then
      If TypeName(txtDecreaseItemFee) <> "Nothing" Then
         txtDecreaseItemFee = ""
         If m_lngOverItemFee < 0 Then
            txtDecreaseItemFee = m_lngOverItemFee * -1 '本次項數應退還規費
         End If
      End If
      'Add By Sindy 2023/3/14
      If TypeName(txtAddPageFee) <> "Nothing" Then
         txtAddPageFee = m_lngOverPageFee  '本次頁數應加收規費
      End If
      If TypeName(txtDecreasePageFee) <> "Nothing" Then
         txtDecreasePageFee = ""
         If m_lngOverPageFee < 0 Then
            txtDecreasePageFee = m_lngOverPageFee * -1 '本次頁數應退還規費
         End If
      End If
      '2023/3/14 END
      
      'Add By Sindy 2019/3/26 在產生新案翻譯(含檢視中說、核對中說格式)的電子送件申請書，
      '若實審發文日小於系統日且有超頁、超項是和實審同一天發文日且相關文號和實審收文號相同，
      '則規費的算法要減掉938超頁、939超項進度檔的相關資料的規費，大部份會是0元。
      'Modify By Sindy 2023/4/7 + m_WriteNote=Y : 一併送中說
      If cp(10) = "201" Or cp(10) = "209" Or cp(10) = "235" Or m_WriteNote = "Y" Then
         'sum(nvl(cp17,0)),cp10
         strCon1 = "select nvl(cp17,0),cp10,cp09 from caseprogress" & _
                     " where cp01='" & cp(1) & "' and cp02='" & cp(2) & "' and cp03='" & cp(3) & "' and cp04='" & cp(4) & "'" & _
                     " and cp158>0 and cp159=0" & _
                     " and cp10 in('938','939') and nvl(cp17,0)>0"
         If m_strCP27 <> "" Then
            strCon1 = strCon1 & " and cp27<>" & DBDATE(m_strCP27)
         End If
         'strcon1 = strcon1 & " group by cp10"
         intQ = 1
         Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
         If intQ = 1 Then
            rsQ1.MoveFirst
            Do While Not rsQ1.EOF
               If Val("" & rsQ1.Fields(0)) > 0 And txtCP84 > 0 Then
                  txtCP84 = Val(txtCP84) - Val("" & rsQ1.Fields(0))
                  If Val(txtCP84) < 0 Then txtCP84 = 0
               End If
               'Add By Sindy 2019/11/15 EX:FCP-61856 之前有繳過了
               '超頁費
               If rsQ1.Fields(1) = "938" Then
                  If m_str938RecvNo = "" Then m_str938RecvNo = rsQ1.Fields("cp09") 'Add By Sindy 2023/4/7
                  m_lngOverPageFee = Val(m_lngOverPageFee) - Val("" & rsQ1.Fields(0))
                  If Val(m_lngOverPageFee) < 0 Then
                     'Add By Sindy 2023/4/7
                     If TypeName(txtDecreasePageFee) <> "Nothing" Then
                        txtDecreasePageFee = m_lngOverPageFee * -1
                     End If
                     '2023/4/7 END
                     m_lngOverPageFee = 0
                  End If
               End If
               '超項費
               If rsQ1.Fields(1) = "939" Then
                  If m_str939RecvNo = "" Then m_str939RecvNo = rsQ1.Fields("cp09") 'Add By Sindy 2023/4/7
                  m_lngOverItemFee = Val(m_lngOverItemFee) - Val("" & rsQ1.Fields(0))
                  If Val(m_lngOverItemFee) < 0 Then
                     'Add By Sindy 2023/4/7
                     If TypeName(txtDecreaseItemFee) <> "Nothing" Then
                        txtDecreaseItemFee = m_lngOverItemFee * -1
                     End If
                     '2023/4/7 END
                     m_lngOverItemFee = 0
                  End If
               End If
               rsQ1.MoveNext
            Loop
            '2019/11/15 END
            'Add By Sindy 2023/4/7
            If TypeName(txtAddItemFee) <> "Nothing" Then
               txtAddItemFee = m_lngOverItemFee '本次項數應加收規費
            End If
            If TypeName(txtAddPageFee) <> "Nothing" Then
               txtAddPageFee = m_lngOverPageFee '本次頁數應加收規費
            End If
            'Modify By Sindy 2024/4/25 mark: FCP-70934新案翻譯會帶出500
'            '敏莉說費用是負數的就不要扣掉
'            If TypeName(txtDecreaseItemFee) <> "Nothing" Then
'               txtCP84 = Val(txtCP84) + Val(txtDecreaseItemFee)
'            End If
'            If TypeName(txtDecreasePageFee) <> "Nothing" Then
'               txtCP84 = Val(txtCP84) + Val(txtDecreasePageFee)
'            End If
            '2024/4/25 END
            '2023/4/7 END
         End If
      End If
      '2019/3/26 END
   End If 'Added by Morgan 2013/1/3
   
   Set rsQ1 = Nothing 'Added by Lydia 2024/03/29
End Sub

'Added by Morgan 2018/3/28
Public Function Utf8BytesFromString(strInput As String) As Byte()
    Dim nBytes As Long
    Dim abBuffer() As Byte
    ' CodePage constant for UTF-8
    Const CP_UTF8 = 65001
    ' Get length in bytes *including* terminating null
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, vbNull, 0&, 0&, 0&)
    ' We don't want the terminating null in our byte array, so ask for `nBytes-1` bytes
    ReDim abBuffer(nBytes - 2)  ' NB ReDim with one less byte than you need
    nBytes = WideCharToMultiByte(CP_UTF8, 0&, ByVal StrPtr(strInput), -1, ByVal VarPtr(abBuffer(0)), nBytes - 1, 0&, 0&)
    Utf8BytesFromString = abBuffer
End Function

'Added by Morgan 2018/3/28
'轉UTF8 URL 編碼
Public Function PUB_UTF8URLEncode(pEncodeStr As String) As String
   Dim bUTF8() As Byte
   Dim stRtn As String
   Dim ii As Integer
   
   bUTF8 = Utf8BytesFromString(pEncodeStr)
   For ii = LBound(bUTF8) To UBound(bUTF8)
      stRtn = stRtn & "%" & Hex(bUTF8(ii))
   Next
   PUB_UTF8URLEncode = stRtn
End Function

'Added by Lydia 2018/04/11  FCP外專翻譯承辦單列印(下載Word樣本檔，套印)
'Modified by Lydia 2020/04/10 +是否產生Word檔 bolSaveFile
'Modify By Sindy 2023/9/14 + , ByRef strColName() As String, ByRef strColText() As String : 回傳變數值
'                          + , Optional ByVal bolOnlyGetVal As Boolean = False : 是否純取值
'                          + , Optional ByRef intColCnt As Integer = 0 : 幾個欄位值
Public Function Pub_PrintFCP201Form(m_strPA01 As String, m_strPA02 As String, m_strPA03 As String, m_strPA04 As String, _
   m_strCP09 As String, ByRef strColName() As String, ByRef strColText() As String, _
   Optional ByVal bolSaveFile As Boolean = False, Optional ByVal bolOnlyGetVal As Boolean = False, _
   Optional ByRef intColCnt As Integer = 0) As Boolean
Dim m_strCP14 As String
Dim m_EP08 As String '核稿期限
Dim m_EP09 As String '完稿日
Dim m_strCP10 As String, m_strCP10n As String '案件性質
Dim rsA As New ADODB.Recordset
Dim m_Na16 As String 'FCP管制人
Dim m_CaseName As String '案件名稱
Dim m_PA150 As String '工程師組別
Dim m_Date924 As String '會稿所限
Dim m_Date201 As String '翻譯所限
Dim m_DateEP08 As String '核稿期限
Dim m_CP118t As String '是否電子送件
Dim m_CP14t As String '翻譯人員
Dim LenStar As Long
Dim intR As Integer
Dim DrawCount As Integer
Dim intFixedHigh As Integer
Dim intQ As Integer
'Added by Lydia 2018/05/03
Dim m_TCT118 As String '彩圖提申
Dim m_PA26n As String '申請人1
Dim m_DualCase As String '一案兩申請的關聯案號
Dim m_strCP113 As String '上班翻譯-工作時數
Dim m_TF19 As String, m_TF20 As String 'Added by Lydia 2018/05/21 相似度和相似案
'Added by Lydia 2019/04/22  使用樣本檔
Dim bVisible As Boolean
Dim iCall As Integer '樣本的變數
Dim m_FileName As String
'Dim strName As String, strText As String
Dim m_WordLeft As Long, m_WordTop As Long 'Word開啟位置
Dim m_EP31 As String, bHave924 As Boolean '是否會稿
Dim m_str924CP27 As String 'Added by Lydia 2019/05/14 會稿發文日
Dim m_CP64 As String 'Added by Lydia 2019/08/23 中說進度之備註
Dim m_str924CP09 As String 'Added by Lydia 2019/09/20 會稿收文號
Dim m_TF37 As String 'Added by Lydia 2019/10/25 翻譯瑕疵備註
Dim m_FilePath As String 'Added by Lydia 2020/04/10 Word完整路徑
Dim m_Murgitroyd As String 'Added byLydia 2020/10/14 Murgitroyd呈送期限設定: 代理人範圍
Dim bolPa175Y As Boolean 'Add By Sindy 2021/8/9
Dim strBCase(1 To 4) As String   'Added by Lydia 2024/03/29

   Pub_PrintFCP201Form = False 'Add By Sindy 2023/9/15
   '粗線深度
'   DrawCount = 20
'   intFixedHigh = 400 '行高
   
   'Modified by Lydia 2019/07/30 +翻譯時數m_strCP113
   'iCall = 12
   iCall = 13
   intColCnt = iCall 'Add By Sindy 2023/9/19
   
   'Add By Sindy 2023/9/14
   ReDim Preserve strColName(iCall) As String
   ReDim Preserve strColText(iCall) As String
   '2023/9/14 END
   
On Error GoTo ErrHand 'Added by Lydia 2019/04/22
   
   '進度檔
   'Modified by Lydia 2018/05/03 +cp113
   'Modified by Lydia 2018/05/21 翻譯費用檔(新案建檔預設代入命名作業的相似案和相似度，程序人員可修改相似案和相似度，該數值記錄在翻譯費用檔)
   'strSql = "select cp06,cp07,cp10,cp14,st02,cp159,cpm03,ep08,ep09,cp113 From caseprogress, casepropertymap, EngineerProgress,staff " & _
               "where cp09='" & m_strCP09 & "' and cp01=cpm01(+) and cp10=cpm02(+) and cp09=ep02(+) and cp14=st01(+)"
   'Modified by Lydia 2019/04/22 +EP31
   'Modified by Lydia 2019/08/23 +CP64
   'Modified by Lydia 2019/10/25 +TF37
   strSql = "select cp06,cp07,cp10,cp14,st02,cp159,cpm03,ep08,ep09,cp113,CP64,TF19,TF20,EP31,TF37 " & _
                "From caseprogress, casepropertymap, EngineerProgress,staff,transfee " & _
               "where cp09='" & m_strCP09 & "' and cp01=cpm01(+) and cp10=cpm02(+) " & _
               "and cp09=ep02(+) and cp14=st01(+) and cp09=tf01(+) "
   intQ = 1
   Set rsA = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
        If Val("" & rsA.Fields("cp159")) > 0 Then
            'Add By Sindy 2023/9/14
            If bolOnlyGetVal = True Then
               MsgBox "已取消收文，不可執行承辦歷程！"
               Exit Function
            Else
            '2023/9/14 END
               MsgBox "已取消收文，不可列印！"
               Exit Function
            End If
        End If
        m_strCP10 = "" & rsA.Fields("cp10")
        m_strCP10n = "" & rsA.Fields("cpm03")
        m_strCP14 = "" & rsA.Fields("cp14")
        m_CP14t = "" & rsA.Fields("st02")
        m_EP08 = TransDate("" & rsA.Fields("ep08"), 1)
        m_EP09 = TransDate("" & rsA.Fields("ep09"), 1)
        m_Date201 = TransDate("" & rsA.Fields("cp06"), 1)
        m_strCP113 = "" & rsA.Fields("cp113") 'Added by Lydia 2018/05/03
        m_CP64 = "" & rsA.Fields("CP64") 'Added by Lydia 2019/08/23
        'Added by Lydia 2018/05/21 相似度和相似案
        If "" & rsA.Fields("tf20") <> "" Then
             m_TF20 = "" & rsA.Fields("tf20")
             m_TF19 = "" & rsA.Fields("tf19")
        End If
        'end 2018/05/21
        m_EP31 = "" & rsA.Fields("EP31") 'Added by Lydia 2019/04/22
        m_TF37 = "" & rsA.Fields("TF37") 'Added by Lydia 2019/10/25 翻譯瑕疵備註
   End If
   '新案翻譯(201),控制需有上完稿日,才能列印
   If m_strCP10 = "201" And Val(m_EP09) = 0 Then
      'Add By Sindy 2023/9/14
      If bolOnlyGetVal = True Then
         MsgBox "尚未輸入完稿日，不可執行承辦歷程！"
         Exit Function
      Else
      '2023/9/14 END
         MsgBox "尚未輸入完稿日，不可列印！"
         Exit Function
      End If
   End If
   
   m_Murgitroyd = Pub_GetSpecMan("外專MURGITROYD設定")  'Added by Lydia 2020/10/14 Murgitroyd呈送期限設定: 代理人範圍
   
   '新案資料
   'Modified by Lydia 2018/05/03 +申請人,命名-彩圖提申
   'strSql = "select nvl(pa05,nvl(pa06,pa07)) casename,pa150,cp118 from patent,caseprogress " & _
               "Where PA01='" & m_strPA01 & "' AND PA02='" & m_strPA02 & "' AND PA03='" & m_strPA03 & "' AND PA04='" & m_strPA04 & "' " & _
               " and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and cp31='Y' "
   'Modified by Lydia 2018/05/21 +命名作業-相似度tct24和相似案tct23
   'Modified by Lydia 2020/10/14 +PA75
   'Modified by Sindy 2021/8/9 +pa175
   'Modified by Lydia 2023/05/22 +PA63,CP27
   strSql = "select nvl(pa05,nvl(pa06,pa07)) casename,pa150,cp118,pa26,nvl(cu04,nvl(cu05,cu06)) pa26n,tct118,tct23,tct24,PA75,pa175,PA63,CP27 " & _
               "from patent,caseprogress,customer,transcasetitle " & _
               "Where PA01='" & m_strPA01 & "' AND PA02='" & m_strPA02 & "' AND PA03='" & m_strPA03 & "' AND PA04='" & m_strPA04 & "' " & _
               "and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and cp31='Y' " & _
               "and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and cp09=tct01(+) "
   intQ = 1
   'Modify By Sindy 2023/9/19
   If bolOnlyGetVal = True Then
      m_CP118t = "非電子送件"
   Else
   '2023/9/19 END
      m_CP118t = "□電子送件　　■非電子送件　　□優先處理"
   End If
   Set rsA = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
       m_CaseName = "" & rsA.Fields("casename")
       m_PA150 = "" & rsA.Fields("pa150")
       If "" & rsA.Fields("CP118") = "Y" Or "" & rsA.Fields("CP118") = "A" Then
            'Modify By Sindy 2023/9/19
            If bolOnlyGetVal = True Then
               m_CP118t = "電子送件"
            Else
            '2023/9/19 END
               m_CP118t = "■電子送件　　□非電子送件　　□優先處理"
            End If
       End If
       'Added by Lydia 2018/05/03 申請人1和彩圖提申
       m_PA26n = "" & rsA.Fields("pa26") & " " & rsA.Fields("pa26n")
       
       'Added by Lydia 2023/05/22 以案件工程師組別區分英文組直接用基本檔的設定，日文組照舊用命名作業的設定。
       If m_PA150 <> "3" And "" & rsA.Fields("cp27") >= "20230501" Then
         If "" & rsA.Fields("pa63") = "Y" Then
              m_TCT118 = "■彩圖提申"
         End If
       Else
       'end 2023/05/22
         If "" & rsA.Fields("tct118") = "Y" Then
              m_TCT118 = "■彩圖提申"
         End If
         'end 2018/05/03
       End If 'Add ed by Lydia 2023/05/22
       
       'Added by Lydia 2018/05/21 相似度和相似案
       If m_TF20 = "" And "" & rsA.Fields("tct23") <> "" Then
            m_TF20 = "" & rsA.Fields("tct23")
            m_TF19 = "" & rsA.Fields("tct24")
       End If
       'end 2018/05/21
       'Added by Lydia 2020/10/14 Murgitroyd呈送期限設定:預設優先處理
       If m_Murgitroyd <> "" And InStr(m_Murgitroyd, "" & rsA.Fields("PA75")) > 0 Then
            'Modify By Sindy 2023/9/19
            If bolOnlyGetVal = True Then
               m_CP118t = "優先處理"
            Else
            '2023/9/19 END
               m_CP118t = Replace(m_CP118t, "□優先處理", "■優先處理")
            End If
       End If
       'end 2020/10/14
       
       'Modified by Sindy 2021/8/9 +pa175
       If "" & rsA.Fields("pa175") = "Y" Then
         bolPa175Y = True
       End If
       '2021/8/9 END
   End If
   
   'Modified by Lydia 2019/04/22 去掉最前面的空白
   'm_PA150 = " 工程師組別：" & IIf(m_PA150 <> "", PUB_GetFCPGrpName(m_PA150), "")
   m_PA150 = "工程師組別：" & IIf(m_PA150 <> "", PUB_GetFCPGrpName(m_PA150), "")
   
   '會稿期限
   'Modified by Lydia 2019/05/14 +CP27
   strSql = "select cp09,cp06,cp07,cp27 from caseprogress Where cp01='" & m_strPA01 & "' AND cp02='" & m_strPA02 & "' AND cp03='" & m_strPA03 & "' AND cp04='" & m_strPA04 & "' and cp10='924' and cp159=0 order by cp09 desc"
   intQ = 1
   Set rsA = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
       m_Date924 = TransDate("" & rsA.Fields("cp06"), 1)
       bHave924 = True 'Added by Lydia 2019/04/22
       m_str924CP27 = "" & rsA.Fields("cp27")
       m_str924CP09 = "" & rsA.Fields("cp09") 'Added by Lydia 2019/09/20
   End If
   If Trim(m_Date924) = "" Then
        m_Date924 = "　　年　　月　　日"
   Else
        m_Date924 = Mid(m_Date924, 1, 3) & " 年 " & Mid(m_Date924, 4, 2) & " 月 " & Mid(m_Date924, 6, 2) & " 日 "
   End If
   '翻譯所限
   If Trim(m_Date201) = "" Then
        m_Date201 = "　　年　　月　　日"
   Else
       m_Date201 = Mid(m_Date201, 1, 3) & " 年 " & Mid(m_Date201, 4, 2) & " 月 " & Mid(m_Date201, 6, 2) & " 日 "
   End If
   '核稿期限
   m_DateEP08 = TransDate(m_EP08, 1)
   If Trim(m_DateEP08) = "" Then
        m_DateEP08 = "　　年　　月　　日"
   Else
       m_DateEP08 = Mid(m_DateEP08, 1, 3) & " 年 " & Mid(m_DateEP08, 4, 2) & " 月 " & Mid(m_DateEP08, 6, 2) & " 日 "
   End If
   '程序
   'Modify By Sindy 2023/9/19
   If bolOnlyGetVal = True Then
      m_Na16 = PUB_GetFCPHandler(m_strPA01, m_strPA02, m_strPA03, m_strPA04)
   Else
   '2023/9/19 END
      m_Na16 = GetStaffName(PUB_GetFCPHandler(m_strPA01, m_strPA02, m_strPA03, m_strPA04), True)
   End If
   '翻譯人員（譯者）
   'Modified by Lydia 2025/03/13 改用模組取得
   'If m_CP14t <> "" And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, m_strCP14) = 0 Then
   If m_CP14t <> "" And InStr(Pub_SetF51Order("F", ""), m_strCP14) = 0 Then
      If Left(m_strCP14, 1) = "F" Then
          m_CP14t = m_CP14t & "-下班"
          m_strCP113 = "" 'Added by Lydia 2018/05/03
      Else
          m_CP14t = m_CP14t & "-上班"
      End If
   'Added by Lydia 2018/05/03
   Else
      m_strCP113 = ""
   'end 2018/05/03
   End If
   'Added by Lydia 2018/05/03 一案兩申請的關聯案號
   strSql = "SELECT PA01||'-'||PA02||DECODE(PA03,'0',NULL,'-'||PA03)||DECODE(PA04,'00',NULL,'-'||PA04) X01" & _
               " FROM ( SELECT CM01,CM02,CM03,CM04 FROM CASEMAP" & _
                            " Where CM10 = 3 AND CM05='" & m_strPA01 & "' AND CM06='" & m_strPA02 & "' AND CM07='" & m_strPA03 & "' AND CM08='" & m_strPA04 & "'" & _
                            " Union All  SELECT CM05,CM06,CM07,CM08 FROM CASEMAP" & _
                            " Where CM10 = 3  AND CM01='" & m_strPA01 & "' AND CM02='" & m_strPA02 & "' AND CM03='" & m_strPA03 & "' AND CM04='" & m_strPA04 & "'" & _
               " ) X, PATENT " & _
       " WHERE PA01(+)=CM01 AND PA02(+)=CM02 AND PA03(+)=CM03 AND PA04(+)=CM04 "
   intQ = 1
   Set rsA = ClsLawReadRstMsg(intQ, strSql)
   If intQ = 1 Then
       rsA.MoveFirst
       Do While Not rsA.EOF
           m_DualCase = m_DualCase & rsA.Fields("X01") & ","
           rsA.MoveNext
       Loop
   End If
   'end 2018/05/03
   
   'Add By Sindy 2023/9/14 讀取變數值
   For intR = 0 To iCall
      strColName(intR) = ""
      strColText(intR) = ""
      If intR = 0 Then '速別
          strColName(intR) = "速別"
          strColText(intR) = m_CP118t
      ElseIf intR = 1 Then '受文者
          strColName(intR) = "受文者"
          strColText(intR) = "智慧局"
      ElseIf intR = 2 Then '本所案號
          strColName(intR) = "本所案號"
          strColText(intR) = m_strPA01 & "-" & m_strPA02 & "-" & m_strPA03 & "-" & m_strPA04
      ElseIf intR = 3 Then '主旨
          strColName(intR) = "主旨"
          strColText(intR) = "中說送件"
      ElseIf intR = 4 Then '會稿期限
          strColName(intR) = "會稿期限"
          strColText(intR) = m_Date924
      ElseIf intR = 5 Then '譯者/案件性質
          strColName(intR) = "Tit01"
          If m_strCP10 = "201" Then
               strColText(intR) = "譯　者"
          Else
               strColText(intR) = "案件性質"
          End If
      ElseIf intR = 6 Then '譯者/案件性質
          strColName(intR) = "譯者"
          If m_strCP10 = "201" Then
               'Modify By Sindy 2023/9/19
               If bolOnlyGetVal = True Then
                  strColText(intR) = m_strCP14
               Else
               '2023/9/19 END
                  strColText(intR) = m_CP14t
               End If
          Else
               strColText(intR) = m_strCP10n
          End If
      ElseIf intR = 7 Then '核稿期限
          strColName(intR) = "核稿期限"
          strColText(intR) = m_DateEP08
      ElseIf intR = 8 Then '翻譯時數/檔案名稱
          strColName(intR) = "Tit02"
          If m_strCP10 = "201" Then
               strColText(intR) = "翻譯時數"
          Else
               strColText(intR) = "檔案名稱"
          End If
      ElseIf intR = 9 Then 'IPO期限
          strColName(intR) = "IPO所限"
          strColText(intR) = m_Date201
      ElseIf intR = 10 Then 'FCP管制人
          strColName(intR) = "FCP管制"
          strColText(intR) = m_Na16
      ElseIf intR = 11 Then '備註
          strColName(intR) = "備註"
          'Modify By Sindy 2023/9/19
          If bolOnlyGetVal = False Then
          '2023/9/19 END
            strColText(intR) = m_PA150 & vbCrLf
            strColText(intR) = strColText(intR) & convForm("申請人1：" & m_PA26n, 70) & vbCrLf
            strColText(intR) = strColText(intR) & "案件名稱：" & m_CaseName & vbCrLf & vbCrLf
          End If
          '彩圖提申
          If m_TCT118 <> "" Then strColText(intR) = strColText(intR) & m_TCT118 & vbCrLf
          '一案兩申請
          If m_DualCase <> "" Then strColText(intR) = strColText(intR) & "■本案與" & Mid(m_DualCase, 1, Len(m_DualCase) - 1) & "為一案兩請" & vbCrLf
          '相似案和相似度
          If m_TF20 <> "" Then
             Call ChgCaseNo(m_TF20, strBCase)
             strColText(intR) = strColText(intR) & "■與" & strBCase(1) & "-" & strBCase(2) & IIf(strBCase(3) & strBCase(4) = "000", "", "-" & strBCase(3) & "-" & strBCase(4)) & "有" & m_TF19 & "%相似" & vbCrLf
          End If
          'Added by Lydia 2019/04/29 加註Claims完稿日
          'Remove by Lydia 2019/09/17 因為已帶出進度備註,所以取消
          'If m_EP31 <> "" Then
          '     strColText(intR) = strColText(intR) & "■" & ChangeTStringToTDateString(TransDate(m_EP31, 1)) & "已交稿Claims" & vbCrLf
          'End If
          
          'Add By Sindy 2021/8/9
          If bolPa175Y = True Then
              'Modified by Lydia 2023/05/08 因應智慧局4/25起對序列表翻譯的變更; +(請工程師使用提申時的XML檔製作中文本)
              strColText(intR) = strColText(intR) & "■有序列表(請工程師使用提申時的XML檔製作中文本)" & vbCrLf
          End If
          '2021/8/9 END
          strColText(intR) = strColText(intR) & "□修改案件名稱：" & vbCrLf
          'Modify By Sindy 2024/1/23 淑華說中說備註內容的規費,刪除不顯示
'          strColText(intR) = strColText(intR) & "□有規費：" & vbCrLf
          '2024/1/23 END
          'Remove by Lydia 2023/05/08
          'strColText(intR) = strColText(intR) & "□序列表附光碟：" & vbCrLf
          strColText(intR) = strColText(intR) & "□申請書加註：" & vbCrLf
          'Added by Lydia 2019/10/25 翻譯瑕疵備註
          If m_TF37 <> "" Then
                strColText(intR) = strColText(intR) & "■翻譯瑕疵：" & m_TF37 & vbCrLf
          'Added by Lydia 2019/10/29 固定有勾選項
          Else
                strColText(intR) = strColText(intR) & "□翻譯瑕疵：" & vbCrLf
          'end 2019/10/29
          End If
          
          'Added by Lydia 2019/08/23 中說進度之備註
          If m_CP64 <> "" Then
              strColText(intR) = strColText(intR) & "進度備註：" & PUB_StringFilter(m_CP64) & vbCrLf
          End If
      ElseIf intR = 12 Then '*請消管制期限
           strColName(intR) = "Tit03"
           'Modified by Lydia 2019/05/14 只有會稿尚未發文,才要
           'If bHave924 = True Then
           If bHave924 = True And m_str924CP27 = "" Then
               strColText(intR) = "*請消管制期限"
           Else
               strColText(intR) = "　"
           End If
      'Added by Lydia 2019/07/30
      ElseIf intR = 13 Then
           strColName(intR) = "翻譯時數"
           strColText(intR) = m_strCP113
      End If
   Next intR
   
   Set rsA = Nothing
   'Add By Sindy 2023/9/14 純取值並沒有要產出承辦單
   If bolOnlyGetVal = True Then
      Pub_PrintFCP201Form = True 'Add By Sindy 2023/9/15
      Set rsA = Nothing
      Exit Function
   End If
   '2023/9/14 END
   
   'Modified by Lydia 2019/04/22 改成樣本
'   Printer.EndDoc
'
'   '畫線 *****************************
'   '裝訂線
'   Printer.DrawStyle = 2
'   Printer.Line (500, 0)-(500, 6300)
'   Printer.Line (500, 6300 + Printer.TextHeight("裝"))-(500, 7400)
'   Printer.Line (500, 7400 + Printer.TextHeight("訂"))-(500, 8600)
'   Printer.Line (500, 8600 + Printer.TextHeight("線"))-(500, 17000)
'   '打字
'   Printer.CurrentX = 500 - (Printer.TextWidth("裝") / 2)
'   Printer.CurrentY = 6300
'   Printer.Print "裝"
'   Printer.CurrentX = 500 - (Printer.TextWidth("訂") / 2)
'   Printer.CurrentY = 7400
'   Printer.Print "訂"
'   Printer.CurrentX = 500 - (Printer.TextWidth("線") / 2)
'   Printer.CurrentY = 8600
'   Printer.Print "線"
'   Printer.DrawStyle = 0
'   '粗線
'   For intR = 1 To DrawCount
'      Printer.Line (1000 + i, 5 * intFixedHigh + i)-(11200 + i, 40 * intFixedHigh + i), , B '最外框
'   Next i
'   '方格
'   '------橫線
'   Printer.Line (1000, 5 * intFixedHigh)-(11200, 6.5 * intFixedHigh), , B    '速別
'   Printer.Line (1000, 6.5 * intFixedHigh)-(11200, 8 * intFixedHigh), , B     '受文者
'   Printer.Line (1000, 8 * intFixedHigh)-(11200, 9.5 * intFixedHigh), , B    '主旨
'   Printer.Line (1000, 9.5 * intFixedHigh)-(11200, 11 * intFixedHigh), , B     '譯者
'   Printer.Line (1000, 11 * intFixedHigh)-(11200, 12.5 * intFixedHigh), , B     '翻譯時數
'   Printer.Line (1000, 12.5 * intFixedHigh)-(11200, 14.5 * intFixedHigh), , B    '判行
'   Printer.Line (1000, 14.5 * intFixedHigh)-(11200, 24.5 * intFixedHigh), , B
'   Printer.Line (1000, 24.5 * intFixedHigh)-(11200, 26 * intFixedHigh), , B '備註
'   Printer.Line (1000, 26 * intFixedHigh)-(11200, 38 * intFixedHigh), , B
'   '------直線
'   Printer.Line (2800, 5 * intFixedHigh)-(2800, 12.5 * intFixedHigh)     '小標題
'   Printer.Line (6600, 6.5 * intFixedHigh)-(6600, 12.5 * intFixedHigh)     '本所案號
'   Printer.Line (8400, 6.5 * intFixedHigh)-(8400, 12.5 * intFixedHigh)
'   Printer.Line (3600, 12.5 * intFixedHigh)-(3600, 24.5 * intFixedHigh)     '判行
'   Printer.Line (7300, 12.5 * intFixedHigh)-(7300, 24.5 * intFixedHigh)     '核稿 & 核對
'   Printer.Line (3600, 23 * intFixedHigh)-(7300, 23 * intFixedHigh)
'   Printer.Line (5500, 23 * intFixedHigh)-(5500, 24.5 * intFixedHigh)
'   Printer.Line (9100, 12.5 * intFixedHigh)-(9100, 24.5 * intFixedHigh)     '中打
'   Printer.Line (2800, 38 * intFixedHigh)-(2800, 40 * intFixedHigh) '發文日期
'
'   '抬頭
'   Printer.Font.Name = "標楷體"
'   Printer.Font.Size = 22
'   Printer.CurrentX = 3550
'   Printer.CurrentY = 460
'   Printer.Print "台一國際專利法律事務所"
'   Printer.CurrentX = 4880
'   Printer.CurrentY = 1020
'   Printer.Print "外專翻譯承辦單"
'
'   Printer.Font.Size = 14
'   PUB_PrintFontIntoBox "速　別", 1000, 5 * intFixedHigh, 2800, 6.5 * intFixedHigh
'   PUB_PrintFontIntoBox m_CP118t, 2900, 5 * intFixedHigh, 11200, 6.5 * intFixedHigh, , False
'   PUB_PrintFontIntoBox "受文者", 1000, 6.5 * intFixedHigh, 2800, 8 * intFixedHigh
'   PUB_PrintFontIntoBox "智慧局", 2900, 6.5 * intFixedHigh, 6600, 8 * intFixedHigh, , False
'   PUB_PrintFontIntoBox "本所案號", 6600, 6.5 * intFixedHigh, 8400, 8 * intFixedHigh
'   PUB_PrintFontIntoBox m_strPA01 & "-" & m_strPA02 & "-" & m_strPA03 & "-" & m_strPA04, 8500, 6.5 * intFixedHigh, 11200, 8 * intFixedHigh, , False
'   PUB_PrintFontIntoBox "主　旨", 1000, 8 * intFixedHigh, 2800, 9.5 * intFixedHigh
'   PUB_PrintFontIntoBox "中說送件", 2900, 8 * intFixedHigh, 6600, 9.5 * intFixedHigh, , False
'   PUB_PrintFontIntoBox "會稿期限", 6600, 8 * intFixedHigh, 8400, 9.5 * intFixedHigh
'   PUB_PrintFontIntoBox m_Date924, 8500, 8 * intFixedHigh, 11200, 9.5 * intFixedHigh, , False
'   If m_strCP10 = "201" Then
'        PUB_PrintFontIntoBox "譯　者", 1000, 9.5 * intFixedHigh, 2800, 11 * intFixedHigh
'        PUB_PrintFontIntoBox m_CP14t, 2900, 9.5 * intFixedHigh, 6600, 11 * intFixedHigh, , False
'   Else
'        PUB_PrintFontIntoBox "案件性質", 1000, 9.5 * intFixedHigh, 2800, 11 * intFixedHigh
'        PUB_PrintFontIntoBox m_strCP10n, 2900, 9.5 * intFixedHigh, 6600, 11 * intFixedHigh, , False
'   End If
'   PUB_PrintFontIntoBox "核稿期限", 6600, 9.5 * intFixedHigh, 8400, 11 * intFixedHigh
'   PUB_PrintFontIntoBox m_DateEP08, 8500, 9.5 * intFixedHigh, 11200, 11 * intFixedHigh, , False
'   If m_strCP10 = "201" Then
'       PUB_PrintFontIntoBox "翻譯時數", 1000, 11 * intFixedHigh, 2800, 12.5 * intFixedHigh
'       'Added by Lydia 2018/05/03 上班翻譯-工作時數
'       If m_strCP113 <> "" Then
'           PUB_PrintFontIntoBox m_strCP113, 2900, 11 * intFixedHigh, 6600, 12.5 * intFixedHigh, , False
'       End If
'       'end 2018/05/03
'   'Added by Lydia 2018/05/08 檢視中說的承辦單增加檔案日期欄位,由程序人工填寫,中打室依據該日期讀取檔案(by Sharon, Jack)
'   Else
'        'Modified by Lydia 2018/05/17 改成檔案名稱
'        'PUB_PrintFontIntoBox "檔案日期", 1000, 11 * intFixedHigh, 2800, 12.5 * intFixedHigh
'        PUB_PrintFontIntoBox "檔案名稱", 1000, 11 * intFixedHigh, 2800, 12.5 * intFixedHigh
'   'end 2018/05/08
'   End If
'   PUB_PrintFontIntoBox "IPO 所限", 6600, 11 * intFixedHigh, 8400, 12.5 * intFixedHigh
'   PUB_PrintFontIntoBox m_Date201, 8500, 11 * intFixedHigh, 11200, 12.5 * intFixedHigh, , False
'   PUB_PrintFontIntoBox "判　行", 1000, 12.5 * intFixedHigh, 3500, 14.5 * intFixedHigh
'   PUB_PrintFontIntoBox "核稿　＆　校對", 3600, 12.5 * intFixedHigh, 7200, 14.5 * intFixedHigh
'   PUB_PrintFontIntoBox "中打室", 7300, 12.5 * intFixedHigh, 9000, 14.5 * intFixedHigh
'   PUB_PrintFontIntoBox "程序管制人員", 9100, 12.5 * intFixedHigh, 11200, 14.5 * intFixedHigh
'   PUB_PrintFontIntoBox m_Na16, 9100, 14.5 * intFixedHigh, 11200, 15.5 * intFixedHigh '代入-程序
'   PUB_PrintFontIntoBox "承辦時數", 3600, 23 * intFixedHigh, 5400, 24.5 * intFixedHigh
'   PUB_PrintFontIntoBox "備　　　　註", 1000, 24.5 * intFixedHigh, 11200, 26 * intFixedHigh
'   PUB_PrintFontIntoBox m_PA150, 1000, 26 * intFixedHigh, 11200, 27 * intFixedHigh, , False '工程師組別
'   PUB_PrintFontIntoBox convForm(" 申請人1：" & m_PA26n, 70), 1000, 27 * intFixedHigh, 11200, 28 * intFixedHigh, , False 'Added by Lydia 2018/05/03 申請人1
'   '案件名稱
'   'Modified by Lydia 2018/05/03
''   PUB_PrintFontIntoBox " 案件名稱：", 1000, 27 * intFixedHigh, 2600, 28 * intFixedHigh, , False
''   LenStar = -Int(-(GetTextLength(m_CaseName) / 60)) '無條件進位
''   intQ = 1000 + Printer.TextWidth(" 案件名稱：")
''   PUB_PrintFontIntoBox m_CaseName, intQ, 27 * intFixedHigh, 11200, (27 + LenStar) * intFixedHigh, , False
''   intQ = 27 + LenStar + 1
'   i = 28
'   PUB_PrintFontIntoBox " 案件名稱：", 1000, i * intFixedHigh, 2600, (i + 1) * intFixedHigh, , False
'   LenStar = -Int(-(GetTextLength(m_CaseName) / 60)) '無條件進位
'   intQ = 1000 + Printer.TextWidth(" 案件名稱：")
'   PUB_PrintFontIntoBox m_CaseName, intQ, i * intFixedHigh, 11200, (i + LenStar) * intFixedHigh, , False
'   intQ = i + LenStar + 1
'   '彩圖提申
'   If m_TCT118 <> "" Then
'      PUB_PrintFontIntoBox " " & m_TCT118, 1000, intQ * intFixedHigh, 11200, (intQ + 1) * intFixedHigh, , False
'      intQ = intQ + 1
'   End If
'   '一案兩申請
'   If m_DualCase <> "" Then
'      PUB_PrintFontIntoBox " ■本案與" & Mid(m_DualCase, 1, Len(m_DualCase) - 1) & "為一案兩請", 1000, intQ * intFixedHigh, 11200, (intQ + 1) * intFixedHigh, , False
'      intQ = intQ + 1
'   End If
'   'end 2018/05/03
'   'Added by Lydia 2018/05/21 相似案和相似度
'   If m_TF20 <> "" Then
'      Call ChgCaseNo(m_TF20, strBCase)
'      PUB_PrintFontIntoBox " ■與" & strBCase(1) & "-" & strBCase(2) & IIf(strBCase(3) & strBCase(4) = "000", "", "-" & strBCase(3) & "-" & strBCase(4)) & "有" & m_TF19 & "%相似", 1000, intQ * intFixedHigh, 11200, (intQ + 1) * intFixedHigh, , False
'      intQ = intQ + 1
'   End If
'   'end 2018/05/21
'
'   PUB_PrintFontIntoBox " □修改案件名稱：", 1000, intQ * intFixedHigh, 11200, (intQ + 1) * intFixedHigh, , False
'   PUB_PrintFontIntoBox " □有規費：", 1000, (intQ + 1) * intFixedHigh, 11200, (intQ + 2) * intFixedHigh, , False
'   PUB_PrintFontIntoBox " □序列表附光碟：", 1000, (intQ + 2) * intFixedHigh, 11200, (intQ + 3) * intFixedHigh, , False
'   PUB_PrintFontIntoBox " □申請書加註：", 1000, (intQ + 3) * intFixedHigh, 11200, (intQ + 4) * intFixedHigh, , False
'   PUB_PrintFontIntoBox "發文日期", 1000, 38 * intFixedHigh, 2800, 39.5 * intFixedHigh
'   PUB_PrintFontIntoBox "　　年　　月　　日　(" & IIf(m_EP09 <> "", Mid(TransDate(m_EP09, 1), 1, 3), "   ") & ")" & _
'                                      "晉專外字　　　　　　　　號", 2900, 38 * intFixedHigh, 11200, 39.5 * intFixedHigh, , False
'   Set rsA = Nothing
'   Printer.EndDoc
   
   'Added by Lydia 2020/04/10 檢查檔案是否存在
   m_FilePath = PUB_Getdesktop
   If bolSaveFile = True Then
        strSql = m_FilePath & "\" & m_strPA01 & m_strPA02 & IIf(m_strPA03 & m_strPA04 <> "000", m_strPA03 & m_strPA04, "") & "翻譯承辦單.doc"
        If Dir(strSql) <> "" Then
             If PUB_ChkFileOpening(strSql, , False) = True Then
                  MsgBox strSql & vbCrLf & "檔案正在使用中，本次產生的電子檔會自動加上現在日期和時間。", vbExclamation
                  strSql = m_FilePath & "\" & m_strPA01 & m_strPA02 & IIf(m_strPA03 & m_strPA04 <> "000", m_strPA03 & m_strPA04, "") & "_" & strSrvDate(1) & Format(ServerTime, "000000") & "翻譯承辦單.doc"
             Else
                  If PUB_DelPCOrgFile(strSql, , False) = False Then
                      MsgBox strSql & vbCrLf & "檔案無法刪除，本次產生的電子檔會自動加上現在日期和時間。", vbExclamation
                      strSql = m_FilePath & "\" & m_strPA01 & m_strPA02 & IIf(m_strPA03 & m_strPA04 <> "000", m_strPA03 & m_strPA04, "") & "_" & strSrvDate(1) & Format(ServerTime, "000000") & "翻譯承辦單.doc"
                  End If
             End If
        End If
        m_FilePath = strSql
   End If
   'Add By Sindy 2022/8/12 借用此變數(m_strContactSheetA4)來傳回電子檔名
   If m_FilePath <> "" Then
      If InStr(m_strContactSheetA4, m_FilePath) = 0 Then
         m_strContactSheetA4 = IIf(m_strContactSheetA4 <> "", m_strContactSheetA4 & ";", "") & m_FilePath
      End If
   End If
   '2022/8/12 END
   'end 2020/04/10
   
    m_FileName = "外專翻譯_承辦單_樣本.doc"
    If Dir(App.path & "\" & m_FileName) <> "" Then Kill App.path & "\" & m_FileName
    
    Call PUB_GetSampleFile(m_FileName, "M51-000299-0-02")
    If Dir(App.path & "\" & m_FileName) <> "" Then
         Screen.MousePointer = vbHourglass
         '判斷word是否已開啟
RestarWord:
         If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = True Then
             g_WordAp.Documents.Open App.path & "\" & m_FileName
             With g_WordAp
                .Selection.WholeStory
                .Selection.Copy
                For intR = 0 To iCall
                   'Modify By Sindy 2023/9/14 mark,提到前面取變數值
'                   strName = ""
'                   strText = ""
'                   If intR = 0 Then '速別
'                       strName = "速別"
'                       strText = m_CP118t
'                   ElseIf intR = 1 Then '受文者
'                       strName = "受文者"
'                       strText = "智慧局"
'                   ElseIf intR = 2 Then '本所案號
'                       strName = "本所案號"
'                       strText = m_strPA01 & "-" & m_strPA02 & "-" & m_strPA03 & "-" & m_strPA04
'                   ElseIf intR = 3 Then '主旨
'                       strName = "主旨"
'                       strText = "中說送件"
'                   ElseIf intR = 4 Then '會稿期限
'                       strName = "會稿期限"
'                       strText = m_Date924
'                   ElseIf intR = 5 Then '譯者/案件性質
'                       strName = "Tit01"
'                       If m_strCP10 = "201" Then
'                            strText = "譯　者"
'                       Else
'                            strText = "案件性質"
'                       End If
'                   ElseIf intR = 6 Then '譯者/案件性質
'                       strName = "譯者"
'                       If m_strCP10 = "201" Then
'                            strText = m_CP14t
'                       Else
'                            strText = m_strCP10n
'                       End If
'                   ElseIf intR = 7 Then '核稿期限
'                       strName = "核稿期限"
'                       strText = m_DateEP08
'                   ElseIf intR = 8 Then '翻譯時數/檔案名稱
'                       strName = "Tit02"
'                       If m_strCP10 = "201" Then
'                            strText = "翻譯時數"
'                       Else
'                            strText = "檔案名稱"
'                       End If
'                   ElseIf intR = 9 Then 'IPO期限
'                       strName = "IPO所限"
'                       strText = m_Date201
'                   ElseIf intR = 10 Then 'FCP管制人
'                       strName = "FCP管制"
'                       strText = m_Na16
'                   ElseIf intR = 11 Then '備註
'                       strName = "備註"
'                       strText = m_PA150 & vbCrLf
'                       strText = strText & convForm("申請人1：" & m_PA26n, 70) & vbCrLf
'                       strText = strText & "案件名稱：" & m_CaseName & vbCrLf & vbCrLf
'                       '彩圖提申
'                       If m_TCT118 <> "" Then strText = strText & m_TCT118 & vbCrLf
'                       '一案兩申請
'                       If m_DualCase <> "" Then strText = strText & "■本案與" & Mid(m_DualCase, 1, Len(m_DualCase) - 1) & "為一案兩請" & vbCrLf
'                       '相似案和相似度
'                       If m_TF20 <> "" Then
'                          Call ChgCaseNo(m_TF20, strBCase)
'                          strText = strText & "■與" & strBCase(1) & "-" & strBCase(2) & IIf(strBCase(3) & strBCase(4) = "000", "", "-" & strBCase(3) & "-" & strBCase(4)) & "有" & m_TF19 & "%相似" & vbCrLf
'                       End If
'                       'Added by Lydia 2019/04/29 加註Claims完稿日
'                       'Remove by Lydia 2019/09/17 因為已帶出進度備註,所以取消
'                       'If m_EP31 <> "" Then
'                       '     strText = strText & "■" & ChangeTStringToTDateString(TransDate(m_EP31, 1)) & "已交稿Claims" & vbCrLf
'                       'End If
'
'                       'Add By Sindy 2021/8/9
'                       If bolPa175Y = True Then
'                           'Modified by Lydia 2023/05/08 因應智慧局4/25起對序列表翻譯的變更; +(請工程師使用提申時的XML檔製作中文本)
'                           strText = strText & "■有序列表(請工程師使用提申時的XML檔製作中文本)" & vbCrLf
'                       End If
'                       '2021/8/9 END
'                       strText = strText & "□修改案件名稱：" & vbCrLf
'                       strText = strText & "□有規費：" & vbCrLf
'                       'Remove by Lydia 2023/05/08
'                       'strText = strText & "□序列表附光碟：" & vbCrLf
'                       strText = strText & "□申請書加註：" & vbCrLf
'                       'Added by Lydia 2019/10/25 翻譯瑕疵備註
'                       If m_TF37 <> "" Then
'                             strText = strText & "■翻譯瑕疵：" & m_TF37 & vbCrLf
'                       'Added by Lydia 2019/10/29 固定有勾選項
'                       Else
'                             strText = strText & "□翻譯瑕疵：" & vbCrLf
'                       'end 2019/10/29
'                       End If
'
'                       'Added by Lydia 2019/08/23 中說進度之備註
'                       If m_CP64 <> "" Then
'                           strText = strText & "進度備註：" & PUB_StringFilter(m_CP64) & vbCrLf
'                       End If
'                   ElseIf intR = 12 Then '*請消管制期限
'                        strName = "Tit03"
'                        'Modified by Lydia 2019/05/14 只有會稿尚未發文,才要
'                        'If bHave924 = True Then
'                        If bHave924 = True And m_str924CP27 = "" Then
'                            strText = "*請消管制期限"
'                        Else
'                            strText = "　"
'                        End If
'                   'Added by Lydia 2019/07/30
'                   ElseIf intR = 13 Then
'                        strName = "翻譯時數"
'                        strText = m_strCP113
'                   End If
    
                   'Find並且置換
                   If Trim(strColName(intR)) <> "" Then
                      .Selection.Find.ClearFormatting
                      .Selection.Find.Text = "|#" & strColName(intR) & "#|"
                      .Selection.Find.Replacement.Text = ""
                      .Selection.Find.Forward = True
                      .Selection.Find.Wrap = wdFindContinue
                      .Selection.Find.Format = False
                      .Selection.Find.MatchCase = False
                      .Selection.Find.MatchWholeWord = False
                      .Selection.Find.MatchWildcards = False
                      .Selection.Find.MatchSoundsLike = False
                      .Selection.Find.MatchAllWordForms = False
                      .Selection.Find.MatchByte = True
                      .Selection.Find.Execute
                      .Selection.Delete
                      .Selection.TypeText strColText(intR)
                      .Selection.Find.Execute
                      .Selection.Font.Underline = wdUnderlineSingle
                   End If
ReadNext:
                Next intR
                
               'Added by Lydia 2020/04/10
               If bolSaveFile = True Then
                    '產生Word檔
                    .ActiveDocument.SaveAs m_FilePath
               Else
               'end 2020/04/10
                    '直接列印
                    .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
               End If 'Added by Lydia 2020/04/10
            End With
         End If
         Screen.MousePointer = vbDefault
         '還原Word位置
         Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop
         'Modified by Lydia 2020/04/10
         'g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
         If bolSaveFile = False Then g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
         
         g_WordAp.Quit wdDoNotSaveChanges
         Set g_WordAp = Nothing
         Pub_PrintFCP201Form = True 'Add By Sindy 2023/9/15
    Else
         MsgBox "無承辦單的樣本!", vbCritical
    End If
    
    Set rsA = Nothing
    
    'Added by Lydia 2019/04/22 不輸入Claims完稿日,於後面列印翻譯承辦單+會稿說明書承辦單
    'Modified by Lydia 2019/05/14 只有會稿尚未發文,才要印承辦單
    'If bHave924 = True And m_EP31 = "" Then
    If bHave924 = True And m_EP31 = "" And m_str924CP27 = "" Then
        'Modified by Lydia 2019/09/20 改傳入會稿收文號
        'Call Pub_PrintFCP924Form(m_strPA01, m_strPA02, m_strPA03, m_strPA04, m_strCP09)
        'Modified by Lydia 2020/04/16 +是否存檔
        'Call Pub_PrintFCP924Form(m_strPA01, m_strPA02, m_strPA03, m_strPA04, m_str924CP09)
        Call Pub_PrintFCP924Form(m_strPA01, m_strPA02, m_strPA03, m_strPA04, m_str924CP09, strColName, strColText, bolSaveFile)
    End If
    
    Exit Function
    
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   End If
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function

'Add By Sindy 2016/8/3
'內商發郵件資訊
'm_ET01=10:業務員期限管制表 ; Added by Lydia 2016/12/22 1728=收款寄證
'Modified by Lydia 2016/12/22 +總收文號 strNcp09,未結清款項 strA1kdata
'Modify By Sindy 2018/5/14 + oForm As Form
'                          , Optional ByVal strCP09 As String = ""
'                          , Optional ByVal lblSendData As LABEL
'Add By Sindy 2019/11/19 + , Optional ByVal strLP01 As String = "" : 信函總收文號
'Add By Sindy 2020/1/31 + , Optional ByVal m_bolAttFromCpp As Boolean = False : 附件是否選自卷宗區
'Add By Sindy 2020/7/7 + , Optional ByVal m_RetrunRecv As String = "" : 發指示信,多文號
'Add By Sindy 2020/7/21 + , Optional ByRef m_strSendDate As String = "", Optional ByRef m_strSendTime As String = "" : 回傳寄信日期時間
'Add By Sindy 2020/7/21 + , Optional ByVal m_bolTLetter As Boolean = False : 是否為商標指示信
'Modified by Sindy 2024/10/16 +bolReadLP42:是否要讀"定稿合併收文號"的案件資料
Public Function PUB_SettingTeMail(oForm As Form, ByVal strTemplatePath As String, _
         ByVal strCP01 As String, ByVal strCP02 As String, ByVal strCP03 As String, ByVal strCP04 As String, _
         Optional ByVal strAttach As String = "", Optional ByVal strCP10 As String = "", Optional ByVal strCP09 As String = "", _
         Optional ByVal m_ET01 As String = "", Optional strNCP09 As String = "", Optional strA1kdata As String = "", _
         Optional ByVal lblSendData As Object, Optional ByVal strLP01 As String = "", _
         Optional ByVal m_bolAttFromCpp As Boolean = False, Optional ByVal m_RetrunRecv As String = "", _
         Optional ByRef m_strSendDate As String = "", Optional ByRef m_strSendTime As String = "", _
         Optional ByVal m_bolTLetter As Boolean = False, Optional ByVal bolReadLP42 As Boolean = False) As Boolean
Dim adoRst As ADODB.Recordset
Dim objOutLook As Object
Dim objMail As Object
Dim strTM05 As String, strTM09 As String, strTM12 As String, strTM15 As String, strTM45 As String, strTM44 As String
Dim strST17 As String, strCuName As String, strContent As String
Dim strMsg As String, ii As Integer
Dim strTo As String, strCP10Nm As String, strTM13 As String
Dim bolIsECase As Boolean
Dim ArrStr As Variant
Dim strCP43m As String 'Added by Lydia 2016/12/22 相關總收文號的案件性質
Dim strSubject As String 'Add By Sindy 2018/5/14
Dim strTM10 As String
Dim strFA119 As String, strType As String 'Added by Sindy 2019/12/6
Dim strP20Mgr As String 'Add By Sindy 2020/1/13
Dim strCP06 As String 'Add By Sindy 2020/1/30
Dim strLP44 As String, strLP45 As String 'Add By Sindy 2020/2/24
Dim strCompName As String 'Add By Sindy 2020/3/26
Dim strCaseNoData As String, strContentCaseNo As String 'Add By Sindy 2020/7/8
Dim strCP14 As String, StrCount As String 'Add By Sindy 2020/10/13
Dim strMail As String 'Add By Sindy 2020/10/13
Dim pCustNo As String, salesNo As String, salesArea As String, strTFeeFile As String
Dim strCPMNm As String, strLD10 As String
Dim intQ As Integer, strCon1 As String 'Added by Lydia 2024/03/29
Dim strTM78 As String 'Added by Sindy 2024/7/16

   'Add By Sindy 2024/10/16 檢查是否有多個案號
   If m_RetrunRecv = "" And bolReadLP42 = True Then
      'Modify by Amy 2024/10/25 原:lp42='" & strCP09 & " ->少一個單引號
      strSql = "Select lp01 From letterprogress Where lp42='" & strCP09 & "' and lp01<>'" & strCP09 & "' order by lp01 asc"
      intI = 1
      Set RsTemp = ClsLawReadRstMsg(intI, strSql)
      If intI = 1 Then
         m_RetrunRecv = strCP09
         RsTemp.MoveFirst
         Do While Not RsTemp.EOF
            m_RetrunRecv = m_RetrunRecv & "," & RsTemp.Fields("lp01")
            RsTemp.MoveNext
         Loop
         '必須為多個案號
         strSql = "Select cp01,cp02,cp03,cp04 From caseprogress Where cp09 in('" & Replace(m_RetrunRecv, ",", "','") & "') group by cp01,cp02,cp03,cp04"
         intI = 1
         Set RsTemp = ClsLawReadRstMsg(intI, strSql)
         If intI = 1 Then
            If RsTemp.RecordCount = 1 Then
               m_RetrunRecv = ""
            End If
         End If
      End If
   End If
   '2024/10/16 END

   '查詢商標資料
   'Modified by Lydia 2016/12/22 + tm10
   'Modify by Sindy 2019/12/6 + tm44,fa119 + fagent
   'Modify By Sindy 2024/7/16 + ,tm78
   strCon1 = "select tm05,tm09,tm12,tm13,tm15,TM44,decode('" & strCP10 & "','102',nvl(TM65,TM45),TM45) TM45,st17,NVL(CU04,Decode(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) cuname,tm78,decode(tm10,'000',cpm03,cpm04) CP10Nm,tm10,tm44,fa119,'T' as TmType" & _
               " from trademark,staff,customer,casepropertymap,fagent" & _
               " where tm01='" & strCP01 & "' and tm02='" & strCP02 & "' and tm03='" & strCP03 & "' and tm04='" & strCP04 & "'" & _
               " and st01='" & strUserNum & "'" & _
               " and substr(tm23,1,8)=cu01(+) and substr(tm23,9,1)=cu02(+)" & _
               " and tm01=cpm01(+) and '" & strCP10 & "'=cpm02(+)" & _
               " and substr(tm44,1,8)=fa01(+) and substr(tm44,9,1)=fa02(+)"
   'Add by Sindy 2019/12/6 增加讀取服務業務基本資料檔
   'Modify By Sindy 2024/7/16 + ,sp58 to tm78
   'Modify by Amy 2024/07/17 原 ,sp58 to tm78->,sp58 as tm78
   strCon1 = strCon1 & " union select sp05||sp06||sp07,'' tm09,sp11,sp12,sp32,sp26,sp27,st17,NVL(CU04,Decode(cu05,null,CU06,cu05||' '||cu88||' '||cu89||' '||cu90)) cuname,sp58 as tm78,decode(sp09,'000',cpm03,cpm04) CP10Nm,sp09,sp26,fa119,'S' as TmType" & _
               " from servicepractice,staff,customer,casepropertymap,fagent" & _
               " where sp01='" & strCP01 & "' and sp02='" & strCP02 & "' and sp03='" & strCP03 & "' and sp04='" & strCP04 & "'" & _
               " and st01='" & strUserNum & "'" & _
               " and substr(sp08,1,8)=cu01(+) and substr(sp08,9,1)=cu02(+)" & _
               " and sp01=cpm01(+) and '" & strCP10 & "'=cpm02(+)" & _
               " and substr(sp26,1,8)=fa01(+) and substr(sp26,9,1)=fa02(+)"
   intQ = 1
   Set adoRst = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      strTM05 = "" & adoRst.Fields("tm05")
      strTM09 = "" & adoRst.Fields("tm09")
      strTM12 = "" & adoRst.Fields("tm12")
      strTM13 = "" & adoRst.Fields("tm13")
      strTM15 = "" & adoRst.Fields("tm15")
      strTM44 = "" & adoRst.Fields("tm44") 'Add By Sindy 2025/2/21
      strTM45 = "" & adoRst.Fields("tm45")
      strST17 = "" & adoRst.Fields("st17")
      strCuName = "" & adoRst.Fields("cuname")
      strCP10Nm = "" & adoRst.Fields("CP10Nm")
      strTM10 = "" & adoRst.Fields("tm10") 'Added by Lydia 2016/12/22
      strFA119 = "" & adoRst.Fields("FA119") 'Added by Sindy 2019/12/6
      strType = "" & adoRst.Fields("TmType") 'Added by Sindy 2019/12/6
      strTM78 = "" & adoRst.Fields("TM78") 'Added by Sindy 2024/7/16
   End If
   'Added by Lydia 2016/12/22 相關總收文號的案件性質
   If strNCP09 <> "" And strCP10 <> "1701" Then
      strCon1 = "select " & IIf(strTM10 = "000", "cpm03", "cpm04") & " as cp43m from caseprogress c1,caseprogress c2,casepropertymap x1 " & _
                  "where c1.cp09='" & strNCP09 & "' and c1.cp43=c2.cp09(+) and c2.cp01=cpm01(+) and c2.cp10=cpm02(+)  "
     intQ = 1
      Set adoRst = ClsLawReadRstMsg(intQ, strCon1)
      If intQ = 1 Then
         strCP43m = "" & adoRst.Fields(0)
      End If
   End If
   'end 2016/12/22
   
   'Add By Sindy 2019/11/19 有信函總收文號
   If Trim(strLP01) <> "" Then
      'Modify By Sindy 2021/1/4 + ld10
      strCon1 = "select c1.cp10 as cp10a,c2.cp10 as cp10b," & IIf(strTM10 = "000", "cpm03", "cpm04") & " as cp43m,c1.cp06 as cp06a,lp44,lp45,ld10" & _
                  " from caseprogress c1,caseprogress c2,casepropertymap x1,LetterProgress,letterdemand" & _
                  " where c1.cp09='" & strLP01 & "' and c1.cp43=c2.cp09(+) and c2.cp01=cpm01(+) and c2.cp10=cpm02(+)" & _
                  " and c1.cp09=lp01(+) and ld18(+)=c1.cp09"
      intQ = 1
      Set adoRst = ClsLawReadRstMsg(intQ, strCon1)
      If intQ = 1 Then
         strLP44 = "" & adoRst.Fields("lp44") 'Add By Sindy 2020/2/24
         strLP45 = "" & adoRst.Fields("lp45") 'Add By Sindy 2020/2/24
         strLD10 = "" & adoRst.Fields("ld10") 'Add By Sindy 2021/1/6
         Select Case adoRst.Fields("cp10a")
            '1725.通知期限
            '1729.撤三開拓
            Case "1725", "1729"
               m_ET01 = "10" '業務員期限管制表
            '1717.通知續展
            '1720.通知繳納註冊費
            Case "1717", "1720"
               m_ET01 = "15" '智慧局註冊費通知函
            Case "1001" '核准
               m_ET01 = "03" '核准
            Case "1002" '核駁
               m_ET01 = "04" '核駁
            Case "1701" '註冊證
               m_ET01 = "05" '發註冊證
           'Add By Sindy 2021/1/4
           Case Else
               m_ET01 = "" & adoRst.Fields("ld10")
            '2021/1/4 END
         End Select
         strCP06 = "" & adoRst.Fields("cp06a") 'Add By Sindy 2020/1/30
         'Modify By Sindy 2020/2/24 + And adoRst.Fields("cp10a") <> "1717" And adoRst.Fields("cp10a") <> "1720"
         'Add By Sindy 2021/3/3
         If Left(strLP01, 1) = "C" Then
         '2021/3/3 END
            If Trim("" & adoRst.Fields("cp10b")) <> "" And _
               adoRst.Fields("cp10a") <> "1717" And adoRst.Fields("cp10a") <> "1720" Then
               If strLD10 <> "" Then strCP10Nm = adoRst.Fields("cp43m")
               If strLD10 <> "" Then strCP10 = adoRst.Fields("cp10b") 'Add By Sindy 2021/1/4
            End If
         End If
      End If
   End If
   '2019/11/19 END
   
   'Modify By Sindy 2018/5/11
   '*************************************************************************************************
   'E-Mail呼叫 frm880019:要將寄信的內容及寄信的成功時間儲存在資料庫中，便於事後查詢。
   If TypeName(lblSendData) <> "Nothing" Then
      lblSendData.Caption = "寄件日期:"
      lblSendData.Visible = True
   End If
   frm880019.m_bolSaveMail = True
   frm880019.lblSender = "tm@taie.com.tw" 'Add By Sindy 2018/5/25 寄信者要掛TM信箱
   frm880019.m_CP01 = strCP01
   frm880019.m_CP02 = strCP02
   frm880019.m_CP03 = strCP03
   frm880019.m_CP04 = strCP04
   frm880019.m_CP09 = strCP09
   frm880019.m_CP10 = strCP10
   frm880019.m_LP01 = strLP01 'Add by Amy 2020/01/02
   '通知期限時,要抓此筆通知期限的進度
   'Modify By Sindy 2019/11/19 + And Trim(strLP01) = ""
   If m_ET01 = "10" And Trim(strLP01) = "" Then '業務員期限管制表
      strCon1 = "select cp09,cp10 from caseprogress" & _
                  " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                  " and substr(cp09,1,1)='D'" & _
                  " and cp05>=(select c2.cp05 from caseprogress c2 where c2.cp09='" & strCP09 & "')" & _
                  " order by cp05 desc,cp67 desc"
      intQ = 1
      Set adoRst = ClsLawReadRstMsg(intQ, strCon1)
      If intQ = 1 Then
         strCP09 = adoRst.Fields("cp09")
         frm880019.m_CP09 = strCP09
         frm880019.m_CP10 = adoRst.Fields("cp10")
      End If
   End If
   '副本.cc
   '收件者.To
   'Modified by Lydia 2016/12/22 指定1728收款寄證,收件者為代理人
   'strTo = PUB_GetFCeMailConText("Main_EMail", strCP01, strCP02, strCP03, strCP04, , strCP10)
   'Modify By Sindy 2023/12/4 mark:桂英在反應為什麼有些寄件備份有代理人名稱,有些沒有~ (統一收件人抓法)
   '                               ex:T-236847 xinshangzheng@126.com (代表信箱:遼寧新商政國際專利商標事務所有限公司);
'   If m_ET01 = "1728" Then
'      strTo = PUB_GetFCeMailConText("Main_EMail", strCP01, strCP02, strCP03, strCP04, "FC")
'      strMail = strTo 'Added by Lydia 2022/10/12
   'Add By Sindy 2020/10/13
'   Else
   If strTM10 <> "000" Then
      strMail = PUB_GetFCeMailConText("Main_EMail", strCP01, strCP02, strCP03, strCP04, "CF")
      strTo = PUB_GetFCeMailConText("Main_EMail", strCP01, strCP02, strCP03, strCP04, "CF", , True)
      strFA119 = ""
      If strTo <> "" Then
         strCon1 = "select fa01,fa02,fa119" & _
                     " from fagent" & _
                     " where fa01='" & Left(strTo, 8) & "' and fa02='" & Mid(strTo, 9, 1) & "'"
         intQ = 1
         Set adoRst = ClsLawReadRstMsg(intQ, strCon1)
         If intQ = 1 Then
            strFA119 = "" & adoRst.Fields("FA119")
         End If
      End If
   '2020/10/13 END
   Else
      'Modify By Sindy 2020/10/13
      'strTo = PUB_GetFCeMailConText("Main_EMail", strCP01, strCP02, strCP03, strCP04, , strCP10)
      strMail = PUB_GetFCeMailConText("Main_EMail", strCP01, strCP02, strCP03, strCP04, "FC")
      strTo = PUB_GetFCeMailConText("Main_EMail", strCP01, strCP02, strCP03, strCP04, "FC", , True)
   End If
   'end 2016/12/22
   'Add By Sindy 2021/3/2 PATTA商標監視(第07,08,09,19,20類)，其他 (我方案號：TT-000161)
   If UCase(strMail) = UCase("ipdept@taie.com.tw") Then strMail = ""
   If strTo = "Y00000000" Then strTo = ""
   '2021/3/2 END
   
   '非E化案件或無代表號信箱時, 顯示訊息, 若同時發生時只顯示一次訊息即可.
   'Modify By Sindy 2020/10/13 以申請國家區分抓CF還是FC
   If strTM10 <> "000" Then
      If PUB_GetFCeMailConText("IsECase", strCP01, strCP02, strCP03, strCP04, "CF") = "Y" Then bolIsECase = True
   Else
      If PUB_GetFCeMailConText("IsECase", strCP01, strCP02, strCP03, strCP04, "FC") = "Y" Then bolIsECase = True
   End If
   strMsg = ""
   'Modify by Sindy 2019/12/6 + Or strFA119 <> ""
   If strTo = "" Or bolIsECase = False Or strFA119 <> "" Then
      If bolIsECase = False Then
         strMsg = "非E化案件"
      End If
      If strMail = "" Then
         If strMsg <> "" Then strMsg = strMsg & "且"
         strMsg = strMsg & "無代表號信箱"
      End If
      'Modify by Sindy 2019/12/6
      If strFA119 <> "" Then
         If strMsg <> "" Then strMsg = strMsg & vbCrLf & vbCrLf
         strMsg = strMsg & "【陸代定稿加註】" & vbCrLf & vbCrLf & strFA119
      End If
      '2019/12/6 END
      MsgBox strMsg, vbInformation
   End If
   
   'Add By Sindy 2020/7/8
   '多筆文號
   If InStr(m_RetrunRecv, ",") > 0 Then
      strContentCaseNo = PUB_GetMailManyCaseData(m_RetrunRecv, strCaseNoData, strCP14, True, strCPMNm)
      'Add By Sindy 2023/11/14 ex:T-079999(許可合同)
      If InStr(strCaseNoData, ",") > 0 Then
      '2023/11/14 END
         strCaseNoData = Mid(strCaseNoData, 1, InStr(strCaseNoData, ",") - 1) + " ~"
      End If
      StrCount = "(共" & UBound(Split(m_RetrunRecv, ",")) + 1 & "筆)"
   Else
      If m_RetrunRecv <> "" Then
         strCon1 = "select cp09,cp10,cp14 from caseprogress" & _
                     " where cp09='" & m_RetrunRecv & "'"
         intQ = 1
         Set adoRst = ClsLawReadRstMsg(intQ, strCon1)
         If intQ = 1 Then
            strCP14 = "" & adoRst.Fields("cp14")
         End If
      End If
      strCaseNoData = strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 = "000", "", "-" & strCP03 & "-" & strCP04)
   End If
   '2020/7/8 END
   
   'Add By Sindy 2025/2/21 有代理人及彼所案號 -- 天雲
   strExc(10) = ""
   If strTM44 <> "" And strTM45 <> "" Then
      strExc(10) = "貴方卷號:" & strTM45 & " "
   End If
   '2025/2/21 END
   '密件副本.BCC
   '主旨.Subject
   '主旨：(A)AEGIS(第03,42,0128類)商標，核准通知
   '      (國外專業代號)商標名稱(第類)商標，案件性質
   'Modified by Lydia 2016/12/22 指定1728收款寄證,主旨:客戶名稱發xx通知(本所案號)
   'strSubject = "(" & strST17 & ")" & strTM05 & "(第" & strTM09 & "類)商標，" & strCP10Nm
   'Added by Sindy 2019/12/6
   If strType = "T" Then
   '2019/12/6 END
      If m_ET01 = "1728" Then
         'Modified by Lydia 2017/01/04 +專業代號 "(" & strST17 & ")"
         strSubject = "(" & strST17 & ")" & strCuName & "發" & strCP43m & strCP10Nm & "通知 (" & strExc(10) & "我方案號：" & strCaseNoData & ")"
      ElseIf m_ET01 = "10" Then '業務員期限管制表
         strSubject = "(" & strST17 & ")" & strTM05 & "(第" & strTM09 & "類)商標，通知續展 (" & strExc(10) & "我方案號：" & strCaseNoData & ")"
      'Modify By Sindy 2019/3/20
      ElseIf m_ET01 = "03" Then '核准
         strSubject = "(" & strST17 & ")" & strTM05 & "(第" & strTM09 & "類)商標，" & strCP10Nm & "-核准 (" & strExc(10) & "我方案號：" & strCaseNoData & ")"
      ElseIf m_ET01 = "04" Then '核駁
         strSubject = "(" & strST17 & ")" & strTM05 & "(第" & strTM09 & "類)商標，" & strCP10Nm & "-核駁 (" & strExc(10) & "我方案號：" & strCaseNoData & ")"
      ElseIf m_ET01 = "05" Then '發註冊證
         strSubject = "(" & strST17 & ")" & strTM05 & "(第" & strTM09 & "類)商標，發註冊證 (" & strExc(10) & "我方案號：" & strCaseNoData & ")"
      'Add By Sindy 2020/10/14
      ElseIf m_bolTLetter = True Then '指示信
         frm880019.m_isCFFagent = True 'Add By Sindy 2022/3/14 讓寄信辨識是CF代理人
         strSubject = "(" & strST17 & ")" & strTM05 & IIf(StrCount <> "", StrCount, "(第" & strTM09 & "類)") & "商標，" & IIf(strCPMNm <> "", strCPMNm, strCP10Nm) & " (" & strExc(10) & "我方案號：" & strCaseNoData & ")"
      Else
      '2019/3/20 END
         'Modify By Sindy 2019/3/6 ex:T-220261申請
         'IIf(strCP10Nm = "申請", strCP10Nm & "核准", strCP10Nm) ==> strCP10Nm
         'strSubject = "(" & strST17 & ")" & strTM05 & "(第" & strTM09 & "類)商標，" & IIf(strCP10Nm = "申請", strCP10Nm & "核准", strCP10Nm) & " (我方案號：" & strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 = "000", "", IIf(strCP03 <> "0", "-" & strCP03, IIf(strCP04 <> "00", "-" & strCP04, ""))) & ")"
         '主旨要顯示為:(C)雅戈爾(簡體字)YOUNGOR(第25類)商標，申請 (我方案號：T220261)
         strSubject = "(" & strST17 & ")" & strTM05 & IIf(StrCount <> "", StrCount, "(第" & strTM09 & "類)") & "商標，" & strCP10Nm & " (" & strExc(10) & "我方案號：" & strCaseNoData & ")"
         '2019/3/6 END
      End If
      'end 2016/12/22
   Else
      strSubject = "(" & strST17 & ")" & strTM05 & "，" & strCP10Nm & " (" & strExc(10) & "我方案號：" & strCaseNoData & ")"
   End If
   '主旨
   'Modify By Sindy 2020/2/24 ex:LP44=T-195289,T-195290,T-195292,T-195293
   If strLP44 <> "" And InStr(strSubject, "我方案號：") > 0 Then
      strSubject = Left(strSubject, InStr(strSubject, "我方案號：") + 4) & strLP44 & ")"
   End If
   '2020/2/24 END
   frm880019.txtSubject = strSubject
   '本文
   '內文.Body
   'Modify By Sindy 2020/2/24
   If strLP45 <> "" Then
      strContent = strLP45
   Else
   '2020/2/24 END
      'Add By Sindy 2020/7/8
      If strContentCaseNo <> "" Then
         strContent = strContentCaseNo
      Else
      '2020/7/8 END
         If Trim(strTM45) <> "" Then
            strContent = strContent & "貴方卷號：" & strTM45 & vbCrLf
         End If
         'Modify By Sindy 2020/7/8
         'strContent = strContent & "我方案號：" & strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 = "000", "", IIf(strCP03 <> "0", "-" & strCP03, IIf(strCP04 <> "00", "-" & strCP04, ""))) & vbCrLf & vbCrLf
         strContent = strContent & "我方案號：" & strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 = "000", "", "-" & strCP03 & "-" & strCP04) & vbCrLf
         '2020/7/8 END
         strContent = strContent & "申 請 人：" & strCuName & vbCrLf
         'Add By Sindy 2024/7/16
         'Modify by Amy 2024/07/17 原:Not IsNull(strTM78)
         If strTM78 <> MsgText(601) Then
            strContent = strContent & "　　　　（多人申請）" & vbCrLf
         End If
         '2024/7/16 END
         strContent = strContent & "商 標：" & strTM05 & vbCrLf
         If strType <> "S" Then strContent = strContent & "類 別：" & strTM09 & vbCrLf
         If strTM15 <> "" Then
            strContent = strContent & "註冊號數：" & strTM15 & vbCrLf
         Else
           strContent = strContent & "申請案號：" & strTM12 & vbCrLf
         End If
      End If
   End If
   strContent = strContent & vbCrLf & "請參附件！" & vbCrLf & vbCrLf
   strDate = ""
   'Added by Sindy 2019/12/6
   If strType = "T" Then
   '2019/12/6 END
      'Add By Sindy 2020/10/14
      If m_bolTLetter = True Then '指示信
         '請依指示信提申
         strContent = strContent & "1.請依指示信提出網上申請。" & vbCrLf _
                                 & "2.收到本信請回覆。" & vbCrLf
      ElseIf m_ET01 = "03" Then '核准
         If strCP10 = "101" Then '申請案的1001核准
            If strTM13 <> "" Then
               strDate = CompDate(2, -7, CompDate(1, 2, strTM13))
               strDate = Left(strDate, 4) & "年" & Mid(strDate, 5, 2) & "月" & Right(strDate, 2) & "日"
            Else
               strDate = "    年   月   日"
            End If
            'Modify By Sindy 2022/12/22 + 另因智慧局自 2023 年 1 月起開放申請領取電子商標註冊證...
            'Modify By Sindy 2025/8/8 字體加大加粗
            strContent = strContent & "請於" & strDate & "以前告知是否繳納註冊費！" & vbCrLf & _
                                      "收到此函請回覆！" & vbCrLf & vbCrLf & _
                                      "<font style=""font-size:16pt""><b>另因智慧局自 2023 年 1 月起開放申請領取電子商標註冊證，申請人可選擇註冊證發給形式，故請一併告知欲申請電子註冊證或紙本註冊證。&nbsp;</b></font>" & vbCrLf
         Else '非申請案的1001核准
            'strContent = strContent & "文件俟核准函再一併郵寄！收到此函請回覆！"
            strContent = strContent & "文件不再郵寄，收到此函請回覆！" 'Modify By Sindy 2018/5/17
            'Add By Sindy 2016/8/15 芸如:再依案件性質出現不同內文
            'Modify By Sindy 2019/2/12 Mark,不顯示
   '         If strCP10 = "102" Or strCP10 = "301" Or strCP10 = "501" Then
   '            strContent = strContent & vbCrLf & "若仍需郵寄紙本請告知！" '若需現在郵寄紙本請告知！
   '         End If
            '2016/8/15 END
         End If
      ElseIf m_ET01 = "10" Then '業務員期限管制表
         strContent = strContent & "是否辦理續展請告知，收到此函請回覆！"
      
      'Add By Sindy 2020/2/24
      ElseIf m_ET01 = "15" Then '智慧局註冊費通知函
         strContent = strContent & "是否辦理" & strCP10Nm & "請告知，收到此函請回覆！"
         strCP14 = strUserNum 'Add By Sindy 2024/10/28 預設操作人員
         
      'Added by Lydia 2016/12/22 指定1728收款寄證-內文
      ElseIf m_ET01 = "1728" Then
           'Modified by Lydia 2020/08/14 改變內文
           'strContent = ""
           strContent = "您好：" & vbCrLf & "有關"
           strContent = strContent & strCuName & "的「" & strTM05 & "」商標" & IIf(strCP10 = "1701", "註冊申請案", strCP43m & "案")
           strContent = strContent & "(Our Ref:" & strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 = "000", "", IIf(strCP03 <> "0", "-" & strCP03, IIf(strCP04 <> "00", "-" & strCP04, "")))
           strContent = strContent & IIf(strTM45 <> "", ";Your Ref:" & strTM45, "") & ")，本所已收到智慧局核發之"
           If strTM15 <> "" Then
              strContent = strContent & "註冊第" & Val(strTM15) & "號"
           Else
              strContent = strContent & "申請號" & Val(strTM12) & "號"
           End If
           strContent = strContent & "商標" & IIf(strCP10 = "1701", strCP10Nm, strCP43m & strCP10Nm & "函")
           'Modified by Lydia 2017/02/18  strA1kdata第1碼區分內文
           'strContent = strContent & "，煩請儘速將" & strA1kdata
           'strContent = strContent & "，匯至本所帳戶，收到款項後，本所會將" & IIf(strCP10 = "1701", "註冊證", strCP10Nm & "函")
           'strContent = strContent & "郵寄予 貴公司。" & vbCrLf & vbCrLf
           'Modified by Lydia 2020/08/14 改變內文
           'strContent = strContent & "，煩請儘速將" & Mid(strA1kdata, 2)
           strContent = strContent & "，由於會計作業關係，尚祈儘速將" & Mid(strA1kdata, 2)
           If Mid(strA1kdata, 1, 1) = "Y" Then '特定客戶-直接寄紙本
              'Modified by Morgan 2023/2/3
              'strContent = strContent & "，匯至本所帳戶，本所將於" & ChangeTStringToTDateString(strSrvDate(2)) & "郵寄" & IIf(strCP10 = "1701", "註冊證", strCP10Nm & "函")
              strContent = strContent & "，匯至本所帳戶，本所將於" & ChangeTStringToTDateString(strSrvDate(2)) & "寄送" & IIf(strCP10 = "1701", "註冊證", strCP10Nm & "函")
              'end 2023/2/3
              strContent = strContent & "予 貴公司，請查收。" & vbCrLf & vbCrLf
           Else                             '只發mail-先不寄紙本
              'Modified by Lydia 2020/08/14 改變內文
              'strContent = strContent & "，匯至本所帳戶，收到款項後，本所會將" & IIf(strCP10 = "1701", "註冊證", strCP10Nm & "函")
              'strContent = strContent & "郵寄予 貴公司。" & vbCrLf & vbCrLf
              strContent = strContent & "，匯至本所帳戶，收到款項後，即會安排進入" & IIf(strCP10 = "1701", "註冊證", strCP10Nm & "函")
              'Modified by Morgan 2023/2/3
              'strContent = strContent & "郵寄流程。" & vbCrLf & vbCrLf
              strContent = strContent & "寄送流程。" & vbCrLf & vbCrLf
              'end 2023/2/3
              'end 2020/08/14
           End If
           'end 2017/02/18
           'Modified by Lydia 2020/08/14 改變內文
           'strContent = strContent & "匯款後請將匯款憑證EMAIL予本所，否則本所不知　貴公司已匯款。" & vbCrLf
           strContent = strContent & "匯款後，煩請將匯款憑證EMAIL予本所，以利本所會計核對。謝謝！" & vbCrLf
           If strAttach <> "" Then strContent = strContent & vbCrLf & "請參附件！" & vbCrLf & vbCrLf
      'end 2016/12/22
      
      '(商申大宗) 除4,6字頭和202.申請意見書則為商爭
      ElseIf strLD10 <> "" And _
         ((Len(strCP10) = 3 And Not (Left(strCP10, 1) = "4" Or Left(strCP10, 1) = "6" Or strCP10 = "202")) Or _
          (Len(strCP10) = 4 And Not (Left(strCP10, 2) = "14" Or Left(strCP10, 2) = "16")) _
         ) Then
         strContent = strContent & "文件俟"
         'Add By Sindy 2017/3/2 桂英:再依案件性質出現不同內文
         '102.延展,301.變更,501.移轉,502.授權
         If strCP10 = "102" Or strCP10 = "301" Or strCP10 = "501" Or strCP10 = "502" Then
            strContent = strContent & "核准函"
         Else
            strContent = strContent & "註冊證"
         End If
         strContent = strContent & "再一併郵寄，收到此函請回覆！"
         '2017/3/2 END
         'Add By Sindy 2016/8/15 芸如:再依案件性質出現不同內文
         '101.申請,717.註冊費,201.補正,206.放棄專用權,211.檢送同意書,303.延期
         If strCP10 = "101" Or strCP10 = "717" Or strCP10 = "201" Or _
            strCP10 = "206" Or strCP10 = "211" Or strCP10 = "303" Then
            strContent = strContent & vbCrLf & "若仍需郵寄紙本請告知！" '若需現在郵寄紙本請告知！
         End If
         '2016/8/15 END
      
      'Add By Sindy 2020/1/30
      'Modify By Sindy 2021/3/3 + Left(strCP09, 1) = "C"
      ElseIf Left(strCP09, 1) = "C" And Pub_StrUserSt03 <> "P22" Then 'C類承辦人
         If Val(strCP06) > 0 Then
            strDate = Left(strCP06, 4) & "年" & Mid(strCP06, 5, 2) & "月" & Right(strCP06, 2) & "日"
            strContent = strContent & "請於" & strDate & "以前告知欲如何續行。" & vbCrLf & vbCrLf
         End If
         strContent = strContent & "紙本文件不再郵寄，收到此函請回覆！"
      '2020/1/30 END
      
      Else '其他
         strContent = strContent & "文件不再郵寄，收到此函請回覆！"
      End If
   Else
      strContent = strContent & "文件不再郵寄，收到此函請回覆！"
   End If
   strContent = strContent & vbCrLf & vbCrLf & vbCrLf
   'Modify By Sindy 2020/1/13 Mark,改圖片簽名檔
   'Modify By Sindy 2020/1/30 非商標處程序
   If Pub_StrUserSt03 <> "P22" Then
      'Modify By Sindy 2020/3/26 "事務所名稱：台一國際專利法律事務所" => CompNameQuery("2")
      If strSrvDate(1) >= 事務所合併日 Then
         strCompName = CompNameQuery("2")
      Else
         strCompName = "台一國際專利商標事務所"
      End If
      '2020/3/26 END
      strContent = strContent & "總經理 林景郁" & vbCrLf & _
                                "商標部 " & strUserName & vbCrLf '& _
'                                strCompName & vbCrLf & _
'                                "Tai E International Patent & Law Office" & vbCrLf & _
'                                "地址:104 台灣台北市長安東路二段112號9樓" & vbCrLf & _
'                                "TEL: 886 2 25061023" & vbCrLf & _
'                                "FAX: 886 2 25011666" & vbCrLf & _
'                                "Email: tm@taie.com.tw" & vbCrLf & _
'                                "Website: www.taie.com.tw" & vbCrLf & vbCrLf
'      strContent = strContent & "********************保密警語********************" & vbCrLf & _
'                             "本信件僅授權於指定之收信人取閱之用，信件中可能含有機密性資訊。" & vbCrLf & _
'                             "如果您並非被指定之收信人，任何未經授權而擅自使用此信件所含之機密資訊的行為是被嚴格禁止的。" & vbCrLf & _
'                             "如果您在任何未經授權的情形之下收到本信件，煩請您立即告知原發信人並將此信件回傳至以上地址。" & vbCrLf & _
'                             "謝謝您的合作。"

'   strContent = strContent & "總經理" & vbCrLf & _
'                             "林景郁" & vbCrLf & _
'                             strCompName & vbCrLf & _
'                             "Tai E International Patent & Law Office" & vbCrLf & _
'                             "104 台灣台北市長安東路2段112號9樓" & vbCrLf & _
'                             "TEL: 886 2 25061023 ext 321" & vbCrLf & _
'                             "FAX: 886 2 25011666" & vbCrLf & _
'                             "Email:tm@taie.com.tw; lawoffice@taie.com.tw" & vbCrLf & _
'                             "URL: https://www.taie.com.tw" & vbCrLf & vbCrLf
'   frm880019.txtContent = strContent & "********************保密警語********************" & vbCrLf & _
'                          "本信件僅授權於指定之收信人取閱之用，信件中可能含有機密性資訊。" & vbCrLf & _
'                          "如果您並非被指定之收信人，任何未經授權而擅自使用此信件所含之機密資訊的行為是被嚴格禁止的。" & vbCrLf & _
'                          "如果您在任何未經授權的情形之下收到本信件，煩請您立即告知原發信人並將此信件回傳至以上地址。" & vbCrLf & _
'                          "謝謝您的合作。"
   End If
   frm880019.txtContent = strContent
   '2020/1/30 END
   '2020/1/13 END
   
   '加附件
'   If strAttach <> "" Then
'      ArrStr = Split(strAttach, ";")
'      For ii = 0 To UBound(ArrStr)
'         If Trim(ArrStr(ii)) <> "" Then  'Added by Lydia 2017/03/02 判斷非空白
'            objMail.Attachments.add ArrStr(ii)
'         End If                          'end 2017/03/02
'      Next ii
'   End If
'   '附件
'   m_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum 'Add By Sindy 2017/1/6 以防止上面寄信時有些檔案會被咬住,後面刪檔會有權限問題
'   KillAttach 'Add By Sindy 2017/3/10
'   bolSelFile = False
'   pFiles = ""
'   For ii = 0 To lstAtt(0).ListCount - 1
'      If lstAtt(0).Selected(ii) Then
'         bolSelFile = True
'         stFileName = lstAtt(0).List(ii)
'         If InStrRev(stFileName, " (") > 0 Then
'            stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
'         End If
'         If InStr(stFileName, "\") = 0 Then
'            If GetAttachFile(stFileName, CInt(m_AttEEP02)) = False Then Exit Sub
'         End If
'         pFiles = pFiles & ";" & stFileName
'      End If
'   Next ii
'   If bolSelFile = False Then
'      Call DownloadAllAttachFile(CInt(m_AttEEP02), 0, pFiles)
'   Else
'      If pFiles <> "" Then pFiles = Mid(pFiles, 2)
'   End If
   'Add By Sindy 2020/1/13
   '商標處經理
   strCon1 = "select st01 from staff" & _
               " where st03='P20' and st20='41' and st04='1'"
   intQ = 1
   Set adoRst = ClsLawReadRstMsg(intQ, strCon1)
   strP20Mgr = ""
   If intQ = 1 Then
      strP20Mgr = "" & adoRst.Fields("st01")
   End If
   '2020/1/13 END
   
'   'Add By Sindy 2020/11/26
'   'MCTF的717.註冊費要夾帶商標註冊費繳費單
'   If strTM10 = "000" And strCP10 = "717" Then
'      'MCTF
'      salesArea = GetCuSales(pCustNo, salesNo)
'      If Mid(PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04), 1, 4) = "MCTF" Or _
'         salesNo = "96029" Or salesNo = "96030" Then
'         Call PUB_PrintTFeeForm(strCP01, strCP02, strCP03, strCP04, , , False, strTFeeFile)
'         strAttach = IIf(strAttach <> "", strAttach & ";" & strTFeeFile, strTFeeFile)
'      End If
'   End If
'   '2020/11/26 END
   
   frm880019.SetAttach strAttach
   'Modify By Sindy 2020/1/13 1.副本:江協理(98020) 2.密件副本:林經理
   'Modify By Sindy 2020/10/13 +  & IIf(strCP14 <> "", ";" & strCP14, "")
   frm880019.SetEmail "", "", strTo, , True, "98020", IIf(strP20Mgr <> "", ";" & strP20Mgr, "") & IIf(strCP14 <> "", ";" & strCP14, "")
   frm880019.cmdAttach.Visible = True 'False
   frm880019.m_bolAttFromCpp = m_bolAttFromCpp 'Add By Sindy 2020/1/31
   frm880019.SetParent oForm
   frm880019.m_bolTLetter = m_bolTLetter 'Add By Sindy 2020/7/21
   frm880019.Show vbModal
   PUB_SettingTeMail = frm880019.m_bolDone
   Unload frm880019
   '*************************************************************************************************
   'm_AttachPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") 'Add By Sindy 2017/1/6 以防止上面寄信時有些檔案會被咬住,後面刪檔會有權限問題
   If PUB_SettingTeMail = True Then '寄信成功
      '寄件日期
      Dim rsTmp As New ADODB.Recordset
      
      strSql = "Select *" & _
               " From smailbackup" & _
               " Where smb01='" & strCP09 & "'" & _
               " order by smb02 desc,smb03 desc"
      rsTmp.CursorLocation = adUseClient
      rsTmp.Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If TypeName(lblSendData) <> "Nothing" Then lblSendData.Visible = False
      If rsTmp.RecordCount > 0 Then
         If TypeName(lblSendData) <> "Nothing" Then lblSendData.Visible = True
         m_strSendDate = TAIWANDATE(rsTmp.Fields("smb02"))
         m_strSendTime = rsTmp.Fields("smb03")
         If TypeName(lblSendData) <> "Nothing" Then
            lblSendData.Caption = "寄件日期:" & Format(m_strSendDate, "###/##/##") & " " & Format(m_strSendTime, "##:##:##")
         End If
      End If
      
      rsTmp.Close
      Set rsTmp = Nothing
   End If
   Set adoRst = Nothing
   Exit Function
   '2018/5/11 END
   
   
'   '呼叫新郵件：
'   Set objOutLook = CreateObject("Outlook.Application")
'   If Dir(strTemplatePath) <> "" Then
'      Set objMail = objOutLook.CreateItemFromTemplate(strTemplatePath)
'   Else
'      Set objMail = objOutLook.CreateItem(0)
'   End If
'   '副本.cc
'   '收件者.To
'   'Modified by Lydia 2016/12/22 指定1728收款寄證,收件者為代理人
'   'strTo = PUB_GetFCeMailConText("Main_EMail", strCP01, strCP02, strCP03, strCP04, , strCP10)
'   If m_ET01 = "1728" Then
'       strTo = PUB_GetFCeMailConText("Main_EMail", strCP01, strCP02, strCP03, strCP04, "FC")
'   Else
'       strTo = PUB_GetFCeMailConText("Main_EMail", strCP01, strCP02, strCP03, strCP04, , strCP10)
'   End If
'   'end 2016/12/22
'   objMail.To = strTo
'   '非E化案件或無代表號信箱時, 顯示訊息, 若同時發生時只顯示一次訊息即可.
'   If PUB_GetFCeMailConText("IsECase", strCP01, strCP02, strCP03, strCP04) = "Y" Then bolIsECase = True
'   strMsg = ""
'   If strTo = "" Or bolIsECase = False Then
'      If bolIsECase = False Then
'         strMsg = "非E化案件"
'      End If
'      If strTo = "" Then
'         If strMsg <> "" Then strMsg = strMsg & "且"
'         strMsg = strMsg & "無代表號信箱"
'      End If
'      MsgBox strMsg, vbInformation
'   End If
'
'   '密件副本.BCC
'   '主旨.Subject
'   '主旨：(A)AEGIS(第03,42,0128類)商標，核准通知
'   '      (國外專業代號)商標名稱(第類)商標，案件性質
'   'Modified by Lydia 2016/12/22 指定1728收款寄證,主旨:客戶名稱發xx通知(本所案號)
'   'objMail.Subject = "(" & strST17 & ")" & strTM05 & "(第" & strTM09 & "類)商標，" & strCP10Nm
'   If m_ET01 = "1728" Then
'       'Modified by Lydia 2017/01/04 +專業代號 "(" & strST17 & ")"
'       objMail.Subject = "(" & strST17 & ")" & strCuName & "發" & strCP43m & strCP10Nm & "通知(" & strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 = "000", "", IIf(strCP03 <> "0", "-" & strCP03, IIf(strCP04 <> "00", "-" & strCP04, ""))) & ")"
'   ElseIf m_ET01 = "10" Then '業務員期限管制表
'       objMail.Subject = "(" & strST17 & ")" & strTM05 & "(第" & strTM09 & "類)商標，通知續展"
'   Else
'       objMail.Subject = "(" & strST17 & ")" & strTM05 & "(第" & strTM09 & "類)商標，" & strCP10Nm
'   End If
'   'end 2016/12/22
'
'   '加附件
'   If strAttach <> "" Then
'      ArrStr = Split(strAttach, ";")
'      For ii = 0 To UBound(ArrStr)
'         If Trim(ArrStr(ii)) <> "" Then  'Added by Lydia 2017/03/02 判斷非空白
'            objMail.Attachments.add ArrStr(ii)
'         End If                          'end 2017/03/02
'      Next ii
'   End If
'   '內文.Body
'   strContent = strContent & "貴方卷號：" & strTM45 & vbCrLf
'   strContent = strContent & "我方案號：" & strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 = "000", "", IIf(strCP03 <> "0", "-" & strCP03, IIf(strCP04 <> "00", "-" & strCP04, ""))) & vbCrLf & vbCrLf
'   strContent = strContent & "申 請 人：" & strCuName & vbCrLf
'   strContent = strContent & "商 標：" & strTM05 & vbCrLf
'   strContent = strContent & "類 別：" & strTM09 & vbCrLf
'   If strTM15 <> "" Then
'      strContent = strContent & "註冊號數：" & strTM15 & vbCrLf
'   Else
'      strContent = strContent & "申請案號：" & strTM12 & vbCrLf
'   End If
'   strContent = strContent & vbCrLf & "請參附件！" & vbCrLf & vbCrLf
'   strContent = strContent & "<FONT color=#ff0000>"
'   strDate = ""
'   If m_ET01 = "03" Then '核准
'      If strCP10 = "101" Then '申請案的1001核准
'         If strTM13 <> "" Then
'            strDate = CompDate(2, -7, CompDate(1, 2, strTM13))
'            strDate = Left(strDate, 4) & "年" & Mid(strDate, 5, 2) & "月" & Right(strDate, 2) & "日"
'         Else
'            strDate = "    年   月   日"
'         End If
'         strContent = strContent & "請於" & strDate & "以前告知是否繳納註冊費！" & vbCrLf & _
'                                   "收到此函請回覆！"
'      Else '非申請案的1001核准
'         'strContent = strContent & "文件俟核准函再一併郵寄！收到此函請回覆！"
'         strContent = strContent & "文件不再郵寄！收到此函請回覆！" 'Modify By Sindy 2018/5/17
'         'Add By Sindy 2016/8/15 芸如:再依案件性質出現不同內文
'         If strCP10 = "102" Or strCP10 = "301" Or strCP10 = "501" Then
'            strContent = strContent & vbCrLf & "若仍需郵寄紙本請告知！" '若需現在郵寄紙本請告知！
'         End If
'         '2016/8/15 END
'      End If
'   ElseIf m_ET01 = "10" Then '業務員期限管制表
'      strContent = strContent & "是否辦理續展請告知，收到此函請回覆！"
'
'   'Added by Lydia 2016/12/22 指定1728收款寄證-內文
'   ElseIf m_ET01 = "1728" Then
'        strContent = ""
'        strContent = strContent & strCuName & "的「" & strTM05 & "」商標" & IIf(strCP10 = "1701", "註冊申請案", strCP43m & "案")
'        strContent = strContent & "(Our Ref:" & strCP01 & "-" & strCP02 & IIf(strCP03 & strCP04 = "000", "", IIf(strCP03 <> "0", "-" & strCP03, IIf(strCP04 <> "00", "-" & strCP04, "")))
'        strContent = strContent & IIf(strTM45 <> "", ";Your Ref:" & strTM45, "") & ")，本所已收到智慧局核發之"
'        If strTM15 <> "" Then
'           strContent = strContent & "註冊第" & Val(strTM15) & "號"
'        Else
'           strContent = strContent & "申請號" & Val(strTM12) & "號"
'        End If
'        strContent = strContent & "商標" & IIf(strCP10 = "1701", strCP10Nm, strCP43m & strCP10Nm & "函")
'        'Modified by Lydia 2017/02/18  strA1kdata第1碼區分內文
'        'strContent = strContent & "，煩請儘速將" & strA1kdata
'        'strContent = strContent & "，匯至本所帳戶，收到款項後，本所會將" & IIf(strCP10 = "1701", "註冊證", strCP10Nm & "函")
'        'strContent = strContent & "郵寄予 貴公司。" & vbCrLf & vbCrLf
'        strContent = strContent & "，煩請儘速將" & Mid(strA1kdata, 2)
'        If Mid(strA1kdata, 1, 1) = "Y" Then '特定客戶-直接寄紙本
'           strContent = strContent & "，匯至本所帳戶，本所將於" & ChangeTStringToTDateString(strSrvDate(2)) & "郵寄" & IIf(strCP10 = "1701", "註冊證", strCP10Nm & "函")
'           strContent = strContent & "予 貴公司，請查收。" & vbCrLf & vbCrLf
'        Else                             '只發mail-先不寄紙本
'           strContent = strContent & "，匯至本所帳戶，收到款項後，本所會將" & IIf(strCP10 = "1701", "註冊證", strCP10Nm & "函")
'           strContent = strContent & "郵寄予 貴公司。" & vbCrLf & vbCrLf
'        End If
'        'end 2017/02/18
'        strContent = strContent & "匯款後請將匯款憑證EMAIL予本所，否則本所不知　貴公司已匯款。" & vbCrLf
'        If strAttach <> "" Then strContent = strContent & vbCrLf & "請參附件！" & vbCrLf & vbCrLf
'        strContent = strContent & "<FONT color=#ff0000>"
'   'end 2016/12/22
'
'   '(商申大宗) 除4,6字頭和202.申請意見書則為商爭
'   ElseIf (Len(strCP10) = 3 And Not (Left(strCP10, 1) = "4" Or Left(strCP10, 1) = "6" Or strCP10 = "202")) _
'       Or (Len(strCP10) = 4 And Not (Left(strCP10, 2) = "14" Or Left(strCP10, 2) = "16")) Then
'      strContent = strContent & "文件俟"
'      'Add By Sindy 2017/3/2 桂英:再依案件性質出現不同內文
'      '102.延展,301.變更,501.移轉,502.授權
'      If strCP10 = "102" Or strCP10 = "301" Or strCP10 = "501" Or strCP10 = "502" Then
'         strContent = strContent & "核准函"
'      Else
'         strContent = strContent & "註冊證"
'      End If
'      strContent = strContent & "再一併郵寄！收到此函請回覆！"
'      '2017/3/2 END
'      'Add By Sindy 2016/8/15 芸如:再依案件性質出現不同內文
'      '101.申請,717.註冊費,201.補正,206.放棄專用權,211.檢送同意書,303.延期
'      If strCP10 = "101" Or strCP10 = "717" Or strCP10 = "201" Or _
'         strCP10 = "206" Or strCP10 = "211" Or strCP10 = "303" Then
'         strContent = strContent & vbCrLf & "若仍需郵寄紙本請告知！" '若需現在郵寄紙本請告知！
'      End If
'      '2016/8/15 END
'   Else '其他
'      strContent = strContent & "文件不再郵寄，收到此函請回覆！"
'   End If
'   strContent = strContent & "</FONT>"
'   'strContent = strContent & vbCrLf & vbCrLf & vbCrLf
'   strContent = Replace(strContent, "新細明體", "Times New Roman")
'   strContent = Replace(strContent, vbCrLf, "<BR>")
'   strContent = Replace(strContent, "  ", "&nbsp;&nbsp;")
'   objMail.HTMLBody = "<FONT FACE=""Times New Roman"">" & strContent & "<BR>" & objMail.HTMLBody & "</FONT>"
'   objMail.Display
'
'   Set objMail = Nothing
'   Set objOutLook = Nothing
'   Set adoRst = Nothing
End Function

'Add By Sindy 2020/6/20 商標處發指示信
Public Function PUB_T_AppFormSendMail(m_EEP01 As String, m_RetrunRecv As String, _
   strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String, _
   strCP10 As String, oForm As Form, Optional lstAtt As ListBox) As Boolean
   
Dim bolHadFile As Boolean
Dim pbolDone As Boolean
Dim pFiles As String
Dim stFileName As String
Dim rsA As New ADODB.Recordset
Dim stConCpp As String
Dim ii As Integer
Dim stAttPath As String
Dim varTmp As Variant, strSendDate As String, strSendTime As String
Dim intQ As Integer, strCon1 As String 'Added by Lydia 2024/03/29

   pFiles = ""
   PUB_T_AppFormSendMail = False
   Screen.MousePointer = vbHourglass
   
   '附件:
   If Not lstAtt Is Nothing Then
      '沒點選附件,就預設是全部附件
      'Modify By Sindy 2018/11/15 不管附件的選取狀況,一律全部帶入
   '   bolHadFile = False
   '   For ii = 0 To lstAtt(0).ListCount - 1
   '      If lstAtt(0).Selected(ii) Then
   '         bolHadFile = True
   '         Exit For
   '      End If
   '   Next ii
   '   If bolHadFile = False Then
         For ii = 0 To lstAtt.ListCount - 1
            If InStr(UCase(lstAtt.List(ii)), UCase(".pdf (")) > 0 Then
               lstAtt.Selected(ii) = True
            End If
         Next ii
   '   End If
      '1.先讀取附件區的PDF
      For ii = 0 To lstAtt.ListCount - 1
         If lstAtt.Selected(ii) Then
            stFileName = lstAtt.List(ii)
            If InStrRev(stFileName, " (") > 0 Then
               'Add By Sindy 2021/8/6 排除 C:\Program Files (x86) 狀況
               If UCase(Mid(stFileName, InStrRev(stFileName, " (") + 1, Len("(X86)"))) <> "(X86)" Then
               '2021/8/6 END
                  stFileName = Left(stFileName, InStrRev(stFileName, " (") - 1)
               End If
            End If
            '不含CDATA簡易申請書
            If UCase(Right(stFileName, 4)) = ".PDF" And InStr(UCase(stFileName), ".CDATA.") = 0 Then
               pFiles = pFiles & ";" & stFileName
            End If
         End If
      Next ii
   End If
   
   '2.下載其他文號的PDF
   'Modify By Sindy 2021\5\19 +  & "\otherFile"
   stAttPath = App.path & Pub_GetSpecMan("EmpFlowAttPath") & "\" & strUserNum & "\otherFile"
   If Dir(stAttPath & "\.") <> "" Then
      Kill stAttPath & "\*.*"
   End If
'   'and substr(upper(cpp02),length('.'||cp10||'.PDF')*-1)=upper('.'||cp10||'.PDF') : 官方來函
'   '要排除的附件 altr, worksheet, info, cust.pdf, reply.pdf
'   'Modify By Sindy 2017/6/14 cp01=efc01(+) ==> instr(cp01||',ALL',efc01(+))>0
'   'Modified by Morgan 2018/9/6 +剔除 altr,worksheet
'   'modify by sonia 2018/12/22 '.INFO.'改為'.INFO',否則CFP-028921之INFO1及INFO2不會剔除
'   'stConCpp = " and INSTR(lower(CPP02),'.altr.')=0 and INSTR(lower(CPP02),'.worksheet.')=0 and INSTR(lower(CPP02),'.info.')=0 and instr(lower(cpp02),'.cust.pdf')=0  and instr(lower(cpp02),'.reply.pdf')=0"
'   'Modified by Morgan 2019/1/22 +剔除 ack.收達
'   '下載卷宗區
'   stConCpp = " and INSTR(lower(CPP02),'.altr.')=0 and INSTR(lower(CPP02),'.worksheet.')=0" & _
'              " and INSTR(lower(CPP02),'.ack.')=0 and INSTR(lower(CPP02),'.info')=0" & _
'              " and instr(lower(cpp02),'.cust.pdf')=0 and instr(lower(cpp02),'.reply.pdf')=0 and instr(lower(cpp02),'.order.')=0"
'   strcon1 = "select * from casepaperpdf,caseprogress" & _
'      " where cpp01 in('" & Replace(m_RetrunRecv, ",", "','") & "') and cpp01<>'" & m_EEP01 & "' and cpp01=cp09" & _
'      " and substr(upper(cpp02),length('.PDF')*-1)=upper('.PDF')" & stConCpp
'   intq = 1
'   Set rsA = ClsLawReadRstMsg(intq, strcon1)
'   If intq = 1 Then
'      rsA.MoveFirst
'      Do While Not rsA.EOF
'         stFileName = "" & rsA.Fields("cpp02")
'         If PUB_GetAttachFile_CPP(lblCP09.Caption, stFileName, m_AttachPath) = False Then
'            MsgBox "卷宗區PDF檔下載失敗！" & vbCrLf & _
'               "（" & m_AttachPath & "\" & stFileName & "）", vbCritical
'            GoTo ErrHnd
'         Else
'            pFiles = pFiles & ";" & stFileName
'         End If
'         rsA.MoveNext
'      Loop
'   End If
'   rsA.Close
   '下載歷程附件區
   strCon1 = "SELECT * FROM empelectronfile" & _
            " where eef01 in('" & Replace(m_RetrunRecv, ",", "','") & "') and eef01<>'" & m_EEP01 & "'" & _
            " AND substr(upper(eef03),LENGTH('.PDF')*-1)=upper('.PDF')" & _
            " AND eef02 IN(SELECT MAX(eep02) FROM empelectronprocess" & _
                         " WHERE eep01=eef01 AND eep04 IN('" & EMP_送件 & "','" & EMP_退件重送 & "'))"
   intQ = 1
   Set rsA = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      rsA.MoveFirst
      Do While Not rsA.EOF
         stFileName = "" & rsA.Fields("eef03")
         If PUB_GetAttachFile_EEF(rsA.Fields("eef01"), rsA.Fields("eef02"), stFileName, stAttPath) = False Then
            MsgBox "歷程附件區PDF檔下載失敗！" & vbCrLf & _
               "（" & stAttPath & "\" & stFileName & "）", vbCritical
            GoTo ErrHnd
         Else
            pFiles = pFiles & ";" & stFileName
         End If
         rsA.MoveNext
      Loop
   End If
   rsA.Close
   If pFiles <> "" Then pFiles = Mid(pFiles, 2)
   Screen.MousePointer = vbDefault
   
   '******
   pbolDone = PUB_SettingTeMail(oForm, "", strCP01, strCP02, strCP03, strCP04, _
                                pFiles, strCP10, m_EEP01, , , , , , , m_RetrunRecv, _
                                strSendDate, strSendTime, True)
   '******
   If pbolDone = False Then '寄信失敗
      MsgBox "寄信失敗！" & IIf(Err.Number > 0, vbCrLf & Err.Description, ""), vbCritical
   Else
      '指示信記錄
      strCon1 = "delete from AppForm where AF01 in('" & Replace(m_RetrunRecv, ",", "','") & "')"
      cnnConnection.Execute strCon1, intQ
      varTmp = Split(m_RetrunRecv, ",")
      For ii = 0 To UBound(varTmp)
         'Modify By Sindy 2024/8/29 strSendDate: +DBDATE() 轉換為西元日期
         strCon1 = "insert into AppForm(AF01,AF02,AF03,AF06,AF07,AF08,AF11,AF12,AF14)" & _
                  " values('" & varTmp(ii) & "','" & strUserNum & "',to_char(sysdate,'yyyymmdd')" & _
                  ",'" & strUserNum & "',to_char(sysdate,'yyyymmdd'),to_number(to_char(sysdate,'HH24MIss'))" & _
                  "," & DBDATE(strSendDate) & "," & strSendTime & ",'" & strUserNum & "')"
         cnnConnection.Execute strCon1, intQ
      Next ii
      
      PUB_T_AppFormSendMail = True
   End If
   
   Set rsA = Nothing
   Exit Function
   
ErrHnd:
   Set rsA = Nothing
   If Err.Number = 70 Then
      MsgBox ChgSQL(stFileName) & "檔案已開啟！", vbCritical
   ElseIf Err.Number > 0 Then
      MsgBox Err.Description, vbCritical
   End If
End Function

'Add By Sindy 2014/9/15
'外專發FC郵件資訊
'Modify By Sindy 2015/1/8 加Optional strCP10 As String = "" ex:FCP-036408的年費
'Modify By Sindy 2019/5/23 + Optional ByVal StrProcCP10 As String = ""
Public Function PUB_SettingFCeMail(ByVal pDeptID As String, ByVal strTemplatePath As String, ByVal EMailType As String, _
               ByVal strCP01 As String, ByVal strCP02 As String, ByVal strCP03 As String, ByVal strCP04 As String, _
               ByVal strContent As String, Optional ByVal strAttach As String = "", Optional ByVal strCP10 As String = "", _
               Optional ByVal m_ET01 As String = "", Optional ByVal m_ET03 As String = "", _
               Optional strFCorCF As String = "", Optional ByVal StrProcCP10 As String = "") As Boolean
Dim adoRst As ADODB.Recordset
Dim objOutLook As Object
Dim objMail As Object
Dim strCP14 As String, strCP14sir As String, SirEname As String, strCP14_st17 As String
Dim SirSt17 As String, strMsg As String
Dim strCaseNo As String, strYourRef As String, strOurRef As String
Dim strTo As String
Dim bolIsECase As Boolean, strSignature As String
Dim ArrStr() As String, ii As Integer
Dim m_Encls As String 'Add By Sindy 2015/5/29
Dim StrText4 As String 'Add By Sindy 2019/8/5
Dim strText5 As String 'Add By Sindy 2019/8/5
Dim strST17 As String 'Add By Sindy 2022/11/30
Dim intQ As Integer, strCon1 As String 'Added by Lydia 2024/03/29

   EMailType = UCase(EMailType)
   
   'Add By Sindy 2020/9/8 特殊申請人的出名人簽名檔
   'If pDeptID = "F23" Then 'Add By Sindy 2020/9/23 承辦才Run
   If InStr(strTemplatePath, "TOT-000F23") > 0 Then 'Modify By Sindy 2021/9/22 Run承辦的樣本,才檢查
      If PUB_SpecAPPLOutAgent_FCP(strCP01, strCP02, strCP03, strCP04) = True Then
         strTemplatePath = Mid(strTemplatePath, 1, InStr(strTemplatePath, "$$") - 1) & "$$TOT-000F23-0-04.oft"
      End If
   End If
   '2020/9/8 END
   
   strCP14 = "": strCP14sir = "": SirEname = "": strCP14_st17 = ""
   If pDeptID = "F22" Then
      'Modify By Sindy 2019/8/6
      If InStr(strTemplatePath, "TOT-000F23") > 0 Then
         '目前智權人員
         If strCP01 = "FCP" Or strCP01 = "FG" Then
            strTo = PUB_GetFCPSalesNo(strCP01, strCP02, strCP03, strCP04)
         Else
            strTo = PUB_GetAKindSalesNo(strCP01, strCP02, strCP03, strCP04)
         End If
         '依目前智權人員判斷其組別
         If PUB_GetStaffST16(strTo) = "2" Then '日文
            'Add By Sindy 2019/8/8
            strCon1 = "select oman,st12 from setspecman,staff where OCODE='S' and oman=st01(+)"
            intQ = 1
            Set adoRst = ClsLawReadRstMsg(intQ, strCon1)
            If intQ = 1 Then
               strCP14sir = "" & adoRst.Fields("oman")
               SirEname = "" & adoRst.Fields("st12") '主管的英文名稱
            End If
            '2019/8/8 END
            StrText4 = "Division Executive"
            strText5 = "Japanese Patent Division"
         Else
            StrText4 = "Assistant Manager"
            strText5 = "International Patent Division"
         End If
         
      Else
      '2019/8/6 END
         
         'Add By Sindy 2021/9/22 特殊申請人的出名人簽名檔
         If PUB_SpecAPPLOutAgent_FCP(strCP01, strCP02, strCP03, strCP04) = True Then
            strTemplatePath = Mid(strTemplatePath, 1, InStr(strTemplatePath, "$$") - 1) & "$$TOT-000F22-0-02.oft"
         End If
         '2021/9/22 END
         
         '承辦工程師及他的主管
         'Modify By Sindy 2018/11/5 對客戶報告(所有以email發出信函)之密件副本
         '　　不必再密件副本給我，機電同仁信件之密件副本一律改Akira、化學同仁一律改Owen。
         '　　另外，email郵件之姓名縮寫亦無需再掛我的WW，請分別改為Akira(AL)、Owen(OC)。
   '      strcon1 = "select cp14,s1.st01,s1.st02,s1.st04,s1.st16,s1.st17 as s1_st17,oman,s2.st12,s2.st17 as s2_st17 from caseprogress,staff s1,staff s2,setspecman" & _
   '                  " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
   '                  " and cp09=(select max(cp09) from caseprogress,staff" & _
   '                  " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "' and cp14=st01(+) and st03='F21'" & _
   '                  " and cp05=(select max(cp05) from caseprogress,staff" & _
   '                  " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "' and cp14=st01(+) and st03='F21')" & _
   '                  ") and cp14=s1.st01(+)" & _
   '                  " and decode(s1.st16,'1','T','2','R','3','S','4','T1',s1.st16)=OCODE(+)" & _
   '                  " and oman=s2.st01(+)"
         'Modify By Sindy 2018/11/7 敏莉:與王經理確認後，只有密件附本不用給他，雙署名和WW需維持原狀
         ',decode(oman,null,s3.st01,oman) as oman,decode(s2.st12,null,s3.st12,s2.st12) as s2_st12,decode(s2.st17,null,s3.st17,s2.st17) as s2_st17
         'Modify By Sindy 2018/11/7 + ,s1.st70 as st70
         'Modified by Lydia 2019/08/01 排除F4102 ; 因為每月批次FCP年費自動不續辦會產生B類收文907,承辦人為F4102 'Memo by Lydia 2019/08/02 排除F4102 (FCP年費不續辦)
         strCon1 = "select cp14,s1.st01,s1.st02,s1.st04,s1.st16 as st16,s1.st17 as s1_st17" & _
                     ",oman" & _
                     ",s2.st12 as s2_st12" & _
                     ",s2.st17 as s2_st17" & _
                     ",s3.st01 as s3_st01,s1.st70 as st70" & _
                     " from caseprogress,staff s1,staff s2,staff s3,setspecman" & _
                     " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "'" & _
                     " and cp09=(select max(cp09) from caseprogress,staff" & _
                     " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "' and cp14=st01(+) and st03='F21' and cp14<>'F4102' " & _
                     " and cp05=(select max(cp05) from caseprogress,staff" & _
                     " where cp01='" & strCP01 & "' and cp02='" & strCP02 & "' and cp03='" & strCP03 & "' and cp04='" & strCP04 & "' and cp14=st01(+) and st03='F21'  and cp14<>'F4102' )" & _
                     ") and cp14=s1.st01(+) and s1.st52=s3.st01(+)" & _
                     " and decode(s1.st16,'1','T','2','R','3','S','4','T1',s1.st16)=OCODE(+)" & _
                     " and oman=s2.st01(+)"
         intQ = 1
         Set adoRst = ClsLawReadRstMsg(intQ, strCon1)
         If intQ = 1 Then
            If adoRst.Fields("st04") = "1" Then '在職
               strCP14 = adoRst.Fields("cp14")
               'Modify By Sindy 2022/5/19
               strCP14sir = PUB_GetFCPEngSup(strCP14)
               'Modify Sindy 2023/5/12 工程師撰寫信若由主任判發的信，outlook郵件信尾署名及主旨
               '                       皆帶判發主任名(原程式皆帶該組副理)，故請調整英文組三組, 撰寫信函發FC郵件（工程師署名）:要抓二級主管
               strCP14sir = PUB_GetFCPEngSup(strCP14, , , True)
               
'               'Modify By Sindy 2018/11/7
'               If "" & adoRst.Fields("st16") = "3" And "" & adoRst.Fields("st70") <> "" Then '日文組抓第三級主管
'                  'strCP14sir = "" & adoRst.Fields("s3_st01")
'                  strCP14sir = Pub_GetSpecMan(adoRst.Fields("st16") & adoRst.Fields("st70"))
'                  'Modify By Sindy 2018/11/12 若本身是小組組長,則帶oman
'                  If strCP14sir = strCP14 Then
'                     strCP14sir = "" & adoRst.Fields("oman")
'                  End If
'                  '2018/11/12 END
'                  'Add By Sindy 2019/5/23 若有二級主管,也要帶出
'                  If "" & adoRst.Fields("s3_st01") <> "" Then
'                     If InStr(strCP14sir, adoRst.Fields("s3_st01")) = 0 Then
'                        strCP14sir = adoRst.Fields("s3_st01") & ";" & strCP14sir
'                     End If
'                  End If
'                  '2019/5/23 END
'               Else
'               '2018/11/7 END
'                  strCP14sir = "" & adoRst.Fields("oman")
'               End If
               '2022/5/19 END
            '離職
            Else
               '承辦人離職的話則抓其主管
               'Modify By Sindy 2018/11/12
               If "" & adoRst.Fields("st16") = "3" Then 'And "" & adoRst.Fields("st70") <> "" Then '日文組抓第三級主管
                  'Modify By Sindy 2022/5/19
                  strCP14 = PUB_GetFCPEngSup(adoRst.Fields("cp14"), True)
                  strCP14sir = PUB_GetFCPEngSup(adoRst.Fields("cp14"))
'                  strCP14 = Pub_GetSpecMan(adoRst.Fields("st16") & adoRst.Fields("st70"))
'                  strCP14sir = Pub_GetSpecMan(adoRst.Fields("st16") & adoRst.Fields("st70"))
'                  'Add By Sindy 2019/5/23 若有二級主管,也要帶出
'                  If "" & adoRst.Fields("s3_st01") <> "" Then
'                     If InStr(strCP14sir, adoRst.Fields("s3_st01")) = 0 Then
'                        strCP14sir = adoRst.Fields("s3_st01") & ";" & strCP14sir
'                     End If
'                  End If
'                  '2019/5/23 END
                  '2022/5/19 END
               Else
               '2018/11/12 END
                  strCP14 = "" & adoRst.Fields("oman")
                  strCP14sir = "" & adoRst.Fields("oman")
               End If
            End If
            
            SirEname = "" & adoRst.Fields("s2_st12") '主管的英文名稱
            SirSt17 = UCase("" & adoRst.Fields("s2_st17")) '主管的專業代號-國外
            strCP14_st17 = "" & adoRst.Fields("s1_st17") '工程師的專業代號-國外
            'Add By Sindy 2019/8/5
            If "" & adoRst.Fields("st16") = "3" Then 'And "" & adoRst.Fields("st70") <> "" Then '日文組抓第三級主管
               StrText4 = "Division Executive"
               strText5 = "Japanese Patent Division"
            Else
               'Add By Sindy 2023/5/12 英文組:非3大主管時,要重抓英文名稱,專業代號,職稱署名
               If Len(strCP14sir) > 5 Then strCP14sir = Left(strCP14sir, 5)
               If strCP14sir <> "" & adoRst.Fields("oman") Then
                  strCon1 = "select s2.st12 as s2_st12,s2.st17 as s2_st17" & _
                              " from staff s2" & _
                              " where '" & strCP14sir & "'=s2.st01(+)"
                  intQ = 1
                  Set adoRst = ClsLawReadRstMsg(intQ, strCon1)
                  If intQ = 1 Then
                     SirEname = "" & adoRst.Fields("s2_st12") '主管的英文名稱
                     SirSt17 = UCase("" & adoRst.Fields("s2_st17")) '主管的專業代號-國外
                     'Assistant Manager(42.副理) Section Chief(51.主任)
                     If GetStaffST20(strUserNum) = "副理" Then
                        StrText4 = "Assistant Manager"
                     Else
                        StrText4 = "Section Chief"
                     End If
                  End If
               Else
               '2023/5/12 END
                  StrText4 = "Assistant Manager"
               End If
               strText5 = "International Patent Division"
            End If
            '2019/8/5 END
         End If
      End If
      
   ElseIf pDeptID = "F23" Then
      strCP14 = PUB_GetFCPSalesNo(strCP01, strCP02, strCP03, strCP04)
      strCon1 = "select st04,st52 from staff where st01='" & strCP14 & "'"
      intQ = 1
      Set adoRst = ClsLawReadRstMsg(intQ, strCon1)
      If intQ = 1 Then
         If adoRst.Fields("st04") = "1" Then '在職
            strCP14sir = "" & adoRst.Fields("st52")
         Else
            strCP14 = "" & adoRst.Fields("st52") '離職的話則抓其主管
            strCP14sir = "" & adoRst.Fields("st52")
         End If
      End If
   End If
   
   '呼叫新郵件：
   Set objOutLook = CreateObject("Outlook.Application")
   If Dir(strTemplatePath) <> "" Then
      'Modify By Sindy 2019/8/5
      If pDeptID = "F22" And InStr(UCase(strText5), UCase("Japan")) > 0 Then
         Set objMail = objOutLook.CreateItemFromTemplate(Mid(strTemplatePath, 1, InStrRev(strTemplatePath, "\")) & "$$TOT-000F22-0-00.oft")
      Else
      '2019/8/5 END
         Set objMail = objOutLook.CreateItemFromTemplate(strTemplatePath)
      End If
   Else
      Set objMail = objOutLook.CreateItem(0)
   End If
   '副本.cc
   '收件者.To
   'Modify By Sindy 2021/3/17 + , "FC"
   strTo = PUB_GetFCeMailConText("Main_EMail", strCP01, strCP02, strCP03, strCP04, "FC", strCP10)
   objMail.To = strTo
   'Modify By Sindy 2014/8/26 非E化案件或無代表號信箱時, 顯示訊息, 若同時發生時只顯示一次訊息即可.
   If PUB_GetFCeMailConText("IsECase", strCP01, strCP02, strCP03, strCP04) = "Y" Then bolIsECase = True
   strMsg = ""
   If strTo = "" Or bolIsECase = False Then
      If bolIsECase = False Then
         strMsg = "非E化案件"
      End If
      If strTo = "" Then
         If strMsg <> "" Then strMsg = strMsg & "且"
         strMsg = strMsg & "無代表號信箱"
      End If
      MsgBox strMsg, vbInformation
   End If
   '2014/8/26 END
   strCaseNo = PUB_GetFCeMailConText("CaseNo", strCP01, strCP02, strCP03, strCP04)
   'Modify By Sindy 2020/3/19 + StrProcCP10
   strYourRef = PUB_GetFCeMailConText("YourRef", strCP01, strCP02, strCP03, strCP04, strFCorCF, StrProcCP10)
   strOurRef = PUB_GetFCeMailConText("OurRef", strCP01, strCP02, strCP03, strCP04)
   'Modify By Sindy 2015/5/29 依定稿別作區分
   If InStr("08,13", m_ET01) > 0 And m_ET01 <> "" Then '不加s
      m_Encls = "Encl."
   Else
      m_Encls = "Encls."
   End If
   '2015/5/29 END
   
   'Added by Morgan 2018/12/10 案件備註可能有HTML識別字元"<"及">"要先置換。 Ex:FCP-058806
   If EMailType = "HTML" Then
      strContent = Replace(Replace(strContent, "<", "&lt;"), ">", "&gt;")
   End If
   'end 2018/12/10
   
   If pDeptID = "F22" Then
      'If EMailType = "HTML" And InStr(strTemplatePath, "TOT-000F22-0-00") = 0 Then
      If EMailType = "HTML" And InStr(strTemplatePath, "TOT-000F22") = 0 Then
         strContent = strContent & "<BR>Best regards,<BR>"
         '定稿維護:
         '密件副本.BCC
         objMail.BCC = "fcp.taie@msa.hinet.net"
         '主旨.Subject
         'Modify By Sindy 2018/5/18 +  & " [PROC." & strCP10 & "]"
         'Modify By Sindy 2019/5/23
         'objMail.Subject = "DY/" & Pub_StrUserSt17 & " - " & strCaseNo & "; " & strYourRef & "; " & strOurRef & " [PROC." & strCP10 & "]"
         'Modify By Sindy 2019/8/5
         If InStr(UCase(strText5), UCase("Japan")) > 0 Then
            'Modify By Sindy 2022/11/30 WW/ => 改抓特殊設定
            Call GetPrjSalesNM(Pub_GetSpecMan("S"), , strST17)
            objMail.Subject = UCase(strST17) & "/" & Pub_StrUserSt17 & " - " & strCaseNo & "; " & strYourRef & "; " & strOurRef & " [PROC." & StrProcCP10 & "]"
            strSignature = "<BR>" & UCase(strST17) & "/" & Pub_StrUserSt17 & "<BR>" '& m_Encls & "<BR><BR>"
         Else
         '2019/8/5 END
            'Modify By Sindy 2022/11/30 DY/ => 改抓特殊設定
            Call GetPrjSalesNM(Pub_GetSpecMan("外專承辦英文組主管"), , strST17) '外專承辦(DY/dy)
            If InStr(strST17, "/") > 0 Then
               strST17 = Left(strST17, InStr(strST17, "/") - 1)
            End If
            objMail.Subject = UCase(strST17) & "/" & Pub_StrUserSt17 & " - " & strCaseNo & "; " & strYourRef & "; " & strOurRef & " [PROC." & StrProcCP10 & "]"
            strSignature = "<BR>" & UCase(strST17) & "/" & Pub_StrUserSt17 & "<BR>" & m_Encls & "<BR><BR>"
         End If
      Else
         strContent = strContent & "Best regards," & vbCrLf & vbCrLf & vbCrLf
         
'         strContent = strContent & SirEname & Space(60 - 6 - (Len(SirEname) * 2) + 3) & "Fred C. T. Yen" & vbCrLf
'         'strContent = strContent & "Patent Department" & vbTab & vbTab & vbTab & "Patent Attorney" & vbCrLf
'         strContent = strContent & "Patent Department" & Space(60 - (Len("Patent Department") * 2) + 3) & "Patent Attorney" & vbCrLf
'         strContent = strContent & Space(60 - 17) & "Managing Partner" & vbCrLf & vbCrLf
         
         'Modify By Sindy 2019/1/11 Mark,改寫下面用replace方式置換文字
'         strContent = strContent & convForm(CheckStr(SirEname), 45) & "Fred C. T. Yen" & vbCrLf
'         strContent = strContent & convForm(CheckStr("Patent Department"), 40) & "Patent Attorney" & vbCrLf
'         strContent = strContent & convForm(CheckStr(" "), 55) & "Managing Partner" & vbCrLf & vbCrLf
'         strContent = strContent & SirSt17 & "/" & strCP14_st17 & "/" & Pub_StrUserSt17 & vbCrLf
'         strContent = strContent & m_Encls & vbCrLf
         
         '撰寫信函:
         '密件副本.BCC
         'modify by sonia 2023/5/17 信尾署名改為主任但密件副本還是要發二級主管
         'objMail.BCC = "fcp.taie@msa.hinet.net;" & strCP14sir & IIf(InStr(strCP14sir, strCP14) > 0, "", ";" & strCP14)        'IIf(strCP14 = strCP14sir, "", ";" & strCP14)
         objMail.BCC = "fcp.taie@msa.hinet.net;" & strCP14sir & IIf(InStr(strCP14sir, strCP14) > 0, "", ";" & strCP14) & IIf(InStr(strCP14sir, PUB_GetFCPEngSup(strCP14)) > 0, "", ";" & PUB_GetFCPEngSup(strCP14))     'IIf(strCP14 = strCP14sir, "", ";" & strCP14)
         '主旨.Subject
         'Modify By Sindy 2018/5/18 +  & " [PROC." & strCP10 & "]"
         'Modify By Sindy 2019/5/23
         'objMail.Subject = SirSt17 & "/" & strCP14_st17 & "/" & Pub_StrUserSt17 & " - " & strCaseNo & "; " & strYourRef & "; " & strOurRef & " [REP." & strCP10 & "]"
         objMail.Subject = SirSt17 & "/" & strCP14_st17 & "/" & Pub_StrUserSt17 & " - " & strCaseNo & "; " & strYourRef & "; " & strOurRef & " [REP." & StrProcCP10 & "]"
      End If
   ElseIf pDeptID = "F23" Then
      strContent = strContent & "<BR>Best regards,<BR>"
      '密件副本.BCC
      objMail.BCC = "fcp.taie@msa.hinet.net;" & strCP14sir & IIf(InStr(strCP14sir, strCP14) > 0, "", ";" & strCP14) 'IIf(strCP14 = strCP14sir, "", ";" & strCP14)
      '主旨.Subject
      'Modify By Sindy 2018/5/18 +  & " [PROC." & strCP10 & "]"
      'Modify By Sindy 2019/5/23
      'objMail.Subject = Pub_StrUserSt17 & " - " & strCaseNo & "; " & strYourRef & "; " & strOurRef & " [PROC." & strCP10 & "]"
      objMail.Subject = Pub_StrUserSt17 & " - " & strCaseNo & "; " & strYourRef & "; " & strOurRef & " [PROC." & StrProcCP10 & "]"
      strSignature = "<BR>" & Pub_StrUserSt17 & "<BR><BR>"
   End If
   '加附件
   If strAttach <> "" Then
      ArrStr = Split(strAttach, ";")
      For ii = 0 To UBound(ArrStr)
         objMail.Attachments.add ArrStr(ii)
      Next ii
   End If
   '內文.Body
   If EMailType = "HTML" Then
      '轉HTML格式
      'strContent = Replace(strContent, "新細明體", "Times New Roman")
      strContent = Replace(strContent, vbCrLf, "<BR>")
      strContent = Replace(strContent, "  ", "&nbsp;&nbsp;")
   End If
   m_MySt(1) = strCP01 'Add By Sindy 2014/11/20
   If UCase(EMailType) = "HTML" Then
      'Modify By Sindy 2014/11/20 加 LetterMemo
      objMail.HTMLBody = "<FONT FACE=""Times New Roman"">" & strContent & "<BR>" & Replace(Replace(objMail.HTMLBody, "&lt;ST17&gt;", strSignature), "&lt;LetterMemo&gt;", IIf(ExceptFieldData("公用備註/英") <> "", "<BR>Message:<BR>" & Replace(ChgHTMLFormat(ExceptFieldData("公用備註/英")), vbCrLf, "<BR>") & "<BR><BR>", "&nbsp;")) & "</FONT>"
      If InStr(objMail.HTMLBody, "&lt;<span class=SpellE>LetterMemo</span>&gt;") > 0 Then
         objMail.HTMLBody = Replace(objMail.HTMLBody, "&lt;<span class=SpellE>LetterMemo</span>&gt;", IIf(ExceptFieldData("公用備註/英") <> "", "<BR>Message:<BR>" & Replace(ChgHTMLFormat(ExceptFieldData("公用備註/英")), vbCrLf, "<BR>") & "<BR><BR>", "&nbsp;"))
      End If
'      'Add By Sindy 2017/11/16
'      objMail.HTMLBody = Replace(objMail.HTMLBody, "http://www.taie.com.tw", "https://www.taie.com.tw")
'      '2017/11/16 END
      'Add By Sindy 2018/12/24
      'Modify By Sindy 2019/1/11 字不會排列整齊,改寫用replace方式置換文字
      If pDeptID = "F22" Then
         If InStr(objMail.HTMLBody, "台一置換文字一") > 0 Then
            objMail.HTMLBody = Replace(objMail.HTMLBody, "台一置換文字一", SirEname)
         End If
         If InStr(objMail.HTMLBody, "台一置換文字二") > 0 Then
            objMail.HTMLBody = Replace(objMail.HTMLBody, "台一置換文字二", IIf(strSignature <> "", strSignature, SirSt17 & "/" & strCP14_st17 & "/" & Pub_StrUserSt17))
         End If
         If InStr(objMail.HTMLBody, "台一置換文字三") > 0 Then
            objMail.HTMLBody = Replace(objMail.HTMLBody, "台一置換文字三", m_Encls)
         End If
         'Add By Sindy 2019/8/5
         If InStr(objMail.HTMLBody, "台一置換文字四") > 0 Then
            objMail.HTMLBody = Replace(objMail.HTMLBody, "台一置換文字四", StrText4)
         End If
         If InStr(objMail.HTMLBody, "台一置換文字五") > 0 Then
            objMail.HTMLBody = Replace(objMail.HTMLBody, "台一置換文字五", strText5)
         End If
         '2019/8/5 END
         objMail.HTMLBody = Replace(objMail.HTMLBody, "新細明體", "Times New Roman")
      '2019/1/11 END
         If InStr(objMail.HTMLBody, "SystemWordLetterMemo") > 0 Then
            objMail.HTMLBody = Replace(objMail.HTMLBody, "SystemWordLetterMemo", IIf(ExceptFieldData("公用備註/英") <> "", "<BR>Message:<BR>" & Replace(ChgHTMLFormat(ExceptFieldData("公用備註/英")), vbCrLf, "<BR>") & "<BR><BR>", "&nbsp;"))
         End If
      End If
      objMail.HTMLBody = Replace(objMail.HTMLBody, "http://www.taie.com.tw", "https://www.taie.com.tw")
      '2018/12/24 END
   Else
      'Modify By Sindy 2014/11/20 加 LetterMemo
      objMail.Body = strContent & vbCrLf & Replace(objMail.Body, "<LetterMemo>", IIf(ExceptFieldData("公用備註/純文字") <> "", vbCrLf & "Message:" & vbCrLf & ExceptFieldData("公用備註/純文字") & vbCrLf, ""))
      'Add By Sindy 2017/11/16
      objMail.Body = Replace(objMail.Body, "http://www.taie.com.tw", "https://www.taie.com.tw")
      '2017/11/16 END
   End If
'   objMail.HTMLBody = "<font face=新細明體 size=2>" & strTemp & "<BR><BR>" & _
'"Tai E International Patent & Law Office<BR>" & _
'"9Fl., No. 112, Sec. 2, Chang-An E. Rd.<BR>" & _
'"Taipei 104, Taiwan, R.O.C.<BR>" & _
'"P.O. Box: 46-478, Taipei 104, Taiwan, R.O.C.<BR>" & _
'"Tel: 886-2-25061023, 25081531<BR>" & _
'"Fax: 886-2-25068147, 25064319, 25076571, 25090804<BR>" & _
'"URL: &lt;<A href=""http://www.taie.com.tw/"">http://www.taie.com.tw/</A>&gt;<BR>" & _
'"E-mail: ipdept@taie.com.tw<BR><BR>" & _
'"This e-mail transmission is intended only for the use of the individual or entity to which it is addressed,<BR>" & _
'"And may contain information that is privileged, confidential and exempt from disclosure under applicable law.<BR>" & _
'"If the reader is not the intended recipient, you are hereby notified that any dissemination, distribution or copying<BR>" & _
'"Of this communication is strictly prohibited. If you have received this transmission in error, please notify us<BR>" & _
'"Immediately, and return the original message to us at the above address. We greatly appreciate your cooperation.<BR>" & _
'"</FONT>"
   objMail.Display
   
   Set objMail = Nothing
   Set objOutLook = Nothing
   Set adoRst = Nothing
End Function

'Add By Sindy 2014/8/15
'Download郵件範本
'Modify By Sindy +, Optional ByVal strLanguage As String = "2"
'                1.中 2.英 3.日
Public Function PUB_DownloadOftPath(strST03 As String, strST17 As String, Optional ByRef EMailType As String = "HTML", _
                                    Optional ByVal bolDownLoad As Boolean = True, Optional ByVal strLanguage As String = "2") As String
   PUB_DownloadOftPath = ""
   If strST03 = "F22" Then
      '郵件格式為RTF(附件會在內文裡)
      'Modify By Sindy 2018/12/24 葉敏莉:因有客戶反應程序用撰寫信函(工程師署名)的產生的Outlook附檔打不開，故請將Outlook的格式改成HTML
      'EMailType = "RTF"
      EMailType = "HTML"
      '2018/12/24 END
      
      'Add By Sindy 2021/9/22 特殊申請人的出名人簽名檔
      PUB_DownloadOftPath = "$$TOT-000F22-0-02.oft" '閻副所長+ Wilson 出名:德國客戶 'Modified by Morgan 2024/3/27 Harriet 改 Wilson
      If bolDownLoad = True Or Dir(App.path & "\" & PUB_DownloadOftPath) = "" Then
         Call PUB_GetSampleFile(PUB_DownloadOftPath, Replace(Left(PUB_DownloadOftPath, Len(PUB_DownloadOftPath) - 4), "$$", ""))
      End If
      '2021/9/22 END
      
      PUB_DownloadOftPath = "$$TOT-000F22-0-00.oft"
      
   ElseIf strST03 = "F23" Then
      '郵件格式為HTML(可以加圖片,指定字型,大小等...)
      EMailType = "HTML"
      'Modify By Sindy 2019/8/5
'      If strST17 <> "" Then
'         If Left(strST17, InStr(strST17, "/") - 1) = "EC" Then
'            PUB_DownloadOftPath = "$$TOT-000F23-0-03.oft"
'         ElseIf Left(strST17, InStr(strST17, "/") - 1) = "ELC" Then
'            PUB_DownloadOftPath = "$$TOT-000F23-0-02.oft"
'         Else
'            PUB_DownloadOftPath = "$$TOT-000F23-0-01.oft"
'         End If
'      Else
'         PUB_DownloadOftPath = "$$TOT-000F23-0-01.oft"
'      End If

      'Add By Sindy 2020/9/8 特殊申請人的出名人簽名檔
      PUB_DownloadOftPath = "$$TOT-000F23-0-04.oft" 'Ryan:德國客戶
      If bolDownLoad = True Or Dir(App.path & "\" & PUB_DownloadOftPath) = "" Then
         Call PUB_GetSampleFile(PUB_DownloadOftPath, Replace(Left(PUB_DownloadOftPath, Len(PUB_DownloadOftPath) - 4), "$$", ""))
      End If
      '2020/9/8 END
      
      PUB_DownloadOftPath = "$$TOT-000F23-0-01.oft" '英文
      If PUB_GetStaffST16(strUserNum) = "2" Then '日文
         If strLanguage = "3" Then '日
            PUB_DownloadOftPath = "$$TOT-000F23-0-03.oft"
         Else
            PUB_DownloadOftPath = "$$TOT-000F23-0-02.oft"
         End If
      End If
      '2019/8/5 END
      
   'Add By Sindy 2017/3/6 巨京查名郵件
   ElseIf strST03 = "P22" And strST17 = "P29" Then
      PUB_DownloadOftPath = "$$TOT-000P29-0-00.oft"
      
   '2017/3/6 END
   'Add By Sindy 2016/8/3 內商郵件範本
   ElseIf Left(strST03, 2) = "P2" Then
      PUB_DownloadOftPath = "$$TOT-000P20-0-00.oft"
   '2016/8/3 END
   'Add By Sindy 2024/7/22 外商郵件範本
   ElseIf Left(strST03, 2) = "F1" Then
      PUB_DownloadOftPath = "$$TOT-000F11-0-01.oft" 'CFT催審郵件
   '2024/7/22 END
   End If
   
   If PUB_DownloadOftPath <> "" Then
      'Modify By Sindy 2015/4/27 改檢查無檔案,就需要重新download +Or Dir(App.path & "\" & PUB_DownloadOftPath) = ""
      If bolDownLoad = True Or Dir(App.path & "\" & PUB_DownloadOftPath) = "" Then
      '2015/4/27 END
         Call PUB_GetSampleFile(PUB_DownloadOftPath, Replace(Left(PUB_DownloadOftPath, Len(PUB_DownloadOftPath) - 4), "$$", ""))
      End If
      PUB_DownloadOftPath = App.path & "\" & PUB_DownloadOftPath
   End If
   '2014/9/12 END
End Function

'Add By Sindy 2024/7/22
'外商發郵件資訊
Public Function PUB_SettingF11eMail(ByVal strTemplatePath As String, _
               ByVal strCP01 As String, ByVal strCP02 As String, ByVal strCP03 As String, ByVal strCP04 As String, _
               ByVal strContent As String) As Boolean
Dim adoRst As ADODB.Recordset
Dim intQ As Integer
Dim objOutLook As Object
Dim objMail As Object
Dim strCP14 As String, strCP14sir As String
Dim strCaseNo As String, strYourRef As String, strOurRef As String
Dim strSignature As String
Dim strST17 As String
Dim strRe As String, m_AttachPath As String, strFullFile As String
Dim strRecipients_all As String, strRecipients_1 As String
   
   strCon1 = "select st04,st52 from staff where st01='" & strUserNum & "'"
   intQ = 1
   Set adoRst = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      If adoRst.Fields("st04") = "1" Then '在職
         strCP14 = strUserNum
         strCP14sir = "" & adoRst.Fields("st52")
      Else
         strCP14 = "" & adoRst.Fields("st52") '離職的話則抓其主管
         strCP14sir = "" & adoRst.Fields("st52")
      End If
   End If
   Call GetPrjSalesNM(strCP14sir, , strST17)
   strSignature = strST17 & "/" & Pub_StrUserSt17
   strSignature = UCase(Mid(strSignature, 1, 1)) & Mid(strSignature, 2)
   
   '定義 Outlook APP
   Set objOutLook = CreateObject("Outlook.Application")
   
   m_AttachPath = App.path & "\" & strUserNum
   If Dir(m_AttachPath, vbDirectory) = "" Then
      MkDir m_AttachPath
   End If
   '從卷宗區抓最新一封外來郵件，由於無法帶出原信件內容，將原信件內容以附件的方式夾帶至系統產生之outlook信函內
   '並帶出收件者信箱(該外來郵件的原寄件人)，及副本收件人信箱(該外來郵件的原副本收件人)
   '最新一封外來郵件的定義: 卷宗區附檔名為代理人來函=altr、外來郵件=rx 的msg檔，以最新的檔案修改日期為準
   strExc(0) = "SELECT CASEPAPERPDF.* FROM CASEPAPERPDF,CASEPROGRESS" & _
               " WHERE " & ChgCaseprogress(strCP01 & strCP02 & strCP03 & strCP04) & " AND CP09=CPP01(+)" & _
               " AND (INSTR(upper(CPP02),'.ALTR.')>0 OR INSTR(upper(CPP02),'.RX.')>0)" & _
               " ORDER BY nvl(CPP17,CPP08) desc,nvl(CPP18,CPP09) desc"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
   Dim strSender As String
   If intI = 1 Then
      RsTemp.MoveFirst
      strFullFile = m_AttachPath & "\" & RsTemp.Fields("cpp02")
      If PUB_GetAttachFile_CPP(RsTemp.Fields("cpp01"), RsTemp.Fields("cpp02"), strFullFile, True) = False Then
         MsgBox "無法儲存檔案[ " & RsTemp.Fields("cpp14") & " ]！"
         Exit Function
      Else
         '開啟外來郵件：
         Set objMail = objOutLook.CreateItemFromTemplate(strFullFile)
         
         'Modify By Sindy 2025/2/17
'         If InStr(UCase(objMail.senderemailaddress), UCase("Recipients/cn=")) = 0 _
'            And InStr(UCase(objMail.senderemailaddress), UCase("Public Folder/CN=")) = 0 Then
'            strSender = objMail.senderemailaddress
'         Else
'            strSender = ""
'         End If
         '抓收件者資料
         strRecipients_1 = "" '收件者
         strRecipients_all = ""
         'Call PUB_ReadMailText_CC(objMail, strRecipients_all, strRecipients_1, True)
         Call PUB_ReadMailText(objMail, strRecipients_all, strRecipients_1, True, strSender)
         '2025/2/17 END
'         MsgBox "外來郵件的原寄件人 = " & strSender & vbCrLf & vbCrLf & _
'                "收件者=" & strRecipients_1 & vbCrLf & vbCrLf & _
'                "外來郵件的原副本收件人 = " & strRecipients_all
      End If
   End If
   
   '呼叫新郵件：
'   If Dir(strTemplatePath) <> "" Then
'      Set objMail = objOutLook.CreateItemFromTemplate(strTemplatePath)
'   Else
      Set objMail = objOutLook.CreateItem(0)
'   End If
   
   '收件者.To
   objMail.To = strSender
   strCaseNo = PUB_GetFCeMailConText("CaseNo", strCP01, strCP02, strCP03, strCP04)
   strYourRef = PUB_GetFCeMailConText("YourRef", strCP01, strCP02, strCP03, strCP04, "CF")
   strOurRef = PUB_GetFCeMailConText("OurRef", strCP01, strCP02, strCP03, strCP04)
   strRe = PUB_GetFCeMailConText("RE", strCP01, strCP02, strCP03, strCP04, "CF")
   strContent = strRe & vbCrLf & strContent
   strContent = Replace(Replace(strContent, "<", "&lt;"), ">", "&gt;")
   'strContent = strContent & "<BR>Best regards,<BR>"
   
   '加附件
   If strFullFile <> "" Then objMail.Attachments.add strFullFile
   
   '副本.cc
   '密件副本.BCC
   objMail.BCC = strRecipients_all
   '主旨.Subject
   objMail.Subject = strSignature & " - " & strCaseNo & "; " & IIf(strYourRef <> "", strYourRef & "; ", "") & strOurRef
   'strSignature = "<BR>" & Pub_StrUserSt17 & "<BR><BR>"

   '轉HTML格式
   'strContent = Replace(strContent, "新細明體", "Times New Roman")
   strContent = Replace(strContent, vbCrLf, "<BR>")
   strContent = Replace(strContent, "  ", "&nbsp;&nbsp;")
   'Modify By Sindy 2014/11/20 加 LetterMemo
   objMail.HTMLBody = "<FONT FACE=""Times New Roman"">" & strContent & "<BR>" & Replace(Replace(objMail.HTMLBody, "&lt;ST17&gt;", strSignature), "&lt;LetterMemo&gt;", IIf(ExceptFieldData("公用備註/英") <> "", "<BR>Message:<BR>" & Replace(ChgHTMLFormat(ExceptFieldData("公用備註/英")), vbCrLf, "<BR>") & "<BR><BR>", "&nbsp;")) & "</FONT>"
   If InStr(objMail.HTMLBody, "&lt;<span class=SpellE>LetterMemo</span>&gt;") > 0 Then
      objMail.HTMLBody = Replace(objMail.HTMLBody, "&lt;<span class=SpellE>LetterMemo</span>&gt;", IIf(ExceptFieldData("公用備註/英") <> "", "<BR>Message:<BR>" & Replace(ChgHTMLFormat(ExceptFieldData("公用備註/英")), vbCrLf, "<BR>") & "<BR><BR>", "&nbsp;"))
   End If
   objMail.HTMLBody = Replace(objMail.HTMLBody, "http://www.taie.com.tw", "https://www.taie.com.tw")
   objMail.Display
   
   Set objMail = Nothing
   Set objOutLook = Nothing
   Set adoRst = Nothing
End Function

'Added by Lydia 2016/12/22 收款寄帳(1728)控制-發催款函
'Modified by Lydia 2017/02/18 +pSpec 是否為特定客戶(寄紙本)
Public Sub PUB_SendA1kdataMail(ByRef sForm As Form, ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal nCP09 As String, ByVal nCP10 As String, Optional ByVal pAC2470 As String = "", Optional pSpec As String = "N")
'nCP09/nCP10      收款寄帳(1728)的相關總收文號/案件性質(核准或註冊證)
'pAC2470    FC代理人 'Memo by Lydia 2017/02/18 改成預設都附催款單和請款單
Dim EMailType As String  '指定郵件格式
Dim strTemplatePath As String '郵件範本檔
Dim strAttPath As String '附件電子檔路徑
Dim strTmp1 As String
Dim strA1kdata As String
Dim tmpArr As Variant, inR As Integer, strPath2 As String 'Added by Lydia 2017/04/06

tmpArr = Split(pAC2470, ",") 'Added by Lydia 2017/04/06

    '隨信附上個案催款單，呼叫FC催款單
    If pAC2470 <> "" Then
       If UCase(TypeName(TmpFrmAcc2470)) <> "NOTHING" Then
         'Added by Lydia 2017/04/06 如果一個案子的請款對象超出一個
         For inR = 0 To UBound(tmpArr)
           If Trim(tmpArr(inR)) <> "" Then
            'Modified by Lydia 2017/06/13 改放在個人文件夾底下
            'strAttPath = "C:\收款寄證"
            strAttPath = GetMyDocPath & "\收款寄證"
            'end 2017/06/13
            
            'Modified by Lydia 2017/04/06
            'pAC2470 = ChangeCustomerL(pAC2470)
            strTmp1 = ChangeCustomerS(Trim(tmpArr(inR)))
            TmpFrmAcc2470.SetParent sForm
            'Modified by Lydia 2017/04/06 因為有傳案號,所以請款對象改抓000~ZZZ
            'TmpFrmAcc2470.Text1 = pAC2470
            'TmpFrmAcc2470.Text2 = pAC2470
            TmpFrmAcc2470.Text1 = strTmp1 & String(9 - Len(strTmp1), "0")
            TmpFrmAcc2470.Text2 = strTmp1 & String(9 - Len(strTmp1), "Z")
            'end 2017/04/06
            'Modified by Lydia 2017/04/06
            'strTmp1 = GetPrjNationNumber(pAC2470)
            strTmp1 = GetPrjNationNumber(ChangeCustomerL(strTmp1))
            If strTmp1 = "020" Then
               TmpFrmAcc2470.Text3 = "020"
               TmpFrmAcc2470.Text4 = "020"
            End If
            TmpFrmAcc2470.strCallCase = pCP01 & pCP02 & pCP03 & pCP04 '傳入本所案號
            'Modified by Lydia 2024/12/31 只存PDF檔
            'TmpFrmAcc2470.Text6 = "2" '存PDF檔
            TmpFrmAcc2470.Text6 = "Y"
            TmpFrmAcc2470.Text7 = "Y" 'Added by Lydia 2017/02/18 附請款單
            'Modified by Lydia 2017/02/18 改傳變數
            'TmpFrmAcc2470.Tag = strAttPath
            TmpFrmAcc2470.m_SavePath = strAttPath
            TmpFrmAcc2470.Show
            TmpFrmAcc2470.MaskEdBox2.Text = ChangeTStringToTDateString(strSrvDate(2)) '請款日期止
            Call TmpFrmAcc2470.Command2_Click
            strAttPath = TmpFrmAcc2470.Tag
            
            Unload TmpFrmAcc2470
            If Mid(strAttPath, 1, 1) = "*" Then strAttPath = Mid(strAttPath, 2)
            strAttPath = Replace(strAttPath, "*", ";") 'Added by Lydia 2017/02/18 有多個檔案
           End If
           strPath2 = strPath2 & IIf(Len(strPath2) > 0, ";", "") & strAttPath 'Added by Lydia 2017/04/06
         Next
         'end 2017/04/06
       End If
    End If
    
    strAttPath = strPath2 'Added by Lydia 2017/04/06
    strA1kdata = GetT_020_a1k_data(pCP01, pCP02, pCP03, pCP04, "2", True)
    'Added by Lydia 2017/02/18 第一碼區分是否為特定客戶(寄紙本)
    strA1kdata = pSpec & strA1kdata
    
    '取得郵件範本檔名
    strTemplatePath = PUB_DownloadOftPath("P20", "", EMailType, False)
    'Modify By Sindy 2018/5/14
    'Call PUB_SettingTeMail(strTemplatePath, pCP01, pCP02, pCP03, pCP04, strAttPath, nCP10, "1728", nCP09, strA1kdata)
    Call PUB_SettingTeMail(sForm, strTemplatePath, pCP01, pCP02, pCP03, pCP04, strAttPath, nCP10, nCP09, "1728", nCP09, strA1kdata)
    '2018/5/14 END
End Sub

'Added by Lydia 2018/05/31 FCP案的會稿924自動掛承辦人和承辦期限(會稿沒掛本所及法限的話，自動掛承辦期限)
Public Function PUB_Update924CP(ByVal pa01 As String, ByVal pa02 As String, ByVal pa03 As String, ByVal pa04 As String, ByVal pCP14 As String, ByVal pDate As String) As String
'pCP14 承辦人
'pDate 起算日期(中說所限)
Dim intA As Integer, strA1 As String
Dim rsA1 As New ADODB.Recordset
    
    If pDate = "" And pCP14 = "" Then Exit Function
    
    PUB_Update924CP = ""
    strA1 = "select cp05,cp06,cp07,cp09,cp10,cp14,cp48 from caseprogress where cp01='" & pa01 & "' and cp02='" & pa02 & "' and cp03='" & pa03 & "' and cp04='" & pa04 & "' " & _
               "and cp10='924' and cp158=0 and cp159=0 "
    intA = 1
    Set rsA1 = ClsLawReadRstMsg(intA, strA1)
    If intA = 1 Then
         If Trim("" & rsA1.Fields("cp06") & rsA1.Fields("cp06")) = "" Then '會稿沒掛本所及法限的話，自動掛承辦期限
            '承辦期限=中說所限前7日，碰到假日提前一工作日
            strA1 = CompWorkDay(1, CompDate(2, -7, TransDate(pDate, 2)))
            '承辦期限>會稿所限 ,承辦期限=會稿所限
            If Val("" & rsA1.Fields("cp06")) > 0 And Val(strA1) > Val("" & rsA1.Fields("cp06")) Then
                strA1 = "" & rsA1.Fields("cp06")
            End If
            '若承辦期限小於或等於系統日則不掛期限改彈訊息
            If strA1 <= strSrvDate(1) Then
                 PUB_Update924CP = "[會稿]承辦期限小於系統日請手動key !"
                 If pCP14 <> "" Then
                    strSql = "update caseprogress set cp14='" & Trim(pCP14) & "', cp122='Y' where cp09='" & rsA1.Fields("cp09") & "' "
                    cnnConnection.Execute strSql, intA
                 End If
            Else
                 strSql = "update caseprogress set cp48=" & strA1 & IIf(pCP14 <> "", ", cp14='" & Trim(pCP14) & "', cp122='Y'", "") & " where cp09='" & rsA1.Fields("cp09") & "' "
                 cnnConnection.Execute strSql, intA
            End If
         'Added by Lydia 2018/06/28 有法限或所限,只更新承辦人
         ElseIf pCP14 <> "" Then
                 strSql = "update caseprogress set cp14='" & Trim(pCP14) & "', cp122='Y' where cp09='" & rsA1.Fields("cp09") & "' "
                 cnnConnection.Execute strSql, intA
         'end 2018/06/28
         End If
    End If
    Set rsA1 = Nothing
    
    Exit Function
End Function

'Added by Lydia 2021/04/14 取得工程師未完稿的翻譯案件清單
Public Function Pub_GetEngEP09List(ByVal pUserNo As String, Optional ByVal pCaseNo As String) As String
Dim strQuery As String
Dim rsQuery As New ADODB.Recordset
Dim intQ As Integer
     
    Pub_GetEngEP09List = ""
    
    'staff_idmap參考PUB_GetMapID：抓員工外譯對照資料
    strQuery = "select cp01||'-'||cp02||decode(cp03,'0','','-'||cp03)||decode(cp04,'00','','-'||cp04) caseno,sqldatet(cp48) cp48t " & _
                    "From caseprogress,patent, engineerprogress " & _
                    "where cp01 in ('FCP','P') and cp10='201' and cp158=0 and cp159=0 " & _
                    "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa58||pa108 is null " & _
                    "and cp14 in (select st01 from staff where st01='" & pUserNo & "' " & _
                    "union all select sim02 from staff_idmap,staff where sim01='" & pUserNo & "' and st01=sim02 and st04='1') " & _
                     "and cp09=ep02(+) and ep09 is null "
    If pCaseNo <> "" Then
        strQuery = strQuery & "and cp01||cp02||cp03||cp04 <> " & CNULL(pCaseNo)
    End If
    strQuery = strQuery & "order by cp48 "
    intQ = 1
    Set rsQuery = ClsLawReadRstMsg(intQ, strQuery)
    If intQ = 1 Then
        strQuery = ""
        rsQuery.MoveFirst
        Do While Not rsQuery.EOF
             strQuery = strQuery & vbCrLf & rsQuery.Fields("caseno") & "，承辦期限：" & rsQuery.Fields("cp48t")
             rsQuery.MoveNext
        Loop
        Pub_GetEngEP09List = strQuery
    End If
    Set rsQuery = Nothing
    
End Function

'Added by Lydia 2018/06/12 取得命名作業-欲翻譯人員的員工編號(名稱)
'Modified by Lydia 2018/09/27 +中說收文號iKeyNo
Public Function Pub_GetTct27ID(ByVal iT10 As String, ByVal iT27 As String, ByVal iT28 As String, Optional ByVal iKeyNo As String = "", Optional ByRef iTname As String) As String
Dim strID As String, strType As String
'Added by Lydia 2018/09/27
Dim rsB1 As New ADODB.Recordset
Dim intB As Integer, StrSqlB As String

'Modified by Lydia 2018/09/27 + iKeyNO
If iT10 = "" And iT27 = "" And iT28 = "" And iKeyNo = "" Then Exit Function

Pub_GetTct27ID = ""
iTname = ""

    'Added by Lydia 2018/09/27 傳入中說收文號，抓認領翻譯人員
    If iKeyNo <> "" Then
        StrSqlB = "select tfa04,tfa05,st02,st04 from transfeeassign, staff where tfa01='" & iKeyNo & "' and tfa06 is not null and tfa01=st01(+)"
        intB = 1
        Set rsB1 = ClsLawReadRstMsg(intB, StrSqlB)
        If intB = 1 Then
             StrSqlB = "" & rsB1.Fields("tfa04")
             If "" & rsB1.Fields("tfa05") = "A" Then
                  strID = PUB_GetMapID(StrSqlB, 0)
                  If strID = "" Then '沒有下班翻譯編號
                      strID = "" & rsB1.Fields("tfa04")
                  End If
                  strType = "-下班"
             Else
                  strID = "" & rsB1.Fields("tfa04")
                  If "" & rsB1.Fields("tfa05") = "B" Then
                      strType = "-上班"
                  End If
             End If
             iTname = GetStaffName(strID, True) & strType
        End If
        Set rsB1 = Nothing
    Else
    'end 2018/09/27
        Select Case iT27
            Case "A"  '命名人員下班翻譯
                  strID = PUB_GetMapID(iT10, 0)
                  If strID = "" Then '沒有下班翻譯編號
                      strID = iT10
                  End If
                  strType = "-下班"
            Case "B"  '命名人員上班翻譯
                  strID = iT10
                  strType = "-上班"
            'Mark by Lydia 2025/03/13 改用模組
            'Case "1" '舜禹
            '      strID = 外翻_舜禹
            'Case "2" '捷恩凱
            '      strID = 外翻_捷恩凱
            'Case "3" '迅達
            '      strID = 外翻_迅達
            'end 2025/03/13
            '指定人名-A/B(下/上班翻譯)
            'Modified by Lydia 2025/03/13
            'Case "4"
            Case "4", "Z"
               'Added by Lydia 2025/03/13 新增國外翻譯社
               If Trim(iT28) = "" Then  '114/3/14以後原本的4會改成Z
                  GoTo JumpToNewCase
               Else
               'end 2025/03/13
                  strID = iT28
                  If Right(UCase(strID), 1) = "A" Then
                       strID = Mid(strID, 1, Len(strID) - 1) & "下班"
                  ElseIf Right(UCase(strID), 1) = "B" Then
                       strID = Mid(strID, 1, Len(strID) - 1) & "上班"
                  End If
               End If
               'end 2025/03/13
            'Added by Lydia 2025/03/13 改用模組
            Case Else
JumpToNewCase:
               strID = Pub_SetF51Order("", iT27)
            'end 2025/03/13
        End Select
        'Modified by Lydia 2025/04/13
        'If iT27 <> "4" Then
        If iT27 = "A" Or iT27 = "B" Or iT27 = "Z" Or (strSrvDate(1) < "20250314" And iT27 = "4") Then '114/3/14以後原本的4會改成Z
             iTname = GetStaffName(strID, True) & strType
        Else
             iTname = strID
        End If
    End If 'end 2018/09/27
    Pub_GetTct27ID = strID
End Function

'Added by Lydia 2018/06/13 FC翻譯案件郵件
Public Sub PUB_Translate_SendMail(ByVal pType As String, ByVal pSavePath As String, ByRef pTFile As String, ByVal pTF01 As String, ByVal pCP14 As String, Optional ByRef pTF24 As String, Optional ByRef pTF25 As String)
'pType 對外
'pSavePath 附件存放路徑
'pTF01 中說收文號
'pCP14 輸入的承辦人
'pTF24,pTF25 輸入的外文本和圖示頁數
'pTFile 工作確認單範本路徑
Dim inX As Integer, i As Integer, intQ As Integer
Dim rsAD As New ADODB.Recordset
Dim objOutLook As Object
Dim objMail As Object
Dim strName As String, strText As String
Dim m_TempFileName As String
Dim strFile As String
Dim strAttFList As String '附件檔案路徑(多筆)
Dim fs, f
Dim strSubject As String, strContent As String 'Email主旨, 內文
Dim strConUser As String '台一聯絡人
Dim strConUserNo As String 'Added by Lydia 2021/09/09 台一聯絡人的員工編號
Dim strPWD As String
Dim strA01 As String, strA02 As String
Dim strTemp(0 To 21) As String
Dim strCase(1 To 4) As String '本所案號
Dim strTF19 As String, strTF20 As String '相似度,相似案號
Dim bolGetNew As Boolean '是否取得卷宗區最新一道ORI.PDF
Dim bolGetSeQ As Boolean 'Added by Lydia 2018/07/09 取得序列表或密碼
Dim strTCN01 As String 'Added by Lydia 2018/12/22 急件翻譯號
Dim tmpArr As Variant
Dim strResList As String 'Added by Lydia 2019/02/22 相似比對結果檔名
Dim m_TCT97 As String 'Added by Lydia 2019/08/21 (命名記錄)指定代表圖TCT97
Dim m_TF23 As String 'Added by Lydia 2019/08/21 原文字數
Dim m_TF36 As String 'Added by Lydia 2019/08/23 翻譯特殊指示
Dim strContAdd As String 'Added by Lydia 2021/10/08 特定Email內文
Dim strTmp1(0 To 10) As String, rsQD As New ADODB.Recordset 'Added by Lydia 2024/03/29

On Error GoTo ErrHand
   
   '抓相關資料
   'Modified by Lydia 2018/12/22 抓急件翻譯號tcn01
   'strA02 = "select '20' ord1,c1.cp01,c1.cp02,c1.cp03,c1.cp04 " & _
                 ",decode(pa49||pa50||fa25||fa26||cu36||cu37,null,'','Y') as dcprice,decode(x01||y01,null,'','Y') as fcprice " & _
                 ",nvl(pa05,nvl(pa06,pa07)) casename,pa150,pa75,tf01 as c1_cp09,c2.cp09 as c2_cp09,c2.cp27 as c2_cp27,c1.cp14 as c1_cp14,c1.cp48 as c1_cp48,a01.* " & _
                 "from TransFee a01,CaseProgress c1,CaseProgress c2 ,TransCaseTitle,patent,customer,fagent " & _
                 ",(select aal04 as x01 from addressa4list where aal01='FCPtct' and substr(aal04,1,1)='X') vtb1 " & _
                 ",(select aal04 as y01 from addressa4list where aal01='FCPtct' and substr(aal04,1,1)='Y') vtb2 " & _
                 "where tf01=c1.cp09(+) and c1.cp159=0 and c1.cp09=" & CNULL(pTF01) & _
                 "and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+) " & _
                 "and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+)and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and c2.cp31='Y' " & _
                 "and c2.cp09=tct01(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
                 "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) " & _
                 "and pa26=x01(+) and pa75=y01(+) "
   'strA02 = strA02 & "union select '01' ord1,tcn14 as cp01,null, null, null ,null as dcprice,null as fcprice,null as casename,null as pa150, " & _
                 "null as pa75,null as c1_cp09,null as c2_cp09, null as c2_cp27,tcn15 as c1_cp14, null as c1_cp48, a01.* " & _
                 "from TrackingCaseName,TransFee a01 where nvl(tcn14,'N') <'A'  and tcn14=tf01(+) and tcn14=" & CNULL(pTF01)
   'Modified by Lydia 2019/08/21 +(命名記錄)指定代表圖TCT97
   'Modified by Lydia 2022/04/18 +(命名記錄)提申前主動修正TCT117=2
   'Modified by Lydia 2023/03/16 +PA10申請日
   strA02 = "select '20' ord1,c1.cp01,c1.cp02,c1.cp03,c1.cp04 " & _
                 ",decode(pa49||pa50||fa25||fa26||cu36||cu37,null,'','Y') as dcprice,tcn01" & _
                 ",nvl(pa05,nvl(pa06,pa07)) casename,pa150,pa75,tf01 as c1_cp09,c2.cp09 as c2_cp09,c2.cp27 as c2_cp27,c1.cp14 as c1_cp14,c1.cp48 as c1_cp48,a01.*,TCT97,TCT117,PA10 " & _
                 "from TransFee a01,CaseProgress c1,CaseProgress c2 ,TransCaseTitle,patent,customer,fagent,TrackingCaseName " & _
                 "where tf01=c1.cp09(+) and c1.cp159=0 and c1.cp09=" & CNULL(pTF01) & _
                 "and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+) " & _
                 "and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+)and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) and c2.cp31='Y' " & _
                 "and c2.cp09=tct01(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
                 "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and c1.cp09=tcn14(+)"
   strA02 = strA02 & " union select '01' ord1,tcn14 as cp01,null, null, null ,null as dcprice,null as tcn01,null as casename,null as pa150, " & _
                 "null as pa75,null as c1_cp09,null as c2_cp09, null as c2_cp27,tcn15 as c1_cp14, null as c1_cp48, a01.*,'' as TCT97,'' as TCT117,0 AS PA10 " & _
                 "from TrackingCaseName,TransFee a01 where nvl(tcn14,'N') <'A'  and tcn14=tf01(+) and tcn14=" & CNULL(pTF01)
   'end 2018/12/22
   'Added by Lydia 2024/03/08 因應內專支援機械組OA (含P案) 收文: 927其他翻譯
   strA02 = strA02 & " union select '02' ord1,c1.cp01,c1.cp02,c1.cp03,c1.cp04,decode(pa49||pa50||fa25||fa26||cu36||cu37,null,'','Y') as dcprice" & _
            " ,null as tcn01,nvl(pa05,nvl(pa06,pa07)) casename,pa150,pa75,tf01 as c1_cp09,c2.cp09 as c2_cp09,c2.cp27 as c2_cp27,c1.cp14 as c1_cp14,c1.cp48 as c1_cp48" & _
            " ,a01.*,null as tct97,null tct117,pa10" & _
            " from transfee a01,caseprogress c1,caseprogress c2 ,patent,customer,fagent" & _
            " where tf01=c1.cp09(+) and c1.cp159=0 and c1.cp09='" & pTF01 & "' and c1.cp10='927' and c1.cp43=c2.cp09(+)" & _
            " and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) "
   'end 2024/03/08
   strA02 = strA02 & " order by ord1"
   intQ = 0
   Set rsAD = ClsLawReadRstMsg(intQ, strA02)
   If intQ = 1 Then
        '案號/追蹤號
        If Val("" & rsAD.Fields("cp01")) > 0 Then
            strTemp(0) = "" & rsAD.Fields("cp01")
        Else
            strCase(1) = "" & rsAD.Fields("cp01")
            strCase(2) = "" & rsAD.Fields("cp02")
            strCase(3) = "" & rsAD.Fields("cp03")
            strCase(4) = "" & rsAD.Fields("cp04")
            'Modified by Lydia 2021/08/27
            'strTemp(0) = strCase(1) & Val(strCase(2)) & IIf(strCase(3) & strCase(4) <> "000", strCase(3) & strCase(4), "")
            strTemp(0) = strCase(1) & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", strCase(3) & strCase(4), "")
            'Added by Lydia 2019/08/21
            m_TCT97 = "" & rsAD.Fields("TCT97")
            m_TF23 = "" & rsAD.Fields("TF23")
            'Added by Lydia 2019/08/23
            m_TF36 = "" & rsAD.Fields("TF36")
        End If
        'Added by Lydia 2018/12/22 改模組判斷固定報價
        strTemp(1) = Pub_GetPa62Flag(strCase(1) & strCase(2) & strCase(3) & strCase(4))
        strTCN01 = "" & rsAD.Fields("tcn01")
        'end 2018/12/22
        'Added by Lydia 2024/03/08 因應內專支援機械組OA (含P案) 收文: 927其他翻譯
        If "" & rsAD.Fields("ord1") = "02" Then
           pType = "A"
           strTCN01 = "" & rsAD.Fields("c2_cp09") '相關收文號
        End If
        'end 2024/03/08
        
        '是否可以上班/下班翻譯
        'Modified by Lydia 2018/12/22
        'If Trim("" & rsAD.Fields("dcprice") & rsAD.Fields("fcprice")) <> "" Then
        'Modified by Lydia 2019/08/13 自2019年8月15日起實施，亦即自當日起交稿案件一律以調整後費率計算; 並且取消折扣案件之限制
        'If Trim("" & rsAD.Fields("dcprice") & strTemp(1)) <> "" Then
        If Trim("" & rsAD.Fields("dcprice") & strTemp(1)) <> "" And strSrvDate(1) < "20190815" Then
             If strTemp(1) = "Y" Then
                 strTemp(21) = strTemp(21) & "有固定報價,"
             End If
        'end 2018/12/22
             strTemp(1) = "上班"
             '列出說明有折扣或固定報價
             If "" & rsAD.Fields("dcprice") <> "" Then
                   strTemp(21) = strTemp(21) & "有折扣,"
             End If
             'Remove by Lydia 2018/12/22
             'If "" & rsAD.Fields("fcprice") <> "" Then
             '      strTemp(21) = strTemp(21) & "有固定報價,"
             'End If
        Else
             strTemp(1) = "上班/下班"
        End If
        '原文語種=語系
        If "" & rsAD.Fields("tf27") = "" Then
            Select Case "" & rsAD.Fields("pa150")
                Case "1", "2", "4": strTemp(2) = "英文" '電子電機組、化學組、機械設計組
                Case "3": strTemp(2) = "日文"  '日文組
                Case Else: strTemp(2) = "" & rsAD.Fields("pa75")
            End Select
        Else
            strTemp(2) = Pub_GetTransFeeL("1", "" & rsAD.Fields("tf27"))
        End If
        '翻譯語種=語系2
        If "" & rsAD.Fields("tf28") = "" Then
            If strCase(1) = "P" Then
                 strTemp(3) = "簡體中文"
            Else
                 strTemp(3) = "繁體中文"
            End If
        Else
            strTemp(3) = Pub_GetTransFeeL("2", "" & rsAD.Fields("tf28"))
        End If
        '相似度
        strTF19 = "" & rsAD.Fields("tf19")
        '相似案號
        strTF20 = "" & rsAD.Fields("tf20")
        '專利名稱
        strTemp(10) = "" & rsAD.Fields("casename")
        '外文本頁數=頁數
        If Val("" & rsAD.Fields("tf24")) > 0 Then
            strTemp(11) = Val("" & rsAD.Fields("tf24"))
        Else
            strTemp(11) = Val(pTF24)
        End If
        '外文本圖示=頁圖
        If Val("" & rsAD.Fields("tf25")) > 0 Then
            strTemp(12) = Val("" & rsAD.Fields("tf25"))
        Else
            strTemp(12) = Val(pTF25)
        End If
        '承辦期限和交稿期限比對，取較早期限
        If "" & rsAD.Fields("c1_cp48") <> "" And "" & rsAD.Fields("tf26") <> "" Then
             If Val("" & rsAD.Fields("c1_cp48")) > Val("" & rsAD.Fields("tf26")) Then
                  strTemp(13) = Val("" & rsAD.Fields("tf26"))
             Else
                  strTemp(13) = Val("" & rsAD.Fields("c1_cp48"))
             End If
        Else
             If "" & rsAD.Fields("c1_cp48") <> "" Then
                  strTemp(13) = "" & rsAD.Fields("c1_cp48")
             Else
                  strTemp(13) = "" & rsAD.Fields("tf26")
             End If
        End If
        '只交Claims期限
        strTemp(18) = "" & rsAD.Fields("tf32")
        '新申請案發文日
        strTemp(15) = "" & rsAD.Fields("c2_cp27")
        '英文參考本
        strTemp(16) = "" & rsAD.Fields("tf30")
        
        strTmp1(7) = "" & rsAD.Fields("TCT117") 'Added by Lydia 2022/04/18 (命名記錄)提申前主動修正
        strTmp1(8) = "" & rsAD.Fields("PA10") 'Added by Lydia 2023/03/16 申請日(已提申)
        
        '判斷預設／輸入承辦人
        'Modified by Lydia 2023/04/28 離職員工在薪資結清前的外譯編號部門仍屬於F52,另外判斷對應F21的員工編號是否在職
        'strA01 = " select st01,st02,st03,st04,st09,st10,st18,st65 from staff where st01='" & IIf(pCP14 <> "", pCP14, "" & rsAD.Fields("c1_cp14")) & "' "
        strA01 = "select s1.st01,s1.st02,s1.st03,s1.st04,s1.st09,s1.st10,s1.st18,s1.st65,sim01,s2.st04 as sim01st04 " & _
                      "from staff s1,staff_idmap,staff s2 where s1.st01='" & IIf(pCP14 <> "", pCP14, "" & rsAD.Fields("c1_cp14")) & "' and s1.st01=sim02(+) and sim01=s2.st01(+) "
        intQ = 1
        Set rsAD = ClsLawReadRstMsg(intQ, strA01)
        If intQ = 1 Then
            If "" & rsAD.Fields("st04") <> "1" Then
               MsgBox "承辦人[" & rsAD.Fields("st01") & " " & rsAD.Fields("st02") & "]已離職，請輸入其他人員 !", vbCritical
               Exit Sub
            Else
               strTemp(4) = "" & rsAD.Fields("st01")
               strTemp(5) = "" & rsAD.Fields("st02")
               strTemp(17) = "" & rsAD.Fields("st03")  '部門
               'Added by Lydia 2023/04/28 離職員工在薪資結清前的外譯編號部門仍屬於F52,另外判斷對應F21的員工編號是否在職，若不在職視為所外翻譯人員F51。
                                                       'ex.F5523 ; 要注意員工在F21有多個編號時,staff_idmap的對應編號sim01是否在職
               If strTemp(17) = "F52" And "" & rsAD.Fields("sim01") <> "" And "" & rsAD.Fields("sim01st04") = "2" Then
                  strTemp(17) = "F51"
               End If
               'end 2023/04/28
               'Added by Lydia 2024/06/19 約定薪資的翻譯人員比照外翻人員的規則
               If Right("" & rsAD.Fields("sim01"), 2) >= "9A" Then
                  strTemp(17) = "F51"
               End If
               'end 2024/06/19
               
               '收件人email=外翻email=所內員工信箱
               If strTemp(17) = "F51" Then
                    strTemp(9) = "" & rsAD.Fields("st18")
               Else
                    If Left(strTemp(4), 1) = "F" Then
                         strTemp(9) = PUB_GetMapID(strTemp(4))
                    Else
                         strTemp(9) = strTemp(4)
                    End If
               End If
               'Modified by Lydia 2025/03/13 改用模組取得
               'If strTemp(4) = 外翻_舜禹 Then
               '   strTemp(5) = "江蘇舜禹翻譯"
               'ElseIf strTemp(4) = 外翻_捷恩凱 Then
               '   strTemp(5) = "南京捷恩凱信息技術"
               'ElseIf strTemp(4) = 外翻_迅達 Then
               '   strTemp(5) = "迅達翻譯"
               'End If
               strTmp1(5) = Pub_SetF51Order("T", strTemp(4))
               If strTmp1(5) <> strTemp(4) Then
                  strTemp(5) = strTmp1(5) '判斷為外翻
               End If
               'end 2025/03/13
               
               '外翻聯絡1,2; email稱謂
               strA01 = "" & rsAD.Fields("st65")
               If strA01 <> "" And InStr(strA01, " ") > 0 Then
                   strTemp(6) = Mid(strA01, 1, InStr(strA01, " ") - 1)
                   strTemp(14) = Mid(strA01, 1, 1) & Mid(strA01, InStr(strA01, " ") + 1)
               Else
                    If strA01 <> "" Then
                         strTemp(6) = strA01
                    Else
                         strTemp(6) = strTemp(5)
                    End If
                    strTemp(14) = strTemp(6)
               End If
               '外翻電話
               'Modified by Lydia 2025/03/13 改用模組取得
               'If strTemp(4) = 外翻_舜禹 Then
               If Pub_SetF51Order("", strTemp(4)) = "1" Then
                   strTemp(7) = "" & rsAD.Fields("st09") & vbCrLf & "84699966，" & vbCrLf & "84699955 # 808"
               Else
                   strTemp(7) = "" & rsAD.Fields("st09")
               End If
               '外翻傳真
               strTemp(8) = "" & rsAD.Fields("st10")
            End If
        End If
   Else
      Exit Sub
   End If
   
   strA01 = Pub_GetSpecMan("M")
   strConUserNo = strA01  'Added by Lydia 2021/09/09 台一聯絡人的員工編號
   strA02 = "select st01,st02,st03,st04,st05,st06,st07,st22 from staff where st01='" & strA01 & "' "
   intQ = 0
   Set rsAD = ClsLawReadRstMsg(intQ, strA02)
   If intQ = 1 Then
      If rsAD.Fields("st04") = "2" Then
         MsgBox "工作清單中的台一聯絡人: " & rsAD.Fields("st02") & " 已離職，請自行修改工作清單!"
      End If
        strTemp(19) = Trim("" & rsAD.Fields("st02"))
        strTemp(20) = PUB_chgWord2Num("" & rsAD.Fields("st07"))
        strConUser = Trim("" & rsAD.Fields("st02"))
        'Mark by Lydia 2021/09/09 直接用全名
        'If Len(strConUser) < 4 Then
        '   strConUser = Left(strConUser, 1) & IIf("" & rsAD.Fields("st22") = "F", "小姐", "先生")
        'Else
        '   strConUser = Left(strConUser, 2) & IIf("" & rsAD.Fields("st22") = "F", "小姐", "先生")
        'End If
        'end 2021/09/09
   Else
      Exit Sub
   End If
   
   '國外譯者:已收文案件需夾帶台一工作通知單
   'Modified by Lydia 2025/03/13 改用模組取得
   'If pType <> "0" And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, strTemp(4)) > 0 Then
   If pType <> "0" And InStr(Pub_SetF51Order("F", ""), strTemp(4)) > 0 Then
       '判斷word是否已開啟
       If g_WordAp Is Nothing Then
RestarWord:
          Set g_WordAp = New Word.Application
          g_WordAp.Visible = False
       End If
       m_TempFileName = strTemp(0) & "工作通知單.doc"
       
       If Dir(pSavePath & "\" & m_TempFileName) <> "" Then
          Kill pSavePath & "\" & m_TempFileName
       End If
       
       '範本
       If Dir(pTFile) = "" Then
           'Modified by Lydia 2023/08/18
           'Call PUB_GetSampleFile("外專翻譯_案件工作確認單樣本.doc", "M51-000299-0-01")
           'pTFile = App.path & "\外專翻譯_案件工作確認單樣本.doc"
           Call PUB_GetSampleFile("外專翻譯_案件工作確認單樣本.doc", "M51-000299-0-01", , pSavePath)
           pTFile = pSavePath & "\外專翻譯_案件工作確認單樣本.doc"
           'end 2023/08/18
       End If
       g_WordAp.Documents.Open pTFile
       
       g_WordAp.ActiveDocument.SaveAs pSavePath & "\" & m_TempFileName
       g_WordAp.ActiveDocument.Close
       g_WordAp.Documents.Open pSavePath & "\" & m_TempFileName
       'strAttFList = strAttFList & pSavePath & "\" & m_TempFileName & vbCrLf 'Mark by Lydia 2018/07/09  移到最後面插入by Sharon
       With g_WordAp
          .Selection.WholeStory
          .Selection.Copy
          'Modified by Lydia 2020/03/31 17=>18
          'Modified by Lydia 2022/07/21 18=>19
          For i = 0 To 19
             strName = ""
             strText = ""
             If i = 0 Then
                strName = "外翻公司"
                strText = strTemp(5)
             ElseIf i = 1 Then
                strName = "外翻聯絡1"
                strText = strTemp(6)
             ElseIf i = 2 Then
                strName = "外翻電話"
                strText = strTemp(7)
             ElseIf i = 3 Then
                strName = "外翻傳真"
                strText = strTemp(8)
             ElseIf i = 4 Then
                strName = "外翻email"
                strText = strTemp(9)
             ElseIf i = 5 Then
                strName = "本所案號"
                strText = strTemp(0)
             ElseIf i = 6 Then
                strName = "專利名稱"
                strText = strTemp(10)
             ElseIf i = 7 Then
                strName = "語系"
                'Added by Lydia 2024/03/08 中翻英
                If pType = "A" Then
                   strText = strTemp(3)
                Else
                'end 2024/03/08
                   strText = strTemp(2)  '英翻中
                End If
             ElseIf i = 8 Then
                strName = "頁數"
                'Modified by Lydia 2018/08/22 頁數為說明書頁數+圖示頁數
                'strText = strTemp(11) & "頁"
                strText = Val(strTemp(11)) + Val(strTemp(12)) & "頁"
             ElseIf i = 9 Then
                strName = "頁圖"
                strText = strTemp(12)
             ElseIf i = 10 Then
                strName = "承辦期限"
                strText = ChangeTStringToTDateString(TransDate(strTemp(13), 1))
             ElseIf i = 11 Then
                strName = "外翻聯絡2"
                strText = strTemp(6)
             ElseIf i = 12 Then
                strName = "發信日"
                strText = ChangeWStringToWDateString(strSrvDate(1))
             ElseIf i = 13 Then
                strName = "語系2"
                'Added by Lydia 2024/03/08 中翻英
                If pType = "A" Then
                    strText = strTemp(2)
                Else
                'end 2024/03/08
                    strText = strTemp(3) '英翻中
                End If
             ElseIf i = 14 Then
                strName = "台一聯絡人"
                strText = strTemp(19)
             ElseIf i = 15 Then
                strName = "台一分機"
                strText = strTemp(20)
             ElseIf i = 16 Then
                strName = "台一收信"
                strText = "idept@taie.com.tw"
             ElseIf i = 17 Then
                strName = "台一聯絡人稱"
                strText = strConUser
             'Added by Lydia 2020/03/31 公司Title改抓公司檔
             ElseIf i = 18 Then
                strName = "A0802"
                strText = CompNameQuery("2")
             'Added by Lydia 2022/07/21
             ElseIf i = 19 Then
                strName = "原文字數"
                'Modified by Lydia 2025/03/13 改用模組取得
                'If 外翻_迅達 = strTemp(4) Then   '迅達需要帶入原文字數
                'Modified by Lydia 2025/03/18 百靈也要帶入原文字數
                'If Pub_SetF51Order("", strTemp(4)) = "3" Then
                strA01 = Pub_SetF51Order("", strTemp(4))
                If strA01 = "3" Or strA01 = "4" Then
                'end 2025/03/18
                    strText = m_TF23
                Else
                    strText = ""
                End If
             End If
             If Trim(strName) <> "" Then
                .Selection.Find.ClearFormatting
                .Selection.Find.Text = "|#" & strName & "#|"
                .Selection.Find.Replacement.Text = ""
                .Selection.Find.Forward = True
                .Selection.Find.Wrap = wdFindContinue
                .Selection.Find.Format = False
                .Selection.Find.MatchCase = False
                .Selection.Find.MatchWholeWord = False
                .Selection.Find.MatchWildcards = False
                .Selection.Find.MatchSoundsLike = False
                .Selection.Find.MatchAllWordForms = False
                .Selection.Find.MatchByte = True
                .Selection.Find.Execute
                .Selection.Delete
                .Selection.TypeText strText
             End If
          Next i
       End With
       g_WordAp.ActiveDocument.Save
       g_WordAp.ActiveDocument.Close
       Clipboard.Clear '清除剪貼簿動作
   End If

   inX = 0 '記錄檔案數量
    '下載附件到本機端
   bolGetNew = True
   bolGetSeQ = False  'Added by Lydia 2018/07/09
    '急件翻譯(尚未收文)按下E-Mail：預設抓Tracking_NO資料夾下的*.ORI.PDF做為附件；為了確認外翻是否能接案件，可能會發送多次。
    If pType = "0" Then
            strA01 = "select cpf01,cpf02,cpf13 from casepaperfile where cpf01='" & Val(strTemp(0)) & "' AND (UPPER(CPF02) LIKE '%.ORI.PDF') and cpf10<>'D' and substr(upper(cpf02),-4)<>upper('.del') "
            intQ = 1
            Set rsQD = ClsLawReadRstMsg(intQ, strA01)
            If intQ = 1 Then
               rsQD.MoveFirst
               Do While Not rsQD.EOF
                  If "" & rsQD.Fields("CPF01") <> "" And "" & rsQD.Fields("CPF02") <> "" And "" & rsQD.Fields("CPF13") <> "" Then
                     strFile = pSavePath & "\" & rsQD.Fields("CPF02")  '下載檔案名稱+路徑
                     If PUB_GetFtpFile("" & rsQD.Fields("CPF13"), strFile, "CASEPAPERFILE") = True Then
                        inX = inX + 1
                        strAttFList = strAttFList & strFile & vbCrLf
                        bolGetNew = False
                     End If
                  End If
                  rsQD.MoveNext
               Loop
            End If

         bolGetNew = False
         
    '急件翻譯(已收文)按下分案：只夾帶台一工作通知單。
    ElseIf pType = "1" Then
         bolGetNew = False
         
    '未提申先翻譯：不會有英文參考本，有主動修正203從電子送件暫存區抓FCPxxxxx.ori.pdf，無主動修正203從Eng.vers.抓*.ori.pdf。
                                '沒有可以抓最後一道.ori.pdf；P案也納入翻譯分案作業，因為要在提申前完成翻譯，所以走未提申先翻譯的流程。
    'Modified by Lydia 2023/03/23 FCP排除案件已提申,P案不用 ; ex.FCP69186在3/16提申後進行翻譯分案，同時抓原始檔區和卷宗區的ORI
    'ElseIf pType = "3" Then
    ElseIf pType = "3" And (strCase(1) = "P" Or (strCase(1) = "FCP" And Val(strTmp1(8))) = 0) Then
         'Modified by Lydia 2019/02/19 P案都抓english_vers (ex.P-122128)
         'If PUB_ChkCPExist(strCase, "203") = True Then
         'Modified by Lydia 2022/04/18 FCP案若有工程師收文之提申前主動修正，才從電子送件暫存區抓ori.pdf (ex.FCP-067005,FCP-067006)；其他和P案都從原始檔區Eng.vers.抓*.ori.pdf --- from Sharon
         'If PUB_ChkCPExist(strCase, "203") = True And strCase(1) = "FCP" Then
         strTmp1(5) = ""
         If strCase(1) = "FCP" And strTmp1(7) = "2" Then
             If PUB_ChkCPExist(strCase, "203", , , , "B") = True Then
                 strTmp1(5) = "Y"
             End If
         End If
         If strTmp1(5) = "Y" Then
         'end 2022/04/18
            strA02 = PUB_FCPCaseNo2FileName(strCase(1), strCase(2), strCase(3), strCase(4))
            'Modified by Lydia 2024/07/29 改用變數
            'strA01 = Dir("\\Typing2\電子送件暫存區\" & strA02 & "\*.ori.pdf")
            strA01 = Dir("\\" & strTyping2Path & "\電子送件暫存區\" & strA02 & "\*.ori.pdf")
            If strA01 <> "" Then
                 'Modified by Lydia 2024/07/29 改用變數
                 'FileCopy "\\Typing2\電子送件暫存區\" & strA02 & "\" & strA01, pSavePath & "\" & strA01
                 FileCopy "\\" & strTyping2Path & "\電子送件暫存區\" & strA02 & "\" & strA01, pSavePath & "\" & strA01
                 strAttFList = strAttFList & pSavePath & "\" & strA01 & vbCrLf
                 bolGetNew = False
            End If
            'Added by Lydia 2022/07/22 未提申先翻譯一併附上序列表.SEQ.
            'Mark by Lydia 2023/05/08 因應智慧局4/25起對序列表翻譯的變更,無需再抓取檔案, 卷宗區.SEQ.pdf及原始檔區.xml檔案
            'strA01 = Dir("\\Typing2\電子送件暫存區\" & strA02 & "\*.seq.*")
            'If strA01 <> "" Then
            '     FileCopy "\\Typing2\電子送件暫存區\" & strA02 & "\" & strA01, pSavePath & "\" & strA01
            '     strAttFList = strAttFList & pSavePath & "\" & strA01 & vbCrLf
             '    bolGetNew = False
            'End If
            ''end 2022/07/22
         Else
            'Added by Lydia 2020/02/24 English_Vers檔案：放在原始檔區，記錄收文號
            If PUB_ChkCPExist(strCase, cntEnglish_Vers, , strA02, , "D") = True Then
                'Modified by Lydia 2022/07/22 未提申先翻譯一併附上序列表.SEQ. ; ex.FCP-67581
                'strA01 = "SELECT CPF01,CPF02,CPF13 FROM CASEPROGRESS A,CASEPAPERFILE B " & _
                                  "WHERE CP09='" & strA02 & "' AND CP159=0 AND CP09=CPF01(+) " & _
                                  "AND NVL(CPF10,'N') <> 'D' AND UPPER(CPF02) LIKE '%.ORI.PDF' " & _
                                  "ORDER BY CPF06 DESC, CPF07 DESC "
                'Modified by Lydia 2023/05/08 因應智慧局4/25起對序列表翻譯的變更,無需再抓取檔案, 卷宗區.SEQ.pdf及原始檔區.xml檔案
                'strA01 = "SELECT CPF01,CPF02,CPF13 FROM CASEPROGRESS A,CASEPAPERFILE B " & _
                                  "WHERE CP09='" & strA02 & "' AND CP159=0 AND CP09=CPF01(+) " & _
                                  "AND NVL(CPF10,'N') <> 'D' AND (UPPER(CPF02) LIKE '%.ORI.PDF' OR UPPER(CPF02) LIKE '%.SEQ.%')" & _
                                  "ORDER BY CPF06 DESC, CPF07 DESC "
                strA01 = "SELECT CPF01,CPF02,CPF13 FROM CASEPROGRESS A,CASEPAPERFILE B " & _
                                  "WHERE CP09='" & strA02 & "' AND CP159=0 AND CP09=CPF01(+) " & _
                                  "AND NVL(CPF10,'N') <> 'D' AND (UPPER(CPF02) LIKE '%.ORI.PDF')" & _
                                  "ORDER BY CPF06 DESC, CPF07 DESC "
                intQ = 1
                Set rsQD = ClsLawReadRstMsg(intQ, strA01)
                If intQ = 1 Then
                     rsQD.MoveFirst
                     Do While Not rsQD.EOF
                          If "" & rsQD.Fields("CPF01") <> "" And "" & rsQD.Fields("CPF02") <> "" And "" & rsQD.Fields("CPF13") <> "" Then
                                strFile = pSavePath & "\" & rsQD.Fields("CPF02")  '下載檔案名稱+路徑
                                If PUB_GetFtpFile("" & rsQD.Fields("CPF13"), strFile, "CASEPAPERFILE") = True Then
                                    inX = inX + 1
                                    strAttFList = strAttFList & strFile & vbCrLf
                                    bolGetNew = False 'Added by Lydia 2023/03/16
                                End If
                          End If
                          rsQD.MoveNext
                     Loop
                End If
            Else
            'end 2020/02/24
                'Remove by Lydia 2021/12/06 (109/4/6)已將\\Typing2的"English_Vers"和"專利案件"的案件資料夾，全部搬到原始檔區
'                strA02 = Pub_GetFCPcaseFilePath(strCase(2), , strCase(1))
'                'Modified by Lydia 2019/02/12 有多個檔案需要做排序(ex.P121918翻譯分案應抓最後的.rep1.ori.pdf)
'                'strA01 = Dir(strA02 & "\*.ori.pdf")
'                strA01 = PUB_GetFileListOrderby(strA02, "*.ori.pdf", True)
'                If strA01 <> "" And InStr(strA01, "||") > 0 Then '用||區隔多筆
'                    strA01 = Mid(strA01, 1, InStr(strA01, "||") - 1)
'                End If
'                'end 2019/02/12
'                If strA01 <> "" Then
'                     FileCopy strA02 & "\" & strA01, pSavePath & "\" & strA01
'                     strAttFList = strAttFList & pSavePath & "\" & strA01 & vbCrLf
'                     bolGetNew = False
'                End If
                'end 2021/12/06
            End If 'Added by Lydia 2020/02/24
         End If
         bolGetSeQ = True 'Added by Lydia 2018/07/09
         
    '英文參考本：新案建檔有設「英文本收文號」，抓該收文號在卷宗區掛的sep.pdf。沒有就沒附件
    ElseIf pType = "4" Then
         If strTemp(16) <> "" And strTemp(16) <> "Y" Then
            'Move by Lydia 2019/02/23 從下方移過來
            bolGetNew = False '抓英文本收文號在卷宗區掛的SEP.pdf和DWG.PDF, 不抓其他
            bolGetSeQ = True '有序列表就抓
            'end 2019/02/23
            'Modified by Lydia 2018/08/08 因為有時sep.pdf只有文字，所以加抓圖檔(DWG.pdf)
            'Modified by Morgan 2025/3/27 +cpp19
            strA01 = "SELECT CPP01,CPP02,CPP14,CPP19 FROM CASEPROGRESS A,CASEPAPERPDF B " & _
                              "WHERE CP09='" & strTemp(16) & "' AND CP159=0 AND CP09=CPP01(+) " & _
                              "AND NVL(CPP10,'N') <> 'D' AND (UPPER(CPP02) LIKE '%.SEP.PDF' OR UPPER(CPP02) LIKE '%.DWG.PDF') " & _
                              "ORDER BY CPP06 DESC, CPP07 DESC "
            intQ = 1
            Set rsQD = ClsLawReadRstMsg(intQ, strA01)
            If intQ = 1 Then
                 rsQD.MoveFirst
                 Do While Not rsQD.EOF
                      If "" & rsQD.Fields("CPP01") <> "" And "" & rsQD.Fields("CPP02") <> "" And "" & rsQD.Fields("CPP14") <> "" Then
                            strFile = pSavePath & "\" & rsQD.Fields("CPP02")  '下載檔案名稱+路徑
                            If PUB_GetFtpFile("" & rsQD.Fields("CPP14"), strFile, , , , "" & rsQD.Fields("CPP19") <> "") = True Then
                                inX = inX + 1
                                strAttFList = strAttFList & strFile & vbCrLf
                            End If
                      End If
                      rsQD.MoveNext
                 Loop
                 'Mark by Lydia 2019/02/23 FCP-060255收文英文參考本但是PDF檔未上卷宗區,發現判斷有誤
            End If
         End If
    'Added by Lydia 2024/03/08
    ElseIf pType = "A" Then
         bolGetNew = False
         '下載官方來函
         'Modified by Morgan 2025/3/27 +cpp19
         strA01 = "SELECT CPP01,CPP02,CPP14,CPP19 FROM CASEPROGRESS A,CASEPAPERPDF B " & _
                           "WHERE CP09='" & strTCN01 & "' AND CP159=0 AND CP09=CPP01(+) " & _
                           "AND NVL(CPP10,'N') <> 'D' AND INSTR(UPPER(CPP02),UPPER(CP10||'.PDF')) > 0 " & _
                           "ORDER BY CPP06 DESC, CPP07 DESC "
         intQ = 1
         Set rsQD = ClsLawReadRstMsg(intQ, strA01)
         If intQ = 1 Then
              rsQD.MoveFirst
              Do While Not rsQD.EOF
                   If "" & rsQD.Fields("CPP01") <> "" And "" & rsQD.Fields("CPP02") <> "" And "" & rsQD.Fields("CPP14") <> "" Then
                         strFile = pSavePath & "\" & rsQD.Fields("CPP02")  '下載檔案名稱+路徑
                         If PUB_GetFtpFile("" & rsQD.Fields("CPP14"), strFile, , , , "" & rsQD.Fields("CPP19") <> "") = True Then
                             inX = inX + 1
                             strAttFList = strAttFList & strFile & vbCrLf
                         End If
                   End If
                   rsQD.MoveNext
              Loop
         End If
    'end 2024/03/08
    End If
    
    '其他(一般翻譯):抓卷宗區最後上傳的外文本(*.ORI.PDF、*.ORI.REP.PDF、*.ORI.FIX.PDF)，若有密碼檔(.PWD.)或序列表(.SEQ.)一併附上；因為第2次發信的狀態可能有所不同，所以撰寫信函(frm090401)依舊抓最後一道.ORI.PDF。
    'Modified by Lydia 2018/07/09 +取得序列表
    'If bolGetNew = True Then
    If bolGetNew = True Or bolGetSeQ = True Then
        '從卷宗區抓資料
        strA02 = ""
        'Modified by Lydia 2018/07/09 分成抓卷宗區最後上傳的外文本+其他(SEQ, PWD) 或只抓其他(SEQ, PWD)
        If bolGetNew = True Then
             'Modified by Lydia 2018/09/18 ORI.FIX=>改成FIX.ORI
             'Modified by Lydia 2018/10/02 +.TBL.
             'Modified by Lydia 2018/11/30 判斷最後一道.ORI.%.PDF ,因為.FIX有人加後面
             'strA01 = "AND (UPPER(CPP02) LIKE '%.ORI.PDF' OR UPPER(CPP02) LIKE '%.ORI.REP%.PDF' " & _
                          "OR UPPER(CPP02) LIKE '%.FIX%.ORI.PDF' OR UPPER(CPP02) LIKE '%.SEQ.%' OR UPPER(CPP02) LIKE '%.PWD.%' OR UPPER(CPP02) LIKE '%.TBL.%' ) "
             'Modified by Lydia 2019/10/24 拿掉 OR UPPER(CPP02) LIKE '%.PWD.%'
             'Modified by Lydia 2023/05/08 因應智慧局4/25起對序列表翻譯的變更,無需再抓取檔案, 卷宗區.SEQ.pdf及原始檔區.xml檔案=>拿掉OR UPPER(CPP02) LIKE '%.SEQ.%'
             strA01 = "AND ((UPPER(CPP02) LIKE '%.ORI.%' AND UPPER(CPP02) LIKE '%.PDF' )  " & _
                          "OR UPPER(CPP02) LIKE '%.TBL.%' ) "
        ElseIf bolGetSeQ = True Then
             'Modified by Lydia 2018/10/02 +.TBL.
             'Modified by Lydia 2019/10/24 拿掉 OR UPPER(CPP02) LIKE '%.PWD.%'
             'Modified by Lydia 2023/05/08 因應智慧局4/25起對序列表翻譯的變更,無需再抓取檔案, 卷宗區.SEQ.pdf及原始檔區.xml檔案=>拿掉UPPER(CPP02) LIKE '%.SEQ.%' OR
             strA01 = "AND (UPPER(CPP02) LIKE '%.TBL.%' ) "
        End If
        
        'Modified by Morgan 2025/3/27 +CPP19
        strA01 = "SELECT CPP01,CPP02,CPP14,CPP19 FROM CASEPROGRESS A,CASEPAPERPDF B " & _
                          "WHERE CP01='" & strCase(1) & "' AND CP02='" & strCase(2) & "' AND CP03='" & strCase(3) & "' AND CP04='" & strCase(4) & "' AND CP159=0 AND CP09=CPP01(+) " & _
                          "AND NVL(CPP10,'N') <> 'D' " & strA01 & _
                          "ORDER BY CPP06 DESC, CPP07 DESC "
        'end 2018/07/09
        intQ = 1
        Set rsQD = ClsLawReadRstMsg(intQ, strA01)
        If intQ = 1 Then
             rsQD.MoveFirst
             Do While Not rsQD.EOF
                  If "" & rsQD.Fields("CPP01") <> "" And "" & rsQD.Fields("CPP02") <> "" And "" & rsQD.Fields("CPP14") <> "" Then
                      strFile = pSavePath & "\" & rsQD.Fields("CPP02")  '下載檔案名稱+路徑
                      '說明書
                      'Modified by Lydia 2018/09/18 ORI.FIX=>改成FIX ; ORI.REP => 改成REP
                      'Modified by Lydia 2018/11/30 判斷最後一道.ORI.%.PDF ,因為.FIX有人加後面
                      'If InStr(UCase("" & rsqd.Fields("CPP02")), ".ORI.") > 0 And InStr(UCase(strA02), ".ORI.PDF") = 0 And _
                                   InStr(UCase(strA02), ".REP") = 0 And InStr(UCase(strA02), ".FIX") = 0 Then
                      If InStr(UCase("" & rsQD.Fields("CPP02")), ".ORI.") > 0 And InStr(UCase(strA02), ".ORI.") = 0 Then
                           If PUB_GetFtpFile("" & rsQD.Fields("CPP14"), strFile, , , , "" & rsQD.Fields("CPP19") <> "") = True Then
                               inX = inX + 1
                               strAttFList = strAttFList & strFile & vbCrLf
                               strA02 = strA02 & rsQD.Fields("CPP02") & ";" 'Added by Lydia 2018/09/19
                           End If
                      End If
                      '序列表
                      'Mark by Lydia 2023/05/08 因應智慧局4/25起對序列表翻譯的變更,無需再抓取檔案, 卷宗區.SEQ.pdf及原始檔區.xml檔案
                      'If InStr(UCase("" & rsqd.Fields("CPP02")), ".SEQ.") > 0 And InStr(UCase(strA02), ".SEQ.") = 0 Then
                      '       If PUB_GetFtpFile("" & rsqd.Fields("CPP14"), strFile) = True Then
                      '           inX = inX + 1
                      '           strAttFList = strAttFList & strFile & vbCrLf
                      '           strA02 = strA02 & rsqd.Fields("CPP02") & ";" 'Added by Lydia 2018/09/19
                      '       End If
                      'End If
                      ''密碼檔
                      'If InStr(UCase("" & rsqd.Fields("CPP02")), ".PWD.") > 0 And InStr(UCase(strA02), ".PWD.") = 0 Then
                      '       If PUB_GetFtpFile("" & rsqd.Fields("CPP14"), strFile) = True Then
                      '          strAttFList = strAttFList & strFile & vbCrLf
                      '          strPWD = "" & rsqd.Fields("CPP02")
                      '          strA02 = strA02 & rsqd.Fields("CPP02") & ";" 'Added by Lydia 2018/09/19
                      '       End If
                      'End If
                      'end 2023/05/08
                      'Added by Lydia 2018/10/02 需提供外翻非說明書部分之其他檔案,例如:技術用語對照表
                      If InStr(UCase("" & rsQD.Fields("CPP02")), ".TBL.") > 0 And InStr(UCase(strA02), ".TBL.") = 0 Then
                             If PUB_GetFtpFile("" & rsQD.Fields("CPP14"), strFile, , , , "" & rsQD.Fields("CPP19") <> "") = True Then
                                 inX = inX + 1
                                 strAttFList = strAttFList & strFile & vbCrLf
                                 strA02 = strA02 & rsQD.Fields("CPP02") & ";"
                             End If
                      End If
                      'end 2018/10/02
                  End If
                  rsQD.MoveNext
             Loop
        End If
        If (InStr(UCase(strAttFList), ".SEP.") = 0 And InStr(UCase(strAttFList), ".ORI.") = 0) Or inX = 0 Then
              MsgBox "卷宗區無說明書！", vbCritical
        End If
    End If
    
    
    'Added by Lydia 2022/11/16 若有序列表原本從卷宗區抓seq.pdf檔案，原始檔區的.XML檔案(因為只有序列表才會用XML檔）檔案一併帶出，且內文新增描述;
    '若卷宗區有seq.pdf檔案 , 原始檔區沒有.XML檔案(因為只有序列表才會用XML檔）檔案, 請彈提醒: 原始檔區無XML檔案 , 請工程師上傳;
    'ex.FCP-67840 在English_Vers有承辦上傳.XML檔案, 如果是工程師上傳.SEQ.XML檔案會在專利案件
    'Mark by Lydia 2023/05/08 因應智慧局4/25起對序列表翻譯的變更,無需再抓取檔案, 卷宗區.SEQ.pdf及原始檔區.xml檔案
    'If InStr(UCase(strAttFList), ".SEQ.") > 0 And strAttFList <> "" Then
    '     strA01 = "SELECT CPF01,CPF02,CPF13 FROM CASEPROGRESS A,CASEPAPERFILE B " & _
                      "WHERE CP01='" & strCase(1) & "' AND CP02='" & strCase(2) & "' AND CP03='" & strCase(3) & "' AND CP04='" & strCase(4) & "' AND CP159=0 " & _
                      "AND CP10 in ('" & cntEnglish_Vers & "', '" & cnt專利案件 & "')  AND CP09=CPF01(+) " & _
                      "AND NVL(CPF10,'N') <> 'D' AND UPPER(CPF02) LIKE '%.XML' " & _
                      "ORDER BY CPF06 DESC, CPF07 DESC "
    '     intq = 1
    '     Set rsqd = ClsLawReadRstMsg(intq, strA01)
    '     If intq = 1 Then
    '          strContAdd = strContAdd & "本案有新格式(ST.26 XML檔)的序列表，請依前指示之翻譯要點進行翻譯," & vbCrLf
    '          rsqd.MoveFirst
    '          Do While Not rsqd.EOF
    '               If "" & rsqd.Fields("CPF01") <> "" And "" & rsqd.Fields("CPF02") <> "" And "" & rsqd.Fields("CPF13") <> "" Then
    '                     strFile = pSavePath & "\" & rsqd.Fields("CPF02")  '下載檔案名稱+路徑
    '                     If PUB_GetFtpFile("" & rsqd.Fields("CPF13"), strFile, "CASEPAPERFILE") = True Then
    '                         inX = inX + 1
    '                         strAttFList = strAttFList & strFile & vbCrLf
    '                     End If
    '               End If
    '               rsqd.MoveNext
    '          Loop
    '     End If
    '     If InStr(UCase(strAttFList), ".XML") = 0 Then
    '             MsgBox "原始檔區無XML檔案，請工程師上傳！", vbCritical
    '     End If
    ' End If
    ''end 2022/11/16
    'end 2023/05/08
    
    'Added by Lydia 2018/10/22 翻譯參考用之word版說明書
    'Modified by Lydia 2024/03/08 排除927其他翻譯+And pType <> "A"
    If strCase(1) <> "" And strCase(2) <> "" And pType <> "A" Then 'Added by Lydia 2018/11/30 排除急件翻譯
        strTmp1(5) = Pub_GetSpecMan("FCP相似比對結果暫存")
        strTmp1(6) = Dir(strTmp1(5) & "\" & strCase(1) & "*" & Val(strCase(2)) & "*.sep.*")
        Do While strTmp1(6) <> ""
             inX = inX + 1
             FileCopy strTmp1(5) & "\" & strTmp1(6), pSavePath & "\" & strTmp1(6)
             strAttFList = strAttFList & pSavePath & "\" & strTmp1(6) & vbCrLf
             SetAttr strTmp1(5) & "\" & strTmp1(6), vbNormal 'Added by Lydia 2020/03/19 預設檔案為正常
             Kill strTmp1(5) & "\" & strTmp1(6)
             strTmp1(6) = Dir()
             'Added by Lydia 2021/10/08 增加Email內文
             If InStr(strContAdd & ",", "提供word 檔僅供參考用") = 0 Then
                 'Modified by Lydia 2022/03/17 為更清楚提醒舜禹提供word檔事，分案翻譯有SEP檔案，Email內容請再加上一句，謝謝。
                 'strContAdd = strContAdd & "提供word 檔僅供參考用,"
                 strContAdd = strContAdd & "提供word 檔僅供參考用,翻譯仍需以pdf檔案為主,請確認翻譯字數及完成日期,謝謝。" & vbCrLf
             End If
             'end 2021/10/08
        Loop
    End If
    'end 2018/10/22
    
    'email內文
    strContent = ""
     '國外譯者
     'Modified by Lydia 2025/03/13 改用模組取得
    'If InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, strTemp(4)) > 0 Then
    If InStr(Pub_SetF51Order("F", ""), strTemp(4)) > 0 Then
            'Added by Lydia 2018/08/08 工程師會把比對結果做成Word檔在命名作業上傳到typing2
            'Modified by Lydia 2022/11/08 有相似度或相似案號就檢查; ex.FCP-67931只有相似案號沒有相似度
            'If strTF19 <> "" And strTF20 <> "" Then
            If strTF19 <> "" Or strTF20 <> "" Then
                  strTmp1(5) = Pub_GetSpecMan("FCP相似比對結果暫存")
                  'Modified by Lydia 2018/09/27 開放可上傳多個檔案(含PDF)
                  'strTmp1(6) = Dir(strTmp1(5) & "\" & strCase(1) & "*" & strCase(2) & ".res.doc*")
                  strTmp1(6) = Dir(strTmp1(5) & "\" & strCase(1) & "*" & Val(strCase(2)) & "*.res.*")
                  If strTmp1(6) = "" Then
                       'Modified by Lydia 2019/02/22 strTmp1(6) = > strResList
                       strResList = "(工程師尚未上傳比對結果)"
                  Else
                       Do While strTmp1(6) <> ""  'Added by Lydia 2018/09/27 開放可上傳多個檔案(含PDF)
                            FileCopy strTmp1(5) & "\" & strTmp1(6), pSavePath & "\" & strTmp1(6)
                            strAttFList = strAttFList & pSavePath & "\" & strTmp1(6) & vbCrLf
                            strResList = strResList & IIf(strResList <> "", ",", "") & strTmp1(6) 'Added by Lydia 2019/03/22 改成記錄檔名
                            'Remove by Lydia 2019/04/23 相似比對檔案(RES)保留至新案翻譯發文(每日批次strMenu91)時刪除(供工程師於核稿時可參考用)
                            'Kill strTmp1(5) & "\" & strTmp1(6) '外專程序對\\Typing2\FCP_WorkFlow有讀寫權限，工程師只有讀取所以走FTP
                       'Added by Lydia 2018/09/27
                            strTmp1(6) = Dir()
                       Loop
                       'end 2018/09/27
                       'Remove by Lydia 2019/03/22 改成記錄檔名
                       'strResList = "案號.RES檔" 'Added by Lydia 2019/02/22
                  End If
            End If
            'end 2018/08/08
            strContent = strTemp(14) & ", 您好:" & vbCrLf & vbCrLf
            'Remove by Lydia 2019/09/27 刪除原本email內文帶入密碼之設定。
            'If strPWD <> "" Then
            '   strContent = "pw: (請參考附件: " & strPWD & ") " & "," & vbCrLf & vbCrLf & vbCrLf & strContent
            'End If
            'end 2019/09/27
            If pType = "0" Then '急件翻譯(尚未收文)
                strSubject = "急件翻譯(" & strTemp(0) & ")" & ","
                strContent = strContent & "附檔說明書需翻成" & strTemp(3) & "," & vbCrLf
                'Added by Lydia 2019/08/21 加註"指定代表圖"
                If m_TCT97 <> "" Then
                    strContent = strContent & "指定代表圖：圖" & m_TCT97 & vbCrLf
                End If
                'Added by Lydia 2019/08/23 加註:翻譯特殊指示
                If m_TF36 <> "" Then
                    strContent = strContent & "翻譯特殊指示：" & m_TF36 & vbCrLf
                End If
                'Added by Lydia 2024/06/19
                If strTemp(2) = "德文" Then
                   strContent = strContent & "本案說明書為德文，本所提供的中文名稱僅作為參考，翻譯時請不要受限於本所提供的譯名，若有更適當的中文翻譯，請通知本所修改。" & vbCrLf
                End If
                'end 2024/06/19
                If strTemp(18) = "" Then 'Added by Lydia 2018/12/05 只有交稿期限用舊句子
                    strContent = strContent & "可否於" & ChangeWStringToWDateString(strTemp(13)) & "前完成,若無法" & ChangeWStringToWDateString(strTemp(13)) & "前交稿,也請告知最快能交稿之日期,謝謝！" & vbCrLf & vbCrLf
                Else
                    '可否於2018/xx/xx前完成交稿claims,於2018/xx/xx前完成交稿全文?
                    strContent = strContent & "可否於" & ChangeWStringToWDateString(strTemp(18)) & "完成交稿Claims,"
                    strContent = strContent & "於" & ChangeWStringToWDateString(strTemp(13)) & "前完成交稿全文?" & vbCrLf & vbCrLf
                    strContent = strContent & "若無法,也請告知最快能交稿之日期,謝謝！" & vbCrLf & vbCrLf
                End If
                'end 2018/12/05
            ElseIf pType = "1" Then '急件翻譯(已立案)
                'Modified by Lydia 2018/12/22 主旨改急件翻譯號
                'strSubject = "急件翻譯(" & strTemp(0) & ")" & ","
                'Modified by Lydia 2024/07/31 +Our Ref:
                strSubject = "Our Ref:" & strTemp(0) & " 急件翻譯(" & strTCN01 & ")" & ","
                'Added by Lydia 2019/08/21 加註"指定代表圖"
                If m_TCT97 <> "" Then
                    strContent = strContent & "指定代表圖：圖" & m_TCT97 & vbCrLf
                End If
                'Added by Lydia 2019/08/23 加註:翻譯特殊指示
                If m_TF36 <> "" Then
                    strContent = strContent & "翻譯特殊指示：" & m_TF36 & vbCrLf
                End If
                'Added by Lydia 2024/06/19
                If strTemp(2) = "德文" Then
                   strContent = strContent & "本案說明書為德文，本所提供的中文名稱僅作為參考，翻譯時請不要受限於本所提供的譯名，若有更適當的中文翻譯，請通知本所修改。" & vbCrLf
                End If
                'end 2024/06/19
                'Added by Lydia 2024/07/31 有相似度或相似案號就顯示; ex.FCP-67931只有相似案號沒有相似度
                If strTF19 <> "" Or strTF20 <> "" Then '有相似度和相似案號
                    Call ChgCaseNo(strTF20, strTmp1)
                    strContent = strContent & "<B><U>此件與前件" & strTmp1(1) & Val(strTmp1(2)) & IIf(strTmp1(3) & strTmp1(4) <> "000", strTmp1(3) & strTmp1(4), "") & "有" & Val(strTF19) & "%相似度" & ",</B></U>"
                    strContent = strContent & "附上相似比對檔案: " & strResList & "," & vbCrLf
                End If
                'end 2024/07/31
                If strTemp(18) <> "" Then strContent = strContent & "請於" & ChangeWStringToWDateString(strTemp(18)) & "先行交稿Claims, " & vbCrLf
                'Modified by Lydia 2022/03/17 +增加Email內文 + strContAdd
                strContent = strContent & strContAdd & "附上台一工作通知單,請於" & ChangeWStringToWDateString(strTemp(13)) & "前完成交稿,謝謝！" & vbCrLf & vbCrLf
            'Added by Lydia 2024/03/08 因應內專支援機械組OA (含P案) 收文: 927其他翻譯
            ElseIf pType = "A" Then
                strSubject = "Our Ref:" & strTemp(0) & " OA翻譯"
                'Added by Lydia 2024/03/12 P案內文有變動
                'Modified by Lydia 2024/03/29 增加核駁
                'If strCase(1) = "P" Then
                strTmp1(9) = Pub_GetNoToCPM("1", strTCN01, strTmp1(10))
                If strTmp1(10) = "1002" Then
                   strContent = strContent & strTemp(0) & "案件核駁審定書，由" & strTemp(3) & "翻" & strTemp(2) & "，僅翻譯第2頁第八理由部分(理由以上全不用翻譯)" '，一併提供本案英文說明書供參考，請根據英文說明書用語翻譯使翻譯一致，請確認完成日期,謝謝。" & vbCrLf & vbCrLf
                ElseIf strCase(1) = "P" Then
                'end 2024/03/29
                   strContent = strContent & strTemp(0) & "案件" & strTmp1(9) & "，由" & strTemp(3) & "翻" & strTemp(2) & "，包含意見通知書(第3頁開始到最後)及檢索報告" '，一併提供本案英文說明書供參考，請根據英文說明書用語翻譯使翻譯一致，請確認完成日期,謝謝。" & vbCrLf & vbCrLf
                Else
                'end 2024/03/12
                   'FCP0xxxxx案件審查意見通知書，由繁中翻英文，包含說明部分及兩個附件(引證參考資料及檢索報告) ,一併提供本案英文說明書供參考，請根據英文說明書用語翻譯使翻譯一致，請確認完成日期,謝謝。
                   'Modified by Lydia 2024/03/15 FCP案加註: (說明第二、三、四為固定內容，不用翻譯此部分)
                   'Modified by Lydia 2024/03/29 改內容
                   'strContent = strContent & strTemp(0) & "案件" & strTmp1(9) & "，由" & strTemp(3) & "翻" & strTemp(2) & "，包含說明部分(說明第二、三、四為固定內容，不用翻譯此部分)及兩個附件(引證參考資料及檢索報告)"
                   strContent = strContent & strTemp(0) & "案件" & strTmp1(9) & "，由" & strTemp(3) & "翻" & strTemp(2) & "，僅需翻譯說明部分(說明以上部分及說明第二、三、四為固定內容，不用翻譯)及兩個附件(引證參考資料及檢索報告)"
                End If
                'Added by Lydia 2024/06/19
                If strTemp(2) = "德文" Then
                   strContent = strContent & "本案說明書為德文，本所提供的中文名稱僅作為參考，翻譯時請不要受限於本所提供的譯名，若有更適當的中文翻譯，請通知本所修改。" & vbCrLf
                End If
                'end 2024/06/19
                strContent = strContent & "，一併提供本案英文說明書供參考，請根據英文說明書用語翻譯使翻譯一致，請確認完成日期,謝謝。" & vbCrLf & vbCrLf
            'end 2024/03/08
            Else
                'Added by Lydia 2022/07/28 因應外專郵件沖銷稽核程式
                If strSrvDate(1) >= 外專信件沖銷啟用日 Then
                     strSubject = "Our Ref:" & strTemp(0)
                Else
                'end 2022/07/28
                    strSubject = strTemp(0)
                End If 'Added by Lydia 2022/07/28
                
                strContent = strContent & "附上" & strTemp(0) & "案件電子檔共" & inX & "個及工作通知單一份," & vbCrLf
                'Added by Lydia 2019/08/21 加註"指定代表圖"
                If m_TCT97 <> "" Then
                    strContent = strContent & "指定代表圖：圖" & m_TCT97 & vbCrLf
                End If
                'Added by Lydia 2019/08/23 加註:翻譯特殊指示
                If m_TF36 <> "" Then
                    strContent = strContent & "翻譯特殊指示：" & m_TF36 & vbCrLf
                End If
                'Added by Lydia 2024/06/19
                If strTemp(2) = "德文" Then
                   strContent = strContent & "本案說明書為德文，本所提供的中文名稱僅作為參考，翻譯時請不要受限於本所提供的譯名，若有更適當的中文翻譯，請通知本所修改。" & vbCrLf
                End If
                'end 2024/06/19
                'Modified by Lydia 2022/11/08 有相似度或相似案號就顯示; ex.FCP-67931只有相似案號沒有相似度
                'If strTF19 <> "" And strTF20 <> "" Then '有相似度和相似案號
                If strTF19 <> "" Or strTF20 <> "" Then '有相似度和相似案號
                    Call ChgCaseNo(strTF20, strTmp1)
                    'Modified by Lydia 2018/08/15 字體加粗和底線
                    'strContent = strContent & "此件與前件" & strTmp1(1) & Val(strTmp1(2)) & IIf(strTmp1(3) & strTmp1(4) <> "000", strTmp1(3) & strTmp1(4), "") & "有" & Val(strTF19) & "%相似度" & "," & vbCrLf
                    'strContent = strContent & "附上相似比對檔案: " & strTmp1(6) & "," & vbCrLf
                    'Modified by Lydia 2019/02/22 修改語句
                    'strContent = strContent & "<B><U>此件與前件" & strTmp1(1) & Val(strTmp1(2)) & IIf(strTmp1(3) & strTmp1(4) <> "000", strTmp1(3) & strTmp1(4), "") & "有" & Val(strTF19) & "%相似度" & ",</B></U>" & vbCrLf
                    'strContent = strContent & "附上相似比對檔案: " & strTmp1(6) & "," & vbCrLf
                    'end 2018/08/15
                    strContent = strContent & "<B><U>此件與前件" & strTmp1(1) & Val(strTmp1(2)) & IIf(strTmp1(3) & strTmp1(4) <> "000", strTmp1(3) & strTmp1(4), "") & "有" & Val(strTF19) & "%相似度" & ",</B></U>"
                    strContent = strContent & "附上相似比對檔案: " & strResList & "," & vbCrLf
                End If
                If strTemp(18) <> "" Then strContent = strContent & "請於" & ChangeWStringToWDateString(strTemp(18)) & "先行交稿Claims, " & vbCrLf
                'Modified by Lydia 2021/10/08 增加Email內文 + strContAdd
                'Modified by Lydia 2022/03/24 判斷內文
                'strContent = strContent & strContAdd & "請確認翻譯字數及完成日期,謝謝。" & vbCrLf & vbCrLf
                If InStr(strContent & strContAdd, "請確認翻譯字數及完成日期,謝謝。") = 0 Then
                     strContent = strContent & strContAdd & "請確認翻譯字數及完成日期,謝謝。" & vbCrLf & vbCrLf
                Else
                    strContent = strContent & strContAdd & vbCrLf
                End If
                'end 2022/03/24
            End If
            'Modified by Lydia 2021/09/09 改變尾端署名
            'strContent = strContent & "台一" & strConUser
            'strContent = strContent & "台一國際智慧財產事務所" & vbCrLf 'Mark by Lydia 2022/02/18 改到下方
            If strUserNum <> strConUserNo And Pub_StrUserSt03 <> "M51" Then  '代理人員
                strTmp1(2) = Pub_GetStaffExtn(strUserNum, strTmp1(3)) & "  代發"
                strContent = strContent & "專利國外部  " & strUserName & " #" & strTmp1(2)
            Else
                strTmp1(2) = Pub_GetStaffExtn(strConUserNo, strTmp1(3))
                strContent = strContent & "專利國外部  " & strConUser & " #" & strTmp1(2)
            End If
            'end 2021/09/09
            strContent = strContent & vbCrLf & "台一國際智慧財產事務所"    'Added by Lydia 2022/02/18
            'Added by Lydia 2019/09/27 編號為國外翻譯社(F5588舜禹、F5698迅達、F5714捷恩凱)，若有放假通知(中文)，自動帶入Outlook。
            strTmp1(0) = Pub_GetLetterMemo("P", "1")
            If strTmp1(0) <> "" Then
               strContent = strContent & vbCrLf & vbCrLf & "<font color=""red"">* " & strTmp1(0) & "</font>" & vbCrLf
            End If
    '所內員工
    ElseIf strTemp(17) <> "F51" Then
            '員工編號+案件狀態,判斷上班/下班翻譯
            strA01 = "上班"
            If Left(strTemp(4), 1) = "F" Then
               strA01 = "下班" 'Added by Lydia 2018/07/09 在mail內文列出只能上班翻譯的原因，但是主旨依舊用下班翻譯
               If InStr(strTemp(1), "下班") = 0 Then
                  'Modified by Lydia 2019/06/19 與認翻譯作業統一
                    'strContent = "P.S. 本案只能" & strTemp(1) & "翻譯: " & Mid(strTemp(21), 1, Len(strTemp(21)) - 1) & vbCrLf & vbCrLf
                    strContent = "P.S. 本案不能下班翻譯: " & Mid(strTemp(21), 1, Len(strTemp(21)) - 1) & vbCrLf & vbCrLf
               'Else 'Mark by Lydia 2018/07/09
               '     strA01 = "下班"
               End If
            End If
            
            strSubject = strTemp(0) & "請進行" & strA01 & "翻譯"
            strContent = strContent & strTemp(5) & ", 您好:" & vbCrLf & vbCrLf
            
            'Added by Lydia 2021/04/14 外專翻譯承辦及核稿期限控管：查詢該認領人員，新案翻譯未上完稿日案件彈訊息並加入Email內文
            strText = Pub_GetEngEP09List(pCP14, strCase(1) & strCase(2) & strCase(3) & strCase(4))
            If strText <> "" Then
               strContent = strContent & "尚未完稿案件：" & strText & vbCrLf & vbCrLf
            End If
            'end 2021/04/14
            
            If pType = "3" And strTemp(15) = "" Then  '未提申先翻譯
                strContent = strContent & strTemp(0) & "," & vbCrLf
            Else
                strContent = strContent & strTemp(0) & "已於" & ChangeTStringToTDateString(TransDate(strTemp(15), 1)) & "提出申請," & vbCrLf & vbCrLf
            End If
            strContent = strContent & "案件名稱：" & strTemp(10) & "," & vbCrLf
            'Added by Lydia 2019/08/23 加註:翻譯特殊指示
            If m_TF36 <> "" Then
                strContent = strContent & "翻譯特殊指示：" & m_TF36 & vbCrLf
            End If
            'Added by Lydia 2024/06/19
            If strTemp(2) = "德文" Then
               strContent = strContent & "本案說明書為德文，本所提供的中文名稱僅作為參考，翻譯時請不要受限於本所提供的譯名，若有更適當的中文翻譯，請通知本所修改。" & vbCrLf
            End If
            'end 2024/06/19
            'Added by Lydia 2018/10/05 加註相似度和相似案號
            'Modified by Lydia 2022/11/08 有相似度或相似案號就檢查; ex.FCP-67931只有相似案號沒有相似度
            'If strTF19 <> "" And strTF20 <> "" Then '有相似度和相似案號
            If strTF19 <> "" Or strTF20 <> "" Then
                 Call ChgCaseNo(strTF20, strTmp1)
                 strContent = strContent & "<B><U>此件與前件" & strTmp1(1) & Val(strTmp1(2)) & IIf(strTmp1(3) & strTmp1(4) <> "000", strTmp1(3) & strTmp1(4), "") & "有" & Val(strTF19) & "%相似度" & ",</B></U>" & vbCrLf
            End If
            'end 2018/10/05
            If strTemp(18) <> "" Then strContent = strContent & "交稿Claims期限為：" & ChangeTStringToTDateString(TransDate(strTemp(18), 1)) & "," & vbCrLf
            strContent = strContent & "翻譯交稿期限為：" & ChangeTStringToTDateString(TransDate(strTemp(13), 1)) & "," & vbCrLf
            strContent = strContent & "請進行" & strA01 & "翻譯, 謝謝!" & vbCrLf & vbCrLf
            'Modified by Lydia 2021/09/09 判斷人員的英文別名
            'strContent = strContent & "Sharon"
            If strUserNum <> strConUserNo And Pub_StrUserSt03 <> "M51" Then '代理人員
                strTmp1(2) = Pub_GetStaffExtn(strUserNum, strTmp1(3))
                If strTmp1(3) <> "" Then strTmp1(3) = strTmp1(3) & "  代發"
            Else
                strTmp1(2) = Pub_GetStaffExtn(strConUserNo, strTmp1(3))
            End If
            strContent = strContent & IIf(strTmp1(3) <> "", strTmp1(3), strConUser)
            'end 2021/09/09
    '所外員工
    Else
            'Added by Lydia 2018/08/08 附上比對結果做成Word檔在命名作業上傳到typing2
            'Modified by Lydia 2022/11/08 有相似度或相似案號就檢查; ex.FCP-67931只有相似案號沒有相似度
            'If strTF19 <> "" And strTF20 <> "" Then
            If strTF19 <> "" Or strTF20 <> "" Then
                  strTmp1(5) = Pub_GetSpecMan("FCP相似比對結果暫存")
                  strTmp1(6) = Dir(strTmp1(5) & "\" & strCase(1) & "*" & Val(strCase(2)) & "*.res.*")
                  If strTmp1(6) = "" Then
                       'Modified by Lydia 2019/02/22 strTmp1(6)=>strResList
                       strResList = "(工程師尚未上傳比對結果)"
                  Else
                       Do While strTmp1(6) <> ""
                            FileCopy strTmp1(5) & "\" & strTmp1(6), pSavePath & "\" & strTmp1(6)
                            strAttFList = strAttFList & pSavePath & "\" & strTmp1(6) & vbCrLf
                            strResList = strResList & IIf(strResList <> "", ",", "") & strTmp1(6)   'Added by Lydia 2019/03/22 改成記錄檔名
                            'Remove by Lydia 2019/04/23 相似比對檔案(RES)保留至新案翻譯發文(每日批次strMenu91)時刪除(供工程師於核稿時可參考用)
                            'Kill strTmp1(5) & "\" & strTmp1(6)
                            strTmp1(6) = Dir()
                       Loop
                       'Remove by Lydia 2019/03/22 改成記錄檔名
                       'strResList = "案號.RES檔" 'Added by Lydia 2019/02/22
                  End If
            End If
            'end 2018/08/08
            strSubject = strTemp(0) & "翻譯"
            strContent = strTemp(5) & ", 您好:" & vbCrLf & vbCrLf
            strContent = strContent & "案件名稱：" & strTemp(10) & "," & vbCrLf
            'Added by Lydia 2019/08/21 加註"原文字數和指定代表圖"
            strContent = strContent & "原文字數：" & m_TF23 & vbCrLf
            If m_TCT97 <> "" Then
                strContent = strContent & "指定代表圖：圖" & m_TCT97 & vbCrLf
            End If
            'Added by Lydia 2019/08/23 加註:翻譯特殊指示
            If m_TF36 <> "" Then
                strContent = strContent & "翻譯特殊指示：" & m_TF36 & vbCrLf
            End If
            'Added by Lydia 2024/06/19
            If strTemp(2) = "德文" Then
               strContent = strContent & "本案說明書為德文，本所提供的中文名稱僅作為參考，翻譯時請不要受限於本所提供的譯名，若有更適當的中文翻譯，請通知本所修改。" & vbCrLf
            End If
            'end 2024/06/19
            'Added by Lydia 2018/10/05 加註相似度和相似案號
            'Modified by Lydia 2022/11/08 有相似度或相似案號就檢查; ex.FCP-67931只有相似案號沒有相似度
            'If strTF19 <> "" And strTF20 <> "" Then '有相似度和相似案號
            If strTF19 <> "" Or strTF20 <> "" Then
                 Call ChgCaseNo(strTF20, strTmp1)
                 'Modified by Lydia 2019/02/22 修改語句
                 'strContent = strContent & "<B><U>此件與前件" & strTmp1(1) & Val(strTmp1(2)) & IIf(strTmp1(3) & strTmp1(4) <> "000", strTmp1(3) & strTmp1(4), "") & "有" & Val(strTF19) & "%相似度" & ",</B></U>" & vbCrLf
                 'strContent = strContent & "附上相似比對檔案: " & strTmp1(6) & "," & vbCrLf
                 strContent = strContent & "<B><U>此件與前件" & strTmp1(1) & Val(strTmp1(2)) & IIf(strTmp1(3) & strTmp1(4) <> "000", strTmp1(3) & strTmp1(4), "") & "有" & Val(strTF19) & "%相似度" & ",</B></U>"
                 strContent = strContent & "附上相似比對檔案: " & strResList & "," & vbCrLf
                 'end 2019/02/22
            End If
            'end 2018/10/05
            If strTemp(18) <> "" Then strContent = strContent & "交稿Claims期限為：" & ChangeTStringToTDateString(TransDate(strTemp(18), 1)) & "," & vbCrLf
            strContent = strContent & "翻譯交稿期限為：" & ChangeTStringToTDateString(TransDate(strTemp(13), 1)) & "," & vbCrLf
            'Modified by Lydia 2018/08/22
            'strContent = strContent & "附上檔案,紙本今以掛號寄出,煩請留意,謝謝!" & vbCrLf & vbCrLf
            strContent = strContent & "附上檔案,請勿將檔案外流且於期限內完成交稿,謝謝!" & vbCrLf & vbCrLf
            'Modified by Lydia 2021/09/09 判斷人員的英文別名
            'strContent = strContent & "Sharon"
            If strUserNum <> strConUserNo And Pub_StrUserSt03 <> "M51" Then '代理人員
                strTmp1(2) = Pub_GetStaffExtn(strUserNum, strTmp1(3))
                If strTmp1(3) <> "" Then strTmp1(3) = strTmp1(3) & "  代發"
            Else
                strTmp1(2) = Pub_GetStaffExtn(strConUserNo, strTmp1(3))
            End If
            strContent = strContent & IIf(strTmp1(3) <> "", strTmp1(3), strConUser)
            'end 2021/09/09
    End If
    
    'Added by Lydia 2018/07/09 移到最後面插入by Sharon
    If m_TempFileName <> "" Then
        strAttFList = strAttFList & pSavePath & "\" & m_TempFileName & vbCrLf
    End If
    
    '轉HTML格式
    strContent = Replace(strContent, "新細明體", "Times New Roman")
    strContent = Replace(strContent, vbCrLf, "<BR>")
    strContent = Replace(strContent, "  ", "&nbsp;&nbsp;")
    '呼叫新郵件：
    Set objOutLook = CreateObject("Outlook.Application")
    'Added by Lydia 2019/08/06 對外信件要加信尾(郵件範本)
    'Modified by Lydia 2025/03/13 改用模組取得
    'If InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, strTemp(4)) > 0 Or strTemp(17) = "F51" Then
    If InStr(Pub_SetF51Order("F", ""), strTemp(4)) > 0 Or strTemp(17) = "F51" Then
        'Modified by Lydia 2023/08/18
        'If Dir(App.path & "\$$TOT-000F22-0-01.oft") = "" Then
        '    Call PUB_GetSampleFile("$$TOT-000F22-0-01.oft", "TOT-000F22-0-01")
        'End If
        'Set objMail = objOutLook.CreateItemFromTemplate(App.path & "\$$TOT-000F22-0-01.oft")
        If Dir(pSavePath & "\$$TOT-000F22-0-01.oft") = "" Then
            Call PUB_GetSampleFile("$$TOT-000F22-0-01.oft", "TOT-000F22-0-01", , pSavePath)
        End If
        Set objMail = objOutLook.CreateItemFromTemplate(pSavePath & "\$$TOT-000F22-0-01.oft")
        'end 2023/08/18
    Else
    'end 2019/08/06
        Set objMail = objOutLook.CreateItem(0)
    End If
    
    strTemp(1) = "" 'Added by Lydia 2023/09/06
    objMail.Subject = strSubject
    objMail.To = strTemp(9)
    'Modified by Lydia 2019/08/06 改成細明體
    'objMail.HTMLBody = "<FONT FACE=""Times New Roman"">" & strContent & "<BR>" & Replace(objMail.HTMLBody, "&lt;LetterMemo&gt;", IIf(ExceptFieldData("公用備註/英") <> "", "<BR>Message:<BR>" & Replace(ChgHTMLFormat(ExceptFieldData("公用備註/英")), vbCrLf, "<BR>") & "<BR><BR>", "&nbsp;")) & "</FONT>"
    If strAttFList <> "" Then
       Set fs = CreateObject("Scripting.FileSystemObject") 'Added by Lydia 2023/09/06
       tmpArr = Split(strAttFList, vbCrLf)
       For intQ = 0 To UBound(tmpArr)
          If Trim(tmpArr(intQ)) <> "" Then
             'Added by Lydia 2023/09/06 目前對外寄信限制25M,判斷檔案大小不可超過; EX.FCP-70173
             Set f = fs.GetFile(tmpArr(intQ))
             If f.Size >= 26214400 Then
                 strTemp(1) = strTemp(1) & vbCrLf & GetFileName(tmpArr(intQ))
             Else
             'end 2023/09/06
                 objMail.Attachments.add (tmpArr(intQ))
              End If 'Added by Lydia 2023/09/06
          End If
       Next
       'Added by Lydia
    End If

    'Added by Lydia 2023/09/06 插入內文
    If strTemp(1) <> "" Then
       strContent = "以下檔案超過郵件附件最大25MB的限制，請自行將檔案縮小再插入Outlook郵件：" & vbCrLf & strTemp(1) & vbCrLf & String(20, "=") & vbCrLf & strContent
       strContent = Replace(strContent, vbCrLf, "<BR>")
    End If
    
    'Move by Lydia 2023/09/06 從上面移下來
    objMail.HTMLBody = "<FONT FACE=""細明體"">" & strContent & "<BR>" & Replace(objMail.HTMLBody, "&lt;LetterMemo&gt;", IIf(ExceptFieldData("公用備註/英") <> "", "<BR>Message:<BR>" & Replace(ChgHTMLFormat(ExceptFieldData("公用備註/英")), vbCrLf, "<BR>") & "<BR><BR>", "&nbsp;")) & "</FONT>"
    
    objMail.Display
    
    Set rsAD = Nothing
   
ErrHand:
    If Err.Number = 462 Then '遠端伺服器不存在或無法使用
       GoTo RestarWord
    ElseIf Err.Number <> 0 Then
         If strContent <> "" Then
            MsgBox "開啟撰寫郵件視窗失敗，請人工作業！"
         Else
            MsgBox (Err.Description)
         End If
         'Added by Lydia 2019/06/26 Sharon要求Outlook作業失敗,發Email通知避免不記得案號
         PUB_SendMail strUserNum, strUserNum, "", IIf(strCase(2) = "", "急件翻譯:", "") & strTemp(0) & "開啟撰寫郵件視窗失敗，請人工作業！", "同摘要"
    End If

   Set objMail = Nothing
   Set objOutLook = Nothing
   'Added by Lydia 2024/03/29
   Set rsAD = Nothing
   Set rsQD = Nothing
End Sub

'Added by Lydia 2018/06/20 FCP案發文時，電子送件自動上傳檔案到卷宗區
'不限新案，依據輸入的智慧局收文號(受理號,ex: 1073066637-0)，將本機C:\E-SET\RdcDocDir\(收文號ex: 1073066637-0)的pdf檔自動搬移到卷宗區(by Phoebe);
'Modified by Lydia 2019/03/22 +傳入發文日
Public Function Pub_AutoEsetToCpp(ByVal bolMsg As Boolean, ByVal sPA01 As String, ByVal sPA02 As String, ByVal sPA03 As String, ByVal sPA04 As String, ByVal sPA08 As String, ByVal sCP09 As String, ByVal sCP10 As String, ByVal sFilePath As String, Optional ByVal pCP27 As String) As Boolean
'sPA01~sPA04 本所案號; sPA08 專利種類
'sCP09 收文號;  sCP10 案件性質
'bolMsg 彈訊息是否上傳檔案
Dim strMid As String
Dim strNewName As String
Dim strFilePath As String
'Added by Lydia 2018/08/29
Dim oFileSys As New FileSystemObject
Dim oFile
 'Added by Lydia 2018/12/21
Dim strTitle As String
Dim strCont As String, bCheck As Boolean
'Added by Lydia 2019/03/22
Dim rsChk As New ADODB.Recordset
Dim intA As Integer
Dim strTemp As String

On Error GoTo ErrorHand01 'Added by Lydia 2018/08/29

    Pub_AutoEsetToCpp = True
    If sFilePath = "" Then Exit Function
    strFilePath = "C:\E-SET\RdcDocDir\" & sFilePath
    strTitle = "發文電子送件自動上傳檔案到卷宗區" 'Added by Lydia 2018/12/21

    strMid = Dir(strFilePath & "\*.pdf")
    
    If strMid = "" Then 'Memo by Lydia 2019/10/28 因為一併發文會先發文的先上傳，所以資料夾=無
        If bolMsg = True Then
            'Added by Lydia 2019/03/22 先去判斷是否有同天發文的案件性質,若無其他發文則回到發文key智慧局收文號畫面
            '若有則繼續下一步詢問
            'Modified by Lydia 2024/12/26 +電子送件和同一發文字號; ex.FCP-065116在12/24先輸入"自請撤回-發明申請"
            'strTemp = "select cp09,cp10 from caseprogress where cp01='" & sPA01 & "' and cp02='" & sPA02 & "' and cp03='" & sPA03 & "' and cp04='" & sPA04 & "' and cp158=" & IIf(pCP27 <> "", TransDate(pCP27, 2), strSrvDate(1))
            strTemp = "select cp09,cp10 from caseprogress where cp01='" & sPA01 & "' and cp02='" & sPA02 & "' and cp03='" & sPA03 & "' and cp04='" & sPA04 & "'" & _
                      " and cp158=" & IIf(pCP27 <> "", TransDate(pCP27, 2), strSrvDate(1)) & " and nvl(cp118,'N') <> 'N' and instr(cp64,'" & ChgSQL(sFilePath) & "') > 0 "
            'Added by Lydia 2019/10/28 排除無收文號(發文->直接延期,先檢查檔案)
            If sCP09 <> "" Then strTemp = strTemp & " and cp09 <> " & CNULL(sCP09)
            
            intA = 1
            Set rsChk = ClsLawReadRstMsg(intA, strTemp)
            If intA = 0 Then
                If Pub_StrUserSt03 <> "M51" Then
                    MsgBox "無此收文號，請重新輸入！", vbCritical
                    Pub_AutoEsetToCpp = False
                    Exit Function
                Else
                    If MsgBox("是否跳過同天發文的判斷？", vbInformation + vbYesNo + vbDefaultButton1, "電腦中心") = vbNo Then
                       Pub_AutoEsetToCpp = False
                       Exit Function
                    End If
                End If
            End If
            Set rsChk = Nothing
            'end 2019/03/22
            
            '因為實審和主動修正可併在一起，所以E-Set資料夾查無檔案則彈訊息問是否上傳檔案(by Phoebe)
            'Modified by Lydia 2018/12/21 +Title
            If MsgBox("電子送件是否要上傳檔案到卷宗區？" & vbCrLf & "(Yes=要上傳；No=不要上傳，繼續發文)", vbYesNo + vbDefaultButton1 + vbInformation, strTitle) = vbYes Then
                Pub_AutoEsetToCpp = False
                Exit Function
            End If
        End If
    'Added by Lydia 2019/12/06
    ElseIf sPA01 <> "" And sPA02 <> "" And sCP09 = "" Then  '發文直接做延期，無收文號先檢查.contact
        '檢查基本資料表(.contact.pdf)的案號, 是否正確;用於稽核智慧局發文字號和本所案號是否正確
        If bCheck = False Then
            strCont = Dir(strFilePath & "\*.contact.pdf")
            If strCont = "" Then
                MsgBox "智慧局發文字號資料夾:" & strFilePath & vbCrLf & "無Contact.pdf ，請確認 !", vbCritical, strTitle
                Pub_AutoEsetToCpp = False
                Exit Function
            Else
                If UCase(strCont) <> UCase(sPA01 & sPA02 & ".CONTACT.PDF") And UCase(strCont) <> UCase(sPA01 & Val(sPA02) & ".CONTACT.PDF") Then
                    MsgBox "智慧局收文號和FCP案號不一致 !", vbCritical, strTitle
                    Pub_AutoEsetToCpp = False
                    Exit Function
                End If
            End If
            bCheck = True
        End If
    'end 2019/12/06
    ElseIf sPA01 <> "" And sPA02 <> "" And sPA08 <> "" And sCP09 <> "" And sCP10 <> "" Then
        Do While strMid <> ""
            'Added by Lydia 2018/12/21 檢查基本資料表(.contact.pdf)的案號, 是否正確;用於稽核智慧局發文字號和本所案號是否正確
            If bCheck = False Then
                strCont = Dir(strFilePath & "\*.contact.pdf")
                If strCont = "" Then
                    MsgBox "智慧局發文字號資料夾:" & strFilePath & vbCrLf & "無Contact.pdf ，請確認 !", vbCritical, strTitle
                    Pub_AutoEsetToCpp = False
                    Exit Function
                Else
                    If UCase(strCont) <> UCase(sPA01 & sPA02 & ".CONTACT.PDF") And UCase(strCont) <> UCase(sPA01 & Val(sPA02) & ".CONTACT.PDF") Then
                        MsgBox "智慧局收文號和FCP案號不一致 !", vbCritical, strTitle
                        Pub_AutoEsetToCpp = False
                        Exit Function
                    End If
                End If
                bCheck = True
            End If
            'end 2018/12/21
            
            Pub_AutoEsetToCpp = False 'Added by Lydia 2019/11/27
            strNewName = PUB_CaseNo2FileName(sPA01, sPA02, sPA03, sPA04)
            Set oFile = oFileSys.GetFile(strFilePath & "\" & strMid) 'Added by Lydia 2018/08/29
           '檢查檔案是否正在使用中
            If PUB_ChkFileOpening(strFilePath & "\" & strMid) = True Then
                'Modified by Lydia 2018/12/21 +Title
                MsgBox strFilePath & "\" & strMid & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation, strTitle
                Exit Function
            End If
            'Added by Lydia 2019/02/12 檢查檔案大小為 0 KB 有誤
            If oFile.Size = 0 Then
                  ShowMsg strFilePath & "\" & strMid & vbCrLf & MsgText(9221)
                  Exit Function
            End If
            'end 2019/02/12
            
            '電子送件自動匯入卷宗區請自動排除檔名中有.FIX_U(劃線本)和.COR_U的檔案by Phoebe
            If InStr(UCase(strMid), ".FIX_U") > 0 Or InStr(UCase(strMid), ".COR_U") > 0 Then
               SetAttr strFilePath & "\" & strMid, vbNormal 'Added by Lydia 2020/03/19 預設檔案為正常
               Kill strFilePath & "\" & strMid
            Else
               If InStr(strMid, "中文本") > 0 Then
                     If sPA08 = "1" Then
                          strNewName = strNewName & "." & sCP10 & ".inv.pdf"
                     ElseIf sPA08 = "2" Then
                          strNewName = strNewName & "." & sCP10 & ".utl.pdf"
                     ElseIf sPA08 = "3" Then
                          strNewName = strNewName & "." & sCP10 & ".des.pdf"
                     Else
                          'Modified by Lydia 2018/12/21 +Title
                          MsgBox "下列檔案請手動上傳卷宗區後，再做發文!" & vbCrLf & strFilePath & "\" & strMid
                          Exit Function
                     End If
               'Modified by Lydia 2024/11/07 更正說明書---無劃線全份與(修正說明書---無劃線全份)相同--- Winfrey
               ElseIf InStr(strMid, "修正說明書") > 0 Or InStr(strMid, "更正說明書") > 0 Then
                     'Modified by Lydia 2018/08/08 敏莉要求與英文版區別，改成只要+fix
                     'strNewName = strNewName & "." & sCP10 & ".ori.fix.pdf"
                     strNewName = strNewName & "." & sCP10 & ".fix.pdf"
               ElseIf InStr(strMid, "申請書") > 0 Or InStr(strMid, "專簡") > 0 Then
                     strNewName = strNewName & "." & sCP10 & ".data.pdf"
               'Added by Lydia 2022/10/14 若有.SEQ_XXX.pdf(智慧局下載) 請先更名為.SEQ.pdf 再匯入卷宗區 --- Phoebe
               ElseIf InStr(UCase(strMid), ".SEQ_") > 0 Or InStr(UCase(strMid), "SEQ_") > 0 Then
                     strNewName = strNewName & "." & sCP10 & ".SEQ." & Mid(strMid, InStrRev(strMid, ".") + 1)
               'end 2022/10/14
               'Added by Lydia 2022/11/23 若檔案名稱有”訂正說明書“，請更名為.COR再匯入卷宗區，使其副檔名說明為"訂正本“ --- Phoebe
               ElseIf InStr(UCase(strMid), "訂正說明書") > 0 Then
                     strNewName = strNewName & "." & sCP10 & ".COR." & Mid(strMid, InStrRev(strMid, ".") + 1)
               'end 2022/11/23
               '其他
               Else
                     If Left(UCase(strMid), Len(strNewName)) = strNewName Then
                           strNewName = strNewName & "." & sCP10 & "." & PUB_GetSimpleName(Mid(strMid, Len(strNewName) + 1))
                     'FCPXXXXX(6碼) 開頭
                     ElseIf Left(UCase(strMid), Len(sPA01 & sPA02)) = sPA01 & sPA02 Then
                           strNewName = strNewName & "." & sCP10 & "." & PUB_GetSimpleName(Mid(strMid, Len(sPA01 & sPA02) + 1))
                     '+案號
                     Else
                           strNewName = strNewName & "." & sCP10 & "." & PUB_GetSimpleName(strMid)
                     End If
               End If
               strNewName = Replace(strNewName, "..", ".")
               'Modified by Lydia 2018/08/29 改成檔案日期,並且CPP09存6碼 (ex.FCP-58867卷宗區顯示錯誤)
               'If SaveAttFile_PDF(sCP09, strFilePath & "\" & strMid, strNewName, Val(strSrvDate(1)), Val(Left(Format(ServerTime, "000000"), 4)), False) Then
               '      Kill strFilePath & "\" & strMid
               If SaveAttFile_PDF(sCP09, strFilePath & "\" & strMid, strNewName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False) Then
                     SetAttr strFilePath & "\" & strMid, vbNormal 'Added by Lydia 2020/03/19 預設檔案為正常
                     oFileSys.DeleteFile strFilePath & "\" & strMid, True
               'end 2018/08/29
               Else
                     Exit Function
               End If
            End If
            strMid = Dir(strFilePath & "\*.pdf")
        Loop
        Pub_AutoEsetToCpp = True 'Added by Lydia 2019/11/27
        '無檔案，刪除資料夾
        If Dir(strFilePath & "\*.*") = "" Then
             RmDir strFilePath
        End If
    End If
    
'Added by Lydia 2018/08/29
ErrorHand01:
     If Err.Number <> 0 Then
         MsgBox Err.Description, vbCritical
     End If
     Set oFileSys = Nothing
     Set oFile = Nothing
End Function

'Added by Lydia 2019/12/04 P案發文時，電子送件自動上傳檔案到卷宗區
'依據輸入的智慧局收文號(受理號,ex: 1073066637-0)，將本機C:\E-SET\RdcDocDir\(收文號ex: 1073066637-0)的pdf檔自動搬移到卷宗區;
Public Function Pub_AutoEsetToCppByP(ByVal bolMsg As Boolean, ByVal sPA01 As String, ByVal sPA02 As String, ByVal sPA03 As String, ByVal sPA04 As String, ByVal sPA08 As String, ByVal sCP09 As String, ByVal sCP10 As String, ByVal sFilePath As String, Optional ByVal pCP27 As String) As Boolean
'sPA01~sPA04 本所案號; sPA08 專利種類
'sCP09 收文號;  sCP10 案件性質
'bolMsg 彈訊息是否上傳檔案
Dim strMid As String
Dim strNewName As String
Dim strFilePath As String
Dim oFileSys As New FileSystemObject
Dim oFile
Dim strTitle As String
Dim strCont As String, bCheck As Boolean
Dim rsChk As New ADODB.Recordset
Dim intA As Integer
Dim strTemp As String

On Error GoTo ErrorHand01

    Pub_AutoEsetToCppByP = True
    If sFilePath = "" Then Exit Function
    strFilePath = "C:\E-SET\RdcDocDir\" & sFilePath
    strTitle = "發文電子送件自動上傳檔案到卷宗區"

    strMid = Dir(strFilePath & "\*.pdf")
    
    If strMid = "" Then '因為一併發文會先發文的先上傳，所以資料夾=無
        If bolMsg = True Then
            '先去判斷是否有同天發文的案件性質,若無其他發文則回到發文key智慧局收文號畫面
            '若有則繼續下一步詢問
            'Modified by Lydia 2024/12/26 +電子送件和同一發文字號; ex.FCP-065116在12/24先輸入"自請撤回-發明申請"
            'strTemp = "select cp09,cp10 from caseprogress where cp01='" & sPA01 & "' and cp02='" & sPA02 & "' and cp03='" & sPA03 & "' and cp04='" & sPA04 & "' and cp158=" & IIf(pCP27 <> "", TransDate(pCP27, 2), strSrvDate(1))
            strTemp = "select cp09,cp10 from caseprogress where cp01='" & sPA01 & "' and cp02='" & sPA02 & "' and cp03='" & sPA03 & "' and cp04='" & sPA04 & "'" & _
                      " and cp158=" & IIf(pCP27 <> "", TransDate(pCP27, 2), strSrvDate(1)) & " and nvl(cp118,'N') <> 'N' and instr(cp64,'" & ChgSQL(sFilePath) & "') > 0 "
            '排除無收文號(發文->直接延期,先檢查檔案)
            If sCP09 <> "" Then strTemp = strTemp & " and cp09 <> " & CNULL(sCP09)

            intA = 1
            Set rsChk = ClsLawReadRstMsg(intA, strTemp)
            If intA = 0 Then
                If Pub_StrUserSt03 <> "M51" Then
                    MsgBox "無此收文號，請重新輸入！", vbCritical
                    Pub_AutoEsetToCppByP = False
                    Exit Function
                Else
                    If MsgBox("是否跳過同天發文的判斷？", vbInformation + vbYesNo + vbDefaultButton1, "電腦中心") = vbNo Then
                       Pub_AutoEsetToCppByP = False
                       Exit Function
                    End If
                End If
            End If
            Set rsChk = Nothing
            
            '因為實審和主動修正可併在一起，所以E-Set資料夾查無檔案則彈訊息問是否上傳檔案(by Phoebe)
            If MsgBox("電子送件是否要上傳檔案到卷宗區？" & vbCrLf & "(Yes=要上傳；No=不要上傳，繼續發文)", vbYesNo + vbDefaultButton1 + vbInformation, strTitle) = vbYes Then
                Pub_AutoEsetToCppByP = False
                Exit Function
            End If
        End If
   
    ElseIf sPA01 <> "" And sPA02 <> "" And sCP09 = "" Then  '發文直接做延期，無收文號先檢查.contact
        '檢查本所案號.pdf 是否存在，用於稽核智慧局發文字號和本所案號是否正確（內專後續案不用上傳.contact，只需檢查有符合本所案號.PDF）
        If bCheck = False Then
            Pub_AutoEsetToCppByP = False
            'Added by Lydia 2019/12/12 有則檢查, 若無本所案號開頭.pdf存在，則跳過檢查，繼續上傳作業。
            strCont = Dir(strFilePath & "\" & sPA01 & "*.pdf")
            If strCont <> "" Then
            'end 2019/12/12
                strCont = Dir(strFilePath & "\" & sPA01 & "*" & Val(sPA02) & "*.pdf")
                If strCont = "" Then
                    MsgBox "智慧局收文號和P案號不一致 !", vbCritical, strTitle
                    Exit Function
                End If
            End If 'end 2019/12/12
            'Added by Lydia 2020/04/15 內專統一在發文前,先檢查檔案是否開啟
            strMid = Dir(strFilePath & "\*.pdf")
            Do While strMid <> ""
                 Set oFile = oFileSys.GetFile(strFilePath & "\" & strMid)
                '檢查檔案是否正在使用中
                 If PUB_ChkFileOpening(strFilePath & "\" & strMid) = True Then
                     MsgBox strFilePath & "\" & strMid & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation, strTitle
                     Exit Function
                 End If
                 '檢查檔案大小為 0 KB 有誤
                 If oFile.Size = 0 Then
                       ShowMsg strFilePath & "\" & strMid & vbCrLf & MsgText(9221)
                       Exit Function
                 End If
                 strMid = Dir()
            Loop
            'end 2020/04/15
            Pub_AutoEsetToCppByP = True
            bCheck = True
        End If
    ElseIf sPA01 <> "" And sPA02 <> "" And sPA08 <> "" And sCP09 <> "" And sCP10 <> "" Then '上傳到卷宗區
        Do While strMid <> ""
             '檢查本所案號.pdf 是否存在，用於稽核智慧局發文字號和本所案號是否正確（內專後續案不用上傳.contact，只需檢查有符合本所案號.PDF）
            If bCheck = False Then
                'Added by Lydia 2019/12/12 有則檢查, 若無本所案號開頭.pdf存在，則跳過檢查，繼續上傳作業。
                strCont = Dir(strFilePath & "\" & sPA01 & "*.pdf")
                If strCont <> "" Then
                'end 2019/12/12
                    strCont = Dir(strFilePath & "\" & sPA01 & "*" & Val(sPA02) & "*.pdf")
                    If strCont = "" Then
                        MsgBox "智慧局收文號和P案號不一致 !", vbCritical, strTitle
                        Pub_AutoEsetToCppByP = False
                        Exit Function
                    End If
                End If 'end 2019/12/12
                bCheck = True
            End If
            
            Pub_AutoEsetToCppByP = False
            strNewName = PUB_CaseNo2FileName(sPA01, sPA02, sPA03, sPA04)
            Set oFile = oFileSys.GetFile(strFilePath & "\" & strMid)
           '檢查檔案是否正在使用中
            If PUB_ChkFileOpening(strFilePath & "\" & strMid) = True Then
                MsgBox strFilePath & "\" & strMid & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation, strTitle
                Exit Function
            End If
            '檢查檔案大小為 0 KB 有誤
            If oFile.Size = 0 Then
                  ShowMsg strFilePath & "\" & strMid & vbCrLf & MsgText(9221)
                  Exit Function
            End If
            
            '與FCP案不同，所有檔案均要上傳；包含.FIX_U(劃線本)和.COR_U的檔案。
            'If InStr(UCase(strMid), ".FIX_U") > 0 Or InStr(UCase(strMid), ".COR_U") > 0 Then
            '   Kill strFilePath & "\" & strMid
            'Else
               If InStr(strMid, "中文本") > 0 Then
                     If sPA08 = "1" Then
                          strNewName = strNewName & "." & sCP10 & ".inv.pdf"
                     ElseIf sPA08 = "2" Then
                          strNewName = strNewName & "." & sCP10 & ".utl.pdf"
                     ElseIf sPA08 = "3" Then
                          strNewName = strNewName & "." & sCP10 & ".des.pdf"
                     Else
                          MsgBox "下列檔案請手動上傳卷宗區後，再做發文!" & vbCrLf & strFilePath & "\" & strMid
                          Exit Function
                     End If
               'Modified by Lydia 2024/11/07 更正說明書---比照外專
               ElseIf InStr(strMid, "修正說明書") > 0 Or InStr(strMid, "更正說明書") > 0 Then
                     strNewName = strNewName & "." & sCP10 & ".fix.pdf"
               ElseIf InStr(strMid, "申請書") > 0 Or InStr(strMid, "專簡") > 0 Then
                     strNewName = strNewName & "." & sCP10 & ".data.pdf"
               '其他
               Else
                     If Left(UCase(strMid), Len(strNewName)) = strNewName Then
                           strNewName = strNewName & "." & sCP10 & "." & PUB_GetSimpleName(Mid(strMid, Len(strNewName) + 1))
                     'PXXXXX(6碼) 開頭
                     ElseIf Left(UCase(strMid), Len(sPA01 & sPA02)) = sPA01 & sPA02 Then
                           strNewName = strNewName & "." & sCP10 & "." & PUB_GetSimpleName(Mid(strMid, Len(sPA01 & sPA02) + 1))
                     '+案號
                     Else
                           strNewName = strNewName & "." & sCP10 & "." & PUB_GetSimpleName(strMid)
                     End If
               End If
               strNewName = Replace(strNewName, "..", ".")
               strNewName = UCase(strNewName) '統一為全大寫的檔名
               
               'Added by Lydia 2019/12/12 與FCP案不同，若卷宗區已有相同檔名的檔案，則上傳檔案直接覆蓋原檔案
                strTemp = "SELECT CP09,CPP02 FROM CASEPROGRESS, CASEPAPERPDF " & _
                                 " WHERE CP09='" & sCP09 & "' AND CP09=CPP01(+) AND UPPER(CPP02)=" & CNULL(strNewName)
                intA = 1
                Set rsChk = ClsLawReadRstMsg(intA, strTemp)
                If intA = 1 Then
                    '先刪除原檔
                    If "" & rsChk.Fields("CPP02") <> "" Then
                        If DelAttFile_PDF(sPA01 & "-" & sPA02 & "-" & sPA03 & "-" & sPA04, "" & rsChk.Fields("CP09"), "" & rsChk.Fields("CPP02"), , , True) = False Then
                             MsgBox "刪除卷宗區檔案失敗：" & vbCrLf & "" & rsChk.Fields("CPP02"), vbCritical, "上傳卷宗區作業"
                             Exit Function
                        End If
                    End If
                End If
                'end 2019/12/12
                
               If SaveAttFile_PDF(sCP09, strFilePath & "\" & strMid, strNewName, Format(oFile.DateLastModified, "YYYYMMDD"), Format(oFile.DateLastModified, "HHMMSS"), False) Then
                     oFileSys.DeleteFile strFilePath & "\" & strMid, True
               Else
                     Exit Function
               End If
            'End If
            strMid = Dir(strFilePath & "\*.pdf")
        Loop
        Pub_AutoEsetToCppByP = True
        
        '已完成上傳,檢查電子檔是否齊備
        strTemp = "select cpm26 from casepropertymap where cpm01='" & sPA01 & "' and cpm02='" & sCP10 & "' "
        intA = 1
        Set rsChk = ClsLawReadRstMsg(intA, strTemp)
        If intA = 1 Then
           Call UpdateCP121(sCP09, sCP10, "" & rsChk.Fields("cpm26"))
        End If
        Set rsChk = Nothing
        
        '無檔案，刪除資料夾
        If Dir(strFilePath & "\*.*") = "" Then
             RmDir strFilePath
        End If
        
    End If

ErrorHand01:
     If Err.Number <> 0 Then
         MsgBox Err.Description, vbCritical
     End If
     Set oFileSys = Nothing
     Set oFile = Nothing
End Function


'2011/7/28 自接洽記錄單移過來,並加傳客戶編號
'Add by Morgan 2011/1/10
'客戶應收帳款
'pTotBill:應收總額
'pPaTwFee:台灣專利規費
'pTmTwFee:台灣商標規費
'pTwFee:  台灣案已發文應收規費
'Modify By Sindy 2015/8/28 + Optional strCaseNo As String 本所案號
'Move by Lydia 2018/10/18 計算客戶應收帳款的總額--從basQuery搬來
Public Function GetBillData(ByRef pCustNo As String, ByRef pTotBill As Double, ByRef pPaTwFee As Double _
   , Optional pTmTwFee As Double, Optional pTwFee As Double, Optional strCaseNo As String) As Boolean
Dim stCon As String
Dim stCon1k As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim mRefCust As String 'Added by Lydia 2018/10/18 取得國內帳款管制之關係企業
Dim intQ As Integer, strCon1 As String, rsQ1 As New ADODB.Recordset 'Added by Lydia 2024/03/29

   GetBillData = False
   
   'Modify By Sindy 2015/8/28 抓取個案的應收帳款
   If strCaseNo <> "" Then
      strCP01 = SystemNumber(strCaseNo, 1)
      strCP02 = SystemNumber(strCaseNo, 2)
      strCP03 = SystemNumber(strCaseNo, 3)
      strCP04 = SystemNumber(strCaseNo, 4)
      strCaseNo = Replace(strCaseNo, "-", "")
      
      strDate = Val(DBDATE(DateAdd("m", -6, Format(strSrvDate(1), "####/##/##")))) - 19110000 '半年
      
      stCon = " and a0j02 = '" & strCaseNo & "' and a0k02<" & strDate
      stCon1k = " and a1k13='" & strCP01 & "' and a1k14='" & strCP02 & "' and a1k15='" & strCP03 & "' and a1k16='" & strCP04 & "' and a1k02<" & strDate
   Else
   '2015/8/28 END
      '抓取某客戶的應收帳款
      'Modify By Sindy 2014/8/11 999=>ZZZ
      'stCon = " and a0k03 >= '" & Left(Trim(pCustNo), 6) & "000" & "' and a0k03 <= '" & Left(Trim(pCustNo), 6) & "999" & "'"
      'Modified by Lydia 2018/10/18 取得國內帳款管制之關係企業
      'Memo by Lydia 2020/02/03 應收帳款上限分開管制為個人"應收帳款上限CU183"和"集團應收帳款上限CRA02", 有設定CRA02則需判斷關係企業(6碼)總金額
      stCon = " and a0k03 >= '" & Left(Trim(pCustNo), 6) & "000" & "' and a0k03 <= '" & Left(Trim(pCustNo), 6) & "ZZZ" & "'"
      stCon1k = " and ((a1k03>='" & Left(Trim(pCustNo), 6) & "000" & "' and a1k03<='" & Left(Trim(pCustNo), 6) & "ZZZ" & "')" & _
                  " or (a1k27>='" & Left(Trim(pCustNo), 6) & "000" & "' and a1k27<='" & Left(Trim(pCustNo), 6) & "ZZZ" & "')" & _
                  " or (a1k28>='" & Left(Trim(pCustNo), 6) & "000" & "' and a1k28<='" & Left(Trim(pCustNo), 6) & "ZZZ" & "'))"
      'Remove by Lydia 2020/02/03 改關係企業6碼
      'mRefCust = PUB_GetBillRefCust(pCustNo)
      'stCon = " and a0k03 in (" & GetAddStr(mRefCust) & ") "
      'stCon1k = " and (a1k03 in (" & GetAddStr(mRefCust) & ") or a1k27 in (" & GetAddStr(mRefCust) & ") or a1k28 in (" & GetAddStr(mRefCust) & ")) "
   End If
   
   'Add By Sindy 2015/9/1 國外請款單
   strCon1 = "select sum(A1),sum(A2),sum(A3),sum(A4) from("
   strCon1 = strCon1 & _
               "select nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
               ",nvl(sum(decode(pa09,'000',nvl(a1k09,0),0)),0) as A2" & _
               ",0 as A3" & _
               ",nvl(sum(decode(pa09,'000',nvl(a1k09,0),0)),0) as A4" & _
               " From acc1k0,patent" & _
               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
               " and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04"
   strCon1 = strCon1 & " union all " & _
               "select nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
               ",0 as A2" & _
               ",nvl(sum(decode(tm10,'000',nvl(a1k09,0),0)),0) as A3" & _
               ",nvl(sum(decode(tm10,'000',nvl(a1k09,0),0)),0) as A4" & _
               " From acc1k0,trademark" & _
               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
               " and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04"
   strCon1 = strCon1 & " union all " & _
               "select nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
               ",0 as A2" & _
               ",0 as A3" & _
               ",nvl(sum(decode(lc15,'000',nvl(a1k09,0),0)),0) as A4" & _
               " From acc1k0,lawcase" & _
               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
               " and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04"
   strCon1 = strCon1 & " union all " & _
               "select nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
               ",nvl(sum(decode( instr('1,5',sk02),0,0,decode(sp09,'000',nvl(a1k09,0),0))),0) as A2" & _
               ",nvl(sum(decode( instr('2,6',sk02),0,0,decode(sp09,'000',nvl(a1k09,0),0))),0) as A3" & _
               ",nvl(sum(decode(sp09,'000',nvl(a1k09,0),0)),0) as A4" & _
               " From acc1k0,servicepractice,systemkind" & _
               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
               " and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 and sk01(+)=a1k13"
   strCon1 = strCon1 & " union all "
   'Modify By Sindy 2012/11/02
'   strcon1 = "select sum(nvl(a0j09,0)+nvl(a0j10,0)-nvl(a1u04,0)-nvl(a1u05,0)" & _
'      "-nvl(a1u07,0)-nvl(a1u09,0)+nvl(a1u08,0)+nvl(a1u10,0))" & _
'      ",sum(nvl(decode(a0j04,'000',decode( instr('1,5',sk02),0,0,a0j10)),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0))" & _
'      ",sum(nvl(decode(a0j04,'000',decode( instr('2,6',sk02),0,0,a0j10)),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0))" & _
'      ",sum(nvl(decode(a0j04,'000',a0j10),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0))" & _
'      " From acc0k0, caseprogress, systemkind, acc0j0" & _
'      ",( select a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07" & _
'      ",sum(a1u09) a1u09,sum(a1u08) a1u08,sum(a1u10) a1u10" & _
'      " From acc0k0, caseprogress, acc1u0" & _
'      " where (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & stCon & _
'      " and cp60(+)=a0k01 and cp27>0 and cp79>0 and a1u03(+)=cp09 group by a1u03 ) X" & _
'      " where (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & stCon & _
'      " and cp60(+)=a0k01 and cp27>0 and cp79>0 and sk01(+)=cp01" & _
'      " and a0j01(+)=cp09 and a1u03(+)=cp09"
   'Modified by Lydia 2018/08/30 (應收帳款管控)已發文且已列印收據之國內應收帳款 =>a0k32 is null
   'Modified by Lydia 2023/05/09 排除已銷帳 nvl(a0k37,'Y') <> 'N'
   'Modified by Lydia 2025/06/09 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
   strCon1 = strCon1 & _
      "select sum(nvl(a0j09,0)+nvl(a0j10,0)-nvl(a1u04,0)-nvl(a1u05,0)" & _
      "-nvl(a1u07,0)-nvl(a1u09,0)+nvl(a1u08,0)+nvl(a1u10,0)) as A1" & _
      ",sum( decode(a0j04,'000', nvl(decode(instr('1,5',sk02),0,0,a0j10),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A2" & _
      ",sum( decode(a0j04,'000', nvl(decode(instr('2,6',sk02),0,0,a0j10),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A3" & _
      ",sum( decode(a0j04,'000', nvl(a0j10,0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A4" & _
      " From acc0k0, caseprogress, systemkind, acc0j0" & _
      ",(select a1u02,a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07,sum(a1u09) a1u09,sum(a1u08) a1u08,sum(a1u10) a1u10" & _
      " From acc1u0 where a1u03 in(" & _
      " select distinct a1u03" & _
      " From acc0k0, caseprogress, acc1u0, acc0j0" & _
      " where geta0k32type(a0k01)='1' and nvl(a0k37,'Y') <> 'N' and (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & stCon & _
      " and a0k01=a0j13 and a0j01=cp09 and cp27>0 and cp79>0 and a1u03=cp09) group by a1u02,a1u03) X" & _
      " where geta0k32type(a0k01)='1' and nvl(a0k37,'Y') <> 'N' and (a0k09 is null or a0k09 = 0) and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & stCon & _
      " and a0k01=a0j13 and a0j01=cp09 and cp27>0 and cp79>0 and X.a1u02(+)=a0j13 and X.a1u03(+)=a0j01 and sk01(+)=cp01"
   '2012/11/02 End
   strCon1 = strCon1 & ")"
   intQ = 1
   Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      pTotBill = Val("" & rsQ1.Fields(0))
      pPaTwFee = Val("" & rsQ1.Fields(1))
      pTmTwFee = Val("" & rsQ1.Fields(2))
      pTwFee = Val("" & rsQ1.Fields(3))
   End If
   
   If pTotBill > 0 Then GetBillData = True
   
   Set rsQ1 = Nothing 'Added by Lydia 2024/03/29
End Function

'Added by Lydia 2018/08/22 檢查同一客戶應收帳款是否有超過付款週期
'Move by Lydia 2018/10/18 從basQuery搬來
'Mark by Lydia 2025/06/09 確定無程式使用
'Public Function GetBillDate(ByRef pCustNo As String, Optional ByVal pStaDay As String, Optional ByRef pBillNo As String, Optional ByRef pMemo As String) As Boolean
''pStaDay 收文日
''pBillNo 收據號碼
''pMemo 列印備註
'Dim stCon As String
'Dim stCon1k As String
'Dim intA As Integer, strA1 As String
'Dim rsAD As New ADODB.Recordset
'Dim pMon01 As Integer 'X個月應收帳款
'Dim pMon02 As Integer '付款週期
'Dim strDate1 As String
''Added by Lydia 2018/10/18
'Dim mRefCust As String '取得國內帳款管制之關係企業
'Dim tmpArr1 As Variant, inJ As Integer
'Dim pChkCust As String '應收款的客戶編號
'Dim strCon1 As String 'Added by Lydia 2024/03/29
'
'   GetBillDate = False
'
'   If Trim(pCustNo) = "" Then Exit Function  'Added by Lydia 2018/10/01 新客戶不用檢查
'
''Added by Lydia 2018/10/18 取得國內帳款管制之關係企業,用迴圈組合語法
'mRefCust = PUB_GetBillRefCust(pCustNo)
'tmpArr1 = Empty
'tmpArr1 = Split(mRefCust, ",")
'For inJ = 0 To UBound(tmpArr1)
'   If Trim(tmpArr1(inJ)) <> "" Then
''end 2018/10/18
'   pMemo = "": pBillNo = ""
'   pMon02 = 2 '一般付款週期(暫訂2個月)
'   pMemo = "1"
'   'Modified by Lydia 2018/10/18
'   'strA1 = "select cu175 from customer where cu01||cu02='" & Left(pCustNo & "000", 9) & "' "
'   'Set rsAD = ClsLawReadRstMsg(intA, strA1)
'   strCon1 = "select cu175 from customer where cu01||cu02='" & tmpArr1(inJ) & "' "
'   Set rsAD = ClsLawReadRstMsg(intA, strCon1)
'   'end 2018/10/18
'   If intA = 1 Then
'       If Val("" & rsAD.Fields("cu175")) > 0 Then
'            pMon02 = Val("" & rsAD.Fields("cu175"))
'            pMemo = "2"
'       End If
'   End If
'
'   'If pCustNo = "" And pMon02 = 0 Then Exit Function 'Remove by Lydia 2018/10/01
'   If pStaDay = "" Then pStaDay = strSrvDate(1)
'   '超過Y個月期限: 收文日-1天-付款週期
'   strDate1 = Val(DBDATE(DateAdd("m", pMon02 * -1, Format(CompDate(2, -1, pStaDay), "####/##/##")))) - 19110000
'
'   '國內應收帳款
'   'Modified by Lydia 2018/10/18
'   'stCon = " and nvl(cp27,0) <" & TransDate(strDate1, 2) & " and a0k03 >= '" & Left(Trim(pCustNo), 6) & "000" & "' and a0k03 <= '" & Left(Trim(pCustNo), 6) & "ZZZ" & "'"
'   stCon = " and nvl(cp27,0) <" & TransDate(strDate1, 2) & " and a0k03 = '" & tmpArr1(inJ) & "' "
'   '國外請款單
'   'Modified by Lydia 2018/10/18
'   'stCon1k = " and a1k02<" & strDate1 & _
'                   " and ((a1k03>='" & Left(Trim(pCustNo), 6) & "000" & "' and a1k03<='" & Left(Trim(pCustNo), 6) & "ZZZ" & "')" & _
'                   " or (a1k27>='" & Left(Trim(pCustNo), 6) & "000" & "' and a1k27<='" & Left(Trim(pCustNo), 6) & "ZZZ" & "')" & _
'                   " or (a1k28>='" & Left(Trim(pCustNo), 6) & "000" & "' and a1k28<='" & Left(Trim(pCustNo), 6) & "ZZZ" & "'))"
'     stCon1k = " and a1k02<" & strDate1 & " and (a1k03='" & tmpArr1(inJ) & "' or a1k27='" & tmpArr1(inJ) & "' or a1k28='" & tmpArr1(inJ) & "') "
'
'   '國外請款單
'   'Modified by Lydia 2018/10/18
''   strA1 = "select a1k01 as billno, a1k02 as billdate, nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
''               ",nvl(sum(decode(pa09,'000',nvl(a1k09,0),0)),0) as A2" & _
''               ",0 as A3" & _
''               ",nvl(sum(decode(pa09,'000',nvl(a1k09,0),0)),0) as A4" & _
''               " From acc1k0,patent" & _
''               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
''               " and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04" & _
''               " group by a1k01, a1k02"
''   strA1 = strA1 & " union all " & _
''               "select a1k01 as billno, a1k02 as billdate, nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
''               ",0 as A2" & _
''               ",nvl(sum(decode(tm10,'000',nvl(a1k09,0),0)),0) as A3" & _
''               ",nvl(sum(decode(tm10,'000',nvl(a1k09,0),0)),0) as A4" & _
''               " From acc1k0,trademark" & _
''               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
''               " and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04" & _
''               " group by a1k01, a1k02"
''   strA1 = strA1 & " union all " & _
''               "select a1k01 as billno, a1k02 as billdate, nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
''               ",0 as A2" & _
''               ",0 as A3" & _
''               ",nvl(sum(decode(lc15,'000',nvl(a1k09,0),0)),0) as A4" & _
''               " From acc1k0,lawcase" & _
''               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
''               " and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04" & _
''               " group by a1k01, a1k02"
''   strA1 = strA1 & " union all " & _
''               "select a1k01 as billno, a1k02 as billdate, nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
''               ",nvl(sum(decode( instr('1,5',sk02),0,0,decode(sp09,'000',nvl(a1k09,0),0))),0) as A2" & _
''               ",nvl(sum(decode( instr('2,6',sk02),0,0,decode(sp09,'000',nvl(a1k09,0),0))),0) as A3" & _
''               ",nvl(sum(decode(sp09,'000',nvl(a1k09,0),0)),0) as A4" & _
''               " From acc1k0,servicepractice,systemkind" & _
''               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
''               " and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 and sk01(+)=a1k13" & _
''               " group by a1k01, a1k02"
'   strA1 = strA1 & " union all select a1k01 as billno, a1k02 as billdate, nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
'               ",nvl(sum(decode(pa09,'000',nvl(a1k09,0),0)),0) as A2" & _
'               ",0 as A3" & _
'               ",nvl(sum(decode(pa09,'000',nvl(a1k09,0),0)),0) as A4,a1k03 as cust1, a1k27 as cust2, a1k28 as cust3," & pMon02 & " as cu175" & _
'               " From acc1k0,patent" & _
'               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
'               " and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04" & _
'               " group by a1k01, a1k02,a1k03,a1k27,a1k28"
'   strA1 = strA1 & " union all " & _
'               "select a1k01 as billno, a1k02 as billdate, nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
'               ",0 as A2" & _
'               ",nvl(sum(decode(tm10,'000',nvl(a1k09,0),0)),0) as A3" & _
'               ",nvl(sum(decode(tm10,'000',nvl(a1k09,0),0)),0) as A4,a1k03 as cust1, a1k27 as cust2, a1k28 as cust3," & pMon02 & " as cu175" & _
'               " From acc1k0,trademark" & _
'               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
'               " and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04" & _
'               " group by a1k01, a1k02,a1k03,a1k27,a1k28"
'   strA1 = strA1 & " union all " & _
'               "select a1k01 as billno, a1k02 as billdate, nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
'               ",0 as A2" & _
'               ",0 as A3" & _
'               ",nvl(sum(decode(lc15,'000',nvl(a1k09,0),0)),0) as A4,a1k03 as cust1, a1k27 as cust2, a1k28 as cust3," & pMon02 & " as cu175" & _
'               " From acc1k0,lawcase" & _
'               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
'               " and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04" & _
'               " group by a1k01, a1k02,a1k03,a1k27,a1k28"
'   strA1 = strA1 & " union all " & _
'               "select a1k01 as billno, a1k02 as billdate, nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
'               ",nvl(sum(decode( instr('1,5',sk02),0,0,decode(sp09,'000',nvl(a1k09,0),0))),0) as A2" & _
'               ",nvl(sum(decode( instr('2,6',sk02),0,0,decode(sp09,'000',nvl(a1k09,0),0))),0) as A3" & _
'               ",nvl(sum(decode(sp09,'000',nvl(a1k09,0),0)),0) as A4,a1k03 as cust1, a1k27 as cust2, a1k28 as cust3," & pMon02 & " as cu175" & _
'               " From acc1k0,servicepractice,systemkind" & _
'               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
'               " and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 and sk01(+)=a1k13" & _
'               " group by a1k01, a1k02,a1k03,a1k27,a1k28"
''end 2018/10/18
'   strA1 = strA1 & " union all "
'
'   '國內應收帳款: 已發文且已列印收據(a0k32 is null 財務認定)之應收帳款
'   'Modified by Lydia 2018/10/18
'   'strA1 = strA1 & _
'      "select a0k01 as billno, a0k02 as billdate, sum(nvl(a0j09,0)+nvl(a0j10,0)-nvl(a1u04,0)-nvl(a1u05,0)" & _
'      "-nvl(a1u07,0)-nvl(a1u09,0)+nvl(a1u08,0)+nvl(a1u10,0)) as A1" & _
'      ",sum( decode(a0j04,'000', nvl(decode(instr('1,5',sk02),0,0,a0j10),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A2" & _
'      ",sum( decode(a0j04,'000', nvl(decode(instr('2,6',sk02),0,0,a0j10),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A3" & _
'      ",sum( decode(a0j04,'000', nvl(a0j10,0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A4" & _
'      " From acc0k0, caseprogress, systemkind, acc0j0" & _
'      ",(select a1u02,a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07,sum(a1u09) a1u09,sum(a1u08) a1u08,sum(a1u10) a1u10" & _
'      " From acc1u0 where a1u03 in(" & _
'      " select distinct a1u03" & _
'      " From acc0k0, caseprogress, acc1u0, acc0j0" & _
'      " where a0k32 is null and nvl(a0k09,0)=0 and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & stCon & _
'      " and a0k01=a0j13(+) and a0j01=cp09(+) and nvl(cp27,0)>0 and cp79>0 and a1u03(+)=cp09) group by a1u02,a1u03) X" & _
'      " where a0k32 is null and nvl(a0k09,0)=0 and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & stCon & _
'      " and a0k01=a0j13(+) and a0j01=cp09(+) and nvl(cp27,0)>0 and cp79>0 and X.a1u02(+)=a0j13 and X.a1u03(+)=a0j01 and sk01(+)=cp01" & _
'      " group by a0k01, a0k02"
'   'Modified by Lydia 2019/02/23 顯示日期a0k02改為發文日cp27; (ex.P-111365的AA7028388發文日1070724超過6個月,但是E10715853收據日期改為1080103)
'   'Modified by Lydia 2023/05/09 排除已銷帳 nvl(a0k37,'Y') <> 'N'
'   strA1 = strA1 & _
'      "select a0k01 as billno, cp27 as billdate, sum(nvl(a0j09,0)+nvl(a0j10,0)-nvl(a1u04,0)-nvl(a1u05,0)" & _
'      "-nvl(a1u07,0)-nvl(a1u09,0)+nvl(a1u08,0)+nvl(a1u10,0)) as A1" & _
'      ",sum( decode(a0j04,'000', nvl(decode(instr('1,5',sk02),0,0,a0j10),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A2" & _
'      ",sum( decode(a0j04,'000', nvl(decode(instr('2,6',sk02),0,0,a0j10),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A3" & _
'      ",sum( decode(a0j04,'000', nvl(a0j10,0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A4,a0k03 as cust1, null as cust2, null as cust3," & pMon02 & " as cu175" & _
'      " From acc0k0, caseprogress, systemkind, acc0j0" & _
'      ",(select a1u02,a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07,sum(a1u09) a1u09,sum(a1u08) a1u08,sum(a1u10) a1u10" & _
'      " From acc1u0 where a1u03 in(" & _
'      " select distinct a1u03" & _
'      " From acc0k0, caseprogress, acc1u0, acc0j0" & _
'      " where a0k32 is null and nvl(a0k37,'Y') <> 'N' and nvl(a0k09,0)=0 and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & stCon & _
'      " and a0k01=a0j13(+) and a0j01=cp09(+) and nvl(cp27,0)>0 and cp79>0 and a1u03(+)=cp09) group by a1u02,a1u03) X" & _
'      " where a0k32 is null and nvl(a0k37,'Y') <> 'N' and nvl(a0k09,0)=0 and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & stCon & _
'      " and a0k01=a0j13(+) and a0j01=cp09(+) and nvl(cp27,0)>0 and cp79>0 and X.a1u02(+)=a0j13 and X.a1u03(+)=a0j01 and sk01(+)=cp01" & _
'      " group by a0k01, cp27,a0k03"
''Added by Lydia 2018/10/18 組合語法
'   End If 'If Trim(tmpArr1(inJ)) <> "" Then
'Next inJ
'   strA1 = Mid(strA1, 11)  '去掉開頭" union all "
''end 2018/10/18
'   strA1 = strA1 & " order by billdate asc"
'   intA = 1
'   Set rsAD = ClsLawReadRstMsg(intA, strA1)
'   If intA = 1 Then
'       rsAD.MoveFirst
'       Do While Not rsAD.EOF
'            If "" & rsAD.Fields("billno") <> "" And "" & rsAD.Fields("billdate") <> "" Then
'                 pBillNo = "" & rsAD.Fields("billno")
'                 strA1 = TransDate("" & rsAD.Fields("billdate"), 2)
'                 'Added by Lydia 2018/10/18 判斷付款週期
'                 If "" & rsAD.Fields("cu175") = "2" Then
'                     pMemo = "1"
'                 Else
'                     pMemo = "2"
'                 End If
'                 '記錄應收款的客戶編號
'                 If "" & rsAD.Fields("cust1") <> "" And InStr(mRefCust, "" & rsAD.Fields("cust1")) > 0 Then
'                      pChkCust = "" & rsAD.Fields("cust1")
'                 ElseIf "" & rsAD.Fields("cust2") <> "" And InStr(mRefCust, "" & rsAD.Fields("cust2")) > 0 Then
'                      pChkCust = "" & rsAD.Fields("cust2")
'                 ElseIf "" & rsAD.Fields("cust3") <> "" And InStr(mRefCust, "" & rsAD.Fields("cust3")) > 0 Then
'                      pChkCust = "" & rsAD.Fields("cust3")
'                 End If
'                 'end 2018/10/18
'                 Exit Do
'            End If
'            rsAD.MoveNext
'       Loop
'
'       If pBillNo <> "" Then
'          'X個月應收帳款
'          pMon01 = DateDiff("m", ChangeWStringToWDateString(strA1), ChangeWStringToWDateString(strSrvDate(1)))
'          If pMemo = "1" Then '一般付款週期
'              'Modified by Lydia 2018/10/18 +客戶編號
'              'pMemo = "本客戶有(" & pBillNo & ")" & pMon01 & "個月應收帳款已逾付款週期"
'              pMemo = "本客戶" & IIf(pChkCust <> "", "(" & pChkCust & ")", "") & "有(" & pBillNo & ")" & pMon01 & "個月應收帳款已逾付款週期"
'          ElseIf pMemo = "2" Then '特殊付款週期
'              'Modified by Lydia 2018/10/18 +客戶編號
'              'pMemo = "本客戶有(" & pBillNo & ")" & pMon01 & "個月應收帳款，已逾付款週期" & pMon02 & "個月"
'              pMemo = "本客戶" & IIf(pChkCust <> "", "(" & pChkCust & ")", "") & "有(" & pBillNo & ")" & pMon01 & "個月應收帳款，已逾付款週期" & pMon02 & "個月"
'          End If
'          GetBillDate = True
'       End If
'   End If
'   Set rsAD = Nothing
'End Function
'end 2025/06/09

'Added by Lydia 2018/10/18 取得國內帳款管制之關係企業
Public Function PUB_GetBillRefCust(ByVal pOldNo As String) As String
Dim inJ As Integer
Dim rsChk As New ADODB.Recordset
Dim strTmp1 As String, strTmp2 As String
Dim strMidList As String
Dim strCustNo As String, strCustName As String
Dim strSalesNo As String, strSalesArea As String
 
    If pOldNo = "" Then Exit Function
    'Added by Lydia 2020/02/03 付款週期與應收帳款額度分別管制:
'　根據1081128智權主管會議中，主管建議調整付款週期計算方式及關係企業之付款週期、應收帳款額度分開管制，您在會上亦指示可放寬，故擬定資料請作單如附件，主要調整內容如下：
'　(1)一般付款週期維持2個月，惟計算方式以發文日當月的翌月1日起計。
'　(2)關係企業屬於下列情況者，其付款週期與應收帳款額度分開獨立管制：
'　　A.關係企業之公司名稱不同。
'　　B.相同公司名稱有不同客戶編號，由不同智權人員服務。
'　經上述調整後，本所原先定義的關係企業，基本上均以獨立的付款週期與應收帳款額度分別管制。
    PUB_GetBillRefCust = ChangeCustomerL(pOldNo)
    Exit Function
    'end 2020/02/03
    
'1.以輸入之客戶編號抓出智權人員A
    strTmp1 = "select 1 ord1, cu01||cu02 as custno,nvl(cu04,nvl(cu05,cu06)) custname ,cu13, st15 from customer,staff where cu01||cu02='" & Left(pOldNo & String(9, "0"), 9) & "' and cu13=st01(+) "
    strTmp1 = strTmp1 & "union select 2 ord1, cu01||cu02 as custno,nvl(cu04,nvl(cu05,cu06)) custname ,cu13, st15 from customer,staff where CU01>='" & Left(pOldNo & String(9, "0"), 6) & "00' AND CU01<='" & Left(pOldNo & String(9, "0"), 6) & "zz' and cu13=st01(+) "
    strTmp1 = strTmp1 & "order by 1,2 "
    inJ = 1
    Set rsChk = ClsLawReadRstMsg(inJ, strTmp1)
    If inJ = 1 Then
        rsChk.MoveFirst
        strCustNo = "" & rsChk.Fields("custno")
        strCustName = Trim("" & rsChk.Fields("custname"))
        strSalesNo = "" & rsChk.Fields("cu13")
        strSalesArea = "" & rsChk.Fields("st15")
        Do While Not rsChk.EOF
             If "" & rsChk.Fields("ord1") = "1" Then
                 '符合條件的客戶代號
                 strMidList = strMidList & "," & rsChk.Fields("custno")
                 '因為可能有造字,所以不用trim去空白
                 If "" & rsChk.Fields("custname") <> "" Then strTmp2 = strTmp2 & " or instr(custname,'" & rsChk.Fields("custname") & "') > 0"
'2.關係企業之智權人員業務區ST15與A之業務區相同者，抓出關係企業編號。
             ElseIf "" & rsChk.Fields("st15") = strSalesArea Then
                   If InStr(strMidList, "" & rsChk.Fields("custno")) = 0 Then
                       strMidList = strMidList & "," & rsChk.Fields("custno")
                       '因為可能有造字,所以不用trim去空白
                       If "" & rsChk.Fields("custname") <> "" Then strTmp2 = strTmp2 & " or instr(custname,'" & rsChk.Fields("custname") & "') > 0"
                   End If
             End If
             rsChk.MoveNext
        Loop

'3.再以2之編號的名稱再去抓關係企業中名稱相同者。
        strTmp1 = "select cu01||cu02 as custno,nvl(cu04,nvl(cu05,cu06)) custname,cu13, st15 from customer,staff WHERE CU01>='" & Left(strCustNo, 6) & "00' AND CU01<='" & Left(strCustNo, 6) & "zz' and cu13=st01(+) "
        strTmp2 = "select * from (" & strTmp1 & ") where " & Mid(strTmp2, 4) & " order by custno"
        inJ = 1
        Set rsChk = ClsLawReadRstMsg(inJ, strTmp2)
        If inJ = 1 Then
            rsChk.MoveFirst
            strTmp2 = ""
            Do While Not rsChk.EOF
                '符合條件的客戶代號
                If InStr(strMidList, "" & rsChk.Fields("custno")) = 0 Then
                      strMidList = strMidList & "," & rsChk.Fields("custno")
'4.3抓出之名稱若為更名前名稱要再抓其所有更名編號。
                      If Val(Mid("" & rsChk.Fields("custno"), 9, 1)) > 0 Then
                          For inJ = 0 To Val(Mid("" & rsChk.Fields("custno"), 9, 1)) - 1
                              strMidList = strMidList & "," & Left(rsChk.Fields("custno"), 8) & inJ
                          Next inJ
                      End If
                End If
                rsChk.MoveNext
            Loop
        End If
    End If
    
'5.2+3+4的編號是要檢查應收帳號的編號。
'範例： X00361、X29066、X39289、X48213、X78186、X72747。
    If strMidList <> "" Then PUB_GetBillRefCust = Mid(strMidList, 2)
End Function

'Added by Lydia 2019/05/10 計算客戶應收帳款的總額和判斷付款週期的應收帳款
'合併GetBillData和GetBillDate(P-97630因為關係企業代號多，所以在應收帳款總額和週期判斷語法會變慢)
'從GetBillDate改成不判斷CP27逐筆加總A1~A4，再判斷最早應收帳款
'Modified by Lydia 2022/06/13 傳入收文之本所案號,案件性質(可用,串接)=> ByVal pCaseNO As String, ByVal pCasePTY As String
'Modified by Lydia 2022/06/15 傳入收文之智權人員=> ByVal pSalesNo As String
Public Function PUB_GetBillDataAll(ByVal pKind As String, ByVal pCustNo As String, ByVal pCaseNo As String, ByVal pCasePty As String, ByVal pSalesNo As String _
                                            , Optional ByRef pTotBill As Double, Optional ByRef pPaTwFee As Double, Optional ByRef pTmTwFee As Double, Optional ByRef pTwFee As Double, Optional ByRef strCaseNo As String _
                                            , Optional ByVal pStaDay As String, Optional ByRef pBillNo As String, Optional ByRef pMemo As String) As Boolean
'pKind: 1-計算客戶應收帳款的總額、2-判斷付款週期的應收帳款、3或空白=兩項全做
'pTotBill:應收總額
'pPaTwFee:台灣專利規費
'pTmTwFee:台灣商標規費
'pTwFee:  台灣案已發文應收規費
'strCaseNo: (指定)本所案號
'pStaDay 收文日
'pBillNo 收據號碼
'pMemo 列印備註
Dim stCon As String
Dim stCon1k As String
Dim intA As Integer, strA1 As String
Dim strCP01 As String, strCP02 As String, strCP03 As String, strCP04 As String
Dim rsAD As New ADODB.Recordset
Dim pMon01 As Integer 'X個月應收帳款
Dim pMon02 As Integer '付款週期
Dim strDate1 As String
Dim mRefCust As String '取得國內帳款管制之關係企業
Dim tmpArr1 As Variant, inJ As Integer
Dim pChkCust As String '應收款的客戶編號
Dim iRound As Integer
'Added by Lydia 2019/05/16
Dim rs1 As New ADODB.Recordset
Dim strB1 As String, strMid As String
Dim bolTips As Boolean, m_JumpTips As String 'Added by Lydia 2022/06/13

   PUB_GetBillDataAll = False
   
   If Trim(pCustNo) = "" Then Exit Function  '新客戶不用檢查
   
   'Added by Lydia 2022/06/15 解除杜協理及吳經理收文時，應收帳款總額.寬限期及前一道年費程序未收等之主管簽核程序，即於接洽單上不再出現相關文字，而使收文人員可利於處理。
                                             '前一道年費程序未收=605年費/606維持費/607延展費程序，但仍有應收帳款，需主管簽核後方可收文！
   'Modified by Lydia 2022/09/13 改抓特殊設定
   'If pSalesNo <> "" And InStr("74018,70005", pSalesNo) > 0 Then
   If pSalesNo <> "" And InStr(Pub_GetSpecMan("應收帳款上限檢查排除"), pSalesNo) > 0 Then
       Exit Function
   End If
   'end 2022/06/15
   
   If strCaseNo <> "" Then
        mRefCust = ChangeCustomerL(pCustNo)
   Else
        '取得國內帳款管制之關係企業,用迴圈組合語法
        mRefCust = PUB_GetBillRefCust(ChangeCustomerL(pCustNo)) 'Memo by Lydia 2020/02/03 改成單一客戶管制
   End If
    tmpArr1 = Empty
    tmpArr1 = Split(mRefCust, ",")
    'Added by Lydia 2022/06/13 是否控管ACS案件的TIPS收款
    If pCaseNo <> "" And pCasePty <> "" Then
       If PUB_ChkACSforTIPS(pCaseNo, pCasePty) = True Then
           bolTips = True
       End If
    End If
    'end 2022/06/13
    
For inJ = 0 To UBound(tmpArr1)
   If Trim(tmpArr1(inJ)) <> "" Then
   pMemo = "": pBillNo = ""
   pMon02 = 2 '一般付款週期(暫訂2個月)
   pMemo = "1"
   '改成不判斷CP27逐筆加總
   'strExc(0) = "select cu175 from customer where cu01||cu02='" & tmpArr1(inJ) & "' "
   'intA = 1
   'Set rsAD = ClsLawReadRstMsg(intA, strExc(0))
   'If intA = 1 Then
   '    If Val("" & rsAD.Fields("cu175")) > 0 Then
   '         pMon02 = Val("" & rsAD.Fields("cu175"))
   '         pMemo = "2"
   '    End If
   'End If
   If pStaDay = "" Then pStaDay = strSrvDate(1)
      
   '超過Y個月期限: 收文日-1天-付款週期
   strDate1 = Val(DBDATE(DateAdd("m", pMon02 * -1, Format(CompDate(2, -1, pStaDay), "####/##/##")))) - 19110000

   If strCaseNo <> "" Then '抓取個案的逾期應收帳款
        If strCP01 = "" Then
            strCP01 = SystemNumber(strCaseNo, 1)
            strCP02 = SystemNumber(strCaseNo, 2)
            strCP03 = SystemNumber(strCaseNo, 3)
            strCP04 = SystemNumber(strCaseNo, 4)
            strCaseNo = Replace(strCaseNo, "-", "")
        End If
        strDate1 = Val(DBDATE(DateAdd("m", -6, Format(strSrvDate(1), "####/##/##")))) - 19110000 '半年
        
        stCon = " and a0j02 = '" & strCaseNo & "' and a0k02<" & strDate1
        stCon1k = " and a1k13='" & strCP01 & "' and a1k14='" & strCP02 & "' and a1k15='" & strCP03 & "' and a1k16='" & strCP04 & "' and a1k02<" & strDate1
   Else  ' 抓取某客戶的應收帳款 (含同一區業務的關係企業)
        '國內應收帳款
        'Modified by Lydia 2022/09/13 改用客戶編號前8碼來統計客戶應收帳款的總額和判斷付款週期 ; ex. CFT-022806申請人X82367001已超過應收帳款總額
        'stCon = " and nvl(cp27,0) > 0  and a0k03 = '" & tmpArr1(inJ) & "' "
        stCon = " and nvl(cp27,0) > 0  and a0k03 >= '" & Left(tmpArr1(inJ), 8) & "0" & "' and a0k03 <= '" & Left(tmpArr1(inJ), 8) & "Z" & "' "
        ''國外請款單
        'Modified by Lydia 2022/09/13 改用客戶編號前8碼來統計客戶應收帳款的總額和判斷付款週期
        'stCon1k = " and (a1k03='" & tmpArr1(inJ) & "' or a1k27='" & tmpArr1(inJ) & "' or a1k28='" & tmpArr1(inJ) & "') "
        stCon1k = " and ((a1k03>='" & Left(tmpArr1(inJ), 8) & "0" & "' and a1k03<='" & Left(tmpArr1(inJ), 8) & "Z" & "') or " & _
                        " (a1k27>='" & Left(tmpArr1(inJ), 8) & "0" & "' and a1k27<='" & Left(tmpArr1(inJ), 8) & "Z" & "') or " & _
                        " (a1k28>='" & Left(tmpArr1(inJ), 8) & "0" & "' and a1k28<='" & Left(tmpArr1(inJ), 8) & "Z" & "')) "
   End If
   
   '國外請款單
   'Modified by Lydia 2022/06/13 配合國內收款增加本所案號a1k13||a1k14||a1k15||a1k16 as caseno, '601' as casepty
   strA1 = strA1 & " union all select a1k01 as billno, a1k02 as billdate, nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
               ",nvl(sum(decode(pa09,'000',nvl(a1k09,0),0)),0) as A2" & _
               ",0 as A3" & _
               ",nvl(sum(decode(pa09,'000',nvl(a1k09,0),0)),0) as A4,a1k03 as cust1, a1k27 as cust2, a1k28 as cust3," & pMon02 & " as cu175" & _
               ",a1k13||a1k14||a1k15||a1k16 as caseno, '601' as casepty " & _
               " From acc1k0,patent" & _
               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
               " and a1k13=pa01 and a1k14=pa02 and a1k15=pa03 and a1k16=pa04" & _
               " group by a1k01, a1k02,a1k03,a1k27,a1k28,a1k13||a1k14||a1k15||a1k16"
   'Modified by Lydia 2022/06/13 配合國內收款增加本所案號a1k13||a1k14||a1k15||a1k16 as caseno, '601' as casepty
   strA1 = strA1 & " union all " & _
               "select a1k01 as billno, a1k02 as billdate, nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
               ",0 as A2" & _
               ",nvl(sum(decode(tm10,'000',nvl(a1k09,0),0)),0) as A3" & _
               ",nvl(sum(decode(tm10,'000',nvl(a1k09,0),0)),0) as A4,a1k03 as cust1, a1k27 as cust2, a1k28 as cust3," & pMon02 & " as cu175" & _
               ",a1k13||a1k14||a1k15||a1k16 as caseno, '601' as casepty " & _
               " From acc1k0,trademark" & _
               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
               " and a1k13=tm01 and a1k14=tm02 and a1k15=tm03 and a1k16=tm04" & _
               " group by a1k01, a1k02,a1k03,a1k27,a1k28,a1k13||a1k14||a1k15||a1k16"
   'Modified by Lydia 2022/06/13 配合國內收款增加本所案號a1k13||a1k14||a1k15||a1k16 as caseno, '601' as casepty
   strA1 = strA1 & " union all " & _
               "select a1k01 as billno, a1k02 as billdate, nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
               ",0 as A2" & _
               ",0 as A3" & _
               ",nvl(sum(decode(lc15,'000',nvl(a1k09,0),0)),0) as A4,a1k03 as cust1, a1k27 as cust2, a1k28 as cust3," & pMon02 & " as cu175" & _
               ",a1k13||a1k14||a1k15||a1k16 as caseno, '601' as casepty " & _
               " From acc1k0,lawcase" & _
               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
               " and a1k13=lc01 and a1k14=lc02 and a1k15=lc03 and a1k16=lc04" & _
               " group by a1k01, a1k02,a1k03,a1k27,a1k28,a1k13||a1k14||a1k15||a1k16"
   'Modified by Lydia 2022/06/13 配合國內收款增加本所案號a1k13||a1k14||a1k15||a1k16 as caseno, '601' as casepty
   strA1 = strA1 & " union all " & _
               "select a1k01 as billno, a1k02 as billdate, nvl(sum(nvl(a1k11,0)-nvl(a1k30,0)-nvl(a1k06,0)),0) as A1" & _
               ",nvl(sum(decode( instr('1,5',sk02),0,0,decode(sp09,'000',nvl(a1k09,0),0))),0) as A2" & _
               ",nvl(sum(decode( instr('2,6',sk02),0,0,decode(sp09,'000',nvl(a1k09,0),0))),0) as A3" & _
               ",nvl(sum(decode(sp09,'000',nvl(a1k09,0),0)),0) as A4,a1k03 as cust1, a1k27 as cust2, a1k28 as cust3," & pMon02 & " as cu175" & _
               ",a1k13||a1k14||a1k15||a1k16 as caseno, '601' as casepty " & _
               " From acc1k0,servicepractice,systemkind" & _
               " where nvl(a1k12, 0)=0 And a1k25 Is Null And a1k29 Is Null" & stCon1k & _
               " and a1k13=sp01 and a1k14=sp02 and a1k15=sp03 and a1k16=sp04 and sk01(+)=a1k13" & _
               " group by a1k01, a1k02,a1k03,a1k27,a1k28,a1k13||a1k14||a1k15||a1k16"
   strA1 = strA1 & " union all "

   '國內應收帳款: 已發文且已列印收據(a0k32 is null 財務認定)之應收帳款
   'Memo by Lydia 2019/05/14 去掉(+)速度較快: a0k01=a0j13(+) and a0j01=cp09(+) and nvl(cp27,0)>0 and cp79>0 and a1u03(+)=cp09 => a0k01=a0j13 and a0j01=cp09 and cp27>0 and cp79>0 and a1u03=cp09
                                                                         'a0k01=a0j13(+) and a0j01=cp09(+) and nvl(cp27,0)>0 and cp79>0 => a0k01=a0j13 and a0j01=cp09 and cp27>0 and cp79>0
    'strA1 = strA1 & _
      "select a0k01 as billno, cp27 as billdate, sum(nvl(a0j09,0)+nvl(a0j10,0)-nvl(a1u04,0)-nvl(a1u05,0)" & _
      "-nvl(a1u07,0)-nvl(a1u09,0)+nvl(a1u08,0)+nvl(a1u10,0)) as A1" & _
      ",sum( decode(a0j04,'000', nvl(decode(instr('1,5',sk02),0,0,a0j10),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A2" & _
      ",sum( decode(a0j04,'000', nvl(decode(instr('2,6',sk02),0,0,a0j10),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A3" & _
      ",sum( decode(a0j04,'000', nvl(a0j10,0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A4,a0k03 as cust1, null as cust2, null as cust3," & pMon02 & " as cu175" & _
      " From acc0k0, caseprogress, systemkind, acc0j0" & _
      ",(select a1u02,a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07,sum(a1u09) a1u09,sum(a1u08) a1u08,sum(a1u10) a1u10" & _
      " From acc1u0 where a1u03 in(" & _
      " select distinct a1u03" & _
      " From acc0k0, caseprogress, acc1u0, acc0j0" & _
      " where a0k32 is null and nvl(a0k09,0)=0 and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & stCon & _
      " and a0k01=a0j13(+) and a0j01=cp09(+) and nvl(cp27,0)>0 and cp79>0 and a1u03(+)=cp09) group by a1u02,a1u03) X" & _
      " where a0k32 is null and nvl(a0k09,0)=0 and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & stCon & _
      " and a0k01=a0j13(+) and a0j01=cp09(+) and nvl(cp27,0)>0 and cp79>0 and X.a1u02(+)=a0j13 and X.a1u03(+)=a0j01 and sk01(+)=cp01" & _
      " group by a0k01, cp27,a0k03"
    'Modified by Lydia 2022/06/13 增加本所案號cp01||cp02||cp03||cp04 as caseno, cp10
    'Modified by Lydia 2023/05/09 排除已銷帳 nvl(a0k37,'Y') <> 'N'
    'Modified by Lydia 2023/08/31 抓acc0k0明細只抓未收款 nvl(a0k37,'Y') <> 'N'=> a0k37 is null
    'Modified by Lydia 2025/06/09 a0k32 is null 改用函數判斷：geta0k32type(a0k01)='1'
    strA1 = strA1 & _
      "select a0k01 as billno, cp27 as billdate, sum(nvl(a0j09,0)+nvl(a0j10,0)-nvl(a1u04,0)-nvl(a1u05,0)" & _
      "-nvl(a1u07,0)-nvl(a1u09,0)+nvl(a1u08,0)+nvl(a1u10,0)) as A1" & _
      ",sum( decode(a0j04,'000', nvl(decode(instr('1,5',sk02),0,0,a0j10),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A2" & _
      ",sum( decode(a0j04,'000', nvl(decode(instr('2,6',sk02),0,0,a0j10),0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A3" & _
      ",sum( decode(a0j04,'000', nvl(a0j10,0)-nvl(a1u05,0)-nvl(a1u09,0)+nvl(a1u10,0) ,0) ) as A4,a0k03 as cust1, null as cust2, null as cust3," & pMon02 & " as cu175" & _
      ",cp01||cp02||cp03||cp04 as caseno, cp10 as casepty" & _
      " From acc0k0, caseprogress, systemkind, acc0j0" & _
      ",(select a1u02,a1u03,sum(a1u04) a1u04,sum(a1u05) a1u05,sum(a1u07) a1u07,sum(a1u09) a1u09,sum(a1u08) a1u08,sum(a1u10) a1u10" & _
      " From acc1u0 where a1u03 in(" & _
      " select distinct a1u03" & _
      " From acc0k0, caseprogress, acc1u0, acc0j0" & _
      " where geta0k32type(a0k01)='1' and nvl(a0k37,'Y') <> 'N' and nvl(a0k09,0)=0 and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & stCon & _
      " and a0k01=a0j13 and a0j01=cp09 and cp27>0 and cp79>0 and a1u03=cp09) group by a1u02,a1u03) X" & _
      " where geta0k32type(a0k01)='1' and a0k37 is null and nvl(a0k09,0)=0 and (a0k06+a0k07) > (nvl(a0k17, 0)+nvl(a0k18, 0))" & stCon & _
      " and a0k01=a0j13 and a0j01=cp09 and cp27>0 and cp79>0 and X.a1u02(+)=a0j13 and X.a1u03(+)=a0j01 and sk01(+)=cp01" & _
      " group by a0k01, cp27,a0k03,cp01||cp02||cp03||cp04,cp10 "
'組合語法
   End If
Next inJ
'----------------------------------------------------------------------------
   strA1 = Mid(strA1, 11)  '去掉開頭" union all "
   If pKind = "1" Or strCaseNo <> "" Then
        strA1 = "SELECT SUM(A1) A1,SUM(A2) A2,SUM(A3) A3,SUM(A4) A4 FROM (" & strA1 & ") "
        intA = 1
        Set rsAD = ClsLawReadRstMsg(intA, strA1)
        If intA = 1 Then
            pTotBill = Val("" & rsAD.Fields("A1"))
            pPaTwFee = Val("" & rsAD.Fields("A2"))
            pTmTwFee = Val("" & rsAD.Fields("A3"))
            pTwFee = Val("" & rsAD.Fields("A4"))
        End If
        If pTotBill > 0 Then PUB_GetBillDataAll = True '有應收帳款
   Else
        pTotBill = 0
        pPaTwFee = 0
        pTmTwFee = 0
        pTwFee = 0
        pBillNo = ""
        pMemo = ""
        
        'Modified by Lydia 2023/08/31 +billno asc ;ex.P-127023的AB1012398收據有4張
        strA1 = strA1 & " order by billdate asc, billno asc"
        intA = 1
        Set rsAD = ClsLawReadRstMsg(intA, strA1)
        If intA = 1 Then
            rsAD.MoveFirst
            Do While Not rsAD.EOF
              'Added by Lydia 2022/06/13 非ACS案: 剔除案件ACS的曾有TIPS案件性質的案號的應收帳款
              m_JumpTips = ""
              If bolTips = False Then
                 If PUB_ChkACSforTIPS("" & rsAD.Fields("caseno"), "" & rsAD.Fields("casepty")) = True Then
                    m_JumpTips = "Y"
                 End If
              End If
              If m_JumpTips <> "Y" Then
              'end 2022/06/13
                 If "" & rsAD.Fields("billno") <> "" And "" & rsAD.Fields("billdate") <> "" Then
                      If pBillNo = "" Then '只顯示最早的應收帳款資料
                            '判斷付款週期
                            'Modified by Lydia 2019/05/16 另外抓客戶的付款週期CU175
                            'If "" & rsAD.Fields("cu175") = "2" Then
                            strB1 = "" '第一個符合關係企業的客戶代號
                            If strB1 = "" And "" & rsAD.Fields("cust1") <> "" And InStr(mRefCust, "" & rsAD.Fields("cust1")) > 0 Then strB1 = "" & rsAD.Fields("cust1")
                            If strB1 = "" And "" & rsAD.Fields("cust2") <> "" And InStr(mRefCust, "" & rsAD.Fields("cust2")) > 0 Then strB1 = "" & rsAD.Fields("cust2")
                            If strB1 = "" And "" & rsAD.Fields("cust3") <> "" And InStr(mRefCust, "" & rsAD.Fields("cust3")) > 0 Then strB1 = "" & rsAD.Fields("cust3")
                            If strB1 <> strMid Then
                                 strA1 = "select nvl(cu175,2) cu175 from customer where cu01||cu02='" & strB1 & "' "
                                 intA = 1
                                 Set rs1 = ClsLawReadRstMsg(intA, strA1)
                                 pMon02 = 2
                                 If intA = 1 Then
                                     pMon02 = Val("" & rs1.Fields("cu175"))
                                 End If
                                 strMid = strB1
                            End If
                            If pMon02 = 2 Then
                            'end 2019/05/16
                                pMemo = "1"
                            Else
                                pMemo = "2"
                            End If
                            '超過Y個月期限: 收文日-1天-付款週期
                            'Modified by Lydia 2019/05/16 另外抓客戶的付款週期CU175
                            'strDate1 = CompDate(1, Val(pMemo) * -1, Format(CompDate(2, -1, pStaDay)))
                            'Modified by Lydia 2020/02/03 一般付款週期預設值維持2個月，惟計算方式以發文日當月的翌月1日起計
                            'strDate1 = CompDate(1, pMon02 * -1, Format(CompDate(2, -1, pStaDay)))
                            'If TransDate("" & rsAD.Fields("billdate"), 2) < strDate1 Then
                            'Memo by Lydia 2025/07/25 起算日若有變更，請一併修改PUB_ProcAcctmp08
                            strDate1 = Left(CompDate(1, 1, TransDate("" & rsAD.Fields("billdate"), 2)), 6) & "01" '發文日當月的翌月1日起計
                            strDate1 = CompDate(1, pMon02, strDate1)
                            If strDate1 < pStaDay Then
                            'end 2020/02/03
                                pBillNo = "" & rsAD.Fields("billno")
                                'Modified by Lydia 2020/02/03 發文日當月的翌月1日起計
                                'strA1 = TransDate("" & rsAD.Fields("billdate"), 2)
                                'Memo by Lydia 2025/07/25 起算日若有變更，請一併修改PUB_ProcAcctmp08
                                strA1 = Left(CompDate(1, 1, TransDate("" & rsAD.Fields("billdate"), 2)), 6) & "01"
                                
                                '記錄應收款的客戶編號
                                If "" & rsAD.Fields("cust1") <> "" And InStr(mRefCust, "" & rsAD.Fields("cust1")) > 0 Then
                                     pChkCust = "" & rsAD.Fields("cust1")
                                ElseIf "" & rsAD.Fields("cust2") <> "" And InStr(mRefCust, "" & rsAD.Fields("cust2")) > 0 Then
                                     pChkCust = "" & rsAD.Fields("cust2")
                                ElseIf "" & rsAD.Fields("cust3") <> "" And InStr(mRefCust, "" & rsAD.Fields("cust3")) > 0 Then
                                     pChkCust = "" & rsAD.Fields("cust3")
                                End If
                            End If
                      End If
                      pTotBill = pTotBill + Val("" & rsAD.Fields("A1"))
                      pPaTwFee = pPaTwFee + Val("" & rsAD.Fields("A2"))
                      pTmTwFee = pTmTwFee + Val("" & rsAD.Fields("A3"))
                      pTwFee = pTwFee + Val("" & rsAD.Fields("A4"))
                      If pKind = "2" Then Exit Do
                 End If
              End If  'Added by Lydia 2022/06/13 剔除案件ACS的曾有TIPS案件性質的案號的應收帳款
              rsAD.MoveNext
            Loop
            
            If pBillNo <> "" Then
               'X個月應收帳款
               pMon01 = DateDiff("m", ChangeWStringToWDateString(strA1), ChangeWStringToWDateString(strSrvDate(1)))
               If pMemo = "1" Then '一般付款週期
                   pMemo = "本客戶" & IIf(pChkCust <> "", "(" & pChkCust & ")", "") & "有(" & pBillNo & ")" & pMon01 & "個月應收帳款已逾付款週期"
               ElseIf pMemo = "2" Then '特殊付款週期
                   pMemo = "本客戶" & IIf(pChkCust <> "", "(" & pChkCust & ")", "") & "有(" & pBillNo & ")" & pMon01 & "個月應收帳款，已逾付款週期" & pMon02 & "個月"
               End If
               PUB_GetBillDataAll = True '有逾期的應收帳款
            Else
               pMemo = ""
            End If
        End If
   End If

   Set rsAD = Nothing
   Set rs1 = Nothing 'Added by Lydia 2019/05/16
End Function

'Added by Morgan 2018/10/25
'EMail客戶函
'Modified by Morgan 2021/12/2 +pToSales
Public Function PUB_SendECustLetter(pRecNo As String, Optional pReceiverNo As String, Optional pToSales As Boolean = False) As Boolean
   Dim lngMousePointer As Long
   Dim stSQL As String, intQ As Integer, intA As Integer
   Dim rsQuery As ADODB.Recordset
   Dim stProperty As String
   
   Dim stPath As String, stFiles As String
   Dim stFileName As String
   Dim arrFileName() As String
   Dim stFileNameList As String
   Dim idx As Integer
   
   Dim bDone As Boolean
   Dim stMailDate As String
   Dim stMailTime As String
   Dim stST15 As String
   Dim stName As String
   Dim stRecNo As String, stCP10 As String
   Dim stReceiverNo As String
   Dim stSales As String
   Dim stLP01 As String, stLP11 As String, stLP52 As String, bolRegMail As Boolean 'Added by Morgan 2022/2/18
      
   stST15 = Pub_StrUserSt15
   stName = strUserName
   
   lngMousePointer = Screen.MousePointer
   Screen.MousePointer = vbHourglass
   
On Error GoTo ErrHnd
   
   'Added by Morgan 2022/4/8
   'D類通知增加已結案及收文檢查
   If Left(pRecNo, 1) = "D" Then
      'Modified by Morgan 2024/10/29 副本信函要改用相關收文號判斷 Ex:P-124821視為撤回(C類)的副本信函
      'stSQL = "select np06,pa57||tm29 pa57,decode(nvl(pa09,tm10),'000',cpm03,cpm04) cp10N" & _
         " from caseprogress,nextprogress,patent,trademark,casepropertymap" & _
         " where cp09='" & pRecNo & "' and np01(+)=cp43 and np22(+)=cp30" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
         " and cpm01(+)=np02 and cpm02(+)=np07"
      stSQL = "select np06,pa57||tm29 pa57,decode(nvl(pa09,tm10),'000',cpm03,cpm04) cp10N" & _
         " from (select * from caseprogress where cp09='" & pRecNo & "' and cp10<>'990'" & _
         " union select * from caseprogress where cp09=(select cp43 from caseprogress where cp09='" & pRecNo & "' and cp10='990' and cp43 like 'D%')" & _
         ") X,nextprogress,patent,trademark,casepropertymap" & _
         " where np01(+)=cp43 and np22(+)=cp30" & _
         " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
         " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
         " and cpm01(+)=np02 and cpm02(+)=np07"
      'end 2024/10/29
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         If Not IsNull(rsQuery("pa57")) Then
            MsgBox "本案已閉卷，請確認！", vbExclamation
            GoTo ErrHnd
            
         ElseIf rsQuery("np06") = "N" Then
            MsgBox "本案[ " & rsQuery("cp10N") & " ]已不續辦，請確認！", vbExclamation
            GoTo ErrHnd
            
         ElseIf rsQuery("np06") = "Y" Then
            MsgBox "本案[ " & rsQuery("cp10N") & " ]已收文，請確認！", vbExclamation
            GoTo ErrHnd
            
         End If
      End If
   End If
   'end 2022/4/8
   
   'Modified by Morgan 2021/10/12 +cu186
   'Modified by Morgan 2022/2/18 +lp01,lp11,lp52
   'Modified by Morgan 2025/2/27 -cu186
   stSQL = "select cp01,cp02,cp03,cp04,cp10,cp43,cpp02,lp33,cu01||' '||cu04 Cust,cu176,lp01,lp11,lp52" & _
      " from caseprogress,casepaperpdf,letterprogress,customer" & _
      " where cp09='" & pRecNo & "'" & _
      " and cpp01(+)=cp09 and substr(lower(cpp02(+)),-8)='.cus.pdf' and cpp10<>'D'" & _
      " and lp01(+)=cp09 and cu01(+)=substr(lp33,1,8) and cu02(+)='0'"
      
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ <> 1 Then
      MsgBox "卷宗區沒有客戶函PDF檔！", vbCritical
      GoTo ErrHnd
   End If
   
   'Added by Morgan 2022/2/18
   '是否掛號直寄,若尚無信函進度者(CFT)則彈詢問視窗
   bolRegMail = False
   If pToSales = False Then
      stLP01 = "" & rsQuery("LP01")
      If IsNull(rsQuery("lp01")) Then
         intQ = MsgBox("請問本函是否原為掛號直寄信？", vbYesNoCancel + vbQuestion + vbDefaultButton3)
         If intQ = vbCancel Then
            GoTo ErrHnd
         End If
         
         If intQ = vbYes Then
            bolRegMail = True
            stLP11 = "Y"
            stLP52 = "Y"
         Else
            stLP11 = "1"
            stLP52 = ""
         End If
      Else
         stLP01 = "" & rsQuery("lp01")
         If rsQuery("lp11") = "Y" And rsQuery("lp52") = "Y" Then
            bolRegMail = True
         End If
      End If
   End If
   'end 2022/2/18
   
   '副本信函990 EMail主旨及內容要抓相關收文號
   stCP10 = rsQuery("cp10")
   If stCP10 = "990" Then
      stRecNo = rsQuery("cp43")
   Else
      stRecNo = pRecNo
   End If
   
   '有傳入收受人
   If pReceiverNo <> "" Then
      'Modified by Morgan 2021/10/27 +cu186
      'Modified by Morgan 2025/2/28 -cu186
      stSQL = "select cu01||' '||cu04 Cust,cu176" & _
         " from customer" & _
         " where cu01='" & Left(pReceiverNo, 8) & "' and cu02='0'"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
      If intQ = 1 Then
         If IsNull(rsQuery("cu176")) Then
            MsgBox "收受人【" & rsQuery.Fields("Cust") & "】未設指定信箱無法寄送！", vbCritical
            GoTo ErrHnd
         End If
         stReceiverNo = pReceiverNo
      Else
         MsgBox "無法讀取收受人【" & pReceiverNo & "】資料！", vbCritical
         GoTo ErrHnd
      End If
      
   '信函進度有收受人
   ElseIf Not IsNull(rsQuery("lp33")) Then
      If IsNull(rsQuery("cu176")) Then
         MsgBox "收受人【" & rsQuery.Fields("Cust") & "】未設指定信箱無法寄送！", vbCritical
         GoTo ErrHnd
      End If
      stReceiverNo = "" & rsQuery("lp33")
   End If
   
   'Added by Morgan 2021/10/12
   'Removed by Morgan 2025/2/27 欄位已改他用
   'If stReceiverNo <> "" And pToSales = False Then
   '   If Not IsNull(rsQuery("cu186")) Then
   '      MsgBox rsQuery("cu186"), vbExclamation, "e化客戶備註"
   '   End If
   'End If
   'end 2025/2/27
   'end 2021/10/12
   
   'Modified by Morgan 2021/10/12 +cu186
   'Modified by Morgan 2025/2/27 -cu186
   stSQL = "select cp01,cp02,cp03,cp04,cp10,decode(pa09,'000',cpm03,cpm04) prty,pa26,cu01||' '||cu04 Cust,cu176" & _
      " from (select cp01,cp02,cp03,cp04,cp09,cp10,nvl(pa09,nvl(sp09,tm10)) pa09,nvl(pa26,nvl(tm23,sp08)) pa26" & _
      " from caseprogress,patent,trademark,servicepractice" & _
      " where cp09='" & stRecNo & "'" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and tm01(+)=cp01 and tm02(+)=cp02 and tm03(+)=cp03 and tm04(+)=cp04" & _
      " and sp01(+)=cp01 and sp02(+)=cp02 and sp03(+)=cp03 and sp04(+)=cp04" & _
      ") X,casepropertymap,customer" & _
      " where cpm01(+)=cp01 and cpm02(+)=cp10" & _
      " and cu01(+)=substr(pa26,1,8) and cu02(+)=substr(pa26,9)"
      
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      '收受人為申請人
      If stReceiverNo = "" Then
         If IsNull(rsQuery.Fields("cu176")) Then
            MsgBox "【" & rsQuery.Fields("Cust") & "】未設指定信箱無法寄送！", vbCritical
            GoTo ErrHnd
         End If
         stReceiverNo = rsQuery("pa26")
         
         'Added by Morgan 2021/11/15
         'Removed by Morgan 2025/2/27 欄位已改他用
         'If Not IsNull(rsQuery("cu186")) Then
         '   MsgBox rsQuery("cu186"), vbExclamation, "全E化客戶備註"
         'End If
         'end 2025/2/27
         'end 2021/11/15
      End If
      
      With rsQuery
      'Modified by Morgan 2024/9/19 +收文號,因同案號同案件性質時檔名相同會導致不下載
      'stPath = App.path & "\" & strUserNum
      stPath = App.path & "\" & strUserNum & "\" & pRecNo
      'end 2024/9/19
      
      PUB_KillTempFolder strUserNum & "\" & pRecNo 'Added by Morgan 2024/12/11 '先刪除暫存資料夾,否則若檔案有更新會抓到舊的(檔名相同)
      
      If PUB_GetAttachFile4Cust(pRecNo, stFiles, stPath, , stCP10) = True Then
         arrFileName = Split(stFiles, ";")
         For idx = UBound(arrFileName) To LBound(arrFileName) Step -1
            If arrFileName(idx) <> "" Then
               stFileName = stPath & "\" & arrFileName(idx)
               stFileNameList = stFileName & ";" & stFileNameList
            End If
         Next
         
         stProperty = "" & .Fields("prty")
         stProperty = stProperty & PUB_GetRelateCasePropertyName(stRecNo, "1")
         stSales = PUB_GetAKindSalesNo(.Fields("cp01"), .Fields("cp02"), .Fields("cp03"), .Fields("cp04"))
         Screen.MousePointer = vbDefault
         
         If pToSales = True Then
            'Add By Sindy 2024/10/16 + bolReadLP42=True
            PUB_ShowMailForm pRecNo, stFileNameList, stProperty, bDone, , , , True, stMailDate, stMailTime, , , , , , , , , , , stSales, , True
            If bDone = True Then
               PUB_SendECustLetter = True
            End If
         Else
            strUserNum = stSales
            strUserName = GetStaffName(strUserNum, True)
            Pub_StrUserSt15 = PUB_GetStaffST15(strUserNum, "1")
            'Modified by Morgan 2022/2/18 +bolRegMail
            'Add By Sindy 2024/10/16 + bolReadLP42=True
            PUB_ShowMailForm pRecNo, stFileNameList, stProperty, bDone, , , , True, stMailDate, stMailTime, , , , , , , , , stReceiverNo, , , bolRegMail, True
            If bDone = True Then
               'Memo by Morgan 2022/2/8 此處會紀錄實際發信人員(全E化客戶是由程序寄發而寄信畫面是以智權身分寄送故不傳更新參數 strLP01)
               If stLP01 <> "" Then
                  strSql = "update letterprogress set lp39=" & stMailDate & ",lp40=" & stMailTime & ",lp38='" & strUser1Num & "' where lp01='" & stLP01 & "'"
                  cnnConnection.Execute strSql, intA
               Else
                  strSql = "insert into letterprogress(lp01,lp10,lp11,lp26,lp38,lp39,lp40,lp52) values('" & pRecNo & "','Y','" & stLP11 & "','E','" & strUser1Num & "'," & stMailDate & "," & stMailTime & ",'" & stLP52 & "')"
                  cnnConnection.Execute strSql, intA
                  
                  'Added by Morgan 2022/3/29 上判發(智權人員卷宗區才看得見),上確認(每日批次才不會列出)
                  stSQL = "update letterprogress set lp04='QPGMR',lp05=lp39,lp06='QPGMR',lp07=lp39 where lp01='" & pRecNo & "'"
                  cnnConnection.Execute stSQL, intA
                  
                  stSQL = "update caseprogress set cp127=to_char(sysdate,'YYYYMMDD'),cp128=to_char(sysdate,'HH24MISS') where cp09='" & pRecNo & "'"
                  cnnConnection.Execute stSQL, intA
                  If intA = 1 Then
                     'Trigger 會寫發文人故要另外更新
                     stSQL = "update caseprogress set cp154='QPGMR' where cp09='" & pRecNo & "'"
                     cnnConnection.Execute stSQL, intA
                  End If
                  'end 2022/3/28
               End If
               PUB_SendECustLetter = True
            End If
         End If
      End If
      End With
   End If
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
   
   Set rsQuery = Nothing
   Screen.MousePointer = lngMousePointer
   
   If strUserNum <> strUser1Num Then
      strUserNum = strUser1Num
      Pub_StrUserSt15 = stST15
      strUserName = stName
   End If
End Function

'Added by Morgan 2021/12/2
Public Function PUB_ChkEmailBackUp(pRecNo As String) As Boolean
   Dim stSQL As String, intQ As Integer
   Dim rsQuery As ADODB.Recordset
   
   PUB_ChkEmailBackUp = True
   stSQL = "select cpp02 from casepaperpdf where cpp01='" & pRecNo & "' and substr(lower(cpp02),-11)='.email.menu'"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      If MsgBox("卷宗區已有寄件備份，是否確定要寄通知信？", vbYesNo + vbDefaultButton2 + vbQuestion) = vbNo Then
         PUB_ChkEmailBackUp = False
      End If
   End If
   Set rsQuery = Nothing
End Function

'Added by Lydia 2019/03/22 (FCT)各式申請書-電子送件:申請人
'Modified by Lydia 2019/07/09 +系統別strTM01
'Modified by Lydia 2020/10/07 +案件性質strCP10
'Modified by Sindy 2021/7/6 + , Optional m_strNaName As String = "1" : 1.中文名稱優先,2.日文名稱優先
'Modified by Lydia 2023/12/29 +strJumpCust指定讀取申請人的資料
Public Function PUB_GetApplFCT_EData(ByVal ET01 As String, ByVal ET03 As String, ByVal strReceiveNo As String, ByVal strCP10 As String, _
   tm() As String, Optional bolReadTM As Boolean = True, Optional strApplNum As String, _
   Optional strRepresentative As String, Optional strTM01 As String = "FCT", Optional m_strNaName As String = "1", Optional strJumpCust As String) As Boolean
'bolReadTM: 判讀是否要抓個案資料
'strApplNum: 欲讀取的申請人資料
'strRepresentative : 欲讀取的代表人資料
   Dim strTxt(110) As String, strTmp As String
   Dim intR As Integer, intJ As Integer
   Dim strChaName As String, strEngName As String
   Dim intK As Integer, k_star As Integer, k_end As Integer, intRow As Integer
   Dim strApplEmp(1 To 5) As String, varTemp As Variant
   Dim strRepEmp(1 To 30) As String
   Dim idx As Integer
   'Added by Lydia 2020/10/07
   Dim strQ1 As String, rsQuery As New ADODB.Recordset, intQ As Integer
   Dim strPerType As String '(目的) 申請人x / 申請變更人x /
   Dim strOldPerType As String, strOldArray(1 To 5) As String   '(來源) 原申請人 ,編號
   Dim strProcArray(1 To 5)
   'end 2020/10/07
      
   intR = 0
   '有(變更、移轉、授權)申請人時,已傳入的資料讀取資料
   For intJ = 1 To 5
      strApplEmp(intJ) = ""
   Next intJ
   If Trim(strApplNum) <> "" Then
      varTemp = Split(strApplNum, "@")
      'Added by Lydia 2020/10/21 授權502分別設定：授權人、被授權人
      'Modified by Lydia 2021/02/05 +FCT案授權
      'If tm(1) = "T" And strCP10 = "502" Then '內商預設：基本檔之申請人=授權人、輸入之申請人=被授權人
      If (tm(1) = "T" Or tm(1) = "FCT") And strCP10 = "502" Then
          strApplEmp(1) = tm(23)
          For intJ = 2 To 5
             If tm(76 + intJ) <> "" Then
                strApplEmp(intJ) = tm(76 + intJ)
             End If
          Next intJ
          For intJ = 0 To UBound(varTemp) '輸入之申請人=被授權人
             strOldArray(intJ + 1) = Trim(varTemp(intJ))
          Next intJ
      Else
      'end 2020/10/21
          For intJ = 0 To UBound(varTemp)
             strApplEmp(intJ + 1) = Trim(varTemp(intJ))
          Next intJ
      End If 'Added by Lydia 2020/10/21
   Else
      strApplEmp(1) = tm(23)
      For intJ = 2 To 5
         If tm(76 + intJ) <> "" Then
            strApplEmp(intJ) = tm(76 + intJ)
         End If
      Next intJ
   End If
   'Added by Lydia 2020/10/07 (來源) 原申請人編號
   If Trim(strApplNum) <> "" And strCP10 = "301" And tm(15) = "" Then  '註冊前變更
       strOldArray(1) = tm(23)
       For intJ = 2 To 5
          If tm(76 + intJ) <> "" Then
             strOldArray(intJ) = tm(76 + intJ)
          End If
       Next intJ
   End If
   'end 2020/10/07
   
   '變更代表人
   For intJ = 1 To 30
      strRepEmp(intJ) = ""
   Next intJ
   If Trim(strRepresentative) <> "" Then
      varTemp = Split(strRepresentative, "@")
      For intJ = 0 To UBound(varTemp)
         strRepEmp(intJ + 1) = Trim(varTemp(intJ))
      Next intJ
   End If
   'Added by Lydia 2020/10/07
   If strCP10 = "301" And tm(15) = "" Then '註冊前變更
      strPerType = "申請變更人"
      strOldPerType = "原申請人"
   'Added by Lydia 2020/10/21
   ElseIf strCP10 = "502" Then '授權
      strPerType = "授權人"
      strOldPerType = "被授權人"
   Else
      strPerType = "申請人"
   End If
   'end 2020/10/07
   
JumpToData2:   'Added by Lydia 2020/09/20 原申請人的資料抓法與申請變更人一致
   '申請人
   For intJ = 1 To 5
      If strApplEmp(intJ) <> "" Then
         '申請人
         strQ1 = " SELECT C.*,N1.NA72 X1,N2.NA72 X2" & _
            " FROM CUSTOMER C,NATION N1,NATION N2 WHERE CU01='" & Left(ChangeCustomerL(strApplEmp(intJ)), 8) & "'" & _
            " and cu02='" & Mid(ChangeCustomerL(strApplEmp(intJ)), 9) & "' AND N1.NA01(+)=CU10 AND N2.NA01(+)=CU87"
         intQ = 1
         Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
         If intQ = 1 Then
            'Added by Lydia 2020/10/07 因為只顯示原申請人的名稱/姓名,所以增加顯示判斷
            If strPerType = "原申請人" Then
                intR = intR + 1
                strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                   " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-NO','♀')"
            End If
            'end 2020/10/08
            intR = intR + 1
            strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-國籍','" & rsQuery("X1") & "')"
            
            intR = intR + 1
            If "" & rsQuery("CU15") = "0" Then
               strTmp = "自然人"
            Else
               strTmp = "法人公司機關學校"
            End If
            strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-身分種類','" & strTmp & "')"
                  
            intR = intR + 1
            If "" & rsQuery("CU10") < "011" Then
               If "" & rsQuery("CU15") = "0" And "" & rsQuery("CU11") = "" Then '個人無ID時也要顯示標題
                  strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-ID','♀')"
               Else
                  strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-ID','" & rsQuery("CU11") & "')"
               End If
            End If
            
            'Modify By Sindy 2021/7/6 優先讀取日文名稱 FCT-01-717-03
            If m_strNaName = "2" Then
               intR = intR + 1
               strTmp = strPerType & intJ & "-中文名稱"
               If "" & rsQuery("CU06") <> "" Then
                  strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL("" & rsQuery("CU06")) & "')"
               Else
                  strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL("" & rsQuery("CU04")) & "')"
               End If
            Else
            '2021/7/6 END
               intR = intR + 1
               If "" & rsQuery("CU15") = "0" Then
                  strTmp = strPerType & intJ & "-中文姓名"
               Else
                  strTmp = strPerType & intJ & "-中文名稱"
               End If
               ' 修法:106/12/01開始中文名稱要加外商國名
               'Modified by Lydia 2024/03/05 改成模組 PUB_GetApplT_CNAME
               'If Val(strSrvDate(2)) >= 1061201 And "" & rsQuery("CU15") = "1" Then '1.公司
               '   'Added by Lydia 2019/07/10 +判斷商標台灣案不用顯示X商
               '   If "" & rsQuery("CU10") <= "010" Then
               '        strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '           " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL("" & rsQuery("CU04")) & "')"
               '   'Added by Lydia 2019/07/10 智慧局的商標申請須知有提到用大陸地區,與專利用的大陸商不同(與嘉雯和阿蓮,還有外專程序確認過)
               '   'Remove by Lydia 2019/08/29 大陸改成大陸商
               '   'ElseIf "" & rsQuery("CU10") = "020" Then
               '   '     'Modified by Lydia 2019/08/05 ８月份有收到智慧局行文於七月討論結果，決定外商後不再加點或空格，並於８月１日開始施行。
               '   '     'strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '   '        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','大陸地區．" & ChgSQL("" & rsQuery("CU04")) & "')"
               '   '     strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '   '        " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','大陸地區" & ChgSQL("" & rsQuery("CU04")) & "')"
               '   'end 2019/08/29
               '   Else
               '   'Modified by Lydia 2019/02/21 阿蓮提出Ｘ商‧名稱
               '        'Modified by Lydia 2019/08/05 ８月份有收到智慧局行文於七月討論結果，決定外商後不再加點或空格，並於８月１日開始施行。
               '        'strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '           " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & GetPrjNationName("" & rsQuery("CU10"), "NA81") & "‧" & ChgSQL("" & rsQuery("CU04")) & "')"
               '        strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '           " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & GetPrjNationName("" & rsQuery("CU10"), "NA81") & ChgSQL("" & rsQuery("CU04")) & "')"
               '   End If
               'Else
               '   '柏翰提個人的姓和名中間要有,號
               '   If "" & rsQuery("CU15") = "0" Then '自然人
               '      'Modified by Lydia 2022/08/24 外商電子送件申請之自然人名稱前加國籍
               '      'strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & PUB_ConvertNameFormat(ChgSQL("" & rsQuery("CU04"))) & "')"
               '      strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & Replace(GetPrjNationName("" & rsQuery("CU10"), "NA81"), "商", "籍") & PUB_ConvertNameFormat(ChgSQL("" & rsQuery("CU04"))) & "')"
               '   Else
               '      strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
               '         " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL("" & rsQuery("CU04")) & "')"
               '   End If
               'End If
               strQ1 = PUB_GetApplT_CNAME("" & rsQuery.Fields("CU01") & rsQuery.Fields("CU02"))
               strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL(strQ1) & "')"
               'end 2024/03/05
            End If
            
            'Modify By Sindy 2021/7/6 優先讀取日文名稱 FCT-01-717-03
            If m_strNaName = "2" Then
               intR = intR + 1
               strTmp = strPerType & intJ & "-英文名稱"
            Else
            '2021/7/6 END
               intR = intR + 1
               If "" & rsQuery("CU15") = "0" Then
                  strTmp = strPerType & intJ & "-英文姓名"
               Else
                  strTmp = strPerType & intJ & "-英文名稱"
               End If
            End If
            'Added by Lydia 2022/02/08 遇見申請英文證明之申請人基本檔沒英文就塞空白 --- by 嘉雯
            'Modified by Lydia 2023/12/21 加上CFT申請英文證明之申請人基本檔沒英文就塞空白 --- by 阿蓮
            If (strTM01 = "T" Or strTM01 = "CFT") And strCP10 = "304" And Trim(RTrim(Trim("" & rsQuery("CU05")) & " " & Trim("" & rsQuery("CU88")) & " " & Trim("" & rsQuery("CU89")) & " " & Trim("" & rsQuery("CU90")))) = "" Then
               strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','♀')"
            Else
            'end 2022/02/08
               strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strTmp & "','" & ChgSQL(RTrim(Trim("" & rsQuery("CU05")) & " " & Trim("" & rsQuery("CU88")) & " " & Trim("" & rsQuery("CU89")) & " " & Trim("" & rsQuery("CU90")))) & "')"
            End If 'Added by Lydia 2022/02/08
            
            '目前抓客戶基本檔資料,等基本檔加欄位後需改抓
            'Modify By Sindy 2019/5/20
            If Left(Pub_StrUserSt03, 1) = "F" Then '外專抓取地址國籍
               intR = intR + 1
               strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-居住國','" & rsQuery("X2") & "')"
            Else
               intR = intR + 1
               strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-居住國','" & rsQuery("X1") & "')"
            End If
            
            '抓個案地址
            'Modified by Lydia 2023/12/23 FCT案之變更和延展申請書：當勾選「變更地址」，基本資料表之申請人地址改抓申請人基本檔之地址。
            'If bolReadTM = True Then
            If bolReadTM = True And InStr(strJumpCust & ",", "申請人地址") = 0 Then
               intR = intR + 1
               strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-中文地址','" & PUB_ChgNumeralStyle(ChgSQL(IIf(intJ = 1, tm(24), tm(80 + intJ)))) & "')"
               intR = intR + 1
               strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-英文地址','" & ChgSQL(IIf(intJ = 1, tm(25), tm(84 + intJ))) & "')"
            Else
               intR = intR + 1
               If "" & rsQuery("CU10") < "011" Then
                  strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-郵遞區號','" & PUB_ChgNumeralStyle("" & rsQuery("CU112")) & "')"
               End If
               intR = intR + 1
               strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-中文地址','" & PUB_ChgNumeralStyle(ChgSQL("" & rsQuery("CU23"))) & "')"
               intR = intR + 1
               'Added by Lydia 2022/02/08 遇見申請英文證明之申請人基本檔沒英文就塞空白 --- by 嘉雯
               'Modified by Lydia 2023/12/21 加上CFT申請英文證明之申請人基本檔沒英文就塞空白 --- by 阿蓮
               If (strTM01 = "T" Or strTM01 = "CFT") And strCP10 = "304" And Trim(RTrim(Trim("" & rsQuery("CU24")) & " " & Trim("" & rsQuery("CU25")) & " " & Trim("" & rsQuery("CU26")) & " " & Trim("" & rsQuery("CU27")) & " " & Trim("" & rsQuery("CU28")) & " " & Trim("" & rsQuery("CU102")))) = "" Then
                    strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-英文地址','♀')"
               Else
               'end 2022/02/08
                    strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                       " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-英文地址','" & ChgSQL(RTrim(Trim("" & rsQuery("CU24")) & " " & Trim("" & rsQuery("CU25")) & " " & Trim("" & rsQuery("CU26")) & " " & Trim("" & rsQuery("CU27")) & " " & Trim("" & rsQuery("CU28")) & " " & Trim("" & rsQuery("CU102")))) & "')"
               End If 'Added by Lydia 2022/02/08
            End If
            
            '依申請人帶出代表人資料
            'Modified by Lydia 2023/12/29 刪除CFT英文證明申請書基本資料之【代表人英文姓名】標題 ; 已向智慧局確認 , 代表人為本國人可不輸英文姓名 ---阿蓮(避免轉檔出錯，不過有沒有代表人都不帶出)
            'If "" & rsQuery("CU15") <> "0" Then '非自然人才要帶出代表人資料
            'Modified by Lydia 2024/02/05 CFT英文證明申請書：無論是否有代表人中文名稱都要帶標題，代表人英文姓名可省略
            'If "" & rsQuery("CU15") <> "0" And strTM01 <> "CFT" Then
            If "" & rsQuery("CU15") <> "0" Then
               strChaName = "": strEngName = ""
               '變更代表人(1~30)
               If Trim(strRepresentative) <> "" Then
                  intRow = 0
                  For intK = 1 To 2
                     intRow = intRow + 1
                     If intK = 1 Then idx = intJ + (intJ - 1) * 5 '1,7,13,19,25
                     If intK = 2 Then idx = idx + 3 '4,10,16,22,28
                     '代表人中文姓名-->非自然人時為必要欄位
                     strTmp = strRepEmp(idx)
                     If strTmp <> "" Then
                        'Remove by Lydia 2019/08/15 嘉雯:代表人中文姓名中間加逗號,造成轉檔錯誤
                        'If Len(strTmp) = 3 Then strTmp = PUB_ConvertNameFormat(strTmp)
                        strChaName = strChaName & " " & intRow & "." & strTmp
                     Else
                        '只有一個代表人時不要有1.
                        'Modified by Lydia 2019/08/08 判斷沒有2.
                        If strChaName <> "" And InStr(strChaName, "2.") = 0 Then
                           strChaName = Replace(strChaName, "1.", "")
                        End If
                     End If
                     '代表人英文姓名-->非必要欄位
                     strTmp = strRepEmp(idx + 1)
                     If strTmp <> "" Then
                        strEngName = strEngName & " " & intRow & "." & strTmp
                     Else
                        '只有一個代表人時不要有1.
                        'Modified by Lydia 2019/08/08 判斷沒有2.
                        If strEngName <> "" And InStr(strEngName, "2.") = 0 Then
                           strEngName = Replace(strEngName, "1.", "")
                        End If
                     End If
                  Next intK
                  
               ElseIf bolReadTM = True Then '抓個案資料
                  If intJ < 3 Then
                     k_star = 1: k_end = 2
                  ElseIf intJ = 3 Then
                     k_star = 3: k_end = 4
                  ElseIf intJ = 4 Then
                     k_star = 5: k_end = 6
                  ElseIf intJ = 5 Then
                     k_star = 7: k_end = 8
                  End If
                  intRow = 0
                  For intK = k_star To k_end
                     intRow = intRow + 1
                     '代表人中文姓名-->非自然人時為必要欄位
                     If intJ = 1 Then
                        strTmp = tm(47 + 3 * (intK - 1))
                     Else
                        strTmp = tm(94 + 3 * (intK - 1))
                     End If
                     If strTmp <> "" Then
                        strChaName = strChaName & " " & intRow & "." & strTmp
                     Else
                        '只有一個代表人時不要有1.
                        'Modified by Lydia 2019/08/08 判斷沒有2.
                        If strChaName <> "" And InStr(strChaName, "2.") = 0 Then
                           strChaName = Replace(strChaName, "1.", "")
                        End If
                     End If
                     '代表人英文姓名-->非必要欄位
                     If intJ = 1 Then
                        strTmp = tm(48 + 3 * (intK - 1))
                     Else
                        strTmp = tm(95 + 3 * (intK - 1))
                     End If
                     If strTmp <> "" Then
                        strEngName = strEngName & " " & intRow & "." & strTmp
                     Else
                        '只有一個代表人時不要有1.
                        'Modified by Lydia 2019/08/08 判斷沒有2.
                        If strEngName <> "" And InStr(strEngName, "2.") = 0 Then
                           strEngName = Replace(strEngName, "1.", "")
                        End If
                     End If
                  Next intK
                  
               Else '抓客戶檔資料
                  intRow = 0
                  For intK = 1 To 6
                     intRow = intRow + 1
                     '代表人中文姓名-->非自然人時為必要欄位
                     strTmp = "" & rsQuery("CU" & CStr(39 + 3 * (intK - 1)))
                     If strTmp <> "" Then
                        'Remove by Lydia 2019/08/15 嘉雯:代表人中文姓名中間加逗號,造成轉檔錯誤
                        'If Len(strTmp) = 3 Then strTmp = PUB_ConvertNameFormat(strTmp)
                        strChaName = strChaName & " " & intRow & "." & strTmp
                     Else
                        '只有一個代表人時不要有1.
                        'Modified by Lydia 2019/08/08 判斷沒有2.
                        If strChaName <> "" And InStr(strChaName, "2.") = 0 Then
                           strChaName = Replace(strChaName, "1.", "")
                        End If
                     End If
                     '代表人英文姓名-->非必要欄位
                     strTmp = "" & rsQuery("CU" & CStr(40 + 3 * (intK - 1)))
                     If strTmp <> "" Then
                        strEngName = strEngName & " " & intRow & "." & strTmp
                     Else
                        '只有一個代表人時不要有1.
                        'Modified by Lydia 2019/08/08 判斷沒有2.
                        If strEngName <> "" And InStr(strEngName, "2.") = 0 Then
                           strEngName = Replace(strEngName, "1.", "")
                        End If
                     End If
                  Next intK
               End If
               '代表人中文姓名
               If strChaName = "" Then
                  'Modified by Lydia 2020/02/05 FCT: 無資料則帶出（容後補呈）
                  'If Left(Pub_StrUserSt03, 2) = "F2" Then '外專要帶後補2個字
                  '   strChaName = "後補"
                  If tm(1) = "FCT" Then
                     strChaName = "（容後補呈）"
                  Else
                     strChaName = "♀"
                  End If
               Else
                  strChaName = Mid(strChaName, 2)
               End If
               intR = intR + 1
               strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                  " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-代表人中文姓名','" & ChgSQL(strChaName) & "')"
               
               'Added by Lydia 2020/02/05 FCT: 無資料則帶出（容後補呈）
               If strEngName = "" Then
                   If tm(1) = "FCT" Then
                       strEngName = "（容後補呈）"
                   Else
                       strEngName = "♀"
                   End If
               Else
                   strEngName = Mid(strEngName, 2)
               End If
               'end 2020/02/05
               'Added by Lydia 2024/02/05 CFT英文證明申請書：無論是否有代表人中文名稱都要帶標題，代表人英文姓名可省略
               If strTM01 = "CFT" And strCP10 = "304" Then
                  strEngName = ""
               End If
               'end 2024/02/05
               
               '代表人英文姓名
               If strEngName <> "" Then
                  'strEngName = Mid(strEngName, 2) 'Mark by Lydia 2020/02/05
                  intR = intR + 1
                  strTxt(intR) = "INSERT INTO EXCEPTCONDITION (ET01,ET02,ET03,ET04,ET05,ET06) " & _
                     " VALUES ('" & ET01 & "','" & strReceiveNo & "','" & ET03 & "','" & strUserNum & "','" & strPerType & intJ & "-代表人英文姓名','" & ChgSQL(strEngName) & "')"
               End If
            End If
         End If
      End If
   Next intJ
   
   'Added by Lydia 2020/10/07 原申請人的資料抓法與申請變更人一致
   If strOldArray(1) <> "" And strApplEmp(1) <> strOldArray(1) Then
      For intJ = 1 To 5
          strApplEmp(intJ) = strOldArray(intJ)
      Next intJ
      strPerType = strOldPerType
      GoTo JumpToData2
   End If
   'end 2020/10/07
   
   Set rsQuery = Nothing 'Added by Lydia 2020/10/07
   
   If Not ClsLawExecSQL(intR, strTxt) Then
      MsgBox "儲存例外欄位失敗，請洽系統管理員 !", vbCritical
   Else
      PUB_GetApplFCT_EData = True
   End If
End Function

'Added by Lydia 2019/04/17  FCP外專會稿承辦單列印(下載Word樣本檔，套印)
'Modified by Lydia 2020/04/10 +是否產生Word檔 bolSaveFile
'Modify By Sindy 2023/9/14 + , ByRef strColName() As String, ByRef strColText() As String : 回傳變數值
'                          + , Optional ByVal bolOnlyGetVal As Boolean = False : 是否純取值
'                          + , Optional ByRef intColCnt As Integer = 0 : 幾個欄位值
Public Function Pub_PrintFCP924Form(m_strPA01 As String, m_strPA02 As String, m_strPA03 As String, m_strPA04 As String, _
   m_strCP09 As String, ByRef strColName() As String, ByRef strColText() As String, _
   Optional ByVal bolSaveFile As Boolean = False, Optional ByVal bolOnlyGetVal As Boolean = False, _
   Optional ByRef intColCnt As Integer = 0) As Boolean
Dim LenStar As Long
Dim i As Integer
Dim DrawCount As Integer
Dim intFixedHigh As Integer
Dim intR As Integer
Dim rsA As New ADODB.Recordset
Dim m_CaseName As String '案件名稱
Dim m_strCP10 As String, m_strCP10n As String '案件性質
Dim m_PA26n As String '申請人1
Dim m_FCNa16 As String 'FCP管制人
Dim m_FCna01 As String, m_FCNa03 As String 'FC代理人之國籍
Dim m_strCP14 As String, m_CP14t As String  '翻譯人員
Dim m_924CP14n As String 'Added by Lydia 2020/04/14 會稿工程師(承辦人)
Dim m_eMail As String '是否E化
Dim m_CP64 As String '案件備註
Dim strSpeed As String '速別
Dim strGetPaper As String '受文者
Dim strMainText As String '主旨
Dim strAppend As String '附件
Dim strNote As String '備註
Dim strAfter As String '發文後
Dim m_PA75 As String 'FC代理人
Dim m_PA26 As String, m_PA27 As String, m_PA28 As String, m_PA29 As String, m_PA30 As String '申請人1~5
Dim m_924CP09 As String '會稿收文號
Dim m_924CP06 As String, m_924CP07 As String '會稿所限和法限
Dim m_EP31t As String
'通知函承辦單備註設定
Dim bolGetSpec As Boolean
Dim gMemo As String, gSpeed As String, gMain As String, gAppend As String, gOther As String '備註內容、速別、主旨、附件、其他
'使用樣本檔
Dim bVisible As Boolean
Dim iCall As Integer '樣本的變數
Dim m_FileName As String
'Dim strName As String, strText As String
Dim m_WordLeft As Long, m_WordTop As Long 'Word開啟位置
Dim m_FilePath As String 'Added by Lydia 2020/04/10 Word完整路徑
   
   Pub_PrintFCP924Form = False 'Add By Sindy 2023/9/15
   '粗線深度
'   DrawCount = 20
'   intFixedHigh = 400 '行高
   
   'Modified by Lydia 2020/04/14 +會稿工程師(承辦人)
   'iCall = 9
   iCall = 10
   intColCnt = iCall 'Add By Sindy 2023/9/19
   
   'Add By Sindy 2023/9/14
   ReDim Preserve strColName(iCall) As String
   ReDim Preserve strColText(iCall) As String
   '2023/9/14 END
   
   'Modifed by Lydia 2019/09/23
   'strSql = "select c1.cp09 as c1cp09,c1.cp10 as c1cp10,c1.cp64,c1.cp159,nvl(pa05,nvl(pa06,pa07)) casename,fa01||fa02 as fano,na01,na03,nvl(fa05,nvl(fa04,fa06)) as faname, " & _
               "pa26,nvl(cu04,nvl(cu05,cu06)) pa26n,pa27,pa28,pa29,pa30, " & _
               "GetEmailFlag(PA01||PA02||PA03||PA04) eMail,s1.st02 as na16n,c1.cp14,s2.st02 as cp14n,c2.cp09 as c2cp09,c2.cp06,c2.cp07,ep31 " & _
               "from caseprogress c1,patent,fagent,customer,nation,staff s1,staff s2,caseprogress c2,engineerprogress " & _
               "where c1.cp09='" & m_strCP09 & "' and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+) " & _
               "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
               "and fa10=na01(+) and na16=s1.st01(+) and c1.cp14=s2.st01(+) " & _
               "and c1.cp01=c2.cp01(+) and c1.cp02=c2.cp02(+) and c1.cp03=c2.cp03(+) and c1.cp04=c2.cp04(+) " & _
               "and c2.cp09 in (select substr(min(cp05||cp09),9,9) mno from caseprogress where cp01='" & m_strPA01 & "' and cp02='" & m_strPA02 & "' and cp03='" & m_strPA03 & "' and cp04='" & m_strPA04 & "' and cp10='924' and cp159=0) " & _
               "and c1.cp09=ep02(+) "
   If m_strCP09 <> "" Then '傳入會稿收文號
       'Modified by Lydia 2020/04/10 拿掉engineerprogress
        strSql = "select c1.cp09 as c1cp09,c1.cp10 as c1cp10,c1.cp64,c1.cp159,nvl(pa05,nvl(pa06,pa07)) casename,fa01||fa02 as fano,na01,na03,nvl(fa05,nvl(fa04,fa06)) as faname, " & _
                    "pa26,nvl(cu04,nvl(cu05,cu06)) pa26n,pa27,pa28,pa29,pa30, " & _
                    "GetEmailFlag(PA01||PA02||PA03||PA04) eMail,s1.st02 as na16n,c1.cp14,s2.st02 as cp14n,c1.cp06,c1.cp07 " & _
                    "from caseprogress c1,patent,fagent,customer,nation,staff s1,staff s2 " & _
                    "where c1.cp09='" & m_strCP09 & "' and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+) " & _
                    "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
                    "and fa10=na01(+) and na16=s1.st01(+) and c1.cp14=s2.st01(+) "
   Else  '用案號尋找
        'Modified by Lydia 2020/04/10 拿掉engineerprogress
        strSql = "select c1.cp09 as c1cp09,c1.cp10 as c1cp10,c1.cp64,c1.cp159,nvl(pa05,nvl(pa06,pa07)) casename,fa01||fa02 as fano,na01,na03,nvl(fa05,nvl(fa04,fa06)) as faname, " & _
                    "pa26,nvl(cu04,nvl(cu05,cu06)) pa26n,pa27,pa28,pa29,pa30, " & _
                    "GetEmailFlag(PA01||PA02||PA03||PA04) eMail,s1.st02 as na16n,c1.cp14,s2.st02 as cp14n,c1.cp06,c1.cp07 " & _
                    "from caseprogress c1,patent,fagent,customer,nation,staff s1,staff s2 " & _
                    "where c1.cp09 in (select substr(min(cp05||cp09),9,9) mno from caseprogress where cp01='" & m_strPA01 & "' and cp02='" & m_strPA02 & "' and cp03='" & m_strPA03 & "' and cp04='" & m_strPA04 & "' and cp10='924' and cp159=0) " & _
                    "and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+) " & _
                    "and substr(pa75,1,8)=fa01(+) and substr(pa75,9,1)=fa02(+) and substr(pa26,1,8)=cu01(+) and substr(pa26,9,1)=cu02(+) " & _
                    "and fa10=na01(+) and na16=s1.st01(+) and c1.cp14=s2.st01(+) "
   End If
   'end 2019/09/23
   intR = 1
   Set rsA = ClsLawReadRstMsg(intR, strSql)
   If intR = 1 Then
        If Val("" & rsA.Fields("cp159")) > 0 Then
            'Add By Sindy 2023/9/14
            If bolOnlyGetVal = True Then
               MsgBox "已取消收文，不可執行承辦歷程！"
               Exit Function
            Else
            '2023/9/14 END
               MsgBox "已取消收文，不可列印！"
               Exit Function
            End If
        End If
        m_CaseName = "" & rsA.Fields("casename")
        m_eMail = "" & rsA.Fields("eMail") '是否Ｅ化
        strGetPaper = "" & Trim(rsA.Fields("na03")) & " " & rsA.Fields("faname") '受文者(FC代理人之國別+英文名稱)
        strMainText = "會稿Claims" '主旨
        m_924CP06 = TransDate("" & rsA.Fields("cp06"), 1) '會稿所限
        m_924CP07 = TransDate("" & rsA.Fields("cp07"), 1) '會稿法限
        'Modified by Lydia 2020/04/14 會稿工程師(承辦人)
        'm_strCP14 = "" & rsA.Fields("cp14")
        'm_CP14t = "" & rsA.Fields("cp14n") '譯者名稱
        m_924CP14n = "" & rsA.Fields("cp14n")
        'end 2020/04/14
        m_CP64 = "" & rsA.Fields("cp64") '進度備註
        m_FCna01 = Left(Trim("" & rsA.Fields("na01")), 3)
        m_FCNa03 = Trim("" & rsA.Fields("na03"))
        m_FCNa16 = "" & rsA.Fields("na16n") 'FCP管制人
        'Modified by Lydia 2019/09/23
        'm_924CP09 = "" & rsA.Fields("c2cp09") '會稿收文號
        m_924CP09 = "" & rsA.Fields("c1cp09")
        'm_EP31t = TransDate("" & rsA.Fields("ep31"), 1)  'Claims完稿日 'Mark by Lydia 2020/04/10 另外抓
        m_PA75 = "" & rsA.Fields("fano") 'FC代理人
        m_PA26 = "" & rsA.Fields("PA26") '第1申請人
        m_PA26n = "" & rsA.Fields("PA26n")
        m_PA27 = "" & rsA.Fields("PA27") '第2申請人
        m_PA28 = "" & rsA.Fields("PA28") '第3申請人
        m_PA29 = "" & rsA.Fields("PA29") '第4申請人
        m_PA30 = "" & rsA.Fields("PA30") '第5申請人
   End If
   
   'Added by Lydia 2020/04/10 因為Claims完稿日是記錄在201中說的ep31,所以要另外抓; ex.FCP-062915
   'Modified by Lydia 2020/04/14 抓譯者
   'strSql = "select ep02,ep31 from Caseprogress,engineerprogress " & _
                "where cp01='" & m_strPA01 & "' and cp02='" & m_strPA02 & "' and cp03='" & m_strPA03 & "' and cp04='" & m_strPA04 & "' and cp10='201' and cp159=0 and cp09=ep02(+) "
   strSql = "select cp14, st02 as cp14n, ep02,ep31 from Caseprogress,engineerprogress,staff " & _
                "where cp01='" & m_strPA01 & "' and cp02='" & m_strPA02 & "' and cp03='" & m_strPA03 & "' and cp04='" & m_strPA04 & "' and cp10='201' and cp159=0 and cp09=ep02(+) and cp14=st01(+)"
   intR = 1
   Set rsA = ClsLawReadRstMsg(intR, strSql)
   If intR = 1 Then
       m_EP31t = TransDate("" & rsA.Fields("ep31"), 1)
       'Added by Lydia 2020/04/14 譯者
       m_strCP14 = "" & rsA.Fields("cp14")
       m_CP14t = "" & rsA.Fields("cp14n")
       'end 2020/04/14
   End If
   'end 2020/04/10
   
   '通知函承辦單備註設定(比照核准函01)
   bolGetSpec = PUB_GetFcpEMPBillSpec(m_strPA01 & m_strPA02 & m_strPA03 & m_strPA04, "01", m_PA75, m_PA26 & "," & m_PA27 & "," & m_PA28 & "," & m_PA29 & "," & m_PA30, , gMemo, gSpeed, gMain, gAppend, gOther)
    '速別
    If m_eMail = "E" Then 'E+寄
       If m_FCna01 = "101" Then '美國
          strSpeed = "Email+掛號"
       Else
          strSpeed = "Email+限時"
       End If
    ElseIf m_eMail = "e" Then 'E化
       strSpeed = "Email"
    Else '非E化
       If m_FCna01 = "101" Then '美國
          strSpeed = "Fax+掛號"
       Else
          strSpeed = "Fax+限時"
       End If
    End If
    If bolGetSpec = True Then '有通知函承辦單備註設定
         If gSpeed <> "" Then strSpeed = gSpeed
    End If

   '會稿所限
   If Trim(m_924CP06) = "" Then
        m_924CP06 = "　　年　　月　　日"
   Else
        m_924CP06 = Mid(m_924CP06, 1, 3) & " 年 " & Mid(m_924CP06, 4, 2) & " 月 " & Mid(m_924CP06, 6, 2) & " 日 "
   End If
   '會稿法限
   If Trim(m_924CP07) = "" Then
        m_924CP07 = "　　年　　月　　日"
   Else
        m_924CP07 = Mid(m_924CP07, 1, 3) & " 年 " & Mid(m_924CP07, 4, 2) & " 月 " & Mid(m_924CP07, 6, 2) & " 日 "
   End If
   
   'Claims完稿日: 有輸入完稿日表示只會稿Claims, 有收文會稿但沒有輸入完稿日表示會稿說明書(整本)
   strMainText = "會稿"
   If Trim(m_EP31t) = "" Then
        m_EP31t = "　　年　　月　　日"
        strMainText = "會稿說明書"
   Else
        m_EP31t = Mid(m_EP31t, 1, 3) & " 年 " & Mid(m_EP31t, 4, 2) & " 月 " & Mid(m_EP31t, 6, 2) & " 日 "
        strMainText = "會稿Claims"
   End If
   
   '翻譯人員（譯者）
   'Modified by Lydia 2025/03/13 改用模組取得
   'If m_CP14t <> "" And InStr(外翻_舜禹 & "," & 外翻_捷恩凱 & "," & 外翻_迅達, m_strCP14) = 0 Then
   If m_CP14t <> "" And InStr(Pub_SetF51Order("F", ""), m_strCP14) = 0 Then
      If Left(m_strCP14, 1) = "F" Then
          m_CP14t = m_CP14t & "-下班"
      Else
          m_CP14t = m_CP14t & "-上班"
      End If
   End If
   
   'FCP程序管制人
   'Modify By Sindy 2023/9/19
   If bolOnlyGetVal = True Then
      m_FCNa16 = PUB_GetFCPHandler(m_strPA01, m_strPA02, m_strPA03, m_strPA04)
   Else
   '2023/9/19 END
      m_FCNa16 = GetStaffName(PUB_GetFCPHandler(m_strPA01, m_strPA02, m_strPA03, m_strPA04), True)
   End If
    
   '備註
   strNote = "檔案已存於Typing2\外專送件" & IIf(strMainText = "會稿Claims", "\中說原始檔", "") & vbCrLf
   'Modify By Sindy 2023/9/19
   If bolOnlyGetVal = False Then
   '2023/9/19 END
      strNote = strNote & convForm("申請人1：" & m_PA26n, 70) & vbCrLf
      strNote = strNote & "案件名稱：" & m_CaseName & vbCrLf
   End If
   'Added by Lydia 2019/08/27 Sharon: +會稿的進度備註
   If m_CP64 <> "" Then
      strNote = strNote & "進度備註：" & PUB_StringFilter(m_CP64) & vbCrLf
   End If
   
   'Add By Sindy 2023/9/14 讀取變數值
   For intR = 0 To iCall
      strColName(intR) = ""
      strColText(intR) = ""
      If intR = 0 Then '速別
          strColName(intR) = "速別"
          strColText(intR) = strSpeed
      ElseIf intR = 1 Then '受文者
          strColName(intR) = "受文者"
          strColText(intR) = strGetPaper
      ElseIf intR = 2 Then '本所案號
          strColName(intR) = "本所案號"
          strColText(intR) = m_strPA01 & "-" & m_strPA02 & "-" & m_strPA03 & "-" & m_strPA04
      ElseIf intR = 3 Then '主旨
          strColName(intR) = "主旨"
          strColText(intR) = strMainText
      ElseIf intR = 4 Then '會稿所限
          strColName(intR) = "會稿所限"
          strColText(intR) = m_924CP06
      ElseIf intR = 5 Then '譯者
          strColName(intR) = "譯者"
          'Modify By Sindy 2023/9/19
          If bolOnlyGetVal = True Then
             If InStr(m_CP14t, "-") > 0 Then
                strColText(intR) = m_CP14t
             Else
                strColText(intR) = m_strCP14
             End If
          Else
          '2023/9/19 END
             strColText(intR) = m_CP14t
          End If
      ElseIf intR = 6 Then '會稿法限
          strColName(intR) = "會稿法限"
          strColText(intR) = m_924CP07
      ElseIf intR = 7 Then 'Claims完稿日
          If strMainText = "會稿Claims" Then
               strColName(intR) = "完稿日"
               strColText(intR) = m_924CP06
          End If
      ElseIf intR = 8 Then 'FCP管制人
          If strMainText = "會稿Claims" Then
               strColName(intR) = "管制人"
               strColText(intR) = m_FCNa16
          End If
      ElseIf intR = 9 Then '備註
          strColName(intR) = "備註"
          strColText(intR) = strNote
      'Added by Lydia 2020/04/14
      ElseIf intR = 10 Then '會稿工程師(承辦人)
          strColName(intR) = "會稿工程師"
          strColText(intR) = m_924CP14n
      End If
   Next intR
   
   'Add By Sindy 2023/9/14 純取值並沒有要產出承辦單
   If bolOnlyGetVal = True Then
      Pub_PrintFCP924Form = True 'Add By Sindy 2023/9/15
      Set rsA = Nothing
      Exit Function
   End If
   '2023/9/14 END
   
   'Added by Lydia 2020/04/10 檢查檔案是否存在
   If bolSaveFile = True Then
        m_FilePath = PUB_Getdesktop
        strSql = m_FilePath & "\" & m_strPA01 & m_strPA02 & IIf(m_strPA03 & m_strPA04 <> "000", m_strPA03 & m_strPA04, "") & strMainText & "承辦單.doc"
        If Dir(strSql) <> "" Then
             If PUB_ChkFileOpening(strSql, , False) = True Then
                  MsgBox strSql & vbCrLf & "檔案正在使用中，本次產生的電子檔會自動加上現在日期和時間。", vbExclamation
                  strSql = m_FilePath & "\" & m_strPA01 & m_strPA02 & IIf(m_strPA03 & m_strPA04 <> "000", m_strPA03 & m_strPA04, "") & "_" & strSrvDate(1) & Format(ServerTime, "000000") & strMainText & "承辦單.doc"
             Else
                  If PUB_DelPCOrgFile(strSql, , False) = False Then
                      MsgBox strSql & vbCrLf & "檔案無法刪除，本次產生的電子檔會自動加上現在日期和時間。", vbExclamation
                      strSql = m_FilePath & "\" & m_strPA01 & m_strPA02 & IIf(m_strPA03 & m_strPA04 <> "000", m_strPA03 & m_strPA04, "") & "_" & strSrvDate(1) & Format(ServerTime, "000000") & strMainText & "承辦單.doc"
                  End If
             End If
        End If
        m_FilePath = strSql
   End If
   'Add By Sindy 2022/8/12 借用此變數(m_strContactSheetA4)來傳回電子檔名
   If m_FilePath <> "" Then
      If InStr(m_strContactSheetA4, m_FilePath) = 0 Then
         m_strContactSheetA4 = IIf(m_strContactSheetA4 <> "", m_strContactSheetA4 & ";", "") & m_FilePath
      End If
   End If
   '2022/8/12 END
   'end 2020/04/10
   
   '下載樣本檔
   If strMainText = "會稿Claims" Then '輸入Claims完稿日,列印承辦單
       m_FileName = "外專翻譯_會稿CLAIMS承辦單_樣本.doc"
       Call PUB_GetSampleFile(m_FileName, "M51-000299-0-03")
   Else                                                '不輸入Claims完稿日,於後面列印翻譯承辦單+會稿說明書承辦單
       m_FileName = "外專翻譯_會稿說明書承辦單_樣本.doc"
       Call PUB_GetSampleFile(m_FileName, "M51-000299-0-04")
   End If
   
    If Dir(App.path & "\" & m_FileName) <> "" Then
         Screen.MousePointer = vbHourglass
         '判斷word是否已開啟
RestarWord:
         If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = True Then
             g_WordAp.Documents.Open App.path & "\" & m_FileName
             With g_WordAp
                .Selection.WholeStory
                .Selection.Copy
                For intR = 0 To iCall
                   'Modify By Sindy 2023/9/14 mark,提到前面取變數值
'                   strName = ""
'                   strText = ""
'                   If intR = 0 Then '速別
'                       strName = "速別"
'                       strText = strSpeed
'                   ElseIf intR = 1 Then '受文者
'                       strName = "受文者"
'                       strText = strGetPaper
'                   ElseIf intR = 2 Then '本所案號
'                       strName = "本所案號"
'                       strText = m_strPA01 & "-" & m_strPA02 & "-" & m_strPA03 & "-" & m_strPA04
'                   ElseIf intR = 3 Then '主旨
'                       strName = "主旨"
'                       strText = strMainText
'                   ElseIf intR = 4 Then '會稿所限
'                       strName = "會稿所限"
'                       strText = m_924CP06
'                   ElseIf intR = 5 Then '譯者
'                       strName = "譯者"
'                       strText = m_CP14t
'                   ElseIf intR = 6 Then '會稿法限
'                       strName = "會稿法限"
'                       strText = m_924CP07
'                   ElseIf intR = 7 Then 'Claims完稿日
'                       If strMainText = "會稿Claims" Then
'                            strName = "完稿日"
'                            strText = m_924CP06
'                       End If
'                   ElseIf intR = 8 Then 'FCP管制人
'                       If strMainText = "會稿Claims" Then
'                            strName = "管制人"
'                            strText = m_FCNa16
'                       End If
'                   ElseIf intR = 9 Then '備註
'                       strName = "備註"
'                       strText = strNote
'                   'Added by Lydia 2020/04/14
'                   ElseIf intR = 10 Then '會稿工程師(承辦人)
'                       strName = "會稿工程師"
'                       strText = m_924CP14n
'                   End If
    
                   'Find並且置換
                   If Trim(strColName(intR)) <> "" Then
                      .Selection.Find.ClearFormatting
                      .Selection.Find.Text = "|#" & strColName(intR) & "#|"
                      .Selection.Find.Replacement.Text = ""
                      .Selection.Find.Forward = True
                      .Selection.Find.Wrap = wdFindContinue
                      .Selection.Find.Format = False
                      .Selection.Find.MatchCase = False
                      .Selection.Find.MatchWholeWord = False
                      .Selection.Find.MatchWildcards = False
                      .Selection.Find.MatchSoundsLike = False
                      .Selection.Find.MatchAllWordForms = False
                      .Selection.Find.MatchByte = True
                      .Selection.Find.Execute
                      .Selection.Delete
                      .Selection.TypeText strColText(intR)
                      .Selection.Find.Execute
                      .Selection.Font.Underline = wdUnderlineSingle
                   End If
ReadNext:
                Next intR
                
               'Added by Lydia 2020/04/10
               If bolSaveFile = True Then
                    '產生Word檔
                    .ActiveDocument.SaveAs m_FilePath
               Else
               'end 2020/04/10
                    '直接列印
                   .ActiveDocument.PrintOut Background:=False, Copies:=1, Collate:=True
               End If 'Added by Lydia 2020/04/10
            End With
         End If
         Screen.MousePointer = vbDefault
         '還原Word位置
         Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop
         'Modified by Lydia 2020/04/10
         'g_WordAp.ActiveDocument.Close wdDoNotSaveChanges
         If bolSaveFile = False Then g_WordAp.ActiveDocument.Close wdDoNotSaveChanges

         g_WordAp.Quit wdDoNotSaveChanges
         Set g_WordAp = Nothing
         Pub_PrintFCP924Form = True 'Add By Sindy 2023/9/15
    Else
         MsgBox "無承辦單的樣本!", vbCritical
    End If
    
    Set rsA = Nothing
    Exit Function
    
ErrHand:
   If Err.Number = 462 Then '遠端伺服器不存在或無法使用
      GoTo RestarWord
   End If
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical
   End If
      
End Function

'Add By Sindy 2019/4/26 檢查是否有處理商標處MCT案件的權限
'Modify By Sindy 2023/5/12 + , Optional bolGetEmpId As Boolean = False : 取得MCTF員工編號
'                            , Optional bolInIsMail As Boolean = False : 是否含收信人員; Ex:MCTF07收信人員
Public Sub PUB_AddComboMCTF(strEmp As String, oCombo As Object, Optional bolGetEmpId As Boolean = False, _
   Optional bolInIsMail As Boolean = False)
   
Dim rsQuery As ADODB.Recordset
Dim j As Integer, ii As Integer
Dim strData As String, strText As String, varMCTFMan As Variant
Dim intQ As Integer, strCon1 As String 'Added by Lydia 2024/03/29

   'Add By Sindy 2023/5/12 取得MCTF員工編號
   If bolGetEmpId = True Then
      strCon1 = "select OCODE,OMAN from setSpecMan where OMAN is not null and instr(OCODE,'MCTF')>0"
   Else
   '2023/5/12 END
      strCon1 = "select OCODE,OMAN from setSpecMan where instr(OCODE,'MCTF')>0"
   End If
   'Add By Sindy 2023/5/12
   If bolInIsMail = False Then
      strCon1 = strCon1 & " and instr(ocode,'收信人員')=0"
   End If
   '2023/5/12 END
   strCon1 = strCon1 & " order by ocode asc"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      rsQuery.MoveFirst
      Do While Not rsQuery.EOF
         If InStr(Pub_GetSpecMan("MCTM"), strEmp) > 0 Or _
            InStr(Pub_GetSpecMan(rsQuery.Fields("OCODE")), strEmp) > 0 Then
            'Modify By Sindy 2023/5/12 取得MCTF員工編號
            If bolGetEmpId = True Then
               strData = rsQuery.Fields("OMAN")
            Else
               strData = rsQuery.Fields("OCODE")
            End If
            strData = Replace(strData, ",", ";")
            varMCTFMan = Split(strData, ";")
            For j = 0 To UBound(varMCTFMan)
               strText = varMCTFMan(j) & " " & GetPrjSalesNM(CStr(varMCTFMan(j)))
               For ii = 0 To oCombo.ListCount - 1
                  If InStr(oCombo.List(ii), Trim(strText)) = 1 Then
                     Exit For
                  End If
               Next ii
               If ii = oCombo.ListCount Then
                  oCombo.AddItem strText
               End If
            Next j
            '2023/5/12 END
         End If
         rsQuery.MoveNext
      Loop
   End If
   'Add By Sindy 2023/5/15 若為MCTM商標處大至台收文管理主管,增加檢查是否有離職的MCTF待處理信件
   If InStr(Pub_GetSpecMan("MCTM"), strEmp) > 0 Then
      strCon1 = "select distinct ir04 from tminput,inputrecord,staff" & _
               " Where nvl(ti16, 0)=0" & _
               " and ti01=ir01 and ti03=ir03 and ir04=st01" & _
               " and ir08=0 and st04='2'" & _
               " order by ir04 desc"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, strCon1)
      If intQ = 1 Then
         If InStr(Pub_GetSpecMan("MCTMember"), rsQuery.Fields(0)) > 0 Then
            strText = rsQuery.Fields(0) & " " & GetPrjSalesNM(rsQuery.Fields(0))
            For ii = 0 To oCombo.ListCount - 1
               If InStr(oCombo.List(ii), Trim(strText)) = 1 Then
                  Exit For
               End If
            Next ii
            If ii = oCombo.ListCount Then
               oCombo.AddItem strText
            End If
         End If
      End If
   End If
   '2023/5/15 END
   
   Set rsQuery = Nothing
End Sub

'Add By Sindy 2019/5/2 商標處信件收件人員拉下式選單
'Add By Sindy 2020/8/25 + Optional bolAddMCTF As Boolean = False
Public Function PUB_AddComboTMailEmp(strEmp As String, oCombo As Object, _
   Optional bolAddMCTF As Boolean = False) As String
Dim rsQuery As ADODB.Recordset
Dim strMCTFMan As String, varMCTFMan As Variant
Dim intQ As Integer, strCon1 As String 'Added by Lydia 2024/03/29

   '商標處人員
   If Left(PUB_GetST03(strEmp), 2) = "P2" Then
      strCon1 = "SELECT a0902,st01,st02 FROM staff,acc090 WHERE st04='1' and st01>'63' and st01<'F' and st03=a0901(+) and substr(st01,4,1)<>'9' and st03>='P20' and st03<='P29' and st01 not in('96029','96030') order by st03,st01 asc"
      intQ = 1
      Set rsQuery = ClsLawReadRstMsg(intQ, strCon1)
      If intQ = 1 Then
         rsQuery.MoveFirst
         Do While Not rsQuery.EOF
            oCombo.AddItem Trim(rsQuery.Fields("st01")) & " " & Trim(rsQuery.Fields("st02"))
            PUB_AddComboTMailEmp = PUB_AddComboTMailEmp & ";" & Trim(rsQuery.Fields("st01"))
            rsQuery.MoveNext
         Loop
      End If
      'Modify By Sindy 2023/7/11 Mark,桂所長已正式退休; 黃咸達已不再支援MCT小組,煩請一併移除
      'oCombo.AddItem "76012 " & GetPrjSalesNM("76012"): PUB_AddComboTMailEmp = PUB_AddComboTMailEmp & ";76012"
      oCombo.AddItem "96003 " & GetPrjSalesNM("96003"): PUB_AddComboTMailEmp = PUB_AddComboTMailEmp & ";96003"
      'oCombo.AddItem "A4009 " & GetPrjSalesNM("A4009"): PUB_AddComboTMailEmp = PUB_AddComboTMailEmp & ";A4009"
      'Add By Sindy 2020/8/25
      If bolAddMCTF = True Then
      '2020/8/25 END
         'Modify By Sindy 2021/11/24
         strMCTFMan = Pub_GetSpecMan("MCTMember")
         varMCTFMan = Split(strMCTFMan, ";")
         For intQ = 0 To UBound(varMCTFMan)
            If Left(varMCTFMan(intQ), 4) = "MCTF" Then
               oCombo.AddItem varMCTFMan(intQ) & " " & varMCTFMan(intQ)
               PUB_AddComboTMailEmp = PUB_AddComboTMailEmp & ";" & varMCTFMan(intQ)
            End If
         Next intQ
'         oCombo.AddItem "MCTF01 MCTF01": PUB_AddComboTMailEmp = PUB_AddComboTMailEmp & ";MCTF01"
'         oCombo.AddItem "MCTF02 MCTF02": PUB_AddComboTMailEmp = PUB_AddComboTMailEmp & ";MCTF02"
'         oCombo.AddItem "MCTF03 MCTF03": PUB_AddComboTMailEmp = PUB_AddComboTMailEmp & ";MCTF03"
'         oCombo.AddItem "MCTF04 MCTF04": PUB_AddComboTMailEmp = PUB_AddComboTMailEmp & ";MCTF04"
'         oCombo.AddItem "MCTF05 MCTF05": PUB_AddComboTMailEmp = PUB_AddComboTMailEmp & ";MCTF05"
         '2021/11/24 END
      End If
   End If
   If PUB_AddComboTMailEmp <> "" Then PUB_AddComboTMailEmp = Mid(PUB_AddComboTMailEmp, 2)
   Set rsQuery = Nothing
End Function

'Added by Lydia 2019/05/27 基本檔維護：若X, Y編號設定"FCP是否核對已准專利"上" N"，則出"核對已准專利"未發文之清單(word)
Public Function Pub_GetFA85CU122List(ByVal pKeyNo As String) As Boolean
Dim rsAD As New ADODB.Recordset
Dim intA As Integer, strA1 As String, strA2 As String
Dim strCase(1 To 4) As String
Dim strAll As String
Dim strFileName As String
Dim m_WordLeft As Long, m_WordTop As Long 'Word開啟位置
Dim bVisible As Boolean

   Pub_GetFA85CU122List = False
   If Left(pKeyNo, 1) = "X" Then
       strA1 = " and instr(pa26||','||pa27||','||pa28||','||pa29||','||pa30,'" & ChangeCustomerL(pKeyNo) & "') > 0 "
       strA2 = "客戶"
   ElseIf Left(pKeyNo, 1) = "Y" Then
       strA1 = " and pa75='" & pKeyNo & "' "
       strA2 = "代理人"
   Else
       Call ChgCaseNo(pKeyNo, strCase)
       If strCase(1) <> "FCP" Then Exit Function
       strA1 = " and cp02='" & strCase(2) & "' and cp03='" & strCase(3) & "' and cp04='" & strCase(4) & "' "
       strA2 = "案號"
   End If
   
   strA1 = " select cp01||'-'||cp02||decode(cp03||cp04,'','-'||cp03||'-'||cp04) caseno,sqldatet(cp05) cp05,pa26,pa27,pa27,pa28,pa29,pa30,pa75" & _
               " From caseprogress, patent where cp01='FCP' and cp158=0 and cp159=0 and cp10='926'" & _
               " and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and pa57 is null and pa108 is null " & strA1
   strA1 = strA1 & " order by cp05 asc "
   
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strA1)
   If intA = 1 Then
        If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then Exit Function
        
        With g_WordAp.Application
           .Selection.Font.Name = "細明體"
           '邊界
           .Selection.PageSetup.LeftMargin = .CentimetersToPoints(1)
           .Selection.PageSetup.RightMargin = .CentimetersToPoints(1)
           .Selection.PageSetup.TopMargin = .CentimetersToPoints(1)
           .Selection.PageSetup.BottomMargin = .CentimetersToPoints(0.8)
            '新增表格(1*4)
            .Selection.Tables.add Range:=.Selection.Range, NumRows:=1, NumColumns:=4
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
            .Selection.Cells.SetHeight RowHeight:=16, HeightRule:=wdRowHeightExactly '固定列高
            .Selection.Cells(1).SetWidth ColumnWidth:=.CentimetersToPoints(2.6), RulerStyle:=wdAdjustProportional
            .Selection.Cells(2).SetWidth ColumnWidth:=.CentimetersToPoints(10), RulerStyle:=wdAdjustProportional
            .Selection.Cells(3).SetWidth ColumnWidth:=.CentimetersToPoints(3.5), RulerStyle:=wdAdjustProportional
            
            .Selection.InsertRows 3
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.SelectRow
            .Selection.Cells.Merge
            .Selection.ParagraphFormat.Alignment = wdAlignParagraphCenter
            .Selection.Cells.SetHeight RowHeight:=28, HeightRule:=wdRowHeightExactly '固定列高
            .Selection.Font.Size = 16
            .Selection.Font.Bold = True
            .Selection.TypeText Text:="核對已准專利未發文之案件清單"
            
            .Selection.MoveRight Unit:=wdCharacter, Count:=2
            .Selection.SelectRow
            .Selection.Font.Size = 12
            .Selection.Font.Bold = False
            
            .Selection.SelectRow
            .Selection.Collapse Direction:=wdCollapseStart
            .Selection.TypeText Text:="修改編號："
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:=pKeyNo
            .Selection.MoveRight Unit:=wdCharacter, Count:=3 '換列
            
            .Selection.TypeText Text:="本所案號"
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:="申請人編號1~5"
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:="FC代理人編號"
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.TypeText Text:="二核收文日"
            .Selection.MoveRight Unit:=wdCharacter, Count:=1
            .Selection.SelectRow
            .Selection.Borders(wdBorderBottom).LineStyle = wdLineStyleSingle '用底部框線當做分隔線
            .Selection.Collapse Direction:=wdCollapseEnd
            
            '清單內容
            rsAD.MoveFirst
            Do While Not rsAD.EOF
                .Selection.TypeText Text:="" & rsAD.Fields("caseno")
                .Selection.MoveRight Unit:=wdCharacter, Count:=1
                strA1 = ""
                For intA = 27 To 30
                   strA1 = strA1 & IIf("" & rsAD.Fields("pa" & intA) <> "", "," & rsAD.Fields("pa" & intA), "")
                Next
                .Selection.TypeText Text:="" & rsAD.Fields("pa26") & strA1
                .Selection.MoveRight Unit:=wdCharacter, Count:=1
                .Selection.TypeText Text:="" & rsAD.Fields("pa75")
                .Selection.MoveRight Unit:=wdCharacter, Count:=1
                .Selection.TypeText Text:="" & rsAD.Fields("cp05")
                .Selection.MoveRight Unit:=wdCharacter, Count:=2
                .Selection.InsertRows
                If rsAD.AbsolutePosition = 1 Then
                   .Selection.Borders(wdBorderTop).LineStyle = wdLineStyleNone
                End If
                rsAD.MoveNext
            Loop
        End With
        '存PDF檔
        Pub_RePosWord g_WordAp, bVisible, m_WordLeft, m_WordTop '還原Word位置
        'Modified by Lydia 2023/04/27 改模組
        'If PUB_PrintWord2PDF(g_WordAp, App.path, pKeyNo, strFileName) = True Then
        If PUB_PrintWord2File(g_WordAp, App.path, pKeyNo, strFileName) = True Then
            strFileName = App.path & "\" & strFileName
        Else
            strFileName = ""
        End If
        Set rsAD = Nothing
   End If
   
   If strFileName = "" Then Exit Function
   
   If Dir(strFileName) <> "" Then
        MsgBox "此" & strA2 & "尚有核對已准專利未發文之案件，系統將出案件清單！", vbExclamation, "FCP是否核對已准專利"
        PUB_SendMail strUserNum, strUserNum, "", pKeyNo & "核對已准專利未發文之案件清單", "請詳參附件。", , strFileName
        Sleep 1000
        SetAttr strFileName, vbNormal 'Added by Lydia 2020/03/19 預設檔案為正常
        Kill strFileName
   End If
End Function

'Add By Sindy 2019/6/10 ex:T-149278
'申請國家非台灣時,下一程序會新增收達997,提申998的期限現在發現有期限超過法定期限的情形,故再修改
'1 西元 '2 民國
Public Function PUB_T997998LimitDate(strNP08 As String, strCP07 As String, intDtType As Integer) As String
   PUB_T997998LimitDate = strNP08
   If Val(strNP08) > 0 Then strNP08 = DBDATE(strNP08)
   If Val(strCP07) > 0 Then strCP07 = DBDATE(strCP07)
   
   '法定期限有值且為系統日或者過期時，收達期限或提申期限都管制為系統日期
   If Val(strCP07) > 0 And Val(strCP07) <= Val(strSrvDate(1)) Then
      If intDtType = 1 Then '西元
         PUB_T997998LimitDate = strSrvDate(1)
      Else
         PUB_T997998LimitDate = strSrvDate(2)
      End If
   'Modify By Sindy 2025/9/15
      Exit Function
   End If
   'Modify By Sindy 2025/9/15 +檢查管制期限<系統日,管制期限為系統日
   If Val(strNP08) > 0 And Val(strNP08) < Val(strSrvDate(1)) Then
      If intDtType = 1 Then '西元
         PUB_T997998LimitDate = strSrvDate(1)
      Else
         PUB_T997998LimitDate = strSrvDate(2)
      End If
      Exit Function
   End If
   '計算出來的期限>=發文進度的法定期限CP07時,
   '期限改掛法定期限的前一工作日，若前一工作日<=系統日時再改為系統日的下一個工作日。
   'ElseIf Val(strNP08) >= Val(strCP07) And Val(strNP08) > 0 And Val(strCP07) > 0 Then
   If Val(strNP08) >= Val(strCP07) And Val(strNP08) > 0 And Val(strCP07) > 0 Then
   '2025/9/15 END
      PUB_T997998LimitDate = PUB_GetWorkDayAfterSysDate(CDbl(strCP07), -1)
      If DBDATE(PUB_T997998LimitDate) <= Val(strSrvDate(1)) Then
         PUB_T997998LimitDate = PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), 1)
      End If
      If intDtType = 1 Then '西元
         PUB_T997998LimitDate = PUB_T997998LimitDate + 19110000
      End If
   End If
End Function

'Add By Sindy 2019/8/8 取得電子送件申請書-優先權資料(by 專利案)
Public Function PUB_GetAppPridate(pa() As String, ET01 As String, strReceiveNo As String, ET03 As String) As String
Dim rsQuery As ADODB.Recordset
Dim ii As String, jj As String
Dim intQ As Integer, strCon1 As String 'Added by Lydia 2024/03/29
   
   PUB_GetAppPridate = ""
   '優先權資料
   'Modify By Sindy 2018/7/25 + order by pd05 asc
   strCon1 = "SELECT sqldatew(pd05) pd05,na72,pd06,pd07,decode(pd08,'1','發明','2','新型','3','設計',pd08) pd08,pd09" & _
            " FROM pridate,nation where pd01='" & pa(1) & "' and pd02='" & pa(2) & "' and pd03='" & pa(3) & "' and pd04='" & pa(4) & "'" & _
            " and na01(+)=pd07" & _
            " order by pd05 asc"
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      jj = 0
      rsQuery.MoveFirst
      Do While Not rsQuery.EOF
         jj = jj + 1
         If PUB_GetAppPridate <> "" Then PUB_GetAppPridate = PUB_GetAppPridate & vbCrLf & vbCrLf
         PUB_GetAppPridate = PUB_GetAppPridate & "【主張優先權" & jj & "】" & vbCrLf
         PUB_GetAppPridate = PUB_GetAppPridate & "　　【申請日】　　　　　　　　" & rsQuery("pd05") & vbCrLf
         PUB_GetAppPridate = PUB_GetAppPridate & "　　【受理國家或地區】　　　　" & rsQuery("na72") & vbCrLf
         PUB_GetAppPridate = PUB_GetAppPridate & "　　【申請案號】　　　　　　　" & rsQuery("pd06")
         'Added by Morgan 2017/8/9
         'Modify By Sindy 2017/11/9 輸入優先權國家代碼時,代表是以電子交換檢送
         'Modify By Sindy 2018/8/10 ex:FCP-59317 韓國,分交換和不交換
         If rsQuery("pd07") = "012" Then '韓國
            If rsQuery("pd07") = "" & rsQuery("pd09") Then
         '2017/11/9 END
               '電子交換
               If PUB_GetAppPridate <> "" Then PUB_GetAppPridate = PUB_GetAppPridate & vbCrLf
               PUB_GetAppPridate = PUB_GetAppPridate & "　　【存取碼】　　　　　　　　交換"
            Else
               '非電子交換
               If PUB_GetAppPridate <> "" Then PUB_GetAppPridate = PUB_GetAppPridate & vbCrLf
               PUB_GetAppPridate = PUB_GetAppPridate & "　　【存取碼】　　　　　　　　不交換"
            End If
            '2018/8/10 END
         ElseIf Not IsNull(rsQuery("pd09")) Then
            '非電子交換
            'Add By Sindy 2017/11/9
            If PUB_GetAppPridate <> "" Then PUB_GetAppPridate = PUB_GetAppPridate & vbCrLf
            PUB_GetAppPridate = PUB_GetAppPridate & "　　【專利類別】　　　　　　　" & ChgSQL("" & rsQuery("pd08"))
            '2017/11/9 END
            If PUB_GetAppPridate <> "" Then PUB_GetAppPridate = PUB_GetAppPridate & vbCrLf
            PUB_GetAppPridate = PUB_GetAppPridate & "　　【存取碼】　　　　　　　　" & ChgSQL(rsQuery("pd09"))
         End If
         rsQuery.MoveNext
      Loop
   End If
   
   Set rsQuery = Nothing
End Function


'Added by Lydia 2019/09/27 依照系統別+定稿語文，取得共同(放假)的信函備註
Public Function Pub_GetLetterMemo(p_Sys As String, p_Language As String) As String
Dim stSQL As String, intR As Integer
Dim rs1 As New ADODB.Recordset
   
   Pub_GetLetterMemo = ""
   stSQL = "select LM05 from lettermemo where LM01='" & p_Sys & "' and LM02='" & p_Language & "' and (" & strSrvDate(1) & " between LM03 and LM04 )"
   intR = 1
   Set rs1 = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      Pub_GetLetterMemo = "" & rs1.Fields(0)
   End If
   
   Set rs1 = Nothing
End Function

'Added by Lydia 2020/01/22 比照公告公報frm060302的路徑，提前檢查Pat3是否開啟 ;
Public Function Pub_CheckGazetteDir(Optional ByVal pDirPath As String) As Boolean

On Error GoTo ErrHandleSpec 'Added by Lydia 2020/03/16

   Pub_CheckGazetteDir = False
   
   If pDirPath = "" Then
       pDirPath = PUB_GetLastDate("frm060302", UCase("txtPath2"))
   End If
   If pDirPath = "" Then  '預設
       pDirPath = "\\Pat3\GAZETTE\PXml\img_1\isu012012\"
   End If
   '雖然Vb有寫ping模組,但是目前無法傳入hostname偵測
   If InStr(UCase(pDirPath), UCase("\PXml")) > 0 Then
       pDirPath = Mid(pDirPath, 1, InStr(UCase(pDirPath), UCase("\PXml")) - 1)
       If Dir(pDirPath & "\PXml", vbDirectory) <> "" Then
           Pub_CheckGazetteDir = True
       End If
   End If
   
   Exit Function
   
ErrHandleSpec:
   If Err.Number <> 0 Then
       MsgBox "無法連接" & pDirPath & vbCrLf & "請確定共用電腦是否正常開啟！", vbCritical, "檢查共用電腦"
   End If
End Function

'Added by Lydia 2020/02/14 外專：案件名稱有特殊字，開啟/維護FCP0xxxxx.新案性質.案件名稱.doc
Public Function Pub_GetPA174toFile(ByVal pType As String, ByVal Cno01 As String, ByVal Cno02 As String, ByVal Cno03 As String, ByVal Cno04 As String, _
                       Optional ByRef frmMe As Form, Optional ByRef frmTmp As Form) As Boolean
'pType : 0-開啟, 1-維護, 2-檢查是否有案件名稱.doc, 3-維護+回傳處理
'Cno01~Cno04: 本所案號
'frmTmp: 原始檔Word維護-修改(frm100101_M_1)
Dim intQ As Integer, strQ As String
Dim rsQD As New ADODB.Recordset
Dim stPA05 As String, stPA06 As String, stPA07 As String
Dim stPA174 As String   '案件名稱有特殊字
Dim stCP09 As String, stCP10 As String
Dim bolChkModify As Boolean '是否可維護: 檢查權限、是否有frm100101_M_1
Dim bolDown As Boolean '是否下載完成
Dim mFilePath As String
Dim bVisible As Boolean '是否顯示Word
Dim m_WordLeft As Long, m_WordTop As Long 'Word開啟位置
Dim hLocalFile As Long
Dim strKind As String '上傳到原始檔的處理方式

    Pub_GetPA174toFile = False
    bolChkModify = False
    
    If pType = "" Or Cno01 = "" Or Cno02 = "" Then Exit Function

    If Cno03 = "" Then Cno03 = "0"
    If Cno04 = "" Then Cno04 = "00"
   
On Error GoTo ErrorHand01

    strQ = "select pa01,pa02,pa03,pa04,pa05,pa06,pa07,pa174,cp09,cp10,tct04,tct07,tct10,st04,st52||','||st53||','||st54||','||st55 as st52st55 " & _
              "From patent, caseprogress, transcasetitle, staff " & _
              "where pa01='" & Cno01 & "' and pa02='" & Cno02 & "' and pa03='" & Cno03 & "' and pa04='" & Cno04 & "' " & _
              "and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and cp31='Y' and cp09=tct01(+) and tct10=st01(+) "
    intQ = 1
    Set rsQD = ClsLawReadRstMsg(intQ, strQ)
    If intQ = 0 Then
        'Added by Lydia 2020/04/23 基本檔維護舊系統轉來的案件,排除無新案收文的狀態 (ex.Amy 補P-018482的准駁有彈此訊息)
        strQ = "select cp09 from caseprogress where cp31='Y' and cp01='" & Cno01 & "' and cp02='" & Cno02 & "' and cp03='" & Cno03 & "' and cp04='" & Cno04 & "' "
        intQ = 1
        Set rsQD = ClsLawReadRstMsg(intQ, strQ)
        If intQ = 0 Then
            GoTo EXITSUB
        Else
        'end 2020/04/
            MsgBox "查無此案號的資料!", vbCritical + vbOKOnly, "案件名稱有特殊字"
            GoTo EXITSUB
        End If 'Added by Lydia 2020/04/23
    Else
    
       stPA05 = "" & rsQD.Fields("pa05")
       stPA06 = "" & rsQD.Fields("pa06")
       stPA07 = "" & rsQD.Fields("pa07")
       stPA174 = "" & rsQD.Fields("pa174")
       stCP09 = "" & rsQD.Fields("cp09")
       stCP10 = "" & rsQD.Fields("cp10")
       '非外專程序、外專承辦或命名之工程師(及其主管)等人員，只開啟檔案不提供上傳功能
       If pType = "1" Or pType = "3" Then  '1-維護模式、3-維護+回傳處理模式
            If (Pub_StrUserSt03 = "M51" Or Pub_StrUserSt03 = "F22" Or Pub_StrUserSt03 = "F23" Or _
               InStr(rsQD.Fields("tct10") & "," & rsQD.Fields("tct04") & "," & rsQD.Fields("tct07") & "," & rsQD.Fields("st52st55"), strUserNum) > 0) Then
               bolChkModify = True
            End If
            If TypeName(frmTmp) <> "frm100101_M_1" Then '無表單,僅供開啟
               bolChkModify = False
            End If
       End If
       If pType = "3" Then
          strKind = "D2" '重複檔案，直接刪除; 是否有執行上傳，回傳給前表單
       Else
          strKind = "D" '預設：重複檔案，直接刪除
       End If
       
       strQ = "select cpf01,cpf02,cpf13 from casepaperfile where cpf01='" & stCP09 & "' " & _
                 "and upper(cpf02) like '%.案件名稱.%' and upper(cpf02) like '%.DOC%' and substr(upper(cpf02),-4)<>upper('.del') "
       intQ = 1
       Set rsQD = ClsLawReadRstMsg(intQ, strQ)
       If pType = "2" Then '2-檢查是否有案件名稱.doc
           If intQ = 1 Then
               Pub_GetPA174toFile = True
           Else
               Pub_GetPA174toFile = False
           End If
           Exit Function
       Else '開啟/維護模式
           If intQ = 0 Then
              If pType = "0" Then  '開啟
                  MsgBox "查無此案件的案件名稱Word檔!", vbCritical + vbOKOnly, "案件名稱有特殊字"
                  GoTo EXITSUB
              ElseIf pType = "1" Or pType = "3" Then  '維護->新增
                  If bolChkModify = False Then
                     MsgBox "無維護權限，並且查無此案件的案件名稱Word檔!", vbCritical + vbOKOnly, "案件名稱有特殊字"
                     GoTo EXITSUB
                  Else
                     '檢查若已定稿維護畫面已開啟時確認是否只開Word並提醒無法直接上傳
                     If PUB_CheckFormExist("frm100101_M_1") Then
                         MsgBox "查無此案件的案件名稱Word檔!" & vbCrLf & "原始檔Word維護畫面已開啟，不得繼續作業!", vbCritical + vbOKOnly, "案件名稱有特殊字"
                         GoTo EXITSUB
                     End If
                     '開啟新的Word檔
                     If Pub_NewWordDoc(g_WordAp, bVisible, m_WordLeft, m_WordTop) = False Then GoTo EXITSUB
                     With g_WordAp.Application
                         .Selection.TypeText Text:="本所案號：" & Cno01 & "-" & Cno02 & "-" & Cno03 & "-" & Cno04
                         .Selection.TypeParagraph
                         .Selection.TypeText Text:="中文名稱：" & stPA05
                         .Selection.TypeParagraph
                         .Selection.TypeText Text:="英文名稱：" & stPA06
                         .Selection.TypeParagraph
                         'Modified by Lydia 2022/04/25 「日文名稱」改為「外文名稱」
                         .Selection.TypeText Text:="外文名稱：" & stPA07
                         .Selection.TypeParagraph
                     End With
                     Call frmTmp.SetFormParent(IIf(UCase(TypeName(frmMe)) <> "NOTHING", frmMe, frmTmp), Cno01 & "-" & Cno02 & "-" & Cno03 & "-" & Cno04, _
                                                  stCP09, PUB_FCPCaseNo2FileName(Cno01, Cno02, Cno03, Cno04) & "." & stCP10 & ".案件名稱.docx", strKind, stCP10)
                     frmTmp.Show
                  End If
              End If
           Else
                mFilePath = App.path & "\" & strUserNum & "\" & rsQD.Fields("cpf02")  '直接以檔案名稱開啟
                If Dir(mFilePath) <> "" Then
                   '檢查檔案是否正在使用中
                   If PUB_ChkFileOpening(mFilePath) = True Then
                        MsgBox mFilePath & vbCrLf & "檔案正在使用中（請關閉），方可繼續操作。", vbExclamation
                        Exit Function
                   End If
                   SetAttr mFilePath, vbNormal '預設檔案為正常
                   Kill mFilePath
                End If
                bolDown = PUB_GetFtpFile("" & rsQD.Fields("cpf13"), mFilePath, "CASEPAPERFILE", True)
                If bolDown = False Then
                     MsgBox "無法開啟檔案[ " & rsQD.Fields("cpf02") & " ]！", "案件名稱有特殊字"
                     GoTo EXITSUB
                Else
                    If bolChkModify = False Then
                        If pType = "1" Or pType = "3" Then  '從維護作業進入，卻無權限
                            MsgBox "無維護權限，此次修改將以 Word 開啟且無法直接上傳！", vbCritical + vbOKOnly, "案件名稱有特殊字"
                        End If
JumpToOpen:   '只開啟檔案
                        'SetAttr mFilePath, vbReadOnly '檔案設定成唯讀屬性 '保留
                        '開啟檔案
                        ShellExecute hLocalFile, "open", mFilePath, vbNullString, vbNullString, 1
                        'GoTo EXITSUB 'Mark by Lydia 2020/02/25
                    Else
                        '檢查若已定稿維護畫面已開啟時確認是否只開Word並提醒無法直接上傳
                        If PUB_CheckFormExist("frm100101_M_1") Then
                            MsgBox "原始檔Word維護畫面已開啟，此次修改將以 Word 開啟且無法直接上傳！", vbExclamation + vbOKOnly, "案件名稱有特殊字"
                            GoTo JumpToOpen
                        Else
                            If PUB_OpenWord(mFilePath) = True Then
                                '預設上傳模式為:有重複檔案，D=刪除舊檔
                                Call frmTmp.SetFormParent(IIf(UCase(TypeName(frmMe)) <> "NOTHING", frmMe, frmTmp), Cno01 & "-" & Cno02 & "-" & Cno03 & "-" & Cno04, _
                                                             stCP09, "" & rsQD.Fields("cpf02"), strKind, stCP10)
                                frmTmp.Show
                            End If
                        End If
                    End If
                End If
           End If
       End If '開啟/維護模式
    End If
    
    Pub_GetPA174toFile = True
    
EXITSUB:
    Set rsQD = Nothing
    Exit Function
    
ErrorHand01:
    If Err.Number <> 0 Then
        If Err.Description <> "" Then MsgBox Err.Description, vbExclamation, "案件名稱有特殊字"
        Resume Next
    End If
End Function

'Added by Lydia 2020/03/09 FCT案輸入註冊證或更正核准(註冊證)前，先掃瞄註冊證至固定資料夾，輸註冊證若缺檔則提醒不可輸入，不缺則自動歸入註冊證那道之卷宗區。
Public Function PUB_FCTCheckPDF(ByVal iCP01 As String, ByVal iCP02 As String, ByVal iCP03 As String, ByVal iCP04 As String, ByVal iCP10 As String, Optional ByVal iCp09 As String, Optional ByRef pFileName As String) As Boolean
Dim fs, f
Dim strFolder As String
Dim strA1 As String
Dim intQ As Integer
Dim rsQuery As New ADODB.Recordset
Dim bolCheck As Boolean
Dim strErr As String
Dim strFileName As String

On Error Resume Next

    strFileName = ""
    PUB_FCTCheckPDF = True
    bolCheck = False
    
    If iCP01 <> "FCT" Or iCP02 = "" Or iCP10 = "" Then
        Exit Function
    End If
    'Remark by Lydia 2022/02/10 FCT紙本公文來函(非電子公文輸入)在確定存檔前會先檢查來源檔案是否存在。
                                              ' 除了原本的核准-註冊證的更正、補換發註冊證103的核准和註冊證1701，加上核准、核駁、審查報告、其它來函輸入
    'If iCP10 <> "1701" And iCP10 <> "1001" Then
    '    Exit Function
    'End If
    'If iCP10 = "1001" And iCp09 <> "" Then
    '    '核准-註冊證的更正
    '    strA1 = "SELECT C2.CP09,C2.CP10 FROM CASEPROGRESS C1,CASEPROGRESS C2 WHERE C1.CP09='" & iCp09 & "' AND C1.CP10='302' AND C1.CP43=C2.CP09(+) "
    '    intQ = 1
    '    Set rsQuery = ClsLawReadRstMsg(intQ, strA1)
    '    If intQ = 1 Then
    '        If "" & rsQuery.Fields("CP10") = "1701" Then
    '             bolCheck = True
    '        End If
    '    End If
    '    'Added by Lydia 2020/07/17補換發註冊證103的核准
    '    If bolCheck = False Then
    '        strA1 = "SELECT CP09,CP10 FROM CASEPROGRESS WHERE CP09='" & iCp09 & "' "
    '        intQ = 1
    '        Set rsQuery = ClsLawReadRstMsg(intQ, strA1)
    '        If intQ = 1 Then
    '           If InStr("103,", rsQuery.Fields("CP10") & ",") > 0 Then
    '               bolCheck = True
    '           End If
    '        End If
    '    End If
    '    'end 2020/07/17
    'Else
    '    bolCheck = True
    'End If
    'If bolCheck = False Then Exit Function
    'end --- Remark by Lydia 2022/02/10
    
    strFolder = Pub_GetSpecMan("FCT註冊證存放路徑")
    strA1 = strFolder
    
    '測試抓桌面的相同資料夾以免誤刪真實檔案
    If Pub_StrUserSt03 = "M51" Or UCase(pub_DbTerminalName) <> 正式資料庫電腦名稱 Then
JumpChk:
       If MsgBox("共用資料夾路徑：" & strFolder & vbCrLf & "是否採用？", vbInformation + vbYesNo + vbDefaultButton2) = vbNo Then
            strA1 = InputBox("本機路徑：", , PUB_Getdesktop & "\" & Mid(strFolder, InStrRev(strFolder, "\") + 1))
            If Dir(strA1, vbDirectory) = "" Then
                MsgBox "路徑不存在！", vbCritical
                GoTo JumpChk
            End If
       End If
    End If
    strFolder = strA1
        
    strFileName = Dir(strFolder & "\" & iCP01 & "*" & Val(iCP02) & IIf(iCP03 <> "0", "*" & iCP03, "") & IIf(iCP04 <> "00", "*" & iCP04, "") & "*.pdf")
    
    If strFileName <> "" Then
        Do While strFileName <> ""
             Set fs = CreateObject("Scripting.FileSystemObject")
             Set f = fs.GetFile(strFolder & "\" & strFileName)
             '檔案大小為 0 KB 有誤
             If f.Size = 0 Then
                  strErr = strErr & vbCrLf & strFolder & "\" & strFileName & "，" & MsgText(9221)
                  GoTo JumpNextDir
             End If
             If PUB_ChkFileOpening(strFolder & "\" & strFileName) = True Then
                 strErr = strErr & vbCrLf & strFolder & "\" & strFileName & "，檔案正在使用中，請關閉或關閉檔案後間隔1分鐘，方能上傳到卷宗區。"
                 GoTo JumpNextDir
             End If
             '檢查檔名規則
             If PUB_ChkEmpFlowFNMRule(iCP01 & "-" & iCP02 & "-" & iCP03 & "-" & iCP04, strFileName, "Y", iCP10, , , False, False, strErr) = False Then
                  GoTo JumpNextDir
             End If
             
             pFileName = pFileName & "*" & strFolder & "\" & strFileName
JumpNextDir:
             strFileName = Dir()
        Loop
        
        If strErr <> "" Then
            MsgBox strErr, vbCritical, "檢查掃瞄檔"
            Exit Function
            PUB_FCTCheckPDF = False
        Else
            pFileName = Mid(pFileName, 2)
        End If
    End If
    If pFileName = "" Then
        'Modified by Lydia 2022/02/10
        'MsgBox "缺少" & IIf(iCP10 = "1001", "", "註冊證") & "掃瞄PDF檔，不可輸入！", vbCritical, "檢查掃瞄檔"
        intQ = ClsPDGetCaseProperty(iCP01, iCP10, strA1)
        MsgBox "缺少〔" & strA1 & "〕掃瞄PDF檔，不可輸入！", vbCritical, "檢查掃瞄檔"
        'end 2022/02/10
        PUB_FCTCheckPDF = False
    End If
    
End Function

'Added by Lydia 2018/07/18 FCT輸入註冊證或核准-註冊證的更正，自動將固定資料夾的掃瞄PDF檔，上傳到卷宗區
Public Function Pub_AutoSavePdf2_FCT(ByVal iCP01 As String, ByVal iCP02 As String, ByVal iCP03 As String, ByVal iCP04 As String, ByVal iCp09 As String, ByVal iCP10 As String, ByVal iFilePath) As Boolean
'iFilePath :完整檔案路徑，用*區隔多個檔案
Dim fs, f
Dim strErr As String
Dim strFileName As String, stReName As String
Dim tmpArr As Variant
Dim strB01 As String, intB As Integer
Dim rsB1 As New ADODB.Recordset
Dim bolConn As Boolean
Dim intR As Integer 'Added by Lydia 2020/04/29

    If iCP01 <> "FCT" Or iCP02 = "" Or iCp09 = "" Or iFilePath = "" Then
        Exit Function
    End If
    
On Error GoTo JumpExit

    tmpArr = Split(iFilePath, "*")
    For intB = LBound(tmpArr) To UBound(tmpArr)
        If tmpArr(intB) <> "" Then
            strFileName = Dir(tmpArr(intB))
            If strFileName <> "" Then
                Set fs = CreateObject("Scripting.FileSystemObject")
                Set f = fs.GetFile(tmpArr(intB))
                If PUB_ChkFileOpening("" & tmpArr(intB)) = True Then
                    strErr = strErr & vbCrLf & "" & tmpArr(intB) & "，檔案正在使用中，請關閉或關閉檔案後間隔1分鐘，方能上傳到卷宗區。"
                    GoTo JumpNextDir
                End If
                
                '更名
                If PUB_GetEmpFlowReNameFile(iCP01, iCP02, iCP03, iCP04, iCP10, strFileName, stReName, True, 1, False, strErr) = False Then
                     GoTo JumpNextDir
                End If

                '檢查卷宗區檔案是否存在
                strB01 = "SELECT cpp01,cpp02 FROM casepaperpdf " & _
                              "WHERE cpp01 ='" & iCp09 & "' and instr(upper(cpp02),'" & UCase(stReName) & "') > 0 and instr(upper(cpp02),'PDF.DEL') = 0 "
                'Modified by Lydia 2020/04/29 debug intB=>intR (ex.FCT043498註冊證多掃描一次)
                intR = 1
                Set rsB1 = ClsLawReadRstMsg(intR, strB01)
                If intR = 1 Then
                'end 2020/04/29
                     strErr = strErr & vbCrLf & rsB1.Fields("cpp02") & "，卷宗區檔案已存在！"
                     GoTo JumpNextDir
                End If
                '上傳到卷宗區
                If bolConn = False Then
                    bolConn = True
                    cnnConnection.BeginTrans
                End If
                If SaveAttFile_PDF(iCp09, "" & tmpArr(intB), stReName, Format(f.DateLastModified, "YYYYMMDD"), Format(f.DateLastModified, "HHMMSS"), False) = False Then
                     strErr = strErr & vbCrLf & "" & tmpArr(intB) & "，存檔失敗！" & vbCrLf & Err.Description
                     bolConn = False
                     GoTo JumpNextDir
                End If
                fs.DeleteFile tmpArr(intB), True '刪檔
            End If
        End If
        
JumpNextDir:
    Next intB
    
    If bolConn = True Then
        cnnConnection.CommitTrans
    End If
    
    Pub_AutoSavePdf2_FCT = True
    
JumpExit:

Set rsB1 = Nothing

If Err.Number <> 0 Then strErr = strErr & vbCrLf & Err.Description
If bolConn = False Then cnnConnection.RollbackTrans

If strErr <> "" Then
   MsgBox "FCT自動將註冊證的掃瞄PDF檔，上傳到卷宗區作業失敗：" & strErr, vbCritical
End If

End Function

'Added by Lydia 2020/03/25 委任書: 設定公司別下拉選項
Public Sub PUB_SetCboTofrm210114(ByVal pFrmName As String, ByRef pCmb As Object, ByRef pCompSeal As String)
Dim strB As String
Dim intB As Integer
Dim rsB As New ADODB.Recordset

    If strSrvDate(1) < 事務所合併日 Then  '1+2公司
        strB = "1,2"
    Else
          '其他
          strB = "2"  '抓2公司
    End If
    If pFrmName = "frm210114_5" And strSrvDate(1) >= 智慧所更名日 Then '常年顧問聘任書: 智慧所更名日起只可為L公司
        strB = "L"
    End If
    'Modified by Lydia 2022/04/01 +著作權案件委任契約書frm210114_7
    If InStr("frm210114_2,frm210114_4,frm210114_7", pFrmName) > 0 Then  'CFP、CFT案,另外抓智權
        strB = strB & ",J"
    End If
    
    pCompSeal = ""
    strB = "select a0801,a0802 from acc080 where a0801 in (" & GetAddStr(strB) & ") order by a0801 desc "
    intB = 1
    Set rsB = ClsLawReadRstMsg(intB, strB)
    If intB = 1 Then
        pCmb.Clear
        rsB.MoveFirst
        Do While Not rsB.EOF
             pCmb.AddItem "" & rsB.Fields("a0802"), 0
             pCompSeal = pCompSeal & "," & rsB.Fields("a0802") & "|"
             Select Case "" & rsB.Fields("a0801") '公司名稱|用印編號(方章圖檔的編號)
                 Case "1"
                      pCompSeal = pCompSeal & "52" '專利商標
                 Case "2"
                      If strSrvDate(1) < 智慧所更名日 Then
                          pCompSeal = pCompSeal & "51" '專利法律
                      Else   '新章
                          pCompSeal = pCompSeal & "66" '智慧所
                      End If
                 Case "J"
                      pCompSeal = pCompSeal & "53" '智權公司
                 Case "L"
                      pCompSeal = pCompSeal & "67" '法務所
             End Select
             rsB.MoveNext
        Loop
        pCmb.AddItem "", 0
        pCompSeal = Mid(pCompSeal, 2)
    End If
    
    Set rsB = Nothing
    
End Sub

'Added by Lydia 2024/08/06 傳入Combo和字串，將ListIndex設定在符合文字
Public Sub Pub_SetCboListIdx(ByRef pObj As Object, ByVal pFindTxt As String)
Dim intJ As Integer
   
   If Trim(pFindTxt) = "" Then
      pObj.ListIndex = 0
   Else
      If pObj.ListCount >= 2 Then
         For intJ = 0 To pObj.ListCount - 1
            If pObj.List(intJ) = pFindTxt Then
               pObj.ListIndex = intJ
               Exit For
            End If
         Next intJ
         pObj.ListIndex = 0
      End If
   End If
   
End Sub

'Added by Lydia 2020/04/09 傳入公司名稱，回傳通訊地址 by 委任契約書 (因為a0802地址太長）
Public Function PUB_SetAddrTofrm210114(ByVal pCompName As String) As String
Dim pAddr As String

    Select Case pCompName
         Case "台一國際專利商標事務所" '1-公司
               pAddr = "台北市長安東路二段一一二號十樓"
         Case "台一國際智慧財產事務所" '2-公司
               pAddr = "台北市長安東路二段一一二號九樓"
         Case "台一智權股份有限公司" 'J-公司
               pAddr = "台北市長安東路二段一一０號四樓"
         Case "台一國際法律事務所" 'L-公司
               pAddr = "台北市長安東路二段一一０號四樓之一"
         Case Else  '預設
               pAddr = "台北市長安東路二段一一二號九樓"
    End Select
    
    PUB_SetAddrTofrm210114 = pAddr
End Function


'Add by Sindy 2020/7/24 傳入中小企業符合減收資格依據代碼,轉換中文
Public Function PUB_AD15ToText(ByVal pAD15 As String, ByVal pAD16 As String) As String
   PUB_AD15ToText = ""
   If Trim(pAD16) = "" Then pAD16 = "　　　　　　"
   Select Case pAD15
      Case "1"
         PUB_AD15ToText = "製造業、營造業、礦業及土石採取業實收資本額八千萬以下：" & pAD16 & "元"
      Case "2"
         PUB_AD15ToText = "前項除外之其他行業前一年營業額一億元以下：" & pAD16 & "元"
      Case "3"
         PUB_AD15ToText = "我國製造業、營造業、礦業及土石採取業實收資本額新台幣八千萬以上但經常僱用員工數未滿200人：員工數" & pAD16 & "人"
      Case "4"
         PUB_AD15ToText = "我國前項除外之其他行業前一年營業額一億元以上者但經常僱用員工數未滿100人：員工數" & pAD16 & "人"
      Case "5"
         PUB_AD15ToText = "依法辦理公司登記或商業登記，實收資本額在新臺幣1億元以下：" & pAD16 & "元"
      Case "6"
         PUB_AD15ToText = "經常僱用員工數未滿200人之事業：員工數" & pAD16 & "人"
   End Select
End Function

'Add by Sindy 2021/6/28
'設定國外潛在客戶類別選項
Public Sub PUB_SetComboPCU11(oCombo As Object, strComboVal As String, Optional bolAddBlank As Boolean = False)
Dim ii As Integer
   
   If strComboVal = "" Then 'Modify By Sindy 2021/6/30 加註:純設定下拉式選單的項目
      oCombo.Clear
      If bolAddBlank = True Then
         oCombo.AddItem ""
      End If
      oCombo.AddItem "1  廠商"
      oCombo.AddItem "2  事務所"
      oCombo.AddItem "3  個人"
      oCombo.AddItem "4  平台"
      oCombo.AddItem "5  供應商"
      oCombo.AddItem "6  媒體"
      oCombo.AddItem "7  協會"
      oCombo.AddItem "8  其他"
      
      oCombo.ListIndex = -1
      
   Else 'Modify By Sindy 2021/6/30 加註:移至下拉式選單的欄位值
      For ii = 0 To oCombo.ListCount - 1
         If Left(oCombo.List(ii), 1) = strComboVal Or Trim(Mid(oCombo.List(ii), 3)) = strComboVal Then
            oCombo.ListIndex = ii
            Exit Sub
         End If
      Next ii
   End If
End Sub

'Added by Lydia 2020/10/07 比對兩個字串內容是否相同(判斷法務案的案件屬性)
Public Function PUB_ChkTwoStrLst(ByVal pStr01 As String, ByVal pStr02 As String, Optional ByVal pSignal As String = ",") As Boolean
'pSingal 分隔符號: 預設,
Dim pArray As Variant
Dim pFind As Boolean
Dim intJJ As Integer
    
     If pStr01 <> pStr02 Then
          '正向
          pArray = Split(pStr01, pSignal)
          For intJJ = 0 To UBound(pArray)
              If InStr(pStr02 & pSignal, pArray(intJJ) & pSignal) = 0 Then
                   GoTo ExitFind
              End If
          Next intJJ
          '反向
          pArray = Empty
          pArray = Split(pStr02, pSignal)
          For intJJ = 0 To UBound(pArray)
              If InStr(pStr01 & pSignal, pArray(intJJ) & pSignal) = 0 Then
                   GoTo ExitFind
              End If
          Next intJJ
     End If
     PUB_ChkTwoStrLst = True
     Exit Function
     
ExitFind:
     PUB_ChkTwoStrLst = False
End Function

'Modify By Sindy 2014/8/4 因為此作業Account和Promoter等都有呼叫到,以防在Account需要加入一堆Form,因此抽出來至Func
'不可以放basFunction因AutoBatchDay有此Func
Public Sub frm210132_SubPubShowNextData(cmdState As Integer, ByRef oForm As Form)
Dim i As Integer, j As Integer
Dim StrTag As String

   Select Case cmdState
      Case 3 '案件基本資料
         With oForm
            .Enabled = False
            For i = 1 To .grdDataList.Rows - 1
               .grdDataList.col = 0
               .grdDataList.row = i
               If Trim(.grdDataList.Text) = "V" Then
                  Dim Str01 As String
                  .grdDataList.col = 0
                  .grdDataList.Text = ""
                  For j = 0 To .grdDataList.Cols - 1
                      .grdDataList.col = j
                      .grdDataList.CellBackColor = QBColor(15)
                  Next j
                  'Modified by Morgan 2011/12/23 調整欄位順序--辜
                  'grdDataList.col = 2
                  .grdDataList.col = 4 '3
                  Str01 = SystemNumber(.grdDataList, 1)
                  If Mid(UCase(Str01), 1, 1) = "N" Then
                      Str01 = Mid(Str01, 2, 3)
                  End If
                  If Not IsNull(.grdDataList.Text) Then
                      If fnSaveParentForm(oForm) = False Then
                          .Enabled = True
                          Exit Sub
                      End If
                      Select Case Pub_RplStr(Str01)
                          Case "CFP", "FCP", "P"   '專利
                                Screen.MousePointer = vbHourglass
                                frm100101_3.Show
                                frm100101_3.Tag = Pub_RplStr(.grdDataList.Text)
                                frm100101_3.StrMenu
                                Screen.MousePointer = vbDefault
                          Case "CFT", "FCT", "T", "TF"   '商標
                                Screen.MousePointer = vbHourglass
                                frm100101_4.Show
                                frm100101_4.Tag = Pub_RplStr(.grdDataList.Text)
                                frm100101_4.StrMenu
                                Screen.MousePointer = vbDefault
                          'Modify By Sindy 2009/07/24 增加LIN系統類別
                          'modify by sonia 2019/7/31 +ACS系統類別
                          Case "CFL", "FCL", "L", "LIN", "ACS"         '法務
                                Screen.MousePointer = vbHourglass
                                frm100101_5.Show
                                frm100101_5.Tag = Pub_RplStr(.grdDataList.Text)
                                frm100101_5.StrMenu
                                Screen.MousePointer = vbDefault
                          Case "LA"            '顧問
                                Screen.MousePointer = vbHourglass
                                frm100101_6.Show
                                frm100101_6.Tag = Pub_RplStr(.grdDataList.Text)
                                frm100101_6.StrMenu
                                Screen.MousePointer = vbDefault
                          Case Else                  '服務
                               Select Case Pub_RplStr(Str01)
                                   Case "TB"    '條碼
                                      Screen.MousePointer = vbHourglass
                                      frm100101_7.Show
                                      frm100101_7.Tag = Pub_RplStr(.grdDataList.Text)
                                      frm100101_7.StrMenu
                                      Screen.MousePointer = vbDefault
                                   Case "TM"
                                      Screen.MousePointer = vbHourglass
                                      frm100101_8.Show
                                      frm100101_8.Tag = Pub_RplStr(.grdDataList.Text)
                                      frm100101_8.StrMenu
                                      Screen.MousePointer = vbDefault
                                   Case "TD"
                                      Screen.MousePointer = vbHourglass
                                      frm100101_9.Show
                                      frm100101_9.Tag = Pub_RplStr(.grdDataList.Text)
                                      frm100101_9.StrMenu
                                      Screen.MousePointer = vbDefault
                                   Case "TC", "CFC"
                                      Screen.MousePointer = vbHourglass
                                      frm100101_A.Show
                                      frm100101_A.Tag = Pub_RplStr(.grdDataList.Text)
                                      frm100101_A.StrMenu
                                      Screen.MousePointer = vbDefault
                                   Case Else
                                      Screen.MousePointer = vbHourglass
                                      frm100101_B.Show
                                      frm100101_B.Tag = Pub_RplStr(.grdDataList.Text)
                                      frm100101_B.StrMenu
                                      Screen.MousePointer = vbDefault
                                End Select
                      End Select
                  End If
                  .Enabled = True
                  Exit Sub
               End If
            Next i
            .Enabled = True
         End With
         
      Case 4 '案件進度
         With oForm
            .Enabled = False
            StrTag = ""
            For i = 1 To .grdDataList.Rows - 1
               .grdDataList.col = 0
               .grdDataList.row = i
               If Trim(.grdDataList.Text) = "V" Then
                  .grdDataList.col = 0
                  .grdDataList.Text = ""
                  For j = 0 To .grdDataList.Cols - 1
                     .grdDataList.col = j
                     .grdDataList.CellBackColor = QBColor(15)
                  Next j
                  'Modified by Morgan 2011/12/23 調整欄位順序--辜
                  'grdDataList.col = 2
                  .grdDataList.col = 4 '3
                  If Not IsNull(.grdDataList.Text) Then
                     If fnSaveParentForm(oForm) = False Then
                        .Enabled = True
                        Exit Sub
                     End If
                     Screen.MousePointer = vbHourglass
                     frm100101_2.Show
                     frm100101_2.Tag = Pub_RplStr(.grdDataList.Text)
                     frm100101_2.StrMenu
                     Screen.MousePointer = vbDefault
                     .Enabled = True
                     Exit Sub
                  End If
               End If
            Next i
            .Enabled = True
         End With
      Case Else
   End Select
End Sub

'Added by Lydia 2021/08/30 各系統之分案作業和內部收文作業：勾選下一程序的期限，且該收文的案件性質與下一程序的案件性質不同，請SHOW訊息提醒
Public Function Pub_CheckNpTheSameShow(ByVal nCP01 As String, ByVal nCP10 As String, ByVal nNP07 As String) As Boolean
'Memo: 有關美專CFP28478的期限因收文、轉案而導致期限管制被沖掉,請協助在分案時增加控管,分案時若有去勾選下一程序的期限,且該收文的案件性質與下一程序的案件性質不相,請SHOW訊息
'nCP01: 收文之系統別、nCP10: 收文之案件性質
'nNP07: 下一程序的案件性質
    
    If nCP01 = "" Or nCP10 = "" Or nNP07 = "" Or nCP10 = nNP07 Then  '無資料 or 相同性質
        Pub_CheckNpTheSameShow = True
        Exit Function
    End If
    '排除分案或內部收文的案件性質為延期
    If (nCP01 = "ACS" And nCP10 = "205") Or (InStr(nCP01, "T") > 0 And nCP10 = "303") Or (InStr(nCP01, "P") > 0 And nCP10 = "404") Then
        Pub_CheckNpTheSameShow = True
        Exit Function
    End If
    
    If nCP10 <> nNP07 Then
        If MsgBox("所勾選下一程序之案件性質與收文之案件性質不同，請再確認是否要沖掉此道下一程序之期限？", vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
            Pub_CheckNpTheSameShow = False
        Else
            Pub_CheckNpTheSameShow = True
        End If
        Exit Function
    End If
End Function

'Added by Lydia 2022/05/23 法律所案源：取得案源類別、發文規費、email加註
Public Function PUB_GetLosCP84(ByVal pLOS15 As String, ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, _
         Optional ByVal pLimit As String, Optional ByRef pLOS02 As String, Optional ByRef pMemo As String) As String
'pLOS15 : 案源單號
'pCP01~pCP04 : 傳入案號
'pLimit：限制案源類別
'pLOS02：回傳案源類別
'pMemo：email加註法務案資料
Dim strQ1 As String, intQ As Integer
Dim rsQD As New ADODB.Recordset
Dim strB1 As String, strB2 As String
    
    PUB_GetLosCP84 = "0"
    pMemo = ""
    pLOS02 = ""
    strQ1 = "select d.cp01,d.cp02,d.cp03,d.cp04,los06,los02,los15,sum(d.cp16) cp16t,sum(d.cp17) cp17t,sum(d.cp16-nvl(d.cp17,0)) /1000 SFee,sum(nvl(d.cp84,0)) cp84t " & _
                "from lawofficesource,caseprogress c,caseprogress d " & _
                "where los15='" & pLOS15 & "' and c.cp09(+)=los06 and d.cp162(+)=los15 and d.cp01=c.cp01 " & _
                "group by d.cp01,d.cp02,d.cp03,d.cp04,los06,los02,los15 "
    intQ = 1
    Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
    If intQ = 1 Then
        pLOS02 = "" & rsQD.Fields("los02")
        If (pLimit <> "" And pLOS02 = pLimit) Or pLimit = "" Then  '限制案源類別
           PUB_GetLosCP84 = Val("" & rsQD.Fields("cp17t")) - Val("" & rsQD.Fields("cp84t"))
           '若有銷帳則要扣除銷帳規費
           If GetCP77Detail(rsQD.Fields("los06"), strB1, strB2) = True Then
              PUB_GetLosCP84 = Val(PUB_GetLosCP84) - Val(strB2)
           End If
           pMemo = "法律所：" & rsQD.Fields("cp01") & "-" & rsQD.Fields("cp02") & "-" & rsQD.Fields("cp03") & "-" & rsQD.Fields("cp04")
           '扣除PT已繳規費
           If Val(PUB_GetLosCP84) > 0 And pCP01 <> rsQD.Fields("cp01") Then
              strQ1 = "select sum(nvl(cp84,0)) cp84t from caseprogress where cp162='" & pLOS15 & "' and cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp159=0 "
              intQ = 1
              Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
              If intQ = 1 Then
                   PUB_GetLosCP84 = Val(PUB_GetLosCP84) - Val("" & rsQD.Fields("cp84t"))
              End If
           End If
           pMemo = pMemo & "，收文規費：" & Format(PUB_GetLosCP84, DDollar)
        End If
    End If
    Set rsQD = Nothing
End Function

'Added by Lydia 2022/07/04 FCP和FMP案之一案兩請僅其中一案更代時，彈提醒
Public Function PUB_ChkFCforChange(ByVal mPA01 As String, ByVal mPA02 As String, ByVal mPA03 As String, ByVal mPA04 As String) As Boolean
Dim strTmp1 As String, strTo As String, strCC As String, strSub As String
Dim strCase(1 To 4) As String, strDCase(1 To 4) As String

   PUB_ChkFCforChange = True  '預設不用檢查
   If mPA01 = "P" Then 'FMP案
       If PUB_ChkIsFMP(mPA01, mPA02, mPA03, mPA04) = False Then
           Exit Function
       End If
   ElseIf mPA01 <> "FCP" Then
       Exit Function
   End If
   
   strCase(1) = mPA01: strCase(2) = mPA02: strCase(3) = mPA03: strCase(4) = mPA04
   If PUB_IsDualApply(strCase, strDCase, , , , , , True) = True Then
       strTmp1 = mPA01 & "-" & mPA02 & IIf(mPA03 & mPA04 <> "000", "-" & mPA03 & "-" & mPA04, "") & "為一案兩請，是否已確認發明案 及/或 新型案更代事項？" & vbCrLf & _
                "另一案件：" & strDCase(1) & "-" & strDCase(2) & IIf(strDCase(3) & strDCase(4) <> "000", "-" & strDCase(3) & "-" & strDCase(4), "") & vbCrLf & vbCrLf & _
                "選""是""將繼續更代作業，選""否""將發email通知承辦人員。"
       If MsgBox(strTmp1, vbExclamation + vbYesNo + vbDefaultButton2) = vbNo Then
           '收件者: 承辦
           strTo = PUB_GetFCPSalesNo(strCase(1), strCase(2), strCase(3), strCase(4))
           If strTo <> "" Then
               '副本收受者: 承辦主管、程序、程序主管、backup
               strCC = PUB_GetFCPProSup(strTo)
               strTmp1 = PUB_GetFCPHandler(strCase(1), strCase(2), strCase(3), strCase(4))
               If strTmp1 <> "" Then strSub = PUB_GetFCPProSup(strTmp1)
               strCC = strCC & ";" & strTmp1 & ";" & strSub & ";backup"
               strCC = Replace(strCC, ";;", ";")
               
               strSub = "一案兩請案件僅其中一案將更代：請去信確認另案是否一併更代；若已確認，請回報程序續行更代作業。Our Ref: " & mPA01 & "-" & mPA02 & IIf(mPA03 & mPA04 <> "000", "-" & mPA03 & "-" & mPA04, "") & " [INCOM.]"
               strTmp1 = mPA01 & "-" & mPA02 & IIf(mPA03 & mPA04 <> "000", "-" & mPA03 & "-" & mPA04, "") & "為一案兩請案件僅其中一案將更代：" & vbCrLf & _
                              "請去信確認另案" & strDCase(1) & "-" & strDCase(2) & IIf(strDCase(3) & strDCase(4) <> "000", "-" & strDCase(3) & "-" & strDCase(4), "") & "是否一併更代；" & vbCrLf & _
                              "若已確認，請回報程序續行更代作業。"
               strTmp1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                                 " values( '" & strUserNum & "' , '" & strTo & "' , '" & strSrvDate(1) & "' , '" & Format(ServerTime, "000000") & "' , '" & strSub & "' , '" & strTmp1 & "', '" & strCC & "')"
               cnnConnection.Execute strTmp1
           End If
           PUB_ChkFCforChange = False
       End If
   End If
End Function

'Added by Lydia 2022/09/21 外商承辦通知主管/人員
Public Function PUB_GetF11ToMan(ByVal pUserNo As String) As String
Dim intA As Integer, strA1 As String
Dim rsAD As New ADODB.Recordset
    
    If pUserNo = "" Then Exit Function
    
    strA1 = "select st01,st02,st03,st04,st16 from staff where st01='" & pUserNo & "' and st03='F11' "
    
    intA = 1
    Set rsAD = ClsLawReadRstMsg(intA, strA1)
    If intA = 1 Then
       '外商: 第2組(英文), 第4組(日文)
       If "" & rsAD.Fields("st16") = "4" Then
           PUB_GetF11ToMan = Pub_GetSpecMan("外商日文組通知主管")
       Else
           PUB_GetF11ToMan = pUserNo
       End If
    End If
    Set rsAD = Nothing
End Function

'Added by Lydia 2022/12/08 儲存出庭律師資料檔: 配合輸入出庭費,改成先存暫存檔再寫入正式Table
'Modified by Lydia 2023/08/14 + 是否只讀取資料 bolOnlyRead
Public Function PUB_SaveCaseLawer(ByVal pCP09 As String, ByVal pSeqNo As String, Optional ByRef pNowNo As String, Optional ByVal bolAuto As Boolean, Optional ByVal bolOnlyRead As Boolean) As Boolean
Dim strR1 As String, intR As Integer
Dim rsRD As New ADODB.Recordset
Dim strEx01 As String

On Error GoTo ErrHandle

   pNowNo = ""
   If pCP09 <> "" And pSeqNo <> "" Then
      If bolAuto = True Then  '更新收文號
         strR1 = "UPDATE RDATAFACTORY SET R004='" & pCP09 & "' WHERE id = '" & strUserNum & "' and formname='frm071018' and seqno='" & pSeqNo & "' "
         cnnConnection.Execute strR1
      End If
      strR1 = "SELECT CP14,R002 AS 出庭律師, R003 AS 出庭費, R004 AS CL01, R005 AS CL02, R006 AS ORD1,ROWSEQ,CL02 AS OLD02, CL03 AS OLD03 " & _
                 "FROM RDATAFACTORY,CASEPROGRESS,CASELAWER WHERE id = '" & strUserNum & "' and formname='frm071018' and seqno='" & pSeqNo & "' AND R004=CP09(+) " & _
                 "AND R004=CL01(+) AND R005=CL02(+) ORDER BY R005 asc "
      intR = 1
      Set rsRD = ClsLawReadRstMsg(intR, strR1)
      '增加／修改記錄
      If intR = 1 Then
          'Added by Lydia 2023/08/14 只讀取資料
          If bolOnlyRead = True Then
             PUB_SaveCaseLawer = True
             Exit Function
          End If
          'end 2023/08/14
          rsRD.MoveFirst
          Do While Not rsRD.EOF
              strEx01 = ""
              If "" & rsRD.Fields("CL02") <> "" & rsRD.Fields("OLD02") Or Val("" & rsRD.Fields("出庭費")) <> Val("" & rsRD.Fields("OLD03")) Then
                 If "" & rsRD.Fields("old02") <> "" Then
                    strEx01 = "Update CaseLawer Set CL03=" & CNULL("" & rsRD.Fields("出庭費"), True) & " Where CL01='" & pCP09 & "' and CL02='" & rsRD.Fields("old02") & "' "
                    pNowNo = pNowNo & rsRD.Fields("CL02") & "|出庭費：" & Val("" & rsRD.Fields("old03")) & "=>" & Val("" & rsRD.Fields("出庭費")) & " ,"
                 Else
                    strEx01 = "Insert Into CaseLawer (CL01,CL02,CL03) Values ('" & pCP09 & "', '" & IIf(Val("" & rsRD.Fields("ord1")) < 2, "" & rsRD.Fields("CP14"), "" & rsRD.Fields("CL02")) & "' ," & CNULL("" & rsRD.Fields("出庭費"), True) & " ) "
                    pNowNo = pNowNo & rsRD.Fields("CL02") & ","
                 End If
                 If strEx01 <> "" Then
                    'Modified by Lydia 2025/07/22 傳入收文號
                    'Pub_SeekTbLog strEx01
                    Pub_SeekTbLog strEx01, , , , , pCP09
                    cnnConnection.Execute strEx01
                 End If
              Else
                 pNowNo = pNowNo & rsRD.Fields("CL02") & ","
              End If
              rsRD.MoveNext
          Loop
      'Added by Lydia 2023/08/14 只讀取資料
      Else
          If bolOnlyRead = True Then
             PUB_SaveCaseLawer = False
             Exit Function
          End If
      'end 2023/08/14
      End If
      
      '刪除記錄
      strR1 = "select * from caselawer where cl01='" & pCP09 & "' and (cl01,cl02) not in (select R004,R005 from rdatafactory where id = '" & strUserNum & "' and formname='frm071018' and seqno='" & pSeqNo & "' ) "
      intR = 1
      Set rsRD = ClsLawReadRstMsg(intR, strR1)
      If intR = 1 Then
          rsRD.MoveFirst
          Do While Not rsRD.EOF
              strEx01 = "Delete From CaseLawer Where CL01='" & rsRD.Fields("CL01") & "' And CL02='" & rsRD.Fields("CL02") & "' "
              'Modified by Lydia 2025/07/22 傳入收文號
              'Pub_SeekTbLog strEx01
              Pub_SeekTbLog strEx01, , , , , pCP09
              cnnConnection.Execute strEx01
              rsRD.MoveNext
          Loop
      End If
   End If
   
   Set rsRD = Nothing
   If bolOnlyRead = False Then 'Added by Lydia 2023/08/14
      PUB_SaveCaseLawer = True
   End If  'Added by Lydia 2023/08/14
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical, "更新失敗=>出庭律師"
   End If
End Function

'Added by Lydia 2023/08/14 須限制案件性質表規費科目(CPM12)為220113的才可點出庭律師。
Public Function Pub_ChkPtyCL(ByVal pCP01 As String, ByVal pCP10 As String) As Boolean
Dim intQ As Integer, strQ1 As String
Dim rsQD As New ADODB.Recordset
   
   Pub_ChkPtyCL = False
   strQ1 = "select cpm12 from casepropertymap where cpm01='" & pCP01 & "' and cpm02='" & pCP10 & "' "
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
   'Modified by Lydia 2024/04/23 改成常數
   'If InStr(rsQD.Fields("cpm12") & ",", "220113") > 0 Then
   If InStr("," & CaseLawerPtyList & ",", "," & rsQD.Fields("cpm12") & ",") > 0 Then
       Pub_ChkPtyCL = True
   End If
   'Added by Lydai 2024/09/30 (113/11/01上線) 增加特殊案件性質可輸出庭費，但也可能不輸
   If Pub_ChkPtyCL = False And InStr(pCP01, "L") > 0 Then
      strQ1 = "select ','||oman||',' from setspecman where ocode='出庭費特殊性質' and instr(','||oman||',','," & pCP10 & ",') > 0 "
      intQ = 1
      Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
      If intQ = 1 Then
         Pub_ChkPtyCL = True
      End If
   End If
   'end 2024/09/30
   Set rsQD = Nothing
End Function

'Add By Sindy 2023/2/24
'傳入值:收文號、變更種類
'變更種類:strChgTy=1.變更申請人之地址
'                  2.變更申請人之代理人
'                  3.變更申請人之代表人
'                  4.變更申請人之姓名或名稱
'                  5.變更申請人之國籍
'回傳值:True-代表有讀取到資料
'strVal:回傳組出來的欄位值
Public Function PUB_GetChangeEvent(strCP09 As String, strChgTy As Integer, Optional ByRef strVal As String) As Boolean
Dim intQ As Integer, strCon1 As String, rsQ1 As New ADODB.Recordset 'Added by Lydia 2024/03/29

   PUB_GetChangeEvent = False
   strVal = ""
   If strCP09 = "" Then Exit Function
   
   If strChgTy = 1 Then
      strCon1 = "select * From ChangeEvent" & _
                  " where ce01='" & strCP09 & "'" & _
                  " and (ce23 is not null or ce24 is not null or ce25 is not null or ce26 is not null or ce27 is not null" & _
                       " or ce28 is not null or ce29 is not null or ce30 is not null or ce31 is not null or ce32 is not null" & _
                       " or ce33 is not null or ce34 is not null or ce35 is not null or ce36 is not null or ce37 is not null)"
   ElseIf strChgTy = 2 Then '2.變更申請人之代理人
      strCon1 = "select * From ChangeEvent" & _
                  " where ce01='" & strCP09 & "' and ce55 is not null"
   ElseIf strChgTy = 3 Then
      strCon1 = "select * From ChangeEvent" & _
                  " where ce01='" & strCP09 & "'" & _
                  " and (ce10 is not null or ce11 is not null or ce12 is not null or ce13 is not null or ce14 is not null" & _
                       " or ce15 is not null or ce68 is not null or ce69 is not null or ce70 is not null or ce71 is not null" & _
                       " or ce72 is not null or ce73 is not null or ce74 is not null or ce75 is not null or ce76 is not null" & _
                       " or ce77 is not null or ce78 is not null or ce79 is not null or ce80 is not null or ce81 is not null" & _
                       " or ce82 is not null or ce83 is not null or ce84 is not null or ce85 is not null or ce86 is not null" & _
                       " or ce87 is not null or ce88 is not null or ce89 is not null or ce90 is not null or ce91 is not null)"
   ElseIf strChgTy = 4 Then '4.變更申請人之姓名或名稱
      strCon1 = "select * From ChangeEvent" & _
                  " where ce01='" & strCP09 & "'" & _
                  " and (ce04 is not null or ce05 is not null or ce06 is not null or ce07 is not null or ce08 is not null" & _
                       " or ce17 is not null or ce18 is not null or ce19 is not null or ce20 is not null or ce21 is not null)"
   ElseIf strChgTy = 5 Then
      strCon1 = "select * From ChangeEvent" & _
                  " where ce01='" & strCP09 & "'" & _
                  " and ce61 is not null and instr(ce61,'申請人國籍')>0"
   Else
      Exit Function
   End If
   intQ = 1
   Set rsQ1 = ClsLawReadRstMsg(intQ, strCon1)
   If intQ = 1 Then
      PUB_GetChangeEvent = True
      If strChgTy = 3 Then '變更申請人之代表人
         strVal = "" & rsQ1.Fields("ce10") & "@" & "" & rsQ1.Fields("ce11") & "@" & "" & rsQ1.Fields("ce12") & "@" & _
                  "" & rsQ1.Fields("ce13") & "@" & "" & rsQ1.Fields("ce14") & "@" & "" & rsQ1.Fields("ce15") & "@" & _
                  "" & rsQ1.Fields("ce68") & "@" & "" & rsQ1.Fields("ce69") & "@" & "" & rsQ1.Fields("ce70") & "@" & _
                  "" & rsQ1.Fields("ce71") & "@" & "" & rsQ1.Fields("ce72") & "@" & "" & rsQ1.Fields("ce73") & "@" & _
                  "" & rsQ1.Fields("ce74") & "@" & "" & rsQ1.Fields("ce75") & "@" & "" & rsQ1.Fields("ce76") & "@" & _
                  "" & rsQ1.Fields("ce77") & "@" & "" & rsQ1.Fields("ce78") & "@" & "" & rsQ1.Fields("ce79") & "@" & _
                  "" & rsQ1.Fields("ce80") & "@" & "" & rsQ1.Fields("ce81") & "@" & "" & rsQ1.Fields("ce82") & "@" & _
                  "" & rsQ1.Fields("ce83") & "@" & "" & rsQ1.Fields("ce84") & "@" & "" & rsQ1.Fields("ce85") & "@" & _
                  "" & rsQ1.Fields("ce86") & "@" & "" & rsQ1.Fields("ce87") & "@" & "" & rsQ1.Fields("ce88") & "@" & _
                  "" & rsQ1.Fields("ce89") & "@" & "" & rsQ1.Fields("ce90") & "@" & "" & rsQ1.Fields("ce91") & "@"
      End If
   End If
   Set rsQ1 = Nothing 'Added by Lydia 2024/03/29
End Function

'Added by Lydia 2023/03/16 法律所案源：P、FCP、T、FCT的延期發文時，若延期的相關總收文號為B2案源時，同時新增法務案之內部收文39延期
Public Sub PUB_InsertLosBCP(ByVal pLOS15 As String, pTCP27 As String, ByVal pTCP06 As String, ByVal pTCP07 As String)
Dim strQ1 As String, intQ As Integer
Dim strB1 As String
Dim rsQD As New ADODB.Recordset
     
    '相關總收文號掛法務案的總收文號LOS06，(B類收文)延期上發文日同時EMAIL給承辦人CP14及法務協辦人員CP29
    strQ1 = "select y.cp01,y.cp02,y.cp03,y.cp04,y.cp06,y.cp07, y.cp09,y.cp12,y.cp13,y.cp14,y.cp29,y.cp10,p.cpm03," & _
                "z.cp13 as cp13pt, z.cp01||'-'||z.cp02||'-'||z.cp03||'-'||z.cp04 as ptcaseno, q.cpm03 as pcpm03 " & _
                "from LAWOFFICESOURCE, caseprogress x, caseprogress y, caseprogress z, casepropertymap p, casepropertymap q " & _
                "where los15='" & pLOS15 & "' and los06=x.cp09(+) and x.cp01=y.cp01(+) and x.cp02=y.cp02(+) and x.cp03=y.cp03(+) and x.cp04=y.cp04(+) " & _
                "and y.cp162=los15 and y.cp01=p.cpm01(+) and y.cp10=p.cpm02(+) and y.cp159=0 and y.cp158=0 and los01=z.cp09 " & _
                "and z.cp01=q.cpm01(+) and z.cp10=q.cpm02(+) "
    intQ = 1
    Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
    If intQ = 1 Then
        strB1 = AutoNo("B", 6)
        '有承辦人才上發文日
        strSql = "INSERT INTO CaseProgress (CP01,CP02,CP03,CP04,CP05,CP06,CP07,CP09,CP10,CP12,CP13,CP14,CP27,CP29,CP43,cp20,cp26,cp32) " & _
                   "VALUES ('" & rsQD.Fields("cp01") & "','" & rsQD.Fields("cp02") & "','" & rsQD.Fields("cp03") & "','" & rsQD.Fields("cp04") & "'," & pTCP27 & "," & _
                       CNULL("" & rsQD.Fields("cp06"), True) & "," & CNULL("" & rsQD.Fields("cp07"), True) & ",'" & strB1 & "','39','" & rsQD.Fields("cp12") & "','" & rsQD.Fields("cp13") & "','" & rsQD.Fields("cp14") & "'," & _
                       IIf("" & rsQD.Fields("cp14") <> "", strSrvDate(1), "null") & ",'" & rsQD.Fields("cp29") & "','" & rsQD.Fields("cp09") & "','N','N','N') "
        cnnConnection.Execute strSql
        If pTCP06 <> "" Or pTCP07 <> "" Then
            strSql = "Update CaseProgress Set " & Mid(IIf(pTCP06 <> "", ", CP06='" & pTCP06 & "' ", "") & IIf(pTCP07 <> "", ", CP07='" & pTCP07 & "' ", ""), 2) & _
                         " Where CP09 = '" & rsQD.Fields("cp09") & "' and cp158=0 "
            cnnConnection.Execute strSql
        End If
        '收件者：承辦人CP14及法務協辦人員CP29
        strB1 = ""
        If "" & rsQD.Fields("cp14") & rsQD.Fields("cp29") <> "" Then
           strB1 = rsQD.Fields("cp14") & IIf("" & rsQD.Fields("cp29") <> "", ";" & rsQD.Fields("cp29"), "")
        Else
           strB1 = "" & rsQD.Fields("cp13")
        End If
        If strB1 <> "" Then
             strQ1 = rsQD.Fields("ptcaseno") & "已於" & ChangeTStringToTDateString(TransDate(pTCP27, 1)) & "提出" & rsQD.Fields("pcpm03") & "之延期" & vbCrLf & _
                        "延期後本所期限：" & ChangeTStringToTDateString(TransDate(pTCP06, 1)) & vbCrLf & _
                        "延期後法定期限：" & ChangeTStringToTDateString(TransDate(pTCP07, 1)) & vbCrLf
             strQ1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                  " values( '" & strUserNum & "','" & strB1 & "',to_char(sysdate,'yyyymmdd')" & _
                  ",to_char(sysdate,'hh24miss'),'" & ChgSQL(rsQD.Fields("cp01") & "-" & rsQD.Fields("cp02") & "-" & rsQD.Fields("cp03") & "-" & rsQD.Fields("cp04") & "產生延期記錄，同時更新期限通知！") & "','" & ChgSQL(strQ1) & "')"
             cnnConnection.Execute strQ1
        End If
    End If
    Set rsQD = Nothing
End Sub

'Added by Lydia 2023/05/17 寰華案無期限之官方來函，系統自動發Mail   'Memo by Lydia 2025/08/19 無期限的承辦人為外專承辦
'Modified by Lydia 2023/06/02 +pCP09來函新增的收文號, pRCP10相關收文號案件性質;
'Memo by Lydia 2023/10/31 將「pCP09來函新增的收文號」改成必需傳入
Public Function Pub_SetFMP2toCMail(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pCP10 As String, ByVal pCP14 As String, ByVal pCP09 As String, Optional ByVal pRCP10 As String) As Boolean
Dim strTo As String, strCC As String, strB1 As String
Dim strSub As String, strContent As String, intB As Integer
Dim strKind As String 'Added by Lydia 2023/06/02
Dim bolChk As Boolean 'Added by Lydia 2024/01/31
Dim strQ1 As String, intQ As Integer, rsQD As New ADODB.Recordset
    Pub_SetFMP2toCMail = False
    'Modified by Lydia 2025/10/16 debug:2024/06/16 +視為撤回1610 ---- Winfrey
    If pCP01 <> "P" Or InStr("1004,1406,1204,1213,1214,1603,1912,1604,1234,1813,1001,1610", pCP10) = 0 Then Exit Function
  
    strCC = strUserNum '程序操作人員
    strB1 = PUB_GetFCPHandler(pCP01, pCP02, pCP03, pCP04)
    If strB1 <> strUserNum Then strCC = strCC & ";" & strB1 '程序操作人員非管制人員
    strCC = strCC & ";backup"
    Call ClsPDGetCaseProperty(pCP01, pCP10, strSub, True)
    
    'Modified by Lydia 2023/06/02 核准區分相關收文號的案件性質
    'Select Case pCP10
    If pCP10 <> "1001" Then
       strKind = pCP10
    Else
       If InStr("908,701,702,401", pRCP10) > 0 Then '代辦退費908,讓與701,合併702,變更401
          strKind = "B"
       ElseIf InStr("431,107", pRCP10) > 0 Then 'PPH 431, 復審申請107
          strKind = "C"
       End If
    End If
    If strKind = "B" Or strKind = "C" Then '核准抓相關收文號的案件性質
       strSub = strSub & PUB_GetRelateCasePropertyName(pCP09, "1")
    End If
    Select Case strKind
    'end 2023/06/02
       'Modified by Lydia 2023/06/02 + strKind = B
       'Modified by Lydia 2024/06/16 +視為撤回1610 ---- Winfrey
       Case "1004", "1406", "1204", "1213", "1214", "1603", "1912", "1604", "1610", "B" '1004延期受理,1406實審屆滿前通知 ; 1204進入實審通知,1213初步審查合格通知,1214初步審查及進入實審,1603專利證書,1912通知已轉他所,1604專利權消滅,1610視為撤回
            '收件者: 智權人員
            'CC：智權人員主管、程序操作人員、backup
            strTo = PUB_GetFCPSalesNo(pCP01, pCP02, pCP03, pCP04)
            strB1 = PUB_GetFCPProSup(strTo)
            If strB1 <> "" Then strCC = strB1 & ";" & strCC
            If InStr("1004,1406", pCP10) > 0 Then
               '主旨：【FMP寰華案-轉發官方來函】案件性質 Our Ref: P-xxxxxx [INCOM.案件性質之編號]
               strSub = "【FMP寰華案-轉發官方來函】" & strSub & " Our Ref: " & pCP01 & "-" & pCP02 & IIf(pCP03 <> "0", "-" & pCP03, "") & IIf(pCP04 <> "00", "-" & pCP04, "") & "[INCOM." & pCP10 & "]"
            Else
               '主旨：【FMP寰華案-已收初審合格通知書#請帶案件性質之中文名稱# 】 請承辦報告 Our Ref: P-xxxxxx [INCOM.案件性質之編號]
               strSub = "【FMP寰華案-" & strSub & "】 請承辦報告 Our Ref: " & pCP01 & "-" & pCP02 & IIf(pCP03 <> "0", "-" & pCP03, "") & IIf(pCP04 <> "00", "-" & pCP04, "") & "[INCOM." & pCP10 & "]"
            End If
            strContent = "通知函、官方來函已匯入卷宗區"
            If strTo <> "" Then
               'Added by Lydia 2023/11/08 增加判斷上一道工程師案件性質是否已經請款，若有通知則不用發無期限Email
               'Modified by Lydia 2023/12/13 1004延期受理,1406實審屆滿前通知=>因為承辦不用報告,所以主旨全部延用
               'If PUB_ChkFCPtoDNUPL(pCP01, pCP02, pCP03, pCP04, pCP10, pCP09, Mid(strSub, 1, InStr(strSub, "】"))) = False Then
               'Modified by Lydia 2024/01/31 寰華案key 1004延期受理時，請不要去判斷是否有無請款未完成
               'If PUB_ChkFCPtoDNUPL(pCP01, pCP02, pCP03, pCP04, pCP10, pCP09, IIf(InStr("1004,1406", pCP10) > 0, strSub, Mid(strSub, 1, InStr(strSub, "】")))) = False Then
               bolChk = False
               If pCP10 <> "1004" Then
                  bolChk = PUB_ChkFCPtoDNUPL(pCP01, pCP02, pCP03, pCP04, pCP10, pCP09, IIf(InStr("1004,1406", pCP10) > 0, strSub, Mid(strSub, 1, InStr(strSub, "】"))))
               End If
               'Modified by Lydia 2024/05/07 不用加發Email(PUB_ChkFCPtoDNUPL的範例3)，但是要發原本的通知函
               'If bolChk = False Then
               If bolChk = False Or InStr("1004,1406,1234,1813,C", strKind) > 0 Then
               'end 2024/01/31
                  strB1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09) " & _
                             "values('" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss') " & _
                             ",'" & ChgSQL(strSub) & "','" & ChgSQL(strContent) & "','" & strCC & "') "
                  cnnConnection.Execute strB1
               End If
               Pub_SetFMP2toCMail = True
            End If
       'Modified by Lydia 2023/06/02 + strKind = C
       Case "1234", "1813", "1001", "C" '1234復審受理通知,1813行政訴訟受理通知,1001核准-PPH
            '收件者: 承辦工程師
            'CC：承辦工程師主官、智權人員、程序操作人員、backup
            strTo = pCP14
            strB1 = PUB_GetFCPSalesNo(pCP01, pCP02, pCP03, pCP04)
            If strB1 <> "" Then strCC = strB1 & ";" & strCC
            strB1 = PUB_GetFCPEngSup(pCP14)
            If strB1 <> "" Then strCC = strB1 & ";" & strCC
            '主旨：【FMP寰華案-轉發官方來函】案件性質 Our Ref: P-xxxxx [INCOM.請案件性質之編號]
            strSub = "【FMP寰華案-轉發官方來函】" & strSub & " Our Ref: " & pCP01 & "-" & pCP02 & IIf(pCP03 <> "0", "-" & pCP03, "") & IIf(pCP04 <> "00", "-" & pCP04, "") & "[INCOM." & pCP10 & "]"
            strContent = "已收官方來函請見卷宗區" & vbCrLf & "若有須向代理人報告 , 請通知承辦收文告代"
            If strTo <> "" Then
               'Added by Lydia 2023/11/08 增加判斷上一道工程師案件性質是否已經請款，若有通知則不用發無期限Email
               'Modified by Lydia 2024/01/31 改用變數
               'If PUB_ChkFCPtoDNUPL(pCP01, pCP02, pCP03, pCP04, pCP10, pCP09, Mid(strSub, 1, InStr(strSub, "】"))) = False Then
               'Modified by Lydia 2024/04/10 + 傳入核准類型 strKind
               bolChk = PUB_ChkFCPtoDNUPL(pCP01, pCP02, pCP03, pCP04, pCP10, pCP09, Mid(strSub, 1, InStr(strSub, "】")), strKind)
               'Modified by Lydia 2024/05/07 不用加發Email(PUB_ChkFCPtoDNUPL的範例3)，但是要發原本的通知函
               'If bolChk = False Then
               If bolChk = False Or InStr("1004,1406,1234,1813,C", strKind) > 0 Then
               'end 2024/01/31
                  strB1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09) " & _
                             "values('" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss') " & _
                             ",'" & ChgSQL(strSub) & "','" & ChgSQL(strContent) & "','" & strCC & "') "
                  cnnConnection.Execute strB1
               End If
               Pub_SetFMP2toCMail = True
            End If
    End Select
    
    'Added by Lydia 2023/10/30 若有發通知Email請增加判斷上一道工程師案件性質是否已經請款（請排除已上不請款N）
    'Mark by Lydia 2023/11/08 增加判斷上一道工程師案件性質是否已經請款，若有通知則不用發無期限Email
    'If Pub_SetFMP2toCMail = True Then
    '   Sleep 100
    '   If PUB_ChkFCPtoDNUPL(pCP01, pCP02, pCP03, pCP04, pCP10, pCP09, Mid(strSub, 1, InStr(strSub, "】"))) = True Then
    '   End If
    'End If
    'end 2023/10/30
End Function

'Added by Lydia 2025/08/19 FCP、FMP（包含寰華）輸入C類來函時，去檢查上一道承辦人掛工程師，是否為未請款，若是，則發Mail通知工程師；
              '與Pub_SetFMP2toCMail的區別，工程師承辦為有期限，所以是不同的通知
Public Function PUB_ChkFCPtoCP14CP60(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pCP10 As String, ByVal pCP09 As String, ByVal pCP14 As String) As Boolean
'pCP09、pCP10: 輸入C類來函的收文號、案件性質
Dim intQ As Integer, rsQD As New ADODB.Recordset
Dim strNA16 As String, strA1 As String, strA2 As String
Dim strTo As String, strCC As String, strSubject As String, strContent As String


   PUB_ChkFCPtoCP14CP60 = False
   If pCP01 <> "P" And pCP01 <> "FCP" Then Exit Function
   If Len(pCP10) <> 4 Or pCP10 = "1001" Or pCP10 = "1008" Then Exit Function  '排除核准
   If pCP14 <> "" Then
      If PUB_GetST03(pCP14) <> "F21" Then
         Exit Function
      End If
   End If
   
On Error GoTo ErrHandle

   strA1 = "select cp09,cp60,cp14,cp10,decode(pa09,'000',nvl(cpm03,cpm04),nvl(cpm03,cpm04)) as cp10name,cp27,cp158 from caseprogress,patent,casepropertymap where cp09= (" & _
            "select max(cp09) mno from caseprogress,staff where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp158 > 0 and cp159=0 " & _
            "and cp14=st01(+) and st03='F21' and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' and nvl(cp20,'Y')<>'N' " & _
            "and cp05 = (select max(cp05) mdate from caseprogress, staff where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp158 > 0 and cp159=0 " & _
            "and cp14=st01(+) and st03='F21' and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' and nvl(cp20,'Y')<>'N' and nvl(cp43,'N') <> '" & pCP09 & "' )) " & _
            "and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) and cp01=cpm01(+) and cp10=cpm02(+) "
   strA1 = strA1 & " and cp60 is null " '未請款
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strA1)
   If intQ = 1 Then
      'Added by Lydia 2025/08/22 不通知的案件性質
      If InStr("421,", "" & rsQD.Fields("cp10")) > 0 Or (pCP10 = "1201" And InStr("407,408,", "" & rsQD.Fields("cp10")) > 0) Then
         '上一道程序：421申請技術報告
         '輸入來函+上一道程序：1201通知修正和407面詢/408請求面詢
      Else
      'end 2025/08/22
         If "" & rsQD.Fields("cp60") = "" Then
            '收件者: 工程師
            '副本收受者：工程師主管（主任+副理）、程序操作人員（若不是程序管制人員，則另加程序管制人員）
            strTo = "" & rsQD.Fields("cp14")
            If Mid(strTo, 4, 1) = "9" Then
               strTo = PUB_GetFCPEngSup(strTo, , , True)
            End If
            strCC = PUB_GetFCPEngSup(strTo, True)
            strNA16 = PUB_GetFCPHandler(pCP01, pCP02, pCP03, pCP04) '程序
            strCC = strCC & ";" & strNA16
            If strNA16 <> strUserNum Then
               strCC = strCC & ";" & strUserNum
            End If
   
            strSubject = "【已收官方來函】上一道程序尚未請款，請工程師儘速處理請款，以利後續流程Our Ref: " & pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 = "000", "", "-" & pCP03 & "-" & pCP04) & " [INCOM." & pCP10 & "]"
            strContent = "以下案件性質未請款，請儘速處理請款：" & vbCrLf & _
                         "發文日期" & String(3, " ") & "總收文號" & String(3, " ") & "案件性質" & vbCrLf & _
                         ChangeWStringToTDateString(rsQD.Fields("cp27")) & String(2, " ") & rsQD.Fields("cp09") & String(2, " ") & rsQD.Fields("cp10name")
            strA2 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                        " values ('" & strUserNum & "','" & strTo & "'," & strSrvDate(1) & ",to_char(sysdate,'hh24miss')" & _
                        ",'" & ChgSQL(strSubject) & "','" & ChgSQL(strContent) & "','" & strCC & "')"
            cnnConnection.Execute strA2, intQ
            PUB_ChkFCPtoCP14CP60 = True
         End If
      End If 'Added by Lydia 2025/08/22
   End If
   Set rsQD = Nothing
   
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical, "判斷上一道工程師案件性質是否已經請款"
   End If
End Function

'Added by Lydia 2023/05/19 外專-簡易聯絡單內容---從frm060104_3抽出，並且改成列印／回傳
Public Function PUB_FCPPrintContactSheetA4(ByVal bolPrint As Boolean, ByVal pCP09 As String, ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, _
                      ByVal pCP10 As String, Optional ByVal bolCall As Boolean = False, Optional ByRef mDate209210 As String, Optional ByRef mDateTF30 As String) As String
Dim iPage As Integer
Dim dblLineHeight As Double '行高
Dim m_dblTitleHeight As Double '抬頭
Dim m_dblTop As Double '上邊界
Dim m_dblLeft As Double '左邊界
Dim m_TBWidth As Double '表格寬
Dim intLine As Integer
Dim dblPrtX As Double
Dim dblPrtY As Double
Dim intFieldWidth
Dim strTemp As String
Dim strTmpA As String '折行-剩的字串
Dim bolFirst As Boolean
Dim m_Line3 As String 'Added by Lydia 2018/01/05
Dim strTmp(0 To 10) As String, intR As Integer
Dim rsRD As New ADODB.Recordset

    PUB_FCPPrintContactSheetA4 = "" 'Add By Sindy 2022/5/11
    strTmp(1) = 0
    'Added by Lydia 2020/02/10 外部呼叫: 預設變數
    If bolCall = True Then
       mDate209210 = ""
       mDateTF30 = ""
    End If
    'end 2020/02/10
    
    If bolPrint = True Then
       intFieldWidth = Array(2000, 2500)
       m_dblTop = 300: m_dblLeft = 600:   dblLineHeight = 200
       iPage = 1
       'Modified by Lydia 2023/05/19
       'Printer.PaperSize = PUB_GetPaperSize(9) '設定紙張 A4
       Printer.PaperSize = 9
       Printer.Orientation = 1 '直印
       Printer.Font.Name = "標楷體"
       m_TBWidth = Printer.ScaleWidth - 700
       
       Call PrintStaticData(pCP01, pCP02, pCP03, pCP04, iPage, dblLineHeight, m_dblTitleHeight, m_dblTop, m_dblLeft, m_TBWidth, intLine, dblPrtX, dblPrtY, intFieldWidth)
       Call PrintTableLine(iPage, dblLineHeight, m_dblTitleHeight, m_dblTop, m_dblLeft, m_TBWidth, intLine, dblPrtX, dblPrtY, intFieldWidth)  '畫表格
       Call PrintField3Title(pCP01, pCP02, iPage, dblLineHeight, m_dblTitleHeight, m_dblTop, m_dblLeft, m_TBWidth, intLine, dblPrtX, dblPrtY, intFieldWidth) '第三欄抬頭-本所案號
    End If
    intLine = intLine + 1
    strTmp(1) = Val(strTmp(1)) + 1
    strTemp = strTmp(1) & "、告申請日、案號"
    If bolPrint = True Then
      dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
      Printer.CurrentX = dblPrtX
      Printer.CurrentY = dblPrtY
      Printer.Print strTemp
    End If
    PUB_FCPPrintContactSheetA4 = PUB_FCPPrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
    
    intLine = intLine + 1
    'Modified by Morgan 2016/9/30 +申請書--何淑華
    'Modified by Lydia 2019/12/09 修正頁->修正本 (by Phoebe)
    strTemp = " ＊同時寄：收據、修正本、申請書"
    If bolPrint = True Then
      dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
      Printer.CurrentX = dblPrtX
      Printer.CurrentY = dblPrtY
      Printer.Print strTemp
    End If
    PUB_FCPPrintContactSheetA4 = PUB_FCPPrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
    
    '*** 第三欄第2點 ***
    '取得202補文件及231寄存證明的本所期限,有期限印NP15並計算NP09筆數(列印份數),無期限印「文件已齊備」
    'Modified by Morgan 2017/7/10 補文件的內容一份前的字數不一定是7，改剔除"一份"兩字就好--敏莉
    'strtmp(0) = "Select NP08,SubStr(NP15,1,7) as NP15,count(NP15) as CNP15 From NextProgress " & _
                      "Where NP01='" & pcp09 & "' And NP07 in ('202','231') And NP06 is null " & _
                      "And InStr(NP15,'專利申請書')=0 Group by NP08,NP15 " & _
            "Union Select NP08,NP15,0 as CNP15 From NextProgress " & _
                      "Where NP01='" & pcp09 & "' And NP07 in ('202','231') And NP06 is null " & _
                      "And InStr(NP15,'專利申請書')>0 Group by NP08,NP15 "
    'Modify By Sindy 2021/7/8 本所期限改抓約定期限 + ,np23
    'Modify By Sindy 2021/7/23 排除 客戶提供中說、英文參考本 在下方獨立判斷, 因非真正智慧局的期限
    strTmp(0) = "Select NP08,np23,replace(NP15,'一份','') as NP15,count(NP15) as CNP15 From NextProgress " & _
                      "Where NP01='" & pCP09 & "' And NP07 in ('202','231') And NP06 is null " & _
                      "And InStr(NP15,'專利申請書')=0 And InStr(NP15,'客戶提供中說')=0 And InStr(NP15,'英文參考本')=0 " & _
                      "Group by NP08,np23,NP15 " & _
            "Union Select NP08,np23,NP15,0 as CNP15 From NextProgress " & _
                      "Where NP01='" & pCP09 & "' And NP07 in ('202','231') And NP06 is null " & _
                      "And InStr(NP15,'專利申請書')>0 And InStr(NP15,'客戶提供中說')=0 And InStr(NP15,'英文參考本')=0 " & _
                      "Group by NP08,np23,NP15 "
    'end 2017/7/10
    intR = 1
    Set rsRD = ClsLawReadRstMsg(intR, strTmp(0))
    If intR = 1 Then
        intLine = intLine + 1
        strTmp(1) = Val(strTmp(1)) + 1
        strTemp = strTmp(1) & "、本案尚缺之文件："
        If bolPrint = True Then
            dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
            Printer.CurrentX = dblPrtX
            Printer.CurrentY = dblPrtY
            Printer.Print strTemp
        End If
        PUB_FCPPrintContactSheetA4 = PUB_FCPPrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
        
        Dim i As Integer
        With rsRD
            .MoveFirst
            Do While Not .EOF
                If bolPrint = True Then
                  If intLine > 24 Then
                      iPage = iPage + 1
                      Printer.NewPage
                      Call PrintStaticData(pCP01, pCP02, pCP03, pCP04, iPage, dblLineHeight, m_dblTitleHeight, m_dblTop, m_dblLeft, m_TBWidth, intLine, dblPrtX, dblPrtY, intFieldWidth)
                      Call PrintTableLine(iPage, dblLineHeight, m_dblTitleHeight, m_dblTop, m_dblLeft, m_TBWidth, intLine, dblPrtX, dblPrtY, intFieldWidth)   '畫表格
                      Call PrintField3Title(pCP01, pCP02, iPage, dblLineHeight, m_dblTitleHeight, m_dblTop, m_dblLeft, m_TBWidth, intLine, dblPrtX, dblPrtY, intFieldWidth)
                  End If
                End If
                
                '專利申請書不需計算份數
                If InStr(.Fields("NP15"), "專利申請書") > 0 Then
                  'Modify By Sindy 2021/7/8 本所期限改抓約定期限
                  If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
                     strTemp = "(期限" & ChangeWStringToTDateString("" & .Fields("NP23")) & ")"
                  Else
                  '2021/7/8 END
                     strTemp = "(期限" & ChangeWStringToTDateString(.Fields("NP08")) & ")"
                  End If
                Else
                  'Modify By Sindy 2021/7/8 本所期限改抓約定期限
                  If strSrvDate(1) >= 外專台灣案約定期限啟用日 Then
                     strTemp = "(" & .Fields("CNP15") & "份)(期限" & ChangeWStringToTDateString("" & .Fields("NP23")) & ")"
                  Else
                  '2021/7/8 END
                     strTemp = "(" & .Fields("CNP15") & "份)(期限" & ChangeWStringToTDateString(.Fields("NP08")) & ")"
                  End If
                End If
                strTemp = " ＊" & .Fields("NP15") & strTemp
                PUB_FCPPrintContactSheetA4 = PUB_FCPPrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
                '超過20個字-換行
                If strTemp <> StrToStr(strTemp, 19) Then
                    bolFirst = True
                    Do While strTemp <> ""
                        If bolPrint = True Then
                           If intLine > 24 Then
                               iPage = iPage + 1
                               Printer.NewPage
                               Call PrintStaticData(pCP01, pCP02, pCP03, pCP04, iPage, dblLineHeight, m_dblTitleHeight, m_dblTop, m_dblLeft, m_TBWidth, intLine, dblPrtX, dblPrtY, intFieldWidth)
                               Call PrintTableLine(iPage, dblLineHeight, m_dblTitleHeight, m_dblTop, m_dblLeft, m_TBWidth, intLine, dblPrtX, dblPrtY, intFieldWidth)   '畫表格
                               Call PrintField3Title(pCP01, pCP02, iPage, dblLineHeight, m_dblTitleHeight, m_dblTop, m_dblLeft, m_TBWidth, intLine, dblPrtX, dblPrtY, intFieldWidth)
                           End If
                        End If
                        If bolFirst = True Then
                            strTmpA = StrToStr(strTemp, 19) '目前要取的字串
                            bolFirst = False
                            strTemp = Mid(strTemp, Len(strTmpA) + 1) '取完剩的字串
                        Else
                            strTmpA = Space(3) & StrToStr(strTemp, 19)
                            strTemp = Mid(strTemp, Len(strTmpA) - 3 + 1) '取完剩的字串
                        End If
                        If bolPrint = True Then
                           intLine = intLine + 1
                           dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
                           Printer.CurrentX = dblPrtX
                           Printer.CurrentY = dblPrtY
                           Printer.Print strTmpA
                        End If
                    Loop
                Else
                    If bolPrint = True Then
                        intLine = intLine + 1
                        dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
                        Printer.CurrentX = dblPrtX
                        Printer.CurrentY = dblPrtY
                        Printer.Print strTemp
                    End If
                End If
                .MoveNext
            Loop
        End With
    Else
        intLine = intLine + 1
        strTmp(1) = Val(strTmp(1)) + 1
        strTemp = strTmp(1) & "、文件已齊備"
        If bolPrint = True Then
            dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
            Printer.CurrentX = dblPrtX
            Printer.CurrentY = dblPrtY
            Printer.Print strTemp
        End If
        PUB_FCPPrintContactSheetA4 = PUB_FCPPrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
    End If
    '*** End 第三欄第2點 ***
    
    'Added by Lydia 2018/01/05 新案發文時，有收文新案翻譯且原文字數為空白，簡易連聯絡單多加註記
    If InStr(NewCasePtyList, pCP10) > 0 Then
        strTmp(0) = "select cp09,cp10,tf01,tf23,tf19,tf20 from caseprogress,transfee " & _
                          "where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' " & _
                          "and cp10='201' and cp09=tf01(+)"
        intR = 1
        Set rsRD = ClsLawReadRstMsg(intR, strTmp(0))
        If intR = 1 Then
             If Val("" & rsRD.Fields("tf23")) = 0 Then
                strTmp(1) = Val(strTmp(1)) + 1
                m_Line3 = strTmp(1) & "、原文字數為空白，請承辦填寫原文字數"
                If bolPrint = True Then
                     intLine = intLine + 1
                     dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
                     Printer.CurrentX = dblPrtX
                     Printer.CurrentY = dblPrtY
                     Printer.Print m_Line3
                End If
                PUB_FCPPrintContactSheetA4 = PUB_FCPPrintContactSheetA4 & m_Line3 & vbCrLf 'Add By Sindy 2022/5/11
                
                m_Line3 = " 　，退程序輸入。"
                If bolPrint = True Then
                     intLine = intLine + 1
                     dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
                     Printer.CurrentX = dblPrtX
                     Printer.CurrentY = dblPrtY
                     Printer.Print m_Line3
                End If
                PUB_FCPPrintContactSheetA4 = PUB_FCPPrintContactSheetA4 & m_Line3 & vbCrLf 'Add By Sindy 2022/5/11
             End If
        End If
    End If
'    If m_Line3 <> "" Then
'         strtmp(1) = "4"
'    Else
'         strtmp(1) = "3"
'    End If
    'end 2018/01/05
    
    strTemp = ""
    'Modify by Amy 2016/04/29
    If pCP10 = "103" Or pCP10 = "125" Then
        'Modified by Lydia 2018/01/05
        'strTemp = "3、退程序主管分案撰中文圖說"
        strTmp(1) = Val(strTmp(1)) + 1
        'Modify By Sindy 2022/12/29
        'strTemp = strtmp(1) & "、退程序主管分案撰中文圖說"
        strTemp = strTmp(1) & "、通知分案撰中文圖說"
        If strTemp <> "" Then
           If bolPrint = True Then
               intLine = intLine + 1
               dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
               Printer.CurrentX = dblPrtX
               Printer.CurrentY = dblPrtY
               Printer.Print strTemp
           End If
           PUB_FCPPrintContactSheetA4 = PUB_FCPPrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
         End If
        '2022/12/29 END
    'Modify By Sindy 2022/12/29 設計案也要抓有告代
    End If
    'Else
    strTemp = ""
    '2022/12/29 END
        'Modified by Lydia 2018/01/05
        'strTemp = "3、退檔 or 退程序主管分案告代函"
        'Modified by Lydia 2018/04/16
        'strTemp = strtmp(1) & "、退檔 or 退程序主管分案告代函"
        
        'strTemp = strtmp(1) & "、退檔 or 退程序(有告代函)"
        'Modify By Sindy 2022/5/11
        '請改成:若有收文告代(901)或主動修正(203)未發文
        '且進度檔是帶提申後告代(ex:066746)or提申後主動修正，
        '二者皆有則帶: 有告代函及主動修正--(請抓承辦人(工程師))，
        '若只有一種則帶其一即可，例: 有告代函--(請抓承辦人)，若皆無則此欄可不帶。
         strTmp(0) = "SELECT * FROM caseprogress,staff" & _
                     " WHERE cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "'" & _
                     " and ((cp10='901' and instr(cp64,'提申後告代')>0) or (cp10='203' and instr(cp64,'提申後主動修正')>0))" & _
                     " and cp27||cp57 is null" & _
                     " and cp14=st01(+)" & _
                     " order by cp10 desc"
         intR = 1
         Set rsRD = ClsLawReadRstMsg(intR, strTmp(0))
         If intR = 1 Then
            rsRD.MoveFirst
            strTmp(10) = 0
            strTmp(9) = "" & rsRD.Fields("st02")
            strTmp(1) = Val(strTmp(1)) + 1
            strTemp = strTmp(1) & "、有"
            Do While Not rsRD.EOF
               strTmp(10) = Val(strTmp(10)) + 1
               If Val(strTmp(10)) > 1 Then
                  strTemp = strTemp & "及"
               End If
               If rsRD.Fields("cp10") = "901" Then
                  strTemp = strTemp & "告代函"
               ElseIf rsRD.Fields("cp10") = "203" Then
                  strTemp = strTemp & "主動修正"
               End If
               rsRD.MoveNext
            Loop
            strTemp = strTemp & "--" & strTmp(9)
         End If
         '2022/5/11 END
'    End If
    'end 2016/04/29
    If strTemp <> "" Then
      If bolPrint = True Then
         intLine = intLine + 1
         dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
         Printer.CurrentX = dblPrtX
         Printer.CurrentY = dblPrtY
         Printer.Print strTemp
      End If
      PUB_FCPPrintContactSheetA4 = PUB_FCPPrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
    End If
    
    'Added by Lydia 2019/01/04 新案發文有扣款日期，則簡易連聯絡單多加註記
    If InStr(NewCasePtyList, pCP10) > 0 Then
        'Modified by Lydia 2023/05/02 +cp84
        'Modified by Lydia 2023/05/19 +CP158
        strTmp(0) = "select cp09,cp152,cp84,CP158  from caseprogress where cp09='" & pCP09 & "' "
        intR = 1
        Set rsRD = ClsLawReadRstMsg(intR, strTmp(0))
        If intR = 1 Then
            If Val("" & rsRD.Fields("cp152")) > 0 Then
                intLine = intLine + 1
                strTmp(1) = Val(strTmp(1)) + 1
                strTemp = strTmp(1) & "、收據下載日期：" & ChangeTStringToTDateString(TransDate("" & rsRD.Fields("cp152"), 1))
                If bolPrint = True Then
                     dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
                     Printer.CurrentX = dblPrtX
                     Printer.CurrentY = dblPrtY
                     Printer.Print strTemp
                End If
                'Added by Lydia 2023/05/02 +規費金額
                If Val("" & rsRD.Fields("cp84")) > 0 Then
                  If bolPrint = True Then
                       intLine = intLine + 1
                       dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
                       Printer.CurrentX = dblPrtX
                       Printer.CurrentY = dblPrtY
                       Printer.Print "　　，規費金額：NTD " & rsRD.Fields("cp84")
                  End If
                  strTemp = strTemp & "，規費金額：NTD " & rsRD.Fields("cp84")
                End If
                'Added by Lydia 2023/05/03 因為先發文新案，所以同時提醒該案亦有同日發文的規費金額 ---- frm060104_k使用
                'Modiried by Lydia 2023/05/19  如果新案同時發文實審，則在發文新案時，key通知申請案號時應該要先跳出，待實審發文後再單獨進入通知申請案號；
                '所以可抓到實審發文規費，且若同日也有938超頁費和939超項費請一併將超頁費和超項費收文規費帶入
                'If PUB_ChkCPExist(cp, "416", 1) = True Then
                '    strTemp = strTemp & "+實體審查:"
                'End If
                ''end 2023/05/03
                strTmp(0) = "select cp09,cp10,decode(cp10,'416','1','2') ord1,cpm03,cp84 from caseprogress, casepropertymap" & _
                                 " where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp158=" & rsRD.Fields("cp158") & _
                                 " and cp10 in ('416','938','939','917') and cp01=cpm01(+) and cp10=cpm02(+) order by ord1,cp09 "
                intR = 1
                Set rsRD = ClsLawReadRstMsg(intR, strTmp(0))
                If intR = 1 Then
                   rsRD.MoveFirst
                   Do While Not rsRD.EOF
                       strTemp = strTemp & " + " & rsRD.Fields("cpm03") & ":" & IIf(Val("" & rsRD.Fields("cp84")) > 0, "" & rsRD.Fields("cp84"), String(5, " "))
                       rsRD.MoveNext
                   Loop
                End If
                'end 2023/05/19
                PUB_FCPPrintContactSheetA4 = PUB_FCPPrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
            End If
        End If
        'Added by Lydia 2020/02/10 新案建檔有重新印「發文簡易聯絡單」之功能,所以要能抓取相關行事曆
        If mDate209210 = "" And InStr("101,102", pCP10) > 0 Then '客戶提供中說期限
            'Add By Sindy 2021/7/23 不產生行事曆了,改新增下一程序
            strTmp(0) = "Select NP08,np23,NP15,0 as CNP15 From NextProgress " & _
                      "Where NP01='" & pCP09 & "' And NP07 in ('202','231') And NP06 is null " & _
                      "And InStr(NP15,'客戶提供中說')>0 " & _
                      "Group by NP08,np23,NP15 "
            intR = 1
            Set rsRD = ClsLawReadRstMsg(intR, strTmp(0))
            If intR = 1 Then
               mDate209210 = "" & rsRD.Fields("np23")
            Else
            '2021/7/23 END
               strTmp(0) = "select sc01 from staff_calendar where instr(sc04,'催客戶提供中說期限') > 0 and sc05='" & pCP01 & "' and sc06='" & pCP02 & "' and sc07='" & pCP03 & "' and sc08='" & pCP04 & "' and sc18 is null "
               strTmp(0) = strTmp(0) & "order by sc01"
               intR = 1
               Set rsRD = ClsLawReadRstMsg(intR, strTmp(0))
               If intR = 1 Then
                  mDate209210 = "" & rsRD.Fields("sc01")
               End If
            End If
        End If
        If mDateTF30 = "" Then
            'Add By Sindy 2021/7/23 不產生行事曆了,改新增下一程序
            strTmp(0) = "Select NP08,np23,NP15,0 as CNP15 From NextProgress " & _
                      "Where NP01='" & pCP09 & "' And NP07 in ('202','231') And NP06 is null " & _
                      "And InStr(NP15,'英文參考本')>0 " & _
                      "Group by NP08,np23,NP15 "
            intR = 1
            Set rsRD = ClsLawReadRstMsg(intR, strTmp(0))
            If intR = 1 Then
               mDateTF30 = "" & rsRD.Fields("np23")
            Else
            '2021/7/23 END
               strTmp(0) = "select sc01 from staff_calendar where instr(sc04,'催客戶提供英文翻譯本') > 0 and sc05='" & pCP01 & "' and sc06='" & pCP02 & "' and sc07='" & pCP03 & "' and sc08='" & pCP04 & "' and sc18 is null "
               strTmp(0) = strTmp(0) & "order by sc01"
               intR = 1
               Set rsRD = ClsLawReadRstMsg(intR, strTmp(0))
               If intR = 1 Then
                  mDateTF30 = "" & rsRD.Fields("sc01")
               End If
            End If
        End If
    End If
    'end 2019/01/04
    
    'Added by Lydia 2019/01/17 FCP新案發文時若檢視中說or核對中說格式未發文，自動設行事曆；一併在簡易聯絡單增加第5點：客戶提供中說期限。
    If mDate209210 <> "" Then
        intLine = intLine + 1
        strTmp(1) = Val(strTmp(1)) + 1
        strTemp = strTmp(1) & "、客戶提供中說期限：" & ChangeTStringToTDateString(TransDate(mDate209210, 1))
        If bolPrint = True Then
            dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
            Printer.CurrentX = dblPrtX
            Printer.CurrentY = dblPrtY
            Printer.Print strTemp
        End If
        PUB_FCPPrintContactSheetA4 = PUB_FCPPrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
    End If
    'end 2019/01/17
    
    'Added by Lydia 2019/12/11 FCP新案發文時檢查有新案翻譯未發文並且尚"待英文本翻譯"，自動設行事曆；一併在簡易聯絡單增加第5點：客戶提供英文翻譯本。
    If mDateTF30 <> "" Then
        intLine = intLine + 1
        strTmp(1) = Val(strTmp(1)) + 1
        strTemp = strTmp(1) & "、客戶提供英文翻譯本：" & ChangeTStringToTDateString(TransDate(mDateTF30, 1))
        If bolPrint = True Then
            dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500
            Printer.CurrentX = dblPrtX
            Printer.CurrentY = dblPrtY
            Printer.Print strTemp
        End If
        PUB_FCPPrintContactSheetA4 = PUB_FCPPrintContactSheetA4 & strTemp & vbCrLf 'Add By Sindy 2022/5/11
    End If
    'end 2019/12/11
    
    If bolPrint = True Then
       Printer.EndDoc
    End If
End Function

'Added by Lydia 2023/05/19 外專-簡易聯絡單內容---從frm060104_3抽出
Private Sub PrintStaticData(ByVal cp01 As String, ByVal cp02 As String, ByVal cp03 As String, ByVal cp04 As String, _
                ByRef iPage As Integer, ByRef dblLineHeight As Double, ByRef m_dblTitleHeight As Double, ByRef m_dblTop As Double, ByRef m_dblLeft As Double, _
                ByRef m_TBWidth As Double, ByRef intLine As Integer, ByRef dblPrtX As Double, ByRef dblPrtY As Double, ByRef intFieldWidth)
'Dim iPage As Integer
'Dim dblLineHeight As Double '行高
'Dim m_dblTitleHeight As Double '抬頭
'Dim m_dblTop As Double '上邊界
'Dim m_dblLeft As Double '左邊界
'Dim m_TBWidth As Double '表格寬
'Dim intLine As Integer
'Dim dblPrtX As Double
'Dim dblPrtY As Double
'Dim intFieldWidth
Dim strCon1 As String 'Added by Lydia 2024/03/29

    intLine = 1
    
    'Removed by Morgan 2020/3/30
    'strcon1 = "台一國際專利商標事務所"
    'm_dblTitleHeight = (intLine + 0.8) * 300
    'Printer.Font.Size = 22
    'dblPrtX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strcon1) / 2)
    'dblPrtY = m_dblTop + m_dblTitleHeight
    'Printer.CurrentX = dblPrtX
    'Printer.CurrentY = dblPrtY
    'Printer.Print strcon1
    'intLine = intLine + 1
    'end 2020/3/30
    
    strCon1 = "簡易聯絡單"
    Printer.Font.Size = 20
    m_dblTitleHeight = (intLine + 1.5) * 300
    dblPrtX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strCon1) / 2)
    dblPrtY = m_dblTop + m_dblTitleHeight
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strCon1
    m_dblTitleHeight = m_dblTitleHeight + 400
    
    intLine = 1
    Printer.Font.Size = 18
    strCon1 = "受 文 者"
    dblPrtX = m_dblLeft + intFieldWidth(0) / 2 - Printer.TextWidth(strCon1) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 350
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strCon1
    
    strCon1 = "發 文 者"
    dblPrtX = m_dblLeft + intFieldWidth(0) + intFieldWidth(1) / 2 - Printer.TextWidth(strCon1) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 350
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strCon1
 
    intLine = intLine + 1
    Printer.Font.Size = 16
    strCon1 = GetStaffName(PUB_GetFCPSalesNo(cp01, cp02, cp03, cp04))
    dblPrtX = m_dblLeft + intFieldWidth(0) / 2 - Printer.TextWidth(strCon1) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + dblLineHeight + intLine * 500
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strCon1

    strCon1 = GetStaffName(PUB_GetFCPHandler(cp01, cp02, cp03, cp04))
    dblPrtX = m_dblLeft + intFieldWidth(0) + intFieldWidth(1) / 2 - Printer.TextWidth(strCon1) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + dblLineHeight + intLine * 500
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strCon1

    intLine = intLine + 2
    Printer.Font.Size = 18
    strCon1 = "發文時間"
    dblPrtX = m_dblLeft + intFieldWidth(0) / 2 - Printer.TextWidth(strCon1) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight * 1.2 + dblLineHeight + intLine * 500  'm_dblTitleHeight * 1.2 for 垂直靠下
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strCon1
    
    Printer.Font.Size = 16
    strCon1 = Year(Now) - 1911 & "年" & "  月" & "  日"
    dblPrtX = m_dblLeft + intFieldWidth(0) + intFieldWidth(1) / 2 - Printer.TextWidth(strCon1) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight * 1.2 + dblLineHeight + intLine * 500
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strCon1
    
    intLine = intLine + 2
    Printer.Font.Size = 18
    strCon1 = "發文地點"
    dblPrtX = m_dblLeft + intFieldWidth(0) / 2 - Printer.TextWidth(strCon1) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + dblLineHeight + intLine * 500
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strCon1
    
    Printer.Font.Size = 16
    strCon1 = "國外部專利處"
    dblPrtX = m_dblLeft + intFieldWidth(0) + intFieldWidth(1) / 2 - Printer.TextWidth(strCon1) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + dblLineHeight + intLine * 500
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strCon1
 
    'Added by Lydia 2017/10/20 新增承辦、判行、確認已報告+Email已回存之簽核欄位
    intLine = intLine + 15
    'Remove by Lydia 2019/03/21 取消承辦、判行(by A4011)
'    Printer.Font.Size = 18
'    strcon1 = "承　辦"
'    dblPrtX = m_dblLeft + intFieldWidth(0) / 2 - Printer.TextWidth(strcon1) / 2
'    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 350
'    Printer.CurrentX = dblPrtX
'    Printer.CurrentY = dblPrtY
'    Printer.Print strcon1
'
'    strcon1 = "判　行"
'    dblPrtX = m_dblLeft + intFieldWidth(0) + intFieldWidth(1) / 2 - Printer.TextWidth(strcon1) / 2
'    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 350
'    Printer.CurrentX = dblPrtX
'    Printer.CurrentY = dblPrtY
'    Printer.Print strcon1
    'end 2019/03/21
    '確認已報告+Email已回存
    intLine = intLine + 11
    Printer.Font.Size = 16
    strCon1 = "確認已報告 + E-mail已回存"
    dblPrtX = m_dblLeft + (intFieldWidth(0) + intFieldWidth(1)) / 2 - Printer.TextWidth(strCon1) / 2
    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 350
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strCon1
    'end 2017/10/20
End Sub

'Added by Lydia 2023/05/19 外專-簡易聯絡單內容---從frm060104_3抽出，並且改成列印／回傳
Private Sub PrintTableLine(ByRef iPage As Integer, ByRef dblLineHeight As Double, ByRef m_dblTitleHeight As Double, ByRef m_dblTop As Double, ByRef m_dblLeft As Double, _
                ByRef m_TBWidth As Double, ByRef intLine As Integer, ByRef dblPrtX As Double, ByRef dblPrtY As Double, ByRef intFieldWidth)
'Dim iPage As Integer
'Dim dblLineHeight As Double '行高
'Dim m_dblTitleHeight As Double '抬頭
'Dim m_dblTop As Double '上邊界
'Dim m_dblLeft As Double '左邊界
'Dim m_TBWidth As Double '表格寬
'Dim intLine As Integer
'Dim dblPrtX As Double
'Dim dblPrtY As Double
'Dim intFieldWidth
Dim i As Integer
Dim startY As Integer 'Added by Lydia 2017/10/20
Dim strCon1 As String 'Added by Lydia 2024/03/29
    '雙橫線-上
    dblPrtY = m_dblTop + m_dblTitleHeight + 50
    startY = dblPrtY 'Added by Lydia 2017/10/20
    Printer.Line (m_dblLeft, dblPrtY)-(m_TBWidth, dblPrtY)
    dblPrtY = m_dblTop + m_dblTitleHeight + 100
    Printer.Line (m_dblLeft + 50, dblPrtY)-(m_TBWidth - 50, dblPrtY)
    
    '雙橫線-下
    Printer.Line (m_dblLeft + 50, 14950)-(m_TBWidth - 50, 14950)
    Printer.Line (m_dblLeft, 15000)-(m_TBWidth, 15000)

    '雙直線-左
    Printer.Line (m_dblLeft, dblPrtY - 50)-(m_dblLeft, 15000)
    Printer.Line (m_dblLeft + 50, dblPrtY)-(m_dblLeft + 50, 14950)
    
    '雙直線-右
    Printer.Line (m_TBWidth - 50, dblPrtY)-(m_TBWidth - 50, 14950)
    Printer.Line (m_TBWidth, dblPrtY - 50)-(m_TBWidth, 15000)
      
    '欄位分隔線-橫線
    'Memo by Lydia 2017/10/20 ex.FCP-57679
    '受文者     |發文者
    '-------------------- i=1
    '羅XX       |蔡XX
    '-------------------- i=2
    '發文時間   |106年 月 日
    '-------------------- i=3
    '發文地點   |國外部專利處
    '(隔4行)    |
    '-----------------------
    '承辦       |判行
    '(隔2行)  |
    '-----------------------
    '確認已報告+Email已回存
    'end 2017/10/20
    For i = 1 To 3
        dblPrtY = m_dblTop + m_dblTitleHeight + i * 1000
        Printer.Line (m_dblLeft + 50, dblPrtY)-(m_dblLeft + intFieldWidth(0) + intFieldWidth(1), dblPrtY)
    Next
    
    'Added by Lydia 2017/10/20 新增承辦、判行、確認已報告+Email已回存之簽核欄位
    '承辦、判行(橫線)
    'Remove by Lydia 2019/03/21 取消承辦、判行(by A4011)
    'dblPrtY = m_dblTop + m_dblTitleHeight + 7 * 1000
    'Printer.Line (m_dblLeft + 50, dblPrtY)-(m_dblLeft + intFieldWidth(0) + intFieldWidth(1), dblPrtY)
    'dblPrtY = m_dblTop + m_dblTitleHeight + 8 * 1000
    'Printer.Line (m_dblLeft + 50, dblPrtY)-(m_dblLeft + intFieldWidth(0) + intFieldWidth(1), dblPrtY)
    'end 2019/03/21
    '確認已報告+Email已回存(橫線),與Anny(A4011)確認過標題下橫線不用畫
    dblPrtY = m_dblTop + m_dblTitleHeight + 11 * 1000
    Printer.Line (m_dblLeft + 50, dblPrtY)-(m_dblLeft + intFieldWidth(0) + intFieldWidth(1), dblPrtY)
    'end 2017/10/20
    
    'Move by Lydia 2017/10/20 從上方移下來
    '欄位分隔線-直線
    dblPrtX = m_dblLeft
    For i = 0 To UBound(intFieldWidth)
        dblPrtX = dblPrtX + intFieldWidth(i)
        'Modified by Lydia 2017/10/20 第1條分隔線只畫到"確認已報告+Email已回存"
        'Printer.Line (dblPrtX, dblPrtY)-(dblPrtX, 14950)
        If i = 0 Then
           Printer.Line (dblPrtX, startY)-(dblPrtX, dblPrtY)
        Else
           Printer.Line (dblPrtX, startY)-(dblPrtX, 14950)
        End If
        'end 2017/10/20
    Next
    'end 2017/10/20
    
    '頁碼
    Printer.Font.Size = 12
    strCon1 = iPage
    dblPrtX = Printer.ScaleWidth / 2 - (Printer.TextWidth(strCon1) / 2)
    dblPrtY = Printer.ScaleHeight - m_dblTop * 2.5
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strCon1
End Sub

Private Sub PrintField3Title(ByVal cp01 As String, ByVal cp02 As String, ByRef iPage As Integer, ByRef dblLineHeight As Double, ByRef m_dblTitleHeight As Double, ByRef m_dblTop As Double, ByRef m_dblLeft As Double, _
                ByRef m_TBWidth As Double, ByRef intLine As Integer, ByRef dblPrtX As Double, ByRef dblPrtY As Double, ByRef intFieldWidth)
'Dim iPage As Integer
'Dim dblLineHeight As Double '行高
'Dim m_dblTitleHeight As Double '抬頭
'Dim m_dblTop As Double '上邊界
'Dim m_dblLeft As Double '左邊界
'Dim m_TBWidth As Double '表格寬
'Dim intLine As Integer
'Dim dblPrtX As Double
'Dim dblPrtY As Double
'Dim intFieldWidth
Dim strCon1 As String 'Added by Lydia 2024/03/29
    '第三欄
    intLine = 1
    dblPrtX = m_dblLeft + intFieldWidth(0) + intFieldWidth(1) + 200
    
    Printer.Font.Size = 16
    strCon1 = cp01 & "-" & cp02
    dblPrtY = m_dblTop + m_dblTitleHeight + intLine * 500 - 100
    Printer.CurrentX = dblPrtX
    Printer.CurrentY = dblPrtY
    Printer.Print strCon1
    
    Printer.FontSize = 14
End Sub

'Added by Lydia 2023/06/09 當寰華案在key閉卷按確認時，請判斷是否有相關香港案及澳門案未不續辦/閉卷，若有則發mail
'Modified by Lydia 2023/06/28 傳入案件性質pCP10
Public Sub PUB_CloseMailto013044(ByVal pType As String, ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pCP10 As String)
'pType: 1-閉卷
Dim strCase(1 To 4) As String
Dim strTo As String, strCC As String, strSub As String, strTmp1 As String
  
   If ChkCMIsExist013(pCP01, pCP02, pCP03, pCP04, strCase(1), strCase(2), strCase(3), strCase(4), , , , True) = True Then
       strSub = strSub & "香港案" & strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", "-" & strCase(3) & "-" & strCase(4), "")
   End If
   If ChkCMIsExist013(pCP01, pCP02, pCP03, pCP04, strCase(1), strCase(2), strCase(3), strCase(4), , , "5", True) = True Then
       strSub = strSub & "澳門案" & strCase(1) & "-" & strCase(2) & IIf(strCase(3) & strCase(4) <> "000", "-" & strCase(3) & "-" & strCase(4), "")
   End If
   If strSub <> "" Then
       strTo = PUB_GetFCPSalesNo(pCP01, pCP02, pCP03, pCP04)
       If strTo <> "" Then
          strTmp1 = PUB_GetFCPProSup(strTo)
          strCC = strCC & ";" & strTmp1
          strTmp1 = PUB_GetFCPHandler(pCP01, pCP02, pCP03, pCP04)
          If strTmp1 <> "" Then strCC = strCC & ";" & strTmp1
          If InStr(strCC, strUserNum) = 0 Then strCC = strCC & ";" & strUserNum
          strCC = Mid(strCC, 2) & ";backup"
          'Modified by Lydia 2023/06/27 改主旨◎寰華案Our Ref: P-124117 [INCOM.913]已上閉卷，請提醒代理人另有相關香港案P-124316澳門案P-124317，待大陸案接獲消滅函將自動閉卷
          'strSub = "寰華案" & pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 <> "000", "-" & pCP03 & "-" & pCP04, "") & "已上" & IIf(pType = "1", "閉卷", "不續辦") & _
                   "，請提醒代理人另有相關" & strSub & "，待大陸案接獲消滅函將自動閉卷"
          strSub = "寰華案Our Ref: " & pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 <> "000", "-" & pCP03 & "-" & pCP04, "") & " [INCOM." & pCP10 & "]已上" & IIf(pType = "1", "閉卷", "不續辦") & _
                   "，請提醒代理人另有相關" & strSub & "，待大陸案接獲消滅函將自動閉卷"
          strTmp1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
             " VALUES ( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
             ",to_char(sysdate,'hh24miss'),'" & strSub & "','如旨','" & strCC & "' )"
          cnnConnection.Execute strTmp1
       End If
   End If

End Sub


'Added by Lydia 2023/10/04 FMP案待客戶最終指示相關控管
Public Function PUB_ChkFMP970mail(ByVal pKind As String, ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, Optional ByRef pMsgList As String) As Boolean
'pKind: 1.自動發mail, 2.Key翻譯完稿日(Outlook草稿), 3.指示信/撰寫信函彈訊息
Dim strCP14 As String, strMemo As String
Dim intQ As Integer, strQ1 As String
Dim rsQD As New ADODB.Recordset
Dim strSub As String, strCont As String
'Added by Lydia 2023/10/26
Dim m_bolFMP2 As Boolean
Dim strBase(0 To 4) As String

   PUB_ChkFMP970mail = False
   pMsgList = ""
   
   If PUB_ChkIsFMP(pCP01, pCP02, pCP03, pCP04) = False Then Exit Function
   'Added by Lydia 2023/10/26 判斷寰華案
   m_bolFMP2 = PUB_FMPtoCheck(1, 2, Pub_strUserST05, pCP01, pCP02, pCP03, pCP04)
   '有可能新案尚未發文
   If m_bolFMP2 = False Then
      m_bolFMP2 = PUB_GetFMP2toP(pCP01, pCP02, pCP03, pCP04)
   End If
   strBase(0) = pCP01 & pCP02 & pCP03 & pCP04
   Call ChgCaseNo(strBase(0), strBase)
   'end 2023/10/26
   
   strQ1 = "select pa01,pa02,pa03,pa04,pa05,pa09,cp10,cp47,cp64,bcp09,bcp14 " & _
           "From patent, caseprogress,(select cp01 as bcp01,cp02 as bcp02,cp03 as bcp03,cp04 as bcp04,cp09 as bcp09,cp14 as bcp14 from caseprogress where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp10='970' and cp158=0 and cp159=0) " & _
           "where pa01='" & pCP01 & "' and pa02='" & pCP02 & "' and pa03='" & pCP03 & "' and pa04='" & pCP04 & "' " & _
           "and pa01=cp01(+) and pa02=cp02(+) and pa03=cp03(+) and pa04=cp04(+) and cp31='Y' " & _
           "and pa01=bcp01(+) and pa02=bcp02(+) and pa03=bcp03(+) and pa04=bcp04(+) "
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      'Added by Lydia 2023/10/26 FMP案不收文970待客戶最終指示
      'Modified by Lydia 2024/02/02 限制發Email + pKind = "1"
      If m_bolFMP2 = False And pKind = "1" Then
         'Mark by Lydia 2024/01/31 與品薇電話確認：只要順稿就要發Email
         'If "" & rsQD.Fields("cp47") <> "" Then
         '   Exit Function
         'End If
         'end 2024/01/31
         strMemo = "" & rsQD.Fields("cp64")
      Else
      'end 2023/10/26
         'Memo by Lydia 2023/10/26 寰華案會收文970待客戶最終指示，因為PCT案提申日也記錄在PA10，改用新案之代理人提申日來判斷未提申
         'Modified by Lydia 2023/11/02 debug
         'If "" & rsQD.Fields("bcp09") = "" Or "" & rsQD.Fields("cp47") <> "" Then
         If "" & rsQD.Fields("bcp09") <> "" And "" & rsQD.Fields("cp47") = "" Then
         Else
         'end 2023/11/02
            Exit Function
         End If
         strMemo = "" & rsQD.Fields("cp64")
         strCP14 = "" & rsQD.Fields("bcp14")
      End If 'Added by Lydia 2023/10/26
   End If
   
   Select Case pKind
      Case "1"  '自動發mail
         If strCP14 = "" Then
            'Modified by Lydia 2023/10/26 debug
            'strCP14 = PUB_GetFCPPromoterNo(pCP01, "936")
            If strBase(2) <> "" Then
               Call PUB_ChkCPExist(strBase, "936", 1, strBase(0), strCP14)
            End If
            'end 2023/10/26
         End If
         If strCP14 <> "" Then
            'Added by Lydia 2023/10/26 區分FMP案和寰華案
            If m_bolFMP2 = False Then 'FMP案
               strSub = "【" & pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 <> "000", "-" & pCP03 & "-" & pCP04, "") & "代理人已返稿，請詳見附件(即該代理人來函msg)內容並續行後續程序。】"
               strCont = "新案進度備註:" & strMemo
            Else
            'end 2023/10/26
               strSub = "【待最終指示】【" & pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 <> "000", "-" & pCP03 & "-" & pCP04, "") & "代理人已返稿，請詳見附件(即該代理人來函msg)內容並續行後續程序。】"
               strCont = "1.本案須獲客戶最終指示才可提申，若已獲客戶指示可提申，請通知外專程序上發文。" & vbCrLf & _
                         "2.新案進度備註:" & strMemo
            End If 'Added by Lydia 2023/10/26
            PUB_SendMail strUserNum, strCP14, "", strSub, strCont
            PUB_ChkFMP970mail = True
         End If
      Case "2", "3" '2.Key翻譯完稿日(Outlook草稿) 3.指示信/撰寫信函彈訊息
         If pKind = "2" Then
            pMsgList = "本案須獲客戶最終指示才可提申，若已獲客戶指示可提申，請通知外專程序上發文。"
         Else
            pMsgList = "本案尚待客戶最終指示，請指示信加入:" & vbCrLf & "本案須獲客戶最終指示才可提申"
         End If
         PUB_ChkFMP970mail = True
   End Select
   Set rsQD = Nothing
End Function

'Added by Lydia 2023/10/31 FCP+FMP寰華案：判斷上一道工程師案件性質是否已經請款，依狀況自動發Email通知；原本為FCP案、寰華案的獨立檢查，改成共用模組
'Modified by Lydia 2024/04/10 +pKind 核准類型
Public Function PUB_ChkFCPtoDNUPL(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pCP10 As String, ByVal pCP09 As String, Optional ByVal PTitle As String, Optional ByVal pKind As String) As Boolean
'----相關公告112051801、112020304、111050201：以原本FMP寰華案的程式，改成共用模組。
'pCP09、pCP10: 輸入C類來函的收文號、案件性質
Dim intQ As Integer, rsQD As New ADODB.Recordset
Dim strNA16 As String, strHandler As String, strA2 As String
Dim strTo As String, strCC As String, strSubject As String, strContent As String
Dim strTmpA(0 To 4) As String
Dim strDefTime As String 'Added by Lydia 2023/11/28
Dim stPA89Memo As String 'Add by Amy 2025/08/05

   PUB_ChkFCPtoDNUPL = False
   If pCP01 <> "P" And pCP01 <> "FCP" Then Exit Function
   
   'Added by Lydia 2024/08/28 113/8/6大批讓與後的X81780請先不管制本案是否有尚未請款事宜，並正常通知承辦報告即可。判斷當程序在輸入核准(相關收文號為讓與)
   strTmpA(0) = "select c2.cp05 from caseprogress c1,caseprogress c2,patent where c1.cp09='" & pCP09 & "' and c1.cp43=c2.cp09(+) " & _
                "and c1.cp01=pa01(+) and c1.cp02=pa02(+) and c1.cp03=pa03(+) and c1.cp04=pa04(+) " & _
                "and pa26='X81780000' and c2.cp10 in ('701','708') and c2.cp05=20240806 and c2.cp65='QPGMR' "
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strTmpA(0))
   If intQ = 1 Then
      Exit Function
   End If
   'end 2024/08/28
   
On Error GoTo ErrHandle

   'Modified by Lydia 2023/11/08 +cp10,cp27,cp158
   'Modified by Lydia 2024/04/29 + and cp158 > 0
   strTmpA(0) = "select cp09,cp60,cp14,cp10,cp27,cp158 from caseprogress where cp09= (" & _
                     "select max(cp09) mno from caseprogress,staff where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp158 > 0 and cp159=0 " & _
                     "and cp14=st01(+) and st03='F21' and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' and nvl(cp20,'Y')<>'N' " & _
                     "and cp05 = (select max(cp05) mdate from caseprogress, staff where cp01='" & pCP01 & "' and cp02='" & pCP02 & "' and cp03='" & pCP03 & "' and cp04='" & pCP04 & "' and cp158 > 0 and cp159=0 " & _
                     "and cp14=st01(+) and st03='F21' and st01<>'F4102' and st01<>'F4104' and st01<>'F4105' and nvl(cp20,'Y')<>'N' and nvl(cp43,'N') <> '" & pCP09 & "' )) "
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strTmpA(0))
   If intQ = 1 Then
      strTmpA(0) = ""
      strNA16 = PUB_GetFCPHandler(pCP01, pCP02, pCP03, pCP04) '程序
      If PTitle = "" And pCP10 = "1001" Then '新案核准
         PTitle = "本案已核准，"
         If pCP01 = "FCP" Then  '外專FCP案核准: 固定CC給特定管制人員
            strHandler = Pub_GetSpecMan("外專告准程序")
         End If
      End If
      strTmpA(1) = "" & rsQD.Fields("cp09")
      strTmpA(2) = "" & rsQD.Fields("cp10")
      strTmpA(3) = "" & rsQD.Fields("cp60")
      strTmpA(4) = "" & rsQD.Fields("cp158")
      strDefTime = Format(ServerTime, "000000") 'Added by Lydia 2023/11/28
      
      If "" & rsQD.Fields("CP60") = "" Then
          '1.上一道工程師案件性質未有請款單號，則自動發Mail
          '收件者: 工程師   副本收受者: 工程師之主管;程序管制人員(Key來函人員不是管制人員也列入收件者);backup
          '主旨: 本案已核准，請工程師儘速處理請款，以利後續告准流程Our Ref: P-060000 [INCOM.1001]
          'Modified by Lydia 20224/06/04 承辦工程師為內專工程師， 則主要收件者，應改為外專對接工程師;副本為下列判發主管的上級主管（只到副理級）
          'strA2 = PUB_GetFCPEngSup("" & rsQD.Fields("CP14"))
          strTo = "" & rsQD.Fields("CP14")
          If Mid(strTo, 4, 1) = "9" Then
             strTo = PUB_GetFCPEngSup(strTo, , , True)
          End If
          strA2 = PUB_GetFCPEngSup(strTo)
          'end 2024/06/04
          
          'Add by Amy 2025/08/05 後續准駁簡單報告=Y,輸C類來函[主旨]最前面加【請簡單報告】-Winfrey
         If Pub_GetField("Patent", "pa01='" & pCP01 & "' And pa02='" & pCP02 & "' And pa03='" & pCP03 & "' And pa04='" & pCP04 & "'", "pa89") = "Y" Then
            stPA89Memo = "【請簡單報告】"
         End If
          
          '主旨
          'Modified by Lydia 2023/11/08 區分核准和非核准
          'Modify by Amy 2025/08/05 後續准駁簡單報告=Y,輸C類來函[主旨]最前面加【請簡單報告】-Winfrey
          If pCP10 = "1001" Then
             strSubject = stPA89Memo & PTitle & "請工程師儘速處理請款，以利後續告准流程Our Ref:" & pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 = "000", "", "-" & pCP03 & "-" & pCP04) & " [INCOM." & pCP10 & "]"
          'Added by Lydia 2023/12/13 1004延期受理,1406實審屆滿前通知=>因為承辦不用報告,所以主旨全部延用
          'Modified by Lydia 2024/04/10 以下性質若有未請款完成，僅在主旨加註：1004延期受理,1406實審屆滿前通知＋1234復審受理通知,1813行政訴訟受理通知,1001核准-PPH(核准的相關收文號掛431 PPH),1001核准-復審申請(核准的相關收文號掛107 復審申請)
          'ElseIf InStr("1004,1406", pCP10) > 0 Then
          ElseIf InStr("1004,1406", pCP10) > 0 And (pCP01 = "FCP" And pCP01 = "FG") Then
             strSubject = stPA89Memo & PTitle
          'end 2023/12/13
          Else
             strSubject = stPA89Memo & PTitle & "請工程師儘速處理請款，以利後續流程Our Ref:" & pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 = "000", "", "-" & pCP03 & "-" & pCP04) & " [INCOM." & pCP10 & "]"
          End If
          'end 2025/08/05
          'end 2023/11/08
          strCC = strA2 & ";" & strNA16 & IIf(strNA16 <> strUserNum, ";" & strUserNum, "")
          '外專FCP案核准: 固定CC給特定管制人員
          If strHandler <> "" And InStr(strCC, strHandler) = 0 Then strCC = strCC & ";" & strHandler
          strCC = strCC & ";backup"
          'Modified by Lydia 2023/11/28 改傳入時間=>Val(strDefTime) + 10
          'Modified by Lydia 2024/06/04 rsQD.Fields("CP14")=> strTo
          strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                 " values( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')," & Val(strDefTime) + 10 & _
                  ",'" & strSubject & "','如旨','" & strCC & "')"
          cnnConnection.Execute strSql, intQ
          PUB_ChkFCPtoDNUPL = True
      Else
          '2.上一道工程師案件性質已有請款單號 , 但卷宗區無REPDN(寄請款函) Or DNUPL(請款單上傳)(有一項就不發email), 則自動發Mail:
          '收件者: 程序管制人員 (Key來函人員不是管制人員也列入收件者) 副本收受者: 程序管制人員主管; backup
          '主旨: 本案已核准，請程序儘速處理請款，以利後續告准流程Our Ref: P-060000 [INCOM.1001]
          'Modified by Lydia 2025/01/16 改成模組取得語法;AND (UPPER(CPP02) LIKE '%.REPDN.%' OR UPPER(CPP02) LIKE '%.DNUPL.%' ) >> PUB_GetFCPforDNsql
          strTmpA(0) = "SELECT CPP01, CPP02 FROM CASEPAPERPDF B " & _
                            "WHERE CPP01 in (select cp09 from caseprogress where cp60='" & rsQD.Fields("CP60") & "')  AND NVL(CPP10,'N') <> 'D' " & PUB_GetFCPforDNsql
          intQ = 1
          Set rsQD = ClsLawReadRstMsg(intQ, strTmpA(0))
          If intQ = 0 Then
            If PUB_ChkFCPtoDNUPLsub(pCP01, pCP02, pCP03, pCP04, strTmpA(1), strTmpA(2), strTmpA(3), strTmpA(4)) = False Then 'Added by Lydia 2023/11/22
              strA2 = PUB_GetFCPProSup(strNA16)
              '主旨
              'Modified by Lydia 2023/11/08 區分核准和非核准
              If pCP10 = "1001" Then
                 strSubject = PTitle & "請程序儘速處理請款，以利後續告准流程Our Ref:" & pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 = "000", "", "-" & pCP03 & "-" & pCP04) & " [INCOM." & pCP10 & "]"
              'Added by Lydia 2023/12/13 1004延期受理,1406實審屆滿前通知=>因為承辦不用報告,所以主旨全部延用
              'Modified by Lydia 2024/04/10 以下性質若有未請款完成，僅在主旨加註：1004延期受理,1406實審屆滿前通知＋1234復審受理通知,1813行政訴訟受理通知,1001核准-PPH(核准的相關收文號掛431 PPH),1001核准-復審申請(核准的相關收文號掛107 復審申請)
              'ElseIf InStr("1004,1406", pCP10) > 0 Then
               ElseIf InStr("1004,1406", pCP10) > 0 And (pCP01 = "FCP" And pCP01 = "FG") Then
                 strSubject = PTitle
              'end 2023/12/13

              Else
                 strSubject = PTitle & "請程序儘速處理請款，以利後續流程Our Ref:" & pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 = "000", "", "-" & pCP03 & "-" & pCP04) & " [INCOM." & pCP10 & "]"
              End If
              'end 2023/11/08
              'CC =>外專FCP案核准: 固定CC給特定管制人員
              If strHandler <> "" And InStr(strNA16 & "," & strUserNum & strCC, strHandler) = 0 Then strCC = strCC & ";" & strHandler
              strCC = strCC & ";backup"
              'Modifie by Lydia 2023/11/28 改傳入時間=>Val(strDefTime) + 10
              strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                     " values( '" & strUserNum & "','" & strNA16 & IIf(strNA16 <> strUserNum, ";" & strUserNum, "") & "',to_char(sysdate,'yyyymmdd')," & Val(strDefTime) + 10 & _
                      ",'" & strSubject & "','如旨','" & strCC & "')"
              cnnConnection.Execute strSql, intQ
              PUB_ChkFCPtoDNUPL = True
            End If 'Added by Lydia 2023/11/22
          End If
      End If
      
      'Modified by Lydia 2024/04/10 以下性質若有未請款完成，僅在主旨加註，不用加發Email(範例3)：1004延期受理,1406實審屆滿前通知＋1234復審受理通知,1813行政訴訟受理通知,C類=>1001核准-PPH(核准的相關收文號掛431 PPH),1001核准-復審申請(核准的相關收文號掛107 復審申請)
      'If PUB_ChkFCPtoDNUPL = True And pCP01 <> "FCP" And pCP01 <> "FG" Then  'FMP寰華案另外通知
      If PUB_ChkFCPtoDNUPL = True And pCP01 <> "FCP" And pCP01 <> "FG" And Not (InStr("1004,1406,1234,1813", pCP10) > 0 Or pKind = "C") Then   'FMP寰華案另外通知
         'Added by Lydia 2023/11/08 若抓到是主動補正（203）和新案的請款單號一致，則不發本封Email=>比照每日批次的StrMenu125外專寄請款函作業稽核(共用模組PUB_ChkFCPtoDNUPLsub)
         If PUB_ChkFCPtoDNUPLsub(pCP01, pCP02, pCP03, pCP04, strTmpA(1), strTmpA(2), strTmpA(3), strTmpA(4)) = False Then
            strA2 = PUB_GetFCPSalesNo(pCP01, pCP02, pCP03, pCP04) '承辦
            '主旨: ◎【FMP(寰華)案核准通知】 Our Ref: P-129322 [INCOM.1001] ，請通知代理人！
            '本案已核准，
            If PTitle = "本案已核准，" And pCP01 = "P" And pCP10 = "1001" Then
                PTitle = "【FMP(寰華)案核准通知】"
            End If
            'Modified by Lydia 2023/11/08 區分核准和非核准
            If pCP10 = "1001" Then
               strSubject = PTitle & "Our Ref:" & pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 = "000", "", "-" & pCP03 & "-" & pCP04) & " [INCOM." & pCP10 & "]，請通知代理人！"
            'Added by Lydia 2023/12/13 1004延期受理,1406實審屆滿前通知=>因為承辦不用報告,所以主旨全部延用
            ElseIf InStr("1004,1406", pCP10) > 0 Then
               strSubject = PTitle
            'end 2023/12/13
            Else
               strSubject = PTitle & "請承辦報告Our Ref:" & pCP01 & "-" & pCP02 & IIf(pCP03 & pCP04 = "000", "", "-" & pCP03 & "-" & pCP04) & " [INCOM." & pCP10 & "]"
            End If
            'end 2023/11/08
            '內文:
            'To:ＸＸＸ（請抓智權管制人員）
            '本案尚有請款程序未完成，待程序E出請款後，再報告核准，謝謝。
            'To:ＸＸＸ（請抓程序管制人員）
            '  請儘速完成請款程序，待E出請款後，通知承辦人員報告核准，謝謝。
            strTmpA(0) = "To " & PUB_ReadUserData(strA2) & ": " & vbCrLf & _
                             "本案尚有請款程序未完成，待程序E出請款後，再報告核准，謝謝。" & vbCrLf & vbCrLf & _
                             "To " & PUB_ReadUserData(strNA16) & ": " & vbCrLf & _
                             "請儘速完成請款程序，待E出請款後，通知承辦人員報告核准，謝謝。"
            If pCP10 <> "1001" Then strTmpA(0) = Replace(strTmpA(0), "核准", "") 'Added by Lydida 2023/11/08 非核准
            
            'Modified by Lydia 2023/06/14 +CC: backup(mc09)
            'Modifie by Lydia 2023/11/28 改傳入時間=>Val(strDefTime)
            strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                       " values( '" & strUserNum & "','" & strA2 & ";" & strNA16 & IIf(strNA16 <> strUserNum, ";" & strUserNum, "") & "',to_char(sysdate,'yyyymmdd')," & Val(strDefTime) & _
                        ",'" & strSubject & "','" & ChgSQL(strTmpA(0)) & "','backup' )"
            cnnConnection.Execute strSql, intQ
         End If
         'end 2023/11/08
      End If
      
JumpToNext: 'Added by Lydia 2023/11/08
   End If
   Set rsQD = Nothing
   
   Exit Function
   
ErrHandle:
   If Err.Number <> 0 Then
      MsgBox Err.Description, vbCritical, "判斷上一道工程師案件性質是否已經請款"
   End If
End Function

'Added by Lydia 2023/11/15 設定商標/特殊商標的下拉清單
Public Sub Pub_SetTMcombo(ByVal pType As String, ByRef pCmb As Control, Optional ByVal pTxt As String, Optional ByVal bolChina As Boolean = False, Optional ByRef p_ItemDataList As String)
Dim stSQL As String, intR As Integer
Dim rsQuery As ADODB.Recordset
Dim strFind As String

   pCmb.Clear
   p_ItemDataList = ""
   stSQL = "select * from " & IIf(pType = "1", "PatentTrademarkMap where PTM01='2'", "SpecialPatentTrademark") & " order by 2 desc"
   intR = 1
   Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      rsQuery.MoveFirst
      Do While Not rsQuery.EOF
         pCmb.AddItem rsQuery.Fields(1) & "  " & IIf(bolChina = True, "" & rsQuery.Fields(3), "" & rsQuery.Fields(2)), 0
         p_ItemDataList = rsQuery.Fields(1) & IIf(p_ItemDataList <> "", ",", "") & p_ItemDataList
         If pTxt <> "" And pTxt = "" & rsQuery.Fields(1) Then
            strFind = pCmb.ListCount
         End If
         rsQuery.MoveNext
      Loop
   End If
   Set rsQuery = Nothing
   If strFind <> "" Then
      pCmb.ListIndex = pCmb.ListCount - Val(strFind)
   Else
      pCmb.Text = ""
   End If

End Sub

'Added by Lydia 2024/03/05 FCT和T案電子送件申請書的申請人名稱; 從PUB_GetApplFCT_EData抽出
Public Function PUB_GetApplT_CNAME(ByVal pCustNo As String) As String
Dim strB1 As String, intB As Integer
Dim rsBD As New ADODB.Recordset
   
   PUB_GetApplT_CNAME = ""
   pCustNo = ChangeCustomerL(pCustNo)
   strB1 = "select na81,cu15,substr(cu10,1,3) as cu10,nvl(cu04,nvl(cu05,cu06)) as cname from customer,nation where cu01='" & Mid(pCustNo, 1, 8) & "' and cu02='" & Mid(pCustNo, 9, 1) & "' and cu10=na01(+) "
   intB = 1
   Set rsBD = ClsLawReadRstMsg(intB, strB1)
   If intB = 1 Then
      If "" & rsBD.Fields("cu15") = "1" Then '1.公司
         If "" & rsBD.Fields("cu10") < "010" Then
            PUB_GetApplT_CNAME = "" & rsBD.Fields("cname")
         Else
            PUB_GetApplT_CNAME = "" & rsBD.Fields("na81") & rsBD.Fields("cname")
         End If
      Else
         '柏翰提個人的姓和名中間要有,號
         If "" & rsBD.Fields("CU15") = "0" Then '自然人
            PUB_GetApplT_CNAME = "" & Replace("" & rsBD.Fields("na81"), "商", "籍") & PUB_ConvertNameFormat("" & rsBD.Fields("cname"))
         Else
            PUB_GetApplT_CNAME = "" & rsBD.Fields("cname")
         End If
      End If
   End If
   Set rsBD = Nothing
End Function

'Added by Lydia 2024/03/06 外專發文：內專協辦工程師完成送件之後，需通知外專工程師進行請款
'Memo by Lydia 2024/04/09 FMP案不用通知--- Phoebe
'Memo by Lydia 2024/04/18 FCP案直接併入frm060104_k的Outlook，所以也不用---Sharon
'Mark by Lydia 2024/05/08 先Mark
'Public Sub Pub_SetEngMail(ByVal pCP09 As String)
'Dim strB1 As String, intB As Integer
'Dim rsBD As New ADODB.Recordset
'Dim strSub As String, strContent As String
''Added by Lydia 2024/03/12
'Dim objOutLook As Object
'Dim objMail As Object
'
'   'Modified by Lydia 2024/03/13 +PA150
'   strB1 = "select cp01,cp02,cp03,cp04,cp10,decode(pa09,'000',cpm03,cpm04) as cp10name,cp06,cp07,cp14,cp13,cp16, " & _
'           "s3.st52,ep02,ep04,s1.st03 as ep04d,ep40,s2.st03 as ep40d,PA150 " & _
'           "from caseprogress, engineerprogress, staff s1, staff s2, staff s3,casepropertymap,patent " & _
'           "where cp09 ='" & pCP09 & "' and substr(cp14,4,1)='9' " & _
'           "and cp09=ep02(+) and ep04=s1.st01(+) and ep40=s2.st01(+) and cp14=s3.st01(+) " & _
'           "and cp01=cpm01(+) and cp10=cpm02(+) and cp01=pa01(+) and cp02=pa02(+) and cp03=pa03(+) and cp04=pa04(+) "
'   intB = 1
'   Set rsBD = ClsLawReadRstMsg(intB, strB1)
'   If intB = 1 Then
'      '當核稿人EP04和判發人EP40全部為內專人員，需通知外專工程師進行請款
'      'Modified by Lydia 2024/03/11 沒有核稿人和承辦人也需要發Email
'      'If "" & rsBD.Fields("ep04d") <> "" And Left("" & rsBD.Fields("ep04d"), 2) <> "F2" And "" & rsBD.Fields("ep40d") <> "" And Left("" & rsBD.Fields("ep40d"), 2) <> "F2" Then
'      'Modified by Lydia 2024/03/12 排除209檢視中說、210製作中說、235核對中說格式、242製作外文提申本。
'      strSub = Pub_GetSpecMan("外專發文-內專工程師案件排除通知性質")  'Memo by Lydia 2024/05/08 已刪除特殊設定:209,210,235,242
'      If Left("" & rsBD.Fields("ep04d"), 2) = "F2" Or Left("" & rsBD.Fields("ep40d"), 2) = "F2" Or InStr(strSub, "" & rsBD.Fields("cp10")) > 0 Then
'      Else
'      'end 2024/03/11
'         If "" & rsBD.Fields("st52") <> "" Then
'            strSub = "【已送件完成_" & rsBD.Fields("cp10name") & IIf("" & rsBD.Fields("cp07") = "", "", "(法限:" & ChangeWStringToTDateString("" & rsBD.Fields("cp07"))) & "】請進行請款 Our Ref: " & rsBD.Fields("cp01") & "-" & rsBD.Fields("cp02") & IIf(rsBD.Fields("cp03") & rsBD.Fields("cp04") <> "000", "-" & rsBD.Fields("cp03") & "-" & rsBD.Fields("cp04"), "")
'            'Added by Lydia 2024/03/13
'            If "" & rsBD.Fields("pa150") = "4" Then
'               strSub = "【機械設計組】" & strSub
'            End If
'            'end 2024/03/13
'            strContent = "1.請工程師處理請款，送件附檔可參卷宗區" & vbCrLf
'            strContent = strContent & "2.請款信、金額、附件請提供給承辦（" & GetStaffName("" & rsBD.Fields("cp13")) & "）及程序人員（" & GetStaffName(PUB_GetFCPHandler(rsBD.Fields("cp01"), rsBD.Fields("cp02"), rsBD.Fields("cp03"), rsBD.Fields("cp04"))) & "）" & vbCrLf
'            'Modified by Lydia 2024/03/12 改使用Outlook草稿
'            'strB1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08) values ('" & strUserNum & "','" & rsBD.Fields("st52") & "',to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss'),'" & ChgSQL(strSub) & "','" & ChgSQL(strContent) & "')"
'            'cnnConnection.Execute strB1
'            '---------------------------------
'            '呼叫新郵件：
'            Set objOutLook = CreateObject("Outlook.Application")
'            Set objMail = objOutLook.CreateItem(0)
'
'            '轉HTML格式
'            strContent = Replace(strContent, "新細明體", "Times New Roman")
'            '&nbsp; 不換行空格
'            '&thinsp; 窄空格
'            '單純只是想要輸入空白？ &nbsp; 就對了
'            '&emsp; 全形空格
'            '&ensp; 半形空格
'            'strContent = Replace(strContent, "　", "&emsp;") '&emsp; 全形空格
'            strContent = Replace(strContent, " ", "&thinsp;") '&ensp; 半形空格
'            strContent = Replace(strContent, vbCrLf, "<BR>")
'            With objMail
'               .To = "" & rsBD.Fields("st52")
'               .Subject = strSub
'               .HTMLBody = strContent
'               .Display
'            End With
'
'            Set objMail = Nothing
'            'end 2024/03/12
'         End If
'      End If
'   End If
'   Set rsBD = Nothing
'
'End Sub
'end 2024/05/08

'Added by Lydia 2024/03/08 收文號之案件性質名稱
'Modified by Lydia 2024/03/29 +收文號案件性質代號pSType
'Modified by Lydia 2024/04/01 +指定收文號之類別pAKind
Public Function Pub_GetNoToCPM(ByVal pStatus As String, ByVal pSNo As String, Optional ByRef pSType As String, Optional ByVal pAKind As String) As String
Dim intP As Integer, strP1 As String
Dim rsAD As New ADODB.Recordset
Dim strNow As String 'Added by Lydia 2024/04/01

   Pub_GetNoToCPM = ""
   pSType = ""
   If pSNo <> "" Then
      strNow = pSNo
      If pStatus = "1" Then
        strP1 = "select CP01||'-'||CP02||'-'||CP03||'-'||CP04 as caseno,CP09 as pBNO,CP10 as ptype from caseprogress where cp09 ='" & strNow & "' "
      Else
JumpToReQuery: 'Added by Lydia 2024/04/01
        strP1 = "select c1.CP01||'-'||c1.CP02||'-'||c1.CP03||'-'||c1.CP04 as caseno,c2.CP09 as pBNO,c2.CP10 as ptype " & _
                "from caseprogress c1,caseprogress c2 where c1.CP09 ='" & strNow & "' and c1.cp43=c2.cp09(+) "
      End If
      intP = 1
      Set rsAD = ClsLawReadRstMsg(intP, strP1)
      If intP = 1 Then
         'Added by Lydia 2024/04/01
         If pStatus = "2" And pAKind <> "" And pAKind <> Mid("" & rsAD.Fields("pBNO"), 1, 1) Then
            strNow = "" & rsAD.Fields("pBNO")
            GoTo JumpToReQuery
         End If
         'end 2024/04/01
         Pub_GetNoToCPM = GetPrjState4("" & rsAD.Fields("caseno"), "" & rsAD.Fields("ptype"))
         pSType = "" & rsAD.Fields("ptype")    'Added by Lydia 2024/03/29
      End If
   End If
   Set rsAD = Nothing
End Function

'Add By Sindy 2024/4/17 是否需要檢查ST16組別
Public Function PUB_NeedChkFCPST16(strEmp As String) As Boolean
Dim strTmp As String, intA As Integer, rsAD As New ADODB.Recordset 'Added by Lydia 2024/05/15

   PUB_NeedChkFCPST16 = True
   'Modify By Sindy 2024/3/5 排除員編第4號是9的人員(支援人員)
   If Mid(Trim(Left(strEmp, 6)), 4, 1) = "9" Then
      PUB_NeedChkFCPST16 = False
      GoTo EXITSUB
   End If
   'Modify By Sindy 2024/3/22 排除機械組
   strTmp = "select st01,st16 from staff where st01='" & strEmp & "'"
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strTmp)
   If intA = 1 Then
      If "" & rsAD("st16") = "4" Then
         PUB_NeedChkFCPST16 = False
         GoTo EXITSUB
      End If
   End If
   
'Added by Lydia 2024/05/15
EXITSUB:
   Set rsAD = Nothing
'end 2024/05/15
End Function

'Add By Sindy 2024/4/25
Public Function PUB_FCPChkCP141(ByVal strCP09 As String, Optional ByRef strReVal As String) As Boolean
Dim strTmp As String, intA As Integer, rsAD As New ADODB.Recordset 'Added by Lydia 2024/05/15
   
   PUB_FCPChkCP141 = True
   strReVal = ""
   
   'Add By Sindy 2022/3/4 若有設定指定送件日，在產生申請書時可彈提醒
   strTmp = "select * from caseprogress where cp09='" & strCP09 & "'"
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strTmp)
   If intA = 1 Then
      If "" & rsAD.Fields("cp142") <> "" Then '有指定送件日期
         '程序人員在產生申請書時(僅限程序產生的申請書)，若那道進度檔有設定指定送件日"當天"or "之後"，
         '若產生申請書的系統日，小於指定送件"當天"or "之後"的日期，則彈提醒
         If ("" & rsAD.Fields("cp164") = "1" Or "" & rsAD.Fields("cp164") = "") Or "" & rsAD.Fields("cp164") = "3" Then
            If rsAD.Fields("cp142") > strSrvDate(1) Then
               MsgBox "此道設定指定" & ChangeWStringToTDateString(rsAD.Fields("cp142")) & "日" & IIf("" & rsAD.Fields("cp164") = "3", "之後", "當天") & "送件，請注意。"
            End If
         End If
      End If
      
      'Add By Sindy 2024/4/25
      'Modify By Sindy 2024/5/2 因年費也有暫不送件的機制,為不衝突;
      '增加檢查除了有勾選「暫不送件」同時備註必須有「需待客戶最終指示」等字眼,才是要管制的
      'Modify By Sindy 2024/5/28 "暫不送件" 改用獨立欄位存放(cp176)
      If "" & rsAD.Fields("cp176") = "Y" And _
         InStr("" & rsAD.Fields("cp64"), "需待客戶最終指示") > 0 Then
         'Add By Sindy 2024/5/2 當年費,進度檔有「取消待客戶最終指示」等字眼時，則不須控管
         If Not ("" & rsAD.Fields("cp10") = 年費 And InStr("" & rsAD.Fields("cp64"), "取消待客戶最終指示") > 0) Then
         '2024/5/2 END
            If MsgBox("本案需待客戶最終指示，請確認是否已獲客戶最終指示？" & vbCrLf & vbCrLf & _
                      "按【是】系統會自動取消暫不送件的設定（請注意!!!）", vbInformation + vbYesNo + vbDefaultButton2) = vbYes Then
               strReVal = "Y"
               'Modify By Sindy 2024/5/2 年費不能點掉暫不送件,僅加註備註
               If "" & rsAD.Fields("cp10") = 年費 Then
                  strSql = "update caseprogress set cp64='於" & ChangeWStringToTDateString(strSrvDate(1)) & "取消待客戶最終指示;'||cp64 where cp09='" & strCP09 & "'"
               Else
               '2024/5/2 END
                  strSql = "update caseprogress set cp176=null,cp64='於" & ChangeWStringToTDateString(strSrvDate(1)) & "取消待客戶最終指示;'||cp64 where cp09='" & strCP09 & "'"
               End If
               Pub_SeekTbLog strSql
               cnnConnection.Execute strSql
               MsgBox "已取消暫不送件!", vbInformation
            Else
               strReVal = "N"
               PUB_FCPChkCP141 = False
               Exit Function
            End If
         End If
      '2024/4/25 END
      End If
   End If
   '2022/3/4 END
   
   Set rsAD = Nothing 'Added by Lydia 2024/05/15
End Function

'Added by Lydia 2024/05/06 比對兩個串接字串，不重複才合併
Public Function Pub_GetTwoListChk(ByVal pVAL01 As String, ByVal pVAL02 As String, Optional ByVal pSignal As String = ",") As String
Dim intA As Integer, tmpArrA As Variant
Dim intB As Integer, tmpArrB As Variant
Dim strMid As String

   If Trim(pVAL01) = "" And Trim(pVAL02) <> "" Then
      strMid = pSignal & pVAL02
   ElseIf Trim(pVAL01) <> "" And Trim(pVAL02) = "" Then
      strMid = pSignal & pVAL01
   Else
      tmpArrA = Split(pVAL01, pSignal)
      For intA = 0 To UBound(tmpArrA)
         strMid = strMid & pSignal & Trim(tmpArrA(intA))
      Next intA
      tmpArrB = Split(pVAL02, pSignal)
      For intB = 0 To UBound(tmpArrB)
         If Trim(tmpArrB(intB)) <> "" Then
            If InStr(strMid & pSignal, pSignal & Trim(tmpArrB(intB)) & pSignal) = 0 Then
               strMid = strMid & pSignal & Trim(tmpArrB(intB))
            End If
         End If
      Next intB
   End If
      
   If strMid <> "" Then
      Pub_GetTwoListChk = Mid(strMid, 2)
   End If
   
   tmpArrA = Empty
   tmpArrB = Empty
End Function

'Added by Morgan 2012/12/27
'Modified by Morgan 2013/9/17 +bolNoAsk
'Modified by Morgan 2015/9/11 +bolIsRefCase=aPA是否為母案
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function PUB_GetDivCaseState(aPA() As String, aCP27 As String, Optional bolNoAsk As Boolean, Optional bolIsRefCase As Boolean = False) As String
   Dim stRtn As String, stSQL As String, intQ As Integer
   Dim adoquery As ADODB.Recordset
   
   'Modified by Morgan 2015/9/11 +母案為分割案判斷
   stSQL = "select a.cp09, a.cp10,a.cp24,a.cp25,pa16,pa20,b.cp09 refNo,c.cp09 refNo2"
   '傳入母案案號(分案母案未存檔前檢查)
   If bolIsRefCase Then
      stSQL = stSQL & " from patent, caseprogress a, caseprogress b, caseprogress c" & _
         " where pa01='" & aPA(1) & "' and pa02='" & aPA(2) & "' and pa03='" & aPA(3) & "' and pa04='" & aPA(4) & "'"
   Else
      stSQL = stSQL & " from divisioncase,patent,caseprogress a,caseprogress b,caseprogress c" & _
         " where dc01='" & aPA(1) & "' and dc02='" & aPA(2) & "' and dc03='" & aPA(3) & "' and dc04='" & aPA(4) & "'" & _
         " and pa01(+)=dc05 and pa02(+)=dc06 and pa03(+)=dc07 and pa04(+)=dc08"
   End If
   stSQL = stSQL & _
      " and a.cp01(+)=pa01 and a.cp02(+)=pa02 and a.cp03(+)=pa03 and a.cp04(+)=pa04 and a.cp10 in ('101','307')" & _
      " and b.cp01(+)=pa01 and b.cp02(+)=pa02 and b.cp03(+)=pa03 and b.cp04(+)=pa04 and b.cp57(+) is null and instr('107,435',b.cp10(+))>0" & _
      " and c.cp01(+)=pa01 and c.cp02(+)=pa02 and c.cp03(+)=pa03 and c.cp04(+)=pa04 and c.cp57(+) is null and c.cp10(+)='416'"
         
   intQ = 1
   Set adoquery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      With adoquery
      
      '初審階段
      '1.A類收文且未審定
      If .Fields("cp10") = "101" And .Fields("cp09") < "B" And IsNull(.Fields("cp24")) And IsNull(.Fields("pa16")) Then
         stRtn = "Y"
      '2.發明申請核准
      ElseIf .Fields("cp10") = "101" And .Fields("cp24") = "1" Then
         stRtn = "Y"
      '3.已收再審或續行母案再審
      ElseIf Not IsNull(.Fields("refNo")) Then
         stRtn = "N"
      'Added by Morgan 2015/9/16
      'Modified by Morgan 2015/10/1
      '4.A類分割未核駁且有收實體審查
      ElseIf .Fields("cp10") = "307" And .Fields("cp09") < "B" And Not IsNull(.Fields("refNo2")) And "" & .Fields("cp24") <> "2" Then
      'end 2015/10/1
         stRtn = "Y"
      'en 2015/9/16
      '非初審階段(發文日輸過去可能會誤判故加比較核駁日)
      ElseIf .Fields("cp24") = "2" And .Fields("cp25") < aCP27 Then
         stRtn = "N"
      End If
      End With
   End If
   
   If stRtn = "" And bolNoAsk = False Then
   
      intQ = MsgBox("資訊不足無法自動設定!!請問本案是否為初審階段提分割??" & vbCrLf & vbCrLf & "注意：本設定將決定日後核准函是否提醒可分割。", vbYesNoCancel + vbDefaultButton3)
      If intQ = vbYes Then
         stRtn = "Y"
      ElseIf intQ = vbNo Then
         stRtn = "N"
      End If
   End If
   PUB_GetDivCaseState = stRtn
   
   Set adoquery = Nothing
End Function

'Add by Morgan 2012/11/8
'檢查分割母案
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function PUB_CheckDivCase(oTxtCaseNo As Object, aPA() As String, Optional ByRef stPA09 As String, Optional ByRef stPA08 As String, Optional ByRef iErrCode As Integer) As Boolean
   Dim stSQL As String
   Dim rsQuery As ADODB.Recordset, intQ As Integer
   'Add by Lydia 2014/10/21
  ' Dim stPA08 As String, stDate As String
    Dim stDate As String
On Error GoTo flgErr
   
   If (oTxtCaseNo(1) = "" Or oTxtCaseNo(2) = "") Then
      MsgBox "分割母案本所案號輸入錯誤！", vbExclamation
      Exit Function
   End If
   
   oTxtCaseNo(1) = Trim(oTxtCaseNo(1))
   oTxtCaseNo(2) = Right("00000" & oTxtCaseNo(2), 6)
   oTxtCaseNo(3) = Right("0" & oTxtCaseNo(3), 1)
   oTxtCaseNo(4) = Right("00" & oTxtCaseNo(4), 2)
   
   If (oTxtCaseNo(1) = aPA(1) And oTxtCaseNo(2) = aPA(2) And oTxtCaseNo(3) = aPA(3) And oTxtCaseNo(4) = aPA(4)) Then
      'modify by sonia 2017/3/24
     ' MsgBox "分割案不可為母案！", vbExclamation
      MsgBox "母案案號不可為分割案本身！", vbExclamation
      Exit Function
   End If
   
   'Modified by Morgan 2014/4/7 +考慮母案是分割案
   'Modified by Morgan 2019/10/3 +考慮母案是再審准
   'Modified by Morgan 2024/6/21 +新型102的分割 Ex:FCP-071896(FCP-069544)
   stSQL = "select PA08, PA09,PA16,nvl(nvl(c5.cp05,c3.cp05),pa20) RD1st,c2.cp25 RD2nd,c1.cp09 cp09_101,c2.cp09 cp09_107,c4.cp09 cp09_307, PA163" & _
      ",nvl(nvl(nvl(c6.cp05,c5.cp05),c3.cp05),pa20) RD3st from patent,caseprogress c1,caseprogress c2,caseprogress c3,caseprogress c4,caseprogress c5" & _
      ",caseprogress c6 where pa01='" & ChgSQL(oTxtCaseNo(1)) & "'" & _
      " and pa02='" & ChgSQL(oTxtCaseNo(2)) & "' and  pa03='" & ChgSQL(oTxtCaseNo(3)) & "' and pa04='" & ChgSQL(oTxtCaseNo(4)) & "'" & _
      " and c1.cp01(+)=pa01 and c1.cp02(+)=pa02 and c1.cp03(+)=pa03 and c1.cp04(+)=pa04 and instr('101,102',c1.cp10(+))>0 and c1.cp27(+)>0" & _
      " and c2.cp01(+)=pa01 and c2.cp02(+)=pa02 and c2.cp03(+)=pa03 and c2.cp04(+)=pa04 and c2.cp10(+)='107' and c2.cp27(+)>0" & _
      " and c3.cp43(+)=c1.cp09 and c3.cp10(+)='1001'" & _
      " and c4.cp01(+)=pa01 and c4.cp02(+)=pa02 and c4.cp03(+)=pa03 and c4.cp04(+)=pa04 and c4.cp10(+)='307' and c4.cp27(+)>0" & _
      " and c5.cp43(+)=c4.cp09 and c5.cp10(+)='1001'" & _
      " and c6.cp43(+)=c2.cp09 and c6.cp10(+)='1001'"
      intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      stPA08 = "" & rsQuery.Fields("PA08")
      stPA09 = "" & rsQuery.Fields("PA09")
      If stPA09 <> aPA(9) Then
         MsgBox "分割案與母案的申請國家需相同！", vbExclamation
      ElseIf stPA08 <> aPA(8) Then
         'Add by Lydia 2014/10/21 當分割案與母案的專利種類不同，自動更新為相同種類
         If iErrCode = -1 Then
            MsgBox "分割案與母案的專利種類需相同，系統已將分割案自動更新為與母案相同的專利種類！", vbExclamation
            iErrCode = 1
         Else
            MsgBox "分割案與母案的專利種類需相同！", vbExclamation
         End If
         
      'Added by Morgan 2012/12/25 --靜芳,Tammy
      'Modified by Morgan 2019/10/3 108.11.1 新法發明/新型准後3個月內可提分割(原發明初審准內30日),108.8.1以後收文分割適用
      'ElseIf stPA09 = "000" And stPA08 <> "1" And Not IsNull(rsQuery.Fields("PA16")) Then
      '   MsgBox "台灣新型/設計母案已審定不可分割！", vbExclamation
      'ElseIf stPA09 = "000" And stPA08 = "1" And rsQuery("RD2nd") > 0 Then
      '   MsgBox "台灣發明母案再審已審定不可分割！", vbExclamation
      'ElseIf stPA09 = "000" And stPA08 = "1" And rsQuery("PA16") = "1" Then
      '   If Not IsNull(rsQuery("cp09_107")) Then
      '      MsgBox "台灣發明母案已再審核准不可分割！", vbExclamation
      '   ElseIf rsQuery("RD1st") < 20121202 Then
      '      MsgBox "台灣發明母案已核准不可分割！(適用102年以前舊法案件)", vbExclamation
      '   'Modified by Morgan 2012/12/25
      '   ElseIf rsQuery("RD1st") >= 20121202 And IsNull(rsQuery("cp09_107")) Then
      '
      '      'Modified by Morgan 2014/4/7 +考慮母案是分割案
      '      'If IsNull(rsQuery("cp09_101")) Then
      '      If Not IsNull(rsQuery("cp09_307")) And IsNull(rsQuery("PA163")) Then
      '         MsgBox "台灣發明母案已核准但非初審階段提的分割，不可再分割！", vbExclamation
      '
      '      ElseIf IsNull(rsQuery("cp09_101")) And IsNull(rsQuery("cp09_307")) Then
      '      'end 2014/4/7
      '         MsgBox "母案已核准但無 [發明申請] 或 [再審申請] 發文致無法判斷核准階段," & vbCrLf & "若為中間來所案件請補收文 [發明申請] (初審核准) 或 [再審申請](再審核准) !!!", vbExclamation
      '
      '      Else
      '         stDate = CompDate(2, 30, rsQuery("RD1st"))
      '         stDate = PUB_GetWorkDay1(stDate, False)
      '         If stDate < strSrvDate(1) Then
      '            MsgBox "台灣發明母案已初審核准超過30天不可分割！", vbExclamation
      '         Else
      '            PUB_CheckDivCase = True
      '         End If
      '      End If
      '   End If
      ElseIf stPA09 = "000" And Not IsNull(rsQuery.Fields("PA16")) Then
         '設計案
         If stPA08 = "3" Then
            'Modified by Morgan 2022/6/1 未提再審或再審已審定時才不能再提分割
            'MsgBox "台灣設計母案已審定不可分割！", vbExclamation
            If rsQuery.Fields("PA16") = "1" Then
               MsgBox "台灣設計母案已核准不可分割！", vbExclamation
            ElseIf rsQuery("RD2nd") > 0 Then
               MsgBox "台灣設計母案再審已審定不可分割！", vbExclamation
            ElseIf IsNull(rsQuery.Fields("cp09_107")) Then
               MsgBox "台灣設計母案核駁後尚未提再審不可分割！", vbExclamation
            Else
               PUB_CheckDivCase = True
            End If
            'end 2022/6/1
         '發明/新型
         Else
            '已核駁
            If rsQuery("PA16") = "2" Then
               '新型案(申請審定前或核准3個月內)
               If stPA08 = "2" Then
                  MsgBox "台灣新型母案已核駁不可分割！", vbExclamation
               '發明案(再審審定前或核准3個月內)
               Else
                  '再審駁
                  If rsQuery("RD2nd") > 0 Then
                     MsgBox "台灣發明母案再審已核駁不可分割！", vbExclamation
                  'Added by Morgan 2022/6/1
                  ElseIf IsNull(rsQuery.Fields("cp09_107")) Then
                     'Added by Morgan 2023/8/30
                     '母案初審核駁後，母案有收文"延期(404)-相關收文號為核駁"已發文，視為已提出" 再審(107)"狀況，判斷
                     '1. 下一程序之再審申請法限未逾期，則無限制" 分割(307)"的分案
                     '2. 下一程序之再審申請法限已逾期，限制分案，且彈訊息："母案之再審申請法限已逾期，請確認是否可提分割？"
                     stSQL = "select np09 from nextprogress,caseprogress where np02='" & oTxtCaseNo(1) & "' and np03='" & oTxtCaseNo(2) & "'" & _
                        " and np04='" & oTxtCaseNo(3) & "' and np05='" & oTxtCaseNo(4) & "' and np07='107' and cp43(+)=np01 and cp10='404' and cp27>0"
                     intQ = 1
                     Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
                     If intQ = 1 Then
                        If rsQuery.Fields("np09") >= strSrvDate(1) Then
                           PUB_CheckDivCase = True
                        Else
                           If MsgBox("母案之再審申請法限已逾期，請確認是否可提分割？" & vbCrLf & "(是:繼續，否:取消)", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                              PUB_CheckDivCase = True
                           End If
                        End If
                     Else
                     'end 2023/8/30
                        'Added by Morgan 2023/11/7
                        '母案有收文再審(107)未發文且未逾法限--Gill
                        stSQL = "select cp09 from caseprogress where cp01='" & oTxtCaseNo(1) & "' and cp02='" & oTxtCaseNo(2) & "'" & _
                           " and cp03='" & oTxtCaseNo(3) & "' and cp04='" & oTxtCaseNo(4) & "' and cp10='107' and cp158=0 and cp159=0 and cp07>=" & strSrvDate(1)
                        intQ = 1
                        Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
                        If intQ = 1 Then
                           PUB_CheckDivCase = True
                        Else
                        'end 2023/11/7
                        
                           MsgBox "台灣發明母案核駁後尚未提再審不可分割！", vbExclamation
                           
                        End If 'Added by Morgan 2023/11/7
                     End If
                  'end 2022/6/1
                  Else
                     PUB_CheckDivCase = True
                  End If
               End If
            '已核准(3個月內)
            Else
               stDate = CompDate(1, 3, rsQuery("RD3st"))
               stDate = PUB_GetWorkDay1(stDate, False)
               If stDate < strSrvDate(1) Then
                  MsgBox "台灣發明/新型母案已核准超過3個月不可分割！", vbExclamation
               Else
                  PUB_CheckDivCase = True
               End If
            End If
         End If
      'end 2019/10/3
      Else
         PUB_CheckDivCase = True
      End If
   Else
      MsgBox "分割母案本所案號不存在！", vbExclamation
   End If
   
flgErr:
   Set rsQuery = Nothing
   If Err.Number <> 0 Then MsgBox Err.Description, vbCritical
    
End Function


'Add by Amy 2023/02/13 取得/釋放資料
'Move by Lydia 2024/05/06 從basUpdate搬過來
'intChoose:0-秀訊息,不可操作/1-秀訊息,可操作/2-不秀訊息,不可操作/3-不秀訊息,可操作
'p_Status:A-新增及確認 / C-只有確認 /D-刪除
'stFormCaption:表單中文名稱
'stOtherKey:其他Key Word
Public Function Pub_ChkLock(ByVal intChoose As Integer, ByVal stFormNo As String, ByVal p_Status As String, Optional ByRef stFormCaption As String = "", Optional ByVal stOtherKey As String = "") As Boolean
    Dim RsQ As New ADODB.Recordset, intQ As Integer
    Dim stKeyLR01 As String, stAllMsg As String, stMsg As String
    Dim stCmd As String, stQ As String, stWhere_Q As String, stD As String, stWhere_D As String
    
On Error GoTo ErrHand
    
    Pub_ChkLock = False
    
    Select Case UCase(stFormNo)
        Case "FRM110101_2", "FRM210149_1" '解除期限,結案單
            stKeyLR01 = stFormNo & "-" & strUserNum & "-" & stOtherKey
            stWhere_D = "And LR01='" & stKeyLR01 & "' "
            stWhere_Q = "And LR01 Like '" & stFormNo & "-%' And InStr(LR01,'" & stOtherKey & "')>0 And LR02<>'" & strUserNum & "' "
            If Right(stOtherKey, 3) = "000" Then
                stAllMsg = Mid(stOtherKey, 1, Len(stOtherKey) - 3)
            Else
                stAllMsg = stOtherKey
            End If
            stAllMsg = "正在" & stFormCaption & "作業" & vbCrLf & _
                              "操作【" & stAllMsg & "】案子！"
        Case "FRM010033" '掃瞄資料匯入(檔案室/內商 使用)
            stKeyLR01 = stFormNo & "-" & strUserNum
            stWhere_D = "And LR01='" & stKeyLR01 & "' "
            stWhere_Q = "And LR01 Like '" & stFormNo & "-%' And st03='" & Pub_StrUserSt03 & "' "
            stAllMsg = "正使用" & stFormCaption & "作業！"
    End Select
    
    stD = "Delete From LockRec Where LR02='" & strUserNum & "' "
    stQ = "Select st02 From LockRec,Staff Where LR02=st01(+) "
    
    '刪除
    If p_Status = "D" Then
        stCmd = stD & stWhere_D
        cnnConnection.Execute stCmd
    Else
        If p_Status <> "C" Then
            '先刪除
            stCmd = stD & stWhere_D
            cnnConnection.Execute stCmd
        End If
        '查詢使用狀況
        stCmd = stQ & stWhere_Q
        intQ = 1
        Set RsQ = ClsLawReadRstMsg(intQ, stCmd)
        If intQ = 1 Then
            stMsg = "" & RsQ.GetString(adClipString, , , ",")
            If intChoose <= 1 Then
                MsgBox "【" & Mid(stMsg, 1, Len(stMsg) - 1) & "】" & vbCrLf & _
                              stAllMsg, vbInformation
            End If
        Else
            Pub_ChkLock = True
        End If
        If p_Status <> "C" Then
            '新增目前記錄
            stCmd = "Insert Into LockRec(LR01,LR02,LR03) Values ('" & stKeyLR01 & "','" & strUserNum & "',to_char(sysdate,'YYYYMMDDHH24MISS'))"
            cnnConnection.Execute stCmd
        End If
    End If
    If intChoose = 1 Or intChoose = 3 Then Pub_ChkLock = True
    Exit Function
   
ErrHand:
    If Err.Number <> 0 Then
        MsgBox Err.Description, vbCritical, "取得/釋放資料異動"
    End If
End Function

'Added by Lydia 2016/08/15
'例外承辦期限控制
'Modify By Sindy 2017/5/9 + Optional ByVal pCP06 As String
'Modified by Morgan 2018/5/22 + Optional ByVal pRefCP10 As String '相關收文號案件性質(只有核駁會傳及用到)
'Modified by Morgan 2018/12/19 + Optional ByVal pCP133 As String '官方發文日
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Sub Pub_SetExceptCP48(ByVal pFA01 As String, ByVal PCU01 As String, ByVal pCP10 As String, ByVal pCP05 As String, ByRef updDTXT As TextBox, _
      Optional ByVal nCP10 As String, Optional ByVal pCP06 As String, Optional ByVal pRefCP10 As String, Optional ByVal pCP133 As String)
   Dim stDate As String 'Added by Morgan 2020/11/13
   Dim iDays As Integer 'Added by Morgan 2022/7/12
'pFa01: FC代理人
'updDTXT: 承辦期限
'pCP10: 案件性質
'nCP10: 下一程序
'pCu01: 客戶
'pCP05: 收文日
   '先正達OA承辦期限設7個工作天,若下一程序為 804,501-509時設2個工作天(24Hr)
   If InStr("Y4830900,Y4830901,Y4830902,Y4830903,Y4830904,Y4830905,Y4830908,Y5132600", Left(pFA01 & "000", 8)) > 0 Then
      If nCP10 = "804" Or nCP10 >= "501" And nCP10 <= "509" Then
         updDTXT = TransDate(CompWorkDay(2, pCP05, 0), 1)
      'modify by sonia 2025/4/22 +1009部分准駁
      ElseIf (pCP10 = "1202" Or pCP10 = "1227" Or pCP10 = 核駁 Or pCP10 = "1009" Or pCP10 = 改變原處分 Or pCP10 = "1007") Then
         updDTXT = TransDate(CompWorkDay(7, pCP05, 0), 1)
      End If
   'Y51753+X45149010 承辦天數:14 起算日期:系統日
   ElseIf Left(pFA01 & "000", 8) = "Y5175300" And Left(PCU01 & "000", 8) = "X4514901" Then
      If (pCP10 = "1202" Or pCP10 = "1227") Then
         updDTXT = TransDate(CompDate(2, 14, strSrvDate(1)), 1)
      End If
      
   'Added by Morgan 2018/12/19
   'Y54491 承辦天數:14 起算日期:官方發文日--劉興杰
   ElseIf Left(pFA01 & "000", 8) = "Y5449100" Then
      If (pCP10 = "1202" Or pCP10 = "1227") Then
         updDTXT = TransDate(CompDate(2, 14, IIf(pCP133 <> "", pCP133, pCP05)), 1)
      End If
   
  'Y20656 (Lerner)+X7072201(Tessera, Inc.) & X70286(Invensas Corporation)之C類接洽單之承辦期限預設為來函收文日當天
   'Modified by Lydia 2017/05/19 指定代理人
   'ElseIf InStr(Left(pFA01 & "000", 8), "Y20656") > 0 And (InStr(Left(PCU01 & "000", 8), "X7072201") > 0 Or InStr(Left(PCU01 & "000", 8), "X70286") > 0) Then
   ElseIf InStr(Left(pFA01 & "000", 8), "Y2065600") > 0 And (InStr(Left(PCU01 & "000", 8), "X7072201") > 0 Or InStr(Left(PCU01 & "000", 8), "X70286") > 0) Then
         updDTXT = TransDate(pCP05, 1)
   
   'Added by Morgan 2016/9/21 --陳亭妙,吳彩菱
   'OA承辦期限設系統日+7天
   ElseIf InStr(Left(pFA01 & "000", 8), "Y20272") > 0 Then
      'modify by sonia 2025/4/22 +1009部分准駁
      If (pCP10 = "1202" Or pCP10 = "1227" Or pCP10 = 核駁 Or pCP10 = "1009" Or pCP10 = 改變原處分 Or pCP10 = "1007") Then
         updDTXT = TransDate(CompDate(2, 7, strSrvDate(1)), 1)
      End If
   'end 2016/9/21
   'Added by Lydia 2017/03/08 Y53942 Tessera之C類接洽單之承辦期限預設為來函收文日當天
   'Modified by Lydia 2017/05/19 指定代理人
   'ElseIf InStr(Left(pFA01 & "000", 8), "Y53942") > 0 Then
   'Modify By Sindy 2017/8/3 + Y339400(Foley&Lardner,LLP)+申請人X48991INTERSIL AMERICAS INC.
   '                         有期限之來函
   'Modify By Sindy 2018/11/16 + X79754 Xcelsis Corporation
   '                         有期限之來函
   ElseIf InStr("Y5394200", Left(pFA01 & "000", 8)) > 0 Or _
      (InStr(Left(pFA01 & "000", 8), "Y3394000") > 0 And InStr(Left(PCU01 & "000", 8), "X4899100") > 0 And Val(pCP06) > 0) Or _
      (InStr(Left(PCU01 & "000", 8), "X7975400") > 0 And Val(pCP06) > 0) Then
      updDTXT = TransDate(pCP05, 1)
   'end 2017/03/08
   'Add By Sindy 2017/5/9 --陳亭妙
   'OA承辦期限設來函收文日加2天
   ElseIf InStr(Left(pFA01 & "000", 8), "Y5285900") > 0 Or InStr(Left(pFA01 & "000", 8), "Y5179901") > 0 Then
      If Val(pCP06) > 0 Then '有本所期限者
         updDTXT = TransDate(CompDate(2, 2, strSrvDate(1)), 1)
      End If
   '2017/5/9 END
   'Added by Morgan 2018/5/22
   'Y48292030(HP) C類報告(1002,1202,1227)承辦期限=收文日起18個工作天,1002僅限101,102,103,125,107,301,302,303,307,308,309--潘子微
   ElseIf InStr(Left(pFA01 & "000", 8), "Y4829203") > 0 Then
      'modify by sonia 2025/4/22 +1009部分准駁
      If (pCP10 = "1202" Or pCP10 = "1227") Or ((pCP10 = 核駁 Or pCP10 = "1009") And InStr("101,102,103,125,107,301,302,303,307,308,309", pRefCP10) > 0) Then
         'Modified by Morgan 2018/10/17 應為工作天
         'updDTXT = TransDate(CompDate(2, 18, pCP05), 1)
         updDTXT = TransDate(CompWorkDay(18, pCP05), 1)
      End If
      
   'Added by Morgan 2019/5/24--葉敏莉
   'Y33844，案件性質：審查意見通知函、最後通知函、核駁函之C類來函，收文日+承辦天數14天(日曆天)=承辦期限
   ElseIf Left(pFA01 & "000", 8) = "Y3384400" Then
      'modify by sonia 2025/4/22 +1009部分准駁
      If (pCP10 = "1202" Or pCP10 = "1227" Or pCP10 = 核駁 Or pCP10 = "1009" Or pCP10 = 改變原處分 Or pCP10 = "1007") Then
         updDTXT = TransDate(CompDate(2, 14, pCP05), 1)
      End If
   
   'Added by Morgan 2020/11/13 --郭怡瑩
   'Y20065+X17707 OA承辦期限=官方發文日+21天
   'Modified by Morgan 2020/11/23 改承辦期限=官方發文日+20天(3週內報告)
   'Modified by Morgan 2021/3/23 改回承辦期限=官方發文日+21天(3週內報告)但遇假日須前推至工作天
   'Modified by Morgan 2022/7/12 1.除審查意見通知函(1202)及最後通知(1227)外，其他改14天內報告 2.Y2006500+X8598700 比照 --林芳如
   'ElseIf Left(pFA01 & "000", 8) = "Y2006500" And Left(PCU01 & "000", 8) = "X1770700" And updDTXT <> "" Then
   ElseIf Left(pFA01 & "000", 8) = "Y2006500" And (Left(PCU01 & "000", 8) = "X1770700" Or Left(PCU01 & "000", 8) = "X8598700") And updDTXT <> "" Then
   'end 2022/7/12
      'modify by sonia 2025/4/22 +1009部分准駁
      If (pCP10 = "1202" Or pCP10 = "1227" Or pCP10 = 核駁 Or pCP10 = "1009" Or pCP10 = 改變原處分 Or pCP10 = "1007") Then
         'stDate = TransDate(CompDate(2, 20, IIf(pCP133 <> "", pCP133, pCP05)), 1)
         'Modified by Morgan 2022/7/12
         'stDate = TransDate(PUB_GetWorkDay1(CompDate(2, 21, IIf(pCP133 <> "", pCP133, pCP05)), True), 1)
         If (pCP10 = "1202" Or pCP10 = "1227") Then
            iDays = 21
         Else
            iDays = 14
         End If
         stDate = TransDate(PUB_GetWorkDay1(CompDate(2, iDays, IIf(pCP133 <> "", pCP133, pCP05)), True), 1)
         'end 2022/7/12
         
         If stDate < updDTXT Then updDTXT = stDate
      End If
   
   End If
End Sub


'Added by Lydia 2016/07/18 國外部行事曆解除期限時,通知相關人員
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Sub PUB_CancelFCPStaffCalendar(ByVal kUser As String, ByVal iUser As String, ByVal stSub As String, ByVal stConT As String, Optional ByVal iCP01 As String, Optional ByVal iCP02 As String, Optional ByVal iCP03 As String, Optional ByVal iCP04 As String)
'kUser: 解除人員
'iUser: 輸入人員
'stTo:  收件人
Dim stTO As String
'Added by Lydia 2020/09/10
Dim strQ1 As String, intQ As Integer
Dim rsQuery As New ADODB.Recordset
'end 2020/09/10
    
    If InStr(stConT, "追蹤會稿結果") > 0 And iCP01 <> "" And iCP02 <> "" And iCP03 <> "" And iCP04 <> "" Then
       '解除追蹤會稿結果時,通知FCP管制人
       stTO = PUB_GetFCPHandler(iCP01, iCP02, iCP03, iCP04)
          
       If stTO = kUser Then stTO = ""
       
    ElseIf kUser <> iUser Then
       '預設:解除人非輸入人員時,mail通知輸入人員
       'Modified by Lydia 2021/05/20 在行事曆事由增加[解除管制不通知]，排除" 解除人員非建立行事曆人員會發email通知建立人員"行事曆已被解除管制"
       'If Left(GetST15(iUser), 1) = "F" Then  'Added by Lydia 2018/09/10 排除非國外部(若是急件翻譯,櫃台收文自動設行事曆,取消行事曆時,不必通知櫃台人員 by Sharon)
       If Left(GetST15(iUser), 1) = "F" And InStr(stConT, "[解除管制不通知]") = 0 Then
           stTO = iUser
           'Added by Lydia 2024/01/05
           If strSrvDate(1) >= 新部門啟用日 Then
              strQ1 = "select st01,st03,st52,st53,st54,st55,nvl(a0924,a0908) as a0908 from staff,acc090,acc090new where st01='" & stTO & "' and st04 <> '1' and st03=a0901(+) and st93=a0921(+) "
           Else
           'end 2024/04/05
              'Added by Lydia 2020/09/10 先抓建檔人員的二級主管，若二級主管已離職則抓部門主管。
              strQ1 = "select st01,st03,st52,st53,st54,st55,a0908 from staff,acc090 where st01='" & stTO & "' and st04 <> '1' and st03=a0901(+) "
           End If
           intQ = 1
           Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
           If intQ = 1 Then
               strQ1 = ""
               If "" & rsQuery.Fields("st52") <> "" Then
                   strQ1 = GetStaffName(rsQuery.Fields("st52"))
                   If strQ1 <> "" Then stTO = "" & rsQuery.Fields("st52")
               End If
               If strQ1 = "" Then
                   stTO = "" & rsQuery.Fields("a0908")
               End If
           End If
           'end 2020/09/10
       End If
    End If
    
    If stTO <> "" Then
       PUB_SendMail kUser, stTO, "", stSub, stConT
    End If
    
    Set rsQuery = Nothing 'Added by Lydia 2020/09/10
End Sub


'Add by Morgan 2009/11/11
'收達期限控管
'Modify by Morgan 2010/12/27 +控制不可大於法限
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Sub PUB_SetArriveDate(ByVal strCP09 As String, Optional strCF03 As String)
Dim stSQL As String, iR As Integer
Dim adoRst As ADODB.Recordset, np08 As String
Dim strCP10 As String 'Added by Morgan 2015/8/7
   
   '所限=法限,FMP案抓 FCP設定
   'Modified by Morgan 2014/3/4 FMP改若FCP未設定時仍抓P的設定,否則新增的性質可能會漏掉
   'Modified by Morgan 2015/8/7
   '所有A類及B類的修正(204),主動修正(203),RCE(424)及答辯(107)皆需管控提申(15天)及收達(7天)
   'stSQL = "select cp01,cp02,cp03,cp04,cp07,cp27" & _
      ",to_char(to_date(cp27,'yyyymmdd')+nvl(b.cf23,a.cf23),'yyyymmdd') np08" & _
      " from (select * from caseprogress,patent" & _
      " where cp09='" & strCP09 & "' and cp27>19221111" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 ) x,casefee a,casefee b" & _
      " where a.cf01(+)=cp01 and a.cf02(+)=pa09 and a.cf03(+)=" & IIf(strCF03 <> "", "'" & strCF03 & "'", "cp10") & _
      " and b.cf01(+)=decode(substr(cp12,1,1),'F','FCP',cp01) and b.cf02(+)=pa09 and b.cf03(+)=" & IIf(strCF03 <> "", "'" & strCF03 & "'", "cp10") & " and nvl(b.cf23,a.cf23)>0"
   stSQL = "select cp01,cp02,cp03,cp04,cp07,cp27,cp10" & _
      ",decode(sign(nvl(b.cf23,a.cf23)),1,to_char(to_date(cp27,'yyyymmdd')+nvl(b.cf23,a.cf23),'yyyymmdd')) np08" & _
      " from (select * from caseprogress,patent" & _
      " where cp09='" & strCP09 & "' and cp27>19221111" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 ) x,casefee a,casefee b" & _
      " where a.cf01(+)=cp01 and a.cf02(+)=pa09 and a.cf03(+)=" & IIf(strCF03 <> "", "'" & strCF03 & "'", "cp10") & _
      " and b.cf01(+)=decode(substr(cp12,1,1),'F','FCP',cp01) and b.cf02(+)=pa09 and b.cf03(+)=" & IIf(strCF03 <> "", "'" & strCF03 & "'", "cp10")
   'end 2015/8/7
   'end 2014/3/4
   iR = 1
   Set adoRst = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      With adoRst
      'Modified by Morgan 2015/8/7
      'If .Fields("cp07") > 0 And Val(.Fields("np08")) > Val("" & .Fields("cp07")) Then
      '   np08 = .Fields("cp07")
      'Else
      '   np08 = .Fields("np08")
      'End If
      np08 = "" & .Fields("np08")
      If Val(np08) = 0 And .Fields("cp01") = "CFP" Then
         strCP10 = "" & .Fields("cp10")
         If strCF03 <> "" Then strCP10 = strCF03
         'Modified by Morgan 2015/8/17 +A類排除 後金909, 補收款911 -- 慧汶
         'Modified by Morgan 2015/8/27 +A類排除 超頁費938, 超項費939, 文件公簽證914 -- 慧汶
         'Modified by Morgan 2015/10/26 +A類排除 急件費920, B類新增 選取208, 訴願501 -- 慧汶
         'Modified by Morgan 2017/11/23 +排除 分析941 --甄妮, A類排除 超圖費947 --慧汶(104/12/3)
         'Modified by Morgan 2017/11/27 +排除 異同分析906 --禧佩
         'Modified by Morgan 2018/8/3 B類 +變更401 --禧佩 ex:CFP-30016
         'Modified by Morgan 2020/8/17 B類 +指定國註冊費224,年費605 --玫音請作單
         'Modified by Morgan 2023/3/7 B類 +UP註冊249
         If (Left(strCP09, 1) = "A" And InStr("906,909,911,938,939,914,920,941,947", strCP10) = 0) Or (Left(strCP09, 1) = "B" And InStr("204,203,424,107,208,401,501,941,224,605,249", strCP10) > 0) Then
            np08 = CompDate(2, 7, .Fields("cp27"))
         End If
         'end 2015/8/17
      End If
      If Val(np08) > 0 Then
         If .Fields("cp07") > 0 And Val(np08) > Val("" & .Fields("cp07")) Then
            np08 = .Fields("cp07")
         End If
      'end 2015/8/7
         
         np08 = PUB_GetWorkDay1(np08, True) 'Added by Lydia 2020/07/07 本所期限檢查：若本所期限非工作天則直接調整至最近的工作天
         stSQL = "update nextprogress set np08=" & np08 & ",np09=" & np08 & " where np01='" & strCP09 & "' and np06 is null and np07='" & 收達 & "'"
         cnnConnection.Execute stSQL, iR
         If iR = 0 Then
            stSQL = "INSERT INTO NEXTPROGRESS (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22)" & _
               " select '" & strCP09 & "','" & .Fields("cp01") & "','" & .Fields("cp02") & "'" & _
               ",'" & .Fields("cp03") & "','" & .Fields("cp04") & "'," & 收達 & _
               "," & np08 & "," & np08 & ",'" & strUserNum & "',NP22" & _
               " from (SELECT NVL(MAX(NP22),0)+1 NP22 FROM NEXTPROGRESS) X"
            cnnConnection.Execute stSQL, iR
         End If
      End If 'Added by Morgan 2015/8/7
      End With
      
   End If
   Set adoRst = Nothing
End Sub

'Added by Morgan 2015/8/7
'新增提申管制
'Modified by Morgan 2020/8/18 +strXAppDate:特殊一般提申期限
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Sub PUB_SetApplyDate(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pCP07 As String, ByVal pCP09 As String, ByVal pCP10 As String, ByVal pCP27 As String, ByVal pPA09 As String, Optional ByVal strXAppDate As String)
   Dim strTemp As String, stSQL As String, intR As Integer
   Dim strDate(3) As String
   
   If strXAppDate = "" Then 'Added by Morgan 2020/8/18
      strTemp = ""
      ClsPDGetCaseDelayDay pCP01, pPA09, pCP10, , , strTemp
   
      'CFP 所有A類及B類的修正(204),主動修正(203),RCE(424)及答辯(107)皆需管控提申(15天)及收達(7天)
      If strTemp = "" And pCP01 = "CFP" Then
         'Modified by Morgan 2015/8/17 +A類排除 後金909, 補收款911 -- 慧汶
         'Modified by Morgan 2015/8/27 +A類排除 超頁費938, 超項費939, 文件公簽證914 -- 慧汶
         'Modified by Morgan 2015/10/26 +A類排除 急件費920, B類新增 選取208, 訴願501 -- 慧汶
         'Modified by Morgan 2017/11/23 +排除 分析941 --甄妮, A類排除 超圖費947 --慧汶(104/12/3)
         'Modified by Morgan 2017/11/27 +排除 異同分析906 --禧佩
         'Modified by Morgan 2018/8/3 B類 +變更401 --禧佩 ex:CFP-30016
         'Modified by Morgan 2020/8/19 B類 +年費605 --玫音請作單
         'Modified by Morgan 2023/8/9 B類 +249UP註冊 --玫音請作單
         'Modified by Morgan 2025/8/4 B類 +701讓渡 --玫音請作單
         If (Left(pCP09, 1) = "A" And InStr("906,909,911,938,939,914,920,941,947", pCP10) = 0) Or (Left(pCP09, 1) = "B" And InStr("204,203,424,107,208,401,501,941,605,249,701", pCP10) > 0) Then
            strTemp = 15
         End If
         'end 2015/8/17
      End If
   
   'Modified by Morgan 2020/8/18
   'If strTemp <> "" Then
   End If
   If strTemp <> "" Or strXAppDate <> "" Then
   'end 2020/8/18
   
      strDate(1) = "": strDate(2) = "": strDate(3) = ""
      '最終提申
      If pCP07 <> "" Then
         'Modified by Morgan 2022/7/21
         '211西班牙(含EPC子案)及117巴西年費繳費期間規定特殊，是於法限起算3個月內繳納，法限前無法繳納
         '提申：法限為「年費法定期限」起算2週，所限提前1天。
         '最終提申：法限為「年費法定期限」起算3個月，所限提前1天。
         If pCP10 = "605" And (pPA09 = "211" Or pPA09 = "117") Then
            strDate(1) = CompDate(1, 3, pCP07)
            If strXAppDate = "" Then strXAppDate = CompDate(2, 14, pCP07)
         Else
            strDate(1) = DBDATE(pCP07)
         End If
         'end 2022/7/21
         strDate(3) = strDate(1)
         strDate(2) = PUB_GetWorkDay1(strDate(1), True)
         '若本所期限非工作天則抓最近的工作天
         stSQL = " insert into nextprogress a (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
            " values('" & pCP09 & "','" & pCP01 & "','" & pCP02 & "','" & pCP03 & "','" & pCP04 & "','996'" & _
            "," & strDate(2) & "," & strDate(1) & ",'" & strUserNum & "',GETNP22)"
         cnnConnection.Execute stSQL, intR
      End If

      '一般提申
      'Added by Morgan 2020/8/18
      If strXAppDate <> "" Then
         strDate(1) = strXAppDate
      Else
      'end 2020/8/18
      
         strDate(1) = CompDate(2, Val(strTemp), pCP27)
         
      End If 'Added by Morgan 2020/8/18
      
      '沒有最終提申或一般提申早於最終提申時才新增
      If strDate(3) = "" Or Val(strDate(1)) < Val(strDate(3)) Then
         strDate(2) = PUB_GetWorkDay1(strDate(1), True)
         stSQL = " insert into nextprogress a (np01,np02,np03,np04,np05,np07,np08,np09,np10,np22)" & _
            " values('" & pCP09 & "','" & pCP01 & "','" & pCP02 & "','" & pCP03 & "','" & pCP04 & "','998'" & _
            "," & strDate(2) & "," & strDate(1) & ",'" & strUserNum & "',GETNP22)"
         cnnConnection.Execute stSQL, intR
      End If
   End If
End Sub

'Add by Morgan 2006/5/23
'取得新的案件備註
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function PUB_GetNewCaseMemo(p_Old As String, p_AddDate As String, Optional p_AddNo As String) As String
   Dim strTmp As String, strReturn As String
   strReturn = p_Old
   strTmp = PUB_GetPCTPriDate(p_Old)
   If strTmp <> "" Then
      strReturn = Replace(strReturn, "PCT優先權日" & strTmp & ";", "")
   End If
   
   If p_AddDate <> "" Then
      strReturn = "PCT優先權日" & p_AddDate & ";" & strReturn
   End If
   
   strTmp = PUB_GetPCTPriNo(p_Old)
   If strTmp <> "" Then
      strReturn = Replace(strReturn, "PCT申請號" & strTmp & ";", "")
   End If
   If p_AddNo <> "" Then
      strReturn = "PCT申請號" & p_AddNo & ";" & strReturn
   End If
   PUB_GetNewCaseMemo = strReturn
End Function


'Add by Morgan 2010/1/20
'檢查是否尚有預繳年費未退
'Modified by Morgan 2021/4/19 T99、T100已無案件，改為檢查T109(預繳可減免退費),2021/4/23 +pChkLetter
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function PUB_ChkRefund(ByRef CaseNo() As String, Optional ByRef lngRefundFee As Long, Optional ByRef stFromYear As String, Optional ByRef stToYear As String, Optional bolIsAppForm As Boolean, Optional strAppDate As String, Optional pChkLetter As Boolean = True) As Boolean
   Dim stCon As String
   Dim stSQL As String, intR As Integer, adoRst As ADODB.Recordset
   
   If strAppDate = "" Then strAppDate = strSrvDate(1)
   '申請書
   If bolIsAppForm Then
      stCon = stCon & " and (nvl(T14,0)=0 or T14=" & strAppDate & ")"
   Else
      stCon = stCon & " and nvl(T14,0)=0"
   End If
   
   'Removed by Morgan 2021/4/19
   'stSQL = "select T09,T10,T13 from T99 where T01='" & CaseNo(1) & "' and T02='" & CaseNo(2) & "'" & _
      " and T03='" & CaseNo(3) & "' and T04='" & CaseNo(4) & "'" & stCon
      
   'Add by Morgan +1000701修法
   'Removed by Morgan 2021/4/19
   'stSQL = stSQL & " union all select T09,T10,T13 from T100 where T01='" & CaseNo(1) & "' and T02='" & CaseNo(2) & "'" & _
      " and T03='" & CaseNo(3) & "' and T04='" & CaseNo(4) & "'" & stCon
   'end 2011/6/17
      
   'Added by Morgan 2021/4/19
   stSQL = " select T09,T10,T13 from T109 where T01='" & CaseNo(1) & "' and T02='" & CaseNo(2) & "'" & _
      " and T03='" & CaseNo(3) & "' and T04='" & CaseNo(4) & "'" & IIf(pChkLetter, " and t18 is not null", "") & stCon
   'end 2021/4/19
   
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      If Not IsNull(adoRst("T09")) Then
         stFromYear = "" & adoRst("T09")
         stToYear = "" & adoRst("T10")
         lngRefundFee = Val("" & adoRst("T13"))
         PUB_ChkRefund = True
      End If
   End If
   Set adoRst = Nothing
End Function

'Add by Morgan 2010/1/20
'計算台灣應繳年費
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function PUB_GetYearFee(ByVal strPA08 As String, ByVal FromYear As Integer, ByVal ToYear As Integer, Optional ByVal bolDiscount As Boolean = False) As Long
   Dim stSQL As String, intR As Integer, adoRst As ADODB.Recordset
   Dim lngFee As Long, iYear As Integer
   
   stSQL = "Select YF05,YF07 From PatentYearFee Where YF01='000' AND YF02='" & strPA08 & "' AND YF03='Y00000001'" & _
         " AND YF04='605' AND YF05>=" & FromYear & " AND YF05<=" & ToYear & " Order By YF05 "
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      With adoRst
      Do While Not .EOF
        lngFee = lngFee + Val(.Fields("YF07").Value)
        '年費減免
        If bolDiscount Then
            iYear = Val(.Fields("YF05").Value)
            If iYear >= 1 And iYear <= 3 Then
               lngFee = lngFee - 800
            ElseIf iYear >= 4 And iYear <= 6 Then
               lngFee = lngFee - 1200
            End If
        End If
         .MoveNext
      Loop
      End With
   End If
   PUB_GetYearFee = lngFee
   Set adoRst = Nothing
End Function

'是否有保密審查未准
'Modify By Sindy 2013/9/27 +strCaseNo
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function PUB_Exists430NotPassed(PField() As String, Optional pbolNoMsg As Boolean = False, _
                                       Optional ByRef strCaseNo As String = "") As Boolean
Dim stVTable As String
Dim stSQL As String, intR As Integer
Dim adoRst As ADODB.Recordset
   
   'ADD BY SONIA 2014/6/19 只限新申請案檢查 P-106726申請優先權證明書405
   If InStr(NewCasePtyList, PField(10)) = 0 Then
      Exit Function
   End If
   'END 2014/6/19
   
   'ADD BY Sindy 2021/10/26 工程師告知目前P128420、P128421無法判發，
   '大陸發明說大陸實用新型的保密審查尚未核准, 不能判發
   '大陸實用新型說大陸發明的保密審查尚未核准, 不能判發
   '是因此兩件大陸案均與P128422建關聯，故會卡住。
   '修正程式 (加條件):若該P大陸案本身有收文保密審查, 則不控管
   stSQL = "select cp09 from caseprogress" & _
           " where cp01='" & PField(1) & "' and cp02='" & PField(2) & "' and cp03='" & PField(3) & "' and cp04='" & PField(4) & "'" & _
           " and cp10='430' and cp57 is null"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      Exit Function
   End If
   '2021/10/26 END
   
   'Modify By Sindy 2016/3/21 + False 不含一案兩請案件 ex:P-114048
   stVTable = PUB_GetRefCaseMapSQL(PField, False)

   stSQL = "select cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)" & _
      " from (" & stVTable & ") X,caseprogress where C01||C02||C03||C04<>'" & PField(1) & PField(2) & PField(3) & PField(4) & "'" & _
      " and cp01(+)=C01 and cp02(+)=C02 and cp03(+)=C03 and cp04(+)=C04 and cp10='430' and cp57 is null and (cp24 is null or cp24<>'1')"
   intR = 1
   Set adoRst = ClsLawReadRstMsg(intR, stSQL)
   If intR = 1 Then
      strCaseNo = adoRst(0) 'Add By Sindy 2013/9/27 回傳案號
      If pbolNoMsg = False Then
         MsgBox "關聯案 " & adoRst(0) & " 保密審查尚未核准，本案不可發文！"
      End If
      PUB_Exists430NotPassed = True
   End If
   Set adoRst = Nothing
End Function

'Add by Morgan 2010/6/2
'保密審查核准可發文通知
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Sub PUB_430OkInform(PField() As String)
   Dim stVTable As String
   Dim stSQL As String, intR As Integer
   Dim adoRst As ADODB.Recordset
   Dim stMsg As String
   
      stVTable = PUB_GetRefCaseMapSQL(PField)
      stSQL = "SELECT CP01||'-'||CP02||DECODE(CP03||CP04,'000','','-'||CP03||'-'||CP04) CNO,CP14" & _
         " FROM (" & stVTable & ") X,CASEPROGRESS WHERE CP01(+)=C01 AND CP02(+)=C02 AND CP03(+)=C03" & _
         " AND CP04(+)=C04 AND CP10 IN (" & NewCasePtyList & ") AND CP57||CP27 IS NULL"
      intR = 1
      Set adoRst = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         With adoRst
         Do While Not .EOF
            stMsg = PField(1) & "-" & PField(2) & IIf(PField(3) & PField(4) = "000", "", "-" & PField(3) & "-" & PField(4))
            stMsg = .Fields("CNO") & " 之相關大陸案 " & stMsg & " 保密審查已核准，可發文！"
            stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
               " VALUES ( '" & strUserNum & "','" & .Fields("CP14") & "',to_char(sysdate,'yyyymmdd')" & _
               ",to_char(sysdate,'hh24miss'),'" & stMsg & "','如旨')"
            cnnConnection.Execute stSQL, intR
            .MoveNext
         Loop
         End With
      End If
      
   Set adoRst = Nothing
End Sub

'Add by Morgan 2010/6/18
'請款單點數重新分配通知
'若已開請款單則換承辦人或核稿人時發Mail通知修改人
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Sub PUB_PointReAssignInform(CaseNo As String, BillNo As String, Optional oldCP14 As String, Optional newCP14 As String, Optional oldEP04 As String, Optional newEP04 As String)
   Dim stDept As String
   Dim stMsg As String, bolMail As Boolean
   Dim stSQL As String, intR As Integer
   Dim stCopy As String
   Dim rsAD As New ADODB.Recordset 'Added by Lydia 2024/05/16
   
   If newCP14 <> oldCP14 Then
      stMsg = "承辦人"
      If oldCP14 <> "" Then
         stDept = GetStaffDepartment(oldCP14)
         If stDept = "F21" Then
            bolMail = True
         End If
      End If
      If Not bolMail And newCP14 <> "" Then
         stDept = GetStaffDepartment(newCP14)
         If stDept = "F21" Then
            bolMail = True
         End If
      End If
   End If
   If Not bolMail And newEP04 <> oldEP04 Then
      stMsg = "核稿人"
      If oldEP04 <> "" Then
         stDept = GetStaffDepartment(oldEP04)
         If stDept = "F21" Then
            bolMail = True
         End If
      End If
      If Not bolMail And newEP04 <> "" Then
         stDept = GetStaffDepartment(newEP04)
         If stDept = "F21" Then
            bolMail = True
         End If
      End If
   End If
   If bolMail Then
      stMsg = stMsg & "已變更," & CaseNo & " 案之 " & BillNo & " 請款單請重新分配點數！"
      'Modified by Morgan 2015/5/15 改只寄 73023 --靜芳
      'stCopy = ""
      'If strUserNum <> "73023" Then stCopy = "73023"  'Add by Morgan 2010/10/23
      'stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,MC09)" & _
         " VALUES ( '" & strUserNum & "','" & strUserNum & "',to_char(sysdate,'yyyymmdd')" & _
         ",to_char(sysdate,'hh24miss'),'" & stMsg & "','如旨','" & stCopy & "')"
      'Modified by Lydia 2016/03/18 收件人改為FCP管制人
      'stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
         " VALUES ( '" & strUserNum & "','73023',to_char(sysdate,'yyyymmdd')" & _
         ",to_char(sysdate,'hh24miss'),'" & stMsg & "','如旨')"
      'end 2015/5/15
      'cnnConnection.Execute stSQL, intR
      'Modified by Lydia 2016/06/15 +a1k30
      stSQL = "select a1k13,a1k14,a1k15,a1k16,nvl(a1k30,0) a1k30 from acc1k0 where a1k01=" & CNULL(BillNo)
      intR = 1
      Set rsAD = ClsLawReadRstMsg(intR, stSQL)
      If intR = 1 Then
         If "" & rsAD(0) <> "" And "" & rsAD(1) <> "" Then
            'Added by Lydia 2016/06/15 若該筆請款單已收款(即nvl(a1k30,0)>0)則改發83002，副本才給FCP管制人；
            If rsAD.Fields("a1k30") > 0 Then
                stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                        " VALUES ( '" & strUserNum & "','" & Pub_GetSpecMan("程式管理人員") & "',to_char(sysdate,'yyyymmdd')" & _
                        ",to_char(sysdate,'hh24miss'),'" & stMsg & "','如旨','" & PUB_GetFCPHandler(rsAD(0), rsAD(1), rsAD(2), rsAD(3)) & "')"
            Else
                stSQL = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08)" & _
                        " VALUES ( '" & strUserNum & "','" & PUB_GetFCPHandler(rsAD(0), rsAD(1), rsAD(2), rsAD(3)) & "',to_char(sysdate,'yyyymmdd')" & _
                        ",to_char(sysdate,'hh24miss'),'" & stMsg & "','如旨')"
            End If
            'end 2016/06/15
            cnnConnection.Execute stSQL, intR
         End If
      End If
      'end 2016/03/18
   End If
   Set rsAD = Nothing 'Added by Lydia 2024/05/16
   
End Sub


'Add by Morgan 2010/11/1
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function PUB_ChkRefCasePA158(p_PA01 As String, p_PA02 As String, p_PA03 As String, p_PA04 As String, p_PA158 As String, Optional p_Silent As Boolean = False) As Boolean
   Dim stSQL As String
   Dim iR As Integer
   Dim adoRst As ADODB.Recordset
   Dim stMsg As String
   
   stSQL = "select pa01||'-'||pa02||decode(pa03||pa04,'000','','-'||pa03||'-'||pa04) CN,pa158" & _
      " from ( select cm01,cm02,cm03,cm04 from casemap" & _
      " where cm05='" & p_PA01 & "' and cm06='" & p_PA02 & "' and cm07='" & p_PA03 & "' and cm08='" & p_PA04 & "'" & _
      " union select cm05,cm06,cm07,cm08 from casemap" & _
      " where cm01='" & p_PA01 & "' and cm02='" & p_PA02 & "' and cm03='" & p_PA03 & "' and cm04='" & p_PA04 & "'" & _
      " union select cr05,cr06,cr07,cr08 from caserelation" & _
      " where cr01='" & p_PA01 & "' and cr02='" & p_PA02 & "' and cr03='" & p_PA03 & "' and cr04='" & p_PA04 & "'" & _
      ") X,patent where pa01(+)=cm01 and pa02(+)=cm02 and pa03(+)=cm03 and pa04(+)=cm04 and pa158<>'" & p_PA158 & "'"
   iR = 1
   Set adoRst = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      With adoRst
      Do While Not .EOF
         stMsg = stMsg & vbCrLf & .Fields(0) & " ( " & PUB_GetCaseAttributeName(.Fields("pa158")) & " )"
         .MoveNext
      Loop
      End With
      If Not p_Silent Then
         'Modified by Morgan 2013/3/29
         '改可確認後繼續
         'stMsg = "案件屬性與下列關聯案不同，請再確認！" & vbCrLf & stMsg
         'MsgBox stMsg
         stMsg = "本案案件屬性與下列關聯案不同，是否確定要繼續？" & vbCrLf & stMsg
         If MsgBox(stMsg, vbYesNo + vbDefaultButton2 + vbQuestion) = vbYes Then
            PUB_ChkRefCasePA158 = True
         End If
         'end 2013/3/29
      End If
   Else
      PUB_ChkRefCasePA158 = True
   End If
   Set adoRst = Nothing
End Function

'Add by Morgan 2011/7/14
'P案內部收文之202,203,204,205本所期限提醒
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Sub PUB_CtrlDateAlert(pCP43 As String)
   Dim stSQL As String, iR As Integer, adoRst As ADODB.Recordset
   'Modified by Lydia 2018/06/05 修改顯示案件性質 '020',CPM04,CPM03 => '000',CPM03,CPM04
   stSQL = "select sqldatet(cp06),decode(pa09,'000',cpm03,cpm04) Pty from caseprogress,patent,casepropertymap" & _
      " where cp43='" & pCP43 & "' and cp01='P' and substr(cp09,1,1)='B' and cp10 in ('202','203','204','205')" & _
      " and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04" & _
      " and cpm01(+)=cp01 and cpm02(+)=cp10 order by cp05 desc"
   iR = 1
   Set adoRst = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      MsgBox "內部收文【" & adoRst.Fields(1) & "】的本所期限為 【" & adoRst.Fields(0) & "】！"
   End If
   Set adoRst = Nothing
End Sub


'Add by Morgan 2011/7/29
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Sub PUB_944Inform(pCP09 As String)
    Dim stMC07 As String
    Dim strTo As String, intA As Integer 'Added by Lydia 2023/04/24
    
    'Modified by Morgan 2011/12/12 + B類收文控管,智權人員會收A類
    stMC07 = "cp01||'-'||cp02||decode(cp03||cp04,'000','','-'||cp03||'-'||cp04)||'申請案的會稿修改尚未處理，請儘快處理。'"
    'Added by Lydia 2023/04/24 修改王副總退休之相關控制
    If strSrvDate(1) >= "20230511" Then
        strTo = "99050"
    ElseIf strSrvDate(1) >= "20230501" Then
        strTo = "71011;99050"
    Else
        strTo = "71011"
    End If
    'end 2023/04/24
    'Modified by Lydia 2023/04/24 '71011' => " & CNULL(strTo) & "
    strSql = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
      " select '" & strUserNum & "'," & CNULL(strTo) & ",to_char(sysdate,'yyyymmdd'),to_char(sysdate,'hh24miss')" & _
      "," & stMC07 & ",'如旨',cp13 from caseprogress" & _
      " where cp43='" & pCP09 & "' and cp09>'B' and cp10='944' and cp57||cp27 is null and rownum<2"
   cnnConnection.Execute strSql, intA
End Sub

'Add by Sindy 2012/3/23
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Sub PUB_T020InsB301(m_TM01 As String, m_TM02 As String, m_TM03 As String, m_TM04 As String, m_CP09 As String, textCP44 As String, _
                           m_TM10 As String, m_CP10 As String, textCP27 As String, m_CP12 As String, m_CP13 As String, strCP45 As String)
Dim strCP44 As String
Dim strCP09 As String
Dim rsTmp As ADODB.Recordset 'Add By Sindy 2012/5/31
Dim intA As Integer, strExSql As String 'Added by Lydia 2024/05/16

   '由巨京代理之案件，請同時假收文變更代理人
   '變更案的收文與發文時間同案件
   '前次發文代理人不是Y52269
   'Modify By Sindy 2023/6/21 增加先判斷是否有前次發文 ex:TF-000832-0-00 104.領土延伸
'   strexsql = "SELECT CP44, CP09 FROM CaseProgress " & _
'                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
'                        "CP02 = '" & m_TM02 & "' AND " & _
'                        "CP03 = '" & m_TM03 & "' AND " & _
'                        "CP04 = '" & m_TM04 & "' AND " & _
'                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP27 Is Not Null"
'   inta = 1
'   Set rsTmp = ClsLawReadRstMsg(inta, strexsql)
'   If inta = 1 Then
   '2023/6/21 END
      'Modify By Sindy 2023/8/2 修改TF的案號判斷方式 ex:TF-000832-0-00 104.領土延伸
      If m_TM01 = "TF" Then
         strExSql = "SELECT CP44, Max(CP27||CP09) FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "substr(CP02,1,5) = '" & Left(m_TM02, 5) & "' AND " & _
                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP44 Is Not Null And CP27 Is Not Null Group By CP44 Order By 2 Desc, 1 "
      Else
      '2023/8/2 END
         strExSql = "SELECT CP44, Max(CP27||CP09) FROM CaseProgress " & _
                  "WHERE CP01 = '" & m_TM01 & "' AND " & _
                        "CP02 = '" & m_TM02 & "' AND " & _
                        "CP03 = '" & m_TM03 & "' AND " & _
                        "CP04 = '" & m_TM04 & "' AND " & _
                        "CP09 <> '" & m_CP09 & "' And CP09<'C' And CP44 Is Not Null And CP27 Is Not Null Group By CP44 Order By 2 Desc, 1 "
      End If
      intA = 1
      Set rsTmp = ClsLawReadRstMsg(intA, strExSql)
      If intA = 0 Then strCP44 = ""
      If intA = 1 Then strCP44 = Trim(rsTmp.Fields("CP44"))
      If Left(strCP44, 6) <> "Y52269" And Left(Trim(textCP44), 6) = "Y52269" Then
         'Modify By Sindy 2012/3/23 A類增加為全部案件性質,排除申請、異議、裁定、撤銷
         'Modify By Sindy 2012/4/26 TF也要,且不用控管申請國家條件
         'Modify By Sindy 2015/7/22 增加排除626.註銷
         'modify by sonia 2016/11/24 T-206816增加排部分撤銷(623), (不可只用strCP44 <> ""條件判斷而不剔除案件性質,否則中間接進來案件T-202368就不會產生2016/12/12)
        If (m_TM01 = "TF" Or (m_TM01 = "T" And m_TM10 = "020")) And _
            Left(m_CP09, 1) = "A" And _
            (m_CP10 <> "101" And m_CP10 <> "601" And m_CP10 <> "603" And m_CP10 <> "605" And m_CP10 <> "626" And m_CP10 <> "623") Then
            'Add By Sindy 2012/5/31 加入若變更事項已有勾選CE55代理人欄,則不要再做假收文
            strExSql = "SELECT CE55 FROM ChangeEvent " & _
                     "WHERE CE01 = '" & m_CP09 & "'"
            intA = 1
            Set rsTmp = ClsLawReadRstMsg(intA, strExSql)
            If intA = 1 Then
               If Not IsNull(rsTmp.Fields("CE55")) Then
                  If rsTmp.Fields("CE55") = "V" Then Exit Sub
               End If
            End If
            '2012/5/31 End
   
            strCP09 = AutoNo("B", 6)
            '新增一筆B類
            'Modify By Sindy 2016/6/30 m_CP12 ==> GetST15(m_CP13)
            'Modify By Sindy 2022/3/3 新增進度檔B類變更之進度備註原為「變更代理人」，請改為「變更文件接收人」。
            '   變更代理人 ==> 變更文件接收人
            strExSql = "insert into caseprogress(cp01,cp02,cp03,cp04,cp05,cp09,cp10,cp12,cp13,cp14,cp20,cp26,cp27,cp32,cp43,cp44,cp45,cp64) " & _
                           "values (" & CNULL(m_TM01) & "," & CNULL(m_TM02) & "," & CNULL(m_TM03) & _
                           "," & CNULL(m_TM04) & "," & CNULL(DBDATE(textCP27)) & "," & CNULL(strCP09) & ",301," & _
                           CNULL(GetST15(m_CP13)) & "," & CNULL(m_CP13) & "," & CNULL(strUserNum) & ",'N','N'," & CNULL(DBDATE(textCP27)) & ",'N'," & _
                           CNULL(m_CP09) & ",'" & textCP44 & "','" & strCP45 & "','變更文件接收人')"
            cnnConnection.Execute strExSql
            '新增變更事項檔
            strExSql = "insert into ChangeEvent(CE01,CE55) values('" & strCP09 & "','V')"
            cnnConnection.Execute strExSql
         End If
      End If
'   End If
End Sub


'Added by Morgan 2012/3/26
'讀取相關案清單
'Move by Lydia 2024/05/06 從basUpdate搬過來
Public Function PUB_GetRefCaseList(pSrc() As String, pList() As String) As Boolean
'pSrc=待查案號(1維),pList=相關案號(2維)
   Dim stSQL As String, intR As Integer, ii As Integer, jj As Integer, kk As Integer
   Dim bFound As Boolean
   Dim stListA() As String '目前要檢查的案號清單
   Dim stListB() As String '下回要檢查的案號清單
   Dim rsQuery As ADODB.Recordset
   
   ReDim pList(4, 0)
   ReDim stListA(4, 1)
   stListA(1, 1) = pSrc(1)
   stListA(2, 1) = pSrc(2)
   stListA(3, 1) = pSrc(3)
   stListA(4, 1) = pSrc(4)
   
   Do While UBound(stListA, 2) > 0
      For ii = 1 To UBound(stListA, 2)
         jj = UBound(pList, 2) + 1
         ReDim Preserve pList(4, jj)
         pList(1, jj) = stListA(1, ii)
         pList(2, jj) = stListA(2, ii)
         pList(3, jj) = stListA(3, ii)
         pList(4, jj) = stListA(4, ii)
      Next
      
      ReDim stListB(4, 0)
      For ii = 1 To UBound(stListA, 2)
         stSQL = "select cm01,cm02,cm03,cm04 from casemap where cm05='" & stListA(1, ii) & "' and cm06='" & stListA(2, ii) & "' and cm07='" & stListA(3, ii) & "' and cm08='" & stListA(4, ii) & "'" & _
            " union select cm05,cm06,cm07,cm08 from casemap where cm01='" & stListA(1, ii) & "' and cm02='" & stListA(2, ii) & "' and cm03='" & stListA(3, ii) & "' and cm04='" & stListA(4, ii) & "'" & _
            " union select cr01,cr02,cr03,cr04 from caserelation where cr05='" & stListA(1, ii) & "' and cr06='" & stListA(2, ii) & "' and cr07='" & stListA(3, ii) & "' and cr08='" & stListA(4, ii) & "'" & _
            " union select dc01,dc02,dc03,dc04 from divisioncase where dc05='" & stListA(1, ii) & "' and dc06='" & stListA(2, ii) & "' and dc07='" & stListA(3, ii) & "' and dc08='" & stListA(4, ii) & "'" & _
            " union select dc05,dc06,dc07,dc08 from divisioncase where dc01='" & stListA(1, ii) & "' and dc02='" & stListA(2, ii) & "' and dc03='" & stListA(3, ii) & "' and dc04='" & stListA(4, ii) & "'"
         intR = 1
         Set rsQuery = ClsLawReadRstMsg(intR, stSQL)
         If intR = 1 Then
            With rsQuery
            Do While Not .EOF
               bFound = False
               For jj = 1 To UBound(pList, 2)
                  If pList(1, jj) & pList(2, jj) & pList(3, jj) & pList(4, jj) = .Fields("cm01") & .Fields("cm02") & .Fields("cm03") & .Fields("cm04") Then
                     bFound = True
                     Exit For
                  End If
               Next
               If Not bFound Then
                  jj = UBound(stListB, 2) + 1
                  ReDim Preserve stListB(4, jj)
                  stListB(1, jj) = .Fields("cm01")
                  stListB(2, jj) = .Fields("cm02")
                  stListB(3, jj) = .Fields("cm03")
                  stListB(4, jj) = .Fields("cm04")
               End If
               .MoveNext
            Loop
            End With
         End If
      Next
      ReDim stListA(4, 0)
      For ii = 1 To UBound(stListB, 2)
         bFound = False
         For jj = 1 To UBound(stListA, 2)
            If stListA(1, jj) & stListA(2, jj) & stListA(3, jj) & stListA(4, jj) = stListB(1, ii) & stListB(2, ii) & stListB(3, ii) & stListB(4, ii) Then
               bFound = True
               Exit For
            End If
         Next
         If Not bFound Then
            jj = UBound(stListA, 2) + 1
            ReDim Preserve stListA(4, jj)
            stListA(1, jj) = stListB(1, ii)
            stListA(2, jj) = stListB(2, ii)
            stListA(3, jj) = stListB(3, ii)
            stListA(4, jj) = stListB(4, ii)
         End If
      Next
   Loop
   PUB_GetRefCaseList = True
   
'   For ii = 1 To UBound(pList, 2)
'      Debug.Print ii & "." & pList(1, ii) & pList(2, ii) & pList(3, ii) & pList(4, ii)
'   Next
   Set rsQuery = Nothing
End Function


'計算承辦人工作進度統計
'  STRINDEX    限制承辦人(空白=全部)
'  STRINDEX2   限制月份(西元)
'
'Move by Lydia 2024/05/17 從basQuery搬過來
Sub CALCUTE_090201(Strindex As String, StrIndex2 As String)
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim StrSqlB As String
Dim rsB As New ADODB.Recordset
'*************************************************************
'2003/07/04
'計算件數時不包含不計件資料, 計算點收時不論計不計件都要算
'*************************************************************
    'Modify By Cheng 2003/05/09
'    cnnConnection.Execute "DELETE FROM R090614_1 WHERE ID='" & strUserNum & "' "
    'edit by nickc 2005/05/04 加欄位
    'adoEng.Execute "DELETE FROM R090614_1 WHERE ID='" & strUserNum & "' "
    adoEng.Execute "drop table R090614_1 "
    adoEng.Execute "create table R090614_1 (R111001 text,R111002 text,R111003 double,R111004 double,R111005 double,R111006 double,R111007 double,R111008 double,R111009 double,R111010 double,R111011 double,R111012 double,R111013 double,R111014 double,R111015 double,R111016 double,R111017 double,R111018 double,R111019 double,R111020 double,R111021 double,R111022 double,R111023 double,R111024 double,ID text) "
    StrSqlB = "Select Top 1 * From R090614_1"
    rsB.CursorLocation = adUseClient
    rsB.Open StrSqlB, adoEng, adOpenDynamic, adLockOptimistic
    '統計其他項目
    '本月收文件數
'    strSQL = "INSERT INTO R090614_1 (R111001,R111002,R111003,ID) select CP14,1,count(*),'" & strUserNum & "' from caseprogress where cp01 not in ('FCP','CFP','P')   AND cp05>=" & StrIndex2 & "01 AND CP05<=" & StrIndex2 & "31 and cp57 is null and cp26 is null " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & "  GROUP BY CP14 "
'    cnnConnection.Execute strSQL
    'Modify By Cheng 2003/07/14
'    strSQLA = "Select CP14,1,Count(*),'" & strUserNum & "' from Caseprogress Where CP01 Not In ('FCP','CFP','P')   AND CP05>=" & StrIndex2 & "01 AND CP05<=" & StrIndex2 & "31 and cp57 is null and cp26 is null " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & "  GROUP BY CP14 "
    'edit by nickc 2005/05/04 加新制統計
    'strSQLA = "Select CP14,1,Count(*),'" & strUserNum & "' from Caseprogress Where CP01 Not In ('FCP','CFP','P','PS','CPS','FG')   AND CP05>=" & StrIndex2 & "01 AND CP05<=" & StrIndex2 & "31 and cp57 is null and cp26 is null " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & "  GROUP BY CP14 "
    'edit by nickc 2006/02/22
    'StrSQLa = "Select CP14,1,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(nvl(cp97,0) * nvl(cp98,0)) from Caseprogress Where CP01 Not In ('FCP','CFP','P','PS','CPS','FG')   AND CP05>=" & StrIndex2 & "01 AND CP05<=" & StrIndex2 & "31 and cp57 is null  " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & "  GROUP BY CP14 "
    StrSQLa = "Select CP14,1,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from Caseprogress Where CP01 Not In ('FCP','CFP','P','PS','CPS','FG')   AND CP05>=" & StrIndex2 & "01 AND CP05<=" & StrIndex2 & "31 and cp57 is null  " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & "  GROUP BY CP14 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            rsB.AddNew
            rsB("R111001").Value = "" & rsA.Fields(0).Value
            rsB("R111002").Value = "" & rsA.Fields(1).Value
            rsB("R111003").Value = Val("" & rsA.Fields(2).Value)
            'add by nickc 2005/05/04
            rsB("R111014").Value = Val("" & rsA.Fields(4).Value)
            rsB("ID").Value = "" & rsA.Fields(3).Value
            rsB.UPDATE
            rsA.MoveNext
        Wend
    Else
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "1"
            rsB("R111003").Value = 0
            'add by nickc 2005/05/04
            rsB("R111014").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '當月發文件數
'    strSQL = "INSERT INTO R090614_1 (R111001,R111002,R111004,ID) select CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),count(*),'" & strUserNum & "' from caseprogress where CP05>=19980101  AND cp27>=" & StrIndex2 & "01 AND CP27<=" & StrIndex2 & "31 and cp57 is  null and cp26 is null " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
'    cnnConnection.Execute strSQL
    'edit by nickc 2005/05/04
    'strSQLA = "Select CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),count(*),'" & strUserNum & "' from caseprogress where CP05>=19980101  AND cp27>=" & StrIndex2 & "01 AND CP27<=" & StrIndex2 & "31 and cp57 is  null and cp26 is null " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    'edit by nickc 2006/02/22
    'StrSQLa = "Select CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(nvl(cp97,0) * nvl(cp98,0)) from caseprogress where CP05>=19980101  AND cp27>=" & StrIndex2 & "01 AND CP27<=" & StrIndex2 & "31 and cp57 is  null " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    StrSQLa = "Select CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from caseprogress where CP05>=19980101  AND cp27>=" & StrIndex2 & "01 AND CP27<=" & StrIndex2 & "31 and cp57 is  null " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            rsB.AddNew
            rsB("R111001").Value = "" & rsA.Fields(0).Value
            rsB("R111002").Value = "" & rsA.Fields(1).Value
            rsB("R111004").Value = Val("" & rsA.Fields(2).Value)
            'add by nickc 2005/05/04
            rsB("R111015").Value = Val("" & rsA.Fields(4).Value)
            rsB("ID").Value = "" & rsA.Fields(3).Value
            rsB.UPDATE
            rsA.MoveNext
        Wend
    Else
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "1"
            rsB("R111004").Value = 0
            'add by nickc 2005/05/04
            rsB("R111015").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "2"
            rsB("R111004").Value = 0
            'add by nickc 2005/05/04
            rsB("R111015").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '當月發文點數
'    strSQL = "INSERT INTO R090614_1 (R111001,R111002,R111005,ID) select CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),SUM(CP18),'" & strUserNum & "' from caseprogress where CP05>=19980101  AND cp27>=" & StrIndex2 & "01 AND CP27<=" & StrIndex2 & "31 and cp57 is  null " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
'    cnnConnection.Execute strSQL
    'edit by nickc 2005/05/04
    'strSQLA = "Select CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),SUM(CP18),'" & strUserNum & "' from caseprogress where CP05>=19980101  AND cp27>=" & StrIndex2 & "01 AND CP27<=" & StrIndex2 & "31 and cp57 is  null " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    'edit by nickc 2005/12/14 加快
    'StrSQLa = "Select CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),SUM(CP18),'" & strUserNum & "',sum(nvl(cp18,0)-nvl(a1u07/1000,0)) from caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 group by a1u03) ABCDE where CP05>=19980101  and cp09=a1u03(+) AND cp27>=" & StrIndex2 & "01 AND CP27<=" & StrIndex2 & "31 and cp57 is  null " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    StrSQLa = "Select CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),SUM(CP18),'" & strUserNum & "',sum(nvl(cp18,0)-nvl(a1u07/1000,0)) from caseprogress,(select a1u03,sum(nvl(a1u07,0)) as a1u07 from acc1u0 where a1u03 in (select cp09 from caseprogress where CP05>=19980101 and cp27>=" & StrIndex2 & "01 AND CP27<=" & StrIndex2 & "31 and cp57 is  null " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & ")  group by a1u03) ABCDE where CP05>=19980101  and cp09=a1u03(+) AND cp27>=" & StrIndex2 & "01 AND CP27<=" & StrIndex2 & "31 and cp57 is  null " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY CP14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            rsB.AddNew
            rsB("R111001").Value = "" & rsA.Fields(0).Value
            rsB("R111002").Value = "" & rsA.Fields(1).Value
            rsB("R111005").Value = Val("" & rsA.Fields(2).Value)
            'add by nickc 2005/05/04
            rsB("R111016").Value = Val("" & rsA.Fields(4).Value)
            rsB("ID").Value = "" & rsA.Fields(3).Value
            rsB.UPDATE
            rsA.MoveNext
        Wend
    Else
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "1"
            rsB("R111005").Value = 0
            'add by nickc 2005/05/04
            rsB("R111016").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "2"
            rsB("R111005").Value = 0
            'add by nickc 2005/05/04
            rsB("R111016").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
   'Q1
   '目前未完稿的件數
'    strSQL = "INSERT INTO R090614_1 (R111001,R111002,R111006,ID) select ep05,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where ep02=cp09(+) AND  cp01 not in ('FCP','CFP','P') AND EP09 IS NULL " & IIf(Len(Strindex) = 0, "", "and ep05='" & Strindex & "' ") & "  and cp57 is  null GROUP BY ep05 "
'    cnnConnection.Execute strSQL
    'Modify By Cheng 2003/07/04
'    strSQLA = "Select ep05,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where ep02=cp09(+) AND  cp01 not in ('FCP','CFP','P') AND EP09 IS NULL " & IIf(Len(Strindex) = 0, "", "and ep05='" & Strindex & "' ") & "  and cp57 is  null GROUP BY ep05 "
    'Modify By Cheng 2003/07/14
'    strSQLA = "Select ep05,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where ep02=cp09(+) AND  cp01 not in ('FCP','CFP','P') AND EP09 IS NULL " & IIf(Len(Strindex) = 0, "", "and ep05='" & Strindex & "' ") & "  and cp57 is  null And CP26 Is Null GROUP BY ep05 "
    'edit by nickc 2005/05/04
    'strSQLA = "Select ep05,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where ep02=cp09(+) AND  cp01 not in ('FCP','CFP','P','PS','CPS','FG') AND EP09 IS NULL " & IIf(Len(Strindex) = 0, "", "and ep05='" & Strindex & "' ") & "  and cp57 is  null And CP26 Is Null GROUP BY ep05 "
    'edit by nickc 2006/02/22
    'StrSQLa = "Select ep05,1,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(nvl(cp97,0) * nvl(cp98,0)) from engineerprogress,caseprogress where ep02=cp09(+) AND  cp01 not in ('FCP','CFP','P','PS','CPS','FG') AND EP09 IS NULL " & IIf(Len(Strindex) = 0, "", "and ep05='" & Strindex & "' ") & "  and cp57 is  null  GROUP BY ep05 "
    StrSQLa = "Select ep05,1,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from engineerprogress,caseprogress where ep02=cp09(+) AND  cp01 not in ('FCP','CFP','P','PS','CPS','FG') AND EP09 IS NULL " & IIf(Len(Strindex) = 0, "", "and ep05='" & Strindex & "' ") & "  and cp57 is  null  GROUP BY ep05 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            rsB.AddNew
            rsB("R111001").Value = "" & rsA.Fields(0).Value
            rsB("R111002").Value = "" & rsA.Fields(1).Value
            rsB("R111006").Value = Val("" & rsA.Fields(2).Value)
            'add by nickc 2005/05/04
            rsB("R111017").Value = Val("" & rsA.Fields(4).Value)
            rsB("ID").Value = "" & rsA.Fields(3).Value
            rsB.UPDATE
            rsA.MoveNext
        Wend
    Else
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "1"
            rsB("R111006").Value = 0
            'add by nickc 2005/05/04
            rsB("R111017").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '會稿中的件數
'    strSQL = "INSERT INTO R090614_1 (R111001,R111002,R111007,ID) select EP05,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 not in ('FCP','CFP','P')   AND EP07 IS NOT NULL  AND EP08 IS NULL " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
'    cnnConnection.Execute strSQL
    'Modify By Cheng 2003/07/04
'    strSQLA = "Select EP05,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 not in ('FCP','CFP','P')   AND EP07 IS NOT NULL  AND EP08 IS NULL " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    'Modify By Cheng 2003/07/14
'    strSQLA = "Select EP05,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 not in ('FCP','CFP','P')   AND EP07 IS NOT NULL  AND EP08 IS NULL " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null And CP26 Is Null GROUP BY EP05 "
    'edit by nickc 2005/05/04
    'strSQLA = "Select EP05,1,count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 not in ('FCP','CFP','P','PS','CPS','FG')   AND EP07 IS NOT NULL  AND EP08 IS NULL " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null And CP26 Is Null GROUP BY EP05 "
    'edit by nickc 2006/02/22
    'StrSQLa = "Select EP05,1,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(nvl(cp97,0) * nvl(cp98,0)) from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 not in ('FCP','CFP','P','PS','CPS','FG')   AND EP07 IS NOT NULL  AND EP08 IS NULL " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    StrSQLa = "Select EP05,1,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from engineerprogress,caseprogress where EP02=CP09(+) AND  cp01 not in ('FCP','CFP','P','PS','CPS','FG')   AND EP07 IS NOT NULL  AND EP08 IS NULL " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            rsB.AddNew
            rsB("R111001").Value = "" & rsA.Fields(0).Value
            rsB("R111002").Value = "" & rsA.Fields(1).Value
            rsB("R111007").Value = Val("0" & rsA.Fields(2).Value)
            'add by nickc 2005/05/04
            rsB("R111018").Value = Val("0" & rsA.Fields(4).Value)
            rsB("ID").Value = "" & rsA.Fields(3).Value
            rsB.UPDATE
            rsA.MoveNext
        Wend
    Else
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "1"
            rsB("R111007").Value = 0
            'add by nickc 2005/05/04
            rsB("R111018").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '超過承辦期限之件數
    '2002/01/24  薛說與汪先生確認過，要變更規則，原本檢查 系統日＞承辦期限，無取消收文，無發文；現在要改成，系統日＞承辦期限，無取消收文，無會稿日，無發文
    'strSQL = "INSERT INTO R090614_1 (R111001,R111002,R111008,ID) select EP05,decode(cp01,'P',2,'CFP',2,'FCP',2,1),count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  CP05>=19980101  AND " & GetTodayDate & ">CP48 AND EP07 IS NOT NULL  AND CP48 IS NOT NULL  " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP26 IS NULL AND CP27 IS NULL GROUP BY EP05,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
'    strSQL = "INSERT INTO R090614_1 (R111001,R111002,R111008,ID) select EP05,decode(cp01,'P',2,'CFP',2,'FCP',2,1),count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  CP05>=19980101  AND " & GetTodayDate & ">CP48 AND CP48 IS NOT NULL  " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP26 IS NULL AND eP07 IS NULL and cp27 is null GROUP BY EP05,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
'    cnnConnection.Execute strSQL
    'edit by nickc 2005/05/04
    'strSQLA = "Select EP05,decode(cp01,'P',2,'CFP',2,'FCP',2,1),count(*),'" & strUserNum & "' from engineerprogress,caseprogress where EP02=CP09(+) AND  CP05>=19980101  AND " & GetTodayDate & ">CP48 AND CP48 IS NOT NULL  " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP26 IS NULL AND eP07 IS NULL and cp27 is null GROUP BY EP05,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    'edit by nickc 2006/02/22
    'StrSQLa = "Select EP05,decode(cp01,'P',2,'CFP',2,'FCP',2,1),sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(nvl(cp97,0) * nvl(cp98,0)) from engineerprogress,caseprogress where EP02=CP09(+) AND  CP05>=19980101  AND " & GetTodayDate & ">CP48 AND CP48 IS NOT NULL  " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND eP07 IS NULL and cp27 is null GROUP BY EP05,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    StrSQLa = "Select EP05,decode(cp01,'P',2,'CFP',2,'FCP',2,1),sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from engineerprogress,caseprogress where EP02=CP09(+) AND  CP05>=19980101  AND " & GetTodayDate & ">CP48 AND CP48 IS NOT NULL  " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND eP07 IS NULL and cp27 is null GROUP BY EP05,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            rsB.AddNew
            rsB("R111001").Value = "" & rsA.Fields(0).Value
            rsB("R111002").Value = "" & rsA.Fields(1).Value
            rsB("R111008").Value = Val("" & rsA.Fields(2).Value)
            'add by nickc 2005/05/04
            rsB("R111019").Value = Val("" & rsA.Fields(4).Value)
            rsB("ID").Value = "" & rsA.Fields(3).Value
            rsB.UPDATE
            rsA.MoveNext
        Wend
    Else
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "1"
            rsB("R111008").Value = 0
            'add by nickc 2005/05/04
            rsB("R111019").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "2"
            rsB("R111008").Value = 0
            'add by nickc 2005/05/04
            rsB("R111019").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    'Q2
    '當日法定期限之件數
'    strSQL = "INSERT INTO R090614_1 (R111001,R111002,R111009,ID) select cp14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),count(*),'" & strUserNum & "' from caseprogress where  cp57 is  null  AND CP27 IS NULL and CP05>=19980101 AND  CP06<=" & GetTodayDate & " " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY cp14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
'    cnnConnection.Execute strSQL
    'Modify By Cheng 2003/07/04
'    strSQLA = "Select cp14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),count(*),'" & strUserNum & "' from caseprogress where  cp57 is  null  AND CP27 IS NULL and CP05>=19980101 AND  CP06<=" & GetTodayDate & " " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY cp14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    'edit by nickc 2005/05/04
    'strSQLA = "Select cp14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),count(*),'" & strUserNum & "' from caseprogress where  cp57 is  null  AND CP27 IS NULL And CP26 Is Null and CP05>=19980101 AND  CP06<=" & GetTodayDate & " " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY cp14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    'edit by nickc 2006/02/22
    'StrSQLa = "Select cp14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(nvl(cp97,0) * nvl(cp98,0)) from caseprogress where  cp57 is  null  AND CP27 IS NULL  and CP05>=19980101 AND  CP06<=" & GetTodayDate & " " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY cp14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    StrSQLa = "Select cp14,decode(cp01,'P',2,'CFP',2,'FCP',2,1),sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from caseprogress where  cp57 is  null  AND CP27 IS NULL  and CP05>=19980101 AND  CP06<=" & GetTodayDate & " " & IIf(Len(Strindex) = 0, "", "and CP14='" & Strindex & "' ") & " GROUP BY cp14,decode(cp01,'P',2,'CFP',2,'FCP',2,1) "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            rsB.AddNew
            rsB("R111001").Value = "" & rsA.Fields(0).Value
            rsB("R111002").Value = "" & rsA.Fields(1).Value
            rsB("R111009").Value = Val("" & rsA.Fields(2).Value)
            'add by nickc 2005/05/04
            rsB("R111020").Value = Val("" & rsA.Fields(4).Value)
            rsB("ID").Value = "" & rsA.Fields(3).Value
            rsB.UPDATE
            rsA.MoveNext
        Wend
    Else
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "1"
            rsB("R111009").Value = 0
            'add by nickc 2005/05/04
            rsB("R111020").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "2"
            rsB("R111009").Value = 0
            'add by nickc 2005/05/04
            rsB("R111020").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '可辦非設計案件
'    strSQL = "INSERT INTO R090614_1 (R111001,R111002,R111010,ID) select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and ep09 is null and EP06 IS NOT NULL AND PA08<>'3' " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL GROUP BY EP05 "
'    cnnConnection.Execute strSQL
    'Modify By Cheng 2003/05/29
'    strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and ep09 is null and EP06 IS NOT NULL AND PA08<>'3' " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL GROUP BY EP05 "
    'Modify By Cheng 2003/07/14
'    strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and ep09 is null and EP06 IS NOT NULL AND (CP10<>'103' And CP10<>'105' ) " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL GROUP BY EP05 "

   'Modify by Morgan 2004/5/19
   '設計加案件性質 113
    'strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and ep09 is null and EP06 IS NOT NULL AND (CP10<>'103' And CP10<>'105' ) " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL GROUP BY EP05 "
    'edit by nickc 2005/05/04
    'strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and ep09 is null and EP06 IS NOT NULL AND (CP10<>'103' And CP10<>'105' ) AND NOT (PA08='3' AND CP10='113') " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL GROUP BY EP05 "
    'edit by nickc 2006/02/22
    'StrSQLa = "Select EP05,2,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(nvl(cp97,0) * nvl(cp98,0)) from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)  and ep09 is null and EP06 IS NOT NULL AND (CP10<>'103' And CP10<>'105' ) AND NOT (PA08='3' AND CP10='113') " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL GROUP BY EP05 "
    StrSQLa = "Select EP05,2,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)  and ep09 is null and EP06 IS NOT NULL AND (CP10<>'103' And CP10<>'105' And CP10<>'125') AND NOT (PA08='3' AND CP10='113') " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL GROUP BY EP05 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            rsB.AddNew
            rsB("R111001").Value = "" & rsA.Fields(0).Value
            rsB("R111002").Value = "" & rsA.Fields(1).Value
            rsB("R111010").Value = Val("" & rsA.Fields(2).Value)
            'add by nickc 2005/05/04
            rsB("R111021").Value = Val("" & rsA.Fields(4).Value)
            rsB("ID").Value = "" & rsA.Fields(3).Value
            rsB.UPDATE
            rsA.MoveNext
        Wend
    Else
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "2"
            rsB("R111010").Value = 0
            'add by nickc 2005/05/04
            rsB("R111021").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '可辦設計案件
'    strSQL = "INSERT INTO R090614_1 (R111001,R111002,R111011,ID) select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and ep09 is null and EP06 IS NOT NULL AND PA08='3' " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL  GROUP BY EP05 "
'    cnnConnection.Execute strSQL
    'Modify By Cheng 2003/05/29
'    strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and ep09 is null and EP06 IS NOT NULL AND PA08='3' " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL  GROUP BY EP05 "
    'Modify By Cheng 2003/07/14
'    strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and ep09 is null and EP06 IS NOT NULL AND (CP10='103' Or CP10='105') " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL  GROUP BY EP05 "

   'Modify by Morgan 2004/5/19
   '設計加案件性質 113
    'strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and ep09 is null and EP06 IS NOT NULL AND (CP10='103' Or CP10='105') " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL  GROUP BY EP05 "
    'edit by nickc 2005/05/04
    'strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and ep09 is null and EP06 IS NOT NULL AND (CP10='103' Or CP10='105' Or ( PA08='3' AND CP10='113')) " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL  GROUP BY EP05 "
    'edit by nickc 2006/02/22
    'StrSQLa = "Select EP05,2,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(nvl(cp97,0) * nvl(cp98,0)) from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and ep09 is null and EP06 IS NOT NULL AND (CP10='103' Or CP10='105' Or ( PA08='3' AND CP10='113')) " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL  GROUP BY EP05 "
    StrSQLa = "Select EP05,2,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and ep09 is null and EP06 IS NOT NULL AND (CP10='103' Or CP10='105' Or CP10='125' Or ( PA08='3' AND CP10='113')) " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null AND CP27 IS NULL  GROUP BY EP05 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            rsB.AddNew
            rsB("R111001").Value = "" & rsA.Fields(0).Value
            rsB("R111002").Value = "" & rsA.Fields(1).Value
            rsB("R111011").Value = Val("" & rsA.Fields(2).Value)
            'add by nickc 2005/05/04
            rsB("R111022").Value = Val("" & rsA.Fields(4).Value)
            rsB("ID").Value = "" & rsA.Fields(3).Value
            rsB.UPDATE
            rsA.MoveNext
        Wend
    Else
            rsB.AddNew
            rsB("R111001").Value = Strindex
            rsB("R111002").Value = "2"
            rsB("R111011").Value = 0
            'add by nickc 2005/05/04
            rsB("R111022").Value = 0
            rsB("ID").Value = strUserNum
            rsB.UPDATE
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '本月已完稿非設計件數
'    strSQL = "INSERT INTO R090614_1 (R111001,R111002,R111012,ID) select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND PA08<>'3' " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
'    cnnConnection.Execute strSQL
    'Modify By Cheng 2003/05/14
    '案件性質非'103','105'為設計
'    strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND PA08<>'3' " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    'Modify By Cheng 2003/07/14
'    strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND (CP10<>'103' And CP10<>'105') " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
   
   'Modify by Morgan 2004/5/19
   '設計加案件性質 113
    'strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND (CP10<>'103' And CP10<>'105') " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    'edit by nickc 2005/05/04
    'strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND (CP10<>'103' And CP10<>'105') AND NOT (PA08='3' AND CP10='113') " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    'edit by nickc 2006/02/22
    'StrSQLa = "Select EP05,2,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(nvl(cp97,0) * nvl(cp98,0)) from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND (CP10<>'103' And CP10<>'105') AND NOT (PA08='3' AND CP10='113') " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    StrSQLa = "Select EP05,2,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND (CP10<>'103' And CP10<>'105' And CP10<>'125') AND NOT (PA08='3' AND CP10='113') " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            rsB.AddNew
            rsB("R111001").Value = "" & rsA.Fields(0).Value
            rsB("R111002").Value = "" & rsA.Fields(1).Value
            rsB("R111012").Value = Val("" & rsA.Fields(2).Value)
            'add by nickc 2005/05/04
            rsB("R111023").Value = Val("" & rsA.Fields(4).Value)
            rsB("ID").Value = "" & rsA.Fields(3).Value
            rsB.UPDATE
            rsA.MoveNext
        Wend
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    '本月已完稿設計件數
'    strSQL = "INSERT INTO R090614_1 (R111001,R111002,R111013,ID) select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND PA08='3' " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
'    cnnConnection.Execute strSQL
    'Modify By Cheng 2003/05/14
    '案件性質'103','105'為設計
'    strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND PA08='3' " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    'Modify By Cheng 2003/07/14
'    strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND (CP10='103' Or CP10='105') " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    'edit by nickc 2005/05/04
    'strSQLA = "Select EP05,2,count(*),'" & strUserNum & "' from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+) AND cp26 is null and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND (CP10='103' Or CP10='105' Or ( PA08='3' AND CP10='113') ) " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    'edit by nickc 2006/02/22
    'StrSQLa = "Select EP05,2,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(nvl(cp97,0) * nvl(cp98,0)) from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)  and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND (CP10='103' Or CP10='105' Or ( PA08='3' AND CP10='113') ) " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    StrSQLa = "Select EP05,2,sum(decode(cp26,null,1,0)),'" & strUserNum & "',sum(decode(cp112,'Y',round(nvl(cp97,0) * nvl(cp98,0) * nvl(cp111,1),2),round(cp97 * cp98,2))) from engineerprogress,caseprogress,PATENT where EP02=CP09(+) AND  cp01 in ('FCP','CFP','P','PS','CPS','FG')  AND CP01=PA01(+) AND CP02=PA02(+) AND CP03=PA03(+) AND CP04=PA04(+)  and EP09>=" & StrIndex2 & "01 and ep09<=" & StrIndex2 & "31 AND (CP10='103' Or CP10='105' Or CP10='125' Or ( PA08='3' AND CP10='113') ) " & IIf(Len(Strindex) = 0, "", "and EP05='" & Strindex & "' ") & " and cp57 is  null GROUP BY EP05 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    If rsA.RecordCount > 0 Then
        While Not rsA.EOF
            rsB.AddNew
            rsB("R111001").Value = "" & rsA.Fields(0).Value
            rsB("R111002").Value = "" & rsA.Fields(1).Value
            rsB("R111013").Value = Val("" & rsA.Fields(2).Value)
            'add by nickc 2005/05/04
            rsB("R111024").Value = Val("" & rsA.Fields(4).Value)
            rsB("ID").Value = "" & rsA.Fields(3).Value
            rsB.UPDATE
            rsA.MoveNext
        Wend
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
    If rsB.State <> adStateClosed Then rsB.Close
    Set rsB = Nothing
End Sub

'Add by Mrogan 2004/11/26 抓分割母案的申請日
'Modify By Sindy 2017/11/29 + Optional ByRef strPA11 As String = ""
'Modify By Sindy 2020/3/10 + Optional ByRef strPA158 As String = ""
'Move by Lydia 2024/05/17 從basQuery搬過來
Public Function PUB_DivAppDate(ByVal stDC01 As String, ByVal stDC02 As String, _
      ByVal stDC03 As String, ByVal stDC04 As String, Optional ByVal bolMessage As Boolean = False, _
      Optional ByRef StrPA11 As String = "", Optional ByRef strPA158 As String = "") As String

Dim stAppDate As String

On Error GoTo ErrHnd
   StrPA11 = "" 'Add By Sindy 2017/11/29
   strPA158 = "" 'Add By Sindy 2020/3/10
   'Modify By Sindy 2017/11/29 +PA11
   'Modify By Sindy 2020/3/10 + PA158
   strSql = "SELECT PA10,PA11,PA158 FROM DIVISIONCASE, PATENT WHERE DC01='" & stDC01 & "' and DC02='" & stDC02 & "' and DC03='" & stDC03 & "' and DC04='" & stDC04 & "'" & _
      " and PA01(+)=DC05 AND PA02(+)=DC06 AND PA03(+)=DC07 AND PA04(+)=DC08"
   CheckOC3
   With AdoRecordSet3
      .CursorLocation = adUseClient
      .Open strSql, cnnConnection, adOpenStatic, adLockReadOnly
      If .RecordCount <> 0 Then
         stAppDate = "" & .Fields(0)
         If stAppDate = "" And bolMessage = True Then
            MsgBox "母案未輸入申請日！", vbExclamation
         End If
         StrPA11 = "" & .Fields("PA11") 'Add By Sindy 2017/11/29
         strPA158 = "" & .Fields("PA158") 'Add By Sindy 2017/11/29
      ElseIf bolMessage = True Then
         MsgBox "該分割案未建立母案關聯！", vbExclamation
      End If
   End With
   PUB_DivAppDate = stAppDate
   
ErrHnd:
   If Err.Number <> 0 Then MsgBox Err.Description
End Function

'add by nickc 2005/06/16 檢查有無香港關聯
'edit by nickc 2006/05/05
'Public Function ChkCMIsExist013(oCP01 As String, oCP02 As String, Optional oCP03 As String = "0", Optional oCP04 As String = "00") As Boolean
'Modified by Morgan 2014/9/23 +strPA08
'Modified by Lydia 2015/07/27 + iTyp 從國內案(大陸案CM05~CM08)檢查有無澳門關聯(CM01~CM04)
'Modified by Morgan 2016/9/7 +bolNotClosed:未閉卷
'Move by Lydia 2024/05/17 從basQuery搬過來
Public Function ChkCMIsExist013(oCP01 As String, oCP02 As String, Optional oCP03 As String = "0", Optional oCP04 As String = "00", Optional oHKCP01 As String, Optional oHKCP02 As String, Optional oHKCP03 As String, Optional oHKCP04 As String, Optional oHKCP09 As String, Optional strPA08 As String, Optional iTyp As String = "4", Optional bolNotClosed As Boolean = False) As Boolean
Dim StrSqlB As String
Dim strCon As String
Dim stPA09 As String 'Added by Lydia 2015/07/27
If strPA08 <> "" Then strCon = " and pa08='" & strPA08 & "'"

If bolNotClosed Then strCon = strCon & " and pa57 is null" 'Added by Morgan 2016/9/7

'Added by Lydia 2015/07/27
Select Case iTyp
    Case "4"
          stPA09 = "013"
    Case "5"
          stPA09 = "044"
End Select

'StrSqlB = "select * from casemap,patent where cm05='" & oCP01 & "' and cm06='" & oCP02 & "' and cm07='" & oCP03 & "' and cm08='" & oCP04 & "' and cm10='4' and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and pa09='013' "
'Modified by Lydia 2015/07/27
'StrSqlB = "select * from casemap,patent,caseprogress where cm05='" & oCP01 & "' and cm06='" & oCP02 & "' and cm07='" & oCP03 & "' and cm08='" & oCP04 & "' and cm10='4' and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and pa09='013' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) " & strCon
StrSqlB = "select * from casemap,patent,caseprogress where cm05='" & oCP01 & "' and cm06='" & oCP02 & "' and cm07='" & oCP03 & "' and cm08='" & oCP04 & "' and cm10='" & iTyp & "' and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and pa09='" & stPA09 & "' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) " & strCon
'end 2015/07/27

CheckOC3
With AdoRecordSet3
   .CursorLocation = adUseClient
   .Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      ChkCMIsExist013 = True
      'add by nickc 2006/05/05
      oHKCP01 = CheckStr(.Fields("cm01"))
      oHKCP02 = CheckStr(.Fields("cm02"))
      oHKCP03 = CheckStr(.Fields("cm03"))
      oHKCP04 = CheckStr(.Fields("cm04"))
      oHKCP09 = CheckStr(.Fields("cp09"))
   Else
      ChkCMIsExist013 = False
      'add by nickc 2006/05/05
      oHKCP01 = ""
      oHKCP02 = ""
      oHKCP03 = ""
      oHKCP04 = ""
      oHKCP09 = ""
   End If
End With
CheckOC3
End Function

'add by nickc 2005/06/16 檢查香港案有無收文 111 (有 np 和 cp 兩種)
'Move by Lydia 2024/05/17 從basQuery搬過來
Public Function Chk013Have111(oCP01 As String, oCP02 As String, oCP03 As String, oCP04 As String, oCP14 As String, Optional NpOrCp As String = "CP", Optional iTyp As String = "4") As String
Dim StrSqlB As String

If NpOrCp = "CP" Then
   StrSqlB = "select cp09,cp14 from casemap,patent,caseprogress where cm05='" & oCP01 & "' and cm06='" & oCP02 & "' and cm07='" & oCP03 & "' and cm08='" & oCP04 & "' and cm10='4' and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and pa09='013' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and cp10='111' "
Else
   StrSqlB = "select np01,'' from casemap,patent,nextprogress where cm05='" & oCP01 & "' and cm06='" & oCP02 & "' and cm07='" & oCP03 & "' and cm08='" & oCP04 & "' and cm10='4' and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and pa09='013' and cm01=np02(+) and cm02=np03(+) and cm03=np04(+) and cm04=np05(+) and np07=111 "
End If
CheckOC3
With AdoRecordSet3
   .CursorLocation = adUseClient
   .Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      Chk013Have111 = CheckStr(.Fields(0))
      oCP14 = CheckStr(.Fields(1))
   Else
      Chk013Have111 = ""
      'oCP14 = ""
   End If
End With
CheckOC3
End Function

'add by nickc 2005/11/22 檢查香港案有無收文 110 (有 np 和 cp 兩種)
'Move by Lydia 2024/05/17 從basQuery搬過來
Public Function Chk013Have110(oCP01 As String, oCP02 As String, oCP03 As String, oCP04 As String, oCP14 As String, Optional NpOrCp As String = "CP") As String
Dim StrSqlB As String

If NpOrCp = "CP" Then
   StrSqlB = "select cp09,cp14 from casemap,patent,caseprogress where cm05='" & oCP01 & "' and cm06='" & oCP02 & "' and cm07='" & oCP03 & "' and cm08='" & oCP04 & "' and cm10='4' and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and pa09='013' and cm01=cp01(+) and cm02=cp02(+) and cm03=cp03(+) and cm04=cp04(+) and cp10='110' "
Else
   StrSqlB = "select np01,'' from casemap,patent,nextprogress where cm05='" & oCP01 & "' and cm06='" & oCP02 & "' and cm07='" & oCP03 & "' and cm08='" & oCP04 & "' and cm10='4' and cm01=pa01(+) and cm02=pa02(+) and cm03=pa03(+) and cm04=pa04(+) and pa09='013' and cm01=np02(+) and cm02=np03(+) and cm03=np04(+) and cm04=np05(+) and np07=110 "
End If
CheckOC3
With AdoRecordSet3
   .CursorLocation = adUseClient
   .Open StrSqlB, cnnConnection, adOpenStatic, adLockReadOnly
   If .RecordCount <> 0 Then
      Chk013Have110 = CheckStr(.Fields(0))
      oCP14 = CheckStr(.Fields(1))
   Else
      Chk013Have110 = ""
   End If
End With
CheckOC3
End Function

'Added by Lydia 2024/07/09 外商：判斷案件國家收費表內有設定提申期限(天)CF11，要加掛提申(998)期限
Public Sub Pub_GetCF11to998(ByVal pNA01 As String, ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, _
        ByVal pCP07 As String, ByVal pCP09 As String, ByVal pCP10 As String, ByVal pCP14 As String, ByVal pCP27 As String)
'pNa01:申請國家
'Memo by Lydia 2024/07/31 與秀玲討論：因為變更NA69會整批更新未發文和未續辦的下一程序，所以傳入模組統一使用CP14
Dim intR As Integer, strR1 As String, strCF11 As String
Dim strNP07 As String, strNP08 As String, strNP22 As String
Dim rsRD As New ADODB.Recordset

   strR1 = "SELECT * FROM CaseFee Where CF01='" & pCP01 & "' AND CF02='" & Mid(pNA01, 1, 3) & "' AND CF03='" & pCP10 & "' "
   intR = 1
   Set rsRD = ClsLawReadRstMsg(intR, strR1)
   If intR = 1 Then
      strCF11 = "" & rsRD.Fields("CF11")
   End If
   
   '指定的案件性質，先以系統類別+申請國家+案件性質抓Casefee的CF11，若NVL(CF11,0)=0或者無Casefee則固定設為30天。
   '案件性質來源：frm050303.strExceptCFT
   'Modified by Lydia 2024/07/31 8/1先啟用，後續確認發文日1100101~1130731+系統別CFT,CFC,S,TF的未提申案件
   'Modified by Lydia 2025/08/15 調整1.提申管制設定僅針對國家代碼為011~999的國家，不包含台灣案; 2.「文件公／簽證」(711)的提申管制只有在S案時才需抓=>區分不同設定
   'If Val(strCF11) = 0 And InStr("CFT,CFC,S,TF,", pCP01 & ",") > 0 And strSrvDate(1) >= "20240801" Then
   '   strR1 = "SELECT * FROM SETSPECMAN WHERE OCODE='CFT_GETCF11' AND INSTR(','||oman||',','," & pCP10 & ",') > 0 "
   If Val(strCF11) = 0 And InStr("CFT,CFC,S,TF,", pCP01 & ",") > 0 And Mid(pNA01, 1, 3) >= "011" And Mid(pNA01, 1, 3) <= "999" Then
      If pCP01 = "CFT" Then
         strR1 = "SELECT * FROM SETSPECMAN WHERE OCODE='CFT提申管制_CFT' AND INSTR(','||oman||',','," & pCP10 & ",') > 0 "
      Else
         strR1 = "SELECT * FROM SETSPECMAN WHERE OCODE='CFT提申管制_CFC_S' AND INSTR(','||oman||',','," & pCP10 & ",') > 0 "
      End If
   'end 2025/08/13
      intR = 1
      Set rsRD = ClsLawReadRstMsg(intR, strR1)
      If intR = 1 Then
         strCF11 = "30"
      End If
   End If
   If Val(strCF11) > 0 Then
      strNP07 = "998"
      strNP08 = DBDATE(IIf(Val(pCP27) = 0, strSrvDate(1), pCP27))
      strNP08 = DBDATE(DateAdd("d", Val(strCF11), ChangeWStringToWDateString(strNP08)))
      '檢查期限是否正確
      strNP08 = PUB_T997998LimitDate(strNP08, pCP07, 1)
      strNP22 = GetNextProgressNo()
      strR1 = "INSERT INTO NextProgress (NP01,NP02,NP03,NP04,NP05,NP07,NP08,NP09,NP10,NP22) " & _
               "VALUES ('" & pCP09 & "','" & pCP01 & "','" & pCP02 & "','" & pCP03 & "','" & pCP04 & "'," & strNP07 & "," & _
                         PUB_GetWorkDay1(strNP08, True) & "," & strNP08 & ",'" & pCP14 & "'," & strNP22 & ")"
      cnnConnection.Execute strR1
   End If
   
   Set rsRD = Nothing
End Sub

'Added by Morgan 2024/8/14
'Base64文字轉存成檔案
Public Function PUB_FromBase64ToFile(pBase64Text As String, pFullFileName As String) As Boolean
   Dim lSize           As Long
   Dim baOutput()      As Byte
   Dim fnum As Integer
    
On Error GoTo ErrHnd

   lSize = Len(pBase64Text) + 1
   ReDim baOutput(0 To lSize - 1) As Byte
   If CryptStringToBinary(StrPtr(pBase64Text), Len(pBase64Text), 1, VarPtr(baOutput(0)), lSize) <> 0 Then
      ReDim Preserve baOutput(0 To lSize - 1) As Byte
   End If
      
   fnum = FreeFile()
   Open pFullFileName For Binary As #fnum
   Put #fnum, 1, baOutput
   Close fnum
   PUB_FromBase64ToFile = True
   Exit Function
   
ErrHnd:
   MsgBox Err.Description, vbCritical
   
End Function


'Added by Lydia 2024/10/30 智慧所案號取得出庭律師清單
Public Function PUB_GetLosCL02list(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String) As String
Dim intQ As Integer, strQ1 As String
Dim rsQuery As New ADODB.Recordset
   
   PUB_GetLosCL02list = ""
   If pCP01 = "" Or pCP02 = "" Then Exit Function
   
   '以輸入的本所案號之所有收文號串法律所案源資料，抓出法律所案件有承辦人且收文日最大的進度，抓承辦人及所有出庭律師。
   strQ1 = "select cp14,cl02 from caseprogress,caselawer where cp09=(select substr(mno,9,9) from (" & _
           "select max(c2.cp05||c2.cp09) as mno from caseprogress c1,lawofficesource,caseprogress c2 " & _
           "where c1.cp01='" & pCP01 & "' and c1.cp02='" & pCP02 & "' and c1.cp03='" & pCP03 & "' and c1.cp04='" & pCP04 & "' and c1.cp159=0 " & _
           "and c1.cp09=los01(+) and los01 is not null and los06=c2.cp09(+) and nvl(c2.cp14,'N') <>'N' " & _
           ")) and cp09=cl01(+) "
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, strQ1)
   If intQ = 1 Then
      strQ1 = ""
      rsQuery.MoveFirst
      Do While Not rsQuery.EOF
         If InStr(strQ1 & ",", "" & rsQuery.Fields("CP14")) = 0 And "" & rsQuery.Fields("CP14") <> "" Then
            strQ1 = strQ1 & ";" & rsQuery.Fields("CP14")
         End If
         If InStr(strQ1 & ",", "" & rsQuery.Fields("CL02")) = 0 And "" & rsQuery.Fields("CL02") <> "" Then
            strQ1 = strQ1 & ";" & rsQuery.Fields("CL02")
         End If
         rsQuery.MoveNext
      Loop
      PUB_GetLosCL02list = Mid(strQ1, 2)
   End If
   Set rsQuery = Nothing
End Function

'Add by Amy 2025/05/16 有回覆單開啟卷宗區,Msg檔自動開啟(開[卷宗區]程式參考frm210147_1)
'intState:0-智權結案單/1-外商結案單/2-外專結案單
'stCaseNo:傳入含-之案號
'stSaveFiles:傳入之檔案
'bolOnly100101_L:只Run 卷宗區,不需檢查是否只有一個MSG
Public Sub Pub_OpenReplayPDFOrMsg(ByVal intState As Integer, ByRef NowFrm As Form, stCaseNo As String, stCCM01 As String, stSaveFiles As String, stLoadPath As String, ByRef stMsg As String _
      , Optional ByVal bolOnly100101_L As Boolean = False)
   Dim RsQ As New ADODB.Recordset, intQ As Integer, stQ As String
   Dim hLocalFile As Long, arrData As Variant
   Dim mFile As FileListBox, mFs
   Dim ii As Integer, jj As Integer, intCntMsg As Integer, bolShowMsg As Boolean
   Dim stRepPath As String, stTPF As String, stTP As String
   Dim stGetMsg As String, stMsg_Open As String, stMsg_EFile As String 'Add by Amy 2025/06/30
   
   Select Case UCase(NowFrm.Name)
      Case UCase("frm110101_2")
      Case UCase("frm210147_1")
      Case UCase("frm210148_1")
      Case UCase("frm210149_1")
   End Select
   
   bolShowMsg = False: stMsg = ""
   stRepPath = stLoadPath
   If Right(stRepPath, 1) <> "\" Then stRepPath = stRepPath & "\"
   
   arrData = Split(stSaveFiles, "&") '多檔改為&串
   
   If bolOnly100101_L = True Then
      'Memo 目前bolOnly100101_L=True 為補看按「完整卷宗區」鈕進入
      '              由frm210148_1進會有兩個按鈕 檢視回覆單(cmdFile) 及 完整卷宗 (Command1)
   Else
      'FC結案單,才判斷是否有Msg
      If intState > 0 Then
         For jj = LBound(arrData) To UBound(arrData)
            '計算 .MSG 檔案數量
            If Right(UCase(arrData(jj)), 4) = ".MSG" Then
               intCntMsg = intCntMsg + 1
            End If
         Next jj
         If intCntMsg = 1 And intCntMsg = UBound(arrData) + 1 Then
            '只有一個.MSG 檔,直接開檔
            bolShowMsg = True
         End If
      End If
   End If
   
   If bolShowMsg = True Then
   '*** 直接DownLoad 資料 ***
      stQ = "Select cpp02,cpp14,cpp19 From CasePaperPdf where cpp11='" & stCCM01 & "' "
      intQ = 1
      Set RsQ = ClsLawReadRstMsg(intQ, stQ)
      If intQ = 1 Then
         Do While RsQ.EOF = False
            'Modify by Amy 2025/06/30 避免訊息彈多次,+stGetMsg (FC結案單可存.MSG及.PDF)
            stGetMsg = "1"
            If PUB_GetFtpFile(RsQ.Fields("cpp14"), stRepPath & RsQ.Fields("cpp02"), , , , "" & RsQ.Fields("cpp19") <> "", stGetMsg) = False Then
               If InStr(stGetMsg, "檔案已開啟") > 0 Then
                  stMsg_Open = stMsg_Open & ";" & stGetMsg
               ElseIf stGetMsg <> "" Then
                  stMsg_EFile = stMsg_EFile & ";" & RsQ.Fields("cpp02")
               End If
            End If
            'end 2025/06/30
            RsQ.MoveNext
         Loop
      End If
      'Modify by Amy 2025/06/30 區分訊息
      If stMsg_Open <> "" Or stMsg_EFile <> "" Then
         If stMsg_Open <> "" Then
            stMsg = Replace(Mid(stMsg_Open, 2), ";", vbCrLf)
         End If
         If stMsg_EFile <> "" Then
            If stMsg <> "" Then stMsg = stMsg & vbCrLf
            stMsg = "附件下載失敗:" & vbCrLf & Replace(Mid(stMsg_EFile, 2), ";", vbCrLf) & vbCrLf & _
                           "請洽電腦中心"
         End If
         Exit Sub
      End If
      '*** End 直接DownLoad 資料 ***
      '*** 直接開檔 ***
      For jj = LBound(arrData) To UBound(arrData)
         '不確定傳入之檔案是否有路徑,故先取代再加入
         stTPF = stLoadPath & Replace(arrData(ii), stLoadPath, "")
         Call PUB_ChkFileTypeOpenExE(stTPF)
         '開啟檔案
         ShellExecute hLocalFile, "open", stTPF, vbNullString, vbNullString, 1
      Next jj
      
   Else
      '開 卷宗區
      frm100101_L.m_strKey = stCaseNo
      frm100101_L.SetParent NowFrm
      If frm100101_L.QueryData(stCCM01) = True Then
         If UCase(NowFrm.Name) = "FRM210148_1" And bolOnly100101_L = True Then
            NowFrm.cmdAction = 9 '預設結束
            frm100101_L.cmdOK(1).Caption = "同意"
            frm100101_L.cmdOK(1).Visible = True
         End If
         For jj = LBound(arrData) To UBound(arrData)
            For ii = frm100101_L.GRD1.Rows - 1 To 1 Step -1
               If InStr(frm100101_L.GRD1.TextMatrix(ii, 4), Replace(arrData(jj), stRepPath, "")) > 0 Then Exit For
            Next ii
            If ii > 0 Then
               Call frm100101_L.FrmCallOpenFile(ii, IIf(UBound(arrData) = jj, True, False))
               If UBound(arrData) = jj Then
                  frm100101_L.Show
                  NowFrm.Hide
               End If
            Else
               Unload frm100101_L
               Screen.MousePointer = vbDefault
               stMsg = "有電子檔:" & arrData(jj) & " (找不到電子檔)"
               NowFrm.Enabled = True
               Exit Sub
            End If
         Next jj
      Else
         Unload frm100101_L
      End If
   End If
End Sub

'Added by Lydia 2025/06/30 內專/外專/內商/外商分案作業，取消閉卷時，若下一程序有未過期且已上N之指定下一程序，都自動取消NP06、NP11、NP12，並彈訊息提醒已還原XXX期限。
Public Function PUB_GetCaseCloseCancel(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pNA01 As String) As String
Dim strQ1 As String, strEx As String, intQ As Integer
Dim rsQD As New ADODB.Recordset
   
   PUB_GetCaseCloseCancel = ""
   If InStr(pCP01, "T") > 0 Then
      '商標：延展102、使用宣誓105期限
      strQ1 = "102,105"
   ElseIf InStr(pCP01, "P") > 0 Then
      '專利：年費605、維持費606、延展費607
      strQ1 = "605,606,607"
   End If
   If strQ1 <> "" And pCP01 <> "" And pCP02 <> "" Then
      strQ1 = "SELECT np01,np22,np02,np03,np04,np05,np07," & IIf(pNA01 = "000", "nvl(cpm03,cpm04)", "nvl(cpm04,cpm03)") & " as np07n FROM nextprogress,casepropertymap" & _
              " WHERE np02='" & pCP01 & "' AND np03='" & pCP02 & "' AND np04='" & IIf(pCP03 = "", "", "0") & "' AND np05='" & IIf(pCP04 = "", "", "00") & "' AND instr('" & strQ1 & "',np07) > 0" & _
              " AND np02=cpm01(+) AND np07=cpm02(+) AND np06='N' AND np09>=to_char(SYSDATE,'yyyymmdd') "
      intQ = 1
      Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
      If intQ = 1 Then
         rsQD.MoveFirst
         Do While Not rsQD.EOF
            strEx = "Update NextProgress set Np06=null, Np11=null, Np12=null where np01='" & rsQD.Fields("np01") & "' and np22='" & rsQD.Fields("np22") & "' and np02='" & rsQD.Fields("np02") & "' and np03='" & rsQD.Fields("np03") & "' and np04='" & rsQD.Fields("np04") & "' and np05='" & rsQD.Fields("np05") & "' "
            cnnConnection.Execute strEx
            PUB_GetCaseCloseCancel = PUB_GetCaseCloseCancel & "、" & rsQD.Fields("np07n")
            rsQD.MoveNext
         Loop
         PUB_GetCaseCloseCancel = Mid(PUB_GetCaseCloseCancel, 2, Len(PUB_GetCaseCloseCancel) - 1)
      End If
      'Added by Lydia 2025/07/01 TF馬德里案:恢復子案的使用宣誓
      If pCP01 = "TF" And pNA01 = "238" Then
         strQ1 = "SELECT np01,np22,np02,np03,np04,np05,np07," & IIf(pNA01 = "000", "nvl(cpm03,cpm04)", "nvl(cpm04,cpm03)") & " as np07n FROM nextprogress,casepropertymap" & _
                 " WHERE np02='" & pCP01 & "' AND substr(np03,1,5)='" & Mid(pCP02, 1, 5) & "' and np04||np05<>'000' AND np07='105' " & _
                 " AND np02=cpm01(+) AND np07=cpm02(+) AND np06='N' AND np09>=to_char(SYSDATE,'yyyymmdd') "
         intQ = 1
         Set rsQD = ClsLawReadRstMsg(intQ, strQ1)
         If intQ = 1 Then
            rsQD.MoveFirst
            Do While Not rsQD.EOF
               strEx = "Update NextProgress set Np06=null, Np11=null, Np12=null where np01='" & rsQD.Fields("np01") & "' and np22='" & rsQD.Fields("np22") & "' and np02='" & rsQD.Fields("np02") & "' and np03='" & rsQD.Fields("np03") & "' and np04='" & rsQD.Fields("np04") & "' and np05='" & rsQD.Fields("np05") & "' "
               cnnConnection.Execute strEx
               PUB_GetCaseCloseCancel = PUB_GetCaseCloseCancel & "、" & rsQD.Fields("np07n")
               rsQD.MoveNext
            Loop
            PUB_GetCaseCloseCancel = Mid(PUB_GetCaseCloseCancel, 2, Len(PUB_GetCaseCloseCancel) - 1)
         End If
      End If
      'end 2025/07/01
   End If
   Set rsQD = Nothing
End Function

'Added by Lydia 2025/09/12 傳入TF基礎案號數+基礎案申請國家+是否要寄Email，傳出基礎案本所案號/TF馬德里案號
Public Function PUB_GetTFbaseInfo(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pBaseNo As String, ByVal pBaseNA01 As String, ByVal pMailType As String, Optional ByVal pBaseNo2 As String, Optional pNowCP09 As String) As String
'pCP01~pCP04：傳入TF案號 or 本所案號
'pBaseNo：傳入TF基礎案號數 or 本所案號(審定號) + pBaseNo2申請案號
'pBaseNA01：傳入TF基礎案申請國家 or 本所案號(申請國)
'pMailType : 1-基本檔維護和分案作業, 2-來函和收文
'pNowCP09：傳入新增收文號
Dim strTF(1 To 11) As String    'TF案號1~4,5案件名稱,6,7申請國家代號+國名,8審定號/申請號,9是否閉卷,10註冊日,11目前准駁
Dim strBase(1 To 11) As String  '基礎案本所案號1~4,5案件名稱,6,7申請國家代號+國名,8審定號/申請號,9是否閉卷,10註冊日,11目前准駁
Dim intQ As Integer, strQuery As String
Dim rsQD As New ADODB.Recordset
Dim intK As Integer, strK1 As String
Dim rsAD As New ADODB.Recordset
Dim strTo As String, strSubject As String, strContent As String, strCC As String

   
   PUB_GetTFbaseInfo = ""
   If pCP01 = "TF" Then
      strTF(1) = pCP01: strTF(2) = pCP02: strTF(3) = pCP03: strTF(4) = pCP04
      strQuery = "select '1' as ord1,tm01,tm02,tm03,tm04,tm15 as pno,tm28 from trademark where tm15='" & pBaseNo & "' and tm10='" & pBaseNA01 & "' " & _
                 "UNION select '2' as ord1,tm01,tm02,tm03,tm04,tm12 as pno,tm28 from trademark where tm12='" & pBaseNo & "' and tm10='" & pBaseNA01 & "' "
   Else
      strBase(1) = pCP01: strBase(2) = pCP02: strBase(3) = pCP03: strBase(4) = pCP04
      'Modified by Lydia 2025/10/?? TF基礎案號改成Table
      'strQuery = "select '1' as ord1,tm01,tm02,tm03,tm04,tm28 from trademark where tm06='" & pBaseNo & "' and tm07='" & pBaseNA01 & "' "
      'If pBaseNo2 <> "" Then strQuery = strQuery & "UNION select '2' as ord1,tm01,tm02,tm03,tm04,tm28 from trademark where tm06='" & pBaseNo2 & "' and tm07='" & pBaseNA01 & "' "
      strQuery = "select '1' as ord1,tm01,tm02,tm03,tm04,tm28 from trademark,tfbaseno where tm01=tfbn01(+) and tm02=tfbn02(+) and tm03=tfbn03(+) and tm04=tfbn04(+) and tfbn05='" & pBaseNo & "' and tfbn06='" & pBaseNA01 & "' "
      If pBaseNo2 <> "" Then strQuery = strQuery & "UNION select '2' as ord1,tm01,tm02,tm03,tm04,tm28 from trademark,tfbaseno where tm01=tfbn01(+) and tm02=tfbn02(+) and tm03=tfbn03(+) and tm04=tfbn04(+) and tfbn05='" & pBaseNo2 & "' and tfbn06='" & pBaseNA01 & "' "
   End If
   strQuery = strQuery & " order by ord1,tm01,tm02,tm03,tm04 "
   intQ = 1
   Set rsQD = ClsLawReadRstMsg(intQ, strQuery)
   If intQ = 1 Then
       rsQD.MoveFirst
       PUB_GetTFbaseInfo = IIf("" & rsQD.Fields("tm28") = "1", "", "N") & rsQD.Fields("tm01") & "-" & rsQD.Fields("tm02") & "-" & rsQD.Fields("tm03") & "-" & rsQD.Fields("tm04")
       If pCP01 = "TF" Then
          strBase(1) = "" & rsQD.Fields("tm01"): strBase(2) = "" & rsQD.Fields("tm02"): strBase(3) = "" & rsQD.Fields("tm03"): strBase(4) = "" & rsQD.Fields("tm04")
          strBase(8) = "" & rsQD.Fields("pno")
       Else
          strTF(1) = "" & rsQD.Fields("tm01"): strTF(2) = "" & rsQD.Fields("tm02"): strTF(3) = "" & rsQD.Fields("tm03"): strTF(4) = "" & rsQD.Fields("tm04")
       End If
       '已超過TF國際註冊日五年=>不發Email；
       '符合下列情況 , 系統自動發送通知信, 提醒馬德里承辦人及其主管:
       'i. 基礎案已閉卷、【703不續辦-續展】(即703之CP43抓NP01且NP07=102者)
       'ii.   官方來函【1002核駁】、【1004敗訴】、【1006部分勝部分敗】，若在商標基本檔維護或內商分案作業，還要判斷無後續AB類收文才發Email
       If pMailType <> "" Then  '要寄Email
          strK1 = "select '1' ord1, tm01,tm02,tm03,tm04,tm05,tm10,tm15,tm12,na03,tm20,decode(tm29,null,null,'Y') as pclose,tm16 from trademark,nation where tm01='" & strTF(1) & "' and tm02='" & strTF(2) & "' and tm03='" & strTF(3) & "' and tm04='" & strTF(4) & "' and tm10=na01(+) " & _
                  "Union select '2' ord1, tm01,tm02,tm03,tm04,tm05,tm10,tm15,tm12,na03,tm20,decode(tm29,null,null,'Y') as pclose,tm16 from trademark,nation where tm01='" & strBase(1) & "' and tm02='" & strBase(2) & "' and tm03='" & strBase(3) & "' and tm04='" & strBase(4) & "' and tm10=na01(+) " & _
                  "order by ord1 "
          intK = 1
          Set rsAD = ClsLawReadRstMsg(intK, strK1)
          If intK = 1 Then
             rsAD.MoveFirst
             Do While Not rsAD.EOF
                '案號1~4,5案件名稱,6,7申請國家代號+國名,8審定號/申請號,9是否閉卷,10註冊日
                If "" & rsAD.Fields("ord1") = "1" Then
                   strTF(5) = "" & rsAD.Fields("tm05")
                   strTF(6) = "" & rsAD.Fields("tm10")
                   strTF(7) = "" & rsAD.Fields("na03")
                   If strTF(8) = "" Then
                      strTF(8) = IIf("" & rsAD.Fields("tm15") <> "", "" & rsAD.Fields("tm15"), "" & rsAD.Fields("tm12"))
                   End If
                   strTF(9) = "" & rsAD.Fields("pclose")
                   strTF(10) = "" & rsAD.Fields("tm20")
                   strTF(11) = "" & rsAD.Fields("tm16")
                Else
                   strBase(5) = "" & rsAD.Fields("tm05")
                   strBase(6) = "" & rsAD.Fields("tm10")
                   strBase(7) = "" & rsAD.Fields("na03")
                   If strBase(8) = "" Or pMailType = "2" Then
                      strBase(8) = IIf("" & rsAD.Fields("tm15") <> "", "" & rsAD.Fields("tm15"), "" & rsAD.Fields("tm12"))
                   End If
                   strBase(9) = "" & rsAD.Fields("pclose")
                   strBase(10) = "" & rsAD.Fields("tm20")
                   strBase(11) = "" & rsAD.Fields("tm16")
                End If
                rsAD.MoveNext
             Loop
          End If
'*************連結通知Email：基本檔維護和分案作業
          If pMailType = "1" Then
             '收受者，請抓TF案A或B類進度最後在職之承辦人
             'Modified by Lydia 2025/10/08 改為TF案A或B類進度最後在職之商標一組(ST93=T11)承辦人，副本為其主管；若無商標一組(ST93=T11)承辦人則直接發T11之部門主管。 +and st93='T11'
             strK1 = "select max(cp05||cp09||cp14) mno from caseprogress,staff where cp01='" & strTF(1) & "' and cp02='" & strTF(2) & "' and cp03='" & strTF(3) & "' and cp04='" & strTF(4) & "' " & _
                     "and cp159=0 and cp09 <'C' and cp14=st01(+) and st04='1' and st93='T11' "
             intK = 1
             Set rsAD = ClsLawReadRstMsg(intK, strK1)
             If intK = 1 Then
                strTo = Mid("" & rsAD.Fields("mno"), 18)
             End If
             If strTo = "" Then
                strTo = GetDeptA09("T11", "24", True)
             Else
                strCC = GetDeptA09(PUB_GetST93(strTo), "24", True)
                If strCC = strTo Then strCC = ""
             End If
             '用(1)排序MailCache
             strSubject = "(1)馬德里基礎案連結建立，請確認！"
             strContent = "有關馬德里" & strTF(1) & "-" & strTF(2) & "-" & strTF(3) & "-" & strTF(4) & "(" & strTF(8) & ")「" & strTF(5) & "」，" & vbCrLf & _
                          strBase(1) & "-" & strBase(2) & "-" & strBase(3) & "-" & strBase(4) & "(" & strBase(7) & ")(" & strBase(8) & ")「" & strBase(5) & "」判斷應為" & strTF(1) & "-" & strTF(2) & "-" & strTF(3) & "-" & strTF(4) & "馬德里基礎案，系統已建立案件連結。" & vbCrLf & vbCrLf
             strContent = strContent & "<B><U>※請承辦人員確認，若資訊有誤請盡速通知程序人員修改資料。</B></U>" & vbCrLf  '字體加粗加底線
             '※基礎案為本所之對造案件，請確認兩造關係。
             If pCP01 = "TF" And Left(PUB_GetTFbaseInfo, 1) = "N" Then
                 strContent = strContent & vbCrLf & "<B><U>※基礎案為本所之對造案件，請確認兩造關係。</B></U>" & vbCrLf
             End If
             
             strContent = Replace(strContent, vbCrLf, "<BR>")
             strContent = Replace(strContent, "  ", "&nbsp;&nbsp;")
             strK1 = ""
             strK1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                    " VALUES ( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
                    ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSubject) & "','" & ChgSQL(strContent) & "','" & strCC & "')"
             cnnConnection.Execute strK1
             Sleep 1000 'Added by Lydia 2025/10/23 避免同一時分秒
          End If
          
'*************基礎案狀態通知Email
          '已超過TF國際註冊日五年=>不發Email
          If strTF(10) <> "" Then
             If CompDate(0, 5, strTF(10)) < strSrvDate(1) Then
                GoTo JumpToNoMail
             End If
          End If
          
          strQuery = ""
          '基礎案已閉卷：單純看基本檔
          If strBase(9) = "Y" And pNowCP09 = "" Then
              strQuery = "閉卷"
          End If

          '有收文703不續辦-延展; CFT案延展（英國）110
          If strQuery = "" Then
             strK1 = "select cp09," & IIf(strBase(6) = "000", "nvl(cpm03,cpm04)", "nvl(cpm04,cpm03)") & "||GetRelateCasePropertyName(cp09,'1') as cp10n " & _
                     "from caseprogress,nextprogress,casepropertymap where cp10='703' and cp01='" & strBase(1) & "' and cp02='" & strBase(2) & "' and cp03='" & strBase(3) & "' and cp04='" & strBase(4) & "' " & _
                     "and cp43=np01(+) and (np07='102' or (np02='CFT' and np07='110')) and np06='N' and cp01=cpm01(+) and cp10=cpm02(+) " & IIf(pNowCP09 <> "", "and cp09='" & pNowCP09 & "' ", "")
             intK = 1
             Set rsAD = ClsLawReadRstMsg(intK, strK1)
             If intK = 1 Then
                strQuery = rsAD.Fields("cp10n")
             End If
          End If
          '有官方來函【1002核駁】、【1004敗訴】、【1006部分勝部分敗】或【307註銷】發文，若在商標基本檔維護或內商分案作業，還要判斷無後續AB類收文才發Email
          'CFT審查報告輸入(類似T之核駁)性質: 1201審查報告、1403重為處分
          If strQuery = "" Then
             strK1 = "select max(cp05||cp09) maxno from caseprogress where cp01='" & strBase(1) & "' and cp02='" & strBase(2) & "' and cp03='" & strBase(3) & "' and cp04='" & strBase(4) & "' " & _
                     "and cp159=0 and (cp10 in ('1002','1004','1006','307') or (cp01='CFT' and cp10 in ('1201','1403'))) " & IIf(pNowCP09 <> "", "and cp09='" & pNowCP09 & "' ", "")
             intK = 1
             Set rsAD = ClsLawReadRstMsg(intK, strK1)
             If intK = 1 Then
                strQuery = "" & rsAD.Fields("maxno")
                If strQuery <> "" Then
                   If pMailType = "1" Then
                      strK1 = "select cp09,cp10 from caseprogress where cp05>=" & Mid(strQuery, 1, 8) & " and cp09<'C' and cp159=0 and cp01='" & strBase(1) & "' and cp02='" & strBase(2) & "' and cp03='" & strBase(3) & "' and cp04='" & strBase(4) & "' "
                      intK = 1
                      Set rsAD = ClsLawReadRstMsg(intK, strK1)
                      If intK = 1 Then
                         strQuery = ""
                      End If
                   End If
                   If strQuery <> "" Then
                      strK1 = "select cp09," & IIf(strBase(6) = "000", "nvl(cpm03,cpm04)", "nvl(cpm04,cpm03)") & "||GetRelateCasePropertyName(cp09,'1') as cp10n " & _
                              "from caseprogress, casepropertymap where cp09='" & Mid(strQuery, 9, 9) & "' and cp01=cpm01(+) and cp10=cpm02(+) "
                      intK = 1
                      Set rsAD = ClsLawReadRstMsg(intK, strK1)
                      If intK = 1 Then
                         strQuery = "" & rsAD.Fields("cp10n")
                      Else
                         strQuery = ""
                      End If
                   End If
                End If
             End If
          End If
          '基礎案被駁：單純看基本檔
          If strQuery = "" And strBase(11) = "2" Then
              strQuery = "核駁"
          End If
          
          If strQuery <> "" Then
             If strTo = "" Then
                '收受者，請抓TF案A或B類進度最後在職之承辦人
                'Modified by Lydia 2025/10/08 改為TF案A或B類進度最後在職之商標一組(ST93=T11)承辦人，副本為其主管；若無商標一組(ST93=T11)承辦人則直接發T11之部門主管。 +and st93='T11'
                strK1 = "select max(cp05||cp09||cp14) mno from caseprogress,staff where cp01='" & strTF(1) & "' and cp02='" & strTF(2) & "' and cp03='" & strTF(3) & "' and cp04='" & strTF(4) & "' " & _
                        "and cp159=0 and cp09 <'C' and cp14=st01(+) and st04='1' and st93='T11' "
                intK = 1
                Set rsAD = ClsLawReadRstMsg(intK, strK1)
                If intK = 1 Then
                   strTo = Mid("" & rsAD.Fields("mno"), 18)
                End If
                If strTo = "" Then
                   strTo = GetDeptA09("T11", "24", True)
                Else
                   strCC = GetDeptA09(PUB_GetST93(strTo), "24", True)
                   If strCC = strTo Then strCC = ""
                End If
             End If
             '用(2)排序MailCache
             strSubject = "(2)通知-馬德里國際註冊基礎案可能失效（請轉知提醒智權人員）"
             strContent = "馬德里案號：" & strTF(1) & "-" & strTF(2) & "-" & strTF(3) & "-" & strTF(4) & "(" & strTF(8) & ")「" & strTF(5) & "」" & vbCrLf & _
                          "所內基礎案：" & strBase(1) & "-" & strBase(2) & "-" & strBase(3) & "-" & strBase(4) & "(" & strBase(7) & ")(" & strBase(8) & ")「" & strBase(5) & "」，" & strQuery & vbCrLf & vbCrLf & _
                          "依馬德里商標相關規定，當馬德里國際註冊日未屆滿5年，若基礎案失效（註銷、撤銷、駁回、放棄等），將可能導致國際註冊將自動部分/全部失效。" & vbCrLf & _
                          "●按規定，基礎國的商標主管機關應主動通知WIPO基礎案失效情況，以啟動中心打擊程序。" & vbCrLf & _
                          "●然而，實務上部分國家（如中國）不會主動通知WIPO。" & vbCrLf & _
                          "●因此，若無第三方提出申請或異議指出基礎案失效，國際註冊可能繼續有效。" & vbCrLf & vbCrLf
             strContent = strContent & "<B><U>※請承辦人員查詢確認基礎案及馬德里國際註冊商標情況後，發信轉知提醒智權人員。</B></U>" '字體加粗加底線
             strContent = Replace(strContent, vbCrLf, "<BR>")
             strContent = Replace(strContent, "  ", "&nbsp;&nbsp;")
             strK1 = ""
             strK1 = "insert into mailcache(mc01,mc02,mc03,mc04,mc07,mc08,mc09)" & _
                    " VALUES ( '" & strUserNum & "','" & strTo & "',to_char(sysdate,'yyyymmdd')" & _
                    ",to_char(sysdate,'hh24miss'),'" & ChgSQL(strSubject) & "','" & ChgSQL(strContent) & "','" & strCC & "')"
             cnnConnection.Execute strK1
             Sleep 1000 'Added by Lydia 2025/10/25 避免同一時分秒
          End If
       End If  'If pMailType <> "" Then---要寄Email
   End If   'If intQ = 1 Then  --- 基礎案為本所案號
   
JumpToNoMail:
   If Trim(PUB_GetTFbaseInfo) = "" Then
      PUB_GetTFbaseInfo = "非本所案件" '基礎案非本所案件
   End If
   Set rsQD = Nothing
   Set rsAD = Nothing
   
End Function

'Added by Lydia 2025/11/18 輸入機關代號/機關名稱，取得相關資料
Public Function PUB_ChkGovIsExist(ByVal pVAL01 As String, Optional ByRef pStrNo As String, Optional ByRef pStrName As String) As Boolean
Dim strA1 As String, intA As Integer
Dim rsAD As New ADODB.Recordset
   
   PUB_ChkGovIsExist = False
   pStrNo = ""
   pStrName = ""
   If pVAL01 <> "" Then
      strA1 = "select '1' as ord1, or01, or02 from organization where or01='" & pVAL01 & "' " & _
              "union select '2' as ord1, or01, or02 from organization where or02 like '%" & pVAL01 & "%' " & _
              "order by ord1,or01,or02 "
      intA = 1
      Set rsAD = ClsLawReadRstMsg(intA, strA1)
      If intA = 1 Then
         rsAD.MoveFirst
         pStrNo = "" & rsAD.Fields("or01")
         pStrName = "" & rsAD.Fields("or02")
         PUB_ChkGovIsExist = True
      Else
         ShowMsg MsgText(9212)
      End If
   End If
   Set rsAD = Nothing
End Function

'Added by Lydia 2025/11/18 預設機關代號清單
Public Sub PUB_SetGovCmb(ByRef pCmb As Object, ByRef p_ItemDataList As String, Optional ByVal p_Default As String, Optional ByVal p_Cond As String)
Dim strA1 As String, intA As Integer
Dim tmpArr As Variant
Dim rsAD As New ADODB.Recordset
   
   pCmb.Clear
   p_ItemDataList = ""
   strA1 = "select or01,or02 from organization " & IIf(p_Cond <> "", "where or02 like '%" & ChgSQL(p_Cond) & "%'", "") & _
           "order by or01 desc "
   intA = 1
   Set rsAD = ClsLawReadRstMsg(intA, strA1)
   If intA = 1 Then
      rsAD.MoveFirst
      Do While Not rsAD.EOF
         pCmb.AddItem rsAD.Fields("or01") & " " & rsAD.Fields("or02"), 0
         p_ItemDataList = rsAD.Fields("or01") & IIf(p_ItemDataList <> "", ",", "") & p_ItemDataList
         rsAD.MoveNext
      Loop
   Else
      If p_Cond <> "" Then
         ShowMsg MsgText(9212)
      End If
   End If
   If p_Default <> "" Then
      'p_ItemDataList = "0," & p_ItemDataList  '不用增加空白項
      tmpArr = Split(p_ItemDataList, ",")
      For intA = 0 To pCmb.ListCount - 1
         If Trim(tmpArr(intA)) <> "0" And tmpArr(intA) = p_Default Then
            pCmb.ListIndex = intA
            Exit For
         End If
      Next
   End If
   '先不預設
   'If pCmb.ListIndex = -1 Then
   '   pCmb.ListIndex = 0
   'End If
   Set rsAD = Nothing
End Sub
