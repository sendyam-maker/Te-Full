Attribute VB_Name = "basStart"
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit

' 89.08.17 Louis 列印國內案件接洽及結案記錄單的物件
Sub Main()

'Added by Morgan 2022/7/15
frmpic002.m_bFixIME = True
frmpic002.Show vbModal
'end 2022/7/15

pub_strCommand = Command()
If pub_strCommand = "" And InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then MsgBox "【" & App.EXEName & "】不可直接執行，請以桌面台一圖示【TeAutoUpd】啟動！", vbCritical: End 'Added by Morgan 2017/10/6
'pub_strCommand = "T" 'Add By Sindy 2022/7/14 TEST用

Dim fso As Object 'Add By Sindy 2014/9/26
Dim stTestAccount As String 'Added by Morgan 2019/3/26 測試帳號

'add by nickc 2005/09/20
Pub_Can_Copy_Pic = False
'Add By Cheng 2001/12/28
'禁止重覆開啟系統
'If App.PrevInstance Then
'   MsgBox App.EXEName & "目前已在執行中...", vbExclamation, "無法重覆開啟"
'   End
'End If
''判斷是否在VB下執行, 若是在VB6下執行, 則不用執行更新動作, 否則需檢查是否要更新
'If StrComp(Right(Pub_GetModuleFileName, Len(App.EXEName & ".EXE")), _
   App.EXEName & ".EXE", vbTextCompare) = 0 Then
'   pub_bln_NeedUpdate = True
'   pub_bln_UpdateActive = False
'   '判斷是否要更新
'   Pub_CheckSysVer
'   Do While pub_bln_NeedUpdate
'      '若執行更新動作則關閉程式
'      If pub_bln_UpdateActive = True Then End
'   Loop
'End If
'Set objPublicData = CreateObject("prjTaieDll.clsPublicData")
'Set objLawDll = CreateObject("prjTaieLawDll.clsLaw")
'Add By Cheng 2003/07/04
'判斷是否在VB下執行, 若非在VB6下執行, 則須先進入登入系統畫面
pub_str_LoginSucceeded = ""
If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
    frmLogin.Show vbModal
End If

'若未進入登入系統畫面
If strUserNum = "" Then
'Add by Morgan 2005/12/14 加連線選擇
'   If objPublicData.ConnectToServer = False Then
'      End
'   End If
'   If objPublicData.ConnectToServer = False Then
'      End
'   End If
   
   If PUB_Connect2DB() = False Then
      End
   End If
   
   ''edit by nickc 2007/02/06 不用 dll 了
   'objPublicData.Connection = cnnConnection
'2005/12/14 end
    ''edit by nickc 2007/02/05 不用 dll 了
    'If objPublicData.SetUserData(strUserNum, strUserName, strGroup) = False Then
    If ClsPDSetUserData(strUserNum, strUserName, strGroup) = False Then
        End
    End If
    pub_str_LoginSucceeded = "1" '登入成功
'Remove by Morgan 2005/12/14 移到上面
'   Set cnnConnection = objPublicData.Connection
    'strUserNo = strUserNum   '2008/12/10 CANCEL BY SONIA 將所有strUserNo改為strUserNum
    'Set cnnConnection = objPublicData.Connection
    ''edit by nickc 2007/02/05 不用 dll 了
    'Set obj003.Connection = cnnConnection
    'Set objLawDll.Connection = cnnConnection
    
    PUB_SetStaffVar
    GetGroupDept
    mdiMain.Show
    DoEvents
    strSrvDate(1) = Format(ServerDate)
    strSrvDate(2) = Format(Val(strSrvDate(1)) - 19110000)
    
    'Add By Amy 2013/05/08 電腦中心看公告 Start
    'Salary無查詢程式修改公告 Account/CaSher/Finance用不同的basStart
    If UCase(App.EXEName) <> "SALARY" Then
        If MsgBox("是否要看程式修改公告?", vbCritical + vbYesNo + vbDefaultButton2, "詢問") = vbYes Then
            frm100131.Show
            MoveFormToCenter frm100131
            'Modify By Amy 2013/06/05 起始日改顯示系統日前5個工作天
            'frm100131.Text1(0) = ACDate(strSrvDate(1) - 7)
            frm100131.Text1(0) = PUB_GetWorkDayAfterSysDate(CDbl(strSrvDate(1)), -5)
            frm100131.Text1(1) = strSrvDate(2)
            frm100131.cmdSearch_Click: DoEvents
        End If
    End If
    '2013/05/08 End
    
    Set adoTaie = cnnConnection
    
    'Modify by Morgan 2008/12/18 人事薪資不用連 mdb
    'Modify by Sindy 2011/8/2 電子簽核不用連 mdb
    If UCase(App.EXEName) <> "SALARY" And UCase(App.EXEName) <> "PERSON" And UCase(App.EXEName) <> "ABSENCE" Then

      'Removed by Morgan 2025/9/9 沒用了
      'cnnRptConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.path & "\db3.mdb"
      'cnnRptConn.Open
      'adoTemp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.path & "\finance.mdb"
      'adoTemp.Open
      'end 2025/9/9

      'Add By Cheng 2003/04/29
      'Modified by Morgan 2011/10/31 +PATPRO1
      'Modify by Sindy 2015/11/3 因經理測試時會改執行檔名稱+數字,所以判斷系統名稱開頭
      'If UCase(App.EXEName) = "PROMOTER" Or UCase(App.EXEName) = "TEPROMOTER" Or UCase(App.EXEName) = "PATPRO" Or UCase(App.EXEName) = "TEPATPRO" Or UCase(App.EXEName) = "PATPRO1" Or UCase(App.EXEName) = "TEPATPRO1" Then
      'Modify By Sindy 2021/7/16 + InStr(UCase(App.EXEName), "LAW") > 0
      'Modify By Sindy 2023/6/20 + InStr(UCase(App.EXEName), "TRADEMARK1") > 0
      If InStr(UCase(App.EXEName), "PROMOTER") > 0 Or _
         InStr(UCase(App.EXEName), "TEPROMOTER") > 0 Or _
         InStr(UCase(App.EXEName), "PATPRO") > 0 Or _
         InStr(UCase(App.EXEName), "TEPATPRO") > 0 Or _
         InStr(UCase(App.EXEName), "PATPRO1") > 0 Or _
         InStr(UCase(App.EXEName), "TEPATPRO1") > 0 Or _
         InStr(UCase(App.EXEName), "LAW") > 0 Or _
         InStr(UCase(App.EXEName), "TRADEMARK1") > 0 Then
      '2015/11/3 END
          adoEng.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.path & "\eng1.mdb"
          adoEng.Open
      End If
   End If
       ' 90.08.16 modify by louis 產生Word物件
       'StartOfficeAp
       
'Remove by Morgan 2007/1/25 移到下面
'    'Add By Cheng 2004/04/20
'    '記錄使用者所別
'    pub_strUserOffice = PUB_GetST06(strUserNum)
'    'add by nick 2004/10/06
'    Pub_StrUserSt03 = PUB_GetST03(strUserNum)
'End 2007/1/25
'Added by Morgan 2013/5/8
Else
   mdiMain.Timer2.Interval = 100
'end 2013/5/8
End If
   
   Call PUB_GetTMdebate 'Add By Sindy 2013/5/15
   
   'add by nick 2004/12/13  鎖 x 變灰色
   DisableControl mdiMain
   
   'add by nick 2004/08/18 將所要定義的欄位數一次抓齊****start
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from caseprogress where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_CP = AdoRecordSet3.Fields.Count
   'cnnConnection.Execute "insert into nicktest values ('TF_CP','" & TF_CP & "') "
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from patent where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_PA = AdoRecordSet3.Fields.Count
   'cnnConnection.Execute "insert into nicktest values ('TF_PA','" & TF_PA & "') "
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from trademark where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_TM = AdoRecordSet3.Fields.Count
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from lawcase where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   
   TF_LC = AdoRecordSet3.Fields.Count
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from hirecase where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_HC = AdoRecordSet3.Fields.Count
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from servicepractice where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   tf_SP = AdoRecordSet3.Fields.Count
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from nextprogress where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_NP = AdoRecordSet3.Fields.Count
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from acc1k0 where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_1K0 = AdoRecordSet3.Fields.Count
   CheckOC3
   'Add by Morgan 2006/10/23
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from customer where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_CU = AdoRecordSet3.Fields.Count
   CheckOC3
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from fagent where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_FA = AdoRecordSet3.Fields.Count
   CheckOC3
   'end 2006/10/23
   
   'Add By Sindy 2022/9/28
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from ConsultRecordList where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_CRL = AdoRecordSet3.Fields.Count
   CheckOC3
   
   'Add by Morgan 2013/1/22
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select * from NHI2ND where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
   TF_NHI = AdoRecordSet3.Fields.Count
   CheckOC3
   
   'Added by Morgan 2019/3/26
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select OMAN from SetSpecMan where ocode='測試帳號' ", cnnConnection, adOpenStatic, adLockReadOnly
   If Not AdoRecordSet3.EOF Then
      stTestAccount = "" & AdoRecordSet3.Fields(0)
   End If
   CheckOC3
   'end 2019/3/26
   
    '**** end
'Modify By Cheng 2003/07/10
'Begin 以下程式碼往上搬, 同時frmLogin也要放
'strUserNum = strUserNum
''Set cnnConnection = objPublicData.Connection
'Set obj003.Connection = cnnConnection
'Set objLawDll.Connection = cnnConnection
'GetGroupDept
'mdiMain.Show
'strSrvDate(1) = Format(ServerDate)
'strSrvDate(2) = Format(Val(strSrvDate(1)) - 19110000)
'
'Set adoTaie = cnnConnection
'cnnRptConn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\db3.mdb"
'cnnRptConn.Open
'adoTemp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\finance.mdb"
'adoTemp.Open
''Add By Cheng 2003/04/29
'If UCase(App.EXEName) = "PROMOTER" Or UCase(App.EXEName) = "TEPROMOTER" Then
'    adoEng.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.Path & "\eng1.mdb"
'    adoEng.Open
'End If
'   ' 90.08.16 modify by louis 產生Word物件
'   'StartOfficeAp
'End

   'Modified by Morgan 2017/1/4 Win7 UAC提高會有權限錯誤,改寫函數可略過錯誤並發通知信
   'Date = Format(strSrvDate(1), "####/##/##") 'Add by Morgan 2005/1/20 校正日期與DB同步
   'time = Format(ServerTime, "##:##:##")   'Add by Morgan 2005/1/14 校正時間與DB同步
   PUB_SyncClientDateTime True
   'end 2017/14
   
   'Add by Morgan 2005/12/13 加DB電腦名稱
    mdiMain.Caption = mdiMain.Caption & " " & PUB_GetDbTerminal
   'Add by Morgan 2006/3/20
   mdiMain.mnu00(0).Visible = False
   If Pub_StrUserSt03 = "M51" Or InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") <> 0 Or InStr(stTestAccount, strUserNum) > 0 Then
       mdiMain.mnu00(0).Visible = True
   End If
   'Add by Morgan 2006/4/11 讀作業系統代碼 1=95 2=NT 其他=未知
   pub_OS = GetVersion32
   pub_HostName = PUB_ReadHostName 'Add by Morgan 2008/3/27
   PUB_SetAppPrinter 'Add by Morgan 2010/2/3
   PUB_SetSystemVar 'Add by Morgan 2009/2/23
   'Add by Morgan 2009/11/19
   strUser1Num = strUserNum '備份原編號
   PUB_KillTempFile "$$*.*" '清除暫存檔
   'end 2009/11/19
   
   PUB_KillTempFile "$*.*" 'Added by Lydia 2020/02/19 清除暫存檔(frm1105,frm100101_M )
   
   'Add By Sindy 2016/4/12 清除執行檔下員工編號資料夾的暫存檔
   PUB_KillTempFile strUserNum & "\*.*"  '清除暫存檔
   '2016/4/12 END
   
   'Modify By Sindy 2021/5/19 By User 清資料夾 + " & strUserNum & "\
   PUB_KillTempFile "SeminarAttach\" & strUserNum & "\*.*" 'Add By Sindy 2021/2/3 清除暫存檔
   
   'Modified by Morgan 2017/3/20 改呼叫函數(Patpro1使用中且資料夾內有檔案被當附件寄出時,執行Patpro會刪除失敗-無權限)
   ''Add By Sindy 2014/9/26
   'Set fso = CreateObject("Scripting.FileSystemObject") '建立FileSystemObject
   'fso.CreateFolder App.path & "\$$" & strUserNum & "TempFolderForDel" '建立一個虛的$$暫存資料夾
   ''If fso.FolderExists(App.path & "\$$*") = True Then 'FolderExists的用法需要完整資料夾名稱
   '   fso.DeleteFolder App.path & "\$$*", True
   ''End If
   'Set fso = Nothing
   ''2014/9/26 END
   PUB_KillTempFolder "$$*"
   'end 2017/3/20
   
   
   'Add By Sindy 2010/10/1
   Call CheckExeFileDateTime
   
   'Add by Morgan 2010/10/20 專利處考核新規則
   If strSrvDate(1) >= "20101026" Then
      bolNewPromoterRule = True
   End If
   
   'Add by Morgan 2010/12/23 申請號新格式
   If strSrvDate(1) >= "20101224" Then
      bolNewAppNoFormat = True
   End If
   
   'Add by Morgan 2011/8/10
   pub_WinSysPath = PUB_GetWinSysPath
   If Dir(pub_WinSysPath & "ablebatchconverter.exe") <> "" Then
      pub_PdfEnable = True
   End If
   
   PUB_AddAuditLog AL_登入 'Added by Morgan 2025/7/31
End Sub

''Add By Sindy 2010/10/1
''iFile:0=每日,1=每月
'Function WLog(oStrLog As String, Optional iFile As Integer = 0)
'Dim ffa As Integer
'ffa = FreeFile
''Add by Morgan 2008/6/2
'If iFile = 1 Then
'   Open App.Path & "\autobatchlog.log" For Append As ffa
'Else
''end 2008/6/2
'   Open App.Path & "\autobatchdaylog.log" For Append As ffa
'End If
'Print #ffa, Trim(Now) & "  ==>  " & oStrLog
'Close ffa
'End Function

'Removed by Moragn 2014/11/6 移到 basQuery
'Public Sub PUB_SetStaffVar()
'   If strUserNum <> "" Then
'      strExc(0) = "Select ST06,ST03,ST05,ST17,ST15 From Staff Where ST01='" & strUserNum & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'        pub_strUserOffice = "" & RsTemp.Fields("ST06")
'        Pub_StrUserSt03 = "" & RsTemp.Fields("ST03")
'        Pub_StrUserSt17 = "" & RsTemp.Fields("ST17")
'        Pub_StrUserSt15 = "" & RsTemp.Fields("ST15")
'        Pub_strUserST05 = "" & RsTemp.Fields("ST05") 'Add by Lydia 2014/10/31 使用者等級=Pub_GetST05(basQuery)
'      End If
'   End If
'End Sub

'move to basquery by nickc 2007/02/07
''Modify by Morgan 2004/3/16
''主管機關來函畫面進入畫面與接洽紀錄分開
'Public Sub Where01ToGo(ByVal intLeaveKind As Integer, Optional stFormName As String = "frm010001")
'
'    Dim oTmp As Form
'
'    If stFormName = "frm010001_1" Then
'        Set oTmp = frm010001_1
'    Else
'        Set oTmp = frm010001
'    End If
'
'    If intLeaveKind = 1 Then
'       Select Case oTmp.intModifyKind
'                 Case 0
'                            If oTmp.intReceiveKind = 0 Then
'                               oTmp.Show
'                            Else
'                               Set obj001 = Nothing
'                            End If
'                 Case 1, 2
'                            oTmp.Show
'       End Select
'    Else
'       Set obj001 = Nothing
'       Unload oTmp
'    End If
'
'End Sub
Public Sub GetGroupDept()
 Dim TmpCls As ClsSysName, i As Integer
   strExc(0) = "SELECT SG02,SG03 FROM STAFF_GROUP WHERE SG01='" & strGroup & "'"
   intI = 1
   Set RsTemp = ClsLawReadRstMsg(intI, strExc(0)) 'edit by nickc 2007/02/05 不用 dll 了 objLawDll.ReadRstMsg(intI, strExc(0))
   If intI = 1 Then
      For i = 1 To ColSysName.Count
         ColSysName.Remove 1
      Next
      i = 1
      Do While Not RsTemp.EOF
         Set TmpCls = New ClsSysName
         TmpCls.SysId = RsTemp.Fields(0).Value
         TmpCls.SysNam = RsTemp.Fields(1).Value
         ColSysName.add TmpCls, Format(i)
         RsTemp.MoveNext
         i = i + 1
      Loop
   End If
End Sub
'move to basquery by nickc 2007/02/07
'Public Function ChkSysName(ByVal SysName As String) As Boolean
' Dim TmpCls As ClsSysName
'   For Each TmpCls In ColSysName
'      If TmpCls.SysId = SysName Then
'         ChkSysName = True
'         Exit For
'      End If
'   Next
'   If ChkSysName = False Then MsgBox "登入之使用者無法使用此系統類別、或系統類別錯誤，請重新輸入 !", vbCritical
'End Function

'*************************************************
'  清除狀態列內容
'
'*************************************************
Public Sub StatusClear()
   mdiMain.StatusBar1.Panels(1).Text = MsgText(601)
   mdiMain.StatusBar1.Panels(2).Text = MsgText(601)
End Sub

'*************************************************
'  訊息顯示
'
'*************************************************
Public Sub StatusView(strMessage As String)
   mdiMain.StatusBar1.Panels(1).Text = strMessage
End Sub

'*************************************************
'  選單為可使用狀態
'
'*************************************************
Public Sub MenuEnabled()
End Sub

'*************************************************
'  顯示資料筆數
'
'*************************************************
Public Sub CountShow(lngCurrent As Long, lngMax As Long)
   mdiMain.StatusBar1.Panels(2).Text = lngCurrent & MsgText(35) & lngMax
End Sub

''將From移至畫面之中心
'Public Sub MoveFormToCenter(ByRef frmTemp As Form)
'Dim intX  As Integer, intY As Integer
'
'If frmTemp.MDIChild Then
'   intX = (mdiMain.ScaleWidth - frmTemp.Width) / 2
'   intY = (mdiMain.ScaleHeight - frmTemp.Height) / 2
'   'If frmTemp.Height > 6110 Then
'   '   intX = 0
'   '   intY = 0
'   'Else
'   '   intX = (mdiMain.ScaleWidth - frmTemp.Width) / 2
'   '   intY = (mdiMain.ScaleHeight - frmTemp.Height) / 2
'   'End If
'   If mdiMain.Width < 10000 And mdiMain.Height < 7000 Then
'      ' 有垂直捲軸
'      If frmTemp.Height > 6110 Then
'         intY = 0
'         If frmTemp.Width > 9200 Then
'            intX = 0
'         Else
'            intX = (mdiMain.ScaleWidth - frmTemp.Width) / 2
'         End If
'      Else
'         intX = (mdiMain.ScaleWidth - frmTemp.Width) / 2
'         intY = (mdiMain.ScaleHeight - frmTemp.Height) / 2
'      End If
'   End If
'Else
'   intX = (Screen.Width - frmTemp.Width) / 2
'   intY = (Screen.Height - frmTemp.Height) / 2
'End If
'frmTemp.Move intX, intY
'End Sub
''end 2007/5/22

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

'Add by Morgan 2010/2/3
'設定程式預設印表機
Public Sub PUB_SetAppPrinter()
   Dim stSQL As String, iR As Integer, strPrinter As String
   Dim idx As Integer
   Dim adoRst As ADODB.Recordset
   
   stSQL = "select PSP06 from PrintStartPoint where PSP01='" & pub_HostName & "' and PSP02='" & App.EXEName & "' and PSP03='APP'"
   iR = 1
   Set adoRst = ClsLawReadRstMsg(iR, stSQL)
   If iR = 1 Then
      For idx = 0 To Printers.Count - 1
         If Printers(idx).DeviceName = "" & adoRst(0) Then
            Set Printer = Printers(idx)
            strPrinter = Printer.DeviceName
            Exit For
         End If
      Next
   End If
   If strPrinter <> "" Then
      Printer.TrackDefault = False
   Else
      Printer.TrackDefault = True
   End If
   Set adoRst = Nothing
End Sub
