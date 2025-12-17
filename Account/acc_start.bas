Attribute VB_Name = "aacc_start"
'2010/8/18 sonia 日期欄已修改
Option Explicit

Sub Main()

'Added by Morgan 2022/7/15
frmpic002.m_bFixIME = True
frmpic002.Show vbModal
'end 2022/7/15

'Modified by Morgan 2021/8/25 參數改放變數方便後續使用
pub_strCommand = Command()
If pub_strCommand = "" And InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then MsgBox "【" & App.EXEName & "】不可直接執行，請以桌面台一圖示【TeAutoUpd】啟動！", vbCritical: End 'Added by Morgan 2017/10/6

Dim fso As Object 'Add By Sindy 2014/10/24
Dim stTestAccount As String 'Added by Morgan 2019/8/16 測試帳號
   
    '判斷是否在VB下執行, 若非在VB6下執行, 則須先進入登入系統畫面
    pub_str_LoginSucceeded = ""
    If InStr(UCase(Pub_GetModuleFileName), "VB6.EXE") = 0 Then
        'Modified by Morgan 2013/5/8 統一用frmLogin
        'frmLogin_1.Show vbModal
        frmLogin.Show vbModal
        Frmacc0000.Show
        'Add by Amy 2020/03/18 開放電腦中心使用
        If Pub_StrUserSt03 = "M51" Then
            Frmacc0000.Main7_0.Visible = True
        Else
            Frmacc0000.Main7_0.Visible = False
        End If
        Frmacc0000.Timer2.Interval = 100
    Else
        Frmacc0000.Show
        Frmacc0000.Main7_0.Visible = True
        
        'Added by Morgan 2013/2/6
        '設定資料庫變數
        strSql = "begin " + _
             "select st02,st03,st05,st11 into user_data.user_name,user_data.user_department," + _
             "user_data.user_level,user_data.user_group from staff where upper(st01)=" + CNULL(strUserNum) + ";" + _
             "user_data.user_num:=" + CNULL(strUserNum) + ";" + _
             "end;"
         cnnConnection.Execute strSql
         'end 2013/2/6
         PUB_SetStaffVar 'Added by Morgan 2014/11/6
    End If
    
    'add by nickc 2007/02/07
    If UCase(App.EXEName) = "PROMOTER" Or UCase(App.EXEName) = "TEPROMOTER" Or UCase(App.EXEName) = "PATPRO" Or UCase(App.EXEName) = "TEPATPRO" Then
        adoEng.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.path & "\eng1.mdb"
        adoEng.Open
    End If
    
    'add by nick 2004/12/14
    DisableControl Frmacc0000
    'add by nick 2004/08/18 將所要定義的欄位數一次抓齊****start
    CheckOC3
    AdoRecordSet3.CursorLocation = adUseClient
    AdoRecordSet3.Open "select * from caseprogress where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
    TF_CP = AdoRecordSet3.Fields.Count
    CheckOC3
    AdoRecordSet3.CursorLocation = adUseClient
    AdoRecordSet3.Open "select * from patent where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
    TF_PA = AdoRecordSet3.Fields.Count
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
    '**** end
    
    'Add by Morgan 2013/4/1
    AdoRecordSet3.CursorLocation = adUseClient
    AdoRecordSet3.Open "select * from NHI2ND where rownum<2 ", cnnConnection, adOpenStatic, adLockReadOnly
    TF_NHI = AdoRecordSet3.Fields.Count
    CheckOC3
    'end 2013/4/1
    
   'Added by Morgan 2019/3/26
   AdoRecordSet3.CursorLocation = adUseClient
   AdoRecordSet3.Open "select OMAN from SetSpecMan where ocode='測試帳號' ", cnnConnection, adOpenStatic, adLockReadOnly
   If Not AdoRecordSet3.EOF Then
      stTestAccount = "" & AdoRecordSet3.Fields(0)
   End If
   CheckOC3
   If InStr(stTestAccount, strUserNum) > 0 Then
      Frmacc0000.Main7_0.Visible = True
   End If
   'end 2019/3/26
    
    
    'add by nickc 2005/11/10
    'edit by nickc 2007/02/07 不用 dll 了
    'Set objLawDll.Connection = cnnConnection
    
'Removed by Morgan 2014/11/6 改前面呼叫 PUB_SetStaffVar 設定
'    'Morgan by Morgan 2007/1/29 用1句以減少連資料庫次數
'    ''add by nickc 2005/05/24
'    'Pub_StrUserSt03 = PUB_GetST03(strUserNum)
'    strExc(0) = "Select ST06,ST03,ST17,ST15 From Staff Where ST01='" & strUserNum & "'"
'    intI = 1
'    'edit by nickc 2007/02/07 不用 dll 了
'    'Set RsTemp = objLawDll.ReadRstMsg(intI, strExc(0))
'    Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'    If intI = 1 Then
'      pub_strUserOffice = "" & RsTemp.Fields("ST06")
'      Pub_StrUserSt03 = "" & RsTemp.Fields("ST03")
'      Pub_StrUserSt17 = "" & RsTemp.Fields("ST17")
'      Pub_StrUserSt15 = "" & RsTemp.Fields("ST15")
'    End If
'    'End 2007/1/29
'end 2014/11/6
    
    'Add by Morgan 2005/12/13 加DB電腦名稱
    Frmacc0000.Caption = Frmacc0000.Caption & " " & PUB_GetDbTerminal
   'Add by Morgan 2006/4/11 讀作業系統代碼 1=95 2=NT 其他=未知
   pub_OS = GetVersion32
   pub_HostName = PUB_ReadHostName 'Add by Morgan 2008/3/27
   PUB_SetSystemVar 'Add by Morgan 2009/2/23
   
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
   
   'Add by Morgan 2012/11/6
   pub_WinSysPath = PUB_GetWinSysPath
   If Dir(pub_WinSysPath & "ablebatchconverter.exe") <> "" Then
      pub_PdfEnable = True
   End If
   
   'Add By Sindy 2014/10/24
   PUB_KillTempFile "$$*.*" '清除暫存檔
   'Modified by Morgan 2021/6/18
   'Set fso = CreateObject("Scripting.FileSystemObject") '建立FileSystemObject
   'fso.CreateFolder App.path & "\$$" & strUserNum & "TempFolderForDel" '建立一個虛的$$暫存資料夾
   ''If fso.FolderExists(App.path & "\$$*") = True Then 'FolderExists的用法需要完整資料夾名稱
   '   fso.DeleteFolder App.path & "\$$*", True
   ''End If
   'Set fso = Nothing
   PUB_KillTempFolder "$$*"
   'end 2021/6/18
   '2014/10/24 END
   
   PUB_AddAuditLog AL_登入 'Added by Morgan 2025/7/31
End Sub

'Modify By Sindy 2013/9/23 Mark
''Add By Sindy 2010/10/1
''iFile:0=每日,1=每月
'Function WLog(oStrLog As String, Optional iFile As Integer = 0)
'Dim ffa As Integer
'ffa = FreeFile
''Add by Morgan 2008/6/2
'If iFile = 1 Then
'   Open App.path & "\autobatchlog.log" For Append As ffa
'Else
''end 2008/6/2
'   Open App.path & "\autobatchdaylog.log" For Append As ffa
'End If
'Print #ffa, Trim(Now) & "  ==>  " & oStrLog
'Close ffa
'End Function

'Modify by Morgan 2011/3/21 因為 Account 及 Casher 要共用所以從 aacc_fun 移來
'Modify By Sindy 2013/12/30 +bolNotInJ
Public Function PUB_AddItem2CboTitle(cboTitle As Object, p_CustNo1 As String, p_CustNo2 As String, p_Year As String, _
                                     Optional bolNotInJ As Boolean = False) As Boolean
   Dim strSql As String, strCon1 As String, strCon2 As String
   Dim adoquery As ADODB.Recordset, iRtn As Integer
   Dim strItem As String
   
On Error GoTo ErrHand

   strCon1 = ""
   If p_Year <> "" Then
      strCon1 = strCon1 & " and a0k16=" & p_Year
   End If
   If p_CustNo1 <> "" Then
      strCon1 = strCon1 & " and a0k03>='" & p_CustNo1 & "'"
   End If
   If p_CustNo2 <> "" Then
      strCon1 = strCon1 & " and a0k03<='" & p_CustNo2 & "'"
   End If
   'Modify By Sindy 2013/12/30 不含J公司
   If bolNotInJ = True Then
      strCon1 = strCon1 & " and a0k11<>'J'"
   End If
   '2013/12/30 END
   
   'Add By Sindy 2015/10/20
   strCon2 = ""
   If p_Year <> "" Then
      strCon2 = strCon2 & " and A1V09=" & p_Year
   End If
   If p_CustNo1 <> "" Then
      strCon2 = strCon2 & " and decode(a0y18,1,a0y07,2,a0y08,a0y09)>='" & p_CustNo1 & "'"
   End If
   If p_CustNo2 <> "" Then
      strCon2 = strCon2 & " and decode(a0y18,1,a0y07,2,a0y08,a0y09)<='" & p_CustNo2 & "'"
   End If
   '2015/10/20 END
   
   'Modify by Morgan 2011/3/10 排除手開收據
   '2011/10/20 MODIFY BY SONIA E10023515
   'strSql = "Select distinct rpad(a0k04, 60,' ') C01, a0k03 C02" & _
      " from Acc0k0 where substrb(a0k01,-5)>'02000' and a0k04 like '" & cboTitle.Text & "%'" & strCon1 & strCon2 & _
      " order by 1,2"
   '2012/5/14 MODIFY BY SONIA 國立中正大學101年都是2000號以下,辜說要拿掉2000號的限制
   'strSql = "Select distinct rpad(a0k04, 60,' ') C01, a0k03 C02" & _
      " from Acc0k0 where substrb(a0k01,-5)>'02000' and instr(upper(a0k04),upper('" & cboTitle.Text & "'))>0" & strCon1 & strCon2 & _
      " order by 1,2"
   'Modify By Sindy 2024/9/30 + ChgSQL
   strSql = "Select distinct rpad(a0k04, 60,' ') C01, a0k03 C02" & _
            " from Acc0k0 where instr(upper(a0k04),upper('" & ChgSQL(cboTitle.Text) & "'))>0" & strCon1
   'Add By Sindy 2015/10/20 +acc1k0
   'Modify By Sindy 2024/9/30 + ChgSQL
   strSql = strSql & " union Select distinct rpad(a1k35, 60,' ') C01, decode(a0y18,1,a0y07,2,a0y08,a0y09) C02" & _
                     " From Acc1k0, acc1v0, acc0z0, acc0y0" & _
                     " where a1k01=a1v02(+) and a1k01=a0z02(+) and a0z01=a0y01(+)" & _
                     " and instr(upper(a1k35),upper('" & ChgSQL(cboTitle.Text) & "'))>0" & strCon2
   '2015/10/20 END
   strSql = strSql & " order by 1,2"
   iRtn = 1
   Set adoquery = ClsLawReadRstMsg(iRtn, strSql)
   If iRtn = 1 Then
      strItem = cboTitle.Text
      cboTitle.Clear
      cboTitle.AddItem strItem
      Do While Not adoquery.EOF
         strItem = "" & adoquery.Fields(0) & " " & adoquery.Fields(1)
         cboTitle.AddItem strItem
         adoquery.MoveNext
      Loop
      cboTitle.ListIndex = 0
      
      'Add by Morgan 2008/2/20
      'SendMessage cboTitle.hWnd, CB_SHOWDROPDOWN, 1, 0 'Modify By Sindy 2021/8/5 Mark
   End If
   
   Set adoquery = Nothing
   PUB_AddItem2CboTitle = True
   Exit Function
   
ErrHand:
   MsgBox Err.Description
End Function

'Copy from basStart by Morgan 2013/5/8
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

'Removed by Moragn 2014/11/6 移到 basQuery
''Copy from basStart by Morgan 2013/5/8
'Public Sub PUB_SetStaffVar()
'   If strUserNum <> "" Then
'      strExc(0) = "Select ST06,ST03,ST17,ST15 From Staff Where ST01='" & strUserNum & "'"
'      intI = 1
'      Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
'      If intI = 1 Then
'        pub_strUserOffice = "" & RsTemp.Fields("ST06")
'        Pub_StrUserSt03 = "" & RsTemp.Fields("ST03")
'        Pub_StrUserSt17 = "" & RsTemp.Fields("ST17")
'        Pub_StrUserSt15 = "" & RsTemp.Fields("ST15")
'      End If
'   End If
'End Sub


'Added by Morgan 2025/5/22
'檢查收據/收文號是否繳款中
'pNo:單號, pType=1:收據號,2:收文號
Public Function PUB_Chk440(pNo As String, pType As String) As Boolean
   Dim stSQL As String, intQ As Integer, stCon As String
   Dim rsQuery As ADODB.Recordset
   
   If pType = "1" Then
      stCon = " and axd04='" & pNo & "'"
   ElseIf pType = "2" Then
      stCon = " and axd05='" & pNo & "'"
   Else
      Exit Function
   End If
   
   stSQL = "select a4416 from acc441,acc440 where a4401(+)=axd01 and a4402(+)=axd02 and a4403(+)=axd03" & stCon
   intQ = 1
   Set rsQuery = ClsLawReadRstMsg(intQ, stSQL)
   If intQ = 1 Then
      If Left("" & rsQuery(0), 1) <> "F" Then
         MsgBox "智權人員已繳款，請通知智權刪除繳款後再續行後續處理！", vbCritical
         PUB_Chk440 = True
      End If
   End If
      
   Set rsQuery = Nothing
End Function
'Added by Morgan 2025/6/18
'支票票期規定的最大日期=收票日期+2個月的最後1天 Ex:6月收票，票期可以到8/31
Public Function PUB_GetCheckMaxDate(pRecDate As String) As String
   PUB_GetCheckMaxDate = CompDate(2, -1, CompDate(1, 3, Left(DBDATE(pRecDate), 6) & "01"))
End Function
