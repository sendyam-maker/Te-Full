Attribute VB_Name = "basLaw"
'Memo By Sindy 2012/12/5 智權人員欄已修改
'Memo By Sindy 2011/2/15 SQLDate已檢查
'Memo By Sindy 2010/8/4 日期欄已修改
Option Explicit


'Siegfried 建立 Temp 之資料庫
Public Function CreateDatabase() As Boolean
 Dim Wo As DAO.Workspace, Db As DAO.Database, Td As DAO.TableDef, St As String
 Dim i As Integer, Fd As DAO.field
On Error GoTo ErrHand
   CreateDatabase = False
   Set Wo = DBEngine.Workspaces(0)
   St = App.path & "\Case.mdb"
   If Dir(St) <> "Case.mdb" Then
      Set Db = Wo.CreateDatabase(St, dbLangChineseTraditional & ";pwd=taie", dbVersion30)
      Set Td = Db.CreateTableDef("TEMP")
      For i = 1 To 17
         Set Fd = Td.CreateField("TMP" & Format(i, "00"), dbText, 50)
         Td.Fields.Append Fd
      Next
      Db.TableDefs.Append Td
      Db.Close
      Wo.Close
      CreateDatabase = True
   Else
      Set Db = Wo.OpenDatabase(St, False, False, ";PWD=taie")
      For Each Td In Db.TableDefs
         If Td.Name = "TEMP" Then
            CreateDatabase = True
            Exit For
         End If
      Next
      If CreateDatabase = False Then
         Set Td = Db.CreateTableDef("TEMP")
         For i = 1 To 17
            Set Fd = Td.CreateField("TMP" & Format(i, "00"), dbText, 50)
            Fd.AllowZeroLength = True
            Td.Fields.Append Fd
         Next
         Db.TableDefs.Append Td
      End If
      Db.Close
      Wo.Close
      CreateDatabase = True
   End If
   Exit Function
ErrHand:
   CreateDatabase = False
End Function

'Add by Morgan 2011/6/23
'列印請款單
Public Sub PUB_PrintBill(pBillNo As String, pPrinter As String, pbEmail As Boolean, pbPlusPaper As Boolean, Optional pName As String, _
                         Optional ByRef pPageCount As Integer, Optional ByVal pCopies As Integer = 0, Optional ByVal pOutMode As String = "1")
   Dim arrBillNo() As String
   Dim ii As Integer
   
   If pBillNo = "" Then Exit Sub
   
   pPageCount = 0
   arrBillNo = Split(pBillNo, ",")
   Load Frmacc2480
   With Frmacc2480
      .Combo1.Text = pPrinter
      For ii = LBound(arrBillNo) To UBound(arrBillNo)
         If arrBillNo(ii) <> "" Then
            .m_iPageCount = 0
            .Text1.Text = arrBillNo(ii)
            .Text2.Text = arrBillNo(ii)
            .m_bBeCalled = True
            .m_CallPrevForm = "PUB_PrintBill"  'Added by Lydia 2020/01/06 呼叫請款單的程式名稱
            .m_bAddDate = True
            .m_bEMail = pbEmail
            .m_bPaper = pbPlusPaper
            .m_iCopies = pCopies
            .txtOutMode = pOutMode 'Add By Sindy 2015/7/9 1.印表機 2.電子檔
            .Command2_Click
            pPageCount = pPageCount + .m_iPageCount
         End If
      Next
      .Combo1.Text = .Combo1.Tag
   End With
   Unload Frmacc2480
   'strFormName = pName 'Removed by Morgan 2015/10/27 呼叫的都不是財務的 Form 不可設定,否則共同程序會被鎖住
End Sub

'Add By Cheng 2003/02/24
'列印DebitNote
'Modified by Lydia 2019/10/01 傳入程式名稱strFrmName
Public Sub PUB_PrintDebitNote(strDNL01 As String, strPrinterName As String, Optional ByVal strFrmName As String = "")
'strDNL01 : 使用者代號
Dim StrSQLa As String
Dim rsA As New ADODB.Recordset
Dim tmpErr As String

    'Modified by Lydia 2019/04/02 PK: 使用者帳號@電腦名稱(pub_HostName)
    'StrSQLa = "Select * From DebitNoteList Where DNL01='" & strDNL01 & "' Order By DNL03 "
    StrSQLa = "Select * From DebitNoteList Where DNL01='" & strDNL01 & "@" & pub_HostName & "' Order By DNL03 "
    rsA.CursorLocation = adUseClient
    rsA.Open StrSQLa, cnnConnection, adOpenStatic, adLockReadOnly
    '若有DebitNote資料
    If rsA.RecordCount > 0 Then
        'Removed by Lydia 2019/10/01 外專請款單的信頭已改成直接列印(黑色)
        'If MsgBox("準備列印請款單，請更換紙張!!!", vbExclamation + vbOKCancel) = vbOK Then
RePrint:
            Load Frmacc2480
            '移至第一筆資料
            rsA.MoveFirst
            With Frmacc2480
               .Combo1.Text = strPrinterName
               .m_bBeCalled = True
               'Added by Lydia 2020/07/17 改成來源表單
               If strFrmName <> "" Then
                    .m_CallPrevForm = strFrmName
               Else
               'end 2020/07/17
                    .m_CallPrevForm = "PUB_PrintDebitNote" 'Added by Lydia 2020/01/06 呼叫請款單的程式名稱
               End If
               While Not rsA.EOF
                  .Text1.Text = "" & rsA("DNL02").Value
                  .Text2.Text = "" & rsA("DNL02").Value
                  'Add by Morgan 2008/4/8 +是否產生電子檔
                  .m_bEMail = False
                  .m_bPaper = False
                  If "" & rsA("DNL04").Value = "Y" Then
                     .m_bEMail = True
                     'Add by Morgan 2009/10/19
                     If "" & rsA("DNL05").Value = "Y" Then
                        .m_bPaper = True
                     End If
                  End If
                  'end 2008/4/8
                  'Added by Lydia 2020/06/23  已付款
                  .m_bPAID = False
                  If "" & rsA("DNL06").Value = "Y" Then
                     .m_bPAID = True
                  End If
                  'end 2020/06/23
                  .Command2_Click: DoEvents
                  'Added by Lydia 2020/09/10 請款單：判斷PDF檔案是否存在
                  If .m_strOutErr <> "" Then
                       tmpErr = tmpErr & .m_strOutErr
                  End If
                  'end 2020/09/10
                  rsA.MoveNext
               Wend
            End With
            Unload Frmacc2480
            'Add By Cheng 2003/02/12
            '可重覆列印DebitNote
            If strFrmName <> "frm060307" Then 'Added by Lydia 2019/10/01 排除年證費請款函(發文自動請款)
                If MsgBox("請款單已列印完畢，您是否要重新列印???", vbExclamation + vbYesNo + vbDefaultButton2) = vbYes Then
                    GoTo RePrint
                End If
            End If 'Added  2019/10/01
            
            'Added by Lydia 2020/09/10
            If tmpErr <> "" Then
                PUB_SendMail strUserNum, strUserNum, "", "請款單電子檔產生失敗", "請款單電子檔產生失敗：" & vbCrLf & Replace(tmpErr, "＆", vbCrLf) & vbCrLf
                MsgBox "請款單電子檔產生失敗：" & vbCrLf & Replace(tmpErr, "＆", vbCrLf) & vbCrLf & "請參考！", vbInformation, strFrmName & "-電子檔產生失敗"
            End If
            'end 2020/09/10
            
        'End If 'Removed 2019/10/01
    End If
    If rsA.State <> adStateClosed Then rsA.Close
    Set rsA = Nothing
End Sub


'*************************************************
'  功能鍵對照函式
'
'*************************************************
Public Sub KeyEnter(InputCode As Integer)
Dim Cancel As Boolean 'Add By Sindy 2009/07/15
Dim bolUnload As Boolean 'Added by Morgan 2016/7/28
'Added by Lydia 2018/12/03
Dim rsA As New ADODB.Recordset, intA As Integer, strTmp As String
Dim stRtn As String
'end 2018/12/03
Dim stCon1 As String, stCon2 As String 'Added by Lydia 2019/01/07
Dim strCP09 As String 'Add By Sindy 2025/10/22

   mdiMain.StatusBar1.Panels(1).Text = MsgText(601)
   Select Case InputCode
      Case vbKeyEscape
         Select Case strFormName
            'Add by Amy 2015/09/03
            Case "frm100114_5"
                Unload frm100114_5
                tool13_enabled
                Exit Sub
            Case "Frmacc2152"
               strCon9 = ""
               strCon10 = ""
               Unload Frmacc2152
               tool2_enabled
               Exit Sub
            Case "Frmacc2162"
               strCon9 = ""
               Unload Frmacc2162
               tool2_enabled
               Exit Sub
         End Select
         If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
            Exit Sub
         End If
         tool4_enabled
         If strFormName = MsgText(601) Then
            Exit Sub
         End If
         Select Case strFormName
            Case "Frmacc2150"
               Unload Frmacc2150
            Case "Frmacc2151"
               Unload Frmacc2151
            Case "Frmacc2153"
               Unload Frmacc2153
            Case "Frmacc2160"
               Unload Frmacc2160
            Case "Frmacc2161"
               Unload Frmacc2161
            Case "Frmacc21g0"
               Unload Frmacc21g0
            Case "Frmacc21g1"
               Unload Frmacc21g1
            Case "Frmacc21h0"
               Unload Frmacc21h0
            Case "Frmacc21h1"
               Unload Frmacc21h1
            Case "Frmacc21h2"
               Unload Frmacc21h2
            Case "Frmacc21i0"
               Unload Frmacc21i0
            Case "Frmacc21i1"
               Unload Frmacc21i1
            Case "Frmacc21j0"
               Unload Frmacc21j0
            Case "Frmacc21k0"
               Unload Frmacc21k0
            Case "Frmacc21m0"
               Unload Frmacc21m0
            Case "Frmacc21o0"
               Unload Frmacc21o0
            Case "Frmacc21o1"
               Unload Frmacc21o1
            Case "Frmacc21p0"
               Unload Frmacc21p0
            'Removed by Morgan 2018/4/27
            'Case "Frmacc21p1"
            '   Unload Frmacc21p1
            'Add By Sindy 2009/06/06
            Case "Frmacc21s0"
               Unload Frmacc21s0
            '2009/06/06 End
            Case "Frmacc2210"
               Unload Frmacc2210
            Case "Frmacc2211"
               Unload Frmacc2211
            Case "Frmacc2212"
               Unload Frmacc2212
            Case "Frmacc2213"
               Unload Frmacc2213
            Case "Frmacc2214"
               Unload Frmacc2214
            Case "Frmacc2215"
               Unload Frmacc2215
            '2011/10/31 ADD BY SONIA
            Case "Frmacc2216"
               Unload Frmacc2216
            '2011/10/31 END
            Case "Frmacc2220"
               Unload Frmacc2220
            Case "Frmacc2230"
               Unload Frmacc2230
            Case "Frmacc2240"
               Unload Frmacc2240
            Case "Frmacc2470"
               Unload Frmacc2470
            Case "Frmacc2480"
               Unload Frmacc2480
            Case "Frmacc24b0"
               Unload Frmacc24b0
            Case "Frmacc24c0"
               Unload Frmacc24c0
            Case "Frmacc24f0"
               Unload Frmacc24f0
            Case "Frmacc24g0"
               Unload Frmacc24g0
            Case "Frmacc24h0"
               Unload Frmacc24h0
            'Add By Cheng 2002/09/05
            Case "Frmacc24i0"
               Unload Frmacc24i0
            Case Else 'Add by Morgan 2010/4/1
               Dim oForm
               For Each oForm In Forms
                  If oForm.Name = strFormName Then Unload oForm: Exit For
               Next
         End Select
      '新增
      Case vbKeyF2
         If mdiMain.Toolbar1.Buttons.Item(4).Enabled = False Then
            Exit Sub
         End If
         If strSaveConfirm = MsgText(4) Then
            Exit Sub
         End If
         strSaveConfirm = MsgText(3)
         Select Case strFormName
            Case "Frmacc2150"
               If CheckUse("Frmacc2150", strAdd) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               
               'Added by Morgan 2016/6/30
               'Modified by Morgan 2018/10/18 CFP電子化,專利處人員輸入都要先匯入電子檔
               'If 內專全面電子化啟用日 <= Val(strSrvDate(1)) And (PUB_GetST05(strUserNum) = "73" Or PUB_GetST05(strUserNum) = "75") Then
               'Modify By Sindy 2021/1/19 + Or Left(Pub_StrUserSt03, 2) = "P2" 內商電子化
               'Modified by Morgan 2022/7/7 因德國年費須在沒有帳單檔案的狀況下輸入(財務處直接繳費)，故取消管控--玫音,郭
               'If Left(Pub_StrUserSt03, 2) = "P1" Or Left(Pub_StrUserSt03, 2) = "P2" Then
               'Modified by Morgan 2023/4/19 +F11,CFT帳單電子化
               If Left(Pub_StrUserSt03, 2) = "P2" Or Pub_StrUserSt03 = "F11" Then
               'end 2018/10/18
                  If Frmacc2150.m_eFileName = "" Then
                     'Modified by Morgan 2016/8/1 彥葶會輸CFP帳單改提醒可繼續
                     'Modified by Morgan 2018/10/18 CFP電子化,專利處程序輸入都要先匯入電子檔
                     'If MsgBox("配合帳單電子化，P案請由[帳單輸入-整批]新增帳單！若非P案點選確定後可繼續。", vbOKCancel + vbDefaultButton2 + vbExclamation) = vbCancel Then
                     MsgBox "配合帳單電子化，請改由【帳單輸入-整批】新增帳單！", vbCritical
                     'end 2018/10/18
                        strSaveConfirm = MsgText(601)
                        Exit Sub
                     'End If
                  End If
               End If
               'end 2016/6/30
               
               Frmacc2150_Clear
               With Frmacc2150
                   'Ken 92/01/03 改為編號依系統年度編號
'                  If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
'                     .MaskEdBox1.Text = CFDate(ACDate(ServerDate))
                     .Text2 = AutoNo(MsgText(812), 5)
'                  Else
'                     If Mid(.MaskEdBox1.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
'                        .Text2 = UpdateNo("acc150", "a1501", 5, .MaskEdBox1.Text, MsgText(812))
'                     Else
'                        .Text2 = AutoNo(MsgText(812), 5)
'                     End If
'                  End If
'                  .strDocNo = .Text2
                  
                  'Added by Sindy 2018/2/22
                  If .m_strIR01 <> "" Then
                     If Val(.m_RDate) > 0 Then
                        .MaskEdBox1.Text = CFDate(Val(.m_RDate))
                        .MaskEdBox1.Mask = DFormat
                     End If
                  End If
                  '2018/2/22 END
                  .FormEnabled
               End With
               adoTaie.BeginTrans
            Case "Frmacc2160"
               If CheckUse("Frmacc2160", strAdd) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               Frmacc2160_Clear
               With Frmacc2160
                  If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
                     .MaskEdBox1.Text = CFDate(ACDate(ServerDate))
                     .Text2 = AutoNo(MsgText(813), 5)
                  Else
                     If Mid(.MaskEdBox1.Text, 1, 3) <> Mid(CFDate(ACDate(ServerDate)), 1, 3) Then
                        .Text2 = UpdateNo("acc160", "a1601", 5, .MaskEdBox1.Text, MsgText(813))
                     Else
                        .Text2 = AutoNo(MsgText(813), 5)
                     End If
                  End If
                  .strDocNo = .Text2
                  
                  'Added by Sindy 2018/2/22
                  If .m_strIR01 <> "" Then
                     If Val(.m_RDate) > 0 Then
                        .MaskEdBox1.Text = CFDate(Val(.m_RDate))
                        .MaskEdBox1.Mask = DFormat
                     End If
                  End If
                  '2018/2/22 END
                  
                  .FormEnabled
               End With
               adoTaie.BeginTrans
            Case "Frmacc21k0"
               If CheckUse("Frmacc21k0", strAdd) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               Frmacc21k0_Clear
            Case "Frmacc21g0"
               If CheckUse("Frmacc21g0", strAdd) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               Frmacc21g0_Clear Frmacc21g0
            Case "Frmacc21h0"
               If CheckUse("Frmacc21h0", strAdd) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               Frmacc21h0_Clear
               With Frmacc21h0
                  'Added by Morgan 2014/6/6
                  '使用預留單號
                  If .Check1.Value = 1 Then
                     If .AddCheck = False Then
                        strSaveConfirm = MsgText(601)
                        Exit Sub
                      End If
                     .Text5 = .Text11
                  Else
                  'end 2014/6/6
                     adoTaie.BeginTrans
                     adoTaie.Execute "update acc1r0 set a1r04 = a1r04 where a1r01 = 'X'"
                     .Text5 = AccAutoNo(MsgText(815), 5)
                     strConTitle = AccSaveAutoNo(MsgText(815), Right(.Text5, 5))
                     adoTaie.CommitTrans
                  End If 'Added by Morgan 2014/6/6
                  .Command1.Enabled = True
                  .Command2.Enabled = False
                  .Command3.Enabled = True
                  .Command4.Enabled = False 'Add by Morgan 2010/5/21
                  .Command5.Enabled = False
                  .Command6.Enabled = False 'Add by Amy 2014/06/26 避免新增時再按此鈕產錯誤
                  .Text5.Enabled = False 'Added by Morgan 2012/9/28
               End With
               adoTaie.BeginTrans
            Case "Frmacc21j0"
               If CheckUse("Frmacc21j0", strAdd) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               Frmacc21j0_Clear
            Case "Frmacc21m0"
               If CheckUse("Frmacc21m0", strAdd) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               Frmacc21m0_Clear Frmacc21m0
            Case "Frmacc21o0"
               If CheckUse("Frmacc21o0", strAdd) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               Frmacc21o0_Clear Frmacc21o0
            'Add By Sindy 2009/06/06
            Case "Frmacc21s0"
               If CheckUse("Frmacc21s0", strAdd) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               Frmacc21s0_Clear Frmacc21s0
            '2009/06/06 End
         End Select
         tool2_enabled
      '修改
      Case vbKeyF3
         If mdiMain.Toolbar1.Buttons.Item(5).Enabled = False Then
            Exit Sub
         End If
         If strSaveConfirm = MsgText(3) Then
            Exit Sub
         End If
         Select Case strFormName
            Case "Frmacc2150"
               If CheckUse("Frmacc2150", strEdit) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               With Frmacc2150
                  .FormEnabled
               End With
               Frmacc2150.Command3.Value = True 'Added by Morgan 2024/12/27 重讀資料,因為瀏覽狀態畫面資料也可能被變動
               adoTaie.BeginTrans
            Case "Frmacc2160"
               If CheckUse("Frmacc2160", strEdit) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               With Frmacc2160
                  .FormEnabled
               End With
               adoTaie.BeginTrans
            Case "Frmacc21h0"
               If CheckUse("Frmacc21h0", strEdit) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               With Frmacc21h0
                  'Added by Morgan 2015/8/28
                  If .Text5.Tag = "" Then
                     MsgBox "請先輸入已存在的請款單號並查詢!!", vbInformation
                     .Text5.SetFocus
                     Exit Sub
                  Else
                     .Text5 = .Text5.Tag
                  End If
                  'end 2015/8/28
                  'Add By Sindy 2009/07/15
                  Cancel = False
                  Call Frmacc21h0.Text1_Validate(Cancel)
                  If Cancel = True Then Exit Sub
                  '2009/07/15 End
                  
                  If .ModifyCheck = False Then Exit Sub 'Added by Morgan 2017/1/6 1k0的修改檢查應與1l0相同
                  
                  .Command1.Enabled = True
                  .Command2.Enabled = False
                  .Command3.Enabled = True
                  .Command4.Enabled = False 'Add by Morgan 2010/5/21
                  .Command5.Enabled = False
                  .Command6.Enabled = False 'Add by Amy 2014/06/26 避免新增時再按此鈕產錯誤
                  .Text5.Enabled = False 'Added by Morgan 2012/9/28
               End With
               adoTaie.BeginTrans
               
            'Add By Sindy 2009/07/15
            Case "Frmacc21i0"
               If CheckUse("Frmacc21i0", strEdit) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               With Frmacc21i0
                  Cancel = False
                  Call Frmacc21i0.Text7_Validate(Cancel)
                  If Cancel = True Then Exit Sub
               End With
               'adoTaie.BeginTrans
               
            'Add By Sindy 2009/07/15
            Case "Frmacc21k0"
               If CheckUse("Frmacc21k0", strEdit) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               With Frmacc21k0
                  Cancel = False
                  Call Frmacc21k0.Text1_Validate(Cancel)
                  If Cancel = True Then Exit Sub
               End With
               'adoTaie.BeginTrans
               
'            Case "Frmacc21h1"
'               With Frmacc21h1
'                  .Text6 = .Text2
'                  .Text8 = .Text2
'                  .FormEnabled
'                  .Text15 = ZeroBeforeNo(0, 3)
'               End With
'               adoTaie.BeginTrans

            'Added by Morgan 2019/7/10
            Case "Frmacc21o0"
               If CheckUse("Frmacc21o0", strEdit) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
            'end 2019/7/10
            
             Case "Frmacc21t0" 'Add by Morgan 2010/11/19
               Frmacc21t0.SetEdit
         End Select
         strSaveConfirm = MsgText(4)
         tool2_enabled
         
         'Added by Morgan 2015/8/28
         If strFormName = "Frmacc21h0" Then
            Frmacc21h0.AdodcRefresh
         End If
         'end 2015/8/28
      '存檔
      Case vbKeyF9
         If strSaveConfirm = MsgText(601) Then
            Exit Sub
         End If
         Err.Clear
         Select Case strFormName
            Case "Frmacc2150"
               With Frmacc2150
'                  If Val(.Text14) <> Val(.Text6) Then
'                     MsgBox MsgText(59), , MsgText(5)
'                     Exit Sub
'                  End If
'                  If strSaveConfirm = MsgText(4) Then
                     Frmacc2150_Save
'                  End If
                  If strControlButton <> MsgText(602) Then
                     .FormDisabled
                  End If
                  If Not .m_ParentForm Is Nothing Then bolUnload = True 'Added by Morgan 2016/7/28
               End With
               If strControlButton <> MsgText(602) Then
                  adoTaie.CommitTrans
                  'Added by Lydia 2017/05/17 新增帳單檢查翻譯費的完稿字數是否超過原來預估
                  If strSaveConfirm = MsgText(3) Then
                      Call Frmacc2150.ChkMailTransFee
                  End If
                  'end 2017/05/17
               End If
            Case "Frmacc2160"
               With Frmacc2160
'2009/6/15 cancel by sonia
'                  If Val(.Text8) <> Val(.Text7) Then
'                     MsgBox MsgText(59), , MsgText(5)
'                     Exit Sub
'                  End If
'                  If strSaveConfirm = MsgText(4) Then
                     Frmacc2160_Save
'                  End If
                  If strControlButton <> MsgText(602) Then
                     .FormDisabled
                  End If
               End With
               If strControlButton <> MsgText(602) Then
                  adoTaie.CommitTrans
               End If
               If Not Frmacc2160.m_PrevForm Is Nothing Then bolUnload = True  'Add by Sindy 2018/2/23
            Case "Frmacc21g0"
               Frmacc21g0_Save Frmacc21g0
            Case "Frmacc21h0"
               Frmacc21h0_Save
               'Modify by Morgan 2010/3/2 要控制沒有錯誤才可結束
               'adoTaie.CommitTrans
               If strControlButton <> MsgText(602) Then
                  With Frmacc21h0
                  If .Check1.Value = 1 And .Text5 = .Text11 Then PUB_UpdRsvDN .Text11 'Added by Morgan 2023/6/29
                  adoTaie.CommitTrans
                  
                  .Command1.Enabled = False
                  .Command3.Enabled = False
                  .Command5.Enabled = True
                  .Command6.Enabled = True 'Add by Amy 2014/06/26 避免新增時再按此鈕產錯誤
                  If Len(.Text5) = 10 Then
                     .Command2.Enabled = False
                     .Command4.Enabled = False 'Add by Morgan 2010/5/21
                  Else
                     .Command2.Enabled = True
                     .Command4.Enabled = True 'Add by Morgan 2010/5/21
                  End If
                  .Text5.Enabled = True 'Added by Morgan 2012/9/28
                  'Added by Morgan 2014/6/9
                  '使用預留單號時自動跳下一號
                  If .Check1.Value = 1 Then
                     If .Text5 = .Text11 Then
                     
                        If .Text11 < .Text10 Then
                           .Text11 = Left(.Text11, 1) & (Val(Mid(.Text11, 2)) + 1)
                        Else
                           .Text11 = ""
                        End If
                     End If
                  End If
                  'end 2014/6/9
                  
                  'Added by Lydia 2018/12/03 程序人員(F22)在"A"類請款時，去檢查電子送件暫存區有無相同案件的資料匣，若有則自行刪除資料匣，若無法刪除整個資料匣請發mail給程序管制人員，其mail內容同之前之設定(外專發文)
On Error GoTo ErrHandle
                  If Pub_StrUserSt03 = "F22" And .Text1 = "FCP" And .Text6 <> "" Then
                      strTmp = "select * from caseprogress where cp01='" & .Text1 & "' and cp02='" & .Text6 & "' and cp03='" & .Text7 & "' and cp04='" & .Text8 & "'  and cp09 like 'A%' and cp60='" & .Text5 & "' "
                      intA = 1
                      Set rsA = ClsLawReadRstMsg(intA, strTmp)
                      If intA = 1 Then
                          strTmp = PUB_GetFCPHandler(.Text1, .Text6, .Text7, .Text8)
                          If strTmp <> "" Then
                            'Modified by Lydia 2024/07/22 改用變數
'                            If Dir("\\Typing2\電子送件暫存區\" & .Text1 & .Text6, vbDirectory) <> "" Then
'                                stRtn = convForm(strTmp, 6) & "無法刪除\\Typing2\電子送件暫存區\" & .Text1 & .Text6 & "，請手動刪除資料夾！"
'                                If Dir("\\Typing2\電子送件暫存區\" & .Text1 & .Text6 & "\*.*") <> "" Then
'                                     Kill "\\Typing2\電子送件暫存區\" & .Text1 & .Text6 & "\*.*"
'                                End If
'                                If stRtn <> "" Then
'                                     RmDir "\\Typing2\電子送件暫存區\" & .Text1 & .Text6
'                                End If
                            If Dir("\\" & strTyping2Path & "\電子送件暫存區\" & .Text1 & .Text6, vbDirectory) <> "" Then
                                stRtn = convForm(strTmp, 6) & "無法刪除\\" & strTyping2Path & "\電子送件暫存區\" & .Text1 & .Text6 & "，請手動刪除資料夾！"
                                If Dir("\\" & strTyping2Path & "\電子送件暫存區\" & .Text1 & .Text6 & "\*.*") <> "" Then
                                     Kill "\\" & strTyping2Path & "\電子送件暫存區\" & .Text1 & .Text6 & "\*.*"
                                End If
                                If stRtn <> "" Then
                                     RmDir "\\" & strTyping2Path & "\電子送件暫存區\" & .Text1 & .Text6
                                End If
                            'end 2024/07/22
                            End If
                          End If
                      End If
                      stCon1 = Pub_GetCP31toCP27(.Text1, .Text6, .Text7, .Text8) 'Added by Lydia 2019/01/10 新申請案發文日
                      'Added by Lydia 2019/01/07 主動修正203、修正204、誤譯訂正433和申復、再審發文(有一併修正)時，若工程師沒有上傳中說word檔最終版本，系統會自動發email給工程師提醒
                      '若到了請款階段(203,204)還沒上傳，會再發一次通知提醒工程師並CC主管
                      If stCon1 <> "" Then 'Added by Lydia 2019/01/10 提申後才檢查
                        'Modified by Morgan 2024/3/5 +pa150
                        strTmp = "select cp01,cp02,cp03,cp04,cp09,cp10,cp14,cp27,st04,st16,pa150 from caseprogress,staff,patent where cp01='" & .Text1 & "' and cp02='" & .Text6 & "' and cp03='" & .Text7 & "' and cp04='" & .Text8 & "'  and cp60='" & .Text5 & "' and pa01(+)=cp01 and pa02(+)=cp02 and pa03(+)=cp03 and pa04(+)=cp04 "
                        'Modified by Lydia 2019/01/10 判斷提申後才檢查
                        'strTmp = strTmp & " and (cp10 in ('203','204','433') or (cp10 in ('107','205') and cp148='Y') ) and cp14=st01(+) "
                        'Modified by Lydia 2019/01/28 「主動修正是否已併入中說送件」，若選擇「是」則不用檢查中說最終版是否存在，若選擇「否」否: 則依舊在主動修正發文或請款時檢查中說最終版。
                        'strTmp = strTmp & " and ((cp10='203' and cp158 > " & stCon1 & " ) or (cp10 in ('204','433') and cp158 >= " & stCon1 & ")  or (cp10 in ('107','205') and cp148='Y' and cp158 >= " & stCon1 & " )) and cp14=st01(+)"
                        strTmp = strTmp & " and ((cp10='203' and cp158 > " & stCon1 & " and nvl(cp148,'N') <> 'Y') or (cp10 in ('204','433') and cp158 >= " & stCon1 & ")  or (cp10 in ('107','205') and cp148='Y' and cp158 >= " & stCon1 & " )) and cp14=st01(+)"
                        strTmp = strTmp & " order by cp27 desc "
                        intA = 1
                        Set rsA = ClsLawReadRstMsg(intA, strTmp)
                        If intA = 1 Then
                           'Add By Sindy 2025/10/22
                           rsA.MoveFirst
                           strCP09 = ""
                           Do While Not rsA.EOF
                              strCP09 = strCP09 & "," & rsA.Fields("cp09")
                              rsA.MoveNext
                           Loop
                           rsA.MoveFirst
                           If Left(strCP09, 1) = "," Then strCP09 = Mid(strCP09, 2)
                           '2025/10/22 END
                           
                           'Modified by Lydia 2020/03/03 改成模組
                           ''中說word檔最終版本=>案號-送件日.FIX_U
                           'strTmp = "\\Typing2\專利案件\" & Left(Val(rsA.Fields("cp02")), 3) & "\" & rsA.Fields("cp01") & rsA.Fields("cp02") & "-" & TransDate("" & rsA.Fields("cp27"), 1) & ".fix_u.doc*"
                           'stRtn = Dir(strTmp)
                           ''Modified by Lydia 2019/01/10  +圖式
                           ''If stRtn = "" And "" & rsA.Fields("cp14") <> "" Then
                           'strTmp = "\\Typing2\專利案件\" & Left(Val(rsA.Fields("cp02")), 3) & "\" & rsA.Fields("cp01") & rsA.Fields("cp02") & "-" & TransDate("" & rsA.Fields("cp27"), 1) & ".fig.pdf"
                           'stCon2 = Dir(strTmp)
                           'If stRtn & stCon2 = "" And "" & rsA.Fields("cp14") <> "" Then
                           'Modify By Sindy 2025/10/22 +, strCP09
                           If Pub_ChkFixUExists(rsA.Fields("cp01"), rsA.Fields("cp02"), rsA.Fields("cp03"), rsA.Fields("cp04"), rsA.Fields("cp27"), strCP09) = False And "" & rsA.Fields("cp14") <> "" Then
                           'end 2020/03/03
                               stCon2 = ""
                           'end 2019/01/10
                               stCon1 = "" & rsA.Fields("cp14")
                               If "" & rsA.Fields("st04") = "1" Then
                                  '若人員離職,自然會轉寄給主管,不用cc
                                   stCon2 = Pub_GetFCPGrpMan("" & rsA.Fields("st16"))
                               End If
                               'Modified by Lydia 2019/01/10 寄件者改為FCP管制人
                               'PUB_SendMail strUserNum, stCon1, "", rsA.Fields("cp01") & "-" & rsA.Fields("cp02") & IIf(rsA.Fields("cp03") & rsA.Fields("cp04") <> "000", "-" & rsA.Fields("cp03") & "-" & rsA.Fields("cp04"), "") & "已請款，未上傳中說Word檔最終版本，請儘速上傳！", vbCrLf & "同主旨", , , , , , IIf(stCon1 <> stCon2, stCon2, "")
                               strTmp = ""
                               strTmp = PUB_GetFCPHandler(.Text1, .Text6, .Text7, .Text8)
                               'Modified by Lydia 2019/01/16 stSender並不是寄件者的ID
                               'PUB_SendMail IIf(strTmp <> "", strTmp, strUserNum), stCon1, "", rsA.Fields("cp01") & "-" & rsA.Fields("cp02") & IIf(rsA.Fields("cp03") & rsA.Fields("cp04") <> "000", "-" & rsA.Fields("cp03") & "-" & rsA.Fields("cp04"), "") & "已請款，未上傳中說Word檔最終版本，請儘速上傳！", vbCrLf & "同主旨", , , , , , IIf(stCon1 <> stCon2, stCon2, "")
                               'Modified by Morgan 2024/3/5 機械組案件主旨都加【機械設計組】--Sharon
                               PUB_SendMail strUserNum, stCon1, "", IIf("" & rsA.Fields("pa150") = "4", "【機械設計組】", "") & rsA.Fields("cp01") & "-" & rsA.Fields("cp02") & IIf(rsA.Fields("cp03") & rsA.Fields("cp04") <> "000", "-" & rsA.Fields("cp03") & "-" & rsA.Fields("cp04"), "") & "已請款，未上傳中說Word檔最終版本，請儘速上傳！", vbCrLf & "同主旨", , , , , , IIf(stCon1 <> stCon2, stCon2, ""), IIf(strTmp <> "", strTmp, strUserNum)
                               'end 2019/01/10
                           End If
                        End If
                      End If 'end 2019/01/10 判斷提申後才檢查
                      'end 2019/01/07
                  End If
                  'end 2018/12/03
                  End With
               End If
               
'            Case "Frmacc21h1"
'               With Frmacc21h1
'                  .FormDisabled
'               End With
'               adoTaie.CommitTrans
            Case "Frmacc21i0"
               Frmacc21i0_Save
            Case "Frmacc21j0"
               Frmacc21j0_Save
               If Not Frmacc21j0.m_PrevForm Is Nothing Then bolUnload = True  'Add by Sindy 2018/2/23
            Case "Frmacc21k0"
               Frmacc21k0_Save
            Case "Frmacc21m0"
               Frmacc21m0_Save Frmacc21m0
            Case "Frmacc21o0"
               Frmacc21o0_Save Frmacc21o0
            'Add By Sindy 2009/06/06
            Case "Frmacc21s0"
               Frmacc21s0_Save Frmacc21s0
            '2009/06/06 End
            Case "Frmacc21t0" 'Add by Morgan 2010/11/19
               Frmacc21t0.SaveRec
         End Select
         If strControlButton <> MsgText(602) Then
            strSaveConfirm = MsgText(601)
            Select Case strFormName
               Case "Frmacc21h1"
                  tool7_enabled
               Case "Frmacc21i0"
                  tool8_enabled
               Case "Frmacc21t0" 'Add by Morgan 2010/11/19
               
               Case Else
                  tool1_enabled
            End Select
            mdiMain.StatusBar1.Panels(1).Text = MsgText(17)
            
            'Added by Morgan 2017/1/6
            '請款單修改存檔時一律彈出點數分配畫面(因收文號有可能更換)
            If strFormName = "Frmacc21h0" Then
               With Frmacc21h0
               '檢查有acc1n0為修改
               cnnConnection.Execute "update acc1n0 set a1n01=a1n01 where a1n01='" & .Text5 & "' and rownum<2", intI
               If intI = 1 Then
                  MsgBox "若請款單收文號有變動請重新分配點數！", vbInformation
                  .Command4.Value = True
               End If
               End With
            End If
            'end 2017/1/6
         End If
         strControlButton = MsgText(601)
      '取消
      Case vbKeyF10
         If strSaveConfirm = MsgText(601) Then
            Exit Sub
         End If
         Select Case strFormName
            Case "Frmacc2150"
               With Frmacc2150
                  If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
                     adoTaie.Execute "delete from acc151 where axf01 = '" & .Text2 & "'"
                     adoTaie.Execute "delete from acc150 where a1501 = '" & .Text2 & "'"
                     Frmacc2150_Clear
                     .adoacc150.ReQuery
                     .AdodcRefresh
                     .AdodcClear
                     If .adoacc150.RecordCount <> 0 Then
                        .RecordShow
                     Else
                        StatusClear
                     End If
                  End If
                  .FormDisabled
                  If Not .m_ParentForm Is Nothing Then bolUnload = True 'Added by Morgan 2016/7/28
               End With
               adoTaie.RollbackTrans
               
               'Added by Moran 2024/12/27 修改取消後重讀資料
               If strSaveConfirm = MsgText(4) Then
                  Frmacc2150.Command3.Value = True
               End If
               'end 2024/12/27
            Case "Frmacc2160"
               With Frmacc2160
                  If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
                     adoTaie.Execute "delete from acc161 where axg01 = '" & .Text2 & "'"
                     adoTaie.Execute "delete from acc160 where a1601 = '" & .Text2 & "'"
                     Frmacc2160_Clear
                     .adoacc160.ReQuery
                     .AdodcRefresh
                     .AdodcClear
                     If .adoacc160.RecordCount <> 0 Then
                        .RecordShow
                     Else
                        StatusClear
                     End If
                  End If
                  .FormDisabled
               End With
               adoTaie.RollbackTrans
            Case "Frmacc21h0"
               adoTaie.RollbackTrans
               With Frmacc21h0
                  .Command1.Enabled = False
                  .Command3.Enabled = False
                  .Command5.Enabled = True
                  .Command6.Enabled = True 'Add by Amy 2014/06/26 避免新增時再按此鈕產錯誤
                  If Len(.Text5) = 10 Then
                     .Command2.Enabled = False
                     .Command4.Enabled = False 'Add by Morgan 2010/5/21
                  Else
                     .Command2.Enabled = True
                     .Command4.Enabled = True 'Add by Morgan 2010/5/21
                     .Command5.Value = True 'Add by Morgan 2010/3/2 重新查詢
                  End If
                  .Text5.Enabled = True 'Added by Morgan 2012/9/28
               End With
               
            Case "Frmacc21h1"
               With Frmacc21h1
                  If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
                     adoTaie.Execute "delete from acc1l0 where a1l01 = '" & .Text1 & "'"
                     .adoacc1k0.ReQuery
                     .AdodcRefresh
                     .SumShow
                  End If
                  .FormDisabled
               End With
               adoTaie.RollbackTrans
            Case "Frmacc21t0" 'Add by Morgan 2010/11/19
               Frmacc21t0.CancelEdit
               
         End Select
         strSaveConfirm = MsgText(601)
         Select Case strFormName
            Case "Frmacc21h1"
               tool7_enabled
            Case "Frmacc21i0"
               tool8_enabled
            Case "Frmacc21t0" 'Add by Morgan 2010/11/19
            
            Case Else
               tool1_enabled
         End Select
      '刪除
      Case vbKeyF5
         If mdiMain.Toolbar1.Buttons.Item(8).Enabled = False Then
            Exit Sub
         End If
         If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
            Exit Sub
         End If
         Select Case strFormName
            Case "Frmacc2150"
               If CheckUse("Frmacc2150", strDel) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
            Case "Frmacc2160"
               If CheckUse("Frmacc2160", strDel) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
            Case "Frmacc21g0"
               If CheckUse("Frmacc21g0", strDel) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               'Added by Lydia 2025/03/18 要刪除請款項目時，先以系統類別+項目代號檢查ACC1L0的A1L03+A1L04，若有存在於ACC1L0，則不可刪除。
               strTmp = "SELECT COUNT(*) CNT FROM ACC1L0 WHERE A1L03 = '" & Trim(Frmacc21g0.Text1.Text) & "' AND A1L04='" & Trim(Frmacc21g0.Text2.Text) & "' "
               intA = 1
               Set rsA = ClsLawReadRstMsg(intA, strTmp)
               If intA = 1 Then
                  If Val("" & rsA.Fields("cnt")) > 0 Then
                     MsgBox "請款項目已存在於國外請款資料明細檔，不可刪除！", vbExclamation + vbOKOnly
                     strSaveConfirm = MsgText(601)
                     Exit Sub
                  End If
               End If
               Set rsA = Nothing
               'end 2025/03/18

            Case "Frmacc21h0"
               If CheckUse("Frmacc21h0", strDel) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
            Case "Frmacc21j0"
               'Added by Morgan 2018/3/6
               MsgBox "帳單作廢不可刪除！", vbCritical
               strSaveConfirm = MsgText(601)
               Exit Sub
               'end 2018/3/6
               If CheckUse("Frmacc21j0", strDel) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
            Case "Frmacc21k0"
               'Added by Morgan 2018/3/7
               MsgBox "請款單作廢不可刪除！", vbCritical
               strSaveConfirm = MsgText(601)
               Exit Sub
               'end 2018/3/7
               If CheckUse("Frmacc21k0", strDel) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
            Case "Frmacc21m0"
               If CheckUse("Frmacc21m0", strDel) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
            Case "Frmacc21o0"
               If CheckUse("Frmacc21o0", strDel) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
            'Add By Sindy 2009/06/06
            Case "Frmacc21s0"
               If CheckUse("Frmacc21s0", strDel) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
            '2009/06/06 End
         End Select
         strDelConfirm = MsgBox("確定刪除?", vbOKCancel + vbDefaultButton2, MsgText(5))
         If strDelConfirm = vbCancel Then
            Exit Sub
         End If
         Select Case strFormName
            Case "Frmacc2150"
               Frmacc2150_Delete
               Frmacc2150_Clear
            Case "Frmacc2160"
               Frmacc2160_Delete
               Frmacc2160_Clear
            Case "Frmacc21g0"
               Frmacc21g0_Delete Frmacc21g0
               Frmacc21g0_Clear Frmacc21g0
            Case "Frmacc21h0"
               Frmacc21h0_Delete
               Frmacc21h0_Clear
               With Frmacc21h0
                  .AdodcRefresh
               End With
            Case "Frmacc21j0"
               Frmacc21j0_Delete
               Frmacc21j0_Clear
            Case "Frmacc21k0"
               Frmacc21k0_Delete
               Frmacc21k0_Clear
            Case "Frmacc21m0"
               Frmacc21m0_Delete Frmacc21m0
               Frmacc21m0_Clear Frmacc21m0
            Case "Frmacc21o0"
               Frmacc21o0_Delete Frmacc21o0
               Frmacc21o0_Clear Frmacc21o0
            'Add By Sindy 2009/06/06
            Case "Frmacc21s0"
               Frmacc21s0_Delete Frmacc21s0
               Frmacc21s0_Clear Frmacc21s0
            '2009/06/06 End
            
            Case "Frmacc21t0" 'Add by Morgan 2010/11/19
               Frmacc21t0.DeleteRec
         End Select
      '查詢
      Case vbKeyF4
         If mdiMain.Toolbar1.Buttons.Item(9).Enabled = False Then
            Exit Sub
         End If
         If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
            Exit Sub
         End If
         strExitControl = MsgText(601)
         Select Case strFormName
            Case "Frmacc2150"
               strFormLink = strFormName
               Frmacc2150_Clear
               Frmacc2150.Enabled = False
               Frmacc2151.Show
            Case "Frmacc2160"
               Frmacc2160_Clear
               Frmacc2160.Enabled = False
               Frmacc2161.Show
            Case "Frmacc21g0"
               Frmacc21g0_Clear Frmacc21g0
               Frmacc21g0.Enabled = False
               Frmacc21g1.Show
            Case "Frmacc21h0"
               strFormLink = strFormName
               Frmacc21h0_Clear
               Frmacc21h0.Enabled = False
               Frmacc21h2.Show
            Case "Frmacc21i0"
               Frmacc21i0_Clear
               Frmacc21i0.Enabled = False
               Frmacc21i1.Show
            Case "Frmacc21j0"
               strFormLink = strFormName
               Frmacc21j0_Clear
               Frmacc21j0.Enabled = False
               Frmacc2151.Show
            Case "Frmacc21k0"
               strFormLink = strFormName
               Frmacc21k0_Clear
               Frmacc21k0.Enabled = False
               Frmacc21h2.Show
            Case "Frmacc21m0"
               Exit Sub
            Case "Frmacc21o0"
               Frmacc21o0_Clear Frmacc21o0
               Frmacc21o0.Enabled = False
               Frmacc21o1.Show
            'Add By Sindy 2009/06/06
            Case "Frmacc21s0"
               Exit Sub
            '2009/06/06 End
            'Add by Amy 2015/09/03 開放查客戶代理人
            Case "Frmacc2210"
                Frmacc2210.Enabled = False
                frm100114_5.Tag = "Frmacc2210"
                frm100114_5.Show
                Exit Sub
            Case Else
               tool1_enabled
         End Select
         tool3_enabled
'         strExitControl = "Y"
      Case vbKeyF7
'         strExitControl = MsgText(601)
'         tool3_enabled
'         strExitControl = "Y"
      '第一筆
      Case vbKeyHome
         If mdiMain.Toolbar1.Buttons.Item(13).Enabled = False Then
            Exit Sub
         End If
         If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
            Exit Sub
         End If
         Select Case strFormName
            Case "Frmacc2150"
               Frmacc2150_First
            Case "Frmacc2160"
               Frmacc2160_First
            Case "Frmacc21g0"
               Frmacc21g0_First Frmacc21g0
            Case "Frmacc21h0"
               Frmacc21h0_First
            Case "Frmacc21i0"
               Frmacc21i0_First
            Case "Frmacc21j0"
               Frmacc21j0_First
            Case "Frmacc21k0"
               Frmacc21k0_First
            Case "Frmacc21m0"
               Frmacc21m0_First Frmacc21m0
            Case "Frmacc21o0"
               Frmacc21o0_First Frmacc21o0
            'Add By Sindy 2009/06/06
            Case "Frmacc21s0"
               Frmacc21s0_First Frmacc21s0
            '2009/06/06 End
            Case "Frmacc21t0" 'Add by Morgan 2010/11/19
               Frmacc21t0.MoveFirst
         End Select
      '上一筆
      Case vbKeyPageUp
         If mdiMain.Toolbar1.Buttons.Item(14).Enabled = False Then
            Exit Sub
         End If
         If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
            Exit Sub
         End If
         Select Case strFormName
            Case "Frmacc2150"
               Frmacc2150_Previous
            Case "Frmacc2160"
               Frmacc2160_Previous
            Case "Frmacc21g0"
               Frmacc21g0_Previous Frmacc21g0
            Case "Frmacc21h0"
               Frmacc21h0_Previous
            Case "Frmacc21i0"
               Frmacc21i0_Previous
            Case "Frmacc21j0"
               Frmacc21j0_Previous
            Case "Frmacc21k0"
               Frmacc21k0_Previous
            Case "Frmacc21m0"
               Frmacc21m0_Previous Frmacc21m0
            Case "Frmacc21o0"
               Frmacc21o0_Previous Frmacc21o0
            'Add By Sindy 2009/06/06
            Case "Frmacc21s0"
               Frmacc21s0_Previous Frmacc21s0
            '2009/06/06 End
            Case "Frmacc21t0" 'Add by Morgan 2010/11/19
               Frmacc21t0.MovePrevious
         End Select
      '下一筆
      Case vbKeyPageDown
         If mdiMain.Toolbar1.Buttons.Item(15).Enabled = False Then
            Exit Sub
         End If
         If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
            Exit Sub
         End If
         Select Case strFormName
            Case "Frmacc2150"
               Frmacc2150_Next
            Case "Frmacc2160"
               Frmacc2160_Next
            Case "Frmacc21g0"
               Frmacc21g0_Next Frmacc21g0
            Case "Frmacc21h0"
               Frmacc21h0_Next
            Case "Frmacc21i0"
               Frmacc21i0_Next
            Case "Frmacc21j0"
               Frmacc21j0_Next
            Case "Frmacc21k0"
               Frmacc21k0_Next
            Case "Frmacc21m0"
               Frmacc21m0_Next Frmacc21m0
            Case "Frmacc21o0"
               Frmacc21o0_Next Frmacc21o0
            'Add By Sindy 2009/06/06
            Case "Frmacc21s0"
               Frmacc21s0_Next Frmacc21s0
            '2009/06/06 End
            Case "Frmacc21t0" 'Add by Morgan 2010/11/19
               Frmacc21t0.MoveNext
         End Select
      '最後一筆
      Case vbKeyEnd
         If mdiMain.Toolbar1.Buttons.Item(16).Enabled = False Then
            Exit Sub
         End If
         If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
            Exit Sub
         End If
         Select Case strFormName
            Case "Frmacc2150"
               Frmacc2150_Last
            Case "Frmacc2160"
               Frmacc2160_Last
            Case "Frmacc21g0"
               Frmacc21g0_Last Frmacc21g0
            Case "Frmacc21h0"
               Frmacc21h0_Last
            Case "Frmacc21i0"
               Frmacc21i0_Last
            Case "Frmacc21j0"
               Frmacc21j0_Last
            Case "Frmacc21k0"
               Frmacc21k0_Last
            Case "Frmacc21m0"
               Frmacc21m0_Last Frmacc21m0
            Case "Frmacc21o0"
               Frmacc21o0_Last Frmacc21o0
            'Add By Sindy 2009/06/06
            Case "Frmacc21s0"
               Frmacc21s0_Last Frmacc21s0
            '2009/06/06 End
            Case "Frmacc21t0" 'Add by Morgan 2010/11/19
               Frmacc21t0.MoveLast
         End Select
   End Select

   If strControlButton = MsgText(601) And bolUnload Then KeyEnter vbKeyEscape 'Added by Morgan 2016/7/28
   
'Added by lydia 2018/12/03
   Exit Sub
   
ErrHandle:
   If Err.Number <> 0 Then
      If strFormName = "Frmacc21h0" Then
            If stRtn <> "" Then
                PUB_SendMail strUserNum, Trim(Mid(stRtn, 1, 6)), "", Mid(stRtn, 7), "同主旨"
                stRtn = ""
            End If
            Resume Next
      End If
   End If
'end 2018/12/03
End Sub

'Added by Lydia 2020/03/03 主動修正203、修正204、誤譯訂正433和申復、再審發文(有一併修正)時，若工程師沒有上傳中說word檔最終版本，系統會自動發email給工程師提醒
                           '若到了請款階段(203,204)還沒上傳，會再發一次通知提醒工程師並CC主管
'Modify By Sindy 2025/10/22 + ByVal pCP09 As String:文號, 考慮有多個文號的狀況
Public Function Pub_ChkFixUExists(ByVal pCP01 As String, ByVal pCP02 As String, ByVal pCP03 As String, ByVal pCP04 As String, ByVal pCP27 As String, _
   ByVal pCP09 As String) As Boolean
Dim rsRD As New ADODB.Recordset
Dim intR As Integer
Dim strTempR As String
Dim stRtn As String

   Pub_ChkFixUExists = False
   If pCP01 = "" Or Val(pCP02) < 6 Then Exit Function
       
   '中說word檔最終版本=>案號-送件日.FIX_U
   'Remove by Lydia 2020/04/07 已改放在原始檔區
   'strTempR = "\\Typing2\專利案件\" & Left(Val(pCP02), 3) & "\" & pCP01 & pCP02 & "-" & TransDate(pCP27, 1) & ".fix_u.doc*"
   'stRtn = Dir(strTempR)
   'If stRtn = "" Then
   '    '圖式
   '    strTempR = "\\Typing2\專利案件\" & Left(Val(pCP02), 3) & "\" & pCP01 & pCP02 & "-" & TransDate(pCP27, 1) & ".fig.pdf"
   '    stRtn = Dir(strTempR)
   'End If
   'end 2020/04/07
   
   If stRtn = "" Then '原始檔區的專利案件
      'Modified by Lydia 2020/04/07 最終版自動上傳有＋發文日.時間
      'strTempR = "SELECT CP01,CP02,CP03,CP04,CP09,NVL(A02,0) CNT1,NVL(B02,0) CNT2 " & _
                      "FROM CASEPROGRESS,(SELECT CPF01 A01,COUNT(*) A02 FROM CASEPAPERFILE WHERE UPPER(CPF02) LIKE '%" & TransDate(pCP27, 1) & ".FIX_U.DOC%' GROUP BY CPF01) VT01" & _
                      ",(SELECT CPF01 B01,COUNT(*) B02 FROM CASEPAPERFILE WHERE UPPER(CPF02) LIKE '%" & TransDate(pCP27, 1) & ".FIG.PDF' GROUP BY CPF01) VT02 " & _
                      "WHERE CP01='" & pCP01 & "' AND CP02='" & pCP02 & "' AND CP03='" & pCP03 & "' AND CP04='" & pCP04 & "'  AND SUBSTR(CP09,1,1)='D' AND CP10='" & cnt專利案件 & "' AND CP159=0 " & _
                      "AND CP09=A01(+) AND CP09=B01(+) "
      'Modified by Lydia 2021/04/19 檔名有「送件日.FIX_U.、送件日.FIG.PDF、送件日.FIX.」就不發email通知工程師；「.FIX_U.、.FIX.」將不限制檔案型態，例如Word檔,PDF檔, Txt檔
      'strTempR = "SELECT CP01,CP02,CP03,CP04,CP09,NVL(A02,0) CNT1,NVL(B02,0) CNT2 " & _
                      "FROM CASEPROGRESS,(SELECT CPF01 A01,COUNT(*) A02 FROM CASEPAPERFILE WHERE UPPER(CPF02) LIKE '%" & TransDate(pCP27, 1) & "%.FIX_U.DOC%' GROUP BY CPF01) VT01" & _
                      ",(SELECT CPF01 B01,COUNT(*) B02 FROM CASEPAPERFILE WHERE UPPER(CPF02) LIKE '%" & TransDate(pCP27, 1) & "%.FIG.PDF' GROUP BY CPF01) VT02 " & _
                      "WHERE CP01='" & pCP01 & "' AND CP02='" & pCP02 & "' AND CP03='" & pCP03 & "' AND CP04='" & pCP04 & "'  AND SUBSTR(CP09,1,1)='D' AND CP10='" & cnt專利案件 & "' AND CP159=0 " & _
                      "AND CP09=A01(+) AND CP09=B01(+) "
      strTempR = "SELECT CP01,CP02,CP03,CP04,CP09,NVL(A02,0) CNT1,NVL(B02,0) CNT2 " & _
                 "FROM CASEPROGRESS,(SELECT CPF01 A01,COUNT(*) A02 FROM CASEPAPERFILE WHERE UPPER(CPF02) LIKE '%" & TransDate(pCP27, 1) & "%.FIX_U.%' OR UPPER(CPF02) LIKE '%" & TransDate(pCP27, 1) & "%.FIX.%' GROUP BY CPF01) VT01" & _
                 ",(SELECT CPF01 B01,COUNT(*) B02 FROM CASEPAPERFILE WHERE UPPER(CPF02) LIKE '%" & TransDate(pCP27, 1) & "%.FIG.PDF' GROUP BY CPF01) VT02 " & _
                 "WHERE CP01='" & pCP01 & "' AND CP02='" & pCP02 & "' AND CP03='" & pCP03 & "' AND CP04='" & pCP04 & "'  AND SUBSTR(CP09,1,1)='D' AND CP10='" & cnt專利案件 & "' AND CP159=0 " & _
                 "AND CP09=A01(+) AND CP09=B01(+) "
      'Add By Sindy 2025/10/22
      strTempR = strTempR & "union " & _
                 "SELECT CP01,CP02,CP03,CP04,CP09,NVL(A02,0) CNT1,NVL(B02,0) CNT2 " & _
                 "FROM CASEPROGRESS,(SELECT CPF01 A01,COUNT(*) A02 FROM CASEPAPERFILE WHERE UPPER(CPF02) LIKE '%.FIX_U.%' OR UPPER(CPF02) LIKE '%.FIX.%' GROUP BY CPF01) VT01" & _
                 ",(SELECT CPF01 B01,COUNT(*) B02 FROM CASEPAPERFILE WHERE UPPER(CPF02) LIKE '%.FIG.PDF' GROUP BY CPF01) VT02 " & _
                 "WHERE CP09 in('" & Replace(pCP09, ",", "','") & "') AND CP159=0 " & _
                 "AND CP09=A01(+) AND CP09=B01(+) "
      '加總
      strTempR = "select sum(CNT1) CNT1,sum(CNT2) CNT2 from (" & strTempR & ")"
      '2025/10/22 END
      intR = 1
      Set rsRD = ClsLawReadRstMsg(intR, strTempR)
      If intR = 0 Then
      ElseIf intR = 1 Then
           If Val("" & rsRD.Fields("cnt1")) + Val("" & rsRD.Fields("cnt2")) > 0 Then
               stRtn = "Y"
           End If
      End If
   End If
   
   If stRtn <> "" Then
      Pub_ChkFixUExists = True
   End If
End Function

'Add by Amy 2025/11/12 開啟請款單輸入
Public Function Pub_Open21H0(stF0301 As String, stFormN As String, mPrev As Form, stCaseNo1 As String, stCaseNo2 As String, stCaseNo3 As String, stCaseNo4 As String _
  , ByRef stErrMsg As String, Optional ByVal stNP07 As String = "", Optional ByVal stClose As String = "") As Boolean
   Dim RsQ As New ADODB.Recordset, strQ As String, intQ As Integer, ii As Integer, strA As String, strD As String, strShowMsg As String, strCPMsg As String
   '        不續辦/閉卷           / 可直接加入之案件性質/ 需確認之案件性質 / 不在CP未收款之案件性質中/可直接加入之總收文號 /需確認之總收文號/請款單號
   Dim stNowCP10 As String, stCP10 As String, stChkCP10 As String, stNotInCP10 As String, stInsCP09 As String, stChkCP09 As String, stInvNo As String
   Dim Gofrm As Form, MnuFrm As Form, arrCP
On Error GoTo ErrHnd

   Select Case UCase(stFormN)
      Case "FRM110101_2" '解除期限
         stNowCP10 = "703"
         If stClose = "Y" Then stNowCP10 = "704"
      Case "FRM110103_3" '閉卷
         stNowCP10 = "704"
   End Select
   Pub_Open21H0 = False
   stErrMsg = ""
   
   '取得請款項目代號資料與未請款金額
   strA = GetAcc21H0Sql("0", "Pub_Open21H0", stF0301, stCaseNo1 & "-" & stCaseNo2 & "-" & stCaseNo3 & "-" & stCaseNo4)
   strD = GetAcc21H0Sql("1", "Pub_Open21H0", stF0301, , stNowCP10)
   '解除期限/閉卷 or 請款代號小於3碼,其他項目 (避免只有請別道 ex:X11400533 FCT-052536)
   strQ = "Select ccd01,SubStr(CCD04,1,3) as CCD04,cp10,CPTol,Total,CP10CNT,Sort From (" & _
                    "Select cp10,Sum(Nvl(cp16,0)) as CPTol,Count(cp10) as CP10CNT From(" & strA & ") Group by cp10 )" & _
                  ",(" & strD & " Union " & _
                     "Select CCD01,SubStr(CCD04,1,3) as CCD04,Sum(CCD05) as Total,2 as Sort From CloseCaseDetail " & _
                     "Where CCD01='" & stF0301 & "' And CCD02='1' And SubStr(CCD04,1,3)<>'" & stNowCP10 & "' And length(CCD04)>=3 " & _
                     "Group by ccd01,SubStr(CCD04,1,3)" & _
                     " ) Where cp10(+)=ccd04 " & _
               " Order by sort"
   intQ = 1
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      RsQ.MoveFirst
      Do While Not RsQ.EOF
         '解除期限/閉卷當道,進度不會有金額
         If "" & RsQ.Fields("CCD04") = stNowCP10 Then
            stCP10 = stCP10 & "," & RsQ.Fields("CCD04") '可直接加入之案件性質
         '其他道進度只有符合1筆案件性質且進度金額=請款項目金額,才可直接加入
         ElseIf "" & RsQ.Fields("CCD04") <> stNowCP10 And "" & RsQ.Fields("CP10CNT") = "1" And Val("" & RsQ.Fields("CPTol")) = Val("" & RsQ.Fields("Total")) Then
            '  ex:FCP-050875 有2筆 303金額都是1500 未請款需讓user自行加入-秀玲
            stCP10 = stCP10 & "," & RsQ.Fields("CCD04") '可直接加入之案件性質
         Else
            stChkCP10 = stChkCP10 & "," & RsQ.Fields("CCD04") '需確認之案件性質
            '其他道無此案案件性質未請款
            If "" & RsQ.Fields("cp10") = "" Then
               stNotInCP10 = stNotInCP10 & "," & RsQ.Fields("CCD04") '不在CP未收款之案件性質中
               strCPMsg = strCPMsg & ",[" & RsQ.Fields("CCD04") & "-" & GetCaseTypeName(stCaseNo1, "" & RsQ.Fields("CCD04"), 0) & "] 無此案件性質未請款"
            Else
               strCPMsg = strCPMsg & ",[" & RsQ.Fields("CCD04") & "-" & GetCaseTypeName(stCaseNo1, "" & RsQ.Fields("CCD04"), 0) & "] 有 " & RsQ.Fields("CP10CNT") & "筆"
               If Val("" & RsQ.Fields("CPTol")) <> Val("" & RsQ.Fields("Total")) Then
                  strCPMsg = strCPMsg & "金額與進度不同"
               End If
            End If
         End If
        
         RsQ.MoveNext
      Loop
   End If
   If stCP10 <> "" Then stCP10 = Mid(stCP10, 2)
   If stChkCP10 <> "" Then stChkCP10 = Mid(stChkCP10, 2)
   If stNotInCP10 <> "" Then stNotInCP10 = Mid(stNotInCP10, 2)
   
   '以抓到之案件性質抓總收文號
   strQ = "Select cp05,cp09,cp10,1 as State From (" & strA & ") Where cp10 in('" & Replace(stCP10, ",", "','") & "') "
   If stChkCP10 <> "" Then
      strQ = strQ & "Union Select cp05,cp09,cp10,2 as State From (" & strA & ") Where cp10 in('" & Replace(stChkCP10, ",", "','") & "') "
   End If
   strQ = strQ & "Order by cp05, cp09"
   Set RsQ = ClsLawReadRstMsg(intQ, strQ)
   If intQ = 1 Then
      RsQ.MoveFirst
      Do While Not RsQ.EOF
         If RsQ.Fields("State") = 1 Then
            stInsCP09 = stInsCP09 & "," & RsQ.Fields("cp09")
         Else
            stChkCP09 = stChkCP09 & "," & RsQ.Fields("cp09") '需確認之總收文號
         End If
         RsQ.MoveNext
      Loop
   End If
   If stInsCP09 <> "" Then stInsCP09 = Mid(stInsCP09, 2)
   If stChkCP09 <> "" Then stChkCP09 = Mid(stChkCP09, 2)
   
   '以下參考frm030203_02
   Set Gofrm = Forms(0).GetForm("Frmacc21h0")
   Gofrm.Show
   PUB_InitForm Gofrm, , 5850
   Set MnuFrm = Forms(0).GetForm("mdiMain")
   MnuFrm.ToolShow
   Call tool1_enabled
   Set Gofrm.frmlink = mPrev
   strFormName = Gofrm.Name
   Call KeyEnter(vbKeyF2)
   Gofrm.Text1 = stCaseNo1
   Gofrm.Text6 = stCaseNo2
   Gofrm.Text7 = stCaseNo3
   Gofrm.Text8 = stCaseNo4
   Call Gofrm.CaseQuery
   Gofrm.stF0301 = stF0301
   Gofrm.stNotInCP10 = stNotInCP10
   Gofrm.stNP07 = stNP07
   stInvNo = Gofrm.Text5 '請款單輸入.請款單號
   
   '抓不到任何未請款之進度資料 ex:FCT-052536 銷[612-補充理由]期限,但請款代號:614 (X11400533)
   If stInsCP09 = "" Then
      stErrMsg = "結案單請款項目未對應到進度資料" & Replace(strCPMsg, ",", vbCrLf) & vbCrLf & "請自行操作"
      GoTo ErrHnd
   End If
   
   '可以直接寫入之總收文號-帶入請款單輸入第1個畫面
   arrCP = Split(stInsCP09, ",")
   Gofrm.stCP09 = ""
   For ii = LBound(arrCP) To UBound(arrCP)
      Gofrm.stCP09 = arrCP(ii)
      Call Gofrm.Command1_Click
   Next ii

   '*** 帶入請款單輸入第2個畫面 ***
   '請款項目=總收文號數 且無其他請款項目,畫面開至[請款單內容輸入],並將資料寫入
   '避免只有請別道 ex:X11400533 FCT-052536
   If Gofrm.adoadodc2.RecordCount = UBound(arrCP) + 1 And stChkCP09 = "" Then
      Gofrm.stUpdCP09 = stInsCP09
      Gofrm.stNowCP10 = stNowCP10 '不續辦or閉卷
      Call KeyEnter(vbKeyF9)
      Call Gofrm.Command2_Click
 
      '有對應不到未請款之進度資料,先寫入第2畫面再彈訊息-秀玲
      '  ex:FCT-052536 銷[612-補充理由]期限,但請款代號:614 (X11400533)
      '  ex:FCT-050785 無303(原要測式FCT-050875 有303) 輸錯案號
      If stNotInCP10 <> "" Then
         stErrMsg = "結案單無法對應之請款項目如下：" & Replace(strCPMsg, ",", vbCrLf) & vbCrLf & _
                              "已新增至明細,請再確認選擇之總收文號(請款單輸入選擇總收文號畫面)"
         GoTo ErrHnd
      End If
   End If
   '*** End 帶入請款單輸入第2個畫面 ***
   
   '有訊息需彈
   If strCPMsg <> "" Then
      stErrMsg = "結案單無法對應之請款項目如下：" & Replace(strCPMsg, ",", vbCrLf) & vbCrLf & "請確認"
      GoTo ErrHnd
   End If
   Pub_Open21H0 = True
   
ErrHnd:
   If Err.Number <> 0 Then
      'Resume
      stErrMsg = "請款單輸入有誤！(Pub_Open21H0)" & vbCrLf & _
                        Err.Description & vbCrLf & "請通知電腦中心！"
   End If
   Set RsQ = Nothing
End Function

