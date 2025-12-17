Attribute VB_Name = "acc_var"
'Memo by Morgan2010/8/18 日期欄已修改
Option Explicit

'*************************************************
'  功能鍵對照函式
'
'*************************************************
Public Sub KeyEnter(InputCode As Integer)
On Error GoTo Checking
   Frmacc0000.StatusBar1.Panels(1).Text = ""
   Select Case InputCode
      Case vbKeyEscape '離開
         FormExit
      Case vbKeyF2 '新增
         If Frmacc0000.Toolbar1.Buttons.Item(4).Enabled = False Or strSaveConfirm = MsgText(4) Then
            Exit Sub
         End If
         strSaveConfirm = MsgText(3)
         Select Case strFormName
            'Added by  Lydia 2020/03/26 從account.aacc_var複製
            Case "Frmacc1130"
               If Frmacc0000.str中所收據人員 <> "" And InStr(Frmacc0000.str中所收據人員, strUserNum) > 0 Then
                    Frmacc1130_Clear
               Else
                    If CheckUse("Frmacc1130", strAdd) = False Then
                       strSaveConfirm = MsgText(601)
                       Exit Sub
                    End If
                    Frmacc1130_Clear
               End If
            Case "Frmacc1140"
               If Frmacc0000.str中所收據人員 <> "" And InStr(Frmacc0000.str中所收據人員, strUserNum) > 0 Then
                    Frmacc1140.Frmacc1140_Clear
               Else
                    If CheckUse("Frmacc1140", strAdd) = False Then
                       strSaveConfirm = MsgText(601)
                       Exit Sub
                    End If
                    Frmacc1140.Frmacc1140_Clear 'Modify by Amy 2015/04/17 搬回form
               End If
            'end 2020/03/26
            Case "Frmacc7100"
               If CheckUse("Frmacc7100", strAdd) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
                    Frmacc7100.oState = "1"
                    Frmacc7100_UnLock
                    Frmacc7100_Clear
                    'add by nick 2004/12/08
                    Frmacc7100.Label18.Caption = "所別：" & pub_strUserOffice & " (1.北所 2.中所 3.南所 4.高所 5.其他)"
                    'add by nick 2004/10/08
                    Frmacc7100.Command1.Enabled = False
            'Add by Morgan 2005/5/24
            Case "Frmacc41e0"
               If CheckUse("Frmacc41e0", strAdd) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               With Frmacc41e0
                  .txtA2301.Tag = .txtA2301
                  .FormClear
                  .FormEnable "1"
                  .txtNo.SetFocus
               End With
               'adoTaie.BeginTrans 'Mark by Amy 2018/02/12 改存檔控制就好
            'Add By Sindy 2013/12/19
            Case "Frmacc11p0"
               'Added by Lydia 2020/03/27 開放權限
               If Frmacc0000.str中所收據人員 <> "" And InStr(Frmacc0000.str中所收據人員, strUserNum) > 0 Then
                    Frmacc11p0.FormEnabled
                    Frmacc11p0.Frmacc11p0_Clear
                    If strControlButton <> MsgText(602) Then '+if 保留畫面資料,只清收據抬頭欄
                       Frmacc11p0.Frmacc11p0_Clear
                       Frmacc11p0.textA4221 = "1" '繳款書寄件處預設為1客戶
                    Else
                       Frmacc11p0.textA4201.Text = ""
                    End If
               Else
               'end 2020/03/27
                    If CheckUse("Frmacc11p0", strAdd) = False Then
                       strSaveConfirm = MsgText(601)
                       Exit Sub
                    End If
                    Frmacc11p0.FormEnabled
                    If strControlButton <> MsgText(602) Then '+if 保留畫面資料,只清收據抬頭欄
                       Frmacc11p0.Frmacc11p0_Clear
                       Frmacc11p0.textA4221 = "1" '繳款書寄件處預設為1客戶
                    Else
                       Frmacc11p0.textA4201.Text = ""
                    End If
                    'end 2020/03/27
               End If 'Added by Lydia 2020/03/27

            '2013/12/19 End
            'Add By Sindy 2012/8/29
            Case "Frmacc11n0"
               If CheckUse("Frmacc11n1", strAdd) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               Frmacc11n0.FormEnabled
               Frmacc11n0.Frmacc11n0_Clear
               Frmacc11n0.Text4.Enabled = False
               Frmacc11n0.Command1.Enabled = False
            '2012/8/29 End
         End Select
         tool2_enabled
      Case vbKeyF3 '修改
         If Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = False Or strSaveConfirm = MsgText(3) Then
            Exit Sub
         End If
         strSaveConfirm = MsgText(4)
         Select Case strFormName
            Case "Frmacc7100"
               If CheckUse("Frmacc7100", strEdit) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               '93.12.16 ADD BY SONIA
               If Frmacc7100.M_REC = "Y" Then
                    strSaveConfirm = MsgText(601)
                    MsgBox "此筆電腦收據資料已收款或已銷帳!!! 不可修改", vbExclamation + vbOKOnly
                    Exit Sub
               End If
               '93.12.16 END
               'add by nick 2004/08/19
               If Frmacc7100.DBOffice <> pub_strUserOffice And UCase(strUserDept) <> "M51" Then
                    strSaveConfirm = MsgText(601)
                    MsgBox "不能修改它所資料", , MsgText(5)
                    Exit Sub
               Else
                    Frmacc7100.oState = "2"
                    Frmacc7100_UnLock
                    'add by nick 2004/10/08
                    Frmacc7100.Command1.Enabled = False
               End If
'                'add by nick 2004/08/19
'                If Frmacc7100.DBOffice = pub_strUserOffice Or UCase(strUserDept) = "M51" Then
                   tool2_enabled
'                End If

            'Add by Morgan 2005/5/24
            Case "Frmacc41e0"
               If CheckUse("Frmacc41e0", strEdit) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               With Frmacc41e0
                  If .ReadData(.txtA2301) = True Then
                     .FormEnable "2"
                     If .txtNo.Enabled = True Then
                        .txtNo.SetFocus
                     ElseIf .txtA2309.Enabled = True Then
                        .txtA2309.SetFocus
                     End If
                     'adoTaie.BeginTrans 'adoTaie.BeginTrans 'Mark by Amy 2018/02/12 改存檔控制就好
                  Else
                     Exit Sub
                  End If
               End With
            'Add By Sindy 2013/12/19
            Case "Frmacc11p0"
               'Added by Lydia 2020/03/27 開放權限
               If Frmacc0000.str中所收據人員 <> "" And InStr(Frmacc0000.str中所收據人員, strUserNum) > 0 Then
                    Frmacc11p0.FormEnabled
               Else
               'end 2020/03/27
                    If CheckUse("Frmacc11p0", strEdit) = False Then
                       strSaveConfirm = MsgText(601)
                       Exit Sub
                    End If
                    Frmacc11p0.FormEnabled
               End If 'Added by Lydia 2020/03/27
            '2013/12/19 End
            'Add By Sindy 2012/9/4
            Case "Frmacc11n0"
               If CheckUse("Frmacc11n1", strEdit) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               Frmacc11n0.FormEnabled
               If Frmacc11n0.Text1 = "" Then
                  Frmacc11n0.Text1 = "X"
               End If
            '2012/9/4 End
            'add by sonia 2014/11/14
            Case "Frmacc21r0"
               If CheckUse("Frmacc21r0", strEdit) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               With Frmacc21r0
                  .Command1.Enabled = False
                  .txtKey.Locked = True
                  .txtBox(1).Locked = False
                  .txtBox(1).SetFocus
                  .txtInform(0).Locked = False
                  .txtInform(1).Locked = False
                  .SetCheck1 (True)
               End With
            'end 2014/11/14
            
         End Select
         tool2_enabled
      Case vbKeyF9 '存檔
         If strSaveConfirm = MsgText(601) Then
            Exit Sub
         End If
         Err.Clear
         Select Case strFormName
            'Added by  Lydia 2020/03/26 從account.aacc_var複製
            Case "Frmacc1130"
               Frmacc1130_Save
            Case "Frmacc1140"
               Frmacc1140.Frmacc1140_Save
            'end 2020/03/26
            Case "Frmacc7100"
                Frmacc7100_Save
                'add by nick 2004/08/19
                Frmacc7100.oState = ""
                'Frmacc7100_Lock
                'add by nick 2004/10/08
                Frmacc7100.Command1.Enabled = True
                'add by nick 2004/12/14
                Frmacc7100_UnLock
            'Add by Morgan 2005/5/24
            Case "Frmacc41e0"
               With Frmacc41e0
                  If .SaveData = True Then
                     'adoTaie.CommitTrans 'adoTaie.BeginTrans 'Mark by Amy 2018/02/12 改存檔控制就好
                     .MailCheck
                     .FormEnable
                  Else
                     Exit Sub
                  End If
               End With
            'Add By Sindy 2013/12/19
            Case "Frmacc11p0"
               Frmacc11p0.Frmacc11p0_Save
            '2013/12/19 End
            'Add By Sindy 2012/8/29
            Case "Frmacc11n0"
               Frmacc11n0.Frmacc11n0_Save
            '2012/8/29 End
            'add by sonia 2014/11/14
            Case "Frmacc21r0"
               With Frmacc21r0
               If .FormSave = True Then
                  strControlButton = MsgText(601)
                  .Command1.Enabled = True
                  .txtKey.Locked = False
                  .txtBox(1).Locked = False
                  .txtInform(0).Locked = True
                  .txtInform(1).Locked = True
                  .SetCheck1 (False)
               Else
                  strControlButton = MsgText(602)
               End If
               End With
         
         End Select
         If strControlButton <> MsgText(602) Then
            strSaveConfirm = MsgText(601)
            
            Select Case strFormName
               'Added by  Lydia 2020/03/26 從account.aacc_var複製
               Case "Frmacc1130"
                   tool14_enabled
               Case "Frmacc1140"
                   tool8_enabled
               'end 2020/03/26
               Case "Frmacc21r0"
                  tool3_enabled   '分所只能改該所資料故取消前後筆功能
                  Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
               'Add by Sindy 2015/7/9
               Case "Frmacc11p0"
                  'Modified by Lydia 2020/03/27 參考aacc_var
                  'If UCase(strUserLevel) = UCase("Frmacc44t0") Then
                  If UCase(strUserLevel) = UCase("Frmacc44t0") Or UCase(strUserLevel) = UCase("Frmacc11b0") Then
                     tool6_enabled
                     Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
                  Else
                     tool1_enabled
                  End If
               '2015/7/9 END
               Case Else
                  tool1_enabled
            End Select
            
            Frmacc0000.StatusBar1.Panels(1).Text = MsgText(17)
         End If
         strControlButton = MsgText(601)
      Case vbKeyF10 '取消
         If strSaveConfirm = MsgText(601) Then
            Exit Sub
         End If
         Select Case strFormName
            'Added by  Lydia 2020/03/26 從account.aacc_var複製
            Case "Frmacc1130"
                tool14_enabled
            Case "Frmacc1140"
                tool8_enabled
            'end 2020/03/26
            Case "Frmacc7100"
               'add by nick 2004/08/19
               Frmacc7100.oState = ""
               'Frmacc7100_Lock
               strSaveConfirm = MsgText(601)
               'add by nick 2004/12/08
               Frmacc7100_Clear
               'add by nick 2004/10/08
               Frmacc7100.Command1.Enabled = True
               'add by nick 2004/12/14
               Frmacc7100_UnLock
               tool1_enabled
               
            'Add by Morgan 2005/5/24
            Case "Frmacc41e0"
               'adoTaie.RollbackTrans 'adoTaie.BeginTrans 'Mark by Amy 2018/02/12 改存檔控制就好
               With Frmacc41e0
                  .FormClear False
                  .FormEnable
                  If .txtA2301 = "" Then .txtA2301 = .txtA2301.Tag
                  If .txtA2301 <> "" Then
                     .ReadData .txtA2301
                  End If
                  .txtA2301.SetFocus
               End With
               tool1_enabled
            'Add By Sindy 2013/12/19
            Case "Frmacc11p0"
               With Frmacc11p0
                  .AdodcRefresh
                  .FormDisabled
               End With
               'Add by Sindy 2015/7/9
               'Modified by Lydia 2020/03/27 參考aacc_var
               'If UCase(strUserLevel) <> UCase("Frmacc44t0") Then
               If UCase(strUserLevel) = UCase("Frmacc11b0") Or _
                   UCase(strUserLevel) = UCase("Frmacc44w1") Then
                   tool6_enabled
                   Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
               ElseIf UCase(strUserLevel) <> UCase("Frmacc44t0") Then
               'end 2020/03/27
               '2015/7/9 END
                  Frmacc11p0.Frmacc11p0_Clear
               End If 'Added by Lydia 2020/03/27
            '2013/12/19 End
            'Add By Sindy 2012/9/4
            Case "Frmacc11n0"
               With Frmacc11n0
                  .FormDisabled
               End With
               Frmacc11n0.Frmacc11n0_Clear
               tool1_enabled
            '2012/9/4 End
            'add by sonia 2014/11/14
            Case "Frmacc21r0"
               With Frmacc21r0
                  .Command1.Enabled = True
                  .txtKey.Locked = False
                  .txtBox(1).Locked = True
                  .txtInform(0).Locked = True
                  .txtInform(1).Locked = True
                  .SetCheck1 (False)
                  .FormShow
               End With
               tool3_enabled   '分所只能改該所資料故取消前後筆功能
            'end 2014/11/14
         End Select
         strSaveConfirm = MsgText(601)
'         Select Case strFormName
'            'Add by Morgan 2006/12/18
'            Case "Frmacc21r0"
'               tool6_enabled
'               Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
'               Frmacc21r0.FormShow
'            Case Else
'               tool1_enabled
'         End Select
         'Added by Lydia 2020/03/27
         Select Case strFormName
         Case "Frmacc11p0"
            If UCase(strUserLevel) <> UCase("Frmacc44w1") Then
               tool1_enabled
            End If
        End Select
        'end 2020/03/27
      Case vbKeyF5 '刪除
         If Frmacc0000.Toolbar1.Buttons.Item(8).Enabled = False Or strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
            Exit Sub
         End If
         'Added by Lydia 2020/03/27 參考aacc_var
        If strFormName = "Frmacc11p0" Then
           strDelConfirm = MsgBox("若為 收據抬頭已新建客戶 的資料刪除，請先仔細確認所有欄位是否於客戶檔都有設定！！！請確認完再刪除！" & vbCrLf & _
                     "是否確定要刪除？", vbOKCancel + vbDefaultButton2, MsgText(5))
        Else
           strDelConfirm = MsgBox(MsgText(6), vbOKCancel + vbDefaultButton2, MsgText(5))
        End If
        'end 2020/03/27
        
         Select Case strFormName
            'Added by  Lydia 2020/03/26 從account.aacc_var複製
             Case "Frmacc1140"
                 If CheckUse("Frmacc1140", strDel) = False Then
                    strSaveConfirm = MsgText(601)
                    Exit Sub
                 End If
            'end 2020/03/26
            'Added by Morgan 2018/11/12 分所簽收資料要請財務處刪除--辜
            Case "Frmacc41e0"
               strSaveConfirm = MsgText(601)
               MsgBox "請通知財務處人員刪除！", vbExclamation + vbOKOnly
               Exit Sub
            'end 2018/11/12
            
            Case "Frmacc7100"
               If CheckUse("Frmacc7100", strDel) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
               '93.12.16 ADD BY SONIA
               If Frmacc7100.M_REC = "Y" Then
                    strSaveConfirm = MsgText(601)
                    MsgBox "此筆電腦收據資料已收款或已銷帳!!! 不可刪除", vbExclamation + vbOKOnly
                    Exit Sub
               End If
               '93.12.16 END
               'add by nick 2004/10/11
               If Frmacc7100.DBOffice <> pub_strUserOffice And UCase(strUserDept) <> "M51" Then
                    strSaveConfirm = MsgText(601)
                    MsgBox "不能修改它所資料", , MsgText(5)
                    Exit Sub
               Else
                    'add by nick 2004/08/19
                    Frmacc7100.oState = "3"
                    'Frmacc7100_Lock
                End If
            'Add By Sindy 2013/12/19
            Case "Frmacc11p0"
               'Added by Lydia 2020/03/27 開放權限
               If Frmacc0000.str中所收據人員 <> "" And InStr(Frmacc0000.str中所收據人員, strUserNum) > 0 Then
               Else
                    If CheckUse("Frmacc11p0", strDel) = False Then
                       strSaveConfirm = MsgText(601)
                       Exit Sub
                    End If
               End If 'Added by Lydia 2020/
            '2013/12/19 End
            'Add By Sindy 2012/8/29
            Case "Frmacc11n0"
               If CheckUse("Frmacc11n1", strDel) = False Then
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
            '2012/8/29 End
         End Select
         strDelConfirm = MsgBox(MsgText(6), vbOKCancel + vbDefaultButton2, MsgText(5))
         If strDelConfirm = vbCancel Then
            Exit Sub
         End If
         Select Case strFormName
            'Added by  Lydia 2020/03/26 從account.aacc_var複製
            Case "Frmacc1140"
               Frmacc1140_Delete
               Frmacc1140.Frmacc1140_Clear 'Modify by Amy 2015/04/17 搬回form
               Frmacc1140.AdodcRefresh
            'end 2020/03/26
            Case "Frmacc7100"
               Frmacc7100_Delete
               Frmacc7100_Clear
            'Add By Sindy 2013/12/19
            Case "Frmacc11p0"
               Frmacc11p0.Frmacc11p0_Delete
               If strControlButton <> MsgText(602) Then 'Added by Lydia 2020/03/27
                    Frmacc11p0.Frmacc11p0_Clear
               End If
            '2013/12/19 End
            'Add By Sindy 2012/8/29
            Case "Frmacc11n0"
               Frmacc11n0.Frmacc11n0_Delete
               Frmacc11n0.Frmacc11n0_Clear
            '2012/8/29 End
         End Select
      Case vbKeyF4 '查詢
         If Frmacc0000.Toolbar1.Buttons.Item(9).Enabled = False Or strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
            Exit Sub
         End If
         strExitControl = MsgText(601)
         Select Case strFormName
            'Added by  Lydia 2020/03/26 從account.aacc_var複製
            Case "Frmacc1130"
               strFormLink = strFormName
               Frmacc1130_Clear
               Frmacc1130.Enabled = False
               Frmacc1131.Show
            Case "Frmacc1140"
               strFormLink = strFormName
               Frmacc1140.Frmacc1140_Clear 'Modify by Amy 2015/04/17 搬回form
               Frmacc1140.Enabled = False
               Frmacc1141.Show
            'end 2020/03/26
            Case "Frmacc7100"
               strFormLink = strFormName
               Frmacc7100_Clear
               Frmacc7100.Enabled = False
               'add by nick 2004/08/19
               Frmacc7100.oState = "4"
               'add by nick 2004/12/14
               Frmacc7100_UnLock
               'Frmacc7100_Lock
               Frmacc7101.Show
            'Add by Morgan 2005/5/24
            Case "Frmacc41e0"
               With Frmacc41e0
                  .Enabled = False
               End With
               Frmacc41e1.Show
            'Add By Sindy 2013/12/19
            Case "Frmacc11p0"
               Frmacc11p0.Frmacc11p0_Clear
               Frmacc11p0.Enabled = False
               Frmacc11p1.Show
            '2013/12/19 End
            'Add By Sindy 2012/8/29
            Case "Frmacc11n0"
               Frmacc11n0.Frmacc11n0_Clear
               Frmacc11n0.Enabled = False
               Frmacc11n1.Show
            '2012/8/29 End
        End Select
         tool3_enabled
      Case vbKeyF7
      Case vbKeyHome '第一筆
         FormMoveFirst
      Case vbKeyPageUp '上一筆
         FormMovePrevious
      Case vbKeyPageDown '下一筆
         FormMoveNext
      Case vbKeyEnd '最後一筆
         FormMoveLast
   End Select
Checking:
   If Err.Number = 0 Or Err.Number = -2147168242 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  資料庫連線
'
'*************************************************
Public Sub Main_C()
'   Set objOraSession = CreateObject("OracleInProcServer.XOraSession")
'   Set objOraDatabase = objOraSession.OpenDatabase(strOraDatabaseName, strOraConnect, 0&)
    Frmacc0000.MousePointer = vbHourglass
    If adoTaie.State <> adStateOpen Then
    
'Add by Morgan 2005/12/14 加連線選擇
'         adoTaie.ConnectionString = strAdoConnect
'         adoTaie.Open
'         Set cnnConnection = adoTaie
         Frmacc0000.Main7_0.Visible = True
         If PUB_Connect2DB() = False Then
            End
         Else
            Set adoTaie = cnnConnection
         End If
'2005/12/14 end

        'Removed by Morgan 2025/9/9 沒用了        
        'adoTemp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.path & "\finance.mdb"
        ''adoTemp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.3.51;Persist Security Info=False;Data Source=" & App.path & "\finance.mdb"
        'adoTemp.Open
        'end 2025/9/9

        'Add By Cheng 2003/05/02
        'edit by nickc 2007/02/09 不用 dll 了
        'Set objLawDll.Connection = cnnConnection
    End If
    strAccount = "1"
    Frmacc0000.MousePointer = vbDefault
End Sub

'*************************************************
'  清除狀態列內容
'
'*************************************************
Public Sub StatusClear()
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   Frmacc0000.StatusBar1.Panels(2).Text = MsgText(601)
End Sub

'*************************************************
'  訊息顯示
'
'*************************************************
Public Sub StatusView(strMessage As String)
   Frmacc0000.StatusBar1.Panels(1).Text = strMessage
End Sub

'*************************************************
'  選單為可使用狀態
'
'*************************************************
Public Sub MenuEnabled()
   With Frmacc0000
      .Main1.Enabled = True
      .Main2.Enabled = True
      .Main3.Enabled = True
      .Main4.Enabled = True  'Added by Lydia 2020/03/26 收據作業
      .Main5.Enabled = True
      .Main7.Enabled = True
      .Main6.Enabled = True  'add by sonia 2023/5/12
   End With
End Sub

'*************************************************
'  選單為不可使用狀態
'
'*************************************************
Public Sub MenuDisabled()
   With Frmacc0000
      .Main1.Enabled = False
      .Main2.Enabled = False
      .Main3.Enabled = False
      '.Main5.Enabled = False 'Removed by Morgan 2016/4/8
      .Main4.Enabled = False  'Added by Lydia 2020/03/26 收據作業
      .Main7.Enabled = False
      .Main6.Enabled = False  'add by sonia 2023/5/12
      
   End With
End Sub

'*************************************************
'  顯示資料筆數
'
'*************************************************
Public Sub CountShow(lngCurrent As Long, lngMax As Long)
   Frmacc0000.StatusBar1.Panels(2).Text = lngCurrent & MsgText(35) & lngMax
End Sub

'*************************************************
'  最後一筆
'
'*************************************************
Public Sub FormMoveLast()
   If Frmacc0000.Toolbar1.Buttons.Item(16).Enabled = False Or strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   Select Case strFormName
      'Added by Lydia 2020/03/26 收據作業
      Case "Frmacc1130"
         Frmacc1130_Last
      Case "Frmacc1140"
         Frmacc1140_Last
      'end 2020/03/26
      Case "Frmacc7100"
         Frmacc7100_Last
      'Add by Morgan 2005/5/24
      Case "Frmacc41e0"
         With Frmacc41e0
            .ReadData .txtA2301, 4
         End With
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         Frmacc11p0.Frmacc11p0_Last
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         Frmacc11n0.Frmacc11n0_Last
      '2012/8/29 End
   End Select
End Sub

'*************************************************
'  下一筆
'
'*************************************************
Public Sub FormMoveNext()
   If Frmacc0000.Toolbar1.Buttons.Item(15).Enabled = False Or strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   Select Case strFormName
      'Added by Lydia 2020/03/26 收據作業
      Case "Frmacc1130"
         Frmacc1130_Next
      Case "Frmacc1140"
         Frmacc1140_Next
      'end 2020/03/26
      Case "Frmacc7100"
         Frmacc7100_Next
      'Add by Morgan 2005/5/24
      Case "Frmacc41e0"
         With Frmacc41e0
            .ReadData .txtA2301, 3
         End With
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         Frmacc11p0.Frmacc11p0_Next
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         Frmacc11n0.Frmacc11n0_Next
      '2012/8/29 End
   End Select
End Sub

'*************************************************
'  上一筆
'
'*************************************************
Public Sub FormMovePrevious()
   If Frmacc0000.Toolbar1.Buttons.Item(14).Enabled = False Or strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   Select Case strFormName
      'Added by Lydia 2020/03/26 收據作業
      Case "Frmacc1130"
         Frmacc1130_Previous
      Case "Frmacc1140"
         Frmacc1140_Previous
      'end 2020/03/26
      Case "Frmacc7100"
         Frmacc7100_Previous
      'Add by Morgan 2005/5/24
      Case "Frmacc41e0"
         With Frmacc41e0
            .ReadData .txtA2301, 2
         End With
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         Frmacc11p0.Frmacc11p0_Previous
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         Frmacc11n0.Frmacc11n0_Previous
      '2012/8/29 End
   End Select
End Sub

'*************************************************
'  第一筆
'
'*************************************************
Public Sub FormMoveFirst()
   If Frmacc0000.Toolbar1.Buttons.Item(13).Enabled = False Or strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   Select Case strFormName
      'Added by Lydia 2020/03/26 收據作業
      Case "Frmacc1130"
         Frmacc1130_First
      Case "Frmacc1140"
         Frmacc1140_First
      'end 2020/03/26
      Case "Frmacc7100"
         Frmacc7100_First
      'Add by Morgan 2005/5/24
      Case "Frmacc41e0"
         With Frmacc41e0
            .ReadData .txtA2301, 1
         End With
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         Frmacc11p0.Frmacc11p0_First
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         Frmacc11n0.Frmacc11n0_First
      '2012/8/29 End
   End Select
End Sub

'*************************************************
'  離開
'
'*************************************************
Public Sub FormExit()
   Dim oForm
   If strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
       Exit Sub
   End If
   tool4_enabled
   If strFormName = "" Then
       Exit Sub
   End If
   Select Case strFormName
     'Added by Lydia 2020/03/26 收據作業
      Case "Frmacc1120"
         Unload Frmacc1120
      Case "Frmacc1121"
         Unload Frmacc1121
      Case "Frmacc1122"
         Unload Frmacc1122
      Case "Frmacc1130"
         Unload Frmacc1130
      Case "Frmacc1140"
         Unload Frmacc1140
      Case "Frmacc1141"
         Unload Frmacc1141
      Case "Frmacc1420"
         Unload Frmacc1420
      'end 2020/03/26
      Case "Frmacc1210"
          Unload Frmacc1210
      Case "Frmacc1211"
          Unload Frmacc1211
      Case "Frmacc1212"
          Unload Frmacc1212
      Case "Frmacc1220"
          Unload Frmacc1220
      Case "Frmacc1221"
          Unload Frmacc1221
      Case "Frmacc1222"
          Unload Frmacc1222
      Case "Frmacc1223"
          Unload Frmacc1223
      Case "Frmacc1224"
          Unload Frmacc1224
      Case "Frmacc1230"
          Unload Frmacc1230
      Case "Frmacc1240"
          Unload Frmacc1240
      Case "Frmacc1440"
          Unload Frmacc1440
      Case "Frmacc1450"
          Unload Frmacc1450
      Case "Frmacc1460"
          Unload Frmacc1460
      Case "Frmacc1490"
          Unload Frmacc1490
      Case "Frmacc14a0"
          Unload Frmacc14a0
      Case "Frmacc4250"
          Unload Frmacc4250
      Case "Frmacc7100"
          Unload Frmacc7100
      Case "Frmacc7101"
          Unload Frmacc7101
      Case "Frmacc7110"
          Unload Frmacc7110
      Case "Frmacc7111"
          Unload Frmacc7111
      Case "Frmacc7112"
          Unload Frmacc7112
      Case "Frmacc7120"
          Unload Frmacc7120
      Case "Frmacc7130"
          Unload Frmacc7130
      'Add by Morgan 2005/3/28
      Case "Frmacc44t0"
          Unload Frmacc44t0
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         Unload Frmacc11p0
      'Added by Lydia 2020/03/27
      Case "Frmacc11p1"
         Unload Frmacc11p1
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         Unload Frmacc11n0
      Case "Frmacc11n1"
         Unload Frmacc11n1
      '2012/8/29 End
      Case Else
         For Each oForm In Forms
            If oForm.Name = strFormName Then Unload oForm: Exit For
         Next
   End Select
End Sub
