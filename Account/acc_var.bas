Attribute VB_Name = "aacc_var"
'Memo by Morgan2010/8/19 日期欄已修改
'Modified by Morgan 2012/1/9 ACDate(ServerDate) -> strSrvDate(2), ServerDate -> strSrvDate(1)
Option Explicit
Public Const Pub_A2b05Begin = 1060531 'Added by Lydia 2017/05/11 財產目錄判斷補資料的日期
Public Const Pub_DBtype As String = "19D 專業技術支出"  'Added by Lydia 2017/09/14 與銀行結匯的匯款性質代號+說明
Public Const J_RMB As String = "CNY" 'Added by Lydia 2017/09/30 華銀整批結匯幣別RMB改成CNY

'*************************************************
'  新增

'*************************************************
Private Sub KeyEnterF2()
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
On Error GoTo Checking
   
   If Frmacc0000.Toolbar1.Buttons.Item(4).Enabled = False Or strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   strSaveConfirm = MsgText(3)
   Select Case strFormName
      Case "Frmacc1110"
         If CheckUse("Frmacc1110", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc1110_Clear
      Case "Frmacc1130"
         If CheckUse("Frmacc1130", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc1130_Clear
      'Add By Sindy 2014/1/9
      Case "Frmacc11q0"
         If CheckUse("Frmacc11q0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc11q0.Frmacc11q0_Clear
      '2014/1/9 END
      Case "Frmacc1140"
         If CheckUse("Frmacc1140", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc1140.Frmacc1140_Clear 'Modify by Amy 2015/04/17 搬回form
      Case "Frmacc1150"
         If CheckUse("Frmacc1150", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Add by Amy 2018/06/05 41字頭及7121不可輸小數
         If Frmacc1150.ChkDot = True Then
           tool1_enabled
           MsgBox "41字頭或7121科目不可輸入小數！", , MsgText(5)
           strSaveConfirm = MsgText(601)
           Exit Sub
         End If
        'end 2018/06/05
         If Frmacc1150.CreDebCheck <> MsgText(602) Then
            tool1_enabled
            MsgBox MsgText(11), , MsgText(5)
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc1150_Clear
         With Frmacc1150
            .AdodcClear
            If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
               .MaskEdBox1.Text = CFDate(strSrvDate(2))
               .Text2 = AutoNo(MsgText(803), 5)
            Else
               If Mid(.MaskEdBox1.Text, 1, 3) <> Mid(CFDate(strSrvDate(2)), 1, 3) Then
                  .Text2 = UpdateNo("acc0l0", "a0l01", 5, .MaskEdBox1.Text, MsgText(803))
               Else
                  .Text2 = AutoNo(MsgText(803), 5)
               End If
            End If
            .strDocNo = .Text2
            .Combo2.Clear
            .FormEnabled
            'Modified by Morgan 2014/1/2
            'If .Text4.Enabled Then .Text4.SetFocus
            If .Text19.Enabled Then .Text19.SetFocus
            .Command5.Enabled = True 'Added by Morgan 2014/1/27
         End With
         adoTaie.BeginTrans
      Case "Frmacc1160"
         If CheckUse("Frmacc1160", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc1160.Frmacc1160_Clear 'Modify by Amy 2022/03/11 原:Frmacc1160_Clear
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         If CheckUse("Frmacc11p0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc11p0.FormEnabled
         If strControlButton <> MsgText(602) Then 'Add By Sindy 2016/5/31 +if 保留畫面資料,只清收據抬頭欄
            Frmacc11p0.Frmacc11p0_Clear
            Frmacc11p0.textA4221 = "1" 'Add By Sindy 2016/12/2 繳款書寄件處預設為1客戶
         Else
            Frmacc11p0.textA4201.Text = ""
         End If
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         If CheckUse("Frmacc11n0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc11n0.FormEnabled
         Frmacc11n0.Frmacc11n0_Clear
         Frmacc11n0.Text4.Enabled = False
         Frmacc11n0.Command1.Enabled = False
      '2012/8/29 End
      'Add By Amy 2013/12/02
      Case "Frmacc11o0"
         If CheckUse("Frmacc11o0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc11o0.FormClear
         Frmacc11o0.FormEnabled
         Frmacc11o0.Text1(0).SetFocus
      Case "Frmacc1170"
         If CheckUse("Frmacc1170", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc1170
            .Combo1.Clear
            .Frmacc1170_Clear 'Modify by Amy 2013/12/26
         End With
         
         With Frmacc1170
            If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
               .MaskEdBox1.Text = CFDate(strSrvDate(2))
               .Text1 = AutoNo(MsgText(804), 5)
            Else
               If Mid(.MaskEdBox1.Text, 1, 3) <> Mid(CFDate(strSrvDate(2)), 1, 3) Then
                  .Text1 = UpdateNo("acc0o0", "a0o01", 5, .MaskEdBox1.Text, MsgText(804))
               Else
                  .Text1 = AutoNo(MsgText(804), 5)
               End If
            End If
            .strDocNo = .Text1
            .FormEnabled
         End With
         adoTaie.BeginTrans
      Case "Frmacc1180"
         If CheckUse("Frmacc1180", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc1180
            .FormEnabled
            .Frmacc1180_Clear (True) 'Modify by Amy 2014/01/28 +參數
         End With
         
         With Frmacc1180
            adoTaie.BeginTrans
            If .MaskEdBox7.Text = MsgText(601) Or .MaskEdBox7.Text = MsgText(29) Then
               .MaskEdBox7.Text = CFDate(strSrvDate(2))
               .Text17 = AccAutoNo(MsgText(818), 3, Mid(.MaskEdBox7.Text, 1, 3), Mid(.MaskEdBox7.Text, 5, 2))
            Else
               If Mid(.MaskEdBox7.Text, 1, 3) <> Mid(CFDate(strSrvDate(2)), 1, 3) Then
                  .Text17 = UpdateNo("acc0q0", "a0q17", 5, .MaskEdBox7.Text, MsgText(818))
               Else
                  .Text17 = AccAutoNo(MsgText(818), 3, Mid(.MaskEdBox7.Text, 1, 3), Mid(.MaskEdBox7.Text, 5, 2))
               End If
            End If
            .strDocNo = .Text17
            strCon9 = AccSaveAutoNo(MsgText(818), Mid(.Text17, 7, 3), Mid(.MaskEdBox7.Text, 1, 3), Mid(.MaskEdBox7.Text, 5, 2))
            adoTaie.CommitTrans
            If .MaskEdBox7.Text = MsgText(29) Or .MaskEdBox7.Text = MsgText(601) Then
               .MaskEdBox7.Mask = ""
               .MaskEdBox7.Text = CFDate(strSrvDate(2))
               .MaskEdBox7.Mask = DFormat
            End If
            If .MaskEdBox5.Text = MsgText(29) Or .MaskEdBox5.Text = MsgText(601) Then
               .MaskEdBox5.Mask = ""
               .MaskEdBox5.Text = CFDate(strSrvDate(2))
               .MaskEdBox5.Mask = DFormat
            End If
            If .MaskEdBox6.Text = MsgText(29) Or .MaskEdBox6.Text = MsgText(601) Then
               .MaskEdBox6.Mask = ""
               .MaskEdBox6.Text = ""
               .MaskEdBox6.Mask = DFormat
            End If
         End With
         adoTaie.BeginTrans
      Case "Frmacc1190"
         If CheckUse("Frmacc1190", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         strCon1 = ""
         strCon2 = ""
         strCon3 = ""
         strCon4 = ""
         With Frmacc1190
            .Frmacc1190_Clear
            If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
               .MaskEdBox1.Text = CFDate(strSrvDate(2))
               .Text2 = AutoNo(MsgText(805), 5)
            Else
               If Mid(.MaskEdBox1.Text, 1, 3) <> Mid(CFDate(strSrvDate(2)), 1, 3) Then
                  .Text2 = UpdateNo("acc0s0", "a0s01", 5, .MaskEdBox1.Text, MsgText(805))
               Else
                  .Text2 = AutoNo(MsgText(805), 5)
               End If
            End If
            .strDocNo = .Text2
            .cboCaseNo.Enabled = True 'Add by Morgan 2011/10/17
         End With
      Case "Frmacc11a0"
         If CheckUse("Frmacc11a0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc11a0.Frmacc11a0_Clear 'Modify by Amy 2014/10/29
         With Frmacc11a0
            If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
               .MaskEdBox1.Text = CFDate(strSrvDate(2))
               .Text1 = AutoNo(MsgText(806), 5)
            Else
               If Mid(.MaskEdBox1.Text, 1, 3) <> Mid(CFDate(strSrvDate(2)), 1, 3) Then
                  .Text1 = UpdateNo("acc0t0", "a0t01", 5, .MaskEdBox1.Text, MsgText(806))
               Else
                  .Text1 = AutoNo(MsgText(806), 5)
               End If
            End If
            .strDocNo = .Text1
            .Combo1.Clear
            .SetData ("F2") 'Add by Amy 2014/10/28
            .ObjectEnabled_1
         End With
         adoTaie.BeginTrans
         
      'Add by Morgan 2004/1/12
      '按新增清除畫面並設定Focus
      Case "Frmacc11d0"
         Frmacc11d0_Clear
         
      Case "Frmacc11f0"
         If CheckUse("Frmacc11f0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc11f0_Clear
         
      'Add by Morgan 2007/4/13
      Case "Frmacc11i0"
         Frmacc11i0.FormClear
         Frmacc11i0.FormEnable
         Frmacc11i0.txtA2502 = "4"
         Frmacc11i0.txtA2505 = "J"
         
      'Add by Morgan 2007/5/16
      Case "Frmacc11j0"
         With Frmacc11j0
            .FormClear
            .FormEnable
            .Text1.Text = "F5"
            .Text1.SetFocus
            .Text1.SelStart = 2
         End With
         
         
      'Add by Morgan 2007/10/4
      Case "Frmacc11k0"
         With Frmacc11k0
            .FormClear
            .FormEnable
            .txtCaseNo(1).SetFocus
         End With
         
      'Add by Morgan 2011/4/8
      Case "Frmacc11l0"
         If CheckUse("Frmacc11l0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc11l0
            .FormClear
            .FormEnable
            .txtCNo(1).SetFocus
         End With

      Case "Frmacc2110"
         If CheckUse("Frmacc2110", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         
         'Add by Morgan 2006/6/20
         If Frmacc2110.bolForm2 = True Then
            MsgBox "匯率有異動，請點【收款資料】以便重新計算台幣收款金額！", vbExclamation
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'end 2006/6/20
         
         Frmacc2110_Clear
         With Frmacc2110
            .AdodcClear
            If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
               .MaskEdBox1.Text = CFDate(strSrvDate(2))
               .Text2 = AutoNo(MsgText(808), 5)
            Else
               If Mid(.MaskEdBox1.Text, 1, 3) <> Mid(CFDate(strSrvDate(2)), 1, 3) Then
                  .Text2 = UpdateNo("acc0y0", "a0y01", 5, .MaskEdBox1.Text, MsgText(808))
               Else
                  .Text2 = AutoNo(MsgText(808), 5)
               End If
            End If
            .strDocNo = .Text2
            .FormEnabled
            .AdodcRefresh
            .SumShow
            If Val(.Text3) <> 0 Then
               .Text1.SetFocus
            End If
         End With
         
         adoTaie.BeginTrans
         
      Case "Frmacc2120"
         If CheckUse("Frmacc2120", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc2120_Clear
         With Frmacc2120
            If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
               .MaskEdBox1.Text = CFDate(strSrvDate(2))
               .Text2 = AutoNo(MsgText(809), 5)
            Else
               If Mid(.MaskEdBox1.Text, 1, 3) <> Mid(CFDate(strSrvDate(2)), 1, 3) Then
                  .Text2 = UpdateNo("acc120", "a1201", 5, .MaskEdBox1.Text, MsgText(809))
               Else
                  .Text2 = AutoNo(MsgText(809), 5)
               End If
            End If
            .strDocNo = .Text2
            .FormEnable MsgText(3) 'Add by Morgan 2011/3/10
         End With
         
         adoTaie.BeginTrans
         
      Case "Frmacc2130"
         If CheckUse("Frmacc2130", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc2130_Clear
         With Frmacc2130
            If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
               .MaskEdBox1.Text = CFDate(strSrvDate(2))
               .Text2 = AutoNo(MsgText(810), 5)
            Else
               If Mid(.MaskEdBox1.Text, 1, 3) <> Mid(CFDate(strSrvDate(2)), 1, 3) Then
                  .Text2 = UpdateNo("acc130", "a1301", 5, .MaskEdBox1.Text, MsgText(810))
               Else
                  .Text2 = AutoNo(MsgText(810), 5)
               End If
            End If
            .strDocNo = .Text2
         End With
      Case "Frmacc2140"
         If CheckUse("Frmacc2140", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc2140_Clear
         With Frmacc2140
            If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
               .MaskEdBox1.Text = CFDate(strSrvDate(2))
               .Text2 = AutoNo(MsgText(811), 5)
            Else
               If Mid(.MaskEdBox1.Text, 1, 3) <> Mid(CFDate(strSrvDate(2)), 1, 3) Then
                  .Text2 = UpdateNo("acc140", "a1401", 5, .MaskEdBox1.Text, MsgText(811))
               Else
                  .Text2 = AutoNo(MsgText(811), 5)
               End If
            End If
            .strDocNo = .Text2
            .m_Old_a1k01 = "" 'Add By Sindy 2011/8/10
         End With
      Case "Frmacc2150"
         If CheckUse("Frmacc2150", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc2150_Clear
         With Frmacc2150
             'Ken 92/01/03 改為編號依系統年度編號
'                  If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
               .Text2 = AutoNo(MsgText(812), 5)
'                  Else
'                     If Mid(.MaskEdBox1.Text, 1, 3) <> Mid(CFDate(strSrvDate(2)), 1, 3) Then
'                        .Text2 = UpdateNo("acc150", "a1501", 5, .MaskEdBox1.Text, MsgText(812))
'                     Else
'                        .Text2 = AutoNo(MsgText(812), 5)
'                     End If
'                  End If
'                  .strDocNo = .Text2
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
               .Text2 = AutoNo(MsgText(813), 5)
            Else
               If Mid(.MaskEdBox1.Text, 1, 3) <> Mid(CFDate(strSrvDate(2)), 1, 3) Then
                  .Text2 = UpdateNo("acc160", "a1601", 5, .MaskEdBox1.Text, MsgText(813))
               Else
                  .Text2 = AutoNo(MsgText(813), 5)
               End If
            End If
            .strDocNo = .Text2
            .FormEnabled
         End With
         adoTaie.BeginTrans
         
      'Added by Morgan 2023/3/30
      Case "Frmacc2172"
         If CheckUse("Frmacc2172", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc2172.FormClear
         Frmacc2172.FormEnabled
         Frmacc2172.MaskEdBox1.SetFocus
      'edn 2023/3/30
      
      Case "Frmacc21d0"
         If CheckUse("Frmacc21d0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc21d0.Frmacc21d0_Clear 'Modify by Amy 2014/11/04搬回form
         With Frmacc21d0
            .FormEnabled
         End With
         adoTaie.BeginTrans
      Case "Frmacc21e0"
         If CheckUse("Frmacc21e0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc21e0_Clear
         With Frmacc21e0
            'Modify by Amy 2014/11/06
            '.Combo1.Clear
            '.FormEnabled
            .SetData ("F2")
         End With
         adoTaie.BeginTrans
      Case "Frmacc21f0"
         If CheckUse("Frmacc21f0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc21f0_Clear
         With Frmacc21f0
            .Text9 = AutoNo(MsgText(817), 5)
            .FormEnabled
         End With
         adoTaie.BeginTrans
'Remove by Morgan 2005/1/14 財務不需要
'            Case "Frmacc21g0"
'               If CheckUse("Frmacc21g0", strAdd) = False Then
'                  strSaveConfirm = MsgText(601)
'                  Exit Sub
'               End If
'               Frmacc21g0_Clear
      Case "Frmacc21h0"
         If CheckUse("Frmacc21h0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc21h0_Clear
         With Frmacc21h0
            'Added by Morgan 2023/7/19
            '使用預留單號
            If .Check1.Value = 1 Then
               If .AddCheck = False Then
                 strSaveConfirm = MsgText(601)
                 Exit Sub
               End If
               .Text5 = .Text11
            Else
            'end 20236/62/8
            
               adoTaie.BeginTrans
               adoTaie.Execute "update acc1r0 set a1r04 = a1r04 where a1r01 = 'X'"
               .Text5 = AccAutoNo(MsgText(815), 5)
               strConTitle = AccSaveAutoNo(MsgText(815), Right(.Text5, 5))
               adoTaie.CommitTrans
            End If 'Added by Morgan 2023/7/19
            .Command1.Enabled = True
            .Command2.Enabled = False
            .Command3.Enabled = True
            .Command5.Enabled = False
         End With
         adoTaie.BeginTrans
      Case "Frmacc21j0"
         If CheckUse("Frmacc21j0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc21j0_Clear
      Case "Frmacc21k0"
         If CheckUse("Frmacc21k0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc21k0_Clear
         
'Remove by Morgan 2005/1/14 財務不用
'Modified by Morgan 2019/10/5 改回由財務維護
      Case "Frmacc21m0"
         If CheckUse("Frmacc21m0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc21m0_Clear Frmacc21m0
         
      Case "Frmacc21n0"
         If CheckUse("Frmacc21n0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc21n0_Clear
'Remove by Morgan 2005/8/3 財務不用
'Modified by Morgan 2019/7/10 又改要用
      Case "Frmacc21o0"
         If CheckUse("Frmacc21o0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc21o0_Clear Frmacc21o0

      'Add By Cheng 2003/07/22
      Case "Frmacc21q0"
         If CheckUse("Frmacc21q0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc21q0_Clear
         Frmacc21q0.Text14.Text = "2"
         Frmacc21q0.Combo1.Text = "Swift Code" 'Added by Lydia 2017/09/14 匯款銀行資料預設為Swift Code
         Frmacc21q0.Command1.Enabled = False
         
      'Added by Morgan 2019/10/5
      Case "Frmacc21s0"
         If CheckUse("Frmacc21s0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc21s0_Clear Frmacc21s0
         
      'Added by Lydia 2016/11/07
      Case "Frmacc21w0"
         If CheckUse("Frmacc21w0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc21w0
             .FormClear
             .Command1.Enabled = False
             .txtKey.Locked = False
             .txtKey.SetFocus
         End With

      Case "Frmacc3110"
         If CheckUse("Frmacc3110", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc3110.Frmacc3110_Clear
         With Frmacc3110
            .FormEnabled
            .SetData ("F2") 'Add by Amy 2014/11/12
         End With
         adoTaie.BeginTrans
      Case "Frmacc3120"
         If CheckUse("Frmacc3120", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc3120.Frmacc3120_Clear 'Modify by Amy 2020/07/14
         With Frmacc3120
            .FormEnabled
         End With
         adoTaie.BeginTrans
      Case "Frmacc3130"
         If CheckUse("Frmacc3130", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc3130_Clear
      Case "Frmacc3140"
         If CheckUse("Frmacc3140", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc3140.Frmacc3140_Clear 'Modify by Amy 2020/07/16
         '2005/12/13 CANCEL BY SONIA
         'With Frmacc3140
         '   .MaskEdBox1.Mask = MsgText(601)
         '   .MaskEdBox1.Text = CFDate(strSrvDate(2))
         '   .MaskEdBox1.Mask = DFormat
         'End With
      Case "Frmacc3150"
         If CheckUse("Frmacc3150", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc3150.Frmacc3150_Clear 'Modify by Amy 2020/07/17
      Case "Frmacc3160"
         If CheckUse("Frmacc3160", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc3160.Frmacc3160_Clear 'Modify by Amy 2020/07/17
      Case "Frmacc3170"
         If CheckUse("Frmacc3170", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Dim objaccnum As Object
         Frmacc3170.Frmacc3170_Clear 'Modify by Amy 2020/07/17
         With Frmacc3170
            .AdodcRefresh
            .AdodcClear
            .MaskEdBox1.Text = CFDate(strSrvDate(2))
            .adoaccnum.CursorLocation = adUseClient
            .adoaccnum.Open "select * from autonumber where au01 = '" & MsgText(816) & "'", adoTaie, adOpenStatic, adLockReadOnly
            If .adoaccnum.RecordCount = 0 Then
               .Text2 = ZeroBeforeNo("0", 3)
            Else
               If .adoaccnum.Fields("au02").Value <> Val(Mid(strSrvDate(1), 5, 4)) Then
                  .Text2 = ZeroBeforeNo("0", 3)
               Else
                  .Text2 = ZeroBeforeNo(str(.adoaccnum.Fields("au03").Value), 3)
               End If
            End If
            .adoaccnum.Close
            .FormEnabled
         End With
         
         adoTaie.BeginTrans
         
      Case "Frmacc3180"
         If CheckUse("Frmacc3180", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc3180_Clear
         Frmacc3180.FormEnabled True, True 'Add by Morgan 2007/2/7
         
      Case "Frmacc3190"
         If CheckUse("Frmacc3190", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc3190_Clear
      Case "Frmacc31a0"
         If CheckUse("Frmacc31a0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc31a0.Frmacc31a0_Clear 'Modify by Amy 2020/07/21
      Case "Frmacc31c0"
         If CheckUse("Frmacc31c0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc31c0.Frmacc31c0_Clear 'Modify by Amy 2020/07/21
      Case "Frmacc4110"
         If CheckUse("Frmacc4110", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc4110.Frmacc4110_Clear 'Modify by Amy 2015/06/11搬回form
      Case "Frmacc4120"
         If CheckUse("Frmacc4120", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Modify by Amy 2024/07/31 原:FormF2Check 整合檢查程式,避免有未改到的,並調整相關函數
         Call Frmacc4120.SetData("F2-1")
         'Memo by Amy 2024/07/31 日期先判斷避免彈訊息後仍可操作
         '  ex:查 D113062234 (傳票日 6/28 已月結)->改日期7/1->新增->彈 目前最大傳票日7/29->仍可操作
         If Frmacc4120.ChkA0205(True) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc4120.SetData ("F2-2") 'Modify by Amy 2014/11/17
         'end 2024/07/31
         adoTaie.BeginTrans
      Case "Frmacc4130"
         If CheckUse("Frmacc4130", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc4130_Clear
      Case "Frmacc4140"
         If CheckUse("Frmacc4140", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc4140_Clear
         With Frmacc4140
            .Text2 = MsgText(603)
         End With
      Case "Frmacc4150"
         If CheckUse("Frmacc4150", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc4150_Clear
'               With Frmacc4150
'                  .Text15 = "1"
'               End With
      Case "Frmacc4160"
         If CheckUse("Frmacc4160", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Modify by Amy 2024/08/23 原寫於 adoTaie.BeginTrans 後,檢查資料已存在不可用新增
         'Modified by Lydia 2017/01/26 新增時預設上一筆輸入的下一會計科目
         'Frmacc4160.Frmacc4160_Clear 'Modify by Amy 2013/12/24
         Frmacc4160.Frmacc4160_Clear True
         If strSaveConfirm = MsgText(601) Then
            Exit Sub
         End If
         Frmacc4160.FormEnabled
         'end 2024/08/19
         adoTaie.BeginTrans
      Case "Frmacc4170"
         If CheckUse("Frmacc4170", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc4170
            .CreDebCheck
            If .CreDebCheck <> MsgText(602) Then
               MsgBox MsgText(11), , MsgText(5)
               strSaveConfirm = MsgText(601)
               Exit Sub
            End If
            .Combo1.Clear
            .Frmacc4170_Clear 'Modify by Amy 2013/12/24
            .AdodcClear
            .AdodcRefresh
            'Modify by Amy 2013/12/24 +if
            If strSrvDate(1) >= InvoiceStartDate Then
                .Text1 = ""
            Else
                .Text1 = "1"
            End If
            'end 2013/12/24
            .FormEnabled
            .Text6 = ZeroBeforeNo(MsgText(12), 3)
            .SumShow
            .Text1.SetFocus
         End With
         adoTaie.BeginTrans
      Case "Frmacc4180"
         If CheckUse("Frmacc4180", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc4180_Clear
      Case "Frmacc4190"
         If CheckUse("Frmacc4190", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Modify by Amy 2024/08/12 原程式搬回表單中
         Frmacc4190.SetData ("F2")
         adoTaie.BeginTrans
      Case "Frmacc41a0"
         If CheckUse("Frmacc41a0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Frmacc41a0_Clear 'Modify by Amy 2013/12/17
         With Frmacc41a0
            .FormClear 'Add 2013/12/17
            .FormEnabled
         End With
         'add by nickc 2008/03/12
         Frmacc41a0.Text10 = "R"
         Frmacc41a0.Text10.SelStart = 1
         adoTaie.BeginTrans
      Case "Frmacc41b0"
         If CheckUse("Frmacc41b0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Frmacc41b0_Clear 'Modify by Amy 2013/12/17
         With Frmacc41b0
            .FormClear 'Add 2013/12/17
            .AdodcRefresh
            .SumShow
         End With
      Case "Frmacc41d0"
         If CheckUse("Frmacc41d0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc41d0.FormF2Set  'Modify by Amy 2014/02/06 原程式搬回form
         adoTaie.BeginTrans
      'Add by Morgan 2005/4/6
      Case "Frmacc41e0"
         If CheckUse("Frmacc41e0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc41e0
            .txtA2301.Tag = .txtA2301
            .FormClear
            .FormEnable "1"
            '.txtNo.SetFocus 'Removed by Morgan 2015/7/20 改智權繳款
         End With
         'adoTaie.BeginTrans 'Removed by Morgan 2018/2/9 改存檔控制就好
      'Added by Lydia 2017/03/03
      Case "Frmacc41i0" '財產目錄
         If CheckUse("Frmacc41i0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         
         With Frmacc41i0
            .Frmacc41i0_Clear
            If .MaskEdBox1.Text = MsgText(601) Or .MaskEdBox1.Text = MsgText(29) Then
               .MaskEdBox1.Text = CFDate(strSrvDate(2))
            End If
            If .MaskEdBox2.Text = MsgText(601) Or .MaskEdBox2.Text = MsgText(29) Then
               .MaskEdBox2.Text = CFDate(strSrvDate(2))
            End If
            .txtA2B01 = .GetA2b01No
            .Acc2b0Refresh
            .AdodcClear
            .AdodcRefresh
            .FormEnabled
            .txtA2B16.SetFocus
         End With
         adoTaie.BeginTrans
      'end 2017/03/03
      'Added by Lydia 2017/05/15
      Case "Frmacc41i0_1" '財產作廢作業
         If CheckUse("Frmacc41i0_1", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc41i0_1
            .Frmacc41i0_1_Clear
            .AdodcClear
            .AdodcRefresh
            .FormEnabled
            .txtA2B01.SetFocus
         End With
         adoTaie.BeginTrans
      'end 2017/05/15
      'Add by Amy 2017/04/21
      Case "Frmacc41j0"
        If CheckUse("Frmacc41j0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc41j0
            If .FormCheck(0, "F2") = False Then
                strSaveConfirm = MsgText(601) 'Add by Amy 2017/06/07
                Exit Sub
            End If
            .SetData ("F2")
         End With
      Case "Frmacc5200"
         If CheckUse("Frmacc5200", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc5200_Clear
         With Frmacc5200
            .QueryTable
            .Command1.Enabled = True
            .DataGrid1.AllowUpdate = True
         End With
         adoTaie.BeginTrans
   End Select
   'Modify by Morgan 2004/1/12
   'tool2_enabled
   Select Case strFormName
      Case "Frmacc11d0"
         strSaveConfirm = MsgText(601)
      Case Else
         tool2_enabled
   End Select
   
Checking:
   If Err.Number = 0 Or Err.Number = -2147168242 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   
End Sub

'*************************************************
'  修改

'*************************************************
Private Sub KeyEnterF3()
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim stMsg As String
   
On Error GoTo Checking
         
   If Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = False Or strSaveConfirm = MsgText(3) Then
      Exit Sub
   End If
   Select Case strFormName
      Case "Frmacc1150"
         If CheckUse("Frmacc1150", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc1150
            'Add by Morgan 2005/8/12
            If .Text2 = "" Then
               MsgBox "收款單號不可空白！", vbExclamation
               strSaveConfirm = MsgText(601)
               Exit Sub
            Else
               .Text2.Tag = .Text2
               .RefreshData
               If .Text2 <> .Text2.Tag Then
                  MsgBox "收款單號已異動，請檢查資料是否正確！", vbExclamation
                  strSaveConfirm = MsgText(601)
                  Exit Sub
               End If
            End If
            '2005/8/12 end
            
            .adoaccsum.CursorLocation = adUseClient
            'Modified by Morgan 2014/6/24 會有1或J公司兩家
            'modify by sonia 2020/4/24
            '.m_A1P22_1 = "": .m_A1P22_J = ""
            .strA1P01s = ""
            .strA1P22s = ""
            'end 2020/4/24
            '.adoaccsum.Open "select a1p22 from acc1p0 where a1p01 = '1' and a1p02 = 'A' and a1p04 = '" & .Text2 & "' and a1p22 is not null", adoTaie, adOpenStatic, adLockReadOnly
            .adoaccsum.Open "select distinct a1p22,a1p01 from acc1p0 where a1p02 = 'A' and a1p04 = '" & .Text2 & "' and a1p22 is not null", adoTaie, adOpenStatic, adLockReadOnly
            If .adoaccsum.RecordCount <> 0 Then
               'strCon10 = .adoaccsum.Fields("a1p22").Value
               Do While Not .adoaccsum.EOF
                  'modify by sonia 2020/4/24 考慮會有3家作帳公司別
                  'If .adoaccsum.Fields("a1p01") = "1" Then
                  '   .m_A1P22_1 = .adoaccsum.Fields("a1p22")
                  'ElseIf .adoaccsum.Fields("a1p01") = "J" Then
                  '   .m_A1P22_J = .adoaccsum.Fields("a1p22")
                  'End If
                  .strA1P01s = .strA1P01s & .adoaccsum.Fields("a1p01") & ";"
                  .strA1P22s = .strA1P22s & .adoaccsum.Fields("a1p22") & ";"
                  'end 2020/4/24
                  .adoaccsum.MoveNext
               Loop
            'Else
            '   strCon10 = ""
            End If
            'end 2014/6/24
            
            .adoaccsum.Close
            .FormEnabled
         End With
         
         adoTaie.BeginTrans
         
      Case "Frmacc1170"
         If CheckUse("Frmacc1170", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc1170
            'Add by Morgan 2004/9/27
            If .EditCheck = False Then
               strSaveConfirm = MsgText(601)
               Exit Sub
            End If
            .FormEnabled
         End With
         adoTaie.BeginTrans
      Case "Frmacc1180"
         If CheckUse("Frmacc1180", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc1180
         
            'Added by Morgan 2023/12/19 判斷有傳票且已過帳就不可修改
            If .adoquery.State = adStateOpen Then
               .adoquery.Close
            End If
            .adoquery.CursorLocation = adUseClient
            .adoquery.Open "select ax210 from acc1p0, acc021 where a1p01 = ax201 and a1p22 = ax202 and a1p03 = ax203 and ax210 is not null and a1p04 = '" & .Text17 & "'", adoTaie, adOpenStatic, adLockReadOnly
            If .adoquery.RecordCount <> 0 Then
               MsgBox MsgText(14), , MsgText(5)
               strControlButton = MsgText(602)
               .adoquery.Close
               Exit Sub
            End If
            .adoquery.Close
            'end 2023/12/19
               
            .SetData ("F3") 'Add by Amy 2014/10/27
            .FormEnabled
         End With
         adoTaie.BeginTrans
      Case "Frmacc1190"
         'Add by Morgan 2004/9/29
         If Frmacc1190.EditCheck = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         
         strCon1 = ""
         strCon2 = ""
         strCon3 = ""
         strCon4 = ""
         Frmacc1190.cboCaseNo.Enabled = True 'Add by Morgan 2011/10/17
         
      Case "Frmacc11a0"
         If CheckUse("Frmacc11a0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc11a0
            'Add by Morgan 2005/10/19 檢查收否已沖
            'Modified by Lydia 2024/11/28 +StrSQLa回傳:僅開放修改客戶欄位
            If .CheckUsed(StrSQLa) = True Then
               strSaveConfirm = MsgText(601)
               Exit Sub
            End If
            'Added by Lydia 2024/11/28
            If StrSQLa = "1" Then  '1-僅開放修改客戶欄位
               .SetData ("F3")
               .ObjectEnabled_3
            Else
            'end 2024/11/28
               .SetData ("F3") 'Add by Amy 2014/10/28
               .ObjectEnabled_1
            End If
         End With
         adoTaie.BeginTrans
      'Add by Morgan 2005/6/7
      Case "Frmacc11h0"
         If CheckUse("Frmacc11h0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc11h0
            If .EditCheck = False Then
               strSaveConfirm = MsgText(601)
               Exit Sub
            End If
            .FormLocked True
            .txtA0w04.SetFocus
         End With
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         If CheckUse("Frmacc11p0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc11p0.FormEnabled
      '2013/12/19 End
      'Add By Sindy 2012/9/4
      Case "Frmacc11n0"
         If CheckUse("Frmacc11n0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc11n0.FormEnabled
         If Frmacc11n0.Text1 = "" Then
            Frmacc11n0.Text1 = "X"
         End If
      '2012/9/4 End
      Case "Frmacc2110"
         If CheckUse("Frmacc2110", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc2110
            'Added by Morgan 2021/5/27 目前會有1或L兩家公司
            .strA1P01s = ""
            .strA1P22s = ""
            'end 2021/5/27
            'Add by Morgan 2005/4/27 檢查是否已過帳
            If .AX210Exist = True Then Exit Sub
            
            .adoaccsum.CursorLocation = adUseClient
            'Added by Morgan 2021/5/27 目前會有1或L兩家公司
            '.adoaccsum.Open "select a1p22 from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p04 = '" & .Text2 & "' and a1p22 is not null", adoTaie, adOpenStatic, adLockReadOnly
            'If .adoaccsum.RecordCount <> 0 Then
            '   strCon10 = .adoaccsum.Fields("a1p22").Value
            'Else
            '   strCon10 = ""
            .adoaccsum.Open "select distinct a1p22,a1p01 from acc1p0 where  a1p02 = 'F' and a1p04 = '" & .Text2 & "' and a1p22 is not null", adoTaie, adOpenStatic, adLockReadOnly
            If .adoaccsum.RecordCount <> 0 Then
               Do While Not .adoaccsum.EOF
                  .strA1P01s = .strA1P01s & .adoaccsum.Fields("a1p01") & ";"
                  .strA1P22s = .strA1P22s & .adoaccsum.Fields("a1p22") & ";"
                  .adoaccsum.MoveNext
               Loop
            'end 2021/5/27
            End If
            .adoaccsum.Close
            .FormEnabled
            .Text3.Tag = .Text3.Text 'Add by Morgan 2006/6/20 紀錄原來匯率
         End With
         adoTaie.BeginTrans
      Case "Frmacc2120"
         If CheckUse("Frmacc2120", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
          'Add By Cheng 2004/04/21
          '若傳票已過帳, 不可修改國外暫收款資料
          With Frmacc2120
              StrSQLa = "Select Count(*) From ACC120, ACC1P0, ACC021 Where A1201=A1P04 And A1P22=AX202 And A1201='" & .Text2.Text & "' And AX210 Is Not Null "
              rsA.CursorLocation = adUseClient
              rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
              If rsA.RecordCount > 0 Then
                  If Val("" & rsA.Fields(0).Value) > 0 Then
                      MsgBox "傳票已過帳不可修改國外暫收款資料!!!", vbExclamation + vbOKOnly
                      strSaveConfirm = ""
                      If rsA.State <> adStateClosed Then rsA.Close
                      Set rsA = Nothing
                      Exit Sub
                  End If
              End If
              If rsA.State <> adStateClosed Then rsA.Close
              Set rsA = Nothing
              .FormEnable MsgText(4) 'Add by Morgan 2011/3/10
          End With
          'End
         adoTaie.BeginTrans
      'add by sonia 2017/1/16
      Case "Frmacc2130"
         If CheckUse("Frmacc2130", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'add by sonia 2017/1/16
         '若已產生國外付款資籵, 不可修改國外暫收款退費資料
         With Frmacc2130
         
           StrSQLa = "Select Count(*) From ACC170 Where A1702='" & .Text2.Text & "' "
           rsA.CursorLocation = adUseClient
           rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
           If rsA.RecordCount > 0 Then
               If Val("" & rsA.Fields(0).Value) > 0 Then
                   MsgBox MsgText(34), , MsgText(5)
                   strSaveConfirm = ""
                   If rsA.State <> adStateClosed Then rsA.Close
                   Set rsA = Nothing
                   Exit Sub
               End If
           End If
           If rsA.State <> adStateClosed Then rsA.Close
           Set rsA = Nothing
         End With
      'end 2017/1/16
      Case "Frmacc2150"
         If CheckUse("Frmacc2150", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc2150
            .FormEnabled
         End With
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
      Case "Frmacc21d0"
         If CheckUse("Frmacc21d0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc21d0
            .FormEnabled
         End With
         adoTaie.BeginTrans
      Case "Frmacc21e0"
         If CheckUse("Frmacc21e0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc21e0
            'Modify by Amy 2014/11/06
             '.FormEnabled
            .SetData ("F3")
         End With
         adoTaie.BeginTrans
      Case "Frmacc21f0"
         If CheckUse("Frmacc21f0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc21f0
            .FormEnabled
         End With
         adoTaie.BeginTrans
      Case "Frmacc21f1"
         If CheckUse("Frmacc21f10", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc21f1
            .FormEnabled
         End With
         adoTaie.BeginTrans
      Case "Frmacc21f2"
         If CheckUse("Frmacc21f2", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc21f2
            .FormEnabled
         End With
         adoTaie.BeginTrans
      Case "Frmacc21h0"
         If CheckUse("Frmacc21h0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc21h0
            .Command1.Enabled = True
            .Command2.Enabled = False
            .Command3.Enabled = True
            .Command5.Enabled = False
         End With
         adoTaie.BeginTrans
      Case "Frmacc21i0"
         If CheckUse("Frmacc21i0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc21i0
            .MaskEdBox2.Mask = ""
            .MaskEdBox2.Text = CFDate(strSrvDate(2))
            .MaskEdBox2.Mask = DFormat
         End With
      'Add By Cheng 2003/07/22
      Case "Frmacc21q0"
         If CheckUse("Frmacc21q0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      'Add by Morgan 2006/12/18
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
            .txtBox(5).Locked = False 'Added by Lydia 2018/07/20 財務信箱(CF)
            'Add by Morgan 2007/3/3
            .txtInform(0).Locked = False
            .txtInform(1).Locked = False
            'end 2007/3/3
            'Modify by Amy 2014/04/03 原程式改至form
            .SetCheck1 (True)
         End With
      'Added by Lydia 2016/11/07
      Case "Frmacc21w0"
         If CheckUse("Frmacc21w0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc21w0
             .Command1.Enabled = False
             .txtKey.Locked = True
         End With
      Case "Frmacc3110"
         If CheckUse("Frmacc3110", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc3110
            'Add by Morgan 2005/3/8 要檢查是否已過帳
            If .Adodc2.Recordset.RecordCount <> 0 Then
               If IsNull(.Adodc2.Recordset.Fields("a1p22").Value) = False Then
                  .adoquery.CursorLocation = adUseClient
                  .adoquery.Open "select ax210 from acc021 where ax201 = '" & .Adodc2.Recordset.Fields("a1p01").Value & "' and ax202 = '" & .Adodc2.Recordset.Fields("a1p22").Value & "' AND AX210 IS NOT NULL", adoTaie, adOpenStatic, adLockReadOnly
                  If .adoquery.RecordCount <> 0 Then
                     MsgBox MsgText(155), , MsgText(5)
                     .Text11.SetFocus
                     .adoquery.Close
                     Exit Sub
                  End If
                  .adoquery.Close
               End If
            End If
            '2005/3/8 end
            .FormEnabled
         End With
         adoTaie.BeginTrans
      Case "Frmacc3120"
         If CheckUse("Frmacc3120", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc3120
            'Add by Morgan 2005/3/8 要檢查是否已過帳
            If .Adodc2.Recordset.RecordCount <> 0 Then
               If IsNull(.Adodc2.Recordset.Fields("a1p22").Value) = False Then
                  .adoquery.CursorLocation = adUseClient
                  .adoquery.Open "select ax210 from acc021 where ax201 = '" & .Adodc2.Recordset.Fields("a1p01").Value & "' and ax202 = '" & .Adodc2.Recordset.Fields("a1p22").Value & "' AND AX210 IS NOT NULL", adoTaie, adOpenStatic, adLockReadOnly
                  If .adoquery.RecordCount <> 0 Then
                     MsgBox MsgText(155), , MsgText(5)
                     .Text12.SetFocus
                     .adoquery.Close
                     Exit Sub
                  End If
                  .adoquery.Close
               End If
            End If
            '2005/3/8 end
            .FormEnabled
         End With
         adoTaie.BeginTrans
      Case "Frmacc3170"
         If CheckUse("Frmacc3170", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc3170
            .FormEnabled
         End With
         adoTaie.BeginTrans
      'Add by Morgan 2007/2/7
      Case "Frmacc3180"
         If CheckUse("Frmacc3180", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc3180
            .FormEnabled
         End With
      Case "Frmacc31c0"
         If CheckUse("Frmacc31c0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc31c0
            If IsNull(.Adodc1.Recordset.Fields("A0E24").Value) = False Then
               MsgBox MsgText(86), , MsgText(5)
               Exit Sub
            End If
         End With
      Case "Frmacc4120"
         If CheckUse("Frmacc4120", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Modify by Amy 2014/11/17
         '.FormEnabled
         'Modify by Amy 2022/05/13 bug-系統產生會彈訊息,但仍可操作
         'Modify by Amy 2024/07/31 原:SetData 整合檢查程式,避免有未改到的
         If Frmacc4120.ChkForm("F3") = False Then
            Exit Sub
         End If
         Call Frmacc4120.SetData("F3")
         'end 2022/05/13
         adoTaie.BeginTrans
      Case "Frmacc4190"
         If CheckUse("Frmacc4190", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Add by Amy 2024/08/15 +FormCheck
         If Frmacc4190.FormCheck("F3") = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Modify by Amy 2024/08/12 原程式搬回表單中
         Frmacc4190.SetData ("F3")
         adoTaie.BeginTrans
      Case "Frmacc41d0"
         If CheckUse("Frmacc41d0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc41d0
            'Add by Morgan 2007/1/9
            If Left(.Text2, 1) = "I" And IsNumeric(Mid(.Text2, 2)) Then
               MsgBox "不可修改銷退的分錄資料！"
               Exit Sub
            End If
            'End 2007/1/9
            .FormEnabled
         End With
         adoTaie.BeginTrans
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
               'adoTaie.BeginTrans 'Removed by Morgan 2018/2/9 改存檔控制就好
            Else
               Exit Sub
            End If
         End With
      'Added by Lydia 2017/03/03
      Case "Frmacc41i0"  '財產目錄
         If CheckUse("Frmacc41i0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc41i0
             If IsEmptyText(.txtA2B01) Then
                MsgBox "請輸入財產編號!", vbCritical
                Exit Sub
             Else
               If .adoacc2b0.RecordCount <> 0 Then
                  .adoacc2b0.Find "a2b01 = '" & .txtA2B01 & "'", 0, adSearchForward, 1
                  If .adoacc2b0.EOF = True Then
                     Exit Sub
                  End If
               Else
                  MsgBox "請輸入財產編號後,按Enter鍵查詢!", vbCritical
                  Exit Sub
               End If
             End If
         End With
      'end 2017/03/03
      'Added by Lydia 2017/03/08
      Case "Frmacc41i0_1" '財產報廢作業
         If CheckUse("Frmacc41i0_1", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc41i0_1
             If IsEmptyText(.txtA2B01) Then
                MsgBox "請輸入財產編號!", vbCritical
                Exit Sub
             Else
               If .adoacc2b0.RecordCount <> 0 Then
                  .adoacc2b0.Find "a2b01 = '" & .txtA2B01 & "'", 0, adSearchForward, 1
                  If .adoacc2b0.EOF = True Then
                     Exit Sub
                  End If
               Else
                  MsgBox "請輸入財產編號後,按Enter鍵查詢!", vbCritical
                  Exit Sub
               End If
             End If
         End With
      'end 2017/03/08
      Case "Frmacc5200"
         If CheckUse("Frmacc5200", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc5200
            .DataGrid1.AllowUpdate = True
         End With
         adoTaie.BeginTrans
   End Select
   strSaveConfirm = MsgText(4)
   Select Case strFormName
      'Add by Morgan 2007/4/16
      Case "Frmacc11i0"
         If CheckUse("Frmacc11i0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         If Frmacc11i0.FormCheck = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         Else
            Frmacc11i0.FormEnable
         End If
      'Add by Morgan 2007/5/16
      Case "Frmacc11j0"
         If CheckUse("Frmacc11j0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         If Frmacc11j0.FormCheck = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc11j0.FormEnable
         
      'Add by Morgan 2007/10/5
      Case "Frmacc11k0"
         If CheckUse("Frmacc11k0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         If Frmacc11k0.FormCheck = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc11k0.FormEnable
         
         
      'Add by Morgan 2011/4/11
      Case "Frmacc11l0"
         If CheckUse("Frmacc11l0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc11l0
            If .EditCheck = False Then
               strSaveConfirm = MsgText(601)
               Exit Sub
            End If
            .FormEnable
            .txtCNo1(1).SetFocus
         End With
        'Add By Amy 2013/12/02
      Case "Frmacc11o0"
         If CheckUse("Frmacc11o0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc11o0.FormEnabled
      
      'Added by Morgan 2023/3/30
      Case "Frmacc2172"
         If CheckUse("Frmacc2172", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         If Frmacc2172.EditCheck = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc2172.FormEnabled
      'end 2023/3/30
      
      Case "Frmacc21d0"
         Frmacc21d0.AdodcRefresh
      Case "Frmacc41a0"
         If CheckUse("Frmacc41a0", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         
         'Add by Morgan 2011/6/23
         If Frmacc41a0.EditCheck = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'end 2011/6/23
         
         'edit by nickc 2005/11/04
         'Frmacc41a0_Clear
         With Frmacc41a0
            .FormEnabled
         End With
         adoTaie.BeginTrans
      'Modify by Amy 2013/12/24 由Frmacc4190前搬過來(為抓strSaveConfirm)
      Case "Frmacc4160"
         If CheckUse("Frmacc4160", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Add by Amy 2024/08/23
         If Frmacc4160.Acc040Check("F3") = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc4160.FormEnabled
         adoTaie.BeginTrans
      Case "Frmacc4170"
         If CheckUse("Frmacc4170", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc4170
            .FormEnabled
         End With
         adoTaie.BeginTrans
      'Added by Lydia 2017/03/03
      Case "Frmacc41i0" '財產目錄
         With Frmacc41i0
            .FormEnabled
         End With
         adoTaie.BeginTrans
      'end 2017/03/03
      'Added by Lydia 2017/03/08
      Case "Frmacc41i0_1"  '財產作廢作業
         With Frmacc41i0_1
            .FormEnabled
         End With
         adoTaie.BeginTrans
      'end 2017/03/08
      'Add by Amy 2017/04/28
      Case "Frmacc41j0"
        If CheckUse("Frmacc41j0", strAdd) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc41j0
            If .FormCheck(0, "F3") = False Then
                Exit Sub
            End If
            .SetData ("F3")
         End With
     'Add by Amy 2014/02/14
      Case "Frmacc5100"
         If CheckUse("Frmacc5100", strEdit) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         adoTaie.BeginTrans
         Frmacc5100.FormEnabled
   End Select
   
   tool2_enabled
   
Checking:
   If Err.Number = 0 Or Err.Number = -2147168242 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   
End Sub

'*************************************************
'  存檔

'*************************************************
Private Sub KeyEnterF9()
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim Cancel As Boolean 'Add By Sindy 2014/2/21
   
On Error GoTo Checking

   Select Case strFormName
      Case "Frmacc5100"
        'Modify by Amy 2014/02/14
        If Frmacc5100.FormCheck = False Then
            strControlButton = MsgText(602)
            Exit Sub
        Else
            strControlButton = MsgText(601)
        End If
        If strSaveConfirm = MsgText(4) Then
            Frmacc5100.Frmacc5100_Save
        End If
        If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
        End If
        'end 2014/02/14
      'Add by Morgan 2004/8/27
      Case "Frmacc31e0"
         Frmacc31e0.Frmacc31e0_Save 'Modify by Amy 2014/11/05 搬回form
      
      'Add by Morgan 2008/9/23
      Case "Frmacc31f0"
         Frmacc31f0.Frmacc31f0_Save 'Modify by Amy 2014/11/05 搬回form
         
      'Add by Morgan 2011/6/2
      Case "Frmacc31g0"
         Frmacc31g0.Frmacc31g0_Save 'Modify by Amy 2014/11/05 搬回form
         
      'Add by Morgan 2012/10/11
      Case "Frmacc31h0"
         Frmacc31h0.Frmacc31h0_Save
         
      'Add By Sindy 2013/12/13
      Case "Frmacc11o5"
         Frmacc11o5.Frmacc11o5_Save
   End Select
   If strSaveConfirm = MsgText(601) Then
      Exit Sub
   End If
   Err.Clear
   Select Case strFormName
      Case "Frmacc1110"
         Frmacc1110_Save
         With Frmacc1110
            .PrintDoc .Text4, .Text5, 1
            .PrintDoc .Text6, .Text7, 2
            .PrintDoc .Text10, .Text11, 3
         End With
         Frmacc1110_Clear
      Case "Frmacc1130"
         Frmacc1130_Save
      'Add By Sindy 2014/1/9
      Case "Frmacc11q0"
         Frmacc11q0.Frmacc11q0_Save
      '2014/1/9 END
      Case "Frmacc1140"
         Frmacc1140.Frmacc1140_Save
      Case "Frmacc1150"
         With Frmacc1150
            'Add By Sindy 2013/2/4
            If .Adodc1.Recordset.RecordCount <= 0 Then
               MsgBox "尚未Insert資料!!!", , MsgText(5)
               Exit Sub
            End If
            '2013/2/4 End
            
            strControlButton = MsgText(601) 'Added by Morgan 2014/8/27
            'Added by Morgan 2013/12/26
            If .SaveCheck = False Then
               strControlButton = MsgText(602)
               Exit Sub
            End If
            'end 2013/12/26
            
            If strSaveConfirm = MsgText(4) Then
               .Frmacc1150_Save 'Modify by Amy 2020/06/30
            End If
            If strControlButton <> MsgText(602) Then
               .FormDisabled
               .Command1.SetFocus
               adoTaie.CommitTrans
            End If
         End With
      Case "Frmacc1160"
         'Add by Amy 2023/05/15
         If Frmacc1160.TxtValidate = False Then
                strControlButton = MsgText(602)
               Exit Sub
         End If
         Frmacc1160_Save
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         Frmacc11p0.Frmacc11p0_Save
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         'Add by Amy 2017/12/13 +FormCheck
         If Frmacc11n0.FormCheck = False Then
            strControlButton = MsgText(602)
            Exit Sub
         End If
         'end 2017/12/13
         Frmacc11n0.Frmacc11n0_Save
      '2012/8/29 End
      'Add by Amy 2013/12/02
      Case "Frmacc11o0"
         With Frmacc11o0
            If .FormSave = True Then
                strSaveConfirm = MsgText(601)
                strControlButton = MsgText(601)
                .FormEnabled
            Else
                strControlButton = MsgText(602)
            End If
         End With
      'end 2013/1129
      Case "Frmacc1170"
         With Frmacc1170
            'Add by Amy 2014/01/15 原程式改寫至function 搬回form
            If .FormCheck = False Then
                strControlButton = MsgText(602)
                Exit Sub
            End If
            'end 2013/01/15
            'Modify by Amy 2014/10/27 為資料一致更新acc1p0
            'If strSaveConfirm = MsgText(4) Then
               .Frmacc1170_Save 'Modify by Amy 2013/12/26
            'End If
            If strControlButton <> MsgText(602) Then
                .UpdateAcc1p0
            End If
            'end 2014/10/27
         End With
         If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
            Frmacc1170.FormDisabled
            Frmacc1170.SetData ("F9") 'Add by Amy 2014/10/27
            Frmacc1170.Text1.SetFocus
         End If
      Case "Frmacc1180"
         With Frmacc1180
            strControlButton = MsgText(601)
            'Modify by Amy 2014/01/28 原程式搬回form-FormCheck
            .FormCheck
            If strControlButton = MsgText(602) Then
                Exit Sub
            End If
            'end 2014/01/28
            'Modify by Amy 2014/10/27 為資料一致更新acc1p0
            'If strSaveConfirm = MsgText(4) Then
               .Frmacc1180_Save 'Modify by Amy 2014/01/17
            'End If
            If strControlButton <> MsgText(602) Then
                .UpdateAcc1p0
            End If
            'end 2014/10/27
         End With
         If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
            Frmacc1180.SetData ("F9") 'Add by Amy 2014/10/27
            Frmacc1180.FormDisabled
         End If
      Case "Frmacc1190"
         With Frmacc1190
            '2005/6/15 ADD BY SONIA
            If Val(.Text14) + Val(.Text15) + Val(.Text16) + Val(.Text21) = 0 Then
               MsgBox "請輸入銷退金額", , MsgText(5)
               Exit Sub
            End If
            '2005/6/15 END
            'Added by Morgan 2013/8/7
            '若為部份銷帳且收據未列印時提醒
            If .Option1.Value = True And .Text3 = "1" And Val(.Text14) < Val(.Text13) Then
               'Modified by Lydia 2023/11/13 開立INVOICE，不列印收據=>and NVL(a0k32,'Y') <> 'Z'
               strExc(0) = "select a0k01 from acc0k0 where a0k01='" & .Text1 & "' and nvl(a0k19,0)=0 and NVL(a0k32,'Y') <> 'Z' "
               intI = 1
               Set RsTemp = ClsLawReadRstMsg(intI, strExc(0))
               If intI = 1 Then
                  MsgBox "此為未列印收據, 若確定為不送件案件, 請記得列印收據 ！！", vbExclamation
               End If
            End If
            'end 2013/8/7
            
            .m_AssignNo = .Text1 'Add by Morgan 2011/5/30
            
            .Frmacc1190_Save 'Modified by Morgan 2014/1/10 移至表單內
            
            'Add by Morgan 2011/10/17
            If strControlButton = MsgText(602) Then
               strControlButton = MsgText(601)
               Exit Sub
            Else
               Frmacc1190.cboCaseNo.Enabled = False
            End If
            
            If Frmacc1190.Enabled Then .CheckAssign 'Add by Morgan 2011/5/30
         End With
         
      Case "Frmacc11a0"
         With Frmacc11a0
            If .Text10 <> .Text16 Then
               MsgBox MsgText(11), , MsgText(5)
               Exit Sub
            End If
            'Add by Morgan 2007/4/12 修改時提醒
            If strSaveConfirm = MsgText(4) Then
               If MsgBox("請確認傳票各項次內容是否有逐筆更新資料？", vbYesNo + vbDefaultButton2 + vbExclamation, "提示訊息") = vbNo Then
                  Exit Sub
               End If
            End If
            'end 2007/4/12
            'If strSaveConfirm = MsgText(4) Then
               Frmacc11a0.Frmacc11a0_Save
               'Add by Amy 2014/10/28
               If strControlButton <> MsgText(602) Then
                    Frmacc11a0.UpdateAcc1p0
               End If
               'end 2014/10/28
            'End If
            '.ObjectEnabled_2 '搬至coommit後因11a0_Save時有一些檢查不過時欄位會被鎖住
         End With
         If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
            'Add by Amy 2014/10/28
            Frmacc11a0.ObjectEnabled_2
            Frmacc11a0.SetData ("F9")
         End If
      Case "Frmacc11d0"
         'add by sonia 2019/5/9
         With Frmacc11d0
            If Val(.Text5) + Val(.Text7) <= 0 Then
               MsgBox "服務費及規費合計不可小於或等於０！", , MsgText(5)
               Exit Sub
            End If
         End With
         'end 2019/5/9
         Frmacc11d0.Frmacc11d0_Save
      Case "Frmacc11f0"
         Frmacc11f0_Save
      'Add by Morgan 2005/6/8
      Case "Frmacc11h0"
         If Frmacc11h0.FormSave = True Then
            Frmacc11h0.FormLocked False
            Frmacc11h0.txtA0w02.SetFocus
         Else
            Exit Sub
         End If
         
      'Add by Morgan 2007/4/16
      Case "Frmacc11i0"
         If Frmacc11i0.FormSave = False Then
            Exit Sub
         End If
         
      'Add by Morgan 2007/5/16
      Case "Frmacc11j0"
         If Frmacc11j0.FormSave = False Then
            Exit Sub
         End If
         
      'Add by Morgan 2007/10/5
      Case "Frmacc11k0"
         If Frmacc11k0.FormSave = False Then
            Exit Sub
         End If
         
      'Add by Morgan 2011/4/8
      Case "Frmacc11l0"
         If Frmacc11l0.FormSave = False Then
            Exit Sub
         ElseIf Frmacc11l0.m_sCallType <> "" Then
            Unload Frmacc11l0
         End If
         
      Case "Frmacc2110"
         With Frmacc2110
             Frmacc2110_Save
             If strControlButton <> MsgText(602) Then 'Added by Morgan 2025/10/23
               .FormDisabled
               .Command2.SetFocus
               'Add by Morgan 2006/6/20 修改存檔時若匯率有變則控制一定要進收款資料畫面
                If strSaveConfirm = MsgText(4) And .Text3.Text <> .Text3.Tag Then
                  .bolForm2 = True
                  .Text2.Locked = True
                Else
                  .bolForm2 = False
                End If
             'If strControlButton <> MsgText(602) Then 'Removed by Morgan 2025/10/23 移到上面
               adoTaie.CommitTrans
             End If
         End With
         
      Case "Frmacc2120"
         Frmacc2120.Frmacc2120_Save 'Modify by Amy 2014/11/03 搬回form
         If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
            Frmacc2120.FormEnable 'Add by Morgan 2011/3/10
         End If
         
      Case "Frmacc2130"
         Frmacc2130_Save
      Case "Frmacc2140"
         Frmacc2140_Save
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
            If Val(.Text8) <> Val(.Text7) Then
               MsgBox MsgText(59), , MsgText(5)
               Exit Sub
            End If
            If strSaveConfirm = MsgText(4) Then
               Frmacc2160_Save
            '2012/8/9 add by sonia 新增存檔要再檢查
            Else
               If .Text1 = MsgText(601) Then
                  MsgBox MsgText(149), , MsgText(5)
                  strControlButton = MsgText(602)
                  .Text1.SetFocus
                  Exit Sub
               End If
               If .Combo1 = MsgText(601) Then
                  MsgBox "無抵帳幣別, 請檢核...", , MsgText(5)
                  strControlButton = MsgText(602)
                  .Combo1.SetFocus
                  Exit Sub
               End If
               If Val(.Text7) = 0 Then
                  MsgBox MsgText(58), , MsgText(5)
                  strControlButton = MsgText(602)
                  .Text7.SetFocus
                  Exit Sub
               End If
               .adoacc160.Fields("a1603").Value = .Text1
               .adoacc160.Fields("a1605").Value = .Combo1
               .adoacc160.Fields("a1606").Value = Val(.Text7)
               .adoacc160.UpdateBatch
            '2012/8/9 end
            End If
            If strControlButton <> MsgText(602) Then
               .FormDisabled
            End If
         End With
         If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
         End If
         
      'Added by Morgan 2023/3/30
      Case "Frmacc2172"
         If Frmacc2172.FormSave = False Then
            Exit Sub
         End If
      'end 2023/3/30
      
      Case "Frmacc21d0"
         With Frmacc21d0
            'Add by Morgan 2010/3/24 借貸方檢核
            If .CreDebCheck <> MsgText(602) Then
               MsgBox MsgText(11), , MsgText(5)
               strControlButton = MsgText(602)
               Exit Sub
            End If
             .Frmacc21d0_Save 'Modify by Amy 2014/11/03 搬回from
'            .FormDisabled
'            .Text1.SetFocus
         End With
         If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
            Frmacc21d0.FormDisabled
            Frmacc21d0.Text1.SetFocus
         End If
      Case "Frmacc21e0"
         With Frmacc21e0
            If .CreDebCheck <> MsgText(602) Then
               MsgBox MsgText(11), , MsgText(5)
               strControlButton = MsgText(602)
               Exit Sub
            End If
            If strSaveConfirm = MsgText(4) Then
               .Frmacc21e0_Save 'Modify by Amy 2014/11/05 搬回form
            End If
            '.FormDisabled '2014/11/06 往下搬至SetData中(因frmacc21e0_Save檢查欄位跳離仍跑FormDisabled會鎖住要key 的欄位)
         End With
         If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
            Frmacc21e0.SetData ("F9") 'Modify by Amy 2014/11/06
         End If
      Case "Frmacc21f0"
         With Frmacc21f0
            If strSaveConfirm = MsgText(4) Then
                .Frmacc21f0_Save 'Modify by Amy 2014/11/05 搬回form
            End If
'            .FormDisabled
         End With
         If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
            Frmacc21f0.FormDisabled
         End If
      Case "Frmacc21f1"
         With Frmacc21f1
            If .CreDebCheck <> MsgText(602) Then
               MsgBox MsgText(11), , MsgText(5)
               strControlButton = MsgText(602)
               Exit Sub
            End If
            If strSaveConfirm = MsgText(4) Then
               .Frmacc21f1_Save 'Modify by Amy 2014/11/05 搬回from
            End If
'            .FormDisabled
         End With
         If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
            Frmacc21f1.FormDisabled
         End If
      Case "Frmacc21f2"
         With Frmacc21f2
            If .CreDebCheck <> MsgText(602) Then
               MsgBox MsgText(11), , MsgText(5)
               strControlButton = MsgText(602)
               Exit Sub
            End If
            If strSaveConfirm = MsgText(4) Then
               .Frmacc21f2_Save 'Modify by Amy 2014/11/05 搬回from
            End If
'            .FormDisabled
         End With
         If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
            Frmacc21f2.FormDisabled
         End If

'Remove by Morgan 2005/1/14 財務不需要
'            Case "Frmacc21g0"
'               Frmacc21g0_Save
      Case "Frmacc21h0"
         Frmacc21h0_Save
         If strControlButton <> MsgText(602) Then
            With Frmacc21h0
               .Command1.Enabled = False
               .Command3.Enabled = False
               .Command5.Enabled = True
               If Len(.Text5) = 10 Then
                  .Command2.Enabled = False
               Else
                  .Command2.Enabled = True
               End If
            End With
         End If
         adoTaie.CommitTrans
      Case "Frmacc21i0"
         Frmacc21i0_Save
      Case "Frmacc21j0"
         Frmacc21j0_Save
      Case "Frmacc21k0"
         Frmacc21k0_Save
'Remove by Morgan 2005/1/14 財務不用
'Modified by Morgan 2019/10/5 改回由財務維護
      Case "Frmacc21m0"
         Frmacc21m0_Save Frmacc21m0
         
      Case "Frmacc21n0"
         Frmacc21n0_Save
'Remove by Morgan 2005/8/3 財務不用
'Modified by Morgan 2019/7/10 又改要用
      Case "Frmacc21o0"
         Frmacc21o0_Save Frmacc21o0
         
      'Add By Chenhg 2003/07/22
      Case "Frmacc21q0"
         'Added by Lydia 2015/03/30 改function判斷是否寫入
         'Frmacc21q0_Save
         If Frmacc21q0_Save = False Then
            Exit Sub
         End If
      'Add by Morgan 2006/12/18
      Case "Frmacc21r0"
         With Frmacc21r0
         If .FormSave = True Then
            strControlButton = MsgText(601)
            .Command1.Enabled = True
            .txtKey.Locked = False
            .txtBox(1).Locked = False
            .txtBox(5).Locked = False 'Added by Lydia 2018/07/20 財務信箱(CF)
            'Add by Morgan 2007/3/3
            .txtInform(0).Locked = True
            .txtInform(1).Locked = True
            'end 2007/3/3
            'Modify by Amy 2014/04/03 原程式改至form
            .SetCheck1 (False)
         Else
            strControlButton = MsgText(602)
         End If
         End With
         
      'Added by Morgan 2019/10/5
      Case "Frmacc21s0"
         Frmacc21s0_Save Frmacc21s0
         
      'Added by Lydia 2016/11/07
      Case "Frmacc21w0"
         With Frmacc21w0
            If .FormSave = True Then
               .Command1.Enabled = True
               .txtKey.Locked = False
            Else
               strControlButton = MsgText(602)
            End If
         End With
      Case "Frmacc3110"
         With Frmacc3110
'                  If .Text21 = MsgText(601) And .Text22 = MsgText(601) Then
'                     Frmacc3110_Save
'                  Else
               .Frmacc3110_Save
'                  End If
            'Modify by Morgan 2007/1/5
            '.FormDisabled
            If strControlButton <> MsgText(602) Then
               .FormDisabled
               .SetData ("F9") 'Add by Amy 2014/11/12
               .Text11.SetFocus
            End If
            'end 2007/1/5
         End With
      Case "Frmacc3120"
         With Frmacc3120
'                  If .Text21 = MsgText(601) And .Text22 = MsgText(601) Then
'                     Frmacc3120_Save
'                  Else

            strControlButton = MsgText(601)
               .Frmacc3120_Save 'Modify by Amy 2014/11/05 搬回form
'                  End If
            If strControlButton <> MsgText(602) Then
               .FormDisabled
               .Text12.SetFocus
            End If
         End With
      Case "Frmacc3130"
         Frmacc3130_Save
      Case "Frmacc3140"
         Frmacc3140.Frmacc3140_Save 'Modify by Amy 2020/07/16
      Case "Frmacc3150"
         Frmacc3150.Frmacc3150_Save 'Modify by Amy 2020/07/17
      Case "Frmacc3160"
         Frmacc3160.Frmacc3160_Save 'Modify by Amy 2020/07/17
      Case "Frmacc3170"
         With Frmacc3170
            If strSaveConfirm = MsgText(4) Then
               Frmacc3170.Frmacc3170_Save 'Modify by Amy 2020/07/17
            End If
'            .FormDisabled
         End With
         adoTaie.CommitTrans
         Frmacc3170.FormDisabled
      Case "Frmacc3180"
         Frmacc3180.Frmacc3180_Save 'Add by Amy 2021/10/19 搬回from
         
      Case "Frmacc3190"
         Frmacc3190.Frmacc3190_Save 'Add by Amy 2021/10/19 搬回form
      Case "Frmacc31a0"
         Frmacc31a0.Frmacc31a0_Save 'Modify by Amy 2014/11/05 搬回form
      Case "Frmacc31c0"
         Frmacc31c0.Frmacc31c0_Save 'Modify by Amy 2014/11/05 搬回form
         'Add by Amy 2014/11/14 +if
         If strControlButton <> MsgText(602) Then
             Frmacc31c0.SetToolBar
         End If
      Case "Frmacc4110"
         Frmacc4110.Frmacc4110_Save 'Modify by Amy 2015/06/11搬回form
      Case "Frmacc4120"
         'Modify by Amy 2024/07/31 原:FormF2Check 整合檢查程式,避免有未改到的
         'Frmacc4120.FormCheck 'Modify by Amy 2014/01/06搬回form 並改寫
         If Frmacc4120.ChkForm("F9") = False Then
            strControlButton = MsgText(602)
            Exit Sub
         End If
         If strControlButton <> MsgText(602) Then
            Frmacc4120.SaveData ("Save")
         'end 2024/07/31
            adoTaie.CommitTrans
         End If
      Case "Frmacc4130"
         Frmacc4130.Frmacc4130_Save
      Case "Frmacc4140"
         Frmacc4140_Save
      Case "Frmacc4150"
'               Frmacc4150_Save
      Case "Frmacc4160"
         'Modify by Amy 2024/08/23 +存檔前檢查
         If Frmacc4160.Acc040Check("F9") = False Then
            strControlButton = MsgText(602)
            Exit Sub
         End If
         If strControlButton <> MsgText(602) Then
            'Add by Amy 2024/08/23
            If strSaveConfirm = MsgText(3) Then  '新增
               Frmacc4160.Acc040Save
            End If
            'end 2024/08/19
            adoTaie.CommitTrans
            Frmacc4160.adoacc040T.Requery
            Frmacc4160.FormDisabled
         End If
         'end 2024/08/19
      Case "Frmacc4170"
         With Frmacc4170
            'Added by Lydia 2021/12/22 跳過一次KeyF9檢查; 因為從frmacc41i0新增時自動呼叫frmacc4170會對frmacc4170再執行一次F9
            If .bolA4170Jump = True Then
               .bolA4170Jump = False
               Exit Sub
            End If
            'end 2021/12/22
            
            strControlButton = MsgText(601) 'Added by Lydia 2022/02/07 重置判斷條件
            If .CreDebCheck <> MsgText(602) Or Val(.Text11) = 0 Or Val(.Text12) = 0 Then
               MsgBox MsgText(11), , MsgText(5)
               strControlButton = MsgText(602)
               Exit Sub
            End If
            
            'Mark by Lydia 2022/02/07 改到下方
            'If strSaveConfirm = MsgText(4) Then
            '   .Frmacc4170_Save
            'End If
            'end 2022/02/07
            'Add By Sindy 2012/2/2 從aacc_sav搬過來
            Dim star_month As String, end_month As String
            Dim intMonth As Integer
            star_month = CStr(Val(Left(.MaskEdBox3, 3)) + 1911) & Mid(.MaskEdBox3, 4, Len(.MaskEdBox3))
            end_month = CStr(Val(Left(.MaskEdBox4, 3)) + 1911) & Mid(.MaskEdBox4, 4, Len(.MaskEdBox4))
            intMonth = DateDiff("m", star_month, end_month) + 1
            '
            If Val(Mid(.MaskEdBox4.Text, 1, 3) & Mid(.MaskEdBox4.Text, 5, 2) & String(2 - Len(.MaskEdBox1), "0") & .MaskEdBox1) >= strSrvDate(2) Then
               If Val(Format(.Text7, strPercent)) / Val(Format(.Text11, strPercent)) <> intMonth Then
                  MsgBox "總額除以每月合計不等於有效月數！", , MsgText(5)
                  strControlButton = MsgText(602)
                  .Text7.SetFocus
                  Exit Sub
               End If
            End If
            '2012/2/2 End
'            .FormDisabled
'            .Text1.SetFocus
         End With
         
         'Added by Lydia 2022/02/07 不論新增或修改,再次檢查和更新主檔; ex.推測451輸成*12個月的金額,452是沒有輸入金額
         If strControlButton <> MsgText(602) Then
             Call Frmacc4170.Frmacc4170_Save
         End If
         'end 2022/02/07
         
         If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
            Frmacc4170.FormDisabled
            Frmacc4170.Text1.SetFocus
         End If
      Case "Frmacc4180"
         Frmacc4180_Save
      Case "Frmacc4190"
         With Frmacc4190
            'Modify by Amy 2020/11/17 搬至表單中
            'Modidfy by Amy 2024/08/15 +F9
            If .FormCheck("F9") = False Then
               strControlButton = MsgText(602)
               Exit Sub
            End If
            If strControlButton <> MsgText(602) Then
               .SetData ("F9") 'Modify by Amy 2024/08/12 原程式搬回表單中
               adoTaie.CommitTrans
            End If
         End With
      Case "Frmacc41a0"
        'edit by nickc 2007/08/24 將檢查的動作，從 acc_sav 移出來，才不會檢查錯誤又存檔
        With Frmacc41a0
            strControlButton = ""  '2012/7/9 add by sonia 否則有錯誤後就無法再操作
            If .Text10 = MsgText(601) Then
               MsgBox MsgText(10), , MsgText(5)
               .Text10.SetFocus
               Exit Sub
            Else
               'Add by Morgan 2011/6/23
               If strSaveConfirm = MsgText(3) Then
                  If .AddCheck = False Then
                     .Text10.SetFocus
                     Exit Sub
                  End If
               End If
               'end 2011/6/23
         
               'end 2011/6/23
               'edit by nickc 2008/02/29 婧瑄改若結轉收入為0就不用管智權人員
               'If .Text11 = MsgText(601) Then
               If .Text11 = MsgText(601) And Val(.Text9) <> 0 Then
                  MsgBox MsgText(10), , MsgText(5)
                  strControlButton = MsgText(602)
                  .Text11.SetFocus
                  Exit Sub
               Else
'                  'Add By Sindy 2014/2/21
'                  Cancel = False
'                  .Text11_Validate Cancel
'                  If Cancel = True Then
'                     strControlButton = MsgText(602)
'                     .Text11.SetFocus
'                     Exit Sub
'                  End If
'                  '2014/2/21 END
                  If ExistCheck("staff", "st01", .Text11, .Label9) = False Then
                     MsgBox MsgText(45) & .Label9, , MsgText(5)
                     strControlButton = MsgText(602)
                     .Text11.SetFocus
                     Exit Sub
                  End If
                  'Add By Sindy 2014/2/24
                  StrSQLa = "select st02, st03,st04 from staff where st01 = '" & .Text11 & "'"
                  rsA.CursorLocation = adUseClient
                  rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
                  If rsA.RecordCount > 0 Then
                     If "" & rsA.Fields("st04").Value = "2" Then
                        If MsgBox("此智權人員已離職，確認是否繼續？", vbYesNo + vbDefaultButton2 + vbExclamation, "提示訊息") = vbNo Then
                           strControlButton = MsgText(602)
                           .Text11.SetFocus
                           If rsA.State <> adStateClosed Then rsA.Close
                           Set rsA = Nothing
                           Exit Sub
                        End If
                     End If
                  End If
                  If rsA.State <> adStateClosed Then rsA.Close
                  Set rsA = Nothing
                  '2014/2/24 END
               End If
               If .MaskEdBox2.Text = MsgText(601) Or .MaskEdBox2.Text = MsgText(29) Then
                  MsgBox .Label10 & MsgText(52), , MsgText(5)
                  strControlButton = MsgText(602)
                  .MaskEdBox2.SetFocus
                  Exit Sub
               End If
               If DateCheck(.MaskEdBox2.Text) = MsgText(603) Then
                  MsgBox .Label10 & MsgText(63), , MsgText(5)
                  strControlButton = MsgText(602)
                  .MaskEdBox2.SetFocus
                  Exit Sub
               End If
               'add by sonia 2025/5/23
               StrSQLa = "select distinct a1p18 from acc1p0 where a1p04 = '" & .Text10 & "' and a1p18<>" & Val(FCDate(.MaskEdBox2.Text))
               rsA.CursorLocation = adUseClient
               rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
               If rsA.RecordCount > 0 Then
                  If Val(FCDate(.MaskEdBox2.Text)) <> rsA.Fields("a1p18").Value Then
                     MsgBox "結算日期與分錄日期不同，請進入傳票資料畫面逐筆更新分錄日期 !", , MsgText(5)
                     strControlButton = MsgText(602)
                     .MaskEdBox2.SetFocus
                     If rsA.State <> adStateClosed Then rsA.Close
                     Set rsA = Nothing
                     Exit Sub
                  End If
               End If
               If rsA.State <> adStateClosed Then rsA.Close
               Set rsA = Nothing
               'end 2025/5/23
            End If
        End With
         'Frmacc41a0_Save 'Modify by Amy 2013/12/17
         Frmacc41a0.FormSave 'Add 2013/12/17
         adoTaie.CommitTrans
         strSaveConfirm = MsgText(601)
         With Frmacc41a0
            .CheckBalance 'Add by Morgan 2011/3/23
            .FormDisabled
         End With
      Case "Frmacc41b0"
         'Frmacc41b0_Save 'Modify by Amy 2013/12/17
         Frmacc41b0.FormSave 'Add 2013/12/17
      Case "Frmacc41d0"
         'Modify by Amy 2020/04/14 原程式搬回表單中
         If Frmacc41d0.FormCheck = True Then
            adoTaie.CommitTrans
         End If
      'Add by Morgan 2005/4/6
      Case "Frmacc41e0"
         With Frmacc41e0
            If .SaveData = True Then
               'adoTaie.CommitTrans 'Removed by Morgan 2018/2/9 改存檔控制就好
               .MailCheck
               .FormEnable
               .cmdFind.Value = True 'Added by Morgan 2015/6/17
            Else
               Exit Sub
            End If
         End With
      'Added by Lydia 2017/03/03
      Case "Frmacc41i0"    '財產目錄
         With Frmacc41i0
            If .FormCheck = False Then
                strControlButton = MsgText(602)
                Exit Sub
            End If
            .Frmacc41i0_Save
            If strControlButton <> MsgText(602) Then
                .UpdateAcc1p0
            End If
         End With
         If strControlButton <> MsgText(602) Then
            'Added by Lydia 2017/05/11 自動新增固定傳票
            If strSaveConfirm = MsgText(3) And Val(FCDate(Frmacc41i0.MaskEdBox1.Text)) >= Pub_A2b05Begin Then
               Frmacc41i0.Acc0dxSave
            End If
            'end 2017/05/11
            adoTaie.CommitTrans
            'Added by Lydia 2017/05/11 重整資料集
            Call Frmacc41i0.Acc2b0Refresh(Frmacc41i0.txtA2B01.Text)
            Frmacc41i0.FormDisabled
            'Modified by Lydia 2017/05/11 自動新增固定傳票
             'If strSaveConfirm = MsgText(3) Then
             If strSaveConfirm = MsgText(3) And Val(FCDate(Frmacc41i0.MaskEdBox1.Text)) >= Pub_A2b05Begin Then
                Frmacc41i0.m_Auto = True
             Else
                Frmacc41i0.txtA2B01.SetFocus
             End If
         End If
      'end 2017/03/03
      'Added by Lydia 2017/03/08
      Case "Frmacc41i0_1"    '財產報廢作業
         With Frmacc41i0_1
            If .FormCheck = False Then
                strControlButton = MsgText(602)
                Exit Sub
            End If
            .Frmacc41i0_1_Save
            If strControlButton <> MsgText(602) Then
                .UpdateAcc1p0
            End If
         End With
         If strControlButton <> MsgText(602) Then
            adoTaie.CommitTrans
            Frmacc41i0_1.FormDisabled
            'Frmacc41i0_1.txtA2B01.SetFocus 'Remove by Lydia 2017/05/15
         End If
      'end 2017/03/08
      Case "Frmacc5200"
         With Frmacc5200
            If .Adodc1.Recordset.RecordCount = 0 Then
               MsgBox MsgText(50), , MsgText(5)
               strControlButton = MsgText(602)
            End If
            If strControlButton <> MsgText(602) Then
               .Command1.Enabled = False
               .DataGrid1.AllowUpdate = False
               adoTaie.CommitTrans
            End If
         End With
   End Select
   If strControlButton <> MsgText(602) Then
      strSaveConfirm = MsgText(601)
      Select Case strFormName
         Case "Frmacc1110"
            tool9_enabled
         'Modify by Morgan 2006/3/20
         'Case "Frmacc1140", "Frmacc11d0"
         'Modify by Sindy 2014/1/9 +Frmacc11q0
         'Modify by Amy 2017/09/04 將frmacc114拆出(不需新增鈕)
         Case "Frmacc1140"
            Call Frmacc1140.ReadData    'Add by Amy 2025/07/16
            tool8_enabled
         Case "Frmacc1130", "Frmacc11d0", "Frmacc11q0"
         'end 2017/09/04
            tool14_enabled
            If strFormName = "Frmacc11q0" Then
               Forms(0).Toolbar1.Buttons.Item(5).Enabled = False
            End If
         Case "Frmacc2111", "Frmacc21h1", "Frmacc21f1", "Frmacc21f2"
            tool7_enabled
         'Add by Amy 2017/08/18
         Case "Frmacc4120"
            Call Frmacc4120.SetData("BtExit")
         'Add by Morgan 2005/6/8
         Case "Frmacc11h0"
            tool7_enabled
         Case "Frmacc21i0"
            tool8_enabled
         'Modify end--------------------
         'add by nickc 2005/08/03
         Case "Frmacc41b0"
            tool1_enabled
            Frmacc0000.Toolbar1.Buttons.Item(8).Enabled = False
         
         Case "Frmacc41a0"
            tool1_enabled
            Frmacc41a0.DisabledMoveRecord 'Add by Morgan 2011/6/23
            Frmacc0000.Toolbar1.Buttons.Item(8).Enabled = False
        'Add by Amy 2014/02/11
        Case "Frmacc41d0"
           tool8_enabled
           Frmacc41d0.FormDisabled
'Remove by Morgan 2010/12/7 --找不到加此控制的原因,先還原
'         'Add by Morgan 2010/11/22
'         Case "Frmacc21q0"
'            tool6_enabled
'            Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True

         'Add by Morgan 2006/12/18
         Case "Frmacc21r0"
            tool6_enabled
            Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
         'Add by Amy 2014/11/14
         Case "Frmacc31c0"
            tool1_enabled
            Frmacc31c0.SetToolBar
         'Add by Morgan 2006/12/27
         Case "Frmacc1194"
            tool3_enabled
         'Add by Morgan 2007/4/18
         Case "Frmacc11i0"
            Frmacc11i0.FormEnable
            tool1_enabled
         'Add by Morgan 2007/5/16
         Case "Frmacc11j0"
            Frmacc11j0.FormEnable
            tool1_enabled
         'Add by Morgan 2007/10/5
         Case "Frmacc11k0"
            Frmacc11k0.FormEnable
            tool1_enabled
         'Add by Sindy 2013/12/27
         Case "Frmacc1127"
            tool3_enabled
        'Add by Amy 2022/08/24
        Case "Frmacc1127"
            tool3_enabled
        'Add by Amy 2017/04/21
        Case "Frmacc41j0"
            With Frmacc41j0
                If .FormCheck(0, "F9") = False Then
                    Exit Sub
                End If
                If strControlButton <> MsgText(602) Then
                    .SetData ("F9")
                End If
            End With
            tool1_enabled
            Frmacc0000.Toolbar1.Buttons.Item(9).Enabled = False
        'Add by Amy 2014/02/14
         Case "Frmacc5100"
            tool8_enabled
            Frmacc5100.FormDisabled
         'Add by Sindy 2015/7/9
         Case "Frmacc11p0"
            If UCase(strUserLevel) = UCase("Frmacc44t0") Or UCase(strUserLevel) = UCase("Frmacc11b0") Then
               tool6_enabled
               Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
            Else
               tool1_enabled
            End If
         '2015/7/9 END
         
         'Added by Morgan 2023/3/30
         Case "Frmacc2172"
            tool10_enabled
            Frmacc2172.FormEnabled
         'end 2023/3/30
         
         Case Else
            tool1_enabled
      End Select
      Frmacc0000.StatusBar1.Panels(1).Text = MsgText(17)
   End If
   strControlButton = MsgText(601)
   
   'Added by Morgan 2013/12/26
   If strFormName = "Frmacc1150" Then
      If Frmacc1150.m_AutoRun = True Then
         Frmacc1150.Command1.Value = True
      End If
   End If
   'end 2013/12/26
   
   'Added by Lydia 2017/03/03
   If strFormName = "Frmacc41i0" Then
      If Frmacc41i0.m_Auto = True Then
         Call Frmacc41i0.cmdCall_Click
         Frmacc41i0.m_Auto = False
      End If
   End If
   'end 2017/03/03
   
Checking:
   If Err.Number = 0 Or Err.Number = -2147168242 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
   
End Sub

'*************************************************
'  取消

'*************************************************
Private Sub KeyEnterF10()
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
On Error GoTo Checking
   
   If strSaveConfirm = MsgText(601) Then
      Exit Sub
   End If
   Select Case strFormName
      Case "Frmacc1150"
         'Modified by Morgan 2014/1/14 有做rollback 所以Mark可不做的程式
         adoTaie.RollbackTrans
         With Frmacc1150
            'If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
            '   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'A' and a1p03 = '" & .Text2 & "' and a1p04 = 'A'", intI
            '   adoTaie.Execute "delete from acc0l0 where a0l01 = '" & .Text2 & "'"
            '   Frmacc1150_Clear
               'Modified by Morgan 2022/6/30
               '.adoacc0l0.Requery
               '.AdodcRefresh
               '.SumShow
               'If strSaveConfirm = "A" Then
               If strSaveConfirm = MsgText(3) Then
                  Frmacc1150_Clear
               Else
                  .RefreshData
               End If
               'end 2022/6/30
               '.AdodcClear
               If .adoacc0l0.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            'End If
            .FormDisabled
         End With
         'adoTaie.RollbackTrans
         'end 2014/1/14
         
      Case "Frmacc1170"
         'Modify by Amy 2014/01/06 有做rollback 所以Mark可不做的程式
         adoTaie.RollbackTrans
         With Frmacc1170
            'If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
            '   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'B' and a1p03 = '" & .Text1 & "' and a1p04 = 'B'"
            '   adoTaie.Execute "delete from acc0o0 where a0o01 = '" & .Text1 & "'"
            '   .Frmacc1170_Clear 'Modify by Amy 2013/12/26
               .adoacc0o0.Requery
               .AdodcRefresh
               .SumShow
               .AdodcClear
               If .adoacc0o0.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            'End If
            .FormDisabled
            .SetData ("F10") 'Add by Amy 2016/07/21
         End With
          'end2014/01/06
         'adoTaie.RollbackTrans 'Modify by Amy 2014/01/06  因mark上方部分程式,所以往上搬
      Case "Frmacc1180"
         With Frmacc1180
            adoTaie.RollbackTrans 'Modify by Amy 2014/01/28 有做rollback 所以Mark可不做的程式
            'If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
            '   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'C' and a1p03 = '" & .Text2 & "' and a1p04 = '" & FCDate(.MaskEdBox7.Text) & "'"
            '   adoTaie.Execute "delete from acc0q0 where a0q01 = " & Val(FCDate(.MaskEdBox7.Text)) & " and a0q03 = '" & .Text1 & "'"
               .Frmacc1180_Clear (True) 'Modify by Amy 2014/01/28 +參數
               .adoacc0q0.Requery
               .AdodcRefresh
               .SumShow
               .AdodcClear
               If .adoacc0q0.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            'End If
            .FormDisabled
         End With
         'adoTaie.RollbackTrans '因mark上方部分程式,所以往上搬
         'end 2014/01/28
         
      Case "Frmacc11a0"
         With Frmacc11a0
            adoTaie.RollbackTrans 'Modify by Amy 2014/10/28 有做rollback 所以Mark可不做的程式
            'If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
            '   adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'D' and a1p03 = '" & .Text1 & "' and a1p04 = 'D'"
            '   adoTaie.Execute "delete from acc0t0 where a0t01 = '" & .Text1 & "'"
               .Frmacc11a0_Clear 'Modify by Amy 2014/10/29 搬回Form
               .adoacc0t0.Requery
               .AdodcRefresh
               .SumShow
               .AdodcClear
               If .adoacc0t0.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            'End If
            .ObjectEnabled_2
         End With
         'adoTaie.RollbackTrans '因mark上方部分程式,所以往上搬
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         With Frmacc11p0
            .AdodcRefresh
            .FormDisabled
         End With
         'Add by Sindy 2015/7/9
         'Modify by Sindy 2016/11/29
         If UCase(strUserLevel) = UCase("Frmacc11b0") Or _
            UCase(strUserLevel) = UCase("Frmacc44w1") Then
            tool6_enabled
            Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
         ElseIf UCase(strUserLevel) <> UCase("Frmacc44t0") Then
         '2015/7/9 END
            Frmacc11p0.Frmacc11p0_Clear
         End If
      '2013/12/19 End
      'Add By Sindy 2012/9/4
      Case "Frmacc11n0"
         With Frmacc11n0
            .FormDisabled
         End With
         Frmacc11n0.Frmacc11n0_Clear
      '2012/9/4 End
      Case "Frmacc2110"
         With Frmacc2110
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'F' and a1p03 = '" & .Text2 & "' and a1p04 = 'F'"
               adoTaie.Execute "delete from acc0y0 where a0y01 = '" & .Text2 & "'"
               Frmacc2110_Clear
               .adoacc0y0.Requery
               .AdodcRefresh
               .SumShow
               .AdodcClear
               If .adoacc0y0.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            End If
            .FormDisabled
         End With
         adoTaie.RollbackTrans
      Case "Frmacc2120"
         adoTaie.RollbackTrans
         Frmacc2120.FormEnable 'Add by Morgan 2011/3/10
      Case "Frmacc2150"
         With Frmacc2150
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc151 where axf01 = '" & .Text2 & "'"
               adoTaie.Execute "delete from acc150 where a1501 = '" & .Text2 & "'"
               Frmacc2150_Clear
               .adoacc150.Requery
               .AdodcRefresh
               .AdodcClear
               If .adoacc150.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            End If
            .FormDisabled
         End With
         adoTaie.RollbackTrans
      Case "Frmacc2160"
         With Frmacc2160
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc161 where axg01 = '" & .Text2 & "'"
               adoTaie.Execute "delete from acc160 where a1601 = '" & .Text2 & "'"
               Frmacc2160_Clear
               .adoacc160.Requery
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
      Case "Frmacc21d0"
         With Frmacc21d0
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'I' and a1p03 = '" & .Text3 & "' and a1p04 = '" & .Text1 & "'"
               adoTaie.Execute "update acc190 set a1908 = '' where a1908 = '" & .Text3 & "'"
               adoTaie.Execute "delete from acc1c0 where a1c01 = '" & .Text3 & "' and a1c02 = '" & .Text1 & "'"
               Frmacc21d0.Frmacc21d0_Clear 'Modify by Amy 2014/11/04搬回form
               .adoacc1b0.Requery
               .AdodcRefresh
               .Adodc3Clear
               If .adoacc1b0.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            End If
            .FormDisabled
            .Text1.SetFocus
         End With
         adoTaie.RollbackTrans
      Case "Frmacc21e0"
         With Frmacc21e0
            'Modify by Amy 2014/11/06
            adoTaie.RollbackTrans '有做rollback 所以Mark可不做的程式
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
'               adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'J' and a1p03 = '" & .Text1 & "' and a1p04 = '" & Val(FCDate(.MaskEdBox1.Text)) & "'"
'               adoTaie.Execute "delete from acc1e0 where a1e01 = '" & .Text1 & "' and a1e02 = " & Val(FCDate(.MaskEdBox1.Text)) & ""
               Frmacc21e0_Clear
               .adoacc1e0.Requery
               .AdodcRefresh
               .AdodcClear
               If .adoacc1e0.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            End If
            '.FormDisabled
            Frmacc21e0.SetData ("F10")
         End With
         'adoTaie.RollbackTrans
         'end 2014/11/06
      Case "Frmacc21f0"
         adoTaie.RollbackTrans
         With Frmacc21f0
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "update acc1k0 set a1k17 = '' where a1k17 = '" & .Text9 & "'"
               adoTaie.Execute "update acc150 set a1512 = '' where a1512 = '" & .Text9 & "'"
               Frmacc21f0_Clear
               .adoacc1g0.Requery
               .AdodcRefresh1
               .AdodcRefresh2
               If .adoacc1g0.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            End If
            .FormDisabled
         End With
      Case "Frmacc21f1"
         With Frmacc21f1
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p03 = '" & strItemNo & "' and a1p04 = 'K'"
               adoTaie.Execute "delete from acc1i0 where a1i01 = '" & strItemNo & "'"
               Frmacc21f1_Clear
               .adoacc1i0.Requery
               .AdodcRefresh
               .SumShow
            End If
            .FormDisabled
         End With
         adoTaie.RollbackTrans
         Frmacc21f1.AdodcRefresh
         Frmacc21f1.SumShow
      Case "Frmacc21f2"
         With Frmacc21f2
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'K' and a1p03 = '" & strItemNo & "' and a1p04 = 'K'"
               adoTaie.Execute "delete from acc1h0 where a1h01 = '" & strItemNo & "'"
               Frmacc21f2_Clear
               .adoacc1h0.Requery
               .AdodcRefresh
               .SumShow
            End If
            .FormDisabled
         End With
         adoTaie.RollbackTrans
         Frmacc21f2.AdodcRefresh
         Frmacc21f2.SumShow
      Case "Frmacc21h0"
         With Frmacc21h0
            .Command1.Enabled = False
            .Command3.Enabled = False
            .Command5.Enabled = True
            If Len(.Text5) = 10 Then
               .Command2.Enabled = False
            Else
               .Command2.Enabled = True
            End If
         End With
         adoTaie.RollbackTrans
      Case "Frmacc21h1"
         With Frmacc21h1
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc1l0 where a1l01 = '" & .Text1 & "'"
               .adoacc1k0.Requery
               .AdodcRefresh
               .SumShow
            End If
            .FormDisabled
         End With
         adoTaie.RollbackTrans
      'Add By Cheng 2003/07/23
      Case "Frmacc21q0"
         Frmacc21q0.Command1.Enabled = True
      'Add by Morgan 2006/12/18
      Case "Frmacc21r0"
         With Frmacc21r0
            .Command1.Enabled = True
            .txtKey.Locked = False
            .txtBox(1).Locked = True
            .txtBox(5).Locked = True 'Added by Lydia 2018/07/20 財務信箱(CF)
            'Add by Morgan 2007/3/3
            .txtInform(0).Locked = True
            .txtInform(1).Locked = True
            'end 2007/3/3
            'Modify by Amy 2014/04/03 原程式改至form
            .SetCheck1 (False)
            .FormShow
         End With
      'Added by Lydia 2016/11/07
      Case "Frmacc21w0"
         With Frmacc21w0
             .Command1.Enabled = True
             .txtKey.Locked = False
         End With
      Case "Frmacc3110"
         adoTaie.RollbackTrans 'Add by Morgan 2007/1/5
         With Frmacc3110
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p03 = '" & .Text5 & "' and a1p04 = '" & .Text11 & "1" & "'"
               adoTaie.Execute "delete from acc0e0 where a0e01 = '" & .Text11 & "' and a0e02 = '" & .Text5 & "'"
               .Frmacc3110_Clear
            End If
            .AdodcRefresh
            .Adodc2Refresh
            .Adodc2Clear
            If .Adodc1.Recordset.RecordCount <> 0 Then
               .RecordShow
            Else
               StatusClear
            End If
            .FormDisabled
         End With
         'adoTaie.RollbackTrans 'Remove by Morgan 2007/1/5 搬到上面
      Case "Frmacc3120"
         With Frmacc3120
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc1p0 where a1p01 = '1' and a1p02 = 'L' and a1p03 = '" & .Text5 & "' and a1p04 = '" & .Text12 & "1" & "'"
               adoTaie.Execute "delete from acc0e0 where a0e01 = '" & .Text12 & "' and a0e02 = '" & .Text5 & "'"
               Frmacc3120.Frmacc3120_Clear 'Modify by Amy 2020/07/14
               .AdodcRefresh
               .Adodc2Refresh
               .Adodc2Clear
               If .Adodc1.Recordset.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            End If
            .FormDisabled
         End With
         adoTaie.RollbackTrans
      Case "Frmacc3170"
         With Frmacc3170
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc0f0 where a0f01 = " & Val(FCDate(.MaskEdBox1.Text)) & " and a0f02 = '" & .Text2 & "'"
               adoTaie.Execute "update acc0e0 set a0e17 = '', a0e18 = '' where a0e17 = " & Val(FCDate(.MaskEdBox1.Text)) & " and a0e18 = '" & .Text2 & "'"
               Frmacc3170.Frmacc3170_Clear 'Modify by Amy 2020/07/17
               .Acc0f0Refresh
               .AdodcRefresh
               .AdodcClear
               If .adoacc0f0.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            End If
            .FormDisabled
         End With
         adoTaie.RollbackTrans
      Case "Frmacc4120"
         adoTaie.RollbackTrans 'Modify by Amy 2022/05/14 從下面搬上來
         '按修改->Insert->取消->修改->Insert->存檔會出現「找不到要更新的資料列。最後取的值已被變更」,無法修改
         With Frmacc4120
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc021 where ax201 = '" & .Text1 & "' and ax202 = '" & .Text2 & "'"
               adoTaie.Execute "delete from acc020 where a0201 = '" & .Text1 & "' and a0202 = '" & .Text2 & "'"
               .Frmacc4120_Clear 'Modify by Amy 2014/01/14 搬回form
               .adoacc020.Requery
               .AdodcRefresh
               .SumShow
               .AdodcClear
               If .adoacc020.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            End If
           .SetData ("F10") 'Modify by Amy 2014/01/14 搬回form
         End With
'               If strControlButton = MsgText(602) Then
            'adoTaie.RollbackTrans 'Mark by Amy 2022/05/14 往上搬
'               End If
         With Frmacc4120
            If strSaveConfirm = MsgText(4) Then
               .AdodcRefresh
               .SumShow
            End If
         End With
      Case "Frmacc4160"
         If strSaveConfirm = MsgText(3) Then
            Frmacc4160.Frmacc4160_Delete 'Modify by Amy 2024/08/23 程式搬回表單中
            'Modify by Amy 2013/12/24
            Frmacc4160.Frmacc4160_Clear
         End If
         adoTaie.RollbackTrans
         Frmacc4160.FormDisabled
         'Add by Amy 2024/08/23
         If strSaveConfirm = MsgText(4) Then
            Frmacc4160.SetData ("F10")
         End If
      Case "Frmacc4170"
         With Frmacc4170
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc0d0 where a0d01 = '" & .Text1 & "' and a0d02 = " & Val(.Text3) & ""
               adoTaie.Execute "delete from acc0d1 where axd01 = '" & .Text1 & "' and axd02 = " & Val(.Text3) & ""
               .Frmacc4170_Clear 'Modify by Amy 2013/12/24
               .adoacc0d1.Requery
               .AdodcRefresh
               .SumShow
               .AdodcClear
               If .adoacc0d1.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            End If
            .FormDisabled
            .Text1.SetFocus
         End With
         adoTaie.RollbackTrans
         With Frmacc4170
            If strSaveConfirm = MsgText(4) Then
               .AdodcRefresh
               .SumShow
            End If
         End With
      Case "Frmacc4190"
         With Frmacc4190
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc060 where a0601 = " & Val(.Text6) & " and a0602 = '" & .Text5 & "' and a0603 = '" & .Text1 & "'"
               adoTaie.Execute "delete from acc061 where ax601 = " & Val(.Text6) & " and ax602 = '" & .Text5 & "' and ax603 = '" & .Text1 & "'"
               Frmacc4190_Clear
               .adoacc061.Requery
               .AdodcRefresh
               .SumShow
               If .adoacc061.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            End If
            .Command1.Enabled = False
            .DataGrid1.AllowUpdate = False
         End With
         adoTaie.RollbackTrans
         Frmacc4190.SetData ("F10") 'Add by Amy 2024/08/12
      Case "Frmacc41a0"
         With Frmacc41a0
            'add by sonia 2025/5/20 新增按放棄時刪除分錄ACC1P0資料(2025/5/21因為下面有RollbackTrans故先取消，改每天檢查)
'            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
'               adoTaie.Execute "delete from acc1p0 where a1p04 = '" & .Text10 & "'"
'            End If
            'end 2025/5/20
            .FormDisabled
            If .Text10.Tag <> "" Then 'Add by Morgan 2011/6/23
               'Frmacc41a0_Last 'Modify by Amy 2013/12/17
               .MoveLastRecord 'Add 2013/12/17
            End If 'Add by Morgan 2011/6/23
         End With
         adoTaie.RollbackTrans
      'add by nickc 2005/07/29
      Case "Frmacc41b0"
      With Frmacc41b0
         'Frmacc41b0_Last 'Modify by Amy 2013/12/17
         .MoveLastRecord 'Add 2013/12/17
      End With
      Case "Frmacc41d0"
         With Frmacc41d0
            .Frmacc41d0_Clear 'Modify by Amy 2014/02/06 搬回form
            .AdodcRefresh
            '.FormDisabled 'Mark by Amy 2014/02/11 往下搬
            .Text1.Enabled = True
            .Text2.Enabled = True
            .Text1.SetFocus
         End With
         adoTaie.RollbackTrans
      'Add by Morgan 2005/4/6
      'Mark by Amy 2017/12/18 往下搬,避免當掉
      'Case "Frmacc41e0"
       
      'Added by Lydia 2017/03/03
      Case "Frmacc41i0" '財產目錄
         adoTaie.RollbackTrans
         With Frmacc41i0
               'Modified by Lydia 2017/05/15
               'If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               If strSaveConfirm = MsgText(3) Then
                  adoTaie.Execute "delete from acc2b0 where a2b01 = '" & .txtA2B01 & "' "
                  .txtA2B01 = ""
                  .Frmacc41i0_Clear
                  .Acc2b0Refresh
               ElseIf strSaveConfirm = MsgText(4) Then
                   .FormShow
               End If
               .AdodcRefresh
               .SumShow
               .AdodcClear
               If .adoacc2b0.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            .FormDisabled
         End With
      'end 2017/03/03
      'Added by Lydia 2017/03/08
      Case "Frmacc41i0_1" '財產報廢作業
         adoTaie.RollbackTrans
         With Frmacc41i0_1
            'Modified by Lydia 2017/05/15
            '.FormShow
            .Frmacc41i0_1_Clear
            'end 2017/05/15
               .AdodcRefresh
               .SumShow
               .AdodcClear
               If .adoacc2b0.RecordCount <> 0 Then
                  .RecordShow
               Else
                  StatusClear
               End If
            .FormDisabled
         End With
      'end 2017/03/08
      'Add by Amy 2014/02/14
      Case "Frmacc5100"
         adoTaie.RollbackTrans
         Frmacc5100.FormShow
      Case "Frmacc5200"
         With Frmacc5200
            If strSaveConfirm = MsgText(3) And strControlButton <> MsgText(602) Then
               adoTaie.Execute "delete from acc040 where a0401 = " & Val(.Text6) & " and a0403 = '" & .Text4 & "' and a0404 = '" & .Text1 & "' and a0405 = '" & .Text3 & "'"
               Frmacc5200_Clear
               .QueryTable
            End If
            .Command1.Enabled = False
            .DataGrid1.AllowUpdate = False
            adoTaie.RollbackTrans
         End With
   End Select
   strSaveConfirm = MsgText(601)
   Select Case strFormName
      Case "Frmacc1110"
         tool9_enabled
      'Modify by Morgan 2006/3/20
      'Case "Frmacc1140", "Frmacc11d0"
      'Modify by Sindy 2014/1/9
      'Modify by Amy 2017/09/04 將frmacc114拆出(不需新增鈕)
      Case "Frmacc1140"
        tool8_enabled
      Case "Frmacc1130", "Frmacc11d0", "Frmacc11q0"
      'end 2017/09/04
         tool14_enabled
         If strFormName = "Frmacc11q0" Then
            Forms(0).Toolbar1.Buttons.Item(5).Enabled = False
         End If
      'Add by Morgan 2011/10/17
      Case "Frmacc1190"
         If Frmacc1190.Option1.Value = True Then
            Frmacc1190.FormShowE
         End If
         tool1_enabled
         
      'Add by Morgan 2007/4/17
      Case "Frmacc11i0"
         With Frmacc11i0
            .FormEnable
            .FormRequery
         End With
         tool1_enabled
         
      Case "Frmacc2111", "Frmacc21h1", "Frmacc21f1", "Frmacc21f2"
         tool7_enabled
      Case "Frmacc21i0"
         tool8_enabled
      'Modify end--------------------
      'Add by Morgan 2005/6/7
      Case "Frmacc11h0"
         Frmacc11h0.FormLocked False
         Frmacc11h0.txtA0w02.SetFocus
         tool7_enabled
         
      'Add by Morgan 2007/5/16
      Case "Frmacc11j0"
         With Frmacc11j0
            .FormEnable
            If .Text1.Tag <> "" Then
               .Text1.Text = .Text1.Tag
               .Command1_Click
            End If
         End With
         tool1_enabled
         
      'Add by Morgan 2007/10/8
      Case "Frmacc11k0"
         With Frmacc11k0
            .FormEnable
            If .txtCaseNo(0).Tag <> "" Then
               .txtCaseNo(0).Text = .txtCaseNo(0).Tag
               .txtCaseNo(1).Text = .txtCaseNo(1).Tag
               .txtCaseNo(2).Text = .txtCaseNo(2).Tag
               .txtCaseNo(3).Text = .txtCaseNo(3).Tag
               .txtCaseNo(4).Text = .txtCaseNo(4).Tag
               .Command1_Click
            End If
         End With
         tool1_enabled

      'Add by Morgan 2011/4/11
      Case "Frmacc11l0"
         With Frmacc11l0
            If .txtA0N01.Tag <> "" Then
               .txtA0N01 = .txtA0N01.Tag
               .QueryData
            End If
            .FormEnable
         End With
         tool1_enabled
         
         If Frmacc11l0.m_sCallType <> "" Then
            Unload Frmacc11l0
         End If
         
      'Add by Amy 2013/12/02
      Case "Frmacc11o0"
         tool1_enabled
         Frmacc11o0.FormShow
         Frmacc11o0.FormEnabled
         
      'Add by Morgan 2010/11/22
      Case "Frmacc21q0"
         'Modify by Morgan 2010/12/7 --要有新增功能
         'tool6_enabled
         'Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
         tool1_enabled
         'end 2010/12/7
         Frmacc21q0.FormShow
         
      'Add by Morgan 2006/12/18
      Case "Frmacc21r0"
         tool6_enabled
         Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True
         Frmacc21r0.FormShow
         
      'Added by Lydia 2016/11/07
      Case "Frmacc21w0"
         tool1_enabled
         If Frmacc21w0.txtKey.Tag <> "" Then
            Frmacc21w0.FormShow
         End If
         
      'Add by Amy 2014/11/14
       Case "Frmacc31c0"
            tool1_enabled
            Frmacc31c0.SetToolBar
            Frmacc31c0.Frmacc31c0_Clear 'Modify by Amy 2020/07/21
      'Add by Morgan 2011/6/23
      Case "Frmacc41a0"
         tool1_enabled
         Frmacc41a0.DisabledMoveRecord
      'Add by Amy 2014/02/11
      Case "Frmacc41d0"
         tool8_enabled
         Frmacc41d0.FormDisabled
      'Add by Amy 2017/08/18
      Case "Frmacc4120"
        Call Frmacc4120.SetData("BtExit")
      'Add by Amy 2017/04/27
      Case "Frmacc41j0"
        With Frmacc41j0
            .SetData ("F10")
        End With
        tool1_enabled
        Frmacc0000.Toolbar1.Buttons.Item(9).Enabled = False
      'Add by Amy 2017/12/18 從上面搬下來
      Case "Frmacc41e0"
         'adoTaie.RollbackTrans 'Removed by Morgan 2018/2/9 改存檔控制就好
         With Frmacc41e0
            .FormClear False
            .FormEnable
            If .txtA2301 = "" Then .txtA2301 = .txtA2301.Tag
            If .txtA2301 <> "" Then
               .ReadData .txtA2301
            End If
             tool1_enabled
            .txtA2301.SetFocus
         End With
      'Add by Amy 2014/02/14
      Case "Frmacc5100"
         tool8_enabled
         Frmacc5100.FormDisabled
         
      'Add by Sindy 2015/7/9
      'Modify by Sindy 2015/10/15 Mark
      'Modify by Sindy 2019/12/19
      Case "Frmacc11p0"
         If UCase(strUserLevel) <> UCase("Frmacc44w1") Then 'Frmacc44t0
            'tool6_enabled
            'Frmacc0000.Toolbar1.Buttons.Item(5).Enabled = True '修改
         'Else
            tool1_enabled
         End If
      '2015/7/9 END
      
      'Added by Morgan 2023/3/30
      Case "Frmacc2172"
         tool10_enabled
         Frmacc2172.FormShow
         Frmacc2172.FormEnabled
      'end 2023/3/30
      
      Case Else
         tool1_enabled
   End Select
   
   'Added by Morgan 2013/12/26
   If strFormName = "Frmacc1150" Then
      If Frmacc1150.m_AutoRun = True Then
         Unload Frmacc1150
      End If
   End If
   'end 2013/12/26
   
Checking:
   If Err.Number = 0 Or Err.Number = -2147168242 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  刪除

'*************************************************
Private Sub KeyEnterF5()
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   Dim stMsg As String 'Add by Amy 2022/05/13
   
On Error GoTo Checking

   If Frmacc0000.Toolbar1.Buttons.Item(8).Enabled = False Or strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   Select Case strFormName
      Case "Frmacc1110"
         If CheckUse("Frmacc1110", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
'      Case "Frmacc1130"
'         If CheckUse("Frmacc1130", strDel) = False Then
'            strSaveConfirm = MsgText(601)
'            Exit Sub
'         End If
     Case "Frmacc1140"
         If CheckUse("Frmacc1140", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc1150"
         If CheckUse("Frmacc1150", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc1150
            'Add by Morgan 2005/9/7 控制舊系統資料不可刪除
            If .MaskEdBox1 < "092/02/01" Then
               strSaveConfirm = MsgText(601)
               MsgBox "舊系統資料不可刪除！", vbExclamation
               Exit Sub
            End If
            '2005/9/7 end
            If .Adodc1.Recordset.RecordCount <> 0 Then
               If IsNull(.Adodc1.Recordset.Fields("a1p22").Value) = False Then
                  If .adoquery.State <> adStateClosed Then .adoquery.Close
                  .adoquery.CursorLocation = adUseClient
                  .adoquery.Open "select ax210 from acc021 where ax201 = '" & .Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & .Adodc1.Recordset.Fields("a1p22").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
                  If .adoquery.RecordCount <> 0 Then
                     MsgBox MsgText(158), , MsgText(5)
                     If .Text2.Enabled = True Then .Text2.SetFocus
                     .adoquery.Close
                     Exit Sub
                  End If
                  .adoquery.Close
               End If
            End If
            'Added by Morgan 2015/5/28
            If .DeleteCheck() = False Then
               Exit Sub
            End If
            'end 2015/5/28
         End With
      Case "Frmacc1160"
         If CheckUse("Frmacc1160", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         If CheckUse("Frmacc11p0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         If CheckUse("Frmacc11n0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      '2012/8/29 End
      Case "Frmacc1170"
         If CheckUse("Frmacc1170", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Add by Morgan 2008/1/2
         If Frmacc1170.EditCheck(1) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         
      Case "Frmacc1180"
         '2006/1/10 ADD BY SONIA
         With Frmacc1180
            If .adoquery.State = adStateOpen Then
               .adoquery.Close
            End If
            .adoquery.CursorLocation = adUseClient
            'Modified by Morgan 2023/12/19 改判斷已有傳票就不可刪除
            '.adoquery.Open "select ax210 from acc1p0, acc021 where a1p01 = ax201 and a1p22 = ax202 and a1p03 = ax203 and ax210 is not null and a1p04 = '" & .Text17 & "'", adoTaie, adOpenStatic, adLockReadOnly
            .adoquery.Open "select a1p22 from acc1p0 where a1p04 = '" & .Text17 & "' and a1p22 is not null", adoTaie, adOpenStatic, adLockReadOnly
            If .adoquery.RecordCount <> 0 Then
               'Modified by Morgan 2023/12/19 改判斷已有傳票就不可刪除
               'MsgBox MsgText(14), , MsgText(5)
               MsgBox "已有傳票不可刪除！", vbCritical
               'end 2023/12/19
               strControlButton = MsgText(602)
               .adoquery.Close
               Exit Sub
            End If
            .adoquery.Close
         End With
         '2006/1/10 END
         If CheckUse("Frmacc1180", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc1190"
         If CheckUse("Frmacc1190", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         '2007/6/27 ADD BY SONIA
         With Frmacc1190
            'Added by Morgan 2014/1/3
            If .EditCheck = False Then
               strControlButton = MsgText(602)
               Exit Sub
            End If
            'end 2014/1/3
            
            If .adoquery.State = adStateOpen Then
               .adoquery.Close
            End If
            .adoquery.CursorLocation = adUseClient
            .adoquery.Open "select ax210 from acc1p0, acc021 where a1p01 = ax201 and a1p22 = ax202 and a1p03 = ax203 and a1p04 = '" & .Text2 & "'", adoTaie, adOpenStatic, adLockReadOnly
            If .adoquery.RecordCount <> 0 Then
               MsgBox MsgText(158), , MsgText(5)
               strControlButton = MsgText(602)
               .adoquery.Close
               Exit Sub
            End If
            .adoquery.Close
         End With
         '2007/6/27 END
      Case "Frmacc11a0"
         If CheckUse("Frmacc11a0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc11a0
            'Add by Morgan 2005/10/19 檢查是否已沖
            If .CheckUsed = True Then
               strSaveConfirm = MsgText(601)
               Exit Sub
            End If
            
            If .Adodc1.Recordset.RecordCount <> 0 Then
               If IsNull(.Adodc1.Recordset.Fields("a1p22").Value) = False Then
                  .adoquery.CursorLocation = adUseClient
                  .adoquery.Open "select ax210 from acc021 where ax201 = '" & .Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & .Adodc1.Recordset.Fields("a1p22").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
                  If .adoquery.RecordCount <> 0 Then
                     MsgBox MsgText(158), , MsgText(5)
                     .Text1.SetFocus
                     .adoquery.Close
                     Exit Sub
                  End If
                  .adoquery.Close
               End If
            End If
         End With
      Case "Frmacc11f0"
         If CheckUse("Frmacc11f0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      'Add by Morgan 2007/4/16
      Case "Frmacc11i0"
         If CheckUse("Frmacc11i0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      'Add by Morgan 2007/5/16
      Case "Frmacc11j0"
         If CheckUse("Frmacc11j0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         
      'Add by Morgan 2007/10/5
      Case "Frmacc11k0"
         If CheckUse("Frmacc11k0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         
      'Add by Morgan 2011/4/11
      Case "Frmacc11l0"
         If CheckUse("Frmacc11l0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      'Add by Amy 2013/12/02
      Case "Frmacc11o0"
         If CheckUse("Frmacc11o0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         Frmacc11o0.FormDel
    
      Case "Frmacc2110"
         If CheckUse("Frmacc2110", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc2110
            If .Adodc1.Recordset.RecordCount <> 0 Then
               If IsNull(.Adodc1.Recordset.Fields("a1p22").Value) = False Then
                  .adoquery.CursorLocation = adUseClient
                  .adoquery.Open "select ax210 from acc021 where ax201 = '" & .Adodc1.Recordset.Fields("a1p01").Value & "' and ax202 = '" & .Adodc1.Recordset.Fields("a1p22").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
                  If .adoquery.RecordCount <> 0 Then
                     MsgBox MsgText(158), , MsgText(5)
                     .Text2.SetFocus
                     .adoquery.Close
                     Exit Sub
                  End If
                  .adoquery.Close
               End If
            End If
         End With
      Case "Frmacc2120"
         If CheckUse("Frmacc2120", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
          '2009/6/4 ADD BY SONIA
          '若傳票已過帳, 不可刪除國外暫收款資料
          With Frmacc2120
          
            'Add by Morgan 2011/4/12
            If .Text12 <> "1" Then
               MsgBox "暫收款類別不為1(暫收)時不可刪除！"
               strSaveConfirm = MsgText(601)
               Exit Sub
            End If
            
            '2011/4/15 modify by sonia 改為已產生傳票即不可刪除
            'StrSQLa = "Select Count(*) From ACC120, ACC1P0, ACC021 Where A1201=A1P04 And A1P22=AX202 And A1201='" & .Text2.Text & "' And AX210 Is Not Null "
            StrSQLa = "Select Count(*) From ACC1P0 Where A1P04='" & .Text2.Text & "' And A1p22 Is Not Null "
            rsA.CursorLocation = adUseClient
            rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
            If rsA.RecordCount > 0 Then
                If Val("" & rsA.Fields(0).Value) > 0 Then
                    MsgBox MsgText(158), , MsgText(5)
                    strSaveConfirm = ""
                    If rsA.State <> adStateClosed Then rsA.Close
                    Set rsA = Nothing
                    Exit Sub
                End If
            End If
            If rsA.State <> adStateClosed Then rsA.Close
            Set rsA = Nothing
          End With
          'End
      Case "Frmacc2130"
         If CheckUse("Frmacc2130", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'add by sonia 2017/1/16
         '若已產生國外付款資籵, 不可刪除國外暫收款退費資料
         With Frmacc2130
         
           StrSQLa = "Select Count(*) From ACC170 Where A1702='" & .Text2.Text & "' "
           rsA.CursorLocation = adUseClient
           rsA.Open StrSQLa, adoTaie, adOpenStatic, adLockReadOnly
           If rsA.RecordCount > 0 Then
               If Val("" & rsA.Fields(0).Value) > 0 Then
                   MsgBox MsgText(34), , MsgText(5)
                   strSaveConfirm = ""
                   If rsA.State <> adStateClosed Then rsA.Close
                   Set rsA = Nothing
                   Exit Sub
               End If
           End If
           If rsA.State <> adStateClosed Then rsA.Close
           Set rsA = Nothing
         End With
         'end 2017/1/16
      Case "Frmacc2140"
         If CheckUse("Frmacc2140", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
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
         
      'Added by Morgan 2023/3/30
      Case "Frmacc2172"
         If CheckUse("Frmacc2172", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         If Frmacc2172.EditCheck = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      'end 2023/3/30
      Case "Frmacc21d0"
         If CheckUse("Frmacc21d0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc21d0
            If .Adodc3.Recordset.RecordCount <> 0 Then
               If IsNull(.Adodc3.Recordset.Fields("a1p22").Value) = False Then
                  .adoquery.CursorLocation = adUseClient
                  .adoquery.Open "select ax210 from acc021 where ax201 = '" & .Adodc3.Recordset.Fields("a1p01").Value & "' and ax202 = '" & .Adodc3.Recordset.Fields("a1p22").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
                  If .adoquery.RecordCount <> 0 Then
                     MsgBox MsgText(158), , MsgText(5)
                     .Text1.SetFocus
                     .adoquery.Close
                     Exit Sub
                  End If
                  .adoquery.Close
               End If
            End If
         End With
      Case "Frmacc21e0"
         If CheckUse("Frmacc21e0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc21f0"
         If CheckUse("Frmacc21f0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc21g0"
         If CheckUse("Frmacc21g0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
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
'Remove by Morgan 2005/1/14 財務不用
'Modified by Morgan 2019/10/5 改回由財務維護
      Case "Frmacc21m0"
         If CheckUse("Frmacc21m0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         
      Case "Frmacc21n0"
         If CheckUse("Frmacc21n0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
'Remove by Morgan 2005/8/3 財務不用
'Modified by Morgan 2019/7/10 又改要用
      Case "Frmacc21o0"
         If CheckUse("Frmacc21o0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         
      'Add By Cheng 2003/07/23
      Case "Frmacc21q0"
         If CheckUse("Frmacc21q0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      
      'Added by Morgan 2019/10/5
      Case "Frmacc21s0"
         If CheckUse("Frmacc21s0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         
      'Added by Lydia 2016/11/07
      Case "Frmacc21w0"
         If CheckUse("Frmacc21w0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc3110"
         If CheckUse("Frmacc3110", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc3110
            If .Adodc2.Recordset.RecordCount <> 0 Then
               If IsNull(.Adodc2.Recordset.Fields("a1p22").Value) = False Then
                  .adoquery.CursorLocation = adUseClient
                  .adoquery.Open "select ax210 from acc021 where ax201 = '" & .Adodc2.Recordset.Fields("a1p01").Value & "' and ax202 = '" & .Adodc2.Recordset.Fields("a1p22").Value & "'", adoTaie, adOpenStatic, adLockReadOnly
                  If .adoquery.RecordCount <> 0 Then
                     MsgBox MsgText(158), , MsgText(5)
                     .Text11.SetFocus
                     .adoquery.Close
                     Exit Sub
                  End If
                  .adoquery.Close
               End If
            End If
         End With
      Case "Frmacc3120"
         If CheckUse("Frmacc3120", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Modify by Amy 2020/07/14 原程式搬回修改
         If Frmacc3120.DelCheck = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc3140"
         If CheckUse("Frmacc3140", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc3150"
         If CheckUse("Frmacc3150", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc3160"
         If CheckUse("Frmacc3160", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc3170"
         If CheckUse("Frmacc3170", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc3180"
         If CheckUse("Frmacc3180", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc3190"
         If CheckUse("Frmacc3190", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc31a0"
         If CheckUse("Frmacc31a0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc31c0"
         If CheckUse("Frmacc31c0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc4110"
         If CheckUse("Frmacc4110", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         'Add by Amy  2015/06/11 + 檢查傳票檔沒有該科目的資料才可以刪除
         ElseIf Frmacc4110.FormCheck = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc4120"
         If CheckUse("Frmacc4120", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'Add by Amy 2022/05/13 bug-系統產生會彈訊息,但仍可操作
         'Modify by Amy 2024/07/31 整合檢查程式,避免有未改到的
         Call Frmacc4120.SetData("F5")
         If strSaveConfirm = MsgText(601) Then
            Exit Sub
         End If
         If Frmacc4120.ChkForm("F5") = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'end 2024/07/31
      Case "Frmacc4130"
         If CheckUse("Frmacc4130", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc4140"
         If CheckUse("Frmacc4140", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc4160"
         If CheckUse("Frmacc4160", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc4170"
         If CheckUse("Frmacc4170", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         With Frmacc4170
            If .MaskEdBox2.Text <> MsgText(29) And .MaskEdBox2.Text <> MsgText(601) Then
               MsgBox MsgText(158), , MsgText(5)
               .Text1.SetFocus
               Exit Sub
            End If
         End With
      Case "Frmacc4180"
         If CheckUse("Frmacc4180", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc4190"
         If CheckUse("Frmacc4190", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc41a0"
         If CheckUse("Frmacc41a0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         
         'Add by Morgan 2011/6/23
         If Frmacc41a0.EditCheck = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         'end 2011/6/23
         
      Case "Frmacc41b0"
         If CheckUse("Frmacc41b0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc41d0"
         If CheckUse("Frmacc41d0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         
      'Added by Morgan 2015/6/17
      Case "Frmacc41e0"
         If CheckUse("Frmacc41e0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         If Frmacc41e0.DeleteCheck = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      'Added by Lydia 2017/03/03
      Case "Frmacc41i0" '財產目錄
         If CheckUse("Frmacc41i0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         If Frmacc41i0.EditCheck(1) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      'end 2017/03/03
      'Add by Amy 2017/04/21
      Case "Frmacc41j0"
        If CheckUse("Frmacc41j0", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
         If Frmacc41j0.FormCheck(0, "F5") = False Then
            Exit Sub
         End If
      Case "Frmacc5200"
         If CheckUse("Frmacc5200", strDel) = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
   End Select
   Select Case strFormName
      Case "Frmacc1170"
         With Frmacc1170
            .CreDebCheck
            If .CreDebCheck <> MsgText(602) Then
               MsgBox MsgText(11), , MsgText(5)
               Exit Sub
            End If
         End With
      'Mark by Amy 2023/12/06 刪除不用檢查借貸合計
'      Case "Frmacc4120"
'         With Frmacc4120
'            .CreDebCheck
'            If .CreDebCheck <> MsgText(602) Then
'               MsgBox MsgText(11), , MsgText(5)
'               Exit Sub
'            End If
'         End With
      Case "Frmacc4170"
         With Frmacc4170
            .CreDebCheck
            If .CreDebCheck <> MsgText(602) Then
               MsgBox MsgText(11), , MsgText(5)
               Exit Sub
            End If
         End With
      Case "Frmacc41d0"
         With Frmacc41d0
            .CreDebCheck
            If .CreDebCheck <> MsgText(602) Then
               MsgBox MsgText(11), , MsgText(5)
               Exit Sub
            End If
         End With
   End Select
   'modify by sonia 2017/10/24
   'strDelConfirm = MsgBox(MsgText(6), vbOKCancel + vbDefaultButton2, MsgText(5))
'   'Add By Sindy 2024/9/2
'   PUB_WriteDebugLog ("strFormName=" & strFormName & ";")
'   '2024/9/2 END
   If strFormName = "Frmacc11p0" Then
      strDelConfirm = MsgBox("若為 收據抬頭已新建客戶 的資料刪除，請先仔細確認所有欄位是否於客戶檔都有設定！！！請確認完再刪除！" & vbCrLf & _
                "是否確定要刪除？", vbOKCancel + vbDefaultButton2, MsgText(5))
   Else
      strDelConfirm = MsgBox(MsgText(6), vbOKCancel + vbDefaultButton2, MsgText(5))
   End If
   'end
   If strDelConfirm = vbCancel Then
      Exit Sub
   End If
   
   Select Case strFormName
      Case "Frmacc1110"
         Frmacc1110_Delete
         Frmacc1110_Clear
'      Case "Frmacc1130"
'         Frmacc1130_Delete
'         Frmacc1130_Clear
'         Frmacc1130.AdodcRefresh
      Case "Frmacc1140"
         Frmacc1140_Delete
         Frmacc1140.Frmacc1140_Clear 'Modify by Amy 2015/04/17 搬回form
         Frmacc1140.AdodcRefresh
      Case "Frmacc1150"
         Frmacc1150_Delete
         Frmacc1150_Clear
      Case "Frmacc1160"
         Frmacc1160_Delete
         Frmacc1160.Frmacc1160_Clear 'Modify by Amy 2022/03/11 原:Frmacc1160_Clear
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         Frmacc11p0.Frmacc11p0_Delete
         If strControlButton <> MsgText(602) Then 'Add By Sindy 2016/5/31 +if
            Frmacc11p0.Frmacc11p0_Clear
         End If
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         Frmacc11n0.Frmacc11n0_Delete
         Frmacc11n0.Frmacc11n0_Clear
      '2012/8/29 End
      Case "Frmacc1170"
        'Modify by Amy 2013/12/26
         Frmacc1170.Frmacc1170_Delete
         Frmacc1170.Frmacc1170_Clear
         'end 2013/12/26
      Case "Frmacc1180"
         'Modify by Amy 2014/01/17
         Frmacc1180.Frmacc1180_Delete
         Frmacc1180.Frmacc1180_Clear (True) 'Modify by Amy 2014/01/28 +參數
         'end 2014/01/17
         
      Case "Frmacc1190"
         Frmacc1190.m_AssignNo = Frmacc1190.Text1 'Add by Morgan 2011/5/30
         Frmacc1190.Frmacc1190_Delete
         Frmacc1190.Frmacc1190_Clear
         Frmacc1190.CheckAssign 'Add by Morgan 2011/5/30
         
      Case "Frmacc11a0"
         'Modify by Amy 2014/10/29 搬回form
         Frmacc11a0.Frmacc11a0_Delete
         Frmacc11a0.Frmacc11a0_Clear
      Case "Frmacc11f0"
         Frmacc11f0_Delete
         Frmacc11f0_Clear
      'Add by Morgan 2007/4/16
      Case "Frmacc11i0"
         If Pub_StrUserSt03 = "M51" Then
            If Frmacc11i0.FormDelete = False Then
               strSaveConfirm = MsgText(601)
               Exit Sub
            End If
         Else
            MsgBox "刪除權限未開放！"
         End If
      'Add by Morgan 2007/5/16
      Case "Frmacc11j0"
         If Frmacc11j0.FormDelete = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         End If
      'Add by Morgan 2007/10/5
      Case "Frmacc11k0"
         If Frmacc11k0.FormDelete = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         Else
            Frmacc11k0.FormClear
         End If
      'Add by Morgan 2011/4/11
      Case "Frmacc11l0"
         If Frmacc11l0.FormDelete = False Then
            strSaveConfirm = MsgText(601)
            Exit Sub
         Else
            Frmacc11l0.FormClear
            If Frmacc11l0.MovePrevious(False) = False Then
               Frmacc11l0.MoveFirst
            End If
         End If
         
      Case "Frmacc2110"
         Frmacc2110_Delete
         Frmacc2110_Clear
         Frmacc2110.Text2.Locked = False 'Add by Morgan 2006/6/20
      Case "Frmacc2120"
         Frmacc2120_Delete
         Frmacc2120_Clear
      Case "Frmacc2130"
         Frmacc2130_Delete
         Frmacc2130_Clear
      Case "Frmacc2140"
         Frmacc2140_Delete
         Frmacc2140_Clear
      Case "Frmacc2150"
         Frmacc2150_Delete
         Frmacc2150_Clear
      Case "Frmacc2160"
         Frmacc2160_Delete
         Frmacc2160_Clear
         
      'Added by Morgan 2023/3/30
      Case "Frmacc2172"
         Frmacc2172.FormDelete
      'end 2023/3/30
      
      Case "Frmacc21d0"
         Frmacc21d0_Delete
         Frmacc21d0.Frmacc21d0_Clear 'Modify by Amy 2014/11/04搬回form
         With Frmacc21d0
            .AdodcRefresh
         End With
      Case "Frmacc21e0"
         Frmacc21e0_Delete
         Frmacc21e0_Clear
      Case "Frmacc21f0"
         Frmacc21f0_Delete
         Frmacc21f0_Clear
'Remove by Morgan 2005/1/14 財務不需要
'            Case "Frmacc21g0"
'               Frmacc21g0_Delete
'               Frmacc21g0_Clear
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
'Remove by Morgan 2005/1/14 財務不用
'Modified by Morgan 2019/10/5 改回由財務維護
      Case "Frmacc21m0"
         Frmacc21m0_Delete Frmacc21m0
         Frmacc21m0_Clear Frmacc21m0
         
      Case "Frmacc21n0"
         Frmacc21n0_Delete
         Frmacc21n0_Clear
'Remove by Morgan 2005/8/3 財務不用
'Modified by Morgan 2019/7/10 又改要用
      Case "Frmacc21o0"
         Frmacc21o0_Delete Frmacc21o0
         Frmacc21o0_Clear Frmacc21o0
         
      'Add By Cheng 2003/07/23
      Case "Frmacc21q0"
         Frmacc21q0_Delete
         Frmacc21q0_Clear
      
      'Added by Morgan 2019/10/5
      Case "Frmacc21s0"
         Frmacc21s0_Delete Frmacc21s0
         Frmacc21s0_Clear Frmacc21s0
         
      'Added by Lydia 2016/11/07
      Case "Frmacc21w0"
         Frmacc21w0.FormDelete
      Case "Frmacc3110"
         Frmacc3110.Frmacc3110_Delete
         Frmacc3110.Frmacc3110_Clear
      Case "Frmacc3120"
         'Modify by Amy 2020/07/14 搬回程式
         Frmacc3120.Frmacc3120_Delete
         Frmacc3120.Frmacc3120_Clear
      Case "Frmacc3140"
         'Modify by Amy 2020/07/16
         Frmacc3140.Frmacc3140_Delete
         Frmacc3140.Frmacc3140_Clear
      Case "Frmacc3150"
         'Modify by Amy 2020/07/17
         Frmacc3150.Frmacc3150_Delete
         Frmacc3150.Frmacc3150_Clear
      Case "Frmacc3160"
         'Modify by Amy 2020/07/17
         Frmacc3160.Frmacc3160_Delete
         Frmacc3160.Frmacc3160_Clear
      Case "Frmacc3170"
         Frmacc3170.Frmacc3170_Delete 'Modify by Amy 2020/07/17
      Case "Frmacc3180"
         Frmacc3180_Delete
         Frmacc3180_Clear
      Case "Frmacc3190"
         Frmacc3190_Delete
         Frmacc3190_Clear
      Case "Frmacc31a0"
         'Modify by Amy 2020/07/21
         Frmacc31a0.Frmacc31a0_Delete
         Frmacc31a0.Frmacc31a0_Clear
      Case "Frmacc31c0"
         With Frmacc31c0
            If IsNull(.Adodc1.Recordset.Fields("A0E24").Value) = False Then
               MsgBox MsgText(87), , MsgText(5)
               Exit Sub
            End If
         End With
         'Modify by Amy 2020/07/21
         Frmacc31c0.Frmacc31c0_Delete
         Frmacc31c0.Frmacc31c0_Clear
      Case "Frmacc4110"
         'Modify by Amy 2015/06/11搬回form
         Frmacc4110.Frmacc4110_Delete
         Frmacc4110.Frmacc4110_Clear
      Case "Frmacc4120"
         'Memo by Amy 2023/12/06 原判斷是否過帳,搬回form
         'Modify by Amy 2014/01/14 搬回form
         'Modify by Amy 2024/07/31 搬至SaveData
         Call Frmacc4120.SaveData("F5")
      Case "Frmacc4130"
         Frmacc4130_Delete
         Frmacc4130_Clear
      Case "Frmacc4140"
         Frmacc4140_Delete
         Frmacc4140_Clear
      Case "Frmacc4160"
         Call Frmacc4160.Frmacc4160_Delete 'Modify by Amy 2024/08/23 程式搬回表單中
      Case "Frmacc4170"
         Frmacc4170_Delete
         Frmacc4170.Frmacc4170_Clear 'Modify by Amy 2013/12/24
      Case "Frmacc4180"
         Frmacc4180_Delete
         Frmacc4180_Clear
      Case "Frmacc4190"
         'Modify by Amy 2024/08/12 原程式搬回表單中
         Frmacc4190.SetData ("F5")
      Case "Frmacc41a0"
         Frmacc41a0_Delete
         'Frmacc41a0_Clear 'Modify by Amy 2013/12/17
         Frmacc41a0.FormClear 'Add 2013/12/17
      Case "Frmacc41b0"
         Frmacc41b0_Delete
      Case "Frmacc41d0"
        'Modify by Amy 2014/01/14 搬回form
         Frmacc41d0.Frmacc41d0_Clear
      'Add by Morgan 2006/10/17
      Case "Frmacc41e0"
         'Modified by Morgan 2015/6/17 開放權限(上面加檢查)
         'If Pub_StrUserSt03 = "M51" Then
            Frmacc41e0_Delete
         'Else
         '   MsgBox "刪除權限未開放！"
         'End If
         'end 2015/6/17
      'Added by Lydia 2017/03/03
      Case "Frmacc41i0"
         Frmacc41i0.Frmacc41i0_Delete
         Frmacc41i0.Frmacc41i0_Clear
      'end 2017/03/03
      'Add by Amy 2017/04/25
      Case "Frmacc41j0"
        Frmacc41j0.FormDel (0)
      Case "Frmacc5200"
         Frmacc5200_Delete
   End Select
   
Checking:
   If Err.Number = 0 Or Err.Number = -2147168242 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub

'*************************************************
'  查詢

'*************************************************
Private Sub KeyEnterF4()
   Dim StrSQLa As String
   Dim rsA As New ADODB.Recordset
   
On Error GoTo Checking

   If Frmacc0000.Toolbar1.Buttons.Item(9).Enabled = False Or strSaveConfirm = MsgText(3) Or strSaveConfirm = MsgText(4) Then
      Exit Sub
   End If
   strExitControl = MsgText(601)
   Select Case strFormName
      Case "Frmacc1130"
         strFormLink = strFormName
         Frmacc1130_Clear
         Frmacc1130.Enabled = False
         Frmacc1131.Show
      'Add By Sindy 2014/1/9
      Case "Frmacc11q0"
         strFormLink = strFormName
         Frmacc11q0.Frmacc11q0_Clear
         Frmacc11q0.Enabled = False
         Frmacc11q1.Show
      '2014/1/9 END
      Case "Frmacc1140"
         strFormLink = strFormName
         Frmacc1140.Frmacc1140_Clear 'Modify by Amy 2015/04/17 搬回form
         Frmacc1140.Enabled = False
         Frmacc1141.Show
      Case "Frmacc1150"
         'Add by Amy 2018/06/05 41字頭及7121不可輸小數
         If Frmacc1150.ChkDot = True Then
           MsgBox "41字頭或7121科目不可輸入小數！", , MsgText(5)
           Exit Sub
         End If
        'end 2018/06/05
         'Add by Morgan 2005/9/29 檢查借貸平衡
         If Frmacc1150.CreDebCheck <> MsgText(602) Then
            MsgBox MsgText(11), , MsgText(5)
            Exit Sub
         End If
         '2005/9/29 end
         Frmacc1150_Clear
         Frmacc1150.Enabled = False
         Frmacc1152.Show
      Case "Frmacc1160"
         Frmacc1160.Frmacc1160_Clear 'Modify by Amy 2022/03/11 原:Frmacc1160_Clear
         Frmacc1160.Enabled = False
         Frmacc1161.Show
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
      'Add By Sindy 2013/12/13
      Case "Frmacc11o5"
         Frmacc11o5.Frmacc11o5_Clear
         Frmacc11o5.Enabled = False
         Frmacc11o6.Show
      '2013/12/13 End
      'Add By Amy 2013/12/02
      Case "Frmacc11o0"
         Frmacc11o0.FormClear
         Frmacc11o0.Enabled = False
         Frmacc11o1.Show
      Case "Frmacc1170"
         Frmacc1170.Frmacc1170_Clear 'Modify by Amy 2013/12/26
         Frmacc1170.Enabled = False
         Frmacc1171.Show
      Case "Frmacc1180"
         Frmacc1180.Frmacc1180_Clear (True) 'Modify by Amy 2014/01/28+參數
         Frmacc1180.Enabled = False
         Frmacc1181.Show
      Case "Frmacc1190"
         Frmacc1190.Frmacc1190_Clear
         Frmacc1190.Enabled = False
         Frmacc1193.Show
      Case "Frmacc11a0"
         Frmacc11a0.Frmacc11a0_Clear 'Modify by Amy 2014/10/29 搬回from
         Frmacc11a0.Enabled = False
         Frmacc11a1.Show
      'Add By Sindy 2015/8/11
      Case "Frmacc11c0"
         Frmacc11c0.Enabled = False
         Frmacc11c1.Show
      '2015/8/11 END
      Case "Frmacc11d0"
         Frmacc11d0_Clear
         Frmacc11d0.Enabled = False
         Frmacc11d1.Show
      Case "Frmacc11f0"
         Frmacc11f0_Clear
         Frmacc11f0.Enabled = False
         Frmacc11f1.Show
      'Add by Morgan 2007/4/17
      Case "Frmacc11i0"
         With Frmacc11i0
            .FormClear
            .Enabled = False
         End With
         Frmacc11i1.Show
      'Add by Morgan 2007/5/18
      Case "Frmacc11j0"
         With Frmacc11j0
            .FormClear
            .Enabled = False
         End With
         Frmacc11j1.Show
         
      'Add by Morgan 2007/10/8 暫時沒
      Case "Frmacc11k0"
         Exit Sub
         
      'Add by Morgan 2011/4/12
      Case "Frmacc11l0"
         With Frmacc11l0
            .FormClear
            .Enabled = False
         End With
         Frmacc11l1.Show
         
      Case "Frmacc2110"
         'Add by Morgan 2006/6/20
         If Frmacc2110.bolForm2 = True Then
            MsgBox "匯率有異動，請點【收款資料】以便重新計算台幣收款金額！", vbExclamation
            Exit Sub
         End If
         'end 2006/6/20
         Frmacc2110_Clear
         Frmacc2110.Enabled = False
         Frmacc2112.Show
      Case "Frmacc2120"
         Frmacc2120_Clear
         Frmacc2120.Enabled = False
         Frmacc2121.Show
      Case "Frmacc2130"
         Frmacc2130_Clear
         Frmacc2130.Enabled = False
         Frmacc2131.Show
      Case "Frmacc2140"
         Frmacc2140_Clear
         Frmacc2140.Enabled = False
         Frmacc2141.Show
      Case "Frmacc2150"
         strFormLink = strFormName
         Frmacc2150_Clear
         Frmacc2150.Enabled = False
         Frmacc2151.Show
      Case "Frmacc2160"
         Frmacc2160_Clear
         Frmacc2160.Enabled = False
         Frmacc2161.Show
      Case "Frmacc21d0"
         Frmacc21d0.Frmacc21d0_Clear 'Modify by Amy 2014/11/04搬回form
         Frmacc21d0.Enabled = False
         Frmacc21d1.Show
      Case "Frmacc21e0"
         Frmacc21e0_Clear
         Frmacc21e0.Enabled = False
         Frmacc21e1.Show
      Case "Frmacc21f0"
         Frmacc21f0_Clear
         Frmacc21f0.Enabled = False
         Frmacc21f3.Show
'Remove by Morgan 2005/1/14 財務不需要
'            Case "Frmacc21g0"
'               Frmacc21g0_Clear
'               Frmacc21g0.Enabled = False
'               Frmacc21g1.Show
'2005/1/14 end
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
'Remove by Morgan 2005/1/14 財務不用
'Modified by Morgan 2019/10/5 改回由財務維護
      Case "Frmacc21m0"
         Exit Sub
         
      Case "Frmacc21n0"
         Exit Sub
'Remove by Morgan 2005/8/3 財務不用
'Modified by Morgan 2019/7/10 又改要用
      Case "Frmacc21o0"
         Frmacc21o0_Clear Frmacc21o0
         Frmacc21o0.Enabled = False
         Frmacc21o1.Show
      
      'Added by Morgan 2019/10/5
      Case "Frmacc21s0"
         Exit Sub
         
      'Added by Lydia 2016/11/07
      Case "Frmacc21w0"
         Exit Sub
      'Add by Amy 2015/08/31
      Case "Frmacc2210"
         Frmacc2210.Enabled = False
         frm100114_5.Tag = "Frmacc2210"
         frm100114_5.Show
      Case "Frmacc3110"
         Frmacc3110.Frmacc3110_Clear
         Frmacc3110.Enabled = False
         Frmacc3111.Show
      Case "Frmacc3120"
         Frmacc3120.Frmacc3120_Clear 'Modify by Amy 2020/07/14
         Frmacc3120.Enabled = False
         Frmacc3121.Show
      Case "Frmacc3130"
         Frmacc3130_Clear
         Frmacc3130.Enabled = False
'               Frmacc3131.Show
      Case "Frmacc3140"
         Frmacc3140.Frmacc3140_Clear 'Modify by Amy 2020/07/16
         Frmacc3140.Enabled = False
         'Frmacc3141.Show ''Mark by Amy 2022/02/23 過濾Form2.0 未使用表單,發現此功能未顯示,故先刪
      Case "Frmacc3150"
         Frmacc3150.Frmacc3150_Clear 'Modify by Amy 2020/07/17
         Frmacc3150.Enabled = False
         Frmacc3151.Show
      Case "Frmacc3160"
         Frmacc3160.Frmacc3160_Clear 'Modify by Amy 2020/07/17
         Frmacc3160.Enabled = False
         Frmacc3161.Show
      Case "Frmacc3170"
         Frmacc3170.Frmacc3170_Clear 'Modify by Amy 2020/07/17
         With Frmacc3170
            .AdodcClear
            .AdodcRefresh
         End With
         Frmacc3170.Enabled = False
         Frmacc3171.Show
      Case "Frmacc3180"
         Frmacc3180_Clear
         Frmacc3180.Enabled = False
         Frmacc3181.Show
      Case "Frmacc3190"
         Frmacc3190_Clear
         Frmacc3190.Enabled = False
         Frmacc3191.Show
      Case "Frmacc31a0"
         Frmacc31a0.Frmacc31a0_Clear 'Modify by Amy 2020/07/21
         Frmacc31a0.Enabled = False
         Frmacc31a1.Show
      Case "Frmacc31c0"
         Frmacc31c0.Frmacc31c0_Clear 'Modify by Amy 2020/07/21
         Frmacc31c0.Enabled = False
         Frmacc31c1.Show
      Case "Frmacc4110"
         Frmacc4110.Frmacc4110_Clear 'Modify by Amy 2015/06/11搬回form
         Frmacc4110.Enabled = False
         Frmacc4111.Show
      Case "Frmacc4120"
         With Frmacc4120
            .Frmacc4120_Clear 'Modify by Amy 2014/01/14 搬回form
            .AdodcClear
            .AdodcRefresh
         End With
         Frmacc4120.Enabled = False
         Frmacc4121.Show
      Case "Frmacc4130"
         Frmacc4130_Clear
         Frmacc4130.Enabled = False
         Frmacc4131.Show
      Case "Frmacc4140"
         Frmacc4140_Clear
         Frmacc4140.Enabled = False
         Frmacc4141.Show
      Case "Frmacc4150"
         Frmacc4150_Clear
'               Frmacc4150.Enabled = False
'               Frmacc4151.Show
      Case "Frmacc4160"
         strFormLink = strFormName
         Frmacc4160.Frmacc4160_Clear 'Modify by Amy 2013/12/24
         Frmacc4160.Enabled = False
         Frmacc4161.Show
      Case "Frmacc4170"
         Frmacc4170.Frmacc4170_Clear  'Modify by Amy 2013/12/24
         Frmacc4170.Enabled = False
         Frmacc4171.Show
      Case "Frmacc4180"
         Frmacc4180_Clear
         Frmacc4180.Enabled = False
         Frmacc4181.Show
      Case "Frmacc4190"
         Frmacc4190_Clear
         Frmacc4190.Enabled = False
         Frmacc4191.Show
      Case "Frmacc41a0"
         Frmacc41a0.Enabled = False
         Frmacc41a1.Show
      Case "Frmacc41b0"
         Frmacc41b0.Enabled = False
         Frmacc41b1.Show
      Case "Frmacc41d0"
         'Frmacc41d0_Clear 'Modify by Amy 2014/02/06
         With Frmacc41d0
            .Frmacc41d0_Clear 'Add by Amy 2014/02/06 搬回form
            .AdodcClear
            .AdodcRefresh
         End With
         Frmacc41d0.Enabled = False
         Frmacc41d1.Show
      'Add by Morgan 2005/4/7
      Case "Frmacc41e0"
         With Frmacc41e0
            .Enabled = False
         End With
         Frmacc41e1.Show
'end 2017/03/03
      'Added by Lydia 2017/05/19
      Case "Frmacc41i0"  '財產目錄
         Frmacc41i0.Frmacc41i0_Clear
         Frmacc41i0.Enabled = False
         Frmacc41i1.iType = "1"
         Frmacc41i1.Show
      Case "Frmacc41i0_1" '財產報廢
         Frmacc41i0_1.Frmacc41i0_1_Clear
         Frmacc41i0_1.Enabled = False
         Frmacc41i1.iType = "2"
         Frmacc41i1.Show
      'end 2017/05/19
      Case "Frmacc5200"
         strFormLink = strFormName
         Frmacc5200_Clear
         Frmacc5200.Enabled = False
         Frmacc4161.Show
      Case Else
         tool1_enabled
   End Select
   tool3_enabled
'         strExitControl = "Y"
Checking:
   If Err.Number = 0 Or Err.Number = -2147168242 Then
      Exit Sub
   End If
   MsgBox Err.Description, , MsgText(5)
End Sub
'*************************************************
'  功能鍵對照函式
'
'*************************************************
Public Sub KeyEnter(InputCode As Integer)
   
   Frmacc0000.StatusBar1.Panels(1).Text = MsgText(601)
   Select Case InputCode
      Case vbKeyEscape '離開
         FormExit
      Case vbKeyF2 '新增
         KeyEnterF2
      Case vbKeyF3 '修改
         KeyEnterF3
      Case vbKeyF9 '存檔
         KeyEnterF9
      Case vbKeyF10 '取消
         KeyEnterF10
      Case vbKeyF5 '刪除
         KeyEnterF5
      Case vbKeyF4 '查詢
         KeyEnterF4
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
         If PUB_Connect2DB() = False Then
            End
         Else
            Set adoTaie = cnnConnection
            'edit by nickc 2007/02/07 不用 dll 了
            'objPublicData.Connection = cnnConnection
         End If
'2005/12/14 end

        'Removed by Morgan 2025/9/9 沒用了
        'adoTemp.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Persist Security Info=False;Data Source=" & App.path & "\finance.mdb"
        'adoTemp.Open
        'end 2025/9/9

        'Add By Cheng 2003/05/02
        'edit by nickc 2007/02/07 不用 dll 了
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
      .Main4.Enabled = True
      '.Main5.Enabled = True
      '.Main6.Enabled = True
      .Main7.Enabled = True
      .Main8.Enabled = True
      .Main9.Enabled = True
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
      .Main4.Enabled = False
      '.Main5.Enabled = False
      '.Main6.Enabled = False
      .Main7.Enabled = False
      .Main8.Enabled = False
      .Main9.Enabled = False
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
      Case "Frmacc1130"
         Frmacc1130_Last
      'Add By Sindy 2014/1/9
      Case "Frmacc11q0"
         Frmacc11q0.Frmacc11q0_Last
      '2014/1/9 END
      Case "Frmacc1140"
         Frmacc1140_Last
      Case "Frmacc1150"
         Frmacc1150.Frmacc1150_Last 'Modify by Amy 2018/06/05 搬回form
      Case "Frmacc1160"
         Frmacc1160_Last
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         Frmacc11p0.Frmacc11p0_Last
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         Frmacc11n0.Frmacc11n0_Last
      '2012/8/29 End
      Case "Frmacc1170"
         Frmacc1170.Frmacc1170_Last
      Case "Frmacc1180"
         Frmacc1180.Frmacc1180_Last 'Modify by Amy 2014/01/17
      Case "Frmacc1190"
         Frmacc1190_Last
      Case "Frmacc11a0"
         Frmacc11a0_Last
      Case "Frmacc11d0"
         Frmacc11d0_Last
      Case "Frmacc11f0"
         Frmacc11f0_Last
      'Add by Morgan 2007/4/17
      Case "Frmacc11i0"
         Frmacc11i0.MoveLast
      'Add by Morgan 2007/5/24
      Case "Frmacc11j0"
         Frmacc11j0.MoveLast
         
      'Add by Morgan 2011/4/11
      Case "Frmacc11l0"
         Frmacc11l0.MoveLast
      'Add by Amy 2013/11/29
      Case "Frmacc11o0"
         Frmacc11o0.GetLastRecordVal  'movelast
      Case "Frmacc2110"
         'Add by Morgan 2006/6/20
         If Frmacc2110.bolForm2 = True Then
            MsgBox "匯率有異動，請點【收款資料】以便重新計算台幣收款金額！", vbExclamation
            Exit Sub
         End If
         'end 2006/6/20
         Frmacc2110_Last
      Case "Frmacc2120"
         Frmacc2120_Last
      Case "Frmacc2130"
         Frmacc2130_Last
      Case "Frmacc2140"
         Frmacc2140_Last
      Case "Frmacc2150"
         Frmacc2150_Last
      Case "Frmacc2160"
         Frmacc2160_Last
      Case "Frmacc21d0"
         Frmacc21d0_Last
      Case "Frmacc21e0"
         Frmacc21e0_Last
      Case "Frmacc21f0"
         Frmacc21f0_Last
'Remove by Morgan 2005/1/14 財務不需要
'      Case "Frmacc21g0"
'         Frmacc21g0_Last
      Case "Frmacc21h0"
         Frmacc21h0_Last
      Case "Frmacc21i0"
         Frmacc21i0_Last
      Case "Frmacc21j0"
         Frmacc21j0_Last
      Case "Frmacc21k0"
         Frmacc21k0_Last
'Remove by Morgan 2005/1/14 財務不用
'Modified by Morgan 2019/10/5 改回由財務維護
      Case "Frmacc21m0"
         Frmacc21m0_Last Frmacc21m0
         
      Case "Frmacc21n0"
         Frmacc21n0_Last
'Remove by Morgan 2005/8/3 財務不用
'Modified by Morgan 2019/7/10 又改要用
      Case "Frmacc21o0"
         Frmacc21o0_Last Frmacc21o0
         
      'Add By Cheng 2003/07/23
      Case "Frmacc21q0"
         Frmacc21q0_Last
      'Add by Morgan 2006/12/18
      Case "Frmacc21r0"
         Frmacc21r0.MoveLast
      
      'Added by Morgan 2019/10/5
      Case "Frmacc21s0"
         Frmacc21s0_Last Frmacc21s0
         
      'Added by Lydia 2016/11/07
      Case "Frmacc21w0"
         Frmacc21w0.MoveLast
      Case "Frmacc3110"
         Frmacc3110_Last
      Case "Frmacc3120"
         Frmacc3120_Last
      Case "Frmacc3140"
         Frmacc3140_Last
      Case "Frmacc3150"
         Frmacc3150_Last
      Case "Frmacc3160"
         Frmacc3160_Last
      Case "Frmacc3170"
         Frmacc3170_Last
      Case "Frmacc3180"
         Frmacc3180_Last
      Case "Frmacc3190"
         Frmacc3190_Last
      Case "Frmacc31a0"
         Frmacc31a0_Last
      Case "Frmacc31c0"
         Frmacc31c0_Last
      Case "Frmacc4110"
         Frmacc4110.Frmacc4110_Last 'Modify by Amy 2015/06/11 搬回form
      Case "Frmacc4120"
         Frmacc4120.Frmacc4120_Last 'Modify by Amy 2014/01/14 搬回form
      Case "Frmacc4130"
         Frmacc4130_Last
      Case "Frmacc4140"
         Frmacc4140_Last
      Case "Frmacc4150"
         Frmacc4150_Last
      Case "Frmacc4160"
         Frmacc4160_Last
      Case "Frmacc4170"
         Frmacc4170_Last
      Case "Frmacc4180"
         Frmacc4180_Last
      Case "Frmacc4190"
         Frmacc4190_Last
      Case "Frmacc41a0"
         'Frmacc41a0_Last 'Modify by Amy 2013/12/17
         Frmacc41a0.MoveLastRecord 'Add 2013/12/17
      Case "Frmacc41b0"
         'Frmacc41b0_Last 'Modify by Amy 2013/12/17
         Frmacc41b0.MoveLastRecord 'Add 2013/12/17
      'Add by Morgan 2004/10/27
      Case "Frmacc41c0"
         Frmacc41c0_Last
      Case "Frmacc41d0"
         Frmacc41d0.Frmacc41d0_Last 'Modify by Amy 2014/02/06 搬回form
      'Add by Morgan 2005/4/7
      Case "Frmacc41e0"
         With Frmacc41e0
            .ReadData .txtA2301, 4
         End With
      'Added by Lydia 2017/03/03 財產目錄
      Case "Frmacc41i0"
         Frmacc41i0.Frmacc41i0_Last
      'Added by Lydia 2017/03/08 財產報廢作業
      Case "Frmacc41i0_1"
         Frmacc41i0_1.Frmacc41i0_1_Last
      'Add by Amy 2017/04/25
      Case "Frmacc41j0"
         Frmacc41j0.SetData ("LastRec")
      'Add by Amy 2014/02/14
      Case "Frmacc5100"
         Frmacc5100.MoveLastRecord
      Case "Frmacc5200"
         Frmacc5200_Last
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
      Case "Frmacc1130"
         Frmacc1130_Next
      'Add By Sindy 2014/1/9
      Case "Frmacc11q0"
         Frmacc11q0.Frmacc11q0_Next
      '2014/1/9 END
      Case "Frmacc1140"
         Frmacc1140_Next
      Case "Frmacc1150"
         Frmacc1150.Frmacc1150_Next 'Modify by Amy 2018/06/05
      Case "Frmacc1160"
         Frmacc1160_Next
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         Frmacc11p0.Frmacc11p0_Next
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         Frmacc11n0.Frmacc11n0_Next
      '2012/8/29 End
     'Add By Amy 2013/11/29
      Case "Frmacc11o0"
         Frmacc11o0.GetNextRecordVal '.MoveNext
      Case "Frmacc1170"
         Frmacc1170.Frmacc1170_Next
      Case "Frmacc1180"
         Frmacc1180.Frmacc1180_Next 'Modify by Amy 2014/01/17
      Case "Frmacc1190"
         Frmacc1190_Next
      Case "Frmacc11a0"
         Frmacc11a0_Next
      Case "Frmacc11d0"
         Frmacc11d0_Next
      Case "Frmacc11f0"
         Frmacc11f0_Next
      'Add by Morgan 2007/4/17
      Case "Frmacc11i0"
         Frmacc11i0.MoveNext
      'Add by Morgan 2007/5/24
      Case "Frmacc11j0"
         Frmacc11j0.MoveNext
         
      'Add by Morgan 2011/4/11
      Case "Frmacc11l0"
         Frmacc11l0.MoveNext
         
      Case "Frmacc2110"
         'Add by Morgan 2006/6/20
         If Frmacc2110.bolForm2 = True Then
            MsgBox "匯率有異動，請點【收款資料】以便重新計算台幣收款金額！", vbExclamation
            Exit Sub
         End If
         'end 2006/6/20
         Frmacc2110_Next
      Case "Frmacc2120"
         Frmacc2120_Next
      Case "Frmacc2130"
         Frmacc2130_Next
      Case "Frmacc2140"
         Frmacc2140_Next
      Case "Frmacc2150"
         Frmacc2150_Next
      Case "Frmacc2160"
         Frmacc2160_Next
      Case "Frmacc21d0"
         Frmacc21d0_Next
      Case "Frmacc21e0"
         Frmacc21e0_Next
      Case "Frmacc21f0"
         Frmacc21f0_Next
'Remove by Morgan 2005/1/14 財務不需要
'      Case "Frmacc21g0"
'         Frmacc21g0_Next
      Case "Frmacc21h0"
         Frmacc21h0_Next
      Case "Frmacc21i0"
         Frmacc21i0_Next
      Case "Frmacc21j0"
         Frmacc21j0_Next
      Case "Frmacc21k0"
         Frmacc21k0_Next
'Remove by Morgan 2005/1/14 財務不用
'Modified by Morgan 2019/10/5 改回由財務維護
      Case "Frmacc21m0"
         Frmacc21m0_Next Frmacc21m0
         
      Case "Frmacc21n0"
         Frmacc21n0_Next
'Remove by Morgan 2005/8/3 財務不用
'Modified by Morgan 2019/7/10 又改要用
      Case "Frmacc21o0"
         Frmacc21o0_Next Frmacc21o0
         
      'Add By Cheng 2003/07/23
      Case "Frmacc21q0"
         Frmacc21q0_Next
      'Add by Morgan 2006/12/18
      Case "Frmacc21r0"
         Frmacc21r0.MoveNext
      
      'Added by Moran 2019/10/5
      Case "Frmacc21s0"
         Frmacc21s0_Next Frmacc21s0
         
      'Added by Lydia 2016/11/07
      Case "Frmacc21w0"
         Frmacc21w0.MoveNext
      Case "Frmacc3110"
         Frmacc3110_Next
      Case "Frmacc3120"
         Frmacc3120_Next
      Case "Frmacc3140"
         Frmacc3140_Next
      Case "Frmacc3150"
         Frmacc3150_Next
      Case "Frmacc3160"
         Frmacc3160_Next
      Case "Frmacc3170"
         Frmacc3170_Next
      Case "Frmacc3180"
         Frmacc3180_Next
      Case "Frmacc3190"
         Frmacc3190_Next
      Case "Frmacc31a0"
         Frmacc31a0_Next
      Case "Frmacc31c0"
         Frmacc31c0_Next
      Case "Frmacc4110"
         Frmacc4110.Frmacc4110_Next 'Modify by Amy 2015/06/11 搬回Form
      Case "Frmacc4120"
         Frmacc4120.Frmacc4120_Next 'Modify by Amy 2014/01/14 搬回form
      Case "Frmacc4130"
         Frmacc4130_Next
      Case "Frmacc4140"
         Frmacc4140_Next
      Case "Frmacc4150"
         Frmacc4150_Next
      Case "Frmacc4160"
         Frmacc4160_Next
      Case "Frmacc4170"
         Frmacc4170_Next
      Case "Frmacc4180"
         Frmacc4180_Next
      Case "Frmacc4190"
         Frmacc4190_Next
      Case "Frmacc41a0"
         'Frmacc41a0_Next 'Modify by Amy 2013/12/17
         Frmacc41a0.MoveNextRecord 'Add 2013/12/17
      Case "Frmacc41b0"
         'Frmacc41b0_Next 'Modify by Amy 2013/12/17
         Frmacc41b0.MoveNextRecord 'Add 2013/12/17
      'Add by Morgan 2004/10/27
      Case "Frmacc41c0"
         Frmacc41c0_Next
      Case "Frmacc41d0"
         Frmacc41d0.Frmacc41d0_Next 'Modify by Amy 2014/02/06 搬回form
      'Add by Morgan 2005/4/7
      Case "Frmacc41e0"
         With Frmacc41e0
            .ReadData .txtA2301, 3
         End With
      'Added by Lydia 2017/03/03 財產目錄
      Case "Frmacc41i0"
         Frmacc41i0.Frmacc41i0_Next
      'Added by Lydia 2017/03/08 財產報廢作業
      Case "Frmacc41i0_1"
         Frmacc41i0_1.Frmacc41i0_1_Next
      'Add by Amy 2017/04/25
      Case "Frmacc41j0"
         Frmacc41j0.SetData ("NextRec")
      'Add by Amy 2014/02/14
      Case "Frmacc5100"
         Frmacc5100.MoveNextRecord
      Case "Frmacc5200"
         Frmacc5200_Next
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
      Case "Frmacc1130"
         Frmacc1130_Previous
      'Add By Sindy 2014/1/9
      Case "Frmacc11q0"
         Frmacc11q0.Frmacc11q0_Previous
      '2014/1/9 END
      Case "Frmacc1140"
         Frmacc1140_Previous
      Case "Frmacc1150"
         Frmacc1150.Frmacc1150_Previous 'Modify by Amy 2018/06/05
      Case "Frmacc1160"
         Frmacc1160_Previous
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         Frmacc11p0.Frmacc11p0_Previous
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         Frmacc11n0.Frmacc11n0_Previous
      '2012/8/29 End
     'Add By Amy 2013/11/29
      Case "Frmacc11o0"
         Frmacc11o0.GetPreRecordVal '.MovePrevious
      Case "Frmacc1170"
         Frmacc1170.Frmacc1170_Previous
      Case "Frmacc1180"
         Frmacc1180.Frmacc1180_Previous 'Modify by Amy 2014/01/17
      Case "Frmacc1190"
         Frmacc1190_Previous
      Case "Frmacc11a0"
         Frmacc11a0_Previous
      Case "Frmacc11d0"
         Frmacc11d0_Previous
      Case "Frmacc11f0"
         Frmacc11f0_Previous
      'Add by Morgan 2007/4/17
      Case "Frmacc11i0"
         Frmacc11i0.MovePrevious
      'Add by Morgan 2007/5/24
      Case "Frmacc11j0"
         Frmacc11j0.MovePrevious
         
      'Add by Morgan 2011/4/11
      Case "Frmacc11l0"
         Frmacc11l0.MovePrevious
         
      Case "Frmacc2110"
         'Add by Morgan 2006/6/20
         If Frmacc2110.bolForm2 = True Then
            MsgBox "匯率有異動，請點【收款資料】以便重新計算台幣收款金額！", vbExclamation
            Exit Sub
         End If
         'end 2006/6/20
         Frmacc2110_Previous
      Case "Frmacc2120"
         Frmacc2120_Previous
      Case "Frmacc2130"
         Frmacc2130_Previous
      Case "Frmacc2140"
         Frmacc2140_Previous
      Case "Frmacc2150"
         Frmacc2150_Previous
      Case "Frmacc2160"
         Frmacc2160_Previous
      Case "Frmacc21d0"
         Frmacc21d0_Previous
      Case "Frmacc21e0"
         Frmacc21e0_Previous
      Case "Frmacc21f0"
         Frmacc21f0_Previous
'Remove by Morgan 2005/1/14 財務不需要
'      Case "Frmacc21g0"
'         Frmacc21g0_Previous
      Case "Frmacc21h0"
         Frmacc21h0_Previous
      Case "Frmacc21i0"
         Frmacc21i0_Previous
      Case "Frmacc21j0"
         Frmacc21j0_Previous
      Case "Frmacc21k0"
         Frmacc21k0_Previous
'Remove by Morgan 2005/1/14 財務不用
'Modified by Morgan 2019/10/5 改回由財務維護
      Case "Frmacc21m0"
         Frmacc21m0_Previous Frmacc21m0
         
      Case "Frmacc21n0"
         Frmacc21n0_Previous
'Remove by Morgan 2005/8/3 財務不用
'Modified by Morgan 2019/7/10 又改要用
      Case "Frmacc21o0"
         Frmacc21o0_Previous Frmacc21o0
         
      'Add By Cheng 2003/07/23
      Case "Frmacc21q0"
         Frmacc21q0_Previous
      'Add by Morgan 2006/12/18
      Case "Frmacc21r0"
         Frmacc21r0.MovePrevious
      
      'Added by Morgan 2019/10/5
      Case "Frmacc21s0"
         Frmacc21s0_Previous Frmacc21s0
         
      'Added by Lydia 2016/11/07
      Case "Frmacc21w0"
         Frmacc21w0.MovePrevious
      Case "Frmacc3110"
         Frmacc3110_Previous
      Case "Frmacc3120"
         Frmacc3120_Previous
      Case "Frmacc3140"
         Frmacc3140_Previous
      Case "Frmacc3150"
         Frmacc3150_Previous
      Case "Frmacc3160"
         Frmacc3160_Previous
      Case "Frmacc3170"
         Frmacc3170_Previous
      Case "Frmacc3180"
         Frmacc3180_Previous
      Case "Frmacc3190"
         Frmacc3190_Previous
      Case "Frmacc31a0"
         Frmacc31a0_Previous
      Case "Frmacc31c0"
         Frmacc31c0_Previous
      Case "Frmacc4110"
         Frmacc4110.Frmacc4110_Previous 'Modify by Amy 2015/06/11 搬回form
      Case "Frmacc4120"
         Frmacc4120.Frmacc4120_Previous 'Modify by Amy 2014/01/14 搬回form
      Case "Frmacc4130"
         Frmacc4130_Previous
      Case "Frmacc4140"
         Frmacc4140_Previous
      Case "Frmacc4150"
         Frmacc4150_Previous
      Case "Frmacc4160"
         Frmacc4160_Previous
      Case "Frmacc4170"
         Frmacc4170_Previous
      Case "Frmacc4180"
         Frmacc4180_Previous
      Case "Frmacc4190"
         Frmacc4190_Previous
      Case "Frmacc41a0"
         'Frmacc41a0_Previous 'Modify by Amy 2013/12/17
         Frmacc41a0.MovePreviousRecord 'Add 2013/12/17
      Case "Frmacc41b0"
         'Frmacc41b0_Previous 'Moidify by Amy 2013/12/17
         Frmacc41b0.MovePreviousRecord 'Add 2013/12/17
      'Add by Morgan 2004/10/27
      Case "Frmacc41c0"
         Frmacc41c0_Previous
      Case "Frmacc41d0"
         Frmacc41d0.Frmacc41d0_Previous 'Modify by Amy 2014/02/06 搬回form
      'Add by Morgan 2005/4/7
      Case "Frmacc41e0"
         With Frmacc41e0
            .ReadData .txtA2301, 2
         End With
      'Added by Lydia 2017/03/03 財產目錄
      Case "Frmacc41i0"
         Frmacc41i0.Frmacc41i0_Previous
      'Added by Lydia 2017/03/08 財產報廢作業
      Case "Frmacc41i0_1"
         Frmacc41i0_1.Frmacc41i0_1_Previous
      'Add by Amy 2017/04/25
      Case "Frmacc41j0"
         Frmacc41j0.SetData ("PreRec")
      'Add by Amy 2014/02/14
      Case "Frmacc5100"
         Frmacc5100.MovePreviousRecord
      Case "Frmacc5200"
         Frmacc5200_Previous
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
      Case "Frmacc1130"
         Frmacc1130_First
      'Add By Sindy 2014/1/9
      Case "Frmacc11q0"
         Frmacc11q0.Frmacc11q0_First
      '2014/1/9 END
      Case "Frmacc1140"
         Frmacc1140_First
      Case "Frmacc1150"
         Frmacc1150.Frmacc1150_First 'Modify by Amy 2018/06/05
      Case "Frmacc1160"
         Frmacc1160_First
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         Frmacc11p0.Frmacc11p0_First
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         Frmacc11n0.Frmacc11n0_First
      '2012/8/29 End
      Case "Frmacc1170"
         Frmacc1170.Frmacc1170_First
      Case "Frmacc1180"
         Frmacc1180.Frmacc1180_First 'Modify by Amy 2014/01/17
      Case "Frmacc1190"
         Frmacc1190_First
      Case "Frmacc11a0"
         Frmacc11a0_First
      Case "Frmacc11d0"
         Frmacc11d0_First
      Case "Frmacc11f0"
         Frmacc11f0_First
      'Add by Morgan 2007/4/17
      Case "Frmacc11i0"
         Frmacc11i0.MoveFirst
      'Add by Morgan 2007/5/24
      Case "Frmacc11j0"
         Frmacc11j0.MoveFirst
         
      'Add by Morgan 2011/4/11
      Case "Frmacc11l0"
         Frmacc11l0.MoveFirst
      'Add by Amy 2013/11/29
      Case "Frmacc11o0"
         Frmacc11o0.GetFirstRecordVal '.MoveFirst
      Case "Frmacc2110"
         'Add by Morgan 2006/6/20
         If Frmacc2110.bolForm2 = True Then
            MsgBox "匯率有異動，請點【收款資料】以便重新計算台幣收款金額！", vbExclamation
            Exit Sub
         End If
         'end 2006/6/20
         Frmacc2110_First
      Case "Frmacc2120"
         Frmacc2120_First
      Case "Frmacc2130"
         Frmacc2130_First
      Case "Frmacc2140"
         Frmacc2140_First
      Case "Frmacc2150"
         Frmacc2150_First
      Case "Frmacc2160"
         Frmacc2160_First
      Case "Frmacc21d0"
         Frmacc21d0_First
      Case "Frmacc21e0"
         Frmacc21e0_First
      Case "Frmacc21f0"
         Frmacc21f0_First
'Remove by Morgan 2005/1/14 財務不需要
'      Case "Frmacc21g0"
'         Frmacc21g0_First
      Case "Frmacc21h0"
         Frmacc21h0_First
      Case "Frmacc21i0"
         Frmacc21i0_First
      Case "Frmacc21j0"
         Frmacc21j0_First
      Case "Frmacc21k0"
         Frmacc21k0_First
'Remove by Morgan 2005/1/14 財務不用
'Modified by Morgan 2019/10/5 改回由財務維護
      Case "Frmacc21m0"
         Frmacc21m0_First Frmacc21m0
         
      Case "Frmacc21n0"
         Frmacc21n0_First
'Remove by Morgan 2005/8/3 財務不用
'Modified by Morgan 2019/7/10 又改要用
      Case "Frmacc21o0"
         Frmacc21o0_First Frmacc21o0
         
      'Add By Cheng 2003/07/23
      Case "Frmacc21q0"
         Frmacc21q0_First
      'Add by Morgan 2006/12/18
      Case "Frmacc21r0"
         Frmacc21r0.MoveFirst
      
      'Added by Morgan 2019/10/5
      Case "Frmacc21s0"
         Frmacc21s0_First Frmacc21s0
         
      'Added by Lydia 2016/11/7
      Case "Frmacc21w0"
         Frmacc21w0.MoveFirst
      Case "Frmacc3110"
         Frmacc3110_First
      Case "Frmacc3120"
         Frmacc3120_First
      Case "Frmacc3140"
         Frmacc3140_First
      Case "Frmacc3150"
         Frmacc3150_First
      Case "Frmacc3160"
         Frmacc3160_First
      Case "Frmacc3170"
         Frmacc3170_First
      Case "Frmacc3180"
         Frmacc3180_First
      Case "Frmacc3190"
         Frmacc3190_First
      Case "Frmacc31a0"
         Frmacc31a0_First
      Case "Frmacc31c0"
         Frmacc31c0_First
      Case "Frmacc4110"
         Frmacc4110.Frmacc4110_First 'Modify by Amy 2015/06/11 搬回form
      Case "Frmacc4120"
         Frmacc4120.Frmacc4120_First 'Modify by Amy 2014/01/14 搬回form
      Case "Frmacc4130"
         Frmacc4130_First
      Case "Frmacc4140"
         Frmacc4140_First
      Case "Frmacc4150"
         Frmacc4150_First
      Case "Frmacc4160"
         Frmacc4160_First
      Case "Frmacc4170"
         Frmacc4170_First
      Case "Frmacc4180"
         Frmacc4180_First
      Case "Frmacc4190"
         Frmacc4190_First
      Case "Frmacc41a0"
         'Frmacc41a0_First 'Modify by Amy 2013/12/17
         Frmacc41a0.MoveFirstRecord 'Add 2013/12/17
      Case "Frmacc41b0"
         'Frmacc41b0_First 'Modify by Amy 2013/12/17
         Frmacc41b0.MoveFirstRecord 'Add 2013/12/17
      Case "Frmacc41c0"
         Frmacc41c0_First
      Case "Frmacc41d0"
         Frmacc41d0.Frmacc41d0_First 'Modify by Amy 2014/02/06 搬回form
      'Add by Morgan 2005/4/7
      Case "Frmacc41e0"
         With Frmacc41e0
            .ReadData .txtA2301, 1
         End With
      'Added by Lydia 2017/03/03 財產目錄
      Case "Frmacc41i0"
         Frmacc41i0.Frmacc41i0_First
      'Added by Lydia 2017/03/08 財產報廢作業
      Case "Frmacc41i0_1"
         Frmacc41i0_1.Frmacc41i0_1_First
      'Add by Amy 2017/04/25
      Case "Frmacc41j0"
         Frmacc41j0.SetData ("FirstRec")
      'Add by Amy 2014/02/14
      Case "Frmacc5100"
         Frmacc5100.MoveFirstRecord
      Case "Frmacc5200"
         Frmacc5200_First
   End Select
End Sub

'*************************************************
'  離開
'
'*************************************************
Public Sub FormExit()
   Select Case strFormName
      'Add by Amy 2013/12/27
      Case "Frmacc1172"
         Unload Frmacc1172
         Exit Sub
      'end 2013/12/27
      Case "Frmacc1191"
         Unload Frmacc1191
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
      'add by nickc 2005/11/04 從下面搬上來
      Case "Frmacc41a2"
         Unload Frmacc41a2
         Exit Sub
      'Add by Amy 2017/10/11 從下面搬上來
      Case "Frmacc41g0"
         Unload Frmacc41g0
         Exit Sub
      Case "Frmacc41h0"
         Unload Frmacc41h0
         Exit Sub
      Case "Frmacc41j0"
         Unload Frmacc41j0
         Exit Sub
      'end 2017/06/20
      'Add by Amy 2021/09/17
      Case "Frmacc41k0"
         Unload Frmacc41k0
         Exit Sub
      'Add by Amy 2014/04/16
      Case "Frmacc42b0"
         Unload Frmacc42b0
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
      Case "Frmacc1110"
         Unload Frmacc1110
      Case "Frmacc1120"
         Unload Frmacc1120
      Case "Frmacc1121"
         Unload Frmacc1121
      Case "Frmacc1122"
         Unload Frmacc1122
      Case "Frmacc1130"
         Unload Frmacc1130
      'Add by Sindy 2014/1/9
      Case "Frmacc11q0"
         Unload Frmacc11q0
      Case "Frmacc11q1"
         Unload Frmacc11q1
      '2014/1/9 END
      Case "Frmacc1131"
         Unload Frmacc1131
      Case "Frmacc1140"
         Unload Frmacc1140
      Case "Frmacc1141"
         Unload Frmacc1141
      Case "Frmacc1150"
         Unload Frmacc1150
      Case "Frmacc1151"
         Unload Frmacc1151
      Case "Frmacc1152"
         Unload Frmacc1152
      Case "Frmacc1153"
         Unload Frmacc1153
         If strExitControl = MsgText(602) Then
            tool3_enabled
            strExitControl = MsgText(601)
            Exit Sub
         End If
      Case "Frmacc1160"
         Unload Frmacc1160
      Case "Frmacc1161"
         Unload Frmacc1161
      'Add By Sindy 2013/12/19
      Case "Frmacc11p0"
         Unload Frmacc11p0
      Case "Frmacc11p1"
         Unload Frmacc11p1
      '2013/12/19 End
      'Add By Sindy 2012/8/29
      Case "Frmacc11n0"
         Unload Frmacc11n0
      Case "Frmacc11n1"
         Unload Frmacc11n1
      '2012/8/29 End
      'Add By Sindy 2013/12/13
      Case "Frmacc11o5"
         Unload Frmacc11o5
      '2013/12/13 END
      Case "Frmacc1170"
         Unload Frmacc1170
      Case "Frmacc1171"
         Unload Frmacc1171
      Case "Frmacc1180"
         Unload Frmacc1180
      Case "Frmacc1181"
         Unload Frmacc1181
      Case "Frmacc1190"
         Unload Frmacc1190
      Case "Frmacc1191"
         Unload Frmacc1191
      Case "Frmacc1192"
         Unload Frmacc1192
      Case "Frmacc1193"
         Unload Frmacc1193
      Case "Frmacc11a0"
         Unload Frmacc11a0
      Case "Frmacc11a1"
         Unload Frmacc11a1
      Case "Frmacc11b0"
         Unload Frmacc11b0
      Case "Frmacc11b1"
         Unload Frmacc11b1
      Case "Frmacc11c0"
         Unload Frmacc11c0
      'Add By Sindy 2015/8/11
      Case "Frmacc11c1"
         Unload Frmacc11c1
      '2015/8/11 END
      Case "Frmacc11d0"
         Unload Frmacc11d0
      Case "Frmacc11d1"
         Unload Frmacc11d1
      Case "Frmacc11e0"
         Unload Frmacc11e0
      Case "Frmacc11f0"
         Unload Frmacc11f0
      Case "Frmacc11f1"
         Unload Frmacc11f1
      Case "Frmacc11g0"
         Unload Frmacc11g0
      Case "Frmacc11g1"
         Unload Frmacc11g1
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
      Case "Frmacc1250"
         Unload Frmacc1250
      Case "Frmacc1260"
         Unload Frmacc1260
      Case "Frmacc1270"
         Unload Frmacc1270
      Case "Frmacc1271"
         Unload Frmacc1271
      Case "Frmacc1272"
         Unload Frmacc1272
     'Add by Amy 2014/02/05
     Case "Frmacc1273"
         Unload Frmacc1273
     'end 2014/02/05
      Case "Frmacc1410"
         Unload Frmacc1410
      Case "Frmacc1420"
         Unload Frmacc1420
      Case "Frmacc1430"
         Unload Frmacc1430
      Case "Frmacc1440"
         Unload Frmacc1440
      Case "Frmacc1450"
         Unload Frmacc1450
      Case "Frmacc1460"
         Unload Frmacc1460
      Case "Frmacc1470"
         Unload Frmacc1470
      Case "Frmacc1480"
         Unload Frmacc1480
      Case "Frmacc1490"
         Unload Frmacc1490
      Case "Frmacc14a0"
         Unload Frmacc14a0
      Case "Frmacc14b0"
         Unload Frmacc14b0
      Case "Frmacc14c0"
         Unload Frmacc14c0
      Case "Frmacc14d0"
         Unload Frmacc14d0
      Case "Frmacc14e0"
         Unload Frmacc14e0
      Case "Frmacc14f0"
         Unload Frmacc14f0
      Case "Frmacc14g0"
         Unload Frmacc14g0
        'Add By Cheng 2003/12/04
      Case "Frmacc14h0"
         Unload Frmacc14h0
        'End
       'Add by Amy 20158/08/31
      Case "frm100114_5"
         Unload frm100114_5
      Case "Frmacc2110"
         Unload Frmacc2110
      Case "Frmacc2111"
         Unload Frmacc2111
      Case "Frmacc2112"
         Unload Frmacc2112
      Case "Frmacc2120"
         Unload Frmacc2120
      Case "Frmacc2121"
         Unload Frmacc2121
      Case "Frmacc2130"
         Unload Frmacc2130
      Case "Frmacc2131"
         Unload Frmacc2131
      Case "Frmacc2140"
         Unload Frmacc2140
      Case "Frmacc2141"
         Unload Frmacc2141
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
      Case "Frmacc2170"
         Unload Frmacc2170
      Case "Frmacc2180"
         Unload Frmacc2180
      Case "Frmacc2190"
         Unload Frmacc2190
      Case "Frmacc2191"
         Unload Frmacc2191
      Case "Frmacc21b0"
         Unload Frmacc21b0
      Case "Frmacc21c0"
         Unload Frmacc21c0
      Case "Frmacc21d0"
         Unload Frmacc21d0
      Case "Frmacc21d1"
         Unload Frmacc21d1
      Case "Frmacc21e0"
         Unload Frmacc21e0
      Case "Frmacc21e1"
         Unload Frmacc21e1
      Case "Frmacc21f0"
         Unload Frmacc21f0
      Case "Frmacc21f1"
         Unload Frmacc21f1
      Case "Frmacc21f2"
         Unload Frmacc21f2
      Case "Frmacc21f3"
         Unload Frmacc21f3
'Remove by Morgan 2005/1/14 財務不需要
'      Case "Frmacc21g0"
'         Unload Frmacc21g0
'      Case "Frmacc21g1"
'         Unload Frmacc21g1
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
      Case "Frmacc21l0"
         Unload Frmacc21l0
'Remove by Morgan 2005/1/14 財務不用
'Modified by Morgan 2019/10/5 改回由財務維護
      Case "Frmacc21m0"
         Unload Frmacc21m0
         
      Case "Frmacc21n0"
         Unload Frmacc21n0
'Remove by Morgan 2005/8/3 財務不用
'Modified by Morgan 2019/7/10 又改要用
      Case "Frmacc21o0"
         Unload Frmacc21o0
         
'      Case "Frmacc21o1"
'         Unload Frmacc21o1
'Remove by Morgan 2005/1/14 財務不用
'      Case "Frmacc21p0"
'         Unload Frmacc21p0
'      Case "Frmacc21p1"
'         Unload Frmacc21p1
      'Add By Cheng 2003/07/22
      Case "Frmacc21q0"
         Unload Frmacc21q0
         
      'Added by Morgan 2019/10/5
      Case "Frmacc21s0"
         Unload Frmacc21s0
         
      'Added by Lydia 2016/11/07
      Case "Frmacc21w0"
         Unload Frmacc21w0
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
      Case "Frmacc2210"
         Unload Frmacc2210
      Case "Frmacc2220"
         Unload Frmacc2220
      Case "Frmacc2230"
         Unload Frmacc2230
      Case "Frmacc2240"
         Unload Frmacc2240
      Case "Frmacc2250"
         Unload Frmacc2250
      Case "Frmacc2310"
         Unload Frmacc2310
      Case "Frmacc2430"
         Unload Frmacc2430
      Case "Frmacc2440"
         Unload Frmacc2440
      Case "Frmacc2450"
         Unload Frmacc2450
      Case "Frmacc2460"
         Unload Frmacc2460
      Case "Frmacc2470"
         Unload Frmacc2470
      Case "Frmacc2471" 'Add by Amy 2013/11/01
         Unload Frmacc2471
      Case "Frmacc2480"
         Unload Frmacc2480
      Case "Frmacc2490"
         Unload Frmacc2490
      Case "Frmacc24a0"
         Unload Frmacc24a0
      Case "Frmacc24b0"
         Unload Frmacc24b0
'Remove by Morgan 2005/8/3 財務不用
'      Case "Frmacc24c0"
'         Unload Frmacc24c0
      Case "Frmacc24d0"
         Unload Frmacc24d0
      Case "Frmacc24e0"
         Unload Frmacc24e0
      Case "Frmacc24f0"
         Unload Frmacc24f0
      Case "Frmacc24g0"
         Unload Frmacc24g0
'Remove by Morgan 2005/8/3 財務不用
'      Case "Frmacc24h0"
'         Unload Frmacc24h0
      'Add By Cheng 2002/09/05
      'Remove by Morgan 2004/12/28 不用
'      Case "Frmacc24i0"
'         Unload Frmacc24i0
      Case "Frmacc3110"
         Unload Frmacc3110
      Case "Frmacc3111"
         Unload Frmacc3111
      Case "Frmacc3120"
         Unload Frmacc3120
      Case "Frmacc3121"
         Unload Frmacc3121
      Case "Frmacc3130"
         Unload Frmacc3130
      Case "Frmacc3140"
         Unload Frmacc3140
      'Mark by Amy 2022/02/23 過濾Form2.0 未使用表單,發現此功能未顯示,故先刪
'      Case "Frmacc3141"
'         Unload Frmacc3141
      Case "Frmacc3150"
         Unload Frmacc3150
      Case "Frmacc3151"
         Unload Frmacc3151
      Case "Frmacc3160"
         Unload Frmacc3160
      Case "Frmacc3161"
         Unload Frmacc3161
      Case "Frmacc3170"
         Unload Frmacc3170
      Case "Frmacc3171"
         Unload Frmacc3171
      Case "Frmacc3180"
         Unload Frmacc3180
      Case "Frmacc3181"
         Unload Frmacc3181
      Case "Frmacc3190"
         Unload Frmacc3190
      Case "Frmacc3191"
         Unload Frmacc3191
      Case "Frmacc31a0"
         Unload Frmacc31a0
      Case "Frmacc31a1"
         Unload Frmacc31a1
      Case "Frmacc31b0"
         Unload Frmacc31b0
      Case "Frmacc31c0"
         Unload Frmacc31c0
      Case "Frmacc31c1"
         Unload Frmacc31c1
      Case "Frmacc31d0"
         Unload Frmacc31d0
      Case "Frmacc3210"
         Unload Frmacc3210
      'Mark by Amy 2022/02/23 過濾Form2.0 未使用表單,發現此功能未顯示,故先刪
'      Case "Frmacc3220"
'         Unload Frmacc3220
'      Case "Frmacc3230"
'         Unload Frmacc3230
      'end 2022/02/23
      Case "Frmacc3250"
         Unload Frmacc3250
      Case "Frmacc3260"
         Unload Frmacc3260
      Case "Frmacc3270"
         Unload Frmacc3270
      Case "Frmacc3280"
         Unload Frmacc3280
      Case "Frmacc32a0"
         Unload Frmacc32a0
      Case "Frmacc32b0"
         Unload Frmacc32b0
      Case "Frmacc32c0"
         Unload Frmacc32c0
      Case "Frmacc32f0"
         Unload Frmacc32f0
      Case "Frmacc32g0"
         Unload Frmacc32g0
      Case "Frmacc32h0"
         Unload Frmacc32h0
      Case "Frmacc3310"
         Unload Frmacc3310
      Case "Frmacc3320"
         Unload Frmacc3320
      Case "Frmacc3410"
         Unload Frmacc3410
      Case "Frmacc3420"
         Unload Frmacc3420
      Case "Frmacc3430"
         Unload Frmacc3430
      Case "Frmacc3440"
         Unload Frmacc3440
      Case "Frmacc3450"
         Unload Frmacc3450
      Case "Frmacc3460"
         Unload Frmacc3460
      Case "Frmacc3470"
         Unload Frmacc3470
      Case "Frmacc3480"
         Unload Frmacc3480
      Case "Frmacc3490"
         Unload Frmacc3490
      Case "Frmacc34a0"
         Unload Frmacc34a0
      Case "Frmacc34b0"
         Unload Frmacc34b0
      Case "Frmacc34c0"
         Unload Frmacc34c0
      Case "Frmacc34d0"
         Unload Frmacc34d0
      Case "Frmacc34e0"
         Unload Frmacc34e0
      Case "Frmacc34g0"
         Unload Frmacc34g0
      Case "Frmacc34h0"
         Unload Frmacc34h0
      Case "Frmacc34i0"
         Unload Frmacc34i0
      Case "Frmacc34j0"
         Unload Frmacc34j0
      Case "Frmacc4110"
         Unload Frmacc4110
      Case "Frmacc4111"
         Unload Frmacc4111
      Case "Frmacc4120"
         Unload Frmacc4120
      Case "Frmacc4121"
         Unload Frmacc4121
      Case "Frmacc4130"
         Unload Frmacc4130
      Case "Frmacc4131"
         Unload Frmacc4131
      Case "Frmacc4140"
         Unload Frmacc4140
      Case "Frmacc4141"
         Unload Frmacc4141
      Case "Frmacc4160"
         Unload Frmacc4160
      Case "Frmacc4161"
         Unload Frmacc4161
      Case "Frmacc4170"
         Unload Frmacc4170
     'Add by Amy 2014/09/11
      Case "Frmacc4170_1"
         Unload Frmacc4170_1
      Case "Frmacc4171"
         Unload Frmacc4171
      Case "Frmacc4180"
         Unload Frmacc4180
      Case "Frmacc4181"
         Unload Frmacc4181
      Case "Frmacc4190"
         Unload Frmacc4190
      Case "Frmacc4191"
         Unload Frmacc4191
      Case "Frmacc41a0"
         Unload Frmacc41a0
      Case "Frmacc41a1"
         Unload Frmacc41a1
'edit by nickc 2005/11/04 往上搬
'      Case "Frmacc41a2"
'         Unload Frmacc41a2
      Case "Frmacc41b0"
         Unload Frmacc41b0
      Case "Frmacc41b1"
         Unload Frmacc41b1
      Case "Frmacc41c0"
         Unload Frmacc41c0
      Case "Frmacc41d0"
         Unload Frmacc41d0
      Case "Frmacc41d1"
         Unload Frmacc41d1
      Case "Frmacc4210"
         Unload Frmacc4210
      Case "Frmacc4220"
         Unload Frmacc4220
      Case "Frmacc4230"
         Unload Frmacc4230
      Case "Frmacc4240"
         Unload Frmacc4240
      Case "Frmacc4250"
         Unload Frmacc4250
      Case "Frmacc4260"
         Unload Frmacc4260
      Case "Frmacc4270"
         Unload Frmacc4270
        'Add By Cheng 2003/05/27
      Case "Frmacc4280"
         Unload Frmacc4280
      'Add by Amy 2016/07/11
      'Modified by Lydia 2017/03/03
      'Case "Frmacc41i0"
      '   'Unload Frmacc41i0
      Case "Frmacc41i0"
         Unload Frmacc41i0
      'end 2017/03/03
      'Added by Lydia 2017/03/08
      Case "Frmacc41i0_1"
         Unload Frmacc41i0_1
      'end 2017/03/08
      'Added by Lydia 2017/05/19
      Case "Frmacc41i1"
         Unload Frmacc41i1
      'end 2017/05/19
      'Add by Amy 2016/01/11
      Case "Frmacc43c0"
         Unload Frmacc43c0
      Case "Frmacc4310"
         'Unload Frmacc4310
      Case "Frmacc4320"
         Unload Frmacc4320
      Case "Frmacc4330"
         Unload Frmacc4330
      Case "Frmacc4340"
         Unload Frmacc4340
      Case "Frmacc4350"
         Unload Frmacc4350
      Case "Frmacc4360"
         Unload Frmacc4360
      Case "Frmacc4410"
         Unload Frmacc4410
      Case "Frmacc4420"
         Unload Frmacc4420
      Case "Frmacc4430"
         Unload Frmacc4430
      Case "Frmacc4440"
         Unload Frmacc4440
      Case "Frmacc4450"
         Unload Frmacc4450
      Case "Frmacc4460"
         Unload Frmacc4460
      Case "Frmacc4470"
         Unload Frmacc4470
      Case "Frmacc4480"
         Unload Frmacc4480
      Case "Frmacc4490"
         Unload Frmacc4490
      Case "Frmacc44a0"
         Unload Frmacc44a0
      Case "Frmacc44b0"
         Unload Frmacc44b0
      Case "Frmacc44c0"
         Unload Frmacc44c0
      Case "Frmacc44d0"
         Unload Frmacc44d0
      Case "Frmacc44e0"
         Unload Frmacc44e0
      Case "Frmacc44g0"
         Unload Frmacc44g0
      Case "Frmacc44h0"
         Unload Frmacc44h0
      Case "Frmacc44i0"
         Unload Frmacc44i0
      Case "Frmacc44j0"
         Unload Frmacc44j0
      Case "Frmacc44k0"
         Unload Frmacc44k0
      Case "Frmacc44l0"
         Unload Frmacc44l0
      Case "Frmacc44m0"
         Unload Frmacc44m0
'Removed by Morgan 2013/7/3
'      Case "Frmacc44o0"
'         Unload Frmacc44o0
      Case "Frmacc44p0"
         Unload Frmacc44p0
      Case "Frmacc44q0"
         Unload Frmacc44q0
      Case "Frmacc44r0"
         Unload Frmacc44r0
      Case "Frmacc5100"
         Unload Frmacc5100
      Case "Frmacc5200"
         Unload Frmacc5200
      Case "Frmacc6200"
         Unload Frmacc6200
      Case Else
         Dim oForm
         For Each oForm In Forms
            If oForm.Name = strFormName Then Unload oForm: Exit For
         Next
   End Select
End Sub
